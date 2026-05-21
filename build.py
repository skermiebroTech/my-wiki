#!/usr/bin/env python3
"""
LapCheck — portable build script.

Produces a single-file executable suitable for dropping on a USB stick.
Run on the *target* OS (PyInstaller doesn't cross-compile reliably).

Usage:
    python build.py                 # build for current OS, release mode
    python build.py --debug         # keep console window (Windows) / verbose
    python build.py --clean         # wipe build/ and dist/ first
    python build.py --name MyApp    # override executable name

Output:
    build/dist/LapCheck(.exe)       # the portable binary
    build/dist/launcher.cmd         # Windows only — convenience launcher
"""
from __future__ import annotations

import argparse
import os
import shutil
import subprocess
import sys
from pathlib import Path

# Windows GitHub Actions runners default stdout to cp1252. Any unicode in our
# print output (✓, em-dashes, etc) will then crash with UnicodeEncodeError.
# Force UTF-8 on the streams we own.
for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(encoding="utf-8")
    except Exception:
        pass

ROOT = Path(__file__).resolve().parent
SRC = ROOT / "src"
WEB = SRC / "ui" / "web"
BUILD = ROOT / "build"
DIST = BUILD / "dist"
WORK = BUILD / "work"
SPEC = BUILD / "spec"


def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(description="Build LapCheck into a portable executable.")
    p.add_argument("--name", default="LapCheck", help="Output executable name (default: LapCheck)")
    p.add_argument("--debug", action="store_true", help="Keep console + verbose output")
    p.add_argument("--clean", action="store_true", help="Remove build/ before building")
    p.add_argument("--onedir", action="store_true",
                   help="Produce a folder bundle instead of single-file (faster startup, more files)")
    return p.parse_args()


def ensure_pyinstaller() -> None:
    try:
        import PyInstaller  # noqa: F401
    except ImportError:
        print("PyInstaller not installed. Run: pip install -r requirements.txt", file=sys.stderr)
        sys.exit(1)


def clean() -> None:
    for d in (BUILD,):
        if d.exists():
            print(f"  rm -rf {d}")
            shutil.rmtree(d)


def platform_data_sep() -> str:
    # PyInstaller --add-data uses ':' on POSIX, ';' on Windows
    return ";" if sys.platform.startswith("win") else ":"


def build(args: argparse.Namespace) -> Path:
    BUILD.mkdir(exist_ok=True)
    WORK.mkdir(exist_ok=True)
    SPEC.mkdir(exist_ok=True)

    sep = platform_data_sep()

    cmd: list[str] = [
        sys.executable, "-m", "PyInstaller",
        str(SRC / "main.py"),
        "--name", args.name,
        "--distpath", str(DIST),
        "--workpath", str(WORK),
        "--specpath", str(SPEC),
        "--noconfirm",
        "--clean",
        # Embed the entire UI web folder so the bundled binary can serve it
        "--add-data", f"{WEB}{sep}ui/web",
        # Embed the SKU catalogue CSV
        "--add-data", f"{ROOT / 'data'}{sep}data",
    ]

    if args.onedir:
        cmd.append("--onedir")
    else:
        cmd.append("--onefile")

    if args.debug:
        cmd.append("--debug=all")
    else:
        # Hide console on Windows; harmless elsewhere
        cmd.append("--windowed")

    # Hidden imports — pywebview's backend modules are sometimes missed by
    # static analysis when loaded conditionally per-platform.
    # (The PyPI distribution is 'pywebview' but the import name is 'webview'.)
    hidden = [
        "webview",
        "webview.platforms.edgechromium",   # Windows (WebView2)
        "webview.platforms.gtk",            # Linux
        "webview.platforms.qt",             # Linux fallback
        "webview.platforms.cocoa",          # macOS
        "psutil",
    ]
    if sys.platform.startswith("win"):
        hidden += ["wmi", "win32com.client"]

    for h in hidden:
        cmd += ["--hidden-import", h]

    # Make sure imports inside src/ resolve when frozen
    cmd += ["--paths", str(ROOT)]

    print("==> PyInstaller command:")
    print("    " + " ".join(cmd))

    env = os.environ.copy()
    # Faster, deterministic builds
    env.setdefault("PYTHONHASHSEED", "0")

    rc = subprocess.call(cmd, env=env)
    if rc != 0:
        print(f"\nBuild failed (PyInstaller exit code {rc})", file=sys.stderr)
        sys.exit(rc)

    # Locate output
    exe_name = args.name + (".exe" if sys.platform.startswith("win") else "")
    out = DIST / exe_name
    if not out.exists() and args.onedir:
        out = DIST / args.name / exe_name
    return out


def write_launcher(exe_path: Path) -> None:
    """Drop a launcher.cmd alongside the .exe so users can also double-click."""
    if not sys.platform.startswith("win"):
        return
    launcher = DIST / "launcher.cmd"
    launcher.write_text(
        "@echo off\r\n"
        ":: LapCheck launcher - runs the diagnostics tool from the same folder\r\n"
        ":: (e.g. the USB root) without keeping the cmd window around.\r\n"
        f'start "" /b "%~dp0{exe_path.name}"\r\n',
        encoding="utf-8",
    )
    print(f"==> Wrote {launcher}")


def report(exe_path: Path) -> None:
    size_mb = exe_path.stat().st_size / 1024 / 1024
    print()
    print("=" * 56)
    print(f"  [OK] Built: {exe_path}")
    print(f"  [OK] Size:  {size_mb:.1f} MB")
    print("=" * 56)
    print()
    print("Next steps:")
    print(f"  1. Copy '{exe_path.name}' to the root of your USB drive.")
    if sys.platform.startswith("win"):
        print(f"  2. Optionally copy 'launcher.cmd' next to it.")
        print(f"  3. During OOBE press Shift+F10 (or Win+R when available), then run:")
        print(f"         D:\\{exe_path.name}    (replace D: with your USB letter)")
    else:
        print(f"  2. Make it executable: chmod +x {exe_path.name}")
        print(f"  3. Run from the USB: ./{exe_path.name}")


def main() -> int:
    args = parse_args()
    ensure_pyinstaller()
    if args.clean:
        print("==> Cleaning build directory")
        clean()
    print(f"==> Building LapCheck for {sys.platform} ({'debug' if args.debug else 'release'})")
    exe = build(args)
    write_launcher(exe)
    report(exe)
    return 0


if __name__ == "__main__":
    sys.exit(main())
