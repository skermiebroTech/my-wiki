#!/bin/zsh
# Build a G1 custom-labels template URL prefilled with this Mac's hardware info
# and open it in the default browser. Paste the whole block into Terminal.
# SKU / grade / G1 number are not machine-derivable - edit the three vars below or on the canvas.

SKU="DEMO"
GRADE="B"
G1ID="1234567"

MODEL=$(system_profiler SPHardwareDataType | awk -F': ' '/Model Name/{print $2}')
SERIAL=$(system_profiler SPHardwareDataType | awk -F': ' '/Serial Number/{print $2}')
CPU=$(sysctl -n machdep.cpu.brand_string 2>/dev/null)
[ -z "$CPU" ] && CPU=$(system_profiler SPHardwareDataType | awk -F': ' '/Chip/{print $2}')
RAM_GB=$(( $(sysctl -n hw.memsize) / 1073741824 ))
DRIVE=$(df -H / | awk 'NR==2{print $2}' | sed 's/G/GB/;s/T/TB/')
DESIGN=$(ioreg -rn AppleSmartBattery | awk -F'= ' '/"DesignCapacity"/{print $2}' | head -1)
FCC=$(ioreg -rn AppleSmartBattery | awk -F'= ' '/"AppleRawMaxCapacity"/{print $2}' | head -1)
[ -z "$FCC" ] && FCC=$(ioreg -rn AppleSmartBattery | awk -F'= ' '/"MaxCapacity"/{print $2}' | head -1)
if [ -n "$DESIGN" ] && [ "$DESIGN" -gt 0 ] 2>/dev/null; then
  BATT="$(( FCC * 100 / DESIGN ))%"; CAP="${FCC}/${DESIGN} mAh"
else
  BATT="N/A"; CAP="no battery"
fi
DATE=$(date +%d/%m/%Y)

SKU="$SKU" GRADE="$GRADE" G1ID="$G1ID" MODEL="$MODEL" SERIAL="$SERIAL" CPU="$CPU" \
RAM_GB="$RAM_GB" DRIVE="$DRIVE" BATT="$BATT" CAP="$CAP" DATE="$DATE" python3 - <<'PYEOF'
import json, os, math, base64, urllib.parse, subprocess

e = os.environ
PI = math.pi
UP = PI * 1.5          # vertical text, reads bottom-to-top
def dx(d): return -100 + d * 800 / 816
def dy(d): return d * 400 / 500

idx = [0]
IDX = "123456789ABCDEFG"
def text(x, y, s, size, angle=0.0, centered=False):
    idx[0] += 1
    w = round(len(s) * size * 0.6, 1); h = round(size * 1.35, 1)
    if centered: x -= w / 2; y -= h / 2
    return {"id": f"el-{idx[0]}", "type": "text", "x": round(x, 2), "y": round(y, 2),
        "width": w, "height": h, "angle": angle, "strokeColor": "#000000",
        "backgroundColor": "transparent", "fillStyle": "solid", "strokeWidth": 2,
        "strokeStyle": "solid", "roughness": 0, "opacity": 100, "groupIds": [],
        "frameId": "label-area-frame", "index": "a" + IDX[idx[0] - 1], "roundness": None,
        "seed": 1000 + idx[0], "version": 1, "versionNonce": 2000 + idx[0],
        "isDeleted": False, "boundElements": None, "updated": 1787713917988,
        "link": None, "locked": False, "text": s, "fontSize": size, "fontFamily": 6,
        "textAlign": "left", "verticalAlign": "top", "containerId": None,
        "originalText": s, "autoResize": True, "lineHeight": 1.35}

els = [
    text(dx(85), dy(170), e["SKU"], 90, UP, True),
    text(dx(25), dy(340), "SKU", 10, PI),
    text(dx(172), dy(0), e["GRADE"], 140),
    text(dx(192), dy(180), e["BATT"], 16),
    text(dx(345), dy(5), e["G1ID"], 80),
    text(dx(350), dy(100), e["MODEL"], 12),
    text(dx(350), dy(135), "SN: " + e["SERIAL"], 12),
    text(dx(350), dy(170), e["CAP"], 12),
    text(dx(350), dy(205), e["CPU"], 34),
    text(dx(345), dy(260), e["RAM_GB"] + " GB RAM", 34),
    text(dx(350), dy(320), e["DRIVE"], 34),
    text(dx(770), dy(250), e["DATE"] + " JS", 24, UP, True),
    text(dx(738), dy(250), "SYSPREPPED", 24, UP, True),
]
idx[0] += 1
els.append({"id": "qr-placeholder", "type": "rectangle", "x": round(dx(150), 2),
    "y": round(dy(190), 2), "width": 130, "height": 130, "angle": 0,
    "strokeColor": "#000000", "backgroundColor": "transparent", "fillStyle": "solid",
    "strokeWidth": 2, "strokeStyle": "solid", "roughness": 0, "opacity": 100,
    "groupIds": [], "frameId": "label-area-frame", "index": "a" + IDX[idx[0] - 1],
    "roundness": None, "seed": 999, "version": 1, "versionNonce": 1999,
    "isDeleted": False, "boundElements": None, "updated": 1787713917988,
    "link": None, "locked": False})
els.append({"id": "label-area-frame", "type": "frame", "x": -100, "y": 0,
    "width": 800, "height": 400, "angle": 0, "strokeColor": "#e03131",
    "backgroundColor": "transparent", "fillStyle": "solid", "strokeWidth": 2,
    "strokeStyle": "dashed", "roughness": 0, "opacity": 100, "groupIds": [],
    "frameId": None, "index": "a0", "roundness": None, "seed": 283113, "version": 1,
    "versionNonce": 481323, "isDeleted": False, "boundElements": None,
    "updated": 1787713903959, "link": None, "locked": True,
    "name": "Label Area (100x50mm)"})

app_state = {"showWelcomeScreen": False, "theme": "light", "collaborators": {},
    "activeTool": {"type": "selection", "customType": None, "locked": False,
                   "lastActiveTool": None},
    "penMode": False, "penDetected": False, "exportBackground": True, "exportScale": 1,
    "exportEmbedScene": False, "exportWithDarkMode": False, "gridSize": 20,
    "gridStep": 5, "gridModeEnabled": False, "isBindingEnabled": True,
    "defaultSidebarDockedPreference": False, "isLoading": False,
    "lastPointerDownWith": "mouse", "name": "G1 Sticker",
    "pasteDialog": {"shown": False, "data": None}, "previousSelectedElementIds": {},
    "scrolledOutside": False, "scrollX": 150, "scrollY": 50,
    "selectedElementIds": {}, "hoveredElementIds": {}, "selectedGroupIds": {},
    "stats": {"open": False, "panels": 3}, "suggestedBindings": [],
    "frameRendering": {"enabled": True, "clip": True, "name": True, "outline": True},
    "viewBackgroundColor": "#ffffff", "zenModeEnabled": False, "zoom": {"value": 1},
    "viewModeEnabled": False, "snapLines": [], "objectsSnapModeEnabled": False,
    "followedBy": {}, "searchMatches": []}

tpl = {"name": "G1 Sticker " + e["SERIAL"],
       "data": {"elements": els, "appState": app_state},
       "dimensions": {"width": 100, "height": 50, "rotate_label": False,
                      "value": "100x50", "customWidth": "50", "customHeight": "25",
                      "customRotated": False}}
frag = base64.b64encode(
    urllib.parse.quote(json.dumps(tpl, separators=(",", ":")),
                       safe="!'()*-._~").encode()).decode()
url = "https://advantage.g1.com.au/custom-labels#template=" + frag
subprocess.run(["open", url])
print("Opened. URL length:", len(url))
PYEOF
