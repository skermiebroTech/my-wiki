#!/bin/zsh
# Print a G1-style sticker to the Zebra label printer, autofilled from this Mac's hardware.
# Usage: zsh Print-G1Sticker-mac.sh
# Or run via: curl -s https://raw.githubusercontent.com/skermiebroTech/my-wiki/main/Print-G1Sticker-mac.sh | zsh

PRINTER_IP="172.17.31.195"
PRINTER_PORT=9100

# Read prompts from the terminal so this works when piped from curl.
if [ -r /dev/tty ]; then
  TTY=/dev/tty
else
  TTY=/dev/stdin
fi

printf 'SKU: ' > /dev/tty; read SKU < "$TTY"
printf 'Grade: ' > /dev/tty; read GRADE < "$TTY"
printf 'Initials: ' > /dev/tty; read INITIALS < "$TTY"
printf 'Asset tag: ' > /dev/tty; read G1ID < "$TTY"

[ -z "$G1ID" ] && G1ID="1234567"

MODEL=$(system_profiler SPHardwareDataType | awk -F': ' '/Model Name/{print $2}')
SERIAL=$(system_profiler SPHardwareDataType | awk -F': ' '/Serial Number/{print $2}')
CPU=$(sysctl -n machdep.cpu.brand_string 2>/dev/null)
[ -z "$CPU" ] && CPU=$(system_profiler SPHardwareDataType | awk -F': ' '/Chip/{print $2}')
RAM_GB=$(( $(sysctl -n hw.memsize) / 1073741824 ))
DRIVE=$(df -H / | awk 'NR==2{print $2}' | sed 's/G/GB/;s/T/TB/')
POWER=$(system_profiler SPPowerDataType)
BATT=$(echo "$POWER" | awk -F': ' '/Maximum Capacity/{gsub(/[ \t]/,"",$2); print $2}')
CYCLES=$(echo "$POWER" | awk -F': ' '/Cycle Count/{gsub(/[ \t]/,"",$2); print $2}')
CONDITION=$(echo "$POWER" | awk -F': ' '/Condition/{gsub(/^[ \t]+/,"",$2); print $2}')
if [ -n "$BATT" ]; then
  CAP="${CYCLES} cycles, ${CONDITION}"
else
  BATT="N/A"; CAP="no battery"
fi
DATE="$(date +%d/%m/%Y) ${INITIALS:-JS}"

# SKU font: scalable font 0 sized so the text spans ~320 dots (label edge to SKU caption)
SKULEN=${#SKU}
if [ $SKULEN -gt 0 ]; then
  SKUW=$(( 31000 / (SKULEN * 55) ))
  SKUH=$(( 150 * SKUW / 180 )); [ $SKUH -gt 150 ] && SKUH=150
  SKUX=$(( 10 + (150 - SKUH) / 2 ))
else
  SKUW=150; SKUH=127; SKUX=21
fi

ZPL="^XA
^LS2^SZ2^PW816^PON^PR14,14^PMN^MNY^LS0^MTD^MD30
^FS^FO25,340^ARI,12,70^FDSKU
^FS^FO${SKUX},10^A0B,${SKUH},${SKUW}^FD${SKU}
^FS^FO172,0^ARN,175,225^FD${GRADE}
^FS^FO200,155^ARN,25,12^FD${BATT}
^FS^FO150,190^BQN,2,8^FDMA,${G1ID}
^FS^FO345,5^ARN,120,100^FD${G1ID}
^FS^FO350,100^ARN,10,5^FB350,6,5,L,0^FD${MODEL}
^FS^FO350,135^ARN,10,5^FB350,6,5,L,0^FDSN: ${SERIAL}
^FS^FO350,170^ARN,10,5^FB350,6,5,L,0^FD${CAP}
^FS^FO350,205^ARN,60,50^FB350,6,5,L,0^FD${CPU}
^FS^FO345,260^ARN,60,50^FB350,6,5,L,0^FD${RAM_GB} GB RAM
^FS^FO350,320^ARN,60,50^FB350,6,5,L,0^FD${DRIVE}
^FS^FO760,130^ARB,40,40^FD${DATE}
^FS^FO730,130^ARB,40,40^FDSYSPREPPED
^XZ"

echo "SKU=$SKU GRADE=$GRADE G1=$G1ID | $MODEL $SERIAL | $CPU ${RAM_GB}GB $DRIVE | $BATT $CAP"
printf '%s' "$ZPL" | nc -w 3 "$PRINTER_IP" "$PRINTER_PORT" && echo "Label sent." || echo "Failed to reach printer $PRINTER_IP:$PRINTER_PORT"
