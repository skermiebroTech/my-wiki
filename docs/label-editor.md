# Label Editor

Edit the fields, check the preview, then copy the print command and paste it into a **PowerShell window** (too long for Win+R).

<style>
.le-grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(220px, 1fr)); gap: 8px 12px; margin-bottom: 12px; }
.le-grid label { display: flex; flex-direction: column; font-size: 0.75rem; font-weight: 600; }
.le-grid input { padding: 4px 6px; font-size: 0.8rem; border: 1px solid #8884; border-radius: 4px; background: var(--md-default-bg-color); color: var(--md-default-fg-color); }
#le-zpl { width: 100%; height: 260px; font-family: monospace; font-size: 0.75rem; border: 1px solid #8884; border-radius: 4px; background: var(--md-default-bg-color); color: var(--md-default-fg-color); }
#le-preview { max-width: 100%; border: 1px solid #8884; border-radius: 4px; margin-top: 8px; background: #fff; }
.le-btn { padding: 6px 14px; margin: 6px 6px 6px 0; border: none; border-radius: 4px; background: var(--md-primary-fg-color); color: #fff; cursor: pointer; font-size: 0.8rem; }
#le-cmd, #le-cmd-mac { width: 100%; height: 90px; font-family: monospace; font-size: 0.7rem; }
</style>

## Fields

<div class="le-grid">
<label>SKU <input id="f-sku" value="DEMO"></label>
<label>Grade <input id="f-grade" value="B"></label>
<label>Battery % <input id="f-batt" value="80%"></label>
<label>Asset ID (also QR) <input id="f-asset" value="1234567"></label>
<label>Model <input id="f-model" value="Example Laptop 9000"></label>
<label>Serial <input id="f-serial" value="SN12345678"></label>
<label>Full charge mWh <input id="f-fcc" value="40000"></label>
<label>Design mWh <input id="f-design" value="50000"></label>
<label>CPU <input id="f-cpu" value="i5-1234X"></label>
<label>RAM (GB) <input id="f-ram" value="16"></label>
<label>Drive <input id="f-drive" value="256GB"></label>
<label>Date + initials <input id="f-date" value=""></label>
<label>Sysprep text <input id="f-sysprep" value="SYSPREPPED"></label>
<label>Printer IP <input id="f-ip" value="172.17.21.186"></label>
</div>

<button class="le-btn" id="le-regen">Rebuild ZPL from fields</button>

## ZPL (editable)

<textarea id="le-zpl" spellcheck="false"></textarea>

<button class="le-btn" id="le-prev">Update preview</button>
<button class="le-btn" id="le-copy">Copy print command</button>
<span id="le-msg"></span>

<img id="le-preview" alt="Label preview appears here">

## Print command

Windows — paste into a PowerShell window (not Win+R, not cmd):

<textarea id="le-cmd" readonly spellcheck="false"></textarea>

macOS — paste into Terminal:

<textarea id="le-cmd-mac" readonly spellcheck="false"></textarea>

<script>
(function () {
  function v(id) { return document.getElementById(id).value; }
  function buildZpl(skuH, skuW) {
    // SKU runs vertically (rotated B). Font size is fitted iteratively against the
    // rendered preview (autoFitSku). Uses scalable font 0 (not the original bitmap
    // font R) so the size can be tuned continuously to span the target exactly.
    var skuY = 10;
    var skuX = 10 + Math.max(0, Math.round((150 - skuH) / 2));
    return '^XA\n\n^LS2\n^SZ2\n^PW816\n^PON\n^PR14,14\n^PMN\n^MNY\n^LS0\n^MTD\n^MD30\n\n' +
      '\n^FX String that says SKU\n^FS\n^FO25,340\n^ARI,12,70\n^FDSKU\n' +
      '\n^FX String that says SKU\n^FS\n^FO' + skuX + ',' + skuY + '\n^A0B,' + skuH + ',' + skuW + '\n^FD' + v('f-sku') + '\n' +
      '\n^FX String that says the grade\n^FS\n^FO172,0\n^ARN,175,225\n^FD' + v('f-grade') + '\n' +
      '\n^FX String that says the battery percentage\n^FS\n^FO200,155\n^ARN,25,12\n^FD' + v('f-batt') + '\n' +
      '\n^FX The QR Code\n^FS\n^FO150,190\n^BQN,2,8\n^FDMA,' + v('f-asset') + '\n' +
      '\n^FX String that says Asset Number\n^FS\n^FO345,5\n^ARN,120,100\n^FD' + v('f-asset') + '\n' +
      '\n^FX String that says the Laptop Model\n^FS\n^FO350,100\n^ARN,10,5\n^FB350,6,5,L,0\n^FD' + v('f-model') + '\n' +
      '\n^FX String that says the Serial Number\n^FS\n^FO350,135\n^ARN,10,5\n^FB350,6,5,L,0\n^FDSN: ' + v('f-serial') + '\n' +
      '\n^FX String that says the Battery Capacity\n^FS\n^FO350,170\n^ARN,10,5\n^FB350,6,5,L,0\n^FD' + v('f-fcc') + '/' + v('f-design') + ' mWh\n' +
      '\n^FX String that says the CPU\n^FS\n^FO350,205\n^ARN,60,50\n^FB350,6,5,L,0\n^FD' + v('f-cpu') + '\n' +
      '\n^FX String that says the RAM\n^FS\n^FO345,260\n^ARN,60,50\n^FB350,6,5,L,0\n^FD' + v('f-ram') + ' GB RAM\n' +
      '\n^FX String that says the DRIVE\n^FS\n^FO350,320\n^ARN,60,50\n^FB350,6,5,L,0\n^FD' + v('f-drive') + '\n' +
      '\n^FX String that says the DATE INSTALLED\n^FS\n^FO760,130\n^ARB,40,40\n^FD' + v('f-date') + '\n' +
      '\n^FX String that says the SYS PREP\n^FS\n^FO730,130\n^ARB,40,40\n^FD' + v('f-sysprep') + '\n' +
      '\n^XZ\n';
  }
  function updateCmd() {
    var zpl = document.getElementById('le-zpl').value;
    var b64 = btoa(unescape(encodeURIComponent(zpl)));
    var cmd = "$b=[Convert]::FromBase64String('" + b64 + "');" +
      "$c=New-Object Net.Sockets.TcpClient('" + v('f-ip') + "',9100);" +
      "$s=$c.GetStream();$s.Write($b,0,$b.Length);$c.Close()";
    document.getElementById('le-cmd').value = cmd;
    document.getElementById('le-cmd-mac').value =
      "echo '" + b64 + "' | base64 -d | nc -w 3 " + v('f-ip') + ' 9100';
    return cmd;
  }
  function renderZpl(zpl) {
    return fetch('https://api.labelary.com/v1/printers/8dpmm/labels/4x3/0/', {
      method: 'POST',
      headers: { 'Accept': 'image/png', 'Content-Type': 'application/x-www-form-urlencoded' },
      body: zpl
    }).then(function (r) {
      if (!r.ok) throw new Error('Labelary HTTP ' + r.status);
      return r.blob();
    });
  }
  function showBlob(b) {
    document.getElementById('le-preview').src = URL.createObjectURL(b);
    document.getElementById('le-msg').textContent = '';
  }
  function updatePreview() {
    renderZpl(document.getElementById('le-zpl').value).then(showBlob).catch(function (e) {
      document.getElementById('le-msg').textContent = 'Preview failed: ' + e.message;
    });
  }
  // Measure the vertical pixel extent of the SKU column (x < 160) in a rendered label
  function measureSku(blob) {
    return createImageBitmap(blob).then(function (img) {
      var cv = document.createElement('canvas');
      cv.width = img.width; cv.height = img.height;
      var cx = cv.getContext('2d');
      cx.drawImage(img, 0, 0);
      var px = cx.getImageData(0, 0, 160, img.height).data;
      var minY = -1, maxY = -1;
      for (var y = 0; y < img.height; y++) {
        for (var x = 0; x < 160; x++) {
          if (px[(y * 160 + x) * 4] < 128) { if (minY < 0) minY = y; maxY = y; break; }
        }
      }
      return maxY - minY;
    });
  }
  // Iteratively size the SKU font until the rendered text spans from the label edge
  // (y=10) down to the small SKU caption (which starts at ~y335)
  function autoFitSku() {
    var TARGET = 320;
    var skuLen = Math.max(1, v('f-sku').length);
    var w = Math.floor(TARGET / (skuLen * 0.55));
    var attempt = 0;
    function step() {
      var h = Math.min(150, Math.round(150 * w / 180));
      var x = 10 + Math.max(0, Math.round((150 - h) / 2));
      // measure the SKU field alone so the caption's ink doesn't skew the extent
      var probe = '^XA^PW816^FO' + x + ',10^A0B,' + h + ',' + w + '^FD' + v('f-sku') + '^FS^XZ';
      return renderZpl(probe).then(measureSku).then(function (extent) {
        attempt++;
        if (extent > 0 && Math.abs(extent - TARGET) > 10 && attempt < 5) {
          w = Math.max(8, Math.min(600, Math.round(w * TARGET / extent)));
          return step();
        }
        var zpl = buildZpl(h, w);
        document.getElementById('le-zpl').value = zpl;
        updateCmd();
        return renderZpl(zpl).then(showBlob);
      });
    }
    return step();
  }
  function regen() {
    document.getElementById('le-zpl').value = buildZpl(150, 180);
    updateCmd();
    autoFitSku().catch(function (e) {
      document.getElementById('le-msg').textContent = 'Preview failed: ' + e.message;
    });
  }
  document.getElementById('le-regen').addEventListener('click', regen);
  document.getElementById('le-prev').addEventListener('click', function () { updateCmd(); updatePreview(); });
  document.getElementById('le-copy').addEventListener('click', function () {
    navigator.clipboard.writeText(updateCmd()).then(function () {
      document.getElementById('le-msg').textContent = 'Copied!';
    });
  });
  document.getElementById('le-zpl').addEventListener('input', updateCmd);
  var d = new Date();
  document.getElementById('f-date').value =
    ('0' + d.getDate()).slice(-2) + '/' + ('0' + (d.getMonth() + 1)).slice(-2) + '/' + d.getFullYear() + ' JS';
  regen();
})();
</script>