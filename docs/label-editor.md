# Label Editor

Edit the fields, check the preview, then copy the print command and paste it into a **PowerShell window** (too long for Win+R).

<style>
.le-grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(220px, 1fr)); gap: 8px 12px; margin-bottom: 12px; }
.le-grid label { display: flex; flex-direction: column; font-size: 0.75rem; font-weight: 600; }
.le-grid input { padding: 4px 6px; font-size: 0.8rem; border: 1px solid #8884; border-radius: 4px; background: var(--md-default-bg-color); color: var(--md-default-fg-color); }
#le-zpl { width: 100%; height: 260px; font-family: monospace; font-size: 0.75rem; border: 1px solid #8884; border-radius: 4px; background: var(--md-default-bg-color); color: var(--md-default-fg-color); }
#le-preview { max-width: 100%; border: 1px solid #8884; border-radius: 4px; margin-top: 8px; background: #fff; }
.le-btn { padding: 6px 14px; margin: 6px 6px 6px 0; border: none; border-radius: 4px; background: var(--md-primary-fg-color); color: #fff; cursor: pointer; font-size: 0.8rem; }
#le-cmd { width: 100%; height: 90px; font-family: monospace; font-size: 0.7rem; }
</style>

## Fields

<div class="le-grid">
<label>SKU <input id="f-sku" value="dono"></label>
<label>Grade <input id="f-grade" value="B"></label>
<label>Battery % <input id="f-batt" value="68%"></label>
<label>G1 ID (also QR) <input id="f-g1" value="5742320"></label>
<label>Model <input id="f-model" value="HP EliteBook x360 1040 G6"></label>
<label>Serial <input id="f-serial" value="5CG01773N4"></label>
<label>Full charge mWh <input id="f-fcc" value="38207"></label>
<label>Design mWh <input id="f-design" value="56210"></label>
<label>CPU <input id="f-cpu" value="i5-8365U"></label>
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

Paste into a PowerShell window (not Win+R, not cmd):

<textarea id="le-cmd" readonly spellcheck="false"></textarea>

<script>
(function () {
  function v(id) { return document.getElementById(id).value; }
  function buildZpl() {
    return '^XA\n\n^LS2\n^SZ2\n^PW816\n^PON\n^PR14,14\n^PMN\n^MNY\n^LS0\n^MTD\n^MD30\n\n' +
      '\n^FX String that says SKU\n^FS\n^FO25,340\n^ARI,12,70\n^FDSKU\n' +
      '\n^FX String that says SKU\n^FS\n^FO10,10\n^ARB,150,180\n^FD' + v('f-sku') + '\n' +
      '\n^FX String that says the grade\n^FS\n^FO172,0\n^ARN,175,225\n^FD' + v('f-grade') + '\n' +
      '\n^FX String that says the battery percentage\n^FS\n^FO200,155\n^ARN,25,12\n^FD' + v('f-batt') + '\n' +
      '\n^FX The QR Code\n^FS\n^FO150,190\n^BQN,2,8\n^FDMA,' + v('f-g1') + '\n' +
      '\n^FX String that says G1 Number\n^FS\n^FO345,5\n^ARN,120,100\n^FD' + v('f-g1') + '\n' +
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
    return cmd;
  }
  function updatePreview() {
    var zpl = document.getElementById('le-zpl').value;
    fetch('https://api.labelary.com/v1/printers/8dpmm/labels/4x2/0/', {
      method: 'POST',
      headers: { 'Accept': 'image/png', 'Content-Type': 'application/x-www-form-urlencoded' },
      body: zpl
    }).then(function (r) {
      if (!r.ok) throw new Error('Labelary HTTP ' + r.status);
      return r.blob();
    }).then(function (b) {
      document.getElementById('le-preview').src = URL.createObjectURL(b);
      document.getElementById('le-msg').textContent = '';
    }).catch(function (e) {
      document.getElementById('le-msg').textContent = 'Preview failed: ' + e.message;
    });
  }
  function regen() {
    document.getElementById('le-zpl').value = buildZpl();
    updateCmd();
    updatePreview();
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