# 📊 Driver Installer Analytics

<div class="dia" id="dia-root">
  <div class="dia-toolbar">
    <span class="dia-status" id="dia-status">Loading analytics…</span>
    <button class="dia-refresh" id="dia-refresh" type="button" title="Refresh now">↻ Refresh</button>
  </div>

  <div class="dia-kpis" id="dia-kpis"></div>

  <div class="dia-grid">
    <section class="dia-card">
      <h3>Run results</h3>
      <div class="dia-donut-wrap"><div id="dia-donut"></div><div class="dia-legend" id="dia-legend"></div></div>
    </section>
    <section class="dia-card">
      <h3>Top manufacturers</h3>
      <div id="dia-mfr" class="dia-bars"></div>
    </section>
    <section class="dia-card dia-wide">
      <h3>Top models</h3>
      <div id="dia-models" class="dia-bars dia-bars-wide"></div>
    </section>
  </div>

  <section class="dia-card dia-tablecard">
    <div class="dia-tablehead">
      <h3>Recent runs</h3>
      <input type="search" id="dia-filter" placeholder="Filter by model, manufacturer, result…" />
    </div>
    <div class="dia-tablescroll">
      <table class="dia-table" id="dia-table">
        <thead>
          <tr>
            <th>When</th><th>Result</th><th>Manufacturer</th><th>Model</th>
            <th>Ver</th><th class="num">Missing</th><th class="num">Installed</th><th class="num">Duration</th>
          </tr>
        </thead>
        <tbody id="dia-tbody"></tbody>
      </table>
    </div>
  </section>

  <p class="dia-foot" id="dia-foot"></p>
</div>

<style>
.dia { --dia-ok:#2e9e5b; --dia-fail:#e5484d; --dia-cancel:#f5a623; --dia-other:#8b8b8b;
  --dia-card: var(--md-default-bg-color);
  --dia-line: color-mix(in srgb, var(--md-default-fg-color) 12%, transparent);
  --dia-muted: color-mix(in srgb, var(--md-default-fg-color) 60%, transparent);
  --dia-fill: var(--md-primary-fg-color);
  font-feature-settings:"tnum"; }
.dia * { box-sizing:border-box; }
.dia-toolbar { display:flex; align-items:center; gap:.75rem; margin:.25rem 0 1rem; }
.dia-status { color:var(--dia-muted); font-size:.85rem; }
.dia-refresh { margin-left:auto; border:1px solid var(--dia-line); background:transparent;
  color:var(--md-typeset-color); border-radius:.5rem; padding:.35rem .7rem; cursor:pointer; font-size:.85rem; }
.dia-refresh:hover { border-color:var(--dia-fill); color:var(--dia-fill); }

.dia-kpis { display:grid; grid-template-columns:repeat(auto-fit,minmax(150px,1fr)); gap:.75rem; margin-bottom:1.25rem; }
.dia-kpi { border:1px solid var(--dia-line); border-radius:.75rem; padding:.9rem 1rem; background:var(--dia-card); }
.dia-kpi .v { font-size:1.7rem; font-weight:700; line-height:1.1; }
.dia-kpi .l { color:var(--dia-muted); font-size:.78rem; margin-top:.25rem; text-transform:uppercase; letter-spacing:.03em; }

.dia-grid { display:grid; grid-template-columns:repeat(auto-fit,minmax(280px,1fr)); gap:1rem; margin-bottom:1rem; }
.dia-grid .dia-wide { grid-column:1 / -1; }
.dia-card { border:1px solid var(--dia-line); border-radius:.75rem; padding:1rem 1.1rem; background:var(--dia-card); }
.dia-card h3 { margin:.1rem 0 .9rem; font-size:.95rem; }

.dia-donut-wrap { display:flex; align-items:center; gap:1rem; flex-wrap:wrap; }
.dia-legend { display:flex; flex-direction:column; gap:.4rem; font-size:.85rem; }
.dia-legend .row { display:flex; align-items:center; gap:.5rem; }
.dia-legend .dot { width:.7rem; height:.7rem; border-radius:2px; display:inline-block; }

.dia-bars { display:flex; flex-direction:column; gap:.55rem; }
.dia-bar { display:grid; grid-template-columns:9rem 1fr auto; align-items:center; gap:.6rem; font-size:.85rem; }
.dia-bar .name { overflow:hidden; text-overflow:ellipsis; white-space:nowrap; color:var(--dia-muted); }
.dia-bar .track { display:block; background:color-mix(in srgb, var(--md-default-fg-color) 8%, transparent); border-radius:6px; height:.7rem; overflow:hidden; }
.dia-bar .fill { display:block; height:100%; background:var(--dia-fill); border-radius:6px; min-width:6px; }
.dia-bar .val { font-variant-numeric:tabular-nums; }
.dia-bars-wide .dia-bar { grid-template-columns:16rem 1fr auto; }

.dia-tablecard { padding-bottom:.5rem; }
.dia-tablehead { display:flex; align-items:center; gap:1rem; flex-wrap:wrap; margin-bottom:.6rem; }
.dia-tablehead h3 { margin:0; }
#dia-filter { margin-left:auto; border:1px solid var(--dia-line); background:transparent; color:var(--md-typeset-color);
  border-radius:.5rem; padding:.4rem .7rem; font-size:.85rem; min-width:min(320px,60vw); }
.dia-tablescroll { overflow-x:auto; }
.dia-table { width:100%; border-collapse:collapse; font-size:.83rem; }
.dia-table th, .dia-table td { text-align:left; padding:.45rem .6rem; border-bottom:1px solid var(--dia-line); white-space:nowrap; }
.dia-table th.num, .dia-table td.num { text-align:right; font-variant-numeric:tabular-nums; }
.dia-table tbody tr:hover { background:color-mix(in srgb, var(--md-default-fg-color) 5%, transparent); }
.dia-pill { display:inline-block; padding:.1rem .5rem; border-radius:999px; font-size:.72rem; font-weight:600; }
.dia-pill.ok { background:color-mix(in srgb,var(--dia-ok) 20%,transparent); color:var(--dia-ok); }
.dia-pill.fail { background:color-mix(in srgb,var(--dia-fail) 20%,transparent); color:var(--dia-fail); }
.dia-pill.cancel { background:color-mix(in srgb,var(--dia-cancel) 20%,transparent); color:var(--dia-cancel); }
.dia-pill.other { background:color-mix(in srgb,var(--dia-other) 22%,transparent); color:var(--dia-other); }

.dia-foot { color:var(--dia-muted); font-size:.78rem; margin-top:1rem; }
.dia-empty { color:var(--dia-muted); padding:1rem 0; }
</style>

<script>
(function () {
  // Same Apps Script deployment the installer POSTs analytics to; its doGet now
  // returns the rows as JSON. If the URL ever changes, update it here too.
  var WEBHOOK = "https://script.google.com/macros/s/AKfycbygEF0i6j_6rSstmfQ2sQPLn0KjkqxZwUwIRjyCsd911IP9kALucv2cImMFumGoUUs/exec";

  var root = document.getElementById("dia-root");
  if (!root) return;

  var $ = function (id) { return document.getElementById(id); };
  var num = function (v) { var n = Number(v); return isNaN(n) ? 0 : n; };
  var esc = function (s) { return String(s == null ? "" : s).replace(/[&<>"]/g, function (c) {
    return ({ "&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;" })[c]; }); };
  var ALL = [];

  function resultClass(r) {
    r = String(r || "").toLowerCase();
    if (r === "success") return "ok";
    if (r === "failure") return "fail";
    if (r === "cancelled" || r === "cancel") return "cancel";
    return "other";
  }
  function fmtDur(s) { s = num(s); if (!s) return "—"; if (s < 60) return s + "s";
    var m = Math.floor(s/60), r = s%60; return m + "m" + (r ? " " + r + "s" : ""); }
  function fmtWhen(t) { var d = new Date(t); if (isNaN(d)) return esc(t);
    return d.toLocaleDateString(undefined,{month:"short",day:"numeric"}) + " " +
           d.toLocaleTimeString(undefined,{hour:"2-digit",minute:"2-digit"}); }

  function kpi(v, l) { return '<div class="dia-kpi"><div class="v">' + v + '</div><div class="l">' + l + '</div></div>'; }

  function render(rows) {
    ALL = rows;
    var total = rows.length;
    if (!total) { $("dia-kpis").innerHTML = ""; $("dia-tbody").innerHTML =
        '<tr><td colspan="8" class="dia-empty">No runs recorded yet.</td></tr>'; return; }

    var succ = 0, resolved = 0, instTotal = 0, durs = [];
    var mfr = {}, models = {}, results = {};
    rows.forEach(function (r) {
      var res = String(r.result || "").toLowerCase();
      results[res] = (results[res] || 0) + 1;
      if (res === "success") succ++;
      var mb = r.missing_before, ma = r.missing_after;
      if (mb !== "" && ma !== "" && mb != null && ma != null) {
        var d = num(mb) - num(ma); if (d > 0) resolved += d;
      }
      instTotal += num(r.inf_count);
      if (num(r.duration_sec) > 0) { durs.push(num(r.duration_sec)); }
      var m = (r.manufacturer || "Unknown").trim() || "Unknown"; mfr[m] = (mfr[m]||0)+1;
      var mdl = String(r.model == null ? "" : r.model).trim();
      if (mdl) models[mdl] = (models[mdl]||0)+1;
    });

    var rate = total ? Math.round((succ/total)*100) : 0;
    // Median, not mean - old rows have drifted/garbage duration values that
    // would otherwise skew a mean into the tens of minutes.
    durs.sort(function (a, b) { return a - b; });
    var medDur = durs.length ? durs[Math.floor((durs.length - 1) / 2)] : 0;
    $("dia-kpis").innerHTML =
      kpi(total, "Total runs") +
      kpi(rate + "%", "Success rate") +
      kpi(Object.keys(models).length, "Unique models") +
      kpi(resolved.toLocaleString(), "Devices resolved") +
      kpi(instTotal.toLocaleString(), "Drivers installed") +
      kpi(fmtDur(medDur), "Median run time");

    renderDonut(results, total);
    renderBars("dia-mfr", mfr, 6);
    renderBars("dia-models", models, 10);
    applyFilter();
  }

  function renderDonut(results, total) {
    var order = [["success","ok","var(--dia-ok)","Success"],["failure","fail","var(--dia-fail)","Failure"],
      ["cancelled","cancel","var(--dia-cancel)","Cancelled"],["testmode","other","var(--dia-other)","Test mode"]];
    var known = {}; order.forEach(function(o){ known[o[0]]=true; });
    var otherN = 0; Object.keys(results).forEach(function(k){ if(!known[k]) otherN += results[k]; });
    var segs = []; order.forEach(function(o){ if (results[o[0]]) segs.push({n:results[o[0]],c:o[2],label:o[3]}); });
    if (otherN) segs.push({n:otherN,c:"var(--dia-other)",label:"Other"});

    var C = 2*Math.PI*42, off = 0, ring = "";
    segs.forEach(function (s) {
      var frac = s.n/total, len = frac*C;
      ring += '<circle cx="60" cy="60" r="42" fill="none" stroke="'+s.c+'" stroke-width="16" '+
        'stroke-dasharray="'+len+' '+(C-len)+'" stroke-dashoffset="'+(-off)+'" transform="rotate(-90 60 60)"></circle>';
      off += len;
    });
    $("dia-donut").innerHTML =
      '<svg viewBox="0 0 120 120" width="120" height="120" role="img" aria-label="Run results">'+ring+
      '<text x="60" y="58" text-anchor="middle" font-size="20" font-weight="700" fill="currentColor">'+total+'</text>'+
      '<text x="60" y="74" text-anchor="middle" font-size="9" fill="currentColor" opacity="0.6">RUNS</text></svg>';
    $("dia-legend").innerHTML = segs.map(function (s) {
      return '<div class="row"><span class="dot" style="background:'+s.c+'"></span>'+
        esc(s.label)+' — <strong>'+s.n+'</strong> ('+Math.round(s.n/total*100)+'%)</div>';
    }).join("");
  }

  function renderBars(elId, map, limit) {
    var arr = Object.keys(map).map(function (k) { return [k, map[k]]; })
      .sort(function (a, b) { return b[1]-a[1]; }).slice(0, limit);
    var max = arr.reduce(function (m, x) { return Math.max(m, x[1]); }, 0) || 1;
    $(elId).innerHTML = arr.map(function (x) {
      return '<div class="dia-bar"><span class="name" title="'+esc(x[0])+'">'+esc(x[0])+'</span>'+
        '<span class="track"><span class="fill" style="width:'+Math.max(3,(x[1]/max*100))+'%"></span></span>'+
        '<span class="val">'+x[1]+'</span></div>';
    }).join("") || '<div class="dia-empty">No data.</div>';
  }

  function applyFilter() {
    var q = ($("dia-filter").value || "").toLowerCase().trim();
    var rows = !q ? ALL : ALL.filter(function (r) {
      return (String(r.model||"")+" "+String(r.manufacturer||"")+" "+String(r.result||"")+" "+
              String(r.script_version||"")).toLowerCase().indexOf(q) !== -1;
    });
    var view = rows.slice(0, 200);
    $("dia-tbody").innerHTML = view.map(function (r) {
      var mb = r.missing_before, ma = r.missing_after;
      var miss = (mb===""||mb==null) ? "—" : (num(mb) + " → " + ((ma===""||ma==null)?"?":num(ma)));
      return "<tr>"+
        "<td>"+fmtWhen(r.timestamp)+"</td>"+
        '<td><span class="dia-pill '+resultClass(r.result)+'">'+esc(r.result||"?")+"</span></td>"+
        "<td>"+esc(r.manufacturer||"—")+"</td>"+
        "<td>"+esc(r.model||"—")+"</td>"+
        "<td>"+esc(r.script_version||"—")+"</td>"+
        '<td class="num">'+miss+"</td>"+
        '<td class="num">'+num(r.inf_count)+"</td>"+
        '<td class="num">'+fmtDur(r.duration_sec)+"</td>"+
      "</tr>";
    }).join("") || '<tr><td colspan="8" class="dia-empty">No matching runs.</td></tr>';
    $("dia-foot").textContent = "Showing " + view.length + " of " + ALL.length + " run(s)." +
      (rows.length > view.length ? " (table capped at 200 — use the filter)" : "");
  }

  function load() {
    $("dia-status").textContent = "Loading analytics…";
    fetch(WEBHOOK, { method: "GET" })
      .then(function (r) { return r.json(); })
      .then(function (d) {
        if (d && d.error) throw new Error(d.error);
        var rows = (d && d.rows) ? d.rows : [];
        render(rows);
        $("dia-status").textContent = "Live from Google Sheet · updated " + new Date().toLocaleTimeString();
      })
      .catch(function (err) {
        $("dia-status").textContent = "Could not load analytics: " + err.message +
          " — the Apps Script may need to be redeployed (see code.gs).";
      });
  }

  $("dia-refresh").addEventListener("click", load);
  $("dia-filter").addEventListener("input", applyFilter);
  load();
  setInterval(load, 60000); // auto-refresh every minute
})();
</script>
