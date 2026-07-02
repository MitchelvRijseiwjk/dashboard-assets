// === SETTINGS SUB-TABS ===
function swSetupTab(panel, el) {
  var panels = document.querySelectorAll('.setup-tab-panel');
  for (var i = 0; i < panels.length; i++) { panels[i].classList.remove('active'); }
  var target = document.getElementById('setupTab-' + panel);
  if (target) target.classList.add('active');
  var tabs = document.querySelectorAll('.setup-tab');
  for (var j = 0; j < tabs.length; j++) { tabs[j].classList.remove('active'); }
  if (el) el.classList.add('active');
}

// swSettings removed — Settings modal (daSettings.openSettings) replaces sub-tab switching

// === DATE FILTER (per-entity) ===
var activeFilterValue = {};
var activeFilterLabel = {};
var currentAnalysisEntity = '';
// Global window end + scope (period start is still per-entity via activeFilterValue).
var activeWindowTo = '';
var activeScope = 'active';

function resolveDate(val) {
  if (!val) return '';
  if (val === 'custom') return '';
  var now = new Date();
  function fmt(d) {
    var y = d.getFullYear();
    var m = ('0' + (d.getMonth() + 1)).slice(-2);
    var day = ('0' + d.getDate()).slice(-2);
    return y + '-' + m + '-' + day;
  }
  if (val === 'thisyear') { return now.getFullYear() + '-01-01'; }
  if (val === 'last6m') { now.setDate(1); now.setMonth(now.getMonth() - 6); return fmt(now); }
  if (val === 'last12m') { now.setDate(1); now.setMonth(now.getMonth() - 12); return fmt(now); }
  if (val === 'last24m') { now.setDate(1); now.setMonth(now.getMonth() - 24); return fmt(now); }
  if (val === 'last36m') { now.setDate(1); now.setMonth(now.getMonth() - 36); return fmt(now); }
  if (val === 'last48m') { now.setDate(1); now.setMonth(now.getMonth() - 48); return fmt(now); }
  // Enforce max 4 years back for custom dates
  if (val) {
    var maxBack = new Date();
    maxBack.setMonth(maxBack.getMonth() - 48);
    maxBack.setDate(1);
    var valDate = new Date(val);
    if (valDate < maxBack) return fmt(maxBack);
  }
  return val;
}

function getSelDateVal(sel) {
  if (!sel) return '';
  var v = sel.value;
  if (v === 'custom') {
    var di = sel.parentNode.querySelector('.custom-date-input');
    if (di && di.value) return di.value;
    return '';
  }
  return v;
}

function getSelectLabel(sel) {
  if (!sel || !sel.value) return 'All data';
  if (sel.value === 'custom') {
    var di = sel.parentNode.querySelector('.custom-date-input');
    if (di && di.value) return 'Since ' + di.value;
    return 'Custom date';
  }
  return sel.options[sel.selectedIndex].text;
}

// === CUSTOM DATE PICKER ===
function handleDateSelect(sel) {
  var wrapper = sel.parentNode;
  var dateInput = wrapper.querySelector('.custom-date-input');
  if (sel.value === 'custom') {
    if (!dateInput) {
      dateInput = document.createElement('input');
      dateInput.type = 'date';
      dateInput.className = 'custom-date-input';
      if (envStartDate) dateInput.min = envStartDate;
      dateInput.max = envEndDate;
      sel.insertAdjacentElement('afterend', dateInput);
    }
    dateInput.style.display = '';
    dateInput.focus();
  } else {
    if (dateInput) dateInput.style.display = 'none';
  }
}

function validateCustomDates() {
  var selects = document.querySelectorAll('select');
  for (var i = 0; i < selects.length; i++) {
    if (selects[i].value === 'custom') {
      var di = selects[i].parentNode.querySelector('.custom-date-input');
      if (!di || !di.value) {
        di.style.borderColor = '#c62828';
        di.focus();
        return false;
      }
    }
  }
  return true;
}

document.addEventListener('change', function(e) {
  var t = e.target;
  if (t.tagName === 'SELECT' && (t.id.indexOf('dateFilter') > -1 || t.id.indexOf('DateFilter') > -1 || t.id.indexOf('aaDate') > -1 || t.id.indexOf('setupDate') > -1)) {
    handleDateSelect(t);
  }
  if (t.classList.contains('custom-date-input') && t.value) {
    t.style.borderColor = '';
  }
});

/* ============================================================
   Date range + scope control (front-end).
   Injected into the existing date cards in place of the old
   single dropdown. Behaviour-preserving: the default rolling
   12 months resolves to the same from-date as the old
   "last 12 months" option. Emits dateFrom/dateTo/scope plus a
   dateFilter alias so not-yet-migrated fetch scripts keep working.
   ============================================================ */
var DP_GREEN = '#06423e';

function dpSvg(name, sz) {
  var s = sz || 18;
  var o = '<svg width="' + s + '" height="' + s + '" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true" style="display:block">';
  var p = '';
  if (name === 'calendar') p = '<rect x="3" y="4.5" width="18" height="16" rx="2"/><path d="M3 9h18M8 3v3M16 3v3"/>';
  else if (name === 'arrow') p = '<path d="M5 12h14M13 6l6 6-6 6"/>';
  else if (name === 'activity') p = '<path d="M3 12h4l3 7 4-14 3 7h4"/>';
  else if (name === 'seedling') p = '<path d="M12 21v-8"/><path d="M12 13C12 9 9 7 5 7c0 4 3 6 7 6z"/><path d="M12 13c0-3 3-5 7-5 0 3-3 5-7 5z"/>';
  else if (name === 'database') p = '<ellipse cx="12" cy="5.5" rx="8" ry="3"/><path d="M4 5.5v6c0 1.66 3.58 3 8 3s8-1.34 8-3v-6"/><path d="M4 11.5v6c0 1.66 3.58 3 8 3s8-1.34 8-3v-6"/>';
  else if (name === 'info') p = '<circle cx="12" cy="12" r="9"/><path d="M12 11v5M12 8h.01"/>';
  else if (name === 'clock') p = '<circle cx="12" cy="12" r="9"/><path d="M12 7v5l3 2"/>';
  else if (name === 'check') p = '<path d="M5 12l4 4L19 7"/>';
  else if (name === 'chevron') p = '<path d="M6 9l6 6 6-6"/>';
  return o + p + '</svg>';
}

var DP_SCOPES = [
  { id: 'active', icon: 'activity', title: 'Active base', desc: 'Created or active in period', rec: true },
  { id: 'created', icon: 'seedling', title: 'Newly created', desc: 'Only records added in period', rec: false },
  { id: 'all', icon: 'database', title: 'Full database', desc: 'Everything, period ignored', rec: false }
];

var DP_COV = {
  active: { label: 'Active base', rows: [['Company', 'created or active'], ['Contact', 'created or active'], ['Sale', 'open or active'], ['Project', 'active'], ['Requests', 'open or updated'], ['Activities', 'logged in period']] },
  created: { label: 'Newly created', rows: [['Company', 'created in period'], ['Contact', 'created in period'], ['Sale', 'created in period'], ['Project', 'created in period'], ['Requests', 'created in period'], ['Activities', 'logged in period']] },
  all: { label: 'Full database', rows: [['Company', 'entire database'], ['Contact', 'entire database'], ['Sale', 'entire database'], ['Project', 'entire database'], ['Requests', 'entire database'], ['Activities', 'logged in period']] }
};

function dpControlHtml(P) {
  var lbl = 'color:var(--so-charcoal);opacity:.75;font-size:14px;';
  var inp = 'padding:9px 11px;font-size:14px;font-family:inherit;';
  var h = '';
  h += '<div class="dp-wrap" style="display:flex;flex-direction:column">';

  // ---- Card 1: Period ----
  h += '<div class="da2-card">';
  h += '<div class="da2-head"><span class="da2-ix">1</span><span class="da2-title">' + dpSvg('calendar', 17) + 'Period</span></div>';
  h += '<div class="da2-sub">Analyze data from a specific window or all time.</div>';
  h += '<div class="dp-seg" data-dp="' + P + '" data-mode-sel="rolling">';
  h += '<button type="button" class="dp-seg-btn" data-mode="rolling">Rolling</button>';
  h += '<button type="button" class="dp-seg-btn" data-mode="custom">Custom range</button>';
  h += '</div>';
  h += '<div class="dp-rolling" data-dp-roll="' + P + '" style="display:flex;align-items:center;gap:11px;flex-wrap:wrap;margin-top:14px">';
  h += '<span style="' + lbl + '">Last</span>';
  h += '<input type="number" id="' + P + 'RollN" value="12" min="1" max="120" style="width:72px;' + inp + '">';
  h += '<select id="' + P + 'RollUnit" style="width:132px;' + inp + '"><option value="month" selected>months</option><option value="year">years</option></select>';
  h += '<span style="color:var(--so-text-muted);display:inline-flex">' + dpSvg('arrow', 16) + '</span>';
  h += '<span id="' + P + 'Preview" style="color:var(--so-green-light);font-size:15px;font-weight:700"></span>';
  h += '</div>';
  h += '<div class="dp-custom" data-dp-cust="' + P + '" style="display:none;align-items:center;gap:11px;flex-wrap:wrap;margin-top:14px">';
  h += '<span style="' + lbl + '">From</span>';
  h += '<input type="date" id="' + P + 'DateFrom" style="width:162px;' + inp + '">';
  h += '<span style="' + lbl + '">to</span>';
  h += '<input type="date" id="' + P + 'DateTo" style="width:162px;' + inp + '">';
  h += '<span id="' + P + 'Span" style="color:var(--so-text-muted);font-size:13px"></span>';
  h += '</div>';
  h += '</div>';

  // ---- Card 2: Scope ----
  h += '<div class="da2-card">';
  h += '<div class="da2-head"><span class="da2-ix">2</span><span class="da2-title">' + dpSvg('database', 17) + 'Scope<span title="Scope decides which records count toward the quality and integrity scores. The period always bounds the activity and momentum metrics, whatever the scope." style="color:var(--so-text-muted);cursor:help;display:inline-flex">' + dpSvg('info', 14) + '</span></span></div>';
  h += '<div class="da2-sub">Which records the period selects for the quality and integrity scores.</div>';
  h += '<div class="dp-scope" data-dp="' + P + '" data-scope-sel="active" style="display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:12px;margin-top:16px">';
  for (var i = 0; i < DP_SCOPES.length; i++) {
    var sc = DP_SCOPES[i];
    h += '<div class="dp-scope-btn" data-scope="' + sc.id + '" style="position:relative;background:var(--so-white);border:1px solid var(--so-border);padding:13px 15px 12px;cursor:pointer;min-height:70px">';
    if (sc.rec) h += '<span style="position:absolute;top:-10px;left:14px;font-size:10px;font-weight:700;letter-spacing:.03em;color:var(--so-white);background:var(--so-green);border-radius:999px;padding:3px 10px">Recommended</span>';
    h += '<span class="dp-card-chk" style="position:absolute;top:13px;right:13px;color:var(--so-green-light);display:none">' + dpSvg('check', 16) + '</span>';
    h += '<div style="display:flex;align-items:center;gap:8px"><span style="color:var(--so-green-light);display:inline-flex">' + dpSvg(sc.icon, 18) + '</span><span style="font-size:15px;font-weight:700;color:var(--so-charcoal)">' + sc.title + '</span></div>';
    h += '<div style="font-size:12.5px;color:var(--so-text-muted);margin-top:7px;line-height:1.45">' + sc.desc + '</div>';
    h += '</div>';
  }
  h += '</div>';

  h += dpCountBarHtml(P);

  h += '<button type="button" class="dp-cov-toggle" data-dp-cov="' + P + '" style="display:inline-flex;align-items:center;gap:8px;margin-top:16px;background:none;border:none;color:var(--so-green-light);font-weight:700;font-size:14px;font-family:inherit;cursor:pointer;padding:4px 0">';
  h += '<span class="dp-cov-cx" style="display:inline-flex;transition:transform .2s">' + dpSvg('chevron', 16) + '</span>What gets included';
  h += '</button>';
  h += '<div class="dp-cov-wrap" data-dp-cov-wrap="' + P + '" style="display:none">';
  h += '<div style="display:flex;align-items:center;gap:7px;margin:13px 0 13px">';
  h += '<span title="For each entity, this is which records the scores count under the chosen scope." style="color:var(--so-text-muted);cursor:help;display:inline-flex">' + dpSvg('info', 14) + '</span>';
  h += '<span style="font-size:13px;color:var(--so-charcoal)">With <span id="' + P + 'CovScope" style="font-weight:700;color:var(--so-green-light)">Active base</span>, the scores cover</span>';
  h += '</div>';
  h += '<div id="' + P + 'CovGrid" style="display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:10px"></div>';
  h += '<div style="display:flex;align-items:center;gap:8px;margin-top:15px"><span style="color:var(--so-text-muted);display:inline-flex">' + dpSvg('clock', 15) + '</span><span style="font-size:12px;color:var(--so-text-muted)">Adoption and momentum always measure activity inside the period, whatever the scope.</span></div>';
  h += '</div>';
  h += '</div>';

  h += '</div>';
  return h;
}

function dpPaintGroup(container) {
  if (!container) return;
  var isScope = container.className.indexOf('dp-scope') > -1;
  if (isScope) {
    var sel = container.getAttribute('data-scope-sel');
    var cards = container.getElementsByClassName('dp-scope-btn');
    for (var i = 0; i < cards.length; i++) {
      var on = cards[i].getAttribute('data-scope') === sel;
      cards[i].style.border = on ? '1.5px solid var(--so-green-light)' : '1px solid var(--so-border)';
      cards[i].style.background = on ? '#eaf2ef' : 'var(--so-white)';
      cards[i].style.boxShadow = on ? '0 0 0 3px rgba(15,92,87,.08)' : 'none';
      cards[i].style.padding = on ? '12.5px 14.5px 11.5px' : '13px 15px 12px';
      var chk = cards[i].getElementsByClassName('dp-card-chk')[0];
      if (chk) chk.style.display = on ? 'inline-flex' : 'none';
    }
  } else {
    var selm = container.getAttribute('data-mode-sel');
    var btns = container.getElementsByTagName('button');
    for (var j = 0; j < btns.length; j++) {
      var onm = btns[j].getAttribute('data-mode') === selm;
      btns[j].style.background = onm ? 'var(--so-green)' : 'transparent';
      btns[j].style.color = onm ? 'var(--so-white)' : 'var(--so-charcoal)';
    }
  }
}

function dpShowMode(P) {
  var seg = document.querySelector('.dp-seg[data-dp="' + P + '"]');
  var mode = seg ? (seg.getAttribute('data-mode-sel') || 'rolling') : 'rolling';
  var roll = document.querySelector('.dp-rolling[data-dp-roll="' + P + '"]');
  var cust = document.querySelector('.dp-custom[data-dp-cust="' + P + '"]');
  if (roll) roll.style.display = (mode === 'rolling') ? 'flex' : 'none';
  if (cust) cust.style.display = (mode === 'custom') ? 'flex' : 'none';
}

function dpUpdateCov(P) {
  var key = dpScope(P);
  var c = DP_COV[key] || DP_COV.active;
  var ls = document.getElementById(P + 'CovScope');
  if (ls) ls.textContent = c.label;
  var g = document.getElementById(P + 'CovGrid');
  if (!g) return;
  var h = '';
  for (var i = 0; i < c.rows.length; i++) {
    h += '<div style="display:flex;align-items:center;justify-content:space-between;background:var(--so-dune);border-radius:7px;padding:8px 11px"><span style="font-size:13px;color:var(--so-charcoal)">' + c.rows[i][0] + '</span><span style="font-size:12px;color:var(--so-charcoal);background:var(--so-dune-dark1);border-radius:6px;padding:3px 9px;white-space:nowrap">' + c.rows[i][1] + '</span></div>';
  }
  g.innerHTML = h;
}

function dpUpdatePreview(P) {
  var el = document.getElementById(P + 'Preview');
  if (!el) return;
  var r = dpResolve(P);
  if (r.mode === 'rolling') {
    var M = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
    var d = new Date(r.from + 'T00:00:00');
    el.textContent = d.getDate() + ' ' + M[d.getMonth()] + ' ' + d.getFullYear() + ' to today';
  } else if (r.mode === 'custom' && r.from) {
    el.textContent = r.from + (r.to ? (' \u2192 ' + r.to) : ' \u2192 \u2026');
  } else {
    el.textContent = '';
  }
}

function dpUpdateSpan(P) {
  var el = document.getElementById(P + 'Span');
  if (!el) return;
  var f = document.getElementById(P + 'DateFrom');
  var t = document.getElementById(P + 'DateTo');
  if (f && t && f.value && t.value) {
    var days = Math.round((new Date(t.value) - new Date(f.value)) / 86400000);
    el.textContent = (days >= 0) ? '(' + days + ' days)' : '';
  } else {
    el.textContent = '';
  }
}

// --- Live "companies in scope" count under the scope cards ---
function dpNum(n) { return (n || 0).toString().replace(/\B(?=(\d{3})+(?!\d))/g, '.'); }

function dpCountBarHtml(P) {
  var h = '';
  h += '<div id="' + P + 'ScopeCount" style="margin-top:13px;display:flex;align-items:center;gap:12px;background:var(--so-dune);border:1px solid var(--so-border);border-radius:var(--radius);padding:12px 14px">';
  h += '<span style="color:var(--so-green);display:inline-flex">' + dpSvg('database', 18) + '</span>';
  h += '<div style="flex:1;min-width:0">';
  h += '<div style="display:flex;justify-content:space-between;align-items:baseline;gap:10px;margin-bottom:7px">';
  h += '<span style="font-size:13px;color:var(--so-charcoal)">Companies in this scan</span>';
  h += '<span id="' + P + 'ScopeCountNum" style="font-size:13px;color:var(--so-text-muted)">Calculating\u2026</span>';
  h += '</div>';
  h += '<div style="height:6px;background:var(--so-dune-dark2);border-radius:999px;overflow:hidden">';
  h += '<div id="' + P + 'ScopeCountBar" style="height:100%;width:0;background:var(--so-green);border-radius:999px;transition:width .35s ease"></div>';
  h += '</div>';
  h += '</div>';
  h += '</div>';
  return h;
}

function dpScopeCountParams(P) {
  var r = dpResolve(P);
  var sc = dpScope(P);
  var amp = String.fromCharCode(38);
  var q = '';
  if (r.from) q += amp + 'dateFrom=' + r.from + amp + 'dateFilter=' + r.from;
  if (r.to) q += amp + 'dateTo=' + r.to;
  q += amp + 'scope=' + sc;
  return q;
}

var _dpCountSeq = {};
var _dpCountTimer = {};

function dpUpdateScopeCount(P) {
  var numEl = document.getElementById(P + 'ScopeCountNum');
  var barEl = document.getElementById(P + 'ScopeCountBar');
  if (!numEl) return;
  if (typeof scopeCountUrl === 'undefined' || !scopeCountUrl) { numEl.textContent = ''; return; }
  numEl.textContent = 'Calculating\u2026';
  if (barEl) barEl.style.opacity = '.5';
  var seq = (_dpCountSeq[P] || 0) + 1;
  _dpCountSeq[P] = seq;
  ajax(scopeCountUrl + dpScopeCountParams(P), function(o) {
    if (seq !== _dpCountSeq[P]) return; // a newer request for this control superseded this one
    if (!o || typeof o.scoped === 'undefined') { numEl.textContent = ''; if (barEl) barEl.style.width = '0'; if (typeof dpUpdateBar === 'function') dpUpdateBar(P); return; }
    var s = o.scoped | 0;
    var t = o.total | 0;
    numEl.innerHTML = '<strong style="color:var(--so-green);font-weight:700">' + dpNum(s) + '</strong> of ' + dpNum(t);
    if (barEl) { barEl.style.opacity = '1'; barEl.style.width = (t > 0 ? Math.max(2, Math.round(s / t * 100)) : 0) + '%'; }
    if (typeof dpUpdateBar === 'function') dpUpdateBar(P);
  });
}

function dpScheduleCount(P) {
  if (_dpCountTimer[P]) clearTimeout(_dpCountTimer[P]);
  _dpCountTimer[P] = setTimeout(function() { dpUpdateScopeCount(P); }, 280);
}

// ---- Sticky bottom action bar (setup intake only) ----
function dpEntCount() {
  var card = document.getElementById('setupEntitiesCard');
  if (!card) return 0;
  var boxes = card.querySelectorAll('input[type="checkbox"]');
  var n = 0;
  for (var i = 0; i < boxes.length; i++) { if (boxes[i].checked) n++; }
  return n;
}

function dpBuildActionBar(P) {
  if (P !== 'setup') return;
  var inner = document.querySelector('#setupScreen .setup-inner');
  if (!inner || document.getElementById('setupActionBar')) return;
  var play = '<svg width="16" height="16" viewBox="0 0 24 24" fill="currentColor" aria-hidden="true"><path d="M8 5v14l11-7z"/></svg>';
  var bar = document.createElement('div');
  bar.id = 'setupActionBar';
  bar.innerHTML =
    '<div class="ab-inner">' +
      '<span class="ab-chip" id="abPeriod">' + dpSvg('calendar', 15) + '<span>Last 12 months</span></span>' +
      '<span class="ab-chip" id="abScope">' + dpSvg('database', 15) + '<span>Active base</span></span>' +
      '<span class="ab-chip" id="abComp">' + dpSvg('activity', 15) + '<b>\u2014</b><span>companies</span></span>' +
      '<span class="ab-meta" id="abMeta"></span>' +
      '<button type="button" class="ab-start" id="abStart">' + play + 'Start Analysis</button>' +
    '</div>';
  inner.appendChild(bar);
  var btn = document.getElementById('abStart');
  if (btn) btn.addEventListener('click', function() { if (!btn.disabled && typeof launchDashboard === 'function') launchDashboard(); });
  dpUpdateBar('setup');
}

function dpUpdateBar(P) {
  if (P !== 'setup' || !document.getElementById('setupActionBar')) return;
  var mode = dpMode(P);
  // Period chip label
  var periodLabel;
  if (mode === 'custom') {
    periodLabel = 'Custom range';
  } else {
    var nEl = document.getElementById(P + 'RollN');
    var uEl = document.getElementById(P + 'RollUnit');
    var n = nEl ? parseInt(nEl.value, 10) : NaN;
    var u = uEl ? uEl.value : 'month';
    if (isNaN(n) || n < 1) periodLabel = 'Rolling';
    else periodLabel = 'Last ' + n + ' ' + (u === 'year' ? (n === 1 ? 'year' : 'years') : (n === 1 ? 'month' : 'months'));
  }
  var pe = document.querySelector('#abPeriod span'); if (pe) pe.textContent = periodLabel;
  // Scope chip
  var scId = dpScope(P), scTitle = 'Active base';
  for (var i = 0; i < DP_SCOPES.length; i++) { if (DP_SCOPES[i].id === scId) scTitle = DP_SCOPES[i].title; }
  var se = document.querySelector('#abScope span'); if (se) se.textContent = scTitle;
  // Companies chip (read the resolved scope count)
  var rawEl = document.getElementById(P + 'ScopeCountNum');
  var raw = rawEl ? (rawEl.textContent || '') : '';
  var digits = raw.split(' of ')[0].replace(/[^\d]/g, '');
  var companies = digits ? parseInt(digits, 10) : NaN;
  var ce = document.querySelector('#abComp b'); if (ce) ce.textContent = isNaN(companies) ? '\u2014' : dpNum(companies);
  // Ready state + run meta
  var ent = dpEntCount();
  var ready = true, reason = '';
  if (ent < 1) { ready = false; reason = 'Select at least one entity'; }
  else if (mode === 'custom') {
    var f = (document.getElementById(P + 'DateFrom') || {}).value || '';
    var t = (document.getElementById(P + 'DateTo') || {}).value || '';
    if (!f || !t || f > t) { ready = false; reason = 'Set a valid date range'; }
  } else {
    var rn = parseInt((document.getElementById(P + 'RollN') || {}).value, 10);
    if (isNaN(rn) || rn < 1) { ready = false; reason = 'Enter a length of at least 1'; }
  }
  var meta = document.getElementById('abMeta');
  if (meta) {
    if (!ready) { meta.className = 'ab-meta warn'; meta.textContent = reason; }
    else {
      meta.className = 'ab-meta';
      var est = isNaN(companies) ? null : Math.max(1, Math.round(companies / 1600));
      meta.innerHTML = (est ? 'est. <b>~' + est + ' min</b> \u00b7 ' : '') + ent + ' ' + (ent === 1 ? 'entity' : 'entities');
    }
  }
  var btn = document.getElementById('abStart');
  if (btn) btn.disabled = !ready;
}

document.addEventListener('click', function(e) {
  var b = e.target;
  while (b && b !== document && (!b.className || typeof b.className !== 'string' || (b.className.indexOf('dp-seg-btn') < 0 && b.className.indexOf('dp-scope-btn') < 0 && b.className.indexOf('dp-cov-toggle') < 0))) {
    b = b.parentNode;
  }
  if (!b || b === document) return;
  if (b.className.indexOf('dp-seg-btn') > -1) {
    var seg = b.parentNode;
    seg.setAttribute('data-mode-sel', b.getAttribute('data-mode'));
    dpPaintGroup(seg);
    dpShowMode(seg.getAttribute('data-dp'));
    dpUpdatePreview(seg.getAttribute('data-dp'));
    dpScheduleCount(seg.getAttribute('data-dp'));
  } else if (b.className.indexOf('dp-scope-btn') > -1) {
    var sc = b.parentNode;
    sc.setAttribute('data-scope-sel', b.getAttribute('data-scope'));
    dpPaintGroup(sc);
    // Keep the global activeScope in sync (B-6 fix)
    activeScope = b.getAttribute('data-scope');
    dpUpdateCov(sc.getAttribute('data-dp'));
    dpScheduleCount(sc.getAttribute('data-dp'));
  } else if (b.className.indexOf('dp-cov-toggle') > -1) {
    var cP = b.getAttribute('data-dp-cov');
    var wrap = document.querySelector('.dp-cov-wrap[data-dp-cov-wrap="' + cP + '"]');
    var cx = b.querySelector('.dp-cov-cx');
    if (wrap) {
      var open = wrap.style.display !== 'none';
      wrap.style.display = open ? 'none' : '';
      if (cx) cx.style.transform = open ? 'rotate(0deg)' : 'rotate(180deg)';
    }
  }
});

document.addEventListener('input', function(e) {
  var t = e.target;
  if (!t || !t.id) return;
  if (/RollN$/.test(t.id)) { var P = t.id.replace(/RollN$/, ''); dpUpdatePreview(P); dpScheduleCount(P); }
  else if (/DateFrom$|DateTo$/.test(t.id)) { var P2 = t.id.replace(/DateFrom$|DateTo$/, ''); dpUpdateSpan(P2); dpScheduleCount(P2); }
});

document.addEventListener('change', function(e) {
  var t = e.target;
  if (!t || !t.id) return;
  if (/RollUnit$/.test(t.id)) { var P = t.id.replace(/RollUnit$/, ''); dpUpdatePreview(P); dpScheduleCount(P); }
  else if (/DateFrom$|DateTo$/.test(t.id)) { var P2 = t.id.replace(/DateFrom$|DateTo$/, ''); dpUpdateSpan(P2); dpScheduleCount(P2); }
});

function dpResolve(P) {
  function fmt(d) { var y = d.getFullYear(); var m = ('0' + (d.getMonth() + 1)).slice(-2); var dd = ('0' + d.getDate()).slice(-2); return y + '-' + m + '-' + dd; }
  var seg = document.querySelector('.dp-seg[data-dp="' + P + '"]');
  var mode = seg ? (seg.getAttribute('data-mode-sel') || 'rolling') : 'rolling';
  if (mode === 'custom') {
    var f = document.getElementById(P + 'DateFrom');
    var t = document.getElementById(P + 'DateTo');
    var fv = f ? f.value : '';
    var tv = t ? t.value : '';
    if (fv) { var mb = new Date(); mb.setMonth(mb.getMonth() - 48); mb.setDate(1); if (new Date(fv) < mb) fv = fmt(mb); }
    return { from: fv, to: tv, mode: 'custom' };
  }
  var nEl = document.getElementById(P + 'RollN');
  var uEl = document.getElementById(P + 'RollUnit');
  var n = nEl ? parseInt(nEl.value, 10) : 12;
  if (!n || n < 1) n = 12;
  var months = (uEl && uEl.value === 'year') ? n * 12 : n;
  var d = new Date(); d.setDate(1); d.setMonth(d.getMonth() - months);
  return { from: fmt(d), to: '', mode: 'rolling' };
}
function dpMode(P) { var seg = document.querySelector('.dp-seg[data-dp="' + P + '"]'); return seg ? (seg.getAttribute('data-mode-sel') || 'rolling') : 'rolling'; }
function dpScope(P) { var sc = document.querySelector('.dp-scope[data-dp="' + P + '"]'); return sc ? (sc.getAttribute('data-scope-sel') || 'active') : 'active'; }

function dpMirror(src, dst) {
  var sSeg = document.querySelector('.dp-seg[data-dp="' + src + '"]');
  var dSeg = document.querySelector('.dp-seg[data-dp="' + dst + '"]');
  if (sSeg && dSeg) { dSeg.setAttribute('data-mode-sel', sSeg.getAttribute('data-mode-sel') || 'rolling'); dpPaintGroup(dSeg); }
  var sScope = document.querySelector('.dp-scope[data-dp="' + src + '"]');
  var dScope = document.querySelector('.dp-scope[data-dp="' + dst + '"]');
  if (sScope && dScope) { dScope.setAttribute('data-scope-sel', sScope.getAttribute('data-scope-sel') || 'active'); dpPaintGroup(dScope); }
  var pairs = ['RollN', 'RollUnit', 'DateFrom', 'DateTo'];
  for (var i = 0; i < pairs.length; i++) {
    var a = document.getElementById(src + pairs[i]);
    var c = document.getElementById(dst + pairs[i]);
    if (a && c) c.value = a.value;
  }
  dpShowMode(dst);
  dpUpdateCov(dst);
  dpUpdatePreview(dst);
  dpUpdateScopeCount(dst);
}

function dpBuild(P, oldSelectId) {
  var old = document.getElementById(oldSelectId);
  if (!old) return;
  if (old.getAttribute('data-dp-replaced') === '1') return;
  var wrap = document.createElement('div');
  wrap.innerHTML = dpControlHtml(P);
  old.parentNode.insertBefore(wrap.firstChild, old);
  old.style.display = 'none';
  old.setAttribute('data-dp-replaced', '1');
  dpPaintGroup(document.querySelector('.dp-seg[data-dp="' + P + '"]'));
  dpPaintGroup(document.querySelector('.dp-scope[data-dp="' + P + '"]'));
  dpShowMode(P);
  dpUpdateCov(P);
  dpUpdatePreview(P);
  dpUpdateScopeCount(P);
}

function dpEntitiesSummary(cardId) {
  var card = document.getElementById(cardId);
  if (!card) return '';
  var boxes = card.querySelectorAll('input[type="checkbox"]');
  var checked = 0;
  for (var i = 0; i < boxes.length; i++) { if (boxes[i].checked) checked++; }
  var sels = card.querySelectorAll('.setup-entity-date, .aa-entity-date');
  var overrides = 0;
  for (var j = 0; j < sels.length; j++) { var v = sels[j].value; if (v && v !== '') overrides++; }
  var s = checked + ' included';
  if (overrides > 0) s += ' \u00b7 ' + overrides + (overrides === 1 ? ' custom window' : ' custom windows');
  else s += ' \u00b7 global window';
  return s;
}

function dpCollapseEntities(cardId) {
  var card = document.getElementById(cardId);
  if (!card) return;
  if (card.getAttribute('data-dp-collapsed') === '1') return;
  var h3 = card.querySelector('h3');
  var desc = card.querySelector('.setup-desc, .aa-desc');
  var grid = card.querySelector('.setup-entity-grid, .aa-entity-grid');
  if (!h3 || !grid) return;
  card.setAttribute('data-dp-collapsed', '1');

  h3.style.cursor = 'pointer';
  h3.style.display = 'flex';
  h3.style.alignItems = 'center';
  h3.style.justifyContent = 'space-between';
  h3.style.margin = '0';

  var left = document.createElement('span');
  left.style.cssText = 'display:flex;align-items:center;gap:9px';
  while (h3.firstChild) { left.appendChild(h3.firstChild); }
  var sum = document.createElement('span');
  sum.className = 'dp-ent-sum';
  sum.style.cssText = 'font-size:.78rem;font-weight:500;color:var(--so-text-muted)';
  sum.textContent = dpEntitiesSummary(cardId);
  left.appendChild(sum);

  var chev = document.createElement('span');
  chev.style.cssText = 'color:var(--so-text-muted);display:inline-flex;transition:transform .15s';
  chev.innerHTML = dpSvg('chevron', 18);

  h3.appendChild(left);
  h3.appendChild(chev);

  if (desc) desc.style.display = 'none';
  grid.style.display = 'none';

  h3.addEventListener('click', function() {
    var open = grid.style.display !== 'none';
    grid.style.display = open ? 'none' : '';
    if (desc) desc.style.display = open ? 'none' : '';
    chev.style.transform = open ? 'rotate(0deg)' : 'rotate(180deg)';
  });

  card.addEventListener('change', function() {
    sum.textContent = dpEntitiesSummary(cardId);
  });
}

function _dpHideBoot() {
  var b = document.getElementById('hgxBoot');
  if (b) b.style.display = 'none';
}

function dpBootReveal() {
  if (window._dpBooted) return;
  window._dpBooted = true;
  if (!window.HugoZoom || typeof HugoZoom.reveal !== 'function') {
    _dpHideBoot();
    return;
  }
  // Hold the loading state briefly so Hugo is seen, then reveal the setup screen.
  setTimeout(function() {
    _dpHideBoot();
    HugoZoom.reveal();
  }, 780);
}

function dpInit() {
  dpBuild('setup', 'setupDateFilter');
  dpBuild('aa', 'aaDateFilter');

  dpCollapseEntities('setupEntitiesCard');
  dpCollapseEntities('aaEntitiesCard');
  dpBuildActionBar('setup');
  // Keep the action bar summary and ready state in sync with any intake change.
  var sc = document.getElementById('setupScreen');
  if (sc) {
    var refresh = function() { setTimeout(function() { dpUpdateBar('setup'); }, 0); };
    sc.addEventListener('click', refresh);
    sc.addEventListener('input', refresh);
    sc.addEventListener('change', refresh);
  }
  dpBootReveal();
}
if (document.readyState === 'loading') { document.addEventListener('DOMContentLoaded', dpInit); } else { dpInit(); }

function captureEntityFilter(key) {
  var sel = document.getElementById('dateFilter_' + key);
  activeFilterValue[key] = resolveDate(getSelDateVal(sel));
  activeFilterLabel[key] = sel ? getSelectLabel(sel) : 'All data';
  currentAnalysisEntity = key;
}

function getDateFilterParam() {
  var amp = String.fromCharCode(38);
  var df = activeFilterValue[currentAnalysisEntity] || '';
  var dt = activeWindowTo || '';
  var sc = activeScope || 'active';
  var p = '';
  if (df) p += amp + 'dateFrom=' + df + amp + 'dateFilter=' + df;
  if (dt) p += amp + 'dateTo=' + dt;
  p += amp + 'scope=' + sc;
  return p;
}

// Field value-exclusions for an entity, as "field=idx,idx;field=idx,idx"
function getExclParam(entityKey) {
  var cfg = (typeof daSettings !== 'undefined' && daSettings.getConfig) ? daSettings.getConfig(entityKey) : null;
  if (!cfg || !cfg.fieldExclusions) return '';
  var pairs = [];
  for (var f in cfg.fieldExclusions) { var ids = cfg.fieldExclusions[f]; if (ids && ids.length > 0) pairs.push(f + '=' + ids.join(',')); }
  if (pairs.length === 0) return '';
  return String.fromCharCode(38) + 'excl=' + encodeURIComponent(pairs.join(';'));
}

// Active rule rows for an entity, flattened to "ruleId~row~condField~condOp~condVal~reqField~reqOp"
function _ruleRows(entityKey) {
  var rs = (typeof daSettings !== 'undefined' && daSettings.getRules) ? daSettings.getRules(entityKey) : [];
  var rows = [];
  for (var i = 0; i < rs.length; i++) {
    var r = rs[i]; if (r.active === false) continue;
    for (var j = 0; j < r.rows.length; j++) {
      var w = r.rows[j];
      rows.push([r.ruleId, j, w.condField, w.condOp, w.condVal, w.reqField, w.reqOp, w.owner || ''].join('~'));
    }
  }
  return rows;
}
function getRulesParam(entityKey) { var rows = _ruleRows(entityKey); return rows.length ? String.fromCharCode(38) + 'rules=' + encodeURIComponent(rows.join(';')) : ''; }
function aaHasRules(entityKey) { return _ruleRows(entityKey).length > 0; }

function dateFilterNotice(suppressFilter) {
  var key = currentAnalysisEntity;
  var df = suppressFilter ? '' : activeFilterValue[key];
  var lbl = activeFilterLabel[key];
  var dt = activeWindowTo || '';
  var pill = scanPillHtml();
  if (!df && !pill) return '';
  var sinceSuffix = '';
  if (lbl && lbl.toLowerCase().indexOf('since') === 0) {
    sinceSuffix = '';
  } else if (df && dt) {
    sinceSuffix = ' (' + df + ' \u2192 ' + dt + ')';
  } else if (df) {
    sinceSuffix = ' (since ' + df + ')';
  }
  var left = df
    ? '<span class="fn-icon">&#9202;</span><span>Filtered: <strong>' + lbl + '</strong>' + sinceSuffix + '</span>'
    : '';
  return '<div class="filter-notice" style="display:flex;align-items:center;justify-content:space-between;gap:8px;flex-wrap:wrap">'
    + '<div style="display:flex;align-items:center;gap:8px">' + left + '</div>'
    + '<span class="scan-pill-slot">' + pill + '</span>'
    + '</div>';
}

// Format an ISO date (2026-06-03) as a short readable date (3 Jun 2026).
function _fmtScanDate(iso) {
  var m = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  var p = String(iso).split('-');
  if (p.length !== 3) return String(iso);
  var y = parseInt(p[0], 10), mo = parseInt(p[1], 10), d = parseInt(p[2], 10);
  if (!mo || !d) return String(iso);
  return d + ' ' + (m[mo - 1] || '') + ' ' + y;
}

// Scan-status pill. On the dashboard there is always a scan, so this shows the
// date. Returns empty if no scan state is loaded.
function scanPillHtml() {
  var st = (typeof daSettings !== 'undefined' && daSettings.getScanState) ? daSettings.getScanState() : null;
  if (!st || !st.scannedAt) return '';
  return '<span style="display:inline-flex;align-items:center;gap:6px;font-size:12px;font-weight:500;padding:5px 12px;border-radius:999px;background:#E2EFDC;color:#2E5E2A;border:1px solid #CDE2C4;white-space:nowrap">'
    + '<svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="9"/><path d="M8.5 12.5l2.5 2.5 4.5-5"/></svg>'
    + 'Last scan ' + _fmtScanDate(st.scannedAt) + '</span>';
}

// Refresh the scan pill inside every filter bar. Called when the scan state
// loads and after a fresh scan, so the pill stays in sync without a full
// re-render of the entity views.
function _refreshScanPills() {
  var pill = scanPillHtml();
  var slots = document.querySelectorAll('.filter-notice .scan-pill-slot');
  for (var i = 0; i < slots.length; i++) { slots[i].innerHTML = pill; }
}

// Refresh the pills as soon as the scan state has loaded from the server.
if (typeof window !== 'undefined') { window.daOnScanReady = _refreshScanPills; }

// Coerce a score (number or { total } object) to a rounded number, or null.
function _scoreNum(x) {
  if (typeof x === 'number') return Math.round(x);
  if (x && typeof x.total === 'number') return Math.round(x.total);
  return null;
}

// Gather the per-entity scores after a full analysis and persist them as the
// scan snapshot, so the date and current scores are available on next load.
function _writeScanSnapshot() {
  if (typeof daSettings === 'undefined' || !daSettings.saveScanState) return;
  if (typeof computeEntityScores !== 'function') return;
  var keys = ['company', 'contact', 'sale', 'project'];
  var scores = {};
  for (var i = 0; i < keys.length; i++) {
    var s = computeEntityScores(keys[i]);
    if (s) scores[keys[i]] = {
      dataQuality: _scoreNum(s.dq),
      dataIntegrity: _scoreNum(s.integrity),
      adoption: _scoreNum(s.adoption),
      overall: _scoreNum(s.health)
    };
  }
  var d = new Date();
  function pad(n) { return (n < 10 ? '0' : '') + n; }
  var iso = d.getFullYear() + '-' + pad(d.getMonth() + 1) + '-' + pad(d.getDate());
  daSettings.saveScanState(iso, scores);
}

// v2: invalidate extra cache on reset
function resetEntity(key) {
  delete overviewData[key];
  var ec = entityConfig[key];
  if (ec && ec.udefId > 0) delete udefData[ec.udefId];
  if (key === 'requests') ticketData = null;
  if (key === 'company') {
    companyDetailData = null;
    companyDetailCatValue = '';
    companyCrossData = null;
  }
  if (key === 'activities') momentumData = null;
  delete entExtra[key];
  if (typeof invalidateExtraCache === 'function') invalidateExtraCache();
}

function reAnalyze(key) {
  // A fresh scan invalidates any cached company-detail results.
  if (key === 'company' && typeof companyDetailCache !== 'undefined') companyDetailCache = {};
  resetEntity(key);
  var rs = document.getElementById(key + 'Results');
  var st2 = document.getElementById(key + 'SubTabs');
  var eb = document.getElementById(key + 'ExportBtn') || document.getElementById('ticketExportBtn');
  if (rs) rs.style.display = 'none';
  if (st2) st2.style.display = 'none';
  if (eb) eb.style.display = 'none';
  if (entityConfig[key]) {
    startFullEntity(key);
  } else if (key === 'activities') {
    startMomentum(key);
  } else {
    startEntityOverview(key);
  }
}

// === ANALYZE ALL ===
var aaQueue = [];
var aaIdx = 0;
var aaEntities = ['company','contact','activities','sale','project','requests','selection','marketing'];
var aaEntityNames = {company:'Company',contact:'Contact',activities:'Activities',sale:'Sale',project:'Project',requests:'Requests',selection:'Selection',marketing:'Marketing'};

function onAADateChange() {}

function togGroup(headerRow) {
  var collapsed = headerRow.classList.toggle('group-collapsed');
  var next = headerRow.nextElementSibling;
  while (next && !next.classList.contains('assoc-group-header')) {
    next.style.display = collapsed ? 'none' : '';
    next = next.nextElementSibling;
  }
}

function togAssocGroup() {
  var tbl = document.getElementById('tbl-assoc');
  if (!tbl) return;
  var headers = tbl.querySelectorAll('.assoc-group-header');
  var allCollapsed = true;
  for (var i = 0; i < headers.length; i++) {
    if (!headers[i].classList.contains('group-collapsed')) { allCollapsed = false; break; }
  }
  for (var i = 0; i < headers.length; i++) {
    if (allCollapsed) { headers[i].classList.remove('group-collapsed'); }
    else { headers[i].classList.add('group-collapsed'); }
    var next = headers[i].nextElementSibling;
    while (next && !next.classList.contains('assoc-group-header')) {
      next.style.display = allCollapsed ? '' : 'none';
      next = next.nextElementSibling;
    }
  }
  var badge = document.getElementById('assocGroupToggle');
  if (badge) {
    var chev = badge.querySelector('.badge-chevron');
    if (chev) { if (allCollapsed) { chev.classList.remove('collapsed'); } else { chev.classList.add('collapsed'); } }
  }
}

function onSetupDateChange() {}

function startAnalyzeAll() {
  dpInit();
  var _per = dpResolve('aa');
  if (dpMode('aa') === 'custom') {
    if (!_per.from) { var _af = document.getElementById('aaDateFrom'); if (_af) { _af.style.borderColor = '#c62828'; _af.focus(); } return; }
    if (!_per.to) { var _at = document.getElementById('aaDateTo'); if (_at) { _at.style.borderColor = '#c62828'; _at.focus(); } return; }
  }
  activeScope = dpScope('aa');
  activeWindowTo = _per.to;
  var globalVal = _per.from;
  aaQueue = [];
  for (var i = 0; i < aaEntities.length; i++) {
    var k = aaEntities[i];
    var cb = document.getElementById('aa_' + k);
    if (cb && cb.checked) {
      aaQueue.push(k);
      var overrideSel = document.getElementById('aaDate_' + k);
      var overrideVal = getSelDateVal(overrideSel);
      var useVal = overrideVal ? overrideVal : globalVal;
      var entSel = document.getElementById('dateFilter_' + k);
      if (entSel) {
        if (useVal && useVal.match(/^\d{4}-\d{2}-\d{2}$/)) {
          var found = false;
          for (var j = 0; j < entSel.options.length; j++) {
            if (entSel.options[j].value === useVal) { found = true; break; }
          }
          if (!found) {
            entSel.value = 'custom';
            handleDateSelect(entSel);
            var di = entSel.parentNode.querySelector('.custom-date-input');
            if (di) di.value = useVal;
          } else {
            entSel.value = useVal;
          }
        } else {
          entSel.value = useVal;
        }
      }
    }
  }
  if (aaQueue.length === 0) return;

  document.getElementById('aaStartScreen').style.display = 'none';
  document.getElementById('aaDoneBanner').style.display = 'none';
  var progScreen = document.getElementById('aaProgressScreen');
  progScreen.style.display = '';

  var listHtml = '';
  var svgPending = '<svg width="14" height="14" viewBox="0 0 14 14"><circle cx="7" cy="7" r="6" fill="none" stroke="#d3d1c7" stroke-width="1.5"/></svg>';
  for (var j = 0; j < aaQueue.length; j++) {
    listHtml += '<div class="aa-pill aa-pill-pending" id="aaProg_' + aaQueue[j] + '">';
    listHtml += '<span class="aa-pill-icon" id="aaIco_' + aaQueue[j] + '">' + svgPending + '</span>';
    listHtml += '<span>' + aaEntityNames[aaQueue[j]] + '</span>';
    listHtml += '</div>';
  }
  document.getElementById('aaEntityList').innerHTML = listHtml;
  // Mirror entity list to setup screen
  var setupList = document.getElementById('setupEntityList');
  if (setupList) {
    var setupHtml = '';
    for (var j = 0; j < aaQueue.length; j++) {
      setupHtml += '<div class="aa-pill aa-pill-pending" id="setupProg_' + aaQueue[j] + '">';
      setupHtml += '<span class="aa-pill-icon" id="setupIco_' + aaQueue[j] + '">' + svgPending + '</span>';
      setupHtml += '<span>' + aaEntityNames[aaQueue[j]] + '</span>';
      setupHtml += '</div>';
    }
    setupList.innerHTML = setupHtml;
  }

  // Initialize SVG circle
  var circBar = document.getElementById('aaCircleBar');
  var circPct = document.getElementById('aaCirclePercent');
  if (circBar) circBar.setAttribute('stroke-dashoffset', '314.16');
  if (circPct) circPct.textContent = '0' + P;
  var labelEl = document.getElementById('aaProgressLabel');
  if (labelEl) labelEl.textContent = 'Analyzing data... 0/' + aaQueue.length;
  _aaCurrentPct = 0;
  _aaAnimTarget = 0;
  _aaAnimCurrent = 0;
  _aaCompleted = 0;
  _aaRunning = 0;
  _aaDoneSteps = 0;
  _aaTotalSteps = 0;
  for (var si = 0; si < aaQueue.length; si++) _aaTotalSteps += _aaEntitySteps(aaQueue[si]);

  // Track start time for done label
  _aaStartTime = Date.now();
  if (_aaElapsedTimer) clearInterval(_aaElapsedTimer);

  aaIdx = 0;
  runNextAA();
}

// Steps per entity (= number of AJAX responses expected)
function _aaEntitySteps(key) {
  var ec = entityConfig[key];
  if (!ec) return 1; // simple entity (selection, marketing)
  var steps;
  if (key === 'company') steps = 6; // overview + udef + core + quality + detail + extra
  else if (key === 'activities') steps = 2; // momentum + overview
  else { steps = 2; if (ec.udefId > 0 || ec.hasTicketFields) steps++; }
  if (key !== 'activities' && aaHasRules(key)) steps++; // extra RulesFetch call
  return steps;
}

var _aaCurrentPct = 0;
var _aaTotalSteps = 0;
var _aaDoneSteps = 0;

var _aaAnimTarget = 0;
var _aaAnimCurrent = 0;
var _aaAnimRAF = null;

function _aaAnimateCounter() {
  var diff = _aaAnimTarget - _aaAnimCurrent;
  if (diff > 0.5) {
    _aaAnimCurrent += diff * 0.08;
    var v = Math.round(_aaAnimCurrent);
    if (_aaCompleted < aaQueue.length && v > 99) v = 99; // never show 100 until all entities done
    var circPct = document.getElementById('aaCirclePercent');
    if (circPct) circPct.textContent = v + P;
    var circPct2 = document.getElementById('setupCirclePercent');
    if (circPct2 && setupAnalysisRunning) circPct2.textContent = v + P;
    _aaAnimRAF = requestAnimationFrame(_aaAnimateCounter);
  } else {
    _aaAnimCurrent = _aaAnimTarget;
    var v2 = Math.round(_aaAnimCurrent);
    if (_aaCompleted < aaQueue.length && v2 > 99) v2 = 99;
    var circPct = document.getElementById('aaCirclePercent');
    if (circPct) circPct.textContent = v2 + P;
    var circPct2 = document.getElementById('setupCirclePercent');
    if (circPct2 && setupAnalysisRunning) circPct2.textContent = v2 + P;
    _aaAnimRAF = null;
  }
}

function _aaSetProgress(pct) {
  if (pct < _aaCurrentPct) pct = _aaCurrentPct;
  _aaCurrentPct = pct;
  var offset = 314.16 * (1 - pct / 100);
  var circBar = document.getElementById('aaCircleBar');
  var labelEl = document.getElementById('aaProgressLabel');
  if (circBar) circBar.setAttribute('stroke-dashoffset', offset);
  // Smooth counter animation instead of instant text
  _aaAnimTarget = pct;
  if (!_aaAnimRAF) _aaAnimRAF = requestAnimationFrame(_aaAnimateCounter);
  if (labelEl) labelEl.textContent = 'Analyzing data... ' + _aaCompleted + '/' + aaQueue.length;
  setupProgressUpdate(Math.round(pct));
}

function _aaStepDone() {
  _aaDoneSteps++;
  if (_aaDoneSteps > _aaTotalSteps) _aaDoneSteps = _aaTotalSteps;
  var pct = _aaTotalSteps > 0 ? (_aaDoneSteps / _aaTotalSteps) * 100 : 0;
  if (pct > 99) pct = 99; // only _aaOnAllDone sets 100%
  _aaSetProgress(pct);
}

var _aaConcurrency = 3;
var _aaCompleted = 0;
var _aaRunning = 0;
var _aaElapsedTimer = null;
var _aaStartTime = 0;

function _aaOnAllDone() {
  _aaSetProgress(100);
  // Stop timer
  if (_aaElapsedTimer) { clearInterval(_aaElapsedTimer); _aaElapsedTimer = null; }
  var elapsed = Math.round((Date.now() - _aaStartTime) / 1000);

  // Persist the scan snapshot (date + per-entity scores) for the badge and settings
  _writeScanSnapshot();
  _refreshScanPills();

  // Show "done" state: replace percentage text with a subtle checkmark inside existing circle
  setTimeout(function() {
    // Replace just the text element with a checkmark polyline (same weight as circle stroke)
    var circPct = document.getElementById('aaCirclePercent');
    if (circPct) {
      var ns = 'http://www.w3.org/2000/svg';
      var check = document.createElementNS(ns, 'polyline');
      check.setAttribute('points', '44,62 54,72 76,48');
      check.setAttribute('stroke', '#06423e');
      check.setAttribute('stroke-width', '4');
      check.setAttribute('fill', 'none');
      check.setAttribute('stroke-linecap', 'round');
      check.setAttribute('stroke-linejoin', 'round');
      check.setAttribute('class', 'aa-check-draw');
      circPct.parentNode.appendChild(check);
      circPct.remove();
    }
    var circPct2 = document.getElementById('setupCirclePercent');
    if (circPct2 && setupAnalysisRunning) {
      var ns2 = 'http://www.w3.org/2000/svg';
      var check2 = document.createElementNS(ns2, 'polyline');
      check2.setAttribute('points', '44,62 54,72 76,48');
      check2.setAttribute('stroke', '#06423e');
      check2.setAttribute('stroke-width', '4');
      check2.setAttribute('fill', 'none');
      check2.setAttribute('stroke-linecap', 'round');
      check2.setAttribute('stroke-linejoin', 'round');
      check2.setAttribute('class', 'aa-check-draw');
      circPct2.parentNode.appendChild(check2);
      circPct2.remove();
    }

    var labelEl = document.getElementById('aaProgressLabel');
    if (labelEl) labelEl.textContent = 'Analysis complete \u00B7 ' + elapsed + 's';
    var labelEl2 = document.getElementById('setupProgressLabel');
    if (labelEl2 && setupAnalysisRunning) labelEl2.textContent = 'Analysis complete \u00B7 ' + elapsed + 's';

  }, 400);

  // Transition to dashboard after delay
  setTimeout(function() {
    // Auto-populate standalone Extra Tables tab from cache
    if (typeof _extraCache !== 'undefined' && _extraCache && _extraCache.ready) {
      extraTables = _extraCache.tables;
      extraData = {};
      for (var ei = 0; ei < _extraCache.tables.length; ei++) {
        extraData[_extraCache.tables[ei].id] = _extraCache.data[_extraCache.tables[ei].id];
      }
      var extraStartEl = document.getElementById('extraStart');
      var extraResultsEl = document.getElementById('extraResults');
      var extraExportEl = document.getElementById('extraExportBtn');
      var extraBtn = document.getElementById('extraAnalyzeBtn');
      if (extraStartEl) extraStartEl.style.display = 'none';
      if (extraResultsEl) extraResultsEl.style.display = 'block';
      if (extraResultsEl && !extraResultsEl.querySelector('.filter-notice')) extraResultsEl.insertAdjacentHTML('afterbegin', dateFilterNotice(true));
      if (extraExportEl) extraExportEl.style.display = '';
      if (extraBtn) { extraBtn.disabled = false; extraBtn.onclick = function(){ document.getElementById('extraResults').style.display = 'none'; document.getElementById('extraCards').innerHTML = ''; extraData = {}; invalidateExtraCache(); startExtra(); }; }
      var wr = 0;
      for (var ei = 0; ei < extraTables.length; ei++) {
        var ed = extraData[extraTables[ei].id];
        if (ed && ed.relationFields && ed.relationFields.length > 0) wr++;
      }
      var extraSumEl = document.getElementById('extraSummary');
      if (extraSumEl) extraSumEl.textContent = extraTables.length + ' extra tables, ' + wr + ' with relation fields.';
      var extraCardsEl = document.getElementById('extraCards');
      if (extraCardsEl) {
        extraCardsEl.innerHTML = '';
        for (var ei = 0; ei < extraTables.length; ei++) {
          var ed = extraData[extraTables[ei].id];
          if (ed) extraCardsEl.innerHTML += renderExtra(ed, ei);
        }
      }
    }

    document.getElementById('aaProgressScreen').style.display = 'none';
    document.getElementById('aaStartScreen').style.display = '';
    document.getElementById('aaDoneBanner').style.display = '';
    document.getElementById('aaDoneSummary').textContent = aaQueue.length + ' entities analyzed in ' + elapsed + 's.';
    if (setupAnalysisRunning) { setupAnalysisComplete(); return; }
    document.getElementById('aaStartBtn').innerHTML = '<img src="data:image/svg+xml;base64,' + icoPlayO + '"> Run Again';
  }, 2000);
}

function runNextAA() {
  if (_aaCompleted === 0 && _aaRunning === 0) {
    _aaCurrentPct = 0;
  }

  // Launch entities up to concurrency limit
  while (_aaRunning < _aaConcurrency && aaIdx < aaQueue.length) {
    _aaLaunchEntity(aaQueue[aaIdx]);
    aaIdx++;
  }
}

function _aaSortPills() {
  var containers = [document.getElementById('aaEntityList'), document.getElementById('setupEntityList')];
  for (var ci = 0; ci < containers.length; ci++) {
    var c = containers[ci];
    if (!c) continue;
    var pills = Array.prototype.slice.call(c.children);
    pills.sort(function(a, b) {
      var order = {'aa-pill aa-pill-loading':0, 'aa-pill aa-pill-done':1, 'aa-pill aa-pill-pending':2};
      return (order[a.className] || 2) - (order[b.className] || 2);
    });
    for (var i = 0; i < pills.length; i++) c.appendChild(pills[i]);
  }
}

function _aaLaunchEntity(key) {
  _aaRunning++;

  var svgLoading = '<svg width="14" height="14" viewBox="0 0 14 14"><circle cx="7" cy="7" r="6" fill="none" stroke="#06423e" stroke-width="1.5"/><circle cx="7" cy="7" r="2" fill="#06423e"><animate attributeName="opacity" values="1;0.3;1" dur="1.2s" repeatCount="indefinite"/></circle></svg>';
  var svgDone = '<svg width="14" height="14" viewBox="0 0 14 14"><circle cx="7" cy="7" r="6" fill="#0F6E56"/><polyline points="4,7 6.2,9.2 10,5" stroke="#fff" stroke-width="1.5" fill="none" stroke-linecap="round" stroke-linejoin="round"/></svg>';

  var pillEl = document.getElementById('aaProg_' + key);
  var icoEl = document.getElementById('aaIco_' + key);
  var pillEl2 = document.getElementById('setupProg_' + key);
  var icoEl2 = document.getElementById('setupIco_' + key);
  function setEntityStatus(state) {
    if (state === 'loading') {
      if (pillEl) pillEl.className = 'aa-pill aa-pill-loading';
      if (icoEl) icoEl.innerHTML = svgLoading;
      if (pillEl2) pillEl2.className = 'aa-pill aa-pill-loading';
      if (icoEl2) icoEl2.innerHTML = svgLoading;
    } else if (state === 'done') {
      if (pillEl) pillEl.className = 'aa-pill aa-pill-done';
      if (icoEl) icoEl.innerHTML = svgDone;
      if (pillEl2) pillEl2.className = 'aa-pill aa-pill-done';
      if (icoEl2) icoEl2.innerHTML = svgDone;
    }
    _aaSortPills();
  }
  setEntityStatus('loading');

  // Capture filter params synchronously BEFORE any async calls
  captureEntityFilter(key);

  function markStepDone(stepId) {
    _aaStepDone();
  }

  function onEntityDone() {
    _aaRunning--;
    _aaCompleted++;
    setEntityStatus('done');

    // Update label with completed count
    var labelEl = document.getElementById('aaProgressLabel');
    if (labelEl) labelEl.textContent = 'Analyzing data... ' + _aaCompleted + '/' + aaQueue.length;
    var labelEl2 = document.getElementById('setupProgressLabel');
    if (labelEl2 && setupAnalysisRunning) labelEl2.textContent = 'Analyzing data... ' + _aaCompleted + '/' + aaQueue.length;

    if (_aaCompleted >= aaQueue.length) {
      _aaOnAllDone();
    } else {
      runNextAA();
    }
  }

  // Launch the entity (all URL construction happens synchronously here)
  if (entityConfig[key]) {
    aaRunFullEntity(key, onEntityDone, markStepDone);
  } else if (key === 'activities') {
    aaRunMomentumEntity(key, onEntityDone);
  } else {
    aaRunSimpleEntity(key, onEntityDone);
  }
}

function aaRunSimpleEntity(key, cb) {
  ajax(overviewUrl + String.fromCharCode(38) + 'entity=' + key + getDateFilterParam() + getExclParam(key), function(d) {
    if (d) renderEntityOverview(key, d);
    _aaStepDone();
    var rs = document.getElementById(key + 'Results');
    var ab = document.getElementById(key + 'AnalyzeBtn');
    var eb = document.getElementById(key + 'ExportBtn');
    if (rs) rs.style.display = '';
    if (ab) ab.textContent = 'Re-analyze';
    if (eb) eb.style.display = '';
    if (cb) cb();
  });
}

function aaRunMomentumEntity(key, cb) {
  var dfParam = getDateFilterParam();
  var loadsDone = 0;
  var totalLoads = 2;

  function checkDone() {
    loadsDone++;
    _aaStepDone();
    if (loadsDone < totalLoads) return;
    // Render momentum
    if (typeof renderMomentum === 'function') renderMomentum(key, momentumData);
    var rs = document.getElementById(key + 'Results');
    var st = document.getElementById(key + 'SubTabs');
    var ab = document.getElementById(key + 'AnalyzeBtn');
    var eb = document.getElementById(key + 'ExportBtn');
    if (rs) rs.style.display = '';
    if (st) st.style.display = '';
    if (ab) ab.textContent = 'Re-analyze';
    if (eb) eb.style.display = '';
    if (cb) cb();
  }

  // Parallel 1: Momentum data
  var amp = String.fromCharCode(38);
  var etParam = '';
  if (typeof daSettings !== 'undefined') {
    var mmS = daSettings.getMomentumSettings();
    if (mmS.excludedTypes && mmS.excludedTypes.length > 0) {
      etParam = amp + 'excludeTypes=' + encodeURIComponent(mmS.excludedTypes.join(','));
    }
  }
  ajax(momentumUrl + amp + 'dummy=1' + dfParam + etParam, function(d) {
    momentumData = d;
    checkDone();
  });

  // Parallel 2: Overview (for Overview sub-tab)
  ajax(overviewUrl + String.fromCharCode(38) + 'entity=' + key + dfParam + getExclParam(key), function(d) {
    if (d) renderEntityOverview(key, d);
    checkDone();
  });
}

// v4: PARALLEL sub-loads with checklist progress reporting
function aaRunFullEntity(key, cb, markStepDone) {
  var ec = entityConfig[key];
  var tabKey = ec.tabKey;
  var hasDetails = (key === 'company');

  var dfParam = getDateFilterParam();

  var totalSteps = 1; // overview always
  if (ec.udefId > 0) totalSteps++;
  if (ec.hasTicketFields) totalSteps++;
  if (hasDetails) totalSteps += 3; // core + quality + detail
  totalSteps++; // extra tables always
  var doRules = aaHasRules(key);
  if (doRules) totalSteps++; // RulesFetch when the entity has active rules

  var completed = 0;
  function stepDone(stepId) {
    completed++;
    if (markStepDone) markStepDone(stepId);
    if (completed >= totalSteps) {
      var rs = document.getElementById(tabKey + 'Results');
      var st2 = document.getElementById(tabKey + 'SubTabs');
      var ab = document.getElementById(tabKey + 'AnalyzeBtn');
      var eb = document.getElementById(tabKey + 'ExportBtn') || document.getElementById('ticketExportBtn');
      if (rs) rs.style.display = '';
      if (st2) st2.style.display = '';
      if (ab) ab.textContent = 'Re-analyze';
      if (eb) eb.style.display = '';
      if (typeof renderDQScore === 'function') renderDQScore(key);
      if (typeof renderScoreBanner === 'function') renderScoreBanner(key);
      if (typeof renderAdoptionTab === 'function') renderAdoptionTab(key);
      if (cb) cb();
    }
  }

  // === PARALLEL 1: Overview ===
  ajax(overviewUrl + String.fromCharCode(38) + 'entity=' + key + dfParam + getExclParam(key), function(d) {
    if (d) renderEntityOverview(key, d);
    stepDone('overview');
  });

  // === PARALLEL: Conditional rules ===
  if (doRules) {
    ajax(rulesUrl + String.fromCharCode(38) + 'entity=' + key + dfParam + getRulesParam(key), function(d) {
      if (d && d.rules) { if (!window.rulesData) window.rulesData = {}; window.rulesData[key] = d.rules; }
      stepDone('rules');
    });
  }

  // === PARALLEL 2: UDEF or Ticket Fields ===
  if (ec.udefId > 0) {
    loadEntityUdefQuiet(ec.udefId, tabKey, ec.udefIdx, function() {
      stepDone('fields');
    });
  } else if (ec.hasTicketFields) {
    loadTicketFieldsQuiet(function() {
      stepDone('tfields');
    });
  }

  // === PARALLEL 3: Company Core + Quality + Detail (3 scripts) ===
  if (hasDetails) {
    var catParam = '';
    if (companyDetailCatValue) {
      catParam = String.fromCharCode(38) + 'categoryName=' + encodeURIComponent(companyDetailCatValue);
    }
    if (!companyDetailData) companyDetailData = {};

    // 3a: CompanyCoreFetch (~6s)
    ajax(coreUrl + dfParam + catParam, function(d) {
      if (d) {
        if (d.activityHealth) companyDetailData.activityHealth = d.activityHealth;
        if (d.trend) companyDetailData.trend = d.trend;
        if (d.trendMonthly) companyDetailData.trendMonthly = d.trendMonthly;
        if (d.trendBefore !== undefined) companyDetailData.trendBefore = d.trendBefore;
      }
      if (typeof renderCompanyDetails === 'function') renderCompanyDetails(companyDetailData);
      stepDone('core');
    });

    // 3b: CompanyQualityFetch (~5s)
    ajax(qualityUrl + dfParam + catParam + getExclParam('company'), function(d) {
      if (d) {
        if (d.quality) companyDetailData.quality = d.quality;
        if (d.funnel) companyDetailData.funnel = d.funnel;
        if (d.churnRisk) companyDetailData.churnRisk = d.churnRisk;
      }
      if (companyDetailData.funnel) {
        companyCrossData = companyDetailData;
        if (typeof renderCrossEntityFunnel === 'function') renderCrossEntityFunnel(companyDetailData);
      }
      if (typeof renderCompanyDetails === 'function') renderCompanyDetails(companyDetailData);
      if (typeof renderDQScore === 'function') renderDQScore('company');
      stepDone('quality');
    });

    // 3c: CompanyDetailFetch (~25s)
    ajax(detailUrl + dfParam + catParam, function(d) {
      if (d) {
        if (d.associates) companyDetailData.associates = d.associates;
        if (d.categoryEffectiveness) companyDetailData.categoryEffectiveness = d.categoryEffectiveness;
      }
      if (typeof renderCompanyDetails === 'function') renderCompanyDetails(companyDetailData);
      stepDone('assoc');
    });
  }

  // === PARALLEL 4: Extra Tables (from cache) ===
  aaLoadExtra(key, ec, tabKey, function() {
    stepDone('tables');
  });
}

function aaLoadExtra(key, ec, tabKey, cb) {
  loadEntityExtraWithProgress(tabKey, ec.extraType, null, function() {
    var rs = document.getElementById(tabKey + 'Results');
    var st2 = document.getElementById(tabKey + 'SubTabs');
    // Don't show results here — let stepDone handle it when ALL are complete
    if (cb) cb();
  });
}

var entityIcons = {
  Company: 'data:image/svg+xml;base64,' + cfgEl.getAttribute('data-ico-company'),
  Person:  'data:image/svg+xml;base64,' + cfgEl.getAttribute('data-ico-contact'),
  Project: 'data:image/svg+xml;base64,' + cfgEl.getAttribute('data-ico-project'),
  Sale:    'data:image/svg+xml;base64,' + cfgEl.getAttribute('data-ico-sale')
};

// === NAVIGATION ===
function sw(id, el) {
  document.querySelectorAll('.tab-panel').forEach(function(p) { p.classList.remove('active'); });
  document.querySelectorAll('.nav-item').forEach(function(n) { n.classList.remove('active'); });
  document.getElementById('tab-' + id).classList.add('active');
  if (el) el.classList.add('active');
}

function launchDashboard() {
  var setupDate = document.getElementById('setupDateFilter');
  var _per = dpResolve('setup');
  if (dpMode('setup') === 'custom') {
    if (!_per.from) { var _sf = document.getElementById('setupDateFrom'); if (_sf) { _sf.style.borderColor = '#c62828'; _sf.focus(); } return; }
    if (!_per.to) { var _st = document.getElementById('setupDateTo'); if (_st) { _st.style.borderColor = '#c62828'; _st.focus(); } return; }
  }
  activeScope = dpScope('setup');
  activeWindowTo = _per.to;
  var setupVal = _per.from;
  dpMirror('setup', 'aa');
  for (var i = 0; i < aaEntities.length; i++) {
    var k = aaEntities[i];
    var setupCb = document.getElementById('setup_' + k);
    var aaCb = document.getElementById('aa_' + k);
    if (setupCb && aaCb) aaCb.checked = setupCb.checked;
    var setupEntDate = document.getElementById('setupDate_' + k);
    var aaEntDate = document.getElementById('aaDate_' + k);
    if (setupEntDate && aaEntDate) {
      aaEntDate.value = setupEntDate.value;
      if (setupEntDate.value === 'custom') {
        handleDateSelect(aaEntDate);
        var sdi = setupEntDate.parentNode.querySelector('.custom-date-input');
        var adi = aaEntDate.parentNode.querySelector('.custom-date-input');
        if (sdi && adi) adi.value = sdi.value;
      }
    }
    var entSel = document.getElementById('dateFilter_' + k);
    var setupEntVal = setupEntDate ? getSelDateVal(setupEntDate) : '';
    var effectiveVal = setupEntVal || setupVal;
    if (entSel && effectiveVal) {
      var found = false;
      for (var j = 0; j < entSel.options.length; j++) {
        if (entSel.options[j].value === effectiveVal) { found = true; break; }
      }
      if (!found && effectiveVal.match(/^\d{4}-\d{2}-\d{2}$/)) {
        entSel.value = 'custom';
        handleDateSelect(entSel);
        var edi = entSel.parentNode.querySelector('.custom-date-input');
        if (edi) edi.value = effectiveVal;
      } else if (found) {
        entSel.value = effectiveVal;
      } else if (setupDate) {
        entSel.value = setupDate.value;
      }
    } else if (entSel && setupDate) {
      if (setupDate.value === 'custom' && setupVal) {
        entSel.value = 'custom';
        handleDateSelect(entSel);
        var edi2 = entSel.parentNode.querySelector('.custom-date-input');
        if (edi2) edi2.value = setupVal;
      } else {
        entSel.value = setupDate.value;
      }
    }
  }
  var setupTabs = document.querySelector('.setup-tabs');
  var loadingDiv = document.getElementById('setupLoadingState');
  var tabContent = document.querySelector('.setup-tab-content');
  var actionBar = document.getElementById('setupActionBar');
  if (tabContent) tabContent.style.display = 'none';
  if (setupTabs) setupTabs.style.display = 'none';
  if (actionBar) actionBar.style.display = 'none';
  if (loadingDiv) loadingDiv.style.display = 'flex';
  setupAnalysisRunning = true;
  startAnalyzeAll();
}

var setupAnalysisRunning = false;

function setupProgressUpdate(pct) {
  if (!setupAnalysisRunning) return;
  if (_aaCompleted < aaQueue.length && pct > 99) pct = 99;
  var offset = 314.16 * (1 - pct / 100);
  var circBar = document.getElementById('setupCircleBar');
  var circPct = document.getElementById('setupCirclePercent');
  var labelEl = document.getElementById('setupProgressLabel');
  if (circBar) circBar.setAttribute('stroke-dashoffset', offset);
  if (circPct) circPct.textContent = pct + P;
  if (labelEl) labelEl.textContent = 'Analyzing data... ' + _aaCompleted + '/' + aaQueue.length;
}

function setupAnalysisComplete() {
  setupAnalysisRunning = false;
  var circBar = document.getElementById('setupCircleBar');
  var circPct = document.getElementById('setupCirclePercent');
  var labelEl = document.getElementById('setupProgressLabel');
  if (circBar) circBar.setAttribute('stroke-dashoffset', '0');
  if (circPct) circPct.textContent = '100' + P;
  if (labelEl) labelEl.textContent = 'Analysis complete';

  // Switch the setup overlay out for the first analysed entity's dashboard.
  function _setupNav() {
    document.body.classList.remove('sidebar-hidden');
    if (aaQueue.length > 0) {
      var navItems = document.querySelectorAll('.nav-item');
      for (var n = 0; n < navItems.length; n++) {
        var oc = navItems[n].getAttribute('onclick') || '';
        if (oc.indexOf("'" + aaQueue[0] + "'") > -1) { sw(aaQueue[0], navItems[n]); break; }
      }
    }
  }

  // Let the finished circle sit for a beat, then run the Hugo iris over the swap.
  setTimeout(function() {
    var overlay = document.getElementById('setupScreen');
    if (window.HugoZoom && typeof HugoZoom.show === 'function') {
      HugoZoom.show('Opening dashboard\u2026');
      // Once the iris has fully closed over the screen, swap underneath and reveal.
      setTimeout(function() {
        _setupNav();
        if (overlay) { overlay.classList.add('hidden'); overlay.classList.remove('hiding'); }
        HugoZoom.reveal();
      }, 520);
    } else {
      if (overlay) overlay.classList.add('hiding');
      _setupNav();
      setTimeout(function() { if (overlay) overlay.classList.add('hidden'); }, 500);
    }
  }, 480);
}
function st(ent, sub, el) {
  var p = document.getElementById('tab-' + ent);
  p.querySelectorAll('.sub-panel').forEach(function(s) { s.classList.remove('active'); });
  var t = document.getElementById(ent + '-' + sub);
  if (t) t.classList.add('active');
  p.querySelectorAll('.sub-tab').forEach(function(t) { t.classList.remove('active'); });
  if (el) el.classList.add('active');
  var sr = document.getElementById(ent + 'SubTabsRight');
  if (sr) sr.style.display = (sub === 'adoption') ? '' : 'none';
}

// === HELPERS ===
function barCls(p) {
  if (p >= 70) return 'bar-high';
  if (p >= 40) return 'bar-mid';
  if (p >= 15) return 'bar-low';
  return 'bar-none';
}
function fillBar(p, h, c) {
  var cls = c ? '' : barCls(p);
  var style = 'width:' + p + P;
  if (c) style += ';background:' + c;
  var hStyle = h ? 'height:' + h + 'px' : '';
  return '<div class="fill-bar" style="' + hStyle + '"><div class="bar ' + cls + '" style="' + style + '"></div></div>';
}
function barCell(pct, color) {
  return '<span class="bar-cell">' + fillBar(pct, null, color) + '<span class="pct-text">' + pct + P + '</span></span>';
}
function slColor(pct) {
  if (pct >= 70) return 'var(--sl-good)';
  if (pct >= 40) return 'var(--sl-ok)';
  if (pct >= 15) return 'var(--sl-warn)';
  return 'var(--sl-bad)';
}
function slColorInv(pct) {
  if (pct < 10) return 'var(--sl-good)';
  if (pct < 30) return 'var(--sl-ok)';
  if (pct < 60) return 'var(--sl-warn)';
  return 'var(--sl-bad)';
}
// Verdict word + inline tint colours for a score, thresholds aligned 1:1 with slColor (70/40/15).
// Colours are returned inline so the pill renders correctly even if CSS and JS are briefly out of sync.
function slBand(pct) {
  if (pct >= 70) return { w: 'Good', bg: '#e7f3e8', fg: '#1e5e22', dot: 'var(--sl-good)' };
  if (pct >= 40) return { w: 'Moderate', bg: '#fdf3df', fg: '#9a6800', dot: 'var(--sl-ok)' };
  if (pct >= 15) return { w: 'Needs attention', bg: '#fceee2', fg: '#aa4a00', dot: 'var(--sl-warn)' };
  return { w: 'Critical', bg: '#fbeaea', fg: '#a32d2d', dot: 'var(--sl-bad)' };
}
// Inline info icon with a hover tooltip (no icon font dependency)
var SB_INFO_SVG = '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="9"/><line x1="12" y1="11" x2="12" y2="16"/><line x1="12" y1="7.5" x2="12.01" y2="7.5"/></svg>';
function sbInfoIcon(text) {
  return '<span class="sb-iw">' + SB_INFO_SVG + '<span class="sb-itip">' + text + '</span></span>';
}

// === Attention helper ===
// badness = how-bad percentage (0-100). For completeness/adoption pass the gap (100 - filled%),
// for integrity/quality flags pass the affected%. weightKey is the row's importance or weight.
// Rows that carry no weight in the score return -1 so they sort to the bottom and show no icon.
function attnWeightFactor(weightKey) {
  var w = (weightKey || '').toString().toLowerCase();
  if (w === 'required' || w === 'high') return 1;
  if (w === 'normal' || w === 'medium') return 0.5;
  return 0;
}
function attnScore(badness, weightKey) {
  var f = attnWeightFactor(weightKey);
  if (f === 0) return -1;
  var b = badness; if (b < 0) b = 0; if (b > 100) b = 100;
  return Math.round(b * f);
}
function attnBand(score) {
  if (score < 0) return 'none';
  if (score >= 50) return 'high';
  if (score >= 25) return 'medium';
  return 'low';
}
function attnBandWord(score) {
  var b = attnBand(score);
  return b === 'high' ? 'High' : (b === 'medium' ? 'Medium' : 'Low');
}
var ATTN_HIGH_SVG = '<svg width="18" height="18" viewBox="0 0 24 24"><circle cx="12" cy="12" r="10" fill="#c62828"/><rect x="11" y="6.5" width="2" height="7" rx="1" fill="#fff"/><circle cx="12" cy="16.5" r="1.2" fill="#fff"/></svg>';
var ATTN_MED_SVG = '<svg width="18" height="18" viewBox="0 0 24 24"><path d="M12 3 L22 20 L2 20 Z" fill="#ef6c00"/><rect x="11" y="9" width="2" height="6" rx="1" fill="#fff"/><circle cx="12" cy="17.5" r="1.1" fill="#fff"/></svg>';
// Icon + hover tooltip for an attention cell. Empty string for rows that carry no weight.
function attnIconHtml(score, reason) {
  var band = attnBand(score);
  if (band === 'none') return '';
  var icon = band === 'high' ? ATTN_HIGH_SVG : (band === 'medium' ? ATTN_MED_SVG : '<span class="attn-dot"></span>');
  return '<span class="attn-iw">' + icon + '<span class="attn-tip">' + reason + '</span></span>';
}
// Sort row objects {attn, badness, html} by attention desc, worst-percentage first on ties, then join.
function attnSortRows(rows) {
  rows.sort(function(a, b) {
    if (b.attn !== a.attn) return b.attn - a.attn;
    return b.badness - a.badness;
  });
  var out = '';
  for (var i = 0; i < rows.length; i++) out += rows[i].html;
  return out;
}
// Plain-language health adjective for the verdict sentence
function slHealthWord(pct) {
  if (pct >= 70) return 'healthy';
  if (pct >= 40) return 'moderately healthy';
  if (pct >= 15) return 'below par';
  return 'in poor shape';
}
// Shared colour key shown once per scored page
function slKeyHtml() {
  return '<div class="sb-key">'
    + '<span class="sb-keychip"><span class="sb-kdot" style="background:var(--sl-bad)"></span>&lt;15</span>'
    + '<span class="sb-keychip"><span class="sb-kdot" style="background:var(--sl-warn)"></span>15\u201339</span>'
    + '<span class="sb-keychip"><span class="sb-kdot" style="background:var(--sl-ok)"></span>40\u201369</span>'
    + '<span class="sb-keychip"><span class="sb-kdot" style="background:var(--sl-good)"></span>70+ Good</span>'
    + '<span class="sb-keynote">Absolute, measured against your own data</span>'
    + '</div>';
}
function sortT(tid, col) {
  var t = document.getElementById(tid);
  if (!t) return;
  var tb = t.querySelector('tbody');
  var th = t.querySelectorAll('th')[col];
  var asc = th.dataset.sortDir !== 'asc';
  th.dataset.sortDir = asc ? 'asc' : 'desc';
  t.querySelectorAll('.sort-arrow').forEach(function(i) { i.classList.remove('active'); i.innerHTML = svgSortN; });
  th.querySelector('.sort-arrow').classList.add('active');
  th.querySelector('.sort-arrow').innerHTML = asc ? svgSortA : svgSortD;

  var sortFn = function(a, b) {
    var av = a.cells[col].dataset.sortValue || a.cells[col].textContent.trim();
    var bv = b.cells[col].dataset.sortValue || b.cells[col].textContent.trim();
    var an = parseFloat(av), bn = parseFloat(bv);
    if (!isNaN(an) && !isNaN(bn)) { av = an; bv = bn; }
    if (asc) return av > bv ? 1 : av < bv ? -1 : 0;
    return av < bv ? 1 : av > bv ? -1 : 0;
  };

  var groupHeaders = tb.querySelectorAll('.assoc-group-header');
  if (groupHeaders.length > 0) {
    for (var gi = 0; gi < groupHeaders.length; gi++) {
      var gh = groupHeaders[gi];
      var members = [];
      var next = gh.nextElementSibling;
      while (next && !next.classList.contains('assoc-group-header')) {
        members.push(next);
        next = next.nextElementSibling;
      }
      members.sort(sortFn);
      var ref = gh;
      for (var mi = 0; mi < members.length; mi++) {
        ref.parentNode.insertBefore(members[mi], ref.nextSibling);
        ref = members[mi];
      }
    }
  } else {
    var rows = Array.from(tb.querySelectorAll('tr'));
    rows.sort(sortFn);
    rows.forEach(function(r) { tb.appendChild(r); });
  }
}
function ajax(url, cb) {
  var x = new XMLHttpRequest();
  x.open('GET', url, true);
  x.onreadystatechange = function() {
    if (x.readyState === 4) {
      if (x.status === 200 && x.responseText.length > 0) {
        try { cb(JSON.parse(x.responseText)); } catch(e) { cb(null); }
      } else { cb(null); }
    }
  };
  x.send();
}
function ajaxPost(url, body, cb) {
  var x = new XMLHttpRequest();
  x.open('POST', url, true);
  x.setRequestHeader('Content-Type', 'application/json');
  x.onreadystatechange = function() {
    if (x.readyState === 4) {
      if (x.status === 200 && x.responseText.length > 0) {
        try { cb(JSON.parse(x.responseText)); } catch(e) { cb(null); }
      } else { cb(null); }
    }
  };
  x.send(body);
}
var ddCnt = 0;
function togDD(id, cnt) {
  var el = document.getElementById(id);
  var btn = document.getElementById('btn-' + id);
  if (el.classList.contains('show')) {
    el.classList.remove('show');
    btn.classList.remove('open');
    btn.innerHTML = svgDDChev + ' Show (' + cnt + ')';
  } else {
    el.classList.add('show');
    btn.classList.add('open');
    btn.innerHTML = svgDDChev + ' Hide (' + cnt + ')';
  }
}
