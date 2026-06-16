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

function captureEntityFilter(key) {
  var sel = document.getElementById('dateFilter_' + key);
  activeFilterValue[key] = resolveDate(getSelDateVal(sel));
  activeFilterLabel[key] = sel ? getSelectLabel(sel) : 'All data';
  currentAnalysisEntity = key;
}

function getDateFilterParam() {
  var df = activeFilterValue[currentAnalysisEntity] || '';
  if (!df) return '';
  return String.fromCharCode(38) + 'dateFilter=' + df;
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
      rows.push([r.ruleId, j, w.condField, w.condOp, w.condVal, w.reqField, w.reqOp].join('~'));
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
  var pill = scanPillHtml();
  if (!df && !pill) return '';
  var left = df
    ? '<span class="fn-icon">&#9202;</span><span>Filtered: <strong>' + lbl + '</strong> (since ' + df + ')</span>'
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
  var globalSel = document.getElementById('aaDateFilter');
  if (globalSel && globalSel.value === 'custom') {
    var gdi = globalSel.parentNode.querySelector('.custom-date-input');
    if (!gdi || !gdi.value) {
      if (gdi) { gdi.style.borderColor = '#c62828'; gdi.focus(); }
      return;
    }
  }
  var globalVal = getSelDateVal(globalSel);
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
  if (setupDate && setupDate.value === 'custom') {
    var di = setupDate.parentNode.querySelector('.custom-date-input');
    if (!di || !di.value) {
      if (di) { di.style.borderColor = '#c62828'; di.focus(); }
      return;
    }
  }
  var setupVal = getSelDateVal(setupDate);
  if (setupDate) {
    var globalSel = document.getElementById('aaDateFilter');
    if (globalSel) {
      if (setupDate.value === 'custom' && setupVal) {
        globalSel.value = 'custom';
        handleDateSelect(globalSel);
        var gdi = globalSel.parentNode.querySelector('.custom-date-input');
        if (gdi) gdi.value = setupVal;
      } else {
        globalSel.value = setupDate.value;
      }
    }
  }
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
  if (tabContent) tabContent.style.display = 'none';
  if (setupTabs) setupTabs.style.display = 'none';
  if (loadingDiv) loadingDiv.style.display = '';
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
  setTimeout(function() {
    var overlay = document.getElementById('setupScreen');
    overlay.classList.add('hiding');
    document.body.classList.remove('sidebar-hidden');
    if (aaQueue.length > 0) {
      var navItems = document.querySelectorAll('.nav-item');
      for (var n = 0; n < navItems.length; n++) {
        var oc = navItems[n].getAttribute('onclick') || '';
        if (oc.indexOf("'" + aaQueue[0] + "'") > -1) {
          sw(aaQueue[0], navItems[n]);
          break;
        }
      }
    }
    setTimeout(function() {
      overlay.classList.add('hidden');
    }, 500);
  }, 600);
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
// Verdict word for a score, thresholds aligned 1:1 with slColor (70/40/15)
function slBand(pct) {
  if (pct >= 70) return { w: 'Good', c: 'var(--sl-good)', cls: 'sb-band-good' };
  if (pct >= 40) return { w: 'Moderate', c: 'var(--sl-ok)', cls: 'sb-band-moderate' };
  if (pct >= 15) return { w: 'Needs attention', c: 'var(--sl-warn)', cls: 'sb-band-warn' };
  return { w: 'Critical', c: 'var(--sl-bad)', cls: 'sb-band-bad' };
}
// Inline info icon with a hover tooltip (no icon font dependency)
var SB_INFO_SVG = '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="9"/><line x1="12" y1="11" x2="12" y2="16"/><line x1="12" y1="7.5" x2="12.01" y2="7.5"/></svg>';
function sbInfoIcon(text) {
  return '<span class="sb-iw">' + SB_INFO_SVG + '<span class="sb-itip">' + text + '</span></span>';
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
