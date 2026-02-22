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

function swSettings(panel, el) {
  var panels = document.querySelectorAll('.settings-panel');
  for (var i = 0; i < panels.length; i++) {
    panels[i].classList.remove('active');
  }
  var target = document.getElementById('settings-' + panel);
  if (target) target.classList.add('active');
  var tabs = document.querySelectorAll('.settings-tab');
  for (var j = 0; j < tabs.length; j++) {
    tabs[j].classList.remove('active');
  }
  if (el) el.classList.add('active');
}

// === DATE FILTER (per-entity) ===
var activeFilterValue = {};
var activeFilterLabel = {};
var currentAnalysisEntity = '';

function resolveDate(val) {
  if (!val) return '';
  if (val === 'custom') return '';
  var now = new Date();
  if (val === 'thisyear') { return now.getFullYear() + '-01-01'; }
  if (val === 'last6m') { now.setMonth(now.getMonth() - 6); return now.toISOString().split('T')[0]; }
  if (val === 'last12m') { now.setMonth(now.getMonth() - 12); return now.toISOString().split('T')[0]; }
  if (val === 'last24m') { now.setMonth(now.getMonth() - 24); return now.toISOString().split('T')[0]; }
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

function dateFilterNotice() {
  var key = currentAnalysisEntity;
  var df = activeFilterValue[key];
  var lbl = activeFilterLabel[key];
  if (!df) return '';
  return '<div class="filter-notice"><span class="fn-icon">&#9202;</span> Filtered: <strong>' + lbl + '</strong> (since ' + df + ')</div>';
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
  for (var j = 0; j < aaQueue.length; j++) {
    listHtml += '<div class="aa-progress-item" id="aaProg_' + aaQueue[j] + '">';
    listHtml += '<span>' + aaEntityNames[aaQueue[j]] + '</span>';
    listHtml += '<span class="aa-status aa-status-pending" id="aaSt_' + aaQueue[j] + '">Pending</span>';
    listHtml += '</div>';
  }
  document.getElementById('aaEntityList').innerHTML = listHtml;
  // Mirror entity list to setup screen (with different IDs)
  var setupList = document.getElementById('setupEntityList');
  if (setupList) {
    var setupHtml = '';
    for (var j = 0; j < aaQueue.length; j++) {
      setupHtml += '<div class="aa-progress-item">';
      setupHtml += '<span>' + aaEntityNames[aaQueue[j]] + '</span>';
      setupHtml += '<span class="aa-status aa-status-pending" id="setupSt_' + aaQueue[j] + '">Pending</span>';
      setupHtml += '</div>';
    }
    setupList.innerHTML = setupHtml;
  }

  // Initialize progress bar and parallel counters
  var aaBar = document.getElementById('aaProgressBar');
  if (aaBar) { aaBar.style.width = '0'; aaBar.classList.add('loading'); }
  var aaPctEl = document.getElementById('aaProgressPercent');
  if (aaPctEl) aaPctEl.textContent = '0' + P;
  _aaCurrentPct = 0;
  _aaCompleted = 0;
  _aaRunning = 0;
  _aaStatusTexts = {};

  aaIdx = 0;
  runNextAA();
}

// Estimated relative weights per entity (based on actual timings)
var aaWeights = {
  company: 35, contact: 24, activities: 28, sale: 10,
  project: 2, requests: 3, selection: 2, marketing: 2
};

var _aaFakeTimer = null;
var _aaCurrentPct = 0;

function _aaSetProgress(pct, status) {
  // Never go backward
  if (pct < _aaCurrentPct) pct = _aaCurrentPct;
  _aaCurrentPct = pct;
  var bar = document.getElementById('aaProgressBar');
  var pctEl = document.getElementById('aaProgressPercent');
  var statusEl = document.getElementById('aaProgressStatus');
  if (bar) bar.style.width = pct + P;
  if (pctEl) pctEl.textContent = Math.round(pct) + P;
  if (statusEl && status) statusEl.textContent = status;
  setupProgressUpdate(Math.round(pct), status);
}

function _aaStartSmooth(targetPct, status) {
  _aaStopSmooth();
  _aaSetProgress(_aaCurrentPct, status);
  var maxPct = targetPct - 0.5;
  _aaFakeTimer = setInterval(function() {
    if (_aaCurrentPct < maxPct) {
      var remaining = maxPct - _aaCurrentPct;
      var increment = Math.max(0.15, remaining * 0.04);
      _aaCurrentPct = Math.min(_aaCurrentPct + increment, maxPct);
      var bar = document.getElementById('aaProgressBar');
      var pctEl = document.getElementById('aaProgressPercent');
      if (bar) bar.style.width = _aaCurrentPct + P;
      if (pctEl) pctEl.textContent = Math.round(_aaCurrentPct) + P;
      setupProgressUpdate(Math.round(_aaCurrentPct), '');
    }
  }, 120);
}

function _aaStopSmooth() {
  if (_aaFakeTimer) { clearInterval(_aaFakeTimer); _aaFakeTimer = null; }
}

// Calculate cumulative percentage ranges per entity
function _aaCalcRanges() {
  var totalWeight = 0;
  for (var i = 0; i < aaQueue.length; i++) totalWeight += (aaWeights[aaQueue[i]] || 5);
  var ranges = {};
  var cumulative = 0;
  for (var i = 0; i < aaQueue.length; i++) {
    var w = aaWeights[aaQueue[i]] || 5;
    var startPct = (cumulative / totalWeight) * 100;
    cumulative += w;
    var endPct = (cumulative / totalWeight) * 100;
    ranges[aaQueue[i]] = { start: startPct, end: endPct };
  }
  return ranges;
}

var _aaRanges = {};
var _aaConcurrency = 3; // run 3 entities at once
var _aaCompleted = 0;
var _aaRunning = 0;
var _aaStatusTexts = {}; // track status per running entity

function _aaUpdateCombinedStatus() {
  var parts = [];
  for (var k in _aaStatusTexts) { parts.push(_aaStatusTexts[k]); }
  var combined = parts.join('  |  ');
  var bar = document.getElementById('aaProgressBar');
  var pctEl = document.getElementById('aaProgressPercent');
  var statusEl = document.getElementById('aaProgressStatus');
  var pct = _aaCurrentPct;
  if (bar) bar.style.width = pct + P;
  if (pctEl) pctEl.textContent = Math.round(pct) + P;
  if (statusEl && combined) statusEl.textContent = combined;
  setupProgressUpdate(Math.round(pct), combined);
}

function _aaOnAllDone() {
  _aaStopSmooth();
  document.getElementById('aaProgressBar').classList.remove('loading');
  _aaSetProgress(100, 'All analyses complete!');
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
    document.getElementById('aaDoneSummary').textContent = aaQueue.length + ' entities analyzed.';
    if (setupAnalysisRunning) { setupAnalysisComplete(); return; }
    document.getElementById('aaStartBtn').innerHTML = '<img src="data:image/svg+xml;base64,' + icoPlayO + '"> Run Again';
  }, 800);
}

function runNextAA() {
  // Calculate ranges on first call
  if (_aaCompleted === 0 && _aaRunning === 0) {
    _aaRanges = _aaCalcRanges();
    _aaCurrentPct = 0;
    _aaStatusTexts = {};
  }

  // Launch entities up to concurrency limit
  while (_aaRunning < _aaConcurrency && aaIdx < aaQueue.length) {
    _aaLaunchEntity(aaQueue[aaIdx]);
    aaIdx++;
  }

  // Start smooth animation toward current progress + running entities' weight
  var totalWeight = 0;
  var potentialWeight = 0;
  for (var qi = 0; qi < aaQueue.length; qi++) {
    var qk = aaQueue[qi];
    var w = aaWeights[qk] || 5;
    totalWeight += w;
    var qStEl = document.getElementById('aaSt_' + qk);
    if (qStEl && (qStEl.textContent === 'Done' || qStEl.textContent === 'Loading...')) {
      potentialWeight += w;
    }
  }
  var targetPct = (potentialWeight / totalWeight) * 100;
  if (targetPct > _aaCurrentPct) _aaStartSmooth(targetPct, '');
}

function _aaLaunchEntity(key) {
  _aaRunning++;
  var range = _aaRanges[key];
  var entityName = aaEntityNames[key];

  var stEl = document.getElementById('aaSt_' + key);
  var stEl2 = document.getElementById('setupSt_' + key);
  function setEntityStatus(cls, txt) {
    if (stEl) { stEl.className = cls; stEl.textContent = txt; }
    if (stEl2) { stEl2.className = cls; stEl2.textContent = txt; }
  }
  setEntityStatus('aa-status aa-status-loading', 'Loading...');

  // Capture filter params synchronously BEFORE any async calls
  captureEntityFilter(key);

  // Track which sub-loads are still running for this entity
  var pending = {};
  var ec = entityConfig[key];
  if (ec) {
    pending['overview'] = 'overview';
    if (ec.udefId > 0) pending['fields'] = 'extra fields';
    if (ec.hasTicketFields) pending['tfields'] = 'ticket fields';
    if (key === 'company') { pending['details'] = 'details'; }
    pending['tables'] = 'tables';
  } else {
    pending['overview'] = 'overview';
  }

  function getStatusText() {
    var parts = [];
    for (var k in pending) { parts.push(pending[k]); }
    if (parts.length === 0) return entityName;
    return entityName + ' — ' + parts.join(', ');
  }

  _aaStatusTexts[key] = getStatusText();
  _aaUpdateCombinedStatus();

  function markStepDone(stepId) {
    delete pending[stepId];
    _aaStatusTexts[key] = getStatusText();
    // Bump progress based on overall completed weight
    _aaUpdateCombinedStatus();
  }

  function onEntityDone() {
    _aaRunning--;
    _aaCompleted++;
    delete _aaStatusTexts[key];
    setEntityStatus('aa-status aa-status-done', 'Done');

    // Update progress: percentage = completed weight / total weight
    var totalWeight = 0;
    var doneWeight = 0;
    for (var qi = 0; qi < aaQueue.length; qi++) {
      var qk = aaQueue[qi];
      var w = aaWeights[qk] || 5;
      totalWeight += w;
      var qStEl = document.getElementById('aaSt_' + qk);
      if (qStEl && qStEl.textContent === 'Done') {
        doneWeight += w;
      }
    }
    var donePct = (doneWeight / totalWeight) * 100;
    _aaStopSmooth();
    _aaSetProgress(donePct, _aaCompleted >= aaQueue.length ? 'All analyses complete!' : '');

    if (_aaCompleted >= aaQueue.length) {
      _aaOnAllDone();
    } else {
      runNextAA(); // launch next available entity
    }
  }

  // Launch the entity (all URL construction happens synchronously here)
  if (entityConfig[key]) {
    aaRunFullEntity(key, onEntityDone, markStepDone);
  } else {
    aaRunSimpleEntity(key, onEntityDone);
  }
}

function aaRunSimpleEntity(key, cb) {
  ajax(overviewUrl + String.fromCharCode(38) + 'entity=' + key + getDateFilterParam(), function(d) {
    if (d) renderEntityOverview(key, d);
    var rs = document.getElementById(key + 'Results');
    var ab = document.getElementById(key + 'AnalyzeBtn');
    var eb = document.getElementById(key + 'ExportBtn');
    if (rs) rs.style.display = '';
    if (ab) ab.textContent = 'Re-analyze';
    if (eb) eb.style.display = '';
    if (cb) cb();
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
  if (hasDetails) totalSteps++; // details (includes funnel now)
  totalSteps++; // extra tables always

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
      if (cb) cb();
    }
  }

  // === PARALLEL 1: Overview ===
  ajax(overviewUrl + String.fromCharCode(38) + 'entity=' + key + dfParam, function(d) {
    if (d) renderEntityOverview(key, d);
    stepDone('overview');
  });

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

  // === PARALLEL 3: Company Details + funnel (company only) ===
  if (hasDetails) {
    loadCompanyDetails(function() {
      // Funnel is now included in detail response — render it
      if (typeof loadCompanyCross === 'function') loadCompanyCross(null);
      stepDone('details');
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

function setupProgressUpdate(pct, status) {
  if (!setupAnalysisRunning) return;
  var bar = document.getElementById('setupProgressBar');
  var pctEl = document.getElementById('setupProgressPercent');
  var statusEl = document.getElementById('setupProgressStatus');
  if (bar) bar.style.width = pct + '%';
  if (pctEl) pctEl.textContent = pct + '%';
  if (statusEl && status) statusEl.textContent = status;
}

function setupAnalysisComplete() {
  setupAnalysisRunning = false;
  var pctEl = document.getElementById('setupProgressPercent');
  var statusEl = document.getElementById('setupProgressStatus');
  var bar = document.getElementById('setupProgressBar');
  if (pctEl) pctEl.textContent = '100%';
  if (bar) bar.style.width = '100%';
  if (statusEl) statusEl.textContent = 'Complete!';
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
