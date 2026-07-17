// === UDEF RENDER ===
function renderUdef(ent, idx, entityKey) {
  var h = '<div class="entity-card">';
  h += '<div class="entity-header"><div class="entity-info"><h3>' + ent.name + '</h3> <span class="field-count">\u2014 ' + ent.fields.length + ' active fields</span></div>';
  h += '<span class="record-badge">' + fmtNum(ent.total) + ' records</span></div>';
  if (ent.fields.length === 0) {
    h += '<div style="padding:20px;color:#888"><em>No extra fields configured</em></div>';
  } else {
    var udefCfg = (typeof daSettings !== 'undefined' && entityKey) ? (daSettings.getSettings(entityKey).udefFieldConfig || {}) : {};
    var Q = String.fromCharCode(39);
    var tid = 'tbl-u' + idx;
    // When a period filter is active, UDEF coverage is measured over records
    // CREATED in the period (Option A). The standard fields use the active-in-period
    // population, so the two denominators differ by design. Make that explicit here.
    var udefScopeNote = '';
    if (typeof activeScope !== 'undefined' && activeScope !== 'all' && ent.total !== undefined) {
      udefScopeNote = ' Coverage is measured over the ' + fmtNum(ent.total) + ' records created in the selected period, which can differ from the record counts on the other tabs.';
    }
    h += '<div class="tbl-cap"><b>Higher completeness is better.</b> Custom fields feed the Data Quality score. Sorted by attention, so important fields with the largest gaps sit on top.' + udefScopeNote + '</div>';
    h += '<table class="data-table" id="' + tid + '">';
    h += '<thead><tr>';
    h += '<th class="attn-col" onclick="sortT(' + Q + tid + Q + ',0)"><span class="sort-arrow active">' + svgSortD + '</span></th>';
    h += '<th onclick="sortT(' + Q + tid + Q + ',1)">Field Label <span class="sort-arrow">' + svgSortN + '</span></th>';
    h += '<th onclick="sortT(' + Q + tid + Q + ',2)">Type <span class="sort-arrow">' + svgSortN + '</span></th>';
    h += '<th class="col-right" style="width:110px" onclick="sortT(' + Q + tid + Q + ',3)">Filled <span class="sort-arrow">' + svgSortN + '</span></th>';
    h += '<th style="text-align:center;width:90px">Importance</th>';
    h += '<th class="col-right" style="width:130px" onclick="sortT(' + Q + tid + Q + ',5)">Completeness <span class="sort-arrow">' + svgSortN + '</span></th>';
    h += '</tr></thead><tbody>';
    var uRows = [];
    for (var fi = 0; fi < ent.fields.length; fi++) {
      var f = ent.fields[fi];
      var pid = f.progId || f.label || ('f' + fi);
      var imp = udefCfg[pid] || 'normal';
      var rc = '';
      if (f.filled === 0) rc = 'unused';
      var uBad = 100 - f.percent;
      var uAttn = attnScore(uBad, imp);
      var uReason = '<b>' + attnBandWord(uAttn) + ' attention.</b> ' + _capFirst(imp) + ' field, ' + Math.round(uBad) + '% empty. Higher completeness lifts the Data Quality score.';
      var dimStyle = imp === 'excluded' ? ' style="opacity:0.4"' : '';
      var tr = '<tr class="' + rc + '"' + dimStyle + '>';
      tr += '<td class="attn-col" data-sort-value="' + uAttn + '">' + attnIconHtml(uAttn, uReason) + '</td>';
      tr += '<td data-sort-value="' + f.label + '">' + f.label + '</td>';
      tr += '<td data-sort-value="' + f.type + '"><span class="type-badge">' + f.type + '</span>';
      if (f.items && f.items.length > 0) {
        ddCnt++;
        var did = 'dd-' + ddCnt;
        tr += ' <span class="dd-toggle" id="btn-' + did + '" onclick="togDD(' + Q + did + Q + ',' + f.items.length + ')">';
        tr += svgDDChev + ' Show (' + f.items.length + ')</span>';
        tr += '<div class="dd-panel" id="' + did + '">';
        for (var ii = 0; ii < f.items.length; ii++) {
          var itm = f.items[ii];
          var ic = '';
          if (itm.c === 0) ic = 'unused';
          tr += '<div class="dd-item ' + ic + '">';
          tr += '<span>' + itm.n + '</span>';
          tr += '<span>' + itm.c + ' <span class="dd-pct">' + itm.p.toFixed(1) + P + '</span></span>';
          tr += '</div>';
        }
        tr += '</div>';
      }
      tr += '</td>';
      tr += '<td class="col-right" data-sort-value="' + f.filled + '">' + f.filled + ' / ' + ent.total + '</td>';
      tr += '<td style="text-align:center"><span class="imp-badge ' + imp + '">' + imp + '</span></td>';
      tr += '<td class="col-right" data-sort-value="' + f.percent + '">' + barCell(f.percent, '') + '</td>';
      tr += '</tr>';
      uRows.push({ attn: uAttn, badness: uBad, html: tr });
    }
    h += attnSortRows(uRows);
    h += '</tbody></table>';
  }
  h += '</div>';
  return h;
}

// Re-render the UDEF field tables for every entity whose data is already loaded.
// Called after Apply & Recalculate so importance changes (e.g. fields set to Off)
// are reflected in the field lists, not only in the scores.
function reRenderLoadedUdef() {
  for (var k in entityConfig) {
    if (!entityConfig.hasOwnProperty(k)) continue;
    var ec = entityConfig[k];
    if (!ec || !ec.udefId || ec.udefId <= 0) continue;
    var d = udefData[ec.udefId];
    if (!d || !d.fields) continue;
    var cardsEl = document.getElementById(k + 'UdefCards');
    if (cardsEl) cardsEl.innerHTML = renderUdef(d, ec.udefIdx, k);
  }
}

// === UDEF LOADING ===
var udefData = {};
var overviewData = {}; // Store overview data per entity key for export

function loadEntityUdef(entityId, entityKey, udefIdx, callback) {
  var startEl = document.getElementById(entityKey + 'UdefStart');
  var progressEl = document.getElementById(entityKey + 'UdefProgress');
  var barEl = document.getElementById(entityKey + 'UdefBar');
  var statusEl = document.getElementById(entityKey + 'UdefStatus');
  var resultsEl = document.getElementById(entityKey + 'UdefResults');
  var cardsEl = document.getElementById(entityKey + 'UdefCards');
  var summaryEl = document.getElementById(entityKey + 'UdefSummary');

  if (startEl) {
    var btns = startEl.querySelectorAll('.btn-analyze');
    for (var i = 0; i < btns.length; i++) btns[i].style.display = 'none';
  }
  if (progressEl) progressEl.style.display = 'block';
  if (barEl) barEl.style.width = '30' + P;
  if (statusEl) statusEl.textContent = 'Loading extra fields...';

  ajax(udefUrl + String.fromCharCode(38) + 'entityId=' + entityId + getDateFilterParam(), function(d) {
    udefData[entityId] = d;
    if (typeof daSettings !== 'undefined' && daSettings.notifyUdefLoaded) daSettings.notifyUdefLoaded(entityId, d);
    if (barEl) barEl.style.width = '100' + P;
    if (statusEl) statusEl.textContent = 'Complete!';
    if (startEl) startEl.style.display = 'none';
    if (resultsEl) resultsEl.style.display = 'block';
    if (d && summaryEl) {
      summaryEl.textContent = d.fields.length + ' active extra fields. Click column headers to sort.';
      if (cardsEl) cardsEl.innerHTML = renderUdef(d, udefIdx, entityKey);
    } else if (summaryEl) {
      summaryEl.textContent = 'No extra fields found.';
    }
    if (callback) callback();
  });
}

function loadEntityUdefQuiet(entityId, entityKey, udefIdx, callback) {
  var cardsEl = document.getElementById(entityKey + 'UdefCards');
  var summaryEl = document.getElementById(entityKey + 'UdefSummary');

  ajax(udefUrl + String.fromCharCode(38) + 'entityId=' + entityId + getDateFilterParam(), function(d) {
    udefData[entityId] = d;
    if (typeof daSettings !== 'undefined' && daSettings.notifyUdefLoaded) daSettings.notifyUdefLoaded(entityId, d);
    if (d && summaryEl) {
      summaryEl.textContent = d.fields.length + ' active extra fields. Click column headers to sort.';
      if (cardsEl) cardsEl.innerHTML = renderUdef(d, udefIdx, entityKey);
    } else if (summaryEl) {
      summaryEl.textContent = 'No extra fields found.';
    }
    if (callback) callback();
  });
}

// === ENTITY OVERVIEW ===
var ovLabels = {
  company: {
    title: 'Company Overview',
    stats: [ ['total','Total Companies'],['withPersons','With Persons'],['withActivities','With Activities'],['withTickets','With Tickets'] ],
    completeness: [ ['orgNr','Org. Number'],['email','Email'],['phone','Phone'],['address','Address'],['webpage','Webpage'] ]
  },
  contact: {
    title: 'Contact Overview',
    stats: [ ['total','Total Contacts'],['withEmail','With Email'],['withPhone','With Phone'],['withPosition','With Position'],['withTitle','With Title'],['withSales','With Sales'],['inProjects','In Projects'],['withActivities','With Activities'] ]
  },
  sale: {
    title: 'Sale Overview',
    stats: [ ['total','Total Sales'],['withPersons','With Persons'],['withCompanies','With Companies'],['withProjects','With Projects'],['withActivities','With Activities'],['withQuote','With Quote'],['withStakeholders','With Stakeholders'] ]
  },
  project: {
    title: 'Project Overview',
    stats: [ ['total','Total Projects'],['withActivities','With Activities'],['withMembers','With Members'],['overdue','Overdue'] ]
  },
  requests: {
    title: 'Ticket Overview',
    stats: [ ['total','Total Tickets'],['open','Open'],['closed','Closed'],['postponed','Postponed'] ]
  },
  activities: {
    title: 'Activities Overview',
    stats: [ ['totalActivities','Total Activities'],['actWithCompanies','With Companies'],['actWithPersons','With Persons'],['actWithSales','With Sales'],['actWithProjects','With Projects'] ],
    sections: [ {
      title: 'Document Overview',
      totalKey: 'totalDocuments',
      stats: [ ['totalDocuments','Total Documents'],['docWithCompanies','With Companies'],['docWithPersons','With Persons'],['docWithSales','With Sales'],['docWithProjects','With Projects'] ]
    } ]
  },
  selection: {
    title: 'Selection Overview',
    stats: [ ['total','Total Selections'],['staticSel','Static'],['dynamicSel','Dynamic'],['combinedSel','Combined'],['runLast12Months','Run Last 12 Months'] ]
  },
  marketing: {
    title: 'Mailings',
    stats: [ ['totalMailings','Total Mailings'],['sentMailings','Sent'] ],
    sections: [ {
      title: 'Recipients',
      totalKey: 'totalRecipients',
      stats: [ ['totalRecipients','Total Recipients'] ]
    }, {
      title: 'Forms',
      totalKey: 'totalForms',
      stats: [ ['totalForms','Total Forms'],['activeForms','Active Forms'],['totalSubmissions','Form Submissions',true] ]
    } ]
  }
};

// === ENTITY LOADING ===
var entityConfig = {
  company:  { udefId: 7,  udefIdx: 100, extraType: 'Company', tabKey: 'company' },
  contact:  { udefId: 8,  udefIdx: 200, extraType: 'Person',  tabKey: 'contact' },
  sale:     { udefId: 10, udefIdx: 300, extraType: 'Sale',    tabKey: 'sale' },
  project:  { udefId: 9,  udefIdx: 400, extraType: 'Project', tabKey: 'project' },
  requests: { udefId: 0,  udefIdx: 0,   extraType: 'Ticket',  tabKey: 'requests', hasTicketFields: true }
};

// ===========================================================
// EXTRA TABLES GLOBAL CACHE — load once, reuse across entities
// ===========================================================
var _extraCache = null;        // { tables:[], data:{id:responseData}, ready:false }
var _extraCacheQueue = null;   // array of callbacks waiting, or null if not loading

function getExtraTablesFromCache(callback) {
  // If already cached, return immediately
  if (_extraCache && _extraCache.ready) {
    callback(_extraCache);
    return;
  }
  // If currently loading, queue the callback
  if (_extraCacheQueue) {
    _extraCacheQueue.push(callback);
    return;
  }
  // Start loading
  _extraCacheQueue = [callback];
  ajax(extraUrl, function(d) {
    if (!d || !d.tables || d.tables.length === 0) {
      _extraCache = { tables: [], data: {}, ready: true };
      _flushExtraCacheQueue();
      return;
    }
    // Build entity counts param from listing response (v4)
    var ecParam = '';
    if (d.entityCounts) {
      var parts = [];
      for (var k in d.entityCounts) {
        if (d.entityCounts.hasOwnProperty(k)) {
          parts.push(k + ':' + d.entityCounts[k]);
        }
      }
      ecParam = String.fromCharCode(38) + 'ec=' + parts.join(',');
    }
    _extraCache = { tables: d.tables, data: {}, ready: false };
    var remaining = d.tables.length;
    for (var i = 0; i < d.tables.length; i++) {
      (function(tbl) {
        ajax(extraUrl + String.fromCharCode(38) + 'tableId=' + tbl.id + ecParam, function(td) {
          _extraCache.data[tbl.id] = td;
          remaining--;
          if (remaining <= 0) {
            _extraCache.ready = true;
            _flushExtraCacheQueue();
          }
        });
      })(d.tables[i]);
    }
  });
}

function _flushExtraCacheQueue() {
  var cbs = _extraCacheQueue;
  _extraCacheQueue = null;
  if (cbs) {
    for (var i = 0; i < cbs.length; i++) cbs[i](_extraCache);
  }
}

function invalidateExtraCache() {
  _extraCache = null;
  _extraCacheQueue = null;
}

// ===========================================================
// LOADING PLACEHOLDER HELPERS
// ===========================================================
function loadingPlaceholder(label) {
  return '<div style="display:flex;align-items:center;gap:12px;padding:32px 20px;color:#999">' +
    '<div style="width:20px;height:20px;border:2.5px solid #e0dfdc;border-top-color:var(--so-green);border-radius:50%;animation:spin .8s linear infinite"></div>' +
    '<span style="font-size:.85rem">' + label + '</span></div>';
}

// Spinner/fadeIn keyframes moved to da-styles.css

// ===========================================================
// PROGRESSIVE startFullEntity — parallel loads, show as ready
// ===========================================================
function startFullEntity(key) {
  captureEntityFilter(key);
  var ec = entityConfig[key];
  var tabKey = ec ? ec.tabKey : key;
  var hasDetails = (key === 'company');

  var progressScreen = document.getElementById(tabKey + 'ProgressScreen');
  var progressBar = document.getElementById(tabKey + 'ProgressBar');
  var progressPercent = document.getElementById(tabKey + 'ProgressPercent');
  var progressStatus = document.getElementById(tabKey + 'ProgressStatus');
  var subTabs = document.getElementById(tabKey + 'SubTabs');
  var resultsContainer = document.getElementById(tabKey + 'Results');
  var headerBtn = document.getElementById(tabKey + 'AnalyzeBtn');

  if (headerBtn) headerBtn.disabled = true;

  // === Show results container but with loading overlay ===
  if (progressScreen) progressScreen.style.display = 'none';
  if (subTabs) subTabs.style.display = '';
  if (resultsContainer) resultsContainer.style.display = '';

  // Clear all section content
  var overviewEl = document.getElementById(tabKey + 'OverviewContent');
  if (overviewEl) overviewEl.innerHTML = '';
  var dqScoreEl = document.getElementById(tabKey + 'DQScoreContent');
  if (dqScoreEl) dqScoreEl.innerHTML = '';
  if (ec && ec.udefId > 0) {
    var udefCards = document.getElementById(tabKey + 'UdefCards');
    var udefSummary = document.getElementById(tabKey + 'UdefSummary');
    if (udefCards) udefCards.innerHTML = '';
    if (udefSummary) udefSummary.textContent = '';
  }
  var extraCards = document.getElementById(tabKey + 'ExtraCards');
  var extraSummary = document.getElementById(tabKey + 'ExtraSummary');
  if (extraCards) extraCards.innerHTML = '';
  if (extraSummary) extraSummary.textContent = '';
  if (hasDetails) {
    var crossEl = document.getElementById('companyCrossContent');
    if (crossEl) crossEl.innerHTML = '';
    var detailEl = document.getElementById('companyDetailContent');
    if (detailEl) detailEl.innerHTML = '';
  }

  // === Create unified loading overlay ===
  var existingOverlay = document.getElementById(tabKey + 'LoadingOverlay');
  if (existingOverlay) existingOverlay.parentNode.removeChild(existingOverlay);
  var overlay = document.createElement('div');
  overlay.id = tabKey + 'LoadingOverlay';
  overlay.className = 'loading-overlay';
  overlay.innerHTML = '<div class="loading-overlay-inner">' +
    '<div class="loading-spinner"></div>' +
    '<div class="loading-title" id="' + tabKey + 'LoadTitle">Analyzing...</div>' +
    '<div class="loading-progress-wrap"><div class="loading-progress-bar" id="' + tabKey + 'LoadBar"></div></div>' +
    '<div class="loading-status" id="' + tabKey + 'LoadStatus">Starting analysis</div>' +
    '</div>';
  if (resultsContainer) {
    resultsContainer.style.position = 'relative';
    resultsContainer.appendChild(overlay);
  }

  // === Track completion of parallel loads ===
  var totalSteps = 2; // overview + extra (always)
  if (ec && ec.udefId > 0) totalSteps++;
  if (ec && ec.hasTicketFields) totalSteps++;
  if (hasDetails) totalSteps += 3; // core + quality + detail (3 parallel scripts)
  var completedSteps = 0;
  var dfParam = getDateFilterParam();

  function stepDone(label) {
    completedSteps++;
    var pct = Math.round((completedSteps / totalSteps) * 100);
    var barEl = document.getElementById(tabKey + 'LoadBar');
    var statusEl = document.getElementById(tabKey + 'LoadStatus');
    if (barEl) barEl.style.width = pct + P;
    if (statusEl) statusEl.textContent = label + ' (' + pct + P + ')';

    // Progressively update DQ score as data becomes available
    renderDQScore(key);

    if (completedSteps >= totalSteps) {
      // All done — remove overlay with fade
      var expBtn = document.getElementById(tabKey + 'ExportBtn') || document.getElementById('ticketExportBtn');
      if (expBtn) expBtn.style.display = '';
      if (headerBtn) { headerBtn.disabled = false; headerBtn.onclick = function(){ reAnalyze(key); }; }
      renderDQScore(key);
      renderScoreBanner(key);
      renderAdoptionTab(key);
      var ov = document.getElementById(tabKey + 'LoadingOverlay');
      if (ov) {
        ov.style.opacity = '0';
        ov.style.transition = 'opacity .4s ease';
        setTimeout(function() { if (ov.parentNode) ov.parentNode.removeChild(ov); }, 400);
      }
    }
  }

  // === PARALLEL LOAD 1: Overview ===
  ajax(overviewUrl + String.fromCharCode(38) + 'entity=' + key + dfParam + getExclParam(key), function(d) {
    if (d) renderEntityOverview(key, d);
    if (overviewEl) overviewEl.style.animation = 'fadeIn .3s ease';
    if (key === 'company') populateDetailCatFilter();
    stepDone('Overview loaded');
  });

  // === PARALLEL LOAD 2: UDEF / Ticket Fields ===
  if (ec && ec.udefId > 0) {
    ajax(udefUrl + String.fromCharCode(38) + 'entityId=' + ec.udefId + dfParam, function(d) {
      udefData[ec.udefId] = d;
      if (typeof daSettings !== 'undefined' && daSettings.notifyUdefLoaded) daSettings.notifyUdefLoaded(ec.udefId, d);
      var cardsEl = document.getElementById(tabKey + 'UdefCards');
      var summaryEl = document.getElementById(tabKey + 'UdefSummary');
      if (d && summaryEl) {
        summaryEl.textContent = d.fields.length + ' active extra fields. Click column headers to sort.';
        if (cardsEl) { cardsEl.innerHTML = renderUdef(d, ec.udefIdx, key); cardsEl.style.animation = 'fadeIn .3s ease'; }
      } else {
        if (summaryEl) summaryEl.textContent = 'No extra fields found.';
        if (cardsEl) cardsEl.innerHTML = '';
      }
      var udefResults = document.getElementById(tabKey + 'UdefResults');
      if (udefResults) udefResults.style.display = 'block';
      stepDone('Extra fields loaded');
    });
  } else if (ec && ec.hasTicketFields) {
    ajax(ticketUrl + dfParam, function(d) {
      ticketData = d;
      showTicketQuiet();
      stepDone('Custom fields loaded');
    });
  }

  // === PARALLEL LOAD 3: Company Core + Quality + Detail (3 scripts, parallel) ===
  if (hasDetails) {
    var catParam = '';
    if (companyDetailCatValue) {
      catParam = String.fromCharCode(38) + 'categoryName=' + encodeURIComponent(companyDetailCatValue);
    }

    // Each script renders its part independently as it arrives
    // companyDetailData is progressively built up

    // Call 1: CompanyCoreFetch (activity health, trend) — lightweight, ~6s
    ajax(coreUrl + dfParam + catParam, function(d) {
      if (!companyDetailData) companyDetailData = {};
      if (d) {
        if (d.activityHealth) companyDetailData.activityHealth = d.activityHealth;
        if (d.trend) companyDetailData.trend = d.trend;
        if (d.trendMonthly) companyDetailData.trendMonthly = d.trendMonthly;
        if (d.trendBefore !== undefined) companyDetailData.trendBefore = d.trendBefore;
      }
      // Re-render details with whatever data we have so far
      renderCompanyDetails(companyDetailData);
      renderCompanyOverviewV2();
      stepDone('Activity health loaded');
    });

    // Call 2: CompanyQualityFetch (quality, funnel, segments, churn) — countRows only, ~5s
    ajax(qualityUrl + dfParam + catParam + getExclParam('company'), function(d) {
      if (!companyDetailData) companyDetailData = {};
      if (d) {
        if (d.quality) companyDetailData.quality = d.quality;
        if (d.funnel) companyDetailData.funnel = d.funnel;
        if (d.churnRisk) companyDetailData.churnRisk = d.churnRisk;
      }
      // Render funnel (creates category filter)
      if (companyDetailData.funnel) {
        companyCrossData = companyDetailData;
        renderCrossEntityFunnel(companyDetailData);
        var crossEl2 = document.getElementById('companyCrossContent');
        if (crossEl2) crossEl2.style.animation = 'fadeIn .3s ease';
      }
      populateDetailCatFilter();
      var countEl = document.getElementById('companyDetailFilterCount');
      if (countEl && companyDetailData.activityHealth) {
        countEl.textContent = companyDetailCatValue ? fmtNum(companyDetailData.activityHealth.total) + ' companies' : '';
      }
      // Re-render details
      renderCompanyDetails(companyDetailData);
      // Trigger DQ score recalc (quality data now available)
      renderDQScore('company');
      renderCompanyOverviewV2();
      stepDone('Quality analysis loaded');
    });

    // Call 3: CompanyDetailFetch (associates, category effectiveness) — bulk loop, ~25s
    ajax(detailUrl + dfParam + catParam, function(d) {
      if (!companyDetailData) companyDetailData = {};
      if (d) {
        if (d.associates) companyDetailData.associates = d.associates;
        if (d.categoryEffectiveness) companyDetailData.categoryEffectiveness = d.categoryEffectiveness;
      }
      // Re-render details (now with associate + category data)
      renderCompanyDetails(companyDetailData);
      var detailEl2 = document.getElementById('companyDetailContent');
      if (detailEl2) detailEl2.style.animation = 'fadeIn .3s ease';
      stepDone('Associate breakdown loaded');
    });
  }

  // === PARALLEL LOAD 4: Extra Tables (from cache) ===
  if (ec) {
    getExtraTablesFromCache(function(cache) {
      if (!entExtra[tabKey]) entExtra[tabKey] = { tables: [], cur: 0, data: {} };
      entExtra[tabKey].tables = cache.tables;
      entExtra[tabKey].entityType = ec.extraType;
      for (var i = 0; i < cache.tables.length; i++) {
        var tbl = cache.tables[i];
        entExtra[tabKey].data[tbl.id] = cache.data[tbl.id];
      }
      entExtra[tabKey].cur = cache.tables.length;
      showEntityExtraQuiet(tabKey);
      if (extraCards) extraCards.style.animation = 'fadeIn .3s ease';
      stepDone('Extra tables loaded');
    });
  } else {
    if (extraCards) extraCards.innerHTML = '';
    stepDone('Complete');
  }
}

function startCompanyOverview() { startFullEntity('company'); }

function startEntityOverview(key) {
  if (entityConfig[key]) {
    startFullEntity(key);
    return;
  }

  captureEntityFilter(key);
  var progressScreen = document.getElementById(key + 'ProgressScreen');
  var progressBar = document.getElementById(key + 'ProgressBar');
  var progressPercent = document.getElementById(key + 'ProgressPercent');
  var progressStatus = document.getElementById(key + 'ProgressStatus');
  var resultsContainer = document.getElementById(key + 'Results');
  var headerBtn = document.getElementById(key + 'AnalyzeBtn');

  if (headerBtn) headerBtn.disabled = true;
  if (progressScreen) progressScreen.style.display = '';
  if (progressBar) { progressBar.style.width = '0'; progressBar.classList.add('loading'); }
  if (progressPercent) progressPercent.textContent = '0' + P;
  if (progressStatus) progressStatus.textContent = 'Loading data...';

  var currentPct = 0;
  var fakeTimer = setInterval(function() {
    if (currentPct < 95) {
      var remaining = 95 - currentPct;
      var increment = Math.max(0.5, remaining * 0.06);
      currentPct = Math.min(currentPct + increment, 95);
      if (progressBar) progressBar.style.width = currentPct + P;
      if (progressPercent) progressPercent.textContent = Math.round(currentPct) + P;
    }
  }, 150);

  ajax(overviewUrl + String.fromCharCode(38) + 'entity=' + key + getDateFilterParam() + getExclParam(key), function(d) {
    clearInterval(fakeTimer);
    if (d) renderEntityOverview(key, d);
    if (progressBar) { progressBar.style.width = '100' + P; progressBar.classList.remove('loading'); }
    if (progressPercent) progressPercent.textContent = '100' + P;
    if (progressStatus) progressStatus.textContent = 'Complete!';

    setTimeout(function() {
      if (progressScreen) progressScreen.style.display = 'none';
      if (resultsContainer) resultsContainer.style.display = '';
      var expBtn = document.getElementById(key + 'ExportBtn');
      if (expBtn) expBtn.style.display = '';
      if (headerBtn) { headerBtn.disabled = false; headerBtn.onclick = function(){ reAnalyze(key); }; }
    }, 500);
  });
}

// Map analysis distribution titles to the standard list fields in the settings,
// so those fields can expand and have their values excluded like custom fields.
var STD_DIST_MAP = {
  company: { 'Category': 'category', 'Business': 'business' },
  sale:    { 'Sale Type': 'saleType', 'Stage': 'stage' },
  project: { 'Project Type': 'projectType', 'Project Status': 'status' }
};

function pushStdListValues(key, d) {
  if (typeof daSettings === 'undefined' || !daSettings.notifyStdListLoaded) return;
  if (!d || !d.distributions) return;
  var map = STD_DIST_MAP[key];
  if (!map) { daSettings.notifyStdListLoaded(key, {}); return; }
  var out = {};
  for (var i = 0; i < d.distributions.length; i++) {
    var dist = d.distributions[i];
    if (!dist || !dist.items) continue;
    var sk = map[dist.title];
    if (!sk) continue;
    var tot = dist.total || 0;
    if (!tot) { for (var t = 0; t < dist.items.length; t++) tot += (dist.items[t].count || 0); }
    var items = [];
    for (var j = 0; j < dist.items.length; j++) {
      var it = dist.items[j];
      var nm = it.name;
      if (nm === '(No value)') continue; // empty bucket is not a selectable value
      if (nm === '(Other)') continue; // reconciliation bucket (deleted/missing list items), not selectable
      var c = it.count || 0;
      items.push({ n: nm, c: c, p: tot > 0 ? (c / tot * 100) : 0, i: it.idx });
    }
    out[sk] = items;
  }
  daSettings.notifyStdListLoaded(key, out);
}

// =====================================================
// Company Overview v2 — state + trajectory + cause, scope-aware.
// Reads progressively from overviewData['company'] (stats, distributions) and
// companyDetailData (activityHealth incl. dbTotal, funnel.segments, trend).
// Idempotent: re-callable from every load path; each block guards on its data.
// Re-inserts the score banner at the end so re-renders keep it.
// =====================================================
// ===== Shared Overview components (float-mono redesign) — one source, all entities =====
function hcKpi(label, value, sub, dot, action) {
  var d = (typeof dot === 'string' && dot) ? '<span class="kdot" style="background:' + dot + '"></span> ' : '';
  var a = (typeof action === 'string' && action) ? '<button class="act">' + action + ' <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="M5 12h14M13 6l6 6-6 6"/></svg></button>' : '';
  return '<div class="kpi"><div class="kl">' + d + label + '</div><div class="kv">' + value + '</div><div class="ks">' + sub + '</div>' + a + '</div>';
}
// Engagement and recency segments are a distribution, not a status, so they run the
// teal ramp (prototype order c1 -> c3 -> c5 -> cn) and never borrow the semantic
// colours. Ordered checks: the label "Active, No Pipeline" also contains "active".
function hcEngColor(name) {
  var nm = (name || '').toLowerCase();
  if (nm.indexOf('fully') >= 0) return 'var(--c1)';
  if (nm.indexOf('active (6m)') >= 0) return 'var(--c1)';
  if (nm.indexOf('cooling') >= 0) return 'var(--c3)';
  if (nm.indexOf('empty') >= 0) return 'var(--cn)';
  if (nm.indexOf('no activity') >= 0) return 'var(--cn)';
  if (nm.indexOf('dormant') >= 0) return 'var(--c5)';
  return 'var(--c3)';
}

function hcSecLabel(t) { return '<div class="seclabel">' + t + '</div>'; }
function hcInsight(kind, bold, rest, buttons) {
  // kind: 'primary' (act on this) | 'warn' (worth checking) | 'neutral' (context) | 'good' (all clear)
  var ICONS = {
    primary: '<path d="M3 7h18M6 12h12M10 17h4"/>',
    warn: '<path d="M12 9v4M12 17h.01M10.3 4l-7 12a2 2 0 0 0 1.7 3h14a2 2 0 0 0 1.7-3l-7-12a2 2 0 0 0-3.4 0z"/>',
    neutral: '<circle cx="12" cy="12" r="9"/><path d="M12 11v5M12 8h.01"/>',
    good: '<path d="M20 6L9 17l-5-5"/>'
  };
  var SKIN = {
    primary: '',
    warn: 'background:var(--mod-bg);color:var(--mod-ink);border-color:var(--mod-line)',
    neutral: 'background:var(--chip-fill);color:var(--chip-ink);border-color:var(--chip-line)',
    good: 'background:var(--good-bg);color:var(--good-ink);border-color:var(--good-line)'
  };
  var TAGS = {
    primary: '<span class="tag">Recommended</span> ',
    warn: '<span class="tag warn">Worth checking</span> ',
    neutral: '',
    good: ''
  };
  if (!ICONS[kind]) kind = 'neutral';
  var sk = SKIN[kind] ? ' style="' + SKIN[kind] + '"' : '';
  var ico = '<span class="ico"' + sk + '><svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.9" stroke-linecap="round" stroke-linejoin="round">' + ICONS[kind] + '</svg></span>';
  var btns = ''; buttons = buttons || [];
  for (var i = 0; i < buttons.length; i++) btns += '<button class="btn-s ' + (buttons[i].p ? 'primary' : 'ghost') + '">' + buttons[i].l + '</button>';
  var body = TAGS[kind] + (bold ? '<b>' + bold + '</b> ' : '') + (rest || '');
  return '<div class="insight' + (kind === 'primary' ? ' primary' : '') + '">' + ico + '<span class="tx">' + body + '</span>' + (btns ? '<span class="acts">' + btns + '</span>' : '') + '</div>';
}

function renderCompanyOverviewV2() {
  var el = document.getElementById('companyOverviewContent');
  if (!el) return;

  var ov = (typeof overviewData !== 'undefined' && overviewData['company']) ? overviewData['company'] : null;
  var o = ov ? ov.overview : null;
  var dists = ov ? ov.distributions : null;
  var cd = (typeof companyDetailData !== 'undefined' && companyDetailData) ? companyDetailData : {};
  var ah = cd.activityHealth || null;
  var fn = cd.funnel || null;
  var trend = cd.trend || null;
  var trendBefore = cd.trendBefore || 0;

  var GREEN = 'var(--so-green,#06423e)';
  var GOOD = 'var(--sl-good,#2e7d32)';
  var BAD = 'var(--sl-bad,#c62828)';
  var MUTED = 'var(--so-text-muted,#6b706c)';

  var secLabel = hcSecLabel;
  var kpiCard = hcKpi;
  function dotLeg(color, label, val, pct) {
    var nums = '';
    if (val !== '') nums += '<span style="font-weight:600;color:#1c2b29;margin-left:8px">' + val + '</span>';
    if (pct !== '') nums += '<span style="color:#8a8f8b;margin-left:5px">' + pct + '</span>';
    return '<div style="display:flex;align-items:center;gap:7px;font-size:12px;color:#3c423f;margin:4px 0">' +
      '<span style="width:10px;height:10px;border-radius:3px;flex:none;background:' + color + '"></span>' +
      '<span>' + label + '</span>' + nums + '</div>';
  }

  var h = '';
  h += dateFilterNotice();
  // (score banner is inserted right after the notice by renderScoreBanner at the end)

  // ---- STATE: KPI row ----
  if (o || ah) {
    var totalC = ah ? ah.total : (o ? (o.total || o.totalCompanies || 0) : 0);
    var dbT = (ah && ah.dbTotal) ? ah.dbTotal : totalC;
    var dormantTail = (dbT && totalC) ? (dbT - totalC) : 0;
    var dormantPct = dbT > 0 ? Math.round(dormantTail / dbT * 100) : 0;
    var newThisYear = (trend && trend.length > 0) ? trend[trend.length - 1].count : 0;
    var withPersons = o ? o.withPersons : 0;
    var withPersonsPct = totalC > 0 ? Math.round(withPersons / totalC * 1000) / 10 : 0;

    h += secLabel('State \u00b7 where things stand');
    h += '<div class="kpis">';
    h += kpiCard('Companies', fmtNum(totalC), 'of ' + fmtNum(dbT) + ' total', '', '');
    h += kpiCard('Dormant tail', fmtNum(dormantTail), dormantPct + '% with no recent activity', 'var(--cn)', 'Create selection');
    h += kpiCard('New this year', fmtNum(newThisYear), 'registered this year', '', '');
    h += kpiCard('With contact person', withPersonsPct + '%', fmtNum(withPersons) + ' companies', 'var(--good)', '');
    h += '</div>';
  }

  // ---- STATE: database composition (active vs dormant) + engagement segments ----
  if (ah && ah.dbTotal) {
    var tC = ah.total;
    var dbC = ah.dbTotal;
    var dorm = dbC - tC;
    var actPct = dbC > 0 ? (tC / dbC * 100) : 0;
    var dormPctC = 100 - actPct;
    h += secLabel('Database composition');
    h += '<div class="entity-card" style="background:var(--card);border:1px solid #dcd5c8;border-radius:12px;padding:16px">';
    h += '<div class="compbar" style="margin-bottom:10px">';
    h += '<span style="width:' + actPct.toFixed(1) + '%;background:var(--c1)">' + (actPct >= 6 ? actPct.toFixed(0) + '%' : '') + '</span>';
    h += '<span style="width:' + dormPctC.toFixed(1) + '%;background:var(--cn);color:#6b6657;font-weight:500">Dormant tail \u00b7 ' + fmtNum(dorm) + ' companies with no recent activity</span>';
    h += '</div>';
    h += '<div style="display:flex;gap:20px;flex-wrap:wrap;margin-bottom:6px">';
    h += dotLeg('var(--c1)', 'Active companies', fmtNum(tC), actPct.toFixed(1) + '%');
    h += dotLeg('var(--cn)', 'Dormant tail' + sbInfoIcon('Companies with no registered activity in the selected period. The active companies are the rest. Dormant records are never removed, so they accumulate over time.'), fmtNum(dorm), dormPctC.toFixed(1) + '%');
    h += '</div>';
    if (fn && fn.segments && fn.segments.length > 0) {
      var segs = fn.segments;
      var segTot = fn.total || tC;
            var engCol = hcEngColor;
      h += '<div style="font-size:12px;color:' + MUTED + ';border-top:1px solid #eee5d8;padding-top:12px;margin-top:10px;margin-bottom:10px">How engaged the active companies are</div>';
      h += '<div class="engbar" style="margin-bottom:12px">';
      for (var sa = 0; sa < segs.length; sa++) {
        var sp = segTot > 0 ? (segs[sa].count / segTot * 100) : 0;
        var segIn = sp >= 9 ? sp.toFixed(0) + '%' : '';
        h += '<span style="width:' + sp.toFixed(1) + '%;background:' + engCol(segs[sa].name) + '">' + segIn + '</span>';
      }
      h += '</div>';
      h += '<div style="display:grid;grid-template-columns:1fr 1fr;gap:9px 24px">';
      for (var sb = 0; sb < segs.length; sb++) {
        var sbp = segTot > 0 ? Math.round(segs[sb].count / segTot * 1000) / 10 : 0;
        h += '<div style="display:flex;gap:7px"><span style="width:10px;height:10px;border-radius:3px;flex:none;background:' + engCol(segs[sb].name) + ';margin-top:3px"></span><div style="min-width:0"><div style="font-size:12px;color:#1c2b29"><b style="font-weight:600">' + segs[sb].name + '</b> <span style="color:#6b706c;font-weight:500;margin-left:2px">' + fmtNum(segs[sb].count) + ' \u00b7 ' + sbp + '%</span></div>';
        if (segs[sb].description) h += '<div style="font-size:13px;color:var(--muted);line-height:1.35;margin-top:1px">' + segs[sb].description + '</div>';
        h += '</div></div>';
      }
      h += '</div>';
    }
    h += '</div>';
  }

  // ---- TRAJECTORY: registration history (retention overlay only when it actually varies) ----
  if (trend && trend.length > 0) {
    var maxC = 1, minRet = 100;
    for (var mi = 0; mi < trend.length; mi++) {
      if (trend[mi].count > maxC) maxC = trend[mi].count;
      if (trend[mi].count > 0) { var rr = Math.round(trend[mi].active / trend[mi].count * 100); if (rr < minRet) minRet = rr; }
    }
    // Under the active scope every company is active by definition, so retention sits near 100%
    // for every year and carries no signal. In that case show the age distribution instead.
    var saturated = (minRet >= 90);
    var n = trend.length;
    var sumC = 0, sumA = 0;
    // Prototype chart: a CSS flex row, so the columns always spread evenly across the
    // card. The previous inline SVG had a fixed viewBox and preserved its aspect ratio,
    // so on a wide card it stayed a narrow band with the bars clustered together.
    var cols = '', xax = '', rets = '';
    for (var bi = 0; bi < n; bi++) {
      var t = trend[bi]; sumC += t.count; sumA += t.active;
      var bh = t.count > 0 ? Math.max(2, Math.round(t.count / maxC * 190)) : 0;
      cols += '<div class="col' + (bi === n - 1 ? ' partial' : '') + '"><span class="cval">' + t.count + '</span><div class="cbar" style="height:' + bh + 'px"></div></div>';
      xax += '<span>' + t.year + '</span>';
      if (!saturated) {
        var ret = t.count > 0 ? Math.round(t.active / t.count * 100) : 0;
        rets += '<span style="color:var(--faint);font-size:11px">' + ret + '%</span>';
      }
    }
    var svg = '<div class="chart" role="img" aria-label="Companies by registration year">' + cols + '</div>';
    svg += '<div class="xaxis">' + xax + '</div>';
    if (rets) svg += '<div class="xaxis" style="padding-top:0">' + rets + '</div>';
    h += secLabel('Trajectory \u00b7 over time');
    h += '<div class="entity-card" style="background:var(--card);border:1px solid #e6e1d8;border-radius:12px;padding:18px">';
    if (saturated) {
      h += '<div style="font-size:17px;font-weight:700;color:var(--ink);margin-bottom:3px">Active companies by registration year</div>';
      h += '<div style="font-size:12px;color:' + MUTED + ';margin-bottom:12px">Each bar counts the currently active companies by the year they were first registered. ' + fmtNum(trendBefore) + ' were registered before this range.</div>';
      h += svg;
      h += '<div style="font-size:12px;color:#5b6b5f;background:#eef3ee;border-radius:6px;padding:10px 12px;margin-top:12px;line-height:1.5">Only currently active companies are counted here, grouped by registration year. Companies that are no longer active are left out, so this reflects the registration years of the part of the base still in use.</div>';
    } else {
      var overallRet = sumC > 0 ? Math.round(sumA / sumC * 100) : 0;
      h += '<div style="display:flex;align-items:baseline;margin-bottom:3px"><div style="font-size:17px;font-weight:700;color:var(--ink)">New companies per year</div><div style="margin-left:auto"><span style="font-size:17px;font-weight:800;color:var(--ink)">' + overallRet + '%</span><span style="font-size:12px;color:' + MUTED + '"> still active</span></div></div>';
      h += '<div style="font-size:12px;color:' + MUTED + ';margin-bottom:12px">Bars show how many companies were added each year. The percentage under each bar is how many of that year\u0027s companies are still active today.</div>';
      h += svg;
      h += '<div style="font-size:12px;color:#5b6b5f;background:#eef3ee;border-radius:6px;padding:10px 12px;margin-top:12px;line-height:1.5">A low percentage on the most recent years means newly added companies stop being used quickly. The older years stay high partly because inactive companies are never removed from the database.</div>';
    }
    h += '</div>';
  }

  // ---- CONTEXT: category mix (demoted) ----
  if (dists && dists.length > 0) {
    var catDist = null;
    for (var ci = 0; ci < dists.length; ci++) {
      var tl = (dists[ci].title || '').toLowerCase();
      if (tl.indexOf('categor') >= 0) { catDist = dists[ci]; break; }
    }
    if (catDist && catDist.items && catDist.items.length > 0) {
      var items = catDist.items.slice().sort(function(a, b) { return b.count - a.count; });
      var catTot = catDist.total || 0;
      if (catTot <= 0) { catTot = 0; for (var ck = 0; ck < items.length; ck++) catTot += items[ck].count; }
      var palette = ['#0d5f59', '#2f817a', '#51a399', '#82c2b8', '#b3ddd5', '#dcebe7'];
      var topN = items.slice(0, 6), otherC = 0;
      for (var ok = 6; ok < items.length; ok++) otherC += items[ok].count;
      h += secLabel('Context');
      h += '<div class="entity-card" style="background:var(--card);border:1px solid #e6e1d8;border-radius:12px;padding:16px">';
      h += '<div style="display:flex;align-items:center;margin-bottom:10px"><div style="font-size:16px;font-weight:700;color:var(--ink)">Category mix</div><span style="margin-left:auto;font-size:11px;color:#a09a8e">for context, not something to act on</span></div>';
      h += '<div class="catbar" style="margin-bottom:8px">';
      for (var pi = 0; pi < topN.length; pi++) {
        var pp = catTot > 0 ? (topN[pi].count / catTot * 100) : 0;
        var catIn = pp >= 10 ? pp.toFixed(0) + '%' : '';
        h += '<span style="width:' + pp.toFixed(1) + '%;background:' + palette[pi] + '">' + catIn + '</span>';
      }
      if (otherC > 0) { var op = catTot > 0 ? (otherC / catTot * 100) : 0; h += '<span style="width:' + op.toFixed(1) + '%;background:var(--cn)"></span>'; }
      h += '</div>';
      h += '<div style="display:flex;gap:14px;flex-wrap:wrap;font-size:11px;color:' + MUTED + '">';
      for (var li = 0; li < topN.length; li++) {
        var lp = catTot > 0 ? Math.round(topN[li].count / catTot * 1000) / 10 : 0;
        h += '<span><span style="display:inline-block;width:9px;height:9px;border-radius:2px;vertical-align:-1px;margin-right:4px;background:' + palette[li] + '"></span>' + topN[li].name + ' ' + lp + '%</span>';
      }
      if (otherC > 0) { var lop = catTot > 0 ? Math.round(otherC / catTot * 1000) / 10 : 0; h += '<span><span style="display:inline-block;width:9px;height:9px;border-radius:2px;vertical-align:-1px;margin-right:4px;background:var(--cn)"></span>Other ' + lop + '%</span>'; }
      h += '</div></div>';
    }
  }

  // ---- INSIGHTS (data-driven) ----
  if (ah && ah.dbTotal) {
    var iDorm = ah.dbTotal - ah.total;
    var iDormPct = ah.dbTotal > 0 ? Math.round(iDorm / ah.dbTotal * 100) : 0;
    h += secLabel('Insights \u00b7 recommended actions');
    h += '<div style="position:relative">';
    if (iDormPct >= 40) {
      h += hcInsight('primary', fmtNum(iDorm) + ' companies (' + iDormPct + '%) are a dormant tail.', 'No recent activity, so these are the clearest candidates for cleanup or archiving.', [{ l: 'Create selection', p: true }, { l: 'Export list' }]);
    }
    if (fn && fn.segments) {
      for (var fi = 0; fi < fn.segments.length; fi++) {
        var sg = fn.segments[fi];
        var nm = (sg.name || '').toLowerCase();
        if (nm.indexOf('empty') >= 0 && sg.count > 0) {
          h += hcInsight('warn', fmtNum(sg.count) + ' empty shells.', 'Active companies with no contact person at all. Worth checking whether they are still needed or can be cleaned up.', [{ l: 'Review ' + fmtNum(sg.count) }]);
          break;
        }
      }
    }
    h += '</div>';
  }

  el.innerHTML = h;
  renderScoreBanner('company');
}

// =====================================================
// Sale Overview v2 — performance first (win rate, lost ratio), then pipeline data quality,
// composition by status and type. Built from overviewData['sale'] (stats + distributions).
// "With Activities" and the Adoption score are scope artifacts under the active scope and
// are deliberately not used as headline signals here.
// =====================================================
function renderSaleOverviewV2() {
  var el = document.getElementById('saleOverviewContent');
  if (!el) return;
  var ov = (typeof overviewData !== 'undefined' && overviewData['sale']) ? overviewData['sale'] : null;
  if (!ov) return;
  var o = ov.overview || {};
  var dists = ov.distributions || [];

  var GREEN = 'var(--so-green,#06423e)', OKC = 'var(--sl-ok,#f9a825)', WARN = 'var(--sl-warn,#ef6c00)', BAD = 'var(--sl-bad,#c62828)', GOOD = 'var(--sl-good,#2e7d32)', MUTED = 'var(--so-text-muted,#6b706c)';
  var secLabel = hcSecLabel;
  var kpiCard = hcKpi;
  function dotLeg(color, label, val, pc) { var nums = ''; if (val !== '') nums += '<span style="font-weight:600;color:#1c2b29;margin-left:8px">' + val + '</span>'; if (pc !== '') nums += '<span style="color:#8a8f8b;margin-left:5px">' + pc + '</span>'; return '<div style="display:flex;align-items:center;gap:7px;font-size:12px;color:#3c423f;margin:4px 0"><span style="width:10px;height:10px;border-radius:3px;flex:none;background:' + color + '"></span><span>' + label + '</span>' + nums + '</div>'; }
  function findDist(key) { for (var i = 0; i < dists.length; i++) { if ((dists[i].title || '').toLowerCase().indexOf(key) >= 0) return dists[i]; } return null; }
  function dval(d, name) { if (d && d.items) { for (var i = 0; i < d.items.length; i++) { if ((d.items[i].name || '').toLowerCase() === name) return d.items[i].count; } } return 0; }
  function pct(n, t) { return t > 0 ? Math.round(n / t * 1000) / 10 : 0; }

  var total = o.total || 0;
  var statusD = findDist('status');
  var lost = dval(statusD, 'lost'), open = dval(statusD, 'open'), sold = dval(statusD, 'sold'), stalled = dval(statusD, 'stalled');
  var decided = sold + lost;
  var winRate = decided > 0 ? Math.round(sold / decided * 1000) / 10 : 0;
  var stageD = findDist('stage');
  var sb = ov.stageBuckets || null;
  var noStage = sb ? (sb.noValue || 0) : dval(stageD, '(no value)');
  var noStagePct = total > 0 ? Math.round(noStage / total * 100) : 0;
  var typeD = findDist('type');

  var h = '';
  h += dateFilterNotice();

  // ---- STATE: KPIs ----
  h += secLabel('State \u00b7 sales performance');
  h += '<div class="stat-row" style="display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:10px">';
  h += kpiCard('Win rate', winRate + '%', fmtNum(sold) + ' won of ' + fmtNum(decided) + ' decided', winRate < 35);
  h += kpiCard('Lost', fmtNum(lost), pct(lost, total) + '% of all sales', true);
  h += kpiCard('Open pipeline', fmtNum(open), pct(open, total) + '% still open', false);
  h += kpiCard('Stage not set', noStagePct + '%', fmtNum(noStage) + ' without a stage', noStagePct >= 40);
  h += '</div>';

  // ---- Status split (donut) + Sale type ----
  h += '<div style="display:grid;grid-template-columns:1fr 1fr;gap:14px;margin-top:14px">';
  var segs = [{ n: 'Lost', v: lost, c: BAD }, { n: 'Open', v: open, c: OKC }, { n: 'Sold', v: sold, c: GOOD }, { n: 'Stalled', v: stalled, c: 'var(--cn)' }];
  var cum = 0, circles = '';
  for (var si = 0; si < segs.length; si++) {
    var sp = total > 0 ? (segs[si].v / total * 100) : 0;
    circles += '<circle cx="90" cy="90" r="70" fill="none" stroke="' + segs[si].c + '" stroke-width="26" stroke-dasharray="' + sp.toFixed(1) + ' ' + (100 - sp).toFixed(1) + '" stroke-dashoffset="' + (25 - cum).toFixed(1) + '" pathLength="100"/>';
    cum += sp;
  }
  h += '<div class="entity-card" style="background:var(--card);border:1px solid #e6e1d8;border-radius:12px;padding:16px">';
  h += '<div style="font-size:16px;font-weight:700;color:var(--ink);margin-bottom:10px">Status split</div>';
  h += '<div style="display:flex;gap:12px;align-items:center"><svg viewBox="0 0 180 180" width="118" height="118" role="img" aria-label="Sale status split"><circle cx="90" cy="90" r="70" fill="none" stroke="var(--track)" stroke-width="26"/>' + circles + '<text x="90" y="86" text-anchor="middle" font-size="17" font-weight="500" fill="var(--ink)">' + fmtNum(total) + '</text><text x="90" y="103" text-anchor="middle" font-size="10" fill="var(--muted)">sales</text></svg>';
  h += '<div style="flex:1;min-width:0">' + dotLeg(BAD, 'Lost', fmtNum(lost), pct(lost, total) + '%') + dotLeg(OKC, 'Open', fmtNum(open), pct(open, total) + '%') + dotLeg(GOOD, 'Sold', fmtNum(sold), pct(sold, total) + '%') + dotLeg('var(--cn)', 'Stalled' + sbInfoIcon('A sale put on hold, neither won nor lost yet.'), fmtNum(stalled), pct(stalled, total) + '%') + '</div></div></div>';
  if (typeD && typeD.items && typeD.items.length > 0) {
    var titems = typeD.items.slice().sort(function (a, b) { return b.count - a.count; });
    var ttot = 0; for (var ti = 0; ti < titems.length; ti++) ttot += titems[ti].count;
    var tpal = ['var(--c1)', 'var(--c3)', 'var(--c5)', 'var(--cn)'];
    h += '<div class="entity-card" style="background:var(--card);border:1px solid #e6e1d8;border-radius:12px;padding:16px"><div style="font-size:16px;font-weight:700;color:var(--ink)">Sale type</div><div style="font-size:13px;color:var(--muted);margin:2px 0 11px">From the Sale Type field on each sale, chosen by the user. It is not derived from the data.</div>';
    for (var tj = 0; tj < titems.length && tj < 4; tj++) {
      var tp = pct(titems[tj].count, ttot);
      h += '<div style="margin-bottom:9px"><div style="display:flex;font-size:12px;margin-bottom:3px"><span>' + titems[tj].name + '</span><b style="margin-left:auto;color:#6b706c">' + tp + '%</b></div><div style="height:9px;background:#eef0ec;border-radius:5px"><div style="width:' + Math.max(tp, 1).toFixed(0) + '%;height:9px;background:' + tpal[Math.min(tj, 3)] + ';border-radius:5px"></div></div></div>';
    }
    h += '</div>';
  }
  h += '</div>';

  // ---- The real gap: pipeline (stage) data quality ----
  if (sb || (stageD && stageD.items)) {
    var low = sb ? (sb.low || 0) : dval(stageD, 'low chance'), mid = sb ? (sb.mid || 0) : dval(stageD, 'middle chance'), high = sb ? (sb.high || 0) : dval(stageD, 'high chance');
    var spNo = pct(noStage, total), spLow = pct(low, total), spMid = pct(mid, total), spHigh = pct(high, total);
    // Build the assessment from the actual distribution so it holds for any tenant.
    var stFilled = low + mid + high;
    var stDomName = 'Low chance', stDomCount = low;
    if (mid > stDomCount) { stDomCount = mid; stDomName = 'Middle chance'; }
    if (high > stDomCount) { stDomCount = high; stDomName = 'High chance'; }
    var stDomShare = stFilled > 0 ? (stDomCount / stFilled * 100) : 0;
    var stagePoor = (spNo >= 40) || (stFilled > 0 && stDomShare >= 75);
    var stageMsg;
    if (spNo >= 50) {
      stageMsg = 'Most sales have no stage filled in' + (stFilled > 0 && stDomShare >= 70 ? ', and the ones that do are nearly all ' + stDomName : '') + '. With so little to go on, the stage field gives a weak basis for forecasting the pipeline.';
    } else if (spNo >= 25) {
      stageMsg = 'A large share of sales have no stage filled in' + (stFilled > 0 && stDomShare >= 70 ? ', and the rest cluster at ' + stDomName : '') + ', which weakens stage as a forecasting signal.';
    } else if (stDomShare >= 75) {
      stageMsg = 'Stage is filled in for most sales, but they cluster heavily at ' + stDomName + ', so the field offers little spread to forecast from.';
    } else {
      stageMsg = 'Stage is filled in for most sales and spread across the buckets, which gives a usable basis for pipeline forecasting.';
    }
    h += secLabel(stagePoor ? 'The real gap \u00b7 pipeline data' : 'Pipeline \u00b7 stage tracking');
    h += '<div class="entity-card" style="background:var(--card);border:1px solid #e6e1d8;border-radius:12px;padding:16px"><div style="font-size:16px;font-weight:700;color:var(--ink);margin-bottom:2px">Stage / probability tracking' + sbInfoIcon('The probability field on each sale (Low, Middle or High chance). It is what pipeline forecasting relies on.') + '</div><div style="font-size:12px;color:' + MUTED + ';margin-bottom:12px;line-height:1.5">' + stageMsg + '</div>';
    h += '<div class="engbar" style="height:30px;margin-bottom:12px">';
    h += '<span style="width:' + spNo + '%;background:var(--cn);color:#6b6657;font-weight:500">' + (spNo >= 12 ? spNo.toFixed(0) + '% no stage set' : '') + '</span>';
    h += '<span style="width:' + spLow + '%;background:' + OKC + '">' + (spLow >= 14 ? spLow.toFixed(0) + '% low chance' : (spLow >= 7 ? spLow.toFixed(0) + '%' : '')) + '</span>';
    h += '<span style="width:' + spMid + '%;background:var(--c3)"></span><span style="width:' + spHigh + '%;background:' + GOOD + '"></span></div>';
    h += '<div style="display:grid;grid-template-columns:1fr 1fr;gap:2px 24px">' + dotLeg('var(--cn)', 'No value', fmtNum(noStage), spNo + '%') + dotLeg('var(--c3)', 'Middle chance', fmtNum(mid), spMid + '%') + dotLeg(OKC, 'Low chance', fmtNum(low), spLow + '%') + dotLeg(GOOD, 'High chance', fmtNum(high), spHigh + '%') + '</div></div>';
  }

  // ---- INSIGHTS ----
  h += secLabel('Insights');
  var insH = '';
  if (decided > 0 && winRate < 35) {
    insH += hcInsight('primary', 'Win rate is ' + winRate + '%, with ' + pct(lost, total) + '% of sales lost.', 'A useful first question is what the ' + fmtNum(lost) + ' lost deals had in common.');
  }
  if (noStagePct >= 40) {
    insH += hcInsight('warn', noStagePct + '% of sales have no stage set.', 'The pipeline cannot be forecast reliably until stage is filled in.');
  }
  var noQuote = (o.withQuote === 0), noStake = (o.withStakeholders === 0);
  if (noQuote || noStake) {
    var qsTitle = (noQuote && noStake) ? 'Quotes and stakeholders are unused' : (noQuote ? 'Quotes are unused' : 'Stakeholders are unused');
    var qsDetail = (noQuote && noStake) ? 'Neither is filled in on any sale, so both features are currently left untouched.' : 'It is not filled in on any sale, so the feature is currently left untouched.';
    insH += hcInsight('neutral', qsTitle + '.', qsDetail);
  }
  if (insH === '') {
    insH = hcInsight('good', '', 'No major data issues stand out for sales in this period.');
  }
  h += insH;

  el.innerHTML = h;
}

// =====================================================
// Contact Overview v2 — sober: reachability and field completeness from overviewData['contact'].
// =====================================================
function renderContactOverviewV2() {
  var el = document.getElementById('contactOverviewContent');
  if (!el) return;
  var ov = (typeof overviewData !== 'undefined' && overviewData['contact']) ? overviewData['contact'] : null;
  if (!ov) return;
  var o = ov.overview || {};
  var GREEN = 'var(--so-green,#06423e)', BAD = 'var(--sl-bad,#c62828)', WARN = 'var(--sl-warn,#ef6c00)', MUTED = 'var(--so-text-muted,#6b706c)';
  var secLabel = hcSecLabel;
  var kpiCard = hcKpi;
  var total = o.total || 0;
  function P(n) { return total > 0 ? Math.round((o[n] || 0) / total * 1000) / 10 : 0; }
  function bar(label, n) { var p = total > 0 ? Math.round(n / total * 1000) / 10 : 0; return '<div style="margin-bottom:11px"><div style="display:flex;font-size:12px;margin-bottom:4px;color:#3c423f"><span>' + label + '</span><span style="margin-left:auto;color:#6b706c"><b style="color:#1c2b29">' + p + '%</b> &middot; ' + fmtNum(n) + '</span></div><div style="height:10px;background:#eef0ec;border-radius:5px"><div style="width:' + Math.max(p, 1).toFixed(0) + '%;height:10px;background:#0f5c57;border-radius:5px"></div></div></div>'; }

  var h = '';
  h += dateFilterNotice();
  h += secLabel('State \u00b7 reachability');
  h += '<div class="stat-row" style="display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:10px">';
  h += kpiCard('Contacts', fmtNum(total), 'people in scope', false);
  h += kpiCard('With email', P('withEmail') + '%', fmtNum(o.withEmail || 0) + ' reachable', false);
  h += kpiCard('With phone', P('withPhone') + '%', fmtNum(o.withPhone || 0) + ' reachable', P('withPhone') < 50);
  h += kpiCard('With position', P('withPosition') + '%', fmtNum(o.withPosition || 0) + ' have a role', P('withPosition') < 60);
  h += '</div>';
  h += secLabel('Field completeness');
  h += '<div class="entity-card" style="background:var(--card);border:1px solid #e6e1d8;border-radius:12px;padding:16px"><div style="font-size:12px;color:' + MUTED + ';margin-bottom:12px">How completely the key contact fields are filled in.</div>';
  h += bar('Email address', o.withEmail || 0);
  h += bar('Phone number', o.withPhone || 0);
  h += bar('Position', o.withPosition || 0);
  h += bar('Job title', o.withTitle || 0);
  h += '</div>';
  h += secLabel('Links to the rest of the CRM');
  h += '<div class="stat-row" style="display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:10px">';
  h += kpiCard('Linked to a company', P('withCompany') + '%', fmtNum(o.withCompany || 0) + ' contacts', false);
  h += kpiCard('Linked to sales' + sbInfoIcon('Contacts connected to at least one sale.'), P('withSales') + '%', fmtNum(o.withSales || 0) + ' contacts', false);
  h += kpiCard('In a project', P('inProjects') + '%', fmtNum(o.inProjects || 0) + ' contacts', false);
  h += '</div>';
  h += secLabel('Insights');
  var insC = '';
  var posPct = P('withPosition'), titlePct = P('withTitle'), phonePct = P('withPhone'), emailPct = P('withEmail');
  if (posPct < 60 || titlePct < 60) {
    insC += hcInsight('warn', 'Role data is incomplete.', 'Position is filled for ' + posPct + '% and job title for ' + titlePct + '% of contacts. That limits any targeting or segmentation by role.');
  }
  if (emailPct < 85) {
    insC += hcInsight('warn', (100 - Math.round(emailPct)) + '% of contacts have no email address.', 'Email is the main outreach channel for most teams, so those gaps limit who can be reached.');
  }
  if (phonePct < 75) {
    var emNote = emailPct >= 85 ? ' Email coverage is strong at ' + emailPct + '%, so most contacts are still reachable, just not by phone.' : '';
    insC += hcInsight('neutral', (100 - Math.round(phonePct)) + '% of contacts have no phone number.', emNote);
  }
  if (insC === '') {
    insC = hcInsight('good', '', 'Contact detail coverage looks healthy across the main fields.');
  }
  h += insC;
  el.innerHTML = h;
}

// =====================================================
// Project Overview v2 — sober and honest about a tiny, barely used dataset.
// =====================================================
function renderProjectOverviewV2() {
  var el = document.getElementById('projectOverviewContent');
  if (!el) return;
  var ov = (typeof overviewData !== 'undefined' && overviewData['project']) ? overviewData['project'] : null;
  if (!ov) return;
  var o = ov.overview || {};
  var dists = ov.distributions || [];
  var GREEN = 'var(--so-green,#06423e)', BAD = 'var(--sl-bad,#c62828)', WARN = 'var(--sl-warn,#ef6c00)', MUTED = 'var(--so-text-muted,#6b706c)';
  var secLabel = hcSecLabel;
  var kpiCard = hcKpi;
  function findDist(k) { for (var i = 0; i < dists.length; i++) { if ((dists[i].title || '').toLowerCase().indexOf(k) >= 0) return dists[i]; } return null; }
  function miniList(title, d) { var s = '<div class="entity-card" style="background:var(--card);border:1px solid #e6e1d8;border-radius:12px;padding:16px"><div style="font-size:16px;font-weight:700;color:var(--ink);margin-bottom:10px">' + title + '</div>'; var any = false; if (d && d.items) { for (var i = 0; i < d.items.length; i++) { if (d.items[i].count > 0) { any = true; s += '<div style="display:flex;font-size:12px;margin:5px 0;color:#3c423f"><span>' + d.items[i].name + '</span><b style="margin-left:auto;color:#1c2b29">' + fmtNum(d.items[i].count) + '</b></div>'; } } } if (!any) s += '<div style="font-size:12px;color:' + MUTED + '">No values set.</div>'; return s + '</div>'; }

  var total = o.total || 0, overdue = o.overdue || 0, withMembers = o.withMembers || 0;
  var h = '';
  h += dateFilterNotice();
  h += secLabel('State \u00b7 projects');
  h += '<div class="stat-row" style="display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:10px">';
  h += kpiCard('Projects', fmtNum(total), 'in scope', false);
  h += kpiCard('Overdue' + sbInfoIcon('Projects whose end date has already passed.'), fmtNum(overdue), (total > 0 ? Math.round(overdue / total * 100) : 0) + '% past due', overdue > 0);
  h += kpiCard('With members', fmtNum(withMembers), (total > 0 ? Math.round(withMembers / total * 100) : 0) + '% have members', withMembers === 0);
  h += '</div>';
  h += '<div style="display:grid;grid-template-columns:1fr 1fr;gap:14px;margin-top:14px">' + miniList('Project status', findDist('status')) + miniList('Project type', findDist('type')) + '</div>';
  h += secLabel('Insights');
  var insP = '';
  if (total > 0 && total <= 10) {
    var od = overdue > 0 ? (', ' + (overdue >= total ? 'all' : fmtNum(overdue)) + ' of them overdue') : '';
    var mem = withMembers === 0 ? ', and none have members assigned' : '';
    insP += hcInsight('warn', 'Projects are barely used.', 'There ' + (total === 1 ? 'is' : 'are') + ' only ' + fmtNum(total) + ' project' + (total === 1 ? '' : 's') + ' in this CRM' + od + mem + '. If projects matter to the business this is an adoption gap rather than a data quality one.');
  } else if (total > 10) {
    var odPct = Math.round(overdue / total * 100);
    if (odPct >= 40) {
      insP += hcInsight('warn', odPct + '% of projects are overdue.', 'Their end date has passed while the project is still open, so either the dates or the statuses need maintaining.');
    }
    if (withMembers === 0) {
      insP += hcInsight('warn', 'No projects have members assigned.', 'Without members it is hard to see who is responsible for each project.');
    }
  }
  if (insP === '') {
    insP = hcInsight('good', '', 'No major issues stand out for projects.');
  }
  h += insP;
  el.innerHTML = h;
}

// =====================================================
// Requests (Ticket) Overview v2 — service desk view from overviewData['requests'].
// The named status distribution is the source of truth. Base open/closed/postponed
// status ids are not surfaced because custom statuses can map to any base status,
// so those counts are unreliable across tenants.
// =====================================================
function renderRequestsOverviewV2() {
  var el = document.getElementById('requestsOverviewContent');
  if (!el) return;
  var ov = (typeof overviewData !== 'undefined' && overviewData['requests']) ? overviewData['requests'] : null;
  if (!ov) return;
  var o = ov.overview || {};
  var dists = ov.distributions || [];

  var GREEN = 'var(--so-green,#06423e)', OKC = 'var(--sl-ok,#f9a825)', WARN = 'var(--sl-warn,#ef6c00)', BAD = 'var(--sl-bad,#c62828)', GOOD = 'var(--sl-good,#2e7d32)', MUTED = 'var(--so-text-muted,#6b706c)';
  var secLabel = hcSecLabel;
  var kpiCard = hcKpi;
  function findDist(key) { for (var i = 0; i < dists.length; i++) { if ((dists[i].title || '').toLowerCase().indexOf(key) >= 0) return dists[i]; } return null; }
  function pct(n, t) { return t > 0 ? Math.round(n / t * 1000) / 10 : 0; }

  // Ranked horizontal bars with count and percentage labels, plus an Other rollup.
  function rankCard(title, note, dist, palette, topN, tip) {
    if (!dist || !dist.items) return '';
    var items = [];
    for (var i = 0; i < dist.items.length; i++) {
      var it = dist.items[i];
      if ((it.count || 0) > 0 && (it.name || '') !== '') items.push({ n: it.name, v: it.count });
    }
    if (items.length === 0) return '';
    items.sort(function (a, b) { return b.v - a.v; });
    var tot = 0; for (var j = 0; j < items.length; j++) tot += items[j].v;
    var head = items.slice(0, topN);
    var restV = 0; for (var k = topN; k < items.length; k++) restV += items[k].v;
    var s = '<div class="entity-card" style="background:var(--card);border:1px solid #e6e1d8;border-radius:12px;padding:16px"><div style="font-size:16px;font-weight:700;color:var(--ink)">' + title + (tip ? sbInfoIcon(tip) : '') + '</div>';
    s += note ? '<div style="font-size:13px;color:var(--muted);margin:2px 0 11px">' + note + '</div>' : '<div style="height:9px"></div>';
    for (var m = 0; m < head.length; m++) {
      var p = pct(head[m].v, tot);
      var col = palette[Math.min(m, palette.length - 1)];
      s += '<div style="margin-bottom:9px"><div style="display:flex;font-size:12px;margin-bottom:3px;gap:8px"><span style="overflow:hidden;text-overflow:ellipsis;white-space:nowrap">' + head[m].n + '</span><b style="margin-left:auto;color:#1c2b29;flex:none">' + fmtNum(head[m].v) + '</b><span style="color:#8a8f8b;flex:none;width:48px;text-align:right">' + p + '%</span></div><div style="height:9px;background:#eef0ec;border-radius:5px"><div style="width:' + Math.max(p, 0.6).toFixed(1) + '%;height:9px;background:' + col + ';border-radius:5px"></div></div></div>';
    }
    if (restV > 0) {
      var rp = pct(restV, tot);
      s += '<div style="margin-bottom:2px"><div style="display:flex;font-size:12px;margin-bottom:3px;gap:8px;color:#8a8f8b"><span>Other (' + (items.length - topN) + ')</span><b style="margin-left:auto;color:#6b706c;flex:none">' + fmtNum(restV) + '</b><span style="flex:none;width:48px;text-align:right">' + rp + '%</span></div><div style="height:9px;background:#eef0ec;border-radius:5px"><div style="width:' + Math.max(rp, 0.6).toFixed(1) + '%;height:9px;background:#cfc9bd;border-radius:5px"></div></div></div>';
    }
    return s + '</div>';
  }

  var total = o.total || 0;
  var unassigned = o.unassigned || 0, engaged = o.engaged || 0;
  var unaPct = pct(unassigned, total), engPct = pct(engaged, total);

  var statusD = findDist('status');
  var prioD = findDist('priorit');
  var typeD = findDist('type');

  var h = '';
  h += dateFilterNotice();

  // ---- STATE: KPIs ----
  h += secLabel('State \u00b7 service desk');
  h += '<div class="stat-row" style="display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:10px">';
  h += kpiCard('Total tickets', fmtNum(total), 'in scope', false);
  h += kpiCard('Unassigned' + sbInfoIcon('Tickets with no owner set, still held by the system user.'), unaPct + '%', fmtNum(unassigned) + ' without an owner', unaPct >= 30);
  h += kpiCard('With replies' + sbInfoIcon('Tickets where at least one reply was sent to the requester.'), engPct + '%', fmtNum(engaged) + ' answered', false);
  h += '</div>';

  // ---- Status breakdown (named statuses, the configured truth) ----
  if (statusD) {
    h += secLabel('Status');
    h += rankCard('Status breakdown', 'The ticket statuses configured in this installation.', statusD, ['var(--c1)', 'var(--c2)', 'var(--c3)', 'var(--c4)', 'var(--c5)', 'var(--c6)'], 6, '');
  }

  // ---- Field usage: priority + ticket type ----
  var pc = prioD ? rankCard('Priority', '', prioD, ['var(--c1)', 'var(--c2)', 'var(--c3)', 'var(--c4)', 'var(--cn)'], 5, 'The priority set on each ticket, chosen by the user.') : '';
  var tc = typeD ? rankCard('Ticket type', '', typeD, ['var(--c1)', 'var(--c2)', 'var(--c3)', 'var(--c4)', 'var(--cn)'], 5, 'The ticket type chosen on each ticket. It is not derived from the data.') : '';
  if (pc || tc) {
    h += secLabel('Field usage');
    h += '<div style="display:grid;grid-template-columns:1fr 1fr;gap:14px">' + pc + tc + '</div>';
  }

  // ---- INSIGHTS (data-driven and guarded) ----
  h += secLabel('Insights');
  var insR = '';
  if (total > 0 && unaPct >= 30) {
    insR += hcInsight('warn', unaPct + '% of tickets have no owner.', fmtNum(unassigned) + ' tickets are not assigned to anyone, which weakens any reporting on workload or responsibility per agent.');
  }
  // Priority concentration: if one value dominates, the field gives no triage signal.
  if (prioD && prioD.items) {
    var pi = [];
    for (var pidx = 0; pidx < prioD.items.length; pidx++) { var pit = prioD.items[pidx]; if ((pit.count || 0) > 0 && (pit.name || '') !== '' && (pit.name || '').toLowerCase().indexOf('no value') < 0) pi.push(pit); }
    var ptot = 0; for (var q = 0; q < pi.length; q++) ptot += pi[q].count;
    if (ptot > 0) {
      var pdom = pi[0]; for (var r = 1; r < pi.length; r++) if (pi[r].count > pdom.count) pdom = pi[r];
      var pdomShare = pct(pdom.count, ptot);
      if (pdomShare >= 85) {
        insR += hcInsight('neutral', 'Priority is set to ' + pdom.name + ' for ' + pdomShare + '% of tickets.', 'With almost everything on one value, priority gives little signal for triage or filtering.');
      }
    }
  }
  if (total > 0 && engPct < 30) {
    insR += hcInsight('neutral', engPct + '% of tickets have a reply logged.', 'If most contact runs by phone or outside the ticket that can be expected, otherwise it suggests replies are not being captured in the CRM.');
  }
  if (insR === '') {
    insR = hcInsight('good', '', 'No major issues stand out for tickets in this period.');
  }
  h += insR;

  el.innerHTML = h;
}

function renderEntityOverview(key, d) {
  overviewData[key] = d;
  pushStdListValues(key, d);
  if (key === 'company') { renderCompanyOverviewV2(); return; }
  if (key === 'sale') { renderSaleOverviewV2(); return; }
  if (key === 'contact') { renderContactOverviewV2(); return; }
  if (key === 'project') { renderProjectOverviewV2(); return; }
  if (key === 'requests') { renderRequestsOverviewV2(); return; }
  var cfg = ovLabels[key];
  var o = d.overview;
  var totalKey = cfg.stats[0][0];
  var total = o[totalKey];
  var h = '';

  h += dateFilterNotice();

  h += '<div class="detail-section">';
  h += '<div class="detail-section-head">' + secHead(cfg.title) + '</div>';
  h += '<div class="stat-row">';
  for (var i = 0; i < cfg.stats.length; i++) {
    var s = cfg.stats[i];
    if (s[0] === totalKey) {
      h += ovCard(s[1], total, '', '');
    } else {
      h += ovCard(s[1], o[s[0]], total, '');
    }
  }
  h += '</div></div>';

  if (cfg.sections) {
    for (var si = 0; si < cfg.sections.length; si++) {
      var sec = cfg.sections[si];
      var secTotal = o[sec.totalKey];
      h += '<div class="detail-section">';
      h += '<div class="detail-section-head">' + secHead(sec.title) + '</div>';
      h += '<div class="stat-row">';
      for (var sj = 0; sj < sec.stats.length; sj++) {
        var ss = sec.stats[sj];
        if (ss[0] === sec.totalKey) {
          h += ovCard(ss[1], secTotal, '', '');
        } else if (ss[2]) {
          h += ovCard(ss[1], o[ss[0]], '', '');
        } else {
          h += ovCard(ss[1], o[ss[0]], secTotal, '');
        }
      }
      h += '</div></div>';
    }
  }

  // Completeness is now shown in the Data Quality tab — skip here
  // (renderDQScore handles completeness display)

  if (d.distributions && d.distributions.length > 0) {
    var cols = d.distributions.length >= 2 ? 2 : 1;
    if (cols === 2) {
      h += '<div style="display:grid;grid-template-columns:1fr 1fr;gap:18px;margin-bottom:18px">';
      var dt0 = d.distributions[0].total || total;
      var dt1 = d.distributions[1].total || total;
      h += distTable(d.distributions[0].title, d.distributions[0].items, dt0);
      h += distTable(d.distributions[1].title, d.distributions[1].items, dt1);
      h += '</div>';
      for (var k = 2; k < d.distributions.length; k++) {
        var dtk = d.distributions[k].total || total;
        h += distTable(d.distributions[k].title, d.distributions[k].items, dtk);
      }
    } else {
      for (var k = 0; k < d.distributions.length; k++) {
        var dtk2 = d.distributions[k].total || total;
        h += distTable(d.distributions[k].title, d.distributions[k].items, dtk2);
      }
    }
  }

  document.getElementById(key + 'OverviewContent').innerHTML = h;
}

function secHead(t) {
  return '<h4 style="font-size:.85rem;text-transform:uppercase;color:var(--so-text-muted);margin-bottom:12px;font-weight:600;letter-spacing:.5px">' + t + '</h4>';
}

function ovCard(label, value, total, color) {
  var v = (typeof value === 'number' && !isNaN(value)) ? value : 0;
  var h = '<div class="stat-card">';
  h += '<div class="stat-value">' + fmtNum(v) + '</div>';
  h += '<div class="stat-label">' + label + '</div>';
  if (total && total > 0) {
    var pct = Math.round((v / total) * 1000) / 10;
    if (isNaN(pct)) pct = 0;
    h += '<div style="margin-top:6px">' + fillBar(pct, 8, color) + '</div>';
    h += '<div class="stat-label" style="margin-top:3px">' + pct + P + '</div>';
  }
  h += '</div>';
  return h;
}

function pctCard(label, pct, desc, color, count, base) {
  var v = (typeof pct === 'number' && !isNaN(pct)) ? pct : 0;
  var h = '<div class="stat-card">';
  if (typeof count === 'number' && typeof base === 'number') {
    h += '<div class="stat-value">' + fmtNum(count) + '<span style="font-size:.55em;color:#999;font-weight:400"> / ' + fmtNum(base) + '</span></div>';
    h += '<div class="stat-label">' + label + ' <span style="font-weight:700;color:' + color + '">' + v + P + '</span></div>';
  } else {
    h += '<div class="stat-value">' + v + P + '</div>';
    h += '<div class="stat-label">' + label + '</div>';
  }
  h += '<div style="margin-top:6px">' + fillBar(v, 8, color) + '</div>';
  h += '<div class="stat-label" style="margin-top:4px;font-size:.7rem;color:#999">' + desc + '</div>';
  h += '</div>';
  return h;
}

function fmtNum(n) {
  if (typeof n !== 'number') return n;
  return n.toLocaleString();
}

function fmtDate(d) {
  if (!d || d === '0' || d === '') return '';
  var datePart = d.split(' ')[0];
  var parts = datePart.split('-');
  if (parts.length === 3 && parts[0].length === 4) return parts[2] + '-' + parts[1] + '-' + parts[0];
  return d;
}

function distTable(title, items, total) {
  if (!items || items.length === 0) return '';
  items.sort(function(a, b) { return b.count - a.count; });
  var Q = String.fromCharCode(39);
  var tid = 'tbl-dist-' + title.toLowerCase().replace(/[^a-z]/g, '') + '-' + Math.random().toString(36).substr(2,4);
  var hasLastUsed = items.length > 0 && (items[0].lastUsed !== undefined);
  var h = '<div class="entity-card">';
  h += '<div class="entity-header"><div class="entity-info"><h3>' + title + '</h3></div>';
  h += '<span class="record-badge">' + items.length + ' values</span></div>';
  h += '<table class="data-table" id="' + tid + '">';
  h += '<thead><tr>';
  var colIdx = 0;
  h += '<th onclick="sortT(' + Q + tid + Q + ',' + colIdx + ')">' + title + ' <span class="sort-arrow">' + svgSortN + '</span></th>';
  colIdx++;
  h += '<th class="col-right" onclick="sortT(' + Q + tid + Q + ',' + colIdx + ')">Count <span class="sort-arrow active">' + svgSortD + '</span></th>';
  colIdx++;
  h += '<th class="col-right" onclick="sortT(' + Q + tid + Q + ',' + colIdx + ')">' + P + ' <span class="sort-arrow">' + svgSortN + '</span></th>';
  colIdx++;
  if (hasLastUsed) {
    h += '<th class="col-right" onclick="sortT(' + Q + tid + Q + ',' + colIdx + ')">Last Used <span class="sort-arrow">' + svgSortN + '</span></th>';
    colIdx++;
  }
  h += '</tr></thead><tbody>';
  for (var i = 0; i < items.length; i++) {
    var it = items[i];
    var pct = total > 0 ? Math.round((it.count / total) * 1000) / 10 : 0;
    var cls = it.count === 0 ? ' class="unused"' : '';
    h += '<tr' + cls + '>';
    h += '<td data-sort-value="' + it.name + '">' + it.name + '</td>';
    h += '<td class="col-right" data-sort-value="' + it.count + '">' + fmtNum(it.count) + '</td>';
    h += '<td class="col-right" data-sort-value="' + pct + '">' + pct + P + '</td>';
    if (hasLastUsed) {
      var lu = it.lastUsed || '';
      var luDisplay = (lu && lu !== '0') ? fmtDate(lu) : '<span style="color:#ccc">Never</span>';
      h += '<td class="col-right" data-sort-value="' + lu + '">' + luDisplay + '</td>';
    }
    h += '</tr>';
  }
  h += '</tbody></table></div>';
  return h;
}

function startCompanyUdef() { loadEntityUdef(7, 'company', 100); }
function startContactUdef() { loadEntityUdef(8, 'contact', 200); }
function startSaleUdef() { loadEntityUdef(10, 'sale', 300); }
function startProjectUdef() { loadEntityUdef(9, 'project', 400); }

// === COMPANY DETAILS ===
var companyDetailData = null;
var companyDetailCatValue = '';
// Cache of completed company-detail fetches, keyed by date filter + category.
// Lets category switches reuse a prior result instead of refetching. Cleared on re-analyze.
var companyDetailCache = {};

function loadCompanyDetails(cb) {
  fetchCompanyDetails(cb);
}

function fetchCompanyDetails(cb) {
  var key = getDateFilterParam() + '|' + (companyDetailCatValue || '');

  // Cache hit: restore the previous result and re-render synchronously, no network call.
  if (companyDetailCache[key]) {
    companyDetailData = companyDetailCache[key];
    renderCompanyDetails(companyDetailData);
    if (companyDetailData.funnel) {
      companyCrossData = companyDetailData;
      renderCrossEntityFunnel(companyDetailData);
    }
    populateDetailCatFilter();
    var ce = document.getElementById('companyDetailFilterCount');
    if (ce && companyDetailData.activityHealth) {
      ce.textContent = companyDetailCatValue ? fmtNum(companyDetailData.activityHealth.total) + ' companies' : '';
    }
    if (cb) cb();
    return;
  }

  var catParam = '';
  if (companyDetailCatValue) {
    catParam = String.fromCharCode(38) + 'categoryName=' + encodeURIComponent(companyDetailCatValue);
  }
  var dfParam = getDateFilterParam();
  var pending = 3;
  companyDetailData = {};

  function checkDone() {
    pending--;
    renderCompanyDetails(companyDetailData);
    renderCompanyOverviewV2();
    if (pending <= 0) {
      companyDetailCache[key] = companyDetailData;
      if (cb) cb();
    }
  }

  // Call 1: Core (activity health, trend)
  ajax(coreUrl + dfParam + catParam, function(d) {
    if (d) {
      if (d.activityHealth) companyDetailData.activityHealth = d.activityHealth;
      if (d.trend) companyDetailData.trend = d.trend;
      if (d.trendMonthly) companyDetailData.trendMonthly = d.trendMonthly;
      if (d.trendBefore !== undefined) companyDetailData.trendBefore = d.trendBefore;
    }
    checkDone();
  });

  // Call 2: Quality (quality, funnel, churn)
  ajax(qualityUrl + dfParam + catParam + getExclParam('company'), function(d) {
    if (d) {
      if (d.quality) companyDetailData.quality = d.quality;
      if (d.funnel) companyDetailData.funnel = d.funnel;
      if (d.churnRisk) companyDetailData.churnRisk = d.churnRisk;
    }
    // Re-render funnel
    if (companyDetailData.funnel) {
      companyCrossData = companyDetailData;
      renderCrossEntityFunnel(companyDetailData);
    }
    populateDetailCatFilter();
    var countEl = document.getElementById('companyDetailFilterCount');
    if (countEl && companyDetailData.activityHealth) {
      countEl.textContent = companyDetailCatValue ? fmtNum(companyDetailData.activityHealth.total) + ' companies' : '';
    }
    checkDone();
  });

  // Call 3: Detail (associates, category effectiveness)
  ajax(detailUrl + dfParam + catParam, function(d) {
    if (d) {
      if (d.associates) companyDetailData.associates = d.associates;
      if (d.categoryEffectiveness) companyDetailData.categoryEffectiveness = d.categoryEffectiveness;
    }
    checkDone();
  });
}

function populateDetailCatFilter() {
  // Build the option list once from the company category distribution.
  var opts = [];
  if (overviewData['company'] && overviewData['company'].distributions) {
    var cats = overviewData['company'].distributions[0];
    if (cats && cats.items) {
      var sorted = cats.items.slice().sort(function(a,b) { return b.count - a.count; });
      for (var ci = 0; ci < sorted.length; ci++) {
        if (sorted[ci].count > 0 && sorted[ci].name !== '(No value)') {
          opts.push({ value: sorted[ci].name, label: sorted[ci].name + ' (' + fmtNum(sorted[ci].count) + ')' });
        }
      }
      for (var cj = 0; cj < sorted.length; cj++) {
        if (sorted[cj].name === '(No value)' && sorted[cj].count > 0) {
          opts.push({ value: '__none__', label: '(No category) (' + fmtNum(sorted[cj].count) + ')' });
        }
      }
    }
  }

  // Filtered company count, shown only when a category is active.
  var countText = '';
  if (companyDetailCatValue && typeof companyDetailData !== 'undefined' && companyDetailData && companyDetailData.activityHealth) {
    countText = fmtNum(companyDetailData.activityHealth.total) + ' companies';
  }

  // Apply to every filter bar currently in the DOM (Adoption + Integrity).
  var bars = [
    { sel: 'companyDetailCatFilter', count: 'companyDetailFilterCount', reset: 'companyFilterReset', bar: 'companyCatFilterBar' },
    { sel: 'companyIntegrityDetailCatFilter', count: 'companyIntegrityDetailFilterCount', reset: 'companyIntegrityFilterReset', bar: 'companyIntegrityCatFilterBar' }
  ];
  for (var b = 0; b < bars.length; b++) {
    var sel = document.getElementById(bars[b].sel);
    if (!sel) continue;
    while (sel.options.length > 1) sel.remove(1);
    for (var o = 0; o < opts.length; o++) {
      var opt = document.createElement('option');
      opt.value = opts[o].value;
      opt.textContent = opts[o].label;
      sel.appendChild(opt);
    }
    sel.value = companyDetailCatValue || '';
    var cEl = document.getElementById(bars[b].count); if (cEl) cEl.textContent = countText;
    var rEl = document.getElementById(bars[b].reset); if (rEl) rEl.style.display = companyDetailCatValue ? '' : 'none';
    var barEl = document.getElementById(bars[b].bar); if (barEl) barEl.classList.toggle('active', !!companyDetailCatValue);
  }
}

function daLoadingBlock(text) {
  return '<div class="da-loading"><span class="da-spinner"></span>' + (text || 'Loading\u2026') + '</div>';
}

function reloadCompanyDetails() {
  var lb = daLoadingBlock('Applying filter\u2026');
  var ids = ['companyDetailContent', 'companyCrossContent', 'companyIntegrityContent', 'companyDQScoreContent'];
  for (var i = 0; i < ids.length; i++) {
    var node = document.getElementById(ids[i]);
    if (node) node.innerHTML = lb;
  }
  fetchCompanyDetails(function() {
    renderDQScore('company');
    renderScoreBanner('company');
    loadCompanyCross(null);
    renderAdoptionTab('company');
  });
}

function onDetailFilterChange(key) {
  var sel = document.getElementById(key + 'DetailCatFilter');
  companyDetailCatValue = sel ? sel.value : '';
  reloadCompanyDetails();
}

function resetDetailFilter(key) {
  companyDetailCatValue = '';
  reloadCompanyDetails();
}

// ===========================================================
// CROSS-ENTITY ANALYSIS (Company)
// ===========================================================
var companyCrossData = null;

function loadCompanyCross(cb) {
  // Funnel data is now included in CompanyDetailFetch response
  // No separate AJAX call needed — just render from companyDetailData
  if (companyDetailData && companyDetailData.funnel) {
    companyCrossData = companyDetailData;
    renderCrossEntityFunnel(companyDetailData);
  }
  if (cb) cb();
}

function renderCrossEntityFunnel(d) {
  var el = document.getElementById('companyCrossContent');
  if (!el) return;
  var f = d.funnel;
  if (!f) { el.innerHTML = ''; return; }
  var segs = f.segments;
  var h = '';
  h += dateFilterNotice();

  // -- Inline category filter bar (always visible) --
  h += '<div class="cat-filter-bar" id="companyCatFilterBar">';
  h += '<span style="font-weight:500;font-size:.85rem;color:var(--so-charcoal)">Category filter:</span>';
  h += '<select id="companyDetailCatFilter" class="cat-filter-select" onchange="onDetailFilterChange(\'company\')">';
  h += '<option value="">All categories</option>';
  h += '</select>';
  h += '<span class="filter-count" id="companyDetailFilterCount" style="font-size:.82rem;color:#666"></span>';
  h += '<span class="filter-reset-link" id="companyFilterReset" style="display:none" onclick="resetDetailFilter(\'company\')">Reset</span>';
  h += '</div>';

  // -- CRM Health Pipeline — combined funnel with counts, %, and conversion --
  var pipeLabel = (typeof daSettings !== 'undefined') ? daSettings.getPipelineLabel('company') : 'Open Sale';
  var pipeType = (typeof daSettings !== 'undefined') ? daSettings.getPipelineType('company') : 'sale';
  var showPipeline = pipeType !== 'none';

  var funnelSteps = [
    { label: 'Total Companies', count: f.total, from: f.total, desc: 'All companies in database' },
    { label: 'With Contact Person', count: f.withPerson, from: f.total, desc: 'Has at least one linked person' },
    { label: 'With Activity (12m)', count: f.withPersonActivity, from: f.withPerson, desc: 'Person + recent activity logged' }
  ];
  if (showPipeline) {
    funnelSteps.push({ label: 'With ' + pipeLabel, count: f.withPersonActivitySale, from: f.withPersonActivity, desc: 'Active + ' + pipeLabel.toLowerCase() });
  }

  h += '<div class="entity-card">';
  h += '<div class="entity-header"><div class="entity-info"><h3>CRM Health Pipeline</h3></div>';
  h += '<span class="record-badge">' + fmtNum(f.total) + ' companies</span></div>';
  h += '<table class="data-table"><thead><tr>';
  h += '<th>Stage</th><th class="col-right" style="width:80px">Count</th><th class="col-right" style="width:80px">% of Total</th><th class="col-right" style="width:90px">Conversion</th><th class="col-right" style="width:130px">Completeness</th>';
  h += '</tr></thead><tbody>';
  for (var i = 0; i < funnelSteps.length; i++) {
    var step = funnelSteps[i];
    var pctTotal = f.total > 0 ? Math.round(step.count / f.total * 1000) / 10 : 0;
    var conv = (i === 0) ? '-' : (step.from > 0 ? Math.round(step.count / step.from * 1000) / 10 + P : '0' + P);
    var col = slColor(pctTotal);
    h += '<tr>';
    h += '<td><span style="font-weight:500">' + step.label + '</span>';
    h += '<div style="font-size:.75rem;color:var(--so-text-muted)">' + step.desc + '</div></td>';
    h += '<td class="col-right">' + fmtNum(step.count) + '</td>';
    h += '<td class="col-right">' + pctTotal + P + '</td>';
    h += '<td class="col-right">' + conv + '</td>';
    h += '<td class="col-right">' + barCell(pctTotal, col) + '</td>';
    h += '</tr>';
  }
  h += '</tbody></table></div>';

  // -- Segments (description middle, bar+% right) --
  h += '<div class="entity-card">';
  h += '<div class="entity-header"><div class="entity-info"><h3>Company Segments</h3></div>';
  h += '<span class="record-badge">4 segments</span></div>';
  h += '<table class="data-table"><thead><tr>';
  h += '<th>Segment</th><th class="col-right">Companies</th><th>Description</th><th class="col-right">' + P + '</th>';
  h += '</tr></thead><tbody>';
  for (var si = 0; si < segs.length; si++) {
    var seg = segs[si];
    var segPct = f.total > 0 ? Math.round(seg.count / f.total * 1000) / 10 : 0;
    h += '<tr>';
    h += '<td><span style="display:inline-block;width:10px;height:10px;border-radius:50%;background:' + hcEngColor(seg.name) + ';margin-right:8px;vertical-align:middle"></span>';
    h += '<span style="font-weight:500">' + seg.name + '</span></td>';
    h += '<td class="col-right">' + fmtNum(seg.count) + '</td>';
    h += '<td style="color:var(--so-text-muted);font-size:.82rem">' + seg.description + '</td>';
    h += '<td class="col-right">' + barCell(segPct, hcEngColor(seg.name)) + '</td>';
    h += '</tr>';
  }
  h += '</tbody></table></div>';

  el.innerHTML = h;
  populateDetailCatFilter();
}

// ===========================================================
// ENTITY SCORES — 3 scores + Overall Health (computed frontend)
// ===========================================================

/**
 * gatherEntityData(key) — collects all available data for score computation
 * Maps entity-specific raw data to generic check/component keys used by settings
 */
function gatherEntityData(key) {
  var ov = overviewData[key];
  if (!ov) return null;
  // Only compute scores for entities with settings definitions
  if (typeof daSettings !== 'undefined' && !daSettings.ENTITY_DEFS[key]) return null;
  var o = ov.overview;
  var total = o.total || 0;
  var cpl = ov.completeness || null;
  var ec = entityConfig[key];
  var ud = (ec && ec.udefId > 0 && udefData[ec.udefId]) ? udefData[ec.udefId] : null;
  var checkData = {};
  var compData = {};
  var qData = null;
  var adoptionTotal = total;

  if (key === 'company') {
    var q = (companyDetailData && companyDetailData.quality) ? companyDetailData.quality : null;
    var ah = (companyDetailData && companyDetailData.activityHealth) ? companyDetailData.activityHealth : null;
    var f = (companyDetailData && companyDetailData.funnel) ? companyDetailData.funnel : null;
    qData = q;
    if (q) {
      checkData.noPerson = q.noPerson;
      checkData.unreachable = q.unreachable;
      checkData.noOwner = q.noOwner;
      checkData.noCategory = q.noCategory;
      checkData.noBusiness = q.noBusiness;
      checkData.noOrgNr = q.noOrgNr;
    }
    if (ah) { checkData.noActivity12m = ah.noActivity; }
    if (f) {
      compData.withPerson = f.withPerson;
      compData.withActivity = f.withPersonActivity;
      var pt = (typeof daSettings !== 'undefined') ? daSettings.getPipelineType('company') : 'sale';
      if (pt !== 'none') compData.withPipeline = f.withPersonActivitySale;
      adoptionTotal = f.total || total;
    }
  } else if (key === 'contact') {
    // Derive from overview stats
    if (o.withEmail !== undefined) checkData.noEmail = total - (o.withEmail || 0);
    if (o.withCompany !== undefined) checkData.noCompany = total - (o.withCompany || 0);
    if (o.withActivities !== undefined) checkData.noActivity = total - (o.withActivities || 0);
    compData.withActivity = o.withActivities || 0;
    compData.withSales = o.withSales || 0;
    // Build completeness from overview if not from server
    if (!cpl) {
      cpl = { email: o.withEmail || 0, phone: o.withPhone || 0, position: o.withPosition || 0, mrMrs: o.withTitle || 0 };
      cpl.firstName = total; cpl.lastName = total; // assume always filled
    }
  } else if (key === 'sale') {
    if (o.withPersons !== undefined) checkData.noContact = total - (o.withPersons || 0);
    if (o.staleSale !== undefined) checkData.staleSale = o.staleSale;
    if (o.withActivities !== undefined) checkData.noActivities = total - (o.withActivities || 0);
    if (o.withAmount !== undefined) checkData.noAmount = total - (o.withAmount || 0);
    compData.withActivity = o.withActivities || 0;
    // Real stage adoption: sales with a pipeline stage (probability) set. Read the
    // "(No value)" bucket from the stage distribution so this matches the Overview tab.
    var saleNoStage = 0;
    if (ov.distributions) {
      for (var sdi = 0; sdi < ov.distributions.length; sdi++) {
        var sdd = ov.distributions[sdi];
        if (sdd && sdd.title && sdd.title.toLowerCase().indexOf('stage') >= 0 && sdd.items) {
          for (var sdj = 0; sdj < sdd.items.length; sdj++) {
            if ((sdd.items[sdj].name || '').toLowerCase() === '(no value)') saleNoStage = sdd.items[sdj].count || 0;
          }
        }
      }
    }
    compData.stageProgression = total - saleNoStage;
    if (!cpl) {
      cpl = { amount: o.withAmount || 0, saleType: total, stage: total, probability: total, closeDate: total };
    }
  } else if (key === 'project') {
    if (o.withMembers !== undefined) checkData.noMembers = total - (o.withMembers || 0);
    if (o.withActivities !== undefined) checkData.noActivities = total - (o.withActivities || 0);
    compData.withActivity = o.withActivities || 0;
    compData.memberEngagement = o.withMembers || 0;
    if (!cpl) {
      cpl = { projectType: total, status: total, endDate: total };
    }
  } else if (key === 'requests') {
    // Adoption: resolution rate (closed) and engagement (tickets replied to).
    // Integrity: unassigned tickets (owned_by = 1 system user). No DQ model yet,
    // so completeness stays null and the ticket-fields tab is left untouched.
    if (o.unassigned !== undefined) checkData.noOwner = o.unassigned;
    compData.resolutionRate = o.closed || 0;
    compData.engagement = o.engaged || 0;
  }

  return { total: total, completeness: cpl, qualityData: qData, udefData: ud, checkData: checkData, componentData: compData, adoptionTotal: adoptionTotal, integrityTotal: adoptionTotal };
}

/**
 * computeEntityScores(key) — computes all 4 scores for an entity
 * Returns { dq, integrity, adoption, health } or null
 */
function computeEntityScores(key) {
  if (typeof daSettings === 'undefined') return null;
  var data = gatherEntityData(key);
  if (!data) return null;

  var dq = daSettings.computeDQScore(key, data.completeness, data.qualityData, data.udefData, data.total);
  var integrity = daSettings.computeIntegrity(key, data.checkData, data.integrityTotal);
  var adoption = daSettings.computeAdoption(key, data.componentData, data.adoptionTotal);
  var health = daSettings.computeHealth(key, dq, integrity, adoption);

  return { dq: dq, integrity: integrity, adoption: adoption, health: health };
}

// Cache scores per entity so Overview + DQ tabs both can use them
var entityScores = {};

function refreshEntityScores(key) {
  entityScores[key] = computeEntityScores(key);
  return entityScores[key];
}

/**
 * renderScoreBanner(key) — Overall Health ring left, 3 sub-score rows right
 */
var sbDescs = {
  company:  { dq: 'Field completeness across records', int: 'Missing contacts and reachability', adopt: 'Contact coverage, activities and pipeline' },
  contact:  { dq: 'Field completeness across records', int: 'Missing email and activity gaps', adopt: 'Activities and sales involvement' },
  sale:     { dq: 'Field completeness across records', int: 'Missing contacts and activities', adopt: 'Activities and stage progression' },
  project:  { dq: 'Field completeness across records', int: 'Missing members and activities', adopt: 'Activities and member engagement' },
  requests: { dq: 'Field completeness across records', int: 'Unassigned tickets without an owner', adopt: 'Resolution rate and replies to customers' }
};

var sbTips = {
  health: 'Weighted average of Data Quality, Data Integrity and Adoption, using the weights set in the dashboard settings.',
  dq: 'Weighted completeness of the fields you marked as important. Required fields count double, Normal once, Off is ignored.',
  int: 'Share of records free from structural problems such as missing links or missing owners. 100 means no records are affected.',
  adopt: 'Share of records showing real CRM usage, such as activities and pipeline. For requests it reflects resolution and replies.'
};

function _capFirst(s) { return s ? s.charAt(0).toUpperCase() + s.slice(1) : s; }

// Auto-generated verdict sentence for the score banner.
// Stays factual: names the overall level, the strongest area and the main gap.
// A manual override field is planned on top of this later.
function sbVerdict(key, scores) {
  if (!scores || !scores.health) return '';
  var label = _capFirst(key);
  var subs = [];
  if (scores.dq) subs.push({ k: 'field completeness', t: scores.dq.total });
  if (scores.integrity) subs.push({ k: 'data integrity', t: scores.integrity.total });
  if (scores.adoption) subs.push({ k: 'adoption', t: scores.adoption.total });
  if (subs.length === 0) return '';
  subs.sort(function(a, b) { return b.t - a.t; });
  var best = subs[0], worst = subs[subs.length - 1];
  var parts = [label + ' data is ' + slHealthWord(scores.health.total) + ' overall.'];
  if (best.t >= 70) parts.push('<b>' + _capFirst(best.k) + ' is strong.</b>');
  if (worst.t < 70 && worst.k !== best.k) parts.push('<b>' + _capFirst(worst.k) + '</b> is the main gap and offers the clearest wins.');
  return parts.join(' ');
}

function sbFoot(key, scores) {
  var h = '<div class="sb-foot">';
  var v = sbVerdict(key, scores);
  if (v) h += '<div class="sb-verdict">' + v + '</div>';
  h += slKeyHtml();
  h += '</div>';
  return h;
}

function renderScoreBanner(key) {
  var el = document.getElementById(key + 'OverviewContent');
  if (!el) return;
  var scores = refreshEntityScores(key);
  if (!scores) return;
  var existing = el.querySelector('.scores-banner');
  if (existing) existing.parentNode.removeChild(existing);
  if (!scores.health && !scores.dq && !scores.integrity && !scores.adoption) return;
  var desc = sbDescs[key] || sbDescs.company;

  var subs = [];
  if (scores.dq) subs.push({ k: 'dq', label: 'Data Quality', desc: desc.dq, score: scores.dq.total, tip: sbTips.dq });
  if (scores.integrity) subs.push({ k: 'integrity', label: 'Data Integrity', desc: desc.int, score: scores.integrity.total, tip: sbTips.int });
  if (scores.adoption) subs.push({ k: 'adoption', label: 'Adoption', desc: desc.adopt, score: scores.adoption.total, tip: sbTips.adopt });

  var comps = (scores.health && scores.health.components) ? scores.health.components : [];
  var HPTS = { high: 3, medium: 2, low: 1 };
  var totalPts = 0;
  for (var p = 0; p < comps.length; p++) totalPts += (HPTS[comps[p].weight] || 2);
  function wpctFor(k) { for (var q = 0; q < comps.length; q++) { if (comps[q].key === k) return totalPts > 0 ? Math.round((HPTS[comps[q].weight] || 2) / totalPts * 100) : 0; } return 0; }

  var ht = scores.health ? scores.health.total : (subs.length ? subs[0].score : 0);
  var hbnd = slBand(ht);
  var h = '<div class="scores-banner card hero">';
  h += '<div class="donut-wrap">';
  h += '<div class="donut" style="background:conic-gradient(from -90deg, ' + slColor(ht) + ' 0 ' + ht + '%, var(--track) ' + ht + '% 100%)"><span class="val">' + ht + '<sup>%</sup></span></div>';
  h += '<div class="ttl">Overall Health' + sbInfoIcon(sbTips.health) + '</div>';
  h += '<span class="badge" style="background:' + hbnd.bg + ';color:' + hbnd.fg + '"><span class="dot" style="background:' + hbnd.dot + '"></span> ' + hbnd.w + '</span>';
  h += '</div>';
  h += '<div class="subs">';
  for (var i = 0; i < subs.length; i++) {
    var s = subs[i];
    var col = slColor(s.score);
    var bnd = slBand(s.score);
    var wp = wpctFor(s.k);
    var wc = wp > 0 ? '<span class="w">weighs ' + wp + '%</span>' : '';
    h += '<div class="subrow">';
    h += '<div class="lab"><b>' + s.label + '</b>' + wc + sbInfoIcon(s.tip) + '<p>' + s.desc + '</p></div>';
    h += '<span class="badge" style="background:' + bnd.bg + ';color:' + bnd.fg + '"><span class="dot" style="background:' + bnd.dot + '"></span> ' + bnd.w + '</span>';
    h += '<div class="barwrap"><div class="bar"><i style="width:' + s.score + '%;background:' + col + '"></i></div><span class="pct" style="color:' + col + '">' + s.score + '%</span></div>';
    h += '</div>';
  }
  h += '</div>';
  var verdict = sbVerdict(key, scores);
  h += '<div class="hero-foot" style="grid-column:1 / -1">';
  if (verdict) h += '<span class="summary">' + verdict + '</span>';
  h += '<span class="note">Overall = weighted average of the three \u00b7 measured against your own data</span>';
  h += '</div>';
  h += '<div style="grid-column:1 / -1; display:flex; align-items:center; gap:14px; margin-top:12px">';
  h += '<span style="color:var(--faint); font-size:12px; font-weight:700; letter-spacing:.06em; text-transform:uppercase">Score bands</span>';
  h += '<div class="legend"><span class="lg"><i style="background:var(--bad)"></i> &lt;15</span><span class="lg"><i style="background:var(--mod)"></i> 15\u201339</span><span class="lg"><i style="background:#c9b03a"></i> 40\u201369</span><span class="lg"><i style="background:var(--good)"></i> 70+ Good</span></div>';
  h += '</div>';
  h += '</div>';
  var fn = el.querySelector('.filter-notice');
  if (fn) fn.insertAdjacentHTML('afterend', h);
  else el.insertAdjacentHTML('afterbegin', h);
}

// ===========================================================
// DATA QUALITY SCORE (computed frontend from available data)
// ===========================================================
function renderDQScore(key) {
  var el = document.getElementById(key + 'DQScoreContent');
  if (!el) return;

  var scores = [];
  var h = '';

  // Company-specific: Issues + Completeness side by side as tables
  var hasIssues = (key === 'company' && companyDetailData && companyDetailData.quality);
  var hasCpl = (key === 'company' && overviewData['company'] && overviewData['company'].completeness);
  if (hasIssues || hasCpl) {
    var cols = (hasIssues && hasCpl) ? 2 : 1;
    if (cols === 2) h += '<div style="display:grid;grid-template-columns:1fr 1fr;gap:18px;margin-bottom:18px">';

    if (hasIssues) {
      var q = companyDetailData.quality;
      var total = companyDetailData.activityHealth ? companyDetailData.activityHealth.total : 0;
      if (total > 0) {
        // Use settings-driven quality issue fields
        var qiFields = (typeof daSettings !== 'undefined') ? daSettings.getSettings('company').qualityIssueFields : ['noPerson', 'noCategory', 'noBusiness', 'noOrgNr', 'unreachable'];
        var qiOptions = (typeof daSettings !== 'undefined') ? daSettings.QUALITY_ISSUE_OPTIONS : null;

        var allIssues = [
          { key: 'noPerson',    label: 'No contact person', val: q.noPerson },
          { key: 'noCategory',  label: 'No category', val: q.noCategory },
          { key: 'noBusiness',  label: 'No business type', val: q.noBusiness },
          { key: 'noOrgNr',     label: 'No org. number', val: q.noOrgNr },
          { key: 'unreachable', label: 'Unreachable', val: q.unreachable },
          { key: 'noOwner',     label: 'No owner', val: q.noOwner || 0 }
        ];

        h += '<div class="entity-card">';
        h += '<div class="entity-header"><div class="entity-info"><h3>Quality flags</h3></div>';
        h += '<span class="record-badge">' + fmtNum(total) + ' companies</span></div>';
        h += '<div class="tbl-cap"><b>Lower is better.</b> Each flag counts against the score. Sorted by attention, so the flags hitting the most records sit on top; flags not in the score drop to the bottom.</div>';
        h += '<table class="data-table"><thead><tr><th class="attn-col"></th><th>Flag</th><th class="col-right" style="width:80px">Count</th><th class="col-right" style="width:130px">' + P + '</th></tr></thead><tbody>';
        var qRows = [];
        for (var i = 0; i < allIssues.length; i++) {
          var iss = allIssues[i];
          var isActive = qiFields.indexOf(iss.key) >= 0;
          var pct = Math.round((iss.val / total) * 1000) / 10;
          var col = slColorInv(pct);
          var attn = isActive ? Math.round(pct) : -1;
          var reason = '<b>' + attnBandWord(attn) + ' attention.</b> Active quality flag, ' + Math.round(pct) + '% of companies affected. Fewer affected lifts the score.';
          var dimStyle = isActive ? '' : ' style="opacity:0.4"';
          var badge = isActive ? '' : ' <span style="font-size:.65rem;color:var(--so-text-muted)">(not in score)</span>';
          var row = '<tr' + dimStyle + '><td class="attn-col">' + attnIconHtml(attn, reason) + '</td><td>' + iss.label + badge + '</td>';
          row += '<td class="col-right">' + fmtNum(iss.val) + '</td>';
          row += '<td class="col-right">' + barCell(pct, col) + '</td></tr>';
          qRows.push({ attn: attn, badness: pct, html: row });
        }
        h += attnSortRows(qRows);
        h += '</tbody></table></div>';
      }
    }

    if (hasCpl) {
      var ov = overviewData['company'];
      var c = ov.completeness;
      var total = ov.overview.total || 0;
      var q = (companyDetailData && companyDetailData.quality) ? companyDetailData.quality : null;

      // Show ALL stdFields with importance badge
      if (total > 0 && typeof daSettings !== 'undefined') {
        var def = daSettings.ENTITY_DEFS['company'];
        var cfg = daSettings.getSettings('company');
        h += '<div class="entity-card">';
        h += '<div class="entity-header"><div class="entity-info"><h3>Standard Field Completeness</h3></div>';
        h += '<span class="record-badge">' + fmtNum(total) + ' companies</span></div>';
        h += '<div class="tbl-cap"><b>Higher completeness is better.</b> Drives the Data Quality score. Sorted by attention, so required fields with the largest gaps sit on top.</div>';
        h += '<table class="data-table"><thead><tr><th class="attn-col"></th><th>Field</th><th class="col-right" style="width:80px">Filled</th><th style="text-align:center;width:90px">Importance</th><th class="col-right" style="width:130px">Completeness</th></tr></thead><tbody>';
        var cRows = [];
        for (var j = 0; j < def.stdFields.length; j++) {
          var fk = def.stdFields[j].key;
          var label = def.stdFields[j].label;
          var imp = cfg.stdFieldConfig[fk] || 'excluded';
          var val = daSettings.getCompletenessValue(fk, c, q, total);
          var pct = Math.round((val / total) * 1000) / 10;
          var badness = 100 - pct;
          var attn = attnScore(badness, imp);
          var reason = '<b>' + attnBandWord(attn) + ' attention.</b> ' + _capFirst(imp) + ' field, ' + Math.round(badness) + '% empty. Higher completeness lifts the Data Quality score.';
          var dimStyle = imp === 'excluded' ? ' style="opacity:0.4"' : '';
          var badgeHtml = '<span class="imp-badge ' + imp + '">' + imp + '</span>';
          var row = '<tr' + dimStyle + '><td class="attn-col">' + attnIconHtml(attn, reason) + '</td><td>' + label + '</td>';
          row += '<td class="col-right">' + fmtNum(val) + '</td>';
          row += '<td style="text-align:center">' + badgeHtml + '</td>';
          row += '<td class="col-right">' + barCell(pct, '') + '</td></tr>';
          cRows.push({ attn: attn, badness: badness, html: row });
        }
        h += attnSortRows(cRows);
        h += '</tbody></table></div>';
      }
    }

    if (cols === 2) h += '</div>';
  }

  // Standard Field Completeness for ALL entities (non-company)
  if (key !== 'company' && typeof daSettings !== 'undefined') {
    var def = daSettings.ENTITY_DEFS[key];
    var cfg = daSettings.getSettings(key);
    var data = gatherEntityData(key);
    if (def && cfg && data && data.completeness && data.total > 0) {
      var cpl = data.completeness;
      var eTotal = data.total;
      var eLabel = def.label;
      h += '<div class="entity-card">';
      h += '<div class="entity-header"><div class="entity-info"><h3>Standard Field Completeness</h3></div>';
      h += '<span class="record-badge">' + fmtNum(eTotal) + ' ' + eLabel.toLowerCase() + 's</span></div>';
      h += '<div class="tbl-cap"><b>Higher completeness is better.</b> Drives the Data Quality score. Sorted by attention, so required fields with the largest gaps sit on top.</div>';
      h += '<table class="data-table"><thead><tr><th class="attn-col"></th><th>Field</th><th class="col-right" style="width:80px">Filled</th><th style="text-align:center;width:90px">Importance</th><th class="col-right" style="width:130px">Completeness</th></tr></thead><tbody>';
      var eRows = [];
      for (var fi = 0; fi < def.stdFields.length; fi++) {
        var fk = def.stdFields[fi].key;
        var fLabel = def.stdFields[fi].label;
        var imp = cfg.stdFieldConfig[fk] || 'excluded';
        var fVal = cpl[fk] || 0;
        var fPct = Math.round((fVal / eTotal) * 1000) / 10;
        var fBad = 100 - fPct;
        var fAttn = attnScore(fBad, imp);
        var fReason = '<b>' + attnBandWord(fAttn) + ' attention.</b> ' + _capFirst(imp) + ' field, ' + Math.round(fBad) + '% empty. Higher completeness lifts the Data Quality score.';
        var dimStyle = imp === 'excluded' ? ' style="opacity:0.4"' : '';
        var badgeHtml = '<span class="imp-badge ' + imp + '">' + imp + '</span>';
        var eRow = '<tr' + dimStyle + '><td class="attn-col">' + attnIconHtml(fAttn, fReason) + '</td><td>' + fLabel + '</td>';
        eRow += '<td class="col-right">' + fmtNum(fVal) + '</td>';
        eRow += '<td style="text-align:center">' + badgeHtml + '</td>';
        eRow += '<td class="col-right">' + barCell(fPct, '') + '</td></tr>';
        eRows.push({ attn: fAttn, badness: fBad, html: eRow });
      }
      h += attnSortRows(eRows);
      h += '</tbody></table></div>';
    }
  }

  h = dateFilterNotice() + h; // filter notice always at very top

  el.innerHTML = h;
}

function computeDQScore(key) {
  // Use settings-driven computation for all entities
  if (typeof daSettings !== 'undefined') {
    var data = gatherEntityData(key);
    if (data) {
      var result = daSettings.computeDQScore(key, data.completeness, data.qualityData, data.udefData, data.total);
      if (result) return result.total;
    }
  }

  // Fallback: simple average of available completeness data
  var ov = overviewData[key];
  if (!ov || !ov.completeness) return null;
  var c = ov.completeness; var total = (ov.overview && ov.overview.total) ? ov.overview.total : 0;
  if (total <= 0) return null;
  var sum = 0; var cnt = 0;
  for (var k in c) { if (c.hasOwnProperty(k) && typeof c[k] === 'number') { sum += c[k] / total * 100; cnt++; } }
  return cnt > 0 ? Math.round(sum / cnt) : null;
}

// end of DQ score functions

// ===========================================================
// ADOPTION TAB — score breakdown for all entities
// ===========================================================
// Category filter bar for the Integrity tab (company only). Shares state and
// reload with the Adoption tab bar via companyDetailCatValue.
function integrityCatFilterBar() {
  var h = '<div class="cat-filter-bar" id="companyIntegrityCatFilterBar">';
  h += '<span style="font-weight:500;font-size:.85rem;color:var(--so-charcoal)">Category filter:</span>';
  h += '<select id="companyIntegrityDetailCatFilter" class="cat-filter-select" onchange="onDetailFilterChange(\'companyIntegrity\')">';
  h += '<option value="">All categories</option>';
  h += '</select>';
  h += '<span class="filter-count" id="companyIntegrityDetailFilterCount" style="font-size:.82rem;color:#666"></span>';
  h += '<span class="filter-reset-link" id="companyIntegrityFilterReset" style="display:none" onclick="resetDetailFilter(\'companyIntegrity\')">Reset</span>';
  h += '</div>';
  return h;
}

function renderAdoptionTab(key) {
  var scores = entityScores[key];
  if (!scores) return;

  var ov = overviewData[key];
  var total = (ov && ov.overview) ? (ov.overview.total || 0) : 0;
  var hAd = '';
  var hInt = '';

  // --- Adoption component breakdown ---
  if (scores.adoption && scores.adoption.details && scores.adoption.details.length > 0) {
    var ad = scores.adoption;
    var ac = slColor(ad.total);
    var data = gatherEntityData(key);
    hAd += '<div class="entity-card">';
    hAd += '<div class="entity-header"><div class="entity-info">';
    hAd += '<h3>Adoption Score</h3>';
    hAd += '</div>';
    hAd += '<span class="record-badge" style="background:' + ac + ';color:#fff;font-size:1.1rem;padding:4px 14px;border-radius:12px">' + ad.total + P + '</span></div>';
    hAd += '<div class="tbl-cap"><b>Higher coverage is better.</b> Drives the Adoption score. Sorted by attention, so heavily weighted components with the largest gaps sit on top.</div>';
    hAd += '<table class="data-table"><thead><tr><th class="attn-col"></th><th>Component</th><th class="col-right" style="width:80px">Count</th><th style="text-align:center;width:90px">Weight</th><th class="col-right" style="width:130px">Completeness</th></tr></thead><tbody>';
    var adRows = [];
    for (var i = 0; i < ad.details.length; i++) {
      var d = ad.details[i];
      var cnt = (data && data.componentData) ? (data.componentData[d.key] || 0) : 0;
      var col = slColor(d.pct);
      var wCls = d.weight === 'high' ? 'required' : (d.weight === 'medium' ? 'normal' : 'excluded');
      var adBad = 100 - d.pct;
      var adAttn = attnScore(adBad, d.weight);
      var adReason = '<b>' + attnBandWord(adAttn) + ' attention.</b> ' + _capFirst(d.weight) + ' weight, ' + Math.round(adBad) + '% of records not covered. Higher coverage lifts Adoption.';
      var adRow = '<tr><td class="attn-col">' + attnIconHtml(adAttn, adReason) + '</td><td>' + d.label + (d.desc ? sbInfoIcon(d.desc) : '') + '</td>';
      adRow += '<td class="col-right">' + fmtNum(cnt) + '</td>';
      adRow += '<td style="text-align:center"><span class="imp-badge ' + wCls + '">' + d.weight + '</span></td>';
      adRow += '<td class="col-right">' + barCell(d.pct, col) + '</td></tr>';
      adRows.push({ attn: adAttn, badness: adBad, html: adRow });
    }
    hAd += attnSortRows(adRows);
    hAd += '</tbody></table></div>';
  }

  // --- Integrity check breakdown ---
  if (scores.integrity && scores.integrity.details && scores.integrity.details.length > 0) {
    var ig = scores.integrity;
    var ic = slColor(ig.total);
    var data2 = gatherEntityData(key);
    hInt += '<div class="entity-card">';
    hInt += '<div class="entity-header"><div class="entity-info">';
    hInt += '<h3>Data Integrity</h3>';
    hInt += '</div>';
    hInt += '<span class="record-badge" style="background:' + ic + ';color:#fff;font-size:1.1rem;padding:4px 14px;border-radius:12px">' + ig.total + P + '</span></div>';
    hInt += '<div class="tbl-cap"><b>Lower is better.</b> Drives the Data Integrity score. Sorted by attention, so heavily weighted checks affecting the most records sit on top.</div>';
    hInt += '<table class="data-table"><thead><tr><th class="attn-col"></th><th>Check</th><th class="col-right" style="width:80px">Affected</th><th style="text-align:center;width:90px">Weight</th><th class="col-right" style="width:130px">% Affected</th></tr></thead><tbody>';
    var igRows = [];
    for (var j = 0; j < ig.details.length; j++) {
      var c = ig.details[j];
      var cnt2 = (data2 && data2.checkData) ? (data2.checkData[c.key] || 0) : 0;
      var col2 = slColorInv(c.affected);
      var iTotal = (data2 && data2.integrityTotal) ? data2.integrityTotal : total;
      var wCls2 = c.weight === 'high' ? 'required' : (c.weight === 'medium' ? 'normal' : 'excluded');
      var igAttn = attnScore(c.affected, c.weight);
      var igReason = '<b>' + attnBandWord(igAttn) + ' attention.</b> ' + _capFirst(c.weight) + ' weight, ' + Math.round(c.affected) + '% of records affected. Fewer affected lifts Data Integrity.';
      var igRow = '<tr><td class="attn-col">' + attnIconHtml(igAttn, igReason) + '</td><td>' + c.label + (c.desc ? sbInfoIcon(c.desc) : '') + '</td>';
      igRow += '<td class="col-right">' + fmtNum(cnt2) + '</td>';
      igRow += '<td style="text-align:center"><span class="imp-badge ' + wCls2 + '">' + c.weight + '</span></td>';
      igRow += '<td class="col-right">' + barCell(c.affected, col2) + '</td></tr>';
      igRows.push({ attn: igAttn, badness: c.affected, html: igRow });
    }
    hInt += attnSortRows(igRows);

    // --- Rule checks (from intake conditional rules) ---
    var ruleResults = (window.rulesData && window.rulesData[key]) ? window.rulesData[key] : [];
    if (ruleResults.length > 0) {
      hInt += '<tr><td colspan="5" style="padding:12px 8px 8px;font-size:12px;font-weight:600;letter-spacing:.02em;color:var(--so-green);border-bottom:1px solid var(--so-border);border-top:2px solid var(--so-border)"><span style="display:inline-flex;align-items:center;gap:6px"><svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><line x1="6" y1="3" x2="6" y2="15"/><circle cx="18" cy="6" r="3"/><circle cx="6" cy="18" r="3"/><path d="M18 9a9 9 0 0 1-9 9"/></svg>Rule checks &middot; ' + ruleResults.length + ' evaluated</span></td></tr>';
      var rRows = [];
      for (var ri = 0; ri < ruleResults.length; ri++) {
        var rv = ruleResults[ri];
        var rvPct = (rv.segmentTotal > 0) ? Math.round(rv.violations / rv.segmentTotal * 100 * 10) / 10 : 0;
        var rvCol = slColorInv(rvPct);
        var rvW = rv.weight || 'high';
        var rvWCls = rvW === 'high' ? 'required' : (rvW === 'medium' ? 'normal' : 'excluded');
        var rvAttn = attnScore(rvPct, rvW);
        var rvTip = 'Rule: for category <b>' + (rv.condLabel || '?') + '</b>, field <b>' + (rv.reqLabel || '?') + '</b> must be filled.';
        if (rv.segmentTotal > 0) rvTip += '<br>' + rv.violations + ' of ' + fmtNum(rv.segmentTotal) + ' records in this segment are missing the field.';
        if (rv.ownerLabel) rvTip += '<br><span style="display:inline-flex;align-items:center;gap:4px;margin-top:4px;font-size:11px;font-weight:600;color:var(--so-green-light);background:var(--so-meadow);padding:2px 8px;border-radius:4px">' + rv.ownerLabel + '</span>';
        else rvTip += '<br><span style="display:inline-flex;align-items:center;gap:4px;margin-top:4px;font-size:11px;color:var(--so-text-muted)">No owner assigned</span>';
        var rvReason = '<b>' + attnBandWord(rvAttn) + ' attention.</b> ' + _capFirst(rvW) + ' weight, ' + rvPct + '% of segment affected.';
        var rvRow = '<tr><td class="attn-col">' + attnIconHtml(rvAttn, rvReason) + '</td>';
        rvRow += '<td>' + (rv.condLabel || '?') + ' &rarr; ' + (rv.reqLabel || '?') + sbInfoIcon(rvTip) + '</td>';
        rvRow += '<td class="col-right">' + fmtNum(rv.violations || 0) + '</td>';
        rvRow += '<td style="text-align:center"><span class="imp-badge ' + rvWCls + '">' + rvW + '</span></td>';
        rvRow += '<td class="col-right">' + barCell(rvPct, rvCol) + '</td></tr>';
        rRows.push({ attn: rvAttn, badness: rvPct, html: rvRow });
      }
      hInt += attnSortRows(rRows);
    }

    hInt += '</tbody></table></div>';
  }

  // --- Adoption placement (stays on the Adoption tab) ---
  if (key === 'company') {
    // Insert adoption table AFTER the filter elements in the cross content
    var crossEl = document.getElementById('companyCrossContent');
    if (crossEl) {
      var prev = crossEl.querySelector('.adoption-score-tables');
      if (prev) prev.parentNode.removeChild(prev);
      if (hAd) {
        var wrapper = document.createElement('div');
        wrapper.className = 'adoption-score-tables';
        wrapper.style.marginBottom = '18px';
        wrapper.innerHTML = hAd;
        var catBar = crossEl.querySelector('.cat-filter-bar');
        var filterNotice = crossEl.querySelector('.filter-notice');
        var insertAfter = catBar || filterNotice;
        if (insertAfter) {
          crossEl.insertBefore(wrapper, insertAfter.nextSibling);
        } else {
          crossEl.insertBefore(wrapper, crossEl.firstChild);
        }
      }
    }
  } else {
    var el = document.getElementById(key + 'AdoptionContent');
    if (el) el.innerHTML = dateFilterNotice() + hAd;
  }

  // --- Integrity placement (its own tab) ---
  var igEl = document.getElementById(key + 'IntegrityContent');
  if (igEl) {
    if (hInt) {
      var igFilter = (key === 'company') ? integrityCatFilterBar() : '';
      igEl.innerHTML = dateFilterNotice() + igFilter + hInt;
    } else {
      igEl.innerHTML = '';
    }
  }
  // Populate/sync both category bars now that the integrity bar exists.
  if (key === 'company') populateDetailCatFilter();
}

function renderCompanyDetails(d) {
  var el = document.getElementById('companyDetailContent');
  if (!el) return;
  var h = '';
  // Filter notice is shown at top of Adoption tab (in renderCrossEntityFunnel)
  var ah = d.activityHealth;
  var total = ah ? ah.total : 0;

  // 1. ACTIVITY RECENCY — compact stacked bar (replaces standalone Activity Health)
  var ah = d.activityHealth;
  if (ah && total > 0) {
    h += '<div class="detail-section">';
    h += '<div class="detail-section-head">' + secHead('Activity Recency') + '<span class="record-badge">' + fmtNum(total) + ' companies</span></div>';
    var ahParts = [
      { val: ah.active6m, label: 'Active (6m)' },
      { val: ah.dormant12m, label: 'Cooling (6\u201312m)' },
      { val: ah.dormantOlder, label: 'Dormant (>12m)' },
      { val: ah.noActivity, label: 'No Activity' }
    ];
    for (var ci = 0; ci < ahParts.length; ci++) ahParts[ci].col = hcEngColor(ahParts[ci].label);
    h += '<div class="engbar" style="height:28px">';
    for (var i = 0; i < ahParts.length; i++) {
      var pct = ahParts[i].val / total * 100;
      if (pct > 0) {
        var segLabel = pct >= 10 ? fmtNum(ahParts[i].val) + ' (' + (Math.round(pct * 10) / 10) + P + ')' : '';
        h += '<span style="width:' + pct + P + ';background:' + ahParts[i].col + '">' + segLabel + '</span>';
      }
    }
    h += '</div>';
    h += '<div style="display:flex;gap:16px;margin-top:6px;flex-wrap:wrap">';
    for (var i = 0; i < ahParts.length; i++) {
      var pct = Math.round(ahParts[i].val / total * 1000) / 10;
      h += '<span style="font-size:.75rem;color:#666;display:flex;align-items:center;gap:4px">';
      h += '<span style="width:8px;height:8px;border-radius:50%;background:' + ahParts[i].col + ';flex-shrink:0"></span>';
      h += ahParts[i].label + ': ' + fmtNum(ahParts[i].val) + ' (' + pct + P + ')';
      h += '</span>';
    }
    h += '</div>';
    h += '</div>';
  }


  // 2. CATEGORY EFFECTIVENESS
  var ce = d.categoryEffectiveness;
  if (ce && ce.length > 0) {
    ce.sort(function(a, b) { return b.total - a.total; });
    var Q = String.fromCharCode(39);
    var ctid = 'tbl-cateff';
    h += '<div class="entity-card">';
    h += '<div class="entity-header"><div class="entity-info"><h3>Category Effectiveness</h3></div>';
    h += '<span class="record-badge">' + ce.length + ' categories</span></div>';
    h += '<table class="data-table" id="' + ctid + '">';
    h += '<thead><tr>';
    h += '<th onclick="sortT(' + Q + ctid + Q + ',0)">Category <span class="sort-arrow">' + svgSortN + '</span></th>';
    h += '<th class="col-right" onclick="sortT(' + Q + ctid + Q + ',1)">Companies <span class="sort-arrow active">' + svgSortD + '</span></th>';
    h += '<th class="col-right" onclick="sortT(' + Q + ctid + Q + ',2)">With Person <span class="sort-arrow">' + svgSortN + '</span></th>';
    var catPipeLabel = (typeof daSettings !== 'undefined') ? daSettings.getPipelineLabel('company') : 'Open Sale';
    var catPipeType = (typeof daSettings !== 'undefined') ? daSettings.getPipelineType('company') : 'sale';
    var catShowPipe = catPipeType !== 'none';
    h += '<th class="col-right" onclick="sortT(' + Q + ctid + Q + ',3)">Active (12m) <span class="sort-arrow">' + svgSortN + '</span></th>';
    if (catShowPipe) {
      h += '<th class="col-right" onclick="sortT(' + Q + ctid + Q + ',4)">' + catPipeLabel + ' <span class="sort-arrow">' + svgSortN + '</span></th>';
    }
    h += '<th class="col-right" onclick="sortT(' + Q + ctid + Q + ',' + (catShowPipe ? 5 : 4) + ')">Engagement <span class="sort-arrow">' + svgSortN + '</span></th>';
    h += '</tr></thead><tbody>';
    for (var i = 0; i < ce.length; i++) {
      var c = ce[i];
      var pctPers = c.total > 0 ? Math.round((c.withPerson / c.total) * 1000) / 10 : 0;
      var pctAct = c.total > 0 ? Math.round((c.withActivity / c.total) * 1000) / 10 : 0;
      var engagement = 0;
      if (typeof daSettings !== 'undefined') {
        engagement = daSettings.computeEngagement(c, c.total, 'company');
      } else {
        engagement = c.total > 0 ? Math.round((c.withActivity * 0.5 + c.withPerson * 0.3 + c.withSale * 0.2) / c.total * 100) : 0;
      }
      var engCol = slColor(engagement);
      h += '<tr>';
      h += '<td data-sort-value="' + c.name + '">' + c.name + '</td>';
      h += '<td class="col-right" data-sort-value="' + c.total + '">' + fmtNum(c.total) + '</td>';
      h += '<td class="col-right" data-sort-value="' + pctPers + '">' + fmtNum(c.withPerson) + '<span style="color:#999;font-size:.75rem;margin-left:4px">' + pctPers + P + '</span></td>';
      h += '<td class="col-right" data-sort-value="' + pctAct + '">' + fmtNum(c.withActivity) + '<span style="color:#999;font-size:.75rem;margin-left:4px">' + pctAct + P + '</span></td>';
      if (catShowPipe) {
        var pctSale = c.total > 0 ? Math.round((c.withSale / c.total) * 1000) / 10 : 0;
        h += '<td class="col-right" data-sort-value="' + pctSale + '">' + fmtNum(c.withSale) + '<span style="color:#999;font-size:.75rem;margin-left:4px">' + pctSale + P + '</span></td>';
      }
      h += '<td class="col-right" data-sort-value="' + engagement + '">' + barCell(engagement, engCol) + '</td>';
      h += '</tr>';
    }
    h += '</tbody></table></div>';
  }

  // 3. ASSOCIATE BREAKDOWN
  var assocs = d.associates;
  if (assocs && assocs.length > 0) {
    assocs.sort(function(a, b) { return b.total - a.total; });
    var totalAll = 0;
    for (var i = 0; i < assocs.length; i++) totalAll += assocs[i].total;
    var Q = String.fromCharCode(39);
    var tid = 'tbl-assoc';
    var hasGroups = assocs.length > 0 && assocs[0].groupName;
    var groups = {};
    var groupOrder = [];
    for (var i = 0; i < assocs.length; i++) {
      var gn = assocs[i].groupName || 'Other';
      if (!groups[gn]) { groups[gn] = { name: gn, members: [], total: 0, withPersons: 0, withActivities: 0, withEmail: 0, stale: 0 }; groupOrder.push(gn); }
      groups[gn].members.push(assocs[i]);
      groups[gn].total += assocs[i].total;
      groups[gn].withPersons += assocs[i].withPersons || 0;
      groups[gn].withActivities += assocs[i].withActivities || 0;
      groups[gn].withEmail += assocs[i].withEmail || 0;
      groups[gn].stale += assocs[i].stale || 0;
    }
    groupOrder.sort(function(a, b) { return groups[b].total - groups[a].total; });

    h += '<div class="entity-card">';
    h += '<div class="entity-header"><div class="entity-info"><h3>Associate Breakdown</h3></div>';
    h += '<div style="display:flex;gap:8px;align-items:center">';
    h += '<span class="record-badge">' + assocs.length + ' users</span>';
    if (hasGroups) {
      h += '<span class="record-badge" style="cursor:pointer;user-select:none" onclick="togAssocGroup()" id="assocGroupToggle">' + groupOrder.length + ' groups ' + svgBadgeChev + '</span>';
    }
    h += '</div></div>';
    h += '<table class="data-table" id="' + tid + '">';
    h += '<thead><tr>';
    h += '<th onclick="sortT(' + Q + tid + Q + ',0)">Associate <span class="sort-arrow">' + svgSortN + '</span></th>';
    h += '<th class="col-right" onclick="sortT(' + Q + tid + Q + ',1)">Companies <span class="sort-arrow active">' + svgSortD + '</span></th>';
    h += '<th class="col-right" onclick="sortT(' + Q + tid + Q + ',2)">With Persons <span class="sort-arrow">' + svgSortN + '</span></th>';
    h += '<th class="col-right" onclick="sortT(' + Q + tid + Q + ',3)">With Activities <span class="sort-arrow">' + svgSortN + '</span></th>';
    h += '<th class="col-right" onclick="sortT(' + Q + tid + Q + ',4)">With Email <span class="sort-arrow">' + svgSortN + '</span></th>';
    h += '<th class="col-right" onclick="sortT(' + Q + tid + Q + ',5)">Stale <span class="sort-arrow">' + svgSortN + '</span></th>';
    h += '<th class="col-right" onclick="sortT(' + Q + tid + Q + ',6)">Completeness <span class="sort-arrow">' + svgSortN + '</span></th>';
    h += '</tr></thead><tbody>';

    function assocRow(a) {
      var shareP = totalAll > 0 ? Math.round((a.total / totalAll) * 1000) / 10 : 0;
      var pctPers = a.total > 0 ? Math.round((a.withPersons / a.total) * 1000) / 10 : 0;
      var pctAct = a.total > 0 ? Math.round((a.withActivities / a.total) * 1000) / 10 : 0;
      var pctEm = a.total > 0 ? Math.round((a.withEmail / a.total) * 1000) / 10 : 0;
      var stale = a.stale || 0;
      var pctStale = a.total > 0 ? Math.round((stale / a.total) * 1000) / 10 : 0;
      var compP = a.total > 0 ? Math.round(((a.withPersons + a.withActivities + a.withEmail) / (a.total * 3)) * 100) : 0;
      var r = '<tr class="assoc-row" data-group="' + (a.groupName || 'Other') + '">';
      var displayName = a.name;
      if (/^\(\d+\)$/.test(displayName)) { displayName = '<span style="color:#999;font-style:italic">Unknown user ' + displayName + '</span>'; }
      r += '<td data-sort-value="' + a.name + '">' + displayName + '</td>';
      r += '<td class="col-right" data-sort-value="' + a.total + '">' + fmtNum(a.total) + '<span style="color:#999;font-size:.75rem;margin-left:4px">' + shareP + P + '</span></td>';
      r += '<td class="col-right" data-sort-value="' + a.withPersons + '">' + fmtNum(a.withPersons) + '<span style="color:#999;font-size:.75rem;margin-left:4px">' + pctPers + P + '</span></td>';
      r += '<td class="col-right" data-sort-value="' + a.withActivities + '">' + fmtNum(a.withActivities) + '<span style="color:#999;font-size:.75rem;margin-left:4px">' + pctAct + P + '</span></td>';
      r += '<td class="col-right" data-sort-value="' + a.withEmail + '">' + fmtNum(a.withEmail) + '<span style="color:#999;font-size:.75rem;margin-left:4px">' + pctEm + P + '</span></td>';
      r += '<td class="col-right" data-sort-value="' + stale + '">' + fmtNum(stale) + '<span style="color:#999;font-size:.75rem;margin-left:4px">' + pctStale + P + '</span></td>';
      r += '<td class="col-right" data-sort-value="' + compP + '">' + barCell(compP, '') + '</td>';
      r += '</tr>';
      return r;
    }

    if (hasGroups) {
      for (var gi = 0; gi < groupOrder.length; gi++) {
        var g = groups[groupOrder[gi]];
        var gShareP = totalAll > 0 ? Math.round((g.total / totalAll) * 1000) / 10 : 0;
        var gCompP = g.total > 0 ? Math.round(((g.withPersons + g.withActivities + g.withEmail) / (g.total * 3)) * 100) : 0;
        var gPctStale = g.total > 0 ? Math.round((g.stale / g.total) * 1000) / 10 : 0;
        h += '<tr class="assoc-group-header" onclick="togGroup(this)">';
        h += '<td data-sort-value="' + g.name + '"><svg class="group-chevron" viewBox="0 0 10 6" fill="none" xmlns="http://www.w3.org/2000/svg"><path d="M1 1l4 4 4-4" stroke="#333" stroke-width="1.3" stroke-linecap="round" stroke-linejoin="round"/></svg>' + g.name + ' <span style="color:#999;font-weight:400;font-size:.75rem">(' + g.members.length + ')</span></td>';
        h += '<td class="col-right" data-sort-value="' + g.total + '">' + fmtNum(g.total) + '<span style="color:#999;font-size:.75rem;margin-left:4px">' + gShareP + P + '</span></td>';
        h += '<td class="col-right" data-sort-value="' + g.withPersons + '">' + fmtNum(g.withPersons) + '</td>';
        h += '<td class="col-right" data-sort-value="' + g.withActivities + '">' + fmtNum(g.withActivities) + '</td>';
        h += '<td class="col-right" data-sort-value="' + g.withEmail + '">' + fmtNum(g.withEmail) + '</td>';
        h += '<td class="col-right" data-sort-value="' + g.stale + '">' + fmtNum(g.stale) + '<span style="color:#999;font-size:.75rem;margin-left:4px">' + gPctStale + P + '</span></td>';
        h += '<td class="col-right" data-sort-value="' + gCompP + '">' + barCell(gCompP, '') + '</td>';
        h += '</tr>';
        for (var mi = 0; mi < g.members.length; mi++) {
          h += assocRow(g.members[mi]);
        }
      }
    } else {
      for (var i = 0; i < assocs.length; i++) {
        h += assocRow(assocs[i]);
      }
    }

    h += '</tbody></table></div>';
  }

  el.innerHTML = h;
}

// === EXTRA TABLES ===
var extraTables = [];
var extraCur = 0;
var extraData = {};

function startExtra() {
  currentAnalysisEntity = '';
  var btn = document.getElementById('extraAnalyzeBtn');
  if (btn) btn.disabled = true;
  document.getElementById('extraStart').style.display = '';
  document.getElementById('extraStatus').textContent = 'Loading tables...';

  // Use cache for extra tables tab too
  getExtraTablesFromCache(function(cache) {
    extraTables = cache.tables;
    extraData = {};
    for (var i = 0; i < cache.tables.length; i++) {
      extraData[cache.tables[i].id] = cache.data[cache.tables[i].id];
    }
    showExtra();
  });
}

function showExtra() {
  document.getElementById('extraBar').style.width = '100' + P;
  document.getElementById('extraStatus').textContent = 'Complete!';
  setTimeout(function() {
    document.getElementById('extraStart').style.display = 'none';
    document.getElementById('extraResults').style.display = 'block';
    var extraRes = document.getElementById('extraResults');
    if (extraRes && !extraRes.querySelector('.filter-notice')) extraRes.insertAdjacentHTML('afterbegin', dateFilterNotice(true));
    document.getElementById('extraExportBtn').style.display = '';
    var btn = document.getElementById('extraAnalyzeBtn');
    if (btn) { btn.disabled = false; btn.onclick = function(){ document.getElementById('extraResults').style.display = 'none'; document.getElementById('extraCards').innerHTML = ''; extraData = {}; invalidateExtraCache(); startExtra(); }; }
    var wr = 0;
    for (var i = 0; i < extraTables.length; i++) {
      var d = extraData[extraTables[i].id];
      if (d && d.relationFields && d.relationFields.length > 0) wr++;
    }
    document.getElementById('extraSummary').textContent = extraTables.length + ' extra tables, ' + wr + ' with relation fields.';
    var c = document.getElementById('extraCards');
    for (var i = 0; i < extraTables.length; i++) {
      var d = extraData[extraTables[i].id];
      if (d) c.innerHTML += renderExtra(d, i);
    }
  }, 300);
}

function renderExtra(tbl, idx) {
  var hasRel = tbl.relationFields && tbl.relationFields.length > 0;
  var isEmpty = tbl.totalRows === 0;
  var cc = 'entity-card';
  if (isEmpty) cc = 'entity-card empty-table';
  var Q = String.fromCharCode(39);
  var h = '<div class="' + cc + '">';
  h += '<div class="entity-header"><div class="entity-info"><h3>' + tbl.displayName + '</h3> <span class="field-count">\u2014 ' + tbl.tableName + '</span></div>';
  h += '<div style="display:flex;gap:8px;align-items:center">';
  if (hasRel) h += '<span class="badge-so">Has Relations</span>';
  h += '<span class="record-badge">' + fmtNum(tbl.totalRows) + ' records</span></div></div>';
  if (isEmpty) {
    h += '<div style="padding:20px;color:#888"><em>No records in this table</em></div>';
  } else {
    if (hasRel) {
      h += '<div class="relation-section"><div class="section-label">Relation Fields</div>';
      for (var ri = 0; ri < tbl.relationFields.length; ri++) {
        var rf = tbl.relationFields[ri];
        h += '<div class="relation-block"><div class="relation-block-header">';
        h += '<div><span class="field-name">' + rf.displayName + '</span> <code>(' + rf.fieldName + ')</code></div>';
        h += '<span class="badge-relation">' + rf.entityType + '</span></div>';
        h += '<div class="stat-row">';
        h += '<div class="stat-card"><div class="stat-value">' + rf.filledCount + ' / ' + tbl.totalRows + '</div>';
        h += '<div class="stat-label">Records with ' + rf.entityType + '</div>';
        h += '<div style="margin-top:6px">' + fillBar(rf.fillPercent, 8) + '</div>';
        h += '<div class="stat-label" style="margin-top:3px">' + rf.fillPercent + P + ' filled</div></div>';
        h += '<div class="stat-card"><div class="stat-value">' + rf.uniqueCount + '</div>';
        h += '<div class="stat-label">Unique ' + rf.entityType + 's linked</div></div>';
        h += '<div class="stat-card"><div class="stat-value">' + rf.avgPerEntity + '</div>';
        h += '<div class="stat-label">Avg records per ' + rf.entityType + '</div></div>';
        h += '<div class="stat-card"><div class="stat-value">' + rf.coveragePercent + P + '</div>';
        h += '<div class="stat-label">' + rf.uniqueCount + ' of ' + rf.totalEntityCount + ' ' + rf.entityType + 's</div>';
        h += '<div class="stat-label">have a record in this table</div></div>';
        h += '</div></div>';
      }
      h += '</div>';
    }
    if (tbl.otherFields && tbl.otherFields.length > 0) {
      tbl.otherFields.sort(function(a,b) { return b.fillPercent - a.fillPercent; });
      h += '<div style="padding:0 20px 16px"><div class="section-label">Other Fields</div>';
      h += '<table class="data-table" id="tbl-e' + idx + '" style="border:1px solid #eeecea;border-radius:5px;overflow:hidden">';
      h += '<thead><tr>';
      h += '<th onclick="sortT(' + Q + 'tbl-e' + idx + Q + ',0)">Field <span class="sort-arrow">' + svgSortN + '</span></th>';
      h += '<th onclick="sortT(' + Q + 'tbl-e' + idx + Q + ',1)">Type <span class="sort-arrow">' + svgSortN + '</span></th>';
      h += '<th class="col-right" onclick="sortT(' + Q + 'tbl-e' + idx + Q + ',2)">Filled <span class="sort-arrow">' + svgSortN + '</span></th>';
      h += '<th class="col-right" onclick="sortT(' + Q + 'tbl-e' + idx + Q + ',3)">' + P + ' <span class="sort-arrow active">' + svgSortD + '</span></th>';
      h += '<th style="width:120px">Fill Rate</th></tr></thead><tbody>';
      for (var fi = 0; fi < tbl.otherFields.length; fi++) {
        var f = tbl.otherFields[fi];
        var rc = '';
        if (f.filledCount === 0) rc = 'unused';
        h += '<tr class="' + rc + '">';
        h += '<td data-sort-value="' + f.displayName + '">' + f.displayName + ' <code style="font-size:.75rem;color:#999">(' + f.fieldName + ')</code></td>';
        h += '<td data-sort-value="' + f.fieldType + '"><span class="type-badge">' + f.fieldType + '</span></td>';
        h += '<td class="col-right" data-sort-value="' + f.filledCount + '">' + f.filledCount + '</td>';
        h += '<td class="col-right" data-sort-value="' + f.fillPercent + '">' + f.fillPercent + P + '</td>';
        h += '<td>' + fillBar(f.fillPercent) + '</td></tr>';
      }
      h += '</tbody></table></div>';
    }
  }
  h += '</div>';
  return h;
}

// === ENTITY EXTRA TABLES (now using global cache) ===
var entExtra = {};

function startEntityExtra(key, entityType) {
  if (!entExtra[key]) entExtra[key] = { tables:[], cur:0, data:{} };
  var btns = document.querySelectorAll('#' + key + 'ExtraStart .btn-analyze');
  for (var i = 0; i < btns.length; i++) btns[i].style.display = 'none';
  document.getElementById(key + 'ExtraProgress').style.display = 'block';
  document.getElementById(key + 'ExtraStatus').textContent = 'Loading tables...';

  getExtraTablesFromCache(function(cache) {
    entExtra[key].tables = cache.tables;
    entExtra[key].entityType = entityType;
    for (var i = 0; i < cache.tables.length; i++) {
      entExtra[key].data[cache.tables[i].id] = cache.data[cache.tables[i].id];
    }
    entExtra[key].cur = cache.tables.length;
    showEntityExtra(key);
  });
}

function showEntityExtra(key) {
  var es = entExtra[key];
  document.getElementById(key + 'ExtraBar').style.width = '100' + P;
  document.getElementById(key + 'ExtraStatus').textContent = 'Complete!';
  setTimeout(function() {
    document.getElementById(key + 'ExtraStart').style.display = 'none';
    document.getElementById(key + 'ExtraResults').style.display = 'block';
    var related = [];
    for (var i = 0; i < es.tables.length; i++) {
      var d = es.data[es.tables[i].id];
      if (!d || !d.relationFields) continue;
      var match = false;
      for (var j = 0; j < d.relationFields.length; j++) {
        if (d.relationFields[j].entityType === es.entityType) { match = true; break; }
      }
      if (match) related.push(es.tables[i]);
    }
    document.getElementById(key + 'ExtraSummary').textContent = related.length + ' extra tables with ' + es.entityType + ' relations (out of ' + es.tables.length + ' total).';
    var c = document.getElementById(key + 'ExtraCards');
    var baseIdx = 500;
    if (key === 'contact') baseIdx = 600;
    if (key === 'sale') baseIdx = 700;
    if (key === 'project') baseIdx = 800;
    if (key === 'requests') baseIdx = 900;
    for (var i = 0; i < related.length; i++) {
      var d = es.data[related[i].id];
      if (d) c.innerHTML += renderExtra(d, baseIdx + i);
    }
  }, 300);
}

function startCompanyExtra() { startEntityExtra('company', 'Company'); }

// Used by startFullEntity (progressive) and aaRunFullEntity
function loadEntityExtraWithProgress(key, entityType, onProgress, onComplete) {
  if (!entExtra[key]) entExtra[key] = { tables:[], cur:0, data:{} };

  getExtraTablesFromCache(function(cache) {
    entExtra[key].tables = cache.tables;
    entExtra[key].entityType = entityType;
    for (var i = 0; i < cache.tables.length; i++) {
      entExtra[key].data[cache.tables[i].id] = cache.data[cache.tables[i].id];
      if (onProgress) onProgress(i + 1, cache.tables.length, cache.tables[i].displayName);
    }
    entExtra[key].cur = cache.tables.length;
    showEntityExtraQuiet(key);
    if (onComplete) onComplete();
  });
}

function showEntityExtraQuiet(key) {
  var es = entExtra[key];
  if (!es) return;
  var startEl = document.getElementById(key + 'ExtraStart');
  var resultsEl = document.getElementById(key + 'ExtraResults');
  if (startEl) startEl.style.display = 'none';
  if (resultsEl) resultsEl.style.display = 'block';
  var related = [];
  for (var i = 0; i < es.tables.length; i++) {
    var d = es.data[es.tables[i].id];
    if (!d || !d.relationFields) continue;
    var match = false;
    for (var j = 0; j < d.relationFields.length; j++) {
      if (d.relationFields[j].entityType === es.entityType) { match = true; break; }
    }
    if (match) related.push(es.tables[i]);
  }
  var summaryEl = document.getElementById(key + 'ExtraSummary');
  if (summaryEl) summaryEl.textContent = related.length + ' extra tables with ' + es.entityType + ' relations (out of ' + es.tables.length + ' total).';
  var c = document.getElementById(key + 'ExtraCards');
  var baseIdx = 500;
  if (key === 'contact') baseIdx = 600;
  if (key === 'sale') baseIdx = 700;
  if (key === 'project') baseIdx = 800;
  if (key === 'requests') baseIdx = 900;
  if (c) { c.innerHTML = ''; for (var i = 0; i < related.length; i++) { var d = es.data[related[i].id]; if (d) c.innerHTML += renderExtra(d, baseIdx + i); } }
}

function loadEntityExtraQuiet(key, entityType, callback) {
  loadEntityExtraWithProgress(key, entityType, null, callback);
}

function loadEntityExtraChain(key, callback) {
  // Legacy — now just uses cache
  loadEntityExtraWithProgress(entExtra[key] ? key : key, entExtra[key] ? entExtra[key].entityType : '', null, callback);
}

function loadTicketFields(callback) {
  var startEl = document.getElementById('ticketStart');
  if (startEl) startEl.style.display = 'none';
  ajax(ticketUrl + getDateFilterParam(), function(d) {
    ticketData = d;
    showTicket();
    if (callback) callback();
  });
}

function loadTicketFieldsQuiet(callback) {
  ajax(ticketUrl + getDateFilterParam(), function(d) {
    ticketData = d;
    showTicketQuiet();
    if (callback) callback();
  });
}

function showTicketQuiet() {
  var rc = 0, uc = 0;
  if (ticketData && ticketData.fields) {
    for (var i = 0; i < ticketData.fields.length; i++) {
      if (ticketData.fields[i].isRelation) rc++;
      if (ticketData.fields[i].filledCount > 0) uc++;
    }
    var summaryEl = document.getElementById('ticketSummary');
    if (summaryEl) summaryEl.textContent = ticketData.fields.length + ' custom fields, ' + rc + ' relation fields. ' + uc + ' fields in use.';
    ticketData.fields.sort(function(a,b) { return b.fillPercent - a.fillPercent; });
    var cardsEl = document.getElementById('ticketCards');
    if (cardsEl) cardsEl.innerHTML = renderTicket();
  }
}

// === TICKET FIELDS ===
var ticketData = null;

function startTicket() {
  var btns = document.querySelectorAll('#ticketStart .btn-analyze');
  for (var i = 0; i < btns.length; i++) btns[i].style.display = 'none';
  document.getElementById('ticketProgress').style.display = 'block';
  ajax(ticketUrl + getDateFilterParam(), function(d) { ticketData = d; showTicket(); });
}

function showTicket() {
  document.getElementById('ticketStart').style.display = 'none';
  document.getElementById('ticketResults').style.display = 'block';
  document.getElementById('ticketExportBtn').style.display = '';
  var rc = 0, uc = 0;
  for (var i = 0; i < ticketData.fields.length; i++) {
    if (ticketData.fields[i].isRelation) rc++;
    if (ticketData.fields[i].filledCount > 0) uc++;
  }
  document.getElementById('ticketSummary').textContent = ticketData.fields.length + ' custom fields, ' + rc + ' relation fields. ' + uc + ' fields in use.';
  ticketData.fields.sort(function(a,b) { return b.fillPercent - a.fillPercent; });
  document.getElementById('ticketCards').innerHTML = renderTicket();
}

function renderTicket() {
  var Q = String.fromCharCode(39);
  var h = '<div class="entity-card">';
  h += '<div class="entity-header"><div class="entity-info"><h3>Custom Ticket Fields</h3></div>';
  h += '<span class="record-badge">' + fmtNum(ticketData.totalTickets) + ' tickets</span></div>';
  h += '<table class="data-table" id="tbl-t0">';
  h += '<thead><tr>';
  h += '<th onclick="sortT(' + Q + 'tbl-t0' + Q + ',0)">Field Label <span class="sort-arrow">' + svgSortN + '</span></th>';
  h += '<th onclick="sortT(' + Q + 'tbl-t0' + Q + ',1)">Field Type <span class="sort-arrow">' + svgSortN + '</span></th>';
  h += '<th class="col-right" onclick="sortT(' + Q + 'tbl-t0' + Q + ',2)">Filled <span class="sort-arrow">' + svgSortN + '</span></th>';
  h += '<th class="col-right" onclick="sortT(' + Q + 'tbl-t0' + Q + ',3)">' + P + ' <span class="sort-arrow active">' + svgSortD + '</span></th>';
  h += '<th style="width:140px">Fill Rate</th>';
  h += '<th>Details</th></tr></thead><tbody>';
  for (var i = 0; i < ticketData.fields.length; i++) {
    var f = ticketData.fields[i];
    var rc = '';
    if (f.filledCount === 0) rc = 'unused';
    if (f.isRelation) rc += ' relation-row';
    h += '<tr class="' + rc + '">';
    h += '<td data-sort-value="' + f.displayName + '">' + f.displayName;
    if (f.isRelation) h += ' <span class="badge-relation">Relation</span>';
    h += '</td>';
    h += '<td data-sort-value="' + f.fieldType + '"><span class="type-badge">' + f.fieldType + '</span></td>';
    h += '<td class="col-right" data-sort-value="' + f.filledCount + '">' + f.filledCount + '</td>';
    h += '<td class="col-right" data-sort-value="' + f.fillPercent + '">' + f.fillPercent + P + '</td>';
    h += '<td>' + fillBar(f.fillPercent) + '</td>';
    h += '<td style="font-size:.78rem;color:#888">';
    if (f.isRelation && f.filledCount > 0) {
      h += f.uniqueCount + ' unique ' + f.entityType + 's, ' + f.avgPerEntity + ' avg/entity, ' + f.coveragePercent + P + ' coverage';
    }
    h += '</td></tr>';
  }
  h += '</tbody></table></div>';
  return h;
}


// =====================================================
// CRM MOMENTUM
// =====================================================

var momentumData = null;

function startMomentum(key) {
  captureEntityFilter(key);
  var progressScreen = document.getElementById(key + 'ProgressScreen');
  var progressBar = document.getElementById(key + 'ProgressBar');
  var progressPercent = document.getElementById(key + 'ProgressPercent');
  var progressStatus = document.getElementById(key + 'ProgressStatus');
  var resultsContainer = document.getElementById(key + 'Results');
  var subTabs = document.getElementById(key + 'SubTabs');
  var headerBtn = document.getElementById(key + 'AnalyzeBtn');

  if (headerBtn) headerBtn.disabled = true;
  if (progressScreen) progressScreen.style.display = '';
  if (progressBar) { progressBar.style.width = '0'; progressBar.classList.add('loading'); }
  if (progressPercent) progressPercent.textContent = '0' + P;
  if (progressStatus) progressStatus.textContent = 'Loading momentum data...';

  var currentPct = 0;
  var fakeTimer = setInterval(function() {
    if (currentPct < 95) {
      var remaining = 95 - currentPct;
      var increment = Math.max(0.5, remaining * 0.06);
      currentPct = Math.min(currentPct + increment, 95);
      if (progressBar) progressBar.style.width = currentPct + P;
      if (progressPercent) progressPercent.textContent = Math.round(currentPct) + P;
    }
  }, 150);

  var dfParam = getDateFilterParam();
  var loadsDone = 0;
  var totalLoads = 2;

  function checkDone() {
    loadsDone++;
    if (loadsDone < totalLoads) return;
    clearInterval(fakeTimer);
    if (progressBar) { progressBar.style.width = '100' + P; progressBar.classList.remove('loading'); }
    if (progressPercent) progressPercent.textContent = '100' + P;
    if (progressStatus) progressStatus.textContent = 'Complete!';
    renderMomentum(key, momentumData);
    setTimeout(function() {
      if (progressScreen) progressScreen.style.display = 'none';
      if (subTabs) subTabs.style.display = '';
      if (resultsContainer) resultsContainer.style.display = '';
      var expBtn = document.getElementById(key + 'ExportBtn');
      if (expBtn) expBtn.style.display = '';
      if (headerBtn) { headerBtn.disabled = false; headerBtn.onclick = function(){ reAnalyze(key); }; }
    }, 500);
  }

  // Build excludeTypes param from momentum settings
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
  ajax(overviewUrl + amp + 'entity=' + key + dfParam + getExclParam(key), function(d) {
    if (d) renderEntityOverview(key, d);
    checkDone();
  });
}

// Level ramp (handover): an ordered scale of its own. No blue, no alarm red.
// Inactive is neutral, because dormancy is an opportunity and not an error.
var MM_GREEN = '#2f8f5b';   // Power    dot
var MM_BLUE = '#2f817a';    // Regular  dot (--c2)
var MM_ORANGE = '#b9832f';  // Low      dot (--mod)
var MM_RED = '#c8c2b4';     // Inactive dot (--cn)
var MM_LEVELS = {
  Power:    { bg: '#d3e8d8', fg: '#226a43', dot: MM_GREEN },
  Regular:  { bg: '#d6e7e3', fg: '#1f6a63', dot: MM_BLUE },
  Low:      { bg: '#e8d5b0', fg: '#785012', dot: MM_ORANGE },
  Inactive: { bg: '#e4ded3', fg: '#625d51', dot: MM_RED }
};
function mmLevelStyle(level) { return MM_LEVELS[level] || MM_LEVELS.Inactive; }

function mmLevelColor(level) {
  if (level === 'Power') return MM_GREEN;
  if (level === 'Regular') return MM_BLUE;
  if (level === 'Low') return MM_ORANGE;
  return MM_RED;
}

function mmUserLevel(total, months) {
  var avg = months > 0 ? total / months : total;
  var powerT = 100, regT = 25;
  if (typeof daSettings !== 'undefined' && daSettings.getMomentumSettings) {
    var ms = daSettings.getMomentumSettings();
    powerT = ms.powerThreshold || 100;
    regT = ms.regularThreshold || 25;
  }
  if (avg >= powerT) return 'Power';
  if (avg >= regT) return 'Regular';
  if (total > 0) return 'Low';
  return 'Inactive';
}

function renderMomentum(key, d) {
  var el = document.getElementById(key + 'MomentumContent');
  if (!el || !d) return;
  var monthly = d.monthly || [];
  var users = d.users || [];
  var totalUsers = d.totalUsers || 0;
  // KPI: compare last 2 FULL months (skip current incomplete month)
  // monthly[n-1] = current month (incomplete), monthly[n-2] = last full, monthly[n-3] = previous full
  var lastFullMonth = monthly.length > 1 ? monthly[monthly.length - 2] : null;
  var prevFullMonth = monthly.length > 2 ? monthly[monthly.length - 3] : null;
  var filterMonths = monthly.length;
  var filterLabel = 'All data';
  var dfVal = activeFilterValue['activities'] || '';
  if (dfVal) {
    var now = new Date();
    var filterDate = new Date(dfVal);
    var diffMs = now.getTime() - filterDate.getTime();
    filterMonths = Math.max(1, Math.round(diffMs / (30.44 * 86400000)));
    var mo = filterDate.getMonth() + 1;
    var yr = filterDate.getFullYear();
    filterLabel = 'Since ' + yr + '-' + (mo < 10 ? '0' : '') + mo;
  }
  var monthName = lastFullMonth ? lastFullMonth.label.split("'")[0].trim() : '';
  var h = '';

  // KPI cards — reuse detail-section + stat-row + stat-card
  h += '<div class="detail-section">';
  h += '<div class="detail-section-head">' + secHead('CRM Momentum \u2014 Last 24 Months') + '</div>';
  h += '<div class="stat-row">';
  h += mmKpiCard(lastFullMonth ? fmtNum(lastFullMonth.activities) : '0', 'Activities ' + monthName,
    lastFullMonth && prevFullMonth ? mmTrend(lastFullMonth.activities, prevFullMonth.activities) : null);
  h += mmKpiCard(lastFullMonth ? fmtNum(lastFullMonth.documents) : '0', 'Documents ' + monthName,
    lastFullMonth && prevFullMonth ? mmTrend(lastFullMonth.documents, prevFullMonth.documents) : null);
  var auVal = lastFullMonth ? lastFullMonth.activeUsers : 0;
  h += mmKpiCard(auVal + ' <span style="font-size:.7em;color:#888;font-weight:400">/ ' + totalUsers + '</span>',
    'Active Users ' + monthName,
    lastFullMonth && prevFullMonth ? mmTrendAbs(lastFullMonth.activeUsers, prevFullMonth.activeUsers) : null);
  var avgAct = (auVal > 0 && lastFullMonth) ? Math.round((lastFullMonth.activities + lastFullMonth.documents) / auVal) : 0;
  var prevAvg = (prevFullMonth && prevFullMonth.activeUsers > 0) ? Math.round((prevFullMonth.activities + prevFullMonth.documents) / prevFullMonth.activeUsers) : 0;
  h += mmKpiCard(fmtNum(avgAct), 'Avg per User ' + monthName,
    lastFullMonth && prevFullMonth ? mmTrend(avgAct, prevAvg) : null);
  h += '</div></div>';

  h += renderMomentumChart(monthly);
  h += renderUserAdoption(users, totalUsers, d.totalActivities + d.totalDocuments, filterMonths, filterLabel);

  el.innerHTML = dateFilterNotice() + h;
  mmInitChartTooltip(el, monthly);
}

function mmKpiCard(value, label, trend) {
  var h = '<div class="stat-card">';
  h += '<div class="stat-value" style="color:var(--so-charcoal)">' + value + '</div>';
  h += '<div class="stat-label">' + label + '</div>';
  if (trend) h += '<div style="font-size:.72rem;margin-top:6px;font-weight:600" class="' + trend.cls + '">' + trend.text + '</div>';
  h += '</div>';
  return h;
}

function mmTrend(curr, prev) {
  if (!prev || prev === 0) {
    if (curr > 0) return { cls: 'mm-trend-up', text: '\u25B2 new' };
    return { cls: 'mm-trend-flat', text: '\u2014 no data' };
  }
  var pct = Math.round(((curr - prev) / prev) * 100);
  if (pct > 0) return { cls: 'mm-trend-up', text: '\u25B2 ' + pct + P + ' vs previous month' };
  if (pct < 0) return { cls: 'mm-trend-down', text: '\u25BC ' + Math.abs(pct) + P + ' vs previous month' };
  return { cls: 'mm-trend-flat', text: '\u2014 same as previous month' };
}

function mmTrendAbs(curr, prev) {
  var diff = curr - prev;
  if (diff > 0) return { cls: 'mm-trend-up', text: '\u25B2 ' + diff + ' vs previous month' };
  if (diff < 0) return { cls: 'mm-trend-down', text: '\u25BC ' + Math.abs(diff) + ' vs previous month' };
  return { cls: 'mm-trend-flat', text: '\u2014 same as previous month' };
}

function mmRoundedTopRect(x, y, w, h, r, fill) {
  if (h <= 0) return '';
  if (r > h) r = h;
  if (r > w / 2) r = w / 2;
  var x2 = x + w, yb = y + h;
  return '<path d="M' + x.toFixed(1) + ',' + yb.toFixed(1)
    + ' L' + x.toFixed(1) + ',' + (y + r).toFixed(1)
    + ' Q' + x.toFixed(1) + ',' + y.toFixed(1) + ' ' + (x + r).toFixed(1) + ',' + y.toFixed(1)
    + ' L' + (x2 - r).toFixed(1) + ',' + y.toFixed(1)
    + ' Q' + x2.toFixed(1) + ',' + y.toFixed(1) + ' ' + x2.toFixed(1) + ',' + (y + r).toFixed(1)
    + ' L' + x2.toFixed(1) + ',' + yb.toFixed(1) + ' Z" fill="' + fill + '"/>';
}

function renderMomentumChart(monthly) {
  if (!monthly || monthly.length === 0) return '';
  var MM_GREEN_LIGHT = '#a8d3c4';
  var maxTotal = 0, maxUsers = 0;
  for (var i = 0; i < monthly.length; i++) {
    var tot = (monthly[i].activities || 0) + (monthly[i].documents || 0);
    if (tot > maxTotal) maxTotal = tot;
    if (monthly[i].activeUsers > maxUsers) maxUsers = monthly[i].activeUsers;
  }

  var opts = [1, 1.5, 2, 2.5, 3, 4, 5, 8, 10];
  var niceMaxTot = maxTotal;
  var mag = Math.pow(10, Math.floor(Math.log10(maxTotal || 1)));
  for (var oi = 0; oi < opts.length; oi++) { if (opts[oi] * mag >= maxTotal) { niceMaxTot = opts[oi] * mag; break; } }
  var niceMaxUsr = maxUsers;
  var magU = Math.pow(10, Math.floor(Math.log10(maxUsers || 1)));
  for (var oi = 0; oi < opts.length; oi++) { if (opts[oi] * magU >= maxUsers) { niceMaxUsr = opts[oi] * magU; break; } }

  var cW = 960, cH = 270, padL = 48, padR = 40, padT = 16, padB = 40;
  var plotW = cW - padL - padR, plotH = cH - padT - padB;
  var n = monthly.length;
  var ySteps = 4;
  var yStepTot = niceMaxTot / ySteps;
  var yStepUsr = niceMaxUsr / ySteps;
  var baseline = padT + plotH;
  var colW = plotW / n;
  var barW = Math.min(colW * 0.62, 26);

  function cx(i) { return padL + colW * (i + 0.5); }

  var svg = '<svg viewBox="0 0 ' + cW + ' ' + cH + '" style="width:100%;height:auto;display:block;max-height:280px">';

  // Y grid lines + left (total volume) + right (active users) labels
  for (var gi = 0; gi <= ySteps; gi++) {
    var gy = baseline - (gi / ySteps) * plotH;
    svg += '<line x1="' + padL + '" y1="' + gy.toFixed(1) + '" x2="' + (cW - padR) + '" y2="' + gy.toFixed(1) + '" stroke="#e0dfdc" stroke-width="1"/>';
    svg += '<text x="' + (padL - 8) + '" y="' + (gy + 4).toFixed(1) + '" text-anchor="end" fill="#999" font-size="10" font-family="DM Sans,sans-serif">' + mmFmtK(Math.round(gi * yStepTot)) + '</text>';
    svg += '<text x="' + (cW - padR + 8) + '" y="' + (gy + 4).toFixed(1) + '" text-anchor="start" fill="' + MM_BLUE + '" font-size="10" font-family="DM Sans,sans-serif" opacity=".6">' + Math.round(gi * yStepUsr) + '</text>';
  }

  // Stacked bars: activities (bottom, square) + documents (top, only outer corners rounded)
  for (var i = 0; i < n; i++) {
    var actH = niceMaxTot > 0 ? (monthly[i].activities / niceMaxTot) * plotH : 0;
    var docH = niceMaxTot > 0 ? (monthly[i].documents / niceMaxTot) * plotH : 0;
    var bx = cx(i) - barW / 2;
    var yActTop = baseline - actH;
    var yDocTop = yActTop - docH;
    if (docH >= 1) {
      if (actH > 0) svg += '<rect x="' + bx.toFixed(1) + '" y="' + yActTop.toFixed(1) + '" width="' + barW.toFixed(1) + '" height="' + actH.toFixed(1) + '" fill="' + MM_GREEN + '"/>';
      svg += mmRoundedTopRect(bx, yDocTop, barW, docH, 3, MM_GREEN_LIGHT);
    } else if (actH > 0) {
      svg += mmRoundedTopRect(bx, yActTop, barW, actH, 3, MM_GREEN);
    }
  }

  // Active users line (right axis) + dots
  var usrPts = [];
  for (var i = 0; i < n; i++) {
    var yU = niceMaxUsr > 0 ? baseline - (monthly[i].activeUsers / niceMaxUsr) * plotH : baseline;
    usrPts.push(cx(i).toFixed(1) + ',' + yU.toFixed(1));
  }
  svg += '<polyline points="' + usrPts.join(' ') + '" fill="none" stroke="' + MM_BLUE + '" stroke-width="2" stroke-linejoin="round" stroke-linecap="round"/>';
  for (var i = 0; i < n; i++) {
    if (i === 0 || i === n - 1 || i % 4 === 0) {
      var p = usrPts[i].split(',');
      svg += '<circle cx="' + p[0] + '" cy="' + p[1] + '" r="3" fill="' + MM_BLUE + '" stroke="#fff" stroke-width="1.5"/>';
    }
  }

  // X labels — every ~2 months
  var xStep = Math.max(1, Math.floor(n / 12));
  for (var i = 0; i < n; i += xStep) {
    svg += '<text x="' + cx(i).toFixed(1) + '" y="' + (baseline + 20) + '" text-anchor="middle" fill="#999" font-size="10" font-family="DM Sans,sans-serif">' + monthly[i].label + '</text>';
  }
  if ((n - 1) % xStep !== 0) {
    svg += '<text x="' + cx(n - 1).toFixed(1) + '" y="' + (baseline + 20) + '" text-anchor="middle" fill="#999" font-size="10" font-family="DM Sans,sans-serif">' + monthly[n - 1].label + '</text>';
  }

  // Invisible hover columns for tooltip
  for (var i = 0; i < n; i++) {
    svg += '<rect x="' + (padL + colW * i).toFixed(1) + '" y="' + padT + '" width="' + colW.toFixed(1) + '" height="' + plotH + '" fill="transparent" class="mm-hover-col" data-idx="' + i + '"/>';
  }

  svg += '</svg>';

  // Legend + plain-language definitions, inside the chart block
  var legend = '<div class="mm-legend" style="padding-left:' + padL + 'px">';
  legend += '<span class="mm-leg-item"><span class="mm-leg-sw" style="background:' + MM_GREEN + '"></span> Activities</span>';
  legend += '<span class="mm-leg-item"><span class="mm-leg-sw" style="background:' + MM_GREEN_LIGHT + '"></span> Documents</span>';
  legend += '<span class="mm-leg-item"><span class="mm-leg-line" style="background:' + MM_BLUE + '"></span> Active users (right axis)</span>';
  legend += '</div>';
  legend += '<div class="mm-chart-defs" style="padding-left:' + padL + 'px"><b>Active user</b> = a user with at least one logged activity or document that month. <b>Avg per user</b> divides that month\'s total activity by its active users. Outlook-synced and private activities are excluded.</div>';

  var h = '<div class="entity-card">';
  h += '<div class="entity-header"><div class="entity-info"><h3>Activity Volume</h3>';
  h += '<span class="field-count">\u2014 based on activeDate</span></div>';
  h += '<span class="record-badge">24 months</span></div>';
  h += '<div style="padding:16px 20px 12px;position:relative" id="mmChartArea">';
  h += svg;
  h += '<div class="mm-tooltip" id="mmTooltip"></div>';
  h += legend;
  h += '</div></div>';
  return h;
}

function mmNiceStep(max) {
  if (max <= 0) return 1;
  var r = max / 4, mag = Math.pow(10, Math.floor(Math.log10(r))), n = r / mag;
  if (n <= 1) return mag; if (n <= 2) return 2 * mag; if (n <= 5) return 5 * mag; return 10 * mag;
}
function mmFmtK(n) { if (n >= 1000) return (n / 1000).toFixed(n % 1000 === 0 ? 0 : 1) + 'k'; return n.toString(); }

function mmInitChartTooltip(el, monthly) {
  var wrap = el.querySelector('#mmChartArea');
  var tip = el.querySelector('#mmTooltip');
  if (!wrap || !tip) return;
  var cols = wrap.querySelectorAll('.mm-hover-col');
  for (var i = 0; i < cols.length; i++) {
    (function(col) {
      col.addEventListener('mouseenter', function() {
        var idx = parseInt(col.getAttribute('data-idx'));
        if (idx < 0 || idx >= monthly.length) return;
        var m = monthly[idx];
        tip.innerHTML = '<strong>' + m.label + '</strong><br>Activities: ' + fmtNum(m.activities) + '<br>Documents: ' + fmtNum(m.documents) + '<br>Active Users: ' + m.activeUsers;
        tip.classList.add('visible');
      });
      col.addEventListener('mousemove', function(e) {
        var rect = wrap.getBoundingClientRect();
        tip.style.left = Math.min(e.clientX - rect.left + 12, rect.width - 160) + 'px';
        tip.style.top = Math.max(e.clientY - rect.top - 60, 0) + 'px';
      });
      col.addEventListener('mouseleave', function() { tip.classList.remove('visible'); });
    })(cols[i]);
  }
}

function renderUserAdoption(users, totalUsers, grandTotal, filterMonths, filterLabel) {
  if (!users || users.length === 0) return '';
  var power = [], regular = [], low = [], inactive = [], totalAct = 0;
  for (var i = 0; i < users.length; i++) {
    users[i].level = mmUserLevel(users[i].total, filterMonths);
    users[i].levelColor = mmLevelColor(users[i].level);
    users[i].levelStyle = mmLevelStyle(users[i].level);
    totalAct += users[i].total;
  }
  for (var i = 0; i < users.length; i++) {
    if (users[i].level === 'Power') power.push(users[i]);
    else if (users[i].level === 'Regular') regular.push(users[i]);
    else if (users[i].level === 'Low') low.push(users[i]);
    else inactive.push(users[i]);
  }
  users.sort(function(a, b) { return b.total - a.total; });
  var pAct = 0, rAct = 0, lAct = 0;
  for (var i = 0; i < power.length; i++) pAct += power[i].total;
  for (var i = 0; i < regular.length; i++) rAct += regular[i].total;
  for (var i = 0; i < low.length; i++) lAct += low[i].total;

  var h = '<div class="entity-card">';
  h += '<div class="entity-header"><div class="entity-info"><h3>User Adoption</h3>';
  h += '<span class="field-count">\u2014 based on date filter</span></div>';
  h += '<div style="display:flex;gap:8px;align-items:center">';
  h += '<span class="record-badge" style="background:#e8f5e9;color:' + MM_GREEN + ';font-weight:600">' + filterLabel + '</span>';
  h += '<span class="record-badge">' + totalUsers + ' total users</span></div></div>';

  var powerT = 100, regT = 25;
  if (typeof daSettings !== 'undefined' && daSettings.getMomentumSettings) {
    var ms = daSettings.getMomentumSettings();
    powerT = ms.powerThreshold || 100;
    regT = ms.regularThreshold || 25;
  }

  h += '<div class="mm-user-groups">';
  h += mmGroupCard(power.length, 'Power Users', powerT + '+ activities / month (avg)', totalUsers, totalAct, pAct, MM_GREEN);
  h += mmGroupCard(regular.length, 'Regular Users', regT + '\u2013' + (powerT - 1) + ' activities / month (avg)', totalUsers, totalAct, rAct, MM_BLUE);
  h += mmGroupCard(low.length, 'Low Usage', '1\u2013' + (regT - 1) + ' activities / month (avg)', totalUsers, totalAct, lAct, MM_ORANGE);
  h += mmGroupCardInactive(inactive.length, 'Inactive', '0 activities in period', totalUsers, MM_RED);
  h += '</div>';

  var top10 = users.slice(0, Math.min(10, users.length));
  if (top10.length > 0 && top10[0].total > 0) {
    h += '<div style="padding:16px 16px 8px;border-top:1px solid var(--so-border);font-size:.78rem;font-weight:600;color:#888;text-transform:uppercase;letter-spacing:.04em">Top 10 Most Active Users</div>';
    h += '<table class="data-table mm-user-table"><thead><tr>';
    h += '<th class="mm-col-rank">#</th><th class="mm-col-user">User</th><th class="mm-col-level">Level</th>';
    h += '<th class="mm-col-group">Group</th><th class="col-right mm-col-num">Activities</th><th class="col-right mm-col-num">Documents</th>';
    h += '<th class="col-right mm-col-total">Total</th><th class="col-right mm-col-pct">' + P + ' of Total</th>';
    h += '</tr></thead><tbody>';
    for (var i = 0; i < top10.length; i++) {
      var u = top10[i]; if (u.total === 0) break;
      var pct = totalAct > 0 ? Math.round((u.total / totalAct) * 100) : 0;
      h += '<tr><td><span class="mm-rank" style="background:' + u.levelStyle.bg + ';color:' + u.levelStyle.fg + '">' + (i + 1) + '</span></td>';
      h += '<td><strong>' + u.name + '</strong></td>';
      h += '<td><span class="mm-badge" style="background:' + u.levelStyle.bg + ';color:' + u.levelStyle.fg + '">' + u.level + '</span></td>';
      h += '<td style="color:#888">' + (u.group || '\u2014') + '</td>';
      h += '<td class="col-right">' + fmtNum(u.activities) + '</td><td class="col-right">' + fmtNum(u.documents) + '</td>';
      h += '<td class="col-right"><strong>' + fmtNum(u.total) + '</strong></td>';
      h += '<td class="col-right">' + mmBarCell(pct, u.levelColor) + '</td></tr>';
    }
    h += '</tbody></table>';
  }

  var groups = mmGroupUsers(users);
  if (groups.length > 0) {
    h += '<div style="padding:16px 16px 8px;border-top:1px solid var(--so-border);font-size:.78rem;font-weight:600;color:#888;text-transform:uppercase;letter-spacing:.04em">All Users by Group</div>';
    h += '<table class="data-table mm-user-table"><thead><tr>';
    h += '<th class="mm-col-rank"></th><th class="mm-col-user">User</th><th class="mm-col-level">Level</th><th class="mm-col-group">Group</th>';
    h += '<th class="col-right mm-col-num">Activities</th><th class="col-right mm-col-num">Documents</th>';
    h += '<th class="col-right mm-col-total">Total</th><th class="col-right mm-col-pct">' + P + ' of Total</th>';
    h += '</tr></thead><tbody>';
    for (var gi = 0; gi < groups.length; gi++) {
      var grp = groups[gi];
      var gPct = totalAct > 0 ? Math.round((grp.total / totalAct) * 100) : 0;
      var gInfo = grp.members.length + ' user' + (grp.members.length !== 1 ? 's' : '') + ' \u00B7 ' + fmtNum(grp.total) + ' activities';
      if (gPct > 0) gInfo += ' \u00B7 ' + gPct + P;
      h += '<tr class="assoc-group-header" onclick="togGroup(this)"><td colspan="8">' + svgBadgeChev + ' ' + grp.name;
      h += ' <span style="color:#999;font-weight:400;font-size:.75rem;margin-left:4px">(' + gInfo + ')</span></td></tr>';
      for (var mi = 0; mi < grp.members.length; mi++) {
        var u = grp.members[mi];
        var pct = totalAct > 0 ? Math.round((u.total / totalAct) * 100) : 0;
        var isI = u.level === 'Inactive';
        h += '<tr' + (isI ? ' style="opacity:.5"' : '') + '>';
        h += '<td></td>';
        h += '<td style="padding-left:12px">' + (isI ? u.name : '<strong>' + u.name + '</strong>') + '</td>';
        h += '<td><span class="mm-badge" style="background:' + u.levelStyle.bg + ';color:' + u.levelStyle.fg + '">' + u.level + '</span></td>';
        h += '<td style="color:#888">' + (u.group || '\u2014') + '</td>';
        h += '<td class="col-right">' + fmtNum(u.activities) + '</td><td class="col-right">' + fmtNum(u.documents) + '</td>';
        h += '<td class="col-right"><strong>' + fmtNum(u.total) + '</strong></td>';
        h += '<td class="col-right">' + (isI ? '\u2014' : mmBarCell(pct, u.levelColor)) + '</td></tr>';
      }
    }
    h += '</tbody></table>';
  }
  h += '</div>';
  return h;
}

function mmGroupCard(cnt, label, desc, totalUsers, totalAct, groupAct, color) {
  var uP = totalUsers > 0 ? Math.round((cnt / totalUsers) * 100) : 0;
  var aP = totalAct > 0 ? Math.round((groupAct / totalAct) * 100) : 0;
  return '<div class="mm-user-group"><div class="mm-ug-count" style="color:' + color + '">' + cnt + '</div>'
    + '<div class="mm-ug-label"><span class="mm-ug-dot" style="background:' + color + '"></span>' + label + '</div>'
    + '<div class="mm-ug-desc">' + desc + '</div>'
    + '<div class="mm-ug-share"><i style="width:' + uP + P + ';background:' + color + '"></i></div>'
    + '<div class="mm-ug-pct">' + uP + P + ' of users \u00B7 ' + aP + P + ' of activities</div></div>';
}
function mmGroupCardInactive(cnt, label, desc, totalUsers, color) {
  var uP = totalUsers > 0 ? Math.round((cnt / totalUsers) * 100) : 0;
  return '<div class="mm-user-group"><div class="mm-ug-count" style="color:' + color + '">' + cnt + '</div>'
    + '<div class="mm-ug-label"><span class="mm-ug-dot" style="background:' + color + '"></span>' + label + '</div>'
    + '<div class="mm-ug-desc">' + desc + '</div>'
    + '<div class="mm-ug-share"><i style="width:' + uP + P + ';background:' + color + '"></i></div>'
    + '<div class="mm-ug-pct">' + uP + P + ' of users</div></div>';
}
function mmBarCell(pct, color) {
  return '<div class="bar-cell"><div style="width:80px;height:8px;background:#eee;border-radius:4px;overflow:hidden"><div style="height:100%;width:' + pct + P + ';background:' + color + ';border-radius:4px 0 0 4px"></div></div><span class="pct-text" style="color:' + color + '">' + pct + P + '</span></div>';
}
function mmGroupUsers(users) {
  var gm = {}, go = [];
  for (var i = 0; i < users.length; i++) {
    var gn = users[i].group || 'No Group';
    if (!gm[gn]) { gm[gn] = { name: gn, members: [], total: 0 }; go.push(gn); }
    gm[gn].members.push(users[i]); gm[gn].total += users[i].total;
  }
  var r = []; for (var i = 0; i < go.length; i++) r.push(gm[go[i]]);
  r.sort(function(a, b) { return b.total - a.total; });
  for (var i = 0; i < r.length; i++) r[i].members.sort(function(a, b) { return b.total - a.total; });
  return r;
}
