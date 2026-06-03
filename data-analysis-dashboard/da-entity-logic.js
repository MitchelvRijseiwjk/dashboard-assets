// === UDEF RENDER ===
function renderUdef(ent, idx, entityKey) {
  var h = '<div class="entity-card">';
  h += '<div class="entity-header"><div class="entity-info"><h3>' + ent.name + '</h3> <span class="field-count">\u2014 ' + ent.fields.length + ' active fields</span></div>';
  h += '<span class="record-badge">' + fmtNum(ent.total) + ' records</span></div>';
  if (ent.fields.length === 0) {
    h += '<div style="padding:20px;color:#888"><em>No extra fields configured</em></div>';
  } else {
    ent.fields.sort(function(a,b) { return b.percent - a.percent; });
    var udefCfg = (typeof daSettings !== 'undefined' && entityKey) ? (daSettings.getSettings(entityKey).udefFieldConfig || {}) : {};
    var Q = String.fromCharCode(39);
    h += '<table class="data-table" id="tbl-u' + idx + '">';
    h += '<thead><tr>';
    h += '<th onclick="sortT(' + Q + 'tbl-u' + idx + Q + ',0)">Field Label <span class="sort-arrow">' + svgSortN + '</span></th>';
    h += '<th onclick="sortT(' + Q + 'tbl-u' + idx + Q + ',1)">Type <span class="sort-arrow">' + svgSortN + '</span></th>';
    h += '<th class="col-right" style="width:110px" onclick="sortT(' + Q + 'tbl-u' + idx + Q + ',2)">Filled <span class="sort-arrow">' + svgSortN + '</span></th>';
    h += '<th style="text-align:center;width:90px">Importance</th>';
    h += '<th class="col-right" style="width:130px" onclick="sortT(' + Q + 'tbl-u' + idx + Q + ',4)">Completeness <span class="sort-arrow active">' + svgSortD + '</span></th>';
    h += '</tr></thead><tbody>';
    for (var fi = 0; fi < ent.fields.length; fi++) {
      var f = ent.fields[fi];
      var pid = f.progId || f.label || ('f' + fi);
      var imp = udefCfg[pid] || 'normal';
      var rc = '';
      if (f.filled === 0) rc = 'unused';
      var dimStyle = imp === 'excluded' ? ' style="opacity:0.4"' : '';
      h += '<tr class="' + rc + '"' + dimStyle + '>';
      h += '<td data-sort-value="' + f.label + '">' + f.label + '</td>';
      h += '<td data-sort-value="' + f.type + '"><span class="type-badge">' + f.type + '</span>';
      if (f.items && f.items.length > 0) {
        ddCnt++;
        var did = 'dd-' + ddCnt;
        h += ' <span class="dd-toggle" id="btn-' + did + '" onclick="togDD(' + Q + did + Q + ',' + f.items.length + ')">';
        h += svgDDChev + ' Show (' + f.items.length + ')</span>';
        h += '<div class="dd-panel" id="' + did + '">';
        for (var ii = 0; ii < f.items.length; ii++) {
          var itm = f.items[ii];
          var ic = '';
          if (itm.c === 0) ic = 'unused';
          h += '<div class="dd-item ' + ic + '">';
          h += '<span>' + itm.n + '</span>';
          h += '<span>' + itm.c + ' <span class="dd-pct">' + itm.p.toFixed(1) + P + '</span></span>';
          h += '</div>';
        }
        h += '</div>';
      }
      h += '</td>';
      h += '<td class="col-right" data-sort-value="' + f.filled + '">' + f.filled + ' / ' + ent.total + '</td>';
      h += '<td style="text-align:center"><span class="imp-badge ' + imp + '">' + imp + '</span></td>';
      h += '<td class="col-right" data-sort-value="' + f.percent + '">' + barCell(f.percent, '') + '</td>';
      h += '</tr>';
    }
    h += '</tbody></table>';
  }
  h += '</div>';
  return h;
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

// Add CSS animation for spinner (inject once)
(function() {
  if (document.getElementById('progressiveLoadStyles')) return;
  var style = document.createElement('style');
  style.id = 'progressiveLoadStyles';
  style.textContent = '.section-loaded{animation:fadeIn .3s ease}' +
    '@keyframes fadeIn{from{opacity:0;transform:translateY(8px)}to{opacity:1;transform:translateY(0)}}';
  document.head.appendChild(style);
})();

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

function renderEntityOverview(key, d) {
  overviewData[key] = d;
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

function loadCompanyDetails(cb) {
  fetchCompanyDetails(cb);
}

function fetchCompanyDetails(cb) {
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
    if (pending <= 0 && cb) cb();
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
  var sel = document.getElementById('companyDetailCatFilter');
  if (!sel) return;
  while (sel.options.length > 1) sel.remove(1);
  if (overviewData['company'] && overviewData['company'].distributions) {
    var cats = overviewData['company'].distributions[0];
    if (cats && cats.items) {
      var sorted = cats.items.slice().sort(function(a,b) { return b.count - a.count; });
      for (var ci = 0; ci < sorted.length; ci++) {
        if (sorted[ci].count > 0 && sorted[ci].name !== '(No value)') {
          var opt = document.createElement('option');
          opt.value = sorted[ci].name;
          opt.textContent = sorted[ci].name + ' (' + fmtNum(sorted[ci].count) + ')';
          sel.appendChild(opt);
        }
      }
      for (var ci = 0; ci < sorted.length; ci++) {
        if (sorted[ci].name === '(No value)' && sorted[ci].count > 0) {
          var opt2 = document.createElement('option');
          opt2.value = '__none__';
          opt2.textContent = '(No category) (' + fmtNum(sorted[ci].count) + ')';
          sel.appendChild(opt2);
        }
      }
    }
  }
  if (companyDetailCatValue) sel.value = companyDetailCatValue;
  var resetBtn = document.getElementById('companyFilterReset');
  if (resetBtn) resetBtn.style.display = companyDetailCatValue ? '' : 'none';
  var bar = document.getElementById('companyCatFilterBar');
  if (bar) bar.classList.toggle('active', !!companyDetailCatValue);
}

function reloadCompanyDetails() {
  var el = document.getElementById('companyDetailContent');
  if (el) el.innerHTML = '<div style="text-align:center;padding:40px;color:#999">Loading...</div>';
  var crossEl = document.getElementById('companyCrossContent');
  if (crossEl) crossEl.innerHTML = '<div style="text-align:center;padding:40px;color:#999">Loading...</div>';
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
    funnelSteps.push({ label: 'With ' + pipeLabel, count: f.withPersonActivitySale, from: f.withPersonActivity, desc: 'Active + open ' + pipeLabel.toLowerCase() });
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
    h += '<td><span style="display:inline-block;width:10px;height:10px;border-radius:50%;background:' + seg.color + ';margin-right:8px;vertical-align:middle"></span>';
    h += '<span style="font-weight:500">' + seg.name + '</span></td>';
    h += '<td class="col-right">' + fmtNum(seg.count) + '</td>';
    h += '<td style="color:var(--so-text-muted);font-size:.82rem">' + seg.description + '</td>';
    h += '<td class="col-right">' + barCell(segPct, seg.color) + '</td>';
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
    compData.stageProgression = o.withActivities || 0; // approximate
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
  project:  { dq: 'Field completeness across records', int: 'Missing members and activities', adopt: 'Activities and member engagement' }
};

function renderScoreBanner(key) {
  var el = document.getElementById(key + 'OverviewContent');
  if (!el) return;
  var scores = refreshEntityScores(key);
  if (!scores) return;

  var existing = el.querySelector('.scores-banner');
  if (existing) existing.parentNode.removeChild(existing);

  var hasAnything = scores.health || scores.dq || scores.integrity || scores.adoption;
  if (!hasAnything) return;

  var desc = sbDescs[key] || sbDescs.company;
  var h = '<div class="scores-banner">';

  // Left: Overall Health ring
  if (scores.health) {
    var hc = slColor(scores.health.total);
    h += '<div class="sb-main">';
    h += '<div class="sb-ring sb-ring-lg" style="--score:' + scores.health.total + ';--color:' + hc + '">';
    h += '<span class="sb-val sb-val-lg">' + scores.health.total + '<small>%</small></span>';
    h += '</div>';
    h += '<div class="sb-main-lbl">Overall Health</div>';
    h += '</div>';
  }

  h += '<div class="sb-divider"></div>';

  // Right: 3 sub-score rows
  var subs = [];
  if (scores.dq) subs.push({ label: 'Data Quality', desc: desc.dq, score: scores.dq.total });
  if (scores.integrity) subs.push({ label: 'Data Integrity', desc: desc.int, score: scores.integrity.total });
  if (scores.adoption) subs.push({ label: 'Adoption', desc: desc.adopt, score: scores.adoption.total });

  h += '<div class="sb-list">';
  for (var i = 0; i < subs.length; i++) {
    var s = subs[i];
    var col = slColor(s.score);
    h += '<div class="sb-row">';
    h += '<div class="sb-row-info"><div class="sb-row-lbl">' + s.label + '</div><div class="sb-row-desc">' + s.desc + '</div></div>';
    h += '<div class="sb-row-val" style="color:' + col + '">' + s.score + '%</div>';
    h += '<div class="sb-row-bar"><div class="sb-row-fill" style="width:' + s.score + '%;background:' + col + '"></div></div>';
    h += '</div>';
  }
  h += '</div>';

  h += '</div>';
  el.insertAdjacentHTML('afterbegin', h);
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
        h += '<div class="entity-header"><div class="entity-info"><h3>Data Quality Issues</h3></div>';
        h += '<span class="record-badge">' + fmtNum(total) + ' companies</span></div>';
        h += '<table class="data-table"><thead><tr><th>Issue</th><th class="col-right" style="width:80px">Count</th><th class="col-right" style="width:130px">' + P + '</th></tr></thead><tbody>';
        for (var i = 0; i < allIssues.length; i++) {
          var iss = allIssues[i];
          var isActive = qiFields.indexOf(iss.key) >= 0;
          var pct = Math.round((iss.val / total) * 1000) / 10;
          var col = slColorInv(pct);
          var dimStyle = isActive ? '' : ' style="opacity:0.4"';
          var badge = isActive ? '' : ' <span style="font-size:.65rem;color:var(--so-text-muted)">(not in score)</span>';
          h += '<tr' + dimStyle + '><td>' + iss.label + badge + '</td>';
          h += '<td class="col-right">' + fmtNum(iss.val) + '</td>';
          h += '<td class="col-right">' + barCell(pct, col) + '</td></tr>';
        }
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
        h += '<table class="data-table"><thead><tr><th>Field</th><th class="col-right" style="width:80px">Filled</th><th style="text-align:center;width:90px">Importance</th><th class="col-right" style="width:130px">Completeness</th></tr></thead><tbody>';
        for (var j = 0; j < def.stdFields.length; j++) {
          var fk = def.stdFields[j].key;
          var label = def.stdFields[j].label;
          var imp = cfg.stdFieldConfig[fk] || 'excluded';
          var val = daSettings.getCompletenessValue(fk, c, q, total);
          var pct = Math.round((val / total) * 1000) / 10;
          var dimStyle = imp === 'excluded' ? ' style="opacity:0.4"' : '';
          var badgeHtml = '<span class="imp-badge ' + imp + '">' + imp + '</span>';
          h += '<tr' + dimStyle + '><td>' + label + '</td>';
          h += '<td class="col-right">' + fmtNum(val) + '</td>';
          h += '<td style="text-align:center">' + badgeHtml + '</td>';
          h += '<td class="col-right">' + barCell(pct, '') + '</td></tr>';
        }
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
      h += '<table class="data-table"><thead><tr><th>Field</th><th class="col-right" style="width:80px">Filled</th><th style="text-align:center;width:90px">Importance</th><th class="col-right" style="width:130px">Completeness</th></tr></thead><tbody>';
      for (var fi = 0; fi < def.stdFields.length; fi++) {
        var fk = def.stdFields[fi].key;
        var fLabel = def.stdFields[fi].label;
        var imp = cfg.stdFieldConfig[fk] || 'excluded';
        var fVal = cpl[fk] || 0;
        var fPct = Math.round((fVal / eTotal) * 1000) / 10;
        var dimStyle = imp === 'excluded' ? ' style="opacity:0.4"' : '';
        var badgeHtml = '<span class="imp-badge ' + imp + '">' + imp + '</span>';
        h += '<tr' + dimStyle + '><td>' + fLabel + '</td>';
        h += '<td class="col-right">' + fmtNum(fVal) + '</td>';
        h += '<td style="text-align:center">' + badgeHtml + '</td>';
        h += '<td class="col-right">' + barCell(fPct, '') + '</td></tr>';
      }
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
function renderAdoptionTab(key) {
  var scores = entityScores[key];
  if (!scores) return;

  var ov = overviewData[key];
  var total = (ov && ov.overview) ? (ov.overview.total || 0) : 0;
  var h = '';

  // --- Adoption component breakdown ---
  if (scores.adoption && scores.adoption.details && scores.adoption.details.length > 0) {
    var ad = scores.adoption;
    var ac = slColor(ad.total);
    var data = gatherEntityData(key);
    h += '<div class="entity-card">';
    h += '<div class="entity-header"><div class="entity-info">';
    h += '<h3>Adoption Score</h3>';
    h += '</div>';
    h += '<span class="record-badge" style="background:' + ac + ';color:#fff;font-size:1.1rem;padding:4px 14px;border-radius:12px">' + ad.total + P + '</span></div>';
    h += '<table class="data-table"><thead><tr><th>Component</th><th class="col-right" style="width:80px">Count</th><th class="col-right" style="width:80px">of Total</th><th style="text-align:center;width:90px">Weight</th><th class="col-right" style="width:130px">Completeness</th></tr></thead><tbody>';
    for (var i = 0; i < ad.details.length; i++) {
      var d = ad.details[i];
      var cnt = (data && data.componentData) ? (data.componentData[d.key] || 0) : 0;
      var col = slColor(d.pct);
      var wCls = d.weight === 'high' ? 'required' : (d.weight === 'medium' ? 'normal' : 'excluded');
      h += '<tr><td>' + d.label + '</td>';
      h += '<td class="col-right">' + fmtNum(cnt) + '</td>';
      h += '<td class="col-right">' + fmtNum(data ? data.adoptionTotal : total) + '</td>';
      h += '<td style="text-align:center"><span class="imp-badge ' + wCls + '">' + d.weight + '</span></td>';
      h += '<td class="col-right">' + barCell(d.pct, col) + '</td></tr>';
    }
    h += '</tbody></table></div>';
  }

  // --- Integrity check breakdown ---
  if (scores.integrity && scores.integrity.details && scores.integrity.details.length > 0) {
    var ig = scores.integrity;
    var ic = slColor(ig.total);
    var data2 = gatherEntityData(key);
    h += '<div class="entity-card" style="margin-top:18px">';
    h += '<div class="entity-header"><div class="entity-info">';
    h += '<h3>Data Integrity</h3>';
    h += '</div>';
    h += '<span class="record-badge" style="background:' + ic + ';color:#fff;font-size:1.1rem;padding:4px 14px;border-radius:12px">' + ig.total + P + '</span></div>';
    h += '<table class="data-table"><thead><tr><th>Check</th><th class="col-right" style="width:80px">Affected</th><th class="col-right" style="width:80px">of Total</th><th style="text-align:center;width:90px">Weight</th><th class="col-right" style="width:130px">Completeness</th></tr></thead><tbody>';
    for (var j = 0; j < ig.details.length; j++) {
      var c = ig.details[j];
      var cnt2 = (data2 && data2.checkData) ? (data2.checkData[c.key] || 0) : 0;
      var col2 = slColorInv(c.affected);
      var iTotal = (data2 && data2.integrityTotal) ? data2.integrityTotal : total;
      var wCls2 = c.weight === 'high' ? 'required' : (c.weight === 'medium' ? 'normal' : 'excluded');
      h += '<tr><td>' + c.label + '</td>';
      h += '<td class="col-right">' + fmtNum(cnt2) + '</td>';
      h += '<td class="col-right">' + fmtNum(iTotal) + '</td>';
      h += '<td style="text-align:center"><span class="imp-badge ' + wCls2 + '">' + c.weight + '</span></td>';
      h += '<td class="col-right">' + barCell(c.affected, col2) + '</td></tr>';
    }
    h += '</tbody></table></div>';
  }

  if (!h) return;

  if (key === 'company') {
    // Insert score tables AFTER filter elements (dateFilterNotice + cat-filter-bar)
    var crossEl = document.getElementById('companyCrossContent');
    if (crossEl) {
      var wrapper = document.createElement('div');
      wrapper.className = 'adoption-score-tables';
      wrapper.style.marginBottom = '18px';
      wrapper.innerHTML = h;
      // Remove any previous score tables
      var prev = crossEl.querySelector('.adoption-score-tables');
      if (prev) prev.parentNode.removeChild(prev);
      // Find the last filter element and insert after it
      var catBar = crossEl.querySelector('.cat-filter-bar');
      var filterNotice = crossEl.querySelector('.filter-notice');
      var insertAfter = catBar || filterNotice;
      if (insertAfter) {
        crossEl.insertBefore(wrapper, insertAfter.nextSibling);
      } else {
        crossEl.insertBefore(wrapper, crossEl.firstChild);
      }
    }
  } else {
    var el = document.getElementById(key + 'AdoptionContent');
    if (el) el.innerHTML = dateFilterNotice() + h;
  }
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
      { val: ah.active6m, col: 'var(--sl-good)', label: 'Active (6m)' },
      { val: ah.dormant12m, col: 'var(--sl-ok)', label: 'Cooling (6\u201312m)' },
      { val: ah.dormantOlder, col: 'var(--sl-warn)', label: 'Dormant (>12m)' },
      { val: ah.noActivity, col: 'var(--sl-bad)', label: 'No Activity' }
    ];
    h += '<div class="stacked-bar" style="height:28px;border-radius:6px">';
    for (var i = 0; i < ahParts.length; i++) {
      var pct = ahParts[i].val / total * 100;
      if (pct > 0) {
        var segLabel = pct >= 10 ? '<span style="font-size:.72rem;color:#fff;font-weight:600">' + fmtNum(ahParts[i].val) + ' (' + Math.round(pct) + P + ')</span>' : '';
        h += '<div class="stacked-segment" style="width:' + pct + P + ';background:' + ahParts[i].col + ';display:flex;align-items:center;justify-content:center;overflow:hidden;border-radius:0">' + segLabel + '</div>';
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

  // 1b. REGISTRATION TREND (with active overlay)
  var trend = d.trend;
  var trendMonthly = d.trendMonthly;
  if (trend && trend.length > 0) {
    // Decide: monthly vs yearly based on filter span
    var useMonthly = false;
    var chartData = [];
    var firstIdx = 0;

    // Check if filter is ≤24 months (yearly data has ≤2 non-zero years)
    var nonZeroYears = 0;
    for (var i = 0; i < trend.length; i++) { if (trend[i].count > 0) nonZeroYears++; }
    if (nonZeroYears <= 2 && trendMonthly && trendMonthly.length > 0) {
      useMonthly = true;
      for (var i = 0; i < trendMonthly.length; i++) {
        chartData.push({ label: trendMonthly[i].month.substring(2), count: trendMonthly[i].count, active: trendMonthly[i].active || 0 });
      }
    } else {
      for (var i = 0; i < trend.length; i++) {
        chartData.push({ label: '' + trend[i].year, count: trend[i].count, active: trend[i].active || 0 });
      }
    }

    // Trim leading zeros
    firstIdx = 0;
    for (var i = 0; i < chartData.length; i++) { if (chartData[i].count > 0 || chartData[i].active > 0) { firstIdx = i; break; } }
    var visibleData = chartData.slice(firstIdx);
    var maxCount = 0;
    for (var i = 0; i < visibleData.length; i++) {
      if (visibleData[i].count > maxCount) maxCount = visibleData[i].count;
      if (visibleData[i].active > maxCount) maxCount = visibleData[i].active;
    }

    if (maxCount > 0 && visibleData.length > 1) {
      var beforeTotal = d.trendBefore || 0;
      if (!useMonthly) {
        for (var i = 0; i < firstIdx; i++) { beforeTotal += chartData[i].count; }
      }
      h += '<div class="detail-section">';
      h += '<div class="detail-section-head">';
      h += secHead(useMonthly ? 'New Registrations Per Month' : 'New Registrations Per Year');
      if (beforeTotal > 0 && !useMonthly) {
        h += '<span class="record-badge">' + fmtNum(beforeTotal) + ' before ' + visibleData[0].label + '</span>';
      }
      h += '</div>';
      h += '<div style="font-size:.78rem;color:#999;margin:-4px 0 8px">Dashed line = retention: how many registered companies still have activity within the last 12 months</div>';
      var cW = 960, cH = 270, padL = 40, padR = 10, padT = 20, padB = 60;
      var plotW = cW - padL - padR, plotH = cH - padT - padB;
      var dataInset = 15;
      var niceMax = maxCount;
      var mag = Math.pow(10, Math.floor(Math.log10(maxCount)));
      var options = [1, 1.5, 2, 2.5, 3, 4, 5, 8, 10];
      for (var oi = 0; oi < options.length; oi++) { if (options[oi] * mag >= maxCount) { niceMax = options[oi] * mag; break; } }
      var ySteps = 4;
      var yStep = niceMax / ySteps;
      var step = (plotW - 2 * dataInset) / (visibleData.length - 1);

      // Compute registration line points
      var pts = [];
      for (var i = 0; i < visibleData.length; i++) {
        var px = padL + dataInset + i * step;
        var py = padT + plotH - (visibleData[i].count / niceMax) * plotH;
        pts.push(px.toFixed(1) + ',' + py.toFixed(1));
      }
      // Compute active overlay points
      var hasActiveData = false;
      var aPts = [];
      for (var i = 0; i < visibleData.length; i++) {
        if (visibleData[i].active > 0) hasActiveData = true;
        var px = padL + dataInset + i * step;
        var py = padT + plotH - (visibleData[i].active / niceMax) * plotH;
        aPts.push(px.toFixed(1) + ',' + py.toFixed(1));
      }

      var areaPath = 'M' + (padL + dataInset) + ',' + (padT + plotH) + ' L' + pts.join(' L') + ' L' + (padL + dataInset + (visibleData.length - 1) * step).toFixed(1) + ',' + (padT + plotH) + ' Z';
      h += '<svg viewBox="0 0 ' + cW + ' ' + cH + '" style="width:100%;height:auto;display:block;max-height:280px">';
      // Y grid
      for (var gi = 0; gi <= ySteps; gi++) {
        var gy = padT + plotH - (gi / ySteps) * plotH;
        var yVal = Math.round(gi * yStep);
        h += '<line x1="' + padL + '" y1="' + gy.toFixed(1) + '" x2="' + (cW - padR) + '" y2="' + gy.toFixed(1) + '" stroke="#e0dfdc" stroke-width="1"/>';
        h += '<text x="' + (padL - 8) + '" y="' + (gy + 4).toFixed(1) + '" text-anchor="end" fill="#999" font-size="11" font-family="DM Sans,sans-serif">' + fmtNum(yVal) + '</text>';
      }
      // Registration area + line
      h += '<path d="' + areaPath + '" fill="rgba(22,91,112,0.06)"/>';
      h += '<polyline points="' + pts.join(' ') + '" fill="none" stroke="var(--so-green)" stroke-width="2.5" stroke-linejoin="round" stroke-linecap="round"/>';
      // Active overlay line (dashed)
      if (hasActiveData) {
        h += '<polyline points="' + aPts.join(' ') + '" fill="none" stroke="var(--sl-good)" stroke-width="2" stroke-dasharray="6,4" stroke-linejoin="round" stroke-linecap="round"/>';
      }
      // Data points + labels
      for (var i = 0; i < visibleData.length; i++) {
        var xy = pts[i].split(',');
        h += '<circle cx="' + xy[0] + '" cy="' + xy[1] + '" r="4" fill="var(--so-green)" stroke="#fff" stroke-width="2"/>';
        if (visibleData[i].count > 0) {
          h += '<text x="' + xy[0] + '" y="' + (parseFloat(xy[1]) - 10).toFixed(1) + '" text-anchor="middle" fill="var(--so-charcoal)" font-size="10" font-weight="600" font-family="DM Sans,sans-serif">' + fmtNum(visibleData[i].count) + '</text>';
        }
        // Active dot + label (below the dot, skip when same as count)
        if (hasActiveData) {
          var axy = aPts[i].split(',');
          h += '<circle cx="' + axy[0] + '" cy="' + axy[1] + '" r="3" fill="var(--sl-good)" stroke="#fff" stroke-width="1.5"/>';
          if (visibleData[i].active !== visibleData[i].count) {
            h += '<text x="' + axy[0] + '" y="' + (parseFloat(axy[1]) + 14).toFixed(1) + '" text-anchor="middle" fill="var(--sl-good)" font-size="9" font-weight="600" font-family="DM Sans,sans-serif">' + visibleData[i].active + '</text>';
          }
        }
        // X labels (show all, rotate monthly for readability)
        var displayLabel = useMonthly ? visibleData[i].label.replace('-', '/') : visibleData[i].label;
        if (useMonthly) {
          h += '<text x="' + xy[0] + '" y="' + (padT + plotH + 14) + '" text-anchor="end" fill="#999" font-size="9" font-family="DM Sans,sans-serif" transform="rotate(-45,' + xy[0] + ',' + (padT + plotH + 14) + ')">' + displayLabel + '</text>';
        } else {
          h += '<text x="' + xy[0] + '" y="' + (padT + plotH + 22) + '" text-anchor="middle" fill="#999" font-size="10" font-family="DM Sans,sans-serif">' + displayLabel + '</text>';
        }
      }
      h += '</svg>';
      // Legend
      h += '<div style="display:flex;gap:16px;margin-top:4px;padding-left:' + padL + 'px">';
      h += '<span style="font-size:.75rem;color:#666;display:flex;align-items:center;gap:4px"><span style="width:12px;height:3px;background:var(--so-green);border-radius:2px"></span> Registered</span>';
      if (hasActiveData) {
        h += '<span style="font-size:.75rem;color:#666;display:flex;align-items:center;gap:4px"><span style="width:12px;height:0;border-top:2px dashed var(--sl-good)"></span> Retention (of which still active today)</span>';
      }
      h += '</div>';
      h += '</div>';
    }
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

var MM_GREEN = '#2e7d32';
var MM_BLUE = '#1565c0';
var MM_ORANGE = '#f57c00';
var MM_RED = '#c62828';

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

  el.innerHTML = h;
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

function renderMomentumChart(monthly) {
  if (!monthly || monthly.length === 0) return '';
  var maxAct = 0, maxUsers = 0;
  for (var i = 0; i < monthly.length; i++) {
    if (monthly[i].activities > maxAct) maxAct = monthly[i].activities;
    if (monthly[i].documents > maxAct) maxAct = monthly[i].documents;
    if (monthly[i].activeUsers > maxUsers) maxUsers = monthly[i].activeUsers;
  }

  // Nice axis max (same pattern as Company trend chart)
  var niceMaxAct = maxAct;
  var mag = Math.pow(10, Math.floor(Math.log10(maxAct || 1)));
  var opts = [1, 1.5, 2, 2.5, 3, 4, 5, 8, 10];
  for (var oi = 0; oi < opts.length; oi++) { if (opts[oi] * mag >= maxAct) { niceMaxAct = opts[oi] * mag; break; } }
  var niceMaxUsr = maxUsers;
  var magU = Math.pow(10, Math.floor(Math.log10(maxUsers || 1)));
  for (var oi = 0; oi < opts.length; oi++) { if (opts[oi] * magU >= maxUsers) { niceMaxUsr = opts[oi] * magU; break; } }

  // SVG dimensions with padding (matches Company chart structure)
  var cW = 960, cH = 270, padL = 48, padR = 40, padT = 16, padB = 40;
  var plotW = cW - padL - padR, plotH = cH - padT - padB;
  var n = monthly.length;
  var ySteps = 4;
  var yStepAct = niceMaxAct / ySteps;
  var yStepUsr = niceMaxUsr / ySteps;

  // Compute line points
  var actPts = [], docPts = [], usrPts = [];
  for (var i = 0; i < n; i++) {
    var px = padL + (n > 1 ? (i / (n - 1)) * plotW : plotW / 2);
    var yA = niceMaxAct > 0 ? padT + plotH - (monthly[i].activities / niceMaxAct) * plotH : padT + plotH;
    var yD = niceMaxAct > 0 ? padT + plotH - (monthly[i].documents / niceMaxAct) * plotH : padT + plotH;
    var yU = niceMaxUsr > 0 ? padT + plotH - (monthly[i].activeUsers / niceMaxUsr) * plotH : padT + plotH;
    actPts.push(px.toFixed(1) + ',' + yA.toFixed(1));
    docPts.push(px.toFixed(1) + ',' + yD.toFixed(1));
    usrPts.push(px.toFixed(1) + ',' + yU.toFixed(1));
  }

  // Build SVG — same style as Company: width:100%;height:auto
  var svg = '<svg viewBox="0 0 ' + cW + ' ' + cH + '" style="width:100%;height:auto;display:block;max-height:280px">';

  // Y grid lines + left labels (activity axis) + right labels (user axis)
  for (var gi = 0; gi <= ySteps; gi++) {
    var gy = padT + plotH - (gi / ySteps) * plotH;
    var actVal = Math.round(gi * yStepAct);
    var usrVal = Math.round(gi * yStepUsr);
    svg += '<line x1="' + padL + '" y1="' + gy.toFixed(1) + '" x2="' + (cW - padR) + '" y2="' + gy.toFixed(1) + '" stroke="#e0dfdc" stroke-width="1"/>';
    svg += '<text x="' + (padL - 8) + '" y="' + (gy + 4).toFixed(1) + '" text-anchor="end" fill="#999" font-size="10" font-family="DM Sans,sans-serif">' + mmFmtK(actVal) + '</text>';
    svg += '<text x="' + (cW - padR + 8) + '" y="' + (gy + 4).toFixed(1) + '" text-anchor="start" fill="' + MM_BLUE + '" font-size="10" font-family="DM Sans,sans-serif" opacity=".6">' + usrVal + '</text>';
  }

  // Area fill under activities
  var areaPath = 'M' + actPts[0] + ' L' + actPts.join(' L') + ' L' + (padL + plotW).toFixed(1) + ',' + (padT + plotH) + ' L' + padL + ',' + (padT + plotH) + ' Z';
  svg += '<path d="' + areaPath + '" fill="rgba(46,125,50,0.06)"/>';

  // Lines
  svg += '<polyline points="' + actPts.join(' ') + '" fill="none" stroke="' + MM_GREEN + '" stroke-width="2.5" stroke-linejoin="round" stroke-linecap="round"/>';
  svg += '<polyline points="' + docPts.join(' ') + '" fill="none" stroke="' + MM_ORANGE + '" stroke-width="2" stroke-linejoin="round" stroke-linecap="round" stroke-dasharray="6,4"/>';
  svg += '<polyline points="' + usrPts.join(' ') + '" fill="none" stroke="' + MM_BLUE + '" stroke-width="2" stroke-linejoin="round" stroke-linecap="round"/>';

  // User dots (every 4th + first + last)
  for (var i = 0; i < n; i++) {
    if (i === 0 || i === n - 1 || i % 4 === 0) {
      var p = usrPts[i].split(',');
      svg += '<circle cx="' + p[0] + '" cy="' + p[1] + '" r="3" fill="' + MM_BLUE + '" stroke="#fff" stroke-width="1.5"/>';
    }
  }

  // X labels — show every ~2 months
  var xStep = Math.max(1, Math.floor(n / 12));
  for (var i = 0; i < n; i += xStep) {
    var px = padL + (n > 1 ? (i / (n - 1)) * plotW : plotW / 2);
    svg += '<text x="' + px.toFixed(1) + '" y="' + (padT + plotH + 20) + '" text-anchor="middle" fill="#999" font-size="10" font-family="DM Sans,sans-serif">' + monthly[i].label + '</text>';
  }
  if ((n - 1) % xStep !== 0) {
    var lastPx = padL + plotW;
    svg += '<text x="' + lastPx.toFixed(1) + '" y="' + (padT + plotH + 20) + '" text-anchor="middle" fill="#999" font-size="10" font-family="DM Sans,sans-serif">' + monthly[n - 1].label + '</text>';
  }

  // Invisible hover columns for tooltip
  var colW = plotW / n;
  for (var i = 0; i < n; i++) {
    var cx = padL + (n > 1 ? (i / (n - 1)) * plotW : plotW / 2);
    svg += '<rect x="' + (cx - colW / 2).toFixed(1) + '" y="' + padT + '" width="' + colW.toFixed(1) + '" height="' + plotH + '" fill="transparent" class="mm-hover-col" data-idx="' + i + '"/>';
  }

  svg += '</svg>';

  // Legend (inline, same pattern as Company)
  var legend = '<div style="display:flex;gap:16px;margin-top:4px;padding-left:' + padL + 'px">';
  legend += '<span style="font-size:.75rem;color:#666;display:flex;align-items:center;gap:4px"><span style="width:12px;height:3px;background:' + MM_GREEN + ';border-radius:2px"></span> Activities</span>';
  legend += '<span style="font-size:.75rem;color:#666;display:flex;align-items:center;gap:4px"><span style="width:12px;height:0;border-top:2px dashed ' + MM_ORANGE + '"></span> Documents</span>';
  legend += '<span style="font-size:.75rem;color:#666;display:flex;align-items:center;gap:4px"><span style="width:10px;height:10px;border-radius:50%;background:' + MM_BLUE + '"></span> Active Users (right axis)</span>';
  legend += '</div>';

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
      h += '<tr><td><span class="mm-rank" style="background:' + u.levelColor + '">' + (i + 1) + '</span></td>';
      h += '<td><strong>' + u.name + '</strong></td>';
      h += '<td><span class="mm-badge" style="background:' + u.levelColor + '">' + u.level + '</span></td>';
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
        h += '<td><span class="mm-badge" style="background:' + u.levelColor + '">' + u.level + '</span></td>';
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
  return '<div class="mm-user-group"><div class="mm-ug-count" style="color:' + color + '">' + cnt + '</div><div class="mm-ug-label">' + label + '</div><div class="mm-ug-desc">' + desc + '</div><div class="mm-ug-pct">' + uP + P + ' of users \u00B7 ' + aP + P + ' of activities</div></div>';
}
function mmGroupCardInactive(cnt, label, desc, totalUsers, color) {
  var uP = totalUsers > 0 ? Math.round((cnt / totalUsers) * 100) : 0;
  return '<div class="mm-user-group"><div class="mm-ug-count" style="color:' + color + '">' + cnt + '</div><div class="mm-ug-label">' + label + '</div><div class="mm-ug-desc">' + desc + '</div><div class="mm-ug-pct">' + uP + P + ' of users</div></div>';
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
