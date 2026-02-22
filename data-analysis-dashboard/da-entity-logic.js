// === UDEF RENDER ===
function renderUdef(ent, idx) {
  var h = '<div class="entity-card">';
  h += '<div class="entity-header"><div class="entity-info"><h3>' + ent.name + '</h3> <span class="field-count">\u2014 ' + ent.fields.length + ' active fields</span></div>';
  h += '<span class="record-badge">' + fmtNum(ent.total) + ' records</span></div>';
  if (ent.fields.length === 0) {
    h += '<div style="padding:20px;color:#888"><em>No extra fields configured</em></div>';
  } else {
    ent.fields.sort(function(a,b) { return b.percent - a.percent; });
    var Q = String.fromCharCode(39);
    h += '<table class="data-table" id="tbl-u' + idx + '">';
    h += '<thead><tr>';
    h += '<th onclick="sortT(' + Q + 'tbl-u' + idx + Q + ',0)">Field Label <span class="sort-arrow">' + svgSortN + '</span></th>';
    h += '<th onclick="sortT(' + Q + 'tbl-u' + idx + Q + ',1)">Type <span class="sort-arrow">' + svgSortN + '</span></th>';
    h += '<th class="col-right" onclick="sortT(' + Q + 'tbl-u' + idx + Q + ',2)">Filled <span class="sort-arrow">' + svgSortN + '</span></th>';
    h += '<th class="col-right" onclick="sortT(' + Q + 'tbl-u' + idx + Q + ',3)">' + P + ' <span class="sort-arrow active">' + svgSortD + '</span></th>';
    h += '<th style="width:130px">Fill Rate</th>';
    h += '</tr></thead><tbody>';
    for (var fi = 0; fi < ent.fields.length; fi++) {
      var f = ent.fields[fi];
      var rc = '';
      if (f.filled === 0) rc = 'unused';
      h += '<tr class="' + rc + '">';
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
      h += '<td class="col-right" data-sort-value="' + f.percent + '">' + f.percent.toFixed(1) + P + '</td>';
      h += '<td>' + fillBar(f.percent) + '</td>';
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
    if (barEl) barEl.style.width = '100' + P;
    if (statusEl) statusEl.textContent = 'Complete!';
    if (startEl) startEl.style.display = 'none';
    if (resultsEl) resultsEl.style.display = 'block';
    if (d && summaryEl) {
      summaryEl.textContent = d.fields.length + ' active extra fields. Click column headers to sort.';
      if (cardsEl) cardsEl.innerHTML = renderUdef(d, udefIdx);
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
    if (d && summaryEl) {
      summaryEl.textContent = d.fields.length + ' active extra fields. Click column headers to sort.';
      if (cardsEl) cardsEl.innerHTML = renderUdef(d, udefIdx);
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
    _extraCache = { tables: d.tables, data: {}, ready: false };
    var remaining = d.tables.length;
    for (var i = 0; i < d.tables.length; i++) {
      (function(tbl) {
        ajax(extraUrl + String.fromCharCode(38) + 'tableId=' + tbl.id, function(td) {
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
  style.textContent = '@keyframes spin{to{transform:rotate(360deg)}}' +
    '.section-loaded{animation:fadeIn .3s ease}' +
    '@keyframes fadeIn{from{opacity:0;transform:translateY(8px)}to{opacity:1;transform:translateY(0)}}' +
    '.progress-inline{display:flex;align-items:center;gap:8px;padding:8px 16px;margin:0 0 12px;border-radius:8px;background:var(--so-cream,#faf9f7);font-size:.8rem;color:#666}' +
    '.progress-inline .pi-bar{flex:1;height:4px;background:#e0dfdc;border-radius:2px;overflow:hidden}' +
    '.progress-inline .pi-fill{height:100%;background:var(--so-green);border-radius:2px;transition:width .3s}' +
    '.progress-inline .pi-check{color:var(--so-green);font-weight:600}';
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

  // === PROGRESSIVE: Show results immediately with loading placeholders ===
  if (progressScreen) progressScreen.style.display = 'none';
  if (subTabs) subTabs.style.display = '';
  if (resultsContainer) resultsContainer.style.display = '';

  // Set loading placeholders in each section
  var overviewEl = document.getElementById(tabKey + 'OverviewContent');
  if (overviewEl) overviewEl.innerHTML = loadingPlaceholder('Loading overview data...');

  // Data Quality tab placeholders
  var dqScoreEl = document.getElementById(tabKey + 'DQScoreContent');
  if (dqScoreEl) dqScoreEl.innerHTML = '';

  if (ec && ec.udefId > 0) {
    var udefCards = document.getElementById(tabKey + 'UdefCards');
    var udefSummary = document.getElementById(tabKey + 'UdefSummary');
    if (udefCards) udefCards.innerHTML = loadingPlaceholder('Loading extra fields...');
    if (udefSummary) udefSummary.textContent = '';
  }

  var extraCards = document.getElementById(tabKey + 'ExtraCards');
  var extraSummary = document.getElementById(tabKey + 'ExtraSummary');
  if (extraCards) extraCards.innerHTML = loadingPlaceholder('Loading extra tables...');
  if (extraSummary) extraSummary.textContent = '';

  // Adoption tab placeholders
  if (hasDetails) {
    var crossEl = document.getElementById('companyCrossContent');
    if (crossEl) crossEl.innerHTML = loadingPlaceholder('Loading engagement funnel...');
    var detailEl = document.getElementById('companyDetailContent');
    if (detailEl) detailEl.innerHTML = loadingPlaceholder('Loading adoption analysis...');
  }

  // Show inline progress bar
  var inlineProgress = document.createElement('div');
  inlineProgress.className = 'progress-inline';
  inlineProgress.id = tabKey + 'InlineProgress';
  inlineProgress.innerHTML = '<span id="' + tabKey + 'IPLabel">Loading...</span>' +
    '<div class="pi-bar"><div class="pi-fill" id="' + tabKey + 'IPBar" style="width:0"></div></div>' +
    '<span id="' + tabKey + 'IPPct">0' + P + '</span>';
  if (overviewEl && overviewEl.parentNode) {
    overviewEl.parentNode.insertBefore(inlineProgress, overviewEl);
  }

  // === Track completion of parallel loads ===
  var totalSteps = 2; // overview + extra (always)
  if (ec && ec.udefId > 0) totalSteps++;
  if (ec && ec.hasTicketFields) totalSteps++;
  if (hasDetails) totalSteps++; // details (includes funnel now)
  var completedSteps = 0;
  var dfParam = getDateFilterParam();

  function stepDone(label) {
    completedSteps++;
    var pct = Math.round((completedSteps / totalSteps) * 100);
    var barEl = document.getElementById(tabKey + 'IPBar');
    var pctEl = document.getElementById(tabKey + 'IPPct');
    var labelEl = document.getElementById(tabKey + 'IPLabel');
    if (barEl) barEl.style.width = pct + P;
    if (pctEl) pctEl.textContent = pct + P;
    if (labelEl) labelEl.textContent = label;

    // Progressively update DQ score as data becomes available
    renderDQScore(key);

    if (completedSteps >= totalSteps) {
      // All done
      if (labelEl) labelEl.innerHTML = '<span class="pi-check">\u2713 Analysis complete</span>';
      var expBtn = document.getElementById(tabKey + 'ExportBtn') || document.getElementById('ticketExportBtn');
      if (expBtn) expBtn.style.display = '';
      if (headerBtn) { headerBtn.disabled = false; headerBtn.onclick = function(){ reAnalyze(key); }; }
      // Final DQ score render with all data available
      renderDQScore(key);
      // Auto-hide progress bar after 3 seconds
      setTimeout(function() {
        var ip = document.getElementById(tabKey + 'InlineProgress');
        if (ip) { ip.style.opacity = '0'; ip.style.transition = 'opacity .5s'; setTimeout(function() { if (ip.parentNode) ip.parentNode.removeChild(ip); }, 500); }
      }, 3000);
    }
  }

  // === PARALLEL LOAD 1: Overview ===
  ajax(overviewUrl + String.fromCharCode(38) + 'entity=' + key + dfParam, function(d) {
    if (d) renderEntityOverview(key, d);
    if (overviewEl) overviewEl.style.animation = 'fadeIn .3s ease';
    stepDone('Overview loaded');
  });

  // === PARALLEL LOAD 2: UDEF / Ticket Fields ===
  if (ec && ec.udefId > 0) {
    ajax(udefUrl + String.fromCharCode(38) + 'entityId=' + ec.udefId + dfParam, function(d) {
      udefData[ec.udefId] = d;
      var cardsEl = document.getElementById(tabKey + 'UdefCards');
      var summaryEl = document.getElementById(tabKey + 'UdefSummary');
      if (d && summaryEl) {
        summaryEl.textContent = d.fields.length + ' active extra fields. Click column headers to sort.';
        if (cardsEl) { cardsEl.innerHTML = renderUdef(d, ec.udefIdx); cardsEl.style.animation = 'fadeIn .3s ease'; }
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

  // === PARALLEL LOAD 3: Company Details (company only) ===
  if (hasDetails) {
    var catParam = '';
    if (companyDetailCatValue) {
      catParam = String.fromCharCode(38) + 'categoryName=' + encodeURIComponent(companyDetailCatValue);
    }
    ajax(detailUrl + dfParam + catParam, function(d) {
      companyDetailData = d;
      if (d) renderCompanyDetails(d);
      populateDetailCatFilter();
      var countEl = document.getElementById('companyDetailFilterCount');
      if (countEl && d && d.activityHealth) {
        countEl.textContent = companyDetailCatValue ? fmtNum(d.activityHealth.total) + ' companies' : '';
      }
      var detailEl2 = document.getElementById('companyDetailContent');
      if (detailEl2) detailEl2.style.animation = 'fadeIn .3s ease';
      // Render funnel from detail data (included in same response)
      if (d && d.funnel) {
        companyCrossData = d;
        renderCrossEntityFunnel(d);
        var crossEl2 = document.getElementById('companyCrossContent');
        if (crossEl2) crossEl2.style.animation = 'fadeIn .3s ease';
      }
      stepDone('Details loaded');
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

  ajax(overviewUrl + String.fromCharCode(38) + 'entity=' + key + getDateFilterParam(), function(d) {
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

function pctCard(label, pct, desc, color) {
  var v = (typeof pct === 'number' && !isNaN(pct)) ? pct : 0;
  var h = '<div class="stat-card">';
  h += '<div class="stat-value">' + v + P + '</div>';
  h += '<div class="stat-label">' + label + '</div>';
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
  ajax(detailUrl + getDateFilterParam() + catParam, function(d) {
    companyDetailData = d;
    if (d) renderCompanyDetails(d);
    populateDetailCatFilter();
    var countEl = document.getElementById('companyDetailFilterCount');
    if (countEl && d && d.activityHealth) {
      countEl.textContent = companyDetailCatValue ? fmtNum(d.activityHealth.total) + ' companies' : '';
    }
    if (cb) cb();
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
  var badge = document.getElementById('companyFilterBadge');
  var resetBtn = document.getElementById('companyFilterReset');
  if (badge) badge.style.display = companyDetailCatValue ? '' : 'none';
  if (resetBtn) resetBtn.style.display = companyDetailCatValue ? '' : 'none';
}

function reloadCompanyDetails() {
  var el = document.getElementById('companyDetailContent');
  if (el) el.innerHTML = '<div style="text-align:center;padding:40px;color:#999">Loading...</div>';
  var crossEl = document.getElementById('companyCrossContent');
  if (crossEl) crossEl.innerHTML = '<div style="text-align:center;padding:40px;color:#999">Loading...</div>';
  fetchCompanyDetails(function() {
    renderDQScore('company');
    loadCompanyCross(null);
  });
}

function togDetailFilter(key) {
  var pop = document.getElementById(key + 'FilterPopover');
  if (!pop) return;
  var isOpen = pop.classList.contains('show');
  document.querySelectorAll('.detail-filter-popover.show').forEach(function(p) { p.classList.remove('show'); });
  if (!isOpen) pop.classList.add('show');
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

document.addEventListener('click', function(e) {
  if (!e.target.closest('.detail-filter-wrap')) {
    document.querySelectorAll('.detail-filter-popover.show').forEach(function(p) { p.classList.remove('show'); });
  }
});

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
  var m = f.metrics;
  var h = '';

  // -- CRM Health metrics with context descriptions --
  h += '<div class="detail-section">';
  h += '<div class="detail-section-head">' + secHead('CRM Health Pipeline') + '</div>';
  h += '<div class="stat-row">';
  h += pctCard('Person Coverage', m.personCoverage, 'Companies with a contact person', m.personCoverage >= 60 ? 'var(--so-meadow-dark1)' : m.personCoverage >= 40 ? 'var(--so-mango)' : 'var(--so-coral)');
  h += pctCard('Activity Rate', m.activityRate, 'With person + activity in 12 months', m.activityRate >= 50 ? 'var(--so-meadow-dark1)' : m.activityRate >= 30 ? 'var(--so-mango)' : 'var(--so-coral)');
  h += pctCard('Pipeline Rate', m.pipelineRate, 'Active companies with an open sale', m.pipelineRate >= 30 ? 'var(--so-meadow-dark1)' : m.pipelineRate >= 15 ? 'var(--so-mango)' : 'var(--so-coral)');
  h += pctCard('Overall Health', m.overallHealth, 'Person + activity + sale (fully engaged)', m.overallHealth >= 20 ? 'var(--so-meadow-dark1)' : m.overallHealth >= 10 ? 'var(--so-mango)' : 'var(--so-coral)');
  h += '</div></div>';

  // -- Engagement Funnel --
  var funnelSteps = [
    { label: 'Total Companies', count: f.total, pct: 100, drop: null },
    { label: 'With Contact Person', count: f.withPerson, pct: f.total > 0 ? Math.round(f.withPerson / f.total * 1000) / 10 : 0 },
    { label: 'With Activity (12m)', count: f.withPersonActivity, pct: f.total > 0 ? Math.round(f.withPersonActivity / f.total * 1000) / 10 : 0 },
    { label: 'With Open Sale', count: f.withPersonActivitySale, pct: f.total > 0 ? Math.round(f.withPersonActivitySale / f.total * 1000) / 10 : 0 }
  ];
  for (var i = 1; i < funnelSteps.length; i++) {
    var prev = funnelSteps[i-1].count;
    var cur = funnelSteps[i].count;
    funnelSteps[i].drop = prev > 0 ? Math.round((prev - cur) / prev * 100) : 0;
  }

  h += '<div class="entity-card">';
  h += '<div class="entity-header"><div class="entity-info"><h3>Company Engagement Funnel</h3></div>';
  h += '<span class="record-badge">' + fmtNum(f.total) + ' companies</span></div>';
  h += '<table class="data-table"><thead><tr>';
  h += '<th>Stage</th><th class="col-right">Companies</th><th class="col-right">' + P + ' of total</th><th class="col-right">vs. previous</th>';
  h += '</tr></thead><tbody>';
  for (var i = 0; i < funnelSteps.length; i++) {
    var step = funnelSteps[i];
    h += '<tr>';
    h += '<td><span style="font-weight:500">' + step.label + '</span></td>';
    h += '<td class="col-right">' + fmtNum(step.count) + '</td>';
    h += '<td class="col-right">' + barCell(step.pct, '') + '</td>';
    if (step.drop === null) {
      h += '<td class="col-right" style="color:#ccc">\u2014</td>';
    } else if (step.drop === 0) {
      h += '<td class="col-right" style="color:var(--so-meadow-dark1)">0' + P + '</td>';
    } else {
      h += '<td class="col-right" style="color:var(--so-coral-dark)">\u2212' + step.drop + P + '</td>';
    }
    h += '</tr>';
  }
  h += '</tbody></table></div>';

  // -- Segments as data-table (same pattern as Category/Business tables) --
  h += '<div class="entity-card">';
  h += '<div class="entity-header"><div class="entity-info"><h3>Company Segments</h3></div>';
  h += '<span class="record-badge">4 segments</span></div>';
  h += '<table class="data-table"><thead><tr>';
  h += '<th>Segment</th><th class="col-right">Companies</th><th class="col-right">' + P + '</th><th>Description</th>';
  h += '</tr></thead><tbody>';
  for (var si = 0; si < segs.length; si++) {
    var seg = segs[si];
    var segPct = f.total > 0 ? Math.round(seg.count / f.total * 1000) / 10 : 0;
    h += '<tr>';
    h += '<td><span style="display:inline-block;width:10px;height:10px;border-radius:50%;background:' + seg.color + ';margin-right:8px;vertical-align:middle"></span>';
    h += '<span style="font-weight:500">' + seg.name + '</span></td>';
    h += '<td class="col-right">' + fmtNum(seg.count) + '</td>';
    h += '<td class="col-right">' + barCell(segPct, seg.color) + '</td>';
    h += '<td style="color:var(--so-text-muted);font-size:.82rem">' + seg.description + '</td>';
    h += '</tr>';
  }
  h += '</tbody></table></div>';

  el.innerHTML = h;
}

// ===========================================================
// DATA QUALITY SCORE (computed frontend from available data)
// ===========================================================
function renderDQScore(key) {
  var el = document.getElementById(key + 'DQScoreContent');
  if (!el) return;

  var scores = [];
  var h = '';

  // Company-specific quality issues
  if (key === 'company' && companyDetailData && companyDetailData.quality) {
    var q = companyDetailData.quality;
    var total = companyDetailData.activityHealth ? companyDetailData.activityHealth.total : 0;
    if (total > 0) {
      h += '<div class="detail-section">';
      h += '<div class="detail-section-head">' + secHead('Data Quality Issues') + '</div>';
      h += '<div class="stat-row">';
      var issues = [
        { label: 'No contact person', val: q.noPerson },
        { label: 'No category', val: q.noCategory },
        { label: 'No business type', val: q.noBusiness },
        { label: 'No org. number', val: q.noOrgNr },
        { label: 'Unreachable', val: q.unreachable }
      ];
      for (var i = 0; i < issues.length; i++) {
        var pct = Math.round((issues[i].val / total) * 1000) / 10;
        var col = pct < 10 ? 'var(--so-meadow)' : pct < 30 ? 'var(--so-mango)' : 'var(--so-coral)';
        h += ovCard(issues[i].label, issues[i].val, total, col);
      }
      h += '</div></div>';
    }
  }

  // Overview completeness (company)
  if (key === 'company' && overviewData['company'] && overviewData['company'].completeness) {
    var ov = overviewData['company'];
    var c = ov.completeness;
    var total = ov.overview.total || 0;
    var ovLabCfg = ovLabels['company'];
    if (ovLabCfg && ovLabCfg.completeness && total > 0) {
      h += '<div class="detail-section">';
      h += '<div class="detail-section-head">' + secHead('Standard Field Completeness') + '</div>';
      h += '<div class="stat-row">';
      for (var j = 0; j < ovLabCfg.completeness.length; j++) {
        var cm = ovLabCfg.completeness[j];
        h += ovCard(cm[1], c[cm[0]], total, '');
      }
      h += '</div></div>';
    }
  }

  // DQ Score summary card
  var dqScore = computeDQScore(key);
  if (dqScore !== null) {
    var scoreColor = dqScore >= 70 ? 'var(--so-meadow)' : dqScore >= 40 ? 'var(--so-mango)' : 'var(--so-coral)';
    var scoreHtml = '<div class="dq-score-banner">';
    scoreHtml += '<div class="dq-score-ring" style="--score:' + dqScore + ';--color:' + scoreColor + '">';
    scoreHtml += '<span class="dq-score-value">' + dqScore + '</span>';
    scoreHtml += '</div>';
    scoreHtml += '<div class="dq-score-info">';
    scoreHtml += '<div class="dq-score-label">Data Quality Score</div>';
    scoreHtml += '<div class="dq-score-desc">Based on field completeness, extra field usage and data quality issues</div>';
    scoreHtml += '</div></div>';
    h = scoreHtml + h; // prepend score to top
  }

  h = dateFilterNotice() + h; // filter notice always at very top

  el.innerHTML = h;
}

function computeDQScore(key) {
  var scores = [];

  // Completeness score (company only for now)
  if (key === 'company' && overviewData['company'] && overviewData['company'].completeness) {
    var c = overviewData['company'].completeness;
    var total = overviewData['company'].overview.total || 1;
    var fields = ['orgNr', 'email', 'phone', 'address', 'webpage'];
    var sum = 0;
    for (var i = 0; i < fields.length; i++) {
      sum += (c[fields[i]] || 0) / total * 100;
    }
    scores.push(sum / fields.length);
  }

  // UDEF fill rate score
  var ec = entityConfig[key];
  if (ec && ec.udefId > 0 && udefData[ec.udefId]) {
    var ud = udefData[ec.udefId];
    if (ud.fields && ud.fields.length > 0) {
      var udefSum = 0;
      for (var i = 0; i < ud.fields.length; i++) {
        udefSum += ud.fields[i].percent;
      }
      scores.push(udefSum / ud.fields.length);
    }
  }

  // Quality issues score (company only)
  if (key === 'company' && companyDetailData && companyDetailData.quality && companyDetailData.activityHealth) {
    var q = companyDetailData.quality;
    var total = companyDetailData.activityHealth.total || 1;
    // Invert: lower issues = higher score
    var issuePct = ((q.noPerson + q.noCategory + q.unreachable) / (total * 3)) * 100;
    scores.push(100 - issuePct);
  }

  if (scores.length === 0) return null;
  var avg = 0;
  for (var i = 0; i < scores.length; i++) avg += scores[i];
  return Math.round(avg / scores.length);
}

// end of DQ score functions

function renderCompanyDetails(d) {
  var el = document.getElementById('companyDetailContent');
  if (!el) return;
  var h = '';
  h += dateFilterNotice();
  var ah = d.activityHealth;
  var total = ah ? ah.total : 0;

  var subRight = document.getElementById('companySubTabsRight');
  if (subRight) {
    var sr = '';
    sr += '<span class="record-badge">' + fmtNum(total) + ' companies</span>';
    sr += '<div class="detail-filter-wrap">';
    sr += '<div class="detail-filter-btn" onclick="togDetailFilter(\'company\')">';
    sr += '<svg width="14" height="14" viewBox="0 0 16 16" fill="none"><path d="M14.4 3.1A1 1 0 0 0 13.5 2.5h-11a1 1 0 0 0-.74 1.67l4.24 4.52v3.47a1 1 0 0 0 .55.9l2 1a1 1 0 0 0 1.45-.9v-4.47l4.23-4.52a1 1 0 0 0 .17-1.07z" fill="#06423e"/></svg>';
    sr += ' Filters <span class="filter-badge" id="companyFilterBadge" style="display:none">1</span>';
    sr += '</div>';
    sr += '<div class="detail-filter-popover" id="companyFilterPopover">';
    sr += '<label>Category</label>';
    sr += '<select id="companyDetailCatFilter" onchange="onDetailFilterChange(\'company\')">';
    sr += '<option value="">All categories</option>';
    sr += '</select>';
    sr += '<span class="filter-count" id="companyDetailFilterCount"></span>';
    sr += '<span class="filter-reset" id="companyFilterReset" style="display:none" onclick="resetDetailFilter(\'company\')">Reset</span>';
    sr += '</div></div>';
    subRight.innerHTML = sr;
  }

  // 1. ACTIVITY HEALTH — visualize engagement status
  var ah = d.activityHealth;
  if (ah && total > 0) {
    h += '<div class="detail-section">';
    h += '<div class="detail-section-head">' + secHead('Activity Health') + '</div>';
    h += '<div class="stat-row">';
    var ahItems = [
      { label: 'Active (6m)', val: ah.active6m, col: 'var(--so-meadow)' },
      { label: 'Cooling (6-12m)', val: ah.dormant12m, col: 'var(--so-mango)' },
      { label: 'Dormant (>12m)', val: ah.dormantOlder, col: 'var(--so-coral)' },
      { label: 'No Activity', val: ah.noActivity, col: '#999' }
    ];
    for (var i = 0; i < ahItems.length; i++) {
      h += ovCard(ahItems[i].label, ahItems[i].val, total, ahItems[i].col);
    }
    h += '</div>';
    // Stacked bar with inline labels + legend
    h += '<div style="padding:4px 0 8px">';
    var ahParts = [
      { pct: ah.active6m / total * 100, col: 'var(--so-meadow)', label: 'Active (6m)' },
      { pct: ah.dormant12m / total * 100, col: 'var(--so-mango)', label: 'Cooling (6-12m)' },
      { pct: ah.dormantOlder / total * 100, col: 'var(--so-coral)', label: 'Dormant (>12m)' },
      { pct: ah.noActivity / total * 100, col: '#ccc', label: 'No Activity' }
    ];
    h += '<div class="stacked-bar">';
    for (var i = 0; i < ahParts.length; i++) {
      var rp = Math.round(ahParts[i].pct);
      if (rp > 0) {
        var segLabel = rp >= 10 ? '<span style="font-size:.7rem;color:#fff;font-weight:600;padding:0 6px">' + rp + P + '</span>' : '';
        h += '<div class="stacked-segment" style="width:' + ahParts[i].pct + P + ';background:' + ahParts[i].col + ';display:flex;align-items:center;justify-content:center;overflow:hidden">' + segLabel + '</div>';
      }
    }
    h += '</div>';
    // Legend
    h += '<div style="display:flex;gap:16px;margin-top:6px;flex-wrap:wrap">';
    for (var i = 0; i < ahParts.length; i++) {
      var rp = Math.round(ahParts[i].pct * 10) / 10;
      h += '<span style="font-size:.75rem;color:#666;display:flex;align-items:center;gap:4px">';
      h += '<span style="width:8px;height:8px;border-radius:50%;background:' + ahParts[i].col + ';flex-shrink:0"></span>';
      h += ahParts[i].label + ' (' + rp + P + ')';
      h += '</span>';
    }
    h += '</div></div>';
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
      var cW = 960, cH = 220, padL = 40, padR = 10, padT = 15, padB = 40;
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
      h += '<svg viewBox="0 0 ' + cW + ' ' + cH + '" style="width:100%;height:auto;display:block;max-height:240px">';
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
        h += '<polyline points="' + aPts.join(' ') + '" fill="none" stroke="var(--so-meadow-dark1)" stroke-width="2" stroke-dasharray="6,4" stroke-linejoin="round" stroke-linecap="round"/>';
      }
      // Data points + labels
      for (var i = 0; i < visibleData.length; i++) {
        var xy = pts[i].split(',');
        h += '<circle cx="' + xy[0] + '" cy="' + xy[1] + '" r="4" fill="var(--so-green)" stroke="#fff" stroke-width="2"/>';
        if (visibleData[i].count > 0) {
          h += '<text x="' + xy[0] + '" y="' + (parseFloat(xy[1]) - 10).toFixed(1) + '" text-anchor="middle" fill="var(--so-charcoal)" font-size="10" font-weight="600" font-family="DM Sans,sans-serif">' + fmtNum(visibleData[i].count) + '</text>';
        }
        // Active dot
        if (hasActiveData) {
          var axy = aPts[i].split(',');
          h += '<circle cx="' + axy[0] + '" cy="' + axy[1] + '" r="3" fill="var(--so-meadow-dark1)" stroke="#fff" stroke-width="1.5"/>';
        }
        // X labels (skip some for monthly to avoid overlap)
        var showLabel = !useMonthly || (i === 0 || i === visibleData.length - 1 || i % 3 === 0);
        if (showLabel) {
          var displayLabel = useMonthly ? visibleData[i].label.replace('-', '/') : visibleData[i].label;
          h += '<text x="' + xy[0] + '" y="' + (padT + plotH + 22) + '" text-anchor="middle" fill="#999" font-size="10" font-family="DM Sans,sans-serif">' + displayLabel + '</text>';
        }
      }
      h += '</svg>';
      // Legend
      h += '<div style="display:flex;gap:16px;margin-top:4px;padding-left:' + padL + 'px">';
      h += '<span style="font-size:.75rem;color:#666;display:flex;align-items:center;gap:4px"><span style="width:12px;height:3px;background:var(--so-green);border-radius:2px"></span> Registered</span>';
      if (hasActiveData) {
        h += '<span style="font-size:.75rem;color:#666;display:flex;align-items:center;gap:4px"><span style="width:12px;height:0;border-top:2px dashed var(--so-meadow-dark1)"></span> Still active (12m)</span>';
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
    h += '<th class="col-right" onclick="sortT(' + Q + ctid + Q + ',3)">Active (12m) <span class="sort-arrow">' + svgSortN + '</span></th>';
    h += '<th class="col-right" onclick="sortT(' + Q + ctid + Q + ',4)">Open Sale <span class="sort-arrow">' + svgSortN + '</span></th>';
    h += '<th class="col-right" onclick="sortT(' + Q + ctid + Q + ',5)">Engagement <span class="sort-arrow">' + svgSortN + '</span></th>';
    h += '</tr></thead><tbody>';
    for (var i = 0; i < ce.length; i++) {
      var c = ce[i];
      var pctPers = c.total > 0 ? Math.round((c.withPerson / c.total) * 1000) / 10 : 0;
      var pctAct = c.total > 0 ? Math.round((c.withActivity / c.total) * 1000) / 10 : 0;
      var pctSale = c.total > 0 ? Math.round((c.withSale / c.total) * 1000) / 10 : 0;
      var engagement = c.total > 0 ? Math.round((c.withActivity * 0.5 + c.withPerson * 0.3 + c.withSale * 0.2) / c.total * 100) : 0;
      var engCol = engagement >= 50 ? 'var(--so-meadow)' : engagement >= 25 ? 'var(--so-mango)' : 'var(--so-coral)';
      h += '<tr>';
      h += '<td data-sort-value="' + c.name + '">' + c.name + '</td>';
      h += '<td class="col-right" data-sort-value="' + c.total + '">' + fmtNum(c.total) + '</td>';
      h += '<td class="col-right" data-sort-value="' + pctPers + '">' + fmtNum(c.withPerson) + '<span style="color:#999;font-size:.75rem;margin-left:4px">' + pctPers + P + '</span></td>';
      h += '<td class="col-right" data-sort-value="' + pctAct + '">' + fmtNum(c.withActivity) + '<span style="color:#999;font-size:.75rem;margin-left:4px">' + pctAct + P + '</span></td>';
      h += '<td class="col-right" data-sort-value="' + pctSale + '">' + fmtNum(c.withSale) + '<span style="color:#999;font-size:.75rem;margin-left:4px">' + pctSale + P + '</span></td>';
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
