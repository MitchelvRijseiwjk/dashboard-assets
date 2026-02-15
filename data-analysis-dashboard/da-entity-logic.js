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

  // Hide start button, show progress
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

// Full entity load with unified progress bar
function startFullEntity(key) {
  captureEntityFilter(key);
  var ec = entityConfig[key];
  var tabKey = ec ? ec.tabKey : key;

  var progressScreen = document.getElementById(tabKey + 'ProgressScreen');
  var progressBar = document.getElementById(tabKey + 'ProgressBar');
  var progressPercent = document.getElementById(tabKey + 'ProgressPercent');
  var progressStatus = document.getElementById(tabKey + 'ProgressStatus');
  var subTabs = document.getElementById(tabKey + 'SubTabs');
  var resultsContainer = document.getElementById(tabKey + 'Results');

  var headerBtn = document.getElementById(tabKey + 'AnalyzeBtn');
  if (headerBtn) headerBtn.disabled = true;
  if (progressScreen) progressScreen.style.display = '';
  if (progressBar) { progressBar.style.width = '0'; progressBar.classList.add('loading'); }
  if (progressPercent) progressPercent.textContent = '0' + P;

  var hasDetails = (key === 'company');
  var ranges = {
    overview: { start: 0, end: 15 },
    udef: { start: 15, end: hasDetails ? 22 : 25 },
    details: { start: 22, end: 30 },
    extra: { start: hasDetails ? 30 : 25, end: 100 }
  };

  var currentPct = 0;
  var fakeTimer = null;

  function setProgress(pct, status) {
    currentPct = pct;
    if (progressBar) progressBar.style.width = pct + P;
    if (progressPercent) progressPercent.textContent = Math.round(pct) + P;
    if (progressStatus) progressStatus.innerHTML = status;
  }

  function startFakeProgress(targetPct, status) {
    stopFakeProgress();
    var maxPct = targetPct - 1; // Never quite reach target
    setProgress(currentPct, status);
    fakeTimer = setInterval(function() {
      if (currentPct < maxPct) {
        var remaining = maxPct - currentPct;
        var increment = Math.max(0.3, remaining * 0.08);
        currentPct = Math.min(currentPct + increment, maxPct);
        if (progressBar) progressBar.style.width = currentPct + P;
        if (progressPercent) progressPercent.textContent = Math.round(currentPct) + P;
      }
    }, 150);
  }

  function stopFakeProgress() {
    if (fakeTimer) { clearInterval(fakeTimer); fakeTimer = null; }
  }

  // Step 1: Load overview with fake progress
  startFakeProgress(ranges.overview.end, 'Loading overview data...');

  ajax(overviewUrl + String.fromCharCode(38) + 'entity=' + key + getDateFilterParam(), function(d) {
    stopFakeProgress();
    if (d) renderEntityOverview(key, d);
    setProgress(ranges.overview.end, 'Overview loaded');
    loadStep2();
  });

  // Step 2: UDEF or Ticket Fields
  function loadStep2() {
    if (ec && ec.udefId > 0) {
      startFakeProgress(ranges.udef.end, 'Loading extra fields...');
      loadEntityUdefQuiet(ec.udefId, tabKey, ec.udefIdx, function() {
        stopFakeProgress();
        setProgress(ranges.udef.end, 'UDEF fields loaded');
        loadStepDetails();
      });
    } else if (ec && ec.hasTicketFields) {
      startFakeProgress(ranges.udef.end, 'Loading custom ticket fields...');
      loadTicketFieldsQuiet(function() {
        stopFakeProgress();
        setProgress(ranges.udef.end, 'Custom fields loaded');
        loadStepDetails();
      });
    } else {
      setProgress(ranges.udef.end, 'Preparing extra tables...');
      loadStepDetails();
    }
  }

  // Step 3: Extra tables with real per-table progress

  function loadStepDetails() {
    if (hasDetails) {
      startFakeProgress(ranges.details.end, 'Loading detail analysis...');
      loadCompanyDetails(function() {
        stopFakeProgress();
        setProgress(ranges.details.end, 'Details loaded');
        loadStep3();
      });
    } else {
      loadStep3();
    }
  }

  function loadStep3() {
    if (ec) {
      startFakeProgress(ranges.extra.start + 5, 'Loading extra tables list...');
      loadEntityExtraWithProgress(tabKey, ec.extraType, function(cur, total, tableName) {
        stopFakeProgress();
        var extraPct = ranges.extra.start + ((cur / total) * (ranges.extra.end - ranges.extra.start));
        var status = 'Loading: <span class="progress-detail">' + tableName + '</span> (' + cur + '/' + total + ')';
        setProgress(extraPct, status);
      }, function() {
        stopFakeProgress();
        finishLoad();
      });
    } else {
      finishLoad();
    }
  }

  function finishLoad() {
    stopFakeProgress();
    setProgress(100, 'Complete!');
    if (progressBar) progressBar.classList.remove('loading');

    setTimeout(function() {
      if (progressScreen) progressScreen.style.display = 'none';
      if (subTabs) subTabs.style.display = '';
      if (resultsContainer) resultsContainer.style.display = '';
      var expBtn = document.getElementById(tabKey + 'ExportBtn') || document.getElementById('ticketExportBtn');
      if (expBtn) expBtn.style.display = '';
      if (headerBtn) { headerBtn.disabled = false; headerBtn.onclick = function(){ reAnalyze(key); }; }
    }, 500);
  }
}

function loadEntityExtraWithProgress(key, entityType, onProgress, onComplete) {
  if (!entExtra[key]) entExtra[key] = { tables:[], cur:0, data:{} };
  ajax(extraUrl + getDateFilterParam(), function(d) {
    if (d && d.tables && d.tables.length > 0) {
      entExtra[key].tables = d.tables;
      entExtra[key].cur = 0;
      entExtra[key].entityType = entityType;
      loadExtraChainWithProgress(key, onProgress, onComplete);
    } else {
      showEntityExtraQuiet(key);
      if (onComplete) onComplete();
    }
  });
}

function loadExtraChainWithProgress(key, onProgress, onComplete) {
  var es = entExtra[key];
  if (es.cur >= es.tables.length) {
    showEntityExtraQuiet(key);
    if (onComplete) onComplete();
    return;
  }
  var tbl = es.tables[es.cur];
  // Report progress
  if (onProgress) onProgress(es.cur + 1, es.tables.length, tbl.displayName);
  ajax(extraUrl + String.fromCharCode(38) + 'tableId=' + tbl.id + getDateFilterParam(), function(d) {
    es.data[tbl.id] = d;
    es.cur++;
    loadExtraChainWithProgress(key, onProgress, onComplete);
  });
}

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
function startCompanyOverview() { startFullEntity('company'); }

function renderEntityOverview(key, d) {
  // Store data for export
  overviewData[key] = d;

  var cfg = ovLabels[key];
  var o = d.overview;
  var totalKey = cfg.stats[0][0];
  var total = o[totalKey];
  var h = '';

  // Show date filter notice if active
  h += dateFilterNotice();

  // Stat cards
  h += '<div class="detail-section">';
  h += '<div class="detail-section-head">' + secHead(cfg.title) + '</div>';
  h += '<div class="stat-row">';
  for (var i = 0; i < cfg.stats.length; i++) {
    var s = cfg.stats[i];
    if (s[0] === totalKey) {
      h += ovCard(s[1], total, '', '');
    } else {
      h += ovCard(s[1], o[s[0] ], total, '');
    }
  }
  h += '</div></div>';

  // Extra sections (e.g. Activities has a Document Overview section)
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
          h += ovCard(ss[1], o[ss[0] ], '', '');
        } else {
          h += ovCard(ss[1], o[ss[0] ], secTotal, '');
        }
      }
      h += '</div></div>';
    }
  }

  // Completeness (company only)
  if (cfg.completeness && d.completeness) {
    var c = d.completeness;
    h += '<div class="detail-section">';
    h += '<div class="detail-section-head">' + secHead('Data Completeness') + '</div>';
    h += '<div class="stat-row">';
    for (var j = 0; j < cfg.completeness.length; j++) {
      var cm = cfg.completeness[j];
      h += ovCard(cm[1], c[cm[0] ], total, '');
    }
    h += '</div></div>';
  }

  // Distribution tables (generic from array)
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

function fmtNum(n) {
  if (typeof n !== 'number') return n;
  return n.toLocaleString();
}

function fmtDate(d) {
  if (!d || d === '0' || d === '') return '';
  // Handle "YYYY-MM-DD HH:MM:SS" or "YYYY-MM-DD"
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
var companyDetailCatValue = ''; // Store current filter value

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
    // Populate category dropdown after render
    populateDetailCatFilter();
    // Show filter count
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
  // Clear existing options except first
  while (sel.options.length > 1) sel.remove(1);
  // Populate from overview data
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
  // Restore selected value
  if (companyDetailCatValue) sel.value = companyDetailCatValue;
  // Update badge
  var badge = document.getElementById('companyFilterBadge');
  var resetBtn = document.getElementById('companyFilterReset');
  if (badge) badge.style.display = companyDetailCatValue ? '' : 'none';
  if (resetBtn) resetBtn.style.display = companyDetailCatValue ? '' : 'none';
}

function reloadCompanyDetails() {
  var el = document.getElementById('companyDetailContent');
  if (el) el.innerHTML = '<div style="text-align:center;padding:40px;color:#999">Loading...</div>';
  fetchCompanyDetails(null);
}

// Filter popover
function togDetailFilter(key) {
  var pop = document.getElementById(key + 'FilterPopover');
  if (!pop) return;
  var isOpen = pop.classList.contains('show');
  // Close all open popovers first
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

// Close filter popover on outside click
document.addEventListener('click', function(e) {
  if (!e.target.closest('.detail-filter-wrap')) {
    document.querySelectorAll('.detail-filter-popover.show').forEach(function(p) { p.classList.remove('show'); });
  }
});

function renderCompanyDetails(d) {
  var el = document.getElementById('companyDetailContent');
  if (!el) return;
  var h = '';
  h += dateFilterNotice();
  var ah = d.activityHealth;
  var total = ah ? ah.total : 0;

  // Populate sub-tabs-right with filter + badges
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

  // 1. DATA QUALITY ISSUES
  var q = d.quality;
  if (q && total > 0) {
    h += '<div class="detail-section">';
    h += '<div class="detail-section-head">';
    h += secHead('Data Quality Issues');
    h += '</div>';
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

  // 1b. REGISTRATION TREND
  var trend = d.trend;
  if (trend && trend.length > 0) {
    var firstIdx = 0;
    for (var i = 0; i < trend.length; i++) { if (trend[i].count > 0) { firstIdx = i; break; } }
    var visibleTrend = trend.slice(firstIdx);
    var maxCount = 0;
    for (var i = 0; i < visibleTrend.length; i++) { if (visibleTrend[i].count > maxCount) maxCount = visibleTrend[i].count; }
    if (maxCount > 0 && visibleTrend.length > 1) {
      var beforeTotal = d.trendBefore || 0;
      for (var i = 0; i < firstIdx; i++) { beforeTotal += trend[i].count; }
      h += '<div class="detail-section">';
      h += '<div class="detail-section-head">';
      h += secHead('New Registrations Per Year');
      if (beforeTotal > 0) {
        h += '<span class="record-badge">' + fmtNum(beforeTotal) + ' before ' + visibleTrend[0].year + '</span>';
      }
      h += '</div>';
      // Chart dimensions (pixel-based)
      var cW = 960, cH = 200, padL = 40, padR = 10, padT = 15, padB = 40;
      var plotW = cW - padL - padR, plotH = cH - padT - padB;
      var dataInset = 15; // inner margin so dots don't sit on grid edge
      // Nice Y-axis scale
      var niceMax = maxCount;
      var mag = Math.pow(10, Math.floor(Math.log10(maxCount)));
      var options = [1, 1.5, 2, 2.5, 3, 4, 5, 8, 10];
      for (var oi = 0; oi < options.length; oi++) { if (options[oi] * mag >= maxCount) { niceMax = options[oi] * mag; break; } }
      var ySteps = 4;
      var yStep = niceMax / ySteps;
      var step = (plotW - 2 * dataInset) / (visibleTrend.length - 1);
      var pts = [];
      for (var i = 0; i < visibleTrend.length; i++) {
        var px = padL + dataInset + i * step;
        var py = padT + plotH - (visibleTrend[i].count / niceMax) * plotH;
        pts.push(px.toFixed(1) + ',' + py.toFixed(1));
      }
      var areaPath = 'M' + (padL + dataInset) + ',' + (padT + plotH) + ' L' + pts.join(' L') + ' L' + (padL + dataInset + (visibleTrend.length - 1) * step).toFixed(1) + ',' + (padT + plotH) + ' Z';
      h += '<svg viewBox="0 0 ' + cW + ' ' + cH + '" style="width:100%;height:auto;display:block;max-height:220px">';
      // Y-axis grid + labels
      for (var gi = 0; gi <= ySteps; gi++) {
        var gy = padT + plotH - (gi / ySteps) * plotH;
        var yVal = Math.round(gi * yStep);
        h += '<line x1="' + padL + '" y1="' + gy.toFixed(1) + '" x2="' + (cW - padR) + '" y2="' + gy.toFixed(1) + '" stroke="#e0dfdc" stroke-width="1"/>';
        h += '<text x="' + (padL - 8) + '" y="' + (gy + 4).toFixed(1) + '" text-anchor="end" fill="#999" font-size="11" font-family="DM Sans,sans-serif">' + fmtNum(yVal) + '</text>';
      }
      // Area fill
      h += '<path d="' + areaPath + '" fill="rgba(22,91,112,0.06)"/>';
      // Line
      h += '<polyline points="' + pts.join(' ') + '" fill="none" stroke="var(--so-green)" stroke-width="2.5" stroke-linejoin="round" stroke-linecap="round"/>';
      // Dots + X labels
      for (var i = 0; i < visibleTrend.length; i++) {
        var xy = pts[i].split(',');
        h += '<circle cx="' + xy[0] + '" cy="' + xy[1] + '" r="5" fill="var(--so-green)" stroke="#fff" stroke-width="2"/>';
        // Value above dot
        h += '<text x="' + xy[0] + '" y="' + (parseFloat(xy[1]) - 12).toFixed(1) + '" text-anchor="middle" fill="var(--so-charcoal)" font-size="11" font-weight="600" font-family="DM Sans,sans-serif">' + fmtNum(visibleTrend[i].count) + '</text>';
        // Year below axis
        h += '<text x="' + xy[0] + '" y="' + (padT + plotH + 22) + '" text-anchor="middle" fill="#999" font-size="11" font-family="DM Sans,sans-serif">' + visibleTrend[i].year + '</text>';
      }
      h += '</svg>';
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
      h += '<td class="col-right" data-sort-value="' + engagement + '">' + fillBar(engagement, 8, engCol) + '<span style="color:#999;font-size:.75rem;margin-left:6px">' + engagement + P + '</span></td>';
      h += '</tr>';
    }
    h += '</tbody></table></div>';
  }

  // 3. ASSOCIATE BREAKDOWN (grouped by user group if available)
  var assocs = d.associates;
  if (assocs && assocs.length > 0) {
    assocs.sort(function(a, b) { return b.total - a.total; });
    var totalAll = 0;
    for (var i = 0; i < assocs.length; i++) totalAll += assocs[i].total;
    var Q = String.fromCharCode(39);
    var tid = 'tbl-assoc';

    // Check if user group data is available
    var hasGroups = assocs.length > 0 && assocs[0].groupName;

    // Build group map
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
    // Sort groups by total desc
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
      r += '<td class="col-right" data-sort-value="' + compP + '">' + fillBar(compP, 8, '') + '<span style="color:#999;font-size:.75rem;margin-left:6px">' + compP + P + '</span></td>';
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
        h += '<td class="col-right" data-sort-value="' + gCompP + '">' + fillBar(gCompP, 8, '') + '<span style="color:#999;font-size:.75rem;margin-left:6px">' + gCompP + P + '</span></td>';
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
  ajax(extraUrl + getDateFilterParam(), function(d) {
    if (d) { extraTables = d.tables; extraCur = 0; loadExtra(); }
  });
}

function loadExtra() {
  if (extraCur >= extraTables.length) { showExtra(); return; }
  var tbl = extraTables[extraCur];
  var pct = Math.round((extraCur / extraTables.length) * 100);
  document.getElementById('extraBar').style.width = pct + P;
  document.getElementById('extraStatus').textContent = 'Loading: ' + tbl.displayName + ' (' + (extraCur+1) + '/' + extraTables.length + ')';
  ajax(extraUrl + String.fromCharCode(38) + 'tableId=' + tbl.id + getDateFilterParam(), function(d) { extraData[tbl.id] = d; extraCur++; loadExtra(); });
}

function showExtra() {
  document.getElementById('extraBar').style.width = '100' + P;
  document.getElementById('extraStatus').textContent = 'Complete!';
  setTimeout(function() {
    document.getElementById('extraStart').style.display = 'none';
    document.getElementById('extraResults').style.display = 'block';
    document.getElementById('extraExportBtn').style.display = '';
    var btn = document.getElementById('extraAnalyzeBtn');
    if (btn) { btn.disabled = false; btn.onclick = function(){ document.getElementById('extraResults').style.display = 'none'; document.getElementById('extraCards').innerHTML = ''; extraData = {}; startExtra(); }; }
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

// === ENTITY EXTRA TABLES ===
var entExtra = {}; // { key: { tables:[], cur:0, data:{} } }

function startEntityExtra(key, entityType) {
  if (!entExtra[key]) entExtra[key] = { tables:[], cur:0, data:{} };
  var btns = document.querySelectorAll('#' + key + 'ExtraStart .btn-analyze');
  for (var i = 0; i < btns.length; i++) btns[i].style.display = 'none';
  document.getElementById(key + 'ExtraProgress').style.display = 'block';
  document.getElementById(key + 'ExtraStatus').textContent = 'Loading tables...';
  ajax(extraUrl + getDateFilterParam(), function(d) {
    if (d) {
      entExtra[key].tables = d.tables;
      entExtra[key].cur = 0;
      entExtra[key].entityType = entityType;
      loadEntityExtra(key);
    }
  });
}

function loadEntityExtra(key) {
  var es = entExtra[key];
  if (es.cur >= es.tables.length) { showEntityExtra(key); return; }
  var tbl = es.tables[es.cur];
  var pct = Math.round((es.cur / es.tables.length) * 100);
  document.getElementById(key + 'ExtraBar').style.width = pct + P;
  document.getElementById(key + 'ExtraStatus').textContent = 'Loading: ' + tbl.displayName + ' (' + (es.cur+1) + '/' + es.tables.length + ')';
  ajax(extraUrl + String.fromCharCode(38) + 'tableId=' + tbl.id + getDateFilterParam(), function(d) {
    es.data[tbl.id] = d;
    es.cur++;
    loadEntityExtra(key);
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

function loadEntityExtraQuiet(key, entityType, callback) {
  if (!entExtra[key]) entExtra[key] = { tables:[], cur:0, data:{} };
  ajax(extraUrl + getDateFilterParam(), function(d) {
    if (d) {
      entExtra[key].tables = d.tables;
      entExtra[key].cur = 0;
      entExtra[key].entityType = entityType;
      loadEntityExtraChain(key, function() {
        showEntityExtraQuiet(key);
        if (callback) callback();
      });
    } else {
      if (callback) callback();
    }
  });
}
function loadEntityExtraChain(key, callback) {
  var es = entExtra[key];
  if (es.cur >= es.tables.length) { if (callback) callback(); return; }
  var tbl = es.tables[es.cur];
  ajax(extraUrl + String.fromCharCode(38) + 'tableId=' + tbl.id + getDateFilterParam(), function(d) {
    es.data[tbl.id] = d;
    es.cur++;
    loadEntityExtraChain(key, callback);
  });
}
function showEntityExtraQuiet(key) {
  var es = entExtra[key];
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
  if (c) { for (var i = 0; i < related.length; i++) { var d = es.data[related[i].id]; if (d) c.innerHTML += renderExtra(d, baseIdx + i); } }
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
