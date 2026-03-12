// === EXPORT ===
function exportFullEntity(key, entityName, udefEntityId) {
  var wb = XLSX.utils.book_new();
  var cfg = ovLabels[key];
  var ec = entityConfig[key];

  // === SHEET 1: Overview ===
  var ovRows = [];
  var od = overviewData[key];
  if (od && cfg) {
    var o = od.overview;
    var totalKey = cfg.stats[0][0];
    var total = o[totalKey];

    // Main stats
    ovRows.push(['OVERVIEW STATISTICS', '', '', '']);
    ovRows.push(['Metric', 'Value', 'Total', 'Percentage']);
    for (var i = 0; i < cfg.stats.length; i++) {
      var s = cfg.stats[i];
      var val = o[s[0]];
      var pct = (s[0] === totalKey) ? '' : (total > 0 ? (val / total * 100).toFixed(1) + P : '0' + P);
      ovRows.push([s[1], val, (s[0] === totalKey) ? '' : total, pct]);
    }

    // Extra sections (e.g. Documents for Activities)
    if (cfg.sections) {
      for (var si = 0; si < cfg.sections.length; si++) {
        var sec = cfg.sections[si];
        var secTotal = o[sec.totalKey];
        ovRows.push(['', '', '', '']);
        ovRows.push([sec.title.toUpperCase(), '', '', '']);
        ovRows.push(['Metric', 'Value', 'Total', 'Percentage']);
        for (var sj = 0; sj < sec.stats.length; sj++) {
          var ss = sec.stats[sj];
          var sval = o[ss[0]];
          var spct = (ss[0] === sec.totalKey || ss[2]) ? '' : (secTotal > 0 ? (sval / secTotal * 100).toFixed(1) + P : '0' + P);
          ovRows.push([ss[1], sval, (ss[0] === sec.totalKey || ss[2]) ? '' : secTotal, spct]);
        }
      }
    }

    // Completeness (company only)
    if (cfg.completeness && od.completeness) {
      var c = od.completeness;
      ovRows.push(['', '', '', '']);
      ovRows.push(['DATA COMPLETENESS', '', '', '']);
      ovRows.push(['Field', 'Filled', 'Total', 'Percentage']);
      for (var j = 0; j < cfg.completeness.length; j++) {
        var cm = cfg.completeness[j];
        var cval = c[cm[0]];
        var cpct = total > 0 ? (cval / total * 100).toFixed(1) + P : '0' + P;
        ovRows.push([cm[1], cval, total, cpct]);
      }
    }

    // Distributions
    if (od.distributions && od.distributions.length > 0) {
      for (var di = 0; di < od.distributions.length; di++) {
        var dist = od.distributions[di];
        var distTotal = dist.total || total;
        ovRows.push(['', '', '', '']);
        ovRows.push([dist.title.toUpperCase(), '', '', '']);
        ovRows.push(['Value', 'Count', 'Total', 'Percentage']);
        if (dist.items) {
          dist.items.sort(function(a, b) { return b.count - a.count; });
          for (var dj = 0; dj < dist.items.length; dj++) {
            var it = dist.items[dj];
            var dpct = distTotal > 0 ? (it.count / distTotal * 100).toFixed(1) + P : '0' + P;
            ovRows.push([it.name, it.count, distTotal, dpct]);
          }
        }
      }
    }
  }
  if (ovRows.length > 0) {
    var wsOv = XLSX.utils.aoa_to_sheet(ovRows);
    wsOv['!cols'] = [{wch:30},{wch:12},{wch:12},{wch:12}];
    XLSX.utils.book_append_sheet(wb, wsOv, 'Overview');
  }

  // === SHEET 2: UDEF Fields ===
  if (udefEntityId && udefData[udefEntityId]) {
    var ud = udefData[udefEntityId];
    var udefRows = [];
    udefRows.push(['Field Label', 'Field Type', 'Dropdown Item', 'Filled', 'Total', 'Fill Rate']);
    var sf = ud.fields.slice().sort(function(a,b) { return b.percent - a.percent; });
    for (var uj = 0; uj < sf.length; uj++) {
      var uf = sf[uj];
      udefRows.push([uf.label, uf.type, '', uf.filled, ud.total, uf.percent/100]);
      if (uf.items && uf.items.length > 0) {
        var tu = 0;
        for (var uk = 0; uk < uf.items.length; uk++) tu += uf.items[uk].c;
        for (var ul = 0; ul < uf.items.length; ul++) {
          var uitm = uf.items[ul];
          udefRows.push([uf.label, 'Dropdown Item', uitm.n, uitm.c, tu, uitm.p/100]);
        }
      }
    }
    var wsUdef = XLSX.utils.aoa_to_sheet(udefRows);
    wsUdef['!cols'] = [{wch:25},{wch:15},{wch:25},{wch:10},{wch:8},{wch:12}];
    XLSX.utils.book_append_sheet(wb, wsUdef, 'UDEF Fields');
  }

  // === SHEET 3: Extra Tables ===
  if (ec && entExtra[key] && entExtra[key].tables) {
    var es = entExtra[key];
    var extRows = [];
    extRows.push(['Table', 'Field Name', 'Display Name', 'Field Type', 'Category', 'Filled', 'Total Rows', 'Fill Rate', 'Entity Type', 'Unique', 'Avg/Entity', 'Coverage']);

    for (var ti = 0; ti < es.tables.length; ti++) {
      var tblInfo = es.tables[ti];
      var tblData = es.data[tblInfo.id];
      if (!tblData) continue;

      var hasRelation = false;
      if (tblData.relationFields) {
        for (var ri = 0; ri < tblData.relationFields.length; ri++) {
          if (tblData.relationFields[ri].entityType === ec.extraType) {
            hasRelation = true;
            break;
          }
        }
      }
      if (!hasRelation) continue;

      if (tblData.relationFields) {
        for (var rj = 0; rj < tblData.relationFields.length; rj++) {
          var rf = tblData.relationFields[rj];
          extRows.push([tblData.displayName, rf.fieldName, rf.displayName, 'Relation', 'Relation Field', rf.filledCount, tblData.totalRows, rf.fillPercent / 100, rf.entityType, rf.uniqueCount, rf.avgPerEntity, rf.coveragePercent / 100]);
        }
      }
      if (tblData.otherFields) {
        var sortedFields = tblData.otherFields.slice().sort(function(a,b) { return b.fillPercent - a.fillPercent; });
        for (var fj = 0; fj < sortedFields.length; fj++) {
          var of2 = sortedFields[fj];
          extRows.push([tblData.displayName, of2.fieldName, of2.displayName, of2.fieldType, 'Other Field', of2.filledCount, tblData.totalRows, of2.fillPercent / 100, '', '', '', '']);
        }
      }
    }

    if (extRows.length > 1) {
      var wsExt = XLSX.utils.aoa_to_sheet(extRows);
      wsExt['!cols'] = [{wch:25},{wch:20},{wch:25},{wch:12},{wch:14},{wch:10},{wch:10},{wch:10},{wch:12},{wch:10},{wch:10},{wch:10}];
      XLSX.utils.book_append_sheet(wb, wsExt, 'Extra Tables');
    }
  }

  // === SHEET 4: Details (company only) ===
  if (key === 'company' && companyDetailData) {
    var dd = companyDetailData;
    var detRows = [];

    // Activity Health
    if (dd.activityHealth) {
      var ah = dd.activityHealth;
      detRows.push(['ACTIVITY HEALTH']);
      detRows.push(['Metric', 'Count', '% of Total']);
      detRows.push(['Active (<6 months)', ah.active6m, ah.total > 0 ? ah.active6m / ah.total : 0]);
      detRows.push(['Cooling (6-12 months)', ah.dormant12m, ah.total > 0 ? ah.dormant12m / ah.total : 0]);
      detRows.push(['Dormant (>12 months)', ah.dormantOlder, ah.total > 0 ? ah.dormantOlder / ah.total : 0]);
      detRows.push(['No activity ever', ah.noActivity, ah.total > 0 ? ah.noActivity / ah.total : 0]);
      detRows.push([]);
    }

    // Churn Risk Matrix
    if (dd.churnRisk && dd.activityHealth) {
      var cr = dd.churnRisk;
      var aht = dd.activityHealth;
      var healthy = cr.activeWithSale || 0;
      var atRisk = (cr.withOpenSale || 0) - healthy;
      var noPipe = (aht.active6m || 0) - healthy;
      var churning = aht.total - (aht.active6m || 0) - atRisk;
      detRows.push(['CHURN RISK MATRIX']);
      detRows.push(['Quadrant', 'Count', '% of Total', 'Description']);
      detRows.push(['Healthy', healthy, aht.total > 0 ? healthy / aht.total : 0, 'Active + open sale']);
      detRows.push(['No Pipeline', noPipe, aht.total > 0 ? noPipe / aht.total : 0, 'Active but no open sale']);
      detRows.push(['At Risk', atRisk, aht.total > 0 ? atRisk / aht.total : 0, 'Open sale but going quiet']);
      detRows.push(['Churning', churning, aht.total > 0 ? churning / aht.total : 0, 'No activity + no sale']);
      detRows.push([]);
      detRows.push(['With open sale', cr.withOpenSale]);
      detRows.push(['With won sale', cr.withWonSale]);
      detRows.push([]);
    }

    // Category Effectiveness
    if (dd.categoryEffectiveness && dd.categoryEffectiveness.length > 0) {
      detRows.push(['CATEGORY EFFECTIVENESS']);
      detRows.push(['Category', 'Companies', 'With Person', '% Person', 'Active (12m)', '% Active', 'Open Sale', '% Sale', 'Engagement']);
      var ceExp = dd.categoryEffectiveness.slice().sort(function(a,b) { return b.total - a.total; });
      for (var ci = 0; ci < ceExp.length; ci++) {
        var cc = ceExp[ci];
        var eng = cc.total > 0 ? Math.round((cc.withActivity * 0.5 + cc.withPerson * 0.3 + cc.withSale * 0.2) / cc.total * 100) : 0;
        detRows.push([cc.name, cc.total, cc.withPerson, cc.total > 0 ? cc.withPerson/cc.total : 0, cc.withActivity, cc.total > 0 ? cc.withActivity/cc.total : 0, cc.withSale, cc.total > 0 ? cc.withSale/cc.total : 0, eng / 100]);
      }
      detRows.push([]);
    }

    // Data Quality
    if (dd.quality) {
      var q = dd.quality;
      detRows.push(['DATA QUALITY']);
      detRows.push(['Issue', 'Count', '% of Total']);
      detRows.push(['No contact person', q.noPerson, q.total > 0 ? q.noPerson / q.total : 0]);
      detRows.push(['No category', q.noCategory, q.total > 0 ? q.noCategory / q.total : 0]);
      detRows.push(['No business type', q.noBusiness, q.total > 0 ? q.noBusiness / q.total : 0]);
      detRows.push(['No org. number', q.noOrgNr, q.total > 0 ? q.noOrgNr / q.total : 0]);
      detRows.push(['Unreachable', q.unreachable, q.total > 0 ? q.unreachable / q.total : 0]);
      detRows.push([]);
    }

    // Associate Breakdown
    if (dd.associates && dd.associates.length > 0) {
      detRows.push(['ASSOCIATE BREAKDOWN']);
      detRows.push(['User Group', 'Associate', 'Companies', 'With Persons', '% Persons', 'With Activities', '% Activities', 'With Email', '% Email', 'Stale', '% Stale']);
      for (var ai = 0; ai < dd.associates.length; ai++) {
        var aa = dd.associates[ai];
        var stale = aa.stale || 0;
        detRows.push([aa.groupName || '', aa.name, aa.total, aa.withPersons, aa.total > 0 ? aa.withPersons/aa.total : 0, aa.withActivities, aa.total > 0 ? aa.withActivities/aa.total : 0, aa.withEmail, aa.total > 0 ? aa.withEmail/aa.total : 0, stale, aa.total > 0 ? stale/aa.total : 0]);
      }
      detRows.push([]);
    }

    // New Registrations Per Year
    if (dd.trend && dd.trend.length > 0) {
      detRows.push(['NEW REGISTRATIONS PER YEAR']);
      detRows.push(['Year', 'New Companies']);
      if (dd.trendBefore > 0) detRows.push(['Before ' + dd.trend[0].year, dd.trendBefore]);
      for (var ti2 = 0; ti2 < dd.trend.length; ti2++) {
        detRows.push([dd.trend[ti2].year, dd.trend[ti2].count]);
      }
    }

    if (detRows.length > 1) {
      var wsDet = XLSX.utils.aoa_to_sheet(detRows);
      wsDet['!cols'] = [{wch:25},{wch:14},{wch:14},{wch:12},{wch:16},{wch:12},{wch:14},{wch:12}];
      // Format percentage columns
      XLSX.utils.book_append_sheet(wb, wsDet, 'Details');
    }
  }

  var fileName = entityName + '_Analysis_' + new Date().toISOString().split('T')[0] + '.xlsx';
  XLSX.writeFile(wb, fileName);
}

function exportUdef() { exportFullEntity('company', 'Company', 7); }
function exportContactUdef() { exportFullEntity('contact', 'Contact', 8); }
function exportSaleUdef() { exportFullEntity('sale', 'Sale', 10); }
function exportProjectUdef() { exportFullEntity('project', 'Project', 9); }

function exportTicket() {
  var wb = XLSX.utils.book_new();
  var cfg = ovLabels['requests'];
  var od = overviewData['requests'];
  var ovRows = [];
  if (od && cfg) {
    var o = od.overview;
    var totalKey = cfg.stats[0][0];
    var total = o[totalKey];
    ovRows.push(['OVERVIEW STATISTICS', '', '', '']);
    ovRows.push(['Metric', 'Value', 'Total', 'Percentage']);
    for (var i = 0; i < cfg.stats.length; i++) {
      var s = cfg.stats[i];
      var val = o[s[0]];
      var pct = (s[0] === totalKey) ? '' : (total > 0 ? (val / total * 100).toFixed(1) + P : '0' + P);
      ovRows.push([s[1], val, (s[0] === totalKey) ? '' : total, pct]);
    }
    if (od.distributions && od.distributions.length > 0) {
      for (var di = 0; di < od.distributions.length; di++) {
        var dist = od.distributions[di];
        var distTotal = dist.total || total;
        ovRows.push(['', '', '', '']);
        ovRows.push([dist.title.toUpperCase(), '', '', '']);
        ovRows.push(['Value', 'Count', 'Total', 'Percentage']);
        if (dist.items) {
          for (var dj = 0; dj < dist.items.length; dj++) {
            var it = dist.items[dj];
            var dpct = distTotal > 0 ? (it.count / distTotal * 100).toFixed(1) + P : '0' + P;
            ovRows.push([it.name, it.count, distTotal, dpct]);
          }
        }
      }
    }
  }
  if (ovRows.length > 0) {
    var wsOv = XLSX.utils.aoa_to_sheet(ovRows);
    wsOv['!cols'] = [{wch:30},{wch:12},{wch:12},{wch:12}];
    XLSX.utils.book_append_sheet(wb, wsOv, 'Overview');
  }
  if (ticketData && ticketData.fields) {
    var tfRows = [];
    tfRows.push(['Field Label', 'Field Name', 'Field Type', 'Is Relation', 'Filled', 'Total Tickets', 'Fill Rate', 'Entity Type', 'Unique', 'Avg/Entity', 'Total Entities', 'Coverage']);
    var sf = ticketData.fields.slice().sort(function(a,b) { return b.fillPercent - a.fillPercent; });
    for (var ti = 0; ti < sf.length; ti++) {
      var tf = sf[ti];
      tfRows.push([tf.displayName, tf.fieldName, tf.fieldType, tf.isRelation ? 'Yes' : 'No', tf.filledCount, ticketData.totalTickets, tf.fillPercent/100, tf.entityType || '', tf.uniqueCount || '', tf.avgPerEntity || '', tf.totalEntityCount || '', tf.isRelation ? tf.coveragePercent / 100 : '']);
    }
    var wsTf = XLSX.utils.aoa_to_sheet(tfRows);
    wsTf['!cols'] = [{wch:25},{wch:20},{wch:12},{wch:10},{wch:10},{wch:12},{wch:10},{wch:12},{wch:10},{wch:12},{wch:12},{wch:10}];
    XLSX.utils.book_append_sheet(wb, wsTf, 'Custom Fields');
  }
  var ec = entityConfig['requests'];
  if (ec && entExtra['requests'] && entExtra['requests'].tables) {
    var es = entExtra['requests'];
    var extRows = [];
    extRows.push(['Table', 'Field Name', 'Display Name', 'Field Type', 'Category', 'Filled', 'Total Rows', 'Fill Rate', 'Entity Type', 'Unique', 'Avg/Entity', 'Coverage']);
    for (var ei = 0; ei < es.tables.length; ei++) {
      var tblInfo = es.tables[ei];
      var tblData = es.data[tblInfo.id];
      if (!tblData) continue;
      var hasRelation = false;
      if (tblData.relationFields) {
        for (var ri = 0; ri < tblData.relationFields.length; ri++) {
          if (tblData.relationFields[ri].entityType === ec.extraType) { hasRelation = true; break; }
        }
      }
      if (!hasRelation) continue;
      if (tblData.relationFields) {
        for (var rj = 0; rj < tblData.relationFields.length; rj++) {
          var rf = tblData.relationFields[rj];
          extRows.push([tblData.displayName, rf.fieldName, rf.displayName, 'Relation', 'Relation Field', rf.filledCount, tblData.totalRows, rf.fillPercent / 100, rf.entityType, rf.uniqueCount, rf.avgPerEntity, rf.coveragePercent / 100]);
        }
      }
      if (tblData.otherFields) {
        for (var fj = 0; fj < tblData.otherFields.length; fj++) {
          var of2 = tblData.otherFields[fj];
          extRows.push([tblData.displayName, of2.fieldName, of2.displayName, of2.fieldType, 'Other Field', of2.filledCount, tblData.totalRows, of2.fillPercent / 100, '', '', '', '']);
        }
      }
    }
    if (extRows.length > 1) {
      var wsExt = XLSX.utils.aoa_to_sheet(extRows);
      wsExt['!cols'] = [{wch:25},{wch:20},{wch:25},{wch:12},{wch:14},{wch:10},{wch:10},{wch:10},{wch:12},{wch:10},{wch:10},{wch:10}];
      XLSX.utils.book_append_sheet(wb, wsExt, 'Extra Tables');
    }
  }
  XLSX.writeFile(wb, 'Requests_Analysis_' + new Date().toISOString().split('T')[0] + '.xlsx');
}

function exportExtra() {
  var wb = XLSX.utils.book_new();
  var ov = [];
  ov.push(['Table Name','Display Name','Total Records','Relation Fields','Other Fields','Has Relations']);
  for (var i = 0; i < extraTables.length; i++) {
    var d = extraData[extraTables[i].id];
    if (!d) continue;
    var hr = 'No';
    if (d.relationFields && d.relationFields.length > 0) hr = 'Yes';
    ov.push([d.tableName, d.displayName, d.totalRows, d.relationFields ? d.relationFields.length : 0, d.otherFields ? d.otherFields.length : 0, hr]);
  }
  var ws1 = XLSX.utils.aoa_to_sheet(ov);
  ws1['!cols'] = [{wch:25},{wch:30},{wch:12},{wch:15},{wch:12},{wch:12}];
  XLSX.utils.book_append_sheet(wb, ws1, 'Overview');

  var fr = [];
  fr.push(['Table','Field Name','Display Name','Field Type','Category','Filled','Total','Fill Rate','Entity Type','Unique','Avg/Entity','Coverage']);
  for (var i = 0; i < extraTables.length; i++) {
    var d = extraData[extraTables[i].id];
    if (!d) continue;
    if (d.relationFields) {
      for (var j = 0; j < d.relationFields.length; j++) {
        var rf = d.relationFields[j];
        fr.push([d.displayName, rf.fieldName, rf.displayName, 'Relation', 'Relation Field', rf.filledCount, d.totalRows, rf.fillPercent/100, rf.entityType, rf.uniqueCount, rf.avgPerEntity, rf.coveragePercent/100]);
      }
    }
    if (d.otherFields) {
      var sf = d.otherFields.slice().sort(function(a,b) { return b.fillPercent - a.fillPercent; });
      for (var j = 0; j < sf.length; j++) {
        var of2 = sf[j];
        fr.push([d.displayName, of2.fieldName, of2.displayName, of2.fieldType, 'Other Field', of2.filledCount, d.totalRows, of2.fillPercent/100, '', '', '', '']);
      }
    }
  }
  var ws2 = XLSX.utils.aoa_to_sheet(fr);
  ws2['!cols'] = [{wch:25},{wch:25},{wch:25},{wch:12},{wch:15},{wch:10},{wch:8},{wch:10},{wch:12},{wch:10},{wch:12},{wch:10}];
  XLSX.utils.book_append_sheet(wb, ws2, 'All Fields');
  XLSX.writeFile(wb, 'ExtraTables_Analysis_' + new Date().toISOString().split('T')[0] + '.xlsx');
}

function exportSimpleEntity(key, entityName) {
  var wb = XLSX.utils.book_new();
  var cfg = ovLabels[key];
  var od = overviewData[key];
  var ovRows = [];

  if (od && cfg) {
    var o = od.overview;
    var totalKey = cfg.stats[0][0];
    var total = o[totalKey];

    ovRows.push(['OVERVIEW STATISTICS', '', '', '']);
    ovRows.push(['Metric', 'Value', 'Total', 'Percentage']);
    for (var i = 0; i < cfg.stats.length; i++) {
      var s = cfg.stats[i];
      var val = o[s[0]];
      var pct = (s[0] === totalKey) ? '' : (total > 0 ? (val / total * 100).toFixed(1) + P : '0' + P);
      ovRows.push([s[1], val, (s[0] === totalKey) ? '' : total, pct]);
    }

    // Extra sections
    if (cfg.sections) {
      for (var si = 0; si < cfg.sections.length; si++) {
        var sec = cfg.sections[si];
        var secTotal = o[sec.totalKey];
        ovRows.push(['', '', '', '']);
        ovRows.push([sec.title.toUpperCase(), '', '', '']);
        ovRows.push(['Metric', 'Value', 'Total', 'Percentage']);
        for (var sj = 0; sj < sec.stats.length; sj++) {
          var ss = sec.stats[sj];
          var sval = o[ss[0]];
          var spct = (ss[0] === sec.totalKey || ss[2]) ? '' : (secTotal > 0 ? (sval / secTotal * 100).toFixed(1) + P : '0' + P);
          ovRows.push([ss[1], sval, (ss[0] === sec.totalKey || ss[2]) ? '' : secTotal, spct]);
        }
      }
    }

    // Distributions
    if (od.distributions && od.distributions.length > 0) {
      for (var di = 0; di < od.distributions.length; di++) {
        var dist = od.distributions[di];
        var distTotal = dist.total || total;
        ovRows.push(['', '', '', '']);
        ovRows.push([dist.title.toUpperCase(), '', '', '']);
        ovRows.push(['Value', 'Count', 'Total', 'Percentage']);
        if (dist.items) {
          for (var dj = 0; dj < dist.items.length; dj++) {
            var it = dist.items[dj];
            var dpct = distTotal > 0 ? (it.count / distTotal * 100).toFixed(1) + P : '0' + P;
            ovRows.push([it.name, it.count, distTotal, dpct]);
          }
        }
      }
    }
  }

  if (ovRows.length > 0) {
    var wsOv = XLSX.utils.aoa_to_sheet(ovRows);
    wsOv['!cols'] = [{wch:30},{wch:12},{wch:12},{wch:12}];
    XLSX.utils.book_append_sheet(wb, wsOv, 'Overview');
  }

  XLSX.writeFile(wb, entityName + '_Analysis_' + new Date().toISOString().split('T')[0] + '.xlsx');
}

function exportActivities() { exportSimpleEntity('activities', 'Activities'); }
function exportSelection() { exportSimpleEntity('selection', 'Selection'); }
function exportMarketing() { exportSimpleEntity('marketing', 'Marketing'); }

// Safe roundedRect — falls back to plain rect if roundedRect fails
function pdfRRect(doc, x, y, w, h, rx, ry, style) {
  if (isNaN(x) || isNaN(y) || isNaN(w) || isNaN(h) || w <= 0 || h <= 0) return;
  try { doc.roundedRect(x, y, w, h, rx || 0, ry || 0, style || 'F'); }
  catch(e) {
    try { doc.rect(x, y, w, h, style || 'F'); } catch(e2) { /* skip */ }
  }
}
// =============================================================================
// PDF HEALTH REPORT — Phase 1-3 (Cover + TOC + Executive Summary)
// =============================================================================

var HR_COVER_B64 = 'data:image/jpeg;base64,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';
var HR_LOGO_B64 = 'data:image/png;base64,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';

// --- Constants ---
var HR_PAGE_W = 210, HR_PAGE_H = 297;
var HR_ML = 22, HR_MR = 22, HR_MT = 20, HR_MB = 18;
var HR_CW = HR_PAGE_W - HR_ML - HR_MR; // 166mm content width

var HR_COLORS = {
  soGreen: [6,66,62], soGreenLight: [15,92,87], white: [255,255,255],
  charcoal: [28,28,30], muted: [137,137,145], gray: [96,94,92],
  border: [221,217,212], dune: [242,239,234], duneDark: [232,227,218],
  good: [46,125,50], ok: [249,168,37], warn: [239,108,0], bad: [198,40,40],
  tblBorder: [240,238,234], blue: [21,101,192]
};

function hrScoreColor(pct) {
  if (pct >= 70) return HR_COLORS.good;
  if (pct >= 50) return HR_COLORS.ok;
  if (pct >= 30) return HR_COLORS.warn;
  return HR_COLORS.bad;
}

function hrScoreHex(pct) {
  var c = hrScoreColor(pct);
  return '#' + c.map(function(v){ return ('0'+v.toString(16)).slice(-2); }).join('');
}

function hrScoreLabel(pct) {
  if (pct >= 70) return 'Good';
  if (pct >= 50) return 'Moderate';
  if (pct >= 30) return 'Needs Attention';
  return 'Critical';
}

// --- Shared drawing helpers ---

function hrDrawPageHeader(doc) {
  doc.addImage(HR_LOGO_B64, 'PNG', HR_ML, 5, 40, 5.5);
  doc.setFontSize(8); doc.setFont('helvetica','bold');
  doc.setTextColor.apply(doc, HR_COLORS.muted);
  var rptDate = new Date().toLocaleDateString('en-GB', {year:'numeric',month:'long'});
  doc.text('CRM Health Report \u00b7 ' + rptDate, HR_PAGE_W - HR_MR, 8.5, {align:'right'});
  doc.setDrawColor.apply(doc, HR_COLORS.border);
  doc.setLineWidth(0.3);
  doc.line(HR_ML, 12, HR_PAGE_W - HR_MR, 12);
}

function hrDrawPageFooter(doc, pageNum, totalPages) {
  // Green bar
  doc.setFillColor.apply(doc, HR_COLORS.soGreen);
  doc.rect(0, HR_PAGE_H - 4, HR_PAGE_W, 4, 'F');
  // Footer text
  doc.setFontSize(7.5); doc.setFont('helvetica','normal');
  doc.setTextColor(187,187,187);
  doc.text('SuperOffice Data Analysis \u00b7 CRM Health Report', HR_ML, HR_PAGE_H - 6);
  // Page number
  doc.setFontSize(8); doc.setTextColor.apply(doc, HR_COLORS.muted);
  doc.text(pageNum + ' / ' + totalPages, HR_PAGE_W - HR_MR, HR_PAGE_H - 10, {align:'right'});
}

function hrDrawEntityHeader(doc, title, subtitle, scoreText) {
  var y = 13, h = 22;
  doc.setFillColor.apply(doc, HR_COLORS.soGreen);
  doc.rect(0, y, HR_PAGE_W, h, 'F');
  doc.setFontSize(20); doc.setFont('helvetica','bold');
  doc.setTextColor.apply(doc, HR_COLORS.white);
  doc.text(title, HR_ML, y + 9);
  doc.setFontSize(10); doc.setFont('helvetica','normal');
  doc.setTextColor(255,255,255);
  doc.text(subtitle, HR_ML, y + 14);
  // Score badge
  doc.setFillColor(26,89,85);
  pdfRRect(doc, HR_PAGE_W - HR_MR - 32, y + 4, 32, 14, 2, 2, 'F');
  doc.setFontSize(22); doc.setFont('helvetica','bold');
  doc.setTextColor.apply(doc, HR_COLORS.white);
  doc.text(scoreText, HR_PAGE_W - HR_MR - 16, y + 13, {align:'center'});
  return y + h + 3; // content start y
}

function hrSectionTitle(doc, y, title, badge) {
  doc.setFontSize(12); doc.setFont('helvetica','bold');
  doc.setTextColor.apply(doc, HR_COLORS.soGreen);
  doc.text(title, HR_ML, y);
  if (badge) {
    var tw = doc.getTextWidth(title);
    doc.setFontSize(8); doc.setFont('helvetica','normal');
    doc.setTextColor.apply(doc, HR_COLORS.charcoal);
    doc.setFillColor.apply(doc, HR_COLORS.dune);
    var bw = doc.getTextWidth(badge) + 6;
    pdfRRect(doc, HR_ML + tw + 4, y - 3, bw, 4.5, 1, 1, 'F');
    doc.text(badge, HR_ML + tw + 7, y);
  }
  doc.setDrawColor.apply(doc, HR_COLORS.border);
  doc.setLineWidth(0.4);
  doc.line(HR_ML, y + 1.5, HR_PAGE_W - HR_MR, y + 1.5);
  return y + 5;
}

// --- Donut ring via canvas ---
function hrDrawDonut(doc, cx, cy, radius, pct, label, subLabel) {
  var canvas = document.createElement('canvas');
  var size = 200; // px
  canvas.width = size; canvas.height = size;
  var ctx = canvas.getContext('2d');
  var ctr = size / 2, r = size * 0.4, lw = size * 0.08;

  // Background ring
  ctx.beginPath();
  ctx.arc(ctr, ctr, r, 0, Math.PI * 2);
  ctx.strokeStyle = '#e8e5e0';
  ctx.lineWidth = lw;
  ctx.stroke();

  // Score arc
  var startAngle = -Math.PI / 2;
  var endAngle = startAngle + (pct / 100) * Math.PI * 2;
  ctx.beginPath();
  ctx.arc(ctr, ctr, r, startAngle, endAngle);
  ctx.strokeStyle = hrScoreHex(pct);
  ctx.lineWidth = lw;
  ctx.lineCap = 'round';
  ctx.stroke();

  // Score text
  ctx.fillStyle = hrScoreHex(pct);
  ctx.font = 'bold ' + Math.round(size * 0.22) + 'px Helvetica, Arial, sans-serif';
  ctx.textAlign = 'center';
  ctx.textBaseline = 'middle';
  ctx.fillText(Math.round(pct) + '%', ctr, ctr + 2);

  var imgData = canvas.toDataURL('image/png');
  var imgSize = radius * 2;
  doc.addImage(imgData, 'PNG', cx - radius, cy - radius, imgSize, imgSize);

  // Label below ring
  doc.setFontSize(11); doc.setFont('helvetica','bold');
  doc.setTextColor.apply(doc, HR_COLORS.charcoal);
  doc.text(label, cx, cy + radius + 4, {align:'center'});
  doc.setFontSize(8.5); doc.setFont('helvetica','normal');
  doc.setTextColor.apply(doc, HR_COLORS.muted);
  doc.text(subLabel, cx, cy + radius + 7.5, {align:'center'});
}

// --- Score card row (3 cards for DQ/DI/Adoption) ---
function hrDrawScoreCards(doc, y, scores) {
  var cardW = (HR_CW - 6) / 3;
  var items = [
    {label:'DATA QUALITY', val: scores.dq},
    {label:'DATA INTEGRITY', val: scores.integrity},
    {label:'ADOPTION', val: scores.adoption}
  ];
  for (var i = 0; i < 3; i++) {
    var x = HR_ML + i * (cardW + 3);
    doc.setFillColor.apply(doc, HR_COLORS.dune);
    pdfRRect(doc, x, y, cardW, 16, 2, 2, 'F');
    doc.setFontSize(7); doc.setFont('helvetica','bold');
    doc.setTextColor.apply(doc, HR_COLORS.muted);
    doc.text(items[i].label, x + 3, y + 4);
    var pct = Math.round(items[i].val);
    doc.setFontSize(16); doc.setFont('helvetica','bold');
    var col = hrScoreColor(pct);
    doc.setTextColor.apply(doc, col);
    doc.text(pct + '%', x + 3, y + 11);
    // Mini progress bar
    doc.setFillColor(224,223,220);
    pdfRRect(doc, x + 3, y + 13, cardW - 6, 1.5, 0.5, 0.5, 'F');
    doc.setFillColor.apply(doc, col);
    pdfRRect(doc, x + 3, y + 13, Math.max(1, (cardW - 6) * pct / 100), 1.5, 0.5, 0.5, 'F');
  }
  return y + 20;
}

// --- Metric cards grid (4 cards) ---
function hrDrawMetricCards(doc, y, metrics) {
  var cardW = (HR_CW - 9) / 4;
  for (var i = 0; i < metrics.length && i < 4; i++) {
    var m = metrics[i];
    var x = HR_ML + i * (cardW + 3);
    doc.setFillColor.apply(doc, HR_COLORS.dune);
    pdfRRect(doc, x, y, cardW, 18, 2, 2, 'F');
    var midX = x + cardW / 2;
    // Value
    doc.setFontSize(16); doc.setFont('helvetica','bold');
    doc.setTextColor.apply(doc, HR_COLORS.charcoal);
    doc.text(fmtNum(m.value), midX, y + 6.5, {align:'center'});
    // / total
    if (m.total) {
      doc.setFontSize(8); doc.setFont('helvetica','normal');
      doc.setTextColor.apply(doc, HR_COLORS.muted);
      doc.text('/ ' + fmtNum(m.total), midX, y + 9.5, {align:'center'});
    }
    // Percentage
    var pct = m.total > 0 ? Math.round(m.value / m.total * 1000) / 10 : 0;
    doc.setFontSize(9); doc.setFont('helvetica','bold');
    doc.setTextColor.apply(doc, hrScoreColor(pct));
    doc.text(pct.toFixed(1) + '%', midX, y + 13, {align:'center'});
    // Label
    doc.setFontSize(7.5); doc.setFont('helvetica','normal');
    doc.setTextColor.apply(doc, HR_COLORS.muted);
    doc.text(m.label, midX, y + 16, {align:'center'});
  }
  return y + 22;
}

// --- Summary box ---
function hrDrawSummaryBox(doc, y, text) {
  doc.setFillColor.apply(doc, HR_COLORS.dune);
  var lines = doc.splitTextToSize(text, HR_CW - 8);
  var boxH = lines.length * 4.5 + 6;
  pdfRRect(doc, HR_ML, y, HR_CW, boxH, 2, 2, 'F');
  doc.setFontSize(10); doc.setFont('helvetica','normal');
  doc.setTextColor.apply(doc, HR_COLORS.gray);
  doc.text(lines, HR_ML + 4, y + 5);
  return y + boxH + 3;
}

// =============================================================================
// MAIN EXPORT FUNCTION
// =============================================================================
function exportHealthReport() {
  // Pre-flight check
  var entities = ['company','contact','sale','project'];
  for (var i = 0; i < entities.length; i++) {
    if (!entityScores[entities[i]] || !overviewData[entities[i]]) {
      alert('Please run Analyze All first before exporting the Health Report.');
      return;
    }
  }
  if (!momentumData) {
    alert('Momentum data not loaded. Please run Analyze All first.');
    return;
  }

  var jsPDF = window.jspdf.jsPDF;
  var doc = new jsPDF({unit:'mm', format:'a4'});

  // Gather scores
  var scores = {};
  for (var ei = 0; ei < entities.length; ei++) {
    var k = entities[ei];
    scores[k] = entityScores[k] || computeEntityScores(k);
  }

  // Determine included chapters and page numbers
  var chapters = [];
  chapters.push({id:'executive', title:'Executive Summary', page:3});
  var pageCounter = 4;
  for (var ci = 0; ci < entities.length; ci++) {
    var ek = entities[ci];
    var ov = overviewData[ek].overview;
    var s = scores[ek];
    chapters.push({
      id: ek,
      title: ek.charAt(0).toUpperCase() + ek.slice(1),
      page: pageCounter,
      records: ov.total,
      health: s ? Math.round(s.health) : 0
    });
    pageCounter++;
  }
  chapters.push({id:'momentum', title:'CRM Momentum', page:pageCounter, users: momentumData.totalUsers || 0});
  var totalPages = pageCounter;

  // Report date
  var rptDate = new Date().toLocaleDateString('en-GB', {year:'numeric', month:'long', day:'numeric'});
  var rptMonth = new Date().toLocaleDateString('en-GB', {year:'numeric', month:'long'});

  // ===== PAGE 1: COVER =====
  doc.addImage(HR_COVER_B64, 'JPEG', 0, 0, HR_PAGE_W, HR_PAGE_H);
  // Logo top-right
  doc.addImage(HR_LOGO_B64, 'PNG', HR_PAGE_W - 52, 10, 40, 6);
  // Text block bottom-left
  var ty = HR_PAGE_H - 68;
  doc.setFontSize(10); doc.setFont('helvetica','bold');
  doc.setTextColor.apply(doc, HR_COLORS.muted);
  doc.text('PREPARED FOR', HR_ML, ty);
  ty += 6;
  doc.setFontSize(18); doc.setFont('helvetica','bold');
  doc.setTextColor.apply(doc, HR_COLORS.soGreen);
  doc.text('[Company Name]', HR_ML, ty);
  ty += 12;
  doc.setFontSize(38); doc.setFont('helvetica','bold');
  doc.setTextColor.apply(doc, HR_COLORS.soGreen);
  doc.text('CRM Health', HR_ML, ty);
  ty += 14;
  doc.text('Report', HR_ML, ty);
  ty += 8;
  doc.setFontSize(13); doc.setFont('helvetica','normal');
  doc.setTextColor.apply(doc, HR_COLORS.gray);
  doc.text('Data quality, integrity and adoption analysis', HR_ML, ty);
  ty += 8;
  doc.setFontSize(10); doc.setFont('helvetica','normal');
  doc.setTextColor.apply(doc, HR_COLORS.muted);
  doc.text('Generated: ' + rptDate + ' \u00b7 Period: All data', HR_ML, ty);
  // Green bar bottom
  doc.setFillColor.apply(doc, HR_COLORS.soGreen);
  doc.rect(0, HR_PAGE_H - 5, HR_PAGE_W, 5, 'F');
  // Page number
  doc.setFontSize(8); doc.setTextColor.apply(doc, HR_COLORS.muted);
  doc.text('1 / ' + totalPages, HR_PAGE_W - HR_MR, HR_PAGE_H - 10, {align:'right'});

  // ===== PAGE 2: TABLE OF CONTENTS =====
  doc.addPage();
  hrDrawPageHeader(doc);
  hrDrawPageFooter(doc, 2, totalPages);
  var tocY = 25;
  doc.setFontSize(24); doc.setFont('helvetica','bold');
  doc.setTextColor.apply(doc, HR_COLORS.soGreen);
  doc.text('Table of Contents', HR_ML, tocY);
  tocY += 12;

  for (var ti = 0; ti < chapters.length; ti++) {
    var ch = chapters[ti];
    doc.setFontSize(12); doc.setFont('helvetica','bold');
    doc.setTextColor.apply(doc, HR_COLORS.charcoal);
    doc.text((ti + 1) + '.', HR_ML, tocY);
    doc.setFont('helvetica','normal');
    doc.text(ch.title, HR_ML + 8, tocY);
    // Subtitle
    var subTxt = '';
    if (ch.records) subTxt = fmtNum(ch.records) + ' records \u00b7 ' + ch.health + '%';
    else if (ch.users) subTxt = ch.users + ' users';
    if (subTxt) {
      doc.setFontSize(9); doc.setTextColor.apply(doc, HR_COLORS.muted);
      var titleW = doc.getTextWidth(ch.title);
      doc.text(subTxt, HR_ML + 8 + titleW + 4, tocY);
    }
    // Page number right
    doc.setFontSize(12); doc.setTextColor.apply(doc, HR_COLORS.muted);
    doc.text(String(ch.page), HR_PAGE_W - HR_MR, tocY, {align:'right'});
    // Clickable link to page
    doc.link(HR_ML, tocY - 4, HR_CW, 8, {pageNumber: ch.page});
    // Dotted line
    doc.setDrawColor.apply(doc, HR_COLORS.border);
    doc.setLineDashPattern([0.5, 1.5], 0);
    doc.setLineWidth(0.2);
    doc.line(HR_ML, tocY + 2, HR_PAGE_W - HR_MR, tocY + 2);
    doc.setLineDashPattern([], 0); // reset
    tocY += 9;
  }
  // Config note
  tocY += 8;
  doc.setFontSize(9); doc.setFont('helvetica','italic');
  doc.setTextColor.apply(doc, HR_COLORS.muted);
  doc.text('This report was generated from the SuperOffice Data Analysis Dashboard.', HR_ML, tocY);
  tocY += 4;
  doc.text('Entity chapters can be configured in the export settings.', HR_ML, tocY);

  // ===== PAGE 3: EXECUTIVE SUMMARY =====
  doc.addPage();
  hrDrawPageHeader(doc);
  hrDrawPageFooter(doc, 3, totalPages);
  var ey = 16;

  // Title
  doc.setFontSize(10); doc.setFont('helvetica','bold');
  doc.setTextColor.apply(doc, HR_COLORS.muted);
  doc.text('OVERALL HEALTH SCORES', HR_PAGE_W / 2, ey, {align:'center'});
  ey += 4;

  // Donut rings
  var ringR = 12; // radius in mm
  var ringY = ey + ringR + 2;
  var ringSpacing = HR_CW / 4;
  var ringEntities = ['company','contact','sale','project'];
  var ringLabels = ['Company','Contact','Sale','Project'];
  for (var ri = 0; ri < 4; ri++) {
    var rk = ringEntities[ri];
    var rs = scores[rk];
    var rOv = overviewData[rk].overview;
    var rCx = HR_ML + ringSpacing * ri + ringSpacing / 2;
    hrDrawDonut(doc, rCx, ringY, ringR, rs ? Math.round(rs.health) : 0, ringLabels[ri], fmtNum(rOv.total) + ' records');
  }
  ey = ringY + ringR + 12;

  // Executive summary text
  var summaryText = hrGenerateSummary(scores, entities);
  doc.setFontSize(11); doc.setFont('helvetica','bold');
  doc.setTextColor.apply(doc, HR_COLORS.charcoal);
  doc.text('Executive Summary', HR_ML + 4, ey + 4);
  doc.setFont('helvetica','normal');
  doc.setFontSize(10);
  doc.setTextColor.apply(doc, HR_COLORS.gray);
  var sumLines = doc.splitTextToSize(summaryText, HR_CW - 8);
  var sumBoxH = sumLines.length * 4.2 + 10;
  doc.setFillColor.apply(doc, HR_COLORS.dune);
  pdfRRect(doc, HR_ML, ey, HR_CW, sumBoxH, 2, 2, 'F');
  doc.setFontSize(11); doc.setFont('helvetica','bold');
  doc.setTextColor.apply(doc, HR_COLORS.charcoal);
  doc.text('Executive Summary', HR_ML + 4, ey + 5);
  doc.setFont('helvetica','normal');
  doc.setFontSize(9.5);
  doc.setTextColor.apply(doc, HR_COLORS.gray);
  doc.text(sumLines, HR_ML + 4, ey + 10);
  ey += sumBoxH + 4;

  // Score Breakdown table
  ey = hrSectionTitle(doc, ey, 'Score Breakdown');
  var sbHead = [['Entity','Records','Data Quality','Data Integrity','Adoption','Overall Health']];
  var sbBody = [];
  for (var si = 0; si < entities.length; si++) {
    var sk = entities[si];
    var ss = scores[sk];
    var so = overviewData[sk].overview;
    sbBody.push([
      sk.charAt(0).toUpperCase() + sk.slice(1),
      fmtNum(so.total),
      ss ? Math.round(ss.dq) + '%' : '-',
      ss ? Math.round(ss.integrity) + '%' : '-',
      ss ? Math.round(ss.adoption) + '%' : '-',
      ss ? Math.round(ss.health) + '%' : '-'
    ]);
  }
  doc.autoTable({
    startY: ey, margin:{left:HR_ML, right:HR_MR},
    head: sbHead, body: sbBody,
    styles:{font:'helvetica', fontSize:9.5, cellPadding:{top:1.5,bottom:1.5,left:2,right:2}},
    headStyles:{fillColor:HR_COLORS.white, textColor:HR_COLORS.muted, fontSize:7.5, fontStyle:'bold'},
    columnStyles:{
      0:{fontStyle:'bold'},
      2:{halign:'right'}, 3:{halign:'right'}, 4:{halign:'right'}, 5:{halign:'right', fontStyle:'bold'}
    },
    theme:'plain',
    tableLineColor:HR_COLORS.tblBorder, tableLineWidth:0.2,
    didParseCell: function(data) {
      // Color the percentage cells
      if (data.section === 'body' && data.column.index >= 2) {
        var val = parseInt(data.cell.raw);
        if (!isNaN(val)) {
          data.cell.styles.textColor = hrScoreColor(val);
        }
      }
    }
  });
  ey = doc.lastAutoTable.finalY + 4;

  // Top Issues table
  ey = hrSectionTitle(doc, ey, 'Top Issues Requiring Attention');
  var issues = hrGatherTopIssues(entities, scores);
  var issHead = [['Issue','Entity','Affected','% of Total','Impact']];
  var issBody = [];
  for (var ii = 0; ii < Math.min(issues.length, 8); ii++) {
    var iss = issues[ii];
    issBody.push([iss.name, iss.entity, fmtNum(iss.count), iss.pct.toFixed(1) + '%', iss.impact]);
  }
  doc.autoTable({
    startY: ey, margin:{left:HR_ML, right:HR_MR},
    head: issHead, body: issBody,
    styles:{font:'helvetica', fontSize:9.5, cellPadding:{top:1.5,bottom:1.5,left:2,right:2}},
    headStyles:{fillColor:HR_COLORS.white, textColor:HR_COLORS.muted, fontSize:7.5, fontStyle:'bold'},
    columnStyles:{2:{halign:'right'}, 3:{halign:'right'}, 4:{halign:'right'}},
    theme:'plain',
    tableLineColor:HR_COLORS.tblBorder, tableLineWidth:0.2
  });
  ey = doc.lastAutoTable.finalY + 4;

  // Momentum Snapshot
  if (momentumData) {
    ey = hrSectionTitle(doc, ey, 'CRM Momentum Snapshot');
    var users = momentumData.users || [];
    var totalU = momentumData.totalUsers || 0;
    var monthly = momentumData.monthly || [];
    var filterMonths = monthly.length;
    var levels = {Power:0, Regular:0, Low:0, Inactive:0};
    var levelActs = {Power:0, Regular:0, Low:0, Inactive:0};
    var totalActs = 0;
    for (var ui = 0; ui < users.length; ui++) {
      var u = users[ui];
      var uTotal = (u.activities || 0) + (u.documents || 0);
      var lvl = mmUserLevel(uTotal, filterMonths);
      levels[lvl]++;
      levelActs[lvl] += uTotal;
      totalActs += uTotal;
    }
    var mmCards = [
      {label:'Power Users', value: levels.Power, color: HR_COLORS.good, sub: totalU > 0 ? Math.round(levels.Power / totalU * 100) + '% of users' : ''},
      {label:'Regular Users', value: levels.Regular, color: HR_COLORS.blue, sub: totalU > 0 ? Math.round(levels.Regular / totalU * 100) + '% of users' : ''},
      {label:'Low Usage', value: levels.Low, color: HR_COLORS.warn, sub: totalU > 0 ? Math.round(levels.Low / totalU * 100) + '% of users' : ''},
      {label:'Inactive', value: levels.Inactive, color: HR_COLORS.bad, sub: totalU > 0 ? Math.round(levels.Inactive / totalU * 100) + '% of users' : ''}
    ];
    var mmCardW = (HR_CW - 9) / 4;
    for (var mi = 0; mi < 4; mi++) {
      var mc = mmCards[mi];
      var mx = HR_ML + mi * (mmCardW + 3);
      doc.setFillColor.apply(doc, HR_COLORS.dune);
      pdfRRect(doc, mx, ey, mmCardW, 16, 2, 2, 'F');
      var mmCx = mx + mmCardW / 2;
      doc.setFontSize(18); doc.setFont('helvetica','bold');
      doc.setTextColor.apply(doc, mc.color);
      doc.text(String(mc.value), mmCx, ey + 6.5, {align:'center'});
      doc.setFontSize(8.5); doc.setFont('helvetica','bold');
      doc.setTextColor.apply(doc, HR_COLORS.charcoal);
      doc.text(mc.label, mmCx, ey + 10.5, {align:'center'});
      doc.setFontSize(7.5); doc.setFont('helvetica','normal');
      doc.setTextColor.apply(doc, HR_COLORS.muted);
      doc.text(mc.sub, mmCx, ey + 13.5, {align:'center'});
    }
  }

  // ===== SAVE =====
  doc.save('CRM_Health_Report_' + new Date().toISOString().split('T')[0] + '.pdf');
}

// --- Generate programmatic executive summary ---
function hrGenerateSummary(scores, entities) {
  var parts = [];
  var vals = [];
  for (var i = 0; i < entities.length; i++) {
    var s = scores[entities[i]];
    if (s) vals.push({key: entities[i], health: Math.round(s.health)});
  }
  vals.sort(function(a,b){ return b.health - a.health; });

  var avg = 0;
  for (var j = 0; j < vals.length; j++) avg += vals[j].health;
  avg = Math.round(avg / vals.length);
  var level = avg >= 70 ? 'good' : (avg >= 50 ? 'moderate' : 'needs improvement');
  parts.push('Overall CRM health is ' + level + ' (' + avg + '%).');

  if (vals.length > 0) {
    parts.push(vals[0].key.charAt(0).toUpperCase() + vals[0].key.slice(1) + ' data scores highest at ' + vals[0].health + '%.');
  }
  if (vals.length > 1) {
    var worst = vals[vals.length - 1];
    parts.push(worst.key.charAt(0).toUpperCase() + worst.key.slice(1) + ' scores lowest at ' + worst.health + '% and needs the most attention.');
  }

  // Company-specific insight
  if (overviewData.company) {
    var co = overviewData.company.overview;
    if (co.total > 0) {
      var pctPerson = Math.round((co.withPersons || 0) / co.total * 100);
      parts.push('Company person coverage is ' + pctPerson + '% (' + fmtNum(co.withPersons || 0) + ' of ' + fmtNum(co.total) + ').');
    }
  }

  // Momentum insight
  if (momentumData) {
    parts.push('CRM Momentum shows ' + (momentumData.totalUsers || 0) + ' active users over the last 24 months.');
  }

  return parts.join(' ');
}

// --- Gather cross-entity top issues ---
function hrGatherTopIssues(entities, scores) {
  var issues = [];
  for (var i = 0; i < entities.length; i++) {
    var k = entities[i];
    var data = gatherEntityData(k);
    if (!data) continue;
    var total = data.integrityTotal || data.total;
    var cd = data.checkData;
    var eName = k.charAt(0).toUpperCase() + k.slice(1);
    for (var prop in cd) {
      if (!cd.hasOwnProperty(prop)) continue;
      var count = cd[prop];
      if (count <= 0) continue;
      var pct = total > 0 ? count / total * 100 : 0;
      var name = prop.replace(/([A-Z])/g, ' $1').replace(/^./, function(s){return s.toUpperCase();});
      name = name.replace('No ', 'No ').replace('Stale ', 'Stale ');
      issues.push({name: name, entity: eName, count: count, pct: pct, impact: pct > 40 ? 'High' : (pct > 20 ? 'Medium' : 'Low')});
    }
  }
  issues.sort(function(a,b){ return b.pct - a.pct; });
  return issues;
}
