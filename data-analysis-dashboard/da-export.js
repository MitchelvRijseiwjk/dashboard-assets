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
