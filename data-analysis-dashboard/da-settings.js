// da-settings.js — Dashboard Settings Manager v2
// Manages configurable scoring, completeness (incl. UDEF), engagement, and pipeline
// Persisted to localStorage per installation

(function() {
  'use strict';

  // ============================================================
  // DEFAULT SETTINGS
  // ============================================================
  var DEFAULTS = {
    company: {
      completenessFields: ['email', 'phone', 'person', 'category', 'orgNr'],
      dqWeights: { completeness: 40, udef: 30, quality: 30 },
      qualityIssueFields: ['noPerson', 'noCategory', 'noBusiness', 'unreachable'],
      udefFieldConfig: {},
      engagementWeights: { withActivity: 50, withPerson: 30, withPipeline: 20 },
      pipelineType: 'sale'
    }
  };

  var COMPLETENESS_OPTIONS = [
    { key: 'email',    label: 'Email address',    source: 'completeness' },
    { key: 'phone',    label: 'Phone number',     source: 'completeness' },
    { key: 'person',   label: 'Contact person',   source: 'quality', qualityKey: 'noPerson' },
    { key: 'category', label: 'Category',         source: 'quality', qualityKey: 'noCategory' },
    { key: 'orgNr',    label: 'Org. number',      source: 'completeness' },
    { key: 'business', label: 'Business type',    source: 'quality', qualityKey: 'noBusiness' },
    { key: 'address',  label: 'Address',          source: 'completeness' },
    { key: 'webpage',  label: 'Webpage',          source: 'completeness' }
  ];

  var QUALITY_ISSUE_OPTIONS = [
    { key: 'noPerson',    label: 'No contact person' },
    { key: 'noCategory',  label: 'No category' },
    { key: 'noBusiness',  label: 'No business type' },
    { key: 'unreachable', label: 'Unreachable (no email or phone)' },
    { key: 'noOrgNr',     label: 'No org. number' }
  ];

  var PIPELINE_OPTIONS = [
    { key: 'sale',    label: 'Open Sale',            desc: 'Companies with an open sale opportunity' },
    { key: 'project', label: 'Active Project',       desc: 'Companies with an active project', soon: true },
    { key: 'both',    label: 'Sale or Project',      desc: 'Companies with either sale or project', soon: true },
    { key: 'none',    label: 'None (activity only)',  desc: 'Funnel stops at activity — for relationship-only CRM' }
  ];

  var PRESETS = {
    standard: {
      label: 'Standard',
      desc: 'Balanced scoring for general CRM usage',
      s: { completenessFields: ['email', 'phone', 'person', 'category', 'orgNr'], dqWeights: { completeness: 40, udef: 30, quality: 30 }, qualityIssueFields: ['noPerson', 'noCategory', 'noBusiness', 'unreachable'], engagementWeights: { withActivity: 50, withPerson: 30, withPipeline: 20 }, pipelineType: 'sale' }
    },
    salesFocus: {
      label: 'Sales Focus',
      desc: 'Emphasizes pipeline and commercial data quality',
      s: { completenessFields: ['email', 'phone', 'person', 'category', 'orgNr', 'business'], dqWeights: { completeness: 30, udef: 20, quality: 50 }, qualityIssueFields: ['noPerson', 'noCategory', 'unreachable'], engagementWeights: { withActivity: 35, withPerson: 25, withPipeline: 40 }, pipelineType: 'sale' }
    },
    serviceFocus: {
      label: 'Service Focus',
      desc: 'Emphasizes reachability and contact data for service teams',
      s: { completenessFields: ['email', 'phone', 'person', 'address'], dqWeights: { completeness: 50, udef: 20, quality: 30 }, qualityIssueFields: ['noPerson', 'unreachable'], engagementWeights: { withActivity: 60, withPerson: 30, withPipeline: 10 }, pipelineType: 'none' }
    },
    projectBased: {
      label: 'Project Based',
      desc: 'For organizations that use projects instead of sales',
      s: { completenessFields: ['email', 'phone', 'person', 'category', 'orgNr'], dqWeights: { completeness: 40, udef: 30, quality: 30 }, qualityIssueFields: ['noPerson', 'noCategory', 'noBusiness'], engagementWeights: { withActivity: 50, withPerson: 30, withPipeline: 20 }, pipelineType: 'project' }
    }
  };

  // ============================================================
  // STATE
  // ============================================================
  var STORAGE_KEY = 'da_settings_v1';
  var _settings = null;
  var _discoveredUdefFields = [];
  var _activeContainer = null;

  function deepClone(obj) { return JSON.parse(JSON.stringify(obj)); }

  function loadSettings() {
    try {
      var raw = localStorage.getItem(STORAGE_KEY);
      if (raw) { _settings = mergeDefaults(JSON.parse(raw), DEFAULTS); return; }
    } catch(e) {}
    _settings = deepClone(DEFAULTS);
  }

  function mergeDefaults(saved, defaults) {
    var result = deepClone(defaults);
    if (!saved) return result;
    for (var entity in defaults) {
      if (!saved[entity]) continue;
      var s = saved[entity];
      var r = result[entity];
      if (s.completenessFields) r.completenessFields = s.completenessFields.slice();
      if (s.dqWeights) { for (var w in s.dqWeights) r.dqWeights[w] = s.dqWeights[w]; }
      if (s.qualityIssueFields) r.qualityIssueFields = s.qualityIssueFields.slice();
      if (s.udefFieldConfig) r.udefFieldConfig = deepClone(s.udefFieldConfig);
      if (s.engagementWeights) { for (var w in s.engagementWeights) r.engagementWeights[w] = s.engagementWeights[w]; }
      if (s.pipelineType) r.pipelineType = s.pipelineType;
    }
    return result;
  }

  function saveSettings() {
    try { localStorage.setItem(STORAGE_KEY, JSON.stringify(_settings)); } catch(e) {}
  }

  function getSettings(entity) {
    if (!_settings) loadSettings();
    return _settings[entity] || deepClone(DEFAULTS.company);
  }

  // ============================================================
  // UDEF FIELD DISCOVERY
  // ============================================================
  function notifyUdefLoaded(entityId, udefResult) {
    if (entityId !== 7) return;
    if (!udefResult || !udefResult.fields) return;
    _discoveredUdefFields = [];
    for (var i = 0; i < udefResult.fields.length; i++) {
      var f = udefResult.fields[i];
      _discoveredUdefFields.push({
        progId: f.progId || f.label || ('field_' + i),
        label: f.label, type: f.type, percent: f.percent
      });
    }
    if (_activeContainer) {
      var el = document.getElementById(_activeContainer);
      if (el && el.offsetParent !== null) {
        renderPanel(el, el.getAttribute('data-settings-mode') || 'full');
      }
    }
  }

  // ============================================================
  // COMPUTED HELPERS
  // ============================================================
  function getCompletenessValue(fieldKey, overviewData, qualityData, total) {
    if (!total || total <= 0) return 0;
    if (overviewData && overviewData[fieldKey] !== undefined) return overviewData[fieldKey];
    var qMap = { person: 'noPerson', category: 'noCategory', business: 'noBusiness' };
    if (qMap[fieldKey] && qualityData) {
      var v = qualityData[qMap[fieldKey]];
      if (v !== undefined) return total - v;
    }
    return 0;
  }

  function computeConfiguredDQScore(entity, overviewCpl, qualityData, udefData, total) {
    var s = getSettings(entity);
    var scores = {};

    if (overviewCpl && s.completenessFields.length > 0 && total > 0) {
      var sum = 0;
      for (var i = 0; i < s.completenessFields.length; i++) {
        sum += (getCompletenessValue(s.completenessFields[i], overviewCpl, qualityData, total) / total) * 100;
      }
      scores.completeness = sum / s.completenessFields.length;
    }

    if (udefData && udefData.fields && udefData.fields.length > 0) {
      var cfg = s.udefFieldConfig || {};
      var uSum = 0; var uW = 0;
      for (var i = 0; i < udefData.fields.length; i++) {
        var f = udefData.fields[i];
        var pid = f.progId || f.label || ('field_' + i);
        var imp = cfg[pid] || 'normal';
        if (imp === 'excluded') continue;
        var w = (imp === 'required') ? 2 : 1;
        uSum += f.percent * w;
        uW += w;
      }
      if (uW > 0) scores.udef = uSum / uW;
    }

    if (qualityData && s.qualityIssueFields.length > 0 && total > 0) {
      var issSum = 0;
      for (var i = 0; i < s.qualityIssueFields.length; i++) issSum += (qualityData[s.qualityIssueFields[i]] || 0);
      scores.quality = 100 - ((issSum / (total * s.qualityIssueFields.length)) * 100);
    }

    var wt = s.dqWeights; var totalW = 0; var wSum = 0;
    if (scores.completeness !== undefined) { wSum += scores.completeness * wt.completeness; totalW += wt.completeness; }
    if (scores.udef !== undefined) { wSum += scores.udef * wt.udef; totalW += wt.udef; }
    if (scores.quality !== undefined) { wSum += scores.quality * wt.quality; totalW += wt.quality; }
    if (totalW <= 0) return null;
    return { total: Math.round(wSum / totalW), components: scores, weights: wt };
  }

  function computeEngagement(catRow, total, entity) {
    if (!total || total <= 0) return 0;
    var s = getSettings(entity);
    var w = s.engagementWeights;
    var pVal = 0;
    if (s.pipelineType === 'sale') pVal = catRow.withSale || 0;
    else if (s.pipelineType === 'project') pVal = catRow.withProject || 0;
    else if (s.pipelineType === 'both') pVal = catRow.withSaleOrProject || 0;
    return Math.round(((catRow.withActivity || 0) * (w.withActivity / 100) + (catRow.withPerson || 0) * (w.withPerson / 100) + pVal * (w.withPipeline / 100)) / total * 100);
  }

  function getPipelineLabel(entity) {
    var s = getSettings(entity);
    for (var i = 0; i < PIPELINE_OPTIONS.length; i++) {
      if (PIPELINE_OPTIONS[i].key === s.pipelineType) return PIPELINE_OPTIONS[i].label;
    }
    return 'Open Sale';
  }

  function getPipelineType(entity) { return getSettings(entity).pipelineType; }

  // ============================================================
  // RENDER PANEL — mode: 'full' or 'dq'
  // ============================================================
  function renderPanel(container, mode) {
    if (!container) return;
    var s = getSettings('company');
    var h = '';

    // Preset bar
    h += '<div class="settings-preset-bar">';
    h += '<label class="settings-preset-label">Preset:</label>';
    h += '<select class="settings-preset-select ds-preset" onchange="daSettings.applyPreset(this.value)">';
    h += '<option value="">Custom</option>';
    for (var pk in PRESETS) {
      h += '<option value="' + pk + '"' + (isPresetMatch(pk) ? ' selected' : '') + '>' + PRESETS[pk].label + '</option>';
    }
    h += '</select>';
    h += '<span class="settings-preset-desc ds-preset-desc"></span>';
    h += '</div>';

    // ==== DATA QUALITY ====
    h += '<div class="settings-section">';
    h += '<div class="settings-section-head">Data Quality</div>';
    h += '<div class="settings-grid">';

    // DQ Score Weights
    h += '<div class="settings-card">';
    h += '<h3>DQ Score Weights</h3>';
    h += '<div class="settings-desc">How much each component contributes to the overall DQ Score.</div>';
    h += renderSliders('dq', [
      { key: 'completeness', label: 'Field Completeness', val: s.dqWeights.completeness },
      { key: 'udef',         label: 'Custom Fields (UDEF)', val: s.dqWeights.udef },
      { key: 'quality',      label: 'Quality Issues', val: s.dqWeights.quality }
    ]);
    h += '</div>';

    // Quality Issue Checks
    h += '<div class="settings-card">';
    h += '<h3>Quality Issue Checks</h3>';
    h += '<div class="settings-desc">Which issues are included in the DQ score.</div>';
    h += '<div class="settings-field-grid">';
    for (var i = 0; i < QUALITY_ISSUE_OPTIONS.length; i++) {
      var opt = QUALITY_ISSUE_OPTIONS[i];
      h += '<label class="settings-checkbox"><input type="checkbox" data-group="qi" data-key="' + opt.key + '"' + (s.qualityIssueFields.indexOf(opt.key) >= 0 ? ' checked' : '') + ' onchange="daSettings.markCustom()"> ' + opt.label + '</label>';
    }
    h += '</div></div>';

    h += '</div>'; // end grid

    // Completeness Definition — full width, two subsections
    h += '<div class="settings-completeness-wrap">';
    h += '<div class="settings-card">';
    h += '<h3>Completeness Definition</h3>';
    h += '<div class="settings-desc">Which fields count towards the completeness score. Custom fields marked <strong>Required</strong> count double, <strong>Excluded</strong> fields are ignored.</div>';

    // Sub 1: Standard Fields
    h += '<div class="cpl-sub"><div class="cpl-sub-head">Standard Fields</div>';
    h += '<div class="settings-field-grid">';
    for (var i = 0; i < COMPLETENESS_OPTIONS.length; i++) {
      var opt = COMPLETENESS_OPTIONS[i];
      h += '<label class="settings-checkbox"><input type="checkbox" data-group="cpl" data-key="' + opt.key + '"' + (s.completenessFields.indexOf(opt.key) >= 0 ? ' checked' : '') + ' onchange="daSettings.markCustom()"> ' + opt.label + '</label>';
    }
    h += '</div></div>';

    // Sub 2: Custom Fields (UDEF)
    h += '<div class="cpl-sub"><div class="cpl-sub-head">Custom Fields (UDEF)</div>';
    if (_discoveredUdefFields.length === 0) {
      h += '<div class="udef-config-empty">Custom fields will appear here after running the Company analysis.</div>';
    } else {
      h += '<table class="data-table udef-config-table"><thead><tr>';
      h += '<th>Field</th><th>Type</th><th class="col-right">Fill %</th><th style="width:160px">Importance</th>';
      h += '</tr></thead><tbody>';
      for (var ui = 0; ui < _discoveredUdefFields.length; ui++) {
        var uf = _discoveredUdefFields[ui];
        var curImp = s.udefFieldConfig[uf.progId] || 'normal';
        var fillCol = uf.percent >= 70 ? 'var(--sl-good)' : (uf.percent >= 30 ? 'var(--sl-ok)' : 'var(--sl-bad)');
        var escId = uf.progId.replace(/'/g, "\\'");
        h += '<tr class="udef-row' + (curImp === 'excluded' ? ' udef-excluded' : '') + '">';
        h += '<td><span style="font-weight:500">' + uf.label + '</span>';
        h += '<span class="udef-progid">' + uf.progId + '</span></td>';
        h += '<td style="font-size:.78rem;color:var(--so-text-muted)">' + (uf.type || '') + '</td>';
        h += '<td class="col-right"><span style="font-weight:600;color:' + fillCol + '">' + uf.percent + '%</span></td>';
        h += '<td><div class="udef-imp-toggle">';
        h += '<button class="udef-imp-btn' + (curImp === 'required' ? ' active req' : '') + '" onclick="daSettings.setUdefImp(\'' + escId + '\',\'required\',this)">Required</button>';
        h += '<button class="udef-imp-btn' + (curImp === 'normal' ? ' active norm' : '') + '" onclick="daSettings.setUdefImp(\'' + escId + '\',\'normal\',this)">Normal</button>';
        h += '<button class="udef-imp-btn' + (curImp === 'excluded' ? ' active excl' : '') + '" onclick="daSettings.setUdefImp(\'' + escId + '\',\'excluded\',this)">Excluded</button>';
        h += '</div></td></tr>';
      }
      h += '</tbody></table>';
    }
    h += '</div>'; // end sub
    h += '</div></div>'; // end card + wrap

    h += '</div>'; // end DQ section

    // ==== ADOPTION (full mode only) ====
    if (mode === 'full') {
      h += '<div class="settings-section">';
      h += '<div class="settings-section-head">Adoption</div>';
      h += '<div class="settings-grid">';

      h += '<div class="settings-card">';
      h += '<h3>Engagement Score Weights</h3>';
      h += '<div class="settings-desc">How metrics contribute to the engagement score in Category Effectiveness.</div>';
      h += renderSliders('eng', [
        { key: 'withActivity', label: 'Activity', val: s.engagementWeights.withActivity },
        { key: 'withPerson',   label: 'Contact Person', val: s.engagementWeights.withPerson },
        { key: 'withPipeline', label: 'Pipeline', val: s.engagementWeights.withPipeline }
      ]);
      h += '</div>';

      h += '<div class="settings-card">';
      h += '<h3>Pipeline Definition</h3>';
      h += '<div class="settings-desc">What counts as "pipeline" in the engagement funnel and health metrics.</div>';
      for (var i = 0; i < PIPELINE_OPTIONS.length; i++) {
        var opt = PIPELINE_OPTIONS[i];
        var checked = s.pipelineType === opt.key ? ' checked' : '';
        var disabled = opt.soon ? ' disabled' : '';
        var cls = opt.soon ? ' settings-radio-soon' : '';
        h += '<label class="settings-radio' + cls + '">';
        h += '<input type="radio" name="pipelineType" value="' + opt.key + '"' + checked + disabled + ' onchange="daSettings.markCustom()">';
        h += '<span><strong>' + opt.label + '</strong>';
        h += '<span class="settings-radio-desc">' + opt.desc + '</span>';
        if (opt.soon) h += '<span class="settings-badge-soon">Coming soon</span>';
        h += '</span></label>';
      }
      h += '</div>';

      h += '</div></div>'; // end grid + section
    }

    // Save/Reset
    h += '<div class="settings-actions">';
    h += '<button class="btn-settings-save" onclick="daSettings.save()">Apply &amp; Recalculate</button>';
    h += '<button class="btn-settings-reset" onclick="daSettings.resetDefaults()">Reset to Defaults</button>';
    h += '</div>';

    container.innerHTML = h;
    container.setAttribute('data-settings-mode', mode);
    _activeContainer = container.id;
    updatePresetDesc();
  }

  // ============================================================
  // SLIDERS
  // ============================================================
  function renderSliders(gid, items) {
    var h = '<div class="weight-slider-group">';
    for (var i = 0; i < items.length; i++) {
      var it = items[i];
      h += '<div class="weight-slider-row">';
      h += '<label class="weight-slider-label">' + it.label + '</label>';
      h += '<input type="range" class="weight-slider" id="ws_' + gid + '_' + it.key + '" data-group="' + gid + '" data-key="' + it.key + '" min="0" max="100" step="5" value="' + it.val + '" oninput="daSettings.onSlider(this)">';
      h += '<span class="weight-slider-value" id="wsv_' + gid + '_' + it.key + '">' + it.val + '%</span>';
      h += '</div>';
    }
    h += '<div class="weight-slider-total" id="wst_' + gid + '">Total: 100%</div></div>';
    return h;
  }

  function onSlider(slider) {
    var g = slider.getAttribute('data-group');
    var k = slider.getAttribute('data-key');
    var v = parseInt(slider.value);
    var ve = document.getElementById('wsv_' + g + '_' + k);
    if (ve) ve.textContent = v + '%';
    updateSlTotal(g);
    markCustom();
  }

  function setSl(g, k, v) {
    var s = document.getElementById('ws_' + g + '_' + k);
    if (s) s.value = v;
    var ve = document.getElementById('wsv_' + g + '_' + k);
    if (ve) ve.textContent = v + '%';
  }

  function updateSlTotal(g) {
    var all = document.querySelectorAll('.weight-slider[data-group="' + g + '"]');
    var t = 0;
    for (var i = 0; i < all.length; i++) t += parseInt(all[i].value);
    var te = document.getElementById('wst_' + g);
    if (te) {
      te.textContent = 'Total: ' + t + '%';
      te.className = t === 100 ? 'weight-slider-total' : 'weight-slider-total weight-slider-total-warn';
    }
  }

  // ============================================================
  // PRESETS & HELPERS
  // ============================================================
  function markCustom() {
    var sel = document.querySelector('.ds-preset');
    if (sel) sel.value = '';
    updatePresetDesc();
  }

  function setUdefImp(progId, imp, btn) {
    var btns = btn.parentElement.querySelectorAll('.udef-imp-btn');
    for (var i = 0; i < btns.length; i++) btns[i].className = 'udef-imp-btn';
    btn.className = 'udef-imp-btn active ' + (imp === 'required' ? 'req' : (imp === 'normal' ? 'norm' : 'excl'));
    var row = btn.closest('tr');
    if (row) { if (imp === 'excluded') row.classList.add('udef-excluded'); else row.classList.remove('udef-excluded'); }
    markCustom();
  }

  function applyPreset(pk) {
    if (!pk || !PRESETS[pk]) return;
    var ps = PRESETS[pk].s;
    // Checkboxes
    var cplBoxes = document.querySelectorAll('input[data-group="cpl"]');
    for (var i = 0; i < cplBoxes.length; i++) cplBoxes[i].checked = ps.completenessFields.indexOf(cplBoxes[i].getAttribute('data-key')) >= 0;
    var qiBoxes = document.querySelectorAll('input[data-group="qi"]');
    for (var i = 0; i < qiBoxes.length; i++) qiBoxes[i].checked = ps.qualityIssueFields.indexOf(qiBoxes[i].getAttribute('data-key')) >= 0;
    // Sliders
    setSl('dq', 'completeness', ps.dqWeights.completeness); setSl('dq', 'udef', ps.dqWeights.udef); setSl('dq', 'quality', ps.dqWeights.quality);
    updateSlTotal('dq');
    if (ps.engagementWeights) {
      setSl('eng', 'withActivity', ps.engagementWeights.withActivity); setSl('eng', 'withPerson', ps.engagementWeights.withPerson); setSl('eng', 'withPipeline', ps.engagementWeights.withPipeline);
      updateSlTotal('eng');
    }
    // Pipeline
    var radios = document.querySelectorAll('input[name="pipelineType"]');
    for (var i = 0; i < radios.length; i++) radios[i].checked = radios[i].value === ps.pipelineType;
    // UDEF reset
    var uToggles = document.querySelectorAll('.udef-imp-toggle');
    for (var i = 0; i < uToggles.length; i++) {
      var btns = uToggles[i].querySelectorAll('.udef-imp-btn');
      for (var j = 0; j < btns.length; j++) { btns[j].className = 'udef-imp-btn'; if (btns[j].textContent === 'Normal') btns[j].className = 'udef-imp-btn active norm'; }
      var row = uToggles[i].closest('tr'); if (row) row.classList.remove('udef-excluded');
    }
    updatePresetDesc();
  }

  function isPresetMatch(pk) {
    var s = getSettings('company'); var ps = PRESETS[pk].s;
    if (s.pipelineType !== ps.pipelineType) return false;
    if (JSON.stringify(s.dqWeights) !== JSON.stringify(ps.dqWeights)) return false;
    if (JSON.stringify(s.engagementWeights) !== JSON.stringify(ps.engagementWeights)) return false;
    if (JSON.stringify(s.completenessFields.slice().sort()) !== JSON.stringify(ps.completenessFields.slice().sort())) return false;
    return true;
  }

  function updatePresetDesc() {
    var el = document.querySelector('.ds-preset-desc');
    var sel = document.querySelector('.ds-preset');
    if (!el || !sel) return;
    var pk = sel.value;
    el.textContent = pk && PRESETS[pk] ? PRESETS[pk].desc : 'You have custom settings';
  }

  // ============================================================
  // SAVE
  // ============================================================
  function save() {
    var cplF = []; var qiF = [];
    var cplB = document.querySelectorAll('input[data-group="cpl"]');
    for (var i = 0; i < cplB.length; i++) { if (cplB[i].checked) cplF.push(cplB[i].getAttribute('data-key')); }
    var qiB = document.querySelectorAll('input[data-group="qi"]');
    for (var i = 0; i < qiB.length; i++) { if (qiB[i].checked) qiF.push(qiB[i].getAttribute('data-key')); }

    var dqW = readSl('dq'); var engW = readSl('eng');

    var pipType = 'sale';
    var radios = document.querySelectorAll('input[name="pipelineType"]');
    for (var i = 0; i < radios.length; i++) { if (radios[i].checked) { pipType = radios[i].value; break; } }

    // UDEF
    var uCfg = {};
    var uToggles = document.querySelectorAll('.udef-imp-toggle');
    for (var i = 0; i < uToggles.length; i++) {
      var act = uToggles[i].querySelector('.udef-imp-btn.active');
      if (!act) continue;
      var row = uToggles[i].closest('tr');
      var pidEl = row ? row.querySelector('.udef-progid') : null;
      if (!pidEl) continue;
      var pid = pidEl.textContent;
      if (act.classList.contains('req')) uCfg[pid] = 'required';
      else if (act.classList.contains('excl')) uCfg[pid] = 'excluded';
    }

    // Validate
    var dqT = (dqW.completeness || 0) + (dqW.udef || 0) + (dqW.quality || 0);
    if (dqT !== 100) { alert('DQ Score weights must total 100% (currently ' + dqT + '%)'); return; }
    if (engW.withActivity !== undefined) {
      var engT = (engW.withActivity || 0) + (engW.withPerson || 0) + (engW.withPipeline || 0);
      if (engT !== 100) { alert('Engagement weights must total 100% (currently ' + engT + '%)'); return; }
    }

    var prev = _settings.company;
    _settings.company = {
      completenessFields: cplF,
      dqWeights: { completeness: dqW.completeness || 40, udef: dqW.udef || 30, quality: dqW.quality || 30 },
      qualityIssueFields: qiF,
      udefFieldConfig: uCfg,
      engagementWeights: engW.withActivity !== undefined ? { withActivity: engW.withActivity, withPerson: engW.withPerson, withPipeline: engW.withPipeline } : prev.engagementWeights,
      pipelineType: pipType
    };
    saveSettings();

    if (typeof renderDQScore === 'function') renderDQScore('company');
    if (typeof renderCrossEntityFunnel === 'function' && typeof companyDetailData !== 'undefined' && companyDetailData) renderCrossEntityFunnel(companyDetailData);
    if (typeof renderCompanyDetails === 'function' && typeof companyDetailData !== 'undefined' && companyDetailData) renderCompanyDetails(companyDetailData);

    showToast('Settings applied and scores recalculated');
  }

  function readSl(g) {
    var r = {}; var sl = document.querySelectorAll('.weight-slider[data-group="' + g + '"]');
    for (var i = 0; i < sl.length; i++) r[sl[i].getAttribute('data-key')] = parseInt(sl[i].value);
    return r;
  }

  function resetDefaults() {
    _settings = deepClone(DEFAULTS); saveSettings();
    if (_activeContainer) {
      var el = document.getElementById(_activeContainer);
      if (el) renderPanel(el, el.getAttribute('data-settings-mode') || 'full');
    }
    showToast('Settings reset to defaults');
  }

  function showToast(msg) {
    var ex = document.getElementById('settingsToast'); if (ex) ex.remove();
    var t = document.createElement('div'); t.id = 'settingsToast'; t.className = 'settings-toast'; t.textContent = msg;
    document.body.appendChild(t);
    setTimeout(function() { t.classList.add('show'); }, 10);
    setTimeout(function() { t.classList.remove('show'); setTimeout(function() { t.remove(); }, 300); }, 2500);
  }

  // ============================================================
  // PUBLIC
  // ============================================================
  function init() { loadSettings(); var p = document.getElementById('settings-quality'); if (p) renderPanel(p, 'full'); }
  function renderInto(elId, mode) { loadSettings(); var el = document.getElementById(elId); if (el) renderPanel(el, mode || 'full'); }

  window.daSettings = {
    init: init, renderInto: renderInto,
    getSettings: getSettings, computeDQScore: computeConfiguredDQScore, computeEngagement: computeEngagement,
    getCompletenessValue: getCompletenessValue, getPipelineLabel: getPipelineLabel, getPipelineType: getPipelineType,
    notifyUdefLoaded: notifyUdefLoaded,
    COMPLETENESS_OPTIONS: COMPLETENESS_OPTIONS, QUALITY_ISSUE_OPTIONS: QUALITY_ISSUE_OPTIONS,
    onSlider: onSlider, markCustom: markCustom, setUdefImp: setUdefImp,
    applyPreset: applyPreset, save: save, resetDefaults: resetDefaults
  };

  loadSettings();
})();
