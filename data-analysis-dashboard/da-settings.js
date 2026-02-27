// da-settings.js — Dashboard Settings Manager
// Manages configurable scoring, completeness, engagement, and pipeline definitions
// Persisted to localStorage per installation

(function() {
  'use strict';

  // ============================================================
  // DEFAULT SETTINGS
  // ============================================================
  var DEFAULTS = {
    company: {
      // --- DATA QUALITY ---
      completenessFields: ['email', 'phone', 'person', 'category', 'orgNr'],
      dqWeights: { completeness: 40, udef: 30, quality: 30 },
      qualityIssueFields: ['noPerson', 'noCategory', 'noBusiness', 'unreachable'],

      // --- ADOPTION ---
      engagementWeights: { withActivity: 50, withPerson: 30, withPipeline: 20 },
      pipelineType: 'sale'  // 'sale' | 'project' | 'both' | 'none'
    }
  };

  // All available completeness fields with labels + data source
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
    { key: 'sale',    label: 'Open Sale',           desc: 'Companies with an open sale opportunity' },
    { key: 'project', label: 'Active Project',      desc: 'Companies with an active project', needsBackend: true },
    { key: 'both',    label: 'Sale or Project',     desc: 'Companies with either sale or project', needsBackend: true },
    { key: 'none',    label: 'None (activity only)', desc: 'Funnel stops at activity — for relationship-only CRM' }
  ];

  var PRESETS = {
    standard: {
      label: 'Standard',
      desc: 'Balanced scoring for general CRM usage',
      settings: {
        completenessFields: ['email', 'phone', 'person', 'category', 'orgNr'],
        dqWeights: { completeness: 40, udef: 30, quality: 30 },
        qualityIssueFields: ['noPerson', 'noCategory', 'noBusiness', 'unreachable'],
        engagementWeights: { withActivity: 50, withPerson: 30, withPipeline: 20 },
        pipelineType: 'sale'
      }
    },
    salesFocus: {
      label: 'Sales Focus',
      desc: 'Emphasizes pipeline and commercial data quality',
      settings: {
        completenessFields: ['email', 'phone', 'person', 'category', 'orgNr', 'business'],
        dqWeights: { completeness: 30, udef: 20, quality: 50 },
        qualityIssueFields: ['noPerson', 'noCategory', 'unreachable'],
        engagementWeights: { withActivity: 35, withPerson: 25, withPipeline: 40 },
        pipelineType: 'sale'
      }
    },
    serviceFocus: {
      label: 'Service Focus',
      desc: 'Emphasizes reachability and contact data for service teams',
      settings: {
        completenessFields: ['email', 'phone', 'person', 'address'],
        dqWeights: { completeness: 50, udef: 20, quality: 30 },
        qualityIssueFields: ['noPerson', 'unreachable'],
        engagementWeights: { withActivity: 60, withPerson: 30, withPipeline: 10 },
        pipelineType: 'none'
      }
    },
    projectBased: {
      label: 'Project Based',
      desc: 'For organizations that use projects instead of sales',
      settings: {
        completenessFields: ['email', 'phone', 'person', 'category', 'orgNr'],
        dqWeights: { completeness: 40, udef: 30, quality: 30 },
        qualityIssueFields: ['noPerson', 'noCategory', 'noBusiness'],
        engagementWeights: { withActivity: 50, withPerson: 30, withPipeline: 20 },
        pipelineType: 'project'
      }
    }
  };

  // ============================================================
  // STORAGE KEY
  // ============================================================
  var STORAGE_KEY = 'da_settings_v1';

  // ============================================================
  // SETTINGS STATE
  // ============================================================
  var _settings = null;

  function deepClone(obj) {
    return JSON.parse(JSON.stringify(obj));
  }

  function loadSettings() {
    try {
      var raw = localStorage.getItem(STORAGE_KEY);
      if (raw) {
        var parsed = JSON.parse(raw);
        // Merge with defaults to handle new fields added in updates
        _settings = mergeDefaults(parsed, DEFAULTS);
        return;
      }
    } catch(e) { /* ignore parse errors */ }
    _settings = deepClone(DEFAULTS);
  }

  function mergeDefaults(saved, defaults) {
    var result = deepClone(defaults);
    if (!saved) return result;
    // For each entity in defaults, overlay saved values
    for (var entity in defaults) {
      if (!saved[entity]) continue;
      var s = saved[entity];
      var r = result[entity];
      if (s.completenessFields) r.completenessFields = s.completenessFields.slice();
      if (s.dqWeights) {
        for (var w in s.dqWeights) { r.dqWeights[w] = s.dqWeights[w]; }
      }
      if (s.qualityIssueFields) r.qualityIssueFields = s.qualityIssueFields.slice();
      if (s.engagementWeights) {
        for (var w in s.engagementWeights) { r.engagementWeights[w] = s.engagementWeights[w]; }
      }
      if (s.pipelineType) r.pipelineType = s.pipelineType;
    }
    return result;
  }

  function saveSettings() {
    try {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(_settings));
    } catch(e) { /* storage full? ignore */ }
  }

  function getSettings(entity) {
    if (!_settings) loadSettings();
    return _settings[entity] || deepClone(DEFAULTS.company);
  }

  // ============================================================
  // COMPUTED HELPERS — used by da-entity-logic.js
  // ============================================================

  // Get completeness value for a field, combining overview.completeness + quality data
  function getCompletenessValue(fieldKey, overviewData, qualityData, total) {
    if (!total || total <= 0) return 0;
    // Direct completeness fields
    if (overviewData && overviewData[fieldKey] !== undefined) {
      return overviewData[fieldKey];
    }
    // Derived from quality (inverted: noPerson → hasPerson = total - noPerson)
    var qMap = { person: 'noPerson', category: 'noCategory', business: 'noBusiness' };
    if (qMap[fieldKey] && qualityData) {
      var issueCount = qualityData[qMap[fieldKey]];
      if (issueCount !== undefined) return total - issueCount;
    }
    return 0;
  }

  // Compute DQ score with configurable weights
  function computeConfiguredDQScore(entity, overviewCompleteness, qualityData, udefData, total) {
    var s = getSettings(entity);
    var scores = {};

    // 1. Completeness score
    if (overviewCompleteness && s.completenessFields.length > 0 && total > 0) {
      var sum = 0;
      for (var i = 0; i < s.completenessFields.length; i++) {
        var val = getCompletenessValue(s.completenessFields[i], overviewCompleteness, qualityData, total);
        sum += (val / total) * 100;
      }
      scores.completeness = sum / s.completenessFields.length;
    }

    // 2. UDEF fill rate
    if (udefData && udefData.fields && udefData.fields.length > 0) {
      var udefSum = 0;
      for (var i = 0; i < udefData.fields.length; i++) {
        udefSum += udefData.fields[i].percent;
      }
      scores.udef = udefSum / udefData.fields.length;
    }

    // 3. Quality issues (inverted: fewer issues = higher score)
    if (qualityData && s.qualityIssueFields.length > 0 && total > 0) {
      var issueSum = 0;
      for (var i = 0; i < s.qualityIssueFields.length; i++) {
        var issueVal = qualityData[s.qualityIssueFields[i]] || 0;
        issueSum += issueVal;
      }
      var issuePct = (issueSum / (total * s.qualityIssueFields.length)) * 100;
      scores.quality = 100 - issuePct;
    }

    // Weighted average using configured weights
    var w = s.dqWeights;
    var totalWeight = 0;
    var weightedSum = 0;

    if (scores.completeness !== undefined) {
      weightedSum += scores.completeness * w.completeness;
      totalWeight += w.completeness;
    }
    if (scores.udef !== undefined) {
      weightedSum += scores.udef * w.udef;
      totalWeight += w.udef;
    }
    if (scores.quality !== undefined) {
      weightedSum += scores.quality * w.quality;
      totalWeight += w.quality;
    }

    if (totalWeight <= 0) return null;
    return {
      total: Math.round(weightedSum / totalWeight),
      components: scores,
      weights: w
    };
  }

  // Compute engagement score for a category row
  function computeEngagement(catRow, total, entity) {
    if (!total || total <= 0) return 0;
    var s = getSettings(entity);
    var w = s.engagementWeights;

    var actPart = (catRow.withActivity || 0) * (w.withActivity / 100);
    var persPart = (catRow.withPerson || 0) * (w.withPerson / 100);

    var pipelineVal = 0;
    if (s.pipelineType === 'sale') {
      pipelineVal = catRow.withSale || 0;
    } else if (s.pipelineType === 'project') {
      pipelineVal = catRow.withProject || 0;
    } else if (s.pipelineType === 'both') {
      pipelineVal = catRow.withSaleOrProject || 0;
    }
    // pipelineType === 'none' → pipelineVal stays 0

    var pipePart = pipelineVal * (w.withPipeline / 100);
    return Math.round((actPart + persPart + pipePart) / total * 100);
  }

  // Get pipeline label for funnel display
  function getPipelineLabel(entity) {
    var s = getSettings(entity);
    for (var i = 0; i < PIPELINE_OPTIONS.length; i++) {
      if (PIPELINE_OPTIONS[i].key === s.pipelineType) return PIPELINE_OPTIONS[i].label;
    }
    return 'Open Sale';
  }

  function getPipelineType(entity) {
    return getSettings(entity).pipelineType;
  }

  // ============================================================
  // SETTINGS PANEL — render into existing DA_Main structure
  // ============================================================

  function renderSettingsPanel() {
    var panel = document.getElementById('settings-quality');
    if (!panel) return;

    var s = getSettings('company');
    var h = '';

    // Preset selector
    h += '<div class="settings-preset-bar">';
    h += '<label class="settings-preset-label">Preset:</label>';
    h += '<select id="settingsPreset" class="settings-preset-select" onchange="daSettings.applyPreset(this.value)">';
    h += '<option value="">Custom</option>';
    for (var pk in PRESETS) {
      var p = PRESETS[pk];
      var sel = isPresetActive(pk) ? ' selected' : '';
      h += '<option value="' + pk + '"' + sel + '>' + p.label + '</option>';
    }
    h += '</select>';
    h += '<span class="settings-preset-desc" id="presetDesc"></span>';
    h += '</div>';

    // ---- DATA QUALITY SECTION ----
    h += '<div class="settings-section">';
    h += '<div class="settings-section-head">Data Quality</div>';

    h += '<div class="settings-grid">';

    // DQ Score Weights
    h += '<div class="settings-card">';
    h += '<h3>DQ Score Weights</h3>';
    h += '<div class="settings-desc">How much each component contributes to the overall Data Quality Score.</div>';
    h += renderWeightSliders('dq', [
      { key: 'completeness', label: 'Field Completeness', val: s.dqWeights.completeness },
      { key: 'udef',         label: 'Custom Fields (UDEF)', val: s.dqWeights.udef },
      { key: 'quality',      label: 'Quality Issues', val: s.dqWeights.quality }
    ]);
    h += '</div>';

    // Completeness Definition
    h += '<div class="settings-card">';
    h += '<h3>Completeness Definition</h3>';
    h += '<div class="settings-desc">Which standard fields count towards the completeness score.</div>';
    h += '<div class="settings-field-grid">';
    for (var i = 0; i < COMPLETENESS_OPTIONS.length; i++) {
      var opt = COMPLETENESS_OPTIONS[i];
      var checked = s.completenessFields.indexOf(opt.key) >= 0 ? ' checked' : '';
      h += '<label class="settings-checkbox"><input type="checkbox" data-group="cpl" data-key="' + opt.key + '"' + checked + ' onchange="daSettings.onFieldToggle()"> ' + opt.label + '</label>';
    }
    h += '</div>';
    h += '</div>';

    // Quality Issues
    h += '<div class="settings-card full-width">';
    h += '<h3>Quality Issue Checks</h3>';
    h += '<div class="settings-desc">Which data quality issues are included in the DQ score calculation.</div>';
    h += '<div class="settings-field-grid">';
    for (var i = 0; i < QUALITY_ISSUE_OPTIONS.length; i++) {
      var opt = QUALITY_ISSUE_OPTIONS[i];
      var checked = s.qualityIssueFields.indexOf(opt.key) >= 0 ? ' checked' : '';
      h += '<label class="settings-checkbox"><input type="checkbox" data-group="qi" data-key="' + opt.key + '"' + checked + ' onchange="daSettings.onFieldToggle()"> ' + opt.label + '</label>';
    }
    h += '</div>';
    h += '</div>';

    h += '</div>'; // end grid
    h += '</div>'; // end section

    // ---- ADOPTION SECTION ----
    h += '<div class="settings-section">';
    h += '<div class="settings-section-head">Adoption</div>';

    h += '<div class="settings-grid">';

    // Engagement Weights
    h += '<div class="settings-card">';
    h += '<h3>Engagement Score Weights</h3>';
    h += '<div class="settings-desc">How metrics contribute to the engagement score in Category Effectiveness.</div>';
    h += renderWeightSliders('eng', [
      { key: 'withActivity', label: 'Activity', val: s.engagementWeights.withActivity },
      { key: 'withPerson',   label: 'Contact Person', val: s.engagementWeights.withPerson },
      { key: 'withPipeline', label: 'Pipeline', val: s.engagementWeights.withPipeline }
    ]);
    h += '</div>';

    // Pipeline Definition
    h += '<div class="settings-card">';
    h += '<h3>Pipeline Definition</h3>';
    h += '<div class="settings-desc">What counts as "pipeline" in the engagement funnel and health metrics.</div>';
    for (var i = 0; i < PIPELINE_OPTIONS.length; i++) {
      var opt = PIPELINE_OPTIONS[i];
      var checked = s.pipelineType === opt.key ? ' checked' : '';
      var disabled = opt.needsBackend ? ' disabled' : '';
      var dimClass = opt.needsBackend ? ' style="opacity:0.5"' : '';
      h += '<label class="settings-radio"' + dimClass + '>';
      h += '<input type="radio" name="pipelineType" value="' + opt.key + '"' + checked + disabled + ' onchange="daSettings.onPipelineChange(this.value)">';
      h += '<span><strong>' + opt.label + '</strong>';
      h += '<span class="settings-radio-desc">' + opt.desc + '</span>';
      if (opt.needsBackend) h += '<span class="settings-badge-soon">requires re-analysis</span>';
      h += '</span>';
      h += '</label>';
    }
    h += '</div>';

    h += '</div>'; // end grid
    h += '</div>'; // end section

    // Save/Reset buttons
    h += '<div class="settings-actions">';
    h += '<button class="btn-settings-save" onclick="daSettings.save()">Apply &amp; Recalculate</button>';
    h += '<button class="btn-settings-reset" onclick="daSettings.resetToDefaults()">Reset to Defaults</button>';
    h += '</div>';

    panel.innerHTML = h;
    updatePresetDesc();
    initSliders();
  }

  // ============================================================
  // WEIGHT SLIDERS — linked group that sums to 100
  // ============================================================

  function renderWeightSliders(groupId, items) {
    var h = '<div class="weight-slider-group" data-group="' + groupId + '">';
    for (var i = 0; i < items.length; i++) {
      var it = items[i];
      h += '<div class="weight-slider-row">';
      h += '<label class="weight-slider-label">' + it.label + '</label>';
      h += '<input type="range" class="weight-slider" id="ws_' + groupId + '_' + it.key + '" ';
      h += 'data-group="' + groupId + '" data-key="' + it.key + '" ';
      h += 'min="0" max="100" step="5" value="' + it.val + '" ';
      h += 'oninput="daSettings.onSliderChange(this)">';
      h += '<span class="weight-slider-value" id="wsv_' + groupId + '_' + it.key + '">' + it.val + '%</span>';
      h += '</div>';
    }
    h += '<div class="weight-slider-total" id="wst_' + groupId + '">Total: 100%</div>';
    h += '</div>';
    return h;
  }

  function initSliders() {
    // Set accent colors on range inputs
    var sliders = document.querySelectorAll('.weight-slider');
    for (var i = 0; i < sliders.length; i++) {
      updateSliderTrack(sliders[i]);
    }
  }

  function updateSliderTrack(slider) {
    var val = parseInt(slider.value);
    var pct = val + '%';
    slider.style.setProperty('--val', pct);
  }

  function onSliderChange(slider) {
    var group = slider.getAttribute('data-group');
    var key = slider.getAttribute('data-key');
    var val = parseInt(slider.value);

    // Update display
    var valEl = document.getElementById('wsv_' + group + '_' + key);
    if (valEl) valEl.textContent = val + '%';
    updateSliderTrack(slider);

    // Get all sliders in this group
    var allSliders = document.querySelectorAll('.weight-slider[data-group="' + group + '"]');
    var total = 0;
    for (var i = 0; i < allSliders.length; i++) {
      total += parseInt(allSliders[i].value);
    }

    // Show total (highlight if not 100)
    var totalEl = document.getElementById('wst_' + group);
    if (totalEl) {
      totalEl.textContent = 'Total: ' + total + '%';
      if (total === 100) {
        totalEl.className = 'weight-slider-total';
      } else {
        totalEl.className = 'weight-slider-total weight-slider-total-warn';
      }
    }

    // Update preset to "Custom"
    var presetSel = document.getElementById('settingsPreset');
    if (presetSel) presetSel.value = '';
    updatePresetDesc();
  }

  // ============================================================
  // EVENT HANDLERS
  // ============================================================

  function onFieldToggle() {
    var presetSel = document.getElementById('settingsPreset');
    if (presetSel) presetSel.value = '';
    updatePresetDesc();
  }

  function onPipelineChange(val) {
    var presetSel = document.getElementById('settingsPreset');
    if (presetSel) presetSel.value = '';
    updatePresetDesc();
  }

  function applyPreset(presetKey) {
    if (!presetKey || !PRESETS[presetKey]) return;
    var ps = PRESETS[presetKey].settings;

    // Update checkboxes — completeness
    var cplBoxes = document.querySelectorAll('input[data-group="cpl"]');
    for (var i = 0; i < cplBoxes.length; i++) {
      cplBoxes[i].checked = ps.completenessFields.indexOf(cplBoxes[i].getAttribute('data-key')) >= 0;
    }

    // Update checkboxes — quality issues
    var qiBoxes = document.querySelectorAll('input[data-group="qi"]');
    for (var i = 0; i < qiBoxes.length; i++) {
      qiBoxes[i].checked = ps.qualityIssueFields.indexOf(qiBoxes[i].getAttribute('data-key')) >= 0;
    }

    // Update DQ weight sliders
    setSliderValue('dq', 'completeness', ps.dqWeights.completeness);
    setSliderValue('dq', 'udef', ps.dqWeights.udef);
    setSliderValue('dq', 'quality', ps.dqWeights.quality);
    updateSliderGroupTotal('dq');

    // Update engagement weight sliders
    setSliderValue('eng', 'withActivity', ps.engagementWeights.withActivity);
    setSliderValue('eng', 'withPerson', ps.engagementWeights.withPerson);
    setSliderValue('eng', 'withPipeline', ps.engagementWeights.withPipeline);
    updateSliderGroupTotal('eng');

    // Update pipeline radio
    var radios = document.querySelectorAll('input[name="pipelineType"]');
    for (var i = 0; i < radios.length; i++) {
      radios[i].checked = radios[i].value === ps.pipelineType;
    }

    updatePresetDesc();
  }

  function setSliderValue(group, key, val) {
    var slider = document.getElementById('ws_' + group + '_' + key);
    if (slider) {
      slider.value = val;
      updateSliderTrack(slider);
    }
    var valEl = document.getElementById('wsv_' + group + '_' + key);
    if (valEl) valEl.textContent = val + '%';
  }

  function updateSliderGroupTotal(group) {
    var allSliders = document.querySelectorAll('.weight-slider[data-group="' + group + '"]');
    var total = 0;
    for (var i = 0; i < allSliders.length; i++) {
      total += parseInt(allSliders[i].value);
    }
    var totalEl = document.getElementById('wst_' + group);
    if (totalEl) {
      totalEl.textContent = 'Total: ' + total + '%';
      totalEl.className = total === 100 ? 'weight-slider-total' : 'weight-slider-total weight-slider-total-warn';
    }
  }

  function isPresetActive(presetKey) {
    var s = getSettings('company');
    var ps = PRESETS[presetKey].settings;
    // Compare key fields
    if (s.pipelineType !== ps.pipelineType) return false;
    if (JSON.stringify(s.dqWeights) !== JSON.stringify(ps.dqWeights)) return false;
    if (JSON.stringify(s.engagementWeights) !== JSON.stringify(ps.engagementWeights)) return false;
    if (JSON.stringify(s.completenessFields.sort()) !== JSON.stringify(ps.completenessFields.slice().sort())) return false;
    return true;
  }

  function updatePresetDesc() {
    var el = document.getElementById('presetDesc');
    var sel = document.getElementById('settingsPreset');
    if (!el || !sel) return;
    var pk = sel.value;
    el.textContent = pk && PRESETS[pk] ? PRESETS[pk].desc : 'You have custom settings';
  }

  // ============================================================
  // SAVE — collect panel state into settings + persist
  // ============================================================

  function save() {
    // Read completeness fields
    var cplFields = [];
    var cplBoxes = document.querySelectorAll('input[data-group="cpl"]');
    for (var i = 0; i < cplBoxes.length; i++) {
      if (cplBoxes[i].checked) cplFields.push(cplBoxes[i].getAttribute('data-key'));
    }

    // Read quality issue fields
    var qiFields = [];
    var qiBoxes = document.querySelectorAll('input[data-group="qi"]');
    for (var i = 0; i < qiBoxes.length; i++) {
      if (qiBoxes[i].checked) qiFields.push(qiBoxes[i].getAttribute('data-key'));
    }

    // Read DQ weights
    var dqW = readSliderGroup('dq');
    var engW = readSliderGroup('eng');

    // Read pipeline type
    var pipelineType = 'sale';
    var radios = document.querySelectorAll('input[name="pipelineType"]');
    for (var i = 0; i < radios.length; i++) {
      if (radios[i].checked) { pipelineType = radios[i].value; break; }
    }

    // Validate weights sum to 100
    var dqTotal = (dqW.completeness || 0) + (dqW.udef || 0) + (dqW.quality || 0);
    var engTotal = (engW.withActivity || 0) + (engW.withPerson || 0) + (engW.withPipeline || 0);

    if (dqTotal !== 100 || engTotal !== 100) {
      var msg = 'Weight sliders must total 100%:';
      if (dqTotal !== 100) msg += '\n• DQ Weights: ' + dqTotal + '%';
      if (engTotal !== 100) msg += '\n• Engagement Weights: ' + engTotal + '%';
      alert(msg);
      return;
    }

    // Apply
    _settings.company = {
      completenessFields: cplFields,
      dqWeights: { completeness: dqW.completeness, udef: dqW.udef, quality: dqW.quality },
      qualityIssueFields: qiFields,
      engagementWeights: { withActivity: engW.withActivity, withPerson: engW.withPerson, withPipeline: engW.withPipeline },
      pipelineType: pipelineType
    };

    saveSettings();

    // Recalculate — re-render DQ score and adoption for company
    if (typeof renderDQScore === 'function') renderDQScore('company');
    if (typeof renderCrossEntityFunnel === 'function' && typeof companyDetailData !== 'undefined' && companyDetailData) {
      renderCrossEntityFunnel(companyDetailData);
    }
    if (typeof renderCompanyDetails === 'function' && typeof companyDetailData !== 'undefined' && companyDetailData) {
      renderCompanyDetails(companyDetailData);
    }

    // Show confirmation
    showSettingsToast('Settings applied and scores recalculated');
  }

  function readSliderGroup(group) {
    var result = {};
    var sliders = document.querySelectorAll('.weight-slider[data-group="' + group + '"]');
    for (var i = 0; i < sliders.length; i++) {
      result[sliders[i].getAttribute('data-key')] = parseInt(sliders[i].value);
    }
    return result;
  }

  function resetToDefaults() {
    _settings = deepClone(DEFAULTS);
    saveSettings();
    renderSettingsPanel();
    showSettingsToast('Settings reset to defaults');
  }

  function showSettingsToast(msg) {
    var existing = document.getElementById('settingsToast');
    if (existing) existing.remove();

    var toast = document.createElement('div');
    toast.id = 'settingsToast';
    toast.className = 'settings-toast';
    toast.textContent = msg;
    document.body.appendChild(toast);

    setTimeout(function() { toast.classList.add('show'); }, 10);
    setTimeout(function() {
      toast.classList.remove('show');
      setTimeout(function() { toast.remove(); }, 300);
    }, 2500);
  }

  // ============================================================
  // INIT — called when settings tab is shown
  // ============================================================

  function init() {
    loadSettings();
    renderSettingsPanel();
  }

  // ============================================================
  // PUBLIC API
  // ============================================================

  window.daSettings = {
    init: init,
    getSettings: getSettings,
    computeDQScore: computeConfiguredDQScore,
    computeEngagement: computeEngagement,
    getCompletenessValue: getCompletenessValue,
    getPipelineLabel: getPipelineLabel,
    getPipelineType: getPipelineType,
    COMPLETENESS_OPTIONS: COMPLETENESS_OPTIONS,
    QUALITY_ISSUE_OPTIONS: QUALITY_ISSUE_OPTIONS,

    // Panel event handlers (called from onclick)
    onSliderChange: onSliderChange,
    onFieldToggle: onFieldToggle,
    onPipelineChange: onPipelineChange,
    applyPreset: applyPreset,
    save: save,
    resetToDefaults: resetToDefaults
  };

  // Auto-load settings on script load
  loadSettings();

})();
