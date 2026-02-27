// da-settings.js v3 — Settings Manager
// Unified field importance for standard + UDEF fields
// Collapsible UDEF section, better toggle UX

(function() {
  'use strict';

  // ============================================================
  // FIELD & OPTION DEFINITIONS
  // ============================================================
  var STD_FIELDS = [
    { key: 'email',    label: 'Email address' },
    { key: 'phone',    label: 'Phone number' },
    { key: 'person',   label: 'Contact person' },
    { key: 'category', label: 'Category' },
    { key: 'orgNr',    label: 'Org. number' },
    { key: 'business', label: 'Business type' },
    { key: 'address',  label: 'Address' },
    { key: 'webpage',  label: 'Webpage' }
  ];

  var QUALITY_ISSUES = [
    { key: 'noPerson',    label: 'No contact person' },
    { key: 'noCategory',  label: 'No category' },
    { key: 'noBusiness',  label: 'No business type' },
    { key: 'unreachable', label: 'Unreachable (no email or phone)' },
    { key: 'noOrgNr',     label: 'No org. number' }
  ];

  var PIPE_OPTS = [
    { key: 'sale', label: 'Open Sale',           desc: 'Companies with an open sale opportunity' },
    { key: 'project', label: 'Active Project',   desc: 'Companies with an active project', soon: true },
    { key: 'both', label: 'Sale or Project',     desc: 'Companies with either sale or project', soon: true },
    { key: 'none', label: 'None (activity only)', desc: 'Funnel stops at activity — for relationship-only CRM' }
  ];

  var PRESETS = {
    standard: { label: 'Standard', desc: 'Balanced scoring for general CRM usage',
      s: { std: { email:'required', phone:'required', person:'required', category:'required', orgNr:'normal', business:'excluded', address:'excluded', webpage:'excluded' }, dqW: {c:40,u:30,q:30}, qi: ['noPerson','noCategory','noBusiness','unreachable'], engW: {a:50,p:30,s:20}, pipe:'sale' }},
    salesFocus: { label: 'Sales Focus', desc: 'Emphasizes pipeline and commercial data quality',
      s: { std: { email:'required', phone:'required', person:'required', category:'required', orgNr:'required', business:'normal', address:'excluded', webpage:'excluded' }, dqW: {c:30,u:20,q:50}, qi: ['noPerson','noCategory','unreachable'], engW: {a:35,p:25,s:40}, pipe:'sale' }},
    serviceFocus: { label: 'Service Focus', desc: 'Emphasizes reachability and contact data',
      s: { std: { email:'required', phone:'required', person:'required', category:'normal', orgNr:'excluded', business:'excluded', address:'normal', webpage:'excluded' }, dqW: {c:50,u:20,q:30}, qi: ['noPerson','unreachable'], engW: {a:60,p:30,s:10}, pipe:'none' }},
    projectBased: { label: 'Project Based', desc: 'For organizations using projects, not sales',
      s: { std: { email:'required', phone:'normal', person:'required', category:'required', orgNr:'normal', business:'excluded', address:'excluded', webpage:'excluded' }, dqW: {c:40,u:30,q:30}, qi: ['noPerson','noCategory','noBusiness'], engW: {a:50,p:30,s:20}, pipe:'project' }}
  };

  // Defaults
  var DEF = {
    stdFieldConfig: { email:'required', phone:'required', person:'required', category:'required', orgNr:'normal', business:'excluded', address:'excluded', webpage:'excluded' },
    udefFieldConfig: {},
    dqWeights: { completeness:40, udef:30, quality:30 },
    qualityIssueFields: ['noPerson','noCategory','noBusiness','unreachable'],
    engagementWeights: { withActivity:50, withPerson:30, withPipeline:20 },
    pipelineType: 'sale'
  };

  // ============================================================
  // STATE
  // ============================================================
  var SK = 'da_settings_v2';
  var _s = null;
  var _udefFields = [];
  var _activeId = null;

  function clone(o) { return JSON.parse(JSON.stringify(o)); }

  function load() {
    try {
      var raw = localStorage.getItem(SK);
      if (raw) { _s = migrate(JSON.parse(raw)); return; }
      // Try v1 migration
      var v1 = localStorage.getItem('da_settings_v1');
      if (v1) { _s = migrateV1(JSON.parse(v1)); save(); return; }
    } catch(e) {}
    _s = clone(DEF);
  }

  function migrate(saved) {
    var r = clone(DEF);
    if (saved.stdFieldConfig) r.stdFieldConfig = clone(saved.stdFieldConfig);
    if (saved.udefFieldConfig) r.udefFieldConfig = clone(saved.udefFieldConfig);
    if (saved.dqWeights) r.dqWeights = clone(saved.dqWeights);
    if (saved.qualityIssueFields) r.qualityIssueFields = saved.qualityIssueFields.slice();
    if (saved.engagementWeights) r.engagementWeights = clone(saved.engagementWeights);
    if (saved.pipelineType) r.pipelineType = saved.pipelineType;
    return r;
  }

  function migrateV1(v1) {
    var r = clone(DEF);
    if (v1.company) {
      var c = v1.company;
      // Convert completenessFields array → stdFieldConfig
      if (c.completenessFields) {
        for (var i = 0; i < STD_FIELDS.length; i++) {
          var k = STD_FIELDS[i].key;
          r.stdFieldConfig[k] = c.completenessFields.indexOf(k) >= 0 ? 'normal' : 'excluded';
        }
      }
      if (c.udefFieldConfig) r.udefFieldConfig = clone(c.udefFieldConfig);
      if (c.dqWeights) r.dqWeights = clone(c.dqWeights);
      if (c.qualityIssueFields) r.qualityIssueFields = c.qualityIssueFields.slice();
      if (c.engagementWeights) r.engagementWeights = clone(c.engagementWeights);
      if (c.pipelineType) r.pipelineType = c.pipelineType;
    }
    return r;
  }

  function save() {
    try { localStorage.setItem(SK, JSON.stringify(_s)); } catch(e) {}
  }

  function get() { if (!_s) load(); return _s; }

  // Derive completenessFields for backward compat with entity-logic
  function getCompleteness() {
    var s = get();
    var fields = [];
    for (var i = 0; i < STD_FIELDS.length; i++) {
      var imp = s.stdFieldConfig[STD_FIELDS[i].key] || 'excluded';
      if (imp !== 'excluded') fields.push(STD_FIELDS[i].key);
    }
    return fields;
  }

  // Public getSettings (backward compat)
  function getSettings(entity) {
    var s = get();
    return {
      completenessFields: getCompleteness(),
      stdFieldConfig: s.stdFieldConfig,
      dqWeights: s.dqWeights,
      qualityIssueFields: s.qualityIssueFields,
      udefFieldConfig: s.udefFieldConfig,
      engagementWeights: s.engagementWeights,
      pipelineType: s.pipelineType
    };
  }

  // ============================================================
  // UDEF DISCOVERY
  // ============================================================
  function notifyUdefLoaded(entityId, data) {
    if (entityId !== 7 || !data || !data.fields) return;
    _udefFields = [];
    for (var i = 0; i < data.fields.length; i++) {
      var f = data.fields[i];
      _udefFields.push({ progId: f.progId || f.label || ('f' + i), label: f.label, type: f.type, percent: f.percent });
    }
    if (_activeId) {
      var el = document.getElementById(_activeId);
      if (el && el.offsetParent !== null) renderPanel(el, el.getAttribute('data-mode') || 'full');
    }
  }

  // ============================================================
  // COMPUTE HELPERS
  // ============================================================
  function getCompletenessValue(key, ovCpl, qData, total) {
    if (!total) return 0;
    if (ovCpl && ovCpl[key] !== undefined) return ovCpl[key];
    var qm = { person:'noPerson', category:'noCategory', business:'noBusiness' };
    if (qm[key] && qData && qData[qm[key]] !== undefined) return total - qData[qm[key]];
    return 0;
  }

  function computeDQ(entity, ovCpl, qData, uData, total) {
    var s = get();
    var scores = {};

    // Completeness: standard fields
    var cplFields = getCompleteness();
    if (ovCpl && cplFields.length > 0 && total > 0) {
      var sum = 0; var wt = 0;
      for (var i = 0; i < cplFields.length; i++) {
        var k = cplFields[i];
        var imp = s.stdFieldConfig[k] || 'normal';
        var w = imp === 'required' ? 2 : 1;
        sum += (getCompletenessValue(k, ovCpl, qData, total) / total * 100) * w;
        wt += w;
      }
      if (wt > 0) scores.completeness = sum / wt;
    }

    // UDEF
    if (uData && uData.fields && uData.fields.length > 0) {
      var cfg = s.udefFieldConfig || {};
      var uS = 0; var uW = 0;
      for (var i = 0; i < uData.fields.length; i++) {
        var f = uData.fields[i];
        var pid = f.progId || f.label || ('f' + i);
        var imp = cfg[pid] || 'normal';
        if (imp === 'excluded') continue;
        var w = imp === 'required' ? 2 : 1;
        uS += f.percent * w; uW += w;
      }
      if (uW > 0) scores.udef = uS / uW;
    }

    // Quality
    if (qData && s.qualityIssueFields.length > 0 && total > 0) {
      var issS = 0;
      for (var i = 0; i < s.qualityIssueFields.length; i++) issS += (qData[s.qualityIssueFields[i]] || 0);
      scores.quality = 100 - (issS / (total * s.qualityIssueFields.length) * 100);
    }

    var dw = s.dqWeights; var tW = 0; var wS = 0;
    if (scores.completeness !== undefined) { wS += scores.completeness * dw.completeness; tW += dw.completeness; }
    if (scores.udef !== undefined) { wS += scores.udef * dw.udef; tW += dw.udef; }
    if (scores.quality !== undefined) { wS += scores.quality * dw.quality; tW += dw.quality; }
    if (tW <= 0) return null;
    return { total: Math.round(wS / tW), components: scores, weights: dw };
  }

  function computeEngagement(row, total, entity) {
    if (!total) return 0;
    var s = get();
    var w = s.engagementWeights;
    var pV = 0;
    if (s.pipelineType === 'sale') pV = row.withSale || 0;
    else if (s.pipelineType === 'project') pV = row.withProject || 0;
    else if (s.pipelineType === 'both') pV = row.withSaleOrProject || 0;
    return Math.round(((row.withActivity || 0) * w.withActivity / 100 + (row.withPerson || 0) * w.withPerson / 100 + pV * w.withPipeline / 100) / total * 100);
  }

  function getPipelineLabel() {
    var s = get();
    for (var i = 0; i < PIPE_OPTS.length; i++) if (PIPE_OPTS[i].key === s.pipelineType) return PIPE_OPTS[i].label;
    return 'Open Sale';
  }

  function getPipelineType() { return get().pipelineType; }

  // ============================================================
  // RENDER — 'full' (settings page) or 'dq' (minimal)
  // ============================================================
  function renderPanel(el, mode) {
    var s = get();
    var h = '';

    // Preset
    h += '<div class="s-preset-bar">';
    h += '<span class="s-preset-lbl">Preset:</span>';
    h += '<select class="s-preset-sel" id="sPreset" onchange="daSettings.applyPreset(this.value)">';
    h += '<option value="">Custom</option>';
    for (var pk in PRESETS) h += '<option value="' + pk + '"' + (isPresetMatch(pk) ? ' selected' : '') + '>' + PRESETS[pk].label + '</option>';
    h += '</select>';
    h += '<span class="s-preset-desc" id="sPresetDesc"></span>';
    h += '</div>';

    // DATA QUALITY
    h += '<div class="s-section">';
    h += '<div class="s-section-head">Data Quality</div>';

    // Row 1: DQ Weights + Quality Issues
    h += '<div class="s-grid">';
    h += '<div class="settings-card">';
    h += '<h3>DQ Score Weights</h3>';
    h += '<div class="settings-desc">Component weights for the overall DQ Score.</div>';
    h += sliders('dq', [
      { k:'completeness', l:'Field Completeness', v:s.dqWeights.completeness },
      { k:'udef', l:'Custom Fields (UDEF)', v:s.dqWeights.udef },
      { k:'quality', l:'Quality Issues', v:s.dqWeights.quality }
    ]);
    h += '</div>';
    h += '<div class="settings-card">';
    h += '<h3>Quality Issue Checks</h3>';
    h += '<div class="settings-desc">Which issues count in the DQ score.</div>';
    h += '<div class="settings-field-grid">';
    for (var i = 0; i < QUALITY_ISSUES.length; i++) {
      var qi = QUALITY_ISSUES[i];
      h += '<label class="settings-checkbox"><input type="checkbox" data-group="qi" data-key="' + qi.key + '"' + (s.qualityIssueFields.indexOf(qi.key) >= 0 ? ' checked' : '') + ' onchange="daSettings.mc()"> ' + qi.label + '</label>';
    }
    h += '</div></div>';
    h += '</div>'; // grid

    // Completeness Definition
    h += '<div class="s-cpl-card">';
    h += '<h3>Completeness Definition</h3>';
    h += '<div class="settings-desc">Which fields count towards completeness. <strong>Required</strong> fields count double.</div>';

    // Standard fields table
    h += '<div class="s-cpl-sub-head">Standard Fields</div>';
    h += '<div class="s-field-list">';
    for (var i = 0; i < STD_FIELDS.length; i++) {
      var sf = STD_FIELDS[i];
      var imp = s.stdFieldConfig[sf.key] || 'excluded';
      h += fieldRow(sf.label, null, null, imp, 'std', sf.key);
    }
    h += '</div>';

    // UDEF fields (collapsible)
    h += '<div class="s-cpl-sub-head s-cpl-udef-head" onclick="daSettings.toggleUdef()">';
    h += '<span class="s-udef-chev" id="sUdefChev">&#9654;</span> Custom Fields (UDEF)';
    if (_udefFields.length > 0) h += ' <span class="s-udef-count">' + _udefFields.length + ' fields</span>';
    h += '</div>';
    h += '<div class="s-udef-body" id="sUdefBody" style="display:none">';
    if (_udefFields.length === 0) {
      h += '<div class="s-udef-empty">Run the Company analysis first to discover custom fields.</div>';
    } else {
      h += '<div class="s-field-list">';
      for (var i = 0; i < _udefFields.length; i++) {
        var uf = _udefFields[i];
        var imp = s.udefFieldConfig[uf.progId] || 'normal';
        var fillCol = uf.percent >= 70 ? 'var(--sl-good)' : (uf.percent >= 30 ? 'var(--sl-ok)' : 'var(--sl-bad)');
        h += fieldRow(uf.label, uf.type, '<span style="font-weight:600;color:' + fillCol + '">' + uf.percent + '%</span>', imp, 'udef', uf.progId);
      }
      h += '</div>';
    }
    h += '</div>'; // udef body

    h += '</div>'; // cpl card
    h += '</div>'; // section

    // ADOPTION (full only)
    if (mode === 'full') {
      h += '<div class="s-section">';
      h += '<div class="s-section-head">Adoption</div>';
      h += '<div class="s-grid">';

      h += '<div class="settings-card">';
      h += '<h3>Engagement Score Weights</h3>';
      h += '<div class="settings-desc">Category Effectiveness engagement calculation.</div>';
      h += sliders('eng', [
        { k:'withActivity', l:'Activity', v:s.engagementWeights.withActivity },
        { k:'withPerson', l:'Contact Person', v:s.engagementWeights.withPerson },
        { k:'withPipeline', l:'Pipeline', v:s.engagementWeights.withPipeline }
      ]);
      h += '</div>';

      h += '<div class="settings-card">';
      h += '<h3>Pipeline Definition</h3>';
      h += '<div class="settings-desc">What counts as "pipeline" in the funnel.</div>';
      for (var i = 0; i < PIPE_OPTS.length; i++) {
        var po = PIPE_OPTS[i];
        h += '<label class="settings-radio' + (po.soon ? ' settings-radio-soon' : '') + '">';
        h += '<input type="radio" name="pipelineType" value="' + po.key + '"' + (s.pipelineType === po.key ? ' checked' : '') + (po.soon ? ' disabled' : '') + ' onchange="daSettings.mc()">';
        h += '<span><strong>' + po.label + '</strong><span class="settings-radio-desc">' + po.desc + '</span>';
        if (po.soon) h += '<span class="settings-badge-soon">Coming soon</span>';
        h += '</span></label>';
      }
      h += '</div>';
      h += '</div></div>'; // grid + section
    }

    // Actions
    h += '<div class="s-actions">';
    h += '<button class="btn-settings-save" onclick="daSettings.doSave()">Apply &amp; Recalculate</button>';
    h += '<button class="btn-settings-reset" onclick="daSettings.doReset()">Reset to Defaults</button>';
    h += '</div>';

    el.innerHTML = h;
    el.setAttribute('data-mode', mode);
    _activeId = el.id;
    updPresetDesc();
  }

  // Field row with 3-way toggle
  function fieldRow(label, type, extra, imp, group, key) {
    var esc = key.replace(/'/g, "\\'");
    var h = '<div class="s-fld-row' + (imp === 'excluded' ? ' s-fld-excl' : '') + '" data-grp="' + group + '" data-key="' + key + '">';
    h += '<div class="s-fld-name">' + label;
    if (type) h += '<span class="s-fld-type">' + type + '</span>';
    h += '</div>';
    if (extra) h += '<div class="s-fld-extra">' + extra + '</div>';
    h += '<div class="s-fld-toggle">';
    h += '<button class="s-imp-btn' + (imp === 'required' ? ' act-req' : '') + '" onclick="daSettings.setImp(\'' + group + '\',\'' + esc + '\',\'required\',this)">Required</button>';
    h += '<button class="s-imp-btn' + (imp === 'normal' ? ' act-norm' : '') + '" onclick="daSettings.setImp(\'' + group + '\',\'' + esc + '\',\'normal\',this)">Normal</button>';
    h += '<button class="s-imp-btn' + (imp === 'excluded' ? ' act-excl' : '') + '" onclick="daSettings.setImp(\'' + group + '\',\'' + esc + '\',\'excluded\',this)">Excluded</button>';
    h += '</div></div>';
    return h;
  }

  // ============================================================
  // SLIDERS
  // ============================================================
  function sliders(gid, items) {
    var h = '<div class="weight-slider-group">';
    for (var i = 0; i < items.length; i++) {
      h += '<div class="weight-slider-row">';
      h += '<label class="weight-slider-label">' + items[i].l + '</label>';
      h += '<input type="range" class="weight-slider" id="ws_' + gid + '_' + items[i].k + '" data-group="' + gid + '" data-key="' + items[i].k + '" min="0" max="100" step="5" value="' + items[i].v + '" oninput="daSettings.onSl(this)">';
      h += '<span class="weight-slider-value" id="wsv_' + gid + '_' + items[i].k + '">' + items[i].v + '%</span>';
      h += '</div>';
    }
    h += '<div class="weight-slider-total" id="wst_' + gid + '">Total: 100%</div></div>';
    return h;
  }

  function onSl(el) {
    var g = el.getAttribute('data-group'), k = el.getAttribute('data-key');
    var ve = document.getElementById('wsv_' + g + '_' + k);
    if (ve) ve.textContent = parseInt(el.value) + '%';
    updSlTotal(g); mc();
  }

  function updSlTotal(g) {
    var all = document.querySelectorAll('.weight-slider[data-group="' + g + '"]');
    var t = 0; for (var i = 0; i < all.length; i++) t += parseInt(all[i].value);
    var te = document.getElementById('wst_' + g);
    if (te) { te.textContent = 'Total: ' + t + '%'; te.className = t === 100 ? 'weight-slider-total' : 'weight-slider-total weight-slider-total-warn'; }
  }

  function setSl(g, k, v) {
    var s = document.getElementById('ws_' + g + '_' + k); if (s) s.value = v;
    var ve = document.getElementById('wsv_' + g + '_' + k); if (ve) ve.textContent = v + '%';
  }

  // ============================================================
  // EVENTS
  // ============================================================
  function mc() { var sel = document.getElementById('sPreset'); if (sel) sel.value = ''; updPresetDesc(); }

  function setImp(group, key, imp, btn) {
    var btns = btn.parentElement.querySelectorAll('.s-imp-btn');
    for (var i = 0; i < btns.length; i++) btns[i].className = 's-imp-btn';
    btn.className = 's-imp-btn act-' + (imp === 'required' ? 'req' : (imp === 'normal' ? 'norm' : 'excl'));
    var row = btn.closest('.s-fld-row');
    if (row) { if (imp === 'excluded') row.classList.add('s-fld-excl'); else row.classList.remove('s-fld-excl'); }
    mc();
  }

  function toggleUdef() {
    var body = document.getElementById('sUdefBody');
    var chev = document.getElementById('sUdefChev');
    if (!body) return;
    var open = body.style.display !== 'none';
    body.style.display = open ? 'none' : 'block';
    if (chev) chev.innerHTML = open ? '&#9654;' : '&#9660;';
  }

  function applyPreset(pk) {
    if (!pk || !PRESETS[pk]) return;
    var ps = PRESETS[pk].s;
    // Standard fields
    var rows = document.querySelectorAll('.s-fld-row[data-grp="std"]');
    for (var i = 0; i < rows.length; i++) {
      var k = rows[i].getAttribute('data-key');
      var imp = ps.std[k] || 'excluded';
      var btns = rows[i].querySelectorAll('.s-imp-btn');
      for (var j = 0; j < btns.length; j++) btns[j].className = 's-imp-btn';
      if (imp === 'required') btns[0].className = 's-imp-btn act-req';
      else if (imp === 'normal') btns[1].className = 's-imp-btn act-norm';
      else btns[2].className = 's-imp-btn act-excl';
      if (imp === 'excluded') rows[i].classList.add('s-fld-excl'); else rows[i].classList.remove('s-fld-excl');
    }
    // UDEF reset to normal
    var uRows = document.querySelectorAll('.s-fld-row[data-grp="udef"]');
    for (var i = 0; i < uRows.length; i++) {
      var btns = uRows[i].querySelectorAll('.s-imp-btn');
      for (var j = 0; j < btns.length; j++) btns[j].className = 's-imp-btn';
      btns[1].className = 's-imp-btn act-norm';
      uRows[i].classList.remove('s-fld-excl');
    }
    // DQ sliders
    setSl('dq','completeness',ps.dqW.c); setSl('dq','udef',ps.dqW.u); setSl('dq','quality',ps.dqW.q); updSlTotal('dq');
    // Engagement
    setSl('eng','withActivity',ps.engW.a); setSl('eng','withPerson',ps.engW.p); setSl('eng','withPipeline',ps.engW.s); updSlTotal('eng');
    // Quality issues
    var qiBoxes = document.querySelectorAll('input[data-group="qi"]');
    for (var i = 0; i < qiBoxes.length; i++) qiBoxes[i].checked = ps.qi.indexOf(qiBoxes[i].getAttribute('data-key')) >= 0;
    // Pipeline
    var radios = document.querySelectorAll('input[name="pipelineType"]');
    for (var i = 0; i < radios.length; i++) radios[i].checked = radios[i].value === ps.pipe;
    updPresetDesc();
  }

  function isPresetMatch(pk) {
    var s = get(); var ps = PRESETS[pk].s;
    if (s.pipelineType !== ps.pipe) return false;
    if (s.dqWeights.completeness !== ps.dqW.c || s.dqWeights.udef !== ps.dqW.u || s.dqWeights.quality !== ps.dqW.q) return false;
    for (var k in ps.std) { if ((s.stdFieldConfig[k] || 'excluded') !== ps.std[k]) return false; }
    return true;
  }

  function updPresetDesc() {
    var el = document.getElementById('sPresetDesc');
    var sel = document.getElementById('sPreset');
    if (!el || !sel) return;
    el.textContent = sel.value && PRESETS[sel.value] ? PRESETS[sel.value].desc : 'You have custom settings';
  }

  // ============================================================
  // SAVE
  // ============================================================
  function doSave() {
    // Read standard fields
    var stdCfg = {};
    var stdRows = document.querySelectorAll('.s-fld-row[data-grp="std"]');
    for (var i = 0; i < stdRows.length; i++) {
      var k = stdRows[i].getAttribute('data-key');
      var act = stdRows[i].querySelector('.s-imp-btn.act-req,.s-imp-btn.act-norm,.s-imp-btn.act-excl');
      if (act) {
        if (act.classList.contains('act-req')) stdCfg[k] = 'required';
        else if (act.classList.contains('act-norm')) stdCfg[k] = 'normal';
        else stdCfg[k] = 'excluded';
      }
    }
    // UDEF
    var uCfg = {};
    var uRows = document.querySelectorAll('.s-fld-row[data-grp="udef"]');
    for (var i = 0; i < uRows.length; i++) {
      var k = uRows[i].getAttribute('data-key');
      var act = uRows[i].querySelector('.s-imp-btn.act-req,.s-imp-btn.act-norm,.s-imp-btn.act-excl');
      if (act) {
        if (act.classList.contains('act-req')) uCfg[k] = 'required';
        else if (act.classList.contains('act-excl')) uCfg[k] = 'excluded';
        // normal = default, don't store
      }
    }
    // Quality issues
    var qiF = [];
    var qiB = document.querySelectorAll('input[data-group="qi"]');
    for (var i = 0; i < qiB.length; i++) if (qiB[i].checked) qiF.push(qiB[i].getAttribute('data-key'));
    // Sliders
    var dqW = readSl('dq'); var engW = readSl('eng');
    // Pipeline
    var pipe = 'sale';
    var radios = document.querySelectorAll('input[name="pipelineType"]');
    for (var i = 0; i < radios.length; i++) if (radios[i].checked) { pipe = radios[i].value; break; }
    // Validate
    var dqT = (dqW.completeness||0) + (dqW.udef||0) + (dqW.quality||0);
    if (dqT !== 100) { alert('DQ weights must total 100% (currently ' + dqT + '%)'); return; }
    if (engW.withActivity !== undefined) {
      var engT = (engW.withActivity||0) + (engW.withPerson||0) + (engW.withPipeline||0);
      if (engT !== 100) { alert('Engagement weights must total 100% (currently ' + engT + '%)'); return; }
    }
    // Apply
    _s = {
      stdFieldConfig: stdCfg, udefFieldConfig: uCfg, dqWeights: { completeness: dqW.completeness||40, udef: dqW.udef||30, quality: dqW.quality||30 },
      qualityIssueFields: qiF,
      engagementWeights: engW.withActivity !== undefined ? { withActivity: engW.withActivity, withPerson: engW.withPerson, withPipeline: engW.withPipeline } : get().engagementWeights,
      pipelineType: pipe
    };
    save();
    if (typeof renderDQScore === 'function') renderDQScore('company');
    if (typeof renderCrossEntityFunnel === 'function' && typeof companyDetailData !== 'undefined' && companyDetailData) renderCrossEntityFunnel(companyDetailData);
    if (typeof renderCompanyDetails === 'function' && typeof companyDetailData !== 'undefined' && companyDetailData) renderCompanyDetails(companyDetailData);
    toast('Settings applied and scores recalculated');
  }

  function readSl(g) {
    var r = {}; var sl = document.querySelectorAll('.weight-slider[data-group="' + g + '"]');
    for (var i = 0; i < sl.length; i++) r[sl[i].getAttribute('data-key')] = parseInt(sl[i].value);
    return r;
  }

  function doReset() {
    _s = clone(DEF); save();
    if (_activeId) { var el = document.getElementById(_activeId); if (el) renderPanel(el, el.getAttribute('data-mode') || 'full'); }
    toast('Settings reset to defaults');
  }

  function toast(msg) {
    var ex = document.getElementById('sToast'); if (ex) ex.remove();
    var t = document.createElement('div'); t.id = 'sToast'; t.className = 'settings-toast'; t.textContent = msg;
    document.body.appendChild(t);
    setTimeout(function(){ t.classList.add('show'); }, 10);
    setTimeout(function(){ t.classList.remove('show'); setTimeout(function(){ t.remove(); }, 300); }, 2500);
  }

  // ============================================================
  // INIT
  // ============================================================
  function init() { load(); var p = document.getElementById('settings-quality'); if (p) renderPanel(p, 'full'); }

  window.daSettings = {
    init: init,
    getSettings: getSettings,
    computeDQScore: computeDQ,
    computeEngagement: computeEngagement,
    getCompletenessValue: getCompletenessValue,
    getPipelineLabel: getPipelineLabel,
    getPipelineType: getPipelineType,
    notifyUdefLoaded: notifyUdefLoaded,
    COMPLETENESS_OPTIONS: STD_FIELDS.map(function(f){ return { key: f.key, label: f.label }; }),
    QUALITY_ISSUE_OPTIONS: QUALITY_ISSUES,
    onSl: onSl, mc: mc, setImp: setImp, toggleUdef: toggleUdef,
    applyPreset: applyPreset, doSave: doSave, doReset: doReset
  };

  load();
})();
