// da-settings.js v4 — Multi-Entity Settings Manager
// 3 Scores (Data Quality, Data Integrity, Adoption) + Overall Health
// H/M/L weight system, per-entity configuration

(function() {
  'use strict';

  var POINTS = { high: 3, medium: 2, low: 1 };
  var _mode = 'full';

  // ============================================================
  // ENTITY DEFINITIONS — fields, checks, engagement per entity
  // ============================================================
  var ENTITY_DEFS = {
    company: {
      label: 'Company', entityId: 7,
      stdFields: [
        { key: 'email',    label: 'Email address' },
        { key: 'phone',    label: 'Phone number' },
        { key: 'category', label: 'Category' },
        { key: 'business', label: 'Business type' },
        { key: 'orgNr',    label: 'Org. number' },
        { key: 'address',  label: 'Address' },
        { key: 'postcode', label: 'Postcode' },
        { key: 'country',  label: 'Country' },
        { key: 'webpage',  label: 'Webpage' }
      ],
      integrityChecks: [
        { key: 'noPerson',      label: 'No contact person',  desc: 'Company has zero linked persons' },
        { key: 'unreachable',   label: 'Unreachable company', desc: 'No person has an email or phone number' },
        { key: 'noActivity12m', label: 'No recent activity',  desc: 'No activities logged in the last 12 months', isNew: true },
        { key: 'noOwner',       label: 'No owner',            desc: 'Company has no assigned associate/owner', isNew: true }
      ],
      engagementComponents: [
        { key: 'withPerson',   label: 'Persons',    alwaysOn: false },
        { key: 'withActivity', label: 'Activities',  alwaysOn: false },
        { key: 'withPipeline', label: 'Pipeline',    alwaysOn: false }
      ],
      hasPipelineConfig: true
    },
    contact: {
      label: 'Contact', entityId: 6,
      stdFields: [
        { key: 'firstName', label: 'First name' },
        { key: 'lastName',  label: 'Last name' },
        { key: 'email',     label: 'Email' },
        { key: 'phone',     label: 'Phone' },
        { key: 'position',  label: 'Position/Title' },
        { key: 'mrMrs',     label: 'Mr/Ms' }
      ],
      integrityChecks: [
        { key: 'noEmail',     label: 'No email address',     desc: 'Person has no email — unreachable digitally' },
        { key: 'noCompany',   label: 'Not linked to company', desc: 'Person is not associated with any company' },
        { key: 'noActivity',  label: 'No activity',          desc: 'No logged activity in 12 months', isNew: true }
      ],
      engagementComponents: [
        { key: 'withActivity', label: 'Activities',    alwaysOn: false },
        { key: 'withSales',    label: 'Linked Sales',  alwaysOn: false }
      ],
      hasPipelineConfig: false
    },
    sale: {
      label: 'Sale', entityId: 10,
      stdFields: [
        { key: 'amount',      label: 'Amount' },
        { key: 'saleType',    label: 'Sale type' },
        { key: 'stage',       label: 'Stage' },
        { key: 'probability', label: 'Probability' },
        { key: 'closeDate',   label: 'Close date' },
        { key: 'competitor',  label: 'Competitor' },
        { key: 'source',      label: 'Source' }
      ],
      integrityChecks: [
        { key: 'noContact',     label: 'No contact linked',  desc: 'Sale has no associated contact person' },
        { key: 'staleSale',     label: 'Stale sale',         desc: 'Open sale past expected close date' },
        { key: 'noActivities',  label: 'No activities logged', desc: 'Sale has zero linked activities' },
        { key: 'noAmount',      label: 'No amount',          desc: 'Sale amount is zero or empty', isNew: true }
      ],
      engagementComponents: [
        { key: 'withActivity',     label: 'Activities logged',  alwaysOn: false },
        { key: 'stageProgression', label: 'Stage progression',  alwaysOn: false }
      ],
      hasPipelineConfig: false
    },
    project: {
      label: 'Project', entityId: 11,
      stdFields: [
        { key: 'projectType', label: 'Project type' },
        { key: 'status',      label: 'Status' },
        { key: 'endDate',     label: 'End date' },
        { key: 'description', label: 'Description' }
      ],
      integrityChecks: [
        { key: 'noMembers',    label: 'No members',    desc: 'Project has zero members assigned' },
        { key: 'noActivities', label: 'No activities',  desc: 'Project has zero linked activities' }
      ],
      engagementComponents: [
        { key: 'withActivity',     label: 'Activities logged',   alwaysOn: false },
        { key: 'memberEngagement', label: 'Member engagement',   alwaysOn: false }
      ],
      hasPipelineConfig: false
    }
  };

  var ENTITY_ORDER = ['company', 'contact', 'sale', 'project'];

  var PIPE_OPTS = [
    { key: 'sale', label: 'Open Sale', desc: 'Companies with an open sale opportunity' },
    { key: 'none', label: 'None (activity only)', desc: 'Funnel stops at activity — for relationship-only CRM' }
  ];

  // ============================================================
  // DEFAULT SETTINGS GENERATOR (per entity)
  // ============================================================
  function entityDefaults(entityKey) {
    var def = ENTITY_DEFS[entityKey];
    if (!def) return null;
    var stdCfg = {};
    for (var i = 0; i < def.stdFields.length; i++) {
      var f = def.stdFields[i];
      if (i < 3) stdCfg[f.key] = 'required';
      else if (i < def.stdFields.length - 2) stdCfg[f.key] = 'normal';
      else stdCfg[f.key] = 'excluded';
    }
    var intCfg = {};
    for (var i = 0; i < def.integrityChecks.length; i++) {
      var c = def.integrityChecks[i];
      if (i < 2) intCfg[c.key] = { enabled: true, weight: 'high' };
      else if (c.isNew) intCfg[c.key] = { enabled: false, weight: 'low' };
      else intCfg[c.key] = { enabled: true, weight: 'medium' };
    }
    var engCfg = {};
    for (var i = 0; i < def.engagementComponents.length; i++) {
      var e = def.engagementComponents[i];
      engCfg[e.key] = { enabled: true, weight: i < 2 ? 'high' : 'low' };
    }
    return {
      healthWeights: { dq: 'high', integrity: 'medium', adoption: 'high' },
      healthExclude: { integrity: false, adoption: false },
      stdFieldConfig: stdCfg,
      udefFieldConfig: {},
      integrityConfig: intCfg,
      engagementConfig: engCfg,
      pipelineType: 'sale'
    };
  }

  function allDefaults() {
    var d = {};
    for (var i = 0; i < ENTITY_ORDER.length; i++) d[ENTITY_ORDER[i]] = entityDefaults(ENTITY_ORDER[i]);
    return d;
  }

  // ============================================================
  // STATE
  // ============================================================
  var SK = 'da_settings_v3';
  var _s = null;
  var _udefFields = {};
  var _activeId = null;
  var _activeEntity = 'company';

  function clone(o) { return JSON.parse(JSON.stringify(o)); }

  function load() {
    try {
      var raw = localStorage.getItem(SK);
      if (raw) { _s = migrateV3(JSON.parse(raw)); return; }
      var v2 = localStorage.getItem('da_settings_v2');
      if (v2) { _s = migrateFromV2(JSON.parse(v2)); save(); return; }
    } catch(e) { console.warn('Settings load error:', e); }
    _s = allDefaults();
  }

  function migrateV3(saved) {
    var defs = allDefaults();
    for (var ek in defs) {
      if (!saved[ek]) { saved[ek] = defs[ek]; continue; }
      var d = defs[ek]; var s = saved[ek];
      if (!s.healthWeights) s.healthWeights = d.healthWeights;
      if (!s.healthExclude) s.healthExclude = d.healthExclude;
      if (!s.stdFieldConfig) s.stdFieldConfig = d.stdFieldConfig;
      if (!s.udefFieldConfig) s.udefFieldConfig = {};
      if (!s.integrityConfig) s.integrityConfig = d.integrityConfig;
      if (!s.engagementConfig) s.engagementConfig = d.engagementConfig;
      if (!s.pipelineType) s.pipelineType = d.pipelineType;
      for (var ck in d.integrityConfig) { if (!s.integrityConfig[ck]) s.integrityConfig[ck] = d.integrityConfig[ck]; }
      for (var cek in d.engagementConfig) { if (!s.engagementConfig[cek]) s.engagementConfig[cek] = d.engagementConfig[cek]; }
    }
    return saved;
  }

  function migrateFromV2(v2) {
    var all = allDefaults();
    var c = all.company;
    if (v2.stdFieldConfig) { c.stdFieldConfig = {}; for (var k in v2.stdFieldConfig) c.stdFieldConfig[k] = v2.stdFieldConfig[k]; }
    if (v2.udefFieldConfig) c.udefFieldConfig = clone(v2.udefFieldConfig);
    if (v2.qualityIssueFields) {
      var oldMap = { noPerson: 'noPerson', unreachable: 'unreachable' };
      for (var ck in c.integrityConfig) {
        if (oldMap[ck] && v2.qualityIssueFields.indexOf(oldMap[ck]) >= 0) c.integrityConfig[ck].enabled = true;
      }
    }
    if (v2.pipelineType) c.pipelineType = v2.pipelineType;
    if (v2.engagementWeights) {
      var ew = v2.engagementWeights;
      c.engagementConfig.withActivity = { enabled: true, weight: ew.withActivity >= 40 ? 'high' : (ew.withActivity >= 25 ? 'medium' : 'low') };
      c.engagementConfig.withPerson = { enabled: true, weight: ew.withPerson >= 40 ? 'high' : (ew.withPerson >= 25 ? 'medium' : 'low') };
      c.engagementConfig.withPipeline = { enabled: true, weight: ew.withPipeline >= 25 ? 'high' : (ew.withPipeline >= 15 ? 'medium' : 'low') };
    }
    return all;
  }

  function save() { try { localStorage.setItem(SK, JSON.stringify(_s)); } catch(e) {} }
  function get(entity) { if (!_s) load(); return _s[entity || 'company'] || entityDefaults(entity || 'company'); }

  // ============================================================
  // UDEF DISCOVERY
  // ============================================================
  function notifyUdefLoaded(entityId, data) {
    if (!data || !data.fields) return;
    var eKey = null;
    for (var k in ENTITY_DEFS) { if (ENTITY_DEFS[k].entityId === entityId) { eKey = k; break; } }
    if (!eKey) return;
    _udefFields[eKey] = [];
    for (var i = 0; i < data.fields.length; i++) {
      var f = data.fields[i];
      _udefFields[eKey].push({ progId: f.progId || f.label || ('f' + i), label: f.label, type: f.type, percent: f.percent });
    }
    if (_activeEntity === eKey && _activeId) {
      var el = document.getElementById(_activeId);
      if (el && el.offsetParent !== null) renderPanel(el, 'full');
    }
  }

  // ============================================================
  // WEIGHT CALC HELPERS
  // ============================================================
  function calcWeightPct(items) {
    var total = 0;
    for (var i = 0; i < items.length; i++) { if (items[i].enabled !== false) total += POINTS[items[i].weight] || 0; }
    var result = [];
    for (var i = 0; i < items.length; i++) {
      if (items[i].enabled === false) { result.push({ pct: 0, pts: 0 }); continue; }
      var pts = POINTS[items[i].weight] || 0;
      result.push({ pct: total > 0 ? Math.round(pts / total * 100) : 0, pts: pts });
    }
    return result;
  }

  // ============================================================
  // COMPUTE — backward-compat API for entity-logic
  // ============================================================
  function getCompletenessValue(key, ovCpl, qData, total) {
    if (!total) return 0;
    if (ovCpl && ovCpl[key] !== undefined) return ovCpl[key];
    var qm = { person:'noPerson', category:'noCategory', business:'noBusiness' };
    if (qm[key] && qData && qData[qm[key]] !== undefined) return total - qData[qm[key]];
    return 0;
  }

  function computeDQ(entity, ovCpl, qData, uData, total) {
    var s = get(entity);
    var def = ENTITY_DEFS[entity] || ENTITY_DEFS.company;
    var scores = {};
    if (ovCpl && total > 0) {
      var sum = 0; var wt = 0;
      for (var i = 0; i < def.stdFields.length; i++) {
        var k = def.stdFields[i].key;
        var imp = s.stdFieldConfig[k] || 'excluded';
        if (imp === 'excluded') continue;
        var w = imp === 'required' ? 2 : 1;
        sum += (getCompletenessValue(k, ovCpl, qData, total) / total * 100) * w;
        wt += w;
      }
      if (wt > 0) scores.completeness = sum / wt;
    }
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
    var tW = 0; var wS = 0;
    if (scores.completeness !== undefined) { wS += scores.completeness * 70; tW += 70; }
    if (scores.udef !== undefined) { wS += scores.udef * 30; tW += 30; }
    if (tW <= 0) return null;
    return { total: Math.round(wS / tW), components: scores, weights: { completeness: 70, udef: 30 } };
  }

  function computeEngagement(row, total, entity) {
    if (!total) return 0;
    var s = get(entity);
    var cfg = s.engagementConfig;
    var items = [];
    for (var k in cfg) { if (cfg[k].enabled !== false) items.push({ key: k, weight: cfg[k].weight, enabled: true }); }
    var wts = calcWeightPct(items);
    var score = 0;
    for (var i = 0; i < items.length; i++) {
      var val = 0;
      if (items[i].key === 'withPipeline') {
        var pt = s.pipelineType || 'sale';
        if (pt === 'sale') val = row.withSale || 0;
        else if (pt === 'project') val = row.withProject || 0;
        else if (pt === 'both') val = row.withSaleOrProject || 0;
      } else { val = row[items[i].key] || 0; }
      score += (val / total * 100) * wts[i].pct / 100;
    }
    return Math.round(score);
  }

  function getPipelineLabel(entity) {
    var s = get(entity || 'company');
    for (var i = 0; i < PIPE_OPTS.length; i++) if (PIPE_OPTS[i].key === s.pipelineType) return PIPE_OPTS[i].label;
    return 'Open Sale';
  }
  function getPipelineType(entity) { return get(entity || 'company').pipelineType; }

  // ============================================================
  // SCORE COMPUTATION — Integrity, Adoption, Overall Health
  // ============================================================

  /**
   * computeIntegrity(entity, checkData, total)
   * checkData: { checkKey: affectedCount, ... }
   * Returns { total: 0-100, details: [...] } or null
   */
  function computeIntegrity(entity, checkData, total) {
    if (!total || total <= 0 || !checkData) return null;
    var s = get(entity); var def = ENTITY_DEFS[entity];
    if (!def) return null;
    var checks = def.integrityChecks; var cfg = s.integrityConfig;
    var items = []; var wItems = [];
    for (var i = 0; i < checks.length; i++) {
      var c = checks[i]; var cc = cfg[c.key];
      if (!cc || !cc.enabled) continue;
      var count = checkData[c.key];
      if (count === undefined || count === null) continue;
      var pct = Math.min(100, Math.round(count / total * 1000) / 10);
      items.push({ key: c.key, label: c.label, affected: pct, weight: cc.weight });
      wItems.push({ weight: cc.weight, enabled: true });
    }
    if (items.length === 0) return null;
    var wts = calcWeightPct(wItems);
    var score = 0;
    for (var i = 0; i < items.length; i++) {
      score += (100 - items[i].affected) * wts[i].pct / 100;
    }
    return { total: Math.round(score), details: items };
  }

  /**
   * computeAdoption(entity, componentData, total)
   * componentData: { componentKey: rawCount, ... }
   * Returns { total: 0-100, details: [...] } or null
   */
  function computeAdoption(entity, componentData, total) {
    if (!total || total <= 0 || !componentData) return null;
    var s = get(entity); var def = ENTITY_DEFS[entity];
    if (!def) return null;
    var comps = def.engagementComponents; var cfg = s.engagementConfig;
    var items = []; var wItems = [];
    for (var i = 0; i < comps.length; i++) {
      var c = comps[i]; var cc = cfg[c.key];
      if (!cc || cc.enabled === false) continue;
      var count = componentData[c.key];
      if (count === undefined || count === null) continue;
      var pct = Math.min(100, Math.round(count / total * 1000) / 10);
      items.push({ key: c.key, label: c.label, pct: pct, weight: cc.weight });
      wItems.push({ weight: cc.weight, enabled: true });
    }
    if (items.length === 0) return null;
    var wts = calcWeightPct(wItems);
    var score = 0;
    for (var i = 0; i < items.length; i++) {
      score += items[i].pct * wts[i].pct / 100;
    }
    return { total: Math.round(score), details: items };
  }

  /**
   * computeHealth(entity, dqResult, intResult, adoptResult)
   * Each result: { total: 0-100 } or null
   * Returns { total: 0-100, components: [...] } or null
   */
  function computeHealth(entity, dqResult, intResult, adoptResult) {
    var s = get(entity);
    var hw = s.healthWeights; var he = s.healthExclude;
    var comps = []; var wItems = [];
    if (dqResult && dqResult.total != null) {
      comps.push({ key: 'dq', label: 'Data Quality', score: dqResult.total, weight: hw.dq || 'high' });
      wItems.push({ weight: hw.dq || 'high', enabled: true });
    }
    if (!he.integrity && intResult && intResult.total != null) {
      comps.push({ key: 'integrity', label: 'Data Integrity', score: intResult.total, weight: hw.integrity || 'medium' });
      wItems.push({ weight: hw.integrity || 'medium', enabled: true });
    }
    if (!he.adoption && adoptResult && adoptResult.total != null) {
      comps.push({ key: 'adoption', label: 'Adoption', score: adoptResult.total, weight: hw.adoption || 'high' });
      wItems.push({ weight: hw.adoption || 'high', enabled: true });
    }
    if (comps.length === 0) return null;
    var wts = calcWeightPct(wItems);
    var score = 0;
    for (var i = 0; i < comps.length; i++) score += comps[i].score * wts[i].pct / 100;
    return { total: Math.round(score), components: comps };
  }

  function getCompleteness(entity) {
    var s = get(entity || 'company');
    var def = ENTITY_DEFS[entity || 'company'];
    var fields = [];
    for (var i = 0; i < def.stdFields.length; i++) {
      var imp = s.stdFieldConfig[def.stdFields[i].key] || 'excluded';
      if (imp !== 'excluded') fields.push(def.stdFields[i].key);
    }
    return fields;
  }

  function getSettings(entity) {
    var s = get(entity || 'company');
    return {
      completenessFields: getCompleteness(entity),
      stdFieldConfig: s.stdFieldConfig,
      udefFieldConfig: s.udefFieldConfig,
      integrityConfig: s.integrityConfig,
      engagementConfig: s.engagementConfig,
      healthWeights: s.healthWeights,
      healthExclude: s.healthExclude,
      pipelineType: s.pipelineType,
      dqWeights: { completeness: 70, udef: 30, quality: 0 },
      qualityIssueFields: Object.keys(s.integrityConfig).filter(function(k) { return s.integrityConfig[k].enabled; }),
      engagementWeights: _backCompatEngWeights(s)
    };
  }

  function _backCompatEngWeights(s) {
    var cfg = s.engagementConfig;
    var items = [];
    for (var k in cfg) items.push({ key: k, weight: cfg[k].weight, enabled: cfg[k].enabled });
    var wts = calcWeightPct(items);
    var r = {};
    for (var i = 0; i < items.length; i++) r[items[i].key] = wts[i].pct;
    return r;
  }

  // ============================================================
  // RENDER
  // ============================================================
  function renderPanel(el, mode) {
    _mode = mode || 'full';
    var h = '';

    // Entity tabs
    h += '<div class="s-entity-tabs">';
    for (var i = 0; i < ENTITY_ORDER.length; i++) {
      var ek = ENTITY_ORDER[i];
      var def = ENTITY_DEFS[ek];
      var active = ek === _activeEntity ? ' active' : '';
      var icoAttr = el.getAttribute('data-ico-' + ek);
      var ico = icoAttr ? '<img class="s-entity-ico" src="data:image/svg+xml;base64,' + icoAttr + '">' : '';
      h += '<button class="s-entity-tab' + active + '" onclick="daSettings.switchEntity(\'' + ek + '\')">' + ico + def.label + '</button>';
    }
    h += '</div>';

    h += renderEntityPanel(_activeEntity);

    h += '<div class="s-actions">';
    h += '<button class="btn-settings-save" onclick="daSettings.doSave()">Apply &amp; Recalculate</button>';
    h += '<button class="btn-settings-reset" onclick="daSettings.doReset()">Reset to Defaults</button>';
    h += '</div>';

    el.innerHTML = h;
    el.setAttribute('data-mode', _mode);
    _activeId = el.id;
  }

  function renderEntityPanel(entityKey) {
    var s = get(entityKey);
    var def = ENTITY_DEFS[entityKey];
    var h = '';

    // ── OVERALL HEALTH ──
    h += '<div class="s-health-card">';
    h += '<h3>Overall Health Score <span class="s-info-icon" onclick="daSettings.openInfo()" title="Scoring methodology">i</span></h3>';
    h += '<div class="settings-desc">Combines all three scores into one metric. Adjust weights to reflect priorities.</div>';
    h += renderHealthWeights(entityKey, s);
    h += '</div>';

    // ── ACCORDION: DATA QUALITY ──
    var dqBadge = accBadgeDQ(s, def);
    h += '<div class="s-accordion open" id="sAcc_dq_' + entityKey + '">';
    h += '<div class="s-accordion-head" onclick="daSettings.toggleAcc(\'sAcc_dq_' + entityKey + '\')">';
    h += chevSvg();
    h += '<span class="s-accordion-title">Data Quality</span>';
    h += '<span class="s-accordion-badge">' + dqBadge + '</span>';
    h += '</div>';
    h += '<div class="s-accordion-body">';
    h += renderCompletenessBody(entityKey, s, def);
    h += '</div></div>';

    // ── ACCORDION: DATA INTEGRITY ──
    var intBadge = accBadgeInt(s, def);
    h += '<div class="s-accordion" id="sAcc_int_' + entityKey + '">';
    h += '<div class="s-accordion-head" onclick="daSettings.toggleAcc(\'sAcc_int_' + entityKey + '\')">';
    h += chevSvg();
    h += '<span class="s-accordion-title">Data Integrity</span>';
    h += '<span class="s-accordion-badge">' + intBadge + '</span>';
    h += '</div>';
    h += '<div class="s-accordion-body">';
    h += '<div class="s-acc-desc">Structural quality problems. Each check can be enabled and weighted independently.</div>';
    h += renderIntegrityChecks(entityKey, s, def);
    h += '</div></div>';

    // ── ACCORDION: ADOPTION ──
    var adoptBadge = accBadgeAdopt(s, def);
    h += '<div class="s-accordion" id="sAcc_adopt_' + entityKey + '">';
    h += '<div class="s-accordion-head" onclick="daSettings.toggleAcc(\'sAcc_adopt_' + entityKey + '\')">';
    h += chevSvg();
    h += '<span class="s-accordion-title">Adoption</span>';
    h += '<span class="s-accordion-badge">' + adoptBadge + '</span>';
    h += '</div>';
    h += '<div class="s-accordion-body">';
    h += '<div class="s-acc-desc">How each dimension contributes to the adoption score.</div>';
    h += renderEngagementWeights(entityKey, s, def);
    if (def.hasPipelineConfig) {
      h += '<div class="s-sub-card">';
      h += '<h4>Pipeline Definition</h4>';
      h += '<div class="s-sub-card-desc">What counts as "active pipeline".</div>';
      h += renderPipelineOpts(entityKey, s);
      h += '</div>';
    }
    h += '</div></div>';

    return h;
  }

  function chevSvg() {
    return '<span class="s-accordion-chev"><svg viewBox="0 0 10 6" fill="none"><path d="M1 1l4 4 4-4" stroke="currentColor" stroke-width="1.3" stroke-linecap="round" stroke-linejoin="round"/></svg></span>';
  }

  function accBadgeDQ(s, def) {
    var active = 0;
    for (var i = 0; i < def.stdFields.length; i++) {
      if ((s.stdFieldConfig[def.stdFields[i].key] || 'excluded') !== 'excluded') active++;
    }
    return active + ' of ' + def.stdFields.length + ' fields active';
  }
  function accBadgeInt(s, def) {
    var active = 0;
    for (var i = 0; i < def.integrityChecks.length; i++) {
      var cc = s.integrityConfig[def.integrityChecks[i].key];
      if (cc && cc.enabled) active++;
    }
    return active + ' of ' + def.integrityChecks.length + ' checks active';
  }
  function accBadgeAdopt(s, def) {
    var active = 0;
    for (var i = 0; i < def.engagementComponents.length; i++) {
      var cc = s.engagementConfig[def.engagementComponents[i].key];
      if (cc && cc.enabled !== false) active++;
    }
    return active + ' components';
  }

  function toggleAcc(id) {
    var el = document.getElementById(id);
    if (el) el.classList.toggle('open');
  }

  function renderHealthWeights(ek, s) {
    var scores = [
      { key: 'dq',        label: 'Data Quality',   alwaysOn: true },
      { key: 'integrity',  label: 'Data Integrity',  alwaysOn: false },
      { key: 'adoption',   label: 'Adoption',        alwaysOn: false }
    ];
    var hw = s.healthWeights; var he = s.healthExclude;
    var items = [];
    for (var i = 0; i < scores.length; i++) items.push({ weight: hw[scores[i].key] || 'high', enabled: !he[scores[i].key] });
    var pcts = calcWeightPct(items);

    var h = '<div class="s-weight-rows" data-group="health" data-entity="' + ek + '">';
    for (var i = 0; i < scores.length; i++) {
      var sc = scores[i]; var level = hw[sc.key] || 'high'; var excluded = !!he[sc.key];
      var rowCls = excluded ? ' s-wrow-excl' : '';
      h += '<div class="s-weight-row' + rowCls + '" data-key="' + sc.key + '" data-level="' + level + '" data-enabled="' + !excluded + '">';
      if (!sc.alwaysOn) h += '<label class="s-exclude-toggle"><input type="checkbox"' + (!excluded ? ' checked' : '') + ' onchange="daSettings.toggleHealthComp(this,\'' + ek + '\',\'' + sc.key + '\')"></label>';
      h += '<span class="s-weight-label">' + sc.label + '</span>';
      h += hmlToggle('health', ek, sc.key, level);
      h += '<span class="s-weight-pct ' + pctClass(level, excluded) + '">' + (excluded ? '\u2014' : pcts[i].pct + '%') + '</span>';
      h += '</div>';
    }
    h += '</div>';
    h += '<div class="s-weight-summary" id="healthSum_' + ek + '">' + weightSummaryText(items) + '</div>';
    return h;
  }

  function renderEngagementWeights(ek, s, def) {
    var cfg = s.engagementConfig; var comps = def.engagementComponents;
    var items = [];
    for (var i = 0; i < comps.length; i++) {
      var c = comps[i]; var cc = cfg[c.key] || { enabled: true, weight: 'medium' };
      items.push({ key: c.key, label: c.label, weight: cc.weight, enabled: cc.enabled !== false, alwaysOn: c.alwaysOn });
    }
    var pcts = calcWeightPct(items);
    var h = '<div class="s-weight-rows" data-group="eng" data-entity="' + ek + '">';
    for (var i = 0; i < items.length; i++) {
      var it = items[i]; var excluded = !it.enabled; var rowCls = excluded ? ' s-wrow-excl' : '';
      h += '<div class="s-weight-row' + rowCls + '" data-key="' + it.key + '" data-level="' + it.weight + '" data-enabled="' + it.enabled + '">';
      if (!it.alwaysOn) h += '<label class="s-exclude-toggle"><input type="checkbox"' + (it.enabled ? ' checked' : '') + ' onchange="daSettings.toggleEngComp(this,\'' + ek + '\',\'' + it.key + '\')"></label>';
      h += '<span class="s-weight-label">' + it.label + '</span>';
      h += hmlToggle('eng', ek, it.key, it.weight);
      h += '<span class="s-weight-pct ' + pctClass(it.weight, excluded) + '">' + (excluded ? '\u2014' : pcts[i].pct + '%') + '</span>';
      h += '</div>';
    }
    h += '</div>';
    h += '<div class="s-weight-summary" id="engSum_' + ek + '">' + weightSummaryText(items) + '</div>';
    return h;
  }

  function hmlToggle(group, ek, key, level) {
    var h = '<div class="s-fld-toggle">';
    var vals = ['high', 'medium', 'low'];
    var labels = ['High', 'Medium', 'Low'];
    var acts = { high: 'act-high', medium: 'act-med', low: 'act-low' };
    for (var i = 0; i < vals.length; i++) {
      var cls = level === vals[i] ? ' ' + acts[vals[i]] : '';
      h += '<button class="s-imp-btn' + cls + '" onclick="daSettings.setHML(\'' + group + '\',\'' + ek + '\',\'' + key + '\',\'' + vals[i] + '\',this)">' + labels[i] + '</button>';
    }
    h += '</div>';
    return h;
  }

  function renderIntegrityChecks(ek, s, def) {
    var cfg = s.integrityConfig;
    var h = '<div class="s-integrity-rows" data-entity="' + ek + '">';
    for (var i = 0; i < def.integrityChecks.length; i++) {
      var c = def.integrityChecks[i];
      var cc = cfg[c.key] || { enabled: false, weight: 'medium' };
      var disabled = !cc.enabled; var rowCls = disabled ? ' s-introw-disabled' : '';
      h += '<div class="s-integrity-row' + rowCls + '" data-key="' + c.key + '" data-level="' + cc.weight + '" data-enabled="' + cc.enabled + '">';
      h += '<label class="s-integrity-check"><input type="checkbox"' + (cc.enabled ? ' checked' : '') + ' onchange="daSettings.toggleIntegrity(this,\'' + ek + '\',\'' + c.key + '\')"></label>';
      h += '<div class="s-integrity-label"><span class="s-integrity-name">' + c.label + '</span><span class="s-integrity-desc">' + c.desc + '</span></div>';
      h += hmlToggle('int', ek, c.key, cc.weight);
      h += '</div>';
    }
    h += '</div>';
    return h;
  }

  function renderCompletenessBody(ek, s, def) {
    var h = '';
    h += '<div class="s-acc-desc">Which fields count towards completeness. <strong>Required</strong> fields count double, <strong>Excluded</strong> are ignored.</div>';
    h += '<div class="s-cpl-sub-head">Standard Fields</div>';
    h += '<div class="s-field-list">';
    for (var i = 0; i < def.stdFields.length; i++) {
      var sf = def.stdFields[i]; var imp = s.stdFieldConfig[sf.key] || 'excluded';
      h += fieldRow(sf.label, null, null, imp, 'std', sf.key, ek);
    }
    h += '</div>';
    var ufields = _udefFields[ek] || [];
    h += '<div class="s-cpl-sub-head s-cpl-udef-head" onclick="daSettings.toggleUdef(\'' + ek + '\')">';
    h += '<span class="s-udef-chev" id="sUdefChev_' + ek + '"><svg viewBox="0 0 10 6" fill="none"><path d="M1 1l4 4 4-4" stroke="currentColor" stroke-width="1.3" stroke-linecap="round" stroke-linejoin="round"/></svg></span> Custom Fields (UDEF)';
    if (ufields.length > 0) h += ' <span class="s-udef-count">' + ufields.length + ' fields</span>';
    else h += ' <span class="s-udef-count">\u2014 loaded after analysis</span>';
    h += '</div>';
    h += '<div class="s-udef-body" id="sUdefBody_' + ek + '" style="display:none">';
    if (ufields.length === 0) {
      h += '<div class="s-udef-empty">Run the analysis first to discover custom fields.</div>';
    } else {
      h += '<div class="s-field-list">';
      for (var i = 0; i < ufields.length; i++) {
        var uf = ufields[i]; var imp = s.udefFieldConfig[uf.progId] || 'normal';
        var fillCol = uf.percent >= 70 ? 'var(--sl-good)' : (uf.percent >= 30 ? 'var(--sl-ok)' : 'var(--sl-bad)');
        h += fieldRow(uf.label, uf.type, '<span style="font-weight:600;color:' + fillCol + '">' + uf.percent + '%</span>', imp, 'udef', uf.progId, ek);
      }
      h += '</div>';
    }
    h += '</div>';
    return h;
  }

  function renderPipelineOpts(ek, s) {
    var h = '';
    for (var i = 0; i < PIPE_OPTS.length; i++) {
      var po = PIPE_OPTS[i];
      h += '<label class="settings-radio">';
      h += '<input type="radio" name="pipelineType_' + ek + '" value="' + po.key + '"' + (s.pipelineType === po.key ? ' checked' : '') + '>';
      h += '<span><strong>' + po.label + '</strong><span class="settings-radio-desc">' + po.desc + '</span></span></label>';
    }
    return h;
  }

  function fieldRow(label, type, extra, imp, group, key, ek) {
    var esc = key.replace(/'/g, "\\'");
    var h = '<div class="s-fld-row' + (imp === 'excluded' ? ' s-fld-excl' : '') + '" data-grp="' + group + '" data-key="' + key + '" data-entity="' + ek + '">';
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

  function pctClass(level, excluded) {
    if (excluded) return 's-pct-off';
    return level === 'high' ? 's-pct-high' : (level === 'medium' ? 's-pct-med' : 's-pct-low');
  }
  function weightSummaryText(items) {
    var labels = { high: 'High (3pt)', medium: 'Medium (2pt)', low: 'Low (1pt)' };
    var parts = [];
    for (var i = 0; i < items.length; i++) { if (items[i].enabled !== false) parts.push(labels[items[i].weight]); }
    if (parts.length === 0) return '<span class="s-wsum-dot s-wsum-dot-off"></span> No components enabled';
    return '<span class="s-wsum-dot"></span> Total: 100% \u2014 ' + parts.join(' + ');
  }

  // ============================================================
  // EVENT HANDLERS
  // ============================================================
  function switchEntity(ek) {
    // Save current entity DOM state before switching
    readEntityFromDOM(_activeEntity);
    _activeEntity = ek;
    if (_activeId) { var el = document.getElementById(_activeId); if (el) renderPanel(el, _mode); }
  }


  function setHML(group, ek, key, level, btn) {
    var toggle = btn.parentElement;
    var btns = toggle.querySelectorAll('.s-imp-btn');
    for (var i = 0; i < btns.length; i++) btns[i].classList.remove('act-high', 'act-med', 'act-low');
    btn.classList.add(level === 'high' ? 'act-high' : (level === 'medium' ? 'act-med' : 'act-low'));
    var row = btn.closest('.s-weight-row') || btn.closest('.s-integrity-row');
    if (row) row.setAttribute('data-level', level);
    if (group === 'health') recalcWeightGroup('health', ek);
    else if (group === 'eng') recalcWeightGroup('eng', ek);
  }

  function recalcWeightGroup(group, ek) {
    var container = document.querySelector('.s-weight-rows[data-group="' + group + '"][data-entity="' + ek + '"]');
    if (!container) return;
    var rows = container.querySelectorAll('.s-weight-row');
    var items = [];
    for (var i = 0; i < rows.length; i++) items.push({ weight: rows[i].getAttribute('data-level'), enabled: rows[i].getAttribute('data-enabled') === 'true' });
    var pcts = calcWeightPct(items);
    for (var i = 0; i < rows.length; i++) {
      var pctEl = rows[i].querySelector('.s-weight-pct');
      if (!pctEl) continue;
      if (items[i].enabled) { pctEl.textContent = pcts[i].pct + '%'; pctEl.className = 's-weight-pct ' + pctClass(items[i].weight, false); }
      else { pctEl.textContent = '\u2014'; pctEl.className = 's-weight-pct s-pct-off'; }
    }
    var sumId = (group === 'health' ? 'healthSum_' : 'engSum_') + ek;
    var sumEl = document.getElementById(sumId);
    if (sumEl) sumEl.innerHTML = weightSummaryText(items);
  }

  function toggleHealthComp(cb, ek, key) {
    var row = cb.closest('.s-weight-row');
    if (cb.checked) { row.classList.remove('s-wrow-excl'); row.setAttribute('data-enabled', 'true'); }
    else { row.classList.add('s-wrow-excl'); row.setAttribute('data-enabled', 'false'); }
    recalcWeightGroup('health', ek);
  }

  function toggleEngComp(cb, ek, key) {
    var row = cb.closest('.s-weight-row');
    if (cb.checked) { row.classList.remove('s-wrow-excl'); row.setAttribute('data-enabled', 'true'); }
    else { row.classList.add('s-wrow-excl'); row.setAttribute('data-enabled', 'false'); }
    recalcWeightGroup('eng', ek);
  }

  function toggleIntegrity(cb, ek, key) {
    var row = cb.closest('.s-integrity-row');
    if (cb.checked) { row.classList.remove('s-introw-disabled'); row.setAttribute('data-enabled', 'true'); }
    else { row.classList.add('s-introw-disabled'); row.setAttribute('data-enabled', 'false'); }
  }

  function setImp(group, key, imp, btn) {
    var btns = btn.parentElement.querySelectorAll('.s-imp-btn');
    for (var i = 0; i < btns.length; i++) btns[i].className = 's-imp-btn';
    btn.className = 's-imp-btn ' + (imp === 'required' ? 'act-req' : (imp === 'normal' ? 'act-norm' : 'act-excl'));
    var row = btn.closest('.s-fld-row');
    if (row) { if (imp === 'excluded') row.classList.add('s-fld-excl'); else row.classList.remove('s-fld-excl'); }
  }

  function toggleUdef(ek) {
    var body = document.getElementById('sUdefBody_' + ek);
    var chev = document.getElementById('sUdefChev_' + ek);
    if (!body) return;
    var open = body.style.display !== 'none';
    body.style.display = open ? 'none' : 'block';
    if (chev) { if (open) chev.classList.remove('s-udef-open'); else chev.classList.add('s-udef-open'); }
  }

  // ============================================================
  // INFO PANEL
  // ============================================================
  function openInfo() {
    var overlay = document.getElementById('sInfoOverlay');
    if (overlay) { overlay.classList.add('open'); return; }
    var html = '<div class="s-info-overlay" id="sInfoOverlay" onclick="daSettings.closeInfo()">';
    html += '<div class="s-info-panel" onclick="event.stopPropagation()">';
    html += '<div class="s-info-header"><h2>Scoring Methodology</h2><button class="s-info-close" onclick="daSettings.closeInfo()">\u2715</button></div>';
    html += '<div class="s-info-body">';
    html += '<div class="s-info-section"><h3>Three Scores + Overall</h3>';
    html += '<p>Each entity is measured on three independent scores, combined into one Overall Health score.</p>';
    html += '<div class="s-info-callout"><strong>Data Quality</strong><p>Are fields filled in? Based on Completeness Definition with Required/Normal/Excluded weights.</p></div>';
    html += '<div class="s-info-callout"><strong>Data Integrity</strong><p>Structural problems? Cross-entity relationship checks, each with H/M/L weight.</p></div>';
    html += '<div class="s-info-callout"><strong>Adoption</strong><p>Is the CRM being used? Activities, relationships, and pipeline metrics.</p></div></div>';
    html += '<div class="s-info-section"><h3>How Weights Work (H/M/L)</h3>';
    html += '<p>High = 3pt, Medium = 2pt, Low = 1pt. Weight % = points \u00F7 total points \u00D7 100.</p>';
    html += '<p>Example: H/M/H = 3+2+3 = 8pt \u2192 38% / 25% / 38%</p></div>';
    html += '<div class="s-info-section"><h3>Overall Health</h3>';
    html += '<div class="s-info-formula">Health = (DQ \u00D7 W\u2081) + (Integrity \u00D7 W\u2082) + (Adoption \u00D7 W\u2083)</div>';
    html += '<p>Components can be excluded entirely per entity.</p></div>';
    html += '<div class="s-info-section"><h3>Per-Entity</h3>';
    html += '<p>Each entity has independent settings \u2014 different fields, checks, engagement metrics, and health weights.</p></div>';
    html += '</div></div></div>';
    document.body.insertAdjacentHTML('beforeend', html);
    setTimeout(function() { document.getElementById('sInfoOverlay').classList.add('open'); }, 10);
  }
  function closeInfo() {
    var overlay = document.getElementById('sInfoOverlay');
    if (overlay) { overlay.classList.remove('open'); }
  }

  // ============================================================
  // SAVE / RESET
  // ============================================================
  function readEntityFromDOM(ek) {
    var s = get(ek); var def = ENTITY_DEFS[ek];

    // Health
    var healthRows = document.querySelectorAll('.s-weight-rows[data-group="health"][data-entity="' + ek + '"] .s-weight-row');
    if (healthRows.length > 0) {
      s.healthWeights = {}; s.healthExclude = {};
      for (var i = 0; i < healthRows.length; i++) {
        var key = healthRows[i].getAttribute('data-key');
        s.healthWeights[key] = healthRows[i].getAttribute('data-level');
        s.healthExclude[key] = healthRows[i].getAttribute('data-enabled') !== 'true';
      }
    }
    // Std fields
    var stdRows = document.querySelectorAll('.s-fld-row[data-grp="std"][data-entity="' + ek + '"]');
    if (stdRows.length > 0) {
      s.stdFieldConfig = {};
      for (var i = 0; i < stdRows.length; i++) {
        var k = stdRows[i].getAttribute('data-key');
        var act = stdRows[i].querySelector('.s-imp-btn.act-req,.s-imp-btn.act-norm,.s-imp-btn.act-excl');
        if (act) { if (act.classList.contains('act-req')) s.stdFieldConfig[k] = 'required'; else if (act.classList.contains('act-norm')) s.stdFieldConfig[k] = 'normal'; else s.stdFieldConfig[k] = 'excluded'; }
      }
    }
    // UDEF
    var uRows = document.querySelectorAll('.s-fld-row[data-grp="udef"][data-entity="' + ek + '"]');
    if (uRows.length > 0) {
      s.udefFieldConfig = {};
      for (var i = 0; i < uRows.length; i++) {
        var k = uRows[i].getAttribute('data-key');
        var act = uRows[i].querySelector('.s-imp-btn.act-req,.s-imp-btn.act-norm,.s-imp-btn.act-excl');
        if (act) { if (act.classList.contains('act-req')) s.udefFieldConfig[k] = 'required'; else if (act.classList.contains('act-excl')) s.udefFieldConfig[k] = 'excluded'; }
      }
    }
    // Integrity
    var intRows = document.querySelectorAll('.s-integrity-rows[data-entity="' + ek + '"] .s-integrity-row');
    if (intRows.length > 0) {
      s.integrityConfig = {};
      for (var i = 0; i < intRows.length; i++) {
        var key = intRows[i].getAttribute('data-key');
        s.integrityConfig[key] = { enabled: intRows[i].getAttribute('data-enabled') === 'true', weight: intRows[i].getAttribute('data-level') };
      }
    }
    // Engagement
    var engRows = document.querySelectorAll('.s-weight-rows[data-group="eng"][data-entity="' + ek + '"] .s-weight-row');
    if (engRows.length > 0) {
      s.engagementConfig = {};
      for (var i = 0; i < engRows.length; i++) {
        var key = engRows[i].getAttribute('data-key');
        s.engagementConfig[key] = { enabled: engRows[i].getAttribute('data-enabled') === 'true', weight: engRows[i].getAttribute('data-level') };
      }
    }
    // Pipeline
    var pipeRadio = document.querySelector('input[name="pipelineType_' + ek + '"]:checked');
    if (pipeRadio) s.pipelineType = pipeRadio.value;
    _s[ek] = s;
  }

  function doSave() {
    readEntityFromDOM(_activeEntity);
    save();
    if (typeof renderDQScore === 'function') renderDQScore('company');
    if (typeof renderCrossEntityFunnel === 'function' && typeof companyDetailData !== 'undefined' && companyDetailData) renderCrossEntityFunnel(companyDetailData);
    if (typeof renderCompanyDetails === 'function' && typeof companyDetailData !== 'undefined' && companyDetailData) renderCompanyDetails(companyDetailData);
    toast('Settings applied and scores recalculated');
  }

  function doReset() {
    _s = allDefaults(); save();
    if (_activeId) { var el = document.getElementById(_activeId); if (el) renderPanel(el, _mode); }
    toast('Settings reset to defaults');
  }

  function toast(msg) {
    var ex = document.getElementById('sToast'); if (ex) ex.remove();
    var t = document.createElement('div'); t.id = 'sToast'; t.className = 'settings-toast'; t.textContent = msg;
    document.body.appendChild(t);
    setTimeout(function(){ t.classList.add('show'); }, 10);
    setTimeout(function(){ t.classList.remove('show'); setTimeout(function(){ t.remove(); }, 300); }, 2500);
  }

  function init() { load(); var p = document.getElementById('settings-quality'); if (p) renderPanel(p, 'full'); }

  // ============================================================
  // PUBLIC API
  // ============================================================
  window.daSettings = {
    init: init,
    getSettings: getSettings,
    computeDQScore: computeDQ,
    computeIntegrity: computeIntegrity,
    computeAdoption: computeAdoption,
    computeHealth: computeHealth,
    computeEngagement: computeEngagement,
    getCompletenessValue: getCompletenessValue,
    getPipelineLabel: getPipelineLabel,
    getPipelineType: getPipelineType,
    notifyUdefLoaded: notifyUdefLoaded,
    COMPLETENESS_OPTIONS: ENTITY_DEFS.company.stdFields.map(function(f){ return { key: f.key, label: f.label }; }),
    QUALITY_ISSUE_OPTIONS: ENTITY_DEFS.company.integrityChecks.map(function(c){ return { key: c.key, label: c.label }; }),
    switchEntity: switchEntity, toggleAcc: toggleAcc, setHML: setHML,
    toggleHealthComp: toggleHealthComp, toggleEngComp: toggleEngComp,
    toggleIntegrity: toggleIntegrity, setImp: setImp, toggleUdef: toggleUdef,
    openInfo: openInfo, closeInfo: closeInfo,
    doSave: doSave, doReset: doReset,
    ENTITY_DEFS: ENTITY_DEFS, ENTITY_ORDER: ENTITY_ORDER
  };
  load();
})();
