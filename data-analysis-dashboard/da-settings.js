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
        { key: 'noPerson',      label: 'No contact person',  desc: 'No contacts means nobody to approach within the company', query: 'count(person WHERE contact_id = company) = 0' },
        { key: 'unreachable',   label: 'Unreachable company', desc: 'All contacts lack working email and phone numbers', query: 'has_email = 0 AND has_phone = 0' },
        { key: 'noCategory',    label: 'No category',         desc: 'Cannot segment customers from prospects without category', query: 'company.category_idx = 0' },
        { key: 'noBusiness',    label: 'No business type',    desc: 'Industry analysis and benchmarks become impossible', query: 'company.business_idx = 0' },
        { key: 'noOrgNr',       label: 'No org. number',      desc: 'Duplicate detection and verification rely on this identifier', query: 'company.orgNr IS EMPTY' },
        { key: 'noActivity12m', label: 'No recent activity',  desc: 'Account has gone dormant for over a year', query: 'count(appointment WHERE activeDate >= now - 12 months) = 0', isNew: true },
        { key: 'noOwner',       label: 'No owner',            desc: 'Without owner, accountability and follow-up disappear', query: 'company.associate_id <= 0', isNew: true }
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
        { key: 'noEmail',     label: 'No email address',     desc: 'Contact has no email and cannot be reached digitally', query: 'email_address NOT CONTAINS "@"' },
        { key: 'noCompany',   label: 'Not linked to company', desc: 'Orphan contact has no business context to act on', query: 'person.contact_id <= 0' },
        { key: 'noActivity',  label: 'No activity',          desc: 'No interaction logged for over a year, relationship dormant', query: 'count(appointment WHERE activeDate >= now - 12 months) = 0', isNew: true }
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
        { key: 'noContact',     label: 'No contact linked',  desc: 'Sales without a contact person are hard to follow up', query: 'sale.person_id <= 0' },
        { key: 'staleSale',     label: 'Stale sale',         desc: 'Forecast is overstated as close-date has already passed', query: 'sale.status = Open AND sale.saledate < NOW()' },
        { key: 'noActivities',  label: 'No activities logged', desc: 'No activities suggest the deal has stalled or been forgotten', query: 'count(appointment WHERE sale_id = sale) = 0' },
        { key: 'noAmount',      label: 'No amount',          desc: 'Without amount, pipeline value cannot be forecasted', query: 'sale.amount <= 0', isNew: true }
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
        { key: 'noMembers',    label: 'No members',    desc: 'Project lacks a team, delivery responsibility is unclear', query: 'count(project_member WHERE project_id = project) = 0' },
        { key: 'noActivities', label: 'No activities',  desc: 'No activities suggest delivery has stalled or paused', query: 'count(appointment WHERE project_id = project) = 0' }
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
    d._momentum = clone(MOMENTUM_DEFAULTS);
    return d;
  }

  // ============================================================
  // STATE
  // ============================================================
  var SK = 'da_settings_v3';
  var _s = null;
  var _serverAvailable = false;
  var _rules = {};
  var _intakeFields = {};
  var _currentScores = {};   // optional per-entity current scores set by the host after a scan
  var _scanState = null;     // { scannedAt, scores } from the latest scan snapshot
  var SCORE_TARGET_DEFS = [
    { key: 'dataQuality',   label: 'Data Quality',   desc: 'How completely the important fields are filled in across records.' },
    { key: 'dataIntegrity', label: 'Data Integrity', desc: 'How free the data is of structural problems like missing links or owners.' },
    { key: 'adoption',      label: 'Adoption',       desc: 'How actively the CRM is used: activities, relationships and pipeline.' }
  ];
  var SCORE_TARGET_DEFAULTS = { dataQuality: 80, dataIntegrity: 85, adoption: 70 };
  var _udefFields = {};
  var _activeId = null;
  var _activeEntity = 'company';

  // Momentum defaults (global, not per-entity)
  var MOMENTUM_DEFAULTS = {
    powerThreshold: 100,
    regularThreshold: 25,
    dateField: 'activeDate',
    excludedTypes: []
  };

  function clone(o) { return JSON.parse(JSON.stringify(o)); }

  function load() {
    // Synchronous load from localStorage (instant, used on page load)
    try {
      var raw = localStorage.getItem(SK);
      if (raw) { _s = migrateV3(JSON.parse(raw)); return; }
      var v2 = localStorage.getItem('da_settings_v2');
      if (v2) { _s = migrateFromV2(JSON.parse(v2)); saveLocal(); return; }
    } catch(e) { console.warn('Settings load error:', e); }
    _s = allDefaults();
  }

  // Parse the per-entity rule shards returned by SettingsFetch
  function parseRules(rulesObj) {
    _rules = {};
    if (!rulesObj) return;
    for (var ek in rulesObj) {
      try { _rules[ek] = typeof rulesObj[ek] === 'string' ? JSON.parse(rulesObj[ek]) : rulesObj[ek]; }
      catch (e) { _rules[ek] = []; }
    }
  }

  // Parse the per-entity field-config shards written by the intake form
  function parseIntakeFields(fieldsObj) {
    _intakeFields = {};
    if (!fieldsObj) return;
    for (var ek in fieldsObj) {
      try { _intakeFields[ek] = typeof fieldsObj[ek] === 'string' ? JSON.parse(fieldsObj[ek]) : fieldsObj[ek]; }
      catch (e) { _intakeFields[ek] = null; }
    }
  }

  // Parse the latest scan snapshot summary. Feeds the current-score markers in
  // settings and the scan-status badge. Shape: { scannedAt, scores: { entity: {...} } }.
  function parseScan(scanRaw) {
    _scanState = null;
    if (!scanRaw) return;
    try { _scanState = (typeof scanRaw === 'string') ? JSON.parse(scanRaw) : scanRaw; }
    catch (e) { _scanState = null; }
    if (_scanState && _scanState.scores) _currentScores = _scanState.scores;
  }

  // Write the latest scan snapshot summary. Called by the dashboard after an
  // analysis completes. Stored under the reserved scan: key namespace, which is
  // designed to also carry the full dashboard snapshot and dated history later.
  function saveScanState(scannedAt, scores) {
    _scanState = { scannedAt: scannedAt, scores: scores || {} };
    _currentScores = _scanState.scores;
    if (typeof settingsSaveUrl === 'undefined' || !settingsSaveUrl) return;
    try {
      var x = new XMLHttpRequest();
      x.open('POST', settingsSaveUrl, true);
      x.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded');
      x.send('key=' + encodeURIComponent('scan:latest') + '&config=' + encodeURIComponent(JSON.stringify(_scanState)));
    } catch (e) {}
  }

  // Overlay intake field choices onto the in-memory settings.
  // The intake form is authoritative for the keys it carries; any field it does
  // not mention keeps its engine default. Runs in memory only and is never
  // persisted back into the engine-owned config row. Entities without a scoring
  // model in the dashboard (ticket, activity) are skipped.
  function applyIntakeFields() {
    if (!_s) load();
    for (var ek in _intakeFields) {
      var shard = _intakeFields[ek];
      if (!shard) continue;
      if (!ENTITY_DEFS[ek]) continue;
      var s = _s[ek];
      if (!s) { s = entityDefaults(ek); _s[ek] = s; }
      if (shard.stdFieldConfig) {
        for (var fk in shard.stdFieldConfig) s.stdFieldConfig[fk] = shard.stdFieldConfig[fk];
      }
      if (shard.udefFieldConfig) {
        if (!s.udefFieldConfig) s.udefFieldConfig = {};
        for (var uk in shard.udefFieldConfig) s.udefFieldConfig[uk] = shard.udefFieldConfig[uk];
      }
      if (shard.integrityConfig) {
        if (!s.integrityConfig) s.integrityConfig = {};
        for (var ik in shard.integrityConfig) s.integrityConfig[ik] = shard.integrityConfig[ik];
      }
      // Value-level exclusions are carried for the measurement layer (handled in a later slice).
      if (shard.fieldExclusions) s.fieldExclusions = shard.fieldExclusions;
    }
  }

  // Seed the engine-owned score targets. Targets live in the engine config and are
  // edited in settings. The intake only provides an initial value: when the config
  // has no target for an entity yet, seed it from the intake shard, otherwise from
  // the defaults. Once saved from settings, the config owns it and seeding stops.
  function seedScoreTargets() {
    if (!_s) load();
    for (var i = 0; i < ENTITY_ORDER.length; i++) {
      var ek = ENTITY_ORDER[i];
      if (!_s[ek]) continue;
      if (_s[ek].scoreTargets) continue;
      var seed = (_intakeFields[ek] && _intakeFields[ek].scoreTargets) ? _intakeFields[ek].scoreTargets : null;
      _s[ek].scoreTargets = {
        dataQuality:   (seed && typeof seed.dataQuality === 'number') ? seed.dataQuality : SCORE_TARGET_DEFAULTS.dataQuality,
        dataIntegrity: (seed && typeof seed.dataIntegrity === 'number') ? seed.dataIntegrity : SCORE_TARGET_DEFAULTS.dataIntegrity,
        adoption:      (seed && typeof seed.adoption === 'number') ? seed.adoption : SCORE_TARGET_DEFAULTS.adoption
      };
    }
  }

  // True when the intake form owns this key for the given group (std|udef|int).
  // Owned keys are locked in the settings modal and stripped before persisting,
  // so the engine config row never stores intake-owned choices.
  function intakeOwns(ek, group, key) {
    var shard = _intakeFields[ek];
    if (!shard) return false;
    if (group === 'std') return !!(shard.stdFieldConfig && shard.stdFieldConfig.hasOwnProperty(key));
    if (group === 'udef') return !!(shard.udefFieldConfig && shard.udefFieldConfig.hasOwnProperty(key));
    if (group === 'int') return !!(shard.integrityConfig && shard.integrityConfig.hasOwnProperty(key));
    return false;
  }

  // Deep clone of the settings with intake-owned keys removed, used for persistence.
  // Keeps _s fully overlaid in memory while the stored config stays engine-owned.
  function persistView() {
    var out = clone(_s);
    for (var ek in _intakeFields) {
      var shard = _intakeFields[ek];
      if (!shard || !out[ek]) continue;
      if (shard.stdFieldConfig && out[ek].stdFieldConfig) {
        for (var fk in shard.stdFieldConfig) delete out[ek].stdFieldConfig[fk];
      }
      if (shard.udefFieldConfig && out[ek].udefFieldConfig) {
        for (var uk in shard.udefFieldConfig) delete out[ek].udefFieldConfig[uk];
      }
      if (shard.integrityConfig && out[ek].integrityConfig) {
        for (var ik in shard.integrityConfig) delete out[ek].integrityConfig[ik];
      }
      if (out[ek].fieldExclusions) delete out[ek].fieldExclusions;
    }
    return out;
  }

  // Async load from server — overwrites localStorage if server has data
  function loadFromServer(callback) {
    if (typeof settingsLoadUrl === 'undefined' || !settingsLoadUrl) { if (callback) callback(false); return; }
    try {
      ajax(settingsLoadUrl, function(resp) {
        if (!resp) { if (callback) callback(false); return; }
        _serverAvailable = resp.tableExists || false;
        parseRules(resp.rules);
        parseIntakeFields(resp.fields);
        parseScan(resp.scan);
        if (resp.found && resp.config) {
          try {
            var parsed = typeof resp.config === 'string' ? JSON.parse(resp.config) : resp.config;
            _s = migrateV3(parsed);
            saveLocal(); // sync localStorage with server data
            applyIntakeFields(); // in-memory overlay, after saveLocal so it is not persisted
            seedScoreTargets();  // seed engine-owned targets from intake/defaults when config lacks them
            console.log('Settings loaded from server (updated by associate ' + (resp.updatedBy || '?') + ')');
            if (callback) callback(true);
          } catch(e) {
            console.warn('Settings parse error from server:', e);
            if (callback) callback(false);
          }
        } else {
          // Server table exists but no row yet — push localStorage settings to server
          if (_serverAvailable && _s) saveToServer();
          applyIntakeFields(); // overlay intake choices onto defaults even without an engine config row
          seedScoreTargets();
          if (callback) callback(false);
        }
      });
    } catch(e) {
      console.warn('Settings server load failed:', e);
      if (callback) callback(false);
    }
  }

  function saveLocal() { try { localStorage.setItem(SK, JSON.stringify(persistView())); } catch(e) {} }

  function saveToServer(callback) {
    if (!_serverAvailable || typeof settingsSaveUrl === 'undefined' || !settingsSaveUrl) {
      if (callback) callback(false);
      return;
    }
    try {
      var x = new XMLHttpRequest();
      x.open('POST', settingsSaveUrl, true);
      x.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded');
      x.onreadystatechange = function() {
        if (x.readyState === 4) {
          if (x.status === 200 && x.responseText.length > 0) {
            try {
              var resp = JSON.parse(x.responseText);
              if (resp && resp.success) {
                console.log('Settings saved to server');
                if (callback) callback(true);
              } else {
                console.warn('Settings server save failed:', resp ? resp.error : 'no response');
                if (callback) callback(false);
              }
            } catch(e) { if (callback) callback(false); }
          } else { if (callback) callback(false); }
        }
      };
      x.send('config=' + encodeURIComponent(JSON.stringify(persistView())));
    } catch(e) {
      console.warn('Settings server save error:', e);
      if (callback) callback(false);
    }
  }

  function save() {
    saveLocal();
    saveToServer();
  }

  function get(entity) { if (!_s) load(); return _s[entity || 'company'] || entityDefaults(entity || 'company'); }

  function getRules(entity) { return (_rules && _rules[entity]) ? _rules[entity] : []; }

  function migrateV3(saved) {
    var defs = allDefaults();
    for (var ek in defs) {
      if (!saved[ek]) { saved[ek] = defs[ek]; continue; }
      var d = defs[ek]; var s = saved[ek];
      if (!s.healthWeights) s.healthWeights = d.healthWeights;
      if (!s.healthExclude) s.healthExclude = d.healthExclude;
      if (!s.stdFieldConfig) s.stdFieldConfig = d.stdFieldConfig;
      if (!s.udefFieldConfig) s.udefFieldConfig = {};
      if (!s.fieldExclusions) s.fieldExclusions = {};
      if (!s.integrityConfig) s.integrityConfig = d.integrityConfig;
      if (!s.engagementConfig) s.engagementConfig = d.engagementConfig;
      if (!s.pipelineType) s.pipelineType = d.pipelineType;
      for (var ck in d.integrityConfig) { if (!s.integrityConfig[ck]) s.integrityConfig[ck] = d.integrityConfig[ck]; }
      for (var cek in d.engagementConfig) { if (!s.engagementConfig[cek]) s.engagementConfig[cek] = d.engagementConfig[cek]; }
      for (var sfk in d.stdFieldConfig) { if (s.stdFieldConfig[sfk] === undefined) s.stdFieldConfig[sfk] = d.stdFieldConfig[sfk]; }
    }
    if (!saved._momentum) saved._momentum = clone(MOMENTUM_DEFAULTS);
    else {
      var md = MOMENTUM_DEFAULTS;
      for (var mk in md) { if (saved._momentum[mk] === undefined) saved._momentum[mk] = md[mk]; }
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

  // ============================================================
  // UDEF DISCOVERY
  // ============================================================
  function notifyUdefLoaded(entityId, data) {
    if (!data || !data.fields) return;
    // Map udefId to entity key (udefId differs from entityId for contact and project)
    var udefIdMap = {7:'company', 8:'contact', 10:'sale', 9:'project'};
    var eKey = udefIdMap[entityId] || null;
    if (!eKey) {
      // Fallback: try matching by ENTITY_DEFS.entityId
      for (var k in ENTITY_DEFS) { if (ENTITY_DEFS[k].entityId === entityId) { eKey = k; break; } }
    }
    if (!eKey) return;
    _udefFields[eKey] = [];
    for (var i = 0; i < data.fields.length; i++) {
      var f = data.fields[i];
      _udefFields[eKey].push({ progId: f.progId || f.label || ('f' + i), label: f.label, type: f.type, percent: f.percent });
    }
    if (_activeEntity === eKey) {
      // Re-render entity section in modal if open
      var overlay = document.getElementById('smOverlay');
      if (overlay && overlay.classList.contains('open')) {
        var sec = document.getElementById('sm-' + eKey);
        if (sec) {
          var h = '<div class="sm-title">' + ENTITY_DEFS[eKey].label + ' Scoring</div>';
          h += '<div class="sm-subtitle">Configure how the ' + ENTITY_DEFS[eKey].label + ' health score is calculated.</div>';
          h += renderEntityPanel(eKey);
          sec.innerHTML = h;
        }
      }
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
    var def = ENTITY_DEFS[entity];
    if (!s || !def) return null;
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
    if (!def || !s) return null;
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
    if (!def || !s) return null;
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
    if (!s || !s.healthWeights) return null;
    var hw = s.healthWeights; var he = s.healthExclude || {};
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
    if (!def || !def.stdFields) return fields;
    for (var i = 0; i < def.stdFields.length; i++) {
      var imp = s.stdFieldConfig[def.stdFields[i].key] || 'excluded';
      if (imp !== 'excluded') fields.push(def.stdFields[i].key);
    }
    return fields;
  }

  function getSettings(entity) {
    var s = get(entity || 'company');
    if (!s || !s.stdFieldConfig) return { completenessFields: [], stdFieldConfig: {}, udefFieldConfig: {}, integrityConfig: {}, engagementConfig: {}, healthWeights: {}, healthExclude: {}, pipelineType: 'sale', dqWeights: { completeness: 70, udef: 30, quality: 0 }, qualityIssueFields: [], engagementWeights: [] };
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
  // SETTINGS MODAL — Discord-style overlay
  // ============================================================
  var _smActiveSection = 'company';

  // SVG icons for modal sidebar
  var SM_ICONS = {
    company: '<svg viewBox="0 0 32 32" fill="none"><path d="M30 26h-2V12a2 2 0 0 0-2-2h-8V4a2 2 0 0 0-3.11-1.665L4.89 9A2 2 0 0 0 4 10.668V26H2a1 1 0 0 0 0 2h28a1 1 0 0 0 0-2m-4-14v14h-8V12zM6 10.668 16 4v22H6z" fill="currentColor"/></svg>',
    contact: '<svg viewBox="0 0 32 32" fill="none"><path d="M28.865 26.5c-1.904-3.291-4.837-5.651-8.261-6.77a9 9 0 1 0-9.208 0c-3.424 1.117-6.357 3.477-8.261 6.77a1 1 0 1 0 1.731 1C7.221 23.43 11.384 21 16 21s8.779 2.43 11.134 6.5a.999.999 0 1 0 1.731-1M9 12a7 7 0 1 1 7 7 7.01 7.01 0 0 1-7-7" fill="currentColor"/></svg>',
    sale: '<svg viewBox="0 0 32 32" fill="none"><path d="M16 3a13 13 0 1 0 13 13A13.013 13.013 0 0 0 16 3m0 24a11 11 0 1 1 11-11 11.01 11.01 0 0 1-11 11m5-8.5a3.5 3.5 0 0 1-3.5 3.5H17v1a1 1 0 0 1-2 0v-1h-2a1 1 0 0 1 0-2h4.5a1.5 1.5 0 1 0 0-3h-3a3.5 3.5 0 1 1 0-7h.5V9a1 1 0 0 1 2 0v1h2a1 1 0 0 1 0 2h-4.5a1.5 1.5 0 1 0 0 3h3a3.5 3.5 0 0 1 3.5 3.5" fill="currentColor"/></svg>',
    project: '<svg viewBox="0 0 32 32" fill="none"><path d="M21 19a1 1 0 0 1-1 1h-8a1 1 0 0 1 0-2h8a1 1 0 0 1 1 1m-1-5h-8a1 1 0 0 0 0 2h8a1 1 0 0 0 0-2m7-8v21a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6a2 2 0 0 1 2-2h4.533a5.99 5.99 0 0 1 8.935 0H25a2 2 0 0 1 2 2M12 8h8a4 4 0 1 0-8 0m13-2h-3.344A6 6 0 0 1 22 8v1a1 1 0 0 1-1 1H11a1 1 0 0 1-1-1V8c0-.681.116-1.358.344-2H7v21h18z" fill="currentColor"/></svg>',
    levels: '<svg viewBox="0 0 32 32" fill="none"><path d="M5 11h4.125a4 4 0 0 0 7.75 0H27a1 1 0 0 0 0-2H16.875a4 4 0 0 0-7.75 0H5a1 1 0 0 0 0 2m8-3a2 2 0 1 1 0 4 2 2 0 0 1 0-4m14 13h-2.125a4 4 0 0 0-7.75 0H5a1 1 0 0 0 0 2h12.125a4 4 0 0 0 7.75 0H27a1 1 0 0 0 0-2m-6 3a2 2 0 1 1 0-4 2 2 0 0 1 0 4" fill="currentColor"/></svg>',
    activity: '<svg viewBox="0 0 32 32" fill="none"><path d="M28.707 9.707 12.707 25.707a1 1 0 0 1-1.414 0l-7-7a1 1 0 1 1 1.414-1.414L12 23.586 27.293 8.293a1 1 0 1 1 1.414 1.414z" fill="currentColor"/></svg>'
  };

  // Inject the small style block for intake-locked rows once.
  // da-styles.css is shipped separately; keeping these self-contained here means
  // the lock styling ships with the same asset and needs no second file.
  function ensureIntakeStyles() {
    if (document.getElementById('daIntakeStyles')) return;
    var css = ''
      + '.s-intake-badge{display:inline-block;margin-left:8px;padding:1px 7px;border-radius:10px;'
      + 'font-size:10px;font-weight:700;letter-spacing:.04em;text-transform:uppercase;vertical-align:middle;'
      + 'background:rgba(99,102,241,.14);color:#6366f1;border:1px solid rgba(99,102,241,.35);}'
      + '.s-imp-btn.s-imp-locked{cursor:not-allowed;}'
      + '.s-imp-btn.s-imp-locked:not(.act-req):not(.act-norm):not(.act-excl):not(.act-high):not(.act-med):not(.act-low){opacity:.45;}'
      + '.s-integrity-check input:disabled{cursor:not-allowed;}'
      + '.st-list{display:flex;flex-direction:column;gap:4px;margin-top:6px;}'
      + '.st-row{padding:14px 0;}'
      + '.st-row + .st-row{border-top:0.5px solid #EDE8DD;}'
      + '.st-name{font-size:14px;font-weight:500;color:#1C1C1E;}'
      + '.st-control{display:flex;align-items:center;gap:16px;margin-top:14px;margin-bottom:22px;}'
      + '.st-track-wrap{position:relative;flex:1;height:18px;}'
      + '.st-track{position:absolute;top:50%;left:0;right:0;transform:translateY(-50%);height:4px;border-radius:999px;background:#D8D2C4;z-index:0;}'
      + '.st-fill{position:absolute;top:50%;left:0;transform:translateY(-50%);height:4px;border-radius:999px;background:#06423E;z-index:1;}'
      + '.st-cur{position:absolute;top:50%;transform:translate(-50%,-50%);z-index:2;}'
      + '.st-cur b{display:block;width:16px;height:16px;border-radius:50%;background:#06423E;}'
      + '.st-curlab{position:absolute;top:22px;left:50%;transform:translateX(-50%);display:flex;align-items:center;gap:6px;white-space:nowrap;}'
      + '.st-curlab i{font-style:normal;font-size:11px;font-weight:500;color:#9C9388;}'
      + '.st-curlab em{font-style:normal;font-size:11px;font-weight:500;color:#5F5E5A;background:#F1EFE8;border-radius:999px;padding:2px 8px;}'
      + '.st-thumb{position:absolute;top:50%;transform:translate(-50%,-50%);z-index:3;pointer-events:none;}'
      + '.st-thumb b{display:block;width:18px;height:18px;border-radius:50%;background:#FFFFFF;border:2px solid #06423E;box-sizing:border-box;}'
      + '.st-track-wrap:focus-within .st-thumb b{box-shadow:0 0 0 3px rgba(6,66,62,.18);}'
      + '.st-slider{-webkit-appearance:none !important;appearance:none !important;position:absolute;inset:0;width:100%;height:100%;background:transparent !important;border:none !important;box-shadow:none !important;margin:0;padding:0;z-index:4;cursor:pointer;}'
      + '.st-slider::-webkit-slider-runnable-track{height:18px !important;background:transparent !important;border:none !important;}'
      + '.st-slider::-webkit-slider-thumb{-webkit-appearance:none !important;appearance:none !important;width:18px;height:18px;background:transparent !important;border:none !important;box-shadow:none !important;cursor:pointer;}'
      + '.st-slider::-moz-range-track{height:18px !important;background:transparent !important;border:none !important;}'
      + '.st-slider::-moz-range-thumb{width:18px;height:18px;background:transparent !important;border:none !important;}'
      + '.st-readout{flex:none;width:60px;text-align:right;}'
      + '.st-readout .st-val{font-size:20px;font-weight:500;color:#06423E;font-variant-numeric:tabular-nums;}'
      + '.st-readout .st-pct{font-size:12px;color:#7A7268;margin-left:1px;}';
    var st = document.createElement('style');
    st.id = 'daIntakeStyles';
    st.textContent = css;
    document.head.appendChild(st);
  }

  function openSettings() {
    ensureIntakeStyles();
    var overlay = document.getElementById('smOverlay');
    if (!overlay) {
      overlay = document.createElement('div');
      overlay.id = 'smOverlay';
      overlay.className = 'sm-overlay';
      overlay.onclick = function(e) { if (e.target === overlay) closeSettings(); };
      overlay.innerHTML = buildModalHTML();
      document.body.appendChild(overlay);
    }
    renderAllSections();
    _smActiveSection = 'company';
    showSettingsSection('company');
    var unsaved = document.getElementById('smUnsaved');
    if (unsaved) unsaved.classList.remove('show');
    setTimeout(function() { overlay.classList.add('open'); }, 10);
  }

  function closeSettings() {
    var overlay = document.getElementById('smOverlay');
    if (overlay) overlay.classList.remove('open');
  }

  function buildModalHTML() {
    var h = '<div class="sm-modal">';
    // Header
    h += '<div class="sm-header"><h2>Settings</h2>';
    h += '<button class="sm-close" onclick="daSettings.closeSettings()"><svg width="16" height="16" viewBox="0 0 16 16" fill="none"><path d="M12 4L4 12M4 4l8 8" stroke="currentColor" stroke-width="1.8" stroke-linecap="round"/></svg></button></div>';
    // Body
    h += '<div class="sm-body">';
    // Sidebar
    h += '<div class="sm-sidebar">';
    h += '<div class="sm-sb-label">Scoring</div>';
    for (var i = 0; i < ENTITY_ORDER.length; i++) {
      var ek = ENTITY_ORDER[i];
      h += '<div class="sm-sb-item' + (i === 0 ? ' active' : '') + '" onclick="daSettings.showSection(\'' + ek + '\',this)">' + SM_ICONS[ek] + ' ' + ENTITY_DEFS[ek].label + '</div>';
    }
    h += '<div class="sm-sb-label">CRM Momentum</div>';
    h += '<div class="sm-sb-item" onclick="daSettings.showSection(\'mm-thresholds\',this)">' + SM_ICONS.levels + ' User Levels</div>';
    h += '<div class="sm-sb-item" onclick="daSettings.showSection(\'mm-activity\',this)">' + SM_ICONS.activity + ' Activity Types</div>';
    h += '</div>';
    // Content
    h += '<div class="sm-content" id="smContent">';
    for (var i = 0; i < ENTITY_ORDER.length; i++) {
      h += '<div class="sm-section" id="sm-' + ENTITY_ORDER[i] + '"></div>';
    }
    h += '<div class="sm-section" id="sm-mm-thresholds"></div>';
    h += '<div class="sm-section" id="sm-mm-activity"></div>';
    h += '</div></div>';
    // Footer
    h += '<div class="sm-footer">';
    h += '<button class="btn-settings-save" onclick="daSettings.doSave()">Apply &amp; Recalculate</button>';
    h += '<button class="btn-settings-reset" onclick="daSettings.doReset()">Reset to Defaults</button>';
    h += '<div class="sm-unsaved" id="smUnsaved">\u25CF Unsaved changes</div>';
    h += '</div>';
    h += '</div>';
    return h;
  }

  function showSettingsSection(id, el) {
    _smActiveSection = id;
    var secs = document.querySelectorAll('.sm-section');
    for (var i = 0; i < secs.length; i++) secs[i].classList.remove('active');
    var items = document.querySelectorAll('.sm-sb-item');
    for (var i = 0; i < items.length; i++) items[i].classList.remove('active');
    var sec = document.getElementById('sm-' + id);
    if (sec) sec.classList.add('active');
    // If no el passed (e.g. from openSettings), find matching sidebar item
    if (!el) {
      for (var i = 0; i < items.length; i++) {
        if (items[i].getAttribute('onclick') && items[i].getAttribute('onclick').indexOf("'" + id + "'") >= 0) {
          el = items[i]; break;
        }
      }
    }
    if (el) el.classList.add('active');
  }

  function renderAllSections() {
    // Entity scoring sections
    for (var i = 0; i < ENTITY_ORDER.length; i++) {
      var ek = ENTITY_ORDER[i];
      var sec = document.getElementById('sm-' + ek);
      if (sec) {
        var h = '<div class="sm-title">' + ENTITY_DEFS[ek].label + ' Scoring</div>';
        h += '<div class="sm-subtitle">Configure how the ' + ENTITY_DEFS[ek].label + ' health score is calculated.</div>';
        h += renderEntityPanel(ek);
        sec.innerHTML = h;
      }
    }
    // Momentum sections
    renderModalMomentum();
    renderModalActivity();
  }

  function renderModalMomentum() {
    var m = _s._momentum || clone(MOMENTUM_DEFAULTS);
    var sec = document.getElementById('sm-mm-thresholds');
    if (!sec) return;
    var h = '<div class="sm-title">User Levels</div>';
    h += '<div class="sm-subtitle">Classify users by their average monthly activity count.</div>';
    h += '<div class="sm-card"><h3>Level Thresholds</h3>';
    h += '<div class="sm-desc">Users are classified based on their average activities per month in the selected period.</div>';
    h += '<div class="sm-threshold-row"><div class="sm-threshold-label" style="color:#2e7d32;font-weight:600">Power User</div>';
    h += '<span style="color:var(--so-text-muted)">\u2265</span>';
    h += '<input type="number" class="sm-threshold-input" id="mmPowerThreshold" value="' + m.powerThreshold + '" min="1" onchange="daSettings.markUnsaved()">';
    h += '<div class="sm-threshold-desc">activities / month (avg)</div></div>';
    h += '<div class="sm-threshold-row"><div class="sm-threshold-label" style="color:#1565c0;font-weight:600">Regular User</div>';
    h += '<span style="color:var(--so-text-muted)">\u2265</span>';
    h += '<input type="number" class="sm-threshold-input" id="mmRegularThreshold" value="' + m.regularThreshold + '" min="1" onchange="daSettings.markUnsaved()">';
    h += '<div class="sm-threshold-desc">activities / month (avg)</div></div>';
    h += '<div class="sm-threshold-row"><div class="sm-threshold-label" style="color:#f57c00;font-weight:600">Low Usage</div>';
    h += '<span style="color:var(--so-text-muted)">\u2265 1 activity in period</span></div>';
    h += '<div class="sm-threshold-row"><div class="sm-threshold-label dim" style="color:#c62828">Inactive</div>';
    h += '<span class="sm-threshold-desc" style="opacity:.5">0 activities</span></div>';
    h += '</div>';
    // Date field
    h += '<div class="sm-card"><h3>Date Field</h3>';
    h += '<div class="sm-desc">Which date field to use for counting activities.</div>';
    h += '<div class="sm-radio' + (m.dateField === 'activeDate' ? ' selected' : '') + '" onclick="daSettings.selectRadio(this,\'mmDateField\',\'activeDate\')">';
    h += '<input type="radio" name="mmDateField" value="activeDate"' + (m.dateField === 'activeDate' ? ' checked' : '') + '>';
    h += '<div><strong>Active Date</strong><div class="sm-radio-desc">When the activity takes place (recommended)</div></div></div>';
    h += '<div class="sm-radio' + (m.dateField === 'registered' ? ' selected' : '') + '" onclick="daSettings.selectRadio(this,\'mmDateField\',\'registered\')">';
    h += '<input type="radio" name="mmDateField" value="registered"' + (m.dateField === 'registered' ? ' checked' : '') + '>';
    h += '<div><strong>Registered</strong><div class="sm-radio-desc">When the activity was entered in CRM</div></div></div>';
    h += '</div>';
    sec.innerHTML = h;
  }

  function renderModalActivity() {
    var sec = document.getElementById('sm-mm-activity');
    if (!sec) return;
    // Try to get activity types from loaded overview data
    var types = null;
    if (typeof overviewData !== 'undefined' && overviewData['activities'] && overviewData['activities'].distributions) {
      var dists = overviewData['activities'].distributions;
      for (var i = 0; i < dists.length; i++) {
        if (dists[i].title === 'Activity Types') { types = dists[i].items; break; }
      }
    }
    if (!types) {
      // Auto-fetch activity overview data
      sec.innerHTML = '<div class="sm-title">Activity Types</div><div class="sm-subtitle">Choose which activity types to include in CRM Momentum analysis.</div><div class="sm-card"><h3>Include in Analysis</h3><div class="sm-desc">Loading activity types...</div></div>';
      if (typeof overviewUrl !== 'undefined' && typeof ajax === 'function') {
        var amp = String.fromCharCode(38);
        ajax(overviewUrl + amp + 'entity=activities', function(d) {
          if (d) {
            if (typeof overviewData !== 'undefined') overviewData['activities'] = d;
            renderModalActivityContent(sec);
          } else {
            sec.querySelector('.sm-desc').textContent = 'Could not load activity types.';
          }
        });
      }
      return;
    }
    renderModalActivityContent(sec);
  }

  function renderModalActivityContent(sec) {
    var types = null;
    if (typeof overviewData !== 'undefined' && overviewData['activities'] && overviewData['activities'].distributions) {
      var dists = overviewData['activities'].distributions;
      for (var i = 0; i < dists.length; i++) {
        if (dists[i].title === 'Activity Types') { types = dists[i].items; break; }
      }
    }
    var h = '<div class="sm-title">Activity Types</div>';
    h += '<div class="sm-subtitle">Choose which activity types to include in CRM Momentum analysis.</div>';
    h += '<div class="sm-card"><h3>Include in Analysis</h3>';
    h += '<div class="sm-desc">Uncheck types you want to exclude from activity counting and user level classification.</div>';
    if (types && types.length > 0) {
      var mmSettings = _s._momentum || clone(MOMENTUM_DEFAULTS);
      var excluded = mmSettings.excludedTypes || [];
      h += '<div class="sm-check-list">';
      for (var i = 0; i < types.length; i++) {
        var t = types[i];
        var isExcluded = excluded.indexOf(t.name) >= 0;
        var isOutlook = t.name.toLowerCase().indexOf('outlook') >= 0;
        var checked = isExcluded ? '' : (isOutlook && excluded.length === 0 ? '' : ' checked');
        // Default: exclude Outlook types on first run (when no excludedTypes saved yet)
        if (excluded.length === 0 && !isOutlook) checked = ' checked';
        h += '<label class="sm-check-item"><input type="checkbox"' + checked + ' data-type="' + t.name.replace(/"/g, '&quot;') + '" onchange="daSettings.markUnsaved()"> ' + t.name;
        if (t.count !== undefined) h += ' <span class="sm-check-note">(' + t.count.toLocaleString() + ')</span>';
        h += '</label>';
      }
      h += '</div>';
    } else {
      h += '<div style="padding:10px 0;font-size:.82rem;color:var(--so-text-muted);font-style:italic">No activity types found.</div>';
    }
    h += '</div>';
    sec.innerHTML = h;
  }

  function selectRadio(el, name, value) {
    var radios = document.querySelectorAll('input[name="' + name + '"]');
    for (var i = 0; i < radios.length; i++) {
      radios[i].checked = false;
      radios[i].closest('.sm-radio').classList.remove('selected');
    }
    var radio = el.querySelector('input[type="radio"]');
    if (radio) { radio.checked = true; radio.value = value; }
    el.classList.add('selected');
    markUnsaved();
  }

  function markUnsaved() {
    var el = document.getElementById('smUnsaved');
    if (el) el.classList.add('show');
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

    // ── SCORE TARGETS ──
    h += renderScoreTargets(entityKey, s);

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

  // renderMomentumPanel removed — now rendered inside modal by renderModalMomentum()

  function readMomentumFromDOM() {
    if (!_s._momentum) _s._momentum = clone(MOMENTUM_DEFAULTS);
    var pEl = document.getElementById('mmPowerThreshold');
    var rEl = document.getElementById('mmRegularThreshold');
    if (pEl) _s._momentum.powerThreshold = Math.max(1, parseInt(pEl.value) || 100);
    if (rEl) _s._momentum.regularThreshold = Math.max(1, parseInt(rEl.value) || 25);
    var radios = document.querySelectorAll('input[name="mmDateField"]');
    for (var i = 0; i < radios.length; i++) {
      if (radios[i].checked) { _s._momentum.dateField = radios[i].value; break; }
    }
    // Read activity type exclusions
    var checks = document.querySelectorAll('.sm-check-item input[data-type]');
    if (checks.length > 0) {
      var excluded = [];
      for (var i = 0; i < checks.length; i++) {
        if (!checks[i].checked) excluded.push(checks[i].getAttribute('data-type'));
      }
      _s._momentum.excludedTypes = excluded;
    }
  }

  function getMomentumSettings() {
    if (!_s) load();
    return _s._momentum || clone(MOMENTUM_DEFAULTS);
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

  // Score targets card: the level aimed for per score, with the current score as a
  // floor marker when the host has supplied it. Engine-owned and editable.
  function renderScoreTargets(ek, s) {
    var tg = s.scoreTargets || SCORE_TARGET_DEFAULTS;
    var cur = _currentScores[ek] || null;
    var h = '<div class="s-health-card"><h3>Score Targets</h3>';
    h += '<div class="settings-desc">The level you aim for on each score. Shown as a marker on the dashboard. You can raise it over time, and it cannot be set below the current score.</div>';
    h += '<div class="st-list">';
    for (var i = 0; i < SCORE_TARGET_DEFS.length; i++) {
      var d = SCORE_TARGET_DEFS[i];
      var val = (typeof tg[d.key] === 'number') ? tg[d.key] : (SCORE_TARGET_DEFAULTS[d.key] || 0);
      var curVal = (cur && typeof cur[d.key] === 'number') ? Math.round(cur[d.key]) : null;
      var pos = 'calc(9px + (100% - 18px) * ' + (val / 100) + ')';
      h += '<div class="st-row" data-entity="' + ek + '" data-score="' + d.key + '">';
      h += '<div class="st-name">' + d.label + '</div>';
      h += '<div class="st-control">';
      h += '<div class="st-track-wrap">';
      h += '<div class="st-track"></div>';
      h += '<div class="st-fill" style="width:' + pos + '"></div>';
      if (curVal != null) {
        h += '<div class="st-cur" style="left:calc(9px + (100% - 18px) * ' + (curVal / 100) + ')"><b></b><div class="st-curlab"><i>current score</i><em>' + curVal + '%</em></div></div>';
      }
      h += '<div class="st-thumb" style="left:' + pos + '"><b></b></div>';
      h += '<input type="range" class="st-slider" min="0" max="100" step="5" value="' + val + '"' + (curVal != null ? (' data-current="' + curVal + '"') : '') + ' oninput="daSettings.onTargetInput(this)" aria-label="' + d.label + ' target">';
      h += '</div>';
      h += '<div class="st-readout"><span class="st-val">' + val + '</span><span class="st-pct">%</span></div>';
      h += '</div></div>';
    }
    h += '</div></div>';
    return h;
  }

  // Live update from a target slider: clamp to the current score, move fill and
  // handle, update the readout, and write the value into the engine settings.
  function onTargetInput(slider) {
    var row = slider.closest('.st-row');
    if (!row) return;
    var ek = row.getAttribute('data-entity');
    var key = row.getAttribute('data-score');
    var floor = parseInt(slider.getAttribute('data-current') || '0', 10) || 0;
    var v = Math.round(parseInt(slider.value, 10) || 0);
    if (v < floor) { v = floor; slider.value = floor; }
    var pos = 'calc(9px + (100% - 18px) * ' + (v / 100) + ')';
    var fill = row.querySelector('.st-fill'); if (fill) fill.style.width = pos;
    var thumb = row.querySelector('.st-thumb'); if (thumb) thumb.style.left = pos;
    var out = row.querySelector('.st-val'); if (out) out.textContent = v;
    if (_s[ek]) {
      if (!_s[ek].scoreTargets) _s[ek].scoreTargets = {};
      _s[ek].scoreTargets[key] = v;
    }
    markUnsaved();
  }

  function renderIntegrityChecks(ek, s, def) {
    var cfg = s.integrityConfig;
    var h = '<div class="s-integrity-rows" data-entity="' + ek + '">';
    for (var i = 0; i < def.integrityChecks.length; i++) {
      var c = def.integrityChecks[i];
      var cc = cfg[c.key] || { enabled: false, weight: 'medium' };
      var locked = intakeOwns(ek, 'int', c.key);
      var disabled = !cc.enabled; var rowCls = disabled ? ' s-introw-disabled' : '';
      if (locked) rowCls += ' s-introw-intake';
      h += '<div class="s-integrity-row' + rowCls + '" data-key="' + c.key + '" data-level="' + cc.weight + '" data-enabled="' + cc.enabled + '">';
      if (locked) {
        h += '<label class="s-integrity-check"><input type="checkbox"' + (cc.enabled ? ' checked' : '') + ' disabled aria-disabled="true"></label>';
      } else {
        h += '<label class="s-integrity-check"><input type="checkbox"' + (cc.enabled ? ' checked' : '') + ' onchange="daSettings.toggleIntegrity(this,\'' + ek + '\',\'' + c.key + '\')"></label>';
      }
      h += '<div class="s-integrity-label"><span class="s-integrity-name">' + c.label + (locked ? ' <span class="s-intake-badge" title="Set in the intake form. Managed there and locked here.">Intake</span>' : '') + '</span><span class="s-integrity-desc">' + c.desc + (c.query ? '<span class="s-integrity-query" tabindex="0"><svg viewBox="0 0 16 16" fill="none" aria-hidden="true"><path d="M5.5 4L2 8l3.5 4M10.5 4L14 8l-3.5 4" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round"/></svg><span class="s-query-tip">' + c.query + '</span></span>' : '') + '</span></div>';
      h += locked ? hmlToggleLocked(cc.weight) : hmlToggle('int', ek, c.key, cc.weight);
      h += '</div>';
    }
    h += '</div>';
    return h;
  }

  // Static, non-interactive H/M/L toggle for an intake-owned integrity check
  function hmlToggleLocked(level) {
    var h = '<div class="s-fld-toggle">';
    var vals = ['high', 'medium', 'low'];
    var labels = ['High', 'Medium', 'Low'];
    var acts = { high: 'act-high', medium: 'act-med', low: 'act-low' };
    for (var i = 0; i < vals.length; i++) {
      var cls = level === vals[i] ? ' ' + acts[vals[i]] : '';
      h += '<button class="s-imp-btn s-imp-locked' + cls + '" disabled aria-disabled="true">' + labels[i] + '</button>';
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
    var locked = intakeOwns(ek, group, key);
    var h = '<div class="s-fld-row' + (imp === 'excluded' ? ' s-fld-excl' : '') + (locked ? ' s-fld-intake' : '') + '" data-grp="' + group + '" data-key="' + key + '" data-entity="' + ek + '">';
    h += '<div class="s-fld-name">' + label;
    if (type) h += '<span class="s-fld-type">' + type + '</span>';
    if (locked) h += '<span class="s-intake-badge" title="Set in the intake form. Managed there and locked here.">Intake</span>';
    h += '</div>';
    if (extra) h += '<div class="s-fld-extra">' + extra + '</div>';
    h += '<div class="s-fld-toggle">';
    if (locked) {
      h += impBtnLocked('Required', imp === 'required', 'act-req');
      h += impBtnLocked('Normal', imp === 'normal', 'act-norm');
      h += impBtnLocked('Excluded', imp === 'excluded', 'act-excl');
    } else {
      h += '<button class="s-imp-btn' + (imp === 'required' ? ' act-req' : '') + '" onclick="daSettings.setImp(\'' + group + '\',\'' + esc + '\',\'required\',this)">Required</button>';
      h += '<button class="s-imp-btn' + (imp === 'normal' ? ' act-norm' : '') + '" onclick="daSettings.setImp(\'' + group + '\',\'' + esc + '\',\'normal\',this)">Normal</button>';
      h += '<button class="s-imp-btn' + (imp === 'excluded' ? ' act-excl' : '') + '" onclick="daSettings.setImp(\'' + group + '\',\'' + esc + '\',\'excluded\',this)">Excluded</button>';
    }
    h += '</div></div>';
    return h;
  }

  // Static, non-interactive importance button for an intake-owned field
  function impBtnLocked(label, active, actCls) {
    return '<button class="s-imp-btn s-imp-locked' + (active ? ' ' + actCls : '') + '" disabled aria-disabled="true">' + label + '</button>';
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
    _activeEntity = ek;
    markUnsaved();
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
    markUnsaved();
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
    markUnsaved();
  }

  function toggleEngComp(cb, ek, key) {
    var row = cb.closest('.s-weight-row');
    if (cb.checked) { row.classList.remove('s-wrow-excl'); row.setAttribute('data-enabled', 'true'); }
    else { row.classList.add('s-wrow-excl'); row.setAttribute('data-enabled', 'false'); }
    recalcWeightGroup('eng', ek);
    markUnsaved();
  }

  function toggleIntegrity(cb, ek, key) {
    var row = cb.closest('.s-integrity-row');
    if (cb.checked) { row.classList.remove('s-introw-disabled'); row.setAttribute('data-enabled', 'true'); }
    else { row.classList.add('s-introw-disabled'); row.setAttribute('data-enabled', 'false'); }
    markUnsaved();
  }

  function setImp(group, key, imp, btn) {
    var btns = btn.parentElement.querySelectorAll('.s-imp-btn');
    for (var i = 0; i < btns.length; i++) btns[i].className = 's-imp-btn';
    btn.className = 's-imp-btn ' + (imp === 'required' ? 'act-req' : (imp === 'normal' ? 'act-norm' : 'act-excl'));
    var row = btn.closest('.s-fld-row');
    if (row) { if (imp === 'excluded') row.classList.add('s-fld-excl'); else row.classList.remove('s-fld-excl'); }
    markUnsaved();
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
    // Read ALL entity settings from modal DOM
    for (var i = 0; i < ENTITY_ORDER.length; i++) {
      readEntityFromDOM(ENTITY_ORDER[i]);
    }
    readMomentumFromDOM();
    saveLocal();
    // Recalculate scores for all entities that have loaded data
    for (var i = 0; i < ENTITY_ORDER.length; i++) {
      var ek = ENTITY_ORDER[i];
      if (typeof renderDQScore === 'function') renderDQScore(ek);
      if (typeof renderScoreBanner === 'function') renderScoreBanner(ek);
    }
    if (typeof renderCrossEntityFunnel === 'function' && typeof companyDetailData !== 'undefined' && companyDetailData) renderCrossEntityFunnel(companyDetailData);
    if (typeof renderCompanyDetails === 'function' && typeof companyDetailData !== 'undefined' && companyDetailData) renderCompanyDetails(companyDetailData);
    // Re-render momentum if data is available
    if (typeof renderMomentum === 'function' && typeof momentumData !== 'undefined' && momentumData) renderMomentum('activities', momentumData);
    var unsaved = document.getElementById('smUnsaved');
    if (unsaved) unsaved.classList.remove('show');
    closeSettings();
    // Save to server (async) with toast feedback
    saveToServer(function(ok) {
      if (ok) {
        toast('Settings saved and scores recalculated');
      } else if (_serverAvailable) {
        toast('Settings applied locally (server save failed)');
      } else {
        toast('Settings applied (create y_crm_health_setting table to enable sharing)');
      }
    });
  }

  function doReset() {
    _s = allDefaults(); saveLocal(); saveToServer();
    // Re-render modal sections if open
    var overlay = document.getElementById('smOverlay');
    if (overlay && overlay.classList.contains('open')) {
      renderAllSections();
      showSettingsSection(_smActiveSection);
    }
    var unsaved = document.getElementById('smUnsaved');
    if (unsaved) unsaved.classList.remove('show');
    toast('Settings reset to defaults');
  }

  function toast(msg) {
    var ex = document.getElementById('sToast'); if (ex) ex.remove();
    var t = document.createElement('div'); t.id = 'sToast'; t.className = 'settings-toast'; t.textContent = msg;
    document.body.appendChild(t);
    setTimeout(function(){ t.classList.add('show'); }, 10);
    setTimeout(function(){ t.classList.remove('show'); setTimeout(function(){ t.remove(); }, 300); }, 2500);
  }

  function init() {
    load(); // sync from localStorage (instant)
    // Then try server (async) — server data overwrites localStorage if found
    loadFromServer(function(loaded) {
      if (loaded) {
        // Re-render scores with server settings
        for (var i = 0; i < ENTITY_ORDER.length; i++) {
          var ek = ENTITY_ORDER[i];
          if (typeof renderDQScore === 'function') renderDQScore(ek);
          if (typeof renderScoreBanner === 'function') renderScoreBanner(ek);
        }
      }
    });
  }

  // Escape key handler
  function _onEscKey(e) {
    if (e.key === 'Escape') closeSettings();
  }
  document.addEventListener('keydown', _onEscKey);

  // ============================================================
  // PUBLIC API
  // ============================================================
  window.daSettings = {
    init: init,
    openSettings: openSettings,
    closeSettings: closeSettings,
    showSection: showSettingsSection,
    getSettings: getSettings,
    getConfig: get,
    getRules: getRules,
    getScoreTargets: function(entity) { var s = get(entity); return (s && s.scoreTargets) ? s.scoreTargets : null; },
    onTargetInput: onTargetInput,
    setCurrentScores: function(map) { _currentScores = map || {}; var ov = document.getElementById('smOverlay'); if (ov && ov.classList.contains('open')) { renderAllSections(); showSettingsSection(_smActiveSection); } },
    getScanState: function() { return _scanState; },
    saveScanState: saveScanState,
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
    selectRadio: selectRadio, markUnsaved: markUnsaved,
    openInfo: openInfo, closeInfo: closeInfo,
    doSave: doSave, doReset: doReset,
    getMomentumSettings: getMomentumSettings,
    ENTITY_DEFS: ENTITY_DEFS, ENTITY_ORDER: ENTITY_ORDER
  };
  load();
})();
