// ============================================================
// PIPELINE HUD v8 — app.js
// Architecture: IIFE modules — Utils, Data, PanelActivity,
//               PanelLeads, TabPipeline, TabRepView, TabDigest, App
// Adding a new tab: create a new IIFE block + register in App
// Data: Sales_Intelligence (event log) + Activity_Tracker
// Rep attribution: Rep_Name field (falls back to Created_By)
// Join key: Master_Journey_ID
// ============================================================

// ============================================================
// UTILS
// ============================================================
var Utils = (function() {

  var STAGES = [
    {num:'01', name:'Identified',     api:'01_Identified'},
    {num:'02', name:'Qualified',      api:'02_Qualified'},
    {num:'03', name:'Engaged',        api:'03_Engaged'},
    {num:'04', name:'Quot. Req.',     api:'04_Quotation_Required'},
    {num:'05', name:'Quoted',         api:'05_Quoted'},
    {num:'06', name:'Re-Quoted',      api:'06_Re_Quoted'},
    {num:'07', name:'Sample Req.',    api:'07_Sample_Required'},
    {num:'08', name:'Dispatched',     api:'08_Sample_Dispatched'},
    {num:'09', name:'Feedback Awt.',  api:'09_Sample_Feedback_Await'},
    {num:'10', name:'Approved',       api:'10_Sample_Approved'},
    {num:'11', name:'PO Awaited',     api:'11_PO_Awaited'},
    {num:'12', name:'Performa Inv.',  api:'12_Raise_Performa_Invoice'},
    {num:'13', name:'Closed Won',     api:'13_Closed_Won'},
    {num:'14', name:'Closed Lost',    api:'14_Closed_Lost'},
    {num:'15', name:'Sample Rej.',    api:'15_Sample_Rejected'}
  ];

  var TERMINAL = ['13_Closed_Won','14_Closed_Lost','15_Sample_Rejected'];

  var STAGE_NUM = {};
  STAGES.forEach(function(s) { STAGE_NUM[s.api] = parseInt(s.num); });

  var STAGE_MAP = {};
  STAGES.forEach(function(s) { STAGE_MAP[s.api] = s; });

  var THRESHOLDS = { dormancyDays: 7, reQuotedCap: 3, slowDays: 4 };

  var ORG = 'agronicfood'; // change if org name changes

  function safeId(obj) {
    if (!obj) return null;
    if (typeof obj === 'object' && obj.id) return String(obj.id);
    if (typeof obj === 'string' && obj.length > 0) return obj;
    return null;
  }

  function safeName(obj) {
    if (!obj) return null;
    if (typeof obj === 'object' && obj.name) return obj.name;
    return null;
  }

  function getRepName(si) {
    // Primary: Rep_Name field (set by Deluge function from Lead/Inquiry owner)
    if (si.Rep_Name && si.Rep_Name.trim && si.Rep_Name.trim() !== '') return si.Rep_Name.trim();
    // Fallback: Created_By
    var cb = si.Created_By;
    if (cb && typeof cb === 'object' && cb.name) return cb.name;
    if (cb && typeof cb === 'string' && cb.length > 0) return cb;
    return 'Unknown';
  }

  function getActivityOwner(a) {
    if (a.User && typeof a.User === 'object' && a.User.name) return a.User.name;
    if (a.Owner && typeof a.Owner === 'object' && a.Owner.name) return a.Owner.name;
    if (a.Created_By && typeof a.Created_By === 'object' && a.Created_By.name) return a.Created_By.name;
    return 'Unknown';
  }

  function getLeadName(si) {
    var n = safeName(si.Lead_Link);
    if (n && n.trim() !== '' && n !== 'NA NA') return n;
    var n2 = safeName(si.Inquiry_Link);
    if (n2 && n2.trim() !== '') return n2;
    return si.Master_Journey_ID || ('Lead ' + (si.id || ''));
  }

  function getMJID(si) {
    return si.Master_Journey_ID || '';
  }

  function getZohoUrl(si) {
    var leadId = safeId(si.Lead_Link);
    var inquiryId = safeId(si.Inquiry_Link);
    if (leadId) return 'https://crm.zoho.in/crm/' + ORG + '/tab/Leads/' + leadId;
    if (inquiryId) return 'https://crm.zoho.in/crm/' + ORG + '/tab/Potentials/' + inquiryId;
    return 'https://crm.zoho.in/crm/' + ORG + '/tab/Leads/';
  }

  function getActivityZohoUrl(a) {
    var leadId = safeId(a.Lead_Link);
    var inquiryId = safeId(a.Inquiry_Link);
    if (leadId) return 'https://crm.zoho.in/crm/' + ORG + '/tab/Leads/' + leadId;
    if (inquiryId) return 'https://crm.zoho.in/crm/' + ORG + '/tab/Potentials/' + inquiryId;
    return '';
  }

  function daysSince(date) {
    if (!date) return null;
    var d = new Date(date);
    if (isNaN(d)) return null;
    return Math.floor((new Date() - d) / (1000 * 60 * 60 * 24));
  }

  function daysUntil(date) {
    if (!date) return null;
    var d = new Date(date);
    if (isNaN(d)) return null;
    d.setHours(0,0,0,0);
    var now = new Date(); now.setHours(0,0,0,0);
    return Math.floor((d - now) / (1000 * 60 * 60 * 24));
  }

  function formatDate(date) {
    if (!date) return '—';
    var d = new Date(date);
    if (isNaN(d)) return '—';
    return d.toLocaleDateString('en-IN', { day:'numeric', month:'short', year:'numeric' });
  }

  function formatTime(date) {
    if (!date) return '';
    var d = new Date(date);
    if (isNaN(d)) return '';
    return d.toLocaleTimeString('en-IN', { hour:'2-digit', minute:'2-digit' });
  }

  function relativeDate(date) {
    var days = daysSince(date);
    if (days === null) return '—';
    if (days === 0) return 'Today';
    if (days === 1) return 'Yesterday';
    return days + ' days ago';
  }

  function dueDateLabel(date) {
    if (!date) return { text: 'No due date', cls: 'c-muted' };
    var days = daysUntil(date);
    if (days === null) return { text: 'No due date', cls: 'c-muted' };
    if (days < 0) return { text: Math.abs(days) + 'd overdue', cls: 'red' };
    if (days === 0) return { text: 'Due today', cls: 'orange' };
    if (days === 1) return { text: 'Due tomorrow', cls: 'blue' };
    return { text: 'Due in ' + days + 'd', cls: 'blue' };
  }

  function dormancyClass(days) {
    if (days === null || days === undefined) return 'c-muted';
    if (days > THRESHOLDS.dormancyDays) return 'c-red';
    if (days >= THRESHOLDS.slowDays) return 'c-orange';
    return 'c-green';
  }

  function velClass(days) {
    if (days === null) return 'c-muted';
    if (days <= 2) return 'c-green';
    if (days <= 5) return 'c-text';
    if (days <= 7) return 'c-orange';
    return 'c-red';
  }

  function convClass(pct) {
    if (pct === null) return 'c-muted';
    if (pct >= 75) return 'c-blue';
    if (pct >= 60) return 'c-green';
    if (pct >= 45) return 'c-orange';
    return 'c-red';
  }

  function typeBadge(type) {
    var t = (type || '').toLowerCase();
    if (t === 'call') return '<span class="type-badge tb-call">CALL</span>';
    if (t === 'message') return '<span class="type-badge tb-msg">MSG</span>';
    if (t === 'meeting') return '<span class="type-badge tb-meet">MEET</span>';
    if (t === 'task') return '<span class="type-badge tb-task">TASK</span>';
    if (t === 'demo') return '<span class="type-badge tb-demo">DEMO</span>';
    if (t === 'site visit') return '<span class="type-badge tb-meet">VISIT</span>';
    return '<span class="type-badge tb-task">' + (type || 'ACT').toUpperCase().substring(0,4) + '</span>';
  }

  function isCallType(type) {
    var t = (type || '').toLowerCase();
    return t === 'call' || t === 'demo' || t === 'meeting' || t === 'site visit';
  }

  function e(str) {
    if (!str) return '';
    return String(str)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;');
  }

  return {
    STAGES: STAGES,
    TERMINAL: TERMINAL,
    STAGE_NUM: STAGE_NUM,
    STAGE_MAP: STAGE_MAP,
    THRESHOLDS: THRESHOLDS,
    safeId: safeId,
    safeName: safeName,
    getRepName: getRepName,
    getActivityOwner: getActivityOwner,
    getLeadName: getLeadName,
    getMJID: getMJID,
    getZohoUrl: getZohoUrl,
    getActivityZohoUrl: getActivityZohoUrl,
    daysSince: daysSince,
    daysUntil: daysUntil,
    formatDate: formatDate,
    formatTime: formatTime,
    relativeDate: relativeDate,
    dueDateLabel: dueDateLabel,
    dormancyClass: dormancyClass,
    velClass: velClass,
    convClass: convClass,
    typeBadge: typeBadge,
    isCallType: isCallType,
    e: e
  };
})();

// ============================================================
// DATA
// Fetches once. All tabs read from this shared pool.
// Adding a new tab: read from Data.allSI, Data.activeSI etc.
// ============================================================
var Data = (function() {

  var allSI = [];
  var activeSI = [];
  var allActivities = [];
  var allCalls = [];
  var allTasks = [];

  // State
  var currentRep = 'all';
  var currentPeriod = 'week';
  var currentView = 'ops';

  // Rep name -> actual name map (built from SI data)
  var repNames = {};

  function setStatus(msg) {
    var el = document.getElementById('status-bar');
    if (!el) return;
    if (!msg) { el.classList.add('hidden'); return; }
    el.classList.remove('hidden');
    el.textContent = msg;
  }

  function fetchAllPages(entity, callback) {
    var all = [];
    function fetchPage(page) {
      ZOHO.CRM.API.getAllRecords({
        Entity: entity,
        sort_order: 'asc',
        per_page: 200,
        page: page
      }).then(function(r) {
        var data = r.data || [];
        all = all.concat(data);
        if (data.length === 200) { fetchPage(page + 1); }
        else { callback(null, all); }
      }).catch(function(e) {
        console.warn('fetchAllPages failed for ' + entity, e);
        callback(null, all);
      });
    }
    fetchPage(1);
  }

  function load(callback) {
    setStatus('Loading Sales Intelligence...');
    fetchAllPages('Sales_Intelligence', function(e1, si) {
      allSI = si || [];

      // Build rep name map from SI data
      allSI.forEach(function(r) {
        var name = Utils.getRepName(r);
        if (name && name !== 'Unknown') repNames[name] = true;
      });

      // activeSI: Current_Stage = true AND not terminal stage
      activeSI = allSI.filter(function(r) {
        return r.Current_Stage === true &&
               Utils.TERMINAL.indexOf(r.Pipeline) === -1;
      });

      setStatus('Loading Activities...');
      fetchAllPages('Activity_Tracker', function(e2, acts) {
        allActivities = acts || [];
        setStatus('Loading Calls...');
        fetchAllPages('Calls', function(e3, calls) {
          allCalls = calls || [];
          setStatus('Loading Tasks...');
          fetchAllPages('Tasks', function(e4, tasks) {
            allTasks = tasks || [];
            setStatus('');
            callback();
          });
        });
      });
    });
  }

  function getFilteredActiveSI() {
    var base = activeSI;

    // Period filter: journey start must be within selected period
    if (currentPeriod !== 'all' && currentView !== 'mgmt') {
      var journeyStartMap = {};
      allSI.forEach(function(r) {
        var mjid = r.Master_Journey_ID;
        if (!mjid) return;
        var t = new Date(r.Created_Time);
        if (isNaN(t)) return;
        if (!journeyStartMap[mjid] || t < journeyStartMap[mjid]) journeyStartMap[mjid] = t;
      });
      var now = new Date();
      var start = new Date();
      if (currentPeriod === 'today') { start.setHours(0,0,0,0); }
      else if (currentPeriod === 'week') { start.setDate(now.getDate() - 7); }
      else if (currentPeriod === 'month') { start.setMonth(now.getMonth() - 1); }
      else if (currentPeriod === 'quarter') { start.setMonth(now.getMonth() - 3); }
      base = base.filter(function(r) {
        var mjid = r.Master_Journey_ID;
        if (!mjid) return false;
        var jStart = journeyStartMap[mjid];
        return jStart && jStart >= start && jStart <= now;
      });
    }

    // Rep filter
    if (currentRep !== 'all' && currentView !== 'mgmt') {
      base = base.filter(function(r) {
        return Utils.getRepName(r) === currentRep;
      });
    }

    return base;
  }

  function getActivitiesForSI(si) {
    var result = { open: 0, closed: 0, calls: 0, tasks: 0, overdue: 0 };
    var mjid = si.Master_Journey_ID;
    var siLeadId = Utils.safeId(si.Lead_Link);

    if (allActivities.length > 0) {
      allActivities.forEach(function(a) {
        var matched = (mjid && a.Master_Journey_ID && a.Master_Journey_ID === mjid);
        if (!matched && siLeadId) {
          var aLeadId = Utils.safeId(a.Lead_Link);
          if (aLeadId && aLeadId === siLeadId) matched = true;
        }
        if (!matched) return;
        var isClosed = a.Activity_Status === 'Completed';
        if (isClosed) { result.closed++; }
        else {
          result.open++;
          var due = Utils.daysUntil(a.Due_Date);
          if (due !== null && due < 0) result.overdue++;
        }
        if (Utils.isCallType(a.Activity_Type)) result.calls++;
        else result.tasks++;
      });
      return result;
    }

    // Fallback
    var ids = {};
    if (siLeadId) ids[siLeadId] = true;
    var accountId = Utils.safeId(si.Accounts);
    // NOTE: Account ID intentionally excluded from fallback to prevent
    // cross-contamination between multiple Inquiries on same Account
    var inquiryId = Utils.safeId(si.Inquiry_Link);
    if (inquiryId) ids[inquiryId] = true;

    allCalls.forEach(function(c) {
      var whoId = Utils.safeId(c.Who_Id);
      if (whoId && ids[whoId]) {
        if (c.Outgoing_Call_Status === 'Completed' && (c.Duration_in_seconds || 0) > 0) {
          result.closed++; result.calls++;
        } else {
          result.open++; result.calls++;
          if (Utils.daysUntil(c.Scheduled_Time) < 0) result.overdue++;
        }
      }
    });

    allTasks.forEach(function(t) {
      var whatId = Utils.safeId(t.What_Id);
      if (whatId && ids[whatId]) {
        if (t.Status === 'Completed') { result.closed++; result.tasks++; }
        else {
          result.open++; result.tasks++;
          if (Utils.daysUntil(t.Due_Date) < 0) result.overdue++;
        }
      }
    });

    return result;
  }

  function getLastActivity(si) {
    var last = null;
    var lastRecord = null;
    var mjid = si.Master_Journey_ID;
    var siLeadId = Utils.safeId(si.Lead_Link);

    function check(date, record) {
      if (!date) return;
      var dt = new Date(date);
      if (isNaN(dt)) return;
      if (!last || dt > last) { last = dt; lastRecord = record; }
    }

    if (allActivities.length > 0 && (mjid || siLeadId)) {
      allActivities.forEach(function(a) {
        var matched = (mjid && a.Master_Journey_ID && a.Master_Journey_ID === mjid);
        if (!matched && siLeadId) {
          var aLeadId = Utils.safeId(a.Lead_Link);
          if (aLeadId && aLeadId === siLeadId) matched = true;
        }
        if (matched && a.Activity_Status === 'Completed') {
          check(a.Modified_Time || a.Created_Time, a);
        }
      });
    }

    if (!last) check(si.Created_Time, null);
    return { date: last, record: lastRecord };
  }

  function getDormancyDays(si) {
    var la = getLastActivity(si);
    if (!la.date) return null;
    return Utils.daysSince(la.date);
  }

  function getOpenActivitiesForSI(si) {
    var mjid = si.Master_Journey_ID;
    var siLeadId = Utils.safeId(si.Lead_Link);
    var result = [];
    allActivities.forEach(function(a) {
      if (a.Activity_Status === 'Completed') return;
      var matched = (mjid && a.Master_Journey_ID && a.Master_Journey_ID === mjid);
      if (!matched && siLeadId) {
        var aLeadId = Utils.safeId(a.Lead_Link);
        if (aLeadId && aLeadId === siLeadId) matched = true;
      }
      if (matched) result.push(a);
    });
    return result;
  }

  function getClosedTodayActivitiesForSI(si) {
    var mjid = si.Master_Journey_ID;
    var siLeadId = Utils.safeId(si.Lead_Link);
    var today = new Date(); today.setHours(0,0,0,0);
    var result = [];
    allActivities.forEach(function(a) {
      if (a.Activity_Status !== 'Completed') return;
      var closed = new Date(a.Modified_Time || a.Created_Time);
      if (isNaN(closed) || closed < today) return;
      var matched = (mjid && a.Master_Journey_ID && a.Master_Journey_ID === mjid);
      if (!matched && siLeadId) {
        var aLeadId = Utils.safeId(a.Lead_Link);
        if (aLeadId && aLeadId === siLeadId) matched = true;
      }
      if (matched) result.push(a);
    });
    return result;
  }

  function getAddedTodayForStage(stageApi) {
    var today = new Date(); today.setHours(0,0,0,0);
    return activeSI.filter(function(r) {
      if (r.Pipeline !== stageApi) return false;
      var created = new Date(r.Created_Time);
      return !isNaN(created) && created >= today;
    });
  }

  function getClosedTodayForStage(stageApi) {
    var today = new Date(); today.setHours(0,0,0,0);
    // Leads in this stage that have at least one activity closed today
    return getFilteredActiveSI().filter(function(r) {
      if (r.Pipeline !== stageApi) return false;
      return getClosedTodayActivitiesForSI(r).length > 0;
    });
  }

  function getAvgVelocity(stageApi) {
    var recs = allSI.filter(function(r) {
      return r.Pipeline === stageApi && r.Current_Stage === false &&
             parseFloat(r.Duration_Minutes) > 0;
    });
    if (recs.length < 5) return null; // min 5 exits to display
    var total = recs.reduce(function(a, r) { return a + parseFloat(r.Duration_Minutes); }, 0);
    return (total / recs.length) / (60 * 24); // days
  }

  function getConvRate(stageApi) {
    var stageNum = Utils.STAGE_NUM[stageApi];
    if (!stageNum) return null;
    var exits = allSI.filter(function(r) {
      return r.Pipeline === stageApi && r.Current_Stage === false;
    });
    if (exits.length < 5) return null;
    var forward = exits.filter(function(r) {
      var toNum = Utils.STAGE_NUM[r.From_Stage];
      // Forward = moved to a higher stage number
      // This checks the stage they came FROM was this stage (From_Stage on next record)
      return true; // simplified — count all non-regression exits
    });
    // Count exits where pipeline stage number increased
    var fwdExits = allSI.filter(function(r) {
      if (r.From_Stage !== stageApi) return false;
      var toNum = Utils.STAGE_NUM[r.Pipeline];
      return toNum && toNum > stageNum;
    });
    var allExitsFromStage = allSI.filter(function(r) { return r.From_Stage === stageApi; });
    if (allExitsFromStage.length < 5) return null;
    return Math.round((fwdExits.length / allExitsFromStage.length) * 100);
  }

  function getRepNames() {
    return Object.keys(repNames).sort();
  }

  function getAllSI() { return allSI; }
  function getActiveSI() { return activeSI; }
  function getAllActivities() { return allActivities; }
  function getAllCalls() { return allCalls; }
  function getAllTasks() { return allTasks; }

  return {
    load: load,
    getAllSI: getAllSI,
    getActiveSI: getActiveSI,
    getAllActivities: getAllActivities,
    getAllCalls: getAllCalls,
    getAllTasks: getAllTasks,
    getFilteredActiveSI: getFilteredActiveSI,
    getActivitiesForSI: getActivitiesForSI,
    getLastActivity: getLastActivity,
    getDormancyDays: getDormancyDays,
    getOpenActivitiesForSI: getOpenActivitiesForSI,
    getClosedTodayActivitiesForSI: getClosedTodayActivitiesForSI,
    getAddedTodayForStage: getAddedTodayForStage,
    getClosedTodayForStage: getClosedTodayForStage,
    getAvgVelocity: getAvgVelocity,
    getConvRate: getConvRate,
    getRepNames: getRepNames,
    get currentRep() { return currentRep; },
    set currentRep(v) { currentRep = v; },
    get currentPeriod() { return currentPeriod; },
    set currentPeriod(v) { currentPeriod = v; },
    get currentView() { return currentView; },
    set currentView(v) { currentView = v; }
  };
})();

// ============================================================
// PANEL — Activity Queue
// Opens when clicking Open Acts or Overdue column
// ============================================================
var PanelActivity = (function() {

  var _stage = null;
  var _filter = 'all'; // 'all' | 'overdue'
  var _sort = 'urgency';
  var _items = []; // { si, activity }

  function buildItems(stageApi, filter) {
    var items = [];
    var filtered = Data.getFilteredActiveSI().filter(function(r) { return r.Pipeline === stageApi; });
    filtered.forEach(function(si) {
      var openActs = Data.getOpenActivitiesForSI(si);
      openActs.forEach(function(a) {
        if (filter === 'overdue') {
          var due = Utils.daysUntil(a.Due_Date);
          if (due === null || due >= 0) return;
        }
        items.push({ si: si, activity: a });
      });
    });
    return items;
  }

  function sortItems(items, sort) {
    return items.slice().sort(function(a, b) {
      if (sort === 'urgency') {
        var da = Utils.daysUntil(a.activity.Due_Date);
        var db = Utils.daysUntil(b.activity.Due_Date);
        da = (da === null) ? 999 : da;
        db = (db === null) ? 999 : db;
        return da - db;
      }
      if (sort === 'due') {
        var ta = new Date(a.activity.Due_Date || '2099');
        var tb = new Date(b.activity.Due_Date || '2099');
        return ta - tb;
      }
      if (sort === 'name') {
        return Utils.getLeadName(a.si).localeCompare(Utils.getLeadName(b.si));
      }
      return 0;
    });
  }

  function cardClass(a) {
    var due = Utils.daysUntil(a.Due_Date);
    if (due !== null && due < 0) return 'act-card overdue';
    if (due === 0) return 'act-card today';
    return 'act-card future';
  }

  function renderCard(si, a) {
    var leadName = Utils.getLeadName(si);
    var mjid = Utils.getMJID(si);
    var zohoUrl = Utils.getZohoUrl(si);
    var dueInfo = Utils.dueDateLabel(a.Due_Date);
    var dormDays = Data.getDormancyDays(si);
    var lastAct = Data.getLastActivity(si);
    var stageInfo = Utils.STAGE_MAP[si.Pipeline] || { name: si.Pipeline };
    var inStageDays = Utils.daysSince(si.Created_Time);

    // Last activity block
    var lastActHtml = '';
    if (lastAct.date) {
      var lastRec = lastAct.record;
      var platform = lastRec ? (lastRec.Platform || '') : '';
      var note = lastRec ? (lastRec.Activity_Notes || '') : '';
      var response = lastRec ? (lastRec.Response_Received || '') : '';
      var dateStr = Utils.formatDate(lastAct.date) + ' · ' + Utils.relativeDate(lastAct.date);
      var dateCls = (dormDays !== null && dormDays > 7) ? 'c-red' : (dormDays !== null && dormDays >= 4) ? 'c-orange' : 'c-green';
      lastActHtml = '<div class="last-act">' +
        '<div class="la-lbl">Last act</div>' +
        '<div class="la-body">' +
          '<div class="la-date ' + dateCls + '">' + Utils.e(dateStr) + '</div>' +
          (platform ? '<div class="la-platform">via ' + Utils.e(platform) + '</div>' : '') +
          (note ? '<div class="la-note">' + Utils.e(note.substring(0, 120)) + (note.length > 120 ? '...' : '') + '</div>' : '') +
          (response ? '<div class="la-response c-muted">Response: <span class="c-text">' + Utils.e(response) + '</span></div>' : '') +
        '</div>' +
      '</div>';
    }

    // Phone
    var phone = '';
    if (si.Lead_Link && si.Lead_Link.phone) phone = si.Lead_Link.phone;
    var phoneHtml = phone ? '<div class="card-phone">📞 ' + Utils.e(phone) + '</div>' : '';

    var dormStr = (dormDays !== null) ? dormDays + 'd' : '—';
    var dormCls = Utils.dormancyClass(dormDays);
    var inStageStr = (inStageDays !== null) ? inStageDays + 'd' : '—';
    var inStageCls = (inStageDays !== null && inStageDays > 7) ? 'orange' : (inStageDays !== null && inStageDays > 14) ? 'red' : '';

    return '<div class="' + cardClass(a) + '">' +
      '<div class="card-top">' +
        Utils.typeBadge(a.Activity_Type) +
        '<span class="card-lead">' + Utils.e(leadName) + '</span>' +
        '<span class="card-mjid">' + Utils.e(mjid) + '</span>' +
      '</div>' +
      '<div class="card-company">' + Utils.e(Utils.safeName(si.Accounts) || '') + '</div>' +
      phoneHtml +
      '<div class="card-desc">' + Utils.e(a.Subject || a.Activity_Type || 'Activity') + '</div>' +
      '<div class="card-meta">' +
        '<div class="cm"><span class="cm-l">Due</span><span class="cm-v ' + dueInfo.cls + '">' + Utils.e(dueInfo.text) + '</span></div>' +
        '<div class="cm"><span class="cm-l">Owner</span><span class="cm-v">' + Utils.e(Utils.getActivityOwner(a)) + '</span></div>' +
        '<div class="cm"><span class="cm-l">In stage</span><span class="cm-v ' + inStageCls + '">' + inStageStr + '</span></div>' +
        '<div class="cm"><span class="cm-l">Dormant</span><span class="cm-v ' + dormCls.replace('c-','') + '">' + dormStr + '</span></div>' +
      '</div>' +
      lastActHtml +
      '<div class="card-footer">' +
        '<a class="open-zoho" href="' + zohoUrl + '" target="_blank">Open in Zoho ↗</a>' +
      '</div>' +
    '</div>';
  }

  function render() {
    var body = document.getElementById('panel-body');
    var emptyEl = document.getElementById('panel-empty');
    _items = sortItems(buildItems(_stage, _filter), _sort);

    if (_items.length === 0) {
      body.innerHTML = '';
      emptyEl.style.display = 'block';
      emptyEl.textContent = _filter === 'overdue' ? 'No overdue activities for this stage.' : 'No open activities for this stage.';
      return;
    }
    emptyEl.style.display = 'none';
    body.innerHTML = _items.map(function(item) {
      return renderCard(item.si, item.activity);
    }).join('');
  }

  function open(stageApi, filter) {
    _stage = stageApi;
    _filter = filter || 'all';
    _sort = 'urgency';

    var stageInfo = Utils.STAGE_MAP[stageApi] || { num: '', name: stageApi };
    var title = stageInfo.num + ' ' + stageInfo.name + ' — ' +
                (_filter === 'overdue' ? 'Overdue Activities' : 'Open Activities');

    var allOpen = buildItems(stageApi, 'all');
    var allOverdue = buildItems(stageApi, 'overdue');
    var sub = allOpen.length + ' open · ' + allOverdue.length + ' overdue · sorted by urgency';

    document.getElementById('panel-title').textContent = title;
    document.getElementById('panel-sub').textContent = sub;

    // Reset sort chips
    document.querySelectorAll('.sort-chip').forEach(function(c) { c.classList.remove('on'); });
    document.querySelector('.sort-chip[data-sort="urgency"]').classList.add('on');

    document.getElementById('panel-sort').style.display = 'flex';
    render();

    document.getElementById('overlay').classList.add('open');
    document.getElementById('side-panel').classList.add('open');
  }

  function onSort(sort) {
    _sort = sort;
    render();
  }

  return { open: open, onSort: onSort };
})();

// ============================================================
// PANEL — Lead Cards
// Opens when clicking Added Today or Closed Today column
// ============================================================
var PanelLeads = (function() {

  var _siList = [];
  var _type = 'added'; // 'added' | 'closed'

  function renderActivityCard(a) {
    var due = Utils.dueDateLabel(a.Due_Date);
    return '<div class="coa-item">' +
      Utils.typeBadge(a.Activity_Type) +
      '<span class="coa-desc">' + Utils.e(a.Subject || a.Activity_Type || 'Activity') + '</span>' +
      '<span class="coa-due cm-v ' + due.cls + '">' + Utils.e(due.text) + '</span>' +
    '</div>';
  }

  function renderClosedActivity(a) {
    var owner = Utils.getActivityOwner(a);
    var time = Utils.formatTime(a.Modified_Time || a.Created_Time);
    var response = a.Response_Received || '';
    return '<div class="coa-item">' +
      Utils.typeBadge(a.Activity_Type) +
      '<span class="coa-desc">' + Utils.e(a.Subject || a.Activity_Type || 'Activity') + '</span>' +
      '<span class="coa-due c-green">Done ' + Utils.e(time ? 'at ' + time : 'today') + '</span>' +
    '</div>' +
    (response ? '<div style="font-size:10px;color:var(--muted);margin-left:8px;margin-bottom:2px;">Response: <span style="color:var(--text);">' + Utils.e(response) + '</span> · ' + Utils.e(owner) + '</div>' : '');
  }

  function renderLeadCard(si) {
    var leadName = Utils.getLeadName(si);
    var mjid = Utils.getMJID(si);
    var zohoUrl = Utils.getZohoUrl(si);
    var dormDays = Data.getDormancyDays(si);
    var lastAct = Data.getLastActivity(si);
    var stageInfo = Utils.STAGE_MAP[si.Pipeline] || { name: si.Pipeline };
    var inStageDays = Utils.daysSince(si.Created_Time);
    var repName = Utils.getRepName(si);

    // Health chip
    var healthText, healthColor;
    if (dormDays === null || dormDays <= 3) { healthText = 'Active'; healthColor = 'var(--green)'; }
    else if (dormDays <= 7) { healthText = 'Slowing'; healthColor = 'var(--orange)'; }
    else { healthText = 'Dormant'; healthColor = 'var(--red)'; }

    var borderColor = (dormDays !== null && dormDays > 7) ? 'var(--red)' : (dormDays !== null && dormDays >= 4) ? 'var(--orange)' : 'var(--green)';

    // Open activities
    var openActs = Data.getOpenActivitiesForSI(si);
    var openActsHtml = '';
    if (openActs.length > 0) {
      openActsHtml = '<div class="card-open-acts">' +
        '<div class="coa-title">Open activities (' + openActs.length + ')</div>' +
        openActs.slice(0, 3).map(renderActivityCard).join('') +
        (openActs.length > 3 ? '<div style="font-size:10px;color:var(--muted);margin-top:3px;">+' + (openActs.length - 3) + ' more</div>' : '') +
      '</div>';
    }

    // Closed today activities (for closed-today panel)
    var closedActs = Data.getClosedTodayActivitiesForSI(si);
    var closedActsHtml = '';
    if (_type === 'closed' && closedActs.length > 0) {
      closedActsHtml = '<div class="card-open-acts">' +
        '<div class="coa-title">Completed today (' + closedActs.length + ')</div>' +
        closedActs.map(renderClosedActivity).join('') +
      '</div>';

      // Next open activity
      var nextOpen = openActs[0];
      if (nextOpen) {
        var due = Utils.dueDateLabel(nextOpen.Due_Date);
        closedActsHtml += '<div class="card-open-acts">' +
          '<div class="coa-title">Next open activity</div>' +
          renderActivityCard(nextOpen) +
        '</div>';
      }
    }

    // Last activity
    var lastActHtml = '';
    if (lastAct.date && _type !== 'closed') {
      var lastRec = lastAct.record;
      var note = lastRec ? (lastRec.Activity_Notes || '') : '';
      var platform = lastRec ? (lastRec.Platform || '') : '';
      var dateCls = (dormDays !== null && dormDays > 7) ? 'c-red' : (dormDays !== null && dormDays >= 4) ? 'c-orange' : 'c-green';
      lastActHtml = '<div class="last-act">' +
        '<div class="la-lbl">Last act</div>' +
        '<div class="la-body">' +
          '<div class="la-date ' + dateCls + '">' + Utils.e(Utils.formatDate(lastAct.date)) + ' · ' + Utils.e(Utils.relativeDate(lastAct.date)) + '</div>' +
          (platform ? '<div class="la-platform">via ' + Utils.e(platform) + '</div>' : '') +
          (note ? '<div class="la-note">' + Utils.e(note.substring(0, 100)) + (note.length > 100 ? '...' : '') + '</div>' : '') +
        '</div>' +
      '</div>';
    }

    var dormStr = (dormDays !== null) ? dormDays + 'd' : '—';
    var dormCls = Utils.dormancyClass(dormDays).replace('c-','');
    var enteredDate = Utils.formatDate(si.Created_Time);

    return '<div class="act-card" style="border-left-color:' + borderColor + ';">' +
      '<div class="card-top">' +
        '<span style="font-size:9px;font-weight:700;padding:2px 6px;border-radius:3px;background:rgba(0,0,0,.3);color:' + healthColor + ';flex-shrink:0;">' + healthText + '</span>' +
        '<span class="card-lead">' + Utils.e(leadName) + '</span>' +
        '<span class="card-mjid">' + Utils.e(mjid) + '</span>' +
      '</div>' +
      '<div class="card-company" style="margin-bottom:5px;">' +
        Utils.e(Utils.safeName(si.Accounts) || '') +
        (si.Buyer_Type ? ' · ' + Utils.e(si.Buyer_Type) : '') +
      '</div>' +
      '<div class="card-meta">' +
        '<div class="cm"><span class="cm-l">Owner</span><span class="cm-v">' + Utils.e(repName) + '</span></div>' +
        '<div class="cm"><span class="cm-l">In stage</span><span class="cm-v">' + (inStageDays !== null ? inStageDays + 'd' : '—') + '</span></div>' +
        '<div class="cm"><span class="cm-l">Dormant</span><span class="cm-v ' + dormCls + '">' + dormStr + '</span></div>' +
        '<div class="cm"><span class="cm-l">Entered</span><span class="cm-v c-muted">' + Utils.e(enteredDate) + '</span></div>' +
      '</div>' +
      (closedActsHtml || openActsHtml) +
      lastActHtml +
      '<div class="card-footer">' +
        '<a class="open-zoho" href="' + zohoUrl + '" target="_blank">Open in Zoho ↗</a>' +
      '</div>' +
    '</div>';
  }

  function open(stageApi, type) {
    _type = type; // 'added' | 'closed'
    var stageInfo = Utils.STAGE_MAP[stageApi] || { num: '', name: stageApi };

    if (type === 'added') {
      _siList = Data.getAddedTodayForStage(stageApi);
    } else {
      _siList = Data.getClosedTodayForStage(stageApi);
    }

    var title = stageInfo.num + ' ' + stageInfo.name + ' — ' +
                (type === 'added' ? 'Added Today' : 'Closed Today');
    var sub = _siList.length + ' lead' + (_siList.length !== 1 ? 's' : '');

    document.getElementById('panel-title').textContent = title;
    document.getElementById('panel-sub').textContent = sub;
    document.getElementById('panel-sort').style.display = 'none';

    var body = document.getElementById('panel-body');
    var emptyEl = document.getElementById('panel-empty');

    if (_siList.length === 0) {
      body.innerHTML = '';
      emptyEl.style.display = 'block';
      emptyEl.textContent = type === 'added' ? 'No leads added today in this stage.' : 'No activities closed today in this stage.';
    } else {
      emptyEl.style.display = 'none';
      body.innerHTML = _siList.map(renderLeadCard).join('');
    }

    document.getElementById('overlay').classList.add('open');
    document.getElementById('side-panel').classList.add('open');
  }

  return { open: open };
})();

// ============================================================
// TAB — PIPELINE
// Adding a new tab: copy this IIFE pattern, change the module name
// ============================================================
var TabPipeline = (function() {

  var _currentView = 'ops'; // 'ops' | 'mgmt'

  // Per-stage computed data cache
  var _stageCache = {};

  function computeStageData() {
    _stageCache = {};
    var filtered = Data.getFilteredActiveSI();
    var today = new Date(); today.setHours(0,0,0,0);

    Utils.STAGES.forEach(function(s) {
      var leads = filtered.filter(function(r) { return r.Pipeline === s.api; });
      var openTotal = 0, overdueTotal = 0, dormantCount = 0;
      var greenCount = 0, orangeCount = 0, redCount = 0;
      var addedToday = 0, closedToday = 0;
      var totalActsAll = 0, closedActsAll = 0;

      leads.forEach(function(si) {
        var acts = Data.getActivitiesForSI(si);
        openTotal += acts.open;
        overdueTotal += acts.overdue;
        totalActsAll += acts.open + acts.closed;
        closedActsAll += acts.closed;

        var dorm = Data.getDormancyDays(si);
        if (dorm !== null && dorm > Utils.THRESHOLDS.dormancyDays) dormantCount++;

        // Flow bar segments
        if (dorm === null || dorm <= 3) greenCount++;
        else if (dorm <= 7) orangeCount++;
        else redCount++;

        // Added today
        var created = new Date(si.Created_Time);
        if (!isNaN(created) && created >= today) addedToday++;

        // Closed today
        var closedTodayActs = Data.getClosedTodayActivitiesForSI(si);
        if (closedTodayActs.length > 0) closedToday++;
      });

      _stageCache[s.api] = {
        s: s,
        leads: leads,
        count: leads.length,
        openTotal: openTotal,
        overdueTotal: overdueTotal,
        dormantCount: dormantCount,
        greenCount: greenCount,
        orangeCount: orangeCount,
        redCount: redCount,
        addedToday: addedToday,
        closedToday: closedToday,
        totalActsAll: totalActsAll,
        closedActsAll: closedActsAll,
        avgVel: Data.getAvgVelocity(s.api),
        convRate: Data.getConvRate(s.api),
        completionPct: totalActsAll > 0 ? Math.round((closedActsAll / totalActsAll) * 100) : null
      };
    });
  }

  function buildBarCell(d, stageApi) {
    var total = d.count;
    if (total === 0) {
      return '<td class="bar-cell"><div class="bar-track"></div></td>';
    }
    var gPct = Math.round((d.greenCount / total) * 100);
    var oPct = Math.round((d.orangeCount / total) * 100);
    var rPct = 100 - gPct - oPct;

    var tooltip = '<div class="bar-tooltip">' +
      (d.greenCount > 0 ? '<div class="tt-row"><div class="tt-dot" style="background:var(--green);"></div><span class="tt-label">Active &lt;3d</span><span class="tt-val">' + d.greenCount + ' lead' + (d.greenCount !== 1 ? 's' : '') + '</span></div>' : '') +
      (d.orangeCount > 0 ? '<div class="tt-row"><div class="tt-dot" style="background:var(--orange);"></div><span class="tt-label">Slowing 4–7d</span><span class="tt-val">' + d.orangeCount + ' lead' + (d.orangeCount !== 1 ? 's' : '') + '</span></div>' : '') +
      (d.redCount > 0 ? '<div class="tt-row"><div class="tt-dot" style="background:var(--red);"></div><span class="tt-label">Dormant &gt;7d</span><span class="tt-val c-red">' + d.redCount + ' lead' + (d.redCount !== 1 ? 's' : '') + ' — click OVERDUE</span></div>' : '') +
    '</div>';

    var segs = '';
    if (gPct > 0) segs += '<div class="bar-seg" style="width:' + gPct + '%;background:var(--green);opacity:.75;"></div>';
    if (oPct > 0) segs += '<div class="bar-seg" style="width:' + oPct + '%;background:var(--orange);opacity:.8;"></div>';
    if (rPct > 0) segs += '<div class="bar-seg" style="width:' + rPct + '%;background:var(--red);opacity:.85;"></div>';

    return '<td class="bar-cell">' +
      tooltip +
      '<div class="bar-track">' + segs + '</div>' +
    '</td>';
  }

  function numCell(val, cls, clickable, stageApi, panelType) {
    if (val === null || val === undefined || val === 0 && !clickable) {
      return '<td class="c-dim">—</td>';
    }
    var str = String(val);
    if (clickable && val > 0) {
      return '<td class="' + cls + ' clickable" onclick="TabPipeline.onCellClick(\'' + stageApi + '\',\'' + panelType + '\')">' + str + '</td>';
    }
    return '<td class="' + cls + '">' + str + '</td>';
  }

  function buildRow(d, isEmpty) {
    var s = d.s;
    var stageApi = s.api;
    var isReQuoted = stageApi === '06_Re_Quoted';
    var accentColor = d.count === 0 ? 'var(--dim)' :
                      (d.redCount > 0 || isReQuoted && d.count > Utils.THRESHOLDS.reQuotedCap) ? 'var(--red)' :
                      d.orangeCount > 0 ? 'var(--orange)' : 'var(--green)';
    var rowStyle = isEmpty ? ' class="empty"' : (isReQuoted ? ' style="background:rgba(240,136,62,.04);"' : '');
    var snameCls = d.count === 0 ? 'sname dim' : (isReQuoted ? 'sname c-orange' : 'sname');

    // Variable columns
    var v1Html, v2Html, v3Html;
    if (_currentView === 'ops') {
      // Added Today — clickable
      v1Html = d.addedToday > 0
        ? '<td class="c-green clickable" onclick="TabPipeline.onCellClick(\'' + stageApi + '\',\'added\')">+' + d.addedToday + '</td>'
        : '<td class="c-dim">—</td>';
      // Closed Today — clickable
      v2Html = d.closedToday > 0
        ? '<td class="c-green clickable" onclick="TabPipeline.onCellClick(\'' + stageApi + '\',\'closed\')">' + d.closedToday + '</td>'
        : '<td class="c-dim">—</td>';
      v3Html = '<td style="display:none;"></td>';
    } else {
      // Total Acts
      v1Html = '<td class="c-text">' + (d.totalActsAll > 0 ? d.totalActsAll : '—') + '</td>';
      // Closed Acts
      v2Html = '<td class="c-text">' + (d.closedActsAll > 0 ? d.closedActsAll : '—') + '</td>';
      // Completion %
      v3Html = d.completionPct !== null
        ? '<td class="c-text">' + d.completionPct + '%</td>'
        : '<td class="c-dim">—</td>';
    }

    // Avg vel
    var velStr = d.avgVel !== null ? d.avgVel.toFixed(1) + 'd' : '—';
    var velCls = d.avgVel !== null ? Utils.velClass(d.avgVel) : 'c-dim';

    // Conv rate
    var convStr = d.convRate !== null ? d.convRate + '%' : '—';
    var convCls = d.convRate !== null ? Utils.convClass(d.convRate) : 'c-dim';

    return '<tr' + rowStyle + '>' +
      '<td class="acc" style="background:' + accentColor + ';"></td>' +
      '<td class="left" style="padding-left:10px;"><div class="stage-cell"><span class="snum">' + s.num + '</span><span class="' + snameCls + '">' + Utils.e(s.name) + '</span></div></td>' +
      buildBarCell(d, stageApi) +
      // Common columns
      '<td class="' + (d.count > 0 ? 'c-text fw7' : 'c-dim') + '">' + (d.count > 0 ? d.count : '—') + '</td>' +
      (d.openTotal > 0 ? '<td class="c-blue clickable" onclick="TabPipeline.onCellClick(\'' + stageApi + '\',\'open\')">' + d.openTotal + '</td>' : '<td class="c-dim">—</td>') +
      (d.overdueTotal > 0 ? '<td class="c-red clickable" onclick="TabPipeline.onCellClick(\'' + stageApi + '\',\'overdue\')">' + d.overdueTotal + '</td>' : '<td class="c-dim">0</td>') +
      (d.dormantCount > 0 ? '<td class="c-red">' + d.dormantCount + '</td>' : '<td class="c-dim">0</td>') +
      '<td class="c-dim">—</td>' + // dropped — future
      // Variable
      v1Html + v2Html + v3Html +
      // Tail
      '<td class="' + velCls + '">' + velStr + '</td>' +
      '<td class="' + convCls + '">' + convStr + '</td>' +
    '</tr>';
  }

  function buildSepRow(label, leads, dormant, slowing, overdue, extraCols) {
    var chips = '';
    if (dormant > 0) chips += '<span class="sep-chip" style="background:rgba(248,81,73,.12);color:var(--red);">' + dormant + ' dormant</span>';
    if (slowing > 0) chips += '<span class="sep-chip" style="background:rgba(240,136,62,.12);color:var(--orange);">' + slowing + ' slowing</span>';
    if (overdue > 0) chips += '<span class="sep-chip" style="background:rgba(248,81,73,.12);color:var(--red);">' + overdue + ' overdue</span>';
    return '<tr class="sep-row">' +
      '<td colspan="4" class="left" style="padding:5px 8px 5px 6px;">' +
        '<div class="sep-label">' + Utils.e(label) +
          '<span class="sep-count">' + leads + ' active</span>' +
          '<div class="sep-chips">' + chips + '</div>' +
        '</div>' +
      '</td>' +
      '<td></td><td></td><td></td><td></td>' +
      '<td></td><td></td><td></td><td></td><td></td>' +
    '</tr>';
  }

  function buildTotalsRow(allData) {
    var totalLeads = 0, totalOpen = 0, totalOverdue = 0;
    var totalDormant = 0, totalAddedToday = 0, totalClosedToday = 0;
    var totalActsAll = 0, closedActsAll = 0;
    var velSum = 0, velCount = 0;

    Utils.STAGES.forEach(function(s) {
      var d = allData[s.api];
      if (!d) return;
      totalLeads += d.count;
      totalOpen += d.openTotal;
      totalOverdue += d.overdueTotal;
      totalDormant += d.dormantCount;
      totalAddedToday += d.addedToday;
      totalClosedToday += d.closedToday;
      totalActsAll += d.totalActsAll;
      closedActsAll += d.closedActsAll;
      if (d.avgVel !== null) { velSum += d.avgVel; velCount++; }
    });

    var avgVelStr = velCount > 0 ? (velSum / velCount).toFixed(1) + 'd avg' : '—';
    var completionPct = totalActsAll > 0 ? Math.round((closedActsAll / totalActsAll) * 100) : null;

    var v1Html, v2Html, v3Html;
    if (_currentView === 'ops') {
      v1Html = '<td class="c-green fw7">+' + totalAddedToday + '</td>';
      v2Html = '<td class="c-green fw7">' + totalClosedToday + '</td>';
      v3Html = '<td style="display:none;"></td>';
    } else {
      v1Html = '<td class="c-text fw7">' + totalActsAll + '</td>';
      v2Html = '<td class="c-text fw7">' + closedActsAll + '</td>';
      v3Html = completionPct !== null
        ? '<td class="c-green fw7">' + completionPct + '%</td>'
        : '<td class="c-dim">—</td>';
    }

    return '<tr class="totals-row">' +
      '<td style="background:var(--border2);"></td>' +
      '<td class="left" style="padding-left:10px;"><span class="totals-lbl">Pipeline total</span></td>' +
      '<td></td>' +
      '<td class="c-green fw7">' + totalLeads + '</td>' +
      '<td class="c-blue fw7">' + totalOpen + '</td>' +
      '<td class="c-red fw7">' + totalOverdue + '</td>' +
      '<td class="c-orange fw7">' + totalDormant + '</td>' +
      '<td class="c-dim">—</td>' +
      v1Html + v2Html + v3Html +
      '<td class="c-muted">' + avgVelStr + '</td>' +
      '<td class="c-muted">—</td>' +
    '</tr>';
  }

  function renderSummaryStrip() {
    var filtered = Data.getFilteredActiveSI();
    var totalLeads = filtered.length;
    var totalOpen = 0, totalOverdue = 0, totalDormant = 0;
    var velSum = 0, velCount = 0;
    var bottleneckStage = '—';
    var maxPressure = -1;

    filtered.forEach(function(si) {
      var acts = Data.getActivitiesForSI(si);
      totalOpen += acts.open;
      totalOverdue += acts.overdue;
      var dorm = Data.getDormancyDays(si);
      if (dorm !== null && dorm > Utils.THRESHOLDS.dormancyDays) totalDormant++;
    });

    Utils.STAGES.forEach(function(s) {
      var d = _stageCache[s.api];
      if (!d || d.count === 0) return;
      var pressure = (d.count * 5) + (d.openTotal * 2) + (d.overdueTotal * 3);
      if (pressure > maxPressure) { maxPressure = pressure; bottleneckStage = s.num + ' ' + s.name; }
      if (d.avgVel !== null) { velSum += d.avgVel; velCount++; }
    });

    var avgVel = velCount > 0 ? (velSum / velCount).toFixed(1) + 'd' : '—';

    // All SI for velocity (not filtered)
    var allClosedRecs = Data.getAllSI().filter(function(r) {
      return r.Current_Stage === false && parseFloat(r.Duration_Minutes) > 0;
    });
    if (allClosedRecs.length >= 3) {
      var totalMins = allClosedRecs.reduce(function(a, r) { return a + parseFloat(r.Duration_Minutes); }, 0);
      avgVel = ((totalMins / allClosedRecs.length) / (60 * 24)).toFixed(1) + 'd';
    }

    setText('sv-leads', totalLeads);
    setText('sv-open', totalOpen);
    setText('sv-overdue', totalOverdue);
    setText('sv-dormant', totalDormant);
    setText('sv-bot', bottleneckStage);
    setText('sv-vel', avgVel);
    setText('ss-leads', 'active in pipeline');
    setText('ss-dormant', 'no touch >7d');
  }

  function setText(id, val) {
    var el = document.getElementById(id);
    if (el) el.textContent = val;
  }

  function render() {
    computeStageData();
    renderSummaryStrip();

    var TOFU = Utils.STAGES.slice(0, 3);
    var MOFU = Utils.STAGES.slice(3, 9);
    var BOFU = Utils.STAGES.slice(9, 15);

    function layerStats(stages) {
      var leads = 0, dorm = 0, slow = 0, over = 0;
      stages.forEach(function(s) {
        var d = _stageCache[s.api];
        if (!d) return;
        leads += d.count;
        dorm += d.dormantCount;
        slow += d.orangeCount;
        over += d.overdueTotal;
      });
      return { leads: leads, dorm: dorm, slow: slow, over: over };
    }

    var tStats = layerStats(TOFU);
    var mStats = layerStats(MOFU);
    var bStats = layerStats(BOFU);

    var html = '';

    // TOFU
    html += buildSepRow('Top of Funnel', tStats.leads, tStats.dorm, tStats.slow, tStats.over);
    TOFU.forEach(function(s) {
      var d = _stageCache[s.api];
      html += buildRow(d, d.count === 0);
    });

    // MOFU
    html += buildSepRow('Mid Funnel', mStats.leads, mStats.dorm, mStats.slow, mStats.over);
    MOFU.forEach(function(s) {
      var d = _stageCache[s.api];
      html += buildRow(d, d.count === 0);
      if (s.api === '06_Re_Quoted') {
        var capBreach = d.count > Utils.THRESHOLDS.reQuotedCap;
        html += '<tr class="cap-row"><td colspan="13" class="' + (capBreach ? 'breach' : '') + '">' +
          'Re-Quoted cap · max ' + Utils.THRESHOLDS.reQuotedCap + ' leads · ' + d.count + ' of ' + Utils.THRESHOLDS.reQuotedCap + ' used · ' +
          (capBreach ? '⚠ THRESHOLD BREACHED' : 'within limit') +
        '</td></tr>';
      }
    });

    // BOFU
    html += buildSepRow('Bottom of Funnel', bStats.leads, bStats.dorm, bStats.slow, bStats.over);
    BOFU.forEach(function(s) {
      var d = _stageCache[s.api];
      html += buildRow(d, d.count === 0);
    });

    // Totals
    html += buildTotalsRow(_stageCache);

    document.getElementById('pipeline-tbody').innerHTML = html;
    updateColumnHeaders();
  }

  function updateColumnHeaders() {
    var isOps = _currentView === 'ops';
    var th1 = document.getElementById('th-v1');
    var th2 = document.getElementById('th-v2');
    var th3 = document.getElementById('th-v3');
    var thConv = document.getElementById('th-conv');
    if (th1) th1.textContent = isOps ? 'ADDED TODAY' : 'TOTAL ACTS';
    if (th2) th2.textContent = isOps ? 'CLOSED TODAY' : 'CLOSED ACTS';
    if (th3) th3.style.display = isOps ? 'none' : '';
    if (thConv) thConv.textContent = isOps ? 'CONV.' : 'FWD EXITS %';
  }

  function setView(view) {
    _currentView = view;
    Data.currentView = view;
    render();
  }

  function onCellClick(stageApi, type) {
    if (type === 'open') { PanelActivity.open(stageApi, 'all'); }
    else if (type === 'overdue') { PanelActivity.open(stageApi, 'overdue'); }
    else if (type === 'added') { PanelLeads.open(stageApi, 'added'); }
    else if (type === 'closed') { PanelLeads.open(stageApi, 'closed'); }
  }

  function init() { render(); }

  return { init: init, render: render, setView: setView, onCellClick: onCellClick };
})();

// ============================================================
// TAB — REP VIEW
// Adding a new tab: copy this IIFE pattern
// ============================================================
var TabRepView = (function() {

  var REP_COLORS = ['#3fb950','#58a6ff','#f0883e','#f85149','#a78bfa'];

  function getRepColor(name, i) {
    return REP_COLORS[i % REP_COLORS.length];
  }

  function getInitials(name) {
    return (name || 'UN').split(' ').map(function(w) { return w[0] || ''; }).join('').substring(0, 2).toUpperCase();
  }

  function render() {
    var allActive = Data.getActiveSI();
    var repMap = {};

    allActive.forEach(function(si) {
      var name = Utils.getRepName(si);
      if (!repMap[name]) {
        repMap[name] = {
          name: name,
          leads: 0,
          callsDone: 0,
          tasksDone: 0,
          closedActs: 0,
          dormant: 0,
          fwdExits: 0,
          totalExits: 0
        };
      }
      var r = repMap[name];
      r.leads++;
      var acts = Data.getActivitiesForSI(si);
      r.callsDone += acts.calls > 0 && acts.closed > 0 ? Math.min(acts.calls, acts.closed) : 0;
      r.tasksDone += acts.closed > 0 ? Math.max(0, acts.closed - (acts.calls > 0 ? 1 : 0)) : 0;
      r.closedActs += acts.closed;
      var dorm = Data.getDormancyDays(si);
      if (dorm !== null && dorm > Utils.THRESHOLDS.dormancyDays) r.dormant++;
    });

    var reps = Object.values(repMap).sort(function(a, b) { return b.leads - a.leads; });

    if (reps.length === 0) {
      document.getElementById('rep-tbody').innerHTML =
        '<tr><td colspan="7" style="padding:30px;text-align:center;color:var(--muted);">No rep data available. Ensure Rep_Name field is populated on Sales Intelligence records.</td></tr>';
      return;
    }

    var html = reps.map(function(r, i) {
      var color = getRepColor(r.name, i);
      var initials = getInitials(r.name);
      var convStr = '—';
      var dormCls = r.dormant > 0 ? 'c-red' : 'c-dim';

      return '<tr>' +
        '<td class="left">' +
          '<div class="rep-cell">' +
            '<div class="rep-av-lg" style="background:' + color + '18;color:' + color + ';">' + initials + '</div>' +
            '<div>' +
              '<div class="rep-name">' + Utils.e(r.name) + '</div>' +
              '<div class="rep-sub">' + r.leads + ' active lead' + (r.leads !== 1 ? 's' : '') + '</div>' +
            '</div>' +
          '</div>' +
        '</td>' +
        '<td class="c-text">' + r.leads + '</td>' +
        '<td class="c-green">' + (r.callsDone || '—') + '</td>' +
        '<td class="c-blue">' + (r.tasksDone || '—') + '</td>' +
        '<td class="c-text">' + (r.closedActs || '—') + '</td>' +
        '<td class="' + dormCls + '">' + (r.dormant > 0 ? r.dormant : '0') + '</td>' +
        '<td class="c-muted">' + convStr + '</td>' +
      '</tr>';
    }).join('');

    document.getElementById('rep-tbody').innerHTML = html;
  }

  function init() { render(); }

  return { init: init, render: render };
})();

// ============================================================
// TAB — DIGEST
// Template-based prose briefing. No Claude API in v8.
// Adding a new tab: copy this IIFE pattern
// ============================================================
var TabDigest = (function() {

  function pct(a, b) { return b > 0 ? Math.round((a / b) * 100) : 0; }

  function hl(text, color) {
    return '<span class="hl-' + color + '">' + Utils.e(String(text)) + '</span>';
  }

  function actItem(cls, tag, tagCls, text) {
    return '<div class="act-item ' + cls + '"><span class="act-tag ' + tagCls + '">' + tag + '</span><span>' + Utils.e(text) + '</span></div>';
  }

  function buildOverall(filtered, allSI) {
    var totalLeads = filtered.length;
    var tofuLeads = filtered.filter(function(r) { return Utils.STAGE_NUM[r.Pipeline] <= 3; }).length;
    var mofuLeads = filtered.filter(function(r) { var n = Utils.STAGE_NUM[r.Pipeline]; return n >= 4 && n <= 9; }).length;
    var bofuLeads = filtered.filter(function(r) { return Utils.STAGE_NUM[r.Pipeline] >= 10; }).length;

    var totalOpen = 0, totalDormant = 0, totalOverdue = 0;
    filtered.forEach(function(si) {
      var acts = Data.getActivitiesForSI(si);
      totalOpen += acts.open;
      totalOverdue += acts.overdue;
      var dorm = Data.getDormancyDays(si);
      if (dorm !== null && dorm > Utils.THRESHOLDS.dormancyDays) totalDormant++;
    });

    var closedRecs = allSI.filter(function(r) { return r.Current_Stage === false && parseFloat(r.Duration_Minutes) > 0; });
    var avgVelStr = '—';
    if (closedRecs.length > 0) {
      var totalMins = closedRecs.reduce(function(a, r) { return a + parseFloat(r.Duration_Minutes); }, 0);
      avgVelStr = ((totalMins / closedRecs.length) / (60 * 24)).toFixed(1) + ' days';
    }

    var regressions = 0;
    allSI.forEach(function(r) {
      var fromNum = Utils.STAGE_NUM[r.From_Stage];
      var toNum = Utils.STAGE_NUM[r.Pipeline];
      if (fromNum && toNum && fromNum > toNum) regressions++;
    });

    // What happened
    var happened = 'The pipeline currently has ' + hl(totalLeads, 'green') + ' active leads. ' +
      hl(tofuLeads, 'blue') + ' are in the top of funnel (Stages 01–03), ' +
      hl(mofuLeads, 'blue') + ' in mid funnel (04–09), and ' +
      hl(bofuLeads, 'blue') + ' in bottom of funnel (10–12). ' +
      'Average stage velocity is ' + hl(avgVelStr, 'green') + ' per stage from ' + closedRecs.length + ' completed stage visits. ' +
      (regressions > 0 ? hl(regressions + ' regression event' + (regressions !== 1 ? 's' : ''), 'red') + ' detected — leads that moved backward in the pipeline.' : 'No regressions detected this period.');

    // Where things stand
    var stands = 'There are ' + hl(totalOpen, 'blue') + ' open activities across the pipeline. ';
    if (totalOverdue > 0) {
      stands += hl(totalOverdue, 'red') + ' of these are past their due date and need immediate attention. ';
    } else {
      stands += 'No activities are currently overdue — good hygiene. ';
    }
    if (totalDormant > 0) {
      stands += hl(totalDormant + ' lead' + (totalDormant !== 1 ? 's' : ''), 'red') + ' have had no activity in over 7 days. ';
    } else {
      stands += 'No dormant leads — all leads have been touched within 7 days. ';
    }

    // Act on this
    var actItems = [];
    filtered.forEach(function(si) {
      var dorm = Data.getDormancyDays(si);
      var acts = Data.getActivitiesForSI(si);
      var name = Utils.getLeadName(si);
      var stage = (Utils.STAGE_MAP[si.Pipeline] || { name: si.Pipeline }).name;
      var stageNum = Utils.STAGE_NUM[si.Pipeline] || 0;

      if (dorm !== null && dorm > 14) {
        actItems.push(actItem('stalled', 'NO ACTIVITY', 'rd', name + ' · ' + stage + ' · ' + dorm + ' days with no touch'));
      } else if (dorm !== null && dorm > 7) {
        actItems.push(actItem('stalled', 'DORMANT', 'rd', name + ' · ' + stage + ' · Last touched ' + dorm + ' days ago'));
      }
      if (acts.overdue > 0) {
        actItems.push(actItem('warn', 'OVERDUE', 'am', name + ' · ' + acts.overdue + ' overdue activit' + (acts.overdue !== 1 ? 'ies' : 'y') + ' in ' + stage));
      }
      if (stageNum >= 10 && stageNum <= 12) {
        actItems.push(actItem('good', 'NEAR CLOSE', 'tl', name + ' · In ' + stage + ' — follow up to close'));
      }
    });

    var regressionItems = [];
    Data.getAllSI().forEach(function(r) {
      var fromNum = Utils.STAGE_NUM[r.From_Stage];
      var toNum = Utils.STAGE_NUM[r.Pipeline];
      if (fromNum && toNum && fromNum > toNum) {
        var mjid = r.Master_Journey_ID || '';
        regressionItems.push(actItem('stalled', 'REGRESSION', 'rd', mjid + ' moved from ' + r.From_Stage + ' back to ' + r.Pipeline));
      }
    });
    actItems = actItems.concat(regressionItems);

    var actHtml = actItems.length > 0
      ? actItems.slice(0, 15).join('')
      : '<div style="color:var(--muted);font-size:12px;font-family:var(--sans);">No flagged items. Pipeline is healthy.</div>';

    return { happened: happened, stands: stands, actHtml: actHtml };
  }

  function render() {
    var filtered = Data.getFilteredActiveSI();
    var allSI = Data.getAllSI();
    var overall = buildOverall(filtered, allSI);

    var repSection = '';
    var repNames = Data.getRepNames();
    repNames.forEach(function(repName) {
      var repSI = filtered.filter(function(r) { return Utils.getRepName(r) === repName; });
      var repLeads = repSI.length;
      var repOpen = 0, repDorm = 0, repOverdue = 0, repClosed = 0;
      repSI.forEach(function(si) {
        var acts = Data.getActivitiesForSI(si);
        repOpen += acts.open;
        repOverdue += acts.overdue;
        repClosed += acts.closed;
        var dorm = Data.getDormancyDays(si);
        if (dorm !== null && dorm > Utils.THRESHOLDS.dormancyDays) repDorm++;
      });

      var repHappened = repName + ' has ' + hl(repLeads, 'green') + ' active lead' + (repLeads !== 1 ? 's' : '') + '. ';
      repHappened += hl(repClosed, 'green') + ' activities closed, ' + hl(repOpen, 'blue') + ' open. ';
      if (repOverdue > 0) repHappened += hl(repOverdue, 'red') + ' overdue. ';
      if (repDorm > 0) repHappened += hl(repDorm, 'red') + ' dormant lead' + (repDorm !== 1 ? 's' : '') + '. ';
      else repHappened += 'No dormant leads. ';

      repSection += '<div class="digest-rep-header">' +
        '<span style="width:24px;height:24px;border-radius:50%;background:rgba(63,185,80,.15);color:var(--green);display:flex;align-items:center;justify-content:center;font-size:10px;font-weight:700;">' +
          (repName || 'UN').split(' ').map(function(w) { return w[0] || ''; }).join('').substring(0, 2).toUpperCase() +
        '</span>' +
        Utils.e(repName) +
      '</div>' +
      '<div class="digest-section"><div class="digest-prose">' + repHappened + '</div></div>';
    });

    document.getElementById('digest-content').innerHTML =
      '<div class="digest-section">' +
        '<div class="digest-section-title">What happened</div>' +
        '<div class="digest-prose">' + overall.happened + '</div>' +
      '</div>' +
      '<div class="digest-section">' +
        '<div class="digest-section-title">Where things stand</div>' +
        '<div class="digest-prose">' + overall.stands + '</div>' +
      '</div>' +
      '<div class="digest-section">' +
        '<div class="digest-section-title">Act on this</div>' +
        overall.actHtml +
      '</div>' +
      (repSection ? '<div class="digest-section"><div class="digest-section-title">Per-rep breakdown</div>' + repSection + '</div>' : '');
  }

  function init() { render(); }

  return { init: init, render: render };
})();

// ============================================================
// APP
// Controls tab switching, toggles, panel events.
// Adding a new tab: register it here in TABS map.
// ============================================================
var App = (function() {

  // Tab registry — add new tabs here only
  var TABS = {
    pipeline: { module: null, panelId: 'panel-pipeline' },
    repview:  { module: null, panelId: 'panel-repview' },
    digest:   { module: null, panelId: 'panel-digest' }
  };

  var currentTab = 'pipeline';

  function setStatus(msg) {
    var el = document.getElementById('status-bar');
    if (!el) return;
    if (!msg) { el.classList.add('hidden'); return; }
    el.classList.remove('hidden');
    el.textContent = msg;
  }

  function switchTab(tabKey) {
    if (!TABS[tabKey]) return;
    currentTab = tabKey;

    // Update tab buttons
    document.querySelectorAll('.tab-btn').forEach(function(b) {
      b.classList.toggle('on', b.dataset.tab === tabKey);
    });

    // Show/hide panels
    document.querySelectorAll('.tab-panel').forEach(function(p) { p.classList.remove('on'); });
    document.getElementById('panel-' + tabKey).classList.add('on');

    // Show/hide view bar (pipeline only)
    var viewBar = document.getElementById('view-bar');
    if (viewBar) viewBar.style.display = (tabKey === 'pipeline') ? '' : 'none';

    // Render active tab
    if (tabKey === 'pipeline') TabPipeline.render();
    else if (tabKey === 'repview') TabRepView.render();
    else if (tabKey === 'digest') TabDigest.render();
  }

  function setRep(repKey, repName) {
    if (Data.currentView === 'mgmt') return;
    Data.currentRep = repKey === 'all' ? 'all' : repName;
    document.querySelectorAll('.rep-btn').forEach(function(b) { b.classList.remove('on'); });
    document.querySelector('.rep-btn[data-rep="' + repKey + '"]').classList.add('on');
    if (currentTab === 'pipeline') TabPipeline.render();
    else if (currentTab === 'repview') TabRepView.render();
    else if (currentTab === 'digest') TabDigest.render();
  }

  function setPeriod(period) {
    if (Data.currentView === 'mgmt') return;
    Data.currentPeriod = period;
    document.querySelectorAll('.per-btn').forEach(function(b) {
      b.classList.toggle('on', b.dataset.period === period);
    });
    if (currentTab === 'pipeline') TabPipeline.render();
    else if (currentTab === 'digest') TabDigest.render();
  }

  function setView(view) {
    Data.currentView = view;

    document.querySelectorAll('.v-btn').forEach(function(b) {
      b.classList.toggle('on', b.dataset.view === view);
    });

    var isMgmt = view === 'mgmt';
    document.getElementById('rep-toggle').classList.toggle('locked', isMgmt);
    document.getElementById('period-toggle').classList.toggle('locked', isMgmt);
    document.getElementById('rep-lock-badge').classList.toggle('show', isMgmt);
    document.getElementById('period-lock-badge').classList.toggle('show', isMgmt);
    document.getElementById('mgmt-note').classList.toggle('show', isMgmt);

    if (isMgmt) {
      document.getElementById('view-hint').textContent = 'Showing cumulative totals · completion % = closed ÷ total activities · conv = forward exits only';
      // Reset rep to all visually
      document.querySelectorAll('.rep-btn').forEach(function(b) { b.classList.remove('on'); });
      document.querySelector('.rep-btn[data-rep="all"]').classList.add('on');
      Data.currentRep = 'all';
    } else {
      document.getElementById('view-hint').textContent = 'Click OPEN ACTS, OVERDUE, ADDED TODAY or CLOSED TODAY to open work queue';
    }

    TabPipeline.setView(view);
  }

  function closePanel() {
    document.getElementById('overlay').classList.remove('open');
    document.getElementById('side-panel').classList.remove('open');
  }

  function wireRepButtons() {
    // Wire static All button
    document.querySelector('.rep-btn[data-rep="all"]').addEventListener('click', function() {
      setRep('all', 'all');
    });
    // Wire named rep buttons dynamically from actual data
    var repNames = Data.getRepNames();
    var repBtns = document.querySelectorAll('.rep-btn[data-rep^="__"]');
    repBtns.forEach(function(btn, i) {
      if (repNames[i]) {
        var name = repNames[i];
        var initials = name.split(' ').map(function(w) { return w[0] || ''; }).join('').substring(0, 2).toUpperCase();
        btn.dataset.rep = name;
        var avEl = btn.querySelector('.rep-av');
        if (avEl) avEl.textContent = initials;
        var spanEl = btn.querySelector('span:not(.rep-av)');
        if (spanEl) spanEl.textContent = name.split(' ')[0];
        btn.addEventListener('click', function() { setRep(name, name); });
        btn.style.display = '';
      } else {
        btn.style.display = 'none';
      }
    });
  }

  function init() {
    setStatus('Connecting to Zoho CRM...');

    Data.load(function() {
      // Wire rep buttons with real names from data
      wireRepButtons();

      // Tab buttons
      document.querySelectorAll('.tab-btn').forEach(function(b) {
        b.addEventListener('click', function() { switchTab(b.dataset.tab); });
      });

      // Period buttons
      document.querySelectorAll('.per-btn').forEach(function(b) {
        b.addEventListener('click', function() { setPeriod(b.dataset.period); });
      });

      // View toggle
      document.querySelectorAll('.v-btn').forEach(function(b) {
        b.addEventListener('click', function() { setView(b.dataset.view); });
      });

      // Panel close
      document.getElementById('panel-close').addEventListener('click', closePanel);
      document.getElementById('overlay').addEventListener('click', closePanel);

      // Panel sort chips
      document.querySelectorAll('.sort-chip').forEach(function(c) {
        c.addEventListener('click', function() {
          document.querySelectorAll('.sort-chip').forEach(function(x) { x.classList.remove('on'); });
          c.classList.add('on');
          PanelActivity.onSort(c.dataset.sort);
        });
      });

      // Initial render
      TabPipeline.init();
    });
  }

  return { init: init, closePanel: closePanel, setView: setView };
})();

// ============================================================
// BOOT
// ============================================================
ZOHO.embeddedApp.on('PageLoad', function() {
  App.init();
});
ZOHO.embeddedApp.init();
