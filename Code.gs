/**
 * Work Manager - backend (Google Apps Script).
 * All data is stored in THIS Google Sheet, and files (photos / voice notes) in the
 * admin's Google Drive, so everything stays inside the admin's Google account.
 * All times are Delhi time (Asia/Kolkata), shown in 12-hour AM/PM format.
 */
var TZ = 'Asia/Kolkata';
var DAY = 86400000;

var SCHEMA = {
  Users: ['id','role','name','username','phone','email','salt','hash','shiftStart','shiftEnd','shiftType','salary','otRate','active','mustChange','created'],
  Sessions: ['token','userId','expires'],
  Attendance: ['id','userId','date','inMs','outMs','inLat','inLng','inAcc','outLat','outLng','outAcc','shiftType','schedMin','workedMin','otMin','status'],
  Tasks: ['id','userId','date','title','details','status','workerNote','createdBy','doneMs'],
  Logs: ['id','userId','date','ms','note','photoId','audioId'],
  Stock: ['id','userId','date','ms','item','qty','unit','note'],
  Payments: ['id','userId','month','type','amount','date','note','by']
};

/* ---------- Web entry points ---------- */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index').setTitle('Work Manager')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
function doPost(e) {
  var b = {};
  try { b = JSON.parse(e.postData.contents); } catch (x) {}
  return ContentService.createTextOutput(JSON.stringify(dispatch_(b.token, b.action, b.payload)))
    .setMimeType(ContentService.MimeType.JSON);
}
function api(token, action, payload) { return dispatch_(token, action, payload); }

/* ---------- One-time setup (run once from the editor) ---------- */
function setup() {
  var ss = SS_();
  Object.keys(SCHEMA).forEach(function (n) {
    var s = ss.getSheetByName(n) || ss.insertSheet(n), h = SCHEMA[n];
    s.getRange(1, 1, s.getMaxRows(), h.length).setNumberFormat('@');
    s.getRange(1, 1, 1, h.length).setValues([h]).setFontWeight('bold');
    s.setFrozenRows(1);
  });
  var d = ss.getSheetByName('Sheet1');
  if (d && ss.getSheets().length > 1) ss.deleteSheet(d);
  var P = PropertiesService.getScriptProperties();
  if (!P.getProperty('FOLDER')) P.setProperty('FOLDER', DriveApp.createFolder('WorkManager Files').getId());
  P.setProperty('ADMIN_EMAIL', Session.getEffectiveUser().getEmail());
  if (!rows_('Users').length) {
    addUser_({ role: 'admin', name: 'Admin One', username: 'admin1', mustChange: 'TRUE' }, 'Admin1@Delhi26');
    addUser_({ role: 'admin', name: 'Admin Two', username: 'admin2', mustChange: 'TRUE' }, 'Admin2@Delhi26');
  }
  Logger.log('Setup done. Logins: admin1 / Admin1@Delhi26   and   admin2 / Admin2@Delhi26');
}
/** Optional: emails a full PDF backup to the admin every night at 11 PM Delhi time. */
function enableDailyBackup() {
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'dailyBackup') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('dailyBackup').timeBased().atHour(23).everyDays(1).inTimezone(TZ).create();
}
function dailyBackup() { emailBackup({ role: 'admin', email: '' }, { type: 'full', month: today_().slice(0, 7), userId: '' }); }

/* ---------- Database helpers (Google Sheet) ---------- */
function SS_() { return SpreadsheetApp.getActiveSpreadsheet(); }
function rows_(n) {
  var s = SS_().getSheetByName(n), v = s.getDataRange().getValues(), h = SCHEMA[n], out = [];
  for (var i = 1; i < v.length; i++) {
    var o = { _row: i + 1 };
    for (var j = 0; j < h.length; j++) o[h[j]] = v[i][j] === '' || v[i][j] == null ? '' : String(v[i][j]);
    if (o[h[0]] !== '') out.push(o);
  }
  return out;
}
function add_(n, o) {
  var s = SS_().getSheetByName(n), h = SCHEMA[n], r = s.getLastRow() + 1;
  if (r > s.getMaxRows()) s.insertRowsAfter(s.getMaxRows(), 200);
  var rg = s.getRange(r, 1, 1, h.length);
  rg.setNumberFormat('@');
  rg.setValues([h.map(function (k) { return o[k] == null ? '' : String(o[k]); })]);
}
function upd_(n, row, o) {
  var h = SCHEMA[n], rg = SS_().getSheetByName(n).getRange(row, 1, 1, h.length), cur = rg.getValues()[0];
  h.forEach(function (k, j) { if (o.hasOwnProperty(k)) cur[j] = String(o[k]); });
  rg.setNumberFormat('@'); rg.setValues([cur]);
}
function del_(n, row) { SS_().getSheetByName(n).deleteRow(row); }
function newId_() { return Utilities.getUuid().replace(/-/g, '').slice(0, 10); }

/* ---------- Time helpers (Delhi time, 12-hour) ---------- */
function today_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'); }
function fmtT_(ms) { return Utilities.formatDate(new Date(Number(ms)), TZ, 'hh:mm a'); }
function fmtDT_(ms) { return Utilities.formatDate(new Date(Number(ms)), TZ, 'dd MMM yyyy, hh:mm a'); }
function fmtD_(d) { return Utilities.formatDate(new Date(d + 'T12:00:00+05:30'), TZ, 'dd MMM yyyy, EEE'); }
function hm_(m) { m = Math.round(Number(m) || 0); return Math.floor(m / 60) + 'h ' + ('0' + (m % 60)).slice(-2) + 'm'; }
function t12_(v) { var p = String(v || '09:00').split(':'), h = +p[0]; return (h % 12 || 12) + ':' + p[1] + ' ' + (h < 12 ? 'AM' : 'PM'); }
function toMin_(v) { var p = String(v || '0:0').split(':'); return (+p[0]) * 60 + (+p[1]); }
function sched_(u) { var d = (toMin_(u.shiftEnd) - toMin_(u.shiftStart) + 1440) % 1440; return d || 1440; }
function map_(a, b) { return a !== '' && b !== '' && a != null ? 'https://maps.google.com/?q=' + a + ',' + b : ''; }
function sum_(arr, k) { return arr.reduce(function (t, r) { return t + (Number(r[k]) || 0); }, 0); }
function inM_(d, m) { return !m || m === 'ALL' || String(d).indexOf(m) === 0; }
function otRate_(x) { return Number(x.otRate) || Math.round((Number(x.salary) || 0) / 30 / (sched_(x) / 60)); }

/* ---------- Passwords & sessions ---------- */
function hash_(pw, salt) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + pw)
    .map(function (b) { return ('0' + (b & 0xff).toString(16)).slice(-2); }).join('');
}
function addUser_(o, pw) {
  var salt = Utilities.getUuid(), id = newId_();
  add_('Users', {
    id: id, role: o.role || 'worker', name: o.name, username: o.username, phone: o.phone || '', email: o.email || '',
    salt: salt, hash: hash_(pw, salt), shiftStart: o.shiftStart || '09:00', shiftEnd: o.shiftEnd || '18:00',
    shiftType: o.shiftType || 'Day', salary: o.salary || 0, otRate: o.otRate == null ? '' : o.otRate,
    active: o.active || 'TRUE', mustChange: o.mustChange || 'FALSE', created: fmtDT_(Date.now())
  });
  return id;
}
function setPw_(user, pw) {
  var salt = Utilities.getUuid();
  upd_('Users', user._row, { salt: salt, hash: hash_(pw, salt), mustChange: 'FALSE' });
  rows_('Sessions').forEach(function (s) { if (s.userId === user.id) upd_('Sessions', s._row, { expires: 0 }); });
}
function pub_(u) {
  return { id: u.id, role: u.role, name: u.name, username: u.username, phone: u.phone, email: u.email,
    shiftStart: u.shiftStart, shiftEnd: u.shiftEnd, shiftType: u.shiftType, salary: Number(u.salary) || 0,
    otRate: u.otRate === '' ? '' : Number(u.otRate), rate: otRate_(u), active: u.active === 'TRUE', mustChange: u.mustChange === 'TRUE' };
}
function auth_(token) {
  var s = rows_('Sessions').filter(function (r) { return r.token === token && Number(r.expires) > Date.now(); })[0];
  var u = s && rows_('Users').filter(function (r) { return r.id === s.userId && r.active === 'TRUE'; })[0];
  if (!u) throw new Error('SESSION');
  return u;
}
function login_(p) {
  var id = String(p.username || '').trim().toLowerCase().replace(/\s/g, ''), pw = String(p.password || '');
  if (!id || !pw) throw new Error('Enter username and password');
  var C = CacheService.getScriptCache(), k = 'fail_' + id.replace(/[^a-z0-9]/g, '').slice(0, 40), f = Number(C.get(k) || 0);
  if (f >= 5) throw new Error('Too many wrong attempts. Try again after 10 minutes.');
  var u = rows_('Users').filter(function (r) {
    return r.active === 'TRUE' && [r.username, r.phone, r.email].some(function (x) { return x && x.toLowerCase().replace(/\s/g, '') === id; });
  })[0];
  if (!u || hash_(pw, u.salt) !== u.hash) { C.put(k, String(f + 1), 600); throw new Error('Wrong username or password'); }
  C.remove(k);
  rows_('Sessions').filter(function (s) { return Number(s.expires) < Date.now(); }).reverse().slice(0, 40).forEach(function (s) { del_('Sessions', s._row); });
  var t =Utilities.getUuid() + Utilities.getUuid();
  add_('Sessions', { token: t, userId: u.id, expires: Date.now() + 30 * DAY });
  return { token: t, user: pub_(u) };
}

/* ---------- Dispatcher ---------- */
var ADMIN_ONLY = ['dash', 'users', 'userSave', 'userActive', 'userSetPassword', 'taskAdd', 'taskDelete', 'attDelete', 'payAdd', 'payDelete', 'pdf', 'emailBackup'];
var WORKER_ONLY = ['today', 'checkIn', 'checkOut', 'logAdd', 'stockAdd'];
var READ = ['me', 'today', 'attList', 'taskList', 'logList', 'stockList', 'payList', 'salary', 'fileGet', 'dash', 'users', 'pdf'];

function dispatch_(token, action, p) {
  var lock = null;
  try {
    p = p || {};
    if (action === 'login') { lock = LockService.getScriptLock(); lock.waitLock(20000); return { ok: true, data: login_(p) }; }
    var u = auth_(token);
    if (ADMIN_ONLY.indexOf(action) >= 0 && u.role !== 'admin') throw new Error('Admin only');
    if (WORKER_ONLY.indexOf(action) >= 0 && u.role !== 'worker') throw new Error('Workers only');
    var fn = HANDLERS[action];
    if (!fn) throw new Error('Unknown action');
    if (READ.indexOf(action) < 0) { lock = LockService.getScriptLock(); lock.waitLock(20000); }
    return { ok: true, data: fn(u, p) };
  } catch (e) {
    return { ok: false, error: String(e && e.message ? e.message : e) };
  } finally { if (lock) { try { lock.releaseLock(); } catch (x) {} } }
}

/* ---------- Common ---------- */
function me(u) { return pub_(u); }
function logout(u, p) { rows_('Sessions').forEach(function (s) { if (s.userId === u.id && s.token === p.token) upd_('Sessions', s._row, { expires: 0 }); }); return true; }
function changePassword(u, p) {
  if (hash_(String(p.oldPassword || ''), u.salt) !== u.hash) throw new Error('Old password is wrong');
  if (String(p.newPassword || '').length < 5) throw new Error('New password must be at least 5 characters');
  setPw_(u, String(p.newPassword));
  var t = Utilities.getUuid() + Utilities.getUuid();
  add_('Sessions', { token: t, userId: u.id, expires: Date.now() + 30 * DAY });
  return { token: t };
}
function names_() { var m = {}; rows_('Users').forEach(function (r) { m[r.id] = r.name; }); return m; }
function uid_(u, p) { return u.role === 'admin' ? (p.userId && p.userId !== 'ALL' ? p.userId : '') : u.id; }

/* ---------- Attendance ---------- */
function monthStats_(uid, month) {
  var a = rows_('Attendance').filter(function (r) { return r.userId === uid && inM_(r.date, month); }), d = {};
  a.forEach(function (r) { d[r.date] = 1; });
  return { days: Object.keys(d).length, workedMin: sum_(a, 'workedMin'), otMin: sum_(a, 'otMin') };
}
function today(u) {
  var T = today_(), att = rows_('Attendance').filter(function (r) { return r.userId === u.id; });
  var open = att.filter(function (r) { return !r.outMs; })[0];
  return {
    serverNow: Date.now(), today: T, nowLabel: fmtDT_(Date.now()),
    sched: { start: u.shiftStart, end: u.shiftEnd, type: u.shiftType, minutes: sched_(u) },
    open: open ? { inMs: Number(open.inMs), in12: fmtDT_(open.inMs), shiftType: open.shiftType, sched: Number(open.schedMin) } : null,
    todayRecs: att.filter(function (r) { return r.date === T && r.outMs; }).map(function (r) {
      return { in12: fmtT_(r.inMs), out12: fmtT_(r.outMs), worked: hm_(r.workedMin), ot: hm_(r.otMin) };
    }),
    month: monthStats_(u.id, T.slice(0, 7))
  };
}
function checkIn(u, p) {
  var lat = Number(p.lat), lng = Number(p.lng);
  if (p.lat === '' || p.lat == null || !isFinite(lat) || !isFinite(lng)) throw new Error('Location is required. Please allow location access and try again.');
  var now = Date.now(), open = rows_('Attendance').filter(function (r) { return r.userId === u.id && !r.outMs; })[0];
  if (open) {
    if (now - Number(open.inMs) < 18 * 3600000) throw new Error('You are already checked in. Please check out first.');
    upd_('Attendance', open._row, { outMs: Number(open.inMs) + Number(open.schedMin) * 60000, workedMin: open.schedMin, otMin: 0, status: 'Auto-closed (no check-out)' });
  }
  add_('Attendance', { id: newId_(), userId: u.id, date: today_(), inMs: now, outMs: '', inLat: lat, inLng: lng, inAcc: Math.round(Number(p.acc) || 0),
    shiftType: p.shiftType === 'Night' ? 'Night' : 'Day', schedMin: sched_(u), workedMin: '', otMin: '', status: 'Working' });
  return today(u);
}
function checkOut(u, p) {
  var open = rows_('Attendance').filter(function (r) { return r.userId === u.id && !r.outMs; })[0];
  if (!open) throw new Error('You are not checked in.');
  var now = Date.now(), worked = Math.max(0, Math.round((now - Number(open.inMs)) / 60000)), ot = Math.max(0, worked - Number(open.schedMin));
  upd_('Attendance', open._row, { outMs: now, outLat: p.lat == null ? '' : p.lat, outLng: p.lng == null ? '' : p.lng, outAcc: Math.round(Number(p.acc) || 0), workedMin: worked, otMin: ot, status: 'Completed' });
  return today(u);
}
function attList(u, p) {
  var N = names_(), id = uid_(u, p);
  return rows_('Attendance').filter(function (r) { return (!id || r.userId === id) && inM_(r.date, p.month); })
    .sort(function (a, b) { return Number(b.inMs) - Number(a.inMs); }).map(function (r) {
      return { id: r.id, userId: r.userId, name: N[r.userId] || '?', date: r.date, dateLabel: fmtD_(r.date), in12: fmtT_(r.inMs), out12: r.outMs ? fmtT_(r.outMs) : '',
        shiftType: r.shiftType, worked: Number(r.workedMin) || 0, ot: Number(r.otMin) || 0, status: r.status,
        inMap: map_(r.inLat, r.inLng), outMap: map_(r.outLat, r.outLng) };
    });
}
function attDelete(u, p) {
  var r = rows_('Attendance').filter(function (x) { return x.id === p.id; })[0];
  if (r) del_('Attendance', r._row);
  return true;
}

/* ---------- Daily work (tasks) ---------- */
function taskAdd(u, p) {
  var title = String(p.title || '').trim();
  if (!title) throw new Error('Enter the work title');
  var date = p.date || today_(), ids = p.userId === 'ALL' || !p.userId
    ? rows_('Users').filter(function (x) { return x.role === 'worker' && x.active === 'TRUE'; }).map(function (x) { return x.id; }) : [p.userId];
  if (!ids.length) throw new Error('No worker selected');
  ids.forEach(function (id) { add_('Tasks', { id: newId_(), userId: id, date: date, title: title, details: p.details || '', status: 'Pending', workerNote: '', createdBy: u.name, doneMs: '' }); });
  return ids.length;
}
function taskList(u, p) {
  var N = names_(), id = uid_(u, p), T = today_();
  return rows_('Tasks').filter(function (r) {
    if (id && r.userId !== id) return false;
    if (u.role === 'admin') return (!p.date || r.date === p.date) && inM_(r.date, p.month);
    return r.date === T || (r.status !== 'Done' && r.date < T);
  }).sort(function (a, b) { return a.date < b.date ? 1 : -1; }).map(function (r) {
    return { id: r.id, userId: r.userId, name: N[r.userId] || '?', date: r.date, dateLabel: fmtD_(r.date), title: r.title, details: r.details, status: r.status, note: r.workerNote, done12: r.doneMs ? fmtDT_(r.doneMs) : '' };
  });
}
function taskDone(u, p) {
  var r = rows_('Tasks').filter(function (x) { return x.id === p.id; })[0];
  if (!r || (u.role !== 'admin' && r.userId !== u.id)) throw new Error('Task not found');
  var done = p.status === 'Done';
  upd_('Tasks', r._row, { status: done ? 'Done' : 'Pending', workerNote: done ? (p.note || '') : '', doneMs: done ? Date.now() : '' });
  return true;
}
function taskDelete(u, p) { var r = rows_('Tasks').filter(function (x) { return x.id === p.id; })[0]; if (r) del_('Tasks', r._row); return true; }

/* ---------- Work updates: note + photo + voice ---------- */
function save_(dataUrl, prefix) {
  var m = String(dataUrl).match(/^data:([^;,]+)[^,]*;base64,(.*)$/);
  if (!m) throw new Error('Bad file data');
  if (m[2].length > 30 * 1024 * 1024) throw new Error('File too large');
  var mime = m[1], ext = (mime.split('/')[1] || 'bin').replace('mpeg', 'mp3').replace('x-m4a', 'm4a').replace('quicktime', 'mov');
  var f = DriveApp.getFolderById(PropertiesService.getScriptProperties().getProperty('FOLDER'))
    .createFile(Utilities.newBlob(Utilities.base64Decode(m[2]), mime, prefix + '.' + ext));
  return f.getId();
}
function logAdd(u, p) {
  var note = String(p.note || '').trim();
  if (!note && !p.photo && !p.audio) throw new Error('Write a note or add a photo / voice note');
  var pre = u.name.replace(/[^A-Za-z0-9]/g, '') + '_' + today_() + '_' + Date.now();
  add_('Logs', { id: newId_(), userId: u.id, date: today_(), ms: Date.now(), note: note, photoId: p.photo ? save_(p.photo, pre + '_photo') : '', audioId: p.audio ? save_(p.audio, pre + '_voice') : '' });
  return true;
}
function logList(u, p) {
  var N = names_(), id = uid_(u, p);
  return rows_('Logs').filter(function (r) { return (!id || r.userId === id) && inM_(r.date, p.month); })
    .sort(function (a, b) { return Number(b.ms) - Number(a.ms); }).map(function (r) {
      return { id: r.id, userId: r.userId, name: N[r.userId] || '?', dateLabel: fmtD_(r.date), time: fmtT_(r.ms), note: r.note, photoId: r.photoId, audioId: r.audioId };
    });
}
function fileGet(u, p) {
  var r = rows_('Logs').filter(function (x) { return x.photoId === p.id || x.audioId === p.id; })[0];
  if (!r || (u.role !== 'admin' && r.userId !== u.id)) throw new Error('Not allowed');
  var b = DriveApp.getFileById(p.id).getBlob();
  return { data: 'data:' + b.getContentType() + ';base64,' + Utilities.base64Encode(b.getBytes()) };
}

/* ---------- Stock used ---------- */
function stockAdd(u, p) {
  var item = String(p.item || '').trim(), qty = Number(p.qty);
  if (!item || !(qty > 0)) throw new Error('Enter item name and quantity');
  add_('Stock', { id: newId_(), userId: u.id, date: today_(), ms: Date.now(), item: item, qty: qty, unit: p.unit || 'pcs', note: p.note || '' });
  return true;
}
function stockList(u, p) {
  var N = names_(), id = uid_(u, p);
  return rows_('Stock').filter(function (r) { return (!id || r.userId === id) && inM_(r.date, p.month); })
    .sort(function (a, b) { return Number(b.ms) - Number(a.ms); }).map(function (r) {
      return { id: r.id, name: N[r.userId] || '?', dateLabel: fmtD_(r.date), time: fmtT_(r.ms), item: r.item, qty: Number(r.qty), unit: r.unit, note: r.note };
    });
}

/* ---------- Salary & payments ---------- */
function salaryRows_(month, id) {
  var att = rows_('Attendance').filter(function (r) { return inM_(r.date, month); }), pay = rows_('Payments').filter(function (r) { return r.month === month; });
  return rows_('Users').filter(function (x) {
    return x.role === 'worker' && (!id || x.id === id) &&
      (x.active === 'TRUE' || att.some(function (r) { return r.userId === x.id; }) || pay.some(function (r) { return r.userId === x.id; }));
  }).map(function (x) {
    var a = att.filter(function (r) { return r.userId === x.id; }), d = {}, py = pay.filter(function (r) { return r.userId === x.id; });
    a.forEach(function (r) { d[r.date] = 1; });
    var ot = sum_(a, 'otMin'), rate = otRate_(x), otPay = Math.round(ot / 60 * rate), base = Number(x.salary) || 0;
    var adv = sum_(py.filter(function (r) { return r.type === 'Advance'; }), 'amount'), paid = sum_(py.filter(function (r) { return r.type === 'Salary'; }), 'amount');
    return { userId: x.id, name: x.name, base: base, days: Object.keys(d).length, workedMin: sum_(a, 'workedMin'), otMin: ot, rate: rate, otPay: otPay,
      total: base + otPay, advance: adv, paid: paid, balance: base + otPay - adv - paid };
  });
}
function salary(u, p) { return salaryRows_(p.month || today_().slice(0, 7), uid_(u, p)); }
function payAdd(u, p) {
  var amt = Number(p.amount);
  if (!p.userId || !(amt > 0)) throw new Error('Choose worker and enter amount');
  add_('Payments', { id: newId_(), userId: p.userId, month: p.month || today_().slice(0, 7), type: p.type === 'Advance' ? 'Advance' : 'Salary', amount: amt, date: p.date || today_(), note: p.note || '', by: u.name });
  return true;
}
function payList(u, p) {
  var N = names_(), id = uid_(u, p);
  return rows_('Payments').filter(function (r) { return (!id || r.userId === id) && (!p.month || p.month === 'ALL' || r.month === p.month); })
    .sort(function (a, b) { return a.date < b.date ? 1 : -1; }).map(function (r) {
      return { id: r.id, name: N[r.userId] || '?', month: r.month, type: r.type, amount: Number(r.amount), dateLabel: fmtD_(r.date), note: r.note };
    });
}
function payDelete(u, p) { var r = rows_('Payments').filter(function (x) { return x.id === p.id; })[0]; if (r) del_('Payments', r._row); return true; }

/* ---------- Admin: workers ---------- */
function users(u) {
  return rows_('Users').map(pub_).sort(function (a, b) { return a.role === b.role ? a.name.localeCompare(b.name) : (a.role === 'admin' ? -1 : 1); });
}
function userSave(u, p) {
  var name = String(p.name || '').trim(), un = String(p.username || '').trim().toLowerCase().replace(/\s/g, ''),
    ph = String(p.phone || '').replace(/\s/g, ''), em = String(p.email || '').trim().toLowerCase();
  if (!name || !un) throw new Error('Name and username are required');
  if (!/^\d{1,2}:\d{2}$/.test(p.shiftStart || '09:00') || !/^\d{1,2}:\d{2}$/.test(p.shiftEnd || '18:00')) throw new Error('Invalid working hours');
  var all = rows_('Users'), mine = p.id ? all.filter(function (x) { return x.id === p.id; })[0] : null;
  if (p.id && !mine) throw new Error('Worker not found');
  all.forEach(function (x) {
    if (x.id === p.id) return;
    [un, ph, em].forEach(function (v) {
      if (v && [x.username, x.phone, x.email].some(function (y) { return y && y.toLowerCase().replace(/\s/g, '') === v; })) throw new Error('"' + v + '" is already used by ' + x.name);
    });
  });
  var d = { name: name, username: un, phone: ph, email: em, shiftStart: p.shiftStart || '09:00', shiftEnd: p.shiftEnd || '18:00',
    shiftType: p.shiftType === 'Night' ? 'Night' : 'Day', salary: Number(p.salary) || 0, otRate: p.otRate === '' || p.otRate == null ? '' : (Number(p.otRate) || 0) };
  if (mine) {
    upd_('Users', mine._row, d);
    if (p.password) { if (String(p.password).length < 5) throw new Error('Password must be at least 5 characters'); setPw_(mine, String(p.password)); }
  } else {
    if (!p.password || String(p.password).length < 5) throw new Error('Give the worker a password (at least 5 characters)');
    d.role = 'worker';
    addUser_(d, String(p.password));
  }
  return true;
}
function userActive(u, p) {
  var x = rows_('Users').filter(function (r) { return r.id === p.id; })[0];
  if (!x || x.role === 'admin') throw new Error('Cannot change this user');
  upd_('Users', x._row, { active: p.active ? 'TRUE' : 'FALSE' });
  return true;
}
function userSetPassword(u, p) {
  var x = rows_('Users').filter(function (r) { return r.id === p.id; })[0];
  if (!x) throw new Error('User not found');
  if (String(p.password || '').length < 5) throw new Error('Password must be at least 5 characters');
  setPw_(x, String(p.password));
  return true;
}
function dash(u) {
  var T = today_(), us = rows_('Users').filter(function (x) { return x.role === 'worker' && x.active === 'TRUE'; }), att = rows_('Attendance'), N = names_();
  var open = att.filter(function (r) { return !r.outMs; }), td = att.filter(function (r) { return r.date === T; }), tk = rows_('Tasks').filter(function (r) { return r.date === T; });
  var seen = {}; td.forEach(function (r) { seen[r.userId] = 1; });
  return {
    workers: us.length, presentToday: Object.keys(seen).length,
    working: open.map(function (r) { return { name: N[r.userId], since: fmtT_(r.inMs), shiftType: r.shiftType, map: map_(r.inLat, r.inLng) }; }),
    absent: us.filter(function (x) { return !seen[x.id] && !open.some(function (r) { return r.userId === x.id; }); }).map(function (x) { return x.name; }),
    tasksTotal: tk.length, tasksDone: tk.filter(function (r) { return r.status === 'Done'; }).length, otToday: hm_(sum_(td, 'otMin'))
  };
}

/* ---------- PDF backup ---------- */
function months_() {
  var m = {};
  rows_('Attendance').forEach(function (r) { m[r.date.slice(0, 7)] = 1; });
  rows_('Payments').forEach(function (r) { m[r.month] = 1; });
  return Object.keys(m).sort();
}
function rep_(type, month, uid) {
  var au = { role: 'admin' }, p = { month: month, userId: uid }, S = [], want = function (t) { return type === 'full' || type === t; };
  var money = function (n) { return 'Rs ' + Number(n).toLocaleString('en-IN'); };
  if (want('workers')) S.push({ title: 'Workers', head: ['Name', 'Username', 'Phone', 'Email', 'Working hours', 'Shift', 'Monthly salary', 'OT rate/hr', 'Active'],
    rows: users(au).filter(function (x) { return x.role === 'worker' && (!uid || x.id === uid); }).map(function (x) { return [x.name, x.username, x.phone, x.email, t12_(x.shiftStart) + ' - ' + t12_(x.shiftEnd), x.shiftType, money(x.salary), money(x.rate), x.active ? 'Yes' : 'No']; }) });
  if (want('attendance')) S.push({ title: 'Attendance', head: ['Date', 'Worker', 'In', 'Out', 'Worked', 'Overtime', 'Shift', 'Check-in location'],
    rows: attList(au, p).map(function (r) { return [r.dateLabel, r.name, r.in12, r.out12 || '-', hm_(r.worked), hm_(r.ot), r.shiftType, r.inMap]; }) });
  if (want('salary')) {
    var ms = month === 'ALL' ? months_() : [month];
    ms.forEach(function (mo) {
      S.push({ title: 'Salary summary - ' + mo, head: ['Worker', 'Base salary', 'Days present', 'Overtime', 'OT pay', 'Total due', 'Advance taken', 'Salary paid', 'Balance'],
        rows: salaryRows_(mo, uid).map(function (r) { return [r.name, money(r.base), r.days, hm_(r.otMin), money(r.otPay), money(r.total), money(r.advance), money(r.paid), money(r.balance)]; }) });
    });
  }
  if (want('payments')) S.push({ title: 'Advance and salary payments', head: ['Date', 'Worker', 'For month', 'Type', 'Amount', 'Note'],
    rows: payList(au, p).map(function (r) { return [r.dateLabel, r.name, r.month, r.type, money(r.amount), r.note]; }) });
  if (want('tasks')) S.push({ title: 'Daily work given', head: ['Date', 'Worker', 'Work', 'Details', 'Status', 'Worker note'],
    rows: taskList(au, p).map(function (r) { return [r.dateLabel, r.name, r.title, r.details, r.status, r.note]; }) });
  if (want('logs')) S.push({ title: 'Work updates', head: ['Date', 'Time', 'Worker', 'Note', 'Photo', 'Voice note'],
    rows: logList(au, p).map(function (r) { return [r.dateLabel, r.time, r.name, r.note, r.photoId ? 'Yes' : '-', r.audioId ? 'Yes' : '-']; }) });
  if (want('stock')) S.push({ title: 'Stock used', head: ['Date', 'Time', 'Worker', 'Item', 'Qty', 'Unit', 'Note'],
    rows: stockList(au, p).map(function (r) { return [r.dateLabel, r.time, r.name, r.item, r.qty, r.unit, r.note]; }) });
  return S;
}
function esc_(s) { return String(s == null ? '' : s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); }
function pdfBlob_(p) {
  var type = p.type || 'full', month = p.month || today_().slice(0, 7), uid = p.userId && p.userId !== 'ALL' ? p.userId : '';
  var secs = rep_(type, month, uid), who = uid ? names_()[uid] : 'All workers';
  var html = '<html><head><style>body{font-family:Arial,sans-serif;font-size:9px}h1{font-size:16px;margin:0}h2{font-size:12px;margin:14px 0 4px}p{margin:2px 0;color:#555}' +
    'table{border-collapse:collapse;width:100%}th,td{border:1px solid #999;padding:3px 4px;text-align:left;vertical-align:top}th{background:#e6edf7}</style></head><body>' +
    '<h1>Work Manager - ' + esc_(type === 'full' ? 'Full backup' : type) + '</h1><p>' + esc_(who) + ' | Month: ' + esc_(month === 'ALL' ? 'All months' : month) + ' | Generated: ' + esc_(fmtDT_(Date.now())) + ' (Delhi time)</p>';
  secs.forEach(function (s) {
    html += '<h2>' + esc_(s.title) + ' (' + s.rows.length + ')</h2><table><tr>' + s.head.map(function (h) { return '<th>' + esc_(h) + '</th>'; }).join('') + '</tr>' +
      (s.rows.length ? s.rows.map(function (r) { return '<tr>' + r.map(function (c) { return '<td>' + esc_(c) + '</td>'; }).join('') + '</tr>'; }).join('') : '<tr><td colspan="' + s.head.length + '">No data</td></tr>') + '</table>';
  });
  var name = 'WorkManager_' + type + '_' + (month === 'ALL' ? 'all' : month) + (uid ? '_' + who.replace(/\W+/g, '') : '') + '.pdf';
  return HtmlService.createHtmlOutput(html + '</body></html>').getBlob().getAs('application/pdf').setName(name);
}
function pdf(u, p) { var b = pdfBlob_(p); return { name: b.getName(), b64: Utilities.base64Encode(b.getBytes()) }; }
function emailBackup(u, p) {
  var b = pdfBlob_(p), P = PropertiesService.getScriptProperties(), to = [P.getProperty('ADMIN_EMAIL'), u.email].filter(function (x, i, a) { return x && a.indexOf(x) === i; }).join(',');
  MailApp.sendEmail({ to: to, subject: 'Work Manager backup - ' + b.getName(), body: 'PDF backup attached. Generated ' + fmtDT_(Date.now()) + ' (Delhi time).', attachments: [b] });
  return { sentTo: to };
}

var HANDLERS = { me: me, logout: logout, changePassword: changePassword, today: today, checkIn: checkIn, checkOut: checkOut, attList: attList, attDelete: attDelete,
  taskAdd: taskAdd, taskList: taskList, taskDone: taskDone, taskDelete: taskDelete, logAdd: logAdd, logList: logList, fileGet: fileGet, stockAdd: stockAdd, stockList: stockList,
  salary: salary, payAdd: payAdd, payList: payList, payDelete: payDelete, users: users, userSave: userSave, userActive: userActive, userSetPassword: userSetPassword,
  dash: dash, pdf: pdf, emailBackup: emailBackup };
