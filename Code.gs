// ===== TAŞIMA PLANI - Google Apps Script Web App =====
// Sheet'te şu sekmeler olmalı: Kisiler, Tatiller, Manuel, Ayarlar, Alicilar (Alicilar isteğe bağlı)

var NOTIFY_DAYS = 7; // Bugünden kaç gün sonrasına kadarki değişiklikler mail atsın

function doGet(e) {
  var action = (e && e.parameter.action) || 'schedule';
  try {
    if (action === 'schedule') return jsonResponse(getConfig());
    if (action === 'addHoliday') return handleAddHoliday(e);
    if (action === 'removeHoliday') return handleRemoveHoliday(e);
    if (action === 'setOverride') return handleSetOverride(e);
    if (action === 'removeOverride') return handleRemoveOverride(e);
    if (action === 'sendTomorrow') {
      if (!checkAuth(e)) return jsonResponse({ error: 'yetkisiz - pin yanlış' });
      return jsonResponse({ sonuc: sendTomorrowDriverEmail() });
    }
    return jsonResponse({ error: 'bilinmeyen istek' });
  } catch (err) {
    return jsonResponse({ error: String(err) });
  }
}

function checkAuth(e) {
  var s = readSettings(SpreadsheetApp.getActiveSpreadsheet());
  return e.parameter.pin && String(e.parameter.pin) === String(s.adminPin);
}

function getSheet(ss, wanted) {
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    var name = sheets[i].getName().trim().toLowerCase();
    if (name === wanted.trim().toLowerCase()) return sheets[i];
  }
  var have = sheets.map(function (s) { return s.getName(); }).join(', ');
  throw new Error('"' + wanted + '" sekmesi bulunamadı. Mevcut sekmeler: ' + have);
}

function fmt(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return v ? String(v).trim() : v;
}

function readSettings(ss) {
  var sh = getSheet(ss, 'Ayarlar');
  var data = sh.getDataRange().getValues();
  var s = {};
  data.forEach(function (row) { if (row[0]) s[row[0]] = row[1]; });
  return {
    startDate: fmt(s['BaslangicTarihi']),
    endDate: fmt(s['BitisTarihi']),
    adminPin: String(s['AdminSifre'] || '')
  };
}

function getConfig() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var pdata = getSheet(ss, 'Kisiler').getDataRange().getValues();
  var people = [], rules = {};
  for (var i = 1; i < pdata.length; i++) {
    var row = pdata[i];
    if (!row[0]) continue;
    var name = String(row[0]).trim();
    people.push(name);
    var days = [];
    for (var d = 0; d < 5; d++) {
      var v = row[d + 1];
      if (v === true || v === 'TRUE' || v === 1 || v === '1') days.push(d + 1);
    }
    rules[name] = days;
  }

  var hdata = getSheet(ss, 'Tatiller').getDataRange().getValues();
  var holidays = [];
  for (var i = 1; i < hdata.length; i++) { if (hdata[i][0]) holidays.push(fmt(hdata[i][0])); }

  var odata = getSheet(ss, 'Manuel').getDataRange().getValues();
  var overrides = {};
  for (var i = 1; i < odata.length; i++) { if (odata[i][0]) overrides[fmt(odata[i][0])] = String(odata[i][1]).trim(); }

  var settings = readSettings(ss);
  return {
    people: people, rules: rules,
    startDate: settings.startDate, endDate: settings.endDate,
    holidays: holidays, overrides: overrides
  };
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// ===== SİTEDEN GELEN DEĞİŞİKLİKLER (değişince mail atar) =====

function handleAddHoliday(e) {
  if (!checkAuth(e)) return jsonResponse({ error: 'yetkisiz - pin yanlış' });
  var date = e.parameter.date;
  var old = getDriverForDate(getConfig(), date);
  var sh = getSheet(SpreadsheetApp.getActiveSpreadsheet(), 'Tatiller');
  sh.appendRow([date]);
  SpreadsheetApp.flush();
  notifyChange(date, old);
  return jsonResponse(getConfig());
}

function handleRemoveHoliday(e) {
  if (!checkAuth(e)) return jsonResponse({ error: 'yetkisiz - pin yanlış' });
  var date = e.parameter.date;
  var old = getDriverForDate(getConfig(), date);
  var sh = getSheet(SpreadsheetApp.getActiveSpreadsheet(), 'Tatiller');
  var data = sh.getDataRange().getValues();
  for (var i = data.length - 1; i >= 1; i--) {
    if (fmt(data[i][0]) === date) sh.deleteRow(i + 1);
  }
  SpreadsheetApp.flush();
  notifyChange(date, old);
  return jsonResponse(getConfig());
}

function handleSetOverride(e) {
  if (!checkAuth(e)) return jsonResponse({ error: 'yetkisiz - pin yanlış' });
  var date = e.parameter.date;
  var old = getDriverForDate(getConfig(), date);
  var sh = getSheet(SpreadsheetApp.getActiveSpreadsheet(), 'Manuel');
  var data = sh.getDataRange().getValues();
  var found = false;
  for (var i = 1; i < data.length; i++) {
    if (fmt(data[i][0]) === date) {
      sh.getRange(i + 1, 2).setValue(e.parameter.person);
      found = true;
      break;
    }
  }
  if (!found) sh.appendRow([date, e.parameter.person]);
  SpreadsheetApp.flush();
  notifyChange(date, old);
  return jsonResponse(getConfig());
}

function handleRemoveOverride(e) {
  if (!checkAuth(e)) return jsonResponse({ error: 'yetkisiz - pin yanlış' });
  var date = e.parameter.date;
  var old = getDriverForDate(getConfig(), date);
  var sh = getSheet(SpreadsheetApp.getActiveSpreadsheet(), 'Manuel');
  var data = sh.getDataRange().getValues();
  for (var i = data.length - 1; i >= 1; i--) {
    if (fmt(data[i][0]) === date) sh.deleteRow(i + 1);
  }
  SpreadsheetApp.flush();
  notifyChange(date, old);
  return jsonResponse(getConfig());
}

// ===== YARDIMCI FONKSİYONLAR =====

function parseDate(s) {
  var p = String(s).split('-');
  return new Date(+p[0], +p[1] - 1, +p[2], 12);
}

function toKey(d) {
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

// Kisiler sekmesinden Email (G) / Telefon (H) / Not (I) okur
function getPeopleDetails() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var data = getSheet(ss, 'Kisiler').getDataRange().getValues();
  var info = {};
  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    info[String(data[i][0]).trim()] = {
      email: String(data[i][6] || '').trim(),
      phone: String(data[i][7] || '').trim(),
      note: String(data[i][8] || '').trim()
    };
  }
  return info;
}

// Alicilar sekmesinden (şoför olmayan) alıcıları okur; sekme yoksa boş döner
function getExtraRecipients() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var list = [];
  try {
    var data = getSheet(ss, 'Alicilar').getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      var mail = String(data[i][1] || '').trim();
      if (mail) list.push({ name: String(data[i][0] || '').trim(), email: mail });
    }
  } catch (err) {}
  return list;
}

// Kisiler + Alicilar birleşik alıcı listesi (aynı mail iki kez eklenmez)
function getAllRecipients(cfg, info) {
  var recipients = [], seen = {};
  function add(name, mail) {
    if (!mail) return;
    var k = mail.toLowerCase();
    if (seen[k]) return;
    seen[k] = true;
    recipients.push({ name: name, email: mail });
  }
  cfg.people.forEach(function (p) { add(p, (info[p] || {}).email); });
  getExtraRecipients().forEach(function (r) { add(r.name, r.email); });
  return recipients;
}

// Verilen tarihteki sürücüyü web uygulamasındaki buildSchedule ile AYNI
// mantıkla hesaplar.  Kullanım sayısı eşitse en uzun süredir sürmeyen seçilir.
// Bu fonksiyonu sendTomorrowDriverEmail, sendTodayDriverEmail ve notifyChange
// ortak olarak kullanır; böylece maildeki sürücü ile ekrandaki sürücü eşleşir.
function getDriverForDate(cfg, targetKey) {
  var target = parseDate(targetKey);
  var wd = target.getDay();
  if (wd === 0 || wd === 6) return null;
  if ((cfg.holidays || []).indexOf(targetKey) !== -1) return null;
  if (cfg.startDate && targetKey < cfg.startDate) return null;
  if (cfg.endDate && targetKey > cfg.endDate) return null;

  var usage = {}, lastSeen = {};
  cfg.people.forEach(function (p) {
    usage[p] = 0;
    lastSeen[p] = -1;
  });
  var idx = 0;
  var start = parseDate(cfg.startDate || targetKey);
  var holidays = cfg.holidays || [];
  var overrides = cfg.overrides || {};

  for (var day = new Date(start); toKey(day) <= targetKey; day.setDate(day.getDate() + 1)) {
    var key = toKey(day);
    var dayOfWeek = day.getDay();
    if (dayOfWeek === 0 || dayOfWeek === 6) continue;
    if (holidays.indexOf(key) !== -1) continue;

    var required = cfg.people.filter(function (p) {
      return (cfg.rules[p] || []).indexOf(dayOfWeek) !== -1;
    });
    // Web uygulamasında hiç kimsenin uygun olmadığı gün, tatil sayılır.
    if (required.length === 0) continue;

    var driver = overrides[key];
    if (!driver) {
      if (required.length === 1) {
        driver = required[0];
      } else {
        driver = required.slice().sort(function (a, b) {
          if (usage[a] !== usage[b]) return usage[a] - usage[b];
          return lastSeen[a] - lastSeen[b];
        })[0];
      }
    }

    // Frontenddeki usage[driver]++ ve lastSeen[driver]=idx++ ile eşdeğer.
    usage[driver] = (usage[driver] || 0) + 1;
    lastSeen[driver] = idx++;
    if (key === targetKey) return driver;
  }
  return null;
}

// ===== YARININ ŞOFÖRÜNÜ MAİLLE BİLDİR =====

function sendTomorrowDriverEmail() {
  var cfg = getConfig();
  var info = getPeopleDetails();

  var t = new Date();
  t.setDate(t.getDate() + 1);
  var key = toKey(t);

  var driver = getDriverForDate(cfg, key);
  if (!driver) return 'Yarın (' + key + ') şoförlük yok, mail gönderilmedi.';

  var d = info[driver] || {};
  var gun = Utilities.formatDate(t, Session.getScriptTimeZone(), 'dd.MM.yyyy');
  var recipients = getAllRecipients(cfg, info);

  recipients.forEach(function (r) {
    var body = 'Merhaba ' + r.name + ',\n\n';
    body += (r.name === driver)
      ? 'Yarın (' + gun + ') şoför sizsiniz.\n\n'
      : 'Yarın (' + gun + ') şoförümüz: ' + driver + '\n\n';
    if (d.phone) body += 'Telefon: ' + d.phone + '\n';
    if (d.note)  body += 'Not: ' + d.note + '\n';
    body += '\nİyi yolculuklar!';
    MailApp.sendEmail(r.email, 'Yarının şoförü: ' + driver + ' (' + gun + ')', body);
  });

  return driver + ' için ' + recipients.length + ' kişiye mail gönderildi.';
}

// ===== BUGÜNÜN ŞOFÖRÜNÜ MAİLLE BİLDİR (elle çalıştırmak için) =====

function sendTodayDriverEmail() {
  var cfg = getConfig();
  var info = getPeopleDetails();
  var t = new Date();
  var key = toKey(t);

  var driver = getDriverForDate(cfg, key);
  if (!driver) { Logger.log('Bugün şoförlük yok.'); return; }

  var d = info[driver] || {};
  var gun = Utilities.formatDate(t, Session.getScriptTimeZone(), 'dd.MM.yyyy');
  var recipients = getAllRecipients(cfg, info);

  recipients.forEach(function (r) {
    var body = 'Merhaba ' + r.name + ',\n\n';
    body += (r.name === driver)
      ? 'Bugün (' + gun + ') şoför sizsiniz.\n\n'
      : 'Bugün (' + gun + ') şoförümüz: ' + driver + '\n\n';
    if (d.phone) body += 'Telefon: ' + d.phone + '\n';
    if (d.note)  body += 'Not: ' + d.note + '\n';
    body += '\nİyi yolculuklar!';
    MailApp.sendEmail(r.email, 'Bugünün şoförü: ' + driver + ' (' + gun + ')', body);
  });

  Logger.log(driver + ' için ' + recipients.length + ' kişiye mail gönderildi.');
}

// ===== DEĞİŞİKLİK MAİLİ (siteden gün/şoför değişince) =====

function notifyChange(dateKey, oldDriver) {
  try {
    dateKey = fmt(dateKey);
    var today = toKey(new Date());
    var limit = new Date();
    limit.setDate(limit.getDate() + NOTIFY_DAYS);
    if (dateKey < today || dateKey > toKey(limit)) return;

    var cfg = getConfig();
    var info = getPeopleDetails();
    var newDriver = getDriverForDate(cfg, dateKey);
    if (newDriver === oldDriver) return;

    var gun = Utilities.formatDate(parseDate(dateKey), Session.getScriptTimeZone(), 'dd.MM.yyyy');
    var d = newDriver ? (info[newDriver] || {}) : {};
    var recipients = getAllRecipients(cfg, info);

    var subject = newDriver
      ? 'Şoför değişikliği: ' + gun + ' → ' + newDriver
      : 'Plan değişikliği: ' + gun + ' şoförlük yok';

    recipients.forEach(function (r) {
      var body = 'Merhaba ' + r.name + ',\n\n';
      body += 'Plan değişti.\n\n';
      if (newDriver) {
        body += gun + ' tarihinde ';
        body += (r.name === newDriver) ? 'şoför artık siz olacaksınız.\n' : 'şoförümüz: ' + newDriver + '\n';
        if (oldDriver && oldDriver !== newDriver) body += 'Önceki şoför: ' + oldDriver + '\n';
        if (d.phone) body += '\nTelefon: ' + d.phone + '\n';
        if (d.note)  body += 'Not: ' + d.note + '\n';
      } else {
        body += gun + ' tarihinde şoförlük yok.\n';
        if (oldDriver) body += 'Önceki şoför: ' + oldDriver + '\n';
      }
      MailApp.sendEmail(r.email, subject, body);
    });
  } catch (err) {
    Logger.log('notifyChange hatası: ' + err);
  }
}

// ===== ZAMANLAYICI VE TEST =====

// Bir kez çalıştır: her gün 18:00-19:00 arası yarının şoförünü otomatik gönderir
function createDailyTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'sendTomorrowDriverEmail') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('sendTomorrowDriverEmail')
    .timeBased().everyDays(1).atHour(18).create();
}

// Test: sadece günlüğe yazar, mail GÖNDERMEZ
function testYarinkiSofor() {
  var cfg = getConfig();
  var t = new Date(); t.setDate(t.getDate() + 1);
  Logger.log(getDriverForDate(cfg, toKey(t)));
}