/**
 * YS Scholaria backend (Google Apps Script).
 * Stores visitor counts, work requests, orders and reviews in the Google Sheet this script is attached to.
 * One-time setup is described in SETUP.txt.
 */
var TZ = 'Asia/Dhaka';

function doGet(e) {
  var p = (e && e.parameter) || {};
  var action = p.action || 'stats';
  var out;
  if (action === 'reviews') {
    out = reviews_();
  } else if (action === 'orders') {
    out = orders_(String(p.ids || '').split(',').filter(function (x) { return x; }));
  } else {
    out = stats_();
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  var d = {};
  try {
    d = JSON.parse((e && e.postData && e.postData.contents) || '{}');
  } catch (err) {
    return ContentService.createTextOutput('bad request');
  }
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var when = new Date();
    if (d.type === 'view') {
      bump_(ss, 'views');
    } else if (d.type === 'request') {
      bump_(ss, 'requests');
      append_(ss, 'Requests', [when, d.id, d.name, d.contact, d.services, d.total, d.deadline, d.details]);
    } else if (d.type === 'order') {
      bump_(ss, 'requests');
      var status = d.mode === 'balance' ? 'Balance submitted' : 'Advance submitted';
      append_(ss, 'Orders', [when, d.id, d.mode, d.name, d.contact, d.services, d.total, d.advance, d.balance,
                             d.amount, d.method, d.sender, d.trx, status]);
    } else if (d.type === 'review') {
      append_(ss, 'Reviews', [when, d.name, d.service, d.stars, d.text, 'No']);
    }
  } finally {
    lock.releaseLock();
  }
  return ContentService.createTextOutput('ok');
}

function stats_() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Stats');
  return {
    views: Number(sh.getRange('B1').getValue()) || 0,
    requests: Number(sh.getRange('B2').getValue()) || 0
  };
}

function bump_(ss, key) {
  var sh = ss.getSheetByName('Stats');
  var cell = key === 'views' ? 'B1' : 'B2';
  sh.getRange(cell).setValue((Number(sh.getRange(cell).getValue()) || 0) + 1);
}

function append_(ss, name, row) {
  ss.getSheetByName(name).appendRow(row);
}

/* Status for the dashboard: the owner edits column N (Status) on the advance row of each order. */
function orders_(ids) {
  var rows = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Orders').getDataRange().getValues().slice(1);
  var out = [];
  ids.forEach(function (id) {
    var hit = null;
    rows.forEach(function (r) {
      if (!hit && String(r[1]) === id && String(r[2]) === 'advance') hit = r;
    });
    if (hit) out.push({ id: id, status: String(hit[13] || 'Advance submitted') });
  });
  return out;
}

/* Only rows marked Yes in the Approved column (F) are published. */
function reviews_() {
  var rows = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Reviews').getDataRange().getValues().slice(1);
  return rows
    .filter(function (r) { return String(r[5]).toLowerCase() === 'yes'; })
    .map(function (r) {
      return {
        date: r[0] instanceof Date ? Utilities.formatDate(r[0], TZ, 'yyyy-MM-dd') : String(r[0]),
        name: String(r[1]),
        service: String(r[2]),
        stars: Number(r[3]) || 5,
        text: String(r[4])
      };
    });
}
