/**
 * RSVP backend for the wedding site. Runs as a Google Apps Script web app
 * bound to a Google Sheet with two tabs:
 *
 *   Guests     A: code   B: names   C: seats   D: email (optional)   E: link (filled for you)
 *   Responses  written by this script, one row per reply
 *
 * GET  ?i=CODE   -> { ok, name, seats, email }   used by the page to personalise
 * POST JSON body -> { ok, party }                 validated against the Guests tab, appended to Responses
 *
 * Setup is in README.md next to this file.
 */

var SITE_URL = 'https://example.com/';   // your site, with trailing slash; used to build invitation links
var GUESTS = 'Guests';
var RESPONSES = 'Responses';
var MAX_SEATS = 6;
var CODE_ALPHABET = 'abcdefghjkmnpqrstuvwxyz23456789'; // no 0/o/1/l, so codes survive being read aloud
var CODE_LENGTH = 6;
var RESPONSE_HEADERS = ['Received', 'Code', 'Invited', 'Seats', 'Name', 'Email', 'Attending', 'Party', 'Dietary', 'Note', 'Sent'];

function ss() { return SpreadsheetApp.getActive(); }

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function clean(v, max) {
  return String(v == null ? '' : v).replace(/\s+/g, ' ').trim().slice(0, max);
}

function normaliseCode(code) {
  return String(code || '').trim().toLowerCase().replace(/[^a-z0-9]/g, '');
}

function findGuest(code) {
  code = normaliseCode(code);
  if (!code) return null;
  var sheet = ss().getSheetByName(GUESTS);
  if (!sheet) return null;
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (normaliseCode(rows[i][0]) === code) {
      var seats = parseInt(rows[i][2], 10);
      return {
        code: code,
        name: clean(rows[i][1], 80),
        seats: Math.max(1, Math.min(MAX_SEATS, isNaN(seats) ? 1 : seats)),
        email: clean(rows[i][3], 120)
      };
    }
  }
  return null;
}

/* ---- Web app ------------------------------------------------------------ */

function doGet(e) {
  var code = e && e.parameter ? e.parameter.i : '';
  if (!code) return json({ ok: true, service: 'rsvp' });
  var guest = findGuest(code);
  if (!guest) return json({ ok: false, error: 'not_found' });
  return json({ ok: true, name: guest.name, seats: guest.seats, email: guest.email });
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.tryLock(10000);
  try {
    var data = JSON.parse(e.postData.contents || '{}');
    var guest = findGuest(data.code);
    if (!guest) return json({ ok: false, error: 'not_found' });

    var attending = ['both', 'saturday', 'no'].indexOf(data.attending) >= 0 ? data.attending : '';
    var name = clean(data.name, 80);
    var email = clean(data.email, 120);
    if (!attending || !name || !/^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(email)) {
      return json({ ok: false, error: 'bad_request' });
    }
    var party = attending === 'no' ? 0 : Math.max(1, Math.min(guest.seats, parseInt(data.party, 10) || 1));

    var sheet = ensureResponses();
    var sent = clean(data.submittedAt, 40);
    /* The page retries once if the browser can't read our reply. Same code and same
       client timestamp within the last few rows means we already have this one. */
    var last = sheet.getLastRow();
    if (sent && last > 1) {
      var recent = sheet.getRange(Math.max(2, last - 9), 1, Math.min(10, last - 1), RESPONSE_HEADERS.length).getValues();
      for (var i = 0; i < recent.length; i++) {
        if (recent[i][1] === guest.code && String(recent[i][10]) === sent) return json({ ok: true, party: party, duplicate: true });
      }
    }
    sheet.appendRow([
      new Date(), guest.code, guest.name, guest.seats,
      name, email, attending, party,
      attending === 'no' ? '' : clean(data.dietary, 200),
      clean(data.note, 1000), sent
    ]);
    return json({ ok: true, party: party });
  } catch (err) {
    return json({ ok: false, error: 'server', message: String(err) });
  } finally {
    lock.releaseLock();
  }
}

/* ---- Sheet helpers, run from the Invitations menu ---------------------- */

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Invitations')
    .addItem('Fill missing codes and links', 'fillCodes')
    .addToUi();
}

function ensureResponses() {
  var sheet = ss().getSheetByName(RESPONSES);
  if (!sheet) {
    sheet = ss().insertSheet(RESPONSES);
    sheet.appendRow(RESPONSE_HEADERS);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function randomCode() {
  var out = '';
  for (var i = 0; i < CODE_LENGTH; i++) {
    out += CODE_ALPHABET.charAt(Math.floor(Math.random() * CODE_ALPHABET.length));
  }
  return out;
}

/** Gives every household row a unique code and an invitation link. Never overwrites an existing code. */
function fillCodes() {
  var sheet = ss().getSheetByName(GUESTS);
  if (!sheet) throw new Error('No sheet named "' + GUESTS + '"');
  ensureResponses();
  var range = sheet.getDataRange();
  var rows = range.getValues();
  var used = {};
  for (var i = 1; i < rows.length; i++) used[normaliseCode(rows[i][0])] = true;
  for (var r = 1; r < rows.length; r++) {
    var hasHousehold = clean(rows[r][1], 80) !== '';
    if (!hasHousehold) continue;
    var code = normaliseCode(rows[r][0]);
    if (!code) {
      do { code = randomCode(); } while (used[code]);
      used[code] = true;
      sheet.getRange(r + 1, 1).setValue(code);
    }
    sheet.getRange(r + 1, 5).setValue(SITE_URL + '?i=' + code);
  }
}
