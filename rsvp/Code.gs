/**
 * RSVP backend for the wedding site. Runs as a Google Apps Script web app
 * bound to a Google Sheet with two tabs:
 *
 *   Guests     A: code   B: names   C: seats   D: email (optional)   E: link (filled for you)
 *   Responses  written by this script, one row per reply
 *
 * GET  ?i=CODE   -> { ok, name, seats, email }   used by the page to personalise
 * POST JSON body -> { ok, party, emailed }        validated against the Guests tab, appended to Responses,
 *                                                 and a confirmation is emailed to the guest
 *
 * The Invitations menu in the sheet fills codes and links, and sends the invitation emails.
 * Setup is in README.md next to this file.
 */

var SITE_URL = 'https://example.com/';   // your site, with trailing slash; used to build invitation links
var FROM_NAME = 'Ashley & Charles';       // sender name on every email (the address is the Google account running this)
var REPLY_TO = '';                        // optional: where guest replies to the emails should go
var INVITE_SUBJECT = 'Bali, 28 August 2027: your invitation';
var CONFIRM_SUBJECT = 'We have your reply';
var RSVP_DEADLINE = '1 March 2027';
var GUESTS = 'Guests';
var RESPONSES = 'Responses';
var MAX_SEATS = 6;
var CODE_ALPHABET = 'abcdefghjkmnpqrstuvwxyz23456789'; // no 0/o/1/l, so codes survive being read aloud
var CODE_LENGTH = 6;
var RESPONSE_HEADERS = ['Received', 'Code', 'Invited', 'Seats', 'Name', 'Email', 'Attending', 'Party', 'Dietary', 'Note', 'Sent'];
var ATTENDING_TEXT = { both: 'Friday and Saturday', saturday: 'Saturday only', no: "Can't make it" };

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
        row: i + 1,
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
    var dietary = attending === 'no' ? '' : clean(data.dietary, 200);
    var note = clean(data.note, 1000);
    sheet.appendRow([
      new Date(), guest.code, guest.name, guest.seats,
      name, email, attending, party, dietary, note, sent
    ]);
    var emailed = false;
    try {
      sendConfirmation(guest, { name: name, email: email, attending: attending, party: party, dietary: dietary, note: note });
      emailed = true;
    } catch (mailErr) {
      /* the reply is recorded either way; a failed receipt is not a failed RSVP */
    }
    return json({ ok: true, party: party, emailed: emailed });
  } catch (err) {
    return json({ ok: false, error: 'server', message: String(err) });
  } finally {
    lock.releaseLock();
  }
}

/* ---- Email ---------------------------------------------------------------- */

function escapeHtml(v) {
  return String(v == null ? '' : v).replace(/[&<>"]/g, function (c) {
    return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c];
  });
}

/* A quiet, single-column email that matches the site. Inline styles only: email clients strip the rest. */
function emailShell(title, bodyHtml, buttonText, buttonUrl) {
  var button = buttonUrl
    ? '<p style="margin:28px 0 0"><a href="' + escapeHtml(buttonUrl) + '" style="display:inline-block;background:#242722;color:#F2EDE3;text-decoration:none;padding:14px 26px;border-radius:999px;font-weight:600">' + escapeHtml(buttonText) + '</a></p>' +
      '<p style="margin:14px 0 0;font-size:13px;color:#7A7C74">Or copy this link: <a href="' + escapeHtml(buttonUrl) + '" style="color:#3F5B45">' + escapeHtml(buttonUrl) + '</a></p>'
    : '';
  return '<div style="background:#E9E4D9;padding:32px 16px;font-family:-apple-system,Segoe UI,Helvetica,Arial,sans-serif;color:#2A2C27;line-height:1.55">' +
    '<div style="max-width:560px;margin:0 auto;background:#F2EDE3;padding:36px 32px;border-radius:6px">' +
    '<p style="margin:0 0 6px;font-family:Menlo,Consolas,monospace;font-size:11px;letter-spacing:.18em;text-transform:uppercase;color:#8C6E4F">Ashley &amp; Charles &middot; AYANA Bali &middot; 28 August 2027</p>' +
    '<h1 style="margin:0 0 20px;font-family:Georgia,Times New Roman,serif;font-weight:400;font-size:30px;line-height:1.1">' + escapeHtml(title) + '</h1>' +
    bodyHtml + button +
    '</div></div>';
}

function mailOptions(to, subject, html, text) {
  var opts = { to: to, subject: subject, htmlBody: html, body: text, name: FROM_NAME };
  if (REPLY_TO) opts.replyTo = REPLY_TO;
  return opts;
}

function guestLink(code) { return SITE_URL + '?i=' + code; }

function sendConfirmation(guest, reply) {
  if (!reply.email) return;
  var lines = [
    ['Attending', ATTENDING_TEXT[reply.attending] || reply.attending],
    ['Party', reply.attending === 'no' ? '' : (reply.party === 1 ? 'Just you' : reply.party + ' of you')],
    ['Dietary', reply.dietary],
    ['Note', reply.note]
  ].filter(function (l) { return l[1]; });
  var rows = lines.map(function (l) {
    return '<tr><td style="padding:6px 16px 6px 0;font-family:Menlo,Consolas,monospace;font-size:11px;letter-spacing:.14em;text-transform:uppercase;color:#5B5D55;vertical-align:top;white-space:nowrap">' + escapeHtml(l[0]) + '</td><td style="padding:6px 0;vertical-align:top">' + escapeHtml(l[1]) + '</td></tr>';
  }).join('');
  var lead = reply.attending === 'no'
    ? "Thank you for letting us know. We'll miss you, and we'll raise a glass to you from the cliff."
    : "Thank you. Your reply is in, and we'll be in touch with booking details for the estate.";
  var html = emailShell(
    reply.attending === 'no' ? "We'll miss you" : "We can't wait",
    '<p style="margin:0 0 18px">' + escapeHtml(lead) + '</p>' +
    '<table style="border-collapse:collapse;margin:0 0 6px">' + rows + '</table>' +
    '<p style="margin:22px 0 0;font-size:14px;color:#5B5D55">Need to change something? Open your invitation again and send a new reply any time before ' + escapeHtml(RSVP_DEADLINE) + '. The latest one counts.</p>',
    'Open your invitation', guestLink(guest.code)
  );
  var text = lead + '\n\n' + lines.map(function (l) { return l[0] + ': ' + l[1]; }).join('\n') +
    '\n\nTo change your reply before ' + RSVP_DEADLINE + ', open your invitation again: ' + guestLink(guest.code);
  MailApp.sendEmail(mailOptions(reply.email, CONFIRM_SUBJECT, html, text));
}

function sendInvitation(guest, toList) {
  var html = emailShell(
    'You\'re invited',
    '<p style="margin:0 0 14px">Dear ' + escapeHtml(guest.name) + ',</p>' +
    '<p style="margin:0 0 14px">We\'re getting married on a cliff above the Indian Ocean, and we would love you to be there. Your invitation is below. It opens like a letter, so give it a moment, and it carries your name and your seats.</p>' +
    '<p style="margin:0">Please reply by ' + escapeHtml(RSVP_DEADLINE) + '.</p>',
    'Open your invitation', guestLink(guest.code)
  );
  var text = 'Dear ' + guest.name + ',\n\nWe\'re getting married on a cliff above the Indian Ocean, and we would love you to be there. Your invitation is here: ' +
    guestLink(guest.code) + '\n\nPlease reply by ' + RSVP_DEADLINE + '.\n\nAshley & Charles';
  MailApp.sendEmail(mailOptions(toList.join(','), INVITE_SUBJECT, html, text));
}

/* ---- Sheet helpers, run from the Invitations menu ---------------------- */

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Invitations')
    .addItem('Fill missing codes and links', 'fillCodes')
    .addItem('Send invitations to unsent rows', 'sendInvitations')
    .addItem('Send a test invitation to me', 'sendTestInvitation')
    .addToUi();
}

function splitEmails(v) {
  return clean(v, 400).split(/[,;\s]+/).filter(function (e) { return /^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(e); });
}

/** Emails every Guests row that has a code and an email but no value in the "sent" column (F). */
function sendInvitations() {
  var ui = SpreadsheetApp.getUi();
  if (SITE_URL.indexOf('example.com') >= 0) { ui.alert('Set SITE_URL at the top of the script first.'); return; }
  var sheet = ss().getSheetByName(GUESTS);
  var rows = sheet.getDataRange().getValues();
  var pending = [];
  for (var r = 1; r < rows.length; r++) {
    var code = normaliseCode(rows[r][0]);
    var emails = splitEmails(rows[r][3]);
    var sent = rows[r][5];
    if (code && emails.length && !sent) pending.push({ row: r + 1, code: code, name: clean(rows[r][1], 80), emails: emails });
  }
  if (!pending.length) { ui.alert('Nothing to send: every row with a code and an email is already marked as sent.'); return; }
  var quota = MailApp.getRemainingDailyQuota();
  var answer = ui.alert('Send invitations', pending.length + ' household(s) have not been sent an invitation. Your account can send ' + quota + ' more emails today. Send now?', ui.ButtonSet.YES_NO);
  if (answer !== ui.Button.YES) return;
  var done = 0;
  for (var i = 0; i < pending.length; i++) {
    if (MailApp.getRemainingDailyQuota() < 1) break;
    sendInvitation(pending[i], pending[i].emails);
    sheet.getRange(pending[i].row, 6).setValue(new Date());
    done++;
  }
  ui.alert('Sent ' + done + ' of ' + pending.length + '.' + (done < pending.length ? ' The rest are still unsent; run this again tomorrow when the daily limit resets.' : ''));
}

/** Sends the invitation email for the first Guests row to the account running the script, so you can see it. */
function sendTestInvitation() {
  var sheet = ss().getSheetByName(GUESTS);
  var rows = sheet.getDataRange().getValues();
  for (var r = 1; r < rows.length; r++) {
    var code = normaliseCode(rows[r][0]);
    if (code) {
      sendInvitation({ code: code, name: clean(rows[r][1], 80) }, [Session.getEffectiveUser().getEmail()]);
      SpreadsheetApp.getUi().alert('Test sent to ' + Session.getEffectiveUser().getEmail() + ' using row ' + (r + 1) + '.');
      return;
    }
  }
  SpreadsheetApp.getUi().alert('No row has a code yet. Run "Fill missing codes and links" first.');
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
