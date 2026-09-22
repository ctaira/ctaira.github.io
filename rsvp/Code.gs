/**
 * RSVP backend for the wedding site. Runs as a Google Apps Script web app
 * bound to a Google Sheet with two tabs:
 *
 *   Guests     A: code   B: names   C: seats   D: email (optional)   E: link (filled for you)
 *   Responses  written by this script, one row per reply
 *
 * GET  ?i=CODE   -> { ok, name, seats, email }   used by the page to personalise
 * POST JSON body -> { ok, emailed }               validated against the Households tab, one Responses row per person,
 *                                                 and a confirmation is emailed to the guest
 *
 * The Invitations menu in the sheet fills codes and links, and sends the invitation emails.
 * Setup is in README.md next to this file.
 */

var SITE_URL = 'https://ashleyandcharlesinbali.com/';   // the live site, with trailing slash; used to build invitation links and to load the email artwork
var FROM_NAME = 'Ashley & Charles';       // sender name on every email (the address is the Google account running this)
var REPLY_TO = '';                        // optional: where guest replies to the emails should go
var INVITE_SUBJECT = 'Ashley & Charles are getting married in Bali';
var CONFIRM_SUBJECT_COMING = 'See you in Bali';
var CONFIRM_SUBJECT_NOT = "We'll miss you in Bali";
var RSVP_DEADLINE = 'January 31, 2027';
var GUEST_LIST = 'Guest List';   // the couple's own list: one row per person, with "Linked to another guest?"
var GUESTS = 'Households';       // built from it by the Invitations menu: one row per household
var LINK_COLUMN = 9;             // Guest List column that receives each person's invite link (9 = I)
var SCRIPT_VERSION = 6;          // shown in the Invitations menu's messages, so you can tell which copy is running
var RESPONSES = 'Responses';
var MAX_SEATS = 6;
var CODE_ALPHABET = 'abcdefghjkmnpqrstuvwxyz23456789'; // no 0/o/1/l, so codes survive being read aloud
var CODE_LENGTH = 6;
var RESPONSE_HEADERS = ['Received', 'Code', 'Household', 'Replied by', 'Email', 'Person', 'Answer', 'Dietary', 'Note', 'Sent'];   // one row per person per reply
var ATTENDING_TEXT = { both: 'Both days', saturday: 'Saturday only', no: 'Not coming' };

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
        email: clean(rows[i][3], 120),
        members: String(rows[i][6] || '').split('|').map(function (m) { return clean(m, 60); }).filter(Boolean)
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
  return json({ ok: true, name: guest.name, seats: guest.seats, members: guest.members });
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.tryLock(10000);
  try {
    var data = JSON.parse(e.postData.contents || '{}');
    var guest = findGuest(data.code);
    if (!guest) return json({ ok: false, error: 'not_found' });

    var name = clean(data.name, 80);
    var email = clean(data.email, 120);
    if (!name || !/^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(email)) return json({ ok: false, error: 'bad_request' });
    /* one answer per household member; names must be the household's own, answers one of three, at least one given */
    var list = guest.members.length ? guest.members : [guest.name];
    var byName = {}; list.forEach(function (m) { byName[nameKey(m)] = m; });
    var people = [], seen = {};
    (Array.isArray(data.people) ? data.people : []).forEach(function (p) {
      var who = byName[nameKey(p && p.name)], ans = p && ['both', 'saturday', 'no'].indexOf(p.answer) >= 0 ? p.answer : '';
      if (who && !seen[who]) { seen[who] = true; people.push({ name: who, answer: ans }); }
    });
    if (!people.some(function (p) { return p.answer; })) return json({ ok: false, error: 'bad_request' });
    var coming = people.some(function (p) { return p.answer === 'both' || p.answer === 'saturday'; });

    var sheet = ensureResponses();
    var sent = clean(data.submittedAt, 40);
    /* The page retries once if the browser can't read our reply. Same code and same
       client timestamp within the last few rows means we already have this one. */
    var last = sheet.getLastRow();
    if (sent && last > 1) {
      var recent = sheet.getRange(Math.max(2, last - 19), 1, Math.min(20, last - 1), RESPONSE_HEADERS.length).getValues();
      for (var i = 0; i < recent.length; i++) {
        if (recent[i][1] === guest.code && String(recent[i][9]) === sent) return json({ ok: true, duplicate: true });
      }
    }
    var dietary = coming ? clean(data.dietary, 200) : '';
    var note = clean(data.note, 1000);
    var when = new Date();
    var out = people.filter(function (p) { return p.answer; }).map(function (p) {
      return [when, guest.code, guest.name, name, email, p.name, ATTENDING_TEXT[p.answer], dietary, note, sent];
    });
    sheet.getRange(sheet.getLastRow() + 1, 1, out.length, RESPONSE_HEADERS.length).setValues(out);
    try { writeStatuses(); } catch (e) { /* the Guest List's status column is a convenience; the reply is recorded above */ }
    var emailed = false;
    try {
      sendConfirmation(guest, { name: name, email: email, people: people, coming: coming, dietary: dietary, note: note });
      emailed = true;
    } catch (mailErr) {
      /* the reply is recorded either way; a failed receipt is not a failed RSVP */
    }
    return json({ ok: true, emailed: emailed });
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

function guestLink(code, personIdx) { return SITE_URL + '?i=' + code + (personIdx ? '&p=' + personIdx : ''); }

function sendConfirmation(guest, reply) {
  if (!reply.email) return;
  var lines = reply.people.filter(function (p) { return p.answer; }).map(function (p) { return [p.name, ATTENDING_TEXT[p.answer]]; });
  var unanswered = reply.people.filter(function (p) { return !p.answer; }).map(function (p) { return p.name; });
  if (reply.dietary) lines.push(['Dietary', reply.dietary]);
  if (reply.note) lines.push(['Note', reply.note]);
  var rows = lines.map(function (l) {
    return '<tr><td style="padding:6px 16px 6px 0;font-family:Menlo,Consolas,monospace;font-size:11px;letter-spacing:.14em;text-transform:uppercase;color:#5B5D55;vertical-align:top;white-space:nowrap">' + escapeHtml(l[0]) + '</td><td style="padding:6px 0;vertical-align:top">' + escapeHtml(l[1]) + '</td></tr>';
  }).join('');
  var lead = reply.coming
    ? "Thank you. Your reply is in, and we'll be in touch with booking details for the estate."
    : "Thank you for letting us know. We'll miss you, and we'll raise a glass to you from the cliff.";
  var later = unanswered.length ? '<p style="margin:14px 0 0;font-size:14px;color:#5B5D55">Still to reply: ' + escapeHtml(unanswered.join(', ')) + '. They can answer any time from the same invitation link.</p>' : '';
  var html = emailShell(
    reply.coming ? "We can't wait" : "We'll miss you",
    '<p style="margin:0 0 18px">' + escapeHtml(lead) + '</p>' +
    '<table style="border-collapse:collapse;margin:0 0 6px">' + rows + '</table>' + later +
    '<p style="margin:22px 0 0;font-size:14px;color:#5B5D55">Need to change something? Open your invitation again and send a new reply any time before ' + escapeHtml(RSVP_DEADLINE) + '. The latest answer for each person counts.</p>',
    'Open your invitation', guestLink(guest.code)
  );
  var text = lead + '\n\n' + lines.map(function (l) { return l[0] + ': ' + l[1]; }).join('\n') +
    (unanswered.length ? '\n\nStill to reply: ' + unanswered.join(', ') + '.' : '') +
    '\n\nTo change your reply before ' + RSVP_DEADLINE + ', open your invitation again: ' + guestLink(guest.code);
  MailApp.sendEmail(mailOptions(reply.email, reply.coming ? CONFIRM_SUBJECT_COMING : CONFIRM_SUBJECT_NOT, html, text));
}

/** Fills the Guest List's "RSVP Status" column, if it has one, with each person's latest answer. */
function writeStatuses() {
  var src = ss().getSheetByName(GUEST_LIST), resp = ss().getSheetByName(RESPONSES);
  if (!src || !resp) return;
  var rows = src.getDataRange().getValues(), h = -1, c, r, col = -1;
  for (r = 0; r < rows.length && h < 0; r++) for (c = 0; c < rows[r].length; c++) if (nameKey(rows[r][c]) === 'first name') { h = r; break; }
  if (h < 0) return;
  for (c = 0; c < rows[h].length; c++) if (nameKey(rows[h][c]) === 'rsvp status') col = c;
  if (col < 0) return;
  var codeOf = {}, hs = ss().getSheetByName(GUESTS).getDataRange().getValues();
  for (r = 1; r < hs.length; r++) codeOf[hs[r][6]] = normaliseCode(hs[r][0]);
  var latest = {}, rs = resp.getDataRange().getValues();      /* later rows win */
  for (r = 1; r < rs.length; r++) latest[normaliseCode(rs[r][1]) + '|' + nameKey(rs[r][5])] = rs[r][6];
  var atRow = {};
  readHouseholds().forEach(function (house) { house.rows.forEach(function (rr, i) { atRow[rr] = { code: codeOf[house.members], name: house.fulls[i] }; }); });
  var out = [], last = src.getLastRow();
  for (r = h + 2; r <= last; r++) { var info = atRow[r]; out.push([info && info.code ? (latest[info.code + '|' + nameKey(info.name)] || '') : '']); }
  if (out.length) src.getRange(h + 2, col + 1, out.length, 1).setValues(out);
}

/** The invitation email is the artwork alone, the whole picture a link to the household's invitation.
    The artwork travels inside the email (see heroImage), so nothing has to be fetched when it is opened.
    With images off, the picture's description shows in its place and is still the link. */
function invitationEmail(guest, link) {
  link = link || guestLink(guest.code);
  return '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" bgcolor="#E9E4D9" style="background-color:#E9E4D9"><tr><td align="center" style="padding:24px 12px">' +
    '<a href="' + escapeHtml(link) + '" style="display:block;text-decoration:none">' +
    '<img src="cid:hero" width="600" alt="Ashley &amp; Charles are getting married. Bali, Indonesia, August 28, 2027. We can\'t wait to celebrate with you. Open the invitation." style="display:block;width:100%;max-width:600px;height:auto;border:0;background-color:#EEF4F5;color:#2A2C27;font-family:Georgia,serif;font-size:18px;text-align:center">' +
    '</a></td></tr></table>';
}

/** The beach artwork, fetched once from the site and embedded in each email. */
var heroBlob = null;
function heroImage() {
  if (!heroBlob) {
    var path = 'assets/email-hero-2x.jpg';
    try { heroBlob = UrlFetchApp.fetch(SITE_URL + path).getBlob(); }
    catch (e) { heroBlob = UrlFetchApp.fetch(SITE_URL.replace(/^https:/, 'http:') + path).getBlob(); }   /* before the site's certificate is issued */
    heroBlob.setName('email-hero.jpg');
  }
  return heroBlob;
}

/** One invitation email to one person. `guest` is the household (code, names); `to` is { name, email }. */
function sendInvitation(guest, to) {
  var salutation = (to.name || guest.name).split(' ')[0];
  var idx = 0; (guest.members || []).forEach(function (m, i) { if (nameKey(m) === nameKey(to.name)) idx = i + 1; });   /* the link names the person */
  var link = guestLink(guest.code, idx);
  var html = invitationEmail(guest, link);
  var text = 'Dear ' + salutation + ',\n\nAshley & Charles are getting married in Bali, Indonesia on August 28, 2027, and we can\'t wait to celebrate with you. Your invitation is here: ' +
    link + '\n\nIt opens like a letter, so give it a moment. It carries your names and your seats.\n\nPlease reply by ' + RSVP_DEADLINE + '.\n\nAshley & Charles';
  var opts = mailOptions(to.email, INVITE_SUBJECT, html, text);
  opts.inlineImages = { hero: heroImage() };
  MailApp.sendEmail(opts);
}

/* ---- Sheet helpers, run from the Invitations menu ---------------------- */

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Invitations')
    .addItem('Build households from Guest List', 'buildHouseholds')
    .addItem('Fill missing codes and links', 'fillCodes')
    .addItem('Send invitations to unsent rows', 'sendInvitations')
    .addItem('Send a test invitation to me', 'sendTestInvitation')
    .addSeparator()
    .addItem('Which version is this?', 'showVersion')
    .addToUi();
}

function showVersion() {
  SpreadsheetApp.getUi().alert('Script v' + SCRIPT_VERSION + '. Links go to column ' + LINK_COLUMN + ' of the ' + GUEST_LIST + ' tab.');
}

/** "Sam Nair <sam@x.com>, priya@x.com" -> [{ name: 'Sam Nair', email: 'sam@x.com' }, { name: '', email: 'priya@x.com' }] */
function parseRecipients(v) {
  var out = [], seen = {};
  clean(v, 1000).split(/[,;]+/).forEach(function (part) {
    var m = part.trim().match(/^(.*?)\s*<([^>]+)>$/), name = m ? m[1].trim() : '', email = (m ? m[2] : part).trim();
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(email) || seen[email.toLowerCase()]) return;
    seen[email.toLowerCase()] = true; out.push({ name: name, email: email });
  });
  return out;
}
function splitEmails(v) { return parseRecipients(v).map(function (r) { return r.email; }); }

/** Emails every household row that has a code and addresses but no value in "sent" (F): one email per person,
    each greeted by their own first name, all carrying the household's link. */
function sendInvitations() {
  var ui = SpreadsheetApp.getUi();
  if (!/^https:\/\/[^/]+\/$/.test(SITE_URL)) { ui.alert('SITE_URL at the top of the script must be the live site with a trailing slash.'); return; }
  var sheet = ss().getSheetByName(GUESTS);
  var rows = sheet.getDataRange().getValues();
  var pending = [], people = 0;
  for (var r = 1; r < rows.length; r++) {
    var code = normaliseCode(rows[r][0]);
    var to = parseRecipients(rows[r][3]);
    var sent = rows[r][5];
    if (code && to.length && !sent) { pending.push({ row: r + 1, code: code, name: clean(rows[r][1], 80), members: String(rows[r][6] || '').split('|').map(function (m) { return clean(m, 60); }).filter(Boolean), to: to }); people += to.length; }
  }
  if (!pending.length) { ui.alert('Nothing to send: every row with a code and an email is already marked as sent.'); return; }
  var quota = MailApp.getRemainingDailyQuota();
  var answer = ui.alert('Send invitations', pending.length + ' household(s), ' + people + ' email(s), have not been sent an invitation. Your account can send ' + quota + ' more emails today. Send now?', ui.ButtonSet.YES_NO);
  if (answer !== ui.Button.YES) return;
  var done = 0, stopped = false;
  for (var i = 0; i < pending.length && !stopped; i++) {
    if (MailApp.getRemainingDailyQuota() < pending[i].to.length) { stopped = true; break; }   /* a household is sent whole or not at all */
    for (var j = 0; j < pending[i].to.length; j++) sendInvitation(pending[i], pending[i].to[j]);
    sheet.getRange(pending[i].row, 6).setValue(new Date());
    done++;
  }
  ui.alert('Sent ' + done + ' of ' + pending.length + ' household(s).' + (done < pending.length ? ' The rest are still unsent; run this again tomorrow when the daily limit resets.' : ''));
}

/** Sends the invitation email for the first Guests row to the account running the script, so you can see it. */
function sendTestInvitation() {
  var sheet = ss().getSheetByName(GUESTS);
  var rows = sheet.getDataRange().getValues();
  for (var r = 1; r < rows.length; r++) {
    var code = normaliseCode(rows[r][0]);
    if (code) {
      var first = parseRecipients(rows[r][3])[0];
      sendInvitation({ code: code, name: clean(rows[r][1], 80), members: String(rows[r][6] || '').split('|').map(function (m) { return clean(m, 60); }).filter(Boolean) }, { name: first ? first.name : '', email: Session.getEffectiveUser().getEmail() });
      SpreadsheetApp.getUi().alert('Test sent to ' + Session.getEffectiveUser().getEmail() + ' using row ' + (r + 1) + '.');
      return;
    }
  }
  SpreadsheetApp.getUi().alert('No row has a code yet. Run "Fill missing codes and links" first.');
}

function ensureResponses() {
  var sheet = ss().getSheetByName(RESPONSES);
  if (sheet) {
    /* an older layout: rewrite the header if the tab is empty, otherwise keep it aside and start a fresh one */
    var head = sheet.getRange(1, 1, 1, RESPONSE_HEADERS.length).getValues()[0].map(function (v) { return String(v || ''); });
    if (head.join('|') !== RESPONSE_HEADERS.join('|')) {
      if (sheet.getLastRow() <= 1) { sheet.getRange(1, 1, 1, 20).clearContent(); sheet.getRange(1, 1, 1, RESPONSE_HEADERS.length).setValues([RESPONSE_HEADERS]); }
      else { sheet.setName(RESPONSES + ' (old)'); sheet = null; }
    }
  }
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

var HOUSEHOLD_HEADERS = ['code', 'names', 'seats', 'email', 'link', 'sent', 'members'];

function nameKey(s) { return String(s || '').toLowerCase().replace(/[^a-z ]/g, ' ').replace(/\s+/g, ' ').trim(); }

/** Reads the Guest List (one row per person) and groups people into households: anyone named in
    "Linked to another guest?" joins that person's household, as does anyone sharing an email address.
    Rows count as guests when they have a first name and a side (A, C or Both). */
function readHouseholds() {
  var src = ss().getSheetByName(GUEST_LIST);
  if (!src) throw new Error('No tab called "' + GUEST_LIST + '".');
  var rows = src.getDataRange().getValues(), h = -1, col = {}, r, c;
  for (r = 0; r < rows.length && h < 0; r++) for (c = 0; c < rows[r].length; c++) if (nameKey(rows[r][c]) === 'first name') { h = r; break; }
  if (h < 0) throw new Error('The "' + GUEST_LIST + '" tab has no "First Name" header.');
  for (c = 0; c < rows[h].length; c++) {
    var k = nameKey(rows[h][c]);
    if (k.indexOf('first') === 0) col.first = c; else if (k.indexOf('last') === 0) col.last = c; else if (k.indexOf('side') === 0) col.side = c;
    else if (k.indexOf('email') === 0) col.email = c; else if (k.indexOf('linked') === 0) col.linked = c;
    else if (k === 'room' || k === 'household') col.group = c;   /* optional: people sharing a value are one household */
  }
  for (c = 0; c < rows[h].length; c++) if (nameKey(rows[h][c]) === 'invite link') col.link = c;
  var people = [];
  for (r = h + 1; r < rows.length; r++) {
    var first = clean(rows[r][col.first], 40), last = clean(rows[r][col.last], 40), side = clean(rows[r][col.side], 10).toUpperCase();
    if (!first || !/^(A|C|BOTH)$/.test(side)) continue;
    people.push({ i: people.length, row: r + 1, first: first, last: last, full: (first + ' ' + last).trim(), emails: splitEmails(rows[r][col.email]),
      group: col.group == null ? '' : clean(rows[r][col.group], 20).toLowerCase(),
      linked: String(col.linked == null ? '' : rows[r][col.linked] || '').split(/[,;\/&]|\band\b/).map(function (x) { return x.trim(); }).filter(Boolean) });
  }
  var parent = people.map(function (_, i) { return i; });
  function find(i) { while (parent[i] !== i) { parent[i] = parent[parent[i]]; i = parent[i]; } return i; }
  function union(a, b) { a = find(a); b = find(b); if (a !== b) parent[b] = a; }
  var byFull = {}, byFirst = {}, byEmail = {}, byGroup = {};
  people.forEach(function (p) {
    if (p.group) { if (byGroup[p.group] != null) union(p.i, byGroup[p.group]); else byGroup[p.group] = p.i; }
    byFull[nameKey(p.full)] = p.i; (byFirst[nameKey(p.first)] = byFirst[nameKey(p.first)] || []).push(p.i);
    p.emails.forEach(function (e) { e = e.toLowerCase(); if (byEmail[e] != null) union(p.i, byEmail[e]); else byEmail[e] = p.i; });
  });
  people.forEach(function (p) {
    var poss = p.first.match(/^([A-Za-z]+)'s\b/);        /* "Kevin's Girlfriend", "Leo's Kimberly": joins Kevin, Leo */
    if (poss) p.linked.push(poss[1]);
    p.linked.forEach(function (name) {
      var k = nameKey(name), j = byFull[k];
      if (j == null && byFirst[k] && byFirst[k].length === 1) j = byFirst[k][0];
      if (j != null && j !== p.i) union(p.i, j);
    });
  });
  var groups = {}, order = [];
  people.forEach(function (p) { var g = find(p.i); if (!groups[g]) { groups[g] = []; order.push(g); } groups[g].push(p); });
  return order.map(function (g) {
    var m = groups[g], lasts = m.map(function (p) { return p.last; });
    var sameLast = m.length > 1 && lasts.every(function (l) { return l && l === lasts[0]; });
    var parts = m.map(function (p) { return sameLast ? p.first : p.full; });
    var names = parts.length > 1 ? parts.slice(0, -1).join(', ') + ' & ' + parts[parts.length - 1] : parts[0];
    if (sameLast) names += ' ' + lasts[0];
    var emails = [], seen = {};   /* one entry per person, "First Last <address>", so each can be written to by name */
    m.forEach(function (p) { p.emails.forEach(function (e) { if (!seen[e.toLowerCase()]) { seen[e.toLowerCase()] = true; emails.push(p.full + ' <' + e + '>'); } }); });
    return { names: names.slice(0, 80), seats: Math.min(m.length, MAX_SEATS), email: emails.join(', '), members: m.map(function (p) { return p.full; }).sort().join(' | '), rows: m.map(function (p) { return p.row; }), fulls: m.map(function (p) { return p.full; }) };
  });
}

/** Writes each person's invitation link into the Guest List, in LINK_COLUMN (column I unless changed),
    and heads that column "Invite link". Everyone in a household gets the same link. */
function writeLinks(households) {
  var src = ss().getSheetByName(GUEST_LIST), rows = src.getDataRange().getValues(), h = -1, c, r;
  for (r = 0; r < rows.length && h < 0; r++) for (c = 0; c < rows[r].length; c++) if (nameKey(rows[r][c]) === 'first name') { h = r; break; }
  var col = LINK_COLUMN - 1;
  if (!clean(rows[h][col], 40)) src.getRange(h + 1, col + 1).setValue('Invite link');   /* keeps a header you have already given the column */
  for (c = 0; c < rows[h].length; c++) {      /* links written earlier into another column are cleared, so there is one copy */
    if (c !== col && /^(invite|unique) link$/.test(nameKey(rows[h][c]))) src.getRange(h + 1, c + 1, Math.max(1, src.getLastRow() - h), 1).clearContent();
  }
  var byRow = {};
  households.forEach(function (hh) {
    var sorted = hh.members.split(' | ');
    hh.rows.forEach(function (rr, i) { var idx = sorted.indexOf(hh.fulls[i]); byRow[rr] = hh.link + (idx >= 0 ? '&p=' + (idx + 1) : ''); });
  });
  var last = src.getLastRow(), out = [];
  for (r = h + 2; r <= last; r++) out.push([byRow[r] || '']);
  if (out.length) src.getRange(h + 2, col + 1, out.length, 1).setValues(out);
}

/** Rebuilds the Households tab from the Guest List. Codes and sent dates are kept for households whose
    members have not changed; a household whose members changed gets a new row (and a new code). */
function buildHouseholds() {
  var ui = SpreadsheetApp.getUi();
  var households = readHouseholds();
  var sheet = ss().getSheetByName(GUESTS);
  if (!sheet) { sheet = ss().insertSheet(GUESTS); sheet.appendRow(HOUSEHOLD_HEADERS); sheet.setFrozenRows(1); }
  var old = sheet.getDataRange().getValues(), keep = {};
  for (var r = 1; r < old.length; r++) if (old[r][6]) keep[old[r][6]] = { code: normaliseCode(old[r][0]), sent: old[r][5] };
  var kept = 0, out = households.map(function (h) {
    var k = keep[h.members]; if (k) kept++;
    return [k ? k.code : '', h.names, h.seats, h.email, '', k ? k.sent : '', h.members];
  });
  if (old.length > 1) sheet.getRange(2, 1, old.length - 1, HOUSEHOLD_HEADERS.length).clearContent();
  sheet.getRange(1, 1, 1, HOUSEHOLD_HEADERS.length).setValues([HOUSEHOLD_HEADERS]);
  if (out.length) sheet.getRange(2, 1, out.length, HOUSEHOLD_HEADERS.length).setValues(out);
  fillCodes();
  var now = sheet.getDataRange().getValues();
  households.forEach(function (hh, i) { hh.link = SITE_URL + '?i=' + normaliseCode(now[i + 1][0]); });
  writeLinks(households);
  var dropped = Math.max(0, (old.length - 1) - kept);
  ui.alert('Script v' + SCRIPT_VERSION + ': ' + households.length + ' households from the ' + GUEST_LIST + ' tab (' + kept + ' unchanged, ' + (households.length - kept) + ' new' + (dropped ? ', ' + dropped + ' old row' + (dropped === 1 ? '' : 's') + ' removed' : '') + '). Each person\'s link is in the ' + GUEST_LIST + ' tab under "Invite link". Check names, seats and emails on the ' + GUESTS + ' tab before sending.');
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
