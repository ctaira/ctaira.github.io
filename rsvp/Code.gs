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

var SITE_URL = 'https://ashleyandcharlesinbali.com/';   // the live site, with trailing slash; used to build invitation links and to load the email artwork
var FROM_NAME = 'Ashley & Charles';       // sender name on every email (the address is the Google account running this)
var REPLY_TO = '';                        // optional: where guest replies to the emails should go
var INVITE_SUBJECT = 'Bali, 28 August 2027: your invitation';
var CONFIRM_SUBJECT = 'We have your reply';
var RSVP_DEADLINE = 'January 31, 2027';
var GUEST_LIST = 'Guest List';   // the couple's own list: one row per person, with "Linked to another guest?"
var GUESTS = 'Households';       // built from it by the Invitations menu: one row per household
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
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(email)) email = '';   /* optional: a bad address is ignored, not fatal */
    if (!attending || !name) return json({ ok: false, error: 'bad_request' });
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
      /* the copy goes to the address they typed, or else to the household's addresses on the sheet */
      sendConfirmation(guest, { name: name, email: email || splitEmails(guest.email).join(','), attending: attending, party: party, dietary: dietary, note: note });
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

/** The invitation email: the artwork on top (a link in itself), then a personal note and a button that works with images off.
    The artwork travels inside the email (see heroImage), so nothing has to be fetched from the site when it is opened. */
function invitationEmail(guest, salutation) {
  var link = guestLink(guest.code);
  var serif = "'EB Garamond',Garamond,Georgia,'Times New Roman',serif";
  /* Colours go on bgcolor attributes as well as styles: several mail clients drop background styles. */
  return '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" bgcolor="#E9E4D9" style="background-color:#E9E4D9"><tr><td align="center" style="padding:24px 12px">' +
    '<table role="presentation" width="600" cellpadding="0" cellspacing="0" border="0" bgcolor="#FBF8F1" style="max-width:600px;width:100%;background-color:#FBF8F1;font-family:' + serif + ';color:#2A2C27;line-height:1.5">' +
    '<tr><td style="padding:0"><a href="' + escapeHtml(link) + '" style="display:block;text-decoration:none">' +
    '<img src="cid:hero" width="600" alt="Ashley &amp; Charles are getting married. Bali, Indonesia, August 28, 2027. We can\'t wait to celebrate with you. Open the invitation." style="display:block;width:100%;max-width:600px;height:auto;border:0;background-color:#EEF4F5;color:#2A2C27;font-size:18px;text-align:center"></a></td></tr>' +
    '<tr><td align="center" style="padding:30px 36px 8px;text-align:center">' +
    '<p style="margin:0 0 14px;font-size:19px;color:#2A2C27">Dear ' + escapeHtml(salutation || guest.name) + ',</p>' +
    '<p style="margin:0 0 20px;font-size:17px;color:#2A2C27">Your invitation opens like a letter, so give it a moment. It carries your names and your seats.</p>' +
    '<table role="presentation" cellpadding="0" cellspacing="0" border="0" align="center" style="margin:0 auto"><tr><td bgcolor="#596044" style="background-color:#596044;border:1px solid #9B8B57;border-radius:2px">' +
    '<a href="' + escapeHtml(link) + '" style="display:inline-block;padding:14px 30px;background-color:#596044;color:#F8F3E8;text-decoration:none;font-family:' + serif + ';font-size:15px;letter-spacing:.18em;text-transform:uppercase"><font color="#F8F3E8">Open the invitation &rarr;</font></a></td></tr></table>' +
    '<p style="margin:22px 0 0;font-size:15px;color:#5B5D55">Please reply by ' + escapeHtml(RSVP_DEADLINE) + '.</p>' +
    '<p style="margin:16px 0 0;font-size:13px;color:#7A7C74">Or copy this link: <a href="' + escapeHtml(link) + '" style="color:#3F5B45">' + escapeHtml(link) + '</a></p>' +
    '<p style="margin:26px 0 22px;font-family:Menlo,Consolas,monospace;font-size:11px;letter-spacing:.18em;text-transform:uppercase;color:#9B8B57">Ashley &amp; Charles &middot; AYANA Bali &middot; August 28, 2027</p>' +
    '</td></tr></table></td></tr></table>';
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
  var html = invitationEmail(guest, salutation);
  var text = 'Dear ' + salutation + ',\n\nAshley & Charles are getting married in Bali, Indonesia on August 28, 2027, and we can\'t wait to celebrate with you. Your invitation is here: ' +
    guestLink(guest.code) + '\n\nIt opens like a letter, so give it a moment. It carries your names and your seats.\n\nPlease reply by ' + RSVP_DEADLINE + '.\n\nAshley & Charles';
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
    .addToUi();
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
    if (code && to.length && !sent) { pending.push({ row: r + 1, code: code, name: clean(rows[r][1], 80), to: to }); people += to.length; }
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
      sendInvitation({ code: code, name: clean(rows[r][1], 80) }, { name: first ? first.name : '', email: Session.getEffectiveUser().getEmail() });
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
  var people = [];
  for (r = h + 1; r < rows.length; r++) {
    var first = clean(rows[r][col.first], 40), last = clean(rows[r][col.last], 40), side = clean(rows[r][col.side], 10).toUpperCase();
    if (!first || !/^(A|C|BOTH)$/.test(side)) continue;
    people.push({ i: people.length, first: first, last: last, full: (first + ' ' + last).trim(), emails: splitEmails(rows[r][col.email]),
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
    return { names: names.slice(0, 80), seats: Math.min(m.length, MAX_SEATS), email: emails.join(', '), members: m.map(function (p) { return p.full; }).sort().join(' | ') };
  });
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
  var dropped = Math.max(0, (old.length - 1) - kept);
  ui.alert(households.length + ' households from the ' + GUEST_LIST + ' tab (' + kept + ' unchanged, ' + (households.length - kept) + ' new' + (dropped ? ', ' + dropped + ' old row' + (dropped === 1 ? '' : 's') + ' removed' : '') + '). Check names, seats and emails on the ' + GUESTS + ' tab before sending.');
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
