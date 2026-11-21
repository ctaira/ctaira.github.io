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
var CC_INVITES = 'ashley.plus.charles@gmail.com';   // every invitation is copied here, unless this account is the one sending
var CONFIRM_FROM = '';                    // RSVP confirmations go out from the account that owns the deployment (the wedding account). Set a verified "Send mail as" alias here to send from it instead.
var CONFIRM_CC = 'ashleydong97@gmail.com, charles.h.taira@gmail.com';   // every confirmation is copied here, so you both see each reply in your own inboxes
var REPLY_TO = '';                        // empty: replies go to whoever sent. Invitations sent from the menu go out as the person who clicked it, so each of you can send to your own side from your own account.
var INVITE_SUBJECT = 'Ashley & Charles are Getting Married | Bali 2027';
var CONFIRM_SUBJECT_COMING = 'See you in Bali';
var CONFIRM_SUBJECT_NOT = "We'll miss you in Bali";
var RSVP_DEADLINE = 'January 31, 2027';
var REMIND_SUBJECT = 'A little reminder to RSVP | Ashley & Charles, Bali 2027';
var REMIND_CC = 'charles.h.taira@gmail.com, ashleydong97@gmail.com, ashley.plus.charles@gmail.com';   // every reminder is copied here, except to whichever of these is sending it (its Sent folder already has it)
var REMIND_SIDES = { 'charles.h.taira@gmail.com': 'C', 'ashleydong97@gmail.com': 'A' };   // run from one of these accounts, "everyone who hasn't replied" means that side only (plus anyone marked Both); from any other account, both sides
var REMIND_GAP_DAYS = 7;         // "everyone who hasn't replied" skips anyone reminded in the last week, so a second round later on is possible but a double click is not
var GUEST_LIST = 'Guest List';   // the couple's own list: one row per person, with "Linked to another guest?"
var GUESTS = 'Households';       // built from it by the Invitations menu: one row per household
var LINK_COLUMN = 9;             // Guest List column that receives each person's invite link (9 = I)
var SCRIPT_VERSION = 24;          // shown in the Invitations menu's messages, so you can tell which copy is running
var RESPONSES = 'Responses';
var MAX_SEATS = 6;
var CODE_ALPHABET = 'abcdefghjkmnpqrstuvwxyz23456789'; // no 0/o/1/l, so codes survive being read aloud
var CODE_LENGTH = 6;
var RESPONSE_HEADERS = ['Received', 'Code', 'Household', 'Replied by', 'Email', 'Person', 'Answer', 'Dietary', 'Note', 'Sent'];   // one row per person per reply
var ATTENDING_TEXT = { both: 'Both days', saturday: 'Saturday only', no: 'Not coming' };
var ATTENDING_KEY = { 'Both days': 'both', 'Saturday only': 'saturday', 'Not coming': 'no' };
var ANSWER_LABEL = { both: 'Both days', saturday: 'Saturday only', no: 'Unable to attend' };   /* what guests read in their email; the sheet keeps ATTENDING_TEXT */

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

/** The latest recorded answer for each member of a household, as { name: 'both' | 'saturday' | 'no' }.
    Later Responses rows win, so a changed mind replaces the earlier answer. */
function latestAnswers(code) {
  var out = {}, sheet = ss().getSheetByName(RESPONSES);
  if (!sheet) return out;
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (normaliseCode(rows[i][1]) !== code) continue;
    var key = ATTENDING_KEY[String(rows[i][6] || '')];
    if (key) out[clean(rows[i][5], 60)] = key;
  }
  return out;
}

/* ---- Web app ------------------------------------------------------------ */

function doGet(e) {
  var code = e && e.parameter ? e.parameter.i : '';
  if (!code) return json({ ok: true, service: 'rsvp' });
  var guest = findGuest(code);
  if (!guest) return json({ ok: false, error: 'not_found' });
  try { stampOpen(guest.code, parseInt(e.parameter.p, 10) || 0); } catch (err) { Logger.log('open stamp failed: ' + err); }   /* never in the way of the lookup */
  return json({ ok: true, name: guest.name, seats: guest.seats, members: guest.members, answers: latestAnswers(guest.code) });
}

/** Each time an invitation page loads, the Guest List row it was sent to gets a "First opened" date and time
    (kept from the first visit), a "Last opened" one (the most recent visit) and its "Opens" count goes up by
    one. The row is found by its link in the "Invite link" column: the code picks the household, the link's
    p= picks the person. A link without p= counts for the whole household. The three columns are added at the
    end of the Guest List the first time they are needed. Your own test opens count too, so read it as a
    nudge list rather than a record. */
function stampOpen(code, person) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return;          /* two people opening at once: the second stamp is skipped rather than miscounted */
  try {
    var g = guestListColumns(false); if (!g || g.col.link == null) return;
    var c, col = {};
    for (c = 0; c < g.rows[g.header].length; c++) { var k = nameKey(g.rows[g.header][c]); if (k === 'first opened') col.first = c; else if (k === 'last opened') col.last = c; else if (k === 'opens') col.count = c; }
    var next = g.sheet.getLastColumn();
    if (col.first == null) { col.first = next++; g.sheet.getRange(g.header + 1, col.first + 1).setValue('First opened'); }
    if (col.last == null) { col.last = next++; g.sheet.getRange(g.header + 1, col.last + 1).setValue('Last opened'); }
    if (col.count == null) { col.count = next++; g.sheet.getRange(g.header + 1, col.count + 1).setValue('Opens'); }
    var now = new Date(), tail = '?i=' + code, when = 'M/d/yyyy h:mm am/pm';
    for (var r = g.header + 1; r < g.rows.length; r++) {
      var link = String(g.rows[r][g.col.link] || ''), at = link.indexOf(tail);
      if (at < 0) continue;
      var rest = link.slice(at + tail.length), m = rest.match(/^&p=(\d+)/);
      if (rest && !m) continue;                                        /* another code that merely starts with this one */
      if (person && (!m || parseInt(m[1], 10) !== person)) continue;
      var first = col.first < g.rows[r].length ? g.rows[r][col.first] : '', n = col.count < g.rows[r].length ? parseInt(g.rows[r][col.count], 10) : 0;
      if (!first) g.sheet.getRange(r + 1, col.first + 1).setValue(now).setNumberFormat(when);
      g.sheet.getRange(r + 1, col.last + 1).setValue(now).setNumberFormat(when);
      g.sheet.getRange(r + 1, col.count + 1).setValue((isNaN(n) ? 0 : n) + 1);
    }
  } finally { lock.releaseLock(); }
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
    '<p style="margin:0 0 6px;font-family:Menlo,Consolas,monospace;font-size:11px;letter-spacing:.18em;text-transform:uppercase;color:#8C6E4F">Ashley &amp; Charles &middot; AYANA Bali &middot; August 28, 2027</p>' +
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
  /* the link leads back to the person who replied, so the page opens on their own row rather than the household view */
  var idx = 0; (guest.members || []).forEach(function (m, i) { if (nameKey(m) === nameKey(reply.name)) idx = i + 1; });
  var back = guestLink(guest.code, idx);
  /* the whole household's standing, read back from the sheet: what this reply changed and what was already in */
  var current = latestAnswers(guest.code), all = guest.members.length ? guest.members : [guest.name];
  var lines = all.filter(function (m) { return current[m]; }).map(function (m) { return [m, ANSWER_LABEL[current[m]]]; });
  var unanswered = all.filter(function (m) { return !current[m]; });
  if (reply.dietary) lines.push(['Dietary', reply.dietary]);
  if (reply.note) lines.push(['Note', reply.note]);
  var rows = lines.map(function (l) {
    return '<tr><td style="padding:6px 16px 6px 0;font-family:Menlo,Consolas,monospace;font-size:11px;letter-spacing:.14em;text-transform:uppercase;color:#5B5D55;vertical-align:top">' + escapeHtml(l[0]) + '</td><td style="padding:6px 0;vertical-align:top">' + escapeHtml(l[1]) + '</td></tr>';
  }).join('');
  /* the greeting follows the person who replied: their own answer if they are in the household, else what this reply said */
  var mine = ''; all.forEach(function (m) { if (nameKey(m) === nameKey(reply.name)) mine = current[m] || ''; });
  var coming = mine ? mine !== 'no' : reply.coming;
  var lead = coming
    ? "Thank you. Your reply is in, and we'll be in touch with booking details for the estate."
    : "Thank you for letting us know. We'll miss you, and we'll raise a glass to you from the cliff.";
  var later = unanswered.length ? '<p style="margin:14px 0 0;font-size:14px;color:#5B5D55">Still to reply: ' + escapeHtml(unanswered.join(', ')) + '. They can answer any time from their own invitation.</p>' : '';
  var html = emailShell(
    coming ? "We can't wait" : "We'll miss you",
    '<p style="margin:0 0 18px">' + escapeHtml(lead) + '</p>' +
    '<table style="border-collapse:collapse;margin:0 0 6px">' + rows + '</table>' + later +
    '<p style="margin:22px 0 0;font-size:14px;color:#5B5D55">Need to change something? Open your invitation again and send a new reply any time before ' + escapeHtml(RSVP_DEADLINE) + '. The latest answer for each person counts.</p>',
    'Open your invitation', back
  );
  var text = lead + '\n\n' + lines.map(function (l) { return l[0] + ': ' + l[1]; }).join('\n') +
    (unanswered.length ? '\n\nStill to reply: ' + unanswered.join(', ') + '.' : '') +
    '\n\nTo change your reply before ' + RSVP_DEADLINE + ', open your invitation again: ' + back;
  var opts = mailOptions(reply.email, coming ? CONFIRM_SUBJECT_COMING : CONFIRM_SUBJECT_NOT, html, text);
  if (CONFIRM_CC) opts.cc = CONFIRM_CC;
  sendAs(opts, CONFIRM_FROM);
}

/** Sends from `from` when it is one of this account's verified "Send mail as" aliases (through the Gmail service,
    which can set the From address); otherwise plainly, from the account itself. */
function sendAs(opts, from) {
  var ok = false;
  if (from) { try { ok = GmailApp.getAliases().some(function (a) { return String(a).toLowerCase() === from.toLowerCase(); }); } catch (e) { ok = false; } }
  if (!ok) { MailApp.sendEmail(opts); return; }
  var o = { htmlBody: opts.htmlBody, name: opts.name, from: from };
  if (opts.cc) o.cc = opts.cc; if (opts.replyTo) o.replyTo = opts.replyTo; if (opts.inlineImages) o.inlineImages = opts.inlineImages;
  GmailApp.sendEmail(opts.to, opts.subject, opts.body, o);
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
function invitationEmail(guest, link, salutation) {
  link = link || guestLink(guest.code);
  var font = 'font-family:-apple-system,Segoe UI,Helvetica,Arial,sans-serif;';
  var cal = 'https://calendar.google.com/calendar/render?action=TEMPLATE&text=Ashley+%26+Charles+%C2%B7+Wedding+weekend+in+Bali&dates=20270827T050000Z%2F20270828T150000Z&ctz=Asia%2FMakassar&location=AYANA+Bali%2C+Jimbaran%2C+Bali%2C+Indonesia&details=Friday+Aug+27%2C+1%3A00+PM%3A+tea+ceremony+at+AYANA+Rimba+(optional)%2C+then+the+welcome+party+at+Rock+Bar+from+4%3A30+PM.%0ASaturday+Aug+28%2C+4%3A30+PM%3A+ceremony+at+AYANA+SKY%2C+then+dinner+and+dancing.%0Ahttps%3A%2F%2Fashleyandcharlesinbali.com';
  var maps = 'https://www.google.com/maps/search/?api=1&query=AYANA+Resort+Bali%2C+Jimbaran';
  var label = font + 'font-size:12px;letter-spacing:.14em;text-transform:uppercase;color:#3F5B45;text-decoration:none;font-weight:600;white-space:nowrap';
  var row = function (title, lines, action, url) {
    return '<tr><td style="padding:16px 0;border-top:1px solid #E3DED2;' + font + 'font-size:15px;line-height:1.5;color:#2A2C27;vertical-align:top">' + title + '<br><span style="color:#5B5D55">' + lines + '</span></td>' +
      '<td align="right" style="padding:16px 0 16px 16px;border-top:1px solid #E3DED2;vertical-align:middle"><a href="' + url + '" style="' + label + '">' + action + '</a></td></tr>';
  };
  /* Mostly text, with the artwork below it: the greeting and the ask come first, the way an invitation card service lays it out */
  return '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" bgcolor="#E9E4D9" style="background-color:#E9E4D9"><tr><td align="center" style="padding:28px 12px">' +
    '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="max-width:600px;background:#F7F3EA;border-radius:6px"><tr><td style="padding:36px 32px 28px">' +
      '<p style="margin:0 0 6px;' + font + 'font-family:Menlo,Consolas,monospace;font-size:11px;letter-spacing:.18em;text-transform:uppercase;color:#8C6E4F">Ashley &amp; Charles &middot; AYANA Bali &middot; August 28, 2027</p>' +
      '<p style="margin:0 0 18px;' + font + 'font-size:17px;line-height:1.55;color:#2A2C27">Dear ' + escapeHtml(salutation || 'friends') + ',</p>' +
      '<p style="margin:0 0 14px;' + font + 'font-size:17px;line-height:1.55;color:#2A2C27">We are thrilled to invite you to celebrate with us in Bali. Within your invitation below, you&rsquo;ll find details for the weekend, along with helpful information for planning your time in Bali.</p>' +
      '<p style="margin:0 0 24px;' + font + 'font-size:17px;line-height:1.55;color:#2A2C27">We look forward to celebrating together and kindly ask that you RSVP by ' + escapeHtml(RSVP_DEADLINE) + '.</p>' +
      '<p style="margin:0 0 30px;text-align:center"><a href="' + escapeHtml(link) + '" style="display:inline-block;background:#242722;color:#F2EDE3;text-decoration:none;padding:15px 30px;border-radius:999px;' + font + 'font-size:13px;letter-spacing:.16em;text-transform:uppercase;font-weight:600">Open the invitation</a></p>' +
      '<a href="' + escapeHtml(link) + '" style="display:block;text-decoration:none"><img src="cid:hero" width="536" alt="Ashley &amp; Charles are getting married. Bali, Indonesia, August 28, 2027. We can\'t wait to celebrate with you." style="display:block;width:100%;max-width:536px;height:auto;border:0;border-radius:4px"></a>' +
      '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="margin-top:28px">' +
        row('Friday, August 27 &amp; Saturday, August 28, 2027', 'Tea ceremony and welcome party, then the wedding on the cliff', 'Add to calendar', cal) +
        row('AYANA Bali', 'Jimbaran, Bali, Indonesia', 'View location', maps) +
      '</table>' +
      '<p style="margin:24px 0 0;' + font + 'font-size:13px;line-height:1.5;color:#7A7C74">Or copy this link: <a href="' + escapeHtml(link) + '" style="color:#3F5B45">' + escapeHtml(link) + '</a></p>' +
      '<p style="margin:14px 0 0;' + font + 'font-size:13px;line-height:1.5;color:#7A7C74">Questions? Just reply to this email.</p>' +
    '</td></tr></table>' +
    '</td></tr></table>';
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
  var html = invitationEmail(guest, link, salutation);
  var text = 'Dear ' + salutation + ',\n\nWe are thrilled to invite you to celebrate with us in Bali. Within your invitation, you\u2019ll find details for the weekend, along with helpful information for planning your time in Bali:\n' +
    link + '\n\nWe look forward to celebrating together and kindly ask that you RSVP by ' + RSVP_DEADLINE + '.\n\nFriday, August 27 & Saturday, August 28, 2027\nAYANA Bali, Jimbaran, Bali, Indonesia\n\nQuestions? Just reply to this email.\n\nAshley & Charles';
  var opts = mailOptions(to.email, INVITE_SUBJECT, html, text);
  var me = ''; try { me = Session.getEffectiveUser().getEmail(); } catch (e) {}
  if (CC_INVITES && me.toLowerCase() !== CC_INVITES.toLowerCase()) opts.cc = CC_INVITES;
  opts.inlineImages = { hero: heroImage() };
  MailApp.sendEmail(opts);
}

/* ---- Sheet helpers, run from the Invitations menu ---------------------- */

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Invitations')
    .addItem('Build households from Guest List', 'buildHouseholds')
    .addItem('Fill missing codes and links', 'fillCodes')
    .addItem('Send to the selected guests (Guest List rows)', 'sendSelectedGuests')
    .addItem('Send the next unsent household', 'sendNextHousehold')
    .addItem('Send the selected household', 'sendSelectedHousehold')
    .addItem('Send invitations to unsent rows', 'sendInvitations')
    .addItem('Send a test invitation to me', 'sendTestInvitation')
    .addSeparator()
    .addItem('Who hasn\'t replied yet?', 'showUnreplied')
    .addItem('Send a test reminder to me', 'sendTestReminder')
    .addItem('Send reminders to the selected guests (Guest List rows)', 'sendSelectedReminders')
    .addItem('Send reminders to everyone who hasn\'t replied', 'sendAllReminders')
    .addSeparator()
    .addItem('How many emails can I still send today?', 'showQuota')
    .addItem('Which version is this?', 'showVersion')
    .addToUi();
}

/** Gmail's sending allowance for the account running the menu: a rolling 24-hour window, counted per recipient. */
function showQuota() {
  var left = MailApp.getRemainingDailyQuota(), me = '';
  try { me = Session.getEffectiveUser().getEmail(); } catch (e) {}
  var perInvite = CC_INVITES && me.toLowerCase() !== CC_INVITES.toLowerCase() ? 2 : 1;
  SpreadsheetApp.getUi().alert((me ? me + ' can' : 'You can') + ' send ' + left + ' more email(s) right now. Each invitation counts ' + perInvite + (perInvite === 2 ? ' (the guest plus the CC)' : '') + ', so that is about ' + Math.floor(left / perInvite) + ' invitation(s). The allowance comes back on a rolling 24-hour basis, roughly a day after each email went out.');
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

/** The households still to send, read from the Households tab: one entry per row with a code, addresses and no sent date. */
function pendingHouseholds(sheet, rows) {
  var pending = [];
  for (var r = 1; r < rows.length; r++) {
    var code = normaliseCode(rows[r][0]), to = parseRecipients(rows[r][3]), sent = rows[r][5];
    if (code && to.length && !sent) pending.push({ row: r + 1, code: code, name: clean(rows[r][1], 80), members: String(rows[r][6] || '').split('|').map(function (m) { return clean(m, 60); }).filter(Boolean), to: to });
  }
  return pending;
}
/** The Guest List's columns, found by header, plus the "Invite sent" column (made at the end if missing). */
function guestListColumns(makeSent) {
  var src = ss().getSheetByName(GUEST_LIST), rows = src.getDataRange().getValues(), h = -1, c, col = {};
  for (var r = 0; r < rows.length && h < 0; r++) for (c = 0; c < rows[r].length; c++) if (nameKey(rows[r][c]) === 'first name') { h = r; break; }
  if (h < 0) return null;
  for (c = 0; c < rows[h].length; c++) {
    var k = nameKey(rows[h][c]);
    if (k === 'first name') col.first = c; else if (k === 'last name') col.last = c; else if (k.indexOf('email') === 0) col.email = c;
    else if (/^(invite|unique) link$/.test(k)) col.link = c; else if (k === 'invite sent') col.sent = c;
  }
  if (col.sent == null && makeSent) { col.sent = src.getLastColumn(); src.getRange(h + 1, col.sent + 1).setValue('Invite sent'); rows = src.getDataRange().getValues(); }
  return { sheet: src, rows: rows, header: h, col: col };
}
/** Addresses already stamped "Invite sent" on the Guest List, lower-cased. */
function stampedEmails() {
  var g = guestListColumns(false), out = {};
  if (!g || g.col.sent == null || g.col.email == null) return out;
  for (var r = g.header + 1; r < g.rows.length; r++) if (g.rows[r][g.col.sent]) parseRecipients(g.rows[r][g.col.email]).forEach(function (x) { out[x.email.toLowerCase()] = true; });
  return out;
}
/** Stamps the Guest List rows whose address is among these, so the per-person and per-household sends agree. */
function stampGuests(emails) {
  var g = guestListColumns(true); if (!g || g.col.email == null) return;
  var want = {}; emails.forEach(function (e) { want[e.toLowerCase()] = true; });
  for (var r = g.header + 1; r < g.rows.length; r++) {
    if (g.rows[r][g.col.sent]) continue;
    var hit = parseRecipients(g.rows[r][g.col.email]).some(function (x) { return want[x.email.toLowerCase()]; });
    if (hit) g.sheet.getRange(r + 1, g.col.sent + 1).setValue(new Date());
  }
}
/** One household, whole: every recipient not already sent to is emailed, then the row's sent date written. Returns the number of emails. */
function sendHousehold(sheet, h) {
  var done = stampedEmails(), sent = [];
  for (var j = 0; j < h.to.length; j++) { if (done[h.to[j].email.toLowerCase()]) continue; sendInvitation(h, h.to[j]); sent.push(h.to[j].email); }
  sheet.getRange(h.row, 6).setValue(new Date());
  if (sent.length) stampGuests(sent);
  return sent.length;
}

/** Hand-picked: the people on the rows you have selected on the Guest List tab, one email each, to their own
    address with their own link. Marks each in an "Invite sent" column; a household is marked sent once everyone
    in it with an address has had theirs. */
function sendSelectedGuests() {
  var ui = SpreadsheetApp.getUi(), active = ss().getActiveSheet();
  if (!active || active.getName() !== GUEST_LIST) { ui.alert('Select one or more rows on the ' + GUEST_LIST + ' tab first.'); return; }
  var g = guestListColumns(true); if (!g) { ui.alert('The ' + GUEST_LIST + ' tab has no "First Name" header.'); return; }
  if (g.col.link == null) { ui.alert('The ' + GUEST_LIST + ' tab has no "Invite link" column. Run "Build households from Guest List" first.'); return; }
  var range = active.getActiveRange(), r0 = range.getRow(), n = range.getNumRows();
  var picks = [], skipped = [];
  for (var r = r0; r < r0 + n; r++) {
    if (r <= g.header + 1 || r > g.rows.length) continue;
    var row = g.rows[r - 1], name = clean((row[g.col.first] || '') + ' ' + (row[g.col.last] || ''), 80);
    if (!name.trim()) continue;
    var to = parseRecipients(row[g.col.email])[0];
    var m = String(row[g.col.link] || '').match(/[?&]i=([a-z0-9]{6})/);
    if (!to) { skipped.push(name + ' (no email)'); continue; }
    if (!m) { skipped.push(name + ' (no invite link)'); continue; }
    var guest = findGuest(m[1]);
    if (!guest) { skipped.push(name + ' (household not found)'); continue; }
    picks.push({ row: r, name: name, email: to.email, guest: guest, already: !!row[g.col.sent] });
  }
  if (!picks.length) { ui.alert('Nothing to send on the selected row(s).' + (skipped.length ? '\n\nSkipped: ' + skipped.join(', ') : '')); return; }
  var lines = picks.map(function (x) { return x.name + ' <' + x.email + '>' + (x.already ? '  (already sent: this would be a second email)' : ''); });
  var answer = ui.alert('Send to ' + picks.length + ' guest(s)?', lines.join('\n') + (skipped.length ? '\n\nSkipped: ' + skipped.join(', ') : ''), ui.ButtonSet.YES_NO);
  if (answer !== ui.Button.YES) return;
  if (MailApp.getRemainingDailyQuota() < picks.length) { ui.alert('Not enough of today\'s email limit left for ' + picks.length + ' email(s).'); return; }
  var sent = 0;
  picks.forEach(function (x) {
    sendInvitation(x.guest, { name: x.name, email: x.email });
    g.sheet.getRange(x.row, g.col.sent + 1).setValue(new Date()); sent++;
  });
  settleHouseholds(picks.map(function (x) { return x.guest.code; }));
  ui.alert('Sent ' + sent + ' email(s).');
}
/** A household whose every addressed member is stamped on the Guest List is marked sent on the Households tab. */
function settleHouseholds(codes) {
  var sheet = ss().getSheetByName(GUESTS), rows = sheet.getDataRange().getValues(), done = stampedEmails(), seen = {};
  codes.forEach(function (code) {
    if (seen[code]) return; seen[code] = true;
    for (var r = 1; r < rows.length; r++) {
      if (normaliseCode(rows[r][0]) !== code || rows[r][5]) continue;
      var to = parseRecipients(rows[r][3]);
      if (to.length && to.every(function (t) { return done[t.email.toLowerCase()]; })) sheet.getRange(r + 1, 6).setValue(new Date());
    }
  });
}

/** One at a time, in sheet order: the first household not yet marked sent. Check Sent in Gmail, then run again. */
function sendNextHousehold() {
  var ui = SpreadsheetApp.getUi(), sheet = ss().getSheetByName(GUESTS);
  var pending = pendingHouseholds(sheet, sheet.getDataRange().getValues());
  if (!pending.length) { ui.alert('Nothing left to send: every household with an email is marked as sent.'); return; }
  var h = pending[0];
  var answer = ui.alert('Send to ' + h.name + '?', h.to.map(function (t) { return t.name + ' <' + t.email + '>'; }).join('\n') + '\n\n' + (pending.length - 1) + ' household(s) will remain after this one.', ui.ButtonSet.YES_NO);
  if (answer !== ui.Button.YES) return;
  if (MailApp.getRemainingDailyQuota() < h.to.length) { ui.alert('Not enough of today\'s email limit left for this household (' + h.to.length + ' emails). Try again tomorrow.'); return; }
  var n = sendHousehold(sheet, h);
  ui.alert((n ? 'Sent ' + n + ' email(s) to ' + h.name + '.' : 'Everyone in ' + h.name + ' with an address had already been sent an invitation, so no email went out; the household is now marked sent.') + ' ' + (pending.length - 1) + ' household(s) still to send.');
}

/** The household on the row you have selected in the Households tab, whether or not it is marked sent. */
function sendSelectedHousehold() {
  var ui = SpreadsheetApp.getUi(), sheet = ss().getSheetByName(GUESTS), active = ss().getActiveSheet();
  if (!active || active.getName() !== GUESTS) { ui.alert('Select a row on the ' + GUESTS + ' tab first.'); return; }
  var r = active.getActiveRange().getRow();
  if (r < 2) { ui.alert('Select a household row, not the header.'); return; }
  var rows = sheet.getDataRange().getValues(), row = rows[r - 1];
  var code = normaliseCode(row[0]), to = parseRecipients(row[3]);
  if (!code) { ui.alert('That row has no code yet. Run "Fill missing codes and links" first.'); return; }
  if (!to.length) { ui.alert('That row has no email address to send to.'); return; }
  var h = { row: r, code: code, name: clean(row[1], 80), members: String(row[6] || '').split('|').map(function (m) { return clean(m, 60); }).filter(Boolean), to: to };
  var warn = row[5] ? '\n\nThis household is already marked as sent on ' + row[5] + '. Sending again will email them a second time.' : '';
  var answer = ui.alert('Send to ' + h.name + '?', h.to.map(function (t) { return t.name + ' <' + t.email + '>'; }).join('\n') + warn, ui.ButtonSet.YES_NO);
  if (answer !== ui.Button.YES) return;
  if (MailApp.getRemainingDailyQuota() < h.to.length) { ui.alert('Not enough of today\'s email limit left for this household.'); return; }
  var n = sendHousehold(sheet, h);
  ui.alert(n ? 'Sent ' + n + ' email(s) to ' + h.name + '.' : 'Everyone in ' + h.name + ' with an address had already been sent an invitation (see the Guest List\'s "Invite sent" column), so no email went out; the household is now marked sent.');
}

/** Emails every household row that has a code and addresses but no value in "sent" (F): one email per person,
    each greeted by their own first name, all carrying the household's link. */
function sendInvitations() {
  var ui = SpreadsheetApp.getUi();
  if (!/^https:\/\/[^/]+\/$/.test(SITE_URL)) { ui.alert('SITE_URL at the top of the script must be the live site with a trailing slash.'); return; }
  var sheet = ss().getSheetByName(GUESTS);
  var pending = pendingHouseholds(sheet, sheet.getDataRange().getValues()), people = 0;
  pending.forEach(function (h) { people += h.to.length; });
  if (!pending.length) { ui.alert('Nothing to send: every row with a code and an email is already marked as sent.'); return; }
  var quota = MailApp.getRemainingDailyQuota();
  var answer = ui.alert('Send invitations', pending.length + ' household(s), ' + people + ' email(s), have not been sent an invitation. Your account can send ' + quota + ' more emails today. Send now?', ui.ButtonSet.YES_NO);
  if (answer !== ui.Button.YES) return;
  var done = 0, stopped = false;
  for (var i = 0; i < pending.length && !stopped; i++) {
    if (MailApp.getRemainingDailyQuota() < pending[i].to.length) { stopped = true; break; }   /* a household is sent whole or not at all */
    sendHousehold(sheet, pending[i]);
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
    if (k === 'first name') col.first = c; else if (k === 'last name') col.last = c; else if (k.indexOf('side') === 0) col.side = c;   /* exact: 'First opened' / 'Last opened' are not name columns */
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

/* ---- Reminders ------------------------------------------------------------ */

/** The reminder artwork (Tuco & Elsa on the beach), fetched once from the site and embedded in each email. */
var reminderBlob = null;
function reminderImage() {
  if (!reminderBlob) { reminderBlob = UrlFetchApp.fetch(SITE_URL + 'assets/email-reminder.jpg').getBlob(); reminderBlob.setName('reminder.jpg'); }
  return reminderBlob;
}

/** Everyone on the Guest List who was sent an invitation, has an email address, and has no answer recorded for
    themselves yet. Answers come from the Responses tab (the latest per person), so a partner who answered for
    them counts. Anyone with "Skip" in the "Reminder sent" column is left out. People sharing one address are folded into one email. Each entry:
    { email, names: [full names], firsts, code, link, rows: [sheet rows], side, reminded: Date|'' } */
function unrepliedGuests() {
  var g = guestListColumns(false); if (!g || g.col.link == null) return [];
  var h = g.rows[g.header], c, col = { side: null, reminded: null };
  for (c = 0; c < h.length; c++) { var k = nameKey(h[c]); if (k.indexOf('side') === 0) col.side = c; else if (k === 'reminder sent') col.reminded = c; }
  var answers = {}, rs = ss().getSheetByName(RESPONSES);
  if (rs) rs.getDataRange().getValues().slice(1).forEach(function (r) { if (ATTENDING_KEY[String(r[6] || '')]) answers[normaliseCode(r[1]) + '|' + nameKey(r[5])] = true; });
  var houses = {}, byKey = {}, out = [];
  for (var r = g.header + 1; r < g.rows.length; r++) {
    var row = g.rows[r];
    var name = clean((row[g.col.first] || '') + ' ' + (row[g.col.last] || ''), 80);
    var to = parseRecipients(row[g.col.email])[0];
    var m = String(row[g.col.link] || '').match(/[?&]i=([a-z0-9]{6})(?:&p=(\d+))?/);
    if (!name || !to || !m || (g.col.sent != null && !row[g.col.sent])) continue;   /* no address, no link, or never invited */
    var code = m[1], idx = parseInt(m[2], 10) || 0;
    if (!houses[code]) houses[code] = findGuest(code);
    var guest = houses[code]; if (!guest) continue;
    var member = guest.members[idx - 1] || name;
    if (answers[code + '|' + nameKey(member)]) continue;
    var key = to.email.toLowerCase() + '|' + code, reminded = col.reminded != null ? row[col.reminded] : '';
    if (/^\s*skip/i.test(String(reminded))) continue;   /* "Skip" in the Reminder sent column: never remind this person */
    if (!byKey[key]) { byKey[key] = { email: to.email, names: [], firsts: [], code: code, idxs: [], rows: [], side: col.side != null ? clean(row[col.side], 10).toUpperCase() : '', reminded: reminded }; out.push(byKey[key]); }
    var e = byKey[key]; e.names.push(member); e.firsts.push(clean(row[g.col.first], 40)); e.idxs.push(idx); e.rows.push(r + 1);
    if (reminded && (!e.reminded || reminded > e.reminded)) e.reminded = reminded;
  }
  out.forEach(function (e) { e.link = guestLink(e.code, e.idxs.length === 1 ? e.idxs[0] : 0); });   /* a shared address gets the household link, which asks who is replying */
  return out;
}

function joinNames(list) { return list.length < 2 ? (list[0] || '') : list.slice(0, -1).join(', ') + ' and ' + list[list.length - 1]; }

function reminderEmail(salutation, link) {
  var font = 'font-family:-apple-system,Segoe UI,Helvetica,Arial,sans-serif;';
  var p = 'margin:0 0 14px;' + font + 'font-size:17px;line-height:1.55;color:#2A2C27';
  return '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" bgcolor="#E9E4D9" style="background-color:#E9E4D9"><tr><td align="center" style="padding:28px 12px">' +
    '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="max-width:600px;background:#F7F3EA;border-radius:6px"><tr><td style="padding:36px 32px 28px">' +
      '<p style="margin:0 0 6px;font-family:Menlo,Consolas,monospace;font-size:11px;letter-spacing:.18em;text-transform:uppercase;color:#8C6E4F">Ashley &amp; Charles &middot; AYANA Bali &middot; August 28, 2027</p>' +
      '<p style="' + p + ';margin-bottom:18px">Dear ' + escapeHtml(salutation || 'friends') + ',</p>' +
      '<p style="' + p + '">A gentle nudge from us, and from Tuco &amp; Elsa: we would love to know if you can join us in Bali.</p>' +
      '<p style="' + p + ';margin-bottom:24px">Please RSVP by ' + escapeHtml(RSVP_DEADLINE) + '. It only takes a minute, and it helps us plan rooms and the weekend.</p>' +
      '<a href="' + escapeHtml(link) + '" style="display:block;text-decoration:none"><img src="cid:reminder" width="536" alt="We\'re pawsitively awaiting your reply. A little reminder from Tuco &amp; Elsa. Please RSVP by ' + escapeHtml(RSVP_DEADLINE) + '. Open the invitation." style="display:block;width:100%;max-width:536px;height:auto;border:0;border-radius:4px"></a>' +
      '<p style="margin:24px 0 0;' + font + 'font-size:13px;line-height:1.5;color:#7A7C74">Or copy this link: <a href="' + escapeHtml(link) + '" style="color:#3F5B45">' + escapeHtml(link) + '</a></p>' +
      '<p style="margin:14px 0 0;' + font + 'font-size:13px;line-height:1.5;color:#7A7C74">Already replied? Thank you, and please ignore this note. Questions? Just reply to this email.</p>' +
    '</td></tr></table>' +
    '</td></tr></table>';
}

/** One reminder to one address. `e` is an entry from unrepliedGuests(); `to` overrides the address (for a test). */
function sendReminder(e, to, subjectPrefix) {
  var salutation = joinNames(e.firsts);
  var text = 'Dear ' + salutation + ',\n\nA gentle nudge from us, and from Tuco & Elsa: we would love to know if you can join us in Bali.\n\n' +
    'Please RSVP by ' + RSVP_DEADLINE + '. It only takes a minute, and it helps us plan rooms and the weekend:\n' + e.link +
    '\n\nAlready replied? Thank you, and please ignore this note. Questions? Just reply to this email.\n\nAshley & Charles';
  var opts = mailOptions(to || e.email, (subjectPrefix || '') + REMIND_SUBJECT, reminderEmail(salutation, e.link), text);
  var cc = remindCc(); if (cc.length && !to) opts.cc = cc.join(', ');
  opts.inlineImages = { reminder: reminderImage() };
  MailApp.sendEmail(opts);
}

/** The copies on each reminder, leaving out the account doing the sending. Each one counts against the daily limit. */
function remindCc() {
  var me = ''; try { me = Session.getEffectiveUser().getEmail().toLowerCase(); } catch (err) {}
  return splitEmails(REMIND_CC).filter(function (x) { return x.toLowerCase() !== me; });
}
function reminderCost() { return 1 + remindCc().length; }   /* the guest plus each copy */
/** The side of the list this account sends reminders to ('C', 'A'), or '' for everyone. */
function remindSide() {
  var me = ''; try { me = Session.getEffectiveUser().getEmail().toLowerCase(); } catch (err) {}
  var side = ''; Object.keys(REMIND_SIDES).forEach(function (k) { if (k.toLowerCase() === me) side = REMIND_SIDES[k]; });
  return side;
}
function sideName(side) { return side === 'C' ? 'Charles\'s side (C)' : side === 'A' ? 'Ashley\'s side (A)' : 'both sides'; }

/** The Guest List's "Reminder sent" column, added at the end the first time it is needed. */
function reminderColumn() {
  var g = guestListColumns(false), h = g.rows[g.header];
  for (var c = 0; c < h.length; c++) if (nameKey(h[c]) === 'reminder sent') return { sheet: g.sheet, col: c + 1 };
  var col = g.sheet.getLastColumn() + 1; g.sheet.getRange(g.header + 1, col).setValue('Reminder sent');
  return { sheet: g.sheet, col: col };
}
function stampReminders(list) {
  var rc = reminderColumn(), now = new Date();
  list.forEach(function (e) { e.rows.forEach(function (r) { rc.sheet.getRange(r, rc.col).setValue(now).setNumberFormat('M/d/yyyy'); }); });
}
function describe(e) { return joinNames(e.names) + ' <' + e.email + '>' + (e.reminded ? '  (reminded ' + Utilities.formatDate(new Date(e.reminded), 'America/Los_Angeles', 'MMM d') + ')' : ''); }

function showUnreplied() {
  var list = unrepliedGuests(), people = 0, by = {}; list.forEach(function (e) { people += e.names.length; by[e.side || '?'] = (by[e.side || '?'] || 0) + 1; });
  var sides = Object.keys(by).sort().map(function (k) { return k + ': ' + by[k]; }).join(', ');
  SpreadsheetApp.getUi().alert(people + ' guest(s) in ' + list.length + ' email(s) have not replied yet (emails by side: ' + sides + ')', list.map(function (e) { return e.side + '  ' + describe(e); }).join('\n') || 'Everyone invited has replied.');
}

function sendTestReminder() {
  var me = Session.getEffectiveUser().getEmail(), list = unrepliedGuests();
  var e = list[0] || { email: me, names: ['Test Guest'], firsts: ['Test'], link: SITE_URL };
  sendReminder(e, me, '[Test] ');
  SpreadsheetApp.getUi().alert('Test reminder sent to ' + me + (list[0] ? ', written as if to ' + joinNames(e.names) + '. Nothing was stamped.' : '.'));
}

function sendSelectedReminders() {
  var ui = SpreadsheetApp.getUi(), active = ss().getActiveSheet();
  if (!active || active.getName() !== GUEST_LIST) { ui.alert('Select one or more rows on the ' + GUEST_LIST + ' tab first.'); return; }
  var range = active.getActiveRange(), r0 = range.getRow(), n = range.getNumRows(), picked = {};
  for (var r = r0; r < r0 + n; r++) picked[r] = true;
  var list = unrepliedGuests().filter(function (e) { return e.rows.some(function (r) { return picked[r]; }); });
  if (!list.length) { ui.alert('Nobody on the selected row(s) is waiting on a reminder: they have replied, have no email address, were never sent an invitation, or are marked Skip.'); return; }
  var answer = ui.alert('Send a reminder to ' + list.length + ' email(s)?', list.map(function (e) { return describe(e) + (e.reminded ? ': this would be another reminder' : ''); }).join('\n'), ui.ButtonSet.YES_NO);
  if (answer !== ui.Button.YES) return;
  if (MailApp.getRemainingDailyQuota() < list.length * reminderCost()) { ui.alert('Not enough of today\'s email limit left for ' + list.length + ' reminder(s): with the copies, each one counts ' + reminderCost() + '. Try again tomorrow, or select fewer rows.'); return; }
  list.forEach(function (e) { sendReminder(e); stampReminders([e]); });
  ui.alert('Sent ' + list.length + ' reminder(s). Each person\'s "Reminder sent" date is on the ' + GUEST_LIST + ' tab.');
}

/** Everyone who hasn't replied and wasn't reminded in the last REMIND_GAP_DAYS days, as many as today's email
    limit allows. Running it again later picks up where it stopped. */
function sendAllReminders() {
  var ui = SpreadsheetApp.getUi(), cutoff = new Date(Date.now() - REMIND_GAP_DAYS * 864e5);
  var side = remindSide();
  var all = unrepliedGuests().filter(function (e) { return !side || e.side === side || e.side === 'BOTH'; });
  var due = all.filter(function (e) { return !e.reminded || new Date(e.reminded) < cutoff; });
  if (!due.length) { ui.alert(all.length ? 'Everyone on ' + sideName(side) + ' who hasn\'t replied was reminded in the last ' + REMIND_GAP_DAYS + ' days. Nothing to send.' : 'Everyone invited on ' + sideName(side) + ' has replied. Nothing to send.'); return; }
  var left = MailApp.getRemainingDailyQuota(), cost = reminderCost(), now = due.slice(0, Math.max(0, Math.floor((left - 2) / cost)));   /* each reminder counts the guest plus its copies; keep a couple back for replies */
  if (!now.length) { ui.alert('Not enough of today\'s email limit is left for a reminder and its ' + (cost - 1) + ' copies. ' + due.length + ' reminder(s) are waiting; run this again tomorrow.'); return; }
  var later = due.length - now.length;
  var answer = ui.alert('Send ' + now.length + ' reminder(s) to ' + sideName(side) + ' now?',
    now.map(describe).join('\n') + '\n\nEach is copied to ' + (remindCc().join(', ') || 'nobody') + '.' + (later ? '\n\n' + later + ' more will wait for tomorrow: with the copies, today\'s email limit allows ' + now.length + '.' : ''), ui.ButtonSet.YES_NO);
  if (answer !== ui.Button.YES) return;
  var sent = 0;
  try { now.forEach(function (e) { sendReminder(e); stampReminders([e]); sent++; }); }
  catch (err) { ui.alert('Stopped after ' + sent + ' reminder(s): ' + err.message + '\n\nEveryone already sent is stamped, so running this again carries on from here.'); return; }
  ui.alert('Sent ' + sent + ' reminder(s).' + (later ? ' ' + later + ' still to go: run this again tomorrow and it carries on from there.' : ' That is everyone.'));
}

