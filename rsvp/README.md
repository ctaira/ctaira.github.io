# RSVP setup

Replies are collected in a Google Sheet through a small Apps Script web app.
Each household gets a private link with an invite code. The page looks the
code up to personalise the passport and the RSVP form, and the script checks
the code again on every submission, so nobody without a link can reply and
nobody can reply for more seats than they were given.

## 1. The sheet

The guest list is your own **Wedding Planning** workbook, one row per person,
with the columns it already has: **First Name**, **Last Name**, **Side (C/A)**,
**Email Address** and **Linked to another guest?**. The script reads the tab
called **Guest List** and groups people into households on its own:

- anyone named in **Linked to another guest?** joins that person's household,
  and anyone sharing an email address does too;
- a first name like `Kevin's Girlfriend` or `Leo's Kimberly` joins Kevin's or
  Leo's household;
- if the tab has a column headed **Room** or **Household**, everyone sharing
  a value in it is one household (the room numbers from the Room List tab
  work well here, copied across);
- a row counts as a guest only when it has a first name and a side (`A`, `C`
  or `Both`), so note rows and section headings are skipped;
- seats are the number of people in the household, and the passport reads
  their names, e.g. `Jimmie, Dorothy & Tiffany Dong` or `Michael Lai & Lalita Lai`.

People who should be invited together but are not linked (siblings on
separate rows, say) get separate invitations. Fill in **Linked to another
guest?** for them and rebuild.

The workbook must be a native Google Sheet for the script to run in it. If it
is still an `.xlsx`, open it in Drive and choose **File → Save as Google
Sheets**, then work in the new copy.

The script keeps its own tab, **Households**, with one row per household:
`code`, `names`, `seats`, `email`, `link`, `sent` and `members`. Never edit
this tab by hand; rebuild it from the menu. A **Responses** tab is created
automatically the first time the script runs.

## 2. The script

Emails are sent by whichever Google account deploys the script, so do this
part signed in as the account you want guests to hear from
(ashley.plus.charles@gmail.com). Transferring ownership of the sheet to
that account first keeps everything in one place.

1. In the sheet, open **Extensions → Apps Script**.
2. Delete the placeholder and paste the contents of `Code.gs`.
3. Set `SITE_URL` at the top to your domain, with a trailing slash. The
   other settings there (sender name, reply-to, subjects, deadline) are
   worth a glance too.
4. Save, then **Deploy → New deployment**. Type: **Web app**. Execute as
   **Me**. Who has access: **Anyone**. Click Deploy and authorise it.
5. Copy the **Web app URL** (it ends in `/exec`).

Whenever you change the script later, use **Deploy → Manage deployments →
Edit → Version: New version**, or the live URL keeps running the old code.

## 3. Households, codes and links

Reload the sheet. An **Invitations** menu appears.

1. Run **Build households from Guest List**. The **Households** tab is
   (re)built from the Guest List and every household gets a six-character
   code and a link like `https://ashleyandcharlesinbali.com/?i=k7m3qx`.
   Read the tab through once: names, seats and emails are what the
   invitations will carry.
2. Run it again whenever the Guest List changes. Households whose members
   are unchanged keep their code and their sent date; a household whose
   members changed gets a new row and a new code, and its old row is
   removed. Existing codes are never rewritten.

**Fill missing codes and links** is also on the menu for a Households tab
you have edited by hand. To disable a link, delete its code.

## 4. Sending the invitations

Emails go out from the Google account that deployed the script, under the
sender name set in `FROM_NAME`. Replies from guests land in that account's
inbox unless `REPLY_TO` is set.

1. **Invitations → Send a test invitation to me** emails you the invitation
   for the first row, so you can see it before anyone else does.
2. **Invitations → Send invitations to unsent rows** emails every household
   that has a code and an email but nothing in **sent**, then writes the
   send time into **sent**. It asks before sending and tells you how many
   emails your account can still send today.

The invitation email carries the beach artwork inside it. The script fetches
`assets/email-hero-2x.jpg` from the live site once per run and embeds it, so
publish the site and set `SITE_URL` before sending. With images off, guests
still see the headline as the picture's description, the note, and the button.

A personal Gmail account can send about 100 emails a day. If the list is
longer, run the same menu item again the next day; rows already sent are
skipped. Google Workspace accounts have a much higher limit. To resend to
one household, clear its **sent** cell and run it again.

## 5. Confirmation emails

Every reply triggers a confirmation to the address the guest entered, with
what they said and their link, and a note that sending a new reply before
the deadline replaces the old one. If the email can't be sent, the reply is
still recorded; the guest just doesn't get the receipt.

## 6. The site

In `index.html`, near the top of the script, set:

```js
var RSVP_ENDPOINT = "https://script.google.com/macros/s/.../exec";
```

That's it. With the endpoint set, a bare URL or an unknown code still shows
the whole invitation, but the RSVP form is replaced with a note asking the
guest to use the link they were sent.

With the endpoint empty, the page runs in preview mode: `?to=` and `?seats=`
personalise it directly and submissions are logged to the console instead of
sent. Useful for checking the design before the sheet exists.

## What lands in Responses

One row per submission: received time, code, the invited names and seats from
your list, then the name, email, attending choice (`both`, `saturday`, `no`),
party size, dietary needs and note the guest entered, and the time their
browser sent it. If someone replies twice, both rows are kept; the later one
is their current answer. An accidental double-send of the same reply is
dropped.
