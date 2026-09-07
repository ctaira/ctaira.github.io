# RSVP setup

Replies are collected in a Google Sheet through a small Apps Script web app.
Each household gets a private link with an invite code. The page looks the
code up to personalise the passport and the RSVP form, and the script checks
the code again on every submission, so nobody without a link can reply and
nobody can reply for more seats than they were given.

## 1. The sheet

Create a Google Sheet. The quickest way is **File → Import → Upload** with
`guests-template.csv` from this folder, then rename the tab to **Guests**
and delete the two example rows. Or start blank, name the first tab
**Guests**, and give it these headers in row 1:

| A `code` | B `names` | C `seats` | D `email` | E `link` | F `sent` |
|---|---|---|---|---|---|

Add one row per household. Fill in **names** (what the passport will say, e.g.
`Sam & Priya Nair`) and **seats** (1 to 6). Put every adult's address in
**email**, separated by commas, so each person in the household gets the
same link. Leave **code**, **link** and **sent** empty; the script fills them.

A **Responses** tab is created automatically the first time the script runs.

## 2. The script

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

## 3. Codes and links

Reload the sheet. An **Invitations** menu appears. Run **Fill missing codes
and links**. Every household row now has a six-character code and a link
like `https://yourdomain.com/?i=k7m3qx`. The next step emails them out, or
you can copy a link from column E to send by text.

Rows you add later get codes the next time you run the menu item. Existing
codes are never changed. To disable a link, delete its code.

## 4. Sending the invitations

Emails go out from the Google account that owns the script, under the sender
name set in `FROM_NAME`.

1. **Invitations → Send a test invitation to me** emails you the invitation
   for the first row, so you can see it before anyone else does.
2. **Invitations → Send invitations to unsent rows** emails every household
   that has a code and an email but nothing in **sent**, then writes the
   send time into **sent**. It asks before sending and tells you how many
   emails your account can still send today.

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
