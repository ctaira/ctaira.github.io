# RSVP setup

Replies are collected in a Google Sheet through a small Apps Script web app.
Each household gets a private link with an invite code. The page looks the
code up to personalise the passport and the RSVP form, and the script checks
the code again on every submission, so nobody without a link can reply and
nobody can reply for more seats than they were given.

## 1. The sheet

Create a Google Sheet. Rename the first tab **Guests** and give it these
headers in row 1:

| A `code` | B `names` | C `seats` | D `email` | E `link` |
|---|---|---|---|---|

Add one row per household. Fill in **names** (what the passport will say, e.g.
`Sam & Priya Nair`) and **seats** (1 to 6). Leave **code** and **link** empty;
the script fills them in. **email** is optional and prefills the form.

A **Responses** tab is created automatically the first time the script runs.

## 2. The script

1. In the sheet, open **Extensions → Apps Script**.
2. Delete the placeholder and paste the contents of `Code.gs`.
3. Set `SITE_URL` at the top to your domain, with a trailing slash.
4. Save, then **Deploy → New deployment**. Type: **Web app**. Execute as
   **Me**. Who has access: **Anyone**. Click Deploy and authorise it.
5. Copy the **Web app URL** (it ends in `/exec`).

Whenever you change the script later, use **Deploy → Manage deployments →
Edit → Version: New version**, or the live URL keeps running the old code.

## 3. Codes and links

Reload the sheet. An **Invitations** menu appears. Run **Fill missing codes
and links**. Every household row now has a six-character code and a link
like `https://yourdomain.com/?i=k7m3qx`. Send each household its own link.

Rows you add later get codes the next time you run the menu item. Existing
codes are never changed. To disable a link, delete its code.

## 4. The site

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
