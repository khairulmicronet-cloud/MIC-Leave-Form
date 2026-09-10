# Staff Leave Breakdown — one-time setup

The Staff Leave Breakdown page (`leave-breakdown.html`) stores its data in a
Google Sheet, read and written through a small Google Apps Script "Web App"
that acts as a free API. You only need to do this once.

## 1. Create the Google Sheet

1. Go to [sheets.google.com](https://sheets.google.com) and create a new,
   blank spreadsheet. Name it something like **MIC Staff Leave Ledger**.
2. You don't need to create any tabs or columns yourself — the script
   creates the `StaffConfig` and `LeaveEntries` tabs automatically the
   first time it runs.

## 2. Add the Apps Script

1. In the Sheet, go to **Extensions > Apps Script**.
2. Delete anything in the default `Code.gs` file, and paste in the full
   contents of [`google-apps-script/Code.gs`](google-apps-script/Code.gs)
   from this repo.
3. Click the **Save** icon (or Ctrl/Cmd+S).

## 3. Deploy it as a Web App

1. Still in the Apps Script editor, click **Deploy > New deployment**.
2. Click the gear icon next to "Select type" and choose **Web app**.
3. Set:
   - **Execute as:** Me (your Google account)
   - **Who has access:** Anyone
     (this only exposes the two actions defined in `Code.gs` — reading and
     writing leave rows in this one sheet — not your whole Google account)
4. Click **Deploy**. The first time, Google will ask you to authorize the
   script — click through the "unverified app" warning (it's your own
   script) and allow access.
5. Copy the **Web app URL** it gives you — it looks like
   `https://script.google.com/macros/s/AKfycb.../exec`.

## 4. Connect the site to it

1. Open `config.js` in this repo.
2. Set:
   ```js
   const LEAVE_LEDGER_API_URL = "https://script.google.com/macros/s/AKfycb.../exec";
   ```
   (using the URL you copied in step 3).
3. Commit and push the change. Once deployed, `leave-breakdown.html` will
   load and let you add/view leave entries for each staff member.

## 5. Set each staff member's entitlement

Open the Staff Leave Breakdown page, pick a staff member and year, and use
**"Edit entitlement / carry-forward for this staff member & year"** at the
bottom to enter their Annual/Sick leave entitlement and any balance carried
over from the previous year. This only needs doing once per person per
year — the page computes running balances automatically from there as you
add leave entries.

## Updating the script later

If you ever change `google-apps-script/Code.gs` in this repo, you need to
re-paste it into the Apps Script editor and create a **new deployment**
(Deploy > Manage deployments > pencil icon > New version) for the change to
take effect — editing the file in GitHub alone does not update the live
script.
