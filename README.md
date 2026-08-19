# 🗓️ Make Dates Perfect

**A demo-prep wizard for Microsoft Dynamics 365 / Dataverse.** It slides the dates on your
records — opportunities, leads, cases, contacts, accounts — and their related touchpoints
(emails, phone calls, appointments, tasks) **forward in time**, so aging demo data looks
fresh and current again **without changing the story**.

> ⚠️ **For demo environments only.** This tool is designed to keep demo/POC/sandbox data
> looking current. It performs bulk date edits on business records and is **not** intended
> for production or any environment containing real customer data. Use it only where
> rewriting dates is safe and expected.

---

## Why it exists

Demo environments are built once but demoed for months. Over time:

- Opportunity **close dates** drift too near (or into the past), breaking your storytelling.
- The **touchpoints** that carry the narrative (emails, calls, meetings, tasks) sink into
  the distant past, so **timelines and dashboards look empty** near "today."

Make Dates Perfect fixes this by computing a single **whole-day offset per record** and
adding it to every relevant date. Because everything moves by the same amount, the
**relative spacing between events is preserved** — the story is identical, just
re-centred on today.

---

## How it works

For **each selected record**, independently:

1. **Gather touchpoints** — all related activities via `activitypointer` (`regardingobjectid`).
2. **Find the anchor** — the most recent touchpoint date. (Close/own dates are excluded from
   the anchor so they land in the future.)
3. **Compute the offset** — `offsetDays = round(today − anchor) − padding`.
4. **Apply uniformly** — add `offsetDays` to every shiftable date on the record and its
   activities. Same whole-day offset everywhere ⇒ time-of-day and spacing preserved.

Anchoring on the latest **activity** (not the close date) is deliberate: it keeps close
dates comfortably in the future, fixing the "too close" problem while the story stays intact.

---

## Features

- **Multi-table** — Opportunities, Leads, Cases, Contacts, Accounts (each with its own
  pastel pill, icon, and live selected-count).
- **Per-record selection** — pick exactly what you want; selections accumulate across tables.
- **Mandatory dry-run preview** — see every record/field, current → new value, and Δ days
  before anything is written. Collapsible per record.
- **Update-only** — it **never deletes** anything.
- **Fully reversible** — every change is captured in an undo log (in-app Undo + JSON download).
- **Runs inside D365** — a model-driven app / web resource using `Xrm.WebApi`, so it
  authenticates automatically. No secrets in the browser.

---

## Prerequisites

- **Node.js 18+** — only for **Option B** (the installer script); not needed if you import the solution.
- A **Dynamics 365 / Dataverse environment** you may customize (create solutions, edit records).
- An account with the **System Customizer** (or System Administrator) role in that environment.

> On Windows PowerShell, if `npm` is blocked by execution policy, use `npm.cmd` instead of `npm`.

---

## Install

Both options install a solution named **`MakeDatesPerfect`** containing a model-driven app,
its site map, and the wizard web resource + icon. Pick whichever you prefer.

### Option A — Import the solution package (no Node.js)

Prebuilt solution zips are in [`/solution`](solution):

1. Go to [make.powerapps.com](https://make.powerapps.com) and select your environment.
2. **Solutions → Import solution → Browse**, and choose:
   - [`solution/MakeDatesPerfect_managed.zip`](solution/MakeDatesPerfect_managed.zip) — recommended for consumers, or
   - [`solution/MakeDatesPerfect_unmanaged.zip`](solution/MakeDatesPerfect_unmanaged.zip) — if you want to customize it.
3. **Next → Import**, then **Publish all customizations** when it finishes.
4. Open the **Make Dates Perfect** app from the **Apps** launcher.

### Option B — Run the installer script

Uses **MSAL device-code sign-in** (no stored passwords).

```bash
# 1. Clone
git clone https://github.com/hyagomatielo/make-dates-perfect.git
cd make-dates-perfect

# 2. Install dependencies
npm install

# 3. Deploy to YOUR environment (pass your org URL)
node deploy.js https://yourorg.crm.dynamics.com
```

You can also set the URL via an environment variable instead of an argument:

```bash
# PowerShell
$env:CRM_URL = "https://yourorg.crm.dynamics.com"; node deploy.js
```

When prompted, open the sign-in URL, enter the device code, and authenticate. The installer
then creates the solution, uploads the app, and publishes. On success it prints the app URL:

```
https://yourorg.crm.dynamics.com/main.aspx?appid=<appid>
```

The app also appears in the **Apps** launcher as **"Make Dates Perfect."**

> **Maintainers:** to refresh the packaged zips from an environment, run
> `node export-solution.js https://yourorg.crm.dynamics.com` (writes to `/solution`).

---

## Using the tool

1. **Open** the app (from the Apps launcher or the printed URL). It must run **inside**
   Dynamics 365 so it can authenticate; opening the raw HTML file standalone won't work.
2. **Step 1 — Select records.** Pick a table pill (Opportunities / Leads / Cases / Contacts /
   Accounts), search, and tick records. Use **Select all** for the current page. Selections
   add up across tables (the pill badges show how many you've picked in each).
3. **Step 2 — Configure.** Choose which touchpoint tables to move (Emails, Phone Calls,
   Appointments, Tasks), whether to shift each record's own dates, and how many days before
   today the latest touchpoint should land (0 = today).
4. **Step 3 — Preview.** A dry-run shows every record/field, current → new, and Δ days.
   Expand each collapsible group to inspect. **Nothing is written yet.**
5. **Step 4 — Apply.** Watch the progress log as updates are written.
6. **Step 5 — Done.** A summary appears with **Undo all changes**, **Download undo log (JSON)**,
   and **Start over**.

---

## What gets changed

| Table | Record's own date fields shifted |
|---|---|
| Opportunity | `estimatedclosedate`, `actualclosedate` |
| Case (incident) | `followupby` |
| Lead / Contact / Account | *(activities only — no own business-date fields)* |

Plus, for **any** selected record, the related **activities** you enabled:

| Activity | Fields shifted |
|---|---|
| Email | `scheduledstart`, `scheduledend`, `actualstart`, `actualend` |
| Phone Call | `scheduledstart`, `scheduledend`, `actualstart`, `actualend` |
| Appointment | `scheduledstart`, `scheduledend`, `actualstart`, `actualend` |
| Task | `scheduledstart`, `scheduledend`, `actualstart`, `actualend` |

---

## Limitations — what is **not** changed, and why

- **Notes / annotations are not shifted.** An annotation only carries `createdon` /
  `modifiedon`, both of which are **system-managed and read-only**. The wizard lists them as
  *skipped* for transparency rather than silently ignoring them.
- **`createdon` / `modifiedon` are never changed.** These system audit stamps cannot be
  updated via the API on existing records. The timeline control sorts primarily by
  `createdon`, so a record's *position* in the timeline follows its shifted **business**
  dates (which drive the visible dates), not a rewritten created-on stamp.
- **Only the listed tables and fields are touched.** Other entities (quotes, orders, custom
  tables) and any date field not in the tables above are left untouched. This is a
  conservative, safe default — see "Extending" below.
- **Records with no activities and no own date fields** produce no change (nothing to anchor on).
- **Per-record selection only.** There is intentionally **no** "fix every record in the table"
  bulk mode — the tool always operates on records you explicitly select, scoped and previewed.
- **Read-only or secured fields may fail to update.** If field-level security or a plugin
  blocks a write, that record is reported as failed in the log; others still succeed.
- **Icons load from a CDN.** The UI pulls Fluent System Icons from `cdn.jsdelivr.net`. If your
  environment's browser has no outbound internet or a strict Content-Security-Policy, the
  icons may not render (functionality is unaffected).

---

## Safety & reversibility

- **Update-only** — the tool never deletes any record under any circumstance.
- **Mandatory preview** — nothing is written until you review Step 3 and click Apply.
- **Undo log** — every applied change stores its original value; use the in-app **Undo** or
  the downloaded JSON to roll everything back.
- **Scoped** — only records you explicitly select are affected — never a blanket table update.

---

## Sharing the app with others

Like any model-driven app, colleagues only see it once a **security role** is associated:

1. Open the app in the maker portal ([make.powerapps.com](https://make.powerapps.com)) →
   your environment → **Apps** → *Make Dates Perfect* → **⋯ → Manage roles** (or Edit →
   Settings → Roles).
2. Add the roles your users have, then **Save**.

---

## Extending it to more tables/fields

The installer is data-driven. In `make-dates-perfect.html`, the `TABLES` registry maps each
root table to its id/name field and its own writable date fields:

```js
const TABLES = {
  opportunity: { label: "Opportunity", idField: "opportunityid", nameField: "name",
    ownFields: [{ name: "estimatedclosedate", dateOnly: true }, { name: "actualclosedate", dateOnly: true }] },
  // add your own entity here…
};
```

Add an entry (and a matching pastel pill in the markup), then re-run `node deploy.js <url>`.

---

## Uninstall

Delete the **`MakeDatesPerfect`** solution from your environment
([make.powerapps.com](https://make.powerapps.com) → Solutions → *Make Dates Perfect* → Delete).
No records created by the tool remain (it only edited dates on existing records).

---

## Tech notes

- Front end: a single self-contained HTML web resource using `Xrm.WebApi`
  (`retrieveMultipleRecords` / `updateRecord`).
- Installer: `deploy.js` (Node.js) authenticates via MSAL device-code flow and creates the
  solution, app module, site map, and web resource via the Dataverse Web API.
- No server, no database, no stored credentials.

---

## License

[MIT](LICENSE) © 2026 Hyago Matielo

