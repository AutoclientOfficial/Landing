# Client Dashboard automation (Google Sheets)

Live dashboard for the spreadsheet
[`1NOxsaHYowzYiiildq9FbMPRE7wyPyNd-s9DtYXtpPaQ`](https://docs.google.com/spreadsheets/d/1NOxsaHYowzYiiildq9FbMPRE7wyPyNd-s9DtYXtpPaQ/edit).

You and the team edit the **working tabs** (every tab except the client-facing
ones). The script aggregates them and renders, on the **Clients** and **POV**
tabs, a status summary table, a **pie chart**, and a **progress bar** — refreshed
automatically so the client always sees the latest version.

## How it works

- **Editable tabs** = all tabs *except* `Clients` and `POV` (and any tab whose
  name starts with `_`).
- Each working tab should have a **`Status`** column (header on row 1). Accepted
  header names: `Status`, `Stage`, `Progress`, `State`.
- The script counts rows per status across all working tabs, then on each client
  tab draws:
  - a **Status / Count** table,
  - a **pie chart** (status breakdown),
  - a **progress bar** = `Done ÷ total` as a percentage.
- It re-runs on every edit (installable `onEdit` trigger), so the client view
  stays in sync.

## Setup (one time, ~2 minutes)

1. Open the sheet → **Extensions → Apps Script**.
2. Delete the default `Code.gs` contents, paste in
   [`ClientDashboard.gs`](./ClientDashboard.gs), and **Save**.
3. In the function dropdown choose **`setup`** and click **Run**. Approve the
   permission prompt (it needs access to this spreadsheet only).
4. Reload the sheet. A **Client Dashboard** menu appears with:
   - **Refresh now** — rebuild the dashboards on demand.
   - **Install auto-refresh (on edit)** — re-installs the trigger if needed.

## Customizing

All settings live in the `CONFIG` object at the top of `ClientDashboard.gs`:

| Setting | What it controls |
|---|---|
| `clientTabs` | The client-facing tabs to render onto. Default `['Clients', 'POV']`. |
| `statusHeaderAliases` | Header names that identify the status column. |
| `headerRow` | Row that holds your headers (default `1`). |
| `statuses` | Status names, **pie colors**, and which ones count as *complete* for the progress bar. |
| `anchorCell` | Top-left cell of the dashboard block on each client tab (default `J1`). Move it so it doesn't overlap existing client content. |

Notes:
- Unknown status values are still counted and shown (with a default color).
- For a donut instead of a pie, set `pieHole` to `0.4` in `rebuildPieChart`.
- Charts float over the grid; only the summary table and progress bar write into
  cells (starting at `anchorCell`), so keep that area clear.

## Why a script (and not just formulas)

A pure-formula version can mirror values live, but **pie charts** and an
embedded **progress bar** can't be created by formulas alone — they require the
Sheets chart API, which Apps Script provides. The script also lets the dashboard
auto-refresh on every edit without manual steps.
