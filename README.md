# Season Ticket Buy-In Review

Web app: pick a **League** and **Year**, upload the **HAL season-ticket
database** file(s) plus the TicketVault **Purchase Details** export(s)
(`.xlsx`, `.xlsm`, or `.csv`), and download a workbook **per company** that
reconciles every HAL record against what is actually loaded in TicketVault and
flags anything that doesn't match.

## What it does

Every seat block in HAL **must** be bought into TicketVault for the correct
amount and the right number of games. Rows that exist only in Purchase Details
(no matching HAL record) are ignored — the check runs **HAL → TicketVault only**.

1. **Match by seat block.** A HAL record is matched to TicketVault on
   **email + team + section/row + individual seat number** (seat ranges are
   expanded, so a block split across several POs in TicketVault still matches).
   The `(P)` columns in HAL are ignored.
2. **Check the amount.** The HAL **Total Cost** must equal the **sum of Total
   Cost** across all matching TicketVault rows, within a tolerance (default
   **$1.00**).
3. **Check the games.** The number of TicketVault games carrying a cost
   (**# Games w/Cost**) must equal the HAL **# Games**.

A record is **Reconciled** only when the amount ties **and** the games match.

## Companies, League & Year

- **Company** — every uploaded HAL file gets a **Company dropdown** (the list is
  the *Short Name* values from `Master_Mapping_List.xlsx`, which is bundled). The
  app pre-fills a best guess from the file name; change it if wrong. Each company
  gets **its own workbook**, and files that share a company merge into one report.
- **League / Year / As Of Date** — chosen at the top. League (MLB…Racing), Year
  (2025-26 … 2028-29), and an **As Of Date** picker. All three apply to the whole
  batch, appear on each **Summary** tab, and drive the file name.
- **Company scoping.** Each broker's HAL is reconciled **only** against the
  Purchase-Details rows whose `Company` belongs to that broker, per the
  *TicketVault Company* column of `Master_Mapping_List.xlsx` (e.g. Chase →
  "Jacks YS", GK → "GK LLC").

Output file name convention:

    Seasons Review - {Company} - {League} - {Year} - As Of {date}.xlsx

When more than one company is produced, all reports are also bundled into
`Seasons Review - {League} - {Year} - As Of {date}.zip`; the UI additionally
offers a per-company download link for each report.

## Different HAL layouts

Every broker formats their HAL differently, so the app maps columns by a list of
known header names per field (see `HAL_SYNONYMS` in `app.py`) rather than fixed
positions. To support a new broker whose headers differ, add its column names to
the relevant list there. Two quirks are handled automatically:

- **Date-corrupted seats.** Excel turns ranges like `3-9` into a date (Mar 9);
  the app converts those back to `month-day`.
- **Year-column cost.** Some HALs keep the total in a season-year column (e.g.
  Levovitz uses `26/27`); the app looks there when there's no plain total column,
  using the selected Year. It also finds year-named plan columns (GK `2026 Plan`).
- **Email not in an obvious column.** If no email header matches, the app finds
  the column whose values look like email addresses (GK `Profiles`, TL `Name`).
- **Team names.** Brokers write teams as nicknames or abbreviations
  (`Bengals`, `49ers`, `ATL`); both sides are normalized to TicketVault's full
  `City Nickname` for the selected League. Teams outside that league pass through
  unchanged (so run each league as its own batch).
- **Non-active rows** (deposits, waitlist, inquiries, cancellations) are dropped
  from the review; the count is reported after processing.
- **Secondary-market vendors** (Ticketmaster, TickPick, StubHub, Ticket Evolution,
  GoTickets) are dropped from Purchase Details before reconciling.

## The UI

- Two dropdowns — **League** and **Year**.
- Two drop zones — **Season Ticket Database** (HAL, with a Company dropdown per
  file) and **Purchase Details** (TicketVault), each accepting one or more files.
- Reconcile stays disabled until League, Year, and at least one file in each zone
  are set. After processing you get an overall verdict and counts, a per-company
  list with individual download links, and a **Download all (ZIP)** button.

## Output tabs (per company)

- **Summary** — title, the **Company / League / Year**, a **RESULT** rollup
  (Reconciled, Not Reconciled, Total # HAL Records) and a **NOT RECONCILED — BY
  REASON** rollup (Not bought in, Partially bought in, TOTAL).
- **Reconciled** — one row per reconciled HAL record. Columns grouped **per HAL**
  (Team, Email, Full/Partial, Section, Row, Seats, Qty, # Games, Total Cost) and
  **per TicketVault** (Total Cost, # Games w/Cost, # Games w/o Cost).
- **Not Reconciled** — same columns preceded by a **Notes** column, plus a gold
  **Variances** section (Total Cost delta, # Games w/Cost delta, and the alternate
  Email Address for different-email hits). Notes are:
  - **Not bought in** — no matching rows in TicketVault for that seat block.
  - **total cost not equal** — present, but the amounts differ by more than the tolerance.
  - **# games not equal** — present, but TicketVault games-with-cost ≠ HAL # games.
  - **# games not in HAL** — present, but the HAL has no game count to check against.
  - **different email address** — the right seats/cost/games exist in TicketVault
    but under a different email; that email is shown in the Variances section.
  - combinations of the above are joined with a comma.

### How edge cases are handled

- **Seat ranges are expanded** to individual seats, so a HAL block of seats
  1–7 still reconciles when TicketVault loaded it as 1–2 and 3–7.
- **Split orders in Purchase Details** (one seat block across several PO rows)
  are summed for the cost, and distinct event dates are counted for the games.
- **Parking** is matched at the team level (slot numbers differ between systems);
  parking loaded at $0 will show as a cost/games discrepancy.
- **Multiple emails in one HAL cell** (e.g. `a@x.com,b@x.com`) are each tried.
- Amounts parse from plain numbers, `$1,234.56`, `=123.45` literals, and `(123)` negatives.

## Run locally

```bash
pip install -r requirements.txt
python app.py        # http://localhost:5000
```

## Deploy: GitHub → Railway

1. Push this folder to a GitHub repo.
2. Railway → New Project → Deploy from GitHub repo → pick it.
3. Railway auto-detects Python (Nixpacks) and uses the start command in `railway.json`.
   No env vars needed; `$PORT` is provided automatically.

## Files

| File | Purpose |
|------|---------|
| `app.py` | Flask backend — parsing, per-company reconciliation, workbook builder |
| `index.html` | Single-page UI: League/Year dropdowns + per-file Company dropdowns |
| `Master_Mapping_List.xlsx` | Source of the Company dropdown (Short Name column) |
| `requirements.txt` | Python dependencies |
| `Procfile` / `railway.json` | Start command for Railway |
