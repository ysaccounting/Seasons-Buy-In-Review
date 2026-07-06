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

- **Company** — records are split by a **`Company`** column in the HAL file, and
  **each company gets its own workbook**. If a HAL file has no `Company` column,
  the whole file is treated as one company named after the file. This lets you
  drop several companies' HAL lists in together and hand each company its report.
- **League / Year** — chosen from dropdowns at the top (League: MLB, MLS, NBA,
  NFL, NHL, NCAAF, NCAAB, WNBA, Racing · Year: 2025-26 … 2028-29). They apply to
  the whole batch, appear on each **Summary** tab, and drive the file name.

Output file name convention:

    Seasons Review - {Company} - {League} - {Year}.xlsx

When more than one company is produced, all reports are also bundled into
`Seasons Review - {League} - {Year}.zip`; the UI additionally offers a per-company
download link for each report.

## The UI

- Two dropdowns — **League** and **Year**.
- Two drop zones — **Season Ticket Database** (HAL) and **Purchase Details**
  (TicketVault), each accepting one or more files.
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
- **Not Reconciled** — same columns preceded by a **Notes** column. Notes are:
  - **Not bought in** — no matching rows in TicketVault for that seat block.
  - **total cost not equal** — present, but the amounts differ by more than the tolerance.
  - **# games not equal** — present, but TicketVault games-with-cost ≠ HAL # games.
  - **total cost not equal, # games not equal** — both.

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
| `index.html` | Single-page UI: League/Year dropdowns + two drop zones |
| `requirements.txt` | Python dependencies |
| `Procfile` / `railway.json` | Start command for Railway |
