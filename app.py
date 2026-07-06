"""
Season Ticket Buy-In Review — backend.

Reconciles the HAL-NFL season-ticket database against the TicketVault
Purchase Details exports. Every HAL record (one seat block — tickets or
parking) is checked against what is actually loaded in TicketVault.

A HAL record is RECONCILED when both are true:
  * the TicketVault total cost for that seat block ties to the HAL total
    cost, within a tolerance (default $1.00), and
  * the number of TicketVault games carrying a cost equals the HAL # games.

Otherwise it is NOT RECONCILED, with a note:
  * "Not bought in"                              — no matching rows in TicketVault
  * "total cost not equal"                       — present, cost doesn't tie
  * "# games not equal"                          — present, games-with-cost differ
  * "total cost not equal, # games not equal"    — both

Match key: email + team + section/row + individual seat number. Direction is
HAL -> TicketVault only (vault rows with no HAL record are ignored).

Multiple companies can be uploaded together. Records are split by a "Company"
column in the HAL file (falling back to the file name), and EACH COMPANY gets
its own workbook. League and Year are chosen in the UI and drive the file name:

    Seasons Review - {Company} - {League} - {Year}.xlsx

Output workbook tabs: Summary (with Company / League / Year), Reconciled,
Not Reconciled.
"""

import io
import os
import csv
import re
import time
import uuid
import zipfile
import tempfile
from collections import defaultdict, OrderedDict

from flask import Flask, request, jsonify, send_file, send_from_directory, abort
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.styles.colors import Color
from openpyxl.utils import get_column_letter

app = Flask(__name__, static_folder=None)
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
STORE_DIR = os.path.join(tempfile.gettempdir(), "recon_store")
os.makedirs(STORE_DIR, exist_ok=True)

DEFAULT_TOLERANCE = 1.00  # dollars; TV total must tie to HAL total within this
LEAGUES = ["MLB", "MLS", "NBA", "NFL", "NHL", "NCAAF", "NCAAB", "WNBA", "Racing"]
YEARS = ["2025-26", "2026-27", "2027-28", "2028-29"]

# --------------------------------------------------------------------------- #
# Generic helpers
# --------------------------------------------------------------------------- #

_NUM_FORMULA = re.compile(r"-?\d+(\.\d+)?")


def _cleanup_old(max_age_seconds=12 * 3600):
    now = time.time()
    for name in os.listdir(STORE_DIR):
        path = os.path.join(STORE_DIR, name)
        try:
            if now - os.path.getmtime(path) > max_age_seconds:
                if os.path.isdir(path):
                    for f in os.listdir(path):
                        os.remove(os.path.join(path, f))
                    os.rmdir(path)
                else:
                    os.remove(path)
        except OSError:
            pass


def _amount(v):
    if v is None or isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip()
    if s in ("", " "):
        return None
    if s.startswith("="):
        body = s[1:]
        return float(body) if _NUM_FORMULA.fullmatch(body) else None
    neg = s.startswith("(") and s.endswith(")")
    s = re.sub(r"[^0-9.\-]", "", s)
    if s in ("", "-", "."):
        return None
    try:
        f = float(s)
        return -f if neg else f
    except ValueError:
        return None


def _rows_from_upload(filename, data):
    low = filename.lower()
    if low.endswith(".csv"):
        text = data.decode("utf-8-sig", errors="replace")
        return [list(r) for r in csv.reader(io.StringIO(text))]
    wb = load_workbook(io.BytesIO(data), data_only=True, read_only=True)
    best, best_score = None, -1
    for ws in wb.worksheets:
        score = (ws.max_row or 0) * (ws.max_column or 1)
        if score > best_score:
            best, best_score = ws, score
    rows = [list(r) for r in best.iter_rows(values_only=True)]
    wb.close()
    return rows


def _norm_header(s):
    return re.sub(r"\s+", " ", str(s or "").strip()).lower()


def _find_header_row(rows):
    for i, row in enumerate(rows[:15]):
        cells = [_norm_header(c) for c in row]
        if len([c for c in cells if c]) >= 2:
            return i, cells
    return 0, [_norm_header(c) for c in (rows[0] if rows else [])]


def _col_index(header, *names):
    wanted = [_norm_header(n) for n in names]
    for i, h in enumerate(header):
        if h in wanted:
            return i
    return None


def _cell(row, i):
    return row[i] if (i is not None and i < len(row)) else None


def _safe_name(s):
    """Sanitize a string for use in a file name."""
    return re.sub(r'[\\/:*?"<>|]+', " ", str(s or "")).strip() or "Unknown"


def _company_from_filename(filename):
    stem = os.path.splitext(os.path.basename(filename))[0]
    return stem.strip() or "Unknown"


# --------------------------------------------------------------------------- #
# Seat / key normalization
# --------------------------------------------------------------------------- #

def _emails(cell):
    return [x.strip().lower() for x in re.split(r"[,;]", str(cell or ""))
            if x.strip() and x.strip().lower() != "nan"]


def _team(cell):
    return str(cell or "").strip()


def _sec(cell):
    return re.sub(r"\s+", " ", str(cell or "").strip().upper())


def _row(cell):
    return str(cell or "").strip().upper()


def _seatnums(cell):
    s = str(cell or "").replace(" ", "")
    m = re.match(r"^(\d+)-(\d+)$", s)
    if m:
        a, b = int(m.group(1)), int(m.group(2))
        if a <= b and b - a < 60:
            return set(range(a, b + 1))
    if re.match(r"^\d+$", s):
        return {int(s)}
    return set()


def _is_parking_hal(full_partial):
    return "parking" in str(full_partial or "").lower()


def _is_parking_pv(sec, row):
    return ("LOT" in sec) or ("PARKING" in row) or ("PARKING" in sec)


# --------------------------------------------------------------------------- #
# Parsing
# --------------------------------------------------------------------------- #

def parse_hal(rows, filename):
    """Parse the HAL-NFL season-ticket database. One record per seat block.
    Each record is tagged with its Company (from a 'Company' column if present,
    otherwise the file name)."""
    hidx, header = _find_header_row(rows)
    ci = {
        "company": _col_index(header, "Company", "Client", "Broker"),
        "email": _col_index(header, "Email"),
        "team": _col_index(header, "Team"),
        "fp": _col_index(header, "Full/Partial"),
        "section": _col_index(header, "Section"),
        "row": _col_index(header, "Row"),
        "seats": _col_index(header, "Seats"),
        "qty": _col_index(header, "Qty"),
        "games": _col_index(header, "Games/Threshold", "# Games", "Games"),
        "total": _col_index(header, "Total", "Total Cost"),
    }
    if ci["email"] is None or ci["team"] is None:
        raise ValueError(f"{os.path.basename(filename)}: HAL file is missing an "
                         f"'Email' or 'Team' column.")
    default_company = _company_from_filename(filename)
    out = []
    for row in rows[hidx + 1:]:
        team = _team(_cell(row, ci["team"]))
        emails = _emails(_cell(row, ci["email"]))
        if not team or not emails:
            continue
        company = str(_cell(row, ci["company"]) or "").strip() or default_company
        games = _amount(_cell(row, ci["games"]))
        out.append({
            "company": company,
            "emails": emails,
            "team": team,
            "Full/Partial": str(_cell(row, ci["fp"]) or "").strip(),
            "Section": str(_cell(row, ci["section"]) or "").strip(),
            "Row": str(_cell(row, ci["row"]) or "").strip(),
            "Seats": str(_cell(row, ci["seats"]) or "").strip(),
            "Qty": str(_cell(row, ci["qty"]) or "").strip(),
            "Email": str(_cell(row, ci["email"]) or "").strip(),
            "games": int(games) if games is not None else None,
            "total": _amount(_cell(row, ci["total"])) or 0.0,
            "sec_n": _sec(_cell(row, ci["section"])),
            "row_n": _row(_cell(row, ci["row"])),
            "is_parking": _is_parking_hal(_cell(row, ci["fp"])),
        })
    return out


def parse_details(rows, filename):
    """Parse a TicketVault Purchase Details export into an index keyed by
    (email, team) -> list of {sec, row, seatset, event, cost, is_parking}."""
    hidx, header = _find_header_row(rows)
    ci = {
        "email": _col_index(header, "PO Email Account", "Email"),
        "team": _col_index(header, "Team/Performer", "Team"),
        "sec": _col_index(header, "Sec", "Section"),
        "row": _col_index(header, "Row"),
        "seats": _col_index(header, "Seats"),
        "event": _col_index(header, "Event Date"),
        "cost": _col_index(header, "Total Cost"),
    }
    if ci["email"] is None or ci["team"] is None or ci["cost"] is None:
        raise ValueError(f"{os.path.basename(filename)}: Purchase Details file is "
                         f"missing 'PO Email Account', 'Team/Performer' or 'Total Cost'.")
    index = defaultdict(list)
    n = 0
    for row in rows[hidx + 1:]:
        email = str(_cell(row, ci["email"]) or "").strip().lower()
        team = _team(_cell(row, ci["team"]))
        if not email or not team:
            continue
        sec = _sec(_cell(row, ci["sec"]))
        rw = _row(_cell(row, ci["row"]))
        index[(email, team)].append({
            "sec": sec,
            "row": rw,
            "seatset": _seatnums(_cell(row, ci["seats"])),
            "event": str(_cell(row, ci["event"]) or ""),
            "cost": _amount(_cell(row, ci["cost"])) or 0.0,
            "is_parking": _is_parking_pv(sec, rw),
        })
        n += 1
    return index, n


# --------------------------------------------------------------------------- #
# Reconciliation
# --------------------------------------------------------------------------- #

def _games_split(matches):
    """Return (games_with_cost, games_without_cost, total_cost) for a set of
    matching vault rows. A game counts as 'with cost' if any matching row for
    that event carries a positive cost."""
    ev = defaultdict(float)
    total = 0.0
    for x in matches:
        ev[x["event"]] = max(ev[x["event"]], x["cost"])
        total += x["cost"]
    wc = sum(1 for v in ev.values() if v > 0)
    woc = sum(1 for v in ev.values() if v == 0)
    return wc, woc, round(total, 2)


def reconcile(hal_rows, index, tolerance):
    reconciled, not_reconciled = [], []
    for r in hal_rows:
        vr = []
        for e in r["emails"]:
            vr.extend(index.get((e, r["team"]), []))

        if not vr:
            matches = []
        elif r["is_parking"]:
            matches = [x for x in vr if x["is_parking"]]
        else:
            hs = _seatnums(r["Seats"])
            matches = [x for x in vr if (not x["is_parking"])
                       and x["sec"] == r["sec_n"] and x["row"] == r["row_n"]
                       and (x["seatset"] & hs)]

        wc, woc, tv_cost = _games_split(matches)

        base = {
            "Team": r["team"], "Email": r["Email"], "Full/Partial": r["Full/Partial"],
            "Section": r["Section"], "Row": r["Row"], "Seats": r["Seats"], "Qty": r["Qty"],
            "# Games": r["games"], "HAL Total Cost": round(r["total"], 2),
            "TV Total Cost": tv_cost, "# Games w/Cost": wc, "# Games w/o Cost": woc,
        }

        if not matches:
            not_reconciled.append({**base, "Notes": "Not bought in", "_p": 0})
            continue

        cost_ok = abs(tv_cost - r["total"]) <= tolerance
        games_ok = (r["games"] is not None and wc == r["games"])
        if cost_ok and games_ok:
            reconciled.append(base)
        else:
            parts = []
            if not cost_ok:
                parts.append("total cost not equal")
            if not games_ok:
                parts.append("# games not equal")
            not_reconciled.append({**base, "Notes": ", ".join(parts), "_p": 1})

    reconciled.sort(key=lambda x: (x["Team"], x["Email"]))
    not_reconciled.sort(key=lambda x: (x["_p"], -x["HAL Total Cost"]))
    return reconciled, not_reconciled


# --------------------------------------------------------------------------- #
# Workbook
# --------------------------------------------------------------------------- #

ARIAL = "Arial"
CUR = "$#,##0"
THIN = Side(style="thin", color="000000")
HAL_FILL = PatternFill("solid", fgColor=Color(theme=8, tint=0.7999816888943144))
TV_FILL = PatternFill("solid", fgColor=Color(theme=8, tint=0.0))
NAVY = "1F3864"
BLUE = "2E5496"
CENTER = Alignment(horizontal="center")

RECON_COLS = ["Team", "Email", "Full/Partial", "Section", "Row", "Seats", "Qty",
              "# Games", "Total Cost", "Total Cost", "# Games w/Cost", "# Games w/o Cost"]
RECON_SRC = ["Team", "Email", "Full/Partial", "Section", "Row", "Seats", "Qty",
             "# Games", "HAL Total Cost", "TV Total Cost", "# Games w/Cost", "# Games w/o Cost"]
RECON_W = [22.3, 46.3, 15.6, 12.4, 9.6, 10.6, 8.6, 13.4, 14.6, 13.0, 22.0, 13.0]

NR_COLS = ["Notes", "Team", "Email", "Full/Partial", "Section", "Row", "Seats", "Qty",
           "# Games", "Total Cost", "Total Cost", "# Games w/Cost", "# Games w/o Cost"]
NR_SRC = ["Notes", "Team", "Email", "Full/Partial", "Section", "Row", "Seats", "Qty",
          "# Games", "HAL Total Cost", "TV Total Cost", "# Games w/Cost", "# Games w/o Cost"]
NR_W = [35.7, 20.0, 28.6, 15.6, 12.4, 9.6, 10.6, 8.6, 13.4, 14.6, 13.0, 20.1, 22.0]


def _build_detail_tab(ws, headers, srcs, widths, rows, hal_last, tv_first, tv_last, cost_cols):
    ws.sheet_view.showGridLines = False
    for j, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(j)].width = w
    n = len(headers)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=hal_last)
    ws.merge_cells(start_row=1, start_column=tv_first, end_row=1, end_column=tv_last)
    a = ws.cell(1, 1, "per HAL"); a.font = Font(name=ARIAL, size=9, bold=True); a.alignment = CENTER; a.fill = HAL_FILL
    b = ws.cell(1, tv_first, "per TicketVault"); b.font = Font(name=ARIAL, size=9, bold=True); b.alignment = CENTER; b.fill = TV_FILL
    for c in range(1, hal_last + 1):
        ws.cell(1, c).fill = HAL_FILL
    for c in range(tv_first, tv_last + 1):
        ws.cell(1, c).fill = TV_FILL
    for j, h in enumerate(headers, 1):
        cell = ws.cell(2, j, h); cell.font = Font(name=ARIAL, size=10, bold=True); cell.alignment = CENTER
    for i, r in enumerate(rows, 3):
        for j, src in enumerate(srcs, 1):
            cell = ws.cell(i, j, r.get(src))
            cell.font = Font(name=ARIAL, size=9); cell.alignment = CENTER
            if j in cost_cols:
                cell.number_format = CUR
    end = 2 + len(rows)
    for rr in range(1, end + 1):
        ws.cell(rr, 1).border = Border(left=THIN, right=(THIN if rr == 1 else None))
        ws.cell(rr, hal_last).border = Border(right=THIN)
        ws.cell(rr, tv_first).border = Border(left=(THIN if rr == 1 else None),
                                              right=(THIN if rr == 1 else None))
        ws.cell(rr, tv_last).border = Border(right=THIN)
    ws.freeze_panes = "A3"
    ws.auto_filter.ref = f"A2:{get_column_letter(n)}{end}"


def build_workbook(company, league, year, reconciled, not_reconciled, hal_total, tolerance):
    rec_n = len(reconciled)
    nr_n = len(not_reconciled)
    nbi = sum(1 for r in not_reconciled if r["Notes"] == "Not bought in")
    part = nr_n - nbi

    wb = Workbook()

    # ---- Summary -------------------------------------------------------- #
    ws = wb.active
    ws.title = "Summary"
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 72.7
    ws.column_dimensions["B"].width = 12.0
    ws.merge_cells("A1:B1")
    t = ws.cell(1, 1, "Seasons Review")
    t.font = Font(name=ARIAL, size=15, bold=True, color=NAVY); t.alignment = CENTER

    def info(r, text, size, bold):
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
        c = ws.cell(r, 1, text); c.font = Font(name=ARIAL, size=size, bold=bold, color=NAVY); c.alignment = CENTER

    info(3, f"Company:  {company}", 11, True)
    info(4, f"League:  {league}", 10, False)
    info(5, f"Year:  {year}", 10, False)

    def bar(r, text):
        c = ws.cell(r, 1, text); c.font = Font(name=ARIAL, size=11, bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor=BLUE); c.alignment = CENTER
        ws.cell(r, 2).fill = PatternFill("solid", fgColor=BLUE)

    def hd(r, a, b):
        for col, val in ((1, a), (2, b)):
            c = ws.cell(r, col, val); c.font = Font(name=ARIAL, size=10, bold=True, color="FFFFFF")
            c.fill = PatternFill("solid", fgColor=NAVY); c.alignment = CENTER

    def line(r, a, b, bold=False):
        ca = ws.cell(r, 1, a); cb = ws.cell(r, 2, b)
        for c in (ca, cb):
            c.font = Font(name=ARIAL, size=10, bold=bold); c.alignment = CENTER

    bar(7, "RESULT"); hd(8, "Metric", "Count")
    line(9, "Reconciled", rec_n)
    line(10, "Not Reconciled", nr_n)
    line(11, "Total # HAL Records", hal_total, bold=True)
    bar(13, "NOT RECONCILED — BY REASON"); hd(14, "Reason", "Count")
    line(15, "Not bought in", nbi)
    line(16, "Partially bought in", part)
    line(17, "TOTAL", nr_n, bold=True)

    # ---- Reconciled ----------------------------------------------------- #
    ws_r = wb.create_sheet("Reconciled")
    _build_detail_tab(ws_r, RECON_COLS, RECON_SRC, RECON_W, reconciled,
                      hal_last=9, tv_first=10, tv_last=12, cost_cols={9, 10})

    # ---- Not Reconciled ------------------------------------------------- #
    ws_nr = wb.create_sheet("Not Reconciled")
    _build_detail_tab(ws_nr, NR_COLS, NR_SRC, NR_W, not_reconciled,
                      hal_last=10, tv_first=11, tv_last=13, cost_cols={10, 11})

    bio = io.BytesIO()
    wb.save(bio)
    bio.seek(0)
    return bio.read(), {"reconciled": rec_n, "not_reconciled": nr_n,
                        "not_bought_in": nbi, "partially": part, "clean": nr_n == 0}


# --------------------------------------------------------------------------- #
# Routes
# --------------------------------------------------------------------------- #

@app.route("/")
def index():
    return send_from_directory(BASE_DIR, "index.html")


@app.route("/options")
def options():
    return jsonify({"leagues": LEAGUES, "years": YEARS})


@app.route("/process", methods=["POST"])
def process():
    hal_files = [f for f in request.files.getlist("hal") if f.filename]
    details_files = [f for f in request.files.getlist("details") if f.filename]
    league = (request.form.get("league") or "").strip()
    year = (request.form.get("year") or "").strip()

    if league not in LEAGUES:
        return jsonify({"error": "Please choose a League."}), 400
    if year not in YEARS:
        return jsonify({"error": "Please choose a Year."}), 400
    if not hal_files:
        return jsonify({"error": "Please upload at least one HAL season ticket database."}), 400
    if not details_files:
        return jsonify({"error": "Please upload at least one TicketVault Purchase Details file."}), 400

    raw_tol = (request.form.get("tolerance") or "").strip()
    tolerance = DEFAULT_TOLERANCE
    if raw_tol:
        try:
            tolerance = abs(float(raw_tol))
        except ValueError:
            return jsonify({"error": "Tolerance must be a number (dollars)."}), 400

    warnings = []
    try:
        hal_rows = []
        for f in hal_files:
            hal_rows.extend(parse_hal(_rows_from_upload(f.filename, f.read()), f.filename))
        if not hal_rows:
            return jsonify({"error": "No season-ticket records found in the HAL file(s)."}), 400

        index = defaultdict(list)
        detail_rows = 0
        for f in details_files:
            part, n = parse_details(_rows_from_upload(f.filename, f.read()), f.filename)
            for k, v in part.items():
                index[k].extend(v)
            detail_rows += n
        if detail_rows == 0:
            return jsonify({"error": "No rows found in the Purchase Details file(s)."}), 400

        hal_teams = {r["team"] for r in hal_rows}
        pv_teams = {t for (_e, t) in index}
        missing_teams = sorted(hal_teams - pv_teams)
        if missing_teams:
            shown = ", ".join(missing_teams[:6]) + ("…" if len(missing_teams) > 6 else "")
            warnings.append(f"{len(missing_teams)} team(s) in HAL have no rows in any "
                            f"Purchase Details file — those records will show as "
                            f"“Not bought in”: {shown}")

        # split records by company, one workbook per company
        by_company = OrderedDict()
        for r in hal_rows:
            by_company.setdefault(r["company"], []).append(r)

        token = uuid.uuid4().hex
        folder = os.path.join(STORE_DIR, token)
        os.makedirs(folder, exist_ok=True)

        reports = []
        tot_rec = tot_nr = 0
        for company, rows in by_company.items():
            reconciled, not_reconciled = reconcile(rows, index, tolerance)
            data, m = build_workbook(company, league, year, reconciled,
                                     not_reconciled, len(rows), tolerance)
            fname = f"Seasons Review - {_safe_name(company)} - {_safe_name(league)} - {_safe_name(year)}.xlsx"
            with open(os.path.join(folder, fname), "wb") as fh:
                fh.write(data)
            tot_rec += m["reconciled"]
            tot_nr += m["not_reconciled"]
            reports.append({
                "company": company, "records": len(rows),
                "reconciled": m["reconciled"], "not_reconciled": m["not_reconciled"],
                "clean": m["clean"], "filename": fname,
                "download_url": f"/download/{token}/{fname}",
            })

        reports.sort(key=lambda x: x["company"].lower())

        # bundle all reports into a single zip when there's more than one company
        if len(reports) > 1:
            zip_name = f"Seasons Review - {_safe_name(league)} - {_safe_name(year)}.zip"
            with zipfile.ZipFile(os.path.join(folder, zip_name), "w",
                                 zipfile.ZIP_DEFLATED) as zf:
                for rep in reports:
                    zf.write(os.path.join(folder, rep["filename"]), rep["filename"])
            download_all_url = f"/download/{token}"
            download_all_name = zip_name
        else:
            download_all_url = reports[0]["download_url"]
            download_all_name = reports[0]["filename"]

        _cleanup_old()
    except Exception as exc:
        return jsonify({"error": f"{type(exc).__name__}: {exc}"}), 500

    return jsonify({
        "league": league, "year": year,
        "companies": len(reports),
        "total_records": len(hal_rows), "reconciled": tot_rec, "not_reconciled": tot_nr,
        "clean": tot_nr == 0,
        "reports": reports,
        "download_all_url": download_all_url, "download_all_name": download_all_name,
        "warnings": warnings,
    })


@app.route("/download/<token>")
def download_all(token):
    folder = os.path.join(STORE_DIR, os.path.basename(token))
    if not os.path.isdir(folder):
        abort(404)
    zips = [f for f in os.listdir(folder) if f.lower().endswith(".zip")]
    if zips:
        pick = zips[0]
        return send_file(os.path.join(folder, pick), mimetype="application/zip",
                         as_attachment=True, download_name=pick)
    xlsx = [f for f in os.listdir(folder) if f.lower().endswith(".xlsx")]
    if len(xlsx) == 1:
        pick = xlsx[0]
        return send_file(os.path.join(folder, pick),
                         mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                         as_attachment=True, download_name=pick)
    abort(404)


@app.route("/download/<token>/<path:name>")
def download_one(token, name):
    folder = os.path.join(STORE_DIR, os.path.basename(token))
    safe = os.path.basename(name)
    path = os.path.join(folder, safe)
    if not os.path.isfile(path):
        abort(404)
    return send_file(path,
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                     as_attachment=True, download_name=safe)


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)), debug=True)
