from __future__ import annotations

import os
import sqlite3
from collections import defaultdict
from datetime import date, datetime, timedelta
from io import BytesIO
from pathlib import Path

import holidays
from flask import (
    Flask,
    abort,
    flash,
    jsonify,
    redirect,
    render_template,
    request,
    send_file,
    url_for,
)
from flask_login import (
    LoginManager,
    UserMixin,
    current_user,
    login_required,
    login_user,
    logout_user,
)
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill
from werkzeug.security import check_password_hash, generate_password_hash
from werkzeug.utils import secure_filename

DB_PATH = Path(os.environ.get("URLAUBSPLANER_DB", "instance/urlaubsplaner.sqlite3"))
ROLE_ORDER = ["admin", "ausbilder", "putzchef", "azubi", "normal"]
ROLE_RANK = {role: index for index, role in enumerate(ROLE_ORDER)}
IGNORED_IMPORT_NAMES = {
    "monat",
    "kw",
    "ferienzeit",
    "geplant oder beantragt",
    "genehmigt",
    "feiertag",
    "grundkurs",
    "berufsschule",
    "ausbildungsmesse",
    "pv+ba+ap",
    "ausbildung",
    "weihnachtsputz",
    "kurzarbeit",
    "weiterbildung",
    "kein arbeitstag",
}
ENTRY_TYPES = {
    "UB": ("Urlaub geplant/beantragt", "FFFF00"),
    "UG": ("Genehmigter Urlaub", "90EE90"),
    "PV": ("PV+BA+AP", "B084CC"),
    "BA": ("PV+BA+AP", "B084CC"),
    "AP": ("PV+BA+AP", "B084CC"),
    "AZ": ("Ausbildung", "006100"),
    "KA": ("Kurzarbeit", "FF0000"),
    "WB": ("Weiterbildung", "FFA500"),
    "AM": ("Ausbildungsmesse", "FF69B4"),
    "BS": ("Berufsschule", "40E0D0"),
    "GK": ("Grundkurs", "C4A484"),
    "WP": ("Weihnachtsputz", "654321"),
    "KAT": ("Kein Arbeitstag", "000000"),
    "KR": ("Krank", "90EE90"),
}
IMPORT_VALUE_TO_CODE = {
    "UrlbGplntOdrBntrgt": "UB",
    "UrlbGnhmgt": "UG",
    "Frtg": None,
    "Grndkrs": "GK",
    "Brfsschl": "BS",
    "Asbldngsmss": "AM",
    "PvBaAp": "PV",
    "Asbldng": "AZ",
    "Whnchtsptz": "WP",
    "Krzrbt": "KA",
    "Wtrbldng": "WB",
    "KnArbtstg": "KAT",
}
IMPORT_VALUE_TO_CODE_NORMALIZED = {
    key.casefold(): value for key, value in IMPORT_VALUE_TO_CODE.items()
}
SPECIAL_KR = "00FE43"
HOLIDAY_COLOR = "ADD8E6"
VACATION_COLOR = "BDD7EE"
ALLOWED_BY_ROLE = {
    "normal": {"UB", "UG"},
    "azubi": {"UB", "UG", "BS"},
    "putzchef": {"UB", "UG", "WP"},
    "ausbilder": {"UB", "UG", "AM", "BS", "GK", "WP"},
    "admin": set(ENTRY_TYPES),
}
GERMAN_WEEKDAYS = ["Mo", "Di", "Mi", "Do", "Fr"]

app = Flask(__name__)
app.config["SECRET_KEY"] = os.environ.get("SECRET_KEY", "dev-change-me")
app.config["UPLOAD_FOLDER"] = "instance/uploads"
login_manager = LoginManager(app)
login_manager.login_view = "login"


class User(UserMixin):
    def __init__(self, row: sqlite3.Row, roles: list[str]):
        self.id = str(row["id"])
        self.first_name = row["first_name"]
        self.last_name = row["last_name"]
        self.username = row["username"]
        self.password_hash = row["password_hash"]
        self.roles = roles

    def has_role(self, role: str) -> bool:
        return role in self.roles

    @property
    def display_name(self) -> str:
        return f"{self.first_name} {self.last_name}"


def db() -> sqlite3.Connection:
    DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db() -> None:
    with db() as conn:
        conn.executescript(
            """
            CREATE TABLE IF NOT EXISTS users (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                username TEXT NOT NULL UNIQUE,
                first_name TEXT NOT NULL,
                last_name TEXT NOT NULL,
                password_hash TEXT NOT NULL
            );
            CREATE TABLE IF NOT EXISTS user_roles (
                user_id INTEGER NOT NULL,
                role TEXT NOT NULL,
                PRIMARY KEY (user_id, role),
                FOREIGN KEY (user_id) REFERENCES users(id) ON DELETE CASCADE
            );
            CREATE TABLE IF NOT EXISTS entries (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                user_id INTEGER NOT NULL,
                entry_date TEXT NOT NULL,
                code TEXT NOT NULL,
                created_by INTEGER,
                UNIQUE(user_id, entry_date, code),
                FOREIGN KEY (user_id) REFERENCES users(id) ON DELETE CASCADE
            );
            """
        )
        if conn.execute("SELECT COUNT(*) FROM users").fetchone()[0] == 0:
            user_id = create_user(
                conn, "admin", "Admin", "Benutzer", "admin", ["admin"]
            )
            conn.execute(
                "INSERT OR IGNORE INTO entries(user_id, entry_date, code, created_by) VALUES (?, ?, ?, ?)",
                (user_id, date.today().isoformat(), "UG", user_id),
            )


def create_user(
    conn, username: str, first: str, last: str, password: str, roles: list[str]
) -> int:
    cur = conn.execute(
        "INSERT INTO users(username, first_name, last_name, password_hash) VALUES (?, ?, ?, ?)",
        (username, first, last, generate_password_hash(password)),
    )
    user_id = cur.lastrowid
    for role in normalized_roles(roles):
        conn.execute(
            "INSERT INTO user_roles(user_id, role) VALUES (?, ?)", (user_id, role)
        )
    return user_id


def normalized_roles(roles: list[str]) -> list[str]:
    valid = [role for role in ROLE_ORDER if role in roles]
    return valid or ["normal"]


@login_manager.user_loader
def load_user(user_id: str) -> User | None:
    with db() as conn:
        row = conn.execute("SELECT * FROM users WHERE id = ?", (user_id,)).fetchone()
        if not row:
            return None
        roles = [
            r["role"]
            for r in conn.execute(
                "SELECT role FROM user_roles WHERE user_id = ?", (user_id,)
            )
        ]
        return User(row, roles)


def weekdays_between(start: date, end: date) -> list[date]:
    days = []
    current = start
    while current <= end:
        if current.weekday() < 5:
            days.append(current)
        current += timedelta(days=1)
    return days


def band_range(user: User, year: int | None = None) -> tuple[date, date]:
    today = date.today()
    base_year = year or today.year
    if user.has_role("admin") or user.has_role("ausbilder"):
        return date(base_year - 1, 7, 2), date(base_year + 1, 7, 2)
    return date(base_year, 1, 1), date(base_year, 12, 31)


def current_week_monday() -> date:
    today = date.today()
    return today - timedelta(days=today.weekday())


def get_users(conn) -> list[sqlite3.Row]:
    rows = conn.execute(
        """
        SELECT u.*, COALESCE(MIN(CASE ur.role
            WHEN 'admin' THEN 0 WHEN 'ausbilder' THEN 1 WHEN 'putzchef' THEN 2
            WHEN 'azubi' THEN 3 ELSE 4 END), 4) AS rank
        FROM users u LEFT JOIN user_roles ur ON u.id = ur.user_id
        GROUP BY u.id ORDER BY rank, u.last_name COLLATE NOCASE, u.first_name COLLATE NOCASE
        """
    ).fetchall()
    return rows


def roles_for_users(conn) -> dict[int, list[str]]:
    result = defaultdict(list)
    for row in conn.execute("SELECT user_id, role FROM user_roles ORDER BY user_id"):
        result[row["user_id"]].append(row["role"])
    return result


def visible_code(
    viewer: User, target_user_id: int, target_roles: list[str], code: str
) -> tuple[str, str]:
    color = ENTRY_TYPES[code][1]
    if code == "KR":
        if (
            viewer.has_role("admin")
            or target_user_id == int(viewer.id)
            or (viewer.has_role("ausbilder") and "azubi" in target_roles)
        ):
            color = SPECIAL_KR
    return code if viewer.has_role("admin") else "", color


def month_spans(days: list[date]) -> list[dict[str, int | str]]:
    spans = []
    previous = None
    for day in days:
        key = (day.year, day.month)
        if key != previous:
            spans.append({"label": day.strftime("%b"), "span": 1})
            previous = key
        else:
            spans[-1]["span"] += 1
    return spans


def year_spans(days: list[date]) -> list[dict[str, int | str]]:
    spans = []
    previous = None
    for day in days:
        if day.year != previous:
            spans.append({"label": str(day.year), "span": 1})
            previous = day.year
        else:
            spans[-1]["span"] += 1
    return spans


def cell_style_and_tooltip(
    viewer: User,
    target_user_id: int,
    target_roles: list[str],
    day: date,
    codes: list[str],
) -> tuple[str, str]:
    if not codes:
        return "", ""
    code = codes[-1]
    label, default_color = ENTRY_TYPES[code]
    _, color = visible_code(viewer, target_user_id, target_roles, code)
    if viewer.has_role("admin"):
        entry_text = f"{code} – {label}"
    elif code == "KR" and color != SPECIAL_KR:
        entry_text = "Abwesend"
    else:
        entry_text = label
    return default_color if code != "KR" else color, f"{day:%d.%m.%Y}: {entry_text}"


def school_vacations_bavaria(years: set[int]) -> dict[str, str]:
    # Ferien nach bayerischem Ferienkalender; feste Berechnung für Schuljahre 2025-2027.
    ranges = {
        2025: [
            ("2025-03-03", "2025-03-07"),
            ("2025-04-14", "2025-04-25"),
            ("2025-06-10", "2025-06-20"),
            ("2025-08-01", "2025-09-15"),
            ("2025-11-03", "2025-11-07"),
            ("2025-12-22", "2026-01-05"),
        ],
        2026: [
            ("2025-12-22", "2026-01-05"),
            ("2026-02-16", "2026-02-20"),
            ("2026-03-30", "2026-04-10"),
            ("2026-05-26", "2026-06-05"),
            ("2026-08-03", "2026-09-14"),
            ("2026-11-02", "2026-11-06"),
            ("2026-12-24", "2027-01-08"),
        ],
        2027: [
            ("2026-12-24", "2027-01-08"),
            ("2027-02-08", "2027-02-12"),
            ("2027-03-22", "2027-04-02"),
            ("2027-05-18", "2027-05-28"),
            ("2027-08-02", "2027-09-13"),
            ("2027-11-02", "2027-11-05"),
            ("2027-12-24", "2028-01-07"),
        ],
    }
    result = {}
    for year in years:
        for start_s, end_s in ranges.get(year, []):
            for day in weekdays_between(
                date.fromisoformat(start_s), date.fromisoformat(end_s)
            ):
                result[day.isoformat()] = "Ferien"
    return result


def bavarian_holidays(years: set[int]) -> dict[str, str]:
    by = holidays.Germany(subdiv="BY", years=years, language="de")
    return {d.isoformat(): name for d, name in by.items() if d.weekday() < 5}


def allowed_codes(user: User) -> list[str]:
    codes = set()
    for role in user.roles:
        codes |= ALLOWED_BY_ROLE.get(role, set())
    return [code for code in ENTRY_TYPES if code in codes]


def may_edit(user: User, target_id: int, code: str) -> bool:
    return user.has_role("admin") or (
        target_id == int(user.id) and code in allowed_codes(user)
    )


@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        with db() as conn:
            row = conn.execute(
                "SELECT * FROM users WHERE username = ?", (request.form["username"],)
            ).fetchone()
        if row and check_password_hash(row["password_hash"], request.form["password"]):
            login_user(load_user(str(row["id"])))
            return redirect(url_for("index"))
        flash("Ungültige Zugangsdaten.")
    return render_template("login.html")


@app.route("/logout")
@login_required
def logout():
    logout_user()
    return redirect(url_for("login"))


@app.route("/")
@login_required
def index():
    year = int(request.args.get("year", date.today().year))
    start, end = band_range(current_user, year)
    days = weekdays_between(start, end)
    today_monday = current_week_monday().isoformat()
    with db() as conn:
        users = get_users(conn)
        role_map = roles_for_users(conn)
        entries = conn.execute(
            "SELECT * FROM entries WHERE entry_date BETWEEN ? AND ?",
            (start.isoformat(), end.isoformat()),
        ).fetchall()
    entry_map = defaultdict(list)
    for e in entries:
        entry_map[(e["user_id"], e["entry_date"])].append(e["code"])
    years = {d.year for d in days}
    return render_template(
        "index.html",
        users=users,
        role_map=role_map,
        days=days,
        entries=entry_map,
        entry_types=ENTRY_TYPES,
        holidays=bavarian_holidays(years),
        vacations=school_vacations_bavaria(years),
        holiday_color=HOLIDAY_COLOR,
        vacation_color=VACATION_COLOR,
        allowed_codes=allowed_codes(current_user),
        today_monday=today_monday,
        visible_code=visible_code,
        cell_style_and_tooltip=cell_style_and_tooltip,
        month_spans=month_spans(days),
        year_spans=year_spans(days),
        year=year,
        german_weekdays=GERMAN_WEEKDAYS,
    )


@app.post("/entry")
@login_required
def entry():
    target_id = int(request.form["user_id"])
    entry_date = request.form["date"]
    code = request.form["code"]
    if code not in ENTRY_TYPES or not may_edit(current_user, target_id, code):
        abort(403)
    with db() as conn:
        if request.form.get("action") == "delete":
            conn.execute(
                "DELETE FROM entries WHERE user_id = ? AND entry_date = ? AND code = ?",
                (target_id, entry_date, code),
            )
        else:
            conn.execute(
                "INSERT OR IGNORE INTO entries(user_id, entry_date, code, created_by) VALUES (?, ?, ?, ?)",
                (target_id, entry_date, code, current_user.id),
            )
    return redirect(request.referrer or url_for("index"))


@app.post("/bulk-entry")
@login_required
def bulk_entry():
    payload = request.get_json(silent=True) or {}
    cells = payload.get("cells", [])
    code = payload.get("code")
    delete = payload.get("delete", False)
    if not cells:
        return jsonify({"updated": 0})
    if not delete and code not in ENTRY_TYPES:
        abort(400)
    updated = 0
    with db() as conn:
        for cell in cells:
            target_id = int(cell["user_id"])
            entry_date = cell["date"]
            if delete:
                removable_codes = (
                    ENTRY_TYPES
                    if current_user.has_role("admin")
                    else allowed_codes(current_user)
                )
                for removable_code in removable_codes:
                    if may_edit(current_user, target_id, removable_code):
                        cur = conn.execute(
                            "DELETE FROM entries WHERE user_id = ? AND entry_date = ? AND code = ?",
                            (target_id, entry_date, removable_code),
                        )
                        updated += cur.rowcount
            elif may_edit(current_user, target_id, code):
                conn.execute(
                    "DELETE FROM entries WHERE user_id = ? AND entry_date = ?",
                    (target_id, entry_date),
                )
                conn.execute(
                    "INSERT OR IGNORE INTO entries(user_id, entry_date, code, created_by) VALUES (?, ?, ?, ?)",
                    (target_id, entry_date, code, current_user.id),
                )
                updated += 1
    return jsonify({"updated": updated})


@app.route("/members", methods=["GET", "POST"])
@login_required
def members():
    if not current_user.has_role("admin"):
        abort(403)
    with db() as conn:
        if request.method == "POST":
            action = request.form["action"]
            if action == "create":
                create_user(
                    conn,
                    request.form["username"],
                    request.form["first_name"],
                    request.form["last_name"],
                    request.form["password"],
                    request.form.getlist("roles"),
                )
            elif action == "delete":
                conn.execute(
                    "DELETE FROM users WHERE id = ?", (request.form["user_id"],)
                )
            elif action == "roles":
                conn.execute(
                    "DELETE FROM user_roles WHERE user_id = ?",
                    (request.form["user_id"],),
                )
                for role in normalized_roles(request.form.getlist("roles")):
                    conn.execute(
                        "INSERT INTO user_roles(user_id, role) VALUES (?, ?)",
                        (request.form["user_id"], role),
                    )
            return redirect(url_for("members"))
        users = get_users(conn)
        role_map = roles_for_users(conn)
    return render_template(
        "members.html", users=users, role_map=role_map, roles=ROLE_ORDER
    )


@app.route("/years")
@login_required
def years():
    with db() as conn:
        rows = conn.execute(
            "SELECT DISTINCT substr(entry_date,1,4) AS year FROM entries WHERE user_id = ? ORDER BY year",
            (current_user.id,),
        ).fetchall()
    years_list = [r["year"] for r in rows] or [str(date.today().year)]
    return render_template("years.html", years=years_list)


def fill_for(viewer: User, target_id: int, roles: list[str], code: str) -> str:
    return visible_code(viewer, target_id, roles, code)[1]


@app.post("/download")
@login_required
def download():
    start_year, end_year = sorted(
        [int(request.form["start_year"]), int(request.form["end_year"])]
    )
    days = weekdays_between(date(start_year, 1, 1), date(end_year, 12, 31))
    with db() as conn:
        users = get_users(conn)
        role_map = roles_for_users(conn)
        entries = conn.execute(
            "SELECT * FROM entries WHERE entry_date BETWEEN ? AND ?",
            (days[0].isoformat(), days[-1].isoformat()),
        ).fetchall()
    entry_map = defaultdict(list)
    active_user_ids = {e["user_id"] for e in entries}
    for e in entries:
        entry_map[(e["user_id"], e["entry_date"])].append(e["code"])
    wb = Workbook()
    ws = wb.active
    ws.title = "Urlaubsübersicht"
    ws.cell(1, 1, "Name")
    for col, day in enumerate(days, 2):
        ws.cell(1, col, f"{GERMAN_WEEKDAYS[day.weekday()]} {day:%d.%m.%Y}")
    years = {d.year for d in days}
    vacations = school_vacations_bavaria(years)
    holidays_map = bavarian_holidays(years)
    ws.cell(2, 1, "Ferien")
    for col, day in enumerate(days, 2):
        if day.isoformat() in vacations:
            ws.cell(2, col, vacations[day.isoformat()]).fill = PatternFill(
                "solid", fgColor=VACATION_COLOR
            )
        elif day.isoformat() in holidays_map:
            ws.cell(2, col, holidays_map[day.isoformat()]).fill = PatternFill(
                "solid", fgColor=HOLIDAY_COLOR
            )
    row_no = 3
    for user in users:
        if user["id"] not in active_user_ids:
            continue
        ws.cell(row_no, 1, f"{user['first_name']} {user['last_name']}")
        for col, day in enumerate(days, 2):
            for code in entry_map.get((user["id"], day.isoformat()), []):
                ws.cell(row_no, col, code if current_user.has_role("admin") else "")
                ws.cell(row_no, col).fill = PatternFill(
                    "solid",
                    fgColor=fill_for(
                        current_user, user["id"], role_map[user["id"]], code
                    ),
                )
                break
        row_no += 1
    legend = row_no + 2
    ws.cell(legend, 1, "Legende")
    for idx, (code, (label, color)) in enumerate(ENTRY_TYPES.items(), legend + 1):
        ws.cell(idx, 1, code)
        ws.cell(idx, 2, label)
        ws.cell(idx, 3).fill = PatternFill(
            "solid",
            fgColor=SPECIAL_KR
            if code == "KR" and current_user.has_role("admin")
            else color,
        )
    stream = BytesIO()
    wb.save(stream)
    stream.seek(0)
    return send_file(
        stream,
        as_attachment=True,
        download_name=f"urlaubsuebersicht_{start_year}_{end_year}.xlsx",
    )


@app.route("/upload", methods=["GET", "POST"])
@login_required
def upload():
    if not current_user.has_role("admin"):
        abort(403)
    if request.method == "POST":
        file = request.files["file"]
        filename = secure_filename(file.filename)
        Path(app.config["UPLOAD_FOLDER"]).mkdir(parents=True, exist_ok=True)
        path = Path(app.config["UPLOAD_FOLDER"]) / filename
        file.save(path)
        import_excel(path)
        flash("Excel-Datei wurde importiert.")
        return redirect(url_for("index"))
    return render_template("upload.html")


def parse_excel_date(value) -> date | None:
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, str):
        text = value.strip()
        for fmt in ("%d.%m.%Y", "%Y-%m-%d"):
            try:
                return datetime.strptime(text[-10:], fmt).date()
            except ValueError:
                pass
    return None


def find_excel_date_columns(ws) -> tuple[int, dict[int, date]]:
    best_row = 1
    best_dates: dict[int, date] = {}
    for row in ws.iter_rows():
        dates = {}
        for cell in row[1:]:
            parsed = parse_excel_date(cell.value)
            if parsed:
                dates[cell.column] = parsed
        if len(dates) > len(best_dates):
            best_row = row[0].row
            best_dates = dates
    return best_row, best_dates


def import_excel(path: Path) -> None:
    wb = load_workbook(path, data_only=True)
    ws = wb.active
    date_row, date_columns = find_excel_date_columns(ws)
    actor_id = getattr(current_user, "id", None)
    with db() as conn:
        for row in ws.iter_rows(min_row=date_row + 1):
            row_label = str(row[0].value).strip() if row[0].value else ""
            if not row_label or row_label.lower() in IGNORED_IMPORT_NAMES:
                continue
            parts = row_label.split()
            first, last = (parts[0], " ".join(parts[1:]) or parts[0])
            username = f"{first}.{last}".lower().replace(" ", ".")
            user = conn.execute(
                "SELECT id FROM users WHERE username = ?", (username,)
            ).fetchone()
            user_id = (
                user["id"]
                if user
                else create_user(conn, username, first, last, "changeme", ["normal"])
            )
            for cell in row[1:]:
                entry_date = date_columns.get(cell.column)
                raw_value = str(cell.value).strip() if cell.value else ""
                code = IMPORT_VALUE_TO_CODE_NORMALIZED.get(raw_value.casefold())
                if code and entry_date:
                    conn.execute(
                        "INSERT OR IGNORE INTO entries(user_id, entry_date, code, created_by) VALUES (?, ?, ?, ?)",
                        (
                            user_id,
                            entry_date.isoformat(),
                            code,
                            actor_id,
                        ),
                    )


@app.context_processor
def inject_globals():
    return {"role_order": ROLE_ORDER, "entry_types": ENTRY_TYPES}


@app.before_request
def ensure_database():
    init_db()


if __name__ == "__main__":
    init_db()
    app.run(host="0.0.0.0", port=4010, debug=True)
