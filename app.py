"""
=====================================================================
 STUDENT ATTENDANCE AUTOMATION  -  DIT Internship 2026, Task 07
 Focus: Education Analytics
=====================================================================
 A complete attendance analytics application in ONE file.

 Run it with:   streamlit run app.py

 FILE LAYOUT (read top to bottom):
   1. Imports and constants
   2. Logging setup
   3. Sample data generator
   4. Data cleaning and validation
   5. Attendance analytics (the maths)
   6. Excel report builders (OpenPyXL)
   7. Test suite (12 test cases)
   8. Streamlit user interface (pages)
   9. main()
=====================================================================
"""

# ---------------------------------------------------------------------
# 1. IMPORTS AND CONSTANTS
# ---------------------------------------------------------------------
import io
import logging
import random
import re
import base64
import pathlib
from datetime import datetime, timedelta

import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st
from openpyxl import Workbook
from openpyxl.formatting.rule import CellIsRule, ColorScaleRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

APP_TITLE = "Student Attendance Automation"
LOG_FILE = "attendance_app.log"

STUDENT_COLS = ["student_id", "student_name", "department", "semester", "section"]
ATT_COLS = ["student_id", "attendance_date", "subject", "attendance_status"]
STATUSES = ["Present", "Absent", "OD", "Holiday"]

# Three ways to treat OD (On Duty) - the user chooses in the sidebar.
OD_ATTENDED = "Count OD as attended"
OD_NOT_ATTENDED = "Count OD as not attended"
OD_EXCLUDED = "Exclude OD from eligible sessions"
OD_MODES = [OD_ATTENDED, OD_NOT_ATTENDED, OD_EXCLUDED]
DEFAULT_THRESHOLD = 75.0

# Alternative column names we quietly accept (mapped to the official names).
COLUMN_ALIASES = {
    "studentid": "student_id", "id": "student_id", "roll_no": "student_id",
    "rollno": "student_id", "reg_no": "student_id",
    "name": "student_name", "studentname": "student_name",
    "dept": "department", "sem": "semester", "sec": "section",
    "date": "attendance_date", "attendancedate": "attendance_date",
    "status": "attendance_status", "attendancestatus": "attendance_status",
    "subject_name": "subject", "course": "subject",
}

# Every spelling of a status we understand (lower-case keys).
STATUS_MAP = {
    "present": "Present", "p": "Present", "pr": "Present", "yes": "Present",
    "y": "Present", "1": "Present", "attended": "Present",
    "absent": "Absent", "a": "Absent", "ab": "Absent", "no": "Absent",
    "n": "Absent", "0": "Absent",
    "od": "OD", "o.d": "OD", "o.d.": "OD", "on duty": "OD", "on-duty": "OD",
    "onduty": "OD",
    "holiday": "Holiday", "h": "Holiday", "hol": "Holiday",
}

# Date layouts we try, in order. Day-first (Indian style) is assumed for
# ambiguous dates such as 03/04/2026 (= 3 April 2026).
DATE_FORMATS = [
    "%Y-%m-%d", "%d-%m-%Y", "%d/%m/%Y", "%Y/%m/%d", "%d.%m.%Y",
    "%d %b %Y", "%d-%b-%Y", "%d %B %Y", "%b %d, %Y",
    "%Y-%m-%d %H:%M:%S", "%d/%m/%Y %H:%M:%S", "%d-%m-%Y %H:%M:%S",
    "%Y-%m-%dT%H:%M:%S",
]

ISSUE_COLS = ["Source", "Row", "Student ID", "Field", "Issue Type",
              "Original Value", "Action"]

DISPLAY = {
    "student_id": "Student ID", "student_name": "Student Name",
    "department": "Department", "semester": "Semester", "section": "Section",
    "od_count": "OD Marked", "holiday_count": "Holidays",
    "eligible_sessions": "Eligible Sessions",
    "attended_sessions": "Attended Sessions",
    "absent_sessions": "Absent Sessions", "attendance_pct": "Attendance %",
    "status": "Status", "shortage_pct": "Shortage %",
    "required_threshold": "Required %", "students": "Students",
    "shortage_students": "Below Threshold",
    "sufficient_students": "Sufficient", "no_data_students": "No Data",
    "subject": "Subject", "month": "Month",
}

STUDENT_VIEW_COLS = [
    "student_id", "student_name", "department", "semester", "section",
    "eligible_sessions", "attended_sessions", "absent_sessions",
    "od_count", "holiday_count", "attendance_pct", "status",
]

SUMMARY_LABELS = [
    ("students_raw", "Student master rows received"),
    ("students_clean", "Students after cleaning"),
    ("attendance_raw", "Attendance rows received"),
    ("attendance_final", "Attendance rows accepted (used in calculations)"),
    ("rejected", "Attendance rows rejected (listed below, not hidden)"),
    ("duplicates", "Duplicate attendance records removed"),
    ("ids_standardized", "Student IDs standardized"),
    ("dates_standardized", "Dates converted to YYYY-MM-DD"),
    ("statuses_normalized", "Status values normalized"),
    ("issues_logged", "Total issues logged"),
]

COLOR_GOOD, COLOR_BAD, COLOR_WARN, COLOR_BLUE = "#16a34a", "#dc2626", "#f59e0b", "#2563eb"
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

log = logging.getLogger("attendance_app")


def add_logo(path="logo.png", height=40):
    try:
        data = base64.b64encode(pathlib.Path(path).read_bytes()).decode()
    except FileNotFoundError:
        return  # no logo file found: site loads normally
    st.markdown(f"""
    <style>
    header[data-testid="stHeader"]::before {{
        content: "";
        position: absolute;
        left: 1rem;
        top: 50%;
        transform: translateY(-50%);
        height: {height}px;
        width: {int(height * 1.08)}px;
        background: url("data:image/png;base64,{data}") no-repeat left center / contain;
        pointer-events: none;
    }}
    @media (max-width: 640px) {{
        header[data-testid="stHeader"]::before {{
            height: {int(height * 0.8)}px;
            width: {int(height * 0.8 * 1.08)}px;
        }}
    }}
    </style>
    """, unsafe_allow_html=True)


# ---------------------------------------------------------------------
# 2. LOGGING (activity log)
# ---------------------------------------------------------------------
class MemoryLogHandler(logging.Handler):
    """Keeps recent log lines in memory so the app can show/download them."""

    def __init__(self, capacity=3000):
        super().__init__()
        self.records = []
        self.capacity = capacity

    def emit(self, record):
        self.records.append(self.format(record))
        if len(self.records) > self.capacity:
            del self.records[: len(self.records) - self.capacity]


@st.cache_resource
def get_log_handler():
    """Create the logger once per server process (cache_resource)."""
    formatter = logging.Formatter(
        "%(asctime)s | %(levelname)-8s | %(message)s", "%Y-%m-%d %H:%M:%S")
    handler = MemoryLogHandler()
    handler.setFormatter(formatter)
    log.setLevel(logging.INFO)
    log.addHandler(handler)
    try:  # also write a log file next to app.py (skipped if not writable)
        file_handler = logging.FileHandler(LOG_FILE, encoding="utf-8")
        file_handler.setFormatter(formatter)
        log.addHandler(file_handler)
    except OSError:
        pass
    log.info("Application startup")
    return handler


# ---------------------------------------------------------------------
# 3. SAMPLE DATA GENERATOR (30 fictional students, 30 working days)
# ---------------------------------------------------------------------
FIRST_NAMES = [
    "Aarav", "Priya", "Karthik", "Meena", "Rahul", "Divya", "Arjun", "Sneha",
    "Vikram", "Lakshmi", "Naveen", "Anitha", "Suresh", "Kavya", "Harish",
    "Deepa", "Manoj", "Swathi", "Ganesh", "Pooja", "Ramesh", "Nithya",
    "Sanjay", "Revathi", "Dinesh", "Janani", "Prakash", "Yamini", "Bala",
    "Ishita"]
LAST_NAMES = ["Kumar", "Raj", "Sharma", "Iyer", "Nair", "Pillai", "Reddy",
              "Menon", "Das", "Varma"]
DEPARTMENTS = ["CSE", "ECE", "MECH", "IT", "CIVIL"]
SUBJECTS = ["Data Structures", "Mathematics", "Physics",
            "Python Programming", "Digital Electronics"]
SHORTAGE_INDEXES = {3, 7, 9, 12, 16, 21, 25, 28}   # students with low attendance


def generate_sample_data(messy=True, seed=42):
    """Return (students_df, attendance_df). All fictional.

    messy=True injects deliberate problems (bad IDs, bad dates, duplicates...)
    so the validation features can be demonstrated live.
    """
    rng = random.Random(seed)

    students = []
    for i in range(30):
        students.append({
            "student_id": f"STU{i + 1:03d}",
            "student_name": f"{FIRST_NAMES[i]} {LAST_NAMES[(i * 3) % 10]}",
            "department": DEPARTMENTS[i % 5],
            "semester": "3" if i < 15 else "5",
            "section": "A" if i % 10 < 5 else "B",
        })

    # 30 working days (Mon-Fri) starting 14 Sep 2026.
    days, cursor = [], datetime(2026, 9, 14)
    while len(days) < 30:
        if cursor.weekday() < 5:
            days.append(cursor)
        cursor += timedelta(days=1)
    holiday_day = days[14]            # 2 Oct 2026 (Gandhi Jayanti)

    # Each student's chance of being present.
    chance = {}
    for i in range(30):
        if i in SHORTAGE_INDEXES:
            chance[i] = rng.uniform(0.52, 0.72)
        elif i % 6 == 0:
            chance[i] = rng.uniform(0.76, 0.80)   # borderline but safe
        else:
            chance[i] = rng.uniform(0.85, 0.98)

    att = []
    for d_idx, day in enumerate(days):
        iso = day.strftime("%Y-%m-%d")
        for i, stu in enumerate(students):
            if day == holiday_day:
                att.append({"student_id": stu["student_id"], "attendance_date": iso,
                            "subject": "All Subjects", "attendance_status": "Holiday"})
                continue
            for k in range(3):                       # 3 subjects per day
                subject = SUBJECTS[(d_idx * 3 + k) % 5]
                roll = rng.random()
                if roll < 0.04:
                    status = "OD"
                elif rng.random() < chance[i]:
                    status = "Present"
                else:
                    status = "Absent"
                att.append({"student_id": stu["student_id"], "attendance_date": iso,
                            "subject": subject, "attendance_status": status})

    if messy:
        n = len(att)
        idx = rng.sample(range(n), 120)
        for i in idx[:12]:                           # lower-case IDs
            att[i]["student_id"] = att[i]["student_id"].lower()
        for i in idx[12:20]:                         # STU-005 style IDs
            sid = att[i]["student_id"]
            att[i]["student_id"] = sid[:3] + "-" + sid[3:]
        for i in idx[20:25]:                         # IDs with spaces
            att[i]["student_id"] = f" {att[i]['student_id']} "
        for i in idx[25:65]:                         # dd/mm/yyyy dates
            d = datetime.strptime(att[i]["attendance_date"], "%Y-%m-%d")
            att[i]["attendance_date"] = d.strftime("%d/%m/%Y")
        for i in idx[65:90]:                         # messy status spellings
            options = {"Present": ["present", "P", "PRESENT"],
                       "Absent": ["absent", "A", "ABSENT"],
                       "OD": ["od", "On Duty"], "Holiday": ["holiday"]}
            att[i]["attendance_status"] = rng.choice(options[att[i]["attendance_status"]])
        extras = [dict(att[i]) for i in idx[90:96]]  # exact duplicates
        conflict = dict(att[idx[96]])                # same slot, different status
        conflict["attendance_status"] = (
            "Absent" if conflict["attendance_status"] == "Present" else "Present")
        extras.append(conflict)
        iso = lambda k: days[k].strftime("%Y-%m-%d")
        subj = SUBJECTS[0]
        extras += [
            {"student_id": "STU003", "attendance_date": "2026-13-45", "subject": subj, "attendance_status": "Present"},
            {"student_id": "STU004", "attendance_date": "31/02/2026", "subject": subj, "attendance_status": "Absent"},
            {"student_id": "STU005", "attendance_date": "not a date", "subject": subj, "attendance_status": "Present"},
            {"student_id": "STU006", "attendance_date": iso(2), "subject": subj, "attendance_status": "Maybe"},
            {"student_id": "STU007", "attendance_date": iso(3), "subject": subj, "attendance_status": "Late"},
            {"student_id": "STU999", "attendance_date": iso(4), "subject": subj, "attendance_status": "Present"},
            {"student_id": "STU098", "attendance_date": iso(5), "subject": subj, "attendance_status": "Absent"},
            {"student_id": "XYZ001", "attendance_date": iso(6), "subject": subj, "attendance_status": "Present"},
            {"student_id": "", "attendance_date": iso(7), "subject": subj, "attendance_status": "Present"},
            {"student_id": "", "attendance_date": iso(8), "subject": subj, "attendance_status": "Absent"},
            {"student_id": "STU010", "attendance_date": iso(5), "subject": "", "attendance_status": "Present"},
        ]
        att.extend(extras)
        students[10]["student_id"] = " stu011 "      # messy master ID
        students.append(dict(students[4]))           # duplicate master row

    return pd.DataFrame(students, dtype=str), pd.DataFrame(att, dtype=str)


# ---------------------------------------------------------------------
# 4. DATA CLEANING AND VALIDATION
# ---------------------------------------------------------------------
def is_blank(value):
    """True for None, NaN, empty text, 'nan', 'none', 'null'..."""
    if value is None:
        return True
    if isinstance(value, float) and np.isnan(value):
        return True
    return str(value).strip().lower() in {"", "nan", "none", "null", "nat", "n/a"}


def standardize_id(value):
    """' stu-005 ' -> 'STU005'.  Returns None if the ID is missing."""
    if is_blank(value):
        return None
    text = re.sub(r"\.0+$", "", str(value).strip())      # 101.0 -> 101
    text = re.sub(r"[^A-Za-z0-9]", "", text).upper()
    return text or None


def parse_date(value):
    """Convert many date styles to a Timestamp. Returns NaT if invalid."""
    if is_blank(value):
        return pd.NaT
    if isinstance(value, (pd.Timestamp, datetime)):
        return pd.Timestamp(value).normalize()
    text = str(value).strip()
    for fmt in DATE_FORMATS:
        try:
            return pd.Timestamp(datetime.strptime(text, fmt))
        except ValueError:
            continue
    return pd.NaT


def normalize_status(value):
    """'p' / 'PRESENT' -> 'Present'. Returns None if not recognised."""
    if is_blank(value):
        return None
    key = re.sub(r"\s+", " ", str(value).strip().lower())
    return STATUS_MAP.get(key)


def normalize_columns(df):
    """Lower-case column names, spaces -> underscores, apply aliases."""
    out = df.copy()
    new_names = []
    for col in out.columns:
        name = re.sub(r"[\s\-]+", "_", str(col).strip().lower())
        new_names.append(COLUMN_ALIASES.get(name, name))
    out.columns = new_names
    return out


def canonicalize(series):
    """Make 'cse', 'CSE', 'Cse' all use the most common spelling."""
    if series.empty:
        return series
    key = series.str.casefold()
    best = series.groupby(key).agg(lambda s: s.value_counts().index[0])
    return key.map(best)


def clean_data(students_raw, att_raw, verbose=True):
    """Validate and clean both tables.

    Returns a dict with: ok, errors, students, attendance, issues, summary.
    Nothing is silently discarded - every rejected or changed record is
    listed in the 'issues' table.
    """
    result = {"ok": False, "errors": [], "summary": {},
              "students": pd.DataFrame(columns=STUDENT_COLS),
              "attendance": pd.DataFrame(columns=ATT_COLS),
              "issues": pd.DataFrame(columns=ISSUE_COLS)}

    if students_raw is None or att_raw is None:
        result["errors"].append("No data loaded. Use the Upload Data page.")
        return result

    s = normalize_columns(students_raw)
    a = normalize_columns(att_raw)
    miss_s = [c for c in STUDENT_COLS if c not in s.columns]
    miss_a = [c for c in ATT_COLS if c not in a.columns]
    if miss_s:
        result["errors"].append(
            f"Student file is missing required column(s): {', '.join(miss_s)}. "
            f"Required: {', '.join(STUDENT_COLS)}")
    if miss_a:
        result["errors"].append(
            f"Attendance file is missing required column(s): {', '.join(miss_a)}. "
            f"Required: {', '.join(ATT_COLS)}")
    if result["errors"]:
        if verbose:
            for err in result["errors"]:
                log.error("Validation error: %s", err)
        return result

    issues, counters = [], {"rejected": 0}

    def add_issue(source, row, sid, field, itype, original, action):
        issues.append({
            "Source": source, "Row": int(row),
            "Student ID": "(blank)" if is_blank(sid) else str(sid),
            "Field": field, "Issue Type": itype,
            "Original Value": "(blank)" if is_blank(original) else str(original),
            "Action": action})

    # ---------- student master ----------
    s = s[STUDENT_COLS].copy().reset_index(drop=True)
    s["_row"] = s.index + 2                       # +2 = Excel row number
    raw_ids = s["student_id"].copy()
    s["student_id"] = raw_ids.map(standardize_id)
    changed = s["student_id"].notna() & (raw_ids.astype(str).str.strip() != s["student_id"])
    ids_standardized = int(changed.sum())

    for _, r in s[s["student_id"].isna()].iterrows():
        add_issue("Student master", r["_row"], "", "student_id",
                  "Missing Student ID", "", "Rejected - student not loaded")
    s = s[s["student_id"].notna()].copy()

    for col in ["student_name", "department", "semester", "section"]:
        for _, r in s[s[col].map(is_blank)].iterrows():
            add_issue("Student master", r["_row"], r["student_id"], col,
                      "Missing Value", "", "Filled with 'Unknown'")
        s[col] = s[col].map(lambda v: "Unknown" if is_blank(v)
                            else re.sub(r"\s+", " ", str(v).strip()))
    s["department"] = canonicalize(s["department"])
    s["section"] = s["section"].str.upper()

    dup_students = s.duplicated("student_id", keep="first")
    for _, r in s[dup_students].iterrows():
        add_issue("Student master", r["_row"], r["student_id"], "student_id",
                  "Duplicate Student ID", r["student_id"], "Removed (first record kept)")
    s = s[~dup_students].drop(columns="_row").reset_index(drop=True)
    valid_ids = set(s["student_id"])

    # ---------- attendance ----------
    a = a[ATT_COLS].copy().reset_index(drop=True)
    a["_row"] = a.index + 2
    attendance_raw = len(a)
    a["raw_id"] = a["student_id"]
    a["student_id"] = a["raw_id"].map(standardize_id)
    a["raw_date"] = a["attendance_date"]
    a["date"] = a["raw_date"].map(parse_date)
    a["raw_status"] = a["attendance_status"]
    a["status"] = a["raw_status"].map(normalize_status)

    def reject(frame, mask_fn, itype, field, raw_col):
        """Remove rows matching the mask but record each one as an issue."""
        mask = mask_fn(frame)
        for _, r in frame[mask].iterrows():
            add_issue("Attendance", r["_row"], r["raw_id"], field, itype,
                      r[raw_col], "Rejected - excluded from calculations")
        counters["rejected"] += int(mask.sum())
        return frame[~mask].copy()

    a = reject(a, lambda f: f["student_id"].isna(), "Missing Student ID", "student_id", "raw_id")
    a = reject(a, lambda f: f["date"].isna() & f["raw_date"].map(is_blank),
               "Missing Date", "attendance_date", "raw_date")
    a = reject(a, lambda f: f["date"].isna(), "Invalid Date", "attendance_date", "raw_date")
    a = reject(a, lambda f: f["status"].isna() & f["raw_status"].map(is_blank),
               "Missing Status", "attendance_status", "raw_status")
    a = reject(a, lambda f: f["status"].isna(), "Unknown Status", "attendance_status", "raw_status")
    a = reject(a, lambda f: ~f["student_id"].isin(valid_ids),
               "Unknown Student ID", "student_id", "raw_id")

    a["date"] = pd.to_datetime(a["date"])
    dates_standardized = int((a["raw_date"].astype(str).str.strip()
                              != a["date"].dt.strftime("%Y-%m-%d")).sum())
    statuses_normalized = int((a["raw_status"].astype(str).str.strip() != a["status"]).sum())

    # Subject: blank -> 'Unspecified' (record kept, but reported)
    blank_subject = a["subject"].map(is_blank)
    for _, r in a[blank_subject].iterrows():
        add_issue("Attendance", r["_row"], r["student_id"], "subject",
                  "Missing Value", "", "Filled with 'Unspecified' (record kept)")
    a["subject"] = a["subject"].map(lambda v: "Unspecified" if is_blank(v)
                                    else re.sub(r"\s+", " ", str(v).strip()))
    a["subject"] = canonicalize(a["subject"])

    # Duplicates: same student + date + subject. First record is kept.
    keys = ["student_id", "date", "subject"]
    dup_mask = a.duplicated(keys, keep="first")
    n_dup = int(dup_mask.sum())
    if n_dup:
        firsts = a.loc[~dup_mask, keys + ["status"]].rename(columns={"status": "first_status"})
        for _, r in a[dup_mask].merge(firsts, on=keys, how="left").iterrows():
            conflict = r["status"] != r["first_status"]
            add_issue("Attendance", r["_row"], r["student_id"], "record",
                      "Conflicting Duplicate" if conflict else "Duplicate Record",
                      f"{r['date']:%Y-%m-%d} | {r['subject']} | {r['status']}",
                      "Removed (first record kept"
                      + (f"; it says {r['first_status']})" if conflict else ")"))
    a = a[~dup_mask]

    att_clean = (a[["student_id", "date", "subject", "status"]]
                 .rename(columns={"date": "attendance_date", "status": "attendance_status"})
                 .sort_values(["attendance_date", "student_id"])
                 .reset_index(drop=True))

    issues_df = pd.DataFrame(issues, columns=ISSUE_COLS)
    if not issues_df.empty:
        issues_df = issues_df.sort_values(["Source", "Issue Type", "Row"]).reset_index(drop=True)

    result["summary"] = {
        "students_raw": len(students_raw), "students_clean": len(s),
        "attendance_raw": attendance_raw, "attendance_final": len(att_clean),
        "rejected": counters["rejected"], "duplicates": n_dup,
        "ids_standardized": ids_standardized,
        "dates_standardized": dates_standardized,
        "statuses_normalized": statuses_normalized,
        "issues_logged": len(issues_df)}
    result.update(ok=True, students=s, attendance=att_clean, issues=issues_df)

    if verbose:
        sm = result["summary"]
        log.info("Data cleaning finished: %d/%d attendance rows accepted, %d rejected, "
                 "%d duplicates removed, %d IDs and %d dates standardized",
                 sm["attendance_final"], sm["attendance_raw"], sm["rejected"],
                 sm["duplicates"], sm["ids_standardized"], sm["dates_standardized"])
        if len(issues_df):
            counts = issues_df["Issue Type"].value_counts().to_dict()
            log.warning("Validation issues found: %s", counts)
    return result


# ---------------------------------------------------------------------
# 5. ATTENDANCE ANALYTICS
# ---------------------------------------------------------------------
def add_flags(att, od_mode):
    """Add 'eligible' and 'attended' (0/1) columns to every record.

    Holiday is never eligible.  OD depends on the chosen mode.
    """
    if od_mode not in OD_MODES:
        raise ValueError(f"Unknown OD mode: {od_mode}")
    df = att.copy()
    status = df["attendance_status"]
    if od_mode == OD_ATTENDED:
        eligible, attended = status.isin(["Present", "Absent", "OD"]), status.isin(["Present", "OD"])
    elif od_mode == OD_NOT_ATTENDED:
        eligible, attended = status.isin(["Present", "Absent", "OD"]), status.eq("Present")
    else:  # OD_EXCLUDED
        eligible, attended = status.isin(["Present", "Absent"]), status.eq("Present")
    df["eligible"] = eligible.astype(int)
    df["attended"] = attended.astype(int)
    return df


def compute_student_summary(students, flagged, threshold):
    """One row per student with sessions, percentage and status."""
    out = students[STUDENT_COLS].copy().reset_index(drop=True)
    if flagged.empty:
        counts = pd.DataFrame(0, index=pd.Index([], dtype=object), columns=STATUSES)
        totals = pd.DataFrame(0, index=pd.Index([], dtype=object), columns=["eligible", "attended"])
    else:
        counts = (pd.crosstab(flagged["student_id"], flagged["attendance_status"])
                  .reindex(columns=STATUSES, fill_value=0))
        totals = flagged.groupby("student_id")[["eligible", "attended"]].sum()

    out["od_count"] = out["student_id"].map(counts["OD"]).fillna(0).astype(int)
    out["holiday_count"] = out["student_id"].map(counts["Holiday"]).fillna(0).astype(int)
    out["eligible_sessions"] = out["student_id"].map(totals["eligible"]).fillna(0).astype(int)
    out["attended_sessions"] = out["student_id"].map(totals["attended"]).fillna(0).astype(int)
    out["absent_sessions"] = out["eligible_sessions"] - out["attended_sessions"]

    # Zero eligible sessions -> percentage is NaN (shown as N/A), never a crash.
    eligible = out["eligible_sessions"].replace(0, np.nan)
    out["attendance_pct"] = out["attended_sessions"] / eligible * 100

    def decide(row):
        if row["eligible_sessions"] == 0:
            return "No Data"
        # integer-style comparison avoids floating point surprises at exactly 75%
        if row["attended_sessions"] * 100 < threshold * row["eligible_sessions"] - 1e-9:
            return "Shortage"
        return "Sufficient"

    out["status"] = out.apply(decide, axis=1) if len(out) else pd.Series(dtype=object)
    out["shortage_pct"] = np.where(out["status"] == "Shortage",
                                   threshold - out["attendance_pct"], 0.0)
    return out


def group_summary(ss, by):
    """Department-wise / section-wise summary (weighted by sessions)."""
    cols = by + ["students", "eligible_sessions", "attended_sessions", "absent_sessions",
                 "attendance_pct", "shortage_students", "sufficient_students", "no_data_students"]
    if ss.empty:
        return pd.DataFrame(columns=cols)
    g = ss.groupby(by).agg(
        students=("student_id", "count"),
        eligible_sessions=("eligible_sessions", "sum"),
        attended_sessions=("attended_sessions", "sum"),
        absent_sessions=("absent_sessions", "sum"),
        shortage_students=("status", lambda s: int((s == "Shortage").sum())),
        sufficient_students=("status", lambda s: int((s == "Sufficient").sum())),
        no_data_students=("status", lambda s: int((s == "No Data").sum())),
    ).reset_index()
    g["attendance_pct"] = g["attended_sessions"] / g["eligible_sessions"].replace(0, np.nan) * 100
    return g[cols]


def analyze(students_raw, att_raw, od_mode=OD_ATTENDED, threshold=DEFAULT_THRESHOLD, verbose=True):
    """The full pipeline: clean -> flag -> summarise. Never raises."""
    base = {"threshold": threshold, "od_mode": od_mode}
    try:
        clean = clean_data(students_raw, att_raw, verbose=verbose)
        if not clean["ok"]:
            return {**clean, **base}

        flagged = add_flags(clean["attendance"], od_mode)
        ss = compute_student_summary(clean["students"], flagged, threshold)

        shortage = ss[ss["status"] == "Shortage"].copy()
        shortage["required_threshold"] = threshold
        shortage = shortage.sort_values("attendance_pct")[[
            "student_id", "student_name", "department", "semester", "section",
            "attendance_pct", "required_threshold", "shortage_pct"]].reset_index(drop=True)

        subject = pd.DataFrame(columns=["subject", "eligible_sessions", "attended_sessions", "attendance_pct"])
        monthly = pd.DataFrame(columns=["month", "eligible_sessions", "attended_sessions", "attendance_pct"])
        if not flagged.empty:
            subject = (flagged.groupby("subject")[["eligible", "attended"]].sum().reset_index()
                       .rename(columns={"eligible": "eligible_sessions", "attended": "attended_sessions"}))
            subject = subject[subject["eligible_sessions"] > 0].copy()
            subject["attendance_pct"] = subject["attended_sessions"] / subject["eligible_sessions"] * 100
            tmp = flagged.assign(month=flagged["attendance_date"].dt.strftime("%Y-%m"))
            monthly = (tmp.groupby("month")[["eligible", "attended"]].sum().reset_index()
                       .rename(columns={"eligible": "eligible_sessions", "attended": "attended_sessions"}))
            monthly = monthly[monthly["eligible_sessions"] > 0].copy()
            monthly["attendance_pct"] = monthly["attended_sessions"] / monthly["eligible_sessions"] * 100

        total_eligible = int(ss["eligible_sessions"].sum())
        total_attended = int(ss["attended_sessions"].sum())
        overall = {
            "total_students": len(ss),
            "total_records": len(clean["attendance"]),
            "eligible": total_eligible, "attended": total_attended,
            "pct": (total_attended / total_eligible * 100) if total_eligible else float("nan"),
            "shortage": int((ss["status"] == "Shortage").sum()),
            "sufficient": int((ss["status"] == "Sufficient").sum()),
            "no_data": int((ss["status"] == "No Data").sum())}

        if verbose:
            log.info("Attendance calculated: %d students, overall %s, %d below %.1f%% (OD mode: %s)",
                     overall["total_students"],
                     "N/A" if total_eligible == 0 else f"{overall['pct']:.2f}%",
                     overall["shortage"], threshold, od_mode)

        return {**clean, **base, "student_summary": ss, "shortage": shortage,
                "dept_summary": group_summary(ss, ["department"]),
                "section_summary": group_summary(ss, ["department", "section"]),
                "subject_summary": subject, "monthly": monthly, "overall": overall}
    except Exception as exc:  # last line of defence: never crash the UI
        log.exception("Unexpected error during analysis")
        return {"ok": False, "errors": [f"Unexpected error: {exc}"], "summary": {},
                "issues": pd.DataFrame(columns=ISSUE_COLS), **base}


# ---------------------------------------------------------------------
# 6. EXCEL REPORT BUILDERS (OpenPyXL)
# ---------------------------------------------------------------------
HEADER_FILL = PatternFill("solid", fgColor="1E3A8A")
HEADER_FONT = Font(bold=True, color="FFFFFF", size=11)
THIN = Side(style="thin", color="CBD5E1")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
RED_FILL = PatternFill("solid", bgColor="FECACA", fgColor="FECACA")
GREEN_FILL = PatternFill("solid", bgColor="BBF7D0", fgColor="BBF7D0")


def _excel_value(v):
    """Convert pandas/numpy values to plain Python values for OpenPyXL."""
    if isinstance(v, pd.Timestamp):
        return v.to_pydatetime()
    if v is None:
        return None
    try:
        if pd.isna(v):
            return None
    except (TypeError, ValueError):
        pass
    if hasattr(v, "item") and not isinstance(v, str):
        return v.item()
    return v


def write_table(ws, df, start_row=1, pct_cols=(), filters=True):
    """Write a DataFrame with a styled header, borders and optional filter.

    Columns named in pct_cols hold 0-100 numbers and are stored as true
    Excel percentages (0.00% format). Returns (first_data_row, last_row).
    """
    for j, col in enumerate(df.columns, start=1):
        cell = ws.cell(row=start_row, column=j, value=str(col))
        cell.font, cell.fill, cell.border = HEADER_FONT, HEADER_FILL, BORDER
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for i, row in enumerate(df.itertuples(index=False), start=start_row + 1):
        for j, (col, val) in enumerate(zip(df.columns, row), start=1):
            val = _excel_value(val)
            cell = ws.cell(row=i, column=j)
            if col in pct_cols:
                cell.value = "N/A" if val is None else float(val) / 100
                cell.number_format = "0.00%"
                cell.alignment = Alignment(horizontal="right")
            else:
                cell.value = val
                if isinstance(val, datetime):
                    cell.number_format = "yyyy-mm-dd"
            cell.border = BORDER
    first, last = start_row + 1, start_row + len(df)
    if filters:
        ws.auto_filter.ref = f"A{start_row}:{get_column_letter(len(df.columns))}{max(last, first)}"
        ws.freeze_panes = ws.cell(row=first, column=1)
    return first, last


def autosize(ws, skip_rows=()):
    widths = {}
    for row in ws.iter_rows():
        for c in row:
            if c.value is None or c.row in skip_rows:
                continue
            text = f"{c.value:.2%}" if isinstance(c.value, float) and c.number_format == "0.00%" else str(c.value)
            widths[c.column_letter] = max(widths.get(c.column_letter, 0), len(text))
    for letter, w in widths.items():
        ws.column_dimensions[letter].width = min(max(w + 3, 12), 48)


def column_letter_of(ws, header_row, name):
    for c in ws[header_row]:
        if c.value == name:
            return c.column_letter
    return None


def add_percentage_rules(ws, header_row, first, last, column_name, threshold):
    """Red below threshold, green at/above (conditional formatting)."""
    letter = column_letter_of(ws, header_row, column_name)
    if not letter or last < first:
        return
    rng = f"{letter}{first}:{letter}{last}"
    ws.conditional_formatting.add(rng, CellIsRule(
        operator="lessThan", formula=[str(threshold / 100)], fill=RED_FILL, font=Font(color="991B1B", bold=True)))
    ws.conditional_formatting.add(rng, CellIsRule(
        operator="greaterThanOrEqual", formula=[str(threshold / 100)], fill=GREEN_FILL, font=Font(color="166534")))


def pretty(df):
    return df.rename(columns=DISPLAY)


def summary_table(summary):
    return pd.DataFrame([(label, summary.get(key, 0)) for key, label in SUMMARY_LABELS],
                        columns=["Check", "Count"])


def build_report_workbook(ctx):
    """Main Excel workbook with the five required sheets."""
    thr, od_mode, ov = ctx["threshold"], ctx["od_mode"], ctx["overall"]
    wb = Workbook()

    # 1. Overall report
    ws = wb.active
    ws.title = "Overall Report"
    ws["A1"] = "Student Attendance Report"
    ws["A1"].font = Font(bold=True, size=16, color="1E3A8A")
    ws["A2"] = f"Generated: {datetime.now():%Y-%m-%d %H:%M}"
    ws["A3"] = f"Shortage threshold: {thr:g}%   |   OD handling: {od_mode}"
    ws["A2"].font = ws["A3"].font = Font(italic=True, color="475569")
    metrics = [
        ("Total students", ov["total_students"], None),
        ("Total attendance records (after cleaning)", ov["total_records"], None),
        ("Total eligible sessions", ov["eligible"], None),
        ("Total attended sessions", ov["attended"], None),
        ("Overall attendance percentage", None if ov["eligible"] == 0 else ov["pct"] / 100, "0.00%"),
        ("Students below threshold (shortage)", ov["shortage"], None),
        ("Students with sufficient attendance", ov["sufficient"], None),
        ("Students with no eligible sessions", ov["no_data"], None),
    ]
    write_table(ws, pd.DataFrame(columns=["Metric", "Value"]), start_row=5, filters=False)
    for i, (label, value, fmt) in enumerate(metrics, start=6):
        a, b = ws.cell(row=i, column=1, value=label), ws.cell(row=i, column=2, value="N/A" if value is None else value)
        a.border = b.border = BORDER
        b.alignment = Alignment(horizontal="right")
        if fmt:
            b.number_format = fmt
    autosize(ws, skip_rows=(1, 2, 3))

    # 2. Student-wise summary
    ws = wb.create_sheet("Student Summary")
    df = pretty(ctx["student_summary"][STUDENT_VIEW_COLS])
    first, last = write_table(ws, df, pct_cols=("Attendance %",))
    add_percentage_rules(ws, 1, first, last, "Attendance %", thr)
    status_col = column_letter_of(ws, 1, "Status")
    if status_col and last >= first:
        ws.conditional_formatting.add(f"{status_col}{first}:{status_col}{last}", CellIsRule(
            operator="equal", formula=['"Shortage"'], fill=RED_FILL, font=Font(color="991B1B", bold=True)))
    autosize(ws)

    # 3. Shortage list
    ws = wb.create_sheet("Shortage List")
    df = pretty(ctx["shortage"])
    first, last = write_table(ws, df, pct_cols=("Attendance %", "Required %", "Shortage %"))
    letter = column_letter_of(ws, 1, "Shortage %")
    if letter and last >= first:
        ws.conditional_formatting.add(f"{letter}{first}:{letter}{last}", ColorScaleRule(
            start_type="min", start_color="FEE2E2", end_type="max", end_color="EF4444"))
    autosize(ws)

    # 4. Department-wise summary
    ws = wb.create_sheet("Department Summary")
    first, last = write_table(ws, pretty(ctx["dept_summary"]), pct_cols=("Attendance %",))
    add_percentage_rules(ws, 1, first, last, "Attendance %", thr)
    autosize(ws)

    # 5. Data validation report
    ws = wb.create_sheet("Data Validation Report")
    ws["A1"] = "Validation summary"
    ws["A1"].font = Font(bold=True, size=13, color="1E3A8A")
    _, end = write_table(ws, summary_table(ctx["summary"]), start_row=2, filters=False)
    ws.cell(row=end + 2, column=1, value="Detailed issues (every rejected or changed record)").font = Font(
        bold=True, size=13, color="1E3A8A")
    write_table(ws, ctx["issues"], start_row=end + 3)
    autosize(ws, skip_rows=(1, end + 2))
    ws.column_dimensions["A"].width = 48

    buffer = io.BytesIO()
    wb.save(buffer)
    return buffer.getvalue()


def df_to_xlsx_bytes(df, sheet_name, pct_cols=()):
    """Small helper: one DataFrame -> formatted .xlsx bytes."""
    wb = Workbook()
    ws = wb.active
    ws.title = sheet_name
    write_table(ws, df, pct_cols=pct_cols)
    autosize(ws)
    buffer = io.BytesIO()
    wb.save(buffer)
    return buffer.getvalue()


# ---------------------------------------------------------------------
# 7. TEST SUITE (12 test cases)
# ---------------------------------------------------------------------
def _dates(n, start="2026-09-14"):
    base = pd.Timestamp(start)
    return [(base + pd.Timedelta(days=i)).strftime("%Y-%m-%d") for i in range(n)]


def _students(*ids):
    return pd.DataFrame({"student_id": list(ids),
                         "student_name": [f"Test Student {i + 1}" for i in range(len(ids))],
                         "department": "CSE", "semester": "3", "section": "A"})


def _att(rows):
    return pd.DataFrame(rows, columns=ATT_COLS)


def _rows(statuses, sid="S001", subject="Maths"):
    return [(sid, d, subject, s) for d, s in zip(_dates(len(statuses)), statuses)]


def _run(rows, ids=("S001",), od=OD_ATTENDED, thr=75.0):
    return analyze(_students(*ids), _att(rows), od, thr, verbose=False)


def _row(res, sid="S001"):
    return res["student_summary"].set_index("student_id").loc[sid]


def _fmt(x):
    return (f"Eligible={int(x['eligible_sessions'])}, Attended={int(x['attended_sessions'])}, "
            f"Percentage={'N/A' if pd.isna(x['attendance_pct']) else format(x['attendance_pct'], '.2f') + '%'}, "
            f"Status={x['status']}")


def _issue_count(res, itype):
    return int((res["issues"]["Issue Type"] == itype).sum())


def t01():
    x = _row(_run(_rows(["Present"] * 8 + ["Absent"] * 2)))
    return _fmt(x), (x["eligible_sessions"] == 10 and x["attended_sessions"] == 8
                     and abs(x["attendance_pct"] - 80) < 1e-9 and x["status"] == "Sufficient")


def t02():
    r = _run(_rows(["Present"] * 5))
    x = _row(r)
    return _fmt(x) + f", Shortage rows={len(r['shortage'])}", (
        abs(x["attendance_pct"] - 100) < 1e-9 and x["status"] == "Sufficient" and r["shortage"].empty)


def t03():
    r = _run(_rows(["Present"] * 6 + ["Absent"] * 4))
    x = _row(r)
    ok = (abs(x["attendance_pct"] - 60) < 1e-9 and x["status"] == "Shortage"
          and len(r["shortage"]) == 1 and abs(r["shortage"].iloc[0]["shortage_pct"] - 15) < 1e-9)
    return _fmt(x) + f", Shortage %={r['shortage']['shortage_pct'].iloc[0]:.2f}", ok


def t04():
    r = _run(_rows(["Present"] * 3 + ["Absent"]))
    x = _row(r)
    return _fmt(x) + f", Shortage rows={len(r['shortage'])}", (
        abs(x["attendance_pct"] - 75) < 1e-9 and x["status"] == "Sufficient" and r["shortage"].empty)


def t05():
    d = _dates(3)
    rows = [("S001", d[0], "Maths", "Present"), ("S001", d[0], "Maths", "Present"),
            ("S001", d[1], "Maths", "Present"), ("S001", d[2], "Maths", "Absent")]
    r = _run(rows)
    x = _row(r)
    n = _issue_count(r, "Duplicate Record")
    return _fmt(x) + f", Duplicates flagged={n}", (
        x["eligible_sessions"] == 3 and x["attended_sessions"] == 2
        and abs(x["attendance_pct"] - 66.6667) < 0.01 and n == 1)


def t06():
    rows = _rows(["Present"]) + [("S001", "2026-13-45", "Maths", "Present"),
                                 ("S001", "31/02/2026", "Maths", "Present"),
                                 ("S001", "not a date", "Maths", "Present")]
    r = _run(rows)
    x = _row(r)
    n = _issue_count(r, "Invalid Date")
    return _fmt(x) + f", Invalid dates flagged={n}", n == 3 and x["eligible_sessions"] == 1


def t07():
    rows = _rows(["Present"]) + [("S001", _dates(3)[1], "Maths", "Maybe"),
                                 ("S001", _dates(3)[2], "Maths", "Late")]
    r = _run(rows)
    x = _row(r)
    n = _issue_count(r, "Unknown Status")
    return _fmt(x) + f", Unknown statuses flagged={n}", n == 2 and x["eligible_sessions"] == 1


def t08():
    d = _dates(3)
    rows = [("S001", d[0], "Maths", "Present"), ("", d[1], "Maths", "Present"),
            (None, d[2], "Maths", "Absent")]
    r = _run(rows)
    x = _row(r)
    n = _issue_count(r, "Missing Student ID")
    return _fmt(x) + f", Missing IDs flagged={n}", n == 2 and x["eligible_sessions"] == 1


def t09():
    r1 = _run([])                                    # header-only attendance file
    x = _row(r1)
    part_a = r1["ok"] and r1["overall"]["total_records"] == 0 and x["status"] == "No Data" and pd.isna(x["attendance_pct"])
    r2 = analyze(pd.DataFrame(), pd.DataFrame(), verbose=False)   # totally empty files
    part_b = (not r2["ok"]) and len(r2["errors"]) >= 1
    return (f"Header-only file: ok={r1['ok']}, records={r1['overall']['total_records']}, status={x['status']}; "
            f"Fully empty files: ok={r2['ok']}, error shown={len(r2['errors']) >= 1}"), bool(part_a and part_b)


def t10():
    r = _run(_rows(["Present", "Present", "Absent", "Holiday", "Holiday", "Holiday"]))
    x = _row(r)
    return _fmt(x) + f", Holidays={int(x['holiday_count'])}", (
        x["eligible_sessions"] == 3 and x["attended_sessions"] == 2 and abs(x["attendance_pct"] - 66.6667) < 0.01)


def t11():
    rows = _rows(["Present", "Present", "Absent", "OD"])
    a = _row(_run(rows, od=OD_ATTENDED))["attendance_pct"]
    n = _row(_run(rows, od=OD_NOT_ATTENDED))["attendance_pct"]
    e = _row(_run(rows, od=OD_EXCLUDED))["attendance_pct"]
    return (f"OD attended={a:.2f}%, OD not attended={n:.2f}%, OD excluded={e:.2f}%",
            abs(a - 75) < 1e-9 and abs(n - 50) < 1e-9 and abs(e - 66.6667) < 0.01)


def t12():
    rows = _rows(["Holiday"] * 3)
    r = _run(rows, ids=("S001", "S002"))             # S002 has no records at all
    x1, x2 = _row(r, "S001"), _row(r, "S002")
    ok = (x1["eligible_sessions"] == 0 and x2["eligible_sessions"] == 0 and pd.isna(x1["attendance_pct"])
          and x1["status"] == "No Data" and x2["status"] == "No Data" and r["shortage"].empty)
    return f"S001 (only holidays): {_fmt(x1)} | S002 (no records): {_fmt(x2)}", bool(ok)


TEST_CASES = [
    ("T01", "Normal attendance calculation", "10 sessions: 8 Present, 2 Absent", "80.00% - Sufficient", t01),
    ("T02", "100% attendance", "5 sessions, all Present", "100.00%, not in shortage list", t02),
    ("T03", "Attendance below 75%", "10 sessions: 6 Present, 4 Absent", "60.00% - Shortage, shortage 15.00%", t03),
    ("T04", "Attendance exactly at 75%", "4 sessions: 3 Present, 1 Absent", "75.00% - Sufficient (not flagged)", t04),
    ("T05", "Duplicate records", "Same student/date/subject entered twice", "Duplicate removed and reported; 2 of 3 = 66.67%", t05),
    ("T06", "Invalid dates", "3 impossible/garbage dates + 1 valid row", "3 rows rejected and listed; valid row counted", t06),
    ("T07", "Unknown attendance statuses", "'Maybe' and 'Late' statuses + 1 valid row", "2 rows rejected and listed", t07),
    ("T08", "Missing student IDs", "Blank and null IDs + 1 valid row", "2 rows rejected and listed", t08),
    ("T09", "Empty datasets", "Header-only files, and completely empty files", "No crash; 'No Data' status / clear error message", t09),
    ("T10", "Holiday exclusion", "2 Present, 1 Absent, 3 Holiday", "3 eligible sessions, 66.67%", t10),
    ("T11", "OD handling", "2 Present, 1 Absent, 1 OD in three modes", "75.00% / 50.00% / 66.67%", t11),
    ("T12", "Zero eligible sessions", "Student with only holidays; student with no records", "Percentage N/A, status 'No Data', no crash", t12),
]


def run_all_tests():
    """Actually execute every test case against the real code above."""
    rows, started = [], datetime.now()
    for tid, name, _inp, expected, fn in TEST_CASES:
        try:
            actual, passed = fn()
        except Exception as exc:
            actual, passed = f"Test raised an error: {exc}", False
        rows.append({"Test ID": tid, "Scenario": name, "Expected Result": expected,
                     "Actual Result": actual, "Result": "PASSED" if passed else "FAILED",
                     "Executed At": started.strftime("%Y-%m-%d %H:%M:%S")})
    df = pd.DataFrame(rows)
    log.info("Test suite executed: %d passed, %d failed", (df["Result"] == "PASSED").sum(),
             (df["Result"] == "FAILED").sum())
    return df


def predefined_tests_df():
    return pd.DataFrame([{"Test ID": t, "Scenario": n, "Input Description": i, "Expected Result": e}
                         for t, n, i, e, _ in TEST_CASES])


def build_test_workbook(results):
    """Test workbook. 'results' is None until the tests were really run."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Predefined Test Cases"
    write_table(ws, predefined_tests_df())
    autosize(ws)

    ws2 = wb.create_sheet("Execution Results")
    if results is None:
        ws2["A1"] = "NOT EXECUTED"
        ws2["A1"].font = Font(bold=True, size=14, color="B91C1C")
        ws2["A2"] = ("No tests had been run when this workbook was generated, so no pass/fail "
                     "result is claimed. Click 'Run all tests now' in the app and download again.")
        ws2.column_dimensions["A"].width = 110
    else:
        first, last = write_table(ws2, results)
        letter = column_letter_of(ws2, 1, "Result")
        ws2.conditional_formatting.add(f"{letter}{first}:{letter}{last}", CellIsRule(
            operator="equal", formula=['"PASSED"'], fill=GREEN_FILL, font=Font(color="166534", bold=True)))
        ws2.conditional_formatting.add(f"{letter}{first}:{letter}{last}", CellIsRule(
            operator="equal", formula=['"FAILED"'], fill=RED_FILL, font=Font(color="991B1B", bold=True)))
        autosize(ws2)

    ws3 = wb.create_sheet("Read Me")
    notes = ["'Predefined Test Cases' = what each test is designed to check and the result we EXPECT.",
             "'Execution Results' = what actually happened when the tests were run by the app.",
             "A result is only PASSED/FAILED if it was really executed (see 'Executed At' timestamp).",
             "Tests run the same cleaning and calculation functions the dashboard uses."]
    for i, line in enumerate(notes, start=1):
        ws3.cell(row=i, column=1, value=line)
    ws3.column_dimensions["A"].width = 100
    buffer = io.BytesIO()
    wb.save(buffer)
    return buffer.getvalue()


# ---------------------------------------------------------------------
# 8. STREAMLIT USER INTERFACE
# ---------------------------------------------------------------------
CSS = """
<style>
.block-container {padding-top: 1.4rem; padding-bottom: 2rem; max-width: 1400px;}
.page-header {background: linear-gradient(120deg,#1e3a8a 0%,#4f46e5 100%); padding: 20px 26px;
  border-radius: 16px; margin-bottom: 18px; box-shadow: 0 4px 14px rgba(30,58,138,.25);}
.page-header h1 {color:#fff; margin:0; padding:0; font-size:1.7rem; font-weight:700;}
.page-header p {color:#e0e7ff; margin:4px 0 0 0; font-size:.95rem;}
.kpi {background:#ffffff; border-radius:14px; padding:14px 18px; box-shadow:0 2px 10px rgba(15,23,42,.10);
  min-height:104px;}
.kpi-label {color:#64748b; font-size:.78rem; font-weight:600; text-transform:uppercase; letter-spacing:.04em;}
.kpi-value {color:#0f172a; font-size:1.9rem; font-weight:700; line-height:1.25;}
.kpi-sub {color:#94a3b8; font-size:.78rem;}
.brand {background: linear-gradient(120deg,#1e3a8a 0%,#4f46e5 100%); color:#fff; padding:14px 16px;
  border-radius:12px; margin-bottom:10px; font-weight:700; font-size:1.05rem;}
.brand small {display:block; font-weight:400; color:#c7d2fe; font-size:.75rem;}
</style>
"""

PAGES = ["📊 Dashboard", "📁 Upload Data", "🧹 Validation & Cleaning", "📈 Attendance Analytics",
         "🚨 Shortage Detection", "📄 Reports", "🧪 Testing", "📝 Activity Logs"]


def _stretch(fn, *args, **kwargs):
    """Call a Streamlit element so it fills the width, in old and new versions."""
    try:
        return fn(*args, width="stretch", **kwargs)
    except Exception:
        return fn(*args, use_container_width=True, **kwargs)


def page_header(title, subtitle):
    st.markdown(f'<div class="page-header"><h1>{title}</h1><p>{subtitle}</p></div>', unsafe_allow_html=True)


def kpi_row(items):
    """items = [(label, value, subtext, color), ...]"""
    cols = st.columns(len(items))
    for col, (label, value, sub, color) in zip(cols, items):
        col.markdown(
            f'<div class="kpi" style="border-left:6px solid {color}"><div class="kpi-label">{label}</div>'
            f'<div class="kpi-value">{value}</div><div class="kpi-sub">{sub}</div></div>',
            unsafe_allow_html=True)


def show_df(df, height=None):
    """Display a DataFrame with friendly headings and a progress bar for %."""
    d = df.rename(columns=DISPLAY).copy()
    for c in ("Attendance %", "Shortage %", "Required %"):
        if c in d.columns:
            d[c] = d[c].astype(float).round(2)
    config = {}
    if "Attendance %" in d.columns:
        config["Attendance %"] = st.column_config.ProgressColumn(
            "Attendance %", min_value=0, max_value=100, format="%.1f")
    kwargs = {"hide_index": True, "column_config": config}
    if height:
        kwargs["height"] = height
    _stretch(st.dataframe, d, **kwargs)


def show_chart(fig, key):
    fig.update_layout(margin=dict(l=10, r=10, t=50, b=10), title_font_size=16)
    _stretch(st.plotly_chart, fig, key=key)


def flash(kind, message):
    st.session_state["flash"] = (kind, message)


def pct_text(value):
    return "N/A" if pd.isna(value) else f"{value:.1f}%"


def init_state():
    ss = st.session_state
    if "students_raw" not in ss:
        s, a = generate_sample_data(messy=True)
        ss.students_raw, ss.att_raw = s, a
        ss.data_source = "Built-in sample data (with demo data-quality issues)"
        ss.data_version = 0
    ss.setdefault("threshold", DEFAULT_THRESHOLD)
    ss.setdefault("od_mode", OD_ATTENDED)
    ss.setdefault("report_bytes", None)
    ss.setdefault("test_results", None)
    if "session_started" not in ss:
        ss.session_started = True
        log.info("New user session started (sample data loaded)")


def require_data(ctx):
    """Show friendly errors and return False if the data cannot be analysed."""
    if ctx.get("ok"):
        return True
    for err in ctx.get("errors", []):
        st.error(err)
    st.info("Go to **📁 Upload Data** to load a valid file or reload the sample data.")
    return False


@st.cache_data(show_spinner=False)
def cached_analyze(students_raw, att_raw, od_mode, threshold):
    """Cached so charts do not recompute on every click."""
    return analyze(students_raw, att_raw, od_mode, threshold)


def read_table(uploaded):
    """Read an uploaded CSV/Excel file as text (so IDs keep their zeros)."""
    name, data = uploaded.name.lower(), uploaded.getvalue()
    if name.endswith(".csv"):
        last_error = None
        for enc in ("utf-8-sig", "latin-1"):
            try:
                return pd.read_csv(io.BytesIO(data), dtype=str, keep_default_na=False, encoding=enc)
            except UnicodeDecodeError as exc:
                last_error = exc
        raise last_error
    if name.endswith((".xlsx", ".xlsm")):
        return pd.read_excel(io.BytesIO(data), dtype=str, keep_default_na=False, engine="openpyxl")
    raise ValueError("Unsupported file type. Please upload .csv, .xlsx or .xlsm files.")


def render_sidebar():
    with st.sidebar:
        st.markdown('<div class="brand">🎓 Attendance Analytics<small>DIT Internship 2026 · Task 07</small></div>',
                    unsafe_allow_html=True)
        page = st.radio("Navigation", PAGES, label_visibility="collapsed")
        st.divider()
        st.markdown("**⚙️ Settings**")
        st.number_input("Attendance shortage threshold (%)", min_value=0.0, max_value=100.0,
                        step=1.0, key="threshold")
        st.selectbox("OD (On Duty) handling", OD_MODES, key="od_mode")
        st.divider()
        st.caption(f"**Data source:** {st.session_state.data_source}")
    return page


# ----- Page: Dashboard -----
def page_dashboard(ctx):
    page_header("Attendance Dashboard", "Live overview of attendance across departments, subjects and months")
    if not require_data(ctx):
        return
    ov, thr = ctx["overall"], ctx["threshold"]
    ss = ctx["student_summary"]
    kpi_row([
        ("Total Students", ov["total_students"], "in the student master", COLOR_BLUE),
        ("Attendance Records", f"{ov['total_records']:,}", "after cleaning", "#7c3aed"),
        ("Overall Attendance", pct_text(ov["pct"]), f"{ov['attended']:,} of {ov['eligible']:,} eligible sessions", "#0891b2"),
        ("Below Threshold", ov["shortage"], f"under {thr:g}% attendance", COLOR_BAD),
        ("Sufficient Attendance", ov["sufficient"], f"{thr:g}% or above", COLOR_GOOD),
    ])
    st.write("")
    if ov["total_records"] == 0:
        st.warning("There are no valid attendance records to chart yet.")
        return

    c1, c2 = st.columns(2)
    with c1:
        counts = pd.DataFrame({"Status": ["Sufficient", "Shortage", "No Data"],
                               "Students": [ov["sufficient"], ov["shortage"], ov["no_data"]]})
        fig = px.pie(counts, names="Status", values="Students", hole=0.55, title="Sufficient vs shortage",
                     color="Status", color_discrete_map={"Sufficient": COLOR_GOOD, "Shortage": COLOR_BAD, "No Data": "#94a3b8"})
        show_chart(fig, "dash_donut")
    with c2:
        data = ss.dropna(subset=["attendance_pct"])
        fig = px.histogram(data, x="attendance_pct", nbins=15, title="Attendance % distribution",
                           labels={"attendance_pct": "Attendance %"}, color_discrete_sequence=[COLOR_BLUE])
        fig.add_vline(x=thr, line_dash="dash", line_color=COLOR_BAD, annotation_text=f"Threshold {thr:g}%")
        fig.update_layout(yaxis_title="Students")
        show_chart(fig, "dash_hist")

    c3, c4 = st.columns(2)
    with c3:
        dept = ctx["dept_summary"]
        fig = px.bar(dept, x="department", y="attendance_pct", text=dept["attendance_pct"].round(1),
                     title="Department-wise attendance %", labels={"department": "Department", "attendance_pct": "Attendance %"},
                     color="attendance_pct", color_continuous_scale="Blues")
        fig.add_hline(y=thr, line_dash="dash", line_color=COLOR_BAD)
        fig.update_layout(coloraxis_showscale=False, yaxis_range=[0, 105])
        show_chart(fig, "dash_dept")
    with c4:
        monthly = ctx["monthly"]
        fig = px.line(monthly, x="month", y="attendance_pct", markers=True, title="Monthly attendance trend",
                      labels={"month": "Month", "attendance_pct": "Attendance %"})
        fig.add_hline(y=thr, line_dash="dash", line_color=COLOR_BAD)
        fig.update_traces(line_color="#4f46e5", line_width=3)
        fig.update_layout(yaxis_range=[0, 105])
        show_chart(fig, "dash_month")

    c5, c6 = st.columns(2)
    with c5:
        subject = ctx["subject_summary"]
        fig = px.bar(subject, x="subject", y="attendance_pct", title="Subject-wise attendance %",
                     labels={"subject": "Subject", "attendance_pct": "Attendance %"}, color_discrete_sequence=["#0891b2"])
        fig.update_layout(yaxis_range=[0, 105])
        show_chart(fig, "dash_subject")
    with c6:
        st.markdown("##### Department-wise statistics")
        show_df(ctx["dept_summary"], height=250)


# ----- Page: Upload -----
def page_upload(ctx):
    page_header("Upload Attendance Data", "Load your own CSV / Excel files, or use the built-in sample dataset")
    ss = st.session_state

    st.subheader("1 · Sample dataset")
    st.caption("30 fictional students, 30 working days, 5 subjects. Includes both sufficient and shortage cases.")
    messy = st.checkbox("Include intentional data-quality problems (to demonstrate validation)", value=True, key="sample_messy")
    s_df, a_df = generate_sample_data(messy=messy)
    c1, c2, c3 = st.columns(3)
    c1.download_button("⬇️ Download sample students (CSV)", s_df.to_csv(index=False).encode("utf-8"),
                       "sample_students.csv", "text/csv")
    c2.download_button("⬇️ Download sample attendance (CSV)", a_df.to_csv(index=False).encode("utf-8"),
                       "sample_attendance.csv", "text/csv")
    if c3.button("🔄 Load this sample into the app"):
        ss.students_raw, ss.att_raw = s_df, a_df
        ss.data_source = "Built-in sample data" + (" (with demo data-quality issues)" if messy else " (clean)")
        ss.data_version += 1
        ss.report_bytes = None
        log.info("Sample data loaded (messy=%s)", messy)
        flash("success", "Sample data loaded.")
        st.rerun()

    st.divider()
    st.subheader("2 · Upload your own files")
    st.caption("Student file columns: student_id, student_name, department, semester, section.  "
               "Attendance file columns: student_id, attendance_date, subject, attendance_status "
               "(Present / Absent / OD / Holiday).")
    u1, u2 = st.columns(2)
    f_students = u1.file_uploader("Student master file", type=["csv", "xlsx", "xlsm"], key="up_students")
    f_att = u2.file_uploader("Attendance file", type=["csv", "xlsx", "xlsm"], key="up_att")
    loaded = False
    if st.button("✅ Process uploaded files", type="primary"):
        if not f_students or not f_att:
            st.error("Please upload BOTH files (student master and attendance) before processing.")
        else:
            try:
                s_new, a_new = read_table(f_students), read_table(f_att)
                log.info("File uploaded: %s (%d rows, %d columns)", f_students.name, *s_new.shape)
                log.info("File uploaded: %s (%d rows, %d columns)", f_att.name, *a_new.shape)
                ss.students_raw, ss.att_raw = s_new, a_new
                ss.data_source = f"Uploaded: {f_students.name} + {f_att.name}"
                ss.data_version += 1
                ss.report_bytes = None
                flash("success", f"Loaded {len(s_new)} student rows and {len(a_new)} attendance rows. "
                                 "Check the Validation & Cleaning page for the results.")
                loaded = True
            except Exception as exc:
                log.error("File upload failed: %s", exc)
                st.error(f"Could not read the files: {exc}")
    if loaded:
        st.rerun()

    st.divider()
    st.subheader("3 · Data currently loaded")
    t1, t2 = st.tabs([f"Students ({len(ss.students_raw)} rows)", f"Attendance ({len(ss.att_raw)} rows)"])
    with t1:
        _stretch(st.dataframe, ss.students_raw.head(200), hide_index=True)
    with t2:
        _stretch(st.dataframe, ss.att_raw.head(200), hide_index=True)
    st.caption("Showing the first 200 rows of the raw (not yet cleaned) data.")


# ----- Page: Validation -----
def page_validation(ctx):
    page_header("Data Validation & Cleaning", "Every problem found is listed here - nothing is silently discarded")
    if not require_data(ctx):
        return
    sm, issues = ctx["summary"], ctx["issues"]
    kpi_row([
        ("Rows Received", f"{sm['attendance_raw']:,}", "attendance file", COLOR_BLUE),
        ("Rows Accepted", f"{sm['attendance_final']:,}", "used in calculations", COLOR_GOOD),
        ("Rows Rejected", sm["rejected"], "invalid - see list below", COLOR_BAD),
        ("Duplicates Removed", sm["duplicates"], "first record kept", COLOR_WARN),
    ])
    st.write("")
    if sm["issues_logged"] == 0:
        st.success("No data-quality problems were found.")
    else:
        st.warning(f"{sm['issues_logged']} issue(s) found. Rejected rows are excluded from calculations "
                   "but are listed below so you can fix the source file.")

    left, right = st.columns([1, 1])
    with left:
        st.markdown("##### Validation summary")
        _stretch(st.dataframe, summary_table(sm), hide_index=True)
    with right:
        if not issues.empty:
            by_type = issues["Issue Type"].value_counts().reset_index()
            by_type.columns = ["Issue Type", "Count"]
            fig = px.bar(by_type, x="Count", y="Issue Type", orientation="h", title="Issues by type",
                         color_discrete_sequence=[COLOR_WARN])
            fig.update_layout(yaxis={"categoryorder": "total ascending"})
            show_chart(fig, "val_types")

    st.markdown("##### Detailed issue list")
    if issues.empty:
        st.info("Nothing to show.")
    else:
        chosen = st.multiselect("Filter by issue type", sorted(issues["Issue Type"].unique()),
                                default=sorted(issues["Issue Type"].unique()))
        view = issues[issues["Issue Type"].isin(chosen)]
        _stretch(st.dataframe, view, hide_index=True, height=360)
        st.download_button("⬇️ Download issue list (CSV)", issues.to_csv(index=False).encode("utf-8"),
                           "validation_issues.csv", "text/csv")
    with st.expander("Preview cleaned data"):
        _stretch(st.dataframe, ctx["attendance"].head(200), hide_index=True)


# ----- Page: Analytics -----
def page_analytics(ctx):
    page_header("Attendance Analytics", "Individual, department-wise, section-wise and subject-wise percentages")
    if not require_data(ctx):
        return
    st.info(f"**Formula:** Attendance % = (Eligible attended sessions ÷ Total eligible sessions) × 100.  "
            f"Holidays are excluded.  OD rule: **{ctx['od_mode']}**.  Threshold: **{ctx['threshold']:g}%**.")
    ss = ctx["student_summary"]
    t1, t2, t3, t4 = st.tabs(["👤 Student-wise", "🏢 Department-wise", "🧩 Section-wise", "📚 Subject-wise"])
    with t1:
        f1, f2, f3 = st.columns(3)
        dept = f1.multiselect("Department", sorted(ss["department"].unique()))
        sec = f2.multiselect("Section", sorted(ss["section"].unique()))
        text = f3.text_input("Search ID or name")
        view = ss
        if dept:
            view = view[view["department"].isin(dept)]
        if sec:
            view = view[view["section"].isin(sec)]
        if text:
            m = view["student_id"].str.contains(text, case=False) | view["student_name"].str.contains(text, case=False)
            view = view[m]
        show_df(view[STUDENT_VIEW_COLS], height=480)
        st.download_button("⬇️ Download this table (CSV)", pretty(view[STUDENT_VIEW_COLS]).to_csv(index=False).encode("utf-8"),
                           "student_attendance.csv", "text/csv")
    with t2:
        show_df(ctx["dept_summary"])
    with t3:
        sec_df = ctx["section_summary"]
        show_df(sec_df)
        if not sec_df.empty:
            plot = sec_df.assign(group=sec_df["department"] + " - " + sec_df["section"])
            fig = px.bar(plot, x="group", y="attendance_pct", title="Section-wise attendance %",
                         labels={"group": "Department - Section", "attendance_pct": "Attendance %"},
                         color_discrete_sequence=["#4f46e5"])
            fig.add_hline(y=ctx["threshold"], line_dash="dash", line_color=COLOR_BAD)
            show_chart(fig, "ana_section")
    with t4:
        show_df(ctx["subject_summary"])


# ----- Page: Shortage -----
def page_shortage(ctx):
    page_header("Shortage Detection", f"Students below the {ctx.get('threshold', DEFAULT_THRESHOLD):g}% attendance requirement")
    if not require_data(ctx):
        return
    sh = ctx["shortage"]
    if sh.empty:
        st.success("🎉 No student is below the threshold with the current settings.")
        if ctx["overall"]["no_data"]:
            st.caption(f"{ctx['overall']['no_data']} student(s) have no eligible sessions and are not evaluated.")
        return
    st.error(f"{len(sh)} student(s) are below {ctx['threshold']:g}% attendance.")
    f1, f2, f3 = st.columns(3)
    dept = f1.multiselect("Department", sorted(sh["department"].unique()), key="sh_dept")
    sec = f2.multiselect("Section", sorted(sh["section"].unique()), key="sh_sec")
    text = f3.text_input("Search ID or name", key="sh_text")
    view = sh
    if dept:
        view = view[view["department"].isin(dept)]
    if sec:
        view = view[view["section"].isin(sec)]
    if text:
        view = view[view["student_id"].str.contains(text, case=False) | view["student_name"].str.contains(text, case=False)]
    st.caption(f"Showing {len(view)} of {len(sh)} students (lowest attendance first).")
    show_df(view, height=420)
    d1, d2 = st.columns(2)
    d1.download_button("⬇️ Download shortage list (CSV)", pretty(view).to_csv(index=False).encode("utf-8"),
                       "attendance_shortage_list.csv", "text/csv")
    d2.download_button("⬇️ Download shortage list (Excel)",
                       df_to_xlsx_bytes(pretty(view), "Shortage List", ("Attendance %", "Required %", "Shortage %")),
                       "attendance_shortage_list.xlsx", XLSX_MIME)
    if not view.empty:
        top = view.head(15).assign(label=lambda d: d["student_id"] + " " + d["student_name"])
        fig = px.bar(top, x="attendance_pct", y="label", orientation="h", title="Lowest attendance (top 15)",
                     labels={"attendance_pct": "Attendance %", "label": ""}, color_discrete_sequence=[COLOR_BAD])
        fig.add_vline(x=ctx["threshold"], line_dash="dash", line_color="#475569")
        fig.update_layout(yaxis={"categoryorder": "total descending"}, xaxis_range=[0, 100])
        show_chart(fig, "short_bar")


# ----- Page: Reports -----
def page_reports(ctx):
    page_header("Report Generation", "Download the complete formatted Excel workbook")
    if not require_data(ctx):
        return
    st.markdown("The workbook contains **5 sheets**: Overall Report · Student Summary · Shortage List · "
                "Department Summary · Data Validation Report. It uses bold headings, filters, percentage "
                "formatting and red/green conditional formatting.")
    signature = (st.session_state.data_version, ctx["threshold"], ctx["od_mode"])
    if st.button("⚙️ Generate Excel report", type="primary"):
        try:
            st.session_state.report_bytes = build_report_workbook(ctx)
            st.session_state.report_sig = signature
            log.info("Excel report generated (%d students, threshold %.1f%%)",
                     ctx["overall"]["total_students"], ctx["threshold"])
            st.toast("Report generated", icon="✅")
        except Exception as exc:
            log.exception("Report generation failed")
            st.error(f"Could not generate the report: {exc}")
    if st.session_state.report_bytes and st.session_state.get("report_sig") == signature:
        st.success("Report is ready.")
        st.download_button("⬇️ Download attendance_report.xlsx", st.session_state.report_bytes,
                           f"attendance_report_{datetime.now():%Y%m%d_%H%M}.xlsx", XLSX_MIME)
    elif st.session_state.report_bytes:
        st.info("Data or settings changed since the last report. Click **Generate Excel report** again.")


# ----- Page: Testing -----
def page_testing(ctx):
    page_header("Testing", "12 test cases for the cleaning and calculation logic")
    st.warning("**Predefined** = what each test is designed to check and the result we expect. "
               "**Executed** = what actually happened when you clicked the button below. "
               "Nothing is marked as passed until the tests are really run.")
    results = st.session_state.test_results
    if st.button("▶️ Run all tests now", type="primary"):
        with st.spinner("Running tests..."):
            st.session_state.test_results = results = run_all_tests()
    if results is None:
        st.info("Tests have **not been executed** in this session yet.")
        st.markdown("##### Predefined test cases (expected results only)")
        _stretch(st.dataframe, predefined_tests_df(), hide_index=True)
    else:
        passed = int((results["Result"] == "PASSED").sum())
        failed = len(results) - passed
        (st.success if failed == 0 else st.error)(
            f"Executed {len(results)} tests at {results['Executed At'].iloc[0]}: {passed} passed, {failed} failed.")
        _stretch(st.dataframe, results, hide_index=True)
        with st.expander("Predefined test cases (inputs and expected results)"):
            _stretch(st.dataframe, predefined_tests_df(), hide_index=True)
    st.download_button("⬇️ Download test workbook (Excel)", build_test_workbook(results),
                       "attendance_test_workbook.xlsx", XLSX_MIME)
    st.caption("The workbook says 'NOT EXECUTED' unless the tests were run before you download it.")


# ----- Page: Logs -----
def page_logs(ctx):
    page_header("Activity Logs", "Startup, uploads, cleaning, validation errors, calculations, reports and errors")
    handler = get_log_handler()
    lines = handler.records
    c1, c2 = st.columns([3, 1])
    c1.caption(f"{len(lines)} log entries this session (newest at the bottom). Also saved to {LOG_FILE} when possible.")
    c2.download_button("⬇️ Download activity log", "\n".join(lines).encode("utf-8"),
                       "attendance_activity_log.txt", "text/plain")
    log_text = "\n".join(lines[-300:]) if lines else "No activity recorded yet."
    try:
        st.code(log_text, language="text", height=520)
    except TypeError:  # older Streamlit versions have no 'height' option
        st.code(log_text, language="text")


ROUTES = {
    "📊 Dashboard": page_dashboard, "📁 Upload Data": page_upload,
    "🧹 Validation & Cleaning": page_validation, "📈 Attendance Analytics": page_analytics,
    "🚨 Shortage Detection": page_shortage, "📄 Reports": page_reports,
    "🧪 Testing": page_testing, "📝 Activity Logs": page_logs,
}


# ---------------------------------------------------------------------
# 9. MAIN
# ---------------------------------------------------------------------
def main():
    st.set_page_config(page_title=APP_TITLE, page_icon="🎓", layout="wide")
    add_logo()
    st.markdown(CSS, unsafe_allow_html=True)
    get_log_handler()
    init_state()
    page = render_sidebar()

    if "flash" in st.session_state:
        kind, message = st.session_state.pop("flash")
        getattr(st, kind)(message)

    ctx = cached_analyze(st.session_state.students_raw, st.session_state.att_raw,
                         st.session_state.od_mode, float(st.session_state.threshold))
    try:
        ROUTES[page](ctx)
    except Exception as exc:  # show a friendly message instead of a crash
        log.exception("Unexpected error on page %s", page)
        st.error(f"Something went wrong on this page: {exc}. The error was written to the Activity Logs.")


if __name__ == "__main__":
    main()
