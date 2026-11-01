# Student Attendance Automation (DIT Internship 2026 – Task 07)

One Python file (`app.py`) = the whole application. No database, no extra folders.

## 1. Architecture (how app.py is organised)

```
Upload / Sample data
        │
        ▼
 clean_data()  ── standardise IDs, dates, statuses
        │          detect duplicates, missing values, invalid dates,
        │          unknown IDs, missing columns  → "issues" table (nothing hidden)
        ▼
 add_flags()   ── per record: eligible? attended?   (Holiday never eligible, OD rule from sidebar)
        ▼
 compute_student_summary() / group_summary()  ── % per student, department, section, subject, month
        ▼
 Streamlit pages:  Dashboard · Upload · Validation · Analytics · Shortage · Reports · Testing · Logs
        ▼
 OpenPyXL → Excel report (5 sheets) · Excel test workbook · CSV downloads
```

Sections inside `app.py`, top to bottom: constants → logging → sample-data generator →
cleaning/validation → analytics → Excel builders → 12 tests → Streamlit pages → `main()`.

**Business rules**

- Attendance % = eligible attended sessions ÷ eligible sessions × 100
- Present = attended, Absent = not attended, Holiday = never counted
- OD is configurable in the sidebar: count as attended / count as not attended / exclude from eligible sessions
- Duplicate (same student + date + subject) → first record kept, extras listed in the validation report
- Zero eligible sessions → shown as "N/A", status "No Data" (never a crash, never flagged as shortage)
- Exactly at the threshold (e.g. 75.00%) is **sufficient**; only strictly below is a shortage
- Day-first dates are assumed for ambiguous text such as `03/04/2026` (3 April)

## 2. Where to save files

Create ONE folder, for example `C:\Users\<you>\attendance_app` (Windows) or `~/attendance_app` (Mac/Linux), and put these two files directly inside it:

```
attendance_app/
├── app.py
└── requirements.txt
```

(The app also creates `attendance_app.log` there when it runs.)

## 3. VS Code setup (exact steps)

1. Install Python 3.10+ from https://www.python.org/downloads/ (tick **"Add Python to PATH"** on Windows).
2. Install VS Code from https://code.visualstudio.com/ and add the **Python** extension (Extensions icon → search "Python" → Install).
3. VS Code → **File → Open Folder…** → choose `attendance_app`.
4. Open the terminal: **Terminal → New Terminal**.

## 4. Install dependencies (type in the VS Code terminal)

**Windows (PowerShell):**
```
python -m venv venv
venv\Scripts\Activate.ps1
pip install -r requirements.txt
```
If PowerShell blocks the activate script, run `Set-ExecutionPolicy -Scope CurrentUser RemoteSigned` once, or use Command Prompt and `venv\Scripts\activate.bat`.

**Mac / Linux:**
```
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
```

## 5. Run the application

```
streamlit run app.py
```
Your browser opens at http://localhost:8501 with sample data already loaded. Stop it with `Ctrl + C`.

## 6. How to test every feature

| Page | What to do | What you should see |
|---|---|---|
| 📊 Dashboard | Open it | 5 KPI cards, donut, histogram, department bars, monthly trend, subject chart, department table |
| ⚙️ Sidebar | Change threshold 75 → 85; switch OD handling | Numbers/charts update; shortage count changes |
| 📁 Upload Data | Untick "Include intentional data-quality problems" → **Load sample**; download both CSVs; upload them back → **Process uploaded files** | Success message; data source in sidebar changes. Uploading only one file shows an error. |
| 🧹 Validation & Cleaning | Load the sample *with* problems | Rejected rows (invalid dates, unknown IDs, missing IDs, unknown statuses), duplicates and conflicting duplicates all listed; CSV download of issues |
| 📈 Attendance Analytics | Use filters and the 4 tabs | Student, department, section, subject tables; section chart |
| 🚨 Shortage Detection | Filter by department/section/name; download CSV and Excel | Student ID, name, department, %, required %, shortage % |
| 📄 Reports | **Generate Excel report** → download | 5 sheets, bold headers, filters, % formats, red/green conditional formatting |
| 🧪 Testing | Before clicking: shows "not executed". Click **Run all tests now** | 12 results marked PASSED/FAILED with timestamp; download workbook (says NOT EXECUTED if downloaded before running) |
| 📝 Activity Logs | Open after using other pages | Startup, uploads, cleaning, validation warnings, calculations, report generation, test runs; download log |
| Error handling | Upload a CSV missing a column, or a `.txt` file | Clear red error message, app keeps working |

You can also try your own files: students need `student_id, student_name, department, semester, section`; attendance needs `student_id, attendance_date, subject, attendance_status`.

## 7. Deployment

**Why not Netlify?** Netlify serves static files (HTML/CSS/JS). A Streamlit app is a running Python server, so it cannot be dropped onto Netlify as an ordinary static site. Use a host that runs Python.

**Streamlit Community Cloud (free, simplest)**
1. Create a free GitHub account and a new **public** repository.
2. Upload `app.py` and `requirements.txt` to the repository root (GitHub → *Add file → Upload files*).
3. Go to https://share.streamlit.io, sign in with GitHub, click **Create app**.
4. Choose your repository, branch `main`, main file path `app.py`, then **Deploy**.
5. After a few minutes you get a public URL to share.

Other options: Render, Railway, Hugging Face Spaces (Streamlit template) or any server where you can run `streamlit run app.py`.

Notes: no database or secrets are needed. Uploaded data lives only in the user's session. The activity log is shared by everyone using the same server instance, so don't upload real student data to a public deployment.

## 8. Troubleshooting

- `streamlit is not recognized` → the virtual environment isn't active; run the activate command again, or use `python -m streamlit run app.py`.
- Port busy → `streamlit run app.py --server.port 8502`.
- Dates not recognised → use `YYYY-MM-DD` or `DD/MM/YYYY`.
