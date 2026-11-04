"""Excel Test Runner - Streamlit UI (single file).   Start with:  python -m streamlit run app.py
Needs only core/rules.py and core/validator.py next to it."""
import base64
import html
import io
import json
import os
import re
import subprocess
import sys
import zipfile
from datetime import datetime
from pathlib import Path
from urllib.parse import urlparse

import streamlit as st
from openpyxl import load_workbook

from core.rules import BROWSERS, read_testcases
from core.validator import validate

ROOT = Path(__file__).resolve().parent
OUTPUT = ROOT / "output"

# ---------------------------------------------------------------------
# Start run.py in a SEPARATE process (safe for Streamlit + Playwright on Windows)
# ---------------------------------------------------------------------
PROGRESS_RE = re.compile(r"^\[(\d+)/(\d+)\]\s+(TC-\d+)\s+(PASS|FAIL)")


def build_command(excel_path, cfg):
    """cfg keys: url, browser, timeout, headed, video, screenshot, username, password, only."""
    cmd = [sys.executable, "run.py", str(excel_path),
           f"--url={cfg['url']}", f"--browser={cfg['browser']}", f"--timeout={cfg['timeout']}"]
    if cfg.get("headed"):
        cmd.append("--headed")
    if not cfg.get("video", True):
        cmd.append("--no-video")
    if not cfg.get("screenshot", True):
        cmd.append("--no-screenshot")
    if cfg.get("username"):
        cmd.append(f"--username={cfg['username']}")
    if cfg.get("password"):
        cmd.append(f"--password={cfg['password']}")
    if cfg.get("only"):
        cmd.append(f"--only={','.join(cfg['only'])}")
    return cmd


def stream(cmd, on_line):
    """Run cmd, call on_line(text) for every output line.
    Returns (exit_code, run_dir or None). Exit code: 0 pass, 1 some failed, 2 Excel/setup problem."""
    env = {**os.environ, "PYTHONIOENCODING": "utf-8", "PYTHONUTF8": "1"}
    proc = subprocess.Popen(
        cmd, cwd=ROOT, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
        text=True, encoding="utf-8", errors="replace", env=env,
    )
    run_dir = None
    for line in proc.stdout:
        line = line.rstrip()
        m = re.match(r"Result Excel:\s*(.+)", line)
        if m:
            run_dir = (ROOT / m.group(1).strip()).parent
        on_line(line)
    return proc.wait(), run_dir


# ---------------------------------------------------------------------
# Read finished runs, build HTML report and evidence zip
# ---------------------------------------------------------------------
def list_runs():
    """All finished runs, newest first."""
    if not OUTPUT.exists():
        return []
    runs = [d for d in OUTPUT.iterdir()
            if d.is_dir() and d.name.startswith("run_") and (d / "results.xlsx").exists()]
    return sorted(runs, reverse=True)


def run_date(run_dir):
    try:
        return datetime.strptime(Path(run_dir).name[4:], "%Y%m%d_%H%M%S").strftime("%d-%m-%Y %H:%M")
    except ValueError:
        return "-"


def load_run(run_dir):
    """Returns (testcase rows, results by TC ID, run info dict)."""
    run_dir = Path(run_dir)
    wb = load_workbook(run_dir / "results.xlsx")
    _, tcs, _ = read_testcases(wb["TestCases"])
    res = {}
    if "Results" in wb.sheetnames:
        for r in wb["Results"].iter_rows(min_row=2, values_only=True):
            if r and r[0]:
                res[str(r[0])] = {
                    "status": r[1] or "", "failed": r[2] or "", "error": r[3] or "",
                    "shot": r[4] or "", "video": r[5] or "", "secs": r[6] or 0,
                }
    info = {}
    f = run_dir / "run_info.json"
    if f.exists():
        try:
            info = json.loads(f.read_text(encoding="utf-8"))
        except ValueError:
            info = {}
    return tcs, res, info


def summarize(tcs, res):
    rows = []
    for t in tcs:
        r = res.get(t["TC ID"])
        if not r:
            continue
        rows.append({
            "id": t["TC ID"], "module": t["Module"], "scenario": t["Test Scenario"],
            "expected": t["Expected Result"], "actual": t["Actual Result"],
            "status": r["status"], "failed": r["failed"], "error": r["error"],
            "shot": r["shot"], "video": r["video"], "secs": r["secs"],
        })
    total = len(rows)
    passed = sum(1 for x in rows if x["status"] == "Pass")
    return {
        "rows": rows, "total": total, "passed": passed, "failed": total - passed,
        "pct": round(passed * 100 / total) if total else 0,
        "secs": round(sum(float(x["secs"] or 0) for x in rows), 1),
    }


def evidence_file(run_dir, kind, path_text):
    """kind = 'screenshots' or 'videos'. Returns a Path if the file exists, else None."""
    if not path_text:
        return None
    name = Path(str(path_text).replace("\\", "/")).name
    p = Path(run_dir) / kind / name
    return p if p.exists() else None


def evidence_zip(run_dir):
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for sub in ("screenshots", "videos"):
            d = Path(run_dir) / sub
            if d.exists():
                for f in sorted(d.iterdir()):
                    z.write(f, f"{sub}/{f.name}")
    return buf.getvalue()


_TEMPLATE = """<!DOCTYPE html><html><head><meta charset="utf-8"><title>Test Execution Report</title>
<style>
body{font-family:Segoe UI,Arial,sans-serif;margin:24px;color:#222}
h1{margin-bottom:4px}.meta{color:#555;margin-bottom:16px}
.cards{display:flex;gap:12px;margin:16px 0}
.card{border:1px solid #ddd;border-radius:8px;padding:10px 18px;min-width:90px}
.card b{display:block;font-size:24px}
table{border-collapse:collapse;width:100%}
th,td{border:1px solid #ddd;padding:6px 8px;text-align:left;vertical-align:top}
th{background:#305496;color:#fff}
.pass{color:#0a7d2c;font-weight:bold}.fail{color:#c0221b;font-weight:bold}
</style></head><body>
<h1>Test Execution Report</h1>
<div class="meta">__META__</div>
<div class="cards">__CARDS__</div>
<h2>Detailed results</h2>
<table><tr><th>TC ID</th><th>Module</th><th>Test Scenario</th><th>Expected Result</th><th>Actual Result</th><th>Status</th></tr>__ROWS__</table>
__FAILS__
</body></html>"""


def build_html(run_dir, summary, info):
    esc = lambda v: html.escape(str(v if v is not None else ""))
    meta = " | ".join(esc(f"{k}: {v}") for k, v in [
        ("Website", info.get("website", "-")), ("Browser", info.get("browser", "-")),
        ("Mode", info.get("mode", "-")), ("Date", run_date(run_dir)),
    ])
    cards = "".join(f'<div class="card"><b>{v}</b>{k}</div>' for k, v in [
        ("Total", summary["total"]), ("Passed", summary["passed"]), ("Failed", summary["failed"]),
        ("Pass %", f"{summary['pct']}%"), ("Seconds", summary["secs"]),
    ])
    rows = "".join(
        f"<tr><td>{esc(x['id'])}</td><td>{esc(x['module'])}</td><td>{esc(x['scenario'])}</td>"
        f"<td>{esc(x['expected'])}</td><td>{esc(x['actual'])}</td>"
        f"<td class=\"{'pass' if x['status'] == 'Pass' else 'fail'}\">{esc(x['status'])}</td></tr>"
        for x in summary["rows"])
    fails = []
    for x in summary["rows"]:
        if x["status"] == "Pass":
            continue
        part = f"<h3>{esc(x['id'])} - {esc(x['scenario'])}</h3><p><b>{esc(x['failed'])}</b>: {esc(x['error'])}</p>"
        vid = evidence_file(run_dir, "videos", x["video"])
        if vid:
            part += f"<p>Video: {esc(vid.name)} (see the evidence zip)</p>"
        img = evidence_file(run_dir, "screenshots", x["shot"])
        if img:
            b64 = base64.b64encode(img.read_bytes()).decode()
            part += f'<img src="data:image/png;base64,{b64}" style="max-width:100%;border:1px solid #ccc">'
        fails.append(part)
    fail_block = "<h2>Failed test cases</h2>" + "".join(fails) if fails else ""
    return (_TEMPLATE.replace("__META__", meta).replace("__CARDS__", cards)
            .replace("__ROWS__", rows).replace("__FAILS__", fail_block))


# ---------------------------------------------------------------------
# The web pages
# ---------------------------------------------------------------------
UPLOADS = ROOT / "uploads"
TEMPLATE = ROOT / "input" / "test_cases.xlsx"
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
RUN_ALL, RUN_SELECTED = "Run all test cases", "Run selected test cases"

st.set_page_config(page_title="Excel Test Runner", layout="wide")


def valid_url(u):
    p = urlparse(u.strip())
    return p.scheme in ("http", "https") and bool(p.netloc)


def template_bytes():
    if not TEMPLATE.exists():
        subprocess.run([sys.executable, "make_template.py"], cwd=ROOT, check=False)
    return TEMPLATE.read_bytes() if TEMPLATE.exists() else None


def open_run(name):
    """Button callback: jump to the Test Report page for this run."""
    st.session_state["selected_run"] = name
    st.session_state["page"] = "Test Report"


# ---------------- sidebar ----------------
st.sidebar.title("Excel Test Runner")
page = st.sidebar.radio("Navigation", ["New Test Run", "Test Report", "Run History"], key="page")
st.sidebar.caption("Keyword-driven tests with Playwright. Results are written back to Excel.")

# =====================================================================
# PAGE 1: NEW TEST RUN
# =====================================================================
if page == "New Test Run":
    st.title("New Test Run")

    st.subheader("A. Website details")
    url = st.text_input("Website URL", placeholder="https://www.saucedemo.com", key="url")
    login = st.radio("Login required?", ["No", "Yes"], horizontal=True, key="login_required")
    username = password = ""
    if login == "Yes":
        c1, c2 = st.columns(2)
        username = c1.text_input("Username / Email", key="username")
        password = c2.text_input("Password", type="password", key="password")
        st.caption("In your Excel steps write {{USERNAME}} and {{PASSWORD}}. "
                   "The password is never saved in the Excel or in the report.")

    st.subheader("B. Browser configuration")
    b1, b2, b3 = st.columns(3)
    browser = b1.selectbox("Browser", BROWSERS, key="browser")
    view_mode = b2.radio("Mode", ["Headless", "Headed"], horizontal=True, key="view_mode")
    timeout = b3.number_input("Wait time (seconds)", min_value=1, max_value=120, value=10, key="timeout")
    d1, d2 = st.columns(2)
    video = d1.checkbox("Record video on failure", value=True, key="video")
    screenshot = d2.checkbox("Screenshot on failure", value=True, key="screenshot")
    if browser != "chromium":
        st.caption(f"First time only: run  playwright install {browser}")

    st.subheader("C. Upload test cases")
    tpl = template_bytes()
    if tpl:
        st.download_button("Download sample Excel template", tpl,
                           file_name="test_cases_template.xlsx", mime=XLSX_MIME)
    up = st.file_uploader("Upload Excel (.xlsx)", type=["xlsx"], key="upload")

    errors, tcs, excel_path = [], [], None
    if up is not None:
        UPLOADS.mkdir(exist_ok=True)
        excel_path = UPLOADS / "current_upload.xlsx"
        excel_path.write_bytes(up.getvalue())
        errors = validate(str(excel_path))
        if errors:
            st.error(f"The Excel file has {len(errors)} problem(s). Nothing will run until they are fixed:")
            st.code("\n".join(errors), language="text")
        else:
            _, tcs, _ = read_testcases(load_workbook(excel_path)["TestCases"])
            st.success(f"Excel is valid. {len(tcs)} test case(s) found.")
            st.dataframe([{"TC ID": t["TC ID"], "Module": t["Module"],
                           "Test Scenario": t["Test Scenario"], "Priority": t["Priority"]}
                          for t in tcs])

    st.subheader("D. Execution settings")
    run_mode = st.radio("What to run?", [RUN_ALL, RUN_SELECTED], horizontal=True, key="run_mode")
    selected = []
    if run_mode == RUN_SELECTED:
        selected = st.multiselect("Select test cases", [t["TC ID"] for t in tcs], key="selected")

    problems = []
    if not valid_url(url):
        problems.append("enter a valid website URL (https://...)")
    if up is None:
        problems.append("upload an Excel file")
    elif errors:
        problems.append("fix the Excel problems shown above")
    if run_mode == RUN_SELECTED and not selected:
        problems.append("select at least one test case")
    if login == "Yes" and not (username and password):
        problems.append("enter username and password")
    if problems:
        st.info("Before you start: " + "; ".join(problems) + ".")

    if st.button("Start Test Execution", type="primary", disabled=bool(problems)):
        cmd = build_command(excel_path, {
            "url": url.strip().rstrip("/"), "browser": browser, "timeout": timeout,
            "headed": view_mode == "Headed", "video": video, "screenshot": screenshot,
            "username": username, "password": password, "only": selected,
        })
        started = datetime.now()
        st.write(f"**Started:** {started.strftime('%d-%m-%Y %H:%M:%S')}")
        bar = st.progress(0.0, text="Starting...")
        box = st.empty()
        lines = []

        def on_line(line):
            lines.append(line)
            m = PROGRESS_RE.match(line)
            if m:
                done, tot = int(m.group(1)), int(m.group(2))
                bar.progress(done / tot, text=f"{done} of {tot} test cases completed")
            box.code("\n".join(lines[-25:]), language="text")

        code, run_dir = stream(cmd, on_line)
        ended = datetime.now()
        st.write(f"**Ended:** {ended.strftime('%d-%m-%Y %H:%M:%S')}  "
                 f"(took {round((ended - started).total_seconds())} s)")

        if run_dir and run_dir.exists():
            info = {"website": url.strip(), "browser": browser, "mode": view_mode,
                    "started": started.strftime("%d-%m-%Y %H:%M:%S"), "login": login}
            (run_dir / "run_info.json").write_text(json.dumps(info), encoding="utf-8")
            st.button("Open Test Report", on_click=open_run, args=(run_dir.name,))
            if code == 0 and run_dir:
                st.success("All test cases passed.")
            elif code == 1 and run_dir:
                st.warning("Some test cases failed. Open the Test Report for screenshots and video.")
            else:
                st.error("The run did not finish. Read the log above (usually an Excel or setup problem).")

# =====================================================================
# PAGE 2: TEST REPORT
# =====================================================================
elif page == "Test Report":
    st.title("Test Report")
    runs = list_runs()
    if not runs:
        st.info("No runs yet. Start one from 'New Test Run'.")
        st.stop()

    names = [r.name for r in runs]
    wanted = st.session_state.get("selected_run")
    choice = st.selectbox("Select a run", names,
                          index=names.index(wanted) if wanted in names else 0,
                          format_func=lambda n: f"{n}   ({run_date(n)})")
    run_dir = OUTPUT / choice
    try:
        tcs, res, info = load_run(run_dir)
    except Exception as e:
        st.error(f"This run folder cannot be read (maybe an old format): {e}")
        st.stop()
    s = summarize(tcs, res)

    st.caption(f"Website: {info.get('website', '-')}  |  Browser: {info.get('browser', '-')}  |  "
               f"Mode: {info.get('mode', '-')}  |  Date: {run_date(run_dir)}")

    m = st.columns(5)
    m[0].metric("Total", s["total"])
    m[1].metric("Passed", s["passed"])
    m[2].metric("Failed", s["failed"])
    m[3].metric("Pass %", f"{s['pct']}%")
    m[4].metric("Duration (s)", s["secs"])

    st.subheader("Detailed results")
    st.dataframe([{"TC ID": x["id"], "Test Scenario": x["scenario"], "Expected": x["expected"],
                   "Actual": x["actual"], "Status": x["status"]} for x in s["rows"]])

    st.subheader("Failed test cases")
    fails = [x for x in s["rows"] if x["status"] != "Pass"]
    if not fails:
        st.success("No failures in this run.")
    for x in fails:
        with st.expander(f"{x['id']} - {x['scenario']}", expanded=True):
            st.write(f"**{x['failed']}**")
            st.code(x["error"] or "-", language="text")
            img = evidence_file(run_dir, "screenshots", x["shot"])
            if img:
                st.image(str(img), caption="Screenshot at the failed step")
            vid = evidence_file(run_dir, "videos", x["video"])
            if vid:
                st.video(str(vid))

    st.subheader("Downloads")
    d = st.columns(3)
    d[0].download_button("Excel report", (run_dir / "results.xlsx").read_bytes(),
                         file_name=f"{choice}_results.xlsx", mime=XLSX_MIME)
    d[1].download_button("HTML report", build_html(run_dir, s, info).encode("utf-8"),
                         file_name=f"{choice}_report.html", mime="text/html")
    d[2].download_button("Screenshots and videos (zip)", evidence_zip(run_dir),
                         file_name=f"{choice}_evidence.zip", mime="application/zip")

# =====================================================================
# PAGE 3: RUN HISTORY
# =====================================================================
else:
    st.title("Run History")
    runs = list_runs()
    if not runs:
        st.info("No runs yet.")
        st.stop()

    table = []
    for r in runs:
        try:
            tcs, res, info = load_run(r)
        except Exception:
            continue
        s = summarize(tcs, res)
        table.append({"Run": r.name, "Website": info.get("website", "-"), "Date": run_date(r),
                      "Total": s["total"], "Passed": s["passed"], "Failed": s["failed"],
                      "Status": "Passed" if s["total"] and s["failed"] == 0 else "Failed"})
    st.dataframe(table)

    if not table:
        st.info("No readable runs yet.")
        st.stop()
    pick = st.selectbox("Open an old report", [t["Run"] for t in table])
    st.button("Open in Test Report", on_click=open_run, args=(pick,))