"""Runs every test case from the Excel file, writes results back, saves evidence."""
import os
import shutil
import time
from datetime import datetime
from pathlib import Path
from openpyxl import load_workbook
from playwright.sync_api import sync_playwright
from core.rules import read_testcases, parse_steps
from core.validator import validate
from core.actions import perform


def short(e):
    text = str(e).strip()
    return text.splitlines()[0][:200] if text else type(e).__name__


def run(excel_path, cfg):
    """Returns exit code: 0 = all passed, 1 = some failed, 2 = Excel/setup problem."""
    errors = validate(excel_path)
    if errors:
        print(f"\n[ERROR] The Excel file has {len(errors)} problem(s). Fix them first:\n")
        for e in errors:
            print("  -", e)
        return 2

    wb = load_workbook(excel_path)
    ws = wb["TestCases"]
    headers, tcs, _ = read_testcases(ws)
    actual_col = headers.index("Actual Result") + 1
    status_col = headers.index("Status") + 1

    if cfg["only"]:
        tcs = [t for t in tcs if t["TC ID"] in cfg["only"]]
        if not tcs:
            print("[ERROR] None of the TC IDs given in --only exist in the Excel file")
            return 2

    all_values = " ".join(s["value"] for t in tcs for s in parse_steps(t["Test Steps"])[0])
    if "{{USERNAME}}" in all_values and not cfg["username"]:
        print("[ERROR] Steps use {{USERNAME}} but --username was not given")
        return 2
    if "{{PASSWORD}}" in all_values and not cfg["password"]:
        print("[ERROR] Steps use {{PASSWORD}} but --password was not given")
        return 2

    pwd = cfg["password"]
    ms = int(cfg["timeout"] * 1000)

    def mask(text):
        return text.replace(pwd, "****") if pwd else text

    def fill(v):
        return v.replace("{{USERNAME}}", cfg["username"]).replace("{{PASSWORD}}", pwd)

    run_dir = Path("output") / datetime.now().strftime("run_%Y%m%d_%H%M%S")
    shots, vids, tmp = run_dir / "screenshots", run_dir / "videos", run_dir / "_tmp"
    shots.mkdir(parents=True)
    vids.mkdir(parents=True)

    if "Results" in wb.sheetnames:
        del wb["Results"]
    ws_res = wb.create_sheet("Results")
    ws_res.append(["TC ID", "Status", "Failed Step", "Error", "Screenshot", "Video", "Seconds"])

    count = {"Pass": 0, "Fail": 0}
    results = []
    print(f"\nRunning {len(tcs)} test case(s) on {cfg['url']} with {cfg['browser']}\n")

    with sync_playwright() as p:
        browser = getattr(p, cfg["browser"]).launch(headless=cfg["headless"])
        for i, tc in enumerate(tcs, 1):
            tid = tc["TC ID"]
            steps, _ = parse_steps(tc["Test Steps"])
            steps.sort(key=lambda s: s["no"])

            ctx = browser.new_context(
                viewport={"width": 1280, "height": 720},
                record_video_dir=str(tmp) if cfg["video"] else None,
            )
            page = ctx.new_page()
            page.set_default_timeout(ms)
            video = page.video

            status, failed, error, shot_path, vid_path = "Pass", "", "", "", ""
            t0 = time.time()
            for s in steps:
                try:
                    perform(page, s["action"], s["locator"], fill(s["value"]), cfg["url"], ms)
                except Exception as e:
                    status = "Fail"
                    failed = f"Step {s['no']} ({s['action']})"
                    error = mask(short(e))
                    if cfg["screenshot"]:
                        try:
                            shot_path = str(shots / f"{tid}_step{s['no']}.png")
                            page.screenshot(path=shot_path, full_page=True)
                        except Exception:
                            shot_path = ""
                    break
            ctx.close()

            if cfg["video"] and video:
                src = video.path()
                if status == "Fail":
                    vid_path = str(vids / f"{tid}.webm")
                    shutil.move(src, vid_path)
                else:
                    os.remove(src)

            actual = (f"As expected: {tc['Expected Result']}" if status == "Pass"
                      else f"{failed}: {error}")
            ws.cell(row=tc["_row"], column=actual_col, value=actual)
            ws.cell(row=tc["_row"], column=status_col, value=status)
            ws_res.append([tid, status, failed, error, shot_path, vid_path,
                           round(time.time() - t0, 1)])

            count[status] += 1
            results.append((tid, status))
            line = f"[{i}/{len(tcs)}] {tid}  {status.upper()}  {tc['Test Scenario']}"
            if status == "Fail":
                line += f"   -> {failed}: {error}"
            print(line)
        browser.close()

    shutil.rmtree(tmp, ignore_errors=True)
    out = run_dir / "results.xlsx"
    wb.save(out)

    print(f"\nTotal Test Cases: {len(tcs)}")
    print(f"Passed: {count['Pass']}")
    print(f"Failed: {count['Fail']}\n")
    for tid, st in results:
        print(f"{tid}: {st.upper()}")
    print(f"\nResult Excel: {out}")
    if count["Fail"]:
        print(f"Screenshots: {shots}")
        print(f"Videos:      {vids}")
    return 1 if count["Fail"] else 0
