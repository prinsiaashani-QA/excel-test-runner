"""Checks the Excel file completely BEFORE any test runs. Returns a list of error messages."""
import re
from openpyxl import load_workbook
from core.rules import (TC_COLUMNS, ACTIONS, PRIORITIES,
                        read_testcases, parse_steps, check_locator)


def validate(path):
    errors = []
    try:
        wb = load_workbook(path)
    except Exception as e:
        return [f"Cannot open the file: {e}"]

    if "TestCases" not in wb.sheetnames:
        return ["Sheet 'TestCases' is missing"]

    headers, rows, _ = read_testcases(wb["TestCases"])
    if headers is None:
        return ["TestCases: header row starting with 'TC ID' not found (must be within the first 10 rows)"]
    missing = [c for c in TC_COLUMNS if c not in headers]
    if missing:
        return [f"TestCases: missing columns -> {missing}"]
    if not rows:
        return ["TestCases: no test case rows found"]

    seen = set()
    for r in rows:
        tid = r["TC ID"]
        where = f"Row {r['_row']} ({tid})"

        if not re.fullmatch(r"TC-\d{3,}", tid):
            errors.append(f"Row {r['_row']}: TC ID must look like TC-001 (found '{tid}')")
        elif tid in seen:
            errors.append(f"{where}: duplicate TC ID")
        seen.add(tid)

        for col in ("Module", "Test Scenario", "Expected Result"):
            if not r[col]:
                errors.append(f"{where}: '{col}' is empty")
        if r["Priority"] not in PRIORITIES:
            errors.append(f"{where}: Priority must be one of {PRIORITIES}")

        steps, parse_errors = parse_steps(r["Test Steps"])
        errors += [f"{where}: {e}" for e in parse_errors]
        if not steps and not parse_errors:
            errors.append(f"{where}: 'Test Steps' is empty")

        nums = [s["no"] for s in steps]
        if nums and nums != list(range(1, len(nums) + 1)):
            errors.append(f"{where}: step numbers must be 1, 2, 3 ... in order")

        for s in steps:
            tag = f"{where} step {s['no']}"
            if s["action"] not in ACTIONS:
                errors.append(f"{tag}: unknown action '{s['action']}'. Allowed: {list(ACTIONS)}")
                continue
            need_loc, need_val = ACTIONS[s["action"]]
            if need_loc and not s["locator"]:
                errors.append(f"{tag}: '{s['action']}' needs a locator")
            if need_val and not s["value"]:
                errors.append(f"{tag}: '{s['action']}' needs a value")
            if s["locator"]:
                msg = check_locator(s["locator"])
                if msg:
                    errors.append(f"{tag}: {msg}")
            if s["action"] == "verify_count" and s["value"] and not s["value"].isdigit():
                errors.append(f"{tag}: verify_count value must be a number")
            if s["action"] == "wait" and s["value"]:
                try:
                    float(s["value"])
                except ValueError:
                    errors.append(f"{tag}: wait value must be a number (seconds)")
            if s["action"] == "open" and s["value"] and not (
                    s["value"].startswith("/") or s["value"].startswith("http")):
                errors.append(f"{tag}: open value must be '/path' or 'https://...'")

        if steps and not any(s["action"].startswith("verify") for s in steps):
            errors.append(f"{where}: no verify_ step (a test without a check is useless)")
    return errors
