"""Shared rules: Excel columns, allowed actions, locator check, step parsing."""
import re

TC_COLUMNS = [
    "TC ID", "Module", "Test Scenario", "Pre-conditions", "Test Steps",
    "Test Data", "Expected Result", "Actual Result", "Priority", "Status",
]

# action: (needs locator?, needs value?)
ACTIONS = {
    "open": (False, False),            # value optional: empty = base URL, /path = base URL + path
    "type": (True, True),
    "clear": (True, False),
    "click": (True, False),
    "select": (True, True),
    "press": (False, True),
    "wait": (False, True),             # seconds
    "verify_text": (True, True),       # text must match exactly
    "verify_contains": (True, True),   # text must contain value
    "verify_url": (False, True),       # URL must contain value
    "verify_visible": (True, False),
    "verify_hidden": (True, False),
    "verify_count": (True, True),
}

STRATEGIES = ["placeholder", "role", "text", "label", "testid", "css", "xpath"]
PRIORITIES = ["High", "Medium", "Low"]
BROWSERS = ["chromium", "firefox", "webkit"]

STEP_RE = re.compile(r"^\s*(\d+)\s*[.)]\s*(.+?)\s*$")


def clean(v):
    """Make any Excel cell value a clean string."""
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v).strip()


def read_testcases(ws):
    """Find the 'TC ID' header row (a title row above is allowed) and read all rows.
    Returns (headers, rows, header_row). headers is None if not found."""
    header_row = None
    for r in range(1, min(ws.max_row, 10) + 1):
        if clean(ws.cell(row=r, column=1).value) == "TC ID":
            header_row = r
            break
    if header_row is None:
        return None, [], None

    headers = [clean(c.value) for c in ws[header_row]]
    rows = []
    for r in range(header_row + 1, ws.max_row + 1):
        values = [ws.cell(row=r, column=c).value for c in range(1, len(headers) + 1)]
        if all(clean(v) == "" for v in values):
            continue
        row = {h: clean(v) for h, v in zip(headers, values)}
        row["_row"] = r
        rows.append(row)
    return headers, rows, header_row


def parse_steps(text):
    """Split the 'Test Steps' cell into steps.
    Format, one step per line:  N. action | locator | value
    Returns (steps, errors)."""
    steps, errors = [], []
    for line in text.splitlines():
        if not line.strip():
            continue
        m = STEP_RE.match(line)
        if not m:
            errors.append(f"line '{line.strip()}' is not in 'N. action | locator | value' format")
            continue
        parts = [p.strip() for p in m.group(2).split("|")]
        action = parts[0].lower()
        locator = value = ""
        if action in ACTIONS:
            if ACTIONS[action][0]:          # action uses a locator
                locator = parts[1] if len(parts) > 1 else ""
                value = parts[2] if len(parts) > 2 else ""
            else:                           # action has only a value
                value = parts[1] if len(parts) > 1 else ""
        steps.append({"no": int(m.group(1)), "action": action,
                      "locator": locator, "value": value})
    return steps, errors


def check_locator(loc):
    """Return an error message if the locator format is wrong, else None."""
    strat, sep, val = loc.partition("=")
    strat = strat.strip().lower()
    if not sep or strat not in STRATEGIES:
        return f"locator '{loc}' has a wrong format. Use strategy=value ({', '.join(STRATEGIES)})"
    if not val.strip():
        return f"locator '{loc}' has an empty value"
    if strat == "role" and not val.partition(":")[0].strip():
        return f"locator '{loc}' should look like role=button:Login"
    return None
