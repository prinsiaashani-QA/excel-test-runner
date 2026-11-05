# Excel Test Runner

A lightweight keyword-driven test runner. Write test cases in an Excel sheet, run them with Playwright (Python) on any website, and get **Actual Result + Status** written back to Excel, with a **screenshot and video on every failure**.

Inspired by keyword-driven tools such as Katalon and Robot Framework.

## What it does
1. Validates the whole Excel file first and shows **all** problems together (nothing runs if it is invalid)
2. Runs each test case in a fresh browser (Chromium, Firefox or WebKit)
3. Fills `Actual Result` and `Status` in your sheet
4. On failure saves a screenshot and a video
5. Passwords are masked in error messages
6. Web UI with live progress, a report page (Excel, HTML and evidence zip downloads) and run history

## Web UI (Streamlit)
    pip install -r requirements.txt
    python -m streamlit run app.py

Pages: **New Test Run** (URL, login, browser, upload Excel, run all or selected test cases), **Test Report** and **Run History**.
The tests run in a separate process, so the UI stays responsive.

## Setup
    pip install -r requirements.txt
    playwright install

## Command line
    python make_template.py
    python run.py input/test_cases.xlsx --url https://www.saucedemo.com --username standard_user --password secret_sauce --headed

Expected: 2 passed, 1 failed (TC-003 is a demo that fails on purpose).

| Option | Meaning |
|---|---|
| `--url` | Website URL (required) |
| `--browser` | chromium (default), firefox, webkit |
| `--headed` | Show the browser window |
| `--timeout` | Wait time in seconds (default 10) |
| `--only TC-001,TC-003` | Run only these test cases |
| `--username`, `--password` | Fill `{{USERNAME}}` / `{{PASSWORD}}` in steps |
| `--no-video`, `--no-screenshot` | Switch evidence off |

Exit code: 0 = all passed, 1 = some failed, 2 = Excel problem (useful for CI).

## Excel format (sheet `TestCases`)
Columns: `TC ID | Module | Test Scenario | Pre-conditions | Test Steps | Test Data | Expected Result | Actual Result | Priority | Status`

A title row above the header is allowed. Leave `Actual Result` and `Status` empty; the runner fills them in a new file.

`Test Steps`: one step per line, in the format `N. action | locator | value`

    1. open
    2. type | placeholder=Username | {{USERNAME}}
    3. click | role=button:Login
    4. verify_url | inventory.html
    5. verify_text | css=.title | Products

### Actions
| Action | Locator | Value | Meaning |
|---|---|---|---|
| open | no | optional | Open URL (empty = base URL, `/path` = base URL + path) |
| type | yes | yes | Type in a box |
| clear | yes | no | Empty a box |
| click | yes | no | Click |
| select | yes | yes | Choose a dropdown option |
| press | no | yes | Press a key, e.g. Enter |
| wait | no | yes | Wait (seconds) |
| verify_text | yes | yes | Text equals value |
| verify_contains | yes | yes | Text contains value |
| verify_url | no | yes | URL contains value |
| verify_visible / verify_hidden | yes | no | Element is visible / hidden |
| verify_count | yes | yes | Number of matching elements |

### Locators: `strategy=value`
`placeholder=Username`, `role=button:Login`, `text=Add to cart`, `label=Email`, `testid=login-btn` (data-testid), `css=.title`, `xpath=//div[@id='a']`

Tip: run `playwright codegen https://your-site.com` and click around. It writes the locators for you.

## Output
    output/run_<date_time>/
        results.xlsx      your sheet with Actual Result + Status, plus a Results sheet
        screenshots/      failed steps
        videos/           failed test cases

## Tests and CI
Tests are in the `tests` folder (run with `pytest`). A GitHub Actions workflow in `.github/workflows` runs on every push.

## Not supported yet
iframes, new tabs, file upload, and logging in once for all test cases.
