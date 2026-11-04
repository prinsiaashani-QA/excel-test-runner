
from openpyxl import Workbook
from core.validator import validate
from core.rules import TC_COLUMNS


def create_excel(path, test_steps):
    wb = Workbook()
    ws = wb.active
    ws.title = "TestCases"

    ws.append(TC_COLUMNS)

    ws.append([
        "TC-001",
        "Login",
        "Valid Login",
        "Website is available",
        test_steps,
        "Username: test",
        "Dashboard should appear",
        "",
        "High",
        ""
    ])

    wb.save(path)


def test_valid_excel(tmp_path):
    file_path = tmp_path / "valid.xlsx"

    create_excel(
        file_path,
        "1. open | https://example.com\n"
        "2. verify_visible | role=heading:Welcome"
    )

    errors = validate(file_path)

    assert errors == []


def test_invalid_excel_file():
    errors = validate("file_does_not_exist.xlsx")

    assert len(errors) > 0
    assert "Cannot open the file" in errors[0]


def test_unknown_action(tmp_path):
    file_path = tmp_path / "invalid.xlsx"

    create_excel(
        file_path,
        "1. unknown_action | something"
    )

    errors = validate(file_path)

    assert any("unknown action" in e for e in errors)


def test_missing_verify_step(tmp_path):
    file_path = tmp_path / "no_verify.xlsx"

    create_excel(
        file_path,
        "1. open | https://example.com"
    )

    errors = validate(file_path)

    assert any("no verify_" in e for e in errors)
