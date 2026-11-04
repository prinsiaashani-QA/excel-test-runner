
from core.rules import clean, parse_steps, check_locator


def test_clean_none():
    assert clean(None) == ""


def test_clean_integer():
    assert clean(123.0) == "123"


def test_clean_text():
    assert clean("  Login  ") == "Login"


def test_parse_valid_steps():
    text = """
    1. open | https://example.com
    2. verify_visible | role=heading:Welcome
    """

    steps, errors = parse_steps(text)

    assert len(steps) == 2
    assert errors == []
    assert steps[0]["action"] == "open"
    assert steps[1]["action"] == "verify_visible"


def test_parse_invalid_step():
    text = "Invalid step format"

    steps, errors = parse_steps(text)

    assert len(errors) == 1


def test_valid_locator():
    assert check_locator("placeholder=Username") is None


def test_invalid_locator():
    result = check_locator("wronglocator=Username")

    assert result is not None


def test_empty_locator_value():
    result = check_locator("css=")

    assert result is not None
