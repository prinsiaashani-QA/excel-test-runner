"""Turns one Excel step (action + locator + value) into a Playwright call."""
import re
from playwright.sync_api import expect


def find(page, loc):
    strat, _, val = loc.partition("=")
    strat, val = strat.strip().lower(), val.strip()
    if strat == "placeholder":
        return page.get_by_placeholder(val)
    if strat == "label":
        return page.get_by_label(val)
    if strat == "text":
        return page.get_by_text(val)
    if strat == "testid":
        return page.get_by_test_id(val)      # uses the data-testid attribute
    if strat == "css":
        return page.locator(val)
    if strat == "xpath":
        return page.locator("xpath=" + val)
    if strat == "role":
        role, _, name = val.partition(":")
        role, name = role.strip(), name.strip()
        return page.get_by_role(role, name=name) if name else page.get_by_role(role)
    raise ValueError(f"Unknown locator strategy: {strat}")


def perform(page, action, loc, val, base_url, ms):
    if action == "open":
        url = val or base_url
        if url.startswith("/"):
            url = base_url.rstrip("/") + url
        page.goto(url)
    elif action == "press":
        page.keyboard.press(val)
    elif action == "wait":
        page.wait_for_timeout(float(val) * 1000)
    elif action == "verify_url":
        expect(page).to_have_url(re.compile(re.escape(val)), timeout=ms)
    else:
        el = find(page, loc)
        if action == "type":
            el.fill(val)
        elif action == "clear":
            el.fill("")
        elif action == "click":
            el.click()
        elif action == "select":
            el.select_option(val)
        elif action == "verify_text":
            expect(el).to_have_text(val, timeout=ms)
        elif action == "verify_contains":
            expect(el).to_contain_text(val, timeout=ms)
        elif action == "verify_visible":
            expect(el).to_be_visible(timeout=ms)
        elif action == "verify_hidden":
            expect(el).to_be_hidden(timeout=ms)
        elif action == "verify_count":
            expect(el).to_have_count(int(val), timeout=ms)
        else:
            raise ValueError(f"Unknown action: {action}")
