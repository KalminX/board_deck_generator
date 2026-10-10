"""
Automated Playwright Script to Record End-to-End Presentation Generator Flow.
Captures high-fidelity 720p video with visible animated mouse cursor and click effects.
"""
import os
import time
import subprocess
from pathlib import Path
from playwright.sync_api import sync_playwright

BASE_URL = "http://127.0.0.1:8088"
RECORDING_DIR = Path("/Users/kalmin/board_deck_generator/recordings")
RECORDING_DIR.mkdir(exist_ok=True)
SAMPLE_EXCEL_PATH = Path("/Users/kalmin/board_deck_generator/sample_data/sample_board_deck_data.xlsx").resolve()

# Visual mouse cursor injection script
MOUSE_HELPER_JS = """
(() => {
    if (window._mouseHelperInjected) return;
    window._mouseHelperInjected = true;

    // Create cursor element
    const cursor = document.createElement('div');
    cursor.id = 'playwright-mouse-pointer';
    cursor.style.cssText = `
        position: fixed;
        top: 0;
        left: 0;
        width: 22px;
        height: 22px;
        background-image: url('data:image/svg+xml;utf8,<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="%232563EB" stroke="%23FFFFFF" stroke-width="1.5"><polygon points="3 3 10 21 13 13 21 10 3 3"/></svg>');
        background-size: contain;
        background-repeat: no-repeat;
        pointer-events: none;
        z-index: 9999999;
        transform: translate(0, 0);
        transition: transform 0.05s ease-out;
        filter: drop-shadow(0 2px 4px rgba(0,0,0,0.4));
    `;
    document.body.appendChild(cursor);

    // Create click ripple effect
    const clickRipple = document.createElement('div');
    clickRipple.id = 'playwright-click-ripple';
    clickRipple.style.cssText = `
        position: fixed;
        width: 36px;
        height: 36px;
        border-radius: 50%;
        background: rgba(37, 99, 235, 0.4);
        border: 2px solid rgba(37, 99, 235, 0.9);
        pointer-events: none;
        z-index: 9999998;
        transform: translate(-50%, -50%) scale(0);
        opacity: 0;
        transition: transform 0.35s ease-out, opacity 0.35s ease-out;
    `;
    document.body.appendChild(clickRipple);

    window.addEventListener('mousemove', (e) => {
        cursor.style.left = `${e.clientX}px`;
        cursor.style.top = `${e.clientY}px`;
    }, { passive: true });

    window.addEventListener('mousedown', (e) => {
        clickRipple.style.left = `${e.clientX}px`;
        clickRipple.style.top = `${e.clientY}px`;
        clickRipple.style.transform = 'translate(-50%, -50%) scale(1)';
        clickRipple.style.opacity = '1';
        setTimeout(() => {
            clickRipple.style.transform = 'translate(-50%, -50%) scale(1.6)';
            clickRipple.style.opacity = '0';
        }, 150);
    }, { passive: true });
})();
"""


def smooth_move_and_click(page, selector_or_locator, steps=25, sleep_after=0.6):
    if isinstance(selector_or_locator, str):
        locator = page.locator(selector_or_locator).first
    else:
        locator = selector_or_locator

    locator.scroll_into_view_if_needed()
    box = locator.bounding_box()
    if box:
        target_x = box["x"] + box["width"] / 2
        target_y = box["y"] + box["height"] / 2
        page.mouse.move(target_x, target_y, steps=steps)
        time.sleep(0.15)
        page.mouse.down()
        time.sleep(0.1)
        page.mouse.up()
        time.sleep(sleep_after)


def smooth_hover(page, selector_or_locator, steps=20, sleep_after=0.4):
    if isinstance(selector_or_locator, str):
        locator = page.locator(selector_or_locator).first
    else:
        locator = selector_or_locator

    locator.scroll_into_view_if_needed()
    box = locator.bounding_box()
    if box:
        target_x = box["x"] + box["width"] / 2
        target_y = box["y"] + box["height"] / 2
        page.mouse.move(target_x, target_y, steps=steps)
        time.sleep(sleep_after)


def main():
    print("Launching Playwright browser with 720p video recording...")
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(
            viewport={"width": 1280, "height": 720},
            record_video_dir=str(RECORDING_DIR),
            record_video_size={"width": 1280, "height": 720}
        )
        page = context.new_page()

        # Inject mouse helper on every navigation
        page.add_init_script(MOUSE_HELPER_JS)

        # -------------------------------------------------------------
        # STEP 1: DASHBOARD
        # -------------------------------------------------------------
        print("Navigating to Dashboard...")
        page.goto(f"{BASE_URL}/")
        page.evaluate(MOUSE_HELPER_JS)
        time.sleep(1.2)

        # Move mouse across summary metric cards
        smooth_hover(page, "text=Total Presentations", steps=15, sleep_after=0.5)
        smooth_hover(page, "text=Approved & Released", steps=15, sleep_after=0.5)
        smooth_hover(page, "text=Awaiting Review", steps=15, sleep_after=0.5)
        smooth_hover(page, "text=PDF Converter Status", steps=15, sleep_after=0.8)

        # Scroll down slightly to observe recent presentations table
        page.mouse.wheel(0, 200)
        time.sleep(0.8)
        smooth_hover(page, "table tbody tr >> nth=0", steps=20, sleep_after=0.8)

        # -------------------------------------------------------------
        # STEP 2: UPLOAD WORKBOOK
        # -------------------------------------------------------------
        print("Navigating to Upload Excel...")
        smooth_move_and_click(page, "a:has-text('New Deck from Excel')", steps=20, sleep_after=1.0)
        page.evaluate(MOUSE_HELPER_JS)
        time.sleep(1.0)

        # Scroll to inspect upload form & supported worksheet guidelines
        smooth_hover(page, "text=Supported Worksheet Conventions:", steps=20, sleep_after=1.0)
        page.mouse.wheel(0, -150)
        time.sleep(0.5)

        # Upload canonical Excel file
        print("Uploading canonical Excel workbook...")
        file_input = page.locator("input[type='file']")
        file_input.set_input_files(str(SAMPLE_EXCEL_PATH))
        time.sleep(1.0)

        # Click Upload & Inspect Workbook
        smooth_move_and_click(page, "button:has-text('Upload & Inspect Workbook')", steps=20, sleep_after=1.5)
        page.evaluate(MOUSE_HELPER_JS)
        time.sleep(1.0)

        # -------------------------------------------------------------
        # STEP 3: CONFIGURE PRESENTATION
        # -------------------------------------------------------------
        print("Configuring presentation...")
        # Hover over Brand & Visual Theme select
        smooth_hover(page, "select[name='brand_theme']", steps=15, sleep_after=0.4)
        page.select_option("select[name='brand_theme']", "Corporate Slate")
        time.sleep(0.5)

        # Hover over Narrative Tone select
        smooth_hover(page, "select[name='narrative_style']", steps=15, sleep_after=0.4)
        page.select_option("select[name='narrative_style']", "standard")
        time.sleep(0.5)

        # Scroll down to slide selection checkboxes
        page.mouse.wheel(0, 250)
        time.sleep(0.8)
        smooth_hover(page, "text=Slide Selection & Presentation Structure", steps=15, sleep_after=0.6)

        # Click Generate PowerPoint Deck
        smooth_move_and_click(page, "button:has-text('Generate PowerPoint Deck')", steps=25, sleep_after=2.0)
        page.evaluate(MOUSE_HELPER_JS)
        time.sleep(1.5)

        # -------------------------------------------------------------
        # STEP 4: REVIEW & INLINE NARRATIVE EDITING
        # -------------------------------------------------------------
        print("Reviewing generated presentation...")
        smooth_hover(page, "h1", steps=15, sleep_after=0.6)
        smooth_hover(page, "text=IN REVIEW", steps=15, sleep_after=0.6)

        # Scroll through slides in review
        page.mouse.wheel(0, 300)
        time.sleep(1.0)

        # Focus on Slide 2 Executive Summary lead narrative textarea
        first_lead_textarea = page.locator("textarea[name='lead_summary']").first
        smooth_move_and_click(page, first_lead_textarea, steps=20, sleep_after=0.5)

        # Add custom executive remark
        first_lead_textarea.fill("Q3 2026 Executive Performance Review: Robust revenue expansion of +11.4% vs budget driven by enterprise SaaS subscriptions, accompanied by disciplined margin optimization.")
        time.sleep(0.8)

        # Click Save Changes & Rebuild PPTX
        first_save_btn = page.locator("button:has-text('Save Changes & Rebuild PPTX')").first
        smooth_move_and_click(page, first_save_btn, steps=20, sleep_after=2.0)
        page.evaluate(MOUSE_HELPER_JS)
        time.sleep(1.2)

        # -------------------------------------------------------------
        # STEP 5: FORMAL APPROVAL
        # -------------------------------------------------------------
        print("Approving Board Pack...")
        smooth_move_and_click(page, "button:has-text('Approve Board Pack')", steps=25, sleep_after=2.0)
        page.evaluate(MOUSE_HELPER_JS)
        time.sleep(1.5)

        # Verify Approved badge
        smooth_hover(page, "text=APPROVED FOR RELEASE", steps=20, sleep_after=1.0)

        # -------------------------------------------------------------
        # STEP 6: DOWNLOAD PPTX AND PDF
        # -------------------------------------------------------------
        print("Demonstrating PPTX and PDF downloads...")
        smooth_hover(page, "a:has-text('Download PPTX')", steps=20, sleep_after=0.6)
        # Click Download PPTX
        with page.expect_download():
            page.locator("a:has-text('Download PPTX')").click()
        time.sleep(1.0)

        smooth_hover(page, "a:has-text('Download PDF')", steps=20, sleep_after=0.6)
        # Click Download PDF
        with page.expect_download():
            page.locator("a:has-text('Download PDF')").click()
        time.sleep(1.2)

        # Hover over Data JSON and Manifest
        smooth_hover(page, "a:has-text('Data JSON')", steps=15, sleep_after=0.4)
        smooth_hover(page, "a:has-text('Manifest')", steps=15, sleep_after=0.5)

        # -------------------------------------------------------------
        # STEP 7: RETURN TO DASHBOARD
        # -------------------------------------------------------------
        print("Returning to Dashboard...")
        smooth_move_and_click(page, "a:has-text('Dashboard')", steps=20, sleep_after=1.5)
        page.evaluate(MOUSE_HELPER_JS)
        time.sleep(1.5)

        # Hover over the latest approved deck row and its download buttons
        smooth_hover(page, "table tbody tr >> nth=0", steps=20, sleep_after=0.8)
        smooth_hover(page, "table tbody tr >> nth=0 >> a:has-text('PPTX')", steps=15, sleep_after=0.5)
        smooth_hover(page, "table tbody tr >> nth=0 >> a:has-text('PDF')", steps=15, sleep_after=1.0)

        time.sleep(1.5)

        # Close context to finish video
        context.close()
        browser.close()
        print("Raw recording complete.")


if __name__ == "__main__":
    main()
