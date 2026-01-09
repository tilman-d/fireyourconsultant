"""Test progress tracking during presentation generation."""

import asyncio
import pytest
from playwright.async_api import async_playwright, expect


@pytest.fixture(scope="module")
def event_loop():
    """Create event loop for async tests."""
    loop = asyncio.new_event_loop()
    yield loop
    loop.close()


@pytest.mark.asyncio
async def test_progress_shows_multiple_stages():
    """Test that progress updates show multiple intermediate values, not just 10% -> 100%."""
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        page = await browser.new_page()

        # Navigate to the app
        await page.goto("http://localhost:8000")

        # Step 1: Enter company URL
        await page.fill("#companyUrl", "https://example.com")
        await page.click("button:has-text('Continue')")

        # Step 2: Enter topic and start generation
        await page.fill("#topic", "Test presentation about AI")
        await page.click("button:has-text('Generate Presentation')")

        # Wait for progress step to appear
        await page.wait_for_selector(".wizard-step[data-step='3'].active", timeout=5000)

        # Collect progress values over time
        progress_values = []
        max_wait_time = 120  # seconds
        check_interval = 0.3  # seconds
        elapsed = 0

        while elapsed < max_wait_time:
            # Check if we're still on progress step or moved to success/error
            progress_step = await page.query_selector(".wizard-step[data-step='3'].active")
            success_step = await page.query_selector(".wizard-step[data-step='4'].active")
            error_step = await page.query_selector(".wizard-step[data-step='error'].active")

            if success_step or error_step:
                # Add 100% if we reached success
                if success_step:
                    progress_values.append(100)
                break

            if progress_step:
                # Get current progress value
                progress_text = await page.text_content("#progressValue")
                if progress_text:
                    try:
                        progress = int(progress_text.replace("%", ""))
                        if not progress_values or progress != progress_values[-1]:
                            progress_values.append(progress)
                            print(f"Progress: {progress}%")
                    except ValueError:
                        pass

            await asyncio.sleep(check_interval)
            elapsed += check_interval

        await browser.close()

        # Assertions
        print(f"\nCollected progress values: {progress_values}")

        # We should have captured multiple different progress values
        assert len(progress_values) >= 3, (
            f"Expected at least 3 different progress values, got {len(progress_values)}: {progress_values}. "
            "Progress jumped too quickly without showing intermediate states."
        )

        # Progress should include values beyond just 10% and 100%
        has_intermediate = any(10 < p < 100 for p in progress_values)
        assert has_intermediate, (
            f"Expected intermediate progress values between 10% and 100%, got: {progress_values}"
        )

        # Should end at 100% (success)
        assert progress_values[-1] == 100, (
            f"Expected progress to end at 100%, but ended at {progress_values[-1]}%"
        )


@pytest.mark.asyncio
async def test_progress_stages_update():
    """Test that progress stages (icons) update as processing proceeds."""
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        page = await browser.new_page()

        await page.goto("http://localhost:8000")

        # Fill form and start generation
        await page.fill("#companyUrl", "https://example.com")
        await page.click("button:has-text('Continue')")
        await page.fill("#topic", "Quick test")
        await page.click("button:has-text('Generate Presentation')")

        # Wait for progress step
        await page.wait_for_selector(".wizard-step[data-step='3'].active", timeout=5000)

        # Track which stages became active
        stages_activated = set()
        max_wait_time = 120
        elapsed = 0
        check_interval = 0.1  # Poll faster to catch stage transitions

        while elapsed < max_wait_time:
            # Check each stage
            for stage_name in ["scraping", "analyzing", "generating", "building"]:
                stage = await page.query_selector(f".stage[data-stage='{stage_name}']")
                if stage:
                    class_attr = await stage.get_attribute("class")
                    if class_attr and ("active" in class_attr or "completed" in class_attr):
                        stages_activated.add(stage_name)

            # Check if done
            success = await page.query_selector(".wizard-step[data-step='4'].active")
            error = await page.query_selector(".wizard-step[data-step='error'].active")
            if success or error:
                break

            await asyncio.sleep(check_interval)
            elapsed += check_interval

        await browser.close()

        print(f"\nStages that were activated: {stages_activated}")

        # At least the first two stages (scraping, analyzing) should be activated
        # Some stages may be too fast to catch with UI polling
        assert "scraping" in stages_activated, "Scraping stage should have been activated"
        assert "analyzing" in stages_activated, "Analyzing stage should have been activated"
        # At least 2 stages should be activated (we might miss fast ones)
        assert len(stages_activated) >= 2, (
            f"Expected at least 2 stages to be activated, got: {stages_activated}"
        )


if __name__ == "__main__":
    asyncio.run(test_progress_shows_multiple_stages())
