"""Playwright test for image upload functionality."""

import pytest
from playwright.sync_api import sync_playwright, expect
import os
from pathlib import Path


@pytest.fixture(scope="module")
def browser():
    """Launch browser for tests."""
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        yield browser
        browser.close()


@pytest.fixture
def page(browser):
    """Create a new page for each test."""
    page = browser.new_page()
    yield page
    page.close()


def test_file_upload_accepts_images(page):
    """Test that the file upload area accepts image files."""
    page.goto("http://localhost:8000")

    # Wait for page to load
    page.wait_for_selector(".wizard-step[data-step='1']")

    # Fill in URL and go to step 2
    page.fill("#companyUrl", "https://example.com")
    page.click("button:has-text('Continue')")

    # Wait for step 2 to be visible
    page.wait_for_selector(".wizard-step[data-step='2']:not([style*='display: none'])")

    # Check that the file input accepts images
    file_input = page.locator("#fileInput")
    accept_attr = file_input.get_attribute("accept")

    assert ".jpg" in accept_attr, "Should accept .jpg files"
    assert ".jpeg" in accept_attr, "Should accept .jpeg files"
    assert ".png" in accept_attr, "Should accept .png files"
    assert ".gif" in accept_attr, "Should accept .gif files"
    assert ".webp" in accept_attr, "Should accept .webp files"
    assert ".pdf" in accept_attr, "Should still accept .pdf files"
    assert ".docx" in accept_attr, "Should still accept .docx files"

    # Check that the upload text mentions images
    upload_text = page.locator(".file-upload-text").text_content()
    assert "images" in upload_text.lower(), "Upload text should mention images"


def test_image_upload_shows_correct_icon(page):
    """Test that uploaded images show the image icon."""
    page.goto("http://localhost:8000")

    # Navigate to step 2
    page.fill("#companyUrl", "https://example.com")
    page.click("button:has-text('Continue')")
    page.wait_for_selector(".wizard-step[data-step='2']:not([style*='display: none'])")

    # Upload a test image
    test_image_path = Path(__file__).parent / "test_image.png"
    if not test_image_path.exists():
        pytest.skip("Test image not found")

    file_input = page.locator("#fileInput")
    file_input.set_input_files(str(test_image_path))

    # Wait for the file to appear in the list
    page.wait_for_selector(".uploaded-file")

    # Check that the image icon is shown
    file_icon = page.locator(".uploaded-file-icon").first
    icon_text = file_icon.text_content()

    # The image icon should be the picture emoji
    assert icon_text.strip() in ["🖼️", "🖼"], f"Expected image icon, got: {icon_text}"

    # Check file name is displayed
    file_name = page.locator(".uploaded-file-info span:nth-child(2)").first.text_content()
    assert "test_image.png" in file_name


def test_mixed_file_upload(page):
    """Test uploading both images and documents."""
    page.goto("http://localhost:8000")

    # Navigate to step 2
    page.fill("#companyUrl", "https://example.com")
    page.click("button:has-text('Continue')")
    page.wait_for_selector(".wizard-step[data-step='2']:not([style*='display: none'])")

    test_image_path = Path(__file__).parent / "test_image.png"
    if not test_image_path.exists():
        pytest.skip("Test image not found")

    # Upload the test image
    file_input = page.locator("#fileInput")
    file_input.set_input_files(str(test_image_path))

    # Wait for the file to appear
    page.wait_for_selector(".uploaded-file")

    # Check that we have one file uploaded
    uploaded_files = page.locator(".uploaded-file")
    assert uploaded_files.count() == 1


def test_upload_hint_mentions_images(page):
    """Test that the upload hint mentions images for slides."""
    page.goto("http://localhost:8000")

    # Navigate to step 2
    page.fill("#companyUrl", "https://example.com")
    page.click("button:has-text('Continue')")
    page.wait_for_selector(".wizard-step[data-step='2']:not([style*='display: none'])")

    # Check the hint text
    hint_text = page.locator(".file-upload-hint").text_content()
    assert "image" in hint_text.lower(), "Hint should mention images"
    assert "slide" in hint_text.lower(), "Hint should mention slides"


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
