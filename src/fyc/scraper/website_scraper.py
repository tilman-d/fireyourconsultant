"""Website scraper using Playwright to extract brand assets."""

import asyncio
import re
import hashlib
import json
from pathlib import Path
from urllib.parse import urljoin, urlparse
from typing import Optional
from datetime import datetime, timedelta

from playwright.async_api import async_playwright, Page, Browser, TimeoutError as PlaywrightTimeout
import httpx
from PIL import Image
import io

from ..config import settings
from ..models import ScrapedImage, ImageCategory


# Simple file-based cache for brand profiles
CACHE_DIR = Path("/tmp/fyc_cache")
CACHE_EXPIRY_HOURS = 24


def get_cached_scrape(url: str) -> dict | None:
    """Get cached scrape data if available and not expired."""
    CACHE_DIR.mkdir(parents=True, exist_ok=True)
    cache_key = hashlib.md5(url.encode()).hexdigest()
    cache_file = CACHE_DIR / f"{cache_key}.json"

    if cache_file.exists():
        try:
            data = json.loads(cache_file.read_text())
            cached_time = datetime.fromisoformat(data.get("_cached_at", "2000-01-01"))
            if datetime.now() - cached_time < timedelta(hours=CACHE_EXPIRY_HOURS):
                print(f"Using cached scrape data for {url}")
                return data
        except Exception:
            pass
    return None


def save_scrape_cache(url: str, data: dict) -> None:
    """Save scrape data to cache."""
    CACHE_DIR.mkdir(parents=True, exist_ok=True)
    cache_key = hashlib.md5(url.encode()).hexdigest()
    cache_file = CACHE_DIR / f"{cache_key}.json"

    data["_cached_at"] = datetime.now().isoformat()
    try:
        cache_file.write_text(json.dumps(data, default=str))
    except Exception as e:
        print(f"Failed to cache scrape data: {e}")


class WebsiteScraper:
    """Scrapes websites for brand assets: colors, fonts, images, and text."""

    def __init__(self, base_url: str):
        self.base_url = base_url.rstrip("/")
        self.domain = urlparse(base_url).netloc
        self.browser: Optional[Browser] = None
        self.visited_urls: set[str] = set()
        self.colors: list[str] = []
        self.fonts: set[str] = set()
        self.images: list[ScrapedImage] = []
        self.text_content: list[str] = []
        self.logo_candidates: list[ScrapedImage] = []

        # Create temp directory for this scrape
        self.temp_dir = settings.temp_dir / hashlib.md5(base_url.encode()).hexdigest()[:12]
        self.temp_dir.mkdir(parents=True, exist_ok=True)

    async def scrape(self) -> dict:
        """Main scraping entry point. Returns all extracted data."""
        async with async_playwright() as p:
            self.browser = await p.chromium.launch(
                headless=True,
                args=['--no-sandbox', '--disable-setuid-sandbox']
            )

            try:
                # Scrape the main pages with retry
                success = await self._scrape_page_with_retry(self.base_url)

                if not success:
                    print(f"Warning: Could not scrape main page {self.base_url}")

                # Try common important pages
                important_paths = ["/about", "/about-us", "/team", "/products", "/services", "/contact"]
                for path in important_paths:
                    if len(self.visited_urls) >= settings.max_pages_to_scrape:
                        break
                    url = f"{self.base_url}{path}"
                    if url not in self.visited_urls:
                        await self._scrape_page_with_retry(url)

                # Download images
                await self._download_images()

            except Exception as e:
                print(f"Scraping error: {e}")
            finally:
                await self.browser.close()

        # Return results even if partial
        return {
            "colors": self._dedupe_colors(),
            "fonts": list(self.fonts),
            "images": self.images,
            "logo_candidates": self.logo_candidates,
            "text_content": self.text_content,
            "temp_dir": str(self.temp_dir),
        }

    async def _scrape_page_with_retry(self, url: str, retries: int = 2) -> bool:
        """Scrape a page with retry logic for blocked sites."""
        # Skip non-http URLs
        if not url.startswith("http"):
            return False

        for attempt in range(retries + 1):
            try:
                await self._scrape_page(url)
                return True
            except PlaywrightTimeout:
                print(f"Timeout on {url} (attempt {attempt + 1})")
                if attempt < retries:
                    await asyncio.sleep(2)
            except Exception as e:
                error_str = str(e).lower()
                # Handle Cloudflare and other blocks
                if any(x in error_str for x in ["cloudflare", "captcha", "blocked", "403", "access denied"]):
                    print(f"Site appears blocked: {url}")
                    return False
                print(f"Error scraping {url}: {e}")
                return False
        return False

    async def _scrape_page(self, url: str) -> None:
        """Scrape a single page for assets."""
        if url in self.visited_urls:
            return

        self.visited_urls.add(url)
        print(f"Scraping: {url}")

        context = await self.browser.new_context(
            viewport={"width": 1920, "height": 1080},
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
        )
        page = await context.new_page()

        try:
            await page.goto(url, timeout=settings.scraper_timeout, wait_until="networkidle")
            await asyncio.sleep(1)  # Let JS finish

            # Extract colors from CSS
            await self._extract_colors(page)

            # Extract fonts
            await self._extract_fonts(page)

            # Extract images
            await self._extract_images(page, url)

            # Extract text content
            await self._extract_text(page)

            # Find internal links for further scraping
            await self._find_links(page)

        except Exception as e:
            print(f"Error scraping {url}: {e}")
        finally:
            await context.close()

    async def _extract_colors(self, page: Page) -> None:
        """Extract color values from the page's CSS."""
        colors = await page.evaluate("""
            () => {
                const colors = new Set();
                const elements = document.querySelectorAll('*');

                for (const el of elements) {
                    const style = window.getComputedStyle(el);
                    const props = ['color', 'backgroundColor', 'borderColor', 'fill', 'stroke'];

                    for (const prop of props) {
                        const value = style.getPropertyValue(prop);
                        if (value && value !== 'rgba(0, 0, 0, 0)' && value !== 'transparent') {
                            colors.add(value);
                        }
                    }
                }

                return Array.from(colors);
            }
        """)
        self.colors.extend(colors)

    async def _extract_fonts(self, page: Page) -> None:
        """Extract font families from the page."""
        fonts = await page.evaluate("""
            () => {
                const fonts = new Set();
                const elements = document.querySelectorAll('h1, h2, h3, h4, h5, h6, p, span, a, button, li');

                for (const el of elements) {
                    const style = window.getComputedStyle(el);
                    const fontFamily = style.getPropertyValue('font-family');
                    if (fontFamily) {
                        // Get first font in the stack
                        const firstFont = fontFamily.split(',')[0].trim().replace(/['"]/g, '');
                        fonts.add(firstFont);
                    }
                }

                return Array.from(fonts);
            }
        """)
        self.fonts.update(fonts)

    async def _extract_images(self, page: Page, page_url: str) -> None:
        """Extract image information from the page."""
        images_data = await page.evaluate(r"""
            () => {
                const images = [];
                const imgElements = document.querySelectorAll('img');

                for (const img of imgElements) {
                    const rect = img.getBoundingClientRect();
                    images.push({
                        src: img.src,
                        alt: img.alt || '',
                        width: rect.width || img.naturalWidth,
                        height: rect.height || img.naturalHeight,
                        isLogo: img.alt?.toLowerCase().includes('logo') ||
                                img.src?.toLowerCase().includes('logo') ||
                                img.className?.toLowerCase().includes('logo'),
                    });
                }

                // Also check for background images
                const bgElements = document.querySelectorAll('[style*="background-image"]');
                for (const el of bgElements) {
                    const style = window.getComputedStyle(el);
                    const bgImage = style.backgroundImage;
                    const match = bgImage.match(/url\(["']?([^"')]+)["']?\)/);
                    if (match) {
                        const rect = el.getBoundingClientRect();
                        images.push({
                            src: match[1],
                            alt: '',
                            width: rect.width,
                            height: rect.height,
                            isLogo: false,
                        });
                    }
                }

                return images;
            }
        """)

        for img_data in images_data:
            src = img_data.get("src", "")
            if not src or src.startswith("data:"):
                continue

            # Make URL absolute
            full_url = urljoin(page_url, src)

            # Skip tiny images (likely icons/tracking)
            width = img_data.get("width", 0)
            height = img_data.get("height", 0)
            if width < settings.min_image_size and height < settings.min_image_size:
                continue

            image = ScrapedImage(
                url=full_url,
                alt_text=img_data.get("alt", ""),
                width=int(width),
                height=int(height),
            )

            if img_data.get("isLogo"):
                self.logo_candidates.append(image)
            else:
                self.images.append(image)

    async def _extract_text(self, page: Page) -> None:
        """Extract meaningful text content from the page."""
        text = await page.evaluate("""
            () => {
                const texts = [];
                const selectors = ['h1', 'h2', 'h3', 'p', 'li', '.hero-text', '.tagline', '.about'];

                for (const selector of selectors) {
                    const elements = document.querySelectorAll(selector);
                    for (const el of elements) {
                        const text = el.innerText?.trim();
                        if (text && text.length > 20 && text.length < 1000) {
                            texts.push(text);
                        }
                    }
                }

                return texts.slice(0, 50);  // Limit to avoid too much data
            }
        """)
        self.text_content.extend(text)

    async def _find_links(self, page: Page) -> None:
        """Find internal links to scrape."""
        if len(self.visited_urls) >= settings.max_pages_to_scrape:
            return

        links = await page.evaluate(f"""
            () => {{
                const links = [];
                const domain = '{self.domain}';
                const anchors = document.querySelectorAll('a[href]');

                for (const a of anchors) {{
                    const href = a.href;
                    if (href.includes(domain) && !href.includes('#')) {{
                        links.push(href);
                    }}
                }}

                return [...new Set(links)].slice(0, 20);
            }}
        """)

        for link in links:
            if link not in self.visited_urls and len(self.visited_urls) < settings.max_pages_to_scrape:
                try:
                    await self._scrape_page(link)
                except Exception:
                    pass

    async def _download_images(self) -> None:
        """Download scraped images to local storage."""
        all_images = self.logo_candidates + self.images
        downloaded = 0

        async with httpx.AsyncClient(timeout=10.0, follow_redirects=True) as client:
            for image in all_images:
                if downloaded >= settings.max_images_to_download:
                    break

                try:
                    response = await client.get(image.url)
                    if response.status_code == 200:
                        # Determine file extension
                        content_type = response.headers.get("content-type", "")
                        ext = ".jpg"
                        if "png" in content_type:
                            ext = ".png"
                        elif "gif" in content_type:
                            ext = ".gif"
                        elif "webp" in content_type:
                            ext = ".webp"
                        elif "svg" in content_type:
                            ext = ".svg"

                        # Save file
                        filename = f"img_{downloaded:03d}{ext}"
                        filepath = self.temp_dir / filename
                        filepath.write_bytes(response.content)
                        image.local_path = str(filepath)

                        # Update dimensions if we didn't get them from browser
                        if image.width == 0 or image.height == 0:
                            try:
                                with Image.open(io.BytesIO(response.content)) as img:
                                    image.width, image.height = img.size
                            except Exception:
                                pass

                        downloaded += 1

                except Exception as e:
                    print(f"Failed to download {image.url}: {e}")

    def _dedupe_colors(self) -> list[str]:
        """Deduplicate and normalize colors."""
        normalized = set()

        for color in self.colors:
            # Convert rgb/rgba to hex
            if color.startswith("rgb"):
                match = re.match(r"rgba?\((\d+),\s*(\d+),\s*(\d+)", color)
                if match:
                    r, g, b = int(match.group(1)), int(match.group(2)), int(match.group(3))
                    hex_color = f"#{r:02x}{g:02x}{b:02x}"
                    normalized.add(hex_color)
            elif color.startswith("#"):
                # Normalize to 6 digits
                if len(color) == 4:
                    color = f"#{color[1]*2}{color[2]*2}{color[3]*2}"
                normalized.add(color.lower())

        return list(normalized)


async def scrape_website(url: str, use_cache: bool = True) -> dict:
    """Convenience function to scrape a website.

    Args:
        url: The website URL to scrape
        use_cache: Whether to use cached data if available (default: True)

    Returns:
        Dict with scraped brand data (colors, fonts, images, etc.)
    """
    # Check cache first
    if use_cache:
        cached = get_cached_scrape(url)
        if cached:
            # Remove cache metadata before returning
            cached.pop("_cached_at", None)
            # Convert image dicts back to ScrapedImage objects
            cached["images"] = [
                ScrapedImage(**img) if isinstance(img, dict) else img
                for img in cached.get("images", [])
            ]
            cached["logo_candidates"] = [
                ScrapedImage(**img) if isinstance(img, dict) else img
                for img in cached.get("logo_candidates", [])
            ]
            return cached

    # Scrape fresh data
    scraper = WebsiteScraper(url)
    result = await scraper.scrape()

    # Cache the result (excluding non-serializable items)
    cache_data = {
        "colors": result.get("colors", []),
        "fonts": result.get("fonts", []),
        "text_content": result.get("text_content", []),
        "temp_dir": result.get("temp_dir", ""),
        # Store image URLs and metadata, not the ScrapedImage objects
        "images": [
            {
                "url": img.url,
                "alt_text": img.alt_text,
                "width": img.width,
                "height": img.height,
                "local_path": img.local_path,
            }
            for img in result.get("images", [])
        ],
        "logo_candidates": [
            {
                "url": img.url,
                "alt_text": img.alt_text,
                "width": img.width,
                "height": img.height,
                "local_path": img.local_path,
            }
            for img in result.get("logo_candidates", [])
        ],
    }
    save_scrape_cache(url, cache_data)

    return result
