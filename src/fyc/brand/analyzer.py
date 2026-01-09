"""Brand analyzer using Claude to interpret scraped website data."""

import base64
import json
from pathlib import Path
from collections import Counter

import anthropic
from langdetect import detect

from ..config import settings
from ..models import (
    BrandProfile,
    BrandColors,
    BrandFonts,
    BrandVoice,
    ScrapedImage,
    ImageCategory,
)


# Font mapping from web fonts to PowerPoint-safe fonts
FONT_MAPPING = {
    # Sans-serif
    "Open Sans": "Calibri",
    "Roboto": "Calibri",
    "Lato": "Calibri",
    "Montserrat": "Arial",
    "Source Sans Pro": "Calibri",
    "Raleway": "Century Gothic",
    "Poppins": "Arial",
    "Nunito": "Calibri",
    "Inter": "Calibri",
    "Work Sans": "Calibri",
    "DM Sans": "Arial",
    # Serif
    "Playfair Display": "Georgia",
    "Merriweather": "Georgia",
    "Lora": "Palatino Linotype",
    "PT Serif": "Times New Roman",
    "Source Serif Pro": "Georgia",
    # Fallbacks
    "sans-serif": "Arial",
    "serif": "Times New Roman",
    "monospace": "Courier New",
}


class BrandAnalyzer:
    """Analyzes scraped data to create a brand profile using Claude."""

    def __init__(self):
        self.client = anthropic.Anthropic(api_key=settings.anthropic_api_key)

    async def analyze(self, scraped_data: dict, url: str = "") -> BrandProfile:
        """Analyze scraped website data and create a brand profile."""
        colors = scraped_data.get("colors", [])
        fonts = scraped_data.get("fonts", [])
        images = scraped_data.get("images", [])
        logo_candidates = scraped_data.get("logo_candidates", [])
        text_content = scraped_data.get("text_content", [])
        temp_dir = scraped_data.get("temp_dir", "")

        # Detect language from text
        language = self._detect_language(text_content)

        # Analyze colors with Claude
        brand_colors = await self._analyze_colors(colors)

        # Map fonts to PowerPoint-safe fonts
        brand_fonts = self._map_fonts(fonts)

        # Analyze brand voice with Claude
        brand_voice = await self._analyze_voice(text_content, language)

        # Categorize images with Claude Vision
        categorized_images = await self._categorize_images(images + logo_candidates)

        # Extract company name and tagline
        company_info = await self._extract_company_info(text_content, url)

        # Find best logo
        logo_path = self._find_logo(logo_candidates, categorized_images)

        return BrandProfile(
            company_name=company_info.get("name", ""),
            tagline=company_info.get("tagline", ""),
            language=language,
            colors=brand_colors,
            fonts=brand_fonts,
            voice=brand_voice,
            logo_path=logo_path,
            images=[img for img in categorized_images if img.category != ImageCategory.LOGO],
            raw_text_samples=text_content[:10],
        )

    def _detect_language(self, text_content: list[str]) -> str:
        """Detect the primary language of the website."""
        combined_text = " ".join(text_content[:20])
        if not combined_text:
            return "en"

        try:
            lang = detect(combined_text)
            return lang
        except Exception:
            return "en"

    async def _analyze_colors(self, colors: list[str]) -> BrandColors:
        """Use Claude to identify the brand color palette."""
        if not colors:
            return BrandColors()

        # Count color frequency and filter common web colors
        color_counts = Counter(colors)

        # Filter out pure white/black and very common defaults
        filtered = {
            c: count
            for c, count in color_counts.items()
            if c not in ["#ffffff", "#000000", "#fff", "#000"]
        }

        top_colors = [c for c, _ in sorted(filtered.items(), key=lambda x: -x[1])[:20]]

        if not top_colors:
            return BrandColors()

        prompt = f"""Analyze these colors extracted from a company website and identify the brand color palette.

Colors (sorted by frequency):
{json.dumps(top_colors, indent=2)}

Identify:
1. Primary color (main brand color, often used in headers/buttons)
2. Secondary color (complementary color)
3. Accent color (call-to-action, highlights)
4. Background color
5. Text color (main body text)
6. Light text color (secondary text, captions)

Return ONLY a JSON object with these exact keys:
{{
    "primary": "#hexcode",
    "secondary": "#hexcode",
    "accent": "#hexcode",
    "background": "#hexcode",
    "text": "#hexcode",
    "text_light": "#hexcode"
}}"""

        response = self.client.messages.create(
            model=settings.claude_model,
            max_tokens=500,
            messages=[{"role": "user", "content": prompt}],
        )

        try:
            # Extract JSON from response
            text = response.content[0].text
            json_match = text[text.find("{") : text.rfind("}") + 1]
            color_data = json.loads(json_match)

            return BrandColors(
                primary=color_data.get("primary", "#1a365d"),
                secondary=color_data.get("secondary", "#2d3748"),
                accent=color_data.get("accent", "#3182ce"),
                background=color_data.get("background", "#ffffff"),
                text=color_data.get("text", "#1a202c"),
                text_light=color_data.get("text_light", "#718096"),
            )
        except Exception as e:
            print(f"Error parsing color response: {e}")
            return BrandColors()

    def _map_fonts(self, fonts: list[str]) -> BrandFonts:
        """Map web fonts to PowerPoint-safe equivalents."""
        if not fonts:
            return BrandFonts()

        # Try to identify heading vs body fonts
        heading_font = None
        body_font = None

        for font in fonts:
            clean_font = font.strip()
            if clean_font in FONT_MAPPING:
                mapped = FONT_MAPPING[clean_font]
                if heading_font is None:
                    heading_font = mapped
                elif body_font is None and mapped != heading_font:
                    body_font = mapped
            elif heading_font is None:
                # Use as-is if not in mapping
                heading_font = clean_font

        return BrandFonts(
            heading=heading_font or "Arial",
            body=body_font or heading_font or "Arial",
            heading_fallback="Arial",
            body_fallback="Calibri",
        )

    async def _analyze_voice(self, text_content: list[str], language: str) -> BrandVoice:
        """Use Claude to analyze the brand's writing style and voice."""
        if not text_content:
            return BrandVoice()

        combined_text = "\n\n".join(text_content[:15])

        prompt = f"""Analyze the following text from a company website and describe their brand voice.

Text samples:
\"\"\"
{combined_text[:4000]}
\"\"\"

Analyze and return a JSON object with:
{{
    "formality": <float 0-1, 0=very casual, 1=very formal>,
    "technicality": <float 0-1, 0=simple language, 1=technical jargon>,
    "enthusiasm": <float 0-1, 0=reserved/professional, 1=enthusiastic/energetic>,
    "sentence_length": "<short/medium/long>",
    "key_phrases": ["phrase 1", "phrase 2", ...],  // 3-5 distinctive phrases they use
    "terminology": ["term 1", "term 2", ...],  // 3-5 industry/company-specific terms
    "tone_description": "<2-3 sentence description of the brand's tone>"
}}

Note: The website is in {language}."""

        response = self.client.messages.create(
            model=settings.claude_model,
            max_tokens=1000,
            messages=[{"role": "user", "content": prompt}],
        )

        try:
            text = response.content[0].text
            json_match = text[text.find("{") : text.rfind("}") + 1]
            voice_data = json.loads(json_match)

            return BrandVoice(
                formality=float(voice_data.get("formality", 0.5)),
                technicality=float(voice_data.get("technicality", 0.5)),
                enthusiasm=float(voice_data.get("enthusiasm", 0.5)),
                sentence_length=voice_data.get("sentence_length", "medium"),
                key_phrases=voice_data.get("key_phrases", [])[:5],
                terminology=voice_data.get("terminology", [])[:5],
                tone_description=voice_data.get("tone_description", ""),
            )
        except Exception as e:
            print(f"Error parsing voice response: {e}")
            return BrandVoice()

    async def _categorize_images(self, images: list[ScrapedImage]) -> list[ScrapedImage]:
        """Use Claude Vision to categorize images."""
        categorized = []

        for image in images:
            if not image.local_path or not Path(image.local_path).exists():
                categorized.append(image)
                continue

            # Skip SVGs for vision (not supported)
            if image.local_path.endswith(".svg"):
                image.category = ImageCategory.LOGO if "logo" in image.alt_text.lower() else ImageCategory.ABSTRACT
                categorized.append(image)
                continue

            try:
                # Read and encode image
                image_data = Path(image.local_path).read_bytes()
                base64_image = base64.b64encode(image_data).decode("utf-8")

                # Determine media type
                media_type = "image/jpeg"
                if image.local_path.endswith(".png"):
                    media_type = "image/png"
                elif image.local_path.endswith(".gif"):
                    media_type = "image/gif"
                elif image.local_path.endswith(".webp"):
                    media_type = "image/webp"

                prompt = """Categorize this image from a company website.

Return ONLY one of these categories:
- team: Photos of people, team members, employees
- product: Product images, screenshots, demos
- office: Office spaces, buildings, workplaces
- abstract: Abstract graphics, patterns, illustrations
- logo: Company logos, brand marks
- hero: Large banner/hero images
- customer: Customer photos, testimonials
- data: Charts, graphs, infographics

Also provide a brief 1-sentence description.

Return JSON:
{"category": "<category>", "description": "<description>"}"""

                response = self.client.messages.create(
                    model=settings.claude_vision_model,
                    max_tokens=200,
                    messages=[
                        {
                            "role": "user",
                            "content": [
                                {
                                    "type": "image",
                                    "source": {
                                        "type": "base64",
                                        "media_type": media_type,
                                        "data": base64_image,
                                    },
                                },
                                {"type": "text", "text": prompt},
                            ],
                        }
                    ],
                )

                text = response.content[0].text
                json_match = text[text.find("{") : text.rfind("}") + 1]
                data = json.loads(json_match)

                category_str = data.get("category", "unknown").upper()
                image.category = ImageCategory[category_str] if category_str in ImageCategory.__members__ else ImageCategory.UNKNOWN
                image.description = data.get("description", "")

            except Exception as e:
                print(f"Error categorizing image {image.local_path}: {e}")
                image.category = ImageCategory.UNKNOWN

            categorized.append(image)

        return categorized

    async def _extract_company_info(self, text_content: list[str], url: str = "") -> dict:
        """Extract company name and tagline from text."""
        if not text_content:
            return {"name": "", "tagline": ""}

        combined_text = "\n".join(text_content[:15])

        # Try to extract domain hint
        domain_hint = ""
        if url:
            from urllib.parse import urlparse
            parsed = urlparse(url)
            domain = parsed.netloc.replace("www.", "")
            domain_hint = f"\nDomain: {domain}"

        prompt = f"""Analyze this website content and extract the company/organization name and their tagline.

IMPORTANT: Look carefully for:
- Company name in headers, navigation, or footer
- The brand name (e.g., "McKinsey", "Stripe", "Anthropic")
- Main tagline or value proposition statement
{domain_hint}

Website text:
\"\"\"
{combined_text[:3000]}
\"\"\"

Return ONLY valid JSON with the company name (just the name, not "Inc" or "Company" unless it's part of the brand):
{{"name": "<company name>", "tagline": "<tagline or main value proposition, empty if not found>"}}"""

        try:
            response = self.client.messages.create(
                model=settings.claude_model,
                max_tokens=200,
                messages=[{"role": "user", "content": prompt}],
            )

            text = response.content[0].text
            json_match = text[text.find("{") : text.rfind("}") + 1]
            return json.loads(json_match)
        except Exception as e:
            print(f"Error extracting company info: {e}")
            return {"name": "", "tagline": ""}

    def _find_logo(self, logo_candidates: list[ScrapedImage], all_images: list[ScrapedImage]) -> str | None:
        """Find the best logo image."""
        # First check logo candidates
        for img in logo_candidates:
            if img.local_path and Path(img.local_path).exists():
                return img.local_path

        # Then check categorized images
        for img in all_images:
            if img.category == ImageCategory.LOGO and img.local_path:
                return img.local_path

        return None


async def analyze_brand(scraped_data: dict, url: str = "") -> BrandProfile:
    """Convenience function to analyze brand from scraped data."""
    analyzer = BrandAnalyzer()
    return await analyzer.analyze(scraped_data, url)
