"""Content generator using Claude to create high-quality presentation slides."""

import json
from typing import Optional

import anthropic

from ..config import settings
from ..models import (
    BrandProfile,
    Presentation,
    SlideContent,
    SlideLayout,
    ImageCategory,
)


class ContentGenerator:
    """Generates professional presentation content using Claude."""

    def __init__(self):
        self.client = anthropic.Anthropic(api_key=settings.anthropic_api_key)

    async def generate(
        self,
        topic: str,
        slide_count: int,
        brand: BrandProfile,
        additional_context: str = "",
    ) -> Presentation:
        """Generate presentation content based on topic and brand profile."""

        system_prompt = self._build_system_prompt(brand)
        user_prompt = self._build_user_prompt(topic, slide_count, brand, additional_context)

        response = self.client.messages.create(
            model=settings.claude_model,
            max_tokens=4000,
            system=system_prompt,
            messages=[{"role": "user", "content": user_prompt}],
        )

        return self._parse_response(response.content[0].text, brand)

    def _build_system_prompt(self, brand: BrandProfile) -> str:
        """Build system prompt with brand voice and quality guidelines."""
        voice = brand.voice

        # Determine voice characteristics
        if voice.formality > 0.7:
            formality = "formal, professional, authoritative"
        elif voice.formality < 0.3:
            formality = "casual, friendly, conversational"
        else:
            formality = "balanced professional yet approachable"

        if voice.technicality > 0.7:
            technical = "Use technical terminology appropriate for industry experts."
        elif voice.technicality < 0.3:
            technical = "Use simple, accessible language that anyone can understand."
        else:
            technical = "Balance technical accuracy with accessibility."

        if voice.enthusiasm > 0.7:
            energy = "Be enthusiastic, dynamic, and inspiring."
        elif voice.enthusiasm < 0.3:
            energy = "Be measured, thoughtful, and understated."
        else:
            energy = "Be confident and engaging without being over-the-top."

        key_phrases = ", ".join(brand.voice.key_phrases[:5]) if brand.voice.key_phrases else "none specified"
        terminology = ", ".join(brand.voice.terminology[:5]) if brand.voice.terminology else "none specified"

        return f"""You are an expert presentation designer creating slides for {brand.company_name or 'a company'}.

## BRAND VOICE
- Tone: {formality}
- {technical}
- {energy}
- Key phrases to incorporate naturally: {key_phrases}
- Industry terminology: {terminology}
- Additional voice notes: {voice.tone_description}

## LANGUAGE
Write ALL content in {brand.language.upper()}. This is critical.

## SLIDE DESIGN PRINCIPLES

### Content Quality
1. Each slide must have ONE clear, memorable message
2. Titles should be impactful statements, not just topic labels
   - BAD: "Our Services"
   - GOOD: "Solutions That Scale With You"
3. Bullet points: Maximum 4-5 per slide, each under 12 words
4. Use concrete numbers and specifics, not vague claims
5. End bullets with impact, not prepositions

### Layout Selection
Choose layouts strategically:
- title_slide: ONLY for the opening slide
- bullet_points: For key features, benefits, or steps (limit to 4-5 items)
- two_column: For comparisons, before/after, pros/cons, or contrasting ideas
- stats: For impressive metrics, KPIs, or numeric achievements (use 3-4 stats per slide)
- image_left / image_right: USE THESE when images are available! They add visual interest and branding. Alternate between left and right placement.
- section_divider: To introduce new major sections (use sparingly)
- quote: For testimonials, key statements, or memorable insights
- thank_you: ONLY for the final slide

### Image Usage (IMPORTANT!)
When image categories are available, you MUST include image_left or image_right layouts to showcase them.
- Include at least 2-3 image slides if 3+ image categories are available
- Include at least 1-2 image slides if 1-2 image categories are available
- Specify the image_category field matching an available category
- Use images for: showcasing products, team culture, office environment, customer stories

### Stats Slides (Important!)
When presenting metrics, KPIs, or numeric achievements, use the stats layout:
- Use bullets array with format: "VALUE - Description" (e.g., "73% - Revenue Growth")
- Include 3-4 impactful statistics per slide
- Make numbers prominent and memorable
- Use stats layout when you have quantifiable achievements to highlight

### Structure
1. Start with a compelling title slide with tagline
2. Group content into 2-3 logical sections
3. Use section_divider slides between major topic shifts
4. Vary layouts - never use the same layout twice in a row
5. End with a strong call-to-action thank_you slide

### Two-Column Slides
When using two_column layout:
- left_content and right_content should be STRINGS
- Start each with a header using **Header Text** format
- Then list items with bullet points
- Make the comparison clear and balanced

Example format:
left_content: "**Traditional Approach**\\n• Manual processes\\n• Higher costs\\n• Slower results"
right_content: "**Our Solution**\\n• Automated workflows\\n• Cost-effective\\n• Faster delivery"

## OUTPUT FORMAT
Return ONLY valid JSON. No markdown, no explanation."""

    def _build_user_prompt(
        self,
        topic: str,
        slide_count: int,
        brand: BrandProfile,
        additional_context: str,
    ) -> str:
        """Build the user prompt for content generation."""
        available_images = self._get_available_image_categories(brand)
        images_str = ", ".join(available_images) if available_images else "none available"

        context_block = f"\nADDITIONAL CONTEXT: {additional_context}" if additional_context else ""

        return f"""Create a {slide_count}-slide presentation.

TOPIC: {topic}
COMPANY: {brand.company_name or 'The company'}
TAGLINE: {brand.tagline or 'Not specified'}
{context_block}

AVAILABLE IMAGE CATEGORIES: {images_str}
(Only use image_category if that category is available)

REQUIREMENTS:
1. Slide 1: title_slide with compelling headline
2. Slides 2-{slide_count-1}: Mix of bullet_points, two_column, stats, image_left, image_right, section_divider layouts
3. IMPORTANT: Include at least ONE stats slide with impactful metrics
4. IMPORTANT: If image categories are available above, include 2-3 slides using image_left or image_right layouts with matching image_category field
5. Slide {slide_count}: thank_you with call-to-action
6. Include speaker_notes for each slide (2-3 sentences of talking points)
7. Never use the same layout twice consecutively
8. Make titles action-oriented and specific

Return this exact JSON structure:
{{
    "title": "Main Presentation Title",
    "subtitle": "Supporting tagline or description",
    "slides": [
        {{
            "layout": "title_slide",
            "title": "Compelling Title Here",
            "subtitle": "Supporting subtitle",
            "speaker_notes": "Brief talking points for presenter"
        }},
        {{
            "layout": "bullet_points",
            "title": "Action-Oriented Title",
            "bullets": ["Point 1", "Point 2", "Point 3", "Point 4"],
            "speaker_notes": "Talking points"
        }},
        {{
            "layout": "two_column",
            "title": "Comparison Title",
            "left_content": "**Left Header**\\n• Item 1\\n• Item 2\\n• Item 3",
            "right_content": "**Right Header**\\n• Item 1\\n• Item 2\\n• Item 3",
            "speaker_notes": "Talking points"
        }},
        {{
            "layout": "stats",
            "title": "Impact by the Numbers",
            "bullets": ["73% - Revenue Growth", "$2.5M - Cost Savings", "10x - Productivity Increase", "98% - Customer Satisfaction"],
            "speaker_notes": "Highlight key metrics"
        }},
        {{
            "layout": "section_divider",
            "title": "Section Name",
            "subtitle": "Brief description",
            "speaker_notes": "Transition notes"
        }},
        {{
            "layout": "image_left",
            "title": "Visual Story Title",
            "body_text": "Supporting content that complements the image and tells the story",
            "image_category": "product",
            "speaker_notes": "Talking points about the visual"
        }},
        {{
            "layout": "image_right",
            "title": "Another Visual Slide",
            "bullets": ["Key point one", "Key point two", "Key point three"],
            "image_category": "team",
            "speaker_notes": "Talking points"
        }},
        {{
            "layout": "thank_you",
            "title": "Thank You",
            "body_text": "Call to action or contact info",
            "speaker_notes": "Closing remarks"
        }}
    ]
}}"""

    def _get_available_image_categories(self, brand: BrandProfile) -> list[str]:
        """Get list of image categories available from scraped images."""
        categories = set()
        for img in brand.images:
            if img.category != ImageCategory.UNKNOWN and img.local_path:
                # Skip SVGs
                if not img.local_path.endswith('.svg'):
                    categories.add(img.category.value)
        return list(categories)

    def _normalize_column_content(self, content) -> str | None:
        """Normalize column content to a string."""
        if content is None:
            return None
        if isinstance(content, str):
            return content
        if isinstance(content, list):
            return "\n• ".join([""] + [str(item) for item in content]) if content else ""
        if isinstance(content, dict):
            parts = []
            if "heading" in content:
                parts.append(f"**{content['heading']}**\n")
            for key in ["content", "items", "text", "points"]:
                if key in content:
                    val = content[key]
                    if isinstance(val, list):
                        parts.append("\n• ".join([""] + [str(item) for item in val]))
                    else:
                        parts.append(str(val))
                    break
            return "".join(parts) if parts else str(content)
        return str(content)

    def _parse_response(self, response_text: str, brand: BrandProfile) -> Presentation:
        """Parse Claude's response into a Presentation object."""
        try:
            # Extract JSON from response
            json_start = response_text.find("{")
            json_end = response_text.rfind("}") + 1
            json_str = response_text[json_start:json_end]

            data = json.loads(json_str)

            slides = []
            for slide_data in data.get("slides", []):
                layout_str = slide_data.get("layout", "bullet_points")
                try:
                    layout = SlideLayout(layout_str)
                except ValueError:
                    layout = SlideLayout.BULLETS

                image_cat = None
                if slide_data.get("image_category"):
                    try:
                        image_cat = ImageCategory(slide_data["image_category"])
                    except ValueError:
                        pass

                # Normalize column content
                left_content = self._normalize_column_content(slide_data.get("left_content"))
                right_content = self._normalize_column_content(slide_data.get("right_content"))

                # Handle bullets that might be a string
                bullets = slide_data.get("bullets", [])
                if isinstance(bullets, str):
                    bullets = [b.strip() for b in bullets.split("\n") if b.strip()]

                slide = SlideContent(
                    layout=layout,
                    title=slide_data.get("title", ""),
                    subtitle=slide_data.get("subtitle", ""),
                    bullets=bullets,
                    body_text=slide_data.get("body_text", ""),
                    left_content=left_content,
                    right_content=right_content,
                    quote=slide_data.get("quote", ""),
                    quote_author=slide_data.get("quote_author", ""),
                    image_category=image_cat,
                    speaker_notes=slide_data.get("speaker_notes", ""),
                )
                slides.append(slide)

            return Presentation(
                title=data.get("title", "Presentation"),
                subtitle=data.get("subtitle", ""),
                slides=slides,
            )

        except json.JSONDecodeError as e:
            print(f"Error parsing JSON response: {e}")
            print(f"Response text: {response_text[:1000]}...")
            return Presentation(
                title="Presentation",
                slides=[
                    SlideContent(
                        layout=SlideLayout.TITLE,
                        title="Error generating content",
                        subtitle="Please try again",
                    )
                ],
            )


async def generate_content(
    topic: str,
    slide_count: int,
    brand: BrandProfile,
    additional_context: str = "",
) -> Presentation:
    """Generate presentation content."""
    generator = ContentGenerator()
    return await generator.generate(topic, slide_count, brand, additional_context)
