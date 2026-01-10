"""Pydantic models for FYC application."""

from pydantic import BaseModel, Field, HttpUrl
from typing import Optional
from enum import Enum


class ImageCategory(str, Enum):
    """Categories for scraped images."""
    TEAM = "team"
    PRODUCT = "product"
    OFFICE = "office"
    ABSTRACT = "abstract"
    LOGO = "logo"
    HERO = "hero"
    CUSTOMER = "customer"
    DATA = "data"
    USER_UPLOAD = "user_upload"
    UNKNOWN = "unknown"


class ScrapedImage(BaseModel):
    """Represents an image scraped from a website."""
    url: str
    alt_text: str = ""
    width: int = 0
    height: int = 0
    local_path: Optional[str] = None
    category: ImageCategory = ImageCategory.UNKNOWN
    description: str = ""
    relevance_score: float = 0.0


class BrandColors(BaseModel):
    """Brand color palette extracted from website."""
    primary: str = "#1a365d"
    secondary: str = "#2d3748"
    accent: str = "#3182ce"
    background: str = "#ffffff"
    text: str = "#1a202c"
    text_light: str = "#718096"


class BrandFonts(BaseModel):
    """Brand fonts extracted from website."""
    heading: str = "Arial"
    body: str = "Arial"
    heading_fallback: str = "Helvetica"
    body_fallback: str = "Helvetica"


class BrandVoice(BaseModel):
    """Brand voice characteristics."""
    formality: float = Field(default=0.5, ge=0.0, le=1.0, description="0=casual, 1=formal")
    technicality: float = Field(default=0.5, ge=0.0, le=1.0, description="0=simple, 1=technical")
    enthusiasm: float = Field(default=0.5, ge=0.0, le=1.0, description="0=reserved, 1=enthusiastic")
    sentence_length: str = "medium"  # short, medium, long
    key_phrases: list[str] = Field(default_factory=list)
    terminology: list[str] = Field(default_factory=list)
    tone_description: str = ""


class BrandProfile(BaseModel):
    """Complete brand profile extracted from website."""
    company_name: str = ""
    tagline: str = ""
    language: str = "en"
    colors: BrandColors = Field(default_factory=BrandColors)
    fonts: BrandFonts = Field(default_factory=BrandFonts)
    voice: BrandVoice = Field(default_factory=BrandVoice)
    logo_path: Optional[str] = None
    images: list[ScrapedImage] = Field(default_factory=list)
    raw_text_samples: list[str] = Field(default_factory=list)


class SlideLayout(str, Enum):
    """Available slide layouts."""
    TITLE = "title_slide"
    BULLETS = "bullet_points"
    TWO_COLUMN = "two_column"
    IMAGE_LEFT = "image_left"
    IMAGE_RIGHT = "image_right"
    SECTION_DIVIDER = "section_divider"
    QUOTE = "quote"
    STATS = "stats"  # For metrics/KPI slides
    THANK_YOU = "thank_you"


class Stat(BaseModel):
    """A single statistic/metric for stats slides."""
    value: str = ""  # e.g., "73%", "$2.5M", "10x"
    label: str = ""  # e.g., "Revenue Growth", "Cost Reduction"
    description: str = ""  # Optional longer description


class SlideContent(BaseModel):
    """Content for a single slide."""
    layout: SlideLayout
    title: str = ""
    subtitle: str = ""
    bullets: list[str] = Field(default_factory=list)
    body_text: str = ""
    left_content: Optional[str] = None
    right_content: Optional[str] = None
    quote: str = ""
    quote_author: str = ""
    stats: list[Stat] = Field(default_factory=list)  # For stats layout
    image_category: Optional[ImageCategory] = None
    speaker_notes: str = ""


class Presentation(BaseModel):
    """Complete presentation structure."""
    title: str
    subtitle: str = ""
    slides: list[SlideContent] = Field(default_factory=list)


class GenerationRequest(BaseModel):
    """Request to generate a presentation."""
    company_url: HttpUrl
    topic: str
    slide_count: int = Field(default=10, ge=3, le=30)
    additional_context: str = ""
    output_format: str = "pptx"


class JobStatus(str, Enum):
    """Status of a generation job."""
    PENDING = "pending"
    SCRAPING = "scraping"
    ANALYZING = "analyzing"
    GENERATING = "generating"
    BUILDING = "building"
    COMPLETED = "completed"
    FAILED = "failed"


class JobResponse(BaseModel):
    """Response for a generation job."""
    job_id: str
    status: JobStatus
    message: str = ""
    progress: float = Field(default=0.0, ge=0.0, le=1.0)
    download_url: Optional[str] = None
    brand_profile: Optional[BrandProfile] = None
    error: Optional[str] = None
