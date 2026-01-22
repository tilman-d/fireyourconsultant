"""Configuration settings for FYC application."""

import logging
from logging.handlers import RotatingFileHandler
from pydantic_settings import BaseSettings
from pydantic import Field
from pathlib import Path


def setup_logging(log_dir: Path) -> logging.Logger:
    """Configure application logging with file rotation."""
    log_dir.mkdir(parents=True, exist_ok=True)
    log_file = log_dir / "fyc.log"

    logger = logging.getLogger("fyc")
    logger.setLevel(logging.DEBUG)

    # File handler with rotation (10MB max, keep 5 backups)
    file_handler = RotatingFileHandler(
        log_file,
        maxBytes=10 * 1024 * 1024,
        backupCount=5,
    )
    file_handler.setLevel(logging.DEBUG)
    file_formatter = logging.Formatter(
        "%(asctime)s - %(name)s - %(levelname)s - %(message)s"
    )
    file_handler.setFormatter(file_formatter)

    # Console handler for errors only
    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.ERROR)
    console_handler.setFormatter(file_formatter)

    logger.addHandler(file_handler)
    logger.addHandler(console_handler)

    return logger


class Settings(BaseSettings):
    """Application settings loaded from environment variables."""

    # API Keys
    anthropic_api_key: str = Field(default="", env="ANTHROPIC_API_KEY")

    # Claude Model Settings
    claude_model: str = Field(default="claude-sonnet-4-20250514", env="CLAUDE_MODEL")
    claude_vision_model: str = Field(default="claude-sonnet-4-20250514", env="CLAUDE_VISION_MODEL")

    # Application Settings
    app_name: str = "Fire Your Consultant"
    debug: bool = Field(default=True, env="DEBUG")

    # Paths
    output_dir: Path = Field(default=Path("output"), env="OUTPUT_DIR")
    temp_dir: Path = Field(default=Path("/tmp/fyc"), env="TEMP_DIR")

    # Scraper Settings
    scraper_timeout: int = Field(default=30000, env="SCRAPER_TIMEOUT")  # ms
    max_pages_to_scrape: int = Field(default=10, env="MAX_PAGES_TO_SCRAPE")
    max_images_to_download: int = Field(default=20, env="MAX_IMAGES_TO_DOWNLOAD")
    min_image_size: int = Field(default=100, env="MIN_IMAGE_SIZE")  # pixels

    # Redis Settings (for job queue)
    redis_url: str = Field(default="redis://localhost:6379", env="REDIS_URL")

    # API Settings
    api_host: str = Field(default="0.0.0.0", env="API_HOST")
    api_port: int = Field(default=8000, env="API_PORT")

    class Config:
        env_file = ".env"
        env_file_encoding = "utf-8"


settings = Settings()

# Initialize logging
log_dir = Path(__file__).parent.parent.parent / "logs"
logger = setup_logging(log_dir)
