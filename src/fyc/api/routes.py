"""FastAPI routes for FYC API."""

import uuid
import asyncio
from pathlib import Path
from typing import Optional

from fastapi import APIRouter, HTTPException, BackgroundTasks
from fastapi.responses import FileResponse

from ..config import settings
from ..models import (
    GenerationRequest,
    JobResponse,
    JobStatus,
    BrandProfile,
)
from ..scraper.website_scraper import scrape_website
from ..brand.analyzer import analyze_brand
from ..content.generator import generate_content
from ..pptx_gen.generator import generate_pptx

router = APIRouter()

# In-memory job storage (use Redis in production)
jobs: dict[str, dict] = {}


async def process_presentation(job_id: str, request: GenerationRequest) -> None:
    """Background task to process a presentation generation request."""
    try:
        # Update status: Scraping
        jobs[job_id]["status"] = JobStatus.SCRAPING
        jobs[job_id]["progress"] = 0.1
        jobs[job_id]["message"] = "Scraping website for brand assets..."

        # Scrape the website
        scraped_data = await scrape_website(str(request.company_url))

        # Update status: Analyzing
        jobs[job_id]["status"] = JobStatus.ANALYZING
        jobs[job_id]["progress"] = 0.3
        jobs[job_id]["message"] = "Analyzing brand identity..."

        # Analyze brand
        brand_profile = await analyze_brand(scraped_data, str(request.company_url))
        jobs[job_id]["brand_profile"] = brand_profile

        # Update status: Generating
        jobs[job_id]["status"] = JobStatus.GENERATING
        jobs[job_id]["progress"] = 0.5
        jobs[job_id]["message"] = "Generating slide content..."

        # Generate content
        presentation = await generate_content(
            topic=request.topic,
            slide_count=request.slide_count,
            brand=brand_profile,
            additional_context=request.additional_context,
        )

        # Update status: Building
        jobs[job_id]["status"] = JobStatus.BUILDING
        jobs[job_id]["progress"] = 0.8
        jobs[job_id]["message"] = "Building PowerPoint file..."

        # Generate PPTX
        output_dir = settings.output_dir / job_id
        output_dir.mkdir(parents=True, exist_ok=True)
        output_path = output_dir / "presentation.pptx"

        generate_pptx(presentation, brand_profile, str(output_path))

        # Update status: Completed
        jobs[job_id]["status"] = JobStatus.COMPLETED
        jobs[job_id]["progress"] = 1.0
        jobs[job_id]["message"] = "Presentation ready!"
        jobs[job_id]["download_url"] = f"/api/download/{job_id}"
        jobs[job_id]["output_path"] = str(output_path)

    except Exception as e:
        jobs[job_id]["status"] = JobStatus.FAILED
        jobs[job_id]["error"] = str(e)
        jobs[job_id]["message"] = f"Error: {str(e)}"
        print(f"Error processing job {job_id}: {e}")
        import traceback
        traceback.print_exc()


@router.post("/generate", response_model=JobResponse)
async def generate_presentation(
    request: GenerationRequest,
    background_tasks: BackgroundTasks,
) -> JobResponse:
    """Start generating a presentation."""
    job_id = str(uuid.uuid4())[:8]

    # Initialize job
    jobs[job_id] = {
        "status": JobStatus.PENDING,
        "progress": 0.0,
        "message": "Job queued...",
        "brand_profile": None,
        "download_url": None,
        "error": None,
    }

    # Start background processing
    background_tasks.add_task(process_presentation, job_id, request)

    return JobResponse(
        job_id=job_id,
        status=JobStatus.PENDING,
        message="Job queued for processing",
        progress=0.0,
    )


@router.get("/job/{job_id}", response_model=JobResponse)
async def get_job_status(job_id: str) -> JobResponse:
    """Get the status of a generation job."""
    if job_id not in jobs:
        raise HTTPException(status_code=404, detail="Job not found")

    job = jobs[job_id]

    return JobResponse(
        job_id=job_id,
        status=job["status"],
        message=job["message"],
        progress=job["progress"],
        download_url=job.get("download_url"),
        brand_profile=job.get("brand_profile"),
        error=job.get("error"),
    )


@router.get("/download/{job_id}")
async def download_presentation(job_id: str) -> FileResponse:
    """Download a generated presentation."""
    if job_id not in jobs:
        raise HTTPException(status_code=404, detail="Job not found")

    job = jobs[job_id]

    if job["status"] != JobStatus.COMPLETED:
        raise HTTPException(status_code=400, detail="Presentation not ready")

    output_path = job.get("output_path")
    if not output_path or not Path(output_path).exists():
        raise HTTPException(status_code=404, detail="File not found")

    return FileResponse(
        output_path,
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        filename="presentation.pptx",
    )


@router.get("/health")
async def health_check() -> dict:
    """Health check endpoint."""
    return {"status": "healthy", "service": "fyc"}
