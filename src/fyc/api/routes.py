"""FastAPI routes for FYC API."""

import uuid
import asyncio
from pathlib import Path
from typing import Optional

from fastapi import APIRouter, HTTPException, BackgroundTasks, UploadFile, File, Form
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
from ..utils.file_extractor import extract_text_from_files

router = APIRouter()

# In-memory job storage (use Redis in production)
jobs: dict[str, dict] = {}


async def process_presentation(job_id: str, request: GenerationRequest) -> None:
    """Background task to process a presentation generation request."""
    try:
        # Update status: Scraping
        jobs[job_id]["status"] = JobStatus.SCRAPING
        jobs[job_id]["progress"] = 0.05
        jobs[job_id]["message"] = "Scraping website for brand assets..."

        # Scrape the website
        scraped_data = await scrape_website(str(request.company_url))

        # Scraping complete
        jobs[job_id]["progress"] = 0.20
        jobs[job_id]["message"] = "Website scraped successfully..."
        await asyncio.sleep(0.1)  # Allow event loop to process status updates

        # Update status: Analyzing
        jobs[job_id]["status"] = JobStatus.ANALYZING
        jobs[job_id]["progress"] = 0.25
        jobs[job_id]["message"] = "Analyzing brand identity..."

        # Analyze brand
        brand_profile = await analyze_brand(scraped_data, str(request.company_url))
        jobs[job_id]["brand_profile"] = brand_profile

        # Analysis complete
        jobs[job_id]["progress"] = 0.40
        jobs[job_id]["message"] = "Brand identity extracted..."
        await asyncio.sleep(0.1)  # Allow event loop to process status updates

        # Update status: Generating
        jobs[job_id]["status"] = JobStatus.GENERATING
        jobs[job_id]["progress"] = 0.45
        jobs[job_id]["message"] = "Generating slide content..."

        # Generate content
        presentation = await generate_content(
            topic=request.topic,
            slide_count=request.slide_count,
            brand=brand_profile,
            additional_context=request.additional_context,
        )

        # Content generation complete
        jobs[job_id]["progress"] = 0.70
        jobs[job_id]["message"] = "Slide content generated..."
        await asyncio.sleep(0.1)  # Allow event loop to process status updates

        # Update status: Building
        jobs[job_id]["status"] = JobStatus.BUILDING
        jobs[job_id]["progress"] = 0.75
        jobs[job_id]["message"] = "Building PowerPoint file..."

        # Generate PPTX - run in executor to not block event loop
        output_dir = settings.output_dir / job_id
        output_dir.mkdir(parents=True, exist_ok=True)
        output_path = output_dir / "presentation.pptx"

        # Run sync function in executor to allow progress updates
        loop = asyncio.get_event_loop()
        await loop.run_in_executor(
            None, generate_pptx, presentation, brand_profile, str(output_path)
        )

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
    background_tasks: BackgroundTasks,
    company_url: str = Form(...),
    topic: str = Form(...),
    slide_count: int = Form(10),
    files: list[UploadFile] = File(default=[]),
) -> JobResponse:
    """Start generating a presentation with optional file uploads."""
    job_id = str(uuid.uuid4())[:8]

    # Extract text from uploaded files
    additional_context = ""
    if files:
        file_contents = []
        for file in files:
            if file.filename:
                content = await file.read()
                file_contents.append((file.filename, content))
        if file_contents:
            additional_context = extract_text_from_files(file_contents)

    # Create request object
    request = GenerationRequest(
        company_url=company_url,
        topic=topic,
        slide_count=slide_count,
        additional_context=additional_context if additional_context else None,
    )

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
