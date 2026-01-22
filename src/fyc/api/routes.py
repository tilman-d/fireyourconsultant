"""FastAPI routes for FYC API."""

import uuid
import asyncio
from pathlib import Path
from typing import Optional

from fastapi import APIRouter, HTTPException, BackgroundTasks, UploadFile, File, Form
from fastapi.responses import FileResponse

from ..config import settings, logger
from ..models import (
    GenerationRequest,
    JobResponse,
    JobStatus,
    BrandProfile,
    ScrapedImage,
    ImageCategory,
)
from ..scraper.website_scraper import scrape_website
from ..brand.analyzer import analyze_brand, BrandAnalyzer
from ..content.generator import generate_content
from ..pptx_gen.generator import generate_pptx
from ..utils.file_extractor import extract_text_from_files
from ..template.extractor import extract_template_styles

router = APIRouter()

# In-memory job storage (use Redis in production)
jobs: dict[str, dict] = {}


async def process_presentation(job_id: str, request: GenerationRequest) -> None:
    """Background task to process a presentation generation request."""
    logger.info(f"Starting job {job_id}: topic='{request.topic}', slides={request.slide_count}")
    try:
        output_dir = settings.output_dir / job_id
        output_dir.mkdir(parents=True, exist_ok=True)

        template_profile = None
        brand_profile = None
        analyzer = BrandAnalyzer()

        # Extract template styles if provided
        template_content = jobs[job_id].get("template_content")
        template_filename = jobs[job_id].get("template_filename")

        if template_content:
            jobs[job_id]["status"] = JobStatus.ANALYZING
            jobs[job_id]["progress"] = 0.05
            jobs[job_id]["message"] = "Extracting styles from template..."

            template_profile = await extract_template_styles(
                template_content,
                template_filename,
                str(output_dir),
            )

            jobs[job_id]["progress"] = 0.15
            jobs[job_id]["message"] = "Template styles extracted..."
            await asyncio.sleep(0.1)

        # Scrape website if URL provided
        if request.company_url:
            jobs[job_id]["status"] = JobStatus.SCRAPING
            jobs[job_id]["progress"] = 0.20
            jobs[job_id]["message"] = "Scraping website for brand assets..."

            # Scrape the website
            scraped_data = await scrape_website(str(request.company_url))

            # Scraping complete
            jobs[job_id]["progress"] = 0.30
            jobs[job_id]["message"] = "Website scraped successfully..."
            await asyncio.sleep(0.1)

            # Update status: Analyzing
            jobs[job_id]["status"] = JobStatus.ANALYZING
            jobs[job_id]["progress"] = 0.35
            jobs[job_id]["message"] = "Analyzing brand identity..."

            # Analyze brand
            brand_profile = await analyze_brand(scraped_data, str(request.company_url))

            # Merge with template if both provided (template takes priority)
            if template_profile:
                jobs[job_id]["message"] = "Merging template styles with brand..."
                brand_profile = analyzer.merge_with_template(brand_profile, template_profile)
        else:
            # Create brand profile from template only (no website)
            jobs[job_id]["status"] = JobStatus.ANALYZING
            jobs[job_id]["progress"] = 0.35
            jobs[job_id]["message"] = "Creating brand profile from template..."

            brand_profile = analyzer.create_profile_from_template(
                template_profile,
                company_name="",
            )

        # Add user-uploaded images to the brand profile (prioritize them)
        uploaded_images = jobs[job_id].get("uploaded_images", [])
        if uploaded_images:
            # Prepend uploaded images so they are used first
            brand_profile.images = uploaded_images + brand_profile.images

        jobs[job_id]["brand_profile"] = brand_profile

        # Analysis complete
        jobs[job_id]["progress"] = 0.40
        jobs[job_id]["message"] = "Brand identity extracted..."
        await asyncio.sleep(0.1)

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
        await asyncio.sleep(0.1)

        # Update status: Building
        jobs[job_id]["status"] = JobStatus.BUILDING
        jobs[job_id]["progress"] = 0.75
        jobs[job_id]["message"] = "Building PowerPoint file..."

        # Generate PPTX - run in executor to not block event loop
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
        logger.info(f"Job {job_id} completed successfully: {output_path}")

    except Exception as e:
        jobs[job_id]["status"] = JobStatus.FAILED
        jobs[job_id]["error"] = str(e)
        jobs[job_id]["message"] = f"Error: {str(e)}"
        logger.error(f"Error processing job {job_id}: {e}", exc_info=True)


@router.post("/generate", response_model=JobResponse)
async def generate_presentation(
    background_tasks: BackgroundTasks,
    company_url: Optional[str] = Form(None),
    topic: str = Form(...),
    slide_count: int = Form(10),
    files: list[UploadFile] = File(default=[]),
    template_file: Optional[UploadFile] = File(None),
) -> JobResponse:
    """Start generating a presentation with optional file uploads and template."""
    # Validate: at least one brand source required
    has_url = company_url and company_url.strip()
    has_template = template_file and template_file.filename

    if not has_url and not has_template:
        raise HTTPException(
            status_code=400,
            detail="Either company_url or template_file must be provided",
        )

    job_id = str(uuid.uuid4())[:8]

    # Separate document files from image files
    image_extensions = {'.jpg', '.jpeg', '.png', '.gif', '.webp'}
    document_extensions = {'.pdf', '.docx'}

    document_contents = []
    uploaded_images: list[ScrapedImage] = []

    # Create job output directory for saving images
    output_dir = settings.output_dir / job_id
    output_dir.mkdir(parents=True, exist_ok=True)

    if files:
        for file in files:
            if file.filename:
                ext = Path(file.filename).suffix.lower()
                content = await file.read()

                if ext in document_extensions:
                    document_contents.append((file.filename, content))
                elif ext in image_extensions:
                    # Save image to job directory
                    image_path = output_dir / file.filename
                    with open(image_path, 'wb') as f:
                        f.write(content)

                    # Create ScrapedImage object for the uploaded image
                    uploaded_images.append(ScrapedImage(
                        url=f"uploaded://{file.filename}",
                        alt_text=f"User uploaded: {file.filename}",
                        local_path=str(image_path),
                        category=ImageCategory.USER_UPLOAD,
                        description=f"User-uploaded image: {file.filename}",
                        relevance_score=1.0,  # High relevance for user uploads
                    ))

    # Process template file if provided
    template_content = None
    template_filename = None
    if template_file and template_file.filename:
        template_content = await template_file.read()
        template_filename = template_file.filename
        # Save uploaded template for comparison
        template_save_path = output_dir / f"uploaded_template_{template_filename}"
        with open(template_save_path, 'wb') as f:
            f.write(template_content)

    # Extract text from document files
    additional_context = ""
    if document_contents:
        additional_context = extract_text_from_files(document_contents)

    # Create request object (company_url now optional)
    request = GenerationRequest(
        company_url=company_url if has_url else None,
        topic=topic,
        slide_count=slide_count,
        additional_context=additional_context or "",
    )

    # Initialize job with template info
    jobs[job_id] = {
        "status": JobStatus.PENDING,
        "progress": 0.0,
        "message": "Job queued...",
        "brand_profile": None,
        "download_url": None,
        "error": None,
        "uploaded_images": uploaded_images,
        "template_content": template_content,
        "template_filename": template_filename,
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
