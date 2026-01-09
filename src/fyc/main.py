"""Main FastAPI application entry point."""

from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse
from pathlib import Path

from .config import settings
from .api.routes import router as api_router

# Create FastAPI app
app = FastAPI(
    title=settings.app_name,
    description="AI-powered presentation generator that creates corporate-styled PowerPoint slides",
    version="0.1.0",
)

# Configure CORS
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # Configure appropriately for production
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Include API routes
app.include_router(api_router, prefix="/api")

# Create output directory
settings.output_dir.mkdir(parents=True, exist_ok=True)
settings.temp_dir.mkdir(parents=True, exist_ok=True)

# Static files path
static_path = Path(__file__).parent.parent.parent / "static"


@app.get("/")
async def root():
    """Serve the frontend."""
    index_path = static_path / "index.html"
    if index_path.exists():
        return FileResponse(index_path)
    return {
        "name": settings.app_name,
        "version": "0.1.0",
        "docs": "/docs",
        "api": "/api",
    }


# Mount static files AFTER the root route
if static_path.exists():
    app.mount("/static", StaticFiles(directory=str(static_path)), name="static")


def run():
    """Run the application with uvicorn."""
    import uvicorn

    uvicorn.run(
        "fyc.main:app",
        host=settings.api_host,
        port=settings.api_port,
        reload=settings.debug,
    )


if __name__ == "__main__":
    run()
