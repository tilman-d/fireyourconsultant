# Fire Your Consultant (FYC)

AI-powered presentation generator that creates corporate-styled PowerPoint slides.

## Features

- Scrapes company websites for brand identity (colors, fonts, logo, images)
- Analyzes brand voice and writing style
- Generates slide content using Claude AI
- Creates native, editable PowerPoint files

## Quick Start

```bash
# Install dependencies
pip install -e .

# Install Playwright browsers
playwright install chromium

# Set your Anthropic API key
export ANTHROPIC_API_KEY=your_key_here

# Run the API server
python -m fyc.main
```

## API Usage

### Generate Presentation

```bash
curl -X POST http://localhost:8000/api/generate \
  -H "Content-Type: application/json" \
  -d '{
    "company_url": "https://example.com",
    "topic": "Q4 Strategy",
    "slide_count": 10
  }'
```

### Check Job Status

```bash
curl http://localhost:8000/api/job/{job_id}
```

### Download Presentation

```bash
curl -O http://localhost:8000/api/download/{job_id}
```

## License

MIT
