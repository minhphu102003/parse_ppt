# parseppt

FastAPI service to convert PPT/PPTX files to Markdown using [ssine/pptx2md](https://github.com/ssine/pptx2md).

## Requirements

- Python 3.10+
- [uv](https://github.com/astral-sh/uv) package manager installed

## Setup

```
# Create virtual environment (optional but recommended)
uv venv

# Activate if you don't use `uv run` (Windows PowerShell):
# . .venv/Scripts/Activate.ps1

# Install dependencies
uv sync  # or: uv add fastapi uvicorn python-multipart pptx2md
```

## Run

```
# Development server
uv run uvicorn app.main:app --reload --host 127.0.0.1 --port 8000

# Or run the module directly
uv run python -m app.main
```

Open Swagger UI: http://127.0.0.1:8000/docs

## API

- POST `/convert`
  - Form field: `file` (PPT or PPTX)
  - Response: zip file containing generated Markdown (`.md`) and assets (images)

Example with `curl`:

```
curl -X POST \
  -F "file=@slides.pptx" \
  -o conversion.zip \
  http://127.0.0.1:8000/convert
```

## Notes

- This service shells out to the `pptx2md` CLI. Ensure it is installed in the same environment (it is listed in `pyproject.toml`).
- The output structure depends on `pptx2md` version. This API returns the entire output directory as a zip to include Markdown and referenced images.
- If you prefer returning a single concatenated Markdown file, we can add an alternate endpoint that merges all `.md` files and inlines images.

## License

This project uses the `pptx2md` library by ssine. Refer to their repository for its license and usage details.
