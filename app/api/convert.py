from __future__ import annotations

import io
import os
import subprocess
import sys
import tempfile
from pathlib import Path
from typing import List

from fastapi import APIRouter, File, HTTPException, UploadFile
from fastapi.responses import JSONResponse, StreamingResponse


router = APIRouter()


def _run_pptx2md_cli(input_path: Path, output_md: Path) -> None:
    """Invoke pptx2md CLI, trying a few variants for compatibility.

    Note: pptx2md expects `-o/--output` to be a file path (markdown file),
    not a directory. We pass a concrete file path to avoid PermissionError.
    """
    candidates: List[List[str]] = [
        ["pptx2md", str(input_path), "-o", str(output_md)],
        ["pptx2md", "-f", str(input_path), "-o", str(output_md)],
        ["pptx2md", "--input", str(input_path), "--output", str(output_md)],
        [sys.executable, "-m", "pptx2md", str(input_path), "-o", str(output_md)],
    ]

    last_err: Exception | None = None
    for cmd in candidates:
        try:
            subprocess.run(
                cmd,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                check=True,
                text=True,
            )
            return
        except Exception as e:
            last_err = e
            continue

    if last_err:
        raise last_err


def _zip_directory(dir_path: Path) -> bytes:
    import zipfile

    memfile = io.BytesIO()
    with zipfile.ZipFile(memfile, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
        for root, _, files in os.walk(dir_path):
            for name in files:
                full = Path(root) / name
                arcname = full.relative_to(dir_path)
                zf.write(full, arcname.as_posix())
    memfile.seek(0)
    return memfile.read()


@router.get("/health")
def health() -> JSONResponse:
    return JSONResponse({"status": "ok"})


@router.post("/convert")
async def convert_ppt_to_markdown(file: UploadFile = File(...)):
    """Accept a PPT/PPTX upload and return a ZIP of Markdown + assets."""
    filename = file.filename or "upload.pptx"
    suffix = Path(filename).suffix.lower()
    if suffix not in {".pptx", ".ppt"}:
        raise HTTPException(status_code=400, detail="Only .ppt or .pptx files are supported")

    with tempfile.TemporaryDirectory() as tmpdir:
        tmpdir_path = Path(tmpdir)
        input_path = tmpdir_path / filename
        output_dir = tmpdir_path / "out"
        output_dir.mkdir(parents=True, exist_ok=True)
        output_md = output_dir / "slides.md"

        contents = await file.read()
        input_path.write_bytes(contents)

        try:
            _run_pptx2md_cli(input_path, output_md)
        except FileNotFoundError:
            raise HTTPException(
                status_code=500,
                detail=(
                    "pptx2md executable not found. Ensure 'pptx2md' is installed. Try: uv add pptx2md"
                ),
            )
        except subprocess.CalledProcessError as e:
            raise HTTPException(
                status_code=500,
                detail=f"pptx2md failed: {e.stderr.strip() or e.stdout.strip() or str(e)}",
            )
        except Exception as e:
            raise HTTPException(status_code=500, detail=f"Conversion error: {e}")

        data = _zip_directory(output_dir)
        return StreamingResponse(
            io.BytesIO(data),
            media_type="application/zip",
            headers={
                "Content-Disposition": f"attachment; filename=conversion_{Path(filename).stem}.zip"
            },
        )
