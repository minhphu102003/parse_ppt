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

# Resolve project root (two levels up from this file: app/api -> app -> project)
PROJECT_ROOT = Path(__file__).resolve().parents[2]
# Base output directory; override via env var PARSEPPT_OUTPUT_DIR if desired
OUTPUT_BASE = Path(os.environ.get("PARSEPPT_OUTPUT_DIR", str(PROJECT_ROOT / "output")))


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

        # Project-local output directory: ./output/<stem>/
        stem = Path(filename).stem
        output_dir = OUTPUT_BASE / stem
        output_dir.mkdir(parents=True, exist_ok=True)
        output_md = output_dir / f"{stem}.md"

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


@router.post("/convert/pptx2md")
async def convert_with_pptx2md(file: UploadFile = File(...)):
    return await convert_ppt_to_markdown(file)


@router.post("/convert/markitdown")
async def convert_with_markitdown(file: UploadFile = File(...)):
    filename = file.filename or "upload.pptx"
    suffix = Path(filename).suffix.lower()
    if suffix not in {".pptx", ".ppt"}:
        raise HTTPException(status_code=400, detail="Only .ppt or .pptx files are supported")

    try:
        import markitdown  # type: ignore
    except Exception:
        raise HTTPException(
            status_code=500,
            detail=(
                "Missing dependency: markitdown. Install with: uv add markitdown"
            ),
        )

    with tempfile.TemporaryDirectory() as tmpdir:
        tmpdir_path = Path(tmpdir)
        input_path = tmpdir_path / filename
        contents = await file.read()
        input_path.write_bytes(contents)

        stem = Path(filename).stem
        output_dir = OUTPUT_BASE / stem
        output_dir.mkdir(parents=True, exist_ok=True)
        output_md = output_dir / f"{stem}.md"

        # Run MarkItDown
        try:
            md = markitdown.MarkItDown()
            result = md.convert(str(input_path))
            text: str | None = None
            # Try common return types
            if isinstance(result, str):
                text = result
            elif hasattr(result, "text_content"):
                text = getattr(result, "text_content")
            elif hasattr(result, "content") and isinstance(getattr(result, "content"), str):
                text = getattr(result, "content")
            elif isinstance(result, dict):
                text = result.get("text") or result.get("content") or result.get("markdown")
            elif isinstance(result, tuple) and result:
                first = result[0]
                text = first if isinstance(first, str) else None
            if not text:
                raise RuntimeError("Unexpected MarkItDown result; cannot extract markdown text")
            output_md.write_text(text, encoding="utf-8")
        except Exception as e:
            raise HTTPException(status_code=500, detail=f"markitdown failed: {e}")

        data = _zip_directory(output_dir)
        return StreamingResponse(
            io.BytesIO(data),
            media_type="application/zip",
            headers={"Content-Disposition": f"attachment; filename=conversion_{stem}.zip"},
        )


@router.post("/convert/pandoc")
async def convert_with_pandoc(file: UploadFile = File(...)):
    filename = file.filename or "upload.pptx"
    suffix = Path(filename).suffix.lower()
    if suffix not in {".pptx", ".ppt"}:
        raise HTTPException(status_code=400, detail="Only .ppt or .pptx files are supported")

    with tempfile.TemporaryDirectory() as tmpdir:
        tmpdir_path = Path(tmpdir)
        input_path = tmpdir_path / filename
        contents = await file.read()
        input_path.write_bytes(contents)

        stem = Path(filename).stem
        output_dir = OUTPUT_BASE / stem
        output_dir.mkdir(parents=True, exist_ok=True)
        output_md = output_dir / f"{stem}.md"

        # Try pandoc conversion. Note: pandoc may not support PPT/PPTX input on some versions.
        candidates = [
            ["pandoc", str(input_path), "-t", "gfm", "-o", str(output_md)],
            ["pandoc", "-f", "pptx", str(input_path), "-t", "gfm", "-o", str(output_md)],
        ]
        last_err: Exception | None = None
        for cmd in candidates:
            try:
                subprocess.run(cmd, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
                last_err = None
                break
            except Exception as e:
                last_err = e

        if last_err:
            raise HTTPException(
                status_code=500,
                detail=(
                    f"pandoc failed or not installed: {last_err}. Install pandoc and try again."
                ),
            )

        data = _zip_directory(output_dir)
        return StreamingResponse(
            io.BytesIO(data),
            media_type="application/zip",
            headers={"Content-Disposition": f"attachment; filename=conversion_{stem}.zip"},
        )


@router.post("/convert/pptx_to_md")
async def convert_with_pptx_to_md(file: UploadFile = File(...)):
    filename = file.filename or "upload.pptx"
    suffix = Path(filename).suffix.lower()
    if suffix not in {".pptx", ".ppt"}:
        raise HTTPException(status_code=400, detail="Only .ppt or .pptx files are supported")

    with tempfile.TemporaryDirectory() as tmpdir:
        tmpdir_path = Path(tmpdir)
        input_path = tmpdir_path / filename
        contents = await file.read()
        input_path.write_bytes(contents)

        stem = Path(filename).stem
        output_dir = OUTPUT_BASE / stem
        output_dir.mkdir(parents=True, exist_ok=True)
        output_md = output_dir / f"{stem}.md"

        # Try Python module first
        err: Exception | None = None
        try:
            import pptx_to_md  # type: ignore
            # Try common entry points
            text: str | None = None
            for fname in ("convert", "convert_pptx_to_markdown", "pptx_to_markdown"):
                func = getattr(pptx_to_md, fname, None)
                if callable(func):
                    try:
                        res = func(str(input_path))
                        if isinstance(res, str):
                            text = res
                        elif isinstance(res, tuple) and res and isinstance(res[0], str):
                            text = res[0]
                        elif hasattr(res, "markdown") and isinstance(res.markdown, str):
                            text = res.markdown
                        if text:
                            break
                    except Exception as e:
                        err = e
            if text:
                output_md.write_text(text, encoding="utf-8")
            else:
                raise RuntimeError("pptx_to_md: could not find a suitable conversion function")
        except Exception as e:
            err = e

        # Fallback to CLI if available
        if err:
            candidates = [
                ["pptx_to_md", str(input_path), "-o", str(output_md)],
                [sys.executable, "-m", "pptx_to_md", str(input_path), "-o", str(output_md)],
            ]
            last_err: Exception | None = None
            for cmd in candidates:
                try:
                    subprocess.run(cmd, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
                    last_err = None
                    break
                except Exception as e:
                    last_err = e
            if last_err:
                raise HTTPException(
                    status_code=500,
                    detail=(
                        "pptx_to_md not installed or failed. Install the library/CLI and try again."
                    ),
                )

        data = _zip_directory(output_dir)
        return StreamingResponse(
            io.BytesIO(data),
            media_type="application/zip",
            headers={"Content-Disposition": f"attachment; filename=conversion_{stem}.zip"},
        )


@router.post("/convert/aspose")
async def convert_with_aspose(file: UploadFile = File(...)):
    filename = file.filename or "upload.pptx"
    suffix = Path(filename).suffix.lower()
    if suffix not in {".pptx", ".ppt"}:
        raise HTTPException(status_code=400, detail="Only .ppt or .pptx files are supported")

    try:
        import aspose.slides as slides  # type: ignore
    except Exception:
        raise HTTPException(
            status_code=500,
            detail=(
                "Missing dependency: aspose.slides. Install and license it to use this endpoint."
            ),
        )

    with tempfile.TemporaryDirectory() as tmpdir:
        tmpdir_path = Path(tmpdir)
        input_path = tmpdir_path / filename
        contents = await file.read()
        input_path.write_bytes(contents)

        stem = Path(filename).stem
        output_dir = OUTPUT_BASE / stem
        output_dir.mkdir(parents=True, exist_ok=True)
        output_md = output_dir / f"{stem}.md"

        try:
            # Aspose.Slides can export to a folder with md + assets
            # The exact API may vary by version; this is a common pattern
            pres = slides.Presentation(str(input_path))
            try:
                # Newer API namespace
                from aspose.slides.export import MarkdownSaveOptions, SaveFormat  # type: ignore
                mopts = MarkdownSaveOptions()
                pres.save(str(output_dir), SaveFormat.MARKDOWN, mopts)
            except Exception:
                # Older style
                pres.save(str(output_dir), slides.export.SaveFormat.MARKDOWN)

            # If no main md produced, concatenate slide markdowns if present
            if not output_md.exists():
                # Attempt to find a single md or merge all
                mds = list(output_dir.glob("*.md"))
                if mds:
                    if len(mds) == 1:
                        mds[0].rename(output_md)
                    else:
                        combined = "\n\n".join(p.read_text(encoding="utf-8") for p in sorted(mds))
                        output_md.write_text(combined, encoding="utf-8")
        except Exception as e:
            raise HTTPException(status_code=500, detail=f"Aspose conversion failed: {e}")

        data = _zip_directory(output_dir)
        return StreamingResponse(
            io.BytesIO(data),
            media_type="application/zip",
            headers={"Content-Disposition": f"attachment; filename=conversion_{stem}.zip"},
        )
