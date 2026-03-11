"""
Fast PPTX -> HTML Conversion Route (Job-based with polling)

Uses the same job system as the original /generate-html endpoint:
  1. POST /convert-pptx-fast  (upload file + x-session-id) -> returns jobId
  2. POST /job-status          (existing endpoint, poll)    -> returns status + result

Pipeline (background):
  PPTX -> PDF -> split pages -> parallel Claude per slide -> CSS unify -> upload HTML
"""

from fastapi import APIRouter, UploadFile, File, HTTPException, BackgroundTasks, Header
from pydantic import BaseModel
from typing import Optional
import traceback
import time
import uuid
from datetime import datetime

from app.services.converter import convert_pptx_to_images
from app.services.parallel_html_generator import generate_html_parallel
from app.services.supabase_client import (
    ensure_session,
    create_generation_job,
    update_job_status,
    upload_generated_html,
)

router = APIRouter(tags=["fast-conversion"])


# ---------------------------------------------------------------------------
# Models
# ---------------------------------------------------------------------------

class ConvertJobResponse(BaseModel):
    success: bool
    jobId: Optional[str] = None
    error: Optional[str] = None


# ---------------------------------------------------------------------------
# Background processing task
# ---------------------------------------------------------------------------

async def _process_fast_conversion(
    job_id: str,
    session_id: str,
    pptx_bytes: bytes,
    filename: str,
):
    """Background task: PPTX -> PNG -> parallel Claude -> CSS unify -> upload HTML."""
    pipeline_start = time.time()
    tag = job_id[:8]

    try:
        await update_job_status(job_id, "processing")

        # Step 1/4: PPTX -> PNG
        print(f"[convert-fast] [{tag}] Step 1/4: PPTX -> PNG conversion...")
        step_start = time.time()
        images = convert_pptx_to_images(pptx_bytes, filename=filename)
        conversion_time = time.time() - step_start
        print(f"[convert-fast] [{tag}] Step 1/4: Done - {len(images)} slides in {conversion_time:.1f}s")

        # Step 2/4: Parallel Claude Vision + CSS unification
        print(f"[convert-fast] [{tag}] Step 2/4: Parallel Claude Vision...")
        step_start = time.time()
        result = await generate_html_parallel(images)
        parallel_time = time.time() - step_start
        print(f"[convert-fast] [{tag}] Step 2/4: Done - {result['slide_count']} slides in {parallel_time:.1f}s")

        # Step 3/4: Upload HTML to storage
        print(f"[convert-fast] [{tag}] Step 3/4: Uploading HTML to storage...")
        step_start = time.time()
        timestamp = datetime.utcnow().strftime("%Y%m%d_%H%M%S")
        html_filename = f"fast_report_{timestamp}.html"
        html_url = await upload_generated_html(session_id, result["full_html"], html_filename)
        upload_time = time.time() - step_start
        print(f"[convert-fast] [{tag}] Step 3/4: Done in {upload_time:.1f}s")
        print(f"[convert-fast] [{tag}]          URL: {html_url}")

        # Step 4/4: Update job as completed
        total_time = time.time() - pipeline_start
        print(f"[convert-fast] [{tag}] Step 4/4: Complete")
        print(f"[convert-fast] [{tag}] === Total: {total_time:.1f}s ===")

        await update_job_status(
            job_id,
            "completed",
            result={
                "htmlUrl": html_url,
                "slideCount": result["slide_count"],
                "okCount": result["ok_count"],
                "failCount": result["fail_count"],
                "totalSeconds": round(total_time, 2),
                "conversionSeconds": round(conversion_time, 2),
                "parallelSeconds": result["timings"]["parallel_seconds"],
                "cssUnifySeconds": result["timings"]["css_unify_seconds"],
                "avgPerSlide": result["timings"]["avg_per_slide"],
                "perSlideSeconds": result["timings"]["per_slide"],
            },
        )

    except Exception as e:
        total_time = time.time() - pipeline_start
        error_msg = str(e)
        print(f"[convert-fast] [{tag}] FAILED after {total_time:.1f}s: {error_msg}")
        traceback.print_exc()
        await update_job_status(job_id, "failed", error=error_msg)


# ---------------------------------------------------------------------------
# Endpoints
# ---------------------------------------------------------------------------

@router.post("/convert-pptx-fast", response_model=ConvertJobResponse)
async def convert_pptx_fast(
    background_tasks: BackgroundTasks,
    file: UploadFile = File(...),
    x_session_id: str = Header(..., alias="x-session-id"),
):
    """
    Start fast PPTX -> HTML conversion.

    Accepts a PPTX file + x-session-id header.
    Creates a generation job and starts processing in background.
    Returns jobId immediately. Poll /job-status to track progress.
    """
    try:
        if not file.filename or not file.filename.lower().endswith((".pptx", ".ppt")):
            raise HTTPException(status_code=400, detail="File must be a .pptx file")

        pptx_bytes = await file.read()
        file_size_kb = len(pptx_bytes) / 1024
        print(f"[convert-fast] Received file: {file.filename} ({file_size_kb:.1f} KB) session={x_session_id[:8]}")

        # Ensure session exists (creates minimal one if needed)
        await ensure_session(x_session_id)

        # Create job in DB (reuses existing generation_jobs table)
        job_id = await create_generation_job(
            session_id=x_session_id,
            engine="claude-html-fast",
            input_data={"filename": file.filename, "file_size_kb": round(file_size_kb, 1)},
        )
        print(f"[convert-fast] Job {job_id[:8]} created")

        # Start background processing
        background_tasks.add_task(
            _process_fast_conversion,
            job_id=job_id,
            session_id=x_session_id,
            pptx_bytes=pptx_bytes,
            filename=file.filename,
        )

        return ConvertJobResponse(success=True, jobId=job_id)

    except HTTPException:
        raise
    except Exception as e:
        traceback.print_exc()
        return ConvertJobResponse(success=False, error=str(e))
