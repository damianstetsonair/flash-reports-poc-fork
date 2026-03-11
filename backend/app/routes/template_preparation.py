"""
Template Preparation Routes

Handles the background PPTX → PDF → PNG → HTML conversion pipeline.
Endpoints:
  POST /prepare-template          - Start background conversion
  POST /template-preparation-status - Check conversion status
  POST /list-slides-from-html     - List slides from prepared HTML
"""

from fastapi import APIRouter, BackgroundTasks, Header, HTTPException
from pydantic import BaseModel
from typing import Optional, List, Dict, Any
import traceback
import time
from datetime import datetime

from app.services.converter import convert_pptx_to_images
from app.services.claude_html import generate_html_template
from app.services.supabase_client import (
    get_session,
    download_template,
    upload_generated_html,
    upload_pdf,
    upload_png,
    update_template_preparation_status,
    get_template_preparation_status,
    download_html_template,
)

router = APIRouter()


# --- Models ---

class PrepareTemplateResponse(BaseModel):
    success: bool
    message: Optional[str] = None
    error: Optional[str] = None


class TemplatePreparationStatusResponse(BaseModel):
    success: bool
    status: str  # 'pending' | 'processing' | 'completed' | 'failed'
    htmlTemplateUrl: Optional[str] = None
    templatePngUrls: Optional[List[str]] = None
    templatePdfUrl: Optional[str] = None
    error: Optional[str] = None


class ListSlidesFromHtmlResponse(BaseModel):
    success: bool
    slides: Optional[List[Dict[str, Any]]] = None
    total: Optional[int] = None
    error: Optional[str] = None


# --- Endpoints ---

@router.post("/prepare-template", response_model=PrepareTemplateResponse)
async def prepare_template(
    background_tasks: BackgroundTasks,
    x_session_id: str = Header(..., alias="x-session-id"),
):
    """
    Start template preparation in background.
    Converts PPTX → PDF → PNG → HTML Template using Claude Vision.
    """
    try:
        session = await get_session(x_session_id)
        if not session:
            return PrepareTemplateResponse(success=False, error="Session not found")

        template_path = session.get("template_path")
        if not template_path:
            return PrepareTemplateResponse(
                success=False, error="No template uploaded for this session"
            )

        current_status = session.get("template_preparation_status")
        current_html_url = session.get("html_template_url")

        if current_status == "completed" and current_html_url:
            return PrepareTemplateResponse(
                success=True, message="Template already prepared"
            )

        if current_status == "processing":
            return PrepareTemplateResponse(
                success=True, message="Template preparation already in progress"
            )

        background_tasks.add_task(
            process_template_preparation,
            session_id=x_session_id,
            template_path=template_path,
        )

        return PrepareTemplateResponse(
            success=True, message="Template preparation started"
        )

    except Exception as e:
        traceback.print_exc()
        return PrepareTemplateResponse(success=False, error=str(e))


@router.post(
    "/template-preparation-status",
    response_model=TemplatePreparationStatusResponse,
)
async def check_template_preparation_status(
    x_session_id: str = Header(..., alias="x-session-id"),
):
    """Check the status of template preparation."""
    try:
        status_info = await get_template_preparation_status(x_session_id)

        return TemplatePreparationStatusResponse(
            success=True,
            status=status_info.get("status", "pending"),
            htmlTemplateUrl=status_info.get("html_template_url"),
            templatePngUrls=status_info.get("template_png_urls"),
            templatePdfUrl=status_info.get("template_pdf_url"),
            error=status_info.get("error"),
        )

    except Exception as e:
        return TemplatePreparationStatusResponse(
            success=False, status="failed", error=str(e)
        )


@router.post("/list-slides-from-html", response_model=ListSlidesFromHtmlResponse)
async def list_slides_from_html(
    x_session_id: str = Header(..., alias="x-session-id"),
):
    """
    List all slides from the prepared HTML template.
    Parses the HTML to extract slide information (number, title, field count).
    """
    try:
        from bs4 import BeautifulSoup

        session = await get_session(x_session_id)
        if not session:
            return ListSlidesFromHtmlResponse(success=False, error="Session not found")

        status = session.get("template_preparation_status")
        html_url = session.get("html_template_url")

        if status != "completed" or not html_url:
            return ListSlidesFromHtmlResponse(
                success=False, error=f"Template not ready. Status: {status}"
            )

        html_content = await download_html_template(html_url)
        soup = BeautifulSoup(html_content, "html.parser")

        slides = []
        slide_divs = soup.find_all("div", class_="slide")

        for slide_div in slide_divs:
            slide_number = slide_div.get("data-slide-number")
            if slide_number:
                slide_number = int(slide_number)
            else:
                slide_number = len(slides) + 1

            title_elem = slide_div.find(class_="main-title")
            if not title_elem:
                title_elem = slide_div.find(["h1", "h2", "h3"])

            title = (
                title_elem.get_text(strip=True) if title_elem else f"Slide {slide_number}"
            )
            if len(title) > 60:
                title = title[:57] + "..."

            section_boxes = slide_div.find_all(class_="section-box")
            content_count = len(section_boxes) if section_boxes else 1

            slides.append(
                {
                    "slide_number": slide_number,
                    "title": title,
                    "field_count": content_count,
                    "layout": "content",
                }
            )

        return ListSlidesFromHtmlResponse(
            success=True, slides=slides, total=len(slides)
        )

    except Exception as e:
        traceback.print_exc()
        return ListSlidesFromHtmlResponse(success=False, error=str(e))


# --- Background Tasks ---

async def process_template_preparation(session_id: str, template_path: str):
    """
    Background task to convert PPTX → PDF → PNG → HTML.

    Steps:
    1. Update status to 'processing'
    2. Download PPTX from Storage
    3. Convert PPTX → PDF → PNG
    4. Generate HTML template with Claude Vision
    5. Upload HTML, PNGs, PDF to Storage
    6. Update session with URLs and status
    """
    start_time = time.time()

    try:
        print(f"[prepare-template] Starting for session {session_id}")

        await update_template_preparation_status(session_id, "processing")

        print(f"[prepare-template] Downloading template: {template_path}")
        pptx_bytes = await download_template(template_path)

        print(f"[prepare-template] Converting PPTX to images...")
        images, pdf_bytes = convert_pptx_to_images(pptx_bytes, return_pdf=True)
        print(f"[prepare-template] Generated {len(images)} slide images")

        print(f"[prepare-template] Generating HTML template with Claude Vision...")
        template_result = generate_html_template(images)
        html_template = template_result["full_html"]
        print(
            f"[prepare-template] Generated HTML with {len(template_result.get('fields', []))} fields"
        )

        timestamp = datetime.utcnow().strftime("%Y%m%d_%H%M%S")

        html_filename = f"template_{timestamp}.html"
        html_url = await upload_generated_html(session_id, html_template, html_filename)
        print(f"[prepare-template] Uploaded HTML to: {html_url}")

        png_urls = []
        for i, (img_bytes, _) in enumerate(images):
            png_filename = f"slide_{timestamp}_{i+1:02d}.png"
            png_url = await upload_png(session_id, img_bytes, png_filename)
            png_urls.append(png_url)
        print(f"[prepare-template] Uploaded {len(png_urls)} PNG images")

        pdf_url = None
        if pdf_bytes:
            pdf_filename = f"template_{timestamp}.pdf"
            pdf_url = await upload_pdf(session_id, pdf_bytes, pdf_filename)
            print(f"[prepare-template] Uploaded PDF to: {pdf_url}")

        await update_template_preparation_status(
            session_id,
            "completed",
            html_template_url=html_url,
            template_png_urls=png_urls,
            template_pdf_url=pdf_url,
        )

        elapsed = time.time() - start_time
        print(
            f"[prepare-template] Completed in {elapsed:.1f}s for session {session_id}"
        )

    except Exception as e:
        elapsed = time.time() - start_time
        error_msg = str(e)
        print(f"[prepare-template] Failed after {elapsed:.1f}s: {error_msg}")
        traceback.print_exc()

        await update_template_preparation_status(
            session_id, "failed", error=error_msg
        )
