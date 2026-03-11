"""
Supabase Client Service

Handles all database operations with Supabase.
"""

from supabase import create_client, Client
from typing import Dict, Any, Optional, List
import httpx

from app.config import SUPABASE_URL, SUPABASE_KEY


def get_supabase_client() -> Client:
    """Get Supabase client instance."""
    return create_client(SUPABASE_URL, SUPABASE_KEY)


async def ensure_session(session_id: str) -> Dict[str, Any]:
    """
    Get an existing session or create a minimal one.
    Used by fast-conversion to satisfy the generation_jobs FK.
    """
    session = await get_session(session_id)
    if session:
        return session

    supabase = get_supabase_client()
    result = supabase.table('sessions').insert({
        "id": session_id,
        "current_step": "generating",
    }).execute()

    if result.data:
        return result.data[0]
    raise RuntimeError("Failed to create session")


async def get_session(session_id: str) -> Optional[Dict[str, Any]]:
    """
    Get session data from Supabase.
    """
    supabase = get_supabase_client()
    try:
        result = supabase.table('sessions').select('*').eq('id', session_id).maybe_single().execute()
        return result.data if result.data else None
    except Exception:
        return None


async def get_mapping(session_id: str) -> Optional[Dict[str, Any]]:
    """
    Get mapping configuration for a session.
    """
    supabase = get_supabase_client()
    result = supabase.table('mappings').select('*').eq('session_id', session_id).single().execute()
    return result.data if result.data else None


async def get_fetched_projects(session_id: str) -> Optional[Dict[str, Any]]:
    """
    Get fetched projects data from session.
    """
    session = await get_session(session_id)
    if session and session.get('fetched_projects_data'):
        return session['fetched_projects_data']
    return None


async def download_template(template_path: str) -> bytes:
    """
    Download template file from Supabase Storage.

    Args:
        template_path: Path in the storage bucket (e.g., "session-id/template.pptx")

    Returns:
        File content as bytes
    """
    # Construct the storage URL
    storage_url = f"{SUPABASE_URL}/storage/v1/object/public/templates/{template_path}"

    async with httpx.AsyncClient() as client:
        response = await client.get(storage_url)
        response.raise_for_status()
        return response.content


async def upload_generated_html(session_id: str, html_content: str, filename: str = "report.html") -> str:
    """
    Upload generated HTML to Supabase Storage.

    Returns:
        Public URL of the uploaded file
    """
    supabase = get_supabase_client()

    file_path = f"{session_id}/{filename}"

    # Upload to storage (use 'outputs' bucket)
    result = supabase.storage.from_('outputs').upload(
        file_path,
        html_content.encode('utf-8'),
        {"content-type": "text/html"}
    )

    # Get public URL
    public_url = supabase.storage.from_('outputs').get_public_url(file_path)

    return public_url


async def upload_pdf(session_id: str, pdf_bytes: bytes, filename: str = "template.pdf") -> str:
    """
    Upload PDF file to Supabase Storage.

    Returns:
        Public URL of the uploaded file
    """
    supabase = get_supabase_client()

    file_path = f"{session_id}/{filename}"

    # Upload to storage (use 'outputs' bucket)
    result = supabase.storage.from_('outputs').upload(
        file_path,
        pdf_bytes,
        {"content-type": "application/pdf"}
    )

    # Get public URL
    public_url = supabase.storage.from_('outputs').get_public_url(file_path)

    return public_url


async def create_generation_job(
    session_id: str,
    engine: str = "claude-html",
    input_data: Optional[Dict[str, Any]] = None
) -> str:
    """
    Create a new generation job in the database.

    Returns:
        Job ID
    """
    supabase = get_supabase_client()

    job_data = {
        "session_id": session_id,
        "status": "pending",
        "engine": engine,
        "input_data": input_data or {}
    }

    result = supabase.table('generation_jobs').insert(job_data).execute()

    if result.data:
        return result.data[0]['id']

    raise RuntimeError("Failed to create generation job")


async def update_job_status(
    job_id: str,
    status: str,
    result: Optional[Dict[str, Any]] = None,
    error: Optional[str] = None
) -> None:
    """
    Update a generation job's status.
    """
    supabase = get_supabase_client()

    update_data = {"status": status}

    if status == "processing":
        update_data["started_at"] = "now()"
    elif status in ("completed", "failed"):
        update_data["completed_at"] = "now()"

    if result:
        update_data["result"] = result

    if error:
        update_data["error"] = error

    supabase.table('generation_jobs').update(update_data).eq('id', job_id).execute()


async def save_generated_report(
    session_id: str,
    html_path: str,
    engine: str = "claude-html"
) -> str:
    """
    Save a reference to a generated report.

    Returns:
        Report ID
    """
    supabase = get_supabase_client()

    # Get current iteration count
    count_result = supabase.table('generated_reports').select('id', count='exact').eq('session_id', session_id).execute()
    iteration = (count_result.count or 0) + 1

    report_data = {
        "session_id": session_id,
        "engine": engine,
        "pptx_path": html_path,  # Using pptx_path field for HTML path too
        "iteration": iteration
    }

    result = supabase.table('generated_reports').insert(report_data).execute()

    if result.data:
        return result.data[0]['id']

    raise RuntimeError("Failed to save generated report")


async def upload_pptx(session_id: str, pptx_bytes: bytes, filename: str = "report.pptx") -> str:
    """
    Upload generated PPTX file to Supabase Storage.

    Returns:
        Public URL of the uploaded file
    """
    supabase = get_supabase_client()

    file_path = f"{session_id}/{filename}"

    result = supabase.storage.from_('outputs').upload(
        file_path,
        pptx_bytes,
        {"content-type": "application/vnd.openxmlformats-officedocument.presentationml.presentation"}
    )

    public_url = supabase.storage.from_('outputs').get_public_url(file_path)

    return public_url


async def get_job_status(job_id: str) -> Optional[Dict[str, Any]]:
    """
    Get the current status of a generation job.
    """
    supabase = get_supabase_client()
    result = supabase.table('generation_jobs').select('*').eq('id', job_id).single().execute()
    return result.data if result.data else None


async def update_template_preparation_status(
    session_id: str,
    status: str,
    html_template_url: Optional[str] = None,
    template_png_urls: Optional[List[str]] = None,
    template_pdf_url: Optional[str] = None,
    error: Optional[str] = None
) -> None:
    """
    Update template preparation status in session.

    Args:
        session_id: Session ID
        status: 'pending' | 'processing' | 'completed' | 'failed'
        html_template_url: URL to the generated HTML template
        template_png_urls: List of URLs to PNG images of slides
        template_pdf_url: URL to the PDF version
        error: Error message if failed
    """
    supabase = get_supabase_client()

    update_data: Dict[str, Any] = {
        "template_preparation_status": status,
        "updated_at": "now()"
    }

    if html_template_url:
        update_data["html_template_url"] = html_template_url

    if template_png_urls:
        update_data["template_png_urls"] = template_png_urls

    if template_pdf_url:
        update_data["template_pdf_url"] = template_pdf_url

    if error:
        update_data["template_preparation_error"] = error
    elif status == 'completed':
        # Clear any previous error
        update_data["template_preparation_error"] = None

    supabase.table('sessions').update(update_data).eq('id', session_id).execute()


async def get_template_preparation_status(session_id: str) -> Dict[str, Any]:
    """
    Get template preparation status from session.

    Returns:
        Dict with status, html_template_url, error, etc.
    """
    session = await get_session(session_id)

    if not session:
        return {"status": "pending", "error": "Session not found"}

    return {
        "status": session.get("template_preparation_status", "pending"),
        "html_template_url": session.get("html_template_url"),
        "template_png_urls": session.get("template_png_urls"),
        "template_pdf_url": session.get("template_pdf_url"),
        "error": session.get("template_preparation_error"),
        "template_path": session.get("template_path")
    }


async def upload_png(session_id: str, png_bytes: bytes, filename: str) -> str:
    """
    Upload PNG image to Supabase Storage.

    Returns:
        Public URL of the uploaded file
    """
    supabase = get_supabase_client()

    file_path = f"{session_id}/{filename}"

    result = supabase.storage.from_('outputs').upload(
        file_path,
        png_bytes,
        {"content-type": "image/png"}
    )

    public_url = supabase.storage.from_('outputs').get_public_url(file_path)

    return public_url


async def download_html_template(html_url: str) -> str:
    """
    Download HTML template from Supabase Storage URL.

    Args:
        html_url: Public URL of the HTML file

    Returns:
        HTML content as string
    """
    async with httpx.AsyncClient() as client:
        response = await client.get(html_url)
        response.raise_for_status()
        return response.text
