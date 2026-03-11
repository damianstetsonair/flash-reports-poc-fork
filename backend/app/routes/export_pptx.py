"""
Export HTML → editable PPTX

POST /export-pptx
Body: { "html": "<full HTML document>" }
Response: binary .pptx file
"""

from fastapi import APIRouter, HTTPException
from fastapi.responses import Response
from pydantic import BaseModel

from app.services.pptx_generator import html_to_pptx

router = APIRouter(tags=["export"])


class ExportPptxRequest(BaseModel):
    html: str


@router.post("/export-pptx")
async def export_pptx(req: ExportPptxRequest):
    if not req.html or not req.html.strip():
        raise HTTPException(status_code=400, detail="html field is required")

    try:
        pptx_bytes = html_to_pptx(req.html)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"PPTX conversion failed: {e}")

    return Response(
        content=pptx_bytes,
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        headers={
            "Content-Disposition": 'attachment; filename="export.pptx"',
        },
    )
