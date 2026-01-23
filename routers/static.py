from fastapi import APIRouter, Request, Depends
from fastapi.responses import HTMLResponse, JSONResponse, FileResponse
from fastapi.templating import Jinja2Templates
import config

router = APIRouter()
templates = Jinja2Templates(directory=config.TEMPLATE_DIR)

@router.get("/", response_class=HTMLResponse)
async def main_page(request: Request):
    return templates.TemplateResponse("index.html", {"request": request})

@router.get("/api/specification/{panel_type}")
async def get_specification(panel_type: str):
    """
    패널 타입에 따른 Specification HTML 반환
    """
    if panel_type not in ['GE', 'SA']:
        return JSONResponse({"error": "Invalid panel type"}, status_code=400)

    spec_file = config.TEMPLATE_DIR / f"{panel_type}_Specification.html"

    if not spec_file.exists():
        return JSONResponse({"error": "Specification file not found"}, status_code=404)

    return FileResponse(spec_file, media_type="text/html")

@router.get("/api/gene-content/{content_type}")
async def get_gene_content(content_type: str):
    """
    Gene Content HTML 반환
    """
    valid_types = ['GE_Gene_Content_DRNA', 'SA_Gene_Content_DNA', 'SA_Gene_Content_RNA']

    if content_type not in valid_types:
        return JSONResponse({"error": "Invalid content type"}, status_code=400)

    gene_content_file = config.TEMPLATE_DIR / f"{content_type}.html"

    if not gene_content_file.exists():
        return JSONResponse({"error": "Gene content file not found"}, status_code=404)

    return FileResponse(gene_content_file, media_type="text/html")
