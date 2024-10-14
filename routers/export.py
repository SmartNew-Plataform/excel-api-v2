from fastapi import APIRouter, Request, HTTPException
from .service.export_xlsx import export_record
from .service.export_xlsx_unified import export_record_unified
router = APIRouter()

@router.post("/export")
async def export_xlsx(request: Request):
    try:
        response = await request.json()
    except:
        raise HTTPException(status_code=400, detail="Falha na requisição JSON")
    return export_record(response)

@router.post("/export-unified")
async def export_xlsx(request: Request):
    try:
        response = await request.json()
    except:
        raise HTTPException(status_code=400, detail="Falha na requisição JSON")
    return export_record_unified(response)