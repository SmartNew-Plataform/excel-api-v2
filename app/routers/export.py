from fastapi import APIRouter, Request, HTTPException
from app.service.export_xlsx import export_record
from app.service.export_xlsx_unified import export_record_unified
router = APIRouter()

@router.post("/api/v1/export")
async def export_xlsx(request: Request):
    try:
        response = await request.json()
    except:
        raise HTTPException(status_code=400, detail="Falha na requisição JSON")
    return export_record(response)

@router.post("/api/v1/export-unified")
async def export_xlsx(request: Request):
    try:
        response = await request.json()
    except:
        raise HTTPException(status_code=400, detail="Falha na requisição JSON")
    return export_record_unified(response)