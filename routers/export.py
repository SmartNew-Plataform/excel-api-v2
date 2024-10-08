from fastapi import APIRouter, Request, HTTPException
from fastapi.responses import JSONResponse
from .service.export_xlsx import export_record

router = APIRouter()

@router.post("/export")
async def export_xlsx(request: Request):
    try:
        response = await request.json()
    except:
        raise HTTPException(status_code=400, detail="Falha na requisição JSON")
    return export_record(response)
