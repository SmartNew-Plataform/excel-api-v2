from fastapi import APIRouter, Request, HTTPException, Response
from app.service.export_xlsx import export_record
import logging

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

router = APIRouter()

@router.post("/api/v1/export")
async def export_xlsx(request: Request):
    try:
        response = await request.json()
        logger.info(f"Payload recebido: {response}")

        required_fields = ["filename", "sheets"]
        missing_fields = [field for field in required_fields if field not in response]
        if missing_fields:
            raise HTTPException(
                status_code=400,
                detail=f"Campos obrigatórios ausentes: {', '.join(missing_fields)}. Exemplo: {{'filename': 'teste.xlsx', 'sheets': [...]}}"
            )

        if not isinstance(response["filename"], str) or not response["filename"].strip():
            raise HTTPException(
                status_code=400,
                detail="'filename' deve ser uma string não vazia. Exemplo: 'material.xlsx'"
            )
        if not response["filename"].lower().endswith('.xlsx'):
            raise HTTPException(
                status_code=400,
                detail="'filename' deve ter a extensão .xlsx. Exemplo: 'material.xlsx'"
            )

        if not isinstance(response["sheets"], list):
            raise HTTPException(
                status_code=400,
                detail="'sheets' deve ser uma lista. Exemplo: [{'sheetName': 'Sheet1', 'records': [...]}]"
            )
        if not response["sheets"]:
            raise HTTPException(
                status_code=400,
                detail="'sheets' não pode estar vazio. Deve conter pelo menos uma planilha."
            )

        for sheet in response["sheets"]:
            if not isinstance(sheet, dict):
                raise HTTPException(
                    status_code=400,
                    detail=f"Cada sheet deve ser um dicionário. Sheet inválido: {sheet}"
                )
            if "sheetName" not in sheet:
                raise HTTPException(
                    status_code=400,
                    detail=f"Cada sheet deve ter um 'sheetName'. Sheet inválido: {sheet}"
                )
            if not isinstance(sheet["sheetName"], str) or not sheet["sheetName"].strip():
                raise HTTPException(
                    status_code=400,
                    detail=f"'sheetName' deve ser uma string não vazia. Sheet inválido: {sheet}"
                )
            invalid_chars = ['*', '/', '\\', '[', ']', ':', '?']
            if any(char in sheet["sheetName"] for char in invalid_chars):
                raise HTTPException(
                    status_code=400,
                    detail=f"'sheetName' contém caracteres inválidos: {invalid_chars}. Sheet: {sheet}"
                )

            if "records" in sheet:
                if not isinstance(sheet["records"], list):
                    raise HTTPException(
                        status_code=400,
                        detail=f"'records' deve ser uma lista. Sheet: {sheet}"
                    )
                if sheet["records"]:
                    expected_columns = len(sheet["records"][0])
                    for index, record in enumerate(sheet["records"]):
                        if not isinstance(record, list):
                            raise HTTPException(
                                status_code=400,
                                detail=f"Cada registro em 'records' deve ser uma lista. Registro inválido na linha {index + 1}: {record}"
                            )
                        if len(record) != expected_columns:
                            raise HTTPException(
                                status_code=400,
                                detail=f"Registro na linha {index + 1} tem {len(record)} colunas, esperado {expected_columns}. Registro: {record}"
                            )
                        for value in record:
                            if not isinstance(value, (str, int, float, bool)) and value is not None:
                                raise HTTPException(
                                    status_code=400,
                                    detail=f"Valores em 'records' devem ser string, número, booleano ou null. Valor inválido na linha {index + 1}: {value}"
                                )

        result = export_record(response)
        logger.info("Arquivo Excel gerado com sucesso")
        return Response(
            content=result.getvalue(),
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": f"attachment; filename={response['filename']}"}
        )
    except ValueError as e:
        logger.error(f"Erro de validação: {str(e)}")
        raise HTTPException(status_code=400, detail=f"Erro no JSON: {str(e)}")
    except Exception as e:
        logger.error(f"Erro inesperado: {str(e)}", exc_info=True)
        raise HTTPException(status_code=500, detail=f"Erro inesperado: {str(e)}")