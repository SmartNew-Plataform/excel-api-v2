from io import BytesIO
import xlsxwriter
from app.service.tags.tags import add_tags
import logging

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

def export_record(response):
    logger.info("Iniciando export_record")
    validResponse(response)
    output = BytesIO()
    try:
        wb = xlsxwriter.Workbook(output, {'in_memory': True})
        for sheet in response['sheets']:
            logger.info(f"Processando sheet: {sheet['sheetName']}")
            validSheetName(sheet)
            ws = wb.add_worksheet(sheet['sheetName'])
            
            current_row = 0

            if 'title' in sheet:
                title = sheet['title']
                if not isinstance(title, dict):
                    raise ValueError("'title' deve ser um dicionário. Exemplo: {'text': 'Título', 'mergeColumns': 5, 'style': {...}}")
                title_text = title.get('text', '')
                merge_columns = title.get('mergeColumns', 1)
                if not isinstance(merge_columns, int) or merge_columns < 1:
                    raise ValueError("'mergeColumns' deve ser um inteiro maior ou igual a 1")

                title_style = title.get('style', {})
                title_format = wb.add_format({
                    'bold': title_style.get('bold', True),
                    'font_size': title_style.get('fontSize', 16),
                    'font_color': title_style.get('fontColor', '#000000'),
                    'bg_color': title_style.get('bgColor', '#FFFFFF'),
                    'align': title_style.get('align', 'center'),
                    'valign': title_style.get('valign', 'middle'),
                })
                ws.merge_range(current_row, 0, current_row, merge_columns - 1, title_text, title_format)
                current_row += 1

            records = sheet.get('records', [])
            records_format = sheet.get('records_format', [])
            format_table = sheet.get('format_table', {})
            tags = sheet.get('tags', {})

            header_style = sheet.get('headerStyle', {})
            header_format = wb.add_format({
                'bold': header_style.get('bold', True),
                'bg_color': header_style.get('bgColor', '#E0E0E0'),
            })
            if header_style.get('border', False):
                header_format.set_border()

            cell_style = sheet.get('cellStyle', {})
            cell_format = wb.add_format()
            if cell_style.get('border', False):
                cell_format.set_border()

            logger.info(f"Records: {records}, Records_format: {records_format}")
            createRecords(wb, ws, current_row, records, records_format, format_table, header_format, cell_format)
            add_tags(wb, ws, tags)
        wb.close()
    except Exception as e:
        logger.error(f"Erro ao gerar o Excel: {str(e)}", exc_info=True)
        raise ValueError(f"Erro ao gerar o Excel: {str(e)}")
    output.seek(0)
    logger.info("Arquivo Excel gerado, tamanho do output: %d bytes", output.getbuffer().nbytes)
    if output.getbuffer().nbytes < 100:
        raise ValueError("Arquivo Excel gerado está vazio ou corrompido.")
    return output

def validResponse(response):
    required_attributes = ['filename', 'sheets']
    for attribute in required_attributes:
        if attribute not in response:
            raise ValueError(f"O corpo da requisição não contém o atributo obrigatório '{attribute}'")

def validSheetName(sheet):
    if 'sheetName' not in sheet:
        raise ValueError("O atributo 'sheetName' é obrigatório em cada planilha")
    if not isinstance(sheet['sheetName'], str) or not sheet['sheetName'].strip():
        raise ValueError("O atributo 'sheetName' deve ser uma string não vazia")
    invalid_chars = ['*', '/', '\\', '[', ']', ':', '?']
    if any(char in sheet['sheetName'] for char in invalid_chars):
        raise ValueError(f"'sheetName' contém caracteres inválidos: {invalid_chars}. Sheet: {sheet}")

def createRecords(wb, ws, id_row_records, records, records_format, format_table, header_format, cell_format):
    if not records:
        logger.info("Nenhum record para processar")
        return
    expected_columns = len(records_format) if records_format else len(records[0])
    logger.info(f"Expected columns: {expected_columns}")
    for index, record in enumerate(records):
        if len(record) != expected_columns:
            raise ValueError(
                f"Registro na linha {index + 1} tem {len(record)} colunas, esperado {expected_columns}. Registro: {record}"
            )
        createRecord(wb, ws, records_format, id_row_records, index, record, format_table, header_format if index == 0 else cell_format)

def createRecord(wb, ws, records_format, id_row_records, index, record, format_table, cell_format):
    row = id_row_records + index
    logger.info(f"Escrevendo record na linha {row}: {record}")
    for col, value in enumerate(record):
        logger.info(f"Escrevendo na linha {row}, coluna {col}: {value} (tipo: {type(value)})")
        try:
            ws.write(row, col, value, cell_format)
        except Exception as e:
            logger.error(f"Erro ao escrever na linha {row}, coluna {col}: {str(e)}")
            raise ValueError(f"Erro ao escrever na linha {row}, coluna {col}: {str(e)}")