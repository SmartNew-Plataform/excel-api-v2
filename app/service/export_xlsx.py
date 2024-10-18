import jsonify
import xlsxwriter
import io
import json
from io import BytesIO

from starlette.responses import StreamingResponse
from app.service.tags.tags import add_tags

def export_record(response):   

    try:
        filename = response['filename']

        output = get_output(response)

        return StreamingResponse(output,
                                 media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                                 headers={"Content-Disposition": f"attachment; filename={filename}"})
    except Exception as e:
   
        return jsonify({
            "message": "Erro ao gravar Excel",
            "error": str(e),
            "status": "400"
        }, 400)

def get_output(response) -> BytesIO:

    validResponse(response)

    output = io.BytesIO()

    wb = xlsxwriter.Workbook(output, {'in_memory': True})

    createSheets(wb, response['sheets'])

    wb.close()
    output.seek(0)

    return output

def createSheets(wb,sheets):
    for sheet in sheets:
        createSheet(wb,sheet)    

def formatCell(wb,format):
    if not bool(format):
        return None
    
    return wb.add_format(format)

def checkFormat(value, old_format):

    if (
            isinstance(value, dict) == False
            or 'format' not in value
            or 'value' not in value
    ):
        return value, old_format, None

    new_value = value['value']
    new_format = value['format']
    new_tags = None

    if old_format is not None:
        old_format.update(new_format)
        new_format = old_format

    if 'tags' in value:
        new_tags = value['tags']

    return new_value, new_format, new_tags

def __update_format(format_1, format_2):
    if format_1 is None and format_2 is None:
        return None

    if format_1 is not None and format_2 is None:
        return json.loads(json.dumps(format_1))

    if format_2 is not None and format_1 is None:
        return json.loads(json.dumps(format_2))

    format_1_copy = json.loads(json.dumps(format_1))
    format_2_copy = json.loads(json.dumps(format_2))
    format_1_copy.update(format_2_copy)
    return format_1_copy


def createRecord(wb,ws,records_format,id_row_records,index_record,record,format_table):
    
    line = id_row_records + index_record

    format_row = None
    if len(record) > 0:
        position_0 = record[0]
        if (
                position_0 is not None
                and isinstance(position_0, dict)
                and 'format' in position_0
        ):
            format_row = position_0['format']
    format_row = __update_format(format_table,format_row)

    for col, item_record in enumerate(record):
        if col > 0:
            value = item_record
            format = json.loads(json.dumps(records_format[col-1]))
            format = __update_format(format_row, format)
            tags = None
            value, format, tags = checkFormat(value, format)
            format_wb = formatCell(wb,format)

            ws.write(line, col-1, value, format_wb)
        
        
def createRecords(wb,ws,id_row_records,records,records_format,format_table):
    for index, record in enumerate(records):
            createRecord(wb,ws,records_format,id_row_records+1,index,record,format_table)

def createRecordHeader(wb,ws,id_row_records,id_col,record_header,format_table_top):

    format_header =  record_header['formatHeader'] if 'formatHeader' in record_header else None

    format_header = __update_format(format_table_top,format_header)

    format_wb = formatCell(wb,format_header)

    ws.write(id_row_records, id_col, record_header['nameHeader'],format_wb)
    
def createRecordHeaders(wb,ws,idRowRecords,recordHeaders,format_table_top):
    for index, recordHeader in enumerate(recordHeaders):
        createRecordHeader(wb,ws,idRowRecords,index,recordHeader,format_table_top)

def createBlock(wb,ws,block):

    validAttributes(block)    

    message = block['message']
    colBegin = block['colBegin']
    colEnd = block['colEnd']
    formatConfig = block['format']
    
    ws.merge_range(f"{colBegin}:{colEnd}".upper(), message, formatCell(wb,formatConfig))
        
def createBlocks(wb,ws,blocks):
    for index, block in enumerate(blocks):
        createBlock(wb,ws,block)

def createHeader(wb,ws,header):
    if 'blocks' in header:
        blocks=header['blocks']
        createBlocks(wb,ws,blocks)

def createHeaders(wb,ws,headers):
    for index, header in enumerate(headers):
        createHeader(wb,ws,header)

def createSheet(wb,sheet):
    validSheetName(sheet)
    ws = wb.add_worksheet(sheet['sheetName'])

    if not bool(sheet):
        return 
    
    idRowRecords = 0
    headers = ""

    format_table = None
    if 'formatTable' in sheet:
        format_table = sheet['formatTable']

    format_table_top = None
    if 'formatTableTop' in sheet:
        format_table_top = sheet['formatTableTop']
    format_table_top = __update_format(format_table,format_table_top)

    if 'headers' in sheet:
        headers=sheet['headers']
        idRowRecords = len(headers)
    if 'recordHeader' in sheet:
        records=sheet['recordHeader']
        createRecordHeaders(wb,ws,idRowRecords,records,format_table_top)
    if 'records' in sheet and 'recordsFormat' in sheet:
        records=sheet['records']
        recordsFormat=sheet['recordsFormat']
        createRecords(wb,ws,idRowRecords,records,recordsFormat,format_table)

    ws.autofit()

    if 'tags' in sheet:
        tags=sheet['tags']
        add_tags(wb,ws,tags)

    if 'headers' in sheet:
        createHeaders(wb,ws,headers)                


def validAttributes(block):
    attributes = ['message','colBegin','colEnd','format']
    
    for attribute in attributes:
        if attribute not in block:
            raise Exception(f"attribute {attribute} not found")    

def validResponse(response):
    attributes = ['filename','sheets']
    
    for attribute in attributes:
        if attribute not in response:
            raise Exception(f"body not attribute '{attribute}' ")

def validSheetName(sheet):
    if 'sheetName' not in sheet:
        raise Exception("attribute sheetName not found")