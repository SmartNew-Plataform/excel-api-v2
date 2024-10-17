import jsonify
import xlsxwriter
import io
import json
from io import BytesIO

from starlette.responses import StreamingResponse

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
            or 'tags' not in value
    ):
        return value, old_format, None


    new_value = value['value']
    new_format = value['format']

    if old_format is not None:
        old_format.update(new_format)
        new_format = old_format

    return new_value, new_format, None

def createRecord(wb,ws,recordsFormat,idRowRecords,indexRecord,record):
    
    line = idRowRecords + indexRecord
    
    for col, itemRecord in enumerate(record):
        value = itemRecord
        format = json.loads(json.dumps(recordsFormat[col]))
        tags = None
        print(1, value, format)
        value, format, tags = checkFormat(value, format)
        print(2,value, format)
        formatWb = formatCell(wb,format)
        
        ws.write(line, col, value, formatWb)
        
        
def createRecords(wb,ws,idRowRecords,records,recordsFormat):
    for index, record in enumerate(records):
        createRecord(wb,ws,recordsFormat,idRowRecords+1,index,record) 

def createRecordHeader(wb,ws,idRowRecords,idCol,recordHeader):

    formatWb = formatCell(wb,recordHeader['formatHeader']) if 'formatHeader' in recordHeader else None

    ws.write(idRowRecords, idCol, recordHeader['nameHeader'],formatWb)
    
def createRecordHeaders(wb,ws,idRowRecords,recordHeaders):
    for index, recordHeader in enumerate(recordHeaders):
        createRecordHeader(wb,ws,idRowRecords,index,recordHeader) 

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

    if 'headers' in sheet:
        headers=sheet['headers']
        idRowRecords = len(headers)
    if 'recordHeader' in sheet:
        records=sheet['recordHeader']
        createRecordHeaders(wb,ws,idRowRecords,records)
    if 'records' in sheet and 'recordsFormat' in sheet:
        records=sheet['records']
        recordsFormat=sheet['recordsFormat']
        createRecords(wb,ws,idRowRecords,records,recordsFormat)

    ws.autofit()

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