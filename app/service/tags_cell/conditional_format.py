def conditional_format(wb,ws,cell_index,json,row):
    
    if (
        'type'     in json and 
        'format'   in json and
        'criteria' in json
    ):
        json_type = json['type']
        json_criteria = json['criteria']
        json_criteria = json_criteria.replace('#POSITION#',cell_index).replace('#ROW#',str(row+1))
        json_format =  wb.add_format(json['format'])

        options = {
                'type' : json_type,
                'criteria' : json_criteria,
                'format': json_format
            }
    
        ws.conditional_format( cell_index, options) 

    '''
    Obs: no payload aonde ler ISBLANK(A1) troque por ISBLANK(#POSITION#)
                    aonde ler COUNTA(1:1) troque por COUNTA(#ROW#,row)
    
    format_vazio = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'})

    worksheet.conditional_format('A1', {
        'type':     'formula',
        'criteria': '=AND(ISBLANK(A1), COUNTA(1:1) >= 1)',
        'format':   format_vazio
    })
    '''