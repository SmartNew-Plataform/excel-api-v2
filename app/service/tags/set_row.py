
def set_rows(wb, ws, set_row_json):
    print('__set_rows')
    for set_row in set_row_json:
        __set_row(wb, ws, set_row)

def __set_row(wb, ws, set_row):
    print('__set_row')
    if 'row' in set_row and 'height' in set_row:

        cell_format = None
        if 'format' in set_row:
            cell_format = set_row['format']

        options = None
        if 'options' in set_row:
            options = set_row['options']

        row = set_row['row']
        height = set_row['height']

        ws.set_row(row,height,cell_format,options)