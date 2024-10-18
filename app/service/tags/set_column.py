def set_columns(wb, ws, set_column_json):
    for set_column in set_column_json:
        __set_column(wb, ws, set_column)

def __set_column(wb, ws, set_column):
    if 'first_col' in set_column and 'last_col' in set_column:

        width = None
        if 'width' in set_column:
            width = set_column['width']

        cell_format = None
        if 'cell_format' in set_column:
            cell_format = set_column['cell_format']

        options = None
        if 'options' in set_column:
            options = set_column['options']

        first_col = set_column['first_col']
        last_col = set_column['last_col']


        ws.set_column(first_col,last_col,width,cell_format,options)