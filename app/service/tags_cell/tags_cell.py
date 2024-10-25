from .data_validation import data_validation
from .conditional_format import conditional_format

def add_tags_cell(wb,ws,cell_index,row,col,tags_cell):
    print('add_tags_cell(wb,ws,cell_index,tags_cell)')
    
    if not isinstance(tags_cell, dict):
        return
    
    if 'data_validation' in tags_cell:
        data_validation(wb, ws, cell_index, tags_cell['data_validation'])
    
    if 'conditional_format' in tags_cell:
        conditional_format(wb, ws, cell_index, tags_cell['conditional_format'],row)

