from .set_row import set_rows
from .set_column import set_columns

def add_tags(wb,ws,tags):
    if not isinstance(tags, dict):
        return
    if 'set_row' in tags:
        set_rows(wb, ws, tags['set_row'])
    if 'set_column' in tags:
        set_columns(wb, ws, tags['set_column'])