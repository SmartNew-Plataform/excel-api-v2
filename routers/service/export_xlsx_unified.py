import pandas as pd
from starlette.responses import StreamingResponse
from io import BytesIO

from .export_xlsx import get_output

def export_record_unified(response):
    output = get_output(response)

    df = pd.read_excel(output, sheet_name=None)

    output2 = BytesIO()

    with pd.ExcelWriter(output2, engine='openpyxl') as writer:
        start_row = 0

        for sheet_name, sheet_data in df.items():

            data = {sheet_name}
            df = pd.DataFrame(data)

            df.to_excel(writer, startrow=start_row, index=False, header=True)

            start_row += 1
            sheet_data.to_excel(writer, startrow=start_row, index=False, header=True)

            start_row += len(sheet_data) + 2

    output2.seek(0)

    return StreamingResponse(output2,
                             media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                             headers={"Content-Disposition": f"attachment; filename={'filename.xlsx'}"})
