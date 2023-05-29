# -*- encoding: utf-8 -*-
import pandas as pd
from pptx import Presentation
import pptx


def extract_table(table_value: pptx.table.Table) -> str:
    table_list = []
    for one_row in table_value.rows:
        each = ''
        for cell in one_row.cells:
            each += cell.text_frame.text + ','
        table_list.append(each)
    table_df = pd.DataFrame(table_list)
    return table_df.to_string()


ppt_path = 'test_files/test_1.pptx'


prs = Presentation(ppt_path)
data = []
for i, slide in enumerate(prs.slides):
    print(f'---------------{i}-----------------')
    for shape in slide.shapes:
        if shape.has_text_frame:
            pass
        elif shape.has_table:
            table_str = extract_table(shape.table)
        elif shape.has_chart:
            pass
        else:
            pass

print('ok')
