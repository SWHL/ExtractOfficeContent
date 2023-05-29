# -*- encoding: utf-8 -*-
# @Author: SWHL
# @Contact: liekkaskono@163.com
import pandas as pd
import pptx
from pptx import Presentation


class ExtractPPTText():
    def __init__(self, ppt_path: str):
        self.prs = Presentation(ppt_path)

    def __call__(self, ):
        data = []
        for slide in self.prs.slides:
            for shape in slide.shapes:
                if shape.has_text_frame:
                    txt = self.extract_text(shape.text)
                    if txt:
                        data.append(txt)
                elif shape.has_table:
                    table_str = self.extract_table(shape.table)
                    data.append(table_str)
                elif shape.has_chart:
                    pass
                else:
                    pass

                try:
                    if "image" in shape.image.content_type:
                        imgName = shape.image.filename
                        with open(imgName, "wb") as f:
                            f.write(shape.image.blob)
                        print(imgName)
                except:
                    continue
        return data

    @staticmethod
    def extract_text(shape_text):
        txt = shape_text.strip()
        if txt:
            return txt
        return None

    @staticmethod
    def extract_table(table_value: pptx.table.Table) -> str:
        table_list = []
        for one_row in table_value.rows:
            each = ''
            for cell in one_row.cells:
                each += cell.text_frame.text + ','
            table_list.append(each)
        table_df = pd.DataFrame(table_list)
        return table_df.to_string()

    @staticmethod
    def extract_image(img):
        pass


if __name__ == '__main__':
    ppt_path = 'test_files/test_1.pptx'

    ppt_extracter = ExtractPPTText(ppt_path)

    res = ppt_extracter()
    print(res)
