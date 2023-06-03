# -*- encoding: utf-8 -*-
# @Author: SWHL
# @Contact: liekkaskono@163.com

from extract_office_text import ExtractExcel

excel_extract = ExtractExcel()

excel_path = 'tests/test_files/excel_with_image.xlsx'

res  = excel_extract(excel_path, out_format='markdown',
                     is_save_img=True, save_img_dir='1')

print(res)
print('ok')
