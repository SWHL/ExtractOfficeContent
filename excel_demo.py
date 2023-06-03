# -*- encoding: utf-8 -*-
# @Author: SWHL
# @Contact: liekkaskono@163.com

from extract_office_text import ExtractExcel

excel_extract = ExtractExcel()

excel_path = 'tests/test_files/excel_example.xlsx'

res  = excel_extract(excel_path, out_format='markdown')

print(res)
print('ok')
