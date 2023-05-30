# -*- encoding: utf-8 -*-
# @Author: SWHL
# @Contact: liekkaskono@163.com
from extract_office_text import ExtractPPTText


ppt_path = 'tests/test_files/ppt_example.pptx'

ppt_extracter = ExtractPPTText(ppt_path, save_dir='outputs')

res = ppt_extracter()
print(res)
