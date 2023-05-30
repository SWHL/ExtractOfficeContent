# -*- encoding: utf-8 -*-
# @Author: SWHL
# @Contact: liekkaskono@163.com
from extract_office_text import ExtractPPTText


ppt_path = 'tests/test_files/简约活动策划方案汇报PPT模板.pptx'

ppt_extracter = ExtractPPTText(ppt_path, save_dir='outputs')

res = ppt_extracter()
print(res)
