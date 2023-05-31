# -*- encoding: utf-8 -*-
# @Author: SWHL
# @Contact: liekkaskono@163.com
from pptx import Presentation


ppt_path = 'tests/test_files/ppt_example.pptx'
prs = Presentation(ppt_path)

slides = prs.slides

print('ok')