# -*- encoding: utf-8 -*-
# @Author: SWHL
# @Contact: liekkaskono@163.com
from pathlib import Path

from extract_office_content import ExtractPPT


ppt_path = 'tests/test_files/ppt_example.pptx'

ppt_extracter = ExtractPPT()

save_dir = 'outputs'
save_txt_path = '1.txt'
save_img_dir = Path(save_dir) / Path(ppt_path).stem

res = ppt_extracter(ppt_path,
                    save_img_dir=str(save_img_dir),
                    save_txt_path=save_txt_path)
print(res)
