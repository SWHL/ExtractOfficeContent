# -*- encoding: utf-8 -*-
# @Author: SWHL
# @Contact: liekkaskono@163.com
from pathlib import Path

from extract_office_text import ExtractPPT


ppt_path = 'tests/test_files/ppt_example.pptx'

ppt_extracter = ExtractPPT()

save_dir = 'outputs'
save_img_dir = Path(save_dir) / Path(ppt_path).stem

res = ppt_extracter(ppt_path,
                    is_save_img=True,
                    save_img_dir=str(save_img_dir),
                    is_save_to_txt=True,
                    save_txt_dir=save_dir)
print(res)
