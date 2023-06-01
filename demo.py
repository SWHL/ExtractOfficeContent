# -*- encoding: utf-8 -*-
# @Author: SWHL
# @Contact: liekkaskono@163.com
from extract_office_text import ExtractPPTText


ppt_path = 'tests/test_files/ppt_example.pptx'

ppt_extracter = ExtractPPTText()

save_dir = 'outputs'
res = ppt_extracter(ppt_path,
                    is_save_img=True,
                    save_img_dir=f'{save_dir}/images',
                    is_save_to_txt=True,
                    save_txt_dir=save_dir)
print(res)
