# -*- encoding: utf-8 -*-
from pptx import Presentation


ppt_path = 'test_files/test_1.pptx'

prs = Presentation(ppt_path)
for i, slide in enumerate(prs.slides):
    print(f'---------------{i}-----------------')
    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue

        print(shape.text)
