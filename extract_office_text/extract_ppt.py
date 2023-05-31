# -*- encoding: utf-8 -*-
# @Author: SWHL
# @Contact: liekkaskono@163.com
from pathlib import Path
import numpy as np
import cv2
from typing import List

import pandas as pd
import pptx
from pptx import Presentation

from .utils import mkdir, write_txt


class ExtractPPTText():
    def __init__(self, ):
        pass

    def __call__(self, ppt_path: str,
                 is_save_to_txt: bool = False,
                 save_txt_dir: str = None,
                 is_save_img: bool = False,
                 save_img_dir: str = None) -> List:
        """
        是否将内容写入txt中？
        是否将图像识别OCR也并到txt中？
        是否单独保留图像到指定目录？

        Args:
            ppt_path (str): _description_

        Returns:
            List: _description_
        """
        if is_save_to_txt and save_txt_dir is None:
            raise ValueError(
                'When is_save_to_txt is True, save_txt_dir must not be None.')

        if is_save_img and save_img_dir is None:
            raise ValueError(
                'When is_save_img is True, save_img_dir must be not None.')

        extract_content = self.extract_all(ppt_path)

        if is_save_to_txt and save_txt_dir:
            full_txt_path = Path(save_txt_dir) / f'{Path(ppt_path).stem}.txt'
            write_txt(full_txt_path, extract_content)
        return extract_content

    def extract_all(self, ppt_path: str) -> List:
        prs = Presentation(ppt_path)
        extract_list = []
        for slide in prs.slides:
            cur_page_content = self.extract_one(slide)
            extract_list.extend(cur_page_content)
        return extract_list

    def extract_one(self, slide) -> List:
        cur_page_content, cur_page_imgs = [], []
        for shape in slide.shapes:
            if shape.has_text_frame:
                txt = self.extract_text(shape.text)
                if txt:
                    cur_page_content.append(txt)
            elif shape.has_table:
                table_str = self.extract_table(shape.table)
                cur_page_content.append(table_str)
            elif shape.has_chart:
                pass
            elif hasattr(shape, 'image'):
                img = self.extract_image(shape.image)
                cur_page_imgs.append(img)
            else:
                pass
        return cur_page_content, cur_page_imgs

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
    def extract_image(img_value):
        img_blob = img_value.blob
        img_np = np.frombuffer(img_blob, dtype=np.uint8)
        img_array = cv2.imdecode(img_np, cv2.IMREAD_COLOR)
        return img_array
