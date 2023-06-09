# -*- encoding: utf-8 -*-
# @Author: SWHL
# @Contact: liekkaskono@163.com
import argparse
from pathlib import Path
from typing import Dict, List, Optional, Tuple, Union

import pandas as pd
import pptx
from pptx import Presentation

from .utils import mkdir


class ExtractPPT():
    def __init__(self, ):
        pass

    def __call__(self, ppt_path: str, save_img_dir: str = None) -> List:
        """Extract content and images of ppt.

        Args:
            ppt_path (str): the path of ppt.
            save_img_dir (str, optional): The directory for saving images. Defaults to None.

        Returns:
            List: txts from pptx.
        """
        if not Path(ppt_path).exists():
            raise FileNotFoundError(f'{ppt_path} does not exist.')

        txts, imgs = self.extract_all(ppt_path)

        if save_img_dir:
            self.save_img(imgs, save_img_dir)
        return list(txts.values())

    def extract_all(self, ppt_path: str) -> Tuple[Dict, Dict]:
        prs = Presentation(ppt_path)
        extract_txts, extract_imgs = {}, {}
        for i, slide in enumerate(prs.slides):
            cur_page = i + 1
            cur_txts, cur_imgs = self.extract_one(slide)

            extract_txts[cur_page] = '\n'.join(cur_txts)
            for cur_img in cur_imgs:
                extract_imgs.setdefault(cur_page, []).append(cur_img)
        return extract_txts, extract_imgs

    def extract_one(self, slide: pptx.slide.Slide) -> List:
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
                img_bytes = self.extract_image(shape.image)
                cur_page_imgs.append(img_bytes)
            else:
                pass
        return cur_page_content, cur_page_imgs

    @staticmethod
    def extract_text(shape_text: str) -> Optional[str]:
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
    def extract_image(img_value: pptx.parts.image.Image) -> bytes:
        return img_value.blob

    @staticmethod
    def save_img(imgs: Dict, save_img_dir: Union[str, Path]) -> None:
        mkdir(save_img_dir)
        for page_num, img_list in imgs.items():
            for i, img in enumerate(img_list):
                save_full_path = Path(save_img_dir) / f'{page_num}_{i+1}.png'
                with open(str(save_full_path), 'wb') as f:
                    f.write(img)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('ppt_path', type=str)
    parser.add_argument('-img_dir', '--save_img_dir', type=str, default=None)
    args = parser.parse_args()

    ppt_extracter = ExtractPPT()
    res = ppt_extracter(args.ppt_path, save_img_dir=args.save_img_dir)
    print(res)


if __name__ == '__main__':
    main()
