# -*- encoding: utf-8 -*-
# @Author: SWHL
# @Contact: liekkaskono@163.com
import argparse
import filetype
from typing import Union
from pathlib import Path

from .extract_excel import ExtractExcel
from .extract_ppt import ExtractPPT
from .extract_word import ExtractWord


class ExtractOfficeContent():
    def __init__(self) -> None:
        self.excel = ExtractExcel()
        self.ppt = ExtractPPT()
        self.word = ExtractWord()

        self.doc_suffix = ['doc', 'docx']
        self.excel_suffix = ['xls', 'xlsx']
        self.ppt_suffix = ['ppt', 'pptx']

    def __call__(self, file_content: Union[Path, str],
                 save_img_dir: str = None):
        file_content = str(file_content)

        if not file_content:
            raise ValueError(f'{file_content} must be Path or str.')

        file_type = self.which_type(file_content)
        all_suffix = self.doc_suffix + self.excel_suffix + self.ppt_suffix
        if file_type not in all_suffix:
            raise ValueError(f'{file_type} must in {all_suffix}')

        if file_type in self.doc_suffix:
            return self.word(file_content, save_img_dir)

        if file_type in self.excel_suffix:
            return self.excel(file_content, save_img_dir=save_img_dir)

        if file_type in self.ppt_suffix:
            return self.ppt(file_content, save_img_dir)

    def which_type(self, file_content: Union[str, Path]):
        if isinstance(file_content, str):
            return filetype.guess(file_content).extension

        if isinstance(file_content, bytes):
            with open(file_content, 'rb') as f:
                data = f.read()

            return filetype.guess(data).extension

        raise ValueError(f'{file_content} must be [str, Path] type.')


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('file_path', type=str)
    parser.add_argument('-img_dir', '--save_img_dir', type=str, )
    args = parser.parse_args()

    extracter = ExtractOfficeContent()
    if Path(args.file_path).is_dir():
        file_list = list(Path(args.file_path).iterdir())
    else:
        file_list = [args.file_path]

    for file_one in file_list:
        res = extracter(str(file_one))
        print(res)


if __name__ == '__main__':
    main()
