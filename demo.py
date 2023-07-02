# -*- encoding: utf-8 -*-
# @Author: SWHL
# @Contact: liekkaskono@163.com
from pathlib import Path

from extract_office_content import ExtractOfficeContent

extracter = ExtractOfficeContent()
file_list = list(Path("tests/test_files").iterdir())

for file_path in file_list:
    res = extracter(file_path)
    print(res)
