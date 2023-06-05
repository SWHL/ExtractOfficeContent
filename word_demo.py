# -*- encoding: utf-8 -*-
# @Author: SWHL
# @Contact: liekkaskono@163.com
from extract_office_text import ExtractWord


word_extract = ExtractWord()

word_path = 'tests/test_files/word_example.docx'
text = word_extract(word_path, "outputs/word")
print(text)
