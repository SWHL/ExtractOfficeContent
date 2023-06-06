# -*- encoding: utf-8 -*-
# @Author: SWHL
# @Contact: liekkaskono@163.com
from typing import List, Union

import docx
import pandas as pd
from docx.document import Document
from docx.oxml.table import CT_Tbl
from docx.table import Table, _Cell


class ExtractWordTable():
    def __init__(self,):
        pass

    def __call__(self, docx_path):
        curr_content = ['']
        doc = docx.Document(docx_path)

        for block in self.iter_block_items(doc):
            if self.is_contain(block.style.name, ['Table', 'Table Grid']):
                df = self.get_table_dataframe(block)
                try:
                    curr_content.append(f'\n{df.to_string()}')
                except:
                    curr_content.append(f'\n{df}')
        return curr_content

    def iter_block_items(self, parent):
        if isinstance(parent, Document):
            # 判断传入的是否为word文档对象，是则获取文档内容的全部子对象
            parent_elm = parent.element.body
        elif isinstance(parent, _Cell):
            # 判断传入的是否为单元格，是则获取单元格内全部子对象
            parent_elm = parent._tc
        else:
            raise ValueError("something's not right")

        for child in parent_elm.iterchildren():
            if isinstance(child, CT_Tbl):
                yield Table(child, parent)

    def is_contain(self, sentence: str, key_words: Union[str, List],) -> bool:
        """sentences中是否包含key_words中任意一个"""
        return any(i in sentence for i in key_words)

    def get_table_dataframe(self, table: docx.table.Table):
        '''获取表格数据，转换为dataframe数据结构'''
        text = []
        if len(table.rows) == 1:
            for i in table.rows[0].cells:
                text.append(i.text)
            return text[-1]

        keys, table_data = None, []
        for i, row in enumerate(table.rows):
            # 获取表格一行的数据
            text = (cell.text for cell in row.cells)

            # 判断是否是表头
            if i == 0:
                keys = tuple(text)
                continue

            table_data.append(dict(zip(keys, text)))
        df = pd.DataFrame(table_data)
        return df


if __name__ == '__main__':
    t = DOCX_EXTRACT()

    docx_path = 'tests/test_files/word_example.docx'
    res = t(docx_path)

    print(res)
