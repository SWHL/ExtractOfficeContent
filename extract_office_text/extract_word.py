#! /usr/bin/env python
# Modified from https://github.com/ankushshah89/python-docx2txt
import argparse
import os
import re
import sys
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

import docx
import pandas as pd
from docx.document import Document
from docx.oxml.table import CT_Tbl
from docx.table import Table, _Cell
from .utils import is_contain, mkdir


class ExtractWord():
    def __init__(self, ):
        self.img_suffix = [".jpg", ".jpeg", ".png", ".bmp"]
        self.nsmap = {
            'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
        }
        self.extract_table = ExtractWordTable()

    def __call__(self, docx_path: str, img_dir=None):
        self.table_content = self.extract_table(docx_path)
        text = ''

        # unzip the docx_path in memory
        zipf = zipfile.ZipFile(docx_path)
        filelist = zipf.namelist()

        header_files, footer_files, img_files = [], [], []
        header_xmls = 'word/header[0-9]*.xml'
        footer_xmls = 'word/footer[0-9]*.xml'

        for fname in filelist:
            if re.match(header_xmls, fname):
                header_files.append(fname)
            elif re.match(footer_xmls, fname):
                footer_files.append(fname)
            elif Path(fname).suffix.lower() in self.img_suffix:
                img_files.append(fname)
            else:
                continue

        # get header text
        # there can be 3 header files in the zip
        header_text = [self.xml2text(zipf.read(path)) for path in header_files]
        text += ''.join(header_text)

        # get main text
        doc_xml = 'word/document.xml'
        main_txt = self.xml2text(zipf.read(doc_xml))
        text += main_txt

        # get footer text
        # there can be 3 footer files in the zip
        footer_text = [self.xml2text(zipf.read(path)) for path in footer_files]
        text += ''.join(footer_text)

        if img_dir:
            mkdir(img_dir)
            for img_path in img_files:
                dst_fname = Path(img_dir) / Path(img_path).name
                with open(dst_fname, "wb") as dst_f:
                    dst_f.write(zipf.read(img_path))
        zipf.close()
        return text.strip()

    def qn(self, tag):
        """
        Stands for 'qualified name', a utility function to turn a namespace
        prefixed tag name into a Clark-notation qualified tag name for lxml. For
        example, ``qn('p:cSld')`` returns ``'{http://schemas.../main}cSld'``.
        Source: https://github.com/python-openxml/python-docx/
        """
        prefix, tagroot = tag.split(':')
        uri = self.nsmap[prefix]
        return f'{{{uri}}}{tagroot}'

    def xml2text(self, xml):
        """
        A string representing the textual content of this run, with content
        child elements like ``<w:tab/>`` translated to their Python
        equivalent.
        Adapted from: https://github.com/python-openxml/python-docx/
        """
        text = ''
        root = ET.fromstring(xml)
        for child in root.iter():
            if child.tag == self.qn('w:t'):
                t_text = child.text
                text += t_text if t_text is not None else ''
            elif child.tag == self.qn('w:tab'):
                text += '\t'
            elif child.tag in (self.qn('w:br'), self.qn('w:cr')):
                text += '\n'
            elif child.tag == self.qn("w:p"):
                text += '\n\n'
        return text


class ExtractWordTable():
    def __init__(self,):
        pass

    def __call__(self, docx_path):
        curr_content = []
        doc = docx.Document(docx_path)
        for block in self.iter_block_items(doc):
            if is_contain(block.style.name, ['Table', 'Table Grid']):
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

    def get_table_dataframe(self, table: docx.table.Table) -> pd.DataFrame:
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


def main():
    parser = argparse.ArgumentParser(description='A pure python-based utility '
                                                 'to extract text and images '
                                                 'from docx files.')
    parser.add_argument("docx", help="path of the docx file")
    parser.add_argument('-i', '--img_dir', help='path of directory '
                                                'to extract images')
    args = parser.parse_args()

    if not os.path.exists(args.docx):
        print('File {} does not exist.'.format(args.docx))
        sys.exit(1)

    if args.img_dir is not None:
        if not os.path.exists(args.img_dir):
            try:
                os.makedirs(args.img_dir)
            except OSError:
                print("Unable to create img_dir {}".format(args.img_dir))
                sys.exit(1)
    return args


if __name__ == '__main__':
    main()