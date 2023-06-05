#! /usr/bin/env python
# Modified from https://github.com/ankushshah89/python-docx2txt
import argparse
import os
import re
import sys
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path


class ExtractWord():
    def __init__(self, ):
        self.img_suffix = [".jpg", ".jpeg", ".png", ".bmp"]
        self.nsmap = {'w':
            'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}

    def __call__(self, docx: str, img_dir=None):
        text = ''

        # unzip the docx in memory
        zipf = zipfile.ZipFile(docx)
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
        text += self.xml2text(zipf.read(doc_xml))

        # get footer text
        # there can be 3 footer files in the zip
        footer_text = [self.xml2text(zipf.read(path)) for path in footer_files]
        text += ''.join(footer_text)

        if img_dir:
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
