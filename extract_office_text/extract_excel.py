# -*- encoding: utf-8 -*-
# @Author: SWHL
# @Contact: liekkaskono@163.com
import tempfile
from pathlib import Path

import openpyxl
import pandas as pd


class ExtractExcel():
    def __init__(self,):
        pass

    def __call__(self, excel_path: str):
        df = self.unmerge_cell(excel_path)

    def unmerge_cell(self, file_name: str):
        """读取原始xlsx文件，拆分并填充单元格，然后生成中间临时文件

        Args:
            file_name (str): The file path to extract text of excel.

        Returns:
            _type_: _description_
        """
        wb = openpyxl.load_workbook(file_name)
        sheet_names = wb.sheetnames
        for sheet_name in wb.sheetnames:
            sheet = wb[sheet_name]
            self.unmerge_and_fill_cells(sheet)

        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp_save_path = Path(tmp_dir) / Path(file_name).name
            wb.save(str(tmp_save_path))
            wb.close()

            df_datas = []
            for name in sheet_names:
                data = pd.read_excel(tmp_save_path, index_col=None,
                                     sheet_name=name)
                df_datas.append(data)
            # TODO：不同Sheet 内容合并
        return data

    def unmerge_and_fill_cells(self, worksheet):
        """
        # 拆分所有的合并单元格，并赋予合并之前的值。
        # 由于openpyxl并没有提供拆分并填充的方法，所以使用该方法进行完成
        """

        all_merged_cell_ranges = list(worksheet.merged_cells.ranges)

        for merged_cell_range in all_merged_cell_ranges:
            merged_cell = merged_cell_range.start_cell
            worksheet.unmerge_cells(range_string=merged_cell_range.coord)

            for row_index, col_index in merged_cell_range.cells:
                cell = worksheet.cell(row=row_index, column=col_index)
                cell.value = merged_cell.value
