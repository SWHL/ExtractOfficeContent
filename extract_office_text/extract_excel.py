# -*- encoding: utf-8 -*-
# @Author: SWHL
# @Contact: liekkaskono@163.com
import tempfile
import uuid
from pathlib import Path
from typing import List

import openpyxl
import pandas as pd
from openpyxl.workbook.workbook import Workbook


class ExtractExcel():
    def __init__(self,):
        pass

    def __call__(self, excel_path: str, out_format: str = 'markdown') -> List:
        wb = self.unmerge_cell(excel_path)
        data_table = self.extract(wb, out_format)
        return data_table

    def unmerge_cell(self, file_name: str) -> Workbook:
        wb = openpyxl.load_workbook(file_name)
        for sheet_name in wb.sheetnames:
            sheet = wb[sheet_name]
            self.unmerge_and_fill_cells(sheet)
        return wb

    def unmerge_and_fill_cells(self, worksheet: Workbook) -> None:
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

    def extract(self, wb: Workbook, out_format: str) -> List:
        sheet_names = wb.sheetnames
        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp_save_path = Path(tmp_dir) / f'{uuid.uuid1()}.xlsx'
            wb.save(str(tmp_save_path))
            wb.close()

            df_datas = []
            for name in sheet_names:
                data = pd.read_excel(tmp_save_path, index_col=None,
                                     sheet_name=name)
                cvt_data = self.convert_table(data, out_format)
                df_datas.append(cvt_data)
        return df_datas

    @staticmethod
    def convert_table(df_table: pd.core.frame.DataFrame,
                      out_format: str) -> str:
        if 'to_' not in out_format:
            out_format = f'to_{out_format}'

        try:
            return getattr(df_table, out_format)()
        except AttributeError as exc:
            raise AttributeError(f'{out_format} is not supported.') from exc
