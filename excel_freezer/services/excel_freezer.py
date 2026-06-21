from os import path
from copy import copy
from datetime import datetime
from pathlib import Path

import openpyxl


class ExcelFreezer:
    def __init__(self, source: Path):
        self._source = source

    @property
    def default_output_file_name(self) -> str:
        today = datetime.now().strftime("%Y-%m-%d")
        file_name = path.splitext(path.basename(self._configuration.source))[0]

        return f"{today}_{file_name}.xlsx"

    def run(self):
        wb_data = openpyxl.load_workbook(self._configuration.source, data_only=True)
        wb_styles = openpyxl.load_workbook(self._configuration.source, data_only=False)

        for sheet_name in wb_styles.sheetnames:
            ws_style = wb_styles[sheet_name]
            ws_data = wb_data[sheet_name]

            for col_name, col_dim in ws_style.column_dimensions.items():
                ws_data.column_dimensions[col_name].width = col_dim.width
            for row_index, row_dim in ws_style.row_dimensions.items():
                ws_data.row_dimensions[row_index].height = row_dim.height

            for row in ws_style.iter_rows():
                for cell in row:
                    new_cell = ws_data.cell(row=cell.row, column=cell.column)

                    if cell.has_style:
                        new_cell.font = copy(cell.font)
                        new_cell.border = copy(cell.border)
                        new_cell.fill = copy(cell.fill)
                        new_cell.number_format = copy(cell.number_format)
                        new_cell.protection = copy(cell.protection)
                        new_cell.alignment = copy(cell.alignment)

        wb_data.save(self._configuration.destination)