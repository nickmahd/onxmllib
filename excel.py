from copy import copy
from datetime import datetime
from parsers import ParsedXML
from pathlib import Path
from typing import Any, Callable, Optional

from dateutil.relativedelta import relativedelta
from openpyxl import Workbook, load_workbook
from openpyxl.cell.cell import Cell
from openpyxl.formula.translate import Translator
from openpyxl.worksheet.worksheet import Worksheet

from core import reduce, get_toname


class SheetHandler:
    def __init__(self, path: Path, workbook: Workbook, template: Worksheet, forceful: bool=False) -> None:
        self.path = path
        self.workbook = workbook
        self.template = template
        self.forceful = forceful

        self.worksheet: Worksheet = self.workbook.active

    def _copy_cell(cell: Cell, temp_cell: Cell) -> None:
        for attr in ['font', 'border', 'fill', 'number_format', 'alignment']:
            setattr(cell, attr, copy(getattr(temp_cell, attr)))
        cell.value = Translator(temp_cell.value, temp_cell.coordinate).translate_formula(cell.coordinate)

    def _set_cell(self, row: int, col: int, value: Any) -> None:
        cell = self.worksheet.cell(row=row, column=col)
        if self.forceful or not cell.value:
            cell.value = value
 
    def _search(self, kwd: str, col: int) -> Optional[list]:
        for row in range(1, self.worksheet.max_row):
            if self.worksheet.cell(row=row, column=col).value == kwd:
                return row
        return None

    def _paste(self, row: int) -> None:
        row -= 1  # convert to offset
        for r in range(1, self.template.max_row + 1):
            for c in range(1, self.template.max_column + 1):
                v = self.template.cell(row=r, column=c)
                cell = self.worksheet.cell(row=r+row, column=c)

                if self.forceful or not cell.value:
                    self._copy_cell(cell, v)
    
    def _set_sheet(self, sheet_name: str) -> None:
        if not (self.worksheet and self.worksheet.title == sheet_name):
            try:
                self.worksheet = self.workbook[sheet_name]
            except KeyError:
                self.worksheet = self.workbook.create_sheet(sheet_name)

            for col, dim in self.template.column_dimensions.items():
                self.worksheet.column_dimensions[col].width = dim.width
 
    def write(self) -> None:
        self.workbook.save(self.path)

class InvoiceHandler(SheetHandler):
    def _fill_column(self, head: int, col: int, date: datetime, revenue: float) -> None:
        self._set_cell(head+2, col, revenue)
        self._set_cell(head+4, col, date)

    def fill(self, file: ParsedXML) -> None:
        self.set_sheet(file.fund)

        month = file.date.strftime('%B')
        last_month = (file.date - relativedelta(months=1)).strftime('%B')

        row = self._search(month, 2) 

        if not row:
            row = self.worksheet.max_row + 1
            self._paste(row)
            self.worksheet.cell(row=row, column=2).value = month

        week = (file.date.day - 1) // 7
        col = week + 2

        if file.date.month != (file.date + relativedelta(days=1)).month:
            col += 1

        self._fill_column(row, col, file.date, file.net_rev)

        if col == 2 and self._search_month(last_month):
            self._set_month(last_month)
            self._fill_column(row, 8, file.date, file.net_rev)

class Summary(SheetHandler):
    def trade_results(self, day: int) -> float:
        ahead_energy_amt = self.worksheet.cell(row=3, column=day+1)
        real_energy_amt = self.worksheet.cell(row=4, column=day+1)
        return -sum(filter(None, [ahead_energy_amt, real_energy_amt]))
    
    def total(self, day: int) -> float:
        cols = [self.worksheet.cell(row=r, column=day+1).value for r in range(5, self.template.max_column + 1)]
        return self.trade_results - sum(filter(None, cols))

class FTR(SheetHandler):
    @property
    def total(self, day: int) -> float:
        cols = [self.worksheet.cell(row=r, column=day+1).value for r in range(3, self.template.max_column + 1)]
        return sum(filter(None, cols))


class HandlerRotater:
    def __init__(self, root: Path, template: Worksheet, doctype: str) -> None:
        self.handlers = {}
        self.template = template
        self.to_name = get_toname(doctype, root)
        self.type = reduce(doctype, Summary, FTR, InvoiceHandler)

    def _get_workbook(self, path: Path) -> Workbook:
        try:
            wb = load_workbook(path)
            return wb
        except FileNotFoundError:
            wb = Workbook()
            wb.remove(wb.active)
            return wb

    def get_handler(self, year: int, fund: str) -> SheetHandler:
        path = self.to_name(year, fund)
        if year not in self.handlers:
            self.handlers[year] = self.type(path=path, workbook=self._get_workbook(path), template=self.template)
        return self.handlers[year]

    def write(self) -> None:
        for handler in self.handlers.values():
            handler.write()