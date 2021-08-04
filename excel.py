import pickle
from abc import ABC, abstractmethod, abstractstaticmethod
from copy import copy
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, TypeVar

from dateutil.relativedelta import relativedelta
from openpyxl import Workbook, load_workbook
from openpyxl.cell.cell import Cell
from openpyxl.formula.translate import Translator
from openpyxl.worksheet.worksheet import Worksheet

from parsers import Invoice, Settlement

T = TypeVar(name="T", bound='SheetHandler')

class SheetHandler(ABC):
    def __init__(self, path: Path) -> None:
        self.path = path
        self.modified: Dict[Path, Workbook] = {}
        self.workbook: Optional[Workbook] = None
        self.worksheet: Optional[Worksheet] = None

    @staticmethod
    def get_handler(parsetype: str, path: Path, template: Path) -> T:
        with open(template, 'rb') as file:
            template = pickle.load(file)[parsetype]
        if parsetype == 'invoice':
            return InvoiceHandler(path, template)
        elif parsetype == 'settlement':
            pass

    @staticmethod
    def _get_adjacent_years(date: datetime) -> List[int]:
        years = [date.year]
        if date.month == 1:
            years.append(date.year - 1)
        elif date.month == 12:
            years.append(date.year + 1)
        return years

    def _set_workbook(self, *args) -> None:
        path = self._get_path(*args)
        if path not in self.modified:
            try:
                wb = load_workbook(path)
            except FileNotFoundError:
                wb = Workbook()
                wb.remove(wb.active)
            finally:
                self.modified[path] = wb
        self.workbook = self.modified[path]

    def _set_sheet(self, sheet_name: str) -> None:
        if not (self.worksheet and self.worksheet.title == sheet_name):
            try:
                self.worksheet = self.workbook[sheet_name]
            except KeyError:
                self.worksheet = self.workbook.create_sheet(sheet_name)
    
    def _search(self, kwd: str, col: int) -> Optional[list]:
        for row in range(1, self.worksheet.max_row):
            if self.worksheet.cell(row, col).value == kwd:
                return row
        return None

    def _copy_cell(self, cell: Cell, temp_cell: Cell) -> None:
        for attr in ['font', 'border', 'fill', 'number_format', 'alignment']:
            setattr(cell, attr, copy(getattr(temp_cell, attr)))
        cell.value = Translator(temp_cell.value, temp_cell.coordinate).translate_formula(cell.coordinate)

    def _set_cell(self, row: int, col: int, value: Any) -> None:
        cell = self.worksheet.cell(row, col)
        if not cell.value:
            cell.value = value

    def _fill_template(self, row: int, template: Worksheet) -> None:
        for col, dim in template.column_dimensions.items():
            self.worksheet.column_dimensions[col].width = dim.width
        for r in range(1, template.max_row + 1):
            for c in range(1, template.max_column + 1):
                v = template.cell(r, c)
                cell = self.worksheet.cell(r+(row-1), c)
                if not cell.value:
                    self._copy_cell(cell, v)
 
    def write(self) -> None:
        for path, workbook in self.modified.items():
            workbook.save(path)

    @abstractstaticmethod
    def _date_to_col(self) -> int:
        pass

    @abstractmethod
    def _get_path(self) -> Path:
        pass

    @abstractmethod
    def _paste(self) -> Optional[int]:
        pass
    
    @abstractmethod
    def _fill_column(self) -> None:
        pass

    @abstractmethod
    def fill(self) -> None:
        pass

class InvoiceHandler(SheetHandler):
    def __init__(self, path: Path, templates: Dict[str, Worksheet]) -> None:
        super().__init__(path)
        self.templates = templates

    @staticmethod
    def _date_to_col(date: datetime) -> int:
        week = (date.day - 1) // 7
        col = week + 2
        if date.month != (date + relativedelta(days=1)).month:
            col += 1  # Extra column for the last (incomplete) week
        return col

    def _get_month_row(self, date: datetime) -> Optional[int]:
        month = date.strftime('%B')
        row = self._search(month, 2)
        return row

    def _get_path(self, year: int) -> Path:
        return self.path / f'{year}.xlsx'

    def _paste(self, date: datetime, market: str) -> int:
        row = self.worksheet.max_row + int(self.worksheet.max_row > 1)
        self._fill_template(row, self.templates[market])
        self.worksheet.cell(row, 2).value = date.strftime('%B')
        return row

    def _fill_column(self, head: int, col: int, invoice: Invoice) -> None:
        self._set_cell(head+2, col, invoice.revenue)
        self._set_cell(head+3, col, invoice.fees)
        self._set_cell(head+4, col, invoice.date)

    def fill(self, invoice: Invoice) -> None:   
        for year in self._get_adjacent_years(invoice.date):
            self._set_workbook(year)
            self._set_sheet(invoice.fund)

            if invoice.market == 'miso':
                row = self._get_month_row(invoice.date) or self._paste(invoice.date, invoice.market)
                col = self._date_to_col(invoice.date)

                self._fill_column(row, col, invoice)

                prev_row = self._get_month_row(invoice.date - relativedelta(months=1))
                if col == 2 and prev_row:
                    self._fill_column(prev_row, col, invoice)

class SettlementHandler(SheetHandler):
    def __init__(self, path: Path, templates: Dict[Dict[str, Worksheet]]) -> None:
        super().__init__(path)
        self.templates = templates

    @staticmethod
    def _date_to_col(date: datetime) -> int:
        return date.day + 1

    def _get_path(self, fund: str, doctype: str, year: int) -> Path:
        return self.path / f'{fund} {doctype} {year}.xlsx'

    def _paste(self, market: str, doctype: str):
        self.fill_template(1, self.templates[market][doctype])

    def _fill_column(self, col: int, amounts: Dict[str, float]):
        for name, val in amounts.items():
            self._set_cell(self._search(name), col, val)
    
    def fill(self, settlement: Settlement):
        for year in self._get_adjacent_years(settlement.date)