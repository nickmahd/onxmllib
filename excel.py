#!/usr/bin/env python3

import logging

from copy import copy
from datetime import datetime
from pathlib import Path

from dateutil.relativedelta import relativedelta
from openpyxl import Workbook, load_workbook
from openpyxl.formula.translate import Translator
from openpyxl.worksheet.worksheet import Worksheet

from parsers import ParsedXML


class SheetHandler:
    def __new__(cls, doctype, *args, **kwargs):
        if doctype == "miso":
            return MisoHandler(*args, **kwargs)
        elif doctype == "summary":
            return Summary(*args, **kwargs)
        elif doctype == "ftr":
            return FTR(*args, **kwargs)
        else:
            raise ModuleNotFoundError(f"{doctype} is not a class or has no reference to one")

    def __init__(self, path: Path, workbook: Workbook, template: Worksheet, forceful=False):
        self.path = path
        self._row_at = 1
        self.workbook = workbook
        self.worksheet: Worksheet = self.workbook.active
        self.template = template
        self.forceful = forceful

    def _copy_cell(cell, temp_cell):
        for attr in ['font', 'border', 'fill', 'number_format', 'alignment']:
            setattr(cell, attr, copy(getattr(temp_cell, attr)))
        cell.value = Translator(temp_cell.value, temp_cell.coordinate).translate_formula(cell.coordinate)
    
    def _search(self, kwd: str, col: int):
        for row in range(1, self.worksheet.max_row):
            if self.worksheet.cell(row=row, column=col).value == kwd:
                return row
        return None

    def _paste(self, row: int):
        for r in range(1, self.template.max_row + 1):
            for c in range(1, self.template.max_column + 1):
                v = self.template.cell(row=r, column=c)
                cell = self.worksheet.cell(row=r+row-1, column=c)

                if self.forceful or not cell.value:
                    self._copy_cell(cell, v)

    def _set_cell(self, row, col, value):
        cell = self.worksheet.cell(row=row, column=col)
        if self.forceful or not cell.value:
            cell.value = value
    
    def set_sheet(self, sheet_name: str):
        if not (self.worksheet and self.worksheet.title == sheet_name):
            try:
                self.worksheet = self.workbook[sheet_name]
            except KeyError:
                self.worksheet = self.workbook.create_sheet(sheet_name)

            for col, dim in self.template.column_dimensions.items():
                self.worksheet.column_dimensions[col].width = dim.width
 
    def write(self):
        self.workbook.save(self.path)

class MisoHandler(SheetHandler):
    def _fill_column(self, head: int, col: int, date: datetime, revenue: float):
        self._set_cell(head+2, col, revenue)
        self._set_cell(head+4, col, date)

    def fill(self, file: ParsedXML):
        month = file.date.strftime('%B')
        last_month = (file.date - relativedelta(months=1)).strftime('%B')

        row = self._search(month, 2) 

        if row:
            self._row_at = row
        else:
            self._row_at = self.worksheet.max_row + 1
            self._paste(self._row_at)
            self.worksheet.cell(row=self._row_at, column=2).value = month

        week = (file.date.day - 1) // 7
        col = week + 2

        if file.date.month != (file.date + relativedelta(days=1)).month:
            col += 1

        self._fill_column(self._row_at, col, file.date, file.net_rev)

        if col == 2 and self._search_month(last_month):
            self._set_month(last_month)
            self._fill_column(self._row_at, 8, file.date, file.net_rev)

class Summary(SheetHandler):
    def trade_results(self, day):
        ahead_energy_amt = self.worksheet.cell(row=3, column=day+1)
        real_energy_amt = self.worksheet.cell(row=4, column=day+1)
        return -sum(filter(None, [ahead_energy_amt, real_energy_amt]))
    
    def total(self, day):
        cols = [self.worksheet.cell(row=r, column=day+1).value for r in range(5, self.template.max_column + 1)]
        return self.trade_results - sum(filter(None, cols))

class FTR(SheetHandler):
    @property
    def total(self, day):
        cols = [self.worksheet.cell(row=r, column=day+1).value for r in range(3, self.template.max_column + 1)]
        return sum(filter(None, cols))


class HandlerRotater:
    def __init__(self, template: Worksheet, handler_type: str, to_name):
        self.handlers = {}
        self.template = template
        self.type = handler_type
        self.to_name = to_name

    @classmethod
    def get_workbook(cls, name: Path):
        try:
            wb = load_workbook(name)
            return wb
        except FileNotFoundError:
            wb = Workbook()
            wb.remove(wb.active)
            return wb
    
    def get_handler(self, input):
        name = self.to_name(input)
        if input not in self.handlers:
            self.handlers[input] = SheetHandler(self.type, path=name, workbook=self.get_workbook(name), template=self.template)
        return self.handlers[input]