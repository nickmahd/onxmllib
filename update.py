#!/usr/bin/env python3

from pickle import dump
from openpyxl import load_workbook

t = {'invoice': load_workbook('sheetinvoice.xlsx').active, 'ao': load_workbook('sheetao.xlsx').active} #, 'ftr': load_workbook('sheetftr.xlsx').active}
with open('template.pkl', 'wb') as f: dump(t, f)