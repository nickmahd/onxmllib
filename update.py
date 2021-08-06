#!/usr/bin/env python3

from pickle import dump
from openpyxl import load_workbook

t = {
    'invoice': {
        'miso': load_workbook('invoice.xlsx')['miso'],
        'pjm': load_workbook('invoice.xlsx')['pjm']},
    'settlement': {
        'summary': load_workbook('ao.xlsx').active,
        'ftr': load_workbook('FTR_template.xlsx').active
    }
}
with open('template.pkl', 'wb') as f: dump(t, f)