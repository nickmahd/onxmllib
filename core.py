import pickle

from argparse import ArgumentParser, Namespace
from pathlib import Path
from typing import Any, Callable

from openpyxl.worksheet.worksheet import Worksheet


def parse_args() -> Namespace:
    desc = "Takes a collection of invoices and sends to .xlsx"
    usage = "xl-parse.py [-h] <invoice|settlement> --output=DIR --input=DIR --template=FILE --doctype=<{ca,mkt}|{summary,ftr}> \n"

    parser = ArgumentParser(description=desc, usage=usage)

    subparsers = parser.add_subparsers(dest='parsetype', metavar="one of <invoice> or <settlement>", required=True)
    invoice = subparsers.add_parser('invoice')
    settlement = subparsers.add_parser('settlement')

    parser.add_argument('--output', default='.', metavar="DIR", type=Path, help="output directory")
    parser.add_argument('--input', default='.', metavar="DIR", type=Path, help="input directory")
    parser.add_argument('--template', default='template.pkl', metavar="FILE", type=Path, help="template file (preformatted)")
    # parser.add_argument('--log', const=".", metavar="DIR", nargs='?', type=Path, help="send a log to directory")
    # parser.add_argument('--market', choices=['miso', 'pjm'], default='miso', help="market of document")

    invoice.add_argument('--doctype', choices=['ca', 'mkt'], required=True, help="type of document")
    settlement.add_argument('--doctype', choices=['summary', 'ftr'], required=True, help="type of document")

    return parser.parse_args()

def reduce(doctype: str, summary: Any = 'summary', ftr: Any = 'ftr', invoice: Any = 'invoice') -> Any:
    if doctype == 'summary':
        return summary
    elif doctype == 'ftr':
        return ftr
    elif doctype in ['ca', 'mkt']:
        return invoice
    else:
        raise ValueError("Doctype not recognized")

def load_template(template: Path, doctype: str) -> dict:
    key = reduce(doctype)
    with open(template, 'rb') as file:
        template = pickle.load(file)[key]
    return template

def get_toname(doctype: str, output: Path) -> Callable[[int, str], Path]:
    summary = lambda year, fund: output / fund + ' Summary ' + str(year) + '.xlsx'
    ftr = lambda year, fund: output / fund + ' FTR ' + str(year) + '.xlsx'
    invoice = lambda year, _: output / str(year) + '.xlsx'
    return reduce(doctype, summary, ftr, invoice)