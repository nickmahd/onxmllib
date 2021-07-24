import pickle

from argparse import ArgumentParser, Namespace
from pathlib import Path
from typing import Callable

from excel import HandlerRotater


def parse_args(parser: ArgumentParser) -> Namespace:
    desc = "Takes a collection of invoices and sends to .xlsx"
    usage = "xl-parse.py [-h] --output=DIR --template=FILE --log=DIR [--force] --market={miso,pjm} --CA=DIR --MKT=DIR\n"

    parser = ArgumentParser(description=desc, usage=usage)

    subparsers = parser.add_subparsers(dest='parsetype', required=True)
    invoice = subparsers.add_parser('invoice')
    settlement = subparsers.add_parser('settlement')

    parser.add_argument('--output', default='.', metavar="DIR", type=Path, help="output directory")
    parser.add_argument('--input', default='.', metavar="DIR", type=Path, help="input directory")
    parser.add_argument('--template', default='template.pkl', metavar="FILE", type=Path, help="template file (preformatted)")
    parser.add_argument('--log', const=".", metavar="DIR", nargs='?', type=Path, help="send a log to directory")

    invoice.add_argument('--doctype', choices=['ca', 'mkt'], required=True, help="type of document")
    invoice.add_argument('--market', choices=['miso', 'pjm'], default='miso', help="market of document")

    settlement.add_argument('--doctype', choices=['summary', 'ftr'], required=True, help="type of document")
    settlement.add_argument('--market', choices=['miso', 'pjm'], default='miso', help="market of document")

    return parser.parse_args()

def load_template(template: Path, doctype: str) -> dict:
    with open(template, 'rb') as file:
        template = pickle.load(file)[doctype]
    
    return (template['template'], template['paths'])

def get_toname(doctype: str, output: Path) -> Callable:
    if doctype == 'summary':
        return lambda year, fund: output / fund + ' Summary ' + str(year) + '.xlsx'
    elif doctype == 'ftr':
        return lambda year, fund: output / fund + ' FTR ' + str(year) + '.xlsx'
    elif doctype in ['ca', 'mkt']:
        return lambda year, fund: output / str(year) + '.xlsx'
    
def get_rotator(output: Path, doctype: str, template: Path, market: str) -> HandlerRotater:
    template = load_template(template, doctype)
    return HandlerRotater(template=template['paths'], handler_type=market, to_name=get_toname(doctype, output))