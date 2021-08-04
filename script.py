#!/usr/bin/env python3

from excel import SheetHandler, parse_args
from parsers import ParsedXML

args = parse_args()

rotator = SheetHandler.get_handler(args.parsetype, args.output, args.template)
files = ParsedXML.from_list(list(args.input.glob('*.xml')), args.doctype)

for file in files:
    for year in file.years:
        rotator.get_handler(year, file.fund).fill(file)

rotator.write()