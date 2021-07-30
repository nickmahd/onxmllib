#!/usr/bin/env python3

from core import HandlerRotater, ParsedXML
from core import parse_args

args = parse_args()

rotator = HandlerRotater(args.output, args.template, args.doctype)
files = ParsedXML.from_list(list(args.input.glob('*.xml')), args.doctype)

for file in files:
    for year in file.years:
        rotator.get_handler(year, file.fund).fill(file)

rotator.write()