#!/usr/bin/env python3

import core
from excel import HandlerRotater
from parsers import ParsedXML

args = core.parse_args()
template = core.load_template(args.template, args.parsetype, args.doctype)
rotator = HandlerRotater(args.output, template, args.doctype)

files = [ParsedXML(file, args.doctype) for file in list(args.input.glob('*.xml'))]

for file in files:
    for year in file.years:
        rotator.get_handler(year, file.fund).fill(file)

rotator.write()