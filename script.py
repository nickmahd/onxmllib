#!/usr/bin/env python3

from excel import SheetHandler
from utils import parse_args

args = parse_args()

handler = SheetHandler.get_handler(args.parsetype, args.output, args.template)
handler.process_dir(args.input, args.market)
handler.write()