#!/usr/bin/env python3

from excel import SheetHandler
from utils import parse_args

args = parse_args()

handler = SheetHandler.get_handler(parsetype=args.parsetype,
                                   path=args.output,
                                   template=args.template)

handler.process_dir(input=args.input,
                    market=args.market,
                    move=args.dryrun)

handler.write()