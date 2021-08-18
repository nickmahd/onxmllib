import smtplib
from argparse import ArgumentParser, Namespace
from email.message import EmailMessage
from pathlib import Path
from typing import Union

from excel import SheetHandler

SCRIPT_DESC = "Takes a collection of invoices and sends to .xlsx"
SCRIPT_USAGE = "script.py [-h] <invoice|settlement> [--output DIR] [--input DIR] [--template FILE] [--market {miso,pjm}] [--dryrun]"

TO_EMAIL = 'venkat@oncept.net'
FROM_EMAIL = 'Me_watch@oncept.net'
SMTP_SERVER = '10.10.0.22'

def parse_args() -> Namespace:
    parser = ArgumentParser(description=SCRIPT_DESC, usage=SCRIPT_USAGE)

    parser.add_argument('parsetype', choices=['invoice', 'settlement'], help='parse type')
    parser.add_argument('--output', default='.', metavar="DIR", type=Path, help="output directory")
    parser.add_argument('--input', default='.', metavar="DIR", type=Path, help="input directory")
    parser.add_argument('--template', default='template.pkl', metavar="FILE", type=Path, help="template file (preformatted)")
    parser.add_argument('--market', choices=['miso', 'pjm'], help="market of document")
    parser.add_argument('--dryrun', action='store_false', help="whether to move files")
    # parser.add_argument('--log', const=".", metavar="DIR", nargs='?', type=Path, help="send a log to directory")

    return parser.parse_args()

def run(parsetype: str, output: Union[str, Path], input: Union[str, Path],
        template: Union[str, Path], market: str, move: Union[str, bool]):
    output = Path(output)
    input = Path(input)
    template = Path(template)
    dryrun = bool(move)

    handler = SheetHandler.get_handler(parsetype=parsetype,
                                       path=output,
                                       template=template)

    handler.process_dir(input=input,
                    market=market,
                    move=dryrun)

    handler.write()

def from_cli(args: Namespace = None):
    args = args or parse_args()
    run(parsetype=args.parsetype, output=args.output, input=args.input,
        template=args.template, market=args.market, move=args.dryrun)

def email(message, to_email: str=TO_EMAIL, subject: str="<no subject>",
          from_email: str=FROM_EMAIL, server: str=SMTP_SERVER) -> None:
    msg = EmailMessage()
    msg['To'] = ', '.join(to_email)
    msg['From'] = from_email
    msg['Subject'] = subject
    msg.set_content(message)

    server = smtplib.SMTP(server)
    server.send_message(msg)
    server.quit()