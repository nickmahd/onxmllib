import smtplib
from argparse import ArgumentParser, Namespace
from email.message import EmailMessage
from pathlib import Path

SCRIPT_DESC = "Takes a collection of invoices and sends to .xlsx"
SCRIPT_USAGE = "xl-parse.py [-h] <invoice|settlement> --output=DIR --input=DIR --template=FILE --doctype=<{ca,mkt}|{ao,ftr}> \n"

TO_EMAIL = 'venkat@oncept.net'
FROM_EMAIL = 'Me_watch@oncept.net'
SMTP_SERVER = '10.10.0.22'

def parse_args() -> Namespace:
    parser = ArgumentParser(description=SCRIPT_DESC, usage=SCRIPT_USAGE)

    subparsers = parser.add_subparsers(dest='parsetype', metavar="one of <invoice> or <settlement>", required=True)
    invoice = subparsers.add_parser('invoice')
    settlement = subparsers.add_parser('settlement')

    parser.add_argument('--output', default='.', metavar="DIR", type=Path, help="output directory")
    parser.add_argument('--input', default='.', metavar="DIR", type=Path, help="input directory")
    parser.add_argument('--template', default='template.pkl', metavar="FILE", type=Path, help="template file (preformatted)")
    # parser.add_argument('--log', const=".", metavar="DIR", nargs='?', type=Path, help="send a log to directory")
    # parser.add_argument('--market', choices=['miso', 'pjm'], default='miso', help="market of document")

    invoice.add_argument('--doctype', choices=['ca', 'mkt'], required=True, help="type of document")
    settlement.add_argument('--doctype', choices=['ao', 'ftr'], required=True, help="type of document")

    return parser.parse_args()

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