import core
from parsers import ParsedXML

args = core.parse_args()
template, paths = core.load_template(args.template)
rotator = core.get_rotator(args.output, args.doctype, template, args.market)

files = [ParsedXML(file, args.doctype, ) for file in list(args.input.glob('*.xml'))]