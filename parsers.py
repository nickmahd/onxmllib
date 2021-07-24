from dateutil.relativedelta import relativedelta
from datetime import datetime
from pathlib import Path
from xml.etree import ElementTree

class ParsedXML:
    def __init__(self, file: Path, paths: dict):
        self.file = file

        self.root = ElementTree.parse(self.file).getroot()

        self.net_rev = float(self.root.findtext(paths['net_rev']).replace(',', ''))
        self.fund = self.root.findtext(paths['fund'])
        self._end_date = self.root.findtext(paths['end_date'])
        self.date = datetime.strptime(self._end_date, '%m/%d/%Y') - relativedelta(days=paths['delta'])
        
        self.years = list(set((self.date + relativedelta(months=i)).year for i in range(-1, 2)))