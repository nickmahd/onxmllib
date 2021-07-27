from dateutil.relativedelta import relativedelta
from datetime import datetime
from pathlib import Path
from xml.etree import ElementTree

from core import _reduce

PATHS = {
    'invoice': {
        'net_rev': 'Header/[Page_Num="1"]/Tot_Net_Chg_Rev_Amt',
        'end_date': 'Header/[Page_Num="1"]/Billing_Prd_End_Dte',
        'fund': 'Header/[Page_Num="1"]/Mrkt_Participant_NmAddr',
        'delta': 7
    },
    'summary': {

    },
    'ftr': {

    }
}

class ParsedXML:
    def __init__(self, file: Path, doctype: str):
        self.file = file
        paths = PATHS[_reduce(doctype)]

        self.root = ElementTree.parse(self.file).getroot()

        self.net_rev = float(self.root.findtext(paths['net_rev']).replace(',', ''))
        self.fund = self.root.findtext(paths['fund'])
        self._end_date = self.root.findtext(paths['end_date'])
        self.date = datetime.strptime(self._end_date, '%m/%d/%Y') - relativedelta(days=paths['delta'])
        
        self.years = list(set((self.date + relativedelta(months=i)).year for i in range(-1, 2)))