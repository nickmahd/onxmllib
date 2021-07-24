from dateutil.relativedelta import relativedelta
from datetime import datetime
from pathlib import Path
from xml.etree import ElementTree

PATHS = {
    'invoice': {
        'miso': {
            'net_rev': 'Header/[Page_Num="1"]/Tot_Net_Chg_Rev_Amt',
            'end_date': 'Header/[Page_Num="1"]/Billing_Prd_End_Dte',
            'fund': 'Header/[Page_Num="1"]/Mrkt_Participant_NmAddr',
            'delta': 7
        },
        'pjm': {
            'net_rev': 'ROWSET/ROW/TOTAL_DUE_RECEIVABLE',
            'end_date': 'HEADER/BILLING_PERIOD_END_DATE',
            'fund': 'HEADER/CUSTOMER_ACCOUNT',
            'delta': 0
        }
    },
    'settlement': {
        'summary': {

        }
    }
}

class ParsedXML:
    def __init__(self, file: Path, parsetype: str, doctype: str):
        self.file = file
        paths = PATHS[parsetype][doctype]

        self.root = ElementTree.parse(self.file).getroot()

        self.net_rev = float(self.root.findtext(paths['net_rev']).replace(',', ''))
        self.fund = self.root.findtext(paths['fund'])
        self._end_date = self.root.findtext(paths['end_date'])
        self.date = datetime.strptime(self._end_date, '%m/%d/%Y') - relativedelta(days=paths['delta'])
        
        self.years = list(set((self.date + relativedelta(months=i)).year for i in range(-1, 2)))