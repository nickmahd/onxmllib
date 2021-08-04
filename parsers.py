import fnmatch

from abc import ABC, abstractmethod
from datetime import datetime
from pathlib import Path
from typing import List, Tuple
from xml.etree import ElementTree
from zipfile import ZipFile

from dateutil.relativedelta import relativedelta


class Invoice(ABC):
    @abstractmethod
    def __init__(self, market: str) -> None:
        self.market = market

    @staticmethod
    def from_list(files: list, doctype: str, market: str) -> List:
        if doctype in ['mkt', 'ca']:
            return [Invoice(file, doctype, market) for file in files]
        elif doctype in ['ao', 'ftr']:
            return [Settlement(file, doctype, market) for file in files]
        else:
            raise ValueError("Doctype not recognized")

    @abstractmethod
    def unpack(self):
        pass

class MISOInvoice(ParsedXML):
    def __init__(self, MKT_FILE: Path, CA_FILE: Path, market: str) -> None:
        super().__init__(market)

        MKT_ROOT = ElementTree.parse(MKT_FILE).getroot()
        CA_ROOT = ElementTree.parse(CA_FILE).getroot()

        delta = relativedelta(days=7)
        fund = MKT_ROOT.findtext('Header/[Page_Num="1"]/Mrkt_Participant_NmAddr')
        end_date = MKT_ROOT.findtext('Header/[Page_Num="1"]/Billing_Prd_End_Dte')

        mkt_amt = MKT_ROOT.findtext('Header/[Page_Num="1"]/Tot_Net_Chg_Rev_Amt')
        ca_amt = CA_ROOT.findtext('Header/[Page_Num="1"]/Tot_Net_Chg_Rev_Amt')

        self.fund = fund
        self.date = datetime.strptime(end_date, '%m/%d/%Y') - delta
        self.revenue = float(mkt_amt.replace(',', ''))
        self.fees = float(ca_amt.replace(',', ''))

    def unpack(self, root: ElementTree.Element) -> Tuple[str, str, str]:
        fund = root.findtext('Header/[Page_Num="1"]/Mrkt_Participant_NmAddr')
        end_date = root.findtext('Header/[Page_Num="1"]/Billing_Prd_End_Dte')
        amt = root.findtext('Header/[Page_Num="1"]/Tot_Net_Chg_Rev_Amt')
        return (fund, end_date, amt)

class MISOSettlement(ParsedXML):
    def __init__(self, zipfile: Path, market: str) -> None:
        super().__init__(market)

        with ZipFile(zipfile, 'r') as zip_ref:
            AO_FILESTR = fnmatch.filter(zip_ref.namelist(), 'AO-*.xml')[0]
            FTR_FILESTR = fnmatch.filter(zip_ref.namelist(), 'FTR-*S7.xml')[0]
            
            AO_ROOT = ElementTree.fromstring(zip_ref.read(AO_FILESTR))
            FTR_ROOT = ElementTree.fromstring(zip_ref.read(FTR_FILESTR))

        fund = None
        end_date = AO_ROOT.findtext('.//SCHEDULED_DATE')

        ao_names = AO_ROOT.findall('.//CHG_TYP/CHG_TYP_NM')
        ao_amounts = AO_ROOT.findall('.//CHG_TYP/STLMT_TYP[1]/AMT')
        ao_all = AO_ROOT.findall('.//CHG_TYP/STLMT_TYP/AMT')

        ftr_names = FTR_ROOT.findall('.//CHG_TYP/CHG_TYP_NM')
        ftr_amounts = FTR_ROOT.findall('.//CHG_TYP/STLMT_TYP/AMT')

        self.fund = fund
        self.date = datetime.strptime(end_date, '%m/%d/%Y')

        self.ao_amounts = {name.text: float(amount.text) for name, amount in zip(ao_names, ao_amounts)}
        self.ao_amounts["Other Amount"] = sum([float(elem.text) for elem in ao_all]) - sum(self.ao_amounts.values())
        self.ftr_amounts = {name: amount for name, amount in zip(ftr_names, ftr_amounts)}