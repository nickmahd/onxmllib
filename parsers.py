import fnmatch
import re

from abc import ABC, abstractmethod, abstractstaticmethod
from datetime import datetime
from pathlib import Path
from typing import List, Tuple, TypeVar
from xml.etree import ElementTree
from zipfile import ZipFile

from dateutil.relativedelta import relativedelta

P = TypeVar(name='P', bound='ParsedXML')
class ParsedXML(ABC):
    I = TypeVar(name='I', bound='Invoice')
    S = TypeVar(name='S', bound='Settlement')
    
    @abstractmethod
    def __init__(self) -> None:
        pass

    @abstractstaticmethod
    def _from_list():
        pass

    @abstractstaticmethod
    def from_dir():
        pass


class Invoice(ParsedXML):
    @staticmethod
    def from_dir(base: Path, market: str, move=False) -> List[ParsedXML.I]:
        if market == 'miso':
            return MISOInvoice.from_dir(base, move=move)
        elif market == 'pjm':
            return PJMInvoice.from_dir(base, move=move)
        else:
            raise NotImplementedError(f"Invoice parser for market '{market}' not implemented")

class MISOInvoice(Invoice):
    def __init__(self, MKT_FILE: Path, CA_FILE: Path) -> None:
        MKT_ROOT = ElementTree.parse(MKT_FILE).getroot()
        CA_ROOT = ElementTree.parse(CA_FILE).getroot()

        delta = relativedelta(days=7)
        fund = MKT_ROOT.findtext('../Mrkt_Participant_NmAddr')
        end_date = MKT_ROOT.findtext('../Billing_Prd_End_Dte')

        mkt_amt = MKT_ROOT.findtext('../Tot_Net_Chg_Rev_Amt')
        ca_amt = CA_ROOT.findtext('../Tot_Net_Chg_Rev_Amt')

        self.fund = fund
        self.date = datetime.strptime(end_date, '%m/%d/%Y') - delta
        self.revenue = float(mkt_amt.replace(',', ''))
        self.fees = float(ca_amt.replace(',', ''))

    @staticmethod
    def _from_list(files: List[tuple]) -> List:
        return [MISOInvoice(*pair) for pair in files]

    @staticmethod
    def from_dir(base: Path, move=False) -> None:
        MKT_DIR = list((base / 'MKT').glob('*'))
        CA_DIR = list((base / 'CA').glob('*'))

        mkt_matches = [file for file in MKT_DIR if re.match('^.+?_MKT_.+?\.xml$', file.name)]
        file_ids = [re.findall('^.+?_(.+?)_MKT_(.+?)\.xml$', file.name)[0] for file in mkt_matches]
        ca_matches = [file for fund, id in file_ids for file in CA_DIR if re.match(f'^.+?_{fund}_CA_{int(id)+1}\.xml$', file.name)]
        files = MISOInvoice._from_list(list(zip(mkt_matches, ca_matches)))

        if move:
            PROCESSED_DIR = (base / 'processed')
            PROCESSED_DIR.mkdir(exist_ok=True)
            for file in mkt_matches + ca_matches:
                pdf = Path(file.parent / f'{file.stem}.pdf')
                pdf.rename(PROCESSED_DIR / pdf.name)
                file.rename(PROCESSED_DIR / file.name)

        return files

class PJMInvoice(Invoice):
    def __init__(self, *args) -> None:
        ROOT = ElementTree.parse(MKT_FILE).getroot()

        delta = relativedelta(days=7)
        fund = ROOT.findtext('.//CUSTOMER_ACCOUNT')
        end_date = ROOT.findtext('.//BILLING_PERIOD_END_DATE')

        amt = ROOT.findtext('.//TOTAL_DUE_RECEIVABLE')

        self.fund = fund
        self.date = datetime.strptime(end_date, '%m/%d/%Y') - delta
        self.revenue = float(amt.replace(',', ''))

    @staticmethod
    def _from_list(files: List[tuple]) -> List:
        return [PJMInvoice(*group) for group in files]

    @staticmethod
    def from_dir(base: Path, move=False) -> None:
        DIR = list((base / 'downloads').rglob('*.xml'))

        matches = [file for file in DIR if re.match('')]
        file_ids = set([re.findall()])



class Settlement(ParsedXML):
    @staticmethod
    def _from_list(files: List, market: str) -> List[ParsedXML.S]:
        if market == 'miso':
            return [MISOSettlement(file) for file in files]
        else:
            raise NotImplementedError(f"Settlement parser for market '{market}' not implemented")

    @staticmethod
    def from_dir(base: Path, market: str, move=False) -> List[ParsedXML.I]:
        DIR = list((base / 'downloads').rglob('*.zip'))
        files = Settlement._from_list(DIR, market)

        if move:
            PROCESSED_DIR = (base / 'processed')
            PROCESSED_DIR.mkdir(exist_ok=True)
            for file in DIR:
                file.rename(PROCESSED_DIR / file.name)

        return files

class MISOSettlement(Settlement):
    def __init__(self, zipfile: Path) -> None:
        with ZipFile(zipfile, 'r') as zip_ref:
            AO_FILESTR = fnmatch.filter(zip_ref.namelist(), 'AO-*.xml')[0]
            FTR_FILESTR = fnmatch.filter(zip_ref.namelist(), 'FTR_*S7.xml')[0]

            AO_ROOT = ElementTree.fromstring(zip_ref.read(AO_FILESTR))
            FTR_ROOT = ElementTree.fromstring(zip_ref.read(FTR_FILESTR))

        fund = AO_ROOT.findtext('.//ID')
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
        self.ftr_amounts = {name.text: float(amount.text) for name, amount in zip(ftr_names, ftr_amounts)}
