import fnmatch
import re

from abc import ABC, abstractmethod, abstractstaticmethod
from datetime import datetime
from pathlib import Path
from typing import Iterable, List, TypeVar
from xml.etree import ElementTree
from zipfile import ZipFile

from dateutil.relativedelta import relativedelta

P = TypeVar(name='P', bound='ParsedXML')
class ParsedXML(ABC):
    I = TypeVar(name='I', bound='Invoice')
    S = TypeVar(name='S', bound='Settlement')
    
    @abstractmethod
    def __init__() -> None:
        pass

    @abstractstaticmethod
    def from_list():
        pass

    @abstractstaticmethod
    def from_dir():
        pass


class Invoice(ParsedXML):
    @classmethod
    def from_list(cls, files: List[Iterable]) -> List[ParsedXML.I]:
        file_list = [cls(*group) for group in files]
        return sorted(file_list, key=lambda x: x.date.month)

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
        fund = MKT_ROOT.findtext('.//Mrkt_Participant_NmAddr')
        end_date = MKT_ROOT.findtext('.//Billing_Prd_End_Dte')

        mkt_amt = MKT_ROOT.findtext('.//Tot_Net_Chg_Rev_Amt')
        ca_amt = CA_ROOT.findtext('.//Tot_Net_Chg_Rev_Amt')

        self.fund = fund
        self.date = datetime.strptime(end_date, '%m/%d/%Y') - delta
        self.revenue = float(mkt_amt.replace(',', ''))
        self.fees = float(ca_amt.replace(',', ''))

    @staticmethod
    def from_dir(base: Path, move=False) -> List[ParsedXML.I]:
        PATH = base / 'download'
        MKT_DIR = list((PATH / 'MKT').glob('*'))
        CA_DIR = list((PATH / 'CA').glob('*'))

        mkt_matches = [file for file in MKT_DIR
                       if re.match('.+?_MKT_.+?\.xml$', file.name)]

        file_ids = [re.findall('^(\d{4}-\d{2}-\d{2})_([A-Z]+)_MKT_.+?\.xml$', file.name)[0]
                    for file in mkt_matches]

        ca_matches = [file for match in file_ids 
                      for file in CA_DIR 
                      if re.match(f'{match[0]}_{match[1]}_CA_.+?\.xml$', file.name)]

        files = MISOInvoice.from_list(list(zip(mkt_matches, ca_matches)))

        if move:
            for file in mkt_matches + ca_matches:
                file_new = Path(*('processed' if part == 'download' else part
                                  for part in file.parts))
                file_new.parent.mkdir(parents=True, exist_ok=True)
                file.rename(file_new)

                pdf = Path(file.parent / f'{file.stem}.pdf')
                pdf_new = Path(*('processed' if part == 'download' else part
                                 for part in pdf.parts))
                pdf.rename(pdf_new)

        return files

class PJMInvoice(Invoice):
    def __init__(self, *group) -> None:
        ROOTS = [ElementTree.parse(file).getroot() for file in group]

        fund = ROOTS[0].findtext('.//CUSTOMER_ACCOUNT')
        end_date = ROOTS[0].findtext('.//BILLING_PERIOD_END_DATE')

        names = [re.match('(.+?)_', file.name).group(1) for file in group]
        amts = [root.findtext('.//TOTAL_DUE_RECEIVABLE') for root in ROOTS]

        self.fund = fund #re.match('(.+?(?=\s\())', fund).group(1)
        self.names = names
        self.date = datetime.strptime(end_date, '%Y-%m-%d') if end_date else None
        self.amts = [float(i) if i else None for i in amts]

    @staticmethod
    def _sort_subgroup(x) -> int:
        key = re.match('(.+?)_', x.name).group(1)[-1]
        return int(key) if key.isnumeric() else 0

    @staticmethod
    def from_dir(base: Path, move=False) -> None:
        PATH = base / 'download'
        DIR = list(PATH.glob('*.xml'))

        file_ids = set([re.findall('^(.{3}).+?_(\d{6}_\d{6}).+?\.xml', file.name)[0]
                        for file in DIR])

        groups = [[file for file in DIR if re.match(f'{match[0]}.+?_{match[1]}', file.name)]
                   for match in file_ids]

        sorted_groups = [sorted(l, key=PJMInvoice._sort_subgroup) for l in groups]
        files = PJMInvoice.from_list(sorted_groups)

        if move:
            PROCESSED_DIR = (base / 'processed')
            PROCESSED_DIR.mkdir(parents=True, exist_ok=True)
            for file in DIR:
                pdf_name = re.match('(.+?WEKBILL)CSV', file.name).group(1)
                pdf = next(PATH.glob(f'{pdf_name}PDF*.pdf'))
                
                pdf.rename(PROCESSED_DIR / pdf.name)
                file.rename(PROCESSED_DIR / file.name)

        return files


class Settlement(ParsedXML):
    @classmethod
    def from_list(cls, files: List) -> List[ParsedXML.S]:
        return [cls(file) for file in files]

    @staticmethod
    def from_dir(base: Path, market: str, move=False) -> List[ParsedXML.S]:
        if market == 'miso':
            return MISOSettlement.from_dir(base, move=move)
        else:
            raise NotImplementedError(f"Settlement parser for market '{market}' not implemented")

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

        print([x for x in ao_amounts if not isinstance(x, ElementTree.Element)])

        self.amounts = {}
        self.amounts['ao'] = {name.text: float(amount.text) for name, amount in zip(ao_names, ao_amounts)}
        self.amounts['ao']["Other Amount"] = sum([float(elem.text) for elem in ao_all]) - sum(self.amounts['ao'].values())
        self.amounts['ftr'] = {name.text: float(amount.text) for name, amount in zip(ftr_names, ftr_amounts)}

    @staticmethod
    def from_dir(base: Path, move=False) -> List[ParsedXML.S]:
        DIR = list((base / 'downloads').rglob('*.zip'))
        files = MISOSettlement.from_list(DIR)

        if move:
            for file in DIR:
                file_new = Path(*('processed' if part == 'downloads' else part
                                  for part in file.parts))
                file_new.parent.mkdir(parents=True, exist_ok=True)
                file.rename(file_new)

        return files