import re
import win32com.client
import pprint


def pretty_print(o):
    pp = pprint.PrettyPrinter(indent=4)
    pp.pprint(o)


def orderPrefix(aEnd, bEnd):
    dummy = [aEnd, bEnd]
    dummy = sorted(dummy)
    return "-".join(dummy)

class Excel:

    def __init__(self, path):
        self.path = path
        self.app = win32com.client.gencache.EnsureDispatch('Excel.Application')
        self.wk = self.app.Workbooks.Open(self.path)
        self.sh = self.wk.Worksheets('Cable Schedule')
        data = """
        Primary Server Room|SR1
        Cabinet 1 - Stores|STO
        Cabinet 2 - Maintenance|MTC
        Cabinet 3 - Engineering|ENG
        Cabinet 4 - Admin|ADM
        Cabinet 5 - Hay Shed MCC|HAY
        Cabinet 6 - By-product MCC|BYP
        Cabinet 7 - Processing (In roof space)|PRO
        Cabinet 8 - QA Building|QAB
        Cabinet 9 - Engine room|ERM
        Cabinet 10 - Cold Stores|CLD
        Cabinet 11 - Retail Area MCC|SR2
        CCTV|VID
        Data Point|DAT
        Edge Device|OTH
        """.splitlines()
        data = [x.strip().split('|') for x in data if len(x.strip()) > 0]
        self.data = {k: v for k, v in [[x, y ]for x, y in data]}
        pretty_print(self.data)
        self.cableNos = {}

    
    @property
    def getlastRow(self):
        return self.sh.Range('C65536').End(-4162).Row

    
    def updateCableIDs(self):
        for row in range(5, self.getlastRow+1):
            cableAEnd = self.sh.Range(f'C{row}').Value
            cableBEnd = self.sh.Range(f'E{row}').Value
            desc = self.sh.Range(f'D{row}').Value
            designator = ''
            if re.search(r'OS1', desc):
                designator = 'FOC'
            elif re.search(r'UTP', desc):
                designator = 'UTP'
            if all([cableAEnd, cableBEnd]):
                cableAMarker = self.data[cableAEnd]
                cableBMarker = self.data[cableBEnd]
                if not any([cableAMarker in ['VID', 'DAT', 'OTH'], 
                        cableBMarker in ['VID', 'DAT', 'OTH']]):
                    cable_id_prefix = f'{designator}-{orderPrefix(cableAMarker, cableBMarker)}' 
                else:
                    cable_id_prefix = f'{designator}-{cableAMarker}-{cableBMarker}'
                if cable_id_prefix not in self.cableNos:
                    self.cableNos[cable_id_prefix] = 0
                self.cableNos[cable_id_prefix] +=1
                cable_id = f'{cable_id_prefix}-{str(self.cableNos[cable_id_prefix]).zfill(3)}'
                self.sh.Range(f'B{row}').Value = cable_id


xl = Excel(r'C:\Users\rdapaz\Desktop\Cabling Scope Changes.xlsx')
xl.updateCableIDs()


