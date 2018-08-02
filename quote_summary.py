import re
import win32com.client
import datetime
from decimal import Decimal

file = r'C:\Users\rdapaz\Desktop\Harvey Beef\Budget\Harvey Beef - Infrastructure Refresh Financial Tracker V1.2.xlsx'
xlApp = win32com.client.gencache.EnsureDispatch('Excel.Application')
xlApp.Visible = True
wk = xlApp.Workbooks.Open(file)


targetSheet = re.compile(r'[A-Z][A-Z][A-Z]?\d{6}v\d', re.IGNORECASE)

arr = []
for sht in wk.Worksheets:
    if targetSheet.search(sht.Name):
        supplier = 'Datacom'
        dt = sht.Range('D3').Value
        dt = f'{dt}'[:10]
        dt = datetime.datetime.strptime(dt, '%Y-%m-%d').strftime('%d/%m/%Y')
        quote_ref = sht.Range('D2').Value.strip()
        eof = sht.Range('G65536').End(-4162).Row
        total = sht.Range(f'G{eof}').Value
        total = Decimal(total)/Decimal(1.1)
        print(dt, supplier, total, '', quote_ref, total, sep="|")
