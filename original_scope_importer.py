import re
import win32com.client
import yaml
import pprint


def pretty_printer(o):
    pp = pprint.PrettyPrinter(indent=4)
    pp.pprint(o)


with open(r'C:\Users\rdapaz\Documents\doctools\original_cable_scope.yaml', 'r') as f:
    data = yaml.load(f)


pretty_printer(data)

xlApp = win32com.client.gencache.EnsureDispatch('Excel.Application')
xlApp.Visible = True
wk = xlApp.Workbooks.Open(r'C:\Users\rdapaz\Desktop\Harvey Beef - Cabling Scope Changes v1.2.xlsx')
sh = wk.Worksheets('Cabling Scope')
eof = sh.Range('B65536').End(-4162).Row

for row in range(4, eof+1):
    _id = sh.Range(f'B{row}').Value
    if _id in data:
        sh.Range(f'D{row}').Value = data[_id]