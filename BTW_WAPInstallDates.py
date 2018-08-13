import win32com.client
import datetime
import re
import yaml

file = r'C:\Users\rdapaz\Desktop\Finalised BTS WAP Installation Schedule Revised.xlsx'

xlApp = win32com.client.gencache.EnsureDispatch('Excel.Application')
xlApp.Visible = True

wk = xlApp.Workbooks.Open(file)
sh = wk.Worksheets('Sched')


FIRST_COL = 2
LAST_COL = 15


def list_to_english(lst):
    print(len(lst))
    if len(lst) == 0:
        return ''
    elif len(lst) == 1:
        return lst[0]
    elif len(lst) == 2:
        return f'{lst[0]} and {lst[1]}'
    else:
        return ", ".join(lst[:-1]) + ' and ' + lst[-1] 

data = []
for col in range(FIRST_COL, LAST_COL+1):
    dt = f'{sh.Cells(4, col).Value}'[:10]
    dt = datetime.datetime.strptime(dt, '%Y-%m-%d').strftime('%d/%m/%Y')
    eof = sh.Cells(14, col).End(-4162).Row
    waps = []
    for row in range(7, eof+1):
        wap = sh.Cells(row, col).Value
        regx = re.compile(r'([WAP 0-9]+)')
        m = regx.search(wap)
        if m:
            wap = m.group(1)
            wap = wap.replace(' ', '')
        waps.append(wap)
    if len(waps) > 0:
        data.append([list_to_english(sorted(waps)), dt, dt])

for entry in data:
    print('|'.join(entry))
