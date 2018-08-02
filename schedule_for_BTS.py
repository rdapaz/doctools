import win32com.client


data = '''
WAP 03|Day 1
WAP 05|Day 2
WAP 13|Day 6
WAP 17|Day 5
WAP 23|Day 4
WAP 24|Day 4
WAP 01|Day 1
WAP 02|Day 1
WAP 04|Day 2
WAP 08|Day 2
WAP 10|Day 6
WAP 11|Day 3
WAP 12|Day 5
WAP 18|Day 5
WAP 20|Day 5
WAP 06|Day 7
WAP 07|Day 7
WAP 09|Day 6
WAP 14|Day 6
WAP 15|Day 7
WAP 19|Day 10
WAP 21|Day 10
WAP 22|Day 10
'''.splitlines()

data = [x.split('|') for x in data if len(x) > 0]

sched = {}
for wap, day in data:
    if day not in sched:
        sched[day] = []
    sched[day].append(wap)

xlApp = win32com.client.gencache.EnsureDispatch('Excel.Application')
xlApp.Visible = True
wk = xlApp.Workbooks.Open(r'C:\Users\rdapaz\Desktop\Harvey Beef\BTS_WAP_Installs.xlsx')
sh = wk.Worksheets('Sched')

print(sh.Range('B2').Value)
for col in range(2, 16):
    dd = sh.Cells(6, col).Value
    if dd and dd in sched:
        waps = sorted(sched[dd])
        for row in range(len(waps)):
            sh.Cells(row + 7,col).Value = waps[row]


