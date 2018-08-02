import re
import win32com.client
import datetime
import os


ROOT_FOLDER = r'C:\Users\rdapaz\Desktop\Harvey Beef\Schedule'

packages = [
    'Harvey LAN - Work Package 1',
    'Harvey Firewall/WAN - Work Package 2',
    'Harvey Servers - Work Package 3',
    'Harvey WLAN - Work Package 4 ',
    'Harvey DR - Work Package 5',
    'Fremantle - Work Package 6',
    'Minderoo - Work Package 7',
    'Swires - Work Package 8',
    'South West Express - Work Package 9 ',
]


path = os.path.join(ROOT_FOLDER, 'Minderoo - Project Schedule 20180713.mpp')
app = win32com.client.gencache.EnsureDispatch('MSPRoject.Application')
app.FileOpen(path)
pj = app.ActiveProject


rex = re.compile(r'(Build|Test|Deploy)', re.IGNORECASE)

arr = []
for tsk in pj.Tasks:
    if tsk.Name in packages:
        parent_task = tsk.Name
        for subtsk in tsk.OutlineChildren:
            m = rex.search(subtsk.Name)
            if m:
                what = m.group(1)
                startDate = f'{subtsk.Start}'[:10]
                startDate = datetime.datetime.strptime(startDate, '%Y-%m-%d').strftime('%d/%m/%Y')
                endDate = f'{subtsk.Finish}'[:10]
                endDatedt = datetime.datetime.strptime(endDate, '%Y-%m-%d')
                endDate = endDatedt.strftime('%d/%m/%Y')
                arr.append([parent_task, subtsk.Name, startDate, endDate, endDatedt])


# arr = sorted(arr, key=lambda x: x[-1])
new_arr = [[x[0], x[1], x[2], x[3]] for x in arr]
for x,y,z,w in new_arr:
    print(x,y,z,w, sep="|")