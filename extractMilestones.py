import re
import win32com.client
import datetime
import os


ROOT_FOLDER = r'C:\Users\rdapaz\Desktop\Belmont DC Exit - Implementation'

packages = {
    'Package 1': 'Telecommunications Services Changes',
    'Package 2': 'Telephony Service Relocation',
    'Package 3': 'Relocation to Malaga DC',
    'Package 4': 'Base AWS Production Build',
    'Package 5': 'Data Protection',
    'Package 6': 'Dev & Test Build',
    'Package 7': 'Disaster Recovery',
    'Package 8': 'Equipment Destined for Site Decommissioning)',
    'Package 9': 'Equipment Destined for Decommissioning & Disposal',
    'Package 10': 'Handover Documentation' 
}


path = os.path.join(ROOT_FOLDER, 'Belmont DC Relocation Project.mpp')
app = win32com.client.gencache.EnsureDispatch('MSPRoject.Application')
app.FileOpen(path)
pj = app.ActiveProject

rex = re.compile(r'Completion of (Package \d)', re.IGNORECASE)

arr = []
for tsk in pj.Tasks:
    startDate, endDate = None, None
    m = rex.search(tsk.Name)
    if m:
        package = m.group(1)
        package_name = packages[package]
        print(package)
        endDate = f'{tsk.Finish}'[:10]
        endDate = datetime.datetime.strptime(endDate, '%Y-%m-%d').strftime('%d/%m/%Y')
        arr.append([tsk.WBS, f"Completion of {package} ({package_name})", endDate, tsk.Finish])
    elif tsk.Duration == 0:
        endDate = f'{tsk.Finish}'[:10]
        endDate = datetime.datetime.strptime(endDate, '%Y-%m-%d').strftime('%d/%m/%Y')
        arr.append([tsk.WBS, tsk.Name, endDate, tsk.Finish])


arr = sorted(arr, key=lambda x: x[-1])
new_arr = [[x[0], x[1], x[2]] for x in arr]
for x,y,z in new_arr:
    print(x,y,z, sep="|")