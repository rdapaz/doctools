import re
import win32com.client
import datetime
import os


ROOT_FOLDER = r'C:\Users\rdapaz\Downloads'


path = os.path.join(ROOT_FOLDER, 'Infrastructure Gantt (1).mpp')
app = win32com.client.gencache.EnsureDispatch('MSPRoject.Application')
app.FileOpen(path)
pj = app.ActiveProject

arr = []
for tsk in pj.Tasks:
    startDate, endDate = None, None
    endDate = f'{tsk.Finish}'[:10]
    endDate = datetime.datetime.strptime(endDate, '%Y-%m-%d').strftime('%d/%m')
    startDate = f'{tsk.Start}'[:10]
    startDate = datetime.datetime.strptime(startDate, '%Y-%m-%d').strftime('%d/%m')
    if tsk.Flag1:
        arr.append([tsk.Name, endDate, tsk.Start])


arr = sorted(arr, key=lambda x: x[-1])
for x,y,z in arr:
    print(x,y, sep="|")