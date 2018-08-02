import re
import win32com.client
import datetime

path = r'C:\Users\rdapaz\Dropbox\Projects\Harvey Beef\Infrastructure Gantt.mpp'
app = win32com.client.gencache.EnsureDispatch('MSPRoject.Application')
app.FileOpen(path)
pj = app.ActiveProject

arr = []
for tsk in pj.Tasks:
    startDate, endDate = None, None
    if tsk.Flag1 == True:
        if tsk.Text2.lower() == 'start':
            startDate = f'{tsk.Start}'[:10]
            startDate = datetime.datetime.strptime(startDate, '%Y-%m-%d').strftime('%d %B %Y')
        elif tsk.Text2.lower() == 'finish':
            endDate = f'{tsk.Finish}'[:10]
            endDate = datetime.datetime.strptime(endDate, '%Y-%m-%d').strftime('%d %B %Y')
        dte = startDate if startDate else endDate
        arr.append([f"{tsk.Text2.upper()}: {tsk.Name}", dte, tsk.Finish])

arr = sorted(arr, key=lambda x: x[-1])
new_arr = [[x[0], x[1]] for x in arr]
for x,y in new_arr:
    print(x,y,sep="|")