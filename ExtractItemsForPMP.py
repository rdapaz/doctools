import re
import win32com.client
import datetime

path = r'C:\Users\rdapaz\Desktop\Belmont Relocation Project - Indicative Project Schedule (RDP).mpp'
app = win32com.client.gencache.EnsureDispatch('MSPRoject.Application')
app.FileOpen(path)
pj = app.ActiveProject

arr = []
for tsk in pj.Tasks:
    if tsk.Text1:
        endDate = f'{tsk.Finish}'[:10]
        endDate = datetime.datetime.strptime(endDate, '%Y-%m-%d').strftime('%d %B %Y')
        arr.append([tsk.Text1, endDate, tsk.Finish])

arr = sorted(arr, key=lambda x: x[-1])
new_arr = [[x[0], x[1]] for x in arr]
for x,y in new_arr:
    print(x,y,sep="|")