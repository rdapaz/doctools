
import win32com.client

orderBy = {
    '1.2.6': 0,
    '4.3.7.4': 1,
    '4.3.2.2.3': 2,
    '4.3.2.1.4': 3,
    '4.3.6': 4,
    '4.3.3.3': 5,
    '4.3.1.3': 6,
    '5.1.13': 7,
    '5.2.10': 8,
    '5.2.11': 9,
    '5.3.5': 10
}

projApp = win32com.client.gencache.EnsureDispatch('MSProject.Application')
projApp.Visible = True
path = r'C:\Users\rdapaz\Desktop\Belmont DC Exit - Implementation\Belmont DC Relocation Project_20180406.mpp'
projApp.FileOpen(path)
pj = projApp.ActiveProject

arr = []
for tsk in pj.Tasks:
    if tsk.Text23:
        arr.append([tsk.WBS, tsk.Name, tsk.Text21, tsk.Text23])

arr = sorted(arr, key=lambda x: orderBy[x[0]])
for wbs, name, finish, perc in arr:
    print(wbs, name, finish, perc, sep="|")
