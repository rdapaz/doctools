import win32com.client
import datetime


pjApp = win32com.client.gencache.EnsureDispatch('MSProject.Application')
pjApp.Visible = True
path = r'C:\Users\rdapaz\Desktop\Minderoo - Project Schedule.mpp'
pjApp.FileOpenEx(Name=path)
pj = pjApp.ActiveProject


dtNow = datetime.datetime.now().date()
lastWeek = dtNow+datetime.timedelta(days=-7)
nextWeek = dtNow+datetime.timedelta(days=7)

for tsk in pj.Tasks:
    tskStart = datetime.datetime.strptime(f'{tsk.Start}'[:10], '%Y-%m-%d').date()
    tskFinish = datetime.datetime.strptime(f'{tsk.Finish}'[:10], '%Y-%m-%d').date()
    if (tskStart >= lastWeek and tskFinish < nextWeek) or (tskStart <= lastWeek and tskFinish >= lastWeek) or (tskFinish > nextWeek and tskStart <= nextWeek):
        print(tsk.Name, tskStart, tskFinish, sep="|")

