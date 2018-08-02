# -*- coding: utf-8 -*-
import win32com.client
import re
import os
import pprint
import datetime

# Python 2 to Python 3 fix
try:
    from cStringIO import StringIO
except:
    from io import StringIO

# Writing to a buffer
output = StringIO()


def pretty_print(o):
    pp = pprint.PrettyPrinter(indent=4)
    pp.pprint(o)


def tidyDate(pjDate):
    sDate = f"{pjDate}"[:10]
    return datetime.datetime.strptime(sDate, '%Y-%m-%d').strftime('%d/%m/%Y')

proj = win32com.client.gencache.EnsureDispatch('MSProject.Application')
proj.Visible = True

ROOT = r'C:\Users\rdapaz\AppData\Local\Microsoft\Windows\INetCache\Content.Outlook\TDTIZAWP'
filepath = os.path.join(ROOT, 'HL Project Plan v1.1 20180327 (BM).mpp')

proj.FileOpen(filepath)
my_proj = proj.ActiveProject

for tsk in my_proj.Tasks:

    if tsk is None:
        continue
    else:
        SPACES = '    '
        task_desc = None
        if tsk.OutlineChildren.Count == 0:
            if tsk.PercentComplete == 100:
                task_desc = '{}{}: {}|{}'.format(
                                    SPACES * int(tsk.OutlineLevel -1),
                                    tsk.Name,
                                    tidyDate(tsk.Start),
                                    tidyDate(tsk.Finish)
                                  )
            else:
                task_desc = '{}{}: {}d|{}'.format(
                                                    SPACES * int(tsk.OutlineLevel -1),
                                                    tsk.Name,
                                                    tsk.Duration/480.0,
                                                    tsk.ResourceNames
                                                  )

        else:
            task_desc = '{}{}:'.format(
                                        SPACES * int(tsk.OutlineLevel -1),
                                        tsk.Name
                                      )
        print(task_desc, file=output)

json_text = output.getvalue()
with open('datacom_tasks.yaml', 'w') as f:
    f.write(json_text)