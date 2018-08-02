import win32com.client
import datetime
import yaml

slide_objects = {
    0: 'Flowchart: Extract 39',
    1: 'Flowchart: Extract 40',
    2: 'Flowchart: Extract 41',
    3: 'Flowchart: Extract 42',
    4: 'Flowchart: Extract 43',
    5: 'Flowchart: Extract 44',
    6: 'Flowchart: Extract 45',
}


def calculateOffset(dt_string):
    dt_start = datetime.date(2017, 10, 30)
    dt_finish = datetime.date(2018, 3, 25)
    dt = datetime.datetime.strptime(dt_string, '%d/%m/%Y')
    delta1 = dt.date() - dt_start
    delta2 = dt_finish - dt_start
    offset = 40 + (delta1/delta2)*876
    return offset


def print_date(dt_str):
    if dt_str == 'TBC':
        return 'TBC'
    else:
        dt = datetime.datetime.strptime(dt_str, '%d/%m/%Y')
        dt = dt.strftime('%d/%m')
        return dt



data = """
LAN Implementation (Core) Implementation|30/04/2018|TBC
Server Implementation (Core), Backups and Disaster Recovery Implementation |18/05/2018|TBC
Firewall Implementations|11/05/2018|TBC
Video Conferencing Implementation|13/04/2018|TBC
LAN (Edge) Implementation|23/05/2018|TBC
Wireless Implementation|23/05/2018|TBC
Disaster Recovery Test|31/05/2018|TBC
""".splitlines()

data = [x.split('|') for x in data if len(x) > 0]

pp = win32com.client.gencache.EnsureDispatch('Powerpoint.Application')
pp.Visible = True

deck = pp.Presentations.Open(r'C:\Users\rdapaz\Desktop\Harvey Beef\PSRs\Harvey Beef - Security and Infrastructure - Project Status Report 28.05.18.pptx')
slide = deck.Slides(1)

tbl = slide.Shapes("Table 32").Table


for idx, row in enumerate(data):
    tsk_name, approv_finish, curr_finish = row
    tbl.Cell(idx+2,1).Shape.TextFrame.TextRange.Text = tsk_name
    tbl.Cell(idx+2,2).Shape.TextFrame.TextRange.Text = print_date(approv_finish)
    tbl.Cell(idx+2,3).Shape.TextFrame.TextRange.Text = print_date(curr_finish)


# Updating issues:
data = """
 - |
    Missing information pertaining to Emydex/Empired infrastructure design
 - |
    Lack of floor plans are impacting the wireless survey planning activities
"""

data = yaml.load(data)

tbl = slide.Shapes("Table 46").Table

for idx, row in enumerate(data):
    tbl.Cell(idx+2,1).Shape.TextFrame.TextRange.Text = idx+1
    tbl.Cell(idx+2,2).Shape.TextFrame.TextRange.Text = row.strip()


# Updating risks
data = """
- |
    Leadtime for Huawei equipment: Risk of schedule delay as leadtime from ordering is 4-6 weeks
- |
    Telstra Service Establishment: Risk of schedule delays due to provisioning issues and Telstra not being willing to provide us with an ETA    
- | 
    Risk of cascading delays as a result of ER1 being pushed out
"""

data = yaml.load(data)

tbl = slide.Shapes("Table 50").Table

for idx, row in enumerate(data):
    tbl.Cell(idx+2,1).Shape.TextFrame.TextRange.Text = idx+1
    tbl.Cell(idx+2,2).Shape.TextFrame.TextRange.Text = row.strip()

# Recently Completed
data = """
Completed BOM and issued quotes for approval|26/05/2018
Completed set up of Dev/Test environments|26/05/2018
Collected firewalls from Minderoo Office|26/05/2018
""".splitlines()

data = [x.split('|') for x in data if len(x) > 0]
new_data = []
for entry, dt_string in data:
    try:
        dt = datetime.datetime.strptime(dt_string, '%d/%m/%Y')
    except:
        dt = datetime.datetime.strptime('31/12/2031', '%d/%m/%Y')
    new_data.append([entry, dt.strftime('%d/%m'), dt])
new_data = sorted(new_data, key = lambda x: x[-1])
new_data = [[x,y] for x, y, _ in new_data]

print(new_data)

tbl = slide.Shapes("Table 56").Table

for idx, row in enumerate(new_data):
    entry, dt_string = row
    tbl.Cell(idx+2,1).Shape.TextFrame.TextRange.Text = entry
    tbl.Cell(idx+2,2).Shape.TextFrame.TextRange.Text = dt_string

# Updating Current Activities
data = """
- |
    Currently completing internal review of Network Micro Design/LLD
- | 
    Updating schedule to include expected ETAs
- | 
    Scheduling firewall workshop
- |
    Scheduling wireless audit and wireless bridge installation on site (ER1)
- | 
    Configuration of Domain Trusts
"""

data = yaml.load(data)

tbl = slide.Shapes("Table 49").Table

for idx, row in enumerate(data):
    tbl.Cell(idx+1,1).Shape.TextFrame.TextRange.Text = row.strip()

# Next Actions
data = """
Network Detailed Design Document|TBC
Server Infrastructure Design Document|TBC
Complete configuration of domain trusts|TBC
""".splitlines()

data = [x.split('|') for x in data if len(x) > 0]
new_data = []
for entry, dt_string in data:
    try:
        dt = datetime.datetime.strptime(dt_string, '%d/%m/%Y')
    except:
        dt = datetime.datetime.strptime('31/12/2031', '%d/%m/%Y')
    new_data.append([entry, dt.strftime('%d/%m'), dt])
new_data = sorted(new_data, key = lambda x: x[-1])
new_data = [[x,y] for x, y, _ in new_data]

print(new_data)

tbl = slide.Shapes("Table 33").Table

for idx, row in enumerate(new_data):
    entry, dt_string = row
    tbl.Cell(idx+2,1).Shape.TextFrame.TextRange.Text = entry
    tbl.Cell(idx+2,2).Shape.TextFrame.TextRange.Text = dt_string

'''
tbl = slide.Shapes("Table 154").Table
for idx, dt in enumerate(display_dates):
    tbl.Cell(1, idx+1).Shape.TextFrame.TextRange.Text = dt


dates = [x[1] for x in data]

for idx, dt in enumerate(dates):
    slide.Shapes(f'{slide_objects[idx]}').Left = calculateOffset(dt)
    slide.Shapes(f'{slide_objects[idx]}').TextFrame.TextRange.Text = idx+1

Table 13      ID
Table 154     MON 8/5
Table 11      Planned Completion:  53%
'''