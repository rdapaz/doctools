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


data = """
Complete Package 5 (Data Protection)|25/06/2018|TBC
Complete Package 2 (Telephony Service Relocation)|25/06/2018|TBC
Complete Package 3 (Relocation to Malaga DC)|28/06/2018|TBC
Complete Package 6 (Dev & Test Build)|18/07/2018|TBC
Complete Package 1 (Telco Service Establishment)|25/07/2018|TBC
Complete Package 8 (Equipment Destined for Site Decommissioning))|25/07/2018|TBC
Complete Package 9 (Equipment Destined for Decommissioning & Disposal)|25/07/2018|TBC
Complete Package 7 (Disaster Recovery)|30/07/2018|TBC
Complete Package 1 (Telco Service Cancellation)|21/08/2018|TBC
""".splitlines()

data = [x.split('|') for x in data if len(x) > 0]

pp = win32com.client.gencache.EnsureDispatch('Powerpoint.Application')
pp.Visible = True

deck = pp.Presentations.Open(r'C:\Users\rdapaz\Desktop\Belmont DC Exit - Implementation\20180511 - CPM Belmont Relocation - PSR.pptx')
slide = deck.Slides(1)

tbl = slide.Shapes("Table 32").Table

def print_date(dt_str):
    if dt_str == 'TBC':
        return 'TBC'
    else:
        dt = datetime.datetime.strptime(dt_str, '%d/%m/%Y')
        dt = dt.strftime('%d/%m')
        return dt



for idx, row in enumerate(data):
    tsk_name, approv_finish, curr_finish = row
    tbl.Cell(idx+2,1).Shape.TextFrame.TextRange.Text = tsk_name
    tbl.Cell(idx+2,2).Shape.TextFrame.TextRange.Text = print_date(approv_finish)
    tbl.Cell(idx+2,3).Shape.TextFrame.TextRange.Text = print_date(curr_finish)


# Updating issues:
data = """
- |
    Relocation of  equipment to an unsuitable facility
- |
    Delays in procurement process could significantly impact project's ability to achieve the required timelines
- |
    Delay in the finalization of the estrat contract will cause cascading delays in project delivery
"""

data = yaml.load(data)

tbl = slide.Shapes("Table 46").Table

# for idx, row in enumerate(data):
#     tbl.Cell(idx+2,1).Shape.TextFrame.TextRange.Text = idx
#     tbl.Cell(idx+2,2).Shape.TextFrame.TextRange.Text = row


# Updating risks
data = """
- |
    Delays in award of Telecommunications RFP prevent replacement or upgraded WAN implementation prior to the closure of the Belmont DC.
- | 
    Delays in award of Telecommunications RFP prevent the timely delivery of FWaaS for configuration and testing for the MPLS network 
- | 
    CUCM fails during migration and it is unsupported, thereby causing an extended outage
- |
    Availability of the operations team for project related activities (schedule)
- |
    Resource unavailability due to planned/unplanned leave to be taken by key project personnel
- |
    Risk of equipment damage during transit to site
"""

data = yaml.load(data)

tbl = slide.Shapes("Table 50").Table

# for idx, row in enumerate(data):
#     tbl.Cell(idx+2,1).Shape.TextFrame.TextRange.Text = idx+1
#     tbl.Cell(idx+2,2).Shape.TextFrame.TextRange.Text = row.strip()


data = """
|
|
|
|
|
""".splitlines()

data = [x.split('|') for x in data if len(x) > 0]
new_data = []
for entry, dt_string in data:
    try:
        dt = datetime.datetime.strptime(dt_string, '%d/%m/%Y')
    except:
        dt = datetime.datetime.strptime('31/12/2031', '%d/%m/%Y')
    new_data.append([entry, dt_string, dt])
new_data = sorted(new_data, key = lambda x: x[-1])
new_data = [[x,y] for x, y, _ in new_data]

print(new_data)

# tbl = slide.Shapes("Table 56").Table

# for idx, row in enumerate(new_data):
#     entry, dt_string = row
#     tbl.Cell(idx+2,1).Shape.TextFrame.TextRange.Text = entry
#     tbl.Cell(idx+2,2).Shape.TextFrame.TextRange.Text = dt_string



display_dates = [
                '30/10/2017',
                '13/11/2017',
                '27/11/2017',
                '11/12/2017',
                '25/12/2017',
                '08/01/2018',
                '22/01/2018',
                '05/02/2018',
                '19/02/2018',
                '05/03/2018',
                '19/03/2018',
                ]

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