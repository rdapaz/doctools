import yaml
import win32com.client
import datetime

slide_shape_data = '''
Rectangle 49:
    Idx: 1
Rectangle 54:
    Idx: 2
Rectangle 55:
    Idx: 3
Rectangle 58:
    Idx: 4
Rectangle 59:
    Idx: 5
Rectangle 60:
    Idx: 6
Rectangle 62:
    Idx: 7
Rectangle 64:
    Idx: 8
Rectangle 65:
    Idx: 9
Rectangle 66:
    Idx: 10
Rectangle 67:
    Idx: 11
Rectangle 68:
    Idx: 12
Rectangle 69:
    Idx: 13
Rectangle 70:
    Idx: 14
Rounded Rectangle 56:
    Idx: 1
Rounded Rectangle 71:
    Idx: 2
Rounded Rectangle 72:
    Idx: 3
Rounded Rectangle 73:
    Idx: 4
Rounded Rectangle 74:
    Idx: 5
Rounded Rectangle 75:
    Idx: 6
Rounded Rectangle 76:
    Idx: 7
Rounded Rectangle 77:
    Idx: 8
Rounded Rectangle 78:
    Idx: 9
Rounded Rectangle 79:
    Idx: 10
Rounded Rectangle 80:
    Idx: 11
Rounded Rectangle 81:
    Idx: 12
Rounded Rectangle 82:
    Idx: 13
Rounded Rectangle 83:
    Idx: 14
'''

def calcNextSunday(dt):
    if dt.strftime('%A') == 'Sunday':
        return dt
    else:
        dt1 = dt
        while dt1.strftime('%A') != 'Sunday':
            dt1 += datetime.timedelta(days=1)
        return dt1


def adjustForGaps(dt1, dt2):
    GAP = 10
    sunday1 = calcNextSunday(dt1)
    print(sunday1)
    sunday2 = calcNextSunday(dt2)
    print(sunday2)
    if 0 <= (sunday2 - sunday1).days < 7:
        return 0
    elif 7 <= (sunday2 - sunday1).days < 14:
        return 1*GAP
    elif 14 <= (sunday2 - sunday1).days < 21:
        return 2*GAP
    elif 21 <= (sunday2 - sunday1).days < 28:
        return 3*GAP
     

def calcStartAndLength(stt, fin):
    dt_start = datetime.date(2018, 7, 30)
    dt_finish = datetime.date(2018, 8, 26)
    dt1 = datetime.datetime.strptime(stt, '%d/%m/%Y').date()
    dt2 = datetime.datetime.strptime(fin, '%d/%m/%Y').date()
    gap = adjustForGaps(dt1, dt2)
    delta1 = (dt2 - dt1).days + 1 # if (dt2 - dt1).days > 0 else (dt2 - dt1).days + 1 
    delta2 = (dt_finish - dt_start).days + 1
    start = 255.76 + gap + float(((dt1 - dt_start).days/delta2))*655.5
    length = float((delta1/delta2))*655.5 if float((delta1/delta2))*655.5 < 655 else 655.5
    print(dt1, dt2, gap, sep="|")
    return start, length


shape_data = yaml.load(slide_shape_data)

timedata = '''
Core Switches Installed|30/07/2018|26/08/2018
Servers Installed|30/07/2018|31/07/2018
Access Switches Installed|30/07/2018|31/07/2018
Additional SFPs|30/07/2018|5/08/2018
LANs interconnected|6/08/2018|6/08/2018
Cabling for WAPs|6/08/2018|18/08/2018
Wireless Controller Live|20/08/2018|21/08/2018
WAPS installed and live|21/08/2018|21/08/2018
Join new servers to domain controllers|6/08/2018|7/08/2018
Backup solutin installed|7/08/2018|7/08/2018
Firewalls installed|6/08/2018|9/08/2018
Firewalls configured/MPLS commissioned|9/08/2018|18/08/2018
Test Workloads migrated|7/08/2018|8/08/2018
Backup Tested & workloads migrated|9/08/2018|17/08/2018
'''.splitlines()

timedata ={idx+1:row for idx, row in enumerate([x.split('|') for x in timedata if len(x) > 0])}
print(timedata)


pp = win32com.client.gencache.EnsureDispatch('Powerpoint.Application')
pp.Visible = True

deck = pp.Presentations.Open(r'C:\Users\rdapaz\Desktop\Resources\Timeline templatev1.pptm')
slide = deck.Slides(2)
slide.Select()

for shp in slide.Shapes:
    if shp.Name.startswith('Rectangle') and shp.Name in shape_data:
        shp_idx = shape_data[shp.Name]['Idx']
        if shp_idx in timedata and shp.HasTextFrame:
            shp.TextFrame.TextRange.Text = timedata[shp_idx][0]

    elif shp.Name.startswith('Rounded Rectangle') and shp.Name in shape_data:
        shp_idx = shape_data[shp.Name]['Idx']
        if shp_idx:
            stt = timedata[shp_idx][1]
            fin = timedata[shp_idx][2]
            shp.Left = calcStartAndLength(stt, fin)[0]
            shp.Width = calcStartAndLength(stt, fin)[1]
            stt = datetime.datetime.strptime(stt, '%d/%m/%Y').strftime('%d/%m')
            fin = datetime.datetime.strptime(fin, '%d/%m/%Y').strftime('%d/%m')
            shp.TextFrame.TextRange.Text = f'{stt} -> {fin}'