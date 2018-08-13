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
Diamond 61:
    Idx: 1
Diamond 85:
    Idx: 2
Diamond 86:
    Idx: 3
Diamond 87:
    Idx: 4
Diamond 88:
    Idx: 5
Diamond 89:
    Idx: 6
Diamond 90:
    Idx: 7
Diamond 91:
    Idx: 8
Diamond 92:
    Idx: 9
Diamond 93:
    Idx: 10
Diamond 94:
    Idx: 11
Diamond 95:
    Idx: 12
Diamond 96:
    Idx: 13
Diamond 97:
    Idx: 14
TextBox 1:
    Idx: 1
TextBox 98:
    Idx: 2
TextBox 99:
    Idx: 3
TextBox 100:
    Idx: 4
TextBox 101:
    Idx: 5
TextBox 102:
    Idx: 6
TextBox 103:
    Idx: 7
TextBox 104:
    Idx: 8
TextBox 105:
    Idx: 9
TextBox 106:
    Idx: 10
TextBox 107:
    Idx: 11
TextBox 108:
    Idx: 12
TextBox 109:
    Idx: 13
TextBox 110:
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
    return 0
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
    dt_start = datetime.date(2018, 8, 13)
    dt_finish = datetime.date(2018, 9, 9)
    dt1 = datetime.datetime.strptime(stt, '%d/%m/%Y').date()
    dt2 = datetime.datetime.strptime(fin, '%d/%m/%Y').date()
    gap = adjustForGaps(dt1, dt2)
    delta1 = (dt2 - dt1).days + 1 # if (dt2 - dt1).days > 0 else (dt2 - dt1).days + 1 
    delta2 = (dt_finish - dt_start).days + 1
    start = 242.57 + gap + float(((dt1 - dt_start).days/delta2))*684.3
    length = float((delta1/delta2))*684.3 if float((delta1/delta2))*684.3 < 684.0 else 684.3
    print(dt1, dt2, gap, sep="|")
    return start, length


shape_data = yaml.load(slide_shape_data)

timedata = '''
Install WAP01, WAP02 and WAP03|13/08/2018|13/08/2018
Install WAP04 and WAP05|14/08/2018|14/08/2018
Install WAP06, WAP07, WAP11 and WAP12|15/08/2018|15/08/2018
Install WAP11, WAP12, WAP14, WAP19, WAP20, WAP21 and WAP22|17/08/2018|17/08/2018
Install WAP08, WAP10 and WAP13|22/08/2018|22/08/2018
Install WAP23 and WAP24|23/08/2018|23/08/2018
Install WAP09, WAP15 and WAP17|24/08/2018|24/08/2018
Finalise punchlist items|27/08/2018|7/9/2018
'''.splitlines()

timedata ={idx+1:row for idx, row in enumerate([x.split('|') for x in timedata if len(x) > 0])}
print(timedata)


pp = win32com.client.gencache.EnsureDispatch('Powerpoint.Application')
pp.Visible = True

deck = pp.Presentations.Open(r'C:\Users\rdapaz\Desktop\Resources\Timeline templatev2.pptm')
slide = deck.Slides(2)
slide.Select()

rounded_recs = {}

for shp in slide.Shapes:
    if shp.Name.startswith('Rectangle') and shp.Name in shape_data:
        shp_idx = shape_data[shp.Name]['Idx']
        if shp_idx > len(timedata):
            continue
        else:
            if shp_idx in timedata and shp.HasTextFrame:
                shp.TextFrame.TextRange.Text = timedata[shp_idx][0]

    elif shp.Name.startswith('Rounded Rectangle') and shp.Name in shape_data:
        shp_idx = shape_data[shp.Name]['Idx']
        if shp_idx:
            if shp_idx <= len(timedata):
                stt = timedata[shp_idx][1]
                fin = timedata[shp_idx][2]
                shp.Left = calcStartAndLength(stt, fin)[0]
                shp.Width = calcStartAndLength(stt, fin)[1]
                stt = datetime.datetime.strptime(stt, '%d/%m/%Y').strftime('%d/%m')
                fin = datetime.datetime.strptime(fin, '%d/%m/%Y').strftime('%d/%m')
                rounded_recs[shp_idx] = (shp.Left + shp.Width, fin)
                shp.TextFrame.TextRange.Text = ''
            else:
                continue

    elif shp.Name.startswith('Diamond') and shp.Name in shape_data:
        shp_idx = shape_data[shp.Name]['Idx']
        if shp_idx:
            if shp_idx <= len(timedata):
                shp.Left = rounded_recs[shp_idx][0] - shp.Width/2
                shp.TextFrame.TextRange.Text = ''
            else:
                continue

    elif shp.Name.startswith('TextBox') and shp.Name in shape_data:
        shp_idx = shape_data[shp.Name]['Idx']
        if shp_idx:
            if shp_idx <= len(timedata):
                shp.Left = rounded_recs[shp_idx][0] + 5
                shp.TextFrame.TextRange.Text = rounded_recs[shp_idx][1]
            else:
                continue