import win32com.client
import datetime

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
    dt_start = datetime.date(2017, 1, 1)
    dt_finish = datetime.date(2017, 12, 31)
    dt = datetime.datetime.strptime(dt_string, '%d/%m/%Y')
    delta1 = dt.date() - dt_start
    delta2 = dt_finish - dt_start
    offset = 30.12 + (delta1/delta2)*686.54
    return offset


'''
Public Sub test()
Dim shp As PowerPoint.Shape
Dim sld As PowerPoint.Slide

Set sld = Application.ActivePresentation.Slides(6)

For Each shp In sld.Shapes
    If shp.Name = "Flowchart: Extract 39" Then
        Debug.Print shp.Left
        shp.Select
    End If
Next shp
End Sub
'''


data = """
Project kick-off|23/01/2017|1/10/2016
Scoping complete|30/01/2017|N/A
Planning and Design complete|15/05/2017|N/A
P2S Activities complete|19/06/2017|N/A
Build complete|21/06/2017|N/A
AIS Notification|06/07/2017|N/A
Project Closeout|5/12/2017|31/03/2017
""".splitlines()

data = [x.split('|') for x in data if len(x) > 0]

pp = win32com.client.gencache.EnsureDispatch('Powerpoint.Application')
pp.Visible = True

deck = pp.Presentations.Open(r'E:\__NEW__\__NEW__\APIC-EM\I0431900 APIC-EM Corporate Deployment_2.pptx')
slide = deck.Slides(6)

tbl = slide.Shapes("Table 13").Table

for idx, row in enumerate(data):
    tsk_name, finish, must_finish = row
    tbl.Cell(idx+2,1).Shape.TextFrame.TextRange.Text = idx+1
    tbl.Cell(idx+2,2).Shape.TextFrame.TextRange.Text = tsk_name
    tbl.Cell(idx+2,3).Shape.TextFrame.TextRange.Text = must_finish
    tbl.Cell(idx+2,4).Shape.TextFrame.TextRange.Text = finish


display_dates = [
                'Jan-17',
                'Feb-17',
                'Mar-17',
                'Apr-17',
                'May-17',
                'Jun-17',
                'Jul-17',
                'Aug-17',
                'Sep-17',
                'Oct-17',
                'Nov-17',
                'Dec-17',
                ]

tbl = slide.Shapes("Table 154").Table
for idx, dt in enumerate(display_dates):
    tbl.Cell(1, idx+1).Shape.TextFrame.TextRange.Text = dt


finish_dates = [x[1] for x in data]

for idx, dt in enumerate(finish_dates):
    slide.Shapes(f'{slide_objects[idx]}').Left = calculateOffset(dt)
    slide.Shapes(f'{slide_objects[idx]}').TextFrame.TextRange.Text = idx+1

'''
Table 13      ID
Table 154     MON 8/5
Table 11      Planned Completion:  53%
'''