import win32com.client

ppt = win32com.client.gencache.EnsureDispatch('Powerpoint.Application')
ppt.Visible = True

deck = ppt.Presentations.Open(r'C:\Users\rdapaz\Desktop\Resources\Program or Project Status Report.pptx')

slide = deck.Slides(5)

for shp in slide.Shapes:
    if shp.Name.startswith('Table'):
        tbl = shp.Table
        tbl.Cell(2,1).Shape.TextFrame.TextRange.Text = shp.Name

