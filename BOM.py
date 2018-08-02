import pprint
import win32com.client
from decimal import Decimal
import datetime
import psycopg2
import re


def pretty_printer(o):
	pp = pprint.PrettyPrinter(indent=4)
	pp.pprint(o)


def value_colName(iVal):
	retVal = None
	if iVal <= 26:
		retVal = chr(64+iVal)
	else:
		m = int(iVal/26)
		n = iVal - m*26
		retVal = f'{value_colName(m)}{value_colName(n)}' 
	return retVal


xlApp = win32com.client.gencache.EnsureDispatch('Excel.Application')
xlApp.Visible = True
path = r'C:\Users\rdapaz\Desktop\Harvey Beef\Harvey Beef - Infrastructure Refresh Financial Tracker V1.xlsx'
wk = xlApp.Workbooks.Open(path)
sh = wk.Worksheets('Hardware')

conn = psycopg2.connect("dbname='bom' user=postgres")
cur = conn.cursor()

sql = """
    SELECT
        agg.category,
        agg.ID,
        agg.description,
        agg.quantity,
        agg.unit_cost,
        agg.delivery_time,
        agg.vendor_notes 
    FROM
        (
    SELECT
        'Video Conferencing' AS category,
        vc.ID,
        vc.description,
        vc.quantity,
        vc.unit_cost,
        vc.delivery_time,
        vc.vendor_notes 
    FROM
        "Video Conferencing" vc UNION
    SELECT
        'LAN/Security/WAN' AS category,
        lan.ID,
        lan.description,
        lan.quantity,
        lan.unit_cost,
        lan.delivery_time,
        lan.vendor_notes 
    FROM
        "LAN-Security-WAN" lan UNION
    SELECT
        'Wireless' AS category,
        w.ID,
        w.description,
        w.quantity,
        w.unit_cost,
        w.delivery_time,
        w.comments AS vendor_notes 
    FROM
        "Wireless" w UNION
    SELECT
        'Servers' AS category,
        s.ID,
        s.description,
        s.quantity,
        s.unit_cost,
        s.delivery_time,
        s.vendors_notes 
    FROM
        "Physical Servers" s 
        ) AS agg 
    WHERE
        agg.quantity > 0 
    ORDER BY
        1,2
"""

cur.execute(sql)
START_ROW = 4
row = START_ROW
for idx, row_data in enumerate(cur.fetchall()):
    category, _, description, qty, unit_cost, lead_time, vendor_notes = row_data
    print(category, description, qty, unit_cost, lead_time, sep="|")
    row += 1
    sh.Range(f'D{row}').Value = idx+1
    sh.Range(f'E{row}').Value = category
    sh.Range(f'G{row}').Value = description
    sh.Range(f'H{row}').Value = qty
    sh.Range(f'I{row}').Value = unit_cost
    sh.Range(f'K{row}').Value = lead_time
    rge = sh.Range(f'G{row}')
    if rge.Comment:
        rge.ClearComments()
    rge.AddComment(Text=vendor_notes)
    shp = rge.Comment.Shape
    shp.TextFrame.AutoSize = True
    # lArea = shp.Width * shp.Height
    # shp.Width = 300
    # shp.Height = (lArea / shp.Width)       # used shp.width so that it is less work to change final width
    shp.TextFrame.AutoMargins = False
    shp.TextFrame.MarginBottom = 0      # margins need to be tweaked
    shp.TextFrame.MarginTop = 0
    shp.TextFrame.MarginLeft = 0
    shp.TextFrame.MarginRight = 0

conn.close()
wk.Save()
