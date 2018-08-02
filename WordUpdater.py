# -*- coding: utf-8 -*-

import win32com.client
import re
import pprint
import json

def pretty_print(o):
    pp = pprint.PrettyPrinter(indent=4)
    pp.pprint(o)


class Word:

    def __init__(self, path):
        self.path = path
        self.app = win32com.client.gencache.EnsureDispatch('Word.Application')
        self.app.Visible = True
        self.app.DisplayAlerts = False
        self.app.Documents.Open(self.path)
        self.doc = self.app.ActiveDocument

    def updateTable(self, bookmark, data, heading_rows=1):
        word_range = self.doc.Bookmarks(bookmark).Range
        table = word_range.Tables(1)
        rows_count = table.Rows.Count
        if not rows_count >= len(data) + heading_rows:
            table.Select()
            self.app.Selection.InsertRowsBelow(NumRows=len(data) + heading_rows - rows_count)
        i = heading_rows
        for entry in data: #sorted(data, key=lambda x: (x[0], x[1])):
            i += 1
            for n in range(len(entry)):
                table.Cell(i, n+1).Range.Text = entry[n]

    def updateIDs(self, bookmark, prefix, offset=0):
        rex = re.compile('[A-Z]+', re.IGNORECASE)
        word_range = self.doc.Bookmarks(bookmark).Range
        table = word_range.Tables(1)
        rows_count = table.Rows.Count
        count = offset
        for rid in range(1, rows_count + 1):
            m = rex.search(table.Cell(rid, 1).Range.Text)
            if m:
                pass
            else:
                count += 1
                table.Cell(rid, 1).Range.Text = f"{prefix}-{count:02}"


def make_data():
    flatten = lambda l: [item for sublist in l for item in sublist]
    data = """
Datacom (General)| 125,460.00 | -   | -   | 12,750.00 | 38,250.00 | 38,250.00 | 36,210.00 | -   | -   | -   
estrat| -   | -   | -   | -   | -   | -   | -   | -   | -   | -   
Chris Campbell| 28,000.00 | -   | 5,600.00 | 5,600.00 | 5,600.00 | 5,600.00 | 5,600.00 | -   | -   | -   
Dan Bronkhorst| -   | -   | -   | -   | -   | -   | -   | -   | -   | -   
Ricardo da Paz| 68,250.00 | -   | 13,000.00 | 13,000.00 | 13,000.00 | 13,000.00 | 16,250.00 | -   | -   | -   
Stuart Stafford| 20,250.00 | -   | 4,050.00 | 4,050.00 | 4,050.00 | 1,350.00 | 6,750.00 | -   | -   | -   
Andrew Johnston| 20,250.00 | -   | 4,050.00 | 4,050.00 | 4,050.00 | 1,350.00 | 6,750.00 | -   | -   | -   
Commissioning Consumables| -   | -   | -   | -   | -   | 2,000.00 | -   | -   | -   | -   
Travel & Accommodation| -   | -   | -   | -   | 1,000.00 | 2,000.00 | 2,000.00 | -   | -   | -   
Wireless| 22,937.00 | -   | -   | -   | -   | 22,937.00 | -   | -   | -   | -   
Servers| 71,400.00 | -   | -   | -   | -   | 71,400.00 | -   | -   | -   | -   
Video Conferencing| 17,047.26 | -   | -   | -   | -   | 17,047.26 | -   | -   | -   | -   
Local Area Network| -   | -   | -   | -   | -   | -   | -   | -   | -   | -   
Harvey Beef:| -   | -   | -   | -   | -   | -   | -   | -   | -   | -   
Admin| -   | -   | -   | -   | -   | -   | -   | -   | -   | -   
        Core Switches| 8,670.00 | -   | -   | -   | -   | 8,670.00 | -   | -   | -   | -   
        Large Switches| 6,303.00 | -   | -   | -   | -   | 6,303.00 | -   | -   | -   | -   
    QA Building:| -   | -   | -   | -   | -   | -   | -   | -   | -   | -   
        Core Switches| 8,670.00 | -   | -   | -   | -   | 8,670.00 | -   | -   | -   | -   
        Large Switches| 6,303.00 | -   | -   | -   | -   | 6,303.00 | -   | -   | -   | -   
    Maintenance:| -   | -   | -   | -   | -   | -   | -   | -   | -   | -   
        Large Switches| 3,151.50 | -   | -   | -   | -   | 3,151.50 | -   | -   | -   | -   
        Small Switches| 1,086.75 | -   | -   | -   | -   | 1,086.75 | -   | -   | -   | -   
    Hay Shed MCC:| -   | -   | -   | -   | -   | -   | -   | -   | -   | -   
        Large Switches| -   | -   | -   | -   | -   | -   | -   | -   | -   | -   
        Small Switches| 1,086.75 | -   | -   | -   | -   | 1,086.75 | -   | -   | -   | -   
    By-product MCC:| -   | -   | -   | -   | -   | -   | -   | -   | -   | -   
        Large Switches| -   | -   | -   | -   | -   | -   | -   | -   | -   | -   
        Small Switches| 2,173.50 | -   | -   | -   | -   | 2,173.50 | -   | -   | -   | -   
    Slaughter Floor MCC:| -   | -   | -   | -   | -   | -   | -   | -   | -   | -   
        Large Switches| -   | -   | -   | -   | -   | -   | -   | -   | -   | -   
        Small Switches| 2,173.50 | -   | -   | -   | -   | 2,173.50 | -   | -   | -   | -   
    Cold Stores:| -   | -   | -   | -   | -   | -   | -   | -   | -   | -   
        Large Switches| -   | -   | -   | -   | -   | -   | -   | -   | -   | -   
        Small Switches| 2,173.50 | -   | -   | -   | -   | 2,173.50 | -   | -   | -   | -   
Fremantle:| -   | -   | -   | -   | -   | -   | -   | -   | -   | -   
    Small Switches| 1,086.75 | -   | -   | -   | -   | 1,086.75 | -   | -   | -   | -   
Minderoo:| -   | -   | -   | -   | -   | -   | -   | -   | -   | -   
        Large Switches| 3,151.50 | -   | -   | -   | -   | 3,151.50 | -   | -   | -   | -   
Swires:| -   | -   | -   | -   | -   | -   | -   | -   | -   | -   
    Small Switches| 1,086.75 | -   | -   | -   | -   | 1,086.75 | -   | -   | -   | -   
SRW:| -   | -   | -   | -   | -   | -   | -   | -   | -   | -   
    Small Switches| 1,086.75 | -   | -   | -   | -   | 1,086.75 | -   | -   | -   | -   
Optics and Patch Leads| 4,678.00 | -   | -   | -   | -   | 4,678.00 | -   | -   | -   | -   
Network Monitoring| 10,740.00 | -   | -   | -   | -   | 10,740.00 | -   | -   | -   | -   
Firewalls| 49,321.00 | -   | -   | -   | 49,321.00 | -   | -   | -   | -   | -   
Telstra WAN| 1,100.00 | -   | -   | -   | -   | 1,100.00 | -   | -   | -   | -   
Harvey Cabling/Fibre/Plant| 75,000.00 | -   | -   | -   | 30,000.00 | 45,000.00 | -   | -   | -   | -   
A/C for QA Building| -   | -   | -   | -   | -   | -   | -   | -   | -   | -   
UPS for QA Building| -   | -   | -   | -   | -   | -   | -   | -   | -   | -   
Environmental Monitoring| -   | -   | -   | -   | -   | -   | -   | -   | -   | -   
Production Hardware| -   | -   | -   | -   || -   | -   | -   | -   | -   
Factory Floor Terminals| 45,000.00 | -   | -   | -   | 45,000.00 | -   | -   | -   | -   | -   
Adjustment| 7,473.49 | -   | -   | -   | -   | -   | 7,473.49 | -   | -   | -   
|615,110.00|0.00|26,700.00|39,450.00|190,271.00|284,655.51|81,033.49|0.00|0.00|0.00
||0.00|26,700.00|66,150.00|256,421.00|541,076.51|622,110.00|622,110.00|622,110.00|622,110.00
""".splitlines()

    data = [x.split('|') for x in data if len(x) > 0]
    # with open('risks.json', 'r') as fin:
    #     data = json.load(fin)
    # new_data = []
    # for row in data:
    #     new_data.append(flatten([[''], row]))
    return data

def main(bookmark, data=[], heading_rows=1):
    my_path = r'C:\Users\rdapaz\Downloads\Harvey Beef Project Toolalla (Infrastructure and Security) - PMP (DRAFT 0.1).docx'
    wd = Word(my_path)
    wd.updateTable(bookmark, data, heading_rows)
    # wd.updateIDs(bookmark, prefix="ID", offset=0)

def mock(data, **kwargs):
    pretty_print(data)

if __name__ == "__main__":
    data = make_data()
    mock(bookmark='bk7', data=data, heading_rows=1)
    main(bookmark='bk7', data=data, heading_rows=1)
