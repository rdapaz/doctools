# -*- coding: utf-8 -*-

import win32com.client
import re
import pprint
import yaml


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
    new_data = []
    with open(r'C:\Users\rdapaz\Documents\doctools\minutes_from_schedule_meeting.yaml', 'r') as f:
        data = yaml.load(f)
    pretty_print(data)
    for entry in data:
            new_data.append(['', f"{entry['Topic']}:\n{entry['Description']}", entry['Responsible'], entry['Lodged'], entry['Due']])
    return new_data

def main(bookmark, data=[], heading_rows=1):
    my_path = r'C:\Users\rdapaz\Desktop\Project Toolalla PM Schedule Planning Meeting.docx'
    wd = Word(my_path)
    wd.updateTable(bookmark, data, heading_rows)
    wd.updateIDs(bookmark, prefix="AP", offset=0)

def mock(data, **kwargs):
    pretty_print(data)

if __name__ == "__main__":
    data = make_data()
    mock(bookmark='actions', data=data, heading_rows=1)
    main(bookmark='actions', data=data, heading_rows=1)
