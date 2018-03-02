# -*- coding: utf-8 -*-

import win32com.client
import re
import pprint
import datetime
import json
import yaml
import time

def pretty_print(o):
    pp = pprint.PrettyPrinter(indent=4)
    pp.pprint(o)


def split_chunks(a, n):
    k, m = divmod(len(a), n)
    return (a[i * k + min(i, m):(i + 1) * k + min(i + 1, m)] for i in range(n))

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
            print(len(entry))
            i += 1
            for n in range(len(entry)):
                table.Cell(i, n+1).Range.Text = entry[n]

    def updateIDs(self, bookmark, prefix):
        rex = re.compile('[A-Z]+', re.IGNORECASE)
        word_range = self.doc.Bookmarks(bookmark).Range 
        table = word_range.Tables(1)
        rows_count = table.Rows.Count
        count = 0
        for rid in range(1, rows_count+1):
            m = rex.search(table.Cell(rid, 1).Range.Text)
            if m:
                pass
            else:
                count+=1
                table.Cell(rid,1).Range.Text = f"{prefix}{str(count).zfill(3)}"


def make_data():
    data = """
Macro Design|09 March 2018
PMP Signed Off|15 March 2018
Micro Design (Cloud Platform and Workload Migrations)|29 March 2018
Migration Plan (Cloud Platform and Workload Migrations)|04 April 2018
Micro Design (Data Protection)|06 April 2018
Micro Design (DMZ & VPN Relocation)|11 April 2018
Migration Plan (DMZ & VPN Relocation)|11 April 2018
Migration Plan (Telephony Services)|25 June 2018
Migrate Workloads|09 July 2018
Belmont DC decommissioned|24 July 2018
DR Test PMP|26 July 2018
DR Test Plan Completed|05 September 2018
DR Test PIR and Completion Report|23 October 2018
""".splitlines()
    data = [x.split('|') for x in data if len(x) > 0]
    return data
    # with open(r'C:\Users\rdapaz\Documents\python_assorted\newer_lessons.json', 'r') as f:
    #     data =json.load(f)
    # new_data = [['', a,b,c,d,e] for a,b,c,d,e in data]
    # for k, v in data.items():
    #     new_data.append([k, v])
    # new_data = [x.split('|') for x in data if len(x) > 0]
    # return new_data

def main(bookmark, data=[], heading_rows=1):
    my_path = r'C:\Users\rdapaz\Desktop\Belmont DC Relocation - Project Management Plan.docm'
    wd = Word(my_path)
    wd.updateTable(bookmark, data, heading_rows)
    # time.sleep(1)
    # wd.updateIDs(bookmark, prefix="L")

def mock(data, **kwargs):
    pretty_print(data)
    
if __name__ == "__main__":
    data = make_data()
    mock(bookmark='bk1', data=data, heading_rows=1)
    main(bookmark='bk1', data=data, heading_rows=1)
    # main(bookmark='Financials1', data=data, heading_rows=1)