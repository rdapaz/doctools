# coding: utf-8

import yaml
import os
import pprint
import win32com.client


def pretty_printer(o):
    pp = pprint.PrettyPrinter(indent=4)
    pp.pprint(o)


def topic_order(k):
    hash = {
    'General': 1,
    'Telecommunications RFP': 2,
    'Disaster Recovery Testing': 3,
    'Backup Import': 4,
    'Training': 5
    }
    return hash.setdefault(k, 0)

class Word:
    def __init__(self, path):
        self.path = path
        self.app = win32com.client.gencache.EnsureDispatch("Word.Application")
        self.app.Visible = True
        self.app.DisplayAlerts = False
        self.doc = self.app.Documents.Open(self.path)
        self.rge = self.doc.Range(Start=0, End=0)

    def initialise(self):
        self.rge.InsertAfter("\n\n")
        self.rge.Select

    def insert_table(self, autoTextName = 'a1'):
        self.app.Selection.TypeText(Text=autoTextName)
        self.app.Selection.Range.InsertAutoText()
        rge = self.doc.Range()
        rge.Select()
        rge.Collapse(0)
        rge.InsertAfter("\n\n")
        rge.Select()

    def updateTable(self, tbl_id, data, heading_rows=1):
        table = self.doc.Tables(tbl_id)
        rows_count = table.Rows.Count
        if not rows_count >= len(data) + heading_rows:
            table.Select()
            self.app.Selection.InsertRowsBelow(NumRows=len(data) + heading_rows - rows_count)
        i = heading_rows
        for entry in data:  # sorted(data, key=lambda x: (x[0], x[1])):
            i += 1
            for n in range(len(entry)):
                table.Cell(i, n + 1).Range.Text = entry[n]

    def generateEntry(self, text, textType):
        self.initialise()
        self.app.Selection.TypeText(Text=textType)
        self.app.Selection.Range.InsertAutoText()
        rge = self.doc.Range()
        rge.Select()
        rge.Collapse(0)
        rge.InsertAfter(text)
        rge.Select()

if __name__ == "__main__":
    ROOTDIR = r'C:\Users\rdapaz\Desktop'
    wordTemplatePath = os.path.join(ROOTDIR, 'CPM SteerCo Template.docx')
    yamlFile = r'C:\Users\rdapaz\Documents\scripts\doctools\cpm_steerco3_minutes.yaml'
    with open(yamlFile, 'r') as f:
        data = yaml.load(f)
    word = Word(wordTemplatePath)
    pretty_printer(data)
    topics = sorted(list(set(x['Topic'] for x in data)), key=lambda x: topic_order(x))
    for topic in topics:
        for p in data:
            if p['Action'] and p['Topic'] == topic:
                    print(p['Action'], p['Responsible'], sep="|")
    # word.generateMinutes(data)