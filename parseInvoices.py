import os
import sys
import win32com.client
import re
import time
import datetime
import pprint
from decimal import Decimal
import json
import psycopg2




path = r'C:\Users\rdapaz\Dropbox\Projects\Harvey Beef\Financials\Invoices'
rex = re.compile(r'xls?', re.IGNORECASE)


def pretty_print(o):
    pp = pprint.PrettyPrinter(indent=4)
    pp.pprint(o)


def processFile(xlfile, data=[]):
    print(xlfile)
    xlApp = win32com.client.gencache.EnsureDispatch('Excel.Application')
    xlApp.Visible = True
    wkBook = xlApp.Workbooks.Open(xlfile)
    sh = wkBook.Worksheets('Sheet1')
    eof = sh.Range('A65536').End(-4162).Row
    for row in range(7, eof+1):
        state = sh.Range(f'A{row}').Value
        if type(state) == str:
            m1 = re.search('^Consultant: (.*)', state, re.IGNORECASE)
            m2 = re.search('^Subtotal', state, re.IGNORECASE)
            consultant = None
            if m1:
                consultant = m1.group(1)
            if m2:
                continue
        else:
            dt = f'{state}'[:10]
            dt = datetime.datetime.strptime(dt, '%Y-%m-%d')
            dt = dt.strftime('%Y-%m-%d')
            project = sh.Range(f'B{row}').Value
            effort = sh.Range(f'C{row}').Value
            if type(effort) == float:
                effort = Decimal(effort)
            data.append([dt, project, consultant, effort])


    print(eof)
    wkBook.Close()


data = []
for dirpath, dirnames, filenames in os.walk(path):
    for filename in filenames:
        if rex.search(filename) and not filename.startswith('~'):
            xlfile = os.path.join(dirpath, filename)
            time.sleep(1)
            processFile(xlfile, data=data)

pretty_print(data)

conn = psycopg2.connect("dbname='estrat_timesheets' user=postgres")
if True:
    cur = conn.cursor()

    sql = """
        CREATE TABLE IF NOT EXISTS \"public\".\"minderoo\" (
            id serial primary key,
            dt date,
            project text,
            consultant text,
            days decimal
            )
    """

    cur.execute(sql)
    
    sql =  """ INSERT INTO \"public\".\"minderoo\" (
                dt, project, consultant, days
                ) VALUES 
                (%s, %s, %s, %s) 
            """
    cur.executemany(sql, data)
    conn.commit()
    conn.close()