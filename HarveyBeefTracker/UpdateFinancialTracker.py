# Andrew Johnston
# Chris Campbell
# Paul Cooper
# Ricardo Da Paz
# Stuart Stafford

import win32com.client
import psycopg2
import datetime


def column_name(iVal):
    retVal = None
    if iVal <= 26:
        retVal = chr(64+iVal)
    else:
        m = int(iVal/26)
        n = iVal - m*26
        if n==0:
            m = m-1
            n = 26
        retVal = '{}{}'.format(column_name(m), column_name(n))
    return retVal


conn = psycopg2.connect("dbname='estrat_timesheets' user=postgres")
cur = conn.cursor()

sql = '''
    SELECT DISTINCT
    to_char( dt, 'YYYY-MM' ) AS dt,
    consultant,
    round( SUM ( days ), 2 ) 
    FROM
        "minderoo" 
    WHERE
        project ~ 'Infrastructure' 
    GROUP BY
        1,
        2 
    ORDER BY
    1
'''

cur.execute(sql)


data = {}

for row in cur.fetchall():
    year_mon, res, days = row
    print(year_mon, res, days, sep="|")
    if year_mon not in data:
        data[year_mon] = {}
    if res not in data[year_mon]:
        data[year_mon][res] = 0.0
    data[year_mon][res] += float(days)

xlApp = win32com.client.gencache.EnsureDispatch('Excel.Application') 
xlApp.Visible = True
path = r'C:\Users\rdapaz\Dropbox\Projects\Harvey Beef\Financials\Harvey Beef - Infrastructure Refresh Financial Tracker V1.1.xlsx'
wkBook = xlApp.Workbooks.Open(path)
sh = wkBook.Worksheets('Forecast')

for col in range(33, 58):
    if col == 45:
        continue
    else:
        dt = sh.Range(f'{column_name(col)}28').Value
        dt = f'{dt}'[:10]
        dt = datetime.datetime.strptime(dt, '%Y-%m-%d').strftime('%Y-%m')
        print(dt)
        for row in range(30, 36):
            res = sh.Range(f'B{row}').Value
            if dt in data and res in data[dt]:
                sh.Range(f'{column_name(col)}{row}').Value = data[dt][res]
            elif sh.Range(f'{column_name(col)}{row}').Value:
                pass
            else:
                sh.Range(f'{column_name(col)}{row}').Value = 0.0