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

fields = """
release|int
station|int
area|text
emydex_form|text
hardware|text
network_points|text
existing_or_new|text
make_model_os|text
screen_size|text
screen_resolution|text
_string|text
sending_weight|text
emydex_polling_or_receiving_continuous|text
max_weight_of_scales|text
weight_increment|text
with_alibi_yes_no|text
emydex_install_date|text
installed_and_connected_date|text
tested_with_emydex_date|text
hardware_cost|decimal
pc_specs|text
""".splitlines()


def int_or_same(int_p):
    try:
        return int(int_p)
    except:
        return int_p


fields = [x.split('|') for x in fields if len(x) > 0]

column_names = [value_colName(x+1) for x in range(len(fields))]

data_fields = {k: v for k, v in zip(column_names, fields)}
pretty_printer(data_fields)

arr = []
for col_name, rest in data_fields.items():
    field_name, data_type = rest
    arr.append(dict(col_name=col_name, field_name=field_name, data_type=data_type))

xlApp = win32com.client.gencache.EnsureDispatch('Excel.Application')
xlApp.Visible = True


path = r'C:\Users\rdapaz\Desktop\Harvey Beef Hardware 2018052.xlsx'
wk = xlApp.Workbooks.Open(path)
sh = wk.Worksheets('Sheet1')

EOF = sh.Range('B65536').End(-4162).Row

vals = []
for row in range(2, EOF+1):
    if not sh.Range(f'A{row}').Value and not sh.Range(f'B{row}').Value:
        pass
    else:
        try:
            for p in arr:
                if p['data_type'] == 'text':
                    exec(f"{p['field_name']} = str(sh.Range('{p['col_name']}{row}').Value) if sh.Range('{p['col_name']}{row}').Value else None")
                elif p['data_type'] == 'decimal':
                    exec(f"{p['field_name']} = Decimal(sh.Range('{p['col_name']}{row}').Value) if sh.Range('{p['col_name']}{row}').Value else 0.0")
                elif p['data_type'] in ('int', 'long'):
                    exec(f"{p['field_name']} = int_or_same(sh.Range('{p['col_name']}{row}').Value) if sh.Range('{p['col_name']}{row}').Value else 0")
                elif p['data_type'] == 'date':
                    exec(f"{p['field_name']} = str(sh.Range('{p['col_name']}{row}').Value) if sh.Range('{p['col_name']}{row}').Value else '1970-01-01'")
                    exec(f"{p['field_name']} = {p['field_name']}[:10]")
                    exec(f"{p['field_name']} = datetime.datetime.strptime({p['field_name']}, '%Y-%m-%d')")
            exec("vals.append({})".format([eval(p['field_name']) for p in arr]))
        except Exception as e:
            print('Error found in row {}'.format(row))
            print(e)

pretty_printer(vals)

conn = psycopg2.connect("dbname='HarveyBeef' user=postgres")

if True:
    cur = conn.cursor()
    dummy_arr = []

    # sql1 = f'CREATE TABLE IF NOT EXISTS \"public\".\"{sh.Name}\" (id serial primary key, '
    sql1 = f'CREATE TABLE IF NOT EXISTS \"public\".\"field_equip\" (id serial primary key, '
    dummy_arr = [f"{x['field_name']} {x['data_type']}" for x in arr]
    pretty_printer(dummy_arr)
    dummy = ", ".join(dummy_arr)
    sql2 = "\n)"
    sql = sql1 + dummy + sql2
    print(sql)
    cur.execute(sql)
    conn.commit()

    # sql1 = f'INSERT INTO \"public\".\"{sh.Name}\" (\n'
    sql1 = f'INSERT INTO \"public\".\"field_equip\" (\n'
    dummy_arr1 = [x['field_name'] for x in arr]
    dummy1 = ", ".join(dummy_arr1)
    sql2 = "\n)\nVALUES\n("
    sql2 = re.sub(r'^\s+', '', sql2, re.MULTILINE)
    dummy_arr2 = ['%s'] * len(dummy_arr1)
    dummy2 = ", ".join(dummy_arr2)
    sql3 = "\n)"
    sql = sql1 + dummy1 + sql2 + dummy2 + sql3
    print(sql)
    cur.executemany(sql, vals)
    conn.commit()
    conn.close()
