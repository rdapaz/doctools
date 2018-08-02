import yaml
import pprint
import psycopg2


def pretty_printer(o):
    pp = pprint.PrettyPrinter(indent=4)
    pp.pprint(o)

with open(r'C:\Users\rdapaz\Documents\scripts\doctools\goods_received.yaml', 'r') as fin:
    data = yaml.load(fin)

my_arr = []
for it, props in data.items():
    for img, arr in props.items():
        for serial in arr:
            my_arr.append(['', it, img, serial])

conn = psycopg2.connect("dbname='bom' user=postgres")
cur = conn.cursor()

sql = """
    CREATE TABLE IF NOT EXISTS \"public\".\"hw_recon\" (
        id serial primary key,
        manufacturer text,
        model text,
        evidence_id text,
        serial_no text
        )
"""
cur.execute(sql)

sql =  """ INSERT INTO \"public\".\"hw_recon\" (
            manufacturer, model, evidence_id, serial_no
            ) VALUES 
            (%s, %s, %s, %s) 
        """
cur.executemany(sql, my_arr)
conn.commit()
conn.close()