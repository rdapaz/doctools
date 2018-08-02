import psycopg2
import re



servers = """
PER2NPS01
PER2ARCGISVD01
""".splitlines()

servers = sorted([x for x in servers if len(x) > 0])

s = ", ".join(servers[:-1]) + f' and {servers[-1]}'
print(s)
print()

s = '(' + ",".join(f'\'{x}\'' for x in servers) +')'


conn = psycopg2.connect("dbname='CPM_VMs' user=postgres")

cur = conn.cursor()

sql = """
        SELECT vm, description, cpus, memory_gb FROM \"public\".\"DevTest\" 
        WHERE vm in {} order by vm asc
      """.format(s)

cur.execute(sql)

rex = re.compile(r'\d\d?\d?\.\d\d?\d?\.\d\d?\d?\.\d\d?\d?\s*\-\s*')

for row in cur.fetchall():
    vm, desc, cpu, gb = row
    s = f' - {vm}: {desc} ({cpu} x cpu, {gb}GB RAM)'
    s = rex.sub('', s)
    print(s)
conn.close()
