
import psycopg2
import pprint


def pretty_printer(o):
    pp = pprint.PrettyPrinter(indent=4)
    pp.pprint(o)


def try_to_int(int_p):
    print(int_p)
    try:
        return int(int_p)
    except:
        return int_p


data = """
PER2SAPE01DEV|8|32|m4.2xlarge
PER2SAPE01TST|4|16|m4.xlarge
PER2SAPG01DEV|4|16|m4.xlarge
PER2SAPG01TST|4|16|m4.xlarge
PER2SAPM01DEV|2|8 |m4.large
PER2SAPM01TST|2|8 |m4.large
PER2SAPP01DEV|4|16|m4.xlarge
""".splitlines()

data = [x.split('|') for x in data if len(x) > 0]

new_data= []
for vm, cpu, gb, inst_type in data:
    cpu = int(cpu)
    gb = int(gb)
    new_data.append([cpu, gb, inst_type, vm])

conn = psycopg2.connect("dbname='CPM_VMs' user=postgres")
cur = conn.cursor()

for row_data in new_data:
    sql = '''
        update "DevTest" set 
        cpus = %s, 
        memory_gb = %s,
        inst_type = %s
        where vm = %s
    '''

    cur.execute(sql, row_data)
    conn.commit()

conn.close()