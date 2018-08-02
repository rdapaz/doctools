
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



data = """
PER1UFS01|Canon UniFlow - Frontend|2|4|t2.small
PER2AADS01|MS Azure ActiveDirectory Sync - Prod|2|4|m4.large
PER2ICA01|Subordinate Issuing Certificate Authority |2|4|m4.xlarge
PER1TFC02|MS Exchange - CAS/HUB Frontend|2|8|t2.large
PER1TFS02|Team Foundation Server|1|4|t2.small
PER2RCA01|Root Certificate Authority |1|2|t2.medium
""".splitlines()

data = [x.split('|') for x in data if len(x) > 0]
data = sorted(data, key=lambda x: x[0])

for vm, desc, cpu, gb, inst_type in data:
    print(f'{vm}: {desc} ({cpu} x cpu, {gb}GB RAM, {inst_type})')
