import yaml
import pprint


def pretty_printer(o):
    pp = pprint.PrettyPrinter(indent=4)
    pp.pprint(o)


with open(r'C:\Users\rdapaz\Documents\scripts\doctools\CPMProdServers.yaml', 'r') as fin:
    data = yaml.load(fin)

pretty_printer(data)
for tranche, server_data in data.items():
    for server_type, instances in server_data.items():
        if len(instances) > 1:
            s = ", ".join(instances[:-1]) + ' and ' +  instances[-1]
        else:
            s = instances[0]
        print(tranche, server_type, s, sep="|")

