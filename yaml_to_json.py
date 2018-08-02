import yaml
import json

with open(r'C:\Users\rdapaz\Documents\scripts\doctools\minutes.yaml', 'r') as fin:
        data = yaml.load(fin)

new_data = []
for p in data:
    new_data.append([
        p['Action'],
        '23/05/2018',
        p['Due'],
        p['Resp'],
        'In Progress'
        ])

with open(r'C:\Users\rdapaz\Documents\scripts\doctools\minutes.yaml', 'w') as fout:
    json.dump(new_data, fout, indent=True)