import re

rex = re.compile(r'\d\d?\d?\.\d\d?\d?\.\d\d?\d?\.\d\d?\d?\s\-')

data = """
PER2CSH01|172.22.4.97 - Citrix Session Host|t
PER2NPS01|172.22.1.21 -MS RADIUS - Frontend|f
PER1CBAR01|172.22.4.112 - CBAR Web app Prod (Yun (Claud) Li)|t
PER2ARCGISVD01|172.22.4.55 - ArcGIS Virtual Desktop|f
PER2CDC01|172.22.4.95 - Citrix Delivery Controller|t
PER2RRAS01|172.22.4.80 - Azure RRAS server|t
""".splitlines()


data = [x.split('|') for x in data if len(x) > 0]
data = sorted(data, key=lambda x: x[0])

arr = []
for row in data:
    vm, desc, successful = row
    desc = rex.sub('', desc)
    desc = desc.strip()
    desc = desc.replace('(Yun (Claud) Li)', '')
    if successful == 't':
        arr.append(f'{vm} ({desc})')

print(', '.join(arr[:-1]) + ' and ' + arr[-1])


