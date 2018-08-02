import re

data = """
1.5d|1.75d||1d|11d|32.38d|9.06d|6.3d|0.42d
5.08d|14.25d|14d|||0.75d|||3.77d
|||0.83d|0.17d|8.16d|3.11d|6.76d|
|5.25d|||||||0.38d
1.5d|0.75d||||0.5d|||0.13d
1.5d||||5.17d|3.17d|||
|||||3d|||
""".splitlines()

rex = re.compile(r'(?<=\d)d', re.DOTALL)

for row in (rex.sub('', row) for row in data if len(row)>0):
    print(row)