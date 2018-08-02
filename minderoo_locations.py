import re

data = """
>   Cabinet 2 (Maintenance) to Cabinet 5 (Hay Shed MCC)
>   Cabinet 4 (Server Room/Admin) to Cabinet 2 (Maintenance)
>   Cabinet 5 (Hay Shed MCC) to Cabinet 6 (By-product MCC)
>   Cabinet 6 (By-product MCC) to Cabinet 9 (Engine room)
>   Cabinet 8 (QA Building) to Cabinet 4 (Server Room/Admin)
>   Cabinet 9 (Engine room) to Cabinet 11 (Retail Area MCC)
>   Cabinet 11 (Retail Area MCC) to Cabinet 8 (QA Building)
"""

entries = list(set((x for x in re.findall(r'\(.*?\)', data))))
for entry in sorted(entries):
    entry = entry.replace('(','').replace(')','')
    print(entry)