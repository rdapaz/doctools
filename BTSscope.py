#BTS scope

import re
import pprint


def pretty_printer(o):
    pp = pprint.PrettyPrinter(indent=4)
    pp.pprint(o)


data = """
WAP 01   -    In Lunch Room   -   Admin   -   Internal   -   60m   -  

WAP 02   -   In By Products Office   -   Admin   -   Internal   -   90m/100m   -   

WAP 03   -   Outside By Products Office   -   Admin   -   External   -   90m/100m  -   

WAP 04   -   In Maintenance Office   -   Maint   -   Internal   -   10m  -    

WAP 05   -   Outside Maintenance (halfway)   -   Maint   -   External   -   40/50m  -    

WAP 06   -   In Stores Office   -   Maint   -   Internal   -   Existing  -   

WAP 07   -   In Cardboard Store   -   Hay Shed   -   Internal   -   Existing  -   

WAP 08   -   in Hay Shed   -   Hay Shed   -   Internal   -   30/40m   -   Enclosure required  

WAP 09   -   Corner of Stock Shed   -   Hay Shed   -   External   -   Existing  - 

WAP 10   -   ByProducts SwRm/Plant Wall   -   ByProducts   -   Internal   -   10m   -   Enclosure required

WAP 11   -   Tripe Room Centralised   -   Slaughter Floor   -   90m   -   Internal   -   enclosure required 

WAP 12   -   Slaughter Floor Hallway   -   Slaughter Floor   -   70/80m   -   Internal   -   enclosure required

WAP 13   -   Outside QC Office   -   QC   -   30m   -   External  -      

WAP 14   -   Boning Room   -   QC/Hot Box   -   Existing   -   Internal   -   enclosure required if existing in poor condition 

WAP 15   -   Slaughter Floor Office   -   Slaughter MCC   -   Existing   -   Internal   -   enclosure required if existing in poor condition

WAP 17   -   Above Cattle Entrance   -   Slaughter MCC   -   100m+   -   External  -    

WAP 18   -   Retail Area/Office   -   Cold Stores   -   90/100m   -   Internal   -   enclosure required if existing is in poor condition

WAP 19   -   Chiller Store Wall   -   Cold Stores   -   Existing   -   Internal   -   enclosure required if existing is in poor condition

WAP 20   -   Cold Store Wall   -   Cold Stores   -   30m   -   Internal   -   enclosure required if existing is in poor condition

WAP 21   -   Cold Stores Office Wall   -   Cold Stores   -   10m   -   Internal   -   enclosure required if existing is in poor condition

WAP 22   -   Chiller Store Wall   -   Cold Stores   -   Existing   -   Internal   -   enclosure required if existing is in poor condition

WAP 23   -   Freezer Store Outside Corner   -   Cold Stores -  100m+   -   External   -   

WAP 24   -   Freezer Store Outside Corner   -   Cold Stores - 100m+   -   External   -    
""".splitlines()

rex = re.compile(r'\s+\-\s+')
data = [rex.split(x) for x in data if len(x) > 0]


pretty_printer(data)

for line in data:
    bts_naming, location, to_rack, internal_external, rough_run_length, parts_required = line
    if re.search(r'existing', "|".join(line), re.IGNORECASE):
            print(bts_naming, location, to_rack, internal_external, rough_run_length, parts_required, sep='|')
