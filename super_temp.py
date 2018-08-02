import re
import pprint


def pretty_printer(o):
    pp = pprint.PrettyPrinter(indent=4)
    pp.pprint(o)


packages = {
    'Package 1': 'Telecommunications Services Changes',
    'Package 2': 'Telephony Service Relocation',
    'Package 3': 'Relocation to Malaga DC',
    'Package 4': 'Base AWS Production Build',
    'Package 5': 'Data Protection',
    'Package 6': 'Dev & Test Build',
    'Package 7': 'Disaster Recovery',
    'Package 8': 'Equipment Destined for Site Decommissioning)',
    'Package 9': 'Equipment Destined for Decommissioning & Disposal' 
}
# """.splitlines()


# rex = re.compile(r'(Package \d) \((.*)\)')

# packages = {}
# for line in data:
#     m = rex.search(line)
#     if m:
#         package = m.group(1)
#         print(package)
#         package_name = m.group(2)
#         print(package_name)
#     packages[package] = package_name


pretty_printer(packages)