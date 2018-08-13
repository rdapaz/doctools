import re
import pprint

def pretty_printer(o):
    pp = pprint.PrettyPrinter(indent=4)
    pp.pprint(o)


data = """
Release
Station
Area
Emydex Form
Hardware
Network Points
Existing or New
Make/Model/OS
Screen size
Screen Resolution
String
Sending Weight?
Emydex Polling or Receiving Contineous"
Max Weight of scales
Weight Increment
With Alibi Yes/No
Emydex Install Date
H/W Install  & Connected to Network Date
Date Tested with Emydex
Hardware Cost
PC Specs

""".splitlines()

def replace_spaces(sText):
    rex = re.compile('\s+')
    sText = re.sub(r'\s+\Z', '', sText)
    sText = re.sub(r'\A\s+', '', sText)
    sText = rex.sub(repl='_', string=sText)
    sText = sText.replace('(', '').replace(')', '').replace('-', '').replace('/','_').replace('__','_').replace('#', 'no').replace('.','').replace('"','').replace('\'','')
    sText = re.sub(r'[_]+', '_', sText)
    sText = sText.lower()
    return sText

data = [replace_spaces(x) for x in data if len(x) > 0]

for item in data:
    print(item)