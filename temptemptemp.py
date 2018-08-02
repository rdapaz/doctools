
import pprint


def pretty_printer(o):
    pp = pprint.PrettyPrinter(indent=4)
    pp.pprint(o)

data = """
Gary Wainwright
Neil Stocklmayer
Gavin O'Hara
Clare Pearson
""".splitlines()

data = [[x, ".".join(x.replace('\'','').split()) + '@datacom.com.au'] for x in data if len(x) > 0]

for name, email in data:
    print(f'{name} => {email}')