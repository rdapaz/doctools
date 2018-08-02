import psycopg2
import datetime
import json
import sys
import os



current_path = os.path.dirname(sys.argv[0])
os.chdir(current_path)
JSONDIR = r'.'


conn = psycopg2.connect("dbname='RAID' user=postgres")

cur = conn.cursor()

sql = """
    SELECT
        r.date_raised,
        r.raised_by,
        r.description_of_risk,
        r.description_of_impact,
        r.priority_rating,
        r.preventative_actions,
        r.contingency_actions,
        r.preventative_action_owner,
        r.contingency_action_owner
    FROM
        "Risks" r 
    WHERE
        r.priority_rating < 10 
        AND r.preventative_actions !~* 'issue' 
    ORDER BY
        r.priority_rating,
        r.riskid ASC
"""

cur.execute(sql)

arr = []
for row in cur.fetchall():
    raised, by, risk_desc, risk_impact, priority, prev, cont, prev_owner, cont_owner = row
    prev_owner = prev_owner if prev_owner else ''
    cont_owner = cont_owner if cont_owner else ''
    owners = "\n".join(list(set((prev_owner, cont_owner))))
    arr.append([
                    raised.strftime('%d/%m/%Y'),
                    by,
                    f'Description:\n{risk_desc}\nImpact:\n{risk_impact}',
                    priority,
                    f'Controls:\n{prev}\nContingencies:\n{cont}',
                    owners
                ])
    print(arr)

with open(os.path.join(JSONDIR, 'risks.json'), 'w') as fout:
    json.dump(arr, fout, indent=True)

