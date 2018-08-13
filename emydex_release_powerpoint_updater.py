import psycopg2
import textwrap
import win32com.client
import pprint


def pretty_print(o):
    pp = pprint.PrettyPrinter(indent=4)
    pp.pprint(o)


conn = psycopg2.connect("dbname='HarveyBeef' user=postgres")
cur = conn.cursor()

sql = '''
    SELECT
    "public".field_equip.ID,
    "public".field_equip."release",
    "public".field_equip.station,
    "public".field_equip.area,
    "public".field_equip.network_points,
    "public".field_equip.hardware,
    "public".field_equip.make_model_os 
    FROM
        "public".field_equip 
    WHERE
        "public".field_equip.release = 2
        AND (
        "public".field_equip.network_points ~* '(live|wireless|crossover|drops)' 
        OR "public".field_equip.hardware ~* 'scanner')
    ORDER BY "public".field_equip.release, "public".field_equip.station, "public".field_equip.id
    '''

print(textwrap.dedent(sql))

cur.execute(sql)

data = {}

for row in cur.fetchall():
    _id, release, station, area, network_points, hardware, make_model_os = row
    # print(release, station, area, network_points, hardware, make_model_os, sep='|')
    network_points = network_points.replace('Crossover Cable To Unit', 'Crossover cable').replace('Wireless Access Points', 'Wireless').replace('?','')
    s = f' - {network_points} link for {"" if not make_model_os else make_model_os.strip()} {hardware}'
    if station not in data:
        data[station] = []
    print(s)
    data[station].append(s)

pretty_print(data)

ppApp = win32com.client.gencache.EnsureDispatch('Powerpoint.Application')
ppApp.Visible = True 
deck = ppApp.Presentations.Open(r'C:\Users\rdapaz\Desktop\Emydex Data Collections Areas - ER2.pptm')
for slideIdx in range(2, 9):
    slide = deck.Slides(slideIdx)
    slide.Select()
    for shp in slide.Shapes:
        if slideIdx + 9 < 19:
            if 'Callout' in shp.Name and slideIdx + 9 in data:
                shp.TextFrame.TextRange.Text = "\n".join(data[slideIdx + 9])
        else:
            new_s = "\n".join(data[19]) + "\n".join(data[20]) + "\n".join(data[21])
            print('Got here')
            if 'Callout' in shp.Name:
                shp.TextFrame.TextRange.Text = new_s 