import pymysql
from xlrd import open_workbook,XL_CELL_TEXT
from datetime import datetime

# set file path
file = "C:\\Users\\DIEGO\\Downloads\\Lista-Procesos07-18.xls"

conn = pymysql.connect( host='localhost',
                        port=3306,
                        user='root',
                        passwd='root',
                        db='IA')

cur = conn.cursor()

num_rows = int(cur.execute("SELECT * FROM seace"))
s_id = 'seace_'+str(num_rows-1).zfill(16)
print(s_id)
cur.execute("SELECT * FROM seace where id_seace=%s",(s_id,))

last_time = ""
last_nom = ""
for row in cur:
    last_time = row[3]
    last_nom = row[1]
print(last_time)
if not last_time:
    last_time = '02/02/1995 00:00'

last_datetime = datetime.strptime(last_time, '%d/%m/%Y %H:%M')
print(last_datetime)

access = 0
# load demo.xlsx
book = open_workbook(file)
# activate demo.xlsx
sheet = book.sheet_by_index(0)
# get b3 cell value
col = [1,2,3,5,6,8,9]
nrow = int(sheet.nrows-1)
while nrow>0: #sheet.cell(nrow,0)
    s = []
    s.append('seace_'+str(num_rows).zfill(16))

    for i in col:
        b1=sheet.cell_value(nrow,int(i))
        s.append(b1)

    date_compare = datetime.strptime(s[2], '%d/%m/%Y %H:%M')
    if date_compare > last_datetime:
        access = 1
    nrow-=1
    if access:
        num_rows+=1
        cur.execute('''INSERT INTO seace (id_seace, nomenclatura, nombre_entidad, fecha, objeto_de_contratacion, descripcion, valor_referencial, moneda)
        VALUES (%s,%s,%s,%s,%s,%s,%s,%s)''',(s[0],s[3],s[1],s[2],s[4],s[5],s[6],s[7]))

    if last_datetime == date_compare:
        if last_nom == s[3]:
            access = 1


conn.commit()

cur.close()
conn.close()
