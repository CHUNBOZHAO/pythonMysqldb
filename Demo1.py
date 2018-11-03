import xlwt
import MySQLdb
import datetime

host = '192.168.0.111'
user = 'rspsadmin'
pwd = 'rsps!@#Test'
db = 'rspsdb'
sql = 'select * from rsps_jhyt_box_info'
sheet_name = 'user'
out_path = 'E:\demo'+datetime.datetime.now().strftime('%Y%m%d')+'.xls'

#conn = MySQLdb.connect(host,user,pwd,db,"utf-8")
conn = MySQLdb.connect(host=host , user=user, passwd=pwd , db=db , charset="utf8")

cursor = conn.cursor()
count = cursor.execute(sql)
print(count)

cursor.scroll(0)
results = cursor.fetchall()
fields = cursor.description
workbook = xlwt.Workbook()
sheet = workbook.add_sheet(sheet_name)

for field in range(0,len(fields)):
    sheet.write(0,field,fields[field][0])

row = 1
col = 0
for row in range(1,len(results)+1):
    for col in range(0,len(fields)):
        sheet.write(row,col,u'%s'%results[row-1][col])

workbook.save(out_path)