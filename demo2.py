import xlrd
import MySQLdb
# Open the workbook and define the worksheet
book = xlrd.open_workbook("E:\demo20181103.xls")
# sheet = book.sheet_names()
sheet = book.sheet_by_index(0)

#建立一个MySQL连接
database = MySQLdb.connect (host="localhost", user = "root", passwd = "root", db = "test")

# 获得游标对象, 用于逐行遍历数据库数据
cursor = database.cursor()

# 创建插入SQL语句
query = """INSERT INTO rsps_jhyt_box_info (id, rfid)  VALUES (%s, %s)"""

# 创建一个for循环迭代读取xls文件每行数据的, 从第二行开始是要跳过标题
for r in range(1, sheet.nrows):
    id   = sheet.cell(r,0).value
    rfid = sheet.cell(r,1).value

    values = (id, rfid)

    # 执行sql语句
    cursor.execute(query, values)

# 关闭游标
cursor.close()

# 提交
database.commit()

# 关闭数据库连接
database.close()

# 打印结果
print ("")
print ("Done! ")
print ("")
columns = str(sheet.ncols)
rows = str(sheet.nrows)
print(columns+"   "+rows)