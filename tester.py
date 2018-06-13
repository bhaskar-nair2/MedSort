import sqlite3 as sql

con=sql.connect('medSort')
cur=con.cursor()
cur.execute('select * from paSearchList')
rcList=cur.fetchall()

for i in rcList:
    print(i[0])