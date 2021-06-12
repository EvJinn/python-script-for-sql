# coding=UTF-8

import pyodbc

connStr = (
    r"DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};"
    r"DBQ=C:\Users\User\Desktop\AUTOSTOP.mdb;"
)
cnxn = pyodbc.connect(connStr)
cursor = cnxn.cursor()

cursor.execute("""
            SELECT COUNT(CAR.USERID) AS COUNT, USERS.NAME
            FROM CAR INNER JOIN USERS ON CAR.USERID = USERS.ID
            GROUP BY USERS.NAME
            HAVING COUNT(CAR.USERID) > 50
            ORDER BY COUNT(CAR.USERID) DESC;
            """)
rows = cursor.fetchall()
file = open("result.txt", "w+")
file.write("(COUNT, USERS.NAME)\n")

for row in rows:
    print(row)
    file.write(str(row))
    file.write('\n')

file.close()
cnxn.close()
