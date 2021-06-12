# coding=UTF-8

import pyodbc

# подключение к базе
connStr = (
    r"DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};"
    r"DBQ=C:\Users\User\Desktop\AUTOSTOP.mdb;" # путь к файлу с базой. User заменить на имя своего пользователя в винде
)
cnxn = pyodbc.connect(connStr)
cursor = cnxn.cursor()

# запрос
cursor.execute("""
            SELECT COUNT(CAR.USERID) AS COUNT, USERS.NAME
            FROM CAR INNER JOIN USERS ON CAR.USERID = USERS.ID
            GROUP BY USERS.NAME
            HAVING COUNT(CAR.USERID) > 50
            ORDER BY COUNT(CAR.USERID) DESC;
            """)
rows = cursor.fetchall()

# открываем файл для записи
file = open("result.txt", "w+")
file.write("(COUNT, USERS.NAME)\n")

# выводим результат запроса в консоль и параллельно записываем его в файл
for row in rows:
    print(row)
    file.write(str(row))
    file.write('\n')

# закрываем файл и подключение к базе
file.close()
cnxn.close()
