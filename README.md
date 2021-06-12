# Встраивание запросов SQL в программу-скрипт
_Модифицировать скрипт, приведённый в материалах к заданию, для реализации произвольного, по вашему выбору, SQL-запроса на выборку, содержащего агрегатные функции. Скрипт должен породить текстовый файл с результатом SQL-запроса, например html-файл или другого формата._
------------
Предварительные ласки:
* Установить **драйвер движка базы данных Access**. </br>
Скачать можно отсюда: https://www.microsoft.com/ru-RU/download/details.aspx?id=54920</br>
Для **32**-битной версии Access ставим **32**-битную версию драйвера, для **64**-битной версии Access ставим **64**-битный драйвер соответственно.</br>
Версию Access смотрим здесь:</br>
[![](https://sun9-57.userapi.com/impg/sVr0SaxQeJ-DYPu2RPk97Np0GFnYk32jVQsCEg/8tXdQ05h40s.jpg?size=1440x785&quality=96&sign=b5f94eafb7f5a6ca81a7599ba9bf3799&type=album)](https://sun9-57.userapi.com/impg/sVr0SaxQeJ-DYPu2RPk97Np0GFnYk32jVQsCEg/8tXdQ05h40s.jpg?size=1440x785&quality=96&sign=b5f94eafb7f5a6ca81a7599ba9bf3799&type=album)</br>
[![](https://sun9-6.userapi.com/impg/3IdhvzJKtMXC_2khS7dJufAwW6bReQr7pQCEEw/S0ZxvZ2uQek.jpg?size=1271x798&quality=96&sign=8f072a7bf10f91bf119af59150e97639&type=album)](https://sun9-6.userapi.com/impg/3IdhvzJKtMXC_2khS7dJufAwW6bReQr7pQCEEw/S0ZxvZ2uQek.jpg?size=1271x798&quality=96&sign=8f072a7bf10f91bf119af59150e97639&type=album)</br>
Если после запуска установочного файла выдало такую ошибку:</br>
[![](https://sun9-59.userapi.com/impg/ucL9bVT10ayM92hXDA7CUp3bUz5wgbEetk5K3w/xp4Pq89fc8o.jpg?size=502x258&quality=96&sign=50b388852087a326dc8549f8a67513a7&type=album)](https://sun9-59.userapi.com/impg/ucL9bVT10ayM92hXDA7CUp3bUz5wgbEetk5K3w/xp4Pq89fc8o.jpg?size=502x258&quality=96&sign=50b388852087a326dc8549f8a67513a7&type=album)</br>
то устанавливать придётся через командную строку с помощью следующих команд</br>
`cd Desktop - если скачали установочный файл на рабочий стол`</br>
`accessdatabaseengine.exe /quiet /norestart - для 32-битной версии`</br>
`accessdatabaseengine_X64.exe /quiet /norestart - для 64-битной версии`
* Установить **Python 3.9** **32х** или **64х** в зависимости от **разрядности** вашего Access. </br>
Скачать можно отсюда:  https://www.python.org/downloads/release/python-395/
* После установки Python установить модуль **pyodbc** на сам компьютер, чтобы можно было использовать скрипт вне среды разработки.
* Для этого открываем командную строку и вводим следующие команды: </br>
`cd C:\Users\User\AppData\Local\Programs\Python\Python39-32\Scripts (вместо User подставить имя своего пользователя в винде)` </br>
`pip install pyodbc` </br>
Дождаться завершения установки.
* Скачать и установить PyCharm, чтобы можно было редактировать проект **(внимание он на Английском)** </br>
Скачать можно отсюда: https://www.jetbrains.com/ru-ru/pycharm/download/#section=windows (рекомендую Community Edition, поскольку она бесплатная)
------------
По самому скрипту. Он очень простой, генерирует *.txt* файл с результатом запроса **в папке с проектом** и дополнительно выводит этот же результат в консоль. Разобраться в коде проще простого, так как инфы по pyodbc и работе с файлами в Python в гугле валом. Основная проблема была в том, чтобы нормально подцепиться к Access. </br>
Для демонстрации выполнять скрипт через командную строку перейдя командой `cd` в папку с *.py* файлом и введя его имя в командой строке.
