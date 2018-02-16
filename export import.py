# Fragment kodu pobierającą dane z bazy danych MySQL i importujący dane do lokalnej bazy SQLite 
# w celu integracji danych z kilku źródeł (różne bazy danych np. MySQL, MS SQL) na potrzeby raportowania

#!/usr/bin/env python
# -*- coding: utf-8 -*-

import pymysql
import sqlite3
import csv
import os
import sys

# Podaj ścieżkę
print(os.getcwd())

# Konfiguracja połączenia z MySQL
db = pymysql.connect(host=" xxxx ",
                     user=" xxxx ",
                     password=" xxxx ",
                     db="c0_cms",
                     charset='utf8')


#############################################################################################
################################### Table offer #############################################
#############################################################################################

# Wykonanie zapytania oraz zapisanie wyniku                     
cur = db.cursor()    
cur.execute("SELECT * FROM `offer`")
result=cur.fetchall()


# Połączenie z bazą danych SQLITE
conn = sqlite3.connect('calc.db')

# Usunięcie danych nieaktualnych z tabeli oraz zaimportownanie nowych z MySQL
conn.execute("DELETE FROM offer")
conn.executemany('INSERT INTO offer VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)', result)
conn.commit()