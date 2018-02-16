# Skrypt pobierający dane z db, zapisujący do csv i json i wysyłający json na ftp.
# coding=utf-8

import pyodbc
import json
import os
import csv
import ftplib
import time
import datetime
import sys

# Pętla wykonująca się nieskończenie długo
while True:
	try: 
		os.getcwd()
		os.chdir('C:\\Users\\Skrypt\\Desktop\\python')
		
# Konfiguracja połączenia z serwerem MS SQL
		connection = pyodbc.connect('DRIVER={SQL Server};SERVER=XX.X.X.XXXX;DATABASE=XXXX;UID=XXXX;PWD=XXXXXXXX')
		cursor = connection.cursor()
		cursor.execute(
		""" 
		SELECT *
		FROM table
		""")

# Utworzenie pliku CSV i zapisanie wyniku zapytania
		with open("punkty_yam.csv", "w", newline='') as csv_file:
			csv_writer=csv.writer(csv_file)
			csv_writer.writerows(cursor)

		csvfile = open("punkty.csv", "r")
		
		jsonfile = open("punkty.json", "w")

		fieldnames = ("Column1", "Column2", "Column3", "Column4")

		reader = csv.DictReader(csvfile, fieldnames)
		
# Zapisanie pliku CSV do JSON (można bezpośrednio, w razie jakby były 2 formaty potrzebne)
		out = json.dumps([row for row in reader])
		jsonfile.write(out)
	
# Utworzenie połączenia z FTP i wysłanie pliku JSON
		session = ftplib.FTP('ftp.xxxxx.pl','login','password')
		file = open('punkty.json', 'rb')
		session.storbinary('stor punkty.json', file)
		file.close()
		session.quit()
		
# W przypadku poprawnego wykonania - data i godzina
		now = datetime.datetime.now()	
		print(now)
		
# Ponowne wykonanie pętli za 15 minut
		time.sleep(900)
		
# W przypadku błędu komunikat i spróbuj ponownie za 5 minut
	except:
		print("Wystąpił błąd")
		time.sleep(300)
