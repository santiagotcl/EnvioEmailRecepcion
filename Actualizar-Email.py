#coding:utf-8
import requests
from bs4 import BeautifulSoup
from datetime import date, datetime,timedelta
import time
import pyodbc
from os import getcwd

DRIVER_NAME = "Microsoft Access Driver (*.mdb, *.accdb)"
#DB_PATH ="C:/Users/TCL A3 4500/Desktop/Programa pablo/PanelControlFR_db.accdb"
DB_PATH ="C:/Users/Mantenimiento/Desktop/yella/PanelControlFR_db.accdb"

print("Ingrese legajo del sistema de mantenimiento")
legajo=input()
print("Ingrese Email de la persona:")
email=input()

try:
    conn = pyodbc.connect("Driver={%s};DBQ=%s;" % (DRIVER_NAME, DB_PATH))
    cursor = conn.cursor()
    cursor.execute("UPDATE Personal SET Email = ? WHERE IdPersonal = ?",(email,legajo))
    cursor.commit()
    conn.close()
    print("Actualizado con Exito!")
except:
    print("Error de conexion con la BDD")












