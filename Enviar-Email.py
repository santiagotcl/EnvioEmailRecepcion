#coding:utf-8
import requests
from bs4 import BeautifulSoup
from datetime import date, datetime,timedelta
import time
import pyodbc
from os import getcwd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
temp=""
temp2=""
temp3=""
temp4=""
DRIVER_NAME = "Microsoft Access Driver (*.mdb, *.accdb)"
#DB_PATH ="C:/Users/TCL A3 4500/Desktop/Programa pablo/PanelControlFR_db.accdb"
DB_PATH ="C:/Users/Mantenimiento/Desktop/yella/PanelControlFR_db.accdb"
FI="14/09/2020"
FF="18/09/2020"
linea="LB"
paradas=0
ListaArticulos=[]

def Filtra_Linea(dato):
    i=len(dato)
    j=0
    dato1 = []
    while(i > 0):
        fila=list(dato[j])
        if(fila[1] == linea):
            dato1.append(dato[j])
        i=i-1
        j=j+1
    return dato1

def Suma_Minutos(dato):
    global paradas
    i=len(dato)
    i=i-1
    j=0
    suma=0
    Hora_Parada=0
    Hora_Arranque=0
    while(i != j):
        temp=dato[j]
        if(temp[2] == 'A' or temp[2] == 'AD'):
            Hora_Parada=temp[4]
            while(temp[2] != 'P' or temp[2] != 'PD'):
                if(j!=i):
                    j=j+1
                    temp=dato[j]
                else:
                    break
                break
            Hora_Arranque = temp[4]
            suma=suma+Resta_Tiempo(Hora_Parada,Hora_Arranque)
            paradas=paradas+1 #cuento cantidad de paradas
        if(j!=i):
            j=j+1
        else:
            break
    print("La cantidad de pardas fue: " + str(paradas))
    return suma

def Resta_Tiempo(HP,HA):
    tiempo = timedelta(
    days=0,
    seconds=0,
    microseconds=0,
    milliseconds=0,
    minutes=0,
    hours=0,
    weeks=0
    )
    suma=0
    tiempo=HA-HP
    suma=int(tiempo.seconds/60)
    return suma

def Crear_Lista(id):
    global dato1
    articulos=list()
    i=0
    maximo=len(dato1)
    #articulon=articulo(temp[0],temp[1],temp[2])
    #articulos.append(articulon)
    while(i < maximo):
        temp=dato1[i]
        if(temp[0]==id):
            dato1.pop(i)
            maximo=len(dato1)
            articulon=articulo(temp[0],temp[1],temp[2])
            articulos.append(articulon)
        else:
            i=i+1
    return (articulos)

def Enviar_Email(contenido1):
    # Iniciamos los parámetros del script
    contenido=contenido1
    destinatarios=list()
    conn = pyodbc.connect("Driver={%s};DBQ=%s;" % (DRIVER_NAME, DB_PATH))
    cursor = conn.cursor()
    print(contenido[0].id)
    cursor.execute("SELECT Email,Persona from Personal WHERE IdPersonal = ?;",(contenido[0].id))
    email=cursor.fetchall()
    email=list(email)
    email=list(email[0])
    email1=email[0]
    #email1="queti"
    print(email)
    cursor.close()
    conn.close()
    remitente = 'dmfideosrivoli@gmail.com'
    destinatarios.append(email1)
    asunto = 'Ingreso de Material'
    cuerpo = "Buenos dias "+email[1]+",se informa desde Deposito de materiales que llegaron los siguientes repuestos a su nombre..."
    i=len(contenido)
    while(i > 0):
        cuerpo=cuerpo+"\n"+"*"+contenido[i-1].Descripcion
        i=i-1
        
    try:
        # Creamos el objeto mensaje
        mensaje = MIMEMultipart()
        # Establecemos los atributos del mensaje
        mensaje['From'] = remitente
        mensaje['To'] = ", ".join(destinatarios)
        mensaje['Subject'] = asunto
        # Agregamos el cuerpo del mensaje como objeto MIME de tipo texto
        mensaje.attach(MIMEText(cuerpo, 'plain'))

        sesion_smtp = smtplib.SMTP('smtp.gmail.com', 587)
        # Ciframos la conexión
        sesion_smtp.starttls()
        # Iniciamos sesión en el servidor
        sesion_smtp.login('dmfideosrivoli@gmail.com','juanymelian')
        # Convertimos el objeto mensaje a texto
        texto = mensaje.as_string()
        # Enviamos el mensaje
        sesion_smtp.sendmail(remitente, destinatarios, texto)
        # Cerramos la conexión
        sesion_smtp.quit()
    except:
        remitente = 'dmfideosrivoli@gmail.com'
        destinatarios.pop(0)
        destinatarios=['santiagocuozzo2@gmail.com'] #email del jefe de deposito
        asunto = 'Fallo al enviar E-mail'
        cuerpo = "Error al enviar mensaje, falla de conexion, E-mail inexistente o mal escrito, llegaron para la parsona "+email[1]+" los siguientes materiales:"
        i=len(contenido)
        while(i > 0):
            cuerpo=cuerpo+"\n"+"*"+contenido[i-1].Descripcion
            i=i-1
                # Creamos el objeto mensaje
        mensaje = MIMEMultipart()
        # Establecemos los atributos del mensaje
        mensaje['From'] = remitente
        mensaje['To'] = ", ".join(destinatarios)
        mensaje['Subject'] = asunto
        # Agregamos el cuerpo del mensaje como objeto MIME de tipo texto
        mensaje.attach(MIMEText(cuerpo, 'plain'))
        # Creamos la conexión con el servidor
        sesion_smtp = smtplib.SMTP('smtp.gmail.com', 587)
        # Ciframos la conexión
        sesion_smtp.starttls()
        # Iniciamos sesión en el servidor
        sesion_smtp.login('dmfideosrivoli@gmail.com','juanymelian')
        # Convertimos el objeto mensaje a texto
        texto = mensaje.as_string()
        # Enviamos el mensaje
        sesion_smtp.sendmail(remitente, destinatarios, texto)
        # Cerramos la conexión
        sesion_smtp.quit()



    return

class articulo:
    def __init__(self, id, solicita, Descripcion):
        self.id = id
        self.solicita = solicita
        self.Descripcion = Descripcion



now = datetime.now()
fecha = now.strftime('%d/%m/%Y')
try:
    conn = pyodbc.connect("Driver={%s};DBQ=%s;" % (DRIVER_NAME, DB_PATH))
    cursor = conn.cursor()
    fecha="03/02/2021"
    cursor.execute("SELECT Solicita,IdDetalleSalidaMat,Descripcion from SalidaMaterialDet WHERE IdDetalleSalidaMat IN (SELECT pedido from (recepciones INNER JOIN SC_pendiente ON recepciones.N_ORDEN_CO = SC_pendiente.OC) WHERE FECHA_MOV = ?)  ORDER BY Solicita Desc;",(fecha))
    PedidosNuevos=cursor.fetchall()
    cursor.close()
    conn.close()
    dato1=list(PedidosNuevos)
    while(len(dato1)>0):  
        dato=dato1[0]
        id=dato[0]
        print(id)
        objeto=Crear_Lista(id)
        #if(id==101):
        Enviar_Email(objeto)
        print(len(objeto))
except:
    print("Error de conexion con la BDD")












