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
DRIVER_NAME = "Microsoft Access Driver (*.mdb, *.accdb)"
#DB_PATH ="C:/Users/TCL A3 4500/Desktop/Programa pablo/PanelControlFR_db.accdb"
#DB_PATH ="C:/Users/Mantenimiento/Desktop/yella/PanelControlFR_db.accdb"
DB_PATH ="M:/PanelControlFR_db.accdb"
ListaArticulos=[]

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


while(1):
    now = datetime.now()
    hora = now.strftime('%H')
    minutos = now.strftime('%M')

    if(minutos == "00"):
        try:
            conn = pyodbc.connect("Driver={%s};DBQ=%s;" % (DRIVER_NAME, DB_PATH))
            cursor = conn.cursor()
            cursor.execute("SELECT idrecepcion from recepciones ORDER BY Ultimo desc;")
            UltimoEnviado=cursor.fetchall()
            UltimoEnviado=list(UltimoEnviado[0])
            UltimoEnviado=list(UltimoEnviado)
            UltimoEnviado=UltimoEnviado[0]
            cursor.execute("SELECT idrecepcion from recepciones ORDER BY idrecepcion desc;")
            Ultimo=cursor.fetchall()
            Ultimo=list(Ultimo[0])
            Ultimo=list(Ultimo)
            Ultimo=Ultimo[0]
            cursor.close()
            conn.close()
            if(Ultimo != UltimoEnviado):
                conn = pyodbc.connect("Driver={%s};DBQ=%s;" % (DRIVER_NAME, DB_PATH))
                cursor = conn.cursor()
            #cursor.execute("SELECT Solicita,IdDetalleSalidaMat,Descripcion from SalidaMaterialDet WHERE IdDetalleSalidaMat IN (SELECT pedido from (recepciones INNER JOIN SC_pendiente ON recepciones.N_ORDEN_CO = SC_pendiente.OC) WHERE FECHA_MOV = ?)  ORDER BY Solicita Desc;",(fecha))
                cursor.execute("SELECT Solicita,IdDetalleSalidaMat,Descripcion from SalidaMaterialDet WHERE IdDetalleSalidaMat IN (SELECT pedido from (recepciones INNER JOIN SC_pendiente ON recepciones.N_ORDEN_CO = SC_pendiente.OC) WHERE idrecepcion BETWEEN ? and ?)  ORDER BY Solicita Desc;",((UltimoEnviado+1),Ultimo))
            #cursor.execute("SELECT Solicita,IdDetalleSalidaMat,Descripcion from SalidaMaterialDet WHERE IdDetalleSalidaMat IN (SELECT pedido from (recepciones INNER JOIN SC_pendiente ON recepciones.N_ORDEN_CO = SC_pendiente.OC) WHERE idrecepcion BETWEEN 9193 and 9211)  ORDER BY Solicita Desc;")
                PedidosNuevos=cursor.fetchall()
                print(PedidosNuevos)
                dato1=list(PedidosNuevos)
                print(len(dato1))
                if(len(dato1)>1):
                    while(len(dato1)>0):  
                        dato=dato1[0]
                        id=dato[0]
                        print(id)
                        objeto=Crear_Lista(id)
                        Enviar_Email(objeto)
                        print(len(objeto))
                cursor.execute("UPDATE recepciones SET Ultimo = ? WHERE idrecepcion = ?",(Ultimo,Ultimo))
                cursor.commit()
                cursor.close()
                conn.close()            
        except:
            print("Error de conexion con la BDD")
    time.sleep(60)







