#!/usr/bin/env python
# -*- coding: utf-8 -*-
## Autor: Ruben Dario Coria

from genericpath import exists
import glob
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.mime.text import MIMEText
import smtplib
import time
import pandas as pd
import logging
import os


########################Subprograma para enviar el mail######################################
########################Para que funcione 

def enviarmail(paraquien):
    # create message object instance
    msg = MIMEMultipart()


    f = open ('mensaje.txt','r') #abre el archivo mensaje.txt
    mensaje = f.read()#lo lee y se lo asigna a la variable
    print(mensaje)#muestra en pantalla
    recipients = [paraquien, 'mail@domain.com']#Mails adonde se van a enviar, con CC

    text = mensaje

    part1=MIMEText(text,'plain')

    # setup the parameters of the message
    password = "your_mail_password"
    msg['From'] =  "Description <your_mail@domain.com>"
    msg['To'] = ", ".join(recipients)##Hace un join de la variable
    msg['Subject'] = "Subject"

    # attach image to message body
    attachment_path = archivopdf # replace with path of your PDF file
    with open(attachment_path, "rb") as f:
        attach = MIMEApplication(f.read(),_subtype="pdf")
        attach.add_header('Content-Disposition','attachment',filename=str(solonombrepdf))
    msg.attach(attach)
    msg.attach(part1)

    #create server
    server = smtplib.SMTP('smtp.gmail.com: 587')

    server.starttls()

    # Login Credentials for sending the mail
    server.login('your_mail@domain.com', password)


    # send the message via the server.
    server.sendmail(msg['From'], recipients, msg.as_string())

    server.quit()#cerramos la conexion con el server
    f.close()#cerramos el archivo txt mensaje.txt
    return
#################Aca termina el subprograma para enviar el mail###############################

path = 'readfile.xlsx' ##Nombre del archivo excel

df = pd.read_excel(path)##Leer tabla de excel


# Creación del logging con el archivo llamado logs_info.log.
logging.basicConfig(
format = '%(asctime)-5s %(name)-15s %(levelname)-8s %(message)s',
level  = logging.INFO,      # Nivel de evento INFO
filename = 'logs_info.log', # archivo en donde se escriben los logs
filemode = "a"              # a ("append"), si el archivo de logs ya existe, se abre y añaden nuevas lineas.
)
logging.info('Se inicio Send_mail')


for i in range(1,151):##For que va del valor 1 al 151
    legajo=(str(i).zfill(2))##convierte en string el numero de legajo
    name = (glob.glob('pdf_files/*LEG_'+legajo+'.pdf'))##Busca si existe un archivo con la coincidencia mencionada
    if name == []:##Si no encuentra el archivo y si es distinto a 'vacio' en la tabla va a guardar el resultado en el log
        if str((df.iloc[i]['MAIL'])) != 'vacio':
            logging.info ("No se encontro el archivo pdf para el legajo "+ legajo+' '+str(df.iloc[i]['MAIL']))
    else:##Si no, envia el mail a la persona indicada con el archivo pdf adjunto.
        solonombrepdf=(os.path.basename(name[0]))# variable para solo para obtener solo el nombre del archivo.
        archivopdf=name[0]##Si existe el archivo, lo digo que envie el path completo en la posicion 0 del arreglo.
        enviarmail (str(df.iloc[i]['MAIL']))
        logging.info ('Se envio el archivo '+os.path.basename(name[0]) +' para el lejago ' +legajo+' al mail '+str(df.iloc[i]['MAIL']))
        print('Se envio el archivo '+os.path.basename(name[0]) +' para el lejago ' +legajo+' al mail '+str(df.iloc[i]['MAIL']))
        time.sleep(10)






