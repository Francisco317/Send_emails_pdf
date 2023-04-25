
#IMPORTACION DE LAS PAQUETERIAS 
from email.message import EmailMessage
import ssl
import smtplib
import csv 
import PyPDF2
from pathlib import Path
from reportlab.pdfgen.canvas import Canvas
from reportlab.lib.pagesizes import LETTER
from reportlab.lib.units import inch, cm
from email import encoders
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import COMMASPACE
from datetime import datetime

#INFORMACION DEL EMISOR DEL CORREO
#SE REQUIERE ACTIVAR VERIFICACION DE 2 PASOS EN GMAIL PARA OBTENER CONTRASEÑA
email_emisor = "colaboradorgm1@gmail.com"
email_contrasena = "rlesbmmgbviufztt"

#ASUNTO DEL MENSAJE, SE CREA LA VARIABLE PARA DESPUÉS MANDARLA TRAER"
asunto = "Revisa el PDF que se te ha enviado"


# AQUI DEFINO LA FECHA DEL DIA DE HOY PARA COLOCARLA EN EL DOCUMENTO
fechaactual = datetime.now()
fechaactual2 = datetime.strftime(fechaactual, "%d/%m/%Y")



#AQUI MANDAMOS A TRAER EL ARCHIVO CSV Y DE DONDE SALDRA LA INFORMACION
with open("C:/Users/Fra ck/OneDrive - Universidad Tecmilenio/Desktop/Envio de correos/Workers.txt", mode='r') as csv_file:
    csv_reader = csv.DictReader(csv_file)

    #ESTE FOR HARA QUE SE ENVIEN LOS MENSAJES DE TODOS LOS REGISTROS EN EL CSV HASTA EL ULTIMO
    for row in csv_reader: 

    
        #AQUI CREAMOS EL ARCHIVO PDF
        canvas = Canvas("IMPORTANTE.pdf", pagesize=LETTER)

        #ESTE ES EL CONTENIDO DEL PDF QUE VAMOS A ENVIAR
        width, height = 595.27, 841.89  # A4 en puntos
        canvas.setPageSize((width, height))

        canvas.setFont('Helvetica-Bold', 14)
        canvas.drawString(50, height - 100, 'Querido Sr: ')
        canvas.drawString(50, height - 120, f'{row["Nombre"]}')

        canvas.setFont('Helvetica', 12)
        canvas.drawString(50, height - 140, '-- Por el medio de la siguiente carte queremos agradecerle lo siguiente:')
        canvas.drawString(50, height - 160, '-- Su participación en el proyecto de mejora continua')
        canvas.drawString(50, height - 180, '-- Completar el certificado de Seis Sigma')
        canvas.drawString(50, height - 200, '-- Obtener el trabajador del mes')

        #AQUI LE AGREGAMOS LA FECHA AL DOCUMENTO PDF
        canvas.setFont('Helvetica', 12)
        canvas.drawString(width - 180, height - 90, fechaactual2)

        #SE GUARDA LO CREADO EN EL DOCUMENTO PDF
        canvas.save()
                    

        # TENEMOS QUE ABRIR EL PDF EN MODO BINARIO PARA QUE PODAMOS ENVIARLO
        with open('IMPORTANTE.pdf', 'rb') as pdf_file:
        # Crea un objeto MIMEBase para adjuntar el archivo PDF
            pdf_part = MIMEBase('application', 'octet-stream')
            pdf_part.set_payload(pdf_file.read())
            encoders.encode_base64(pdf_part)
            pdf_part.add_header('Content-Disposition', 'attachment', filename='IMPORTANTE.pdf')



        #INFORMACIÓN SOBRE EL DESTINATARIO Y QUE VA A IR EN EL ASUNTO Y CONTENIDO
        email_receptor = (f'{row["email"]}')
        em = EmailMessage()
        em["From"] = email_emisor
        em["To"] = email_receptor
        em["Subject"] = asunto
        em.set_content(pdf_part)

        


        contexto = ssl.create_default_context()

        #AQUI CON SMTPLIB ES DONDE SE ENVIA EL CORREO CON LA INFORMACION ANTERIOR
        with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=contexto) as smtp:
            smtp.login(email_emisor, email_contrasena)

            smtp.sendmail(email_emisor, email_receptor, em.as_string())

            smtp.quit()
    
