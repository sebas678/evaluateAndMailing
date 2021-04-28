import smtplib 
import xlrd  #Libreria excel
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication


documento=xlrd.open_workbook("documento.xlsx") #Se abre documento

datos= documento.sheet_by_index(2)  #Se lee la hoja 
pdfDoc='ADJUNTO1.pdf'
carta='ADJUNTO2.pdf'
info='ADJUNTO3.pdf'

# Nos conectamos al servidor SMTP de Gmail 
emisor = "ejemplo@correo.com" #Correo de donde se envia

serverSMTP = smtplib.SMTP('smtp.gmail.com',587) 
serverSMTP.ehlo() 
serverSMTP.starttls() 
serverSMTP.ehlo() 
serverSMTP.login(emisor,"PASSWORD") 

for i in range(datos.nrows):#Recorremos el documento de excel 
    x=i+1
    receptor1=datos.cell_value(x,8)
    receptor2=datos.cell_value(x,10)

    if((receptor1.find('@')!=-1)&(receptor2.find('@')!=-1)):#Seleccionamos destinatarios si solamente uno o dos
        receptor=receptor1 +','+ receptor2
    elif(receptor1.find('@')==-1):
        receptor=receptor2
    elif(receptor2.find('@')==-1):
        receptor= receptor1

    # Configuracion del mail 
    mensaje = MIMEMultipart() 
    mensaje['From']=emisor
    mensaje['To']=receptor
    mensaje['Subject']="Envío de credenciales para "+str(datos.cell_value(x,6))+"." 

    #Crea/Adjunta mensaje Texto
    mensajito=MIMEText('''Estimados:
Adjunto encontraran información y credenciles para sus usuarios de plataforma
  1) PDF Bienvenida
  2) Credenciales Institucionales
  
      - Nombre: '''+str(datos.cell_value(x,2))+'''
      - Parqueo asignado: '''+str(datos.cell_value(x,6))+'''
      - Correo institucional: '''+str(datos.cell_value(x,3))+'''
      - Contraseña: '''+str(datos.cell_value(x,4))+'''

  3) Calendario (PDF Adjunto)
  4) Canales de comunicación (Infografía adjunta)
  5) Libros Sugeridos (Infografía Adjunta)

Atentamente,
Recursos Humanos.
''')
    mensaje.attach(mensajito)
    
    #Crea/Adjunta PDF
    pdfH=MIMEApplication(open(pdfDoc,'rb').read())
    pdfH.add_header('Content-Disposition','attachment',filename=pdfDoc)
    mensaje.attach(pdfH)
    
    pdfC=MIMEApplication(open(carta,'rb').read())
    pdfC.add_header('Content-Disposition','attachment',filename=carta)
    mensaje.attach(pdfC)

    pdfI=MIMEApplication(open(info,'rb').read())
    pdfI.add_header('Content-Disposition','attachment',filename=info)
    mensaje.attach(pdfI)
    print(receptor)
    print("\n")

    # Enviamos el mensaje 
    serverSMTP.sendmail(emisor,receptor.split(','),mensaje.as_string())

# Cerramos la conexion 
print("FIN")
serverSMTP.close()