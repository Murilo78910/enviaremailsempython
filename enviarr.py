'''PARA QUE FUNCIONE PREENCHA O ARQUIVO EXEL COM OS DADOSPOR COLUNA DA MANEIRA QUE LÁ ESTÁ E TAMBÉM SUBSTITUA O 'HOST' 
PARA O TIPO DE E-MAIL O QUAL ENVIARÁ; SEMPRE COM: 'smtp.' ANTES'''

import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

#S M T P - Simple Mail transfer protocol
#Para criar o servidor e enviar o e-mail

clientes = pd.read_excel('./clientes.xlsx')

for index, cliente in clientes.iterrows():

    #1- STARTAR o servidor SMTP
    host = "smtp.zoho.com"
    port = 587  # Porta correta para TLS
    login = ""
    senha = ""

    # Start no E-mail
    server = smtplib.SMTP(host, port)
    server.ehlo()
    server.starttls()
    server.login(login, senha)

    #2- Construir o EMAIL tipo MIME
    corpo = f"Olá {cliente['nome']}, você recebeu um email de fulano"
    email_msg = MIMEMultipart()
    email_msg['From'] = login
    email_msg['To'] = cliente['email']
    email_msg['Subject'] = "Meu E-mail via Python, "
    email_msg.attach(MIMEText(corpo, 'plain', 'utf-8'))

    #Para enviar um arquivo --->
    '''

    #Abrimos o arquivo em modo leitura e binary
    cam_arquivo = "C:\Estudo\logo.png"
    attchment = open(cam_arquivo, 'rb')

    #Lemos o arquivo no modo binario e jogamos codificado em base 64 (que é o que o e-mail precisa)
    att = MIMEBase('aplication', 'octet-stream')
    att.set_payload(attchment.read())
    encoders.encode_base64(att)

    #Adicionamos o cabeçalho no tipo anexo de e-mail
    att.add_header('Content-Disposition', f'attachment; filename= {'logo.png'}')

    #Fechamos o arqivo
    attchment.close()

    #Colocamos o anexo no corpo de e-mail
    email_msg.attach(att)

    '''

    #3- Enviar o EMAIL tipo MIME no SERVIDOR SMTP
    server.sendmail(email_msg['From'], email_msg['To'], email_msg.as_string())
    server.quit()