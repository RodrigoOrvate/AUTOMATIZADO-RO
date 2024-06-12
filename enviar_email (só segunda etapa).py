import smtplib

from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

from email.mime.base import MIMEBase
from email import encoders

#Aqui abre o servidor
host = "smtp.gmail.com"
port = "587"
login = "rodrigoa.orvate@gmail.com"
senha = "roah kodp alqg wppu"

server = smtplib.SMTP(host,port)

server.ehlo()
server.starttls()
server.login(login,senha)

#Email tipo MIME

corpo = "<b>Se você está vendo isso é porque deu certo!</b>"

email_msg = MIMEMultipart()
email_msg['From'] = login
email_msg['To'] = login
email_msg['Subject'] = "RESULTADOS"
email_msg.attach(MIMEText(corpo, 'html')) #plain é texto normal, html é texto em formato html

cam_arquivo = "C:\Users\rodri\OneDrive\Área de Trabalho\AUTOMATIZADO\ANÁLISE_DE_PREÇOS_-_INFORMAÇÕES_TÉCNICAS.xlsx"
attchment = open(cam_arquivo,'rb')

att = MIMEBase('application', 'octet-stream')
att.set_payload(attchment.read())
encoders.encode_base64(att)

att.add_header('Contect-Disposition', f'attachment; filename= ANÁLISE_DE_PREÇOS_-_INFORMAÇÕES_TÉCNICAS.xlsx') #ERRO
attchment.close()
email_msg.attach(att)

#Enviar email tipo MIME no servidor SMTP
server.sendmail(email_msg['From'],email_msg['To'],email_msg.as_string())

server.quit()