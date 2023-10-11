import time
import sys
import smtplib, ssl
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import helpers
import config_email
import email_list


try:
    params_config = config_email.params_config()
    sender_address = params_config.username
    sender_host =  params_config.host
    sender_port = int(params_config.port)
    receiver_address = email_list.get_list_email()

    if(len(sys.argv)>1):
        #Setup the MIME
        message = MIMEMultipart()
        start_time = time.time()
        print('Inicianado el envio: ')
        message['Subject'] = str(sys.argv[1])  #The subject line
        mail_content = str(sys.argv[2])
        
        message['From'] = sender_address
        
        message.attach(MIMEText(mail_content, 'html'))
        
        for email_to in receiver_address:
            message['To'] = email_to
            #The body and the attachments for the mail
            #use gmail with port
            sender_context = ssl.create_default_context()
            with smtplib.SMTP_SSL(sender_host, sender_port, context=sender_context) as session:
                # session.starttls() #enable security
                text = message.as_string()
                session.sendmail(sender_address, email_to, text)
                session.quit()
                time_process = str(round((time.time() - start_time),2))
                helpers.put_log("Se envio el mensaje "+email_to+": "+str(time_process),"--","senderEmail")
        
        
except ValueError  as error:
    except_info = sys.exc_info()
    s_message = f'({except_info[2].tb_lineno}) {except_info[0]} {str(error)}'
    helpers.put_log(s_message,"--","senderEmail")
    print('Giskard: ', 'existe un error al enviar el correo revisar el archivo de log')