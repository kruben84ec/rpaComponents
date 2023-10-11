import sys
import smtplib, ssl
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from HoockUtilities import hoock_utilities


class Email(hoock_utilities):
    def __init__(self, config):
        super(Email, self).__init__(config)
        self.configuration = self.get_data_json(config)
        self.log = str(self.configuration.log)
        self.configuration_email = self.get_data_json(self.configuration.email_config)
        self.receivers = self.get_data_json(self.configuration.email_list)
    
    def sender_email(self, subject, content):
        message = MIMEMultipart()
        message['Subject'] = str(subject)  #The subject line
        mail_content = str(content)
        message['From'] = self.configuration_email.username
        message.attach(MIMEText(mail_content, 'html'))
        
        try:
            for email_to in self.receivers.listEmail:
                message['To'] = email_to
                sender_host = self.configuration_email.host
                sender_port = int(self.configuration_email.port)
                with smtplib.SMTP(sender_host, sender_port) as server:
                    email_message = message.as_string()
                    s_message = "Se envió se correo: "+email_message
                    server.starttls()
                    server.sendmail(self.configuration_email.username, email_to, email_message)
                    self.put_log(s_message,"--","Email", self.log+"/Email.txt")
                    server.quit()

        except IOError as error:
            except_info = sys.exc_info()
            s_message = f'({except_info[2].tb_lineno}) {except_info[0]} {str(error)}'
            self.put_log(s_message,"--","Email", self.log+"/Email.txt")
