import smtplib
import os
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from tkinter import messagebox

class EmailSender:

    def __init__(self, smtp_server, port, username, password):
        self.smtp_server = smtp_server
        self.port = port
        self.username = username
        self.password = password

    def send_email(self, to_email, subject, plain_message=None, attachment=None, html_content=None,cc_address=None, bcc_address=None ):
        msg = MIMEMultipart()        
        msg['From'] = self.username
        msg['To'] = to_email     
        msg['Subject'] = subject   
    
        if plain_message:
            part = MIMEText(plain_message, 'plain')
            msg.attach(part)
        
        elif html_content:
            part = MIMEText(html_content, 'html')
            msg.attach(part)

        if cc_address:
            msg['Cc'] = cc_address
        
        if bcc_address:
            msg['Bcc'] = bcc_address

        if attachment:
            for file in attachment:
                filename = os.path.basename(file)
                with open (file, 'rb') as fp:
                    part = MIMEBase('application', 'octet-stream')
                    part.set_payload(fp.read())
                    encoders.encode_base64(part)
                    part.add_header('Content-Disposition','file; filename="{}"'.format(filename))
                    msg.attach(part)

        try: # Note: use 'try' must have 'except'
            with smtplib.SMTP(self.smtp_server, self.port) as server:
                # Enable email STARTTLS encryption:
                server.starttls()
                server.login(self.username, self.password)                   
                server.send_message(msg) # Use send_messaget o send MIME type and send as a single object (msg). More flexible than sendmail()
                # messagebox.showinfo('Information',f'Email has been sent!')

        except Exception as e:
            messagebox.showerror("An error occurred while sending the email: \n",f"{e}")