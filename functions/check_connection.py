import smtplib
from tkinter import messagebox
def check_connection(host, port, user, login):
    try:
        with smtplib.SMTP(host,port) as server: 
            server.starttls()
            server.login(user, login)            
            return True
        
    except smtplib.SMTPConnectError:
        messagebox.showerror("Error","Can not connect to server")
        return False
    # except smtplib.SMTPAuthenticationError:
    #     messagebox.showerror("Error","Your authentication is not correct")
    #     return False
    except smtplib.SMTPAuthenticationError:
        messagebox.showerror("Error","Username or Password incorrect")
        return False
    except TimeoutError:
        messagebox.showerror('Error', 'Time out error, a connection attempt failed')
        return False
    except ConnectionRefusedError:
        messagebox.showerror('Error', 'Hostname refuses to connect')  
        return False
        