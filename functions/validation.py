# Conditions:
# List must have 4 columns
# Column 2, 3, 4 are email format
# Column 2 must have values at email format
from tkinter import messagebox
import socket
import re

# SMTP server validatioin:
def host_validation(host):
     # Try to get the hostname form input
     try:
          socket.gethostbyname(host)
          return True
     # If the hostname cannot be found
     except socket.gaierror:
          messagebox.showerror('Error', 'Hostname is incorrect')
          return False

# Validation email format:
def email_format_check(email):
     regex = r'^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z]+(?:;|$)' 
     if bool(re.match(regex,email)):
          return True
     else:
          return False     

# Custom validation email:
def email_validation(name):       
    # Check if list have 4 columns    
    if len(name[0]) == 4:
            for Display_name, email_address, cc_address, bcc_address in name[1:]: 
                
                # Check if email address is not empty                
                if email_address == '':
                    messagebox.showerror('Error', '"Email Address" can not empty!')
                    return False
                
                # Check if email address format is correct                    
                elif email_format_check(email_address) is False:
                    messagebox.showerror('Error', f"{Display_name}'s email address is not valid!")                 
                    return False

                # Check if cc or bcc address are available
                elif cc_address !='' and email_format_check(cc_address) is False:
                    messagebox.showerror('Error', f"{Display_name}'s CC email address is not valid!")                 
                    return False
                elif bcc_address !='' and email_format_check(bcc_address) is False:
                    messagebox.showerror('Error', f"{Display_name}'s BCC email address is not valid!")                 
                    return False
                               
            return True
    else:
        messagebox.showerror('Error', 'Missing column(s), please try again!')                
        return False