import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from PIL import Image, ImageTk
import sv_ttk
import keyring
import csv
from functions import *

attachment_row = 0
attachment_column = 0
attachment_list = []
send_email_list = []

# This is used for upload_list() to toggle the upload and clear button
button_upload_list = None
attached_frame = None

def send_SMTPemail():

    # Check if smtp setting is connected:      
    if load_smtp_settings()[0] is not None: 
        # Get email, password, smtp_server, smtp_port return from load_smtp_settings()
        email, password, smtp_server, smtp_port = load_smtp_settings()
    
        # Check if send email list is uploaded:
        if len(send_email_list)>0:           
            # Get email's subtject:
            if entry_Subject.get()=='':
                answer=messagebox.askyesno('Your email has no subject line','Send anyway?')
                if answer:
                    subject = '(No subject)'
                    entry_Subject.insert('',subject)
            else:
                subject = entry_Subject.get()

            message = str(body_message.get('1.0', 'end'))
            send_SMTP = EmailSender(smtp_server,smtp_port,email,password)

            for Display_name, To_address, CC_address, BCC_address in send_email_list[1:]:                
                new_message = message.replace("['Display Name']",Display_name)              
                if checkbox_enableHTML_var.get() == 1:
                    send_SMTP.send_email(To_address,subject,attachment=attachment_list,html_content=new_message,cc_address=CC_address,bcc_address=BCC_address)      
                else:
                    send_SMTP.send_email(To_address,subject,attachment=attachment_list,plain_message=new_message,cc_address=CC_address,bcc_address=BCC_address)      
             # Check if Close After Send is checked:
            if checkbox_CloseAfterSend_var.get() == 1:                  
                messagebox.showinfo('Information',f'Email has been sent!')
                root.destroy()
            else:
                messagebox.showinfo('Information',f'Email has been sent!')
                
            
        # Else, using email data input from entry to send:
        else:
            # Check if recipient email address is input:
            if entry_To.get():
                if email_format_check(entry_To.get()):            
                    To_address = entry_To.get()                    
                else:
                    messagebox.showerror("Recipient's email incorrect","Please enter correct email address")
                    clean_entry(entry_To,entry_CC,entry_BCC)
                    # Get CC and BCC address once recipient address is correct:
                if entry_CC.get():
                    if email_format_check(entry_CC.get()):
                        CC_address = entry_CC.get()
                    else:
                        messagebox.showerror("Recipient's CC email incorrect","Please enter correct email CC address")                        
                        clean_entry(entry_CC)
                else:
                    CC_address = ''                
                if entry_BCC.get():
                    if email_format_check(entry_BCC.get()):
                        BCC_address = entry_BCC.get()
                    else:
                        messagebox.showerror("Recipient's BCC email incorrect","Please enter correct email BCC address")                        
                        clean_entry(entry_BCC)
                else:
                    BCC_address = ''       

                # Get email's subtject:
                if entry_Subject.get()=='':
                    answer=messagebox.askyesno('Your email has no subject line','Send anyway?')
                    if answer:
                        subject = '(No subject)'
                        entry_Subject.insert('',subject)
                else:
                    subject = entry_Subject.get()
                
                new_message = str(body_message.get('1.0', 'end'))
                
                send_SMTP = EmailSender(smtp_server,smtp_port,email,password)
                if checkbox_enableHTML_var.get() == 1:
                    send_SMTP.send_email( To_address,subject, attachment=attachment_list,html_content=new_message, cc_address=CC_address, bcc_address=BCC_address)      
                else:
                    send_SMTP.send_email( To_address,subject, attachment=attachment_list,plain_message=new_message, cc_address=CC_address, bcc_address=BCC_address)      
                if checkbox_CloseAfterSend_var.get() == 1:                  
                    messagebox.showinfo('Information',f'Email has been sent!')
                    root.destroy()
                else:
                    messagebox.showinfo('Information',f'Email has been sent!')
                    
            else:
                messagebox.showerror("Recipient's email missing","Please enter receipient(s) email address or upload a list") 
    else:
        messagebox.showerror("SMTP setting missing","Please connect to SMTP server")  
    
        


def add_attachment():
    def clear_attachment():
        attachment_list.clear()
        if attached_frame is not None:        
            attached_frame.destroy()
    global attached_frame         
    # Clear attachment list       
    clear_attachment()
    # Select attachment files
    attachment_fullpath = filedialog.askopenfilenames(title="Add attachment",filetypes=[('Text or Image',('*.csv','*.xlsx','*.pdf','*.jpg','*.png','*.jpg'))] )    
    if attachment_fullpath:
        attached_frame = ttk.Frame(Right_Attachment)
        attached_frame.grid(row=attachment_row, column=attachment_column+1, padx=5, pady=5)    
        attached_file = ttk.Label(attached_frame, text=f'{len(attachment_fullpath)} file(s) added')
        attached_file.pack(side=tk.LEFT)
        # Clear attachment button:
        clear_attachment_button = ttk.Button(attached_frame, text='X', command=clear_attachment )
        clear_attachment_button.pack(side=tk.LEFT, padx=5)
        for file in attachment_fullpath:            
                attachment_list.append(file)         
    # print(attachment_list)      

def load_smtp_settings():
    # Get and store values:
    email = keyring.get_password("smtp_app", "email")
    password = keyring.get_password("smtp_app", "password")
    smtp_server = keyring.get_password("smtp_app", "smtp_server")
    smtp_port = keyring.get_password("smtp_app", "smtp_port")
    return email, password, smtp_server, smtp_port

def clear_listing():
    # Set the To, CC, BCC entry back to normal
    enable_entry(entry_To,entry_CC,entry_BCC)
    
    # Clean current entry (Note: Entry must set to normal before clean):
    clean_entry(entry_To,entry_CC,entry_BCC)

    # Destroy clear button    
    button_clear_list.destroy()

    # Re-Create Upload List button:
    button_upload_list = ttk.Button(emailItem_frame, text='Upload List', command= lambda: upload_list('email'))
    button_upload_list.grid(padx=5, pady=5,row=4,column=3, sticky='we')

def upload_list(name):    
    if name == 'email':        
        email_list= filedialog.askopenfilename(title="Upload Email Sending List",filetypes=[('csv file', '*.csv')], defaultextension='.csv')
        # If upload file is selected:
        if email_list:            
            try:       
                with open(email_list,'r') as csvfile:
                    upload_email_list = list(csv.reader(csvfile))
                    # Clear the entry boxes and the send_email_list list.
                    clean_entry(entry_To,entry_CC,entry_BCC)                
                    send_email_list.clear()                            
                    # If Validation email is True
                    if email_validation(upload_email_list):
                        # Destroy upload button
                        button_upload_list.destroy()
                        entry_To.insert('','Email is using upload list to send')
                        entry_CC.insert('','CC Email is using upload list to send')
                        entry_BCC.insert('','BCC Email is using upload list to send')
                        # Add email into email list
                        for values in upload_email_list:                            
                            send_email_list.append(values)
                        # Disable entry
                        disable_entry(entry_To, entry_CC, entry_BCC)
                        # Create clear button
                        global button_clear_list                                                
                        button_clear_list = ttk.Button(
                            emailItem_frame, 
                            text='Clear', 
                            command=clear_listing)
                        button_clear_list.grid(padx=5, pady=5,row=4,column=3, sticky='we')                        
                
            except FileNotFoundError:
                messagebox.showerror('Error', 'The file was not found.')
            except IOError:
                messagebox.showerror('Error', 'The file is not a CSV file.')
            except:
                messagebox.showerror('Error', 'An unexpected error occurred.')                      

# [SMTP SETTING]:
def SMTP_setting():
    # Define SMTP_setting_window to entire app:
    global SMTP_settings_window 
          
    # If setting window is not pop up:
    if SMTP_settings_window is None or not SMTP_settings_window.winfo_exists():
        SMTP_settings_window = tk.Toplevel(root)
        SMTP_settings_window.resizable(False, False)       
        
        # Create disconect button:
        def disconnecting():
            enable_entry(entry_SMTPemail,entry_SMTPpassword,entry_SMTPport,entry_SMTPserver)                        
                                        
            keyring.delete_password('smtp_app', 'email')
            keyring.delete_password("smtp_app", "password")
            keyring.delete_password("smtp_app", "smtp_server")
            keyring.delete_password("smtp_app", "smtp_port")
            clean_entry(entry_SMTPpassword)

            entry_From.configure(state="normal") 
            entry_From.delete(0,tk.END)
            entry_From.configure(state="disabled")
            
            SMTP_message.config(text='Server disconected!')
            button_connect.config(text='Connect')
            global current_connection_pharse
            current_connection_pharse = connecting
            button_sendTest_setting.config(state='disabled')
        
        def connecting():
            # Get values returned by get_smtp_setting()
            email, password, smtp_server, smtp_port = get_smtp_settings()           
            # Check if email is emtpy:
            if email =='':
                SMTP_message.config(text='Please enter an email address!')
                # # Keep window stay on top:
                # SMTP_settings_window.lift()      
                # # Or:
                # SMTP_settings_window.attributes('-topmost', True)
                # SMTP_settings_window.attributes('-topmost', False)                      
            # Check if email address is valid:
            elif email_format_check(entry_SMTPemail.get()) is not True:
                SMTP_message.config(text='Input email is not valid!')                
            # Check if password is emtpy:
            elif password == '':
                SMTP_message.config(text='Please enter password!')                
            # Check if SMTP server is emtpy:
            elif smtp_server == '':
                SMTP_message.config(text='Please enter SMTP server address!')
            # Check if SMTP port is emtpy:
            elif smtp_port == '':
                SMTP_message.config(text='Please enter an SMTP Port!')
            # Check if SMTP port is digits:                                
            elif not smtp_port.isdigit():
                SMTP_message.config(text='SMTP Port is not valid!')                
            # Test connection:
            elif host_validation(smtp_server) is True:
                if check_connection(smtp_server,smtp_port,email,password):
                        # insert encrypted value into smtp_app:
                        keyring.set_password("smtp_app", "email", email)
                        keyring.set_password("smtp_app", "password", password)
                        keyring.set_password("smtp_app", "smtp_server", smtp_server)
                        keyring.set_password("smtp_app", "smtp_port", smtp_port)  
                        # Disable SMTP entries:
                        disable_entry(entry_SMTPemail,entry_SMTPpassword,entry_SMTPport,entry_SMTPserver)

                        SMTP_message.config(text='Connected to server!')
                        
                        # Update the email address in the main window
                        entry_From.configure(state="normal") # Enable input values
                        entry_From.delete(0, tk.END)  # Clear the old values
                        entry_From.insert(0, email)  # Insert the encrypted email address
                        entry_From.configure(state="disabled") # Disable input values  

                        button_connect.config(text='Disconnect')
                        global current_connection_pharse
                        current_connection_pharse = disconnecting
                        button_sendTest_setting.config(state='enabled')
        
        def toggle_connection():
            global current_connection_pharse            
            current_connection_pharse()

        
        # Get values from setting entry:
        def get_smtp_settings():
            email = entry_SMTPemail.get()
            password = entry_SMTPpassword.get()
            smtp_server = entry_SMTPserver.get()
            smtp_port = entry_SMTPport.get()
            return email, password, smtp_server, smtp_port

        def toggle_password_visibility():
            # Show password:
            if showpass_var.get() == 1:
                entry_SMTPpassword.configure(show="")
            # Hide password:
            else:
                entry_SMTPpassword.configure(show="*")

        def delete_SMTP():            
            enable_entry(entry_SMTPemail,entry_SMTPpassword,entry_SMTPport,entry_SMTPserver)
            keyring.delete_password('smtp_app', 'email')
            keyring.delete_password("smtp_app", "password")
            keyring.delete_password("smtp_app", "smtp_server")
            keyring.delete_password("smtp_app", "smtp_port")
            entry_SMTPemail.delete(0, tk.END)
            entry_SMTPpassword.delete(0, tk.END)
            entry_SMTPserver.delete(0, tk.END)
            entry_SMTPport.delete(0, tk.END)
            entry_From.configure(state="normal") 
            entry_From.delete(0,tk.END)
            entry_From.configure(state="disabled")
            SMTP_message.config(text='SMTP settings have been reset!')
            button_connect.config(text='Connect')

            global current_connection_pharse
            current_connection_pharse = connecting

        def send_testing():
            email, password, smtp_server, smtp_port = get_smtp_settings()           
            # Check if email is emtpy:            
            if email =='':
                SMTP_message.config(text='Please enter an email address!')
                # # Keep window stay on top:
                # SMTP_settings_window.lift()      
                # # Or:
                # SMTP_settings_window.attributes('-topmost', True)
                # SMTP_settings_window.attributes('-topmost', False)                    
            # Check if email address is valid:
            elif email_format_check(entry_SMTPemail.get()) is False:
                SMTP_message.config(text='Input email is not valid!')                
            # Check if password is emtpy:
            elif password == '':
                SMTP_message.config(text='Please enter password!')                
            # Check if SMTP server is emtpy:
            elif smtp_server == '':
                SMTP_message.config(text='Please enter SMTP server address!')              
            # Check if SMTP port is emtpy:
            elif smtp_port == '':
                SMTP_message.config(text='Please enter an SMTP Port!')
            # Check if SMTP port is digits:                                
            elif not smtp_port.isdigit():
                SMTP_message.config(text='SMTP Port is not valid!')
                
            else:
                sending_test_window = tk.Toplevel(SMTP_settings_window)
                # Disable window resize:
                sending_test_window.resizable(False,False)
                # Set window position:
                sending_test_window.geometry(f"+{x+50}+{y+50}")

                # Create def:
                def cancel_test():
                    sending_test_window.destroy()
                    # Bring SMTP setting window to front:
                    SMTP_settings_window.lift()
                    
                def send_test():
                    recipient = entry_test_email.get()                    
                    subject = 'Test email'
                    message = '<p>Hi,</p><p>This is a test email</p>'
                    if recipient =='':
                        messagebox.showerror("Recipient email missing","Please enter recipient's email" )
                        # Bring send test window to front:
                        sending_test_window.lift()
                    elif email_format_check(recipient) is False:
                        messagebox.showerror("Recipient email incorrect","Please enter valid recipient's email" )
                        # Bring send test window to front:
                        sending_test_window.lift()
                    else:
                        # Assign email, password, smtp_server, smtp_port for values returned from load_smtp_setting:
                        # email, password, smtp_server, smtp_port = load_smtp_settings()
                        send_test_SMTP = EmailSender(smtp_server,smtp_port,email,password)
                        send_test_SMTP.send_email(recipient,subject,html_content=message) 
                        messagebox.showinfo('Information',f'Email has been sent!')
                        sending_test_window.destroy()                  
                        SMTP_settings_window.destroy()         

                label_test_email = ttk.Label(sending_test_window, text='Send a test email to:')
                label_test_email.grid(row=0,column=0,padx=5,pady=5)

                entry_test_email = ttk.Entry(sending_test_window, width=30)
                entry_test_email.grid(row=0,column=1,padx=5,pady=5,columnspan=2)

                button_send_test = ttk.Button(sending_test_window, text='Send', command=send_test)
                button_send_test.grid(row=1,column=1,padx=5, pady=5, sticky='we')

                button_close_test = ttk.Button(sending_test_window, text='Cancel', command=cancel_test)
                button_close_test.grid(row=1,column=2,padx=5,pady=5, sticky='we')
        
        global current_connection_pharse
        current_connection_pharse = connecting
         
        label_SMTPemail = ttk.Label(SMTP_settings_window, text='Email:')      
        label_SMTPemail.grid(row=0,column=0,padx=5,pady=5, sticky='e')
        entry_SMTPemail = ttk.Entry(SMTP_settings_window, width=45)
        entry_SMTPemail.grid(row=0,column=1,columnspan=4,padx=5,pady=5, sticky='we')

        label_SMTPpassword = ttk.Label(SMTP_settings_window, text='Password:')      
        label_SMTPpassword.grid(row=1,column=0,padx=5,pady=5, sticky='e')
        entry_SMTPpassword = ttk.Entry(SMTP_settings_window, show='*', width=45)
        entry_SMTPpassword.grid(row=1,column=1,padx=5,columnspan=2,pady=5)
        
        showpass_var = tk.IntVar()
        checkbox_showpass = ttk.Checkbutton(SMTP_settings_window,text='Show', variable=showpass_var, command=toggle_password_visibility)
        checkbox_showpass.grid(row=1,column=3,padx=5,pady=5)

        label_SMTPserver = ttk.Label(SMTP_settings_window, text='SMTP server:')      
        label_SMTPserver.grid(row=2,column=0,padx=5,pady=5, sticky='e')
        entry_SMTPserver = ttk.Entry(SMTP_settings_window, width=25)
        entry_SMTPserver.grid(row=2,column=1,padx=5,pady=5,columnspan=2, sticky='we')

        label_SMTPport = ttk.Label(SMTP_settings_window, text='SMTP port:')      
        label_SMTPport.grid(row=2,column=3,padx=5,pady=5, sticky='w')
        entry_SMTPport = ttk.Entry(SMTP_settings_window, width=5)
        entry_SMTPport.grid(row=2,column=4,padx=5,pady=5, sticky='w')

        button_sendTest_setting = ttk.Button(SMTP_settings_window,text='Send Testing', command=send_testing, state='disabled')        
        button_sendTest_setting.grid(row=3,column=0,padx=5, pady=5)

        SMTP_message = ttk.Label(SMTP_settings_window,text='')
        SMTP_message.grid(row=3,column=1,padx=5, pady=5, sticky='we')

        button_connect = ttk.Button(SMTP_settings_window,text='Connect', command=toggle_connection)
        button_connect.grid(row=3,column=2,padx=5, pady=5, sticky='e')        
                
        button_cancel_setting = ttk.Button(SMTP_settings_window,text='Close', command=SMTP_settings_window.destroy )
        button_cancel_setting.grid(row=3,column=3,padx=5,pady=5)
        
        button_reset_setting = ttk.Button(SMTP_settings_window,text='Reset', command=delete_SMTP )
        button_reset_setting.grid(row=3,column=4,padx=5,pady=5)

        # Center the settings window on the screen
        SMTP_settings_window.update_idletasks()
        window_width = SMTP_settings_window.winfo_width()
        window_height = SMTP_settings_window.winfo_height()
        screen_width = SMTP_settings_window.winfo_screenwidth()
        screen_height = SMTP_settings_window.winfo_screenheight()
        x = int((screen_width - window_width) / 2)
        y = int((screen_height - window_height) / 2)
        SMTP_settings_window.geometry(f"+{x}+{y}")  
        # If smtp setting is available      
        if load_smtp_settings()[0] is not None:            
            # Assign email, password, smtp_server, smtp_port for values returned from load_smtp_setting:
            email, password, smtp_server, smtp_port = load_smtp_settings()
            # Show SMTP Setting if they are available
            entry_SMTPemail.insert(0, email)
            entry_SMTPpassword.insert(0, password)
            entry_SMTPserver.insert(0, smtp_server)
            entry_SMTPport.insert(0, smtp_port)
            disable_entry(entry_SMTPemail,entry_SMTPport,entry_SMTPserver,entry_SMTPpassword)           
            current_connection_pharse = disconnecting
            button_connect.config(text='Disconect')
            button_sendTest_setting.config(state='enabled')
        
def load_email_address():          
        email = keyring.get_password("smtp_app", "email")        
        if email:
            # Update the email address in the main window
            entry_From.configure(state="normal")
            entry_From.delete(0, tk.END)  # Clear the entry
            entry_From.insert(0, email)  # Insert the updated email address        
            entry_From.configure(state="disabled")
        else:
            entry_From.configure(state="normal")
            entry_From.insert(0, 'Please setup sending account in SMTP setting')
            entry_From.configure(state="disabled")
   
def insert_text():
    txtfile = filedialog.askopenfilename(title="Select file",filetypes=[('text file', '*.txt')])   
    
    def cleanmessage():
        body_message.delete('1.0','end')
        # clean_message.config(text='Message Cleaned!')
        button_cleantxt.grid_forget()  
    # If text file is selected    
    if txtfile:        
        # Open text file
        txtfile = open(txtfile,'r')
        # Add text content to text
        text = txtfile.read()     
        # Close text file   
        txtfile.close()
        # Clean current email content:
        body_message.delete('1.0','end')    
        # Clear message:
        clean_message.config(text='')
        # Insert new email content
        body_message.insert(tk.END, text)        
        # Create clear button
        button_cleantxt = ttk.Button(body_frame, text='Clean text', command=cleanmessage)
        button_cleantxt.grid(row=0, column=2, padx=5, pady=5)       
       
# Get current working directory:
current_dir = os.getcwd()
SMTP_settings_window = None
# Create root:
root = tk.Tk()
root.title("SMTP Email Sending")
icon_file = os.path.join(current_dir,'email_icon.png')
icon = tk.PhotoImage(file=icon_file)
root.iconphoto(True, icon)
root.resizable(width=False,height=True)

# Load the image:
image = Image.open(os.path.join(current_dir,'email_icon.png'))
image = image.resize((100,100))

# Convert the image to Tkinter format:
tk_image = ImageTk.PhotoImage(image)

# Select theme
sv_ttk.set_theme("light")

# Create Frame:
frame = ttk.Frame(root)
frame.pack(padx=5, pady=5) # Show the frame

# Email Frame:
email_frame = ttk.Frame(frame)
email_frame.grid(padx=5,row=0,column=0, sticky='we')

# Create a label within the frame and display the image
label_send = ttk.Label(email_frame, image=tk_image)
label_send.grid(row=0,column=0,padx=5, pady=5, sticky='nswe')

# Email item frame:
emailItem_frame = ttk.Frame(email_frame)
emailItem_frame.grid(row=0,column=1, sticky='we')
# Label From:
label_From = ttk.Label(emailItem_frame, text="From:")
label_From.grid(padx=2, pady=5,row=0,column=0, sticky='e' )
# Entry From:
entry_From = ttk.Entry(emailItem_frame, width=70)
entry_From.grid(padx=5, pady=5,row=0,column=1,columnspan=4, sticky='we')
# Label To:
label_To = ttk.Label(emailItem_frame,text="To:")
label_To.grid(padx=2, pady=5,row=1,column=0, sticky='e')
# Entry To:
entry_To = ttk.Entry(emailItem_frame)
entry_To.grid(padx=5, pady=5,row=1,column=1,columnspan=4, sticky='we')
# Label CC:
label_CC = ttk.Label(emailItem_frame,text="CC:")
label_CC.grid(padx=2, pady=5,row=2,column=0, sticky='e')
# Entry CC:
entry_CC = ttk.Entry(emailItem_frame)
entry_CC.grid(padx=5, pady=5,row=2,column=1,columnspan=4, sticky='we')
# Label BCC:
label_BCC = ttk.Label(emailItem_frame,text="BCC:")
label_BCC.grid(padx=2, pady=5,row=3,column=0, sticky='e')
# Entry BCC:
entry_BCC = ttk.Entry(emailItem_frame)
entry_BCC.grid(padx=5, pady=5,row=3,column=1,columnspan=4,sticky='we')
# Button Upload List:
button_upload_list = ttk.Button(emailItem_frame, text='Upload List', command= lambda: upload_list('email') )
button_upload_list.grid(padx=5, pady=5,row=4,column=3, sticky='we')
# Button Sample List:
button_get_email_template = ttk.Button(emailItem_frame, text='Download Sample', command= lambda: dowload_template('email'))
button_get_email_template.grid(padx=5, pady=5,row=4,column=4,sticky='we')
# Label Subject:
label_Subject = ttk.Label(emailItem_frame,text="Subject:")
label_Subject.grid(padx=2, pady=5,row=5,column=0, sticky='e')
# Entry Subject:
entry_Subject = ttk.Entry(emailItem_frame)
entry_Subject.grid(padx=5, pady=5,row=5,column=1,columnspan=4, sticky='we')

# File loading frame:
attachment_frame = ttk.Frame(frame)
attachment_frame.grid(row=2,column=0,sticky='w')

Left_Attachment = ttk.Frame(attachment_frame)
Left_Attachment.grid(row=0,column=0,padx=5,pady=5)

label_Attachment = ttk.Label(Left_Attachment,text="Attachment(s):")
label_Attachment.pack(padx=15, pady=5)

# entry_Attachment =ttk.Treeview(body_frame, height=2)
# entry_Attachment.grid(padx=5, pady=5,row=0,column=1,rowspan=2, sticky='we')
Right_Attachment = ttk.Frame(attachment_frame, width=65)
Right_Attachment.grid(row=0,column=1,padx=5,pady=5)

button_attachment = ttk.Button(Right_Attachment, text='Add...', command= add_attachment)
button_attachment.grid(column=attachment_column,row=attachment_row,padx=5,pady=5)

# Body Frame:
body_frame = ttk.Frame(frame)
body_frame.grid(row=3,column=0, sticky='w')

label_MessageBody = ttk.Label(body_frame,text="Message Body:")
label_MessageBody.grid(padx=15, pady=5,row=0,column=0)

button_insert_text = ttk.Button(body_frame, text='Insert Text Files', command=insert_text)
button_insert_text.grid(padx=15, pady=5,row=0,column=1)

checkbox_enableHTML_var = tk.IntVar(value=1)
checkbox_enable_html = ttk.Checkbutton(body_frame,text='Send as HTML', variable=checkbox_enableHTML_var)
checkbox_enable_html.grid(padx=15, pady=5,row=0,column=3)

clean_message = ttk.Label(body_frame, text='')
clean_message.grid(row=0,column=3, padx=5, pady=5)

message_frame = ttk.Frame(frame)
message_frame.grid(row=4,column=0)

message_y_scroll = ttk.Scrollbar(message_frame)
message_y_scroll.pack(side='right',fill='y')

body_message = tk.Text(message_frame, height=15, yscrollcommand=message_y_scroll.set)
body_message.pack(side='left', padx=5)

message_y_scroll.config(command=body_message.yview)

bottom_frame = ttk.Frame(frame)
bottom_frame.grid(row=5,column=0,padx=2, pady=5, sticky='we')
button_setting = ttk.Button(bottom_frame, text='SMTP Setting', command=SMTP_setting)
button_setting.pack(side='left')


button_send = ttk.Button(bottom_frame, text='Send', command=send_SMTPemail )
button_send.pack(side='right')

checkbox_CloseAfterSend_var = tk.IntVar(value=1)
checkbox_CloseAfterSend = ttk.Checkbutton(bottom_frame, text='Close window after send', variable=checkbox_CloseAfterSend_var)

# email_entry = ttk.Entry(SMTP_settings_window)
# email_entry.grid(row=0,column=1, padx=5, pady=5)

checkbox_CloseAfterSend.pack(side='right')

# Load email address from SMTP setting to From:
load_email_address()

root.mainloop()