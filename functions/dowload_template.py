from tkinter import filedialog
import csv

def dowload_template(name):    
    if name == 'email':
        email_template = filedialog.asksaveasfilename(
            title="Save email template as",filetypes=[('csv file', '*.csv')], 
            defaultextension='.csv', 
            initialfile='Email_template'
        )
        # If template file is selected:
        if email_template:  
            example_rows = [
                    ('John Doe','example1@example.com', 'cc_address@example.com', 'bcc_address@example.com')                    
                    ]
            with open(email_template,'w', newline='') as csvfile:
                writer = csv.writer(csvfile)
                writer.writerow(['Display Name','Email Address','CC address (Optional)', 'BCC address (Optional)'])                
                for Display_Name,Email_Address, CC, BCC in example_rows:
                    writer.writerow([Display_Name,Email_Address,CC, BCC])