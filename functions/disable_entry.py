import tkinter as tk

def disable_entry(*entry_list): #using *args here
    # Disable/Grey entry
    for entry in entry_list:
        entry.config(state='disabled')

def clean_entry(*entry_list):
    # Delete entry content
    for entry in entry_list:
        entry.delete(0,tk.END)

def enable_entry(*entry_list):
    # Set entry back to normal
    for entry in entry_list:
        entry.configure(state='normal')