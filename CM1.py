# -*- coding: utf-8 -*-
"""
Created on Wed Aug 23 01:51:27 2023

@author: H
"""

# IMPORTS
import tkinter as tk
from tkinter import ttk
import openpyxl
import os
# END IMPORTS

# This project attempts to create a small customer relationship management system
# The two main feature at first are:
# Easily view, edit, add and remove clients and their info from a simple database, using a minimal GUI
# Allow the population of the client data by importing Excel sheets
# Add an option for sending WhatsApp messages and notifications to customers

# LOG START
# -23AUG--CREATED GUI INFRASTRUCTURE--(NEED TO REVIEW CONCEPTS)
# -28AUG--ADDED BUTTONS AND ENTRY BOXES--NOT YET FUNCTIONAL
# -30AUG--UPDATED SOME BUTTONS FUNCTIONALITY
# -08SEP--INTRODUCED EXCEL IMPORT FUNCTIONALITY
# -TODO---ADD RECORD BUTTON SHOULD INSERT VALUES TO THE TREE AND TO THE EXCEL FILE
# LOG END


# First Step, create data structure

def load_data():
    current_file_path = os.path.dirname(__file__)
    xl_file_path = current_file_path + "\customers.xlsx"
    workbook = openpyxl.load_workbook(xl_file_path)
    sheet = workbook.active
    
    list_values = list(sheet.values)

    for value_tuple in list_values[1:]:
        my_tree.insert('', tk.END, values=value_tuple)

#############################################################################


# Style
root = tk.Tk()
root.title = ('CRM Test')
root.geometry = ("1000x500")
root.iconbitmap = ('F:\Art\Icons\pc.ico')
style = ttk.Style()

style.theme_use('default')

# Treeview colors
style.configure("Treeview", background="#D3D3D3",
                foreground="black", rowheight=25, fieldbackground="#D3D3D3")

# Change color of selected client
style.map('Treeview', background=[('selected', "347083")])

# Frame and scrollbar
frame = ttk.Frame(root)
frame.pack(pady=10)

treeFrame = ttk.Frame(frame)
treeFrame.grid(row=0, column=1, pady=10)
treeScroll = ttk.Scrollbar(treeFrame)
treeScroll.pack(side="right", fill="y")


# Create the treeview

my_tree = ttk.Treeview(
    treeFrame, yscrollcommand=treeScroll.set, selectmode="extended")





# Creating Columns
my_tree['columns'] = ("Name", "Address", "Area", "Phone",
                      "Phone 2", "Tax Status", "Fees Status", "Notes")

my_tree.column("#0", width=0)
my_tree.column("Name", width=140)
my_tree.column("Address", width=140)
my_tree.column("Area", width=140)
my_tree.column("Phone", width=140)
my_tree.column("Phone 2", width=140)
my_tree.column("Tax Status", width=140)
my_tree.column("Fees Status", width=140)
my_tree.column("Notes", width=140)

# Creating Headings
my_tree.heading("#0", text="")
my_tree.heading("Name", text="Name")
my_tree.heading("Address", text="Address")
my_tree.heading("Area", text="Area")
my_tree.heading("Phone", text="Phone")
my_tree.heading("Phone 2", text="Phone 2")
my_tree.heading("Tax Status", text="Tax Status")
my_tree.heading("Fees Status", text="Fees Status")
my_tree.heading("Notes", text="Notes")



############################################################################


#############################################################################



# Add record entry boxes
data_frame = ttk.LabelFrame(root, text="Record")
data_frame.pack(fill="x", expand="yes", padx=20)


###
name_label = ttk.Label(data_frame, text="Name")
name_label.grid(row=0, column=0, padx=10, pady=10)

name_entry = ttk.Entry(data_frame)
name_entry.grid(row=0, column=1, padx=10, pady=10)
###


adrs_label = ttk.Label(data_frame, text="Address")
adrs_label.grid(row=0, column=2, padx=10, pady=10)

adrs_entry = ttk.Entry(data_frame)
adrs_entry.grid(row=0, column=3, padx=10, pady=10)
###


area_label = ttk.Label(data_frame, text="Area")
area_label.grid(row=0, column=4, padx=10, pady=10)

area_entry = ttk.Entry(data_frame)
area_entry.grid(row=0, column=5, padx=10, pady=10)
###


ph_label = ttk.Label(data_frame, text="Phone")
ph_label.grid(row=1, column=0, padx=10, pady=10)

ph_entry = ttk.Entry(data_frame)
ph_entry.grid(row=1, column=1, padx=10, pady=10)
###


ph2_label = ttk.Label(data_frame, text="Phone 2")
ph2_label.grid(row=1, column=2, padx=10, pady=10)

ph2_entry = ttk.Entry(data_frame)
ph2_entry.grid(row=1, column=3, padx=10, pady=10)
###


tax_label = ttk.Label(data_frame, text="Tax Status")
tax_label.grid(row=1, column=4, padx=10, pady=10)

tax_entry = ttk.Entry(data_frame)
tax_entry.grid(row=1, column=5, padx=10, pady=10)
###


fees_label = ttk.Label(data_frame, text="Office Fees Status")
fees_label.grid(row=1, column=6, padx=10, pady=10)

fees_entry = ttk.Entry(data_frame)
fees_entry.grid(row=1, column=7, padx=10, pady=10)
###


cmts_label = ttk.Label(data_frame, text="Notes")
cmts_label.grid(row=0, column=6, padx=10, pady=10)

cmts_entry = ttk.Entry(data_frame)
cmts_entry.grid(row=0, column=7, padx=10, pady=10)
###

#####################################################################

# Select Record


def select_record(e):
    # first clear the entry boxes
    name_entry.delete(0, 'END')
    adrs_entry.delete(0, 'END')
    area_entry.delete(0, 'END')
    ph_entry.delete(0, 'END')
    ph2_entry.delete(0, 'END')
    tax_entry.delete(0, 'END')
    fees_entry.delete(0, 'END')
    cmts_entry.delete(0, 'END')

    # grab record number
    selected = my_tree.focus()
    # grab record values
    values = my_tree.item(selected, 'values')

    # output to entry boxes
    name_entry.insert(0, values[0])
    adrs_entry.insert(0, values[1])
    area_entry.insert(0, values[2])
    ph_entry.insert(0, values[3])
    ph2_entry.insert(0, values[4])
    tax_entry.insert(0, values[5])
    fees_entry.insert(0, values[6])
    cmts_entry.insert(0, values[7])

#####################################################################
# Remove record functionality


def remove_one():
    x = my_tree.selection()[0]
    my_tree.delete(x)


def remove_all():
    for record in my_tree.get_children():
        my_tree.delete(record)


#####################################################################
# Update record functionality


def update_record():
    # grab rec number
    selected = my_tree.focus()
    # update record
    my_tree.item(selected, text="", values=(name_entry.get(), adrs_entry.get(), area_entry.get(
    ), ph_entry.get(), ph2_entry.get(), tax_entry.get(), fees_entry.get(), cmts_entry.get(),))

    # first clear the entry boxes
    name_entry.delete(0, 'END')
    adrs_entry.delete(0, 'END')
    area_entry.delete(0, 'END')
    ph_entry.delete(0, 'END')
    ph2_entry.delete(0, 'END')
    tax_entry.delete(0, 'END')
    fees_entry.delete(0, 'END')
    cmts_entry.delete(0, 'END')

#####################################################################
# Clear boxes functionality


def clear_boxes():

    #clear the entry boxes
    name_entry.delete(0, 'END')
    adrs_entry.delete(0, 'END')
    area_entry.delete(0, 'END')
    ph_entry.delete(0, 'END')
    ph2_entry.delete(0, 'END')
    tax_entry.delete(0, 'END')
    fees_entry.delete(0, 'END')
    cmts_entry.delete(0, 'END')
    
#####################################################################
# Add buttons
button_frame = ttk.LabelFrame(root, text="Commands")
button_frame.pack(fill="x", expand="yes", padx=20)

add_button = ttk.Button(button_frame, text="Add New Client")
add_button.grid(row=0, column=0, padx=10, pady=10)

upd_button = ttk.Button(button_frame, text="Update Client Details", command=update_record)
upd_button.grid(row=0, column=1, padx=10, pady=10)

rmv_button= ttk.Button(
    button_frame, text="Remove Selected Client Details", command=remove_one)
rmv_button.grid(row=0, column=2, padx=10, pady=10)

rmvall_button= ttk.Button(
    button_frame, text="Remove All Clients Details", command=remove_all)
rmvall_button.grid(row=0, column=3, padx=10, pady=10)

rmvall_button= ttk.Button(
    button_frame, text="Clear", command=clear_boxes)
rmvall_button.grid(row=0, column=4, padx=10, pady=10)

#####################################################################
# Bind the treeview
my_tree.bind("<ButtonRelease-1>", select_record)
my_tree.pack()
treeScroll.config(command=my_tree.yview)
load_data()
root.mainloop()

print("77")