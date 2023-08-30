# -*- coding: utf-8 -*-
"""
Created on Wed Aug 23 01:51:27 2023

@author: H
"""

# IMPORTS
from tkinter import *
from tkinter import ttk
# END IMPORTS

# This project attempts to create a small customer relationship management system
# The two main feature at first are:
# Easily view, edit, add and remove clients and their info from a simple database, using a minimal GUI
# Allow the population of the client data by importing Excel sheets
# Add an option for sending WhatsApp messages and notifications to customers

# LOG
# -23AUG--CREATED GUI INFRASTRUCTURE--NEED TO REVIEW CONCEPTS--SOME CODE WAS COPIED FROM WEB
# -28AUG--ADDED BUTTONS AND ENTRY BOXES--NOT YET FUNCTIONAL
# -30AUG--UPDATED SOME BUTTONS FUNCTIONALITY
# LOG


# First Step, create data structure

# Style
root = Tk()
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
tree_frame = Frame(root)
tree_frame.pack(pady=10)

tree_scroll = Scrollbar(tree_frame)
tree_scroll.pack(side=RIGHT, fill=Y)

# Create the treeview

my_tree = ttk.Treeview(
    tree_frame, yscrollcommand=tree_scroll.set, selectmode="extended")
my_tree.pack()

tree_scroll.config(command=my_tree.yview)


# Creating Columns
my_tree['columns'] = ("Name", "Address", "Area", "Phone",
                      "Phone 2", "Tax Status", "Fees Status", "Notes")

my_tree.column("#0", width=0, stretch=NO)
my_tree.column("Name", anchor=W, width=140)
my_tree.column("Address", anchor=W, width=140)
my_tree.column("Area", anchor=CENTER, width=140)
my_tree.column("Phone", anchor=CENTER, width=140)
my_tree.column("Phone 2", anchor=CENTER, width=140)
my_tree.column("Tax Status", anchor=CENTER, width=140)
my_tree.column("Fees Status", anchor=CENTER, width=140)
my_tree.column("Notes", anchor=W, width=140)

# Creating Headings
my_tree.heading("#0", text="", anchor=W)
my_tree.heading("Name", text="Name", anchor=W)
my_tree.heading("Address", text="Address", anchor=W)
my_tree.heading("Area", text="Area", anchor=W)
my_tree.heading("Phone", text="Phone", anchor=W)
my_tree.heading("Phone 2", text="Phone 2", anchor=W)
my_tree.heading("Tax Status", text="Tax Status", anchor=W)
my_tree.heading("Fees Status", text="Fees Status", anchor=W)
my_tree.heading("Notes", text="Notes", anchor=W)

data = [
    ["Hesham", "October", "Juhayna", "01111",
        "0555", "Good", "Paid", "Call Next Month"],
    ["Hamada", "October", "Zayed", "01311", "0545", "Need", "Paid", "Call Now"]
]

# Create styriped rows
my_tree.tag_configure('oddrow', background="white")
my_tree.tag_configure('evenrow', background="lightblue")

#############################################################################

# Adding data to screen
global count
count = 0

for record in data:
    if count % 2 == 0:
        my_tree.insert(parent='', index='end', iid=count, text='', values=(
            record[0], record[1], record[2], record[3], record[4], record[5], record[6], record[7]), tags=('evenrow',))
    else:
        my_tree.insert(parent='', index='end', iid=count, text='', values=(
            record[0], record[1], record[2], record[3], record[4], record[5], record[6], record[7]), tags=('oddrow',))

    count += 1


# Add record entry boxes
data_frame = LabelFrame(root, text="Record")
data_frame.pack(fill="x", expand="yes", padx=20)


###
name_label = Label(data_frame, text="Name")
name_label.grid(row=0, column=0, padx=10, pady=10)

name_entry = Entry(data_frame)
name_entry.grid(row=0, column=1, padx=10, pady=10)
###


adrs_label = Label(data_frame, text="Address")
adrs_label.grid(row=0, column=2, padx=10, pady=10)

adrs_entry = Entry(data_frame)
adrs_entry.grid(row=0, column=3, padx=10, pady=10)
###


area_label = Label(data_frame, text="Area")
area_label.grid(row=0, column=4, padx=10, pady=10)

area_entry = Entry(data_frame)
area_entry.grid(row=0, column=5, padx=10, pady=10)
###


ph_label = Label(data_frame, text="Phone")
ph_label.grid(row=1, column=0, padx=10, pady=10)

ph_entry = Entry(data_frame)
ph_entry.grid(row=1, column=1, padx=10, pady=10)
###


ph2_label = Label(data_frame, text="Phone 2")
ph2_label.grid(row=1, column=2, padx=10, pady=10)

ph2_entry = Entry(data_frame)
ph2_entry.grid(row=1, column=3, padx=10, pady=10)
###


tax_label = Label(data_frame, text="Tax Status")
tax_label.grid(row=1, column=4, padx=10, pady=10)

tax_entry = Entry(data_frame)
tax_entry.grid(row=1, column=5, padx=10, pady=10)
###


fees_label = Label(data_frame, text="Office Fees Status")
fees_label.grid(row=1, column=6, padx=10, pady=10)

fees_entry = Entry(data_frame)
fees_entry.grid(row=1, column=7, padx=10, pady=10)
###


cmts_label = Label(data_frame, text="Notes")
cmts_label.grid(row=0, column=6, padx=10, pady=10)

cmts_entry = Entry(data_frame)
cmts_entry.grid(row=0, column=7, padx=10, pady=10)
###

#####################################################################

# Select Record


def select_record(e):
    # first clear the entry boxes
    name_entry.delete(0, END)
    adrs_entry.delete(0, END)
    area_entry.delete(0, END)
    ph_entry.delete(0, END)
    ph2_entry.delete(0, END)
    tax_entry.delete(0, END)
    fees_entry.delete(0, END)
    cmts_entry.delete(0, END)

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
    ), ph_entry.get(), ph2_entry.get(), tax_entry.get(), fees_entry.get(), cmts_entry.get()))

    # first clear the entry boxes
    name_entry.delete(0, END)
    adrs_entry.delete(0, END)
    area_entry.delete(0, END)
    ph_entry.delete(0, END)
    ph2_entry.delete(0, END)
    tax_entry.delete(0, END)
    fees_entry.delete(0, END)
    cmts_entry.delete(0, END)


#####################################################################
# Add buttons
button_frame = LabelFrame(root, text="Commands")
button_frame.pack(fill="x", expand="yes", padx=20)

add_button = Button(button_frame, text="Add New Client")
add_button.grid(row=0, column=0, padx=10, pady=10)

upd_button = Button(button_frame, text="Update Client Details", command=update_record
upd_button.grid(row=0, column=1, padx=10, pady=10)

rmv_button=Button(
    button_frame, text="Remove Selected Client Details", command=remove_one)
rmv_button.grid(row=0, column=2, padx=10, pady=10)

rmvall_button=Button(
    button_frame, text="Remove All Clients Details", command=remove_all)
rmvall_button.grid(row=0, column=3, padx=10, pady=10)

slct_button=Button(button_frame, text="Select Record", command=select_record)
slct_button.grid(row=0, column=4, padx=10, pady=10)

#####################################################################
# Bind the treeview
my_tree.bind("<ButtonRelease-1>", select_record)
