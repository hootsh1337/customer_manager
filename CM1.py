# -*- coding: utf-8 -*-
"""
Created on Wed Aug 23 01:51:27 2023

@author: H
"""

#IMPORTS
from tkinter import *
from tkinter import ttk
#END IMPORTS

#This project attempts to create a small customer relationship management system
#The two main feature at first are: 
    #Easily view, edit, add and remove clients and their info from a simple database, using a minimal GUI
    #Allow the population of the client data by importing Excel sheets
    #Add an option for sending WhatsApp messages and notifications to customers
    
#LOG
#-23AUG--CREATED GUI INFRASTRUCTURE--NEED TO REVIEW CONCEPTS--SOME CODE WAS COPIED FROM WEB
#LOG
    
    
#First Step, create data structure

#Style
root = Tk()
root.title = ('CRM Test')
root.geometry = ("1000x500")
root.iconbitmap = ('F:\Art\Icons\pc.ico')
style = ttk.Style()

style.theme_use('default')

#Treeview colors
style.configure("Treeview", background="#D3D3D3",foreground="black",rowheight=25,fieldbackground="#D3D3D3")

#Change color of selected client
style.map('Treeview', background=[('selected',"347083")])

#Frame and scrollbar
tree_frame = Frame(root)
tree_frame.pack(pady=10)

tree_scroll = Scrollbar(tree_frame)
tree_scroll.pack(side=RIGHT, fill=Y)

#Create the treeview

my_tree = ttk.Treeview(tree_frame, yscrollcommand=tree_scroll.set , selectmode="extended")
my_tree.pack()

tree_scroll.config(command=my_tree.yview)


#Creating Columns
my_tree['columns'] = ("Name","Address","Area","Phone","Phone 2","Tax Status","Fees Status","Notes")

my_tree.column("#0", width=0, stretch=NO)
my_tree.column("Name", anchor=W, width=140)
my_tree.column("Address", anchor=W, width=140)
my_tree.column("Area", anchor=CENTER, width=140)
my_tree.column("Phone", anchor=CENTER, width=140)
my_tree.column("Phone 2", anchor=CENTER, width=140)
my_tree.column("Tax Status", anchor=CENTER, width=140)
my_tree.column("Fees Status", anchor=CENTER, width=140)
my_tree.column("Notes", anchor=W, width=140)

#Creating Headings
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
        ["Hesham","October","Juhayna","01111","0555","Good","Paid","Call Next Month"],
        ["Hamada","October","Zayed","01311","0545","Need","Paid","Call Now"]
        ]

#Create styriped rows
my_tree.tag_configure('oddrow', background="white")
my_tree.tag_configure('evenrow', background="lightblue")

#############################################################################

#Adding data to screen
global count
count=0

for record in data:
    if count % 2 == 0:
        my_tree.insert(parent='',index='end',iid=count,text='',values=(record[0],record[1],record[2],record[3],record[4],record[5],record[6],record[7]),tags=('evenrow',))
    else:
        my_tree.insert(parent='',index='end',iid=count,text='',values=(record[0],record[1],record[2],record[3],record[4],record[5],record[6],record[7]),tags=('oddrow',))
        
    count += 1