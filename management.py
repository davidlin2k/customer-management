# -*- coding: utf-8 -*-
"""
Spyder Editor

This is a temporary script file.
"""


import pandas as pd
import webbrowser as wb
import tkinter as tk
from tkinter import messagebox
from tkinter import Frame
import os
from pandastable.core import Table
from pandastable.data import TableModel
from configparser import ConfigParser

script_dir = os.path.dirname(__file__)

config = ConfigParser()
config.read(os.path.join(script_dir, 'config.ini'))

wb.register('chrome', None, wb.BackgroundBrowser(config.get('Dir', 'Chrome_Dir')))

excel_dir = config.get('Dir', 'Excel_Dir')
data = pd.read_excel(excel_dir, sheet_name = 'Sheet1')
track17 = 'https://t.17track.net/en#nums='

fields = 'Names', 'Address', 'Tracking Number', 'Wechat Number', 'Note'

class MyTable(Table):
    def __init__(self, parent=None, **kwargs):
        Table.__init__(self, parent, **kwargs)
        return
    
def make_table(frame, **kwds):
    df = data
    pt = MyTable(frame, dataframe=df, **kwds )
    pt.show()
    return pt

def test1():
    t = tk.Toplevel()
    fr = Frame(t)
    fr.pack(fill=tk.BOTH,expand=1)
    pt = make_table(fr)
    return

def main_menu():
    menu_window.mainloop()

def add_customer_window():
    root = tk.Tk()
    ents = makeform(root, fields)
    root.after(1, lambda: root.focus_force())
    ents[0][1].focus_set()
    root.bind('<Return>', (lambda event, e=ents: add_customer(root, e))) 
    root.bind('<Escape>', (lambda event : root.destroy()))
    b1 = tk.Button(root, text='Save',
                  command=(lambda e=ents: add_customer(root, e)))
    b1.pack(side=tk.LEFT, padx=5, pady=5)
    b2 = tk.Button(root, text='Quit', command=root.destroy)
    b2.pack(side=tk.LEFT, padx=5, pady=5)
    root.mainloop()
    return

def add_customer(root, entries):
    dict_ = {}
    for entry in entries:
        field = entry[0]
        text  = entry[1].get()
        dict_.update({field : text})
    data.loc[data.index.max()+1] = dict_
    data.to_excel(excel_dir, sheet_name = 'Sheet1', index=False)
    messagebox.showinfo('Info', 'Saved!')
    root.destroy()

def makeform(root, fields):
    entries = []
    for field in fields:
        row = tk.Frame(root)
        lab = tk.Label(row, width=15, text=field, anchor='w')
        ent = tk.Entry(row, width=35)
        row.pack(side=tk.TOP, fill=tk.X, padx=5, pady=5)
        lab.pack(side=tk.LEFT)
        ent.pack(side=tk.RIGHT, expand=tk.YES, fill=tk.X)
        entries.append((field, ent))
    return entries

def remove_customer_window():
    root = tk.Tk()
    ents = makeform(root, ['Name'])
    root.after(1, lambda: root.focus_force())
    ents[0][1].focus_set()
    root.bind('<Return>', (lambda event, e=ents: remove_customer(root, e))) 
    root.bind('<Escape>', (lambda event : root.destroy()))
    b1 = tk.Button(root, text='Remove',
                  command=(lambda e=ents: remove_customer(root, e)))
    b1.pack(side=tk.LEFT, padx=5, pady=5)
    b2 = tk.Button(root, text='Quit', command=root.destroy)
    b2.pack(side=tk.LEFT, padx=5, pady=5)
    root.mainloop()
    return

def remove_customer(root, e):
    data.drop(data.index[data['Names'] == e[0][1].get()], inplace=True)
    data.to_excel(excel_dir, sheet_name = 'Sheet1', index=False)
    messagebox.showinfo('Info', 'Removed!')
    root.destroy()

def track_package_window():
    root = tk.Tk()
    ents = makeform(root, ['Name'])
    root.after(1, lambda: root.focus_force())
    ents[0][1].focus_set()
    root.bind('<Return>', (lambda event, e=ents: track_package(root, e)))
    root.bind('<Escape>', (lambda event : root.destroy()))
    b1 = tk.Button(root, text='Track',
                  command=(lambda e=ents: track_package(root, e)))
    b1.pack(side=tk.LEFT, padx=5, pady=5)
    b2 = tk.Button(root, text='Quit', command=root.destroy)
    b2.pack(side=tk.LEFT, padx=5, pady=5)
    root.mainloop()
    return

def track_package(root, e):
    user_index = data.index[data['Names'] == e[0][1].get()].tolist()
    tracking_num = data.loc[user_index[0]]['Tracking Number']
    wb.get('chrome').open(track17 + tracking_num)
    root.destroy()
    
def display_data():
    root = tk.Tk()
    f = Frame(root)
    root.geometry('1000x600+200+100')
    root.bind('<Escape>', (lambda event : root.destroy()))
    f.pack(fill=tk.BOTH,expand=1)
    pt = make_table(f)
    bp = Frame(root)
    bp.pack(side=tk.TOP)
    
def setting_window():
    root = tk.Tk()
    root.geometry('600x100')
    ents = makeform(root, ['Chomre Path', 'Data Sheet Path'])
    ents[0][1].insert(0, config.get('Dir', 'Chrome_Dir'))
    ents[1][1].insert(0, excel_dir)
    root.after(1, lambda: root.focus_force())
    ents[0][1].focus_set()
    root.bind('<Return>', (lambda event, e=ents: reset_dir(root, e)))   
    root.bind('<Escape>', (lambda event : root.destroy()))
    b1 = tk.Button(root, text='Set',
                  command=(lambda e=ents: reset_dir(root, e)))
    b1.pack(side=tk.LEFT, padx=5, pady=5)
    b2 = tk.Button(root, text='Quit', command=root.destroy)
    b2.pack(side=tk.LEFT, padx=5, pady=5)
    root.mainloop()
    return

def reset_dir(root, e):
    config.set('Dir', 'Chrome_Dir', e[0][1].get())
    config.set('Dir', 'Excel_Dir', e[1][1].get())
    chrome = wb.register('chrome', None, wb.BackgroundBrowser(e[0][1].get()))
    excel_dir = e[1][1].get()
    
    with open(os.path.join(script_dir, 'config.ini'), 'w') as configfile:
        config.write(configfile)
    messagebox.showinfo('Info', 'Saved!')
    root.destroy()
    
def edit_customer_window():
    root = tk.Tk()
    ents = makeform(root, fields)
    root.after(1, lambda: root.focus_force())
    ents[0][1].focus_set()
    root.bind('<Return>', (lambda event, e=ents: edit_customer(root, e)))  
    root.bind('<Escape>', (lambda event : root.destroy()))
    b1 = tk.Button(root, text='Save',
                  command=(lambda e=ents: edit_customer(root, e)))
    b1.pack(side=tk.LEFT, padx=5, pady=5)
    b2 = tk.Button(root, text='Quit', command=root.destroy)
    b2.pack(side=tk.LEFT, padx=5, pady=5)
    root.mainloop()
    return

def edit_customer(root, entries):
    dict_ = {}
    user_index = data.index[data['Names'] == entries[0][1].get()].tolist()
    for entry in entries:
        field = entry[0]
        text  = entry[1].get()
        dict_.update({field : text})
    for k, v in dict_.items():
        if v == '':
            dict_[k] = data.loc[user_index[0]][k]
    data.loc[user_index[0]] = list(dict_.values())
    data.to_excel(excel_dir, sheet_name = 'Sheet1', index=False)
    messagebox.showinfo('Info', 'Saved!')
    root.destroy()
    
menu_window = tk.Tk()
menu_window.geometry('250x280')
ac_button = tk.Button(menu_window, text = 'Add Customer', command = add_customer_window)
ed_button = tk.Button(menu_window, text = 'Edit Customer', command = edit_customer_window)
rm_button = tk.Button(menu_window, text = 'Remove Customer', command = remove_customer_window)
tr_button = tk.Button(menu_window, text = 'Track Package', command = track_package_window)
dis_button = tk.Button(menu_window, text = 'Display Data', command = display_data)
set_button = tk.Button(menu_window, text = 'Setting', command = setting_window)
ac_button.pack(side=tk.TOP, fill=tk.X, expand=tk.NO, padx=10, pady=10)
ed_button.pack(side=tk.TOP, fill=tk.X, expand=tk.NO, padx=10, pady=10)
rm_button.pack(side=tk.TOP, fill=tk.X, expand=tk.NO, padx=10, pady=10)
tr_button.pack(side=tk.TOP, fill=tk.X, expand=tk.NO, padx=10, pady=10)
dis_button.pack(side=tk.TOP, fill=tk.X, expand=tk.NO, padx=10, pady=10)
set_button.pack(side=tk.TOP, fill=tk.X, expand=tk.NO, padx=10, pady=10)

main_menu()



   
