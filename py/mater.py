#! /usr/bin/env python
# -*- coding: utf-8 -*-

import os
#from py_mentat import*
from tkinter import*
from tkinter import ttk
import time

def add_material(event,directory,file):
    file_tbl = directory+'\\'+file
    py_send('@popup(table_pm,0) *md_table_read_any "%s"' % file_tbl)
    file = file.replace('.txt','')
    file = file.replace(' ','')
    py_send('*edit_table %s' % file)
    py_send('*set_md_table_type 1')
    py_send('experimental_data')
    py_send('*set_xcurve_method_time_indep least_squares_ds')
    py_send('*xcv_table uniaxial %s' %file)
    py_send('*xcv_model mooney2')
    py_send('*xcv_mode uniaxial on')
    py_send('*xcv_positive on')
    py_send('*xcv_extrapolation on')
    py_send('*xcv_compute')
    py_send('*xcv_create')
    mater_name = py_get_string('mater_name()')
    py_send('*edit_mater %s' % mater_name)
    py_send('*mater_name %s' % file)



data_directory  = 'D:\\buffer\\data'


root = Tk()
tabControl = ttk.Notebook(root)
tab1 = ttk.Frame(tabControl)
tab2 = ttk.Frame(tabControl)
tabControl.add(tab1, text='легковая шина')
tabControl.add(tab2, text='ЦМК')
tabControl.pack(expand=1, fill="both")

def add_all(files):
    for file in files:
        add_material(None,data_directory,file)
def connect():
    files = os.listdir(data_directory)
    for file in files:
        lbl = Label(tab1,text = file)
        lbl.grid(column = 0,row = files.index(file)+2)
        stat = os.stat(data_directory +'\\'+file)
        change_time = time.ctime(stat.st_atime)
        lbl = Label(tab1, text=change_time)
        lbl.grid(column=1, row=files.index(file) + 2)
        btn = Button(tab1,text = 'Добавить')
        btn.bind('<Button-1>',lambda event,d =data_directory,f = file: add_material(event,d,f))
        btn.grid(column=2, row=files.index(file) + 2)
    btn = Button(tab1,text = 'Добавить все', command = lambda f = files: add_all(f))
    btn.grid(column=2, row= len(files)+ 2)
button = Button(tab1,text = 'connect',command = lambda :connect() )
button.grid(column=0,row=0,padx=5,pady=5)

label = Label(tab1,text = 'Материал:')
label.grid(column=0,row=1,padx=5,pady=5)

label = Label(tab1,text = 'Изменен:')
label.grid(column=1,row=1,padx=5,pady=5)
def test():
    pass
btn = Button(tab1,text = 'test', command = lambda : test())
btn.grid(column=1,row=20,padx=5,pady=5)
def test():
    name = py_get_string('mater_name_index(1)')
    out = py_get_float('mater_par(Rezina_19-320-3_germosloi,structural:mooney_c01)')
    print(out)
root.title('Материалы')
root.geometry('1000x1000')
root.mainloop()
