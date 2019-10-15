#! /usr/bin/env python
# -*- coding: utf-8 -*
import os
#from py_mentat import*
import tkinter as tk
from tkinter import ttk
import json
from tkinter.filedialog import *


class Scrollable(tk.Frame):
    """
       Make a frame scrollable with scrollbar on the right.
       After adding or removing widgets to the scrollable frame,
       call the update() method to refresh the scrollable area.
    """

    def __init__(self, frame, width=16):

        scrollbar = tk.Scrollbar(frame, width=width)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y, expand=False)

        self.canvas = tk.Canvas(frame, yscrollcommand=scrollbar.set)
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        scrollbar.config(command=self.canvas.yview)

        self.canvas.bind('<Configure>', self.__fill_canvas)

        # base class initialization
        tk.Frame.__init__(self, frame)

        # assign this obj (the inner frame) to the windows item of the canvas
        self.windows_item = self.canvas.create_window(0,0, window=self, anchor=tk.NW)


    def __fill_canvas(self, event):
        "Enlarge the windows item to the canvas width"
        canvas_width = event.width
        self.canvas.itemconfig(self.windows_item, width = canvas_width)

    def update(self):
        "Update the canvas and the scrollregion"
        self.update_idletasks()
        self.canvas.config(scrollregion=self.canvas.bbox(self.windows_item))

class MaterialWidget():
    def __init__(self):
        self.root = tk.Tk()
        self.root.event_add('<<Paste>>', '<Control-igrave>')
        self.root.event_add("<<Copy>>", "<Control-ntilde>")
        self.data_keys =['Тип', 'Шифр(дата)', 'Шифр(полный)', 'Назначение',
                    'Массовая плотность', 'Твёрдость по шору', 'C10', 'C01', 'K']
        self.conn_on = False

    def mainloop(self):
        tabControl = ttk.Notebook(self.root)
        tab1 = ttk.Frame(tabControl)
        tab2 = ttk.Frame(tabControl)
        tabControl.add(tab1, text='легковая шина')
        tabControl.add(tab2, text='ЦМК')
        tabControl.pack(expand=1, fill="both")
        self.body_1 = tk.LabelFrame(tab1, text='База данных')
        self.body_1.pack(side='left')
        body_2 = ttk.LabelFrame(tab1, text='Задача')
        body_2.pack(side='right')
        data_titles = ['Тип', 'Шифр\n(дата)', 'Шифр\n(полный)', 'Назначение',
                       'Массовая \n плотность', 'Твёрдость \nпо шору', 'C10', 'C01', 'K']
        button_frame = ttk.Frame(self.body_1)
        button_frame.pack(fill='x', expand=True)

        button_connect = ttk.Button(button_frame, text='connect',command = lambda :self.conn())
        button_connect.pack(side='left')
        button_add = ttk.Button(button_frame, text='add+',command = lambda :self.add())
        button_add.pack(side='right')
        title_1 = ttk.Frame(self.body_1)
        title_1.pack()

        for title in data_titles:
            lbl = tk.Label(title_1, bg="lightgreen", text=title, width=13, height=2, borderwidth=2, relief="solid")
            lbl.pack(side='left', anchor='n')

        lbl = tk.Label(title_1,width=20, height=2)
        lbl.pack(side='left', anchor='n')


        self.root.mainloop()

    def conn(self):
        if not self.conn_on:
            with open("data_file.json", "r") as read_file:
                self.materials = json.load(read_file)
        self.data_1 = tk.Frame(self.body_1)
        self.data_1.pack(fill=tk.BOTH, expand=True)
        self.scrollable_body = Scrollable(self.data_1)
        self.lbl_list = []
        for material in self.materials:
            mat_lbl_dict = dict()
            for key in self.data_keys:
                data = str(material[key])
                if len(data) > 10:
                    data = data.replace(' ', '\n')
                strvar = StringVar()
                strvar.set(data)
                lbl = tk.Label(self.scrollable_body, textvariable=strvar, width=13, height=2, borderwidth=2, relief="groove")
                entry = tk.Entry(self.scrollable_body,textvariable=strvar, width=13)
                entry.grid(column=self.data_keys.index(key), row=self.materials.index(material))
                entry.grid_remove()
                lbl.grid(column=self.data_keys.index(key), row=self.materials.index(material))
                action_frame1 = ttk.Frame(self.scrollable_body)
                action_frame1.grid(column=9,row=self.materials.index(material))
                action_frame2 = ttk.Frame(self.scrollable_body)
                action_frame2.grid(column=9, row=self.materials.index(material))
                action_frame2.grid_remove()
                mat_lbl_dict[key] = [lbl, entry, strvar,action_frame1,action_frame2]
                btn_edit = tk.Button(action_frame1, text='edit', width=7,command= lambda id = self.materials.index(material): self.edit(id))
                btn_del = tk.Button(action_frame1, text='del', width=7,command= lambda id = self.materials.index(material): self.delete(id))
                btn_save = tk.Button(action_frame2, text='save', width=7,command= lambda id = self.materials.index(material): self.save_change(id))
                btn_cancel = tk.Button(action_frame2, text='cancel', width=7,command= lambda id = self.materials.index(material): self.cancel_change(id))
                btn_edit.pack(side='left')
                btn_del.pack(side='left')
                btn_save.pack(side='left')
                btn_cancel.pack(side='left')
            self.lbl_list.append(mat_lbl_dict)
        self.scrollable_body.update()
        self.conn_on = True
    def add(self):
        if self.conn_on:
            self.entry_1 = tk.Entry(self.scrollable_body, width=13)
            self.entry_2 = tk.Entry(self.scrollable_body, width=13,)
            self.entry_3 = tk.Entry(self.scrollable_body, width=13)
            self.entry_4 = tk.Entry(self.scrollable_body, width=13)
            self.entry_5 = tk.Entry(self.scrollable_body, width=13)
            self.entry_6 = tk.Entry(self.scrollable_body, width=13)
            self.entry_1.grid(column=0, row=len(self.materials), padx=1,pady=1)
            self.entry_2.grid(column=1, row=len(self.materials), padx=1, pady=1)
            self.entry_3.grid(column=2, row=len(self.materials), padx=1, pady=1)
            self.entry_4.grid(column=3, row=len(self.materials), padx=1, pady=1)
            self.entry_5.grid(column=4, row=len(self.materials), padx=1, pady=1)
            self.entry_6.grid(column=5, row=len(self.materials), padx=1, pady=1)
            self.btn = ttk.Button(self.scrollable_body,text='одноосное растяжение',width =39,command = lambda : self.open_file_odnoos())
            self.btn.grid(column=6, row=len(self.materials), padx=1, pady=1,columnspan=3)
    def open_file_odnoos(self):
        op = askopenfilename(filetypes = (("text files","*.txt"),("all files","*.*")))
        py_send('@popup(table_pm,0) *md_table_read_any "%s"' % op)
        table_name = py_get_string('table_name()')
        py_send('*edit_table %s' % table_name)
        py_send('*set_md_table_type 1')
        py_send('experimental_data')
        py_send('*set_xcurve_method_time_indep least_squares_ds')
        py_send('*xcv_table uniaxial %s' % table_name)
        py_send('*xcv_model mooney2')
        py_send('*xcv_mode uniaxial on')
        py_send('*xcv_positive on')
        py_send('*xcv_extrapolation on')
        py_send('*xcv_compute')
        py_send('*xcv_create')
        mater_name = py_get_string('mater_name()')
        self.c10 = py_get_float('mater_par(%s,structural:mooney_c10)' % mater_name)
        self.c01 = py_get_float('mater_par(%s,structural:mooney_c01)' % mater_name)
        self.bulk_modulus = (self.c10+self.c01)*10000
        f= open(op)
        odnoos = []
        for line in f:
            l = line.split()
            odnoos.append(l)
        self.btn.grid_remove()
        lbl_c10 = tk.Label(self.scrollable_body, text=str(round(self.c10 ,7)), width=13, height=2, borderwidth=2,
                           relief="groove")
        lbl_c10.grid(column=6, row=len(self.materials))
        lbl_c01 = tk.Label(self.scrollable_body, text=str(round(self.c01, 7)), width=13, height=2, borderwidth=2,
                           relief="groove")
        lbl_c01.grid(column=7, row=len(self.materials))
        lbl_bulk = tk.Label(self.scrollable_body, text=str(round(self.bulk_modulus, 2)), width=13, height=2, borderwidth=2,
                           relief="groove")
        lbl_bulk.grid(column=8, row=len(self.materials))
        py_send('*remove_current_mater')
        py_send('*remove_current_table')
        self.action_frame = ttk.Frame(self.scrollable_body)
        self.action_frame.grid(column=9, row=len(self.materials))
        btn_save = tk.Button(self.action_frame,text ='save',width =7,command = lambda : self.save())
        btn_cancel = tk.Button(self.action_frame,text ='cancel',width =7)
        btn_save.pack(side = 'left')
        btn_cancel.pack(side = 'left')
    def save(self):
        type = self.entry_1.get()
        shifr_date = self.entry_2.get()
        shifr =self.entry_3.get()
        destination = self.entry_4.get()
        mass_density = self.entry_5.get()
        hard_shor =self.entry_6.get()
        self.entry_1.grid_remove()
        self.entry_2.grid_remove()
        self.entry_3.grid_remove()
        self.entry_4.grid_remove()
        self.entry_5.grid_remove()
        self.entry_6.grid_remove()
        self.lbl_1 = tk.Label(self.scrollable_body, text=type, width=13, height=2, borderwidth=2,
                           relief="groove")
        self.lbl_2 =tk.Label(self.scrollable_body, text=shifr_date, width=13, height=2, borderwidth=2,
                           relief="groove")
        self.lbl_3 =tk.Label(self.scrollable_body, text=shifr, width=13, height=2, borderwidth=2,
                           relief="groove")
        self.lbl_4 = tk.Label(self.scrollable_body, text=destination, width=13, height=2, borderwidth=2,
                           relief="groove")
        self.lbl_5 = tk.Label(self.scrollable_body, text= mass_density, width=13, height=2, borderwidth=2,
                           relief="groove")
        self.lbl_6 = tk.Label(self.scrollable_body, text= hard_shor, width=13, height=2, borderwidth=2,
                           relief="groove")
        self.lbl_1.grid(column=0, row=len(self.materials))
        self.lbl_2.grid(column=1, row=len(self.materials))
        self.lbl_3.grid(column=2, row=len(self.materials))
        self.lbl_4.grid(column=3, row=len(self.materials))
        self.lbl_5.grid(column=4, row=len(self.materials))
        self.lbl_6.grid(column=5, row=len(self.materials))
        self.action_frame.grid_remove()
        material = {
            'Тип': type,
            'Шифр(дата)': shifr_date,
            'Шифр(полный)': shifr,
            'Назначение': destination,
            'Массовая плотность': mass_density,
            'Твёрдость по шору': hard_shor,
            'C10': self.c10,
            'C01': self.c01,
            'K': self.bulk_modulus
        }
        self.materials.append(material)
    def edit(self,index):
        for key,value in self.lbl_list[index].items():
            value[0].grid_remove()
            value[1].grid()
            value[3].grid_remove()
            value[4].grid()
    def save_change(self,index):
        for key,value in self.lbl_list[index].items():
            value[0].grid()
            value[1].grid_remove()
            value[3].grid()
            value[4].grid_remove()
            self.materials[index][key] = value[2].get()
    def cancel_change(self,index):
        for key, value in self.lbl_list[index].items():
            value[0].grid()
            value[1].grid_remove()
            value[2].set(self.materials[index][key])
            value[3].grid()
            value[4].grid_remove()
    def delete(self,index):
        self.materials.pop(index)
        self.data_1.pack_forget()
        self.conn()
MaterialWidget().mainloop()