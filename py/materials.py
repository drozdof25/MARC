#! /usr/bin/env python
# -*- coding: utf-8 -*
import os
from py_mentat import*
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
        self.changing = False
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
        button_save = ttk.Button(button_frame, text='save_data', command=lambda: self.save_data())
        button_save.pack(side='left')
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
            self.add_row_data(material,self.materials.index(material))
        self.changing = False
        self.scrollable_body.update()
        self.conn_on = True
    def add(self):
        if not self.changing:
            data_dict = dict()
            r = len(self.materials)
            for key in self.data_keys:
                data_dict[key] = ''
            self.add_row_data(data_dict,r)
            for key, value in self.lbl_list[r].items():
                value[0].grid_remove()
                value[1].grid()
                value[3].grid_remove()
                value[4].grid()
                if self.data_keys.index(key) in range(6,9):
                    value[1].grid_remove()
                if self.data_keys.index(key)==8:
                    btn = ttk.Button(self.scrollable_body, text='одноосное растяжение', width=39,
                            command=lambda row =r: self.open_file_odnoos(row))
                    btn.grid(column=6, row=r, padx=1, pady=1, columnspan=3)
                    value.append(btn)
    def open_file_odnoos(self,r):
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
        c10 = py_get_float('mater_par(%s,structural:mooney_c10)' % mater_name)
        c01 = py_get_float('mater_par(%s,structural:mooney_c01)' % mater_name)
        py_send('*remove_current_mater')
        py_send('*remove_current_table')
        bulk_modulus = (c10+c01)*10000
        f= open(op)
        odnoos = []
        for line in f:
            l = line.split()
            odnoos.append(l)
        self.lbl_list[r]['C10'][2].set(round(c10,6))
        self.lbl_list[r]['C10'][1].grid()
        self.lbl_list[r]['C01'][2].set(round(c01,6))
        self.lbl_list[r]['C01'][1].grid()
        self.lbl_list[r]['K'][2].set(round(bulk_modulus,2))
        self.lbl_list[r]['K'][5].grid_remove()
        self.lbl_list[r]['K'][1].grid()
        self.lbl_list[r]['K'].append(odnoos)
    def edit(self,index):
        self.changing = True
        for key,value in self.lbl_list[index].items():
            value[0].grid_remove()
            value[1].grid()
            value[3].grid_remove()
            value[4].grid()
    def save_change(self,index):
        self.changing = False
        data_dict = dict()
        for key,value in self.lbl_list[index].items():
            value[0].grid()
            value[1].grid_remove()
            value[3].grid()
            value[4].grid_remove()
            if index < len(self.materials):
                self.materials[index][key] = value[2].get()
            else:
                data_dict[key] = value[2].get()
                if self.data_keys.index(key)==8:
                    data_dict['Одноосное растяжение'] = value[6]
        if not index < len(self.materials):
            self.materials.append(data_dict)
    def cancel_change(self,index):
        self.changing = False
        if index < len(self.materials):
            for key, value in self.lbl_list[index].items():
                value[0].grid()
                value[1].grid_remove()
                value[2].set(self.materials[index][key])
                value[3].grid()
                value[4].grid_remove()
        else:
            for key, value in self.lbl_list[index].items():
                value[0].grid_forget()
                value[1].grid_forget()
                value[3].grid_forget()
                value[4].grid_forget()
                if self.data_keys.index(key) == 8:
                    value[5].grid_forget()
            self.lbl_list.pop(index)
    def delete(self,index):
        self.materials.pop(index)
        self.data_1.pack_forget()
        self.conn()
    def add_row_data(self,data_dict,r):
        self.changing = True
        object_dict = dict()
        for key in self.data_keys:
            data = str(data_dict[key])
            if len(data) > 10:
                data = data.replace(' ', '\n')
            strvar = StringVar()
            strvar.set(data)
            lbl = tk.Label(self.scrollable_body, textvariable=strvar, width=13, height=2,
            entry = tk.Entry(self.scrollable_body, textvariable=strvar, width=13)
            entry.grid(column=self.data_keys.index(key), row=r)
            entry.grid_remove()
            lbl.grid(column=self.data_keys.index(key), row=r)
            action_frame1 = ttk.Frame(self.scrollable_body)
            action_frame1.grid(column=9, row=r)
            action_frame2 = ttk.Frame(self.scrollable_body)
            action_frame2.grid(column=9, row=r)
            action_frame2.grid_remove()
            object_dict[key] = [lbl, entry, strvar, action_frame1, action_frame2]
            btn_edit = tk.Button(action_frame1, text='edit', width=7,
                                 command=lambda id=r: self.edit(id))
            btn_del = tk.Button(action_frame1, text='del', width=7,
                                command=lambda id=r: self.delete(id))
            btn_save = tk.Button(action_frame2, text='save', width=7,
                                 command=lambda id=r: self.save_change(id))
            btn_cancel = tk.Button(action_frame2, text='cancel', width=7,
                                   command=lambda id=r: self.cancel_change(id))
            btn_edit.pack(side='left')
            btn_del.pack(side='left')
            btn_save.pack(side='left')
            btn_cancel.pack(side='left')
        self.lbl_list.append(object_dict)
    def save_data(self):
        with open("data_file.json", "w") as write_file:
            json.dump(self.materials, write_file)
    def add_to_marc(self,index):
        py_send('*new_md_table 1 1')
        table_name = py_get_string('table_name()')
MaterialWidget().mainloop()