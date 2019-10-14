import os
#from py_mentat import*
import tkinter as tk
from tkinter import ttk
import json



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
        self.data_req =['Тип', 'Шифр(дата)', 'Шифр(полный)', 'Назначение',
                    'Массовая плотность', 'Твёрдость по шору', 'C10', 'C01', 'K']
        self.conn_on = False

    def mainloop(self):
        tabControl = ttk.Notebook(self.root)
        tab1 = ttk.Frame(tabControl)
        tab2 = ttk.Frame(tabControl)
        tabControl.add(tab1, text='легковая шина')
        tabControl.add(tab2, text='ЦМК')
        tabControl.pack(expand=1, fill="both")
        body_1 = ttk.LabelFrame(tab1, text='База данных')
        body_1.pack(side='left')
        body_2 = ttk.LabelFrame(tab1, text='Задача')
        body_2.pack(side='right')
        data_titles = ['Тип', 'Шифр\n(дата)', 'Шифр\n(полный)', 'Назначение',
                       'Массовая \n плотность', 'Твёрдость \nпо шору', 'C10', 'C01', 'K']
        labels_titles = []
        button_frame = ttk.Frame(body_1)
        button_frame.pack(fill='x', expand=True)

        button_connect = ttk.Button(button_frame, text='connect',command = lambda :self.conn())
        button_connect.pack(side='left')
        button_add = ttk.Button(button_frame, text='add+',command = lambda :self.add())
        button_add.pack(side='right')
        title_1 = ttk.Frame(body_1)
        title_1.pack()

        for title in data_titles:
            lbl = tk.Label(title_1, bg="lightgreen", text=title, width=13, height=2, borderwidth=2, relief="solid")
            labels_titles.append(lbl)
            lbl.pack(side='left', anchor='n')

        self.data_1 = tk.Frame(body_1)
        self.data_1.pack(fill=tk.BOTH, expand=True)
        self.root.mainloop()

    def conn(self):
        if not self.conn_on:
            with open("data_file.json", "r") as read_file:
                self.materials = json.load(read_file)
            self.scrollable_body = Scrollable(self.data_1)
            for material in self.materials:
                for req in self.data_req:
                    data = str(material[req])
                    if len(data) > 10:
                        data = data.replace(' ', '\n')
                    lbl = tk.Label(self.scrollable_body, text=data, width=13, height=2, borderwidth=2, relief="groove")
                    lbl.grid(column=self.data_req.index(req), row=self.materials.index(material))
            self.scrollable_body.update()
            self.conn_on = True
    def add(self):
        if self.conn_on:
            type = tk.StringVar()
            shifr_date = tk.StringVar()
            shifr_full = tk.StringVar()
            destination = tk.StringVar()
            mass_density = tk.StringVar()
            hardness = tk.StringVar()
            entry_1 = Entry(textvariable=type)
MaterialWidget().mainloop()