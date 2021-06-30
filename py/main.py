import sys
import re

sys.path.append(r'C:\MSC.Software\Marc\2018.0.0\mentat2018\shlib\win64')
from py_mentat import *
from py_post import *
from tkinter import *
from tkinter.filedialog import *
from tkinter.ttk import Combobox
from collections import OrderedDict, Counter


class GUI(Tk):
    def __init__(self, screenName=None, baseName=None, className='Tk',
                 useTk=1, sync=0, use=None):
        super().__init__(screenName=None, baseName=None, className='Tk', useTk=1, sync=0, use=None)
        self.data = {}
        self.post_file2d = None
        self.post_file3d = None
        self.frame2 = LabelFrame(self, text='Геометрия')
        self.combo2 = Combobox(self.frame2)
        self.combo1 = Combobox(self.frame2)
        self.frame1 = Frame(self)
        self.lbl4 = Label(self.frame1)
        self.lbl3 = Label(self.frame1)

    def run(self):
        self.build_main_widgets()
        self.title('v0')
        self.mainloop()

    def build_main_widgets(self):
        self.frame1.pack()
        btn = Button(self.frame1, text='Обзор', command=lambda: self.open_file('2d'), width=10)
        btn2 = Button(self.frame1, text='Обзор', command=lambda: self.open_file('3d'), width=10)
        lbl = Label(self.frame1, text='Файл решения 2D:')
        lbl2 = Label(self.frame1, text='Файл решения 3D:')
        btn.grid(column=0, row=0)
        btn2.grid(column=0, row=1)
        lbl.grid(column=1, row=0)
        lbl2.grid(column=1, row=1)
        self.lbl3.grid(column=2, row=0, columnspan=10)
        self.lbl4.grid(column=2, row=1, columnspan=10)
        self.frame2.pack()
        lbl5 = Label(self.frame2, text='инкремент')
        lbl6 = Label(self.frame2, text='инкремент')
        lbl5.grid(column=1, row=3)
        self.combo1.grid(column=2, row=3)
        lbl6.grid(column=1, row=4)
        self.combo2.grid(column=2, row=4)
        btn3 = Button(self.frame2, text='Сохранить геометрию', command=lambda: self.save_geometry_2d())
        btn3.grid(column=3, row=3)
        btn4 = Button(self.frame2, text='Сохранить геометрию', command=lambda: self.save_geometry_3d())
        btn4.grid(column=3, row=4)
        btn5 = Button(self.frame2, text='Обновить кол-во инкрементов', command=lambda: self.update_incs())
        btn5.grid(column=2, row=1)

    def open_file(self, task):
        op = askopenfilename(filetypes=(("Binary Post File", "*.t16"), ("all files", "*.*")))
        self.data[task] = {'file': op}
        if task == '2d':
            self.post_file2d = op
            self.lbl3['text'] = self.data[task]['file']
        else:
            self.post_file3d = op
            self.lbl4['text'] = self.data[task]['file']

    def update_incs(self):
        if self.post_file2d:
            self.update_incs_by_task(self.post_file2d, self.combo1)
        if self.post_file3d:
            self.update_incs_by_task(self.post_file3d, self.combo2)

    def update_incs_by_task(self, file_task, combobox):
        p = post_open(file_task)
        incs = [i for i in range(0, p.increments() - 1)]
        combobox['values'] = incs
        combobox.current(max(incs))

    def save_geometry_2d(self):
        inc = int(self.combo1.get())
        border = self.get_border_2d(inc)
        self.save_geometry(border)

    def get_border_2d(self, inc):
        p = post_open(self.post_file2d)
        p.moveto(inc + 1)
        self.nodes = []
        elements = []
        edges = []
        for i in range(0, p.elements()):
            d = dict(id=p.element(i).id, items=p.element(i).items,
                     type=p.element(i).type)
            elements.append(d)
            if elements[i]['type'] == 82:
                elements[i]['items'].pop()
                elements[i]['items'] = list(OrderedDict.fromkeys(elements[i]['items']).keys())
                elements[i]['edges'] = list()
                for n in range(0, len(elements[i]['items'])):
                    node0 = elements[i]['items'][n]
                    node1 = elements[i]['items'][n - 1]
                    edge = [node0, node1]
                    elements[i]['edges'].append(edge)
                    edges.append(edge)
        for element in elements:
            for node_id in element['items']:
                if node_id not in self.nodes:
                    self.nodes.append(node_id)
        edges_tuple = []
        for edge in edges:
            edge.sort()
            edge_tuple = tuple(edge)
            edges_tuple.append(edge_tuple)
        edges = edges_tuple
        repeat = Counter(edges)
        self.border_edges = []
        for edge in edges:
            if repeat[edge] == 1:
                self.border_edges.append(edge)
        return self.getborder_disp(inc, self.post_file2d, self.border_edges)


    def get_elements(self, file_task):
        p = post_open(file_task)
        p.moveto(0)
        return p.elements()

    def save_geometry_3d(self):
        inc = int(self.combo2.get())
        self.get_border_2d(1)
        self.total_rep = int(self.get_elements(self.post_file3d) / self.get_elements(self.post_file2d))
        border_edges = []
        for edge in self.border_edges:
            edg = []
            for node in edge:
                nod = node * self.total_rep * 2 - self.total_rep + node + len(self.nodes)
                edg.append(nod)
            border_edges.append(edg)
        self.getborder_disp(inc, self.post_file3d, border_edges)
        border = self.getborder_disp(inc, self.post_file3d, border_edges)
        self.save_geometry(border)

    def getborder_disp(self, inc, post_file, border_edges):
        p = post_open(post_file)
        p.moveto(inc + 1)
        border_edges_xy = []
        nodes = []
        for i in range(0, p.nodes()):
            nodes.append(p.node(i).id)
        for edge in border_edges:
            edge_xy = []
            for node_id in edge:
                have_disp = p.node_displacements()
                if have_disp:
                    dx, dy, dz = p.node_displacement(nodes.index(node_id))
                else:
                    dx, dy, dz = 0, 0, 0
                edge_xy.append(p.node(nodes.index(node_id)).x + float(dx))
                edge_xy.append(p.node(nodes.index(node_id)).y + float(dy))
                edge_xy.append(p.node(nodes.index(node_id)).z + float(dz))
            border_edges_xy.append(edge_xy)
        return border_edges_xy

    def save_geometry(self, border):
        border_string = str()
        for line in border:
            border_string += str(line) + '\n'
        border_string = re.sub('[[]', '', border_string)
        border_string = re.sub('[]]', '', border_string)
        file_name = asksaveasfilename(filetypes=(("txt files", "*.txt"), ("all files", "*.*")))
        f = open(file_name, 'w')
        f.write(border_string)
        f.close()


if __name__ == '__main__':
    GUI().run()
