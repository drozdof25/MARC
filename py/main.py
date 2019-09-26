import sys
import re
sys.path.append('C:\\MSC.Software\\Marc\\2018.0.0\\mentat2018\\shlib\\win64')
from py_mentat import*
from py_post import*
from tkinter import*
from tkinter.filedialog import *
from tkinter.ttk import Combobox
from collections import OrderedDict,Counter


class GUI(Tk):
    #window.geometry('400x250')
    #chk_state.set(True)
    #chk = Checkbutton(window, text='Посадка', var=chk_state)
    #chk2 = Checkbutton(window, text='Обжатие', var=chk_state)
    #chk3 = Checkbutton(window, text='Качение', var=chk_state)
    #chk.grid(column = 0,row = 0)
    #chk2.grid(column = 0,row = 1)
    #chk3.grid(column=0, row=2)
    #window.mainloop()
    #print(window.grid)
    #py_send('*add_points 10 130 0')
    def run(self):
        self.build_main_widgets()
        self.title('v0')
        self.mainloop()
    def build_main_widgets(self):
        btn = Button(self,text = 'Обзор',command = lambda :self.open_file('2d'))
        btn2 = Button(self,text = 'Обзор',command = lambda :self.open_file('3d'))
        lbl = Label(self, text = 'Файл решения 2D:')
        lbl2 = Label(self, text='Файл решения 3D:')
        self.lbl3 = Label(self)
        self.lbl4 = Label(self)
        btn.grid(column=0,row =1)
        btn2.grid(column=0, row=2)
        lbl.grid(column=1, row=1)
        lbl2.grid(column=1, row=2)
        self.lbl3.grid(column=2, row=1,columnspan = 15)
        self.lbl4.grid(column=2, row=2,columnspan = 15)

        self.check_state1 = BooleanVar()
        self.check_state2 = BooleanVar()
        self.check_state3 = BooleanVar()
        self.check_state4 = BooleanVar()
        chk1 = Checkbutton(self,text = 'Посадка 2D',variable =self.check_state1)
        chk2 = Checkbutton(self, text='Посадка 3D', variable=self.check_state2)
        chk3 = Checkbutton(self, text='Обжатие 3D',variable =self.check_state3)
        chk4 = Checkbutton(self, text='Качение 3D',variable =self.check_state4)
        chk1.grid(column = 0,row =3)
        chk2.grid(column=0, row=4)
        chk3.grid(column=0, row=5)
        chk4.grid(column=0, row=6)

        lbl5 = Label(self,text = 'инкремент')
        self.combo1 = Combobox(self)
        lbl6 = Label(self, text='инкремент')
        self.combo2 = Combobox(self)
        lbl7 = Label(self, text='инкремент')
        self.combo3 = Combobox(self)
        lbl8 = Label(self, text='инкремент')
        self.combo4 = Combobox(self)

        lbl5.grid(column = 1,row =3)
        self.combo1.grid(column=2, row=3)
        lbl6.grid(column=1, row=4)
        self.combo2.grid(column=2, row=4)
        lbl7.grid(column=1, row=5)
        self.combo3.grid(column=2, row=5)
        lbl8.grid(column=1, row=6)
        self.combo4.grid(column=2, row=6)

        btn3 = Button(self,text = 'Сохранить геометрию',command = lambda :self.save_geometry_2d())
        btn3.grid(column=3, row=3)
        btn4 = Button(self, text='Сохранить геометрию', command=lambda:self.save_geometry_3d(self.check_state2.get(),int(self.combo2.get())))
        btn4.grid(column=3, row=4)

    def open_file(self,problem):
        op = askopenfilename(filetypes = (("Binary Post File","*.t16"),("all files","*.*")))
        if problem == '2d':
            self.post_file2d = op
            self.lbl3['text'] = op
            self.check_state1.set(True)
            self.get_data(problem)
            print(self.elements_2d)
            #p = post_open(self.post_file2d)
            #incs = []
            #for i in range(0,p.increments()-1):
                #incs.append(i)
            #self.combo1['values']=incs
        else:
            self.post_file3d = op
            self.lbl4['text'] = op
            self.check_state2.set(True)
            self.check_state3.set(True)
            self.get_data(problem)
            print(self.elements_3d)
            #p = post_open(self.post_file3d)
            #incs = []
            #for i in range(0, p.increments() - 1):
                #incs.append(i)
            #self.combo2['values'] = incs
            #self.combo3['values'] = incs
            #self.combo4['values'] = incs
        if self.lbl3['text'] and self.lbl4['text']:
            self.total_rep = int(self.elements_3d / self.elements_2d)
            print(self.total_rep)
    def get_data(self, problem):
        incs = []
        if problem == '2d':
            p = post_open(self.post_file2d)
            for i in range(0, p.increments() - 1):
                incs.append(i)
            self.combo1['values'] = incs
            self.combo1.current(max(incs))
            p.moveto(0)
            self.elements_2d = p.elements()
        else:
            p = post_open(self.post_file3d)
            for i in range(0, p.increments() - 1):
                incs.append(i)
                self.combo2['values'] = incs
                self.combo2.current(min(incs))
                self.combo3['values'] = incs
                self.combo3.current(max(incs))
                self.combo4['values'] = incs
                p.moveto(0)
                self.elements_3d = p.elements()

    def save_geometry_2d(self):
        inc = int(self.combo1.get())
        border = self.get_border(inc+1)
        border_string = str()
        for line in border:
            border_string += str(line)+'\n'
        border_string = re.sub('[[]', '' , border_string)
        border_string = re.sub('[]]', '', border_string)
        file_name = asksaveasfilename(filetypes=(("TXT files", "*.txt"),("all files","*.*")))
        f = open(file_name, 'w')
        f.write(border_string)
        f.close()
    def get_border(self,inc):
        p = post_open(self.post_file2d)
        p.moveto(inc)
        self.nodes =[]
        nodes = []
        elements =[]
        edges = []
        for i in range(0,p.nodes()):
            d = dict(id=p.node(i).id, pos = [p.node(i).x,p.node(i).y,p.node(i).z])
            nodes.append(d)
        for i in range(0,p.elements()):
            d = dict(id = p.element(i).id,items = p.element(i).items,
                     type =p.element(i).type)
            elements.append(d)
            if  elements[i]['type'] == 82:
                elements[i]['items'].pop()
                elements[i]['items'] = list(OrderedDict.fromkeys(elements[i]['items']).keys())
                elements[i]['edges'] = list()
                for n in range(0,len(elements[i]['items'])):
                    node0 = elements[i]['items'][n]
                    node1 = elements[i]['items'][n-1]
                    edge = [node0,node1]
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
        repeat =Counter(edges)
        #
        self.border_edges = []
        for edge in edges:
            if repeat[edge]==1:
                self.border_edges.append(edge)
        ####
        border_edges_xy = []
        for edge in self.border_edges:
            edge_xy = []
            for node_id in edge:
                have_disp = p.node_displacements()
                if have_disp:
                    dx , dy, dz = p.node_displacement(node_id-1)
                else:
                    dx, dy, dz = 0,0,0
                edge_xy.append(nodes[node_id-1]['pos'][0]+ float(dx))
                edge_xy.append(nodes[node_id - 1]['pos'][1]+ float(dy))
                edge_xy.append(nodes[node_id - 1]['pos'][2]+float(dz))
            border_edges_xy.append(edge_xy)
        #####
        return border_edges_xy
    def save_geometry_3d(self,check,inc):
        self.get_border(1)
        border_edges = []
        for edge in self.border_edges:
            edg = []
            for node in edge:
                nod = node*self.total_rep*2 - self.total_rep + node + len(self.nodes)
                edg.append(nod)
            border_edges.append(edg)
        print(self.border_edges)
        print(border_edges)
        p = post_open(self.post_file3d)
        p.moveto(inc+1)
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
                edge_xy.append(p.node(nodes.index(node_id)).x  + float(dx))
                edge_xy.append(p.node(nodes.index(node_id)).y  + float(dy))
                edge_xy.append(p.node(nodes.index(node_id)).z  + float(dz))
            border_edges_xy.append(edge_xy)
        border = border_edges_xy
        border_string = str()
        for line in border:
            border_string += str(line) + '\n'
        border_string = re.sub('[[]', '', border_string)
        border_string = re.sub('[]]', '', border_string)
        file_name = asksaveasfilename(filetypes=(("TXT files", "*.txt"), ("all files", "*.*")))
        f = open(file_name, 'w')
        f.write(border_string)
        f.close()

if __name__ =='__main__':
    #py_connect('',40007)
    GUI().run()
    #py_disconnect()
