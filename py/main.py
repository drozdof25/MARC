import sys
import re
sys.path.append('C:\\MSC.Software\\Marc\\2018.0.0\\mentat2018\\shlib\\win64')
from py_mentat import*
from py_post import*
from tkinter import*
from tkinter.filedialog import *
from tkinter.ttk import Combobox
from collections import OrderedDict,Counter
import xlwt
import numpy as np


class GUI(Tk):
    def run(self):
        self.build_main_widgets()
        self.title('v0')
        self.mainloop()
    def build_main_widgets(self):
        frame1 = Frame(self)
        frame1.pack()
        btn = Button(frame1,text = 'Обзор',command = lambda :self.open_file('2d'),width =10)
        btn2 = Button(frame1,text = 'Обзор',command = lambda :self.open_file('3d'),width =10)
        lbl = Label(frame1, text = 'Файл решения 2D:')
        lbl2 = Label(frame1, text='Файл решения 3D:')
        self.lbl3 = Label(frame1)
        self.lbl4 = Label(frame1)

        btn.grid(column=0,row =0)
        btn2.grid(column=0, row=1)
        lbl.grid(column=1, row=0)
        lbl2.grid(column=1, row=1)
        self.lbl3.grid(column=2, row=0,columnspan = 10)
        self.lbl4.grid(column=2, row=1,columnspan = 10)

        frame2 = LabelFrame(self,text = 'Геометрия')
        frame2.pack()



        self.check_state1 = BooleanVar()
        self.check_state2 = BooleanVar()
        self.check_state3 = BooleanVar()
        self.check_state4 = BooleanVar()
        chk1 = Checkbutton(frame2,text = '2D',variable =self.check_state1)
        chk2 = Checkbutton(frame2, text='3D', variable=self.check_state2)

        chk1.grid(column = 0,row =3)
        chk2.grid(column=0, row=4)


        lbl5 = Label(frame2,text = 'инкремент')
        self.combo1 = Combobox(frame2)
        lbl6 = Label(frame2, text='инкремент')
        self.combo2 = Combobox(frame2)


        lbl5.grid(column = 1,row =3)
        self.combo1.grid(column=2, row=3)
        lbl6.grid(column=1, row=4)
        self.combo2.grid(column=2, row=4)


        btn3 = Button(frame2,text = 'Сохранить геометрию',command = lambda :self.save_geometry_2d())
        btn3.grid(column=3, row=3)
        btn4 = Button(frame2, text='Сохранить геометрию', command=lambda:self.save_geometry_3d(self.check_state2.get(),int(self.combo2.get())))
        btn4.grid(column=3, row=4)

        self.frame3 = LabelFrame(self, text='Энергия и работа')
        self.frame3.pack()

        e_lbl0 = Label(self.frame3,text = 'Задача')
        e_lbl0.grid(column=0, row=0)
        self.e_lbl0_1 = Label(self.frame3,text = 'Посадка')
        self.e_lbl0_1.grid(column=0, row=1)
        self.e_lbl0_2 = Label(self.frame3, text='Обжатие')
        self.e_lbl0_2.grid(column=0, row=2)
        self.e_lbl0_3 = Label(self.frame3, text='Качение')
        self.e_lbl0_3 .grid(column=0, row=3)

        e_lbl1 = Label(self.frame3,text = 'Инкремент')
        e_lbl1.grid(column = 1 ,row =0)
        self.e_combo1 = Combobox(self.frame3)
        self.e_combo1.grid(column=1, row=1)
        self.e_combo2 = Combobox(self.frame3)
        self.e_combo2.grid(column=1, row=2)
        self.e_combo3 = Combobox(self.frame3)
        self.e_combo3.grid(column=1, row=3)
        e_lbl2 = Label(self.frame3, text='Полная энергия \n деформации, Дж')
        e_lbl2.grid(column=2, row=0)

        e_lbl3 = Label(self.frame3, text='Полная работа, Дж')
        e_lbl3.grid(column=3, row=0)

        e_lbl4 = Label(self.frame3, text='Полная энергия упругой \n деформации, Дж')
        e_lbl4.grid(column=4, row=0)

        e_lbl5 = Label(self.frame3, text='Полная работа приложеных \n сил/перемещений, Дж')
        e_lbl5.grid(column=5, row=0)

        e_lbl6 = Label(self.frame3, text='Полная работа контактных \n сил, Дж')
        e_lbl6.grid(column=6, row=0)

        e_lbl7 = Label(self.frame3, text='Полная работа сил трения \n в контакте, Дж')
        e_lbl7.grid(column=7, row=0)

        e_lbl8 = Label(self.frame3, text='Объём, мм^3')
        e_lbl8.grid(column=8, row=0)

        e_lbl9 = Label(self.frame3, text='Масса, кг')
        e_lbl9.grid(column=9, row=0)

        e_lbl10 = Label(self.frame3,text = 'Радиус шины,мм' )
        e_lbl10.grid(column=10, row=0)

        e_lbl10 = Label(self.frame3, text='Динамический радиус,мм')
        e_lbl10.grid(column=11, row=0)

        e_lbl11 = Label(self.frame3, text='Нагрузка Q,кг')
        e_lbl11.grid(column=12, row=0)


        e_btn = Button(self.frame3,text = 'обновить',command = lambda :self.refresh_energy())
        e_btn.grid(column=8, row=4)
        e_btn2 = Button(self.frame3, text='сохранить',command = lambda : self.save_energy())
        e_btn2.grid(column=9, row=4)

    def open_file(self,problem):
        op = askopenfilename(filetypes = (("Binary Post File","*.t16"),("all files","*.*")))
        if problem == '2d':
            self.post_file2d = op
            self.lbl3['text'] = op
            self.check_state1.set(True)
            self.get_data(problem)
            print(self.elements_2d)
        else:
            self.post_file3d = op
            self.lbl4['text'] = op
            self.check_state2.set(True)
            self.check_state3.set(True)
            self.get_data(problem)
            print(self.elements_3d)
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
            incs.append(NONE)
            self.e_combo1['values'] = incs
            self.e_combo2['values'] = incs
            self.e_combo3['values'] = incs

            p.moveto(0)
            self.elements_3d = p.elements()

    def save_geometry_2d(self):
        inc = int(self.combo1.get())
        self.save_geometry(self.get_border(inc))

    def get_border(self,inc):
        p = post_open(self.post_file2d)
        p.moveto(inc+1)
        self.nodes =[]
        elements =[]
        edges = []
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
        self.border_edges = []
        for edge in edges:
            if repeat[edge]==1:
                self.border_edges.append(edge)
        return self.getborder_disp(inc,self.post_file2d,self.border_edges)

    def save_geometry_3d(self,check,inc):
        self.get_border(1)
        border_edges = []
        for edge in self.border_edges:
            edg = []
            for node in edge:
                nod = node*self.total_rep*2 - self.total_rep + node + len(self.nodes)
                edg.append(nod)
            border_edges.append(edg)
        self.getborder_disp(inc,self.post_file3d,border_edges)
        border = self.getborder_disp(inc,self.post_file3d,border_edges)
        self.save_geometry(border)

    def getborder_disp(self,inc,post_file,border_edges):
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

    def save_geometry(self,border):
        border_string = str()
        for line in border:
            border_string += str(line) + '\n'
        border_string = re.sub('[[]', '', border_string)
        border_string = re.sub('[]]', '', border_string)
        file_name = asksaveasfilename(filetypes=(("TXT files", "*.txt"), ("all files", "*.*")))
        f = open(file_name, 'w')
        f.write(border_string)
        f.close()
    def refresh_energy(self):
        inc_combos = [self.e_combo1,self.e_combo2,self.e_combo3]
        self.energy = []
        p = post_open(self.post_file3d)
        p.moveto(0 + 1)
        node_y = []
        for i in range(0, p.nodes()):
            dx, dy, dz = p.node_displacement(i)
            node_y.append(p.node(i).y + dy)
        radius = max(node_y)
        external_force_id = 0
        for i in range(0,p.node_scalars()):
            if p.node_scalar_label(i) == 'External Force Y':
                external_force_id = i
        node_0 = 0
        for i in range(0,p.nodes()):
            if p.node(i).y == 0 and p.node(i).x ==0 and p.node(i).z == 0:
                node_0 = i
                break
        tire_index = 0
        for i in range(0, p.cbodies()):
            if p.cbody(i).type ==0:
                tire_index = 0
        for inc_combo in inc_combos:
            job = ''
            if inc_combos.index(inc_combo)==0:
                job = 'Посадка'
            elif inc_combos.index(inc_combo)==1:
                job = 'Обжатие'
            elif inc_combos.index(inc_combo) == 2:
                job = 'Качение'
            grid_info = inc_combo.grid_info()
            try:
                inc = int(inc_combo.get())
            except:
                continue
            else:
                p.moveto(inc + 1)
                strainenergy = Label(self.frame3,text = str(p.strainenergy/1000))
                work = Label(self.frame3,text=str(p.work / 1000))
                elasticenergy = Label(self.frame3,text=str(p.elasticenergy / 1000))
                appliedwork = Label(self.frame3,text=str(p.appliedwork / 1000))
                contactwork = Label(self.frame3,text=str(p.contactwork / 1000))
                frictionwork = Label(self.frame3,text=str(p.frictionwork/ 1000))
                volume = Label(self.frame3,text=str(p.volume))
                mass = Label(self.frame3,text=str(p.mass *1000))
                rad = Label(self.frame3,text = str(radius))
                energ = [job,int(inc_combo.get()),p.strainenergy/1000,p.work / 1000,p.elasticenergy / 1000,p.appliedwork / 1000,
                               p.contactwork / 1000,p.frictionwork/ 1000,p.volume,p.mass *1000,radius]
                if energ[0] != 'Посадка':
                    node_y = []
                    for i in range(0, p.nodes()):
                        dx, dy, dz = p.node_displacement(i)
                        node_y.append(p.node(i).y + dy)
                    energ.append(abs(min(node_y)))
                    rad_antiprogib = abs(max(node_y))
                    din_radius = Label(self.frame3,text = str(abs(min(node_y))))
                    din_radius.grid(column=11, row=grid_info['row'])
                    load = round(p.node_scalar(node_0,external_force_id)/9.81)
                    energ.append(load)
                    load_lbl = Label(self.frame3,text = str(load))
                    load_lbl.grid(column=12, row=grid_info['row'])
                    progib = radius - abs(min(node_y))
                    antiprogib = rad_antiprogib - radius
                    energ.append(progib)
                    energ.append(antiprogib)
                    if energ[0] == 'Качение':
                        velx,vely,velz = p.cbody_velocity(tire_index)
                        line_vel = float(velz) *0.0036
                        print(p.cbody(tire_index))
                self.energy.append(energ)
                strainenergy.grid(column = 2, row = grid_info['row'])
                work.grid(column=3, row=grid_info['row'])
                elasticenergy.grid(column=4, row=grid_info['row'])
                appliedwork.grid(column=5, row=grid_info['row'])
                contactwork.grid(column=6, row=grid_info['row'])
                frictionwork.grid(column=7, row=grid_info['row'])
                volume.grid(column=8, row=grid_info['row'])
                mass.grid(column=9, row=grid_info['row'])
                rad.grid(column=10, row=1,rowspan = 3)

    def save_energy(self):
        file_name = asksaveasfilename(defaultextension ='.xls',filetypes=(("Excel book", "*.xls"), ("all files", "*.*")))
        print(file_name)
        wb = xlwt.Workbook()
        ws = wb.add_sheet('Энергия и работа',cell_overwrite_ok=True)
        titles = ['Задача','Инкремент','Полная энергия деформации, Дж','Полная работа, Дж','Полная энергия упругой деформации, Дж',
                  'Полная работа приложеных сил/перемещений, Дж','Полная работа контактных  сил, Дж',
                  'Полная работа сил трения  в контакте, Дж','Объём, мм^3','Масса, кг','Радиус шины, мм','Динамический радиус, мм',
                  'Нагрузка Q, кг','Прогиб, мм','Антипрогиб, мм','Радиус качения, мм','Деформация сжатия протектора в окружном направлении %',
                  'Линейная скорость качения шины, км/ч','Угловая скорость качения шины, рад/с','Fzroad,Н','КСК']
        # Устанавливаем перенос по словам, выравнивание
        alignment = xlwt.Alignment()
        alignment.wrap = 1
        alignment.horz = xlwt.Alignment.HORZ_CENTER  # May be: HORZ_GENERAL, HORZ_LEFT, HORZ_CENTER, HORZ_RIGHT, HORZ_FILLED, HORZ_JUSTIFIED, HORZ_CENTER_ACROSS_SEL, HORZ_DISTRIBUTED
        alignment.vert = xlwt.Alignment.VERT_CENTER  # May be: VERT_TOP, VERT_CENTER, VERT_BOTTOM, VERT_JUSTIFIED, VERT_DISTRIBUTED
        # Устанавливаем шрифт
        font = xlwt.Font()
        font.name = 'Arial Cyr'
        font.bold = True
        # Устанавливаем границы
        borders = xlwt.Borders()
        borders.left = xlwt.Borders.THIN  # May be: NO_LINE, THIN, MEDIUM, DASHED, DOTTED, THICK, DOUBLE, HAIR, MEDIUM_DASHED, THIN_DASH_DOTTED, MEDIUM_DASH_DOTTED, THIN_DASH_DOT_DOTTED, MEDIUM_DASH_DOT_DOTTED, SLANTED_MEDIUM_DASH_DOTTED, or 0x00 through 0x0D.
        borders.right = xlwt.Borders.THIN
        borders.top = xlwt.Borders.THIN
        borders.bottom = xlwt.Borders.THIN
        style_titles = xlwt.XFStyle()
        style_titles.font = font
        style_titles.alignment = alignment
        style_titles.borders = borders
        style_data = xlwt.XFStyle()
        style_data.borders = borders
        for i in range(1,len(titles)+1):
            ws.write_merge(1, 8, i, i,titles[i - 1],style_titles)
            #ws.write(1, i, titles[i - 1], style)
        c = 9
        for problem in self.energy:
            r = 1
            for energy in problem:
                if r == 11:
                    ws.write_merge(9,11,r,r,energy,style_data)
                else:
                    ws.write(c,r,energy,style_data)
                r += 1
            c += 1
        wb.save(file_name)
if __name__ =='__main__':
    GUI().run()

