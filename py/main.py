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
import matplotlib.pyplot as plt
import win32com.client
from sklearn.linear_model import LinearRegression

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

        self.strainenergy = []
        for i in range(0,3):
            self.strainenergy.append(Label(self.frame3, width=15, height=2, borderwidth=2, relief="groove"))
            self.strainenergy[i].grid(column=2, row=i+1)

        self.work = []
        for i in range(0, 3):
            self.work.append(Label(self.frame3, width=15, height=2, borderwidth=2, relief="groove"))
            self.work[i].grid(column=3, row=i + 1)

        self.elasticenergy = []
        for i in range(0, 3):
            self.elasticenergy.append(Label(self.frame3, width=15, height=2, borderwidth=2, relief="groove"))
            self.elasticenergy[i].grid(column=4, row=i + 1)

        self.appliedwork = []
        for i in range(0, 3):
            self.appliedwork.append(Label(self.frame3, width=15, height=2, borderwidth=2, relief="groove"))
            self.appliedwork[i].grid(column=5, row=i + 1)

        self.contactwork = []
        for i in range(0, 3):
            self.contactwork.append(Label(self.frame3, width=15, height=2, borderwidth=2, relief="groove"))
            self.contactwork[i].grid(column=6, row=i + 1)

        self.frictionwork = []
        for i in range(0, 3):
            self.frictionwork.append(Label(self.frame3, width=15, height=2, borderwidth=2, relief="groove"))
            self.frictionwork[i].grid(column=7, row=i + 1)

        self.volume = []
        for i in range(0, 3):
            self.volume.append(Label(self.frame3, width=15, height=2, borderwidth=2, relief="groove"))
            self.volume[i].grid(column=8, row=i + 1)

        self.mass = []
        for i in range(0, 3):
            self.mass.append(Label(self.frame3, width=15, height=2, borderwidth=2, relief="groove"))
            self.mass[i].grid(column=9, row=i + 1)

        self.rad = []
        for i in range(0, 3):
            self.rad.append(Label(self.frame3, width=15, height=2, borderwidth=2, relief="groove"))
            self.rad[i].grid(column=10, row=i + 1)

        self.din_rad = []
        for i in range(0, 3):
            self.din_rad.append(Label(self.frame3, width=15, height=2, borderwidth=2, relief="groove"))
            self.din_rad[i].grid(column=11, row=i + 1)

        self.load_lbl = []
        for i in range(0, 3):
            self.load_lbl.append(Label(self.frame3, width=15, height=2, borderwidth=2, relief="groove"))
            self.load_lbl[i].grid(column=12, row=i + 1)

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

        self.frame4 = LabelFrame(self, text='Жесткости',width = 1700)
        self.frame4.pack()
        frame5 = LabelFrame(self.frame4,text = 'Радиальная жесткость')
        frame5.pack()
        g_lbl = Label(frame5,text = 'Диапозон инкрементов: ')
        g_lbl.grid(column = 0, row = 0)
        self.g_combo1 = Combobox(frame5)
        self.g_combo2 = Combobox(frame5)
        self.g_combo1.grid(column = 1, row =0)
        self.g_combo2.grid(column = 2, row=0)
        g_btn = Button(frame5, text='получить данные',command = lambda : self.radial_hardness())
        g_btn.grid(column=3, row=0)
        g_btn = Button(frame5, text='сохранить данные', command=lambda: self.save_radial_hardness())
        g_btn.grid(column=4, row=0)
    def open_file(self,problem):
        op = askopenfilename(filetypes = (("Binary Post File","*.t16"),("all files","*.*")))
        if problem == '2d':
            self.post_file2d = op
            self.lbl3['text'] = op
            self.check_state1.set(True)
            self.get_data(problem)

        else:
            self.post_file3d = op
            self.lbl4['text'] = op
            self.check_state2.set(True)
            self.check_state3.set(True)
            self.get_data(problem)

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
            self.g_combo1['values'] = incs
            self.g_combo2['values'] = incs
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
        file_name = asksaveasfilename(filetypes=(("txt files", "*.txt"), ("all files", "*.*")))
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
        external_force_id = self.get_scalar_id('External Force Y')
        node_0 = self.get_node_0()
        tire_index = None
        road_index = None
        for i in range(0, p.cbodies()):

            if p.cbody(i).type ==0:
                tire_index = i
            if p.cbody(i).type == 2:
                road_index = i
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
                energ = [job,int(inc_combo.get()),p.strainenergy/1000,p.work / 1000,p.elasticenergy / 1000,p.appliedwork / 1000,
                               p.contactwork / 1000,p.frictionwork/ 1000,p.volume,p.mass *1000,radius]
                if energ[0] != 'Посадка':
                    node_y = []
                    for i in range(0, p.nodes()):
                        dx, dy, dz = p.node_displacement(i)
                        node_y.append(p.node(i).y + dy)
                    energ.append(abs(min(node_y)))
                    rad_antiprogib = abs(max(node_y))

                    self.din_rad[grid_info['row']-1]['text'] = str(round(abs(min(node_y)),4))
                    load = round(p.node_scalar(node_0,external_force_id)/9.81)
                    energ.append(load)
                    self.load_lbl[grid_info['row']-1]['text'] = str(load)
                    progib = radius - abs(min(node_y))
                    antiprogib = rad_antiprogib - radius
                    energ.append(progib)
                    energ.append(antiprogib)
                    if energ[0] == 'Качение':
                        velx,vely,velz = p.cbody_velocity(tire_index)
                        line_vel = round(float(velz) *0.0036)
                        rotat = p.cbody_rotation(tire_index)

                        angle_vel = rotat * 2 * 3.141592653589793238462643
                        rad_kach = line_vel*1000000/3600/angle_vel
                        deformation = (1-rad_kach/radius) *100
                        fx,fy,force_z_road = p.cbody_force(road_index)
                        ksk = abs(force_z_road)/p.node_scalar(node_0,external_force_id)
                        energ.append(rad_kach)
                        energ.append(deformation)
                        energ.append(line_vel)
                        energ.append(angle_vel)
                        energ.append(abs(force_z_road))
                        energ.append(ksk)
                self.energy.append(energ)
                self.strainenergy[grid_info['row']-1]['text'] = str(round(p.strainenergy/1000,4))
                self.work[grid_info['row'] - 1]['text'] = str(round(p.work/1000,4))
                self.elasticenergy[grid_info['row'] - 1]['text'] = str(round(p.elasticenergy / 1000,4))
                self.appliedwork[grid_info['row'] - 1]['text'] = str(round(p.appliedwork / 1000,4))
                self.contactwork[grid_info['row'] - 1]['text'] = str(round(p.contactwork / 1000,4))
                self.frictionwork[grid_info['row'] - 1]['text'] = str(round(p.frictionwork / 1000,4))
                self.volume[grid_info['row'] - 1]['text'] = str(round(p.volume,4))
                self.mass[grid_info['row'] - 1]['text'] = str(round(p.mass*1000,4))
                self.rad[grid_info['row'] - 1]['text'] = str(round(radius,4))


    def save_energy(self):
        file_name = asksaveasfilename(defaultextension ='.xls',filetypes=(("Excel book", "*.xls"), ("all files", "*.*")))
        print(file_name)
        Excel = win32com.client.Dispatch("Excel.Application")
        wb = Excel.Workbooks.Open(u'D:\\Projects\\MARC+Python\\MARC\\py\\shablon.xls')
        sheet = wb.ActiveSheet
        c = 10
        for problem in self.energy:
            r = 2
            for energy in problem:
                if r == 12:
                    sheet.Cells(10,12).value = energy
                elif r==15:
                    sheet.Cells(c,r).value = energy
                elif r == 22:
                    sheet.Cells(c,r).value = energy
                else:
                    sheet.Cells(c,r).value = energy
                r += 1
            c+=1
        # wb = xlwt.Workbook()
        # ws = wb.add_sheet('Энергия и работа',cell_overwrite_ok=True)
        # titles = ['Задача','Инкремент','Полная энергия деформации, Дж','Полная работа, Дж','Полная энергия упругой деформации, Дж',
        #           'Полная работа приложеных сил/перемещений, Дж','Полная работа контактных  сил, Дж',
        #           'Полная работа сил трения  в контакте, Дж','Объём, мм^3','Масса, кг','Радиус шины, мм','Динамический радиус, мм',
        #           'Нагрузка Q, кг','Прогиб, мм','Антипрогиб, мм','Радиус качения, мм','Деформация сжатия протектора в окружном направлении %',
        #           'Линейная скорость качения шины, км/ч','Угловая скорость качения шины, рад/с','Fzroad,Н','КСК']
        # # Устанавливаем перенос по словам, выравнивание
        # alignment = xlwt.Alignment()
        # alignment.wrap = 1
        # alignment.horz = xlwt.Alignment.HORZ_CENTER  # May be: HORZ_GENERAL, HORZ_LEFT, HORZ_CENTER, HORZ_RIGHT, HORZ_FILLED, HORZ_JUSTIFIED, HORZ_CENTER_ACROSS_SEL, HORZ_DISTRIBUTED
        # alignment.vert = xlwt.Alignment.VERT_CENTER  # May be: VERT_TOP, VERT_CENTER, VERT_BOTTOM, VERT_JUSTIFIED, VERT_DISTRIBUTED
        # # Устанавливаем шрифт
        # font = xlwt.Font()
        # font.name = 'Arial Cyr'
        # font.bold = True
        # #Фон
        # patern = xlwt.Pattern()
        # patern.pattern = xlwt.Pattern.SOLID_PATTERN
        # patern.pattern_fore_colour = xlwt.Style.colour_map['yellow']
        #
        # # Устанавливаем границы
        # borders = xlwt.Borders()
        # borders.left = xlwt.Borders.THIN  # May be: NO_LINE, THIN, MEDIUM, DASHED, DOTTED, THICK, DOUBLE, HAIR, MEDIUM_DASHED, THIN_DASH_DOTTED, MEDIUM_DASH_DOTTED, THIN_DASH_DOT_DOTTED, MEDIUM_DASH_DOT_DOTTED, SLANTED_MEDIUM_DASH_DOTTED, or 0x00 through 0x0D.
        # borders.right = xlwt.Borders.THIN
        # borders.top = xlwt.Borders.THIN
        # borders.bottom = xlwt.Borders.THIN
        # style_titles = xlwt.XFStyle()
        # style_titles.font = font
        # style_titles.alignment = alignment
        # style_titles.borders = borders
        # style_data = xlwt.XFStyle()
        # style_data.borders = borders
        # style_data2 = xlwt.XFStyle()
        # style_data2.pattern = patern
        # style_data2.borders = borders
        # style_data2.font = font
        # style_data2.alignment =alignment
        #
        # for i in range(1,len(titles)+1):
        #     ws.write_merge(1, 8, i, i,titles[i - 1],style_titles)
        #     #ws.write(1, i, titles[i - 1], style)
        # c = 9
        # for problem in self.energy:
        #     r = 1
        #     for energy in problem:
        #         if r == 11:
        #             ws.write_merge(9,11,r,r,energy,style_data2)
        #         elif r==14:
        #             ws.write(c,r,energy,style_data2)
        #         elif r == 21:
        #             ws.write(c, r, energy, style_data2)
        #         else:
        #             ws.write(c, r, energy, style_data)
        #         r += 1
        #     c += 1
        # wb.save(file_name)
        file_name = file_name.replace('/', '\\')
        wb.SaveAs(file_name)
        wb.Close()
        Excel.Quit()
    def radial_hardness(self):
        p = post_open(self.post_file3d)
        inc1 = int(self.g_combo1.get())+1
        inc2 = int(self.g_combo2.get())+1
        y_ax = []
        x_ax = []
        ekvator_node_index = self.get_ekvator_node()
        node_0_index = self.get_node_0()
        external_force_id = self.get_scalar_id('External Force Y')
        for i in range(inc1,inc2+1):
            p.moveto(i)
            dx,dy,dz = p.node_displacement(ekvator_node_index)
            load = p.node_scalar(node_0_index, external_force_id)
            y_ax.append(load)
            x_ax.append(dy)
        self.radial_hardness_data = [y_ax ,x_ax]
        x_0 = x_ax[0]
        for i in range(0,len(x_ax)):
            x_ax[i] -= x_0
        x = np.array(x_ax).reshape((-1 , 1))
        y = np.array(y_ax)
        model = LinearRegression().fit(x, y)
        self.intercept = model.intercept_
        self.coef = model.coef_[0]
        plt.plot(y_ax,x_ax,marker = 'o',linestyle = 'dashed')
        plt.title('Радиальная жесткость')
        plt.ylabel('Перемещение')
        plt.xlabel('Нагрузка')
        plt.grid(True)
        plt.show()

    def get_ekvator_node(self):
        p = post_open(self.post_file3d)
        nodes = dict()
        p.moveto(0 + 1)
        for i in range(0, p.nodes()):
            dx, dy, dz = p.node_displacement(i)
            nodes[i] = float(p.node(i).y + dy)
        ekvator_y = min(nodes.values())
        print(ekvator_y)
        for key,item in nodes.items():
            if item == ekvator_y:
                return key
    def get_node_0(self):
        p = post_open(self.post_file3d)
        p.moveto(0 + 1)
        for i in range(0, p.nodes()):
            if p.node(i).y == 0 and p.node(i).x == 0 and p.node(i).z == 0:
                return i
    def get_scalar_id(self,scalar):
        p = post_open(self.post_file3d)
        p.moveto(0 + 1)
        for i in range(0,p.node_scalars()):
            if p.node_scalar_label(i) == scalar:
                return i

    def save_radial_hardness(self):
        file_name = asksaveasfilename(defaultextension='.xls',
                                      filetypes=(("Excel book", "*.xls"), ("all files", "*.*")))
        print(file_name)
        Excel = win32com.client.Dispatch("Excel.Application")
        wb = Excel.Workbooks.Open(u'D:\\Projects\\MARC+Python\\MARC\\py\\shablon.xls')
        sheet = wb.ActiveSheet
        r = 2
        for axe in self.radial_hardness_data:
            c = 23
            for value in axe:
                sheet.Cells(c, r).value = value
                c += 1
            r += 1
        len_data = len(self.radial_hardness_data[0])
        sheet.Cells(23,4).AutoFill(sheet.Range(sheet.Cells(23, 4), sheet.Cells(23 +len_data-1, 4)))
        print(sheet.Cells(23,4))
        sheet.Cells(50, 9).value = self.intercept
        sheet.Cells(50, 8).value = self.coef
        sheet.ChartObjects(1).Activate()
        chart = wb.ActiveChart
        chart.SeriesCollection(1).XValues = "='Энергия и работа'!$D$23:$D${0}".format(23+len_data)
        chart.SeriesCollection(1).Values = "='Энергия и работа'!$B$23:$B${0}".format(23+len_data)
        file_name = file_name.replace('/', '\\')
        wb.SaveAs(file_name)
        wb.Close()
        Excel.Quit()
if __name__ =='__main__':
    GUI().run()

