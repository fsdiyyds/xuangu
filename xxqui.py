# -*- coding: utf-8 -*-

'''
@Time    : 2022/2/21 16:34
@Author  : fenglei
@FileName: xxqui.py
@Software: PyCharm
 
'''
import tkinter as tk
from tkinter import ttk

import os
from tkinter import filedialog
from tkinter import messagebox as mBox


class Login(tk.Tk):
    """登录窗口UI"""

    def __init__(self):
        """初始化"""
        super().__init__()  # 有点相当于tk.Tk()
        self.photo = tk.PhotoImage(file=".\\xxq.png")
        self.photo1 = tk.PhotoImage(file=".\\封面.png")
        self.photo2 = tk.PhotoImage(file=".\\钢筋.png")
        self.shuliang = []
        self.guimian_ele = tk.StringVar()
        self.guimian_podu = tk.StringVar()
        self.guimian_xiangding = tk.StringVar()
        self.guimian_lujian = tk.StringVar()
        self.xiangti_theta = tk.StringVar()
        self.xiangti_qiaokuan = tk.StringVar()
        self.xiangti_kongjin = tk.StringVar()
        self.xiangti_dingban = tk.StringVar()
        self.xiangti_diban = tk.StringVar()
        self.xiangti_dingdao_c = tk.StringVar()
        self.xiangti_dingdao_g = tk.StringVar()
        self.xiangti_didao_c = tk.StringVar()
        self.xiangti_didao_g = tk.StringVar()
        self.xiangti_jinggao = tk.StringVar()
        self.xiangti_bianfu = tk.StringVar()
        self.xiangti_dingpu = tk.StringVar()
        self.xiangti_jichuhou = tk.StringVar()
        self.xiangti_jichuchao = tk.StringVar()
        self.kongshu = tk.StringVar()
        self.xiangti_midfu = tk.StringVar()

        self.yiqiang_di = tk.StringVar()
        self.yiqiang_22 = tk.StringVar()
        self.yiqiang_44 = tk.StringVar()
        self.yiqiang_33 = tk.StringVar()
        self.yiqiang_55 = tk.StringVar()
        self.yiqiang_dingkuan = tk.StringVar()
        self.yiqiang_dingmaoshi = tk.StringVar()
        self.yiqiang_di_zuo = tk.StringVar()
        self.yiqiang_di_you = tk.StringVar()
        self.yiqiang_podu = tk.StringVar()
        self.yiqiang_afa = tk.StringVar()
        self.yiqiang_fangpo = tk.StringVar()
        self.yiqiang_afa2 = tk.StringVar()

        self.xiangti_licheng_qian = tk.StringVar()
        self.xiangti_licheng = tk.StringVar()
        self.licheng_fangxiang = tk.StringVar()
        self.zuo_li = tk.StringVar()
        self.you_li = tk.StringVar()
        self.mid_chengjiang = tk.StringVar()
        self.mid_num = tk.StringVar()
        self.xiangti_yanshen = tk.StringVar()

        self.gouzao = 'X'

        self.btn_excel = ttk.Button()

        self.welcome = tk.StringVar()
        self.user = tk.StringVar()
        self.pwd = tk.StringVar()
        self.msg = tk.StringVar()
        self.btn_ok = ttk.Button()
        self.btn_ok1 = ttk.Button()
        self.btn_pj = ttk.Button()
        self.btn_shuliang = ttk.Button()
        self.btn_select = ttk.Button()
        self.btn_exzample = ttk.Button()
        self.btn_export = ttk.Button()

        self.version = tk.StringVar()
        self.com_version = ttk.Combobox()
        self.r_value = tk.StringVar()

        self.shuchu = tk.StringVar()
        self.text_out = tk.Text()
        self.text_out1 = tk.Text()
        self.tkend = tk.END
        self.excel_path = tk.StringVar()
        self.excel_real_path = ''
        self.file_dir = tk.StringVar(value="'C:\\Windows'")

        self.tag = tk.IntVar()
        self.tag1 = tk.IntVar()

        self.WIDTH = 640
        self.HEIGHT = 500

        self.N1_1 = tk.StringVar()
        self.N1_2 = tk.StringVar()
        self.N2 = tk.StringVar()
        self.N3 = tk.StringVar()
        self.N4 = tk.StringVar()
        self.N5 = tk.StringVar()
        self.N6 = tk.StringVar()
        self.N7 = tk.StringVar()
        self.Nmid_7 = tk.StringVar()
        self.N16 = tk.StringVar()
        self.N17 = tk.StringVar()
        # gouzao gangjing
        self.N12 = tk.StringVar()
        self.jianju_N12 = tk.StringVar()
        self.N8_9 = tk.StringVar()
        self.N_jichu = tk.StringVar()
        self.baohu = tk.StringVar()
        # "E:/exe文件/箱形桥（框架桥）CAD辅助设计FBB(取消限制)/FBB_V2_2.VLX"
        self.file_export = tk.StringVar()
        self.excel_export = tk.StringVar()
        self.file_fbi = tk.StringVar()
        self.shuru_way = ''
        self.pj_file = '.\\pj.fbi'
        self.run()

    def save_excel(self):
        file_op = filedialog.asksaveasfilename(filetypes=[('Excel', 'xlsx'), ('Excel', 'xls'), ('All Files', '*')],
                                               initialdir=self.file_dir)    # 这个是另存为对话框
        self.file_export.set(file_op)

    def ex_excel(self):
        file_op = filedialog.asksaveasfilename(filetypes=[('Excel', 'xlsx'), ('Excel', 'xls'), ('All Files', '*')],
                                               initialdir=self.file_dir)    # 这个是另存为对话框
        self.excel_export.set(file_op)

    def save_fbi(self):
        file_op = filedialog.asksaveasfilename(filetypes=[('工程文件', 'fbi'), ('All Files', '*')],
                                               initialdir=self.file_dir)    # 这个是另存为对话框
        self.file_fbi.set(file_op)

    def run(self):

        def changeTag(tag):
            frame3.pack_forget()
            frame4.pack_forget()
            frame5.pack_forget()
            frame6.pack_forget()
            frame7.pack_forget()

            if tag == 0:
                frame3.pack(fill=tk.X)
            elif tag == 1:
                frame4.pack(fill=tk.X, pady=(2, 0))
                self.excel_real_path = '.\\example.xls'
            elif tag == 2:
                frame5.pack(fill=tk.X)
                self.shuru_way = 'wen'
            elif tag == 3:
                frame6.pack(fill=tk.X)
            elif tag == 4:
                frame7.pack(fill=tk.X)

        shuru_type = 'B'
        self.title('箱形桥智能设计成图软件')
        self.option_add('*Font', ('', 12))
        ttk.Style().configure(".", font=('', 12))

        frame2 = tk.Frame(self)
        frame2.pack(fill=tk.Y, pady=0)

        self.ws = frame2.winfo_screenwidth()
        self.hs = frame2.winfo_screenheight()
        x = (self.ws / 2) - (self.WIDTH / 2)
        y = (self.hs / 2) - (self.HEIGHT / 2)
        self.config(width=self.WIDTH)

        # geometry('%dx{}+%d+%d' % (self.WIDTH, x, y))

        self.tag.set(0)

        tagWidth = 15
        tk.Radiobutton(frame2, text="输入方式", command=lambda: changeTag(0), variable=self.tag, width=tagWidth, value=0,
                       bd=1,
                       indicatoron=0).grid(column=0, row=1)
        tk.Radiobutton(frame2, text="逐项输入", command=lambda: changeTag(1), variable=self.tag, width=tagWidth, value=1,
                       bd=1,
                       indicatoron=0).grid(column=1, row=1)
        tk.Radiobutton(frame2, text="文件输入", command=lambda: changeTag(2), variable=self.tag, width=tagWidth, value=2,
                       bd=1,
                       indicatoron=0).grid(column=2, row=1)
        tk.Radiobutton(frame2, text="绘钢筋图", command=lambda: changeTag(3), variable=self.tag, width=tagWidth, value=3,
                       bd=1,
                       indicatoron=0).grid(column=3, row=1)
        tk.Radiobutton(frame2, text="提数量表", command=lambda: changeTag(4), variable=self.tag, width=tagWidth, value=4,
                       bd=1,
                       indicatoron=0).grid(column=4, row=1)

        # frm3 = tk.Frame(self, bg='grey')
        # frm3.pack(padx=10, pady=10)
        frame3 = tk.Frame(self, height=300)
        frame3.pack(fill=tk.X)

        frame4 = tk.Frame(self, height=300, highlightbackground="black", highlightthickness=1)

        frame5 = tk.Frame(self, height=300, bg="yellow")

        frame6 = tk.Frame(self, height=300, bg="pink")

        frame7 = tk.Frame(self, height=300, bg="black")

        # menuBar = tk.Menu(self)
        # self.config(menu=menuBar)
        # fileMenu = tk.Menu(menuBar, tearoff=0)
        # fileMenu.add_command(label="新建项目", font=('', 8))
        # # fileMenu.add_separator()
        # fileMenu.add_command(label="退出", font=('', 8))
        # menuBar.add_cascade(label="文件", menu=fileMenu)
        frm1 = tk.Frame(frame3, bg='grey')
        frm1.pack(padx=0, pady=(0, 0))
        imageLabel2 = tk.Label(frame3, image=self.photo1).pack(pady=(0, 0))

        frm2 = tk.Frame(frame3, bg='pink')
        frm2.pack(padx=0, pady=2, fill=tk.X)
        frm3 = tk.Frame(frame3, bg='grey')
        frm3.pack(padx=0, pady=2, fill=tk.X)

        self.welcome.set('- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -')
        tk.Label(frm1, textvariable=self.welcome, font=('黑体', 12), foreground='black').pack(pady=0)

        tk.Label(frm2, text="请下拉选择使用的CAD版本", font=('黑体', 12), foreground='black', bg='pink').pack(padx=(0, 15),
                                                                                                  side='left')
        self.com_version = ttk.Combobox(frm2, textvariable=self.version, font=('microsoft yahei', 10, 'italic'),
                                        value=('1------AutoCAD2014------1',
                                               '2------AutoCAD2016------2',
                                               '3------AutoCAD2018------3',
                                               '4------AutoCAD2020------4',
                                               '5------AutoCAD2021------5'), width=50)
        self.com_version.pack(pady=0, side='left', fill=tk.X)
        self.com_version.set(value='3------AutoCAD2018------3')

        # def _msgBox1():
        #     mBox.showinfo('输入方式提示', '目前不支持，逐项输入也请在excel中完成！')

        def zhu_selection():
            global shuru_type
            print('you have selected ' + self.r_value.get())
            shuru_type = self.r_value.get()
            if not shuru_type == 'B':
                self.text_out.pack_forget()
                excel_frm.pack_forget()
                waibu_frm.pack(padx=5, pady=(2, 5), anchor='nw')
                self.text_out.pack(padx=5, pady=(2, 5), fill='y')
            else:
                self.text_out.pack_forget()
                waibu_frm.pack_forget()
                excel_frm.pack(padx=5, pady=5, anchor='nw')
                self.text_out.pack(padx=5, pady=5, fill='y')

        waibu_frm = tk.Frame(frame3, relief='ridge', borderwidth=1)

        self.r_value.set('B')

        tk.Label(frm3, text="注意：①选择CAD版本(已默认勾选CAD2018)之后选择一种输入方式即可\n      ②出图后请反核图面，保证设计质量\n      ③本软件已申请著作权,未经允许不得拷贝传播",
                 bd=1, justify='left', anchor="w").pack(pady=(10, 2), fill=tk.X)

        # tk.Radiobutton(frm3, text="逐项输入", variable=self.tag1, width=8, value=2,
        #                bd=1).grid(column=1, row=1)
        # tk.Radiobutton(frm3, text="文件输入", variable=self.tag1, width=8, value=3,
        #                bd=1).grid(column=2, row=1)
        # button_1 = ttk.Button(frm3, text='确定', command=lambda: changeTag(self.tag1.get()), style='TButton')
        # button_1.grid(column=3, row=1)

        def insert(text1, duiying, row, col):
            tk.Label(frame4, text=text1, font=('', 8), padx=2, anchor='w').grid(row=row, column=col, padx=(5, 0),
                                                                                pady=5)
            tk.Entry(frame4, textvariable=duiying, width=6).grid(row=row, column=col + 1, padx=(0, 55), pady=5)

        items1 = ['箱体中心里程前缀', '里程', '里程方向 ', '左侧里程名', '右侧里程名', '轨面中心高程', '轨面坡度', '轨面箱顶距离', '轨面路肩距离']
        dy1 = [self.xiangti_licheng_qian, self.xiangti_licheng, self.licheng_fangxiang, self.zuo_li, self.you_li,
               self.guimian_ele, self.guimian_podu, self.guimian_xiangding, self.guimian_lujian]
        items2 = ['角度theta', '孔数', '孔径', '桥宽（斜的长度）', '净高', '顶板厚度', '底板厚度', '中腹板厚度', '边腹板', '顶板倒角长',
                  '顶板倒角高', '底板倒角长', '底板倒角高', '顶面铺设混凝土厚度', '基础厚', '基础超出箱体两侧距离']
        dy2 = [self.xiangti_theta, self.kongshu, self.xiangti_kongjin, self.xiangti_qiaokuan, self.xiangti_jinggao,
               self.xiangti_dingban, self.xiangti_diban, self.xiangti_midfu, self.xiangti_bianfu,
               self.xiangti_dingdao_c,
               self.xiangti_dingdao_g, self.xiangti_didao_c, self.xiangti_didao_g, self.xiangti_dingpu,
               self.xiangti_jichuhou, self.xiangti_jichuchao]
        items3 = ['翼墙左侧偏角α', '翼墙右侧偏角α2', '左翼墙大剖面高', '左翼墙小剖面高', '右翼墙大剖面高', '右翼墙小剖面高', '翼墙底部高度',
                  '翼墙顶部宽度', '翼墙顶部帽石边宽', '翼墙底左侧宽度', '翼墙底右侧宽度', '翼墙坡度', '翼墙放坡', '中间沉降缝个数',
                  '边沉降缝个数', '箱体延伸']
        dy3 = [self.yiqiang_afa, self.yiqiang_afa2, self.yiqiang_22, self.yiqiang_44, self.yiqiang_33, self.yiqiang_55,
               self.yiqiang_di, self.yiqiang_dingkuan, self.yiqiang_dingmaoshi, self.yiqiang_di_zuo,
               self.yiqiang_di_you,
               self.yiqiang_podu, self.yiqiang_fangpo, self.mid_chengjiang, self.mid_num, self.xiangti_yanshen]

        imageLabel = tk.Label(frame4, image=self.photo).grid(row=0, column=4, columnspan=2, rowspan=5,
                                                             padx=(0, 0), pady=5)

        self.btn_exzample = tk.Button(frame4, text='输入范例', font=('黑体', 10), bg="Honeydew", fg="blue")
        self.btn_exzample.grid(row=5 + len(items1), column=4, columnspan=1, rowspan=1, padx=(0, 0))
        self.btn_export = tk.Button(frame4, text='保存数据',font=('黑体', 10), bg="Honeydew", fg="blue")
        self.btn_export.grid(row=6 + len(items1), column=4, columnspan=1, rowspan=1, padx=(0, 0))
        self.btn_ok1 = ttk.Button(frame4, text=' 开   始 \n 绘   图 ', style='TButton')
        self.btn_ok1.grid(row=5 + len(items1), column=5, columnspan=1, rowspan=2, padx=(0, 0), ipadx=0, ipady=2)
        for i in range(len(items1)):
            insert(items1[i], dy1[i], i + 5, 4)
        for i in range(len(items2)):
            insert(items2[i], dy2[i], i, 0)
        for i in range(len(items3)):
            insert(items3[i], dy3[i], i, 2)

        self.text_out = tk.Text(frame4, height=6)
        self.text_out.grid(row=len(items3) + 1, column=0, columnspan=6, rowspan=1, padx=(0, 0))

        style = ttk.Style()
        style.configure("TButton", font=("Calibri", 12, "bold"), borderwidth=4, foreground="green", background="blue",
                        relief='RAISED')

        style1 = ttk.Style()
        style1.configure("Tutton", font=("Calibri", 12, "bold"), borderwidth=4, foreground="yellow",
                         background="black",
                         relief='RAISED')
        # Changes will be reflected
        # by the movement of mouse.
        style.map("TButton", foreground=[("active", "disabled", "green")],
                  background=[("active", "black")])

        def selectPath():
            path_ = filedialog.askopenfilename(filetypes=[('Excel', 'xlsx'), ('Excel', 'xls'), ('All Files', '*')],
                                               initialdir=self.file_dir)
            self.excel_path.set(path_)
            self.excel_real_path = self.excel_path.get()
            self.file_dir = os.path.dirname(self.excel_real_path)
            print(self.excel_real_path)

        excel_frm = tk.Frame(frame5, relief='ridge', borderwidth=1)
        excel_frm.pack(padx=0, pady=5, anchor='nw', fill=tk.X)
        Entyr_dwg = tk.Entry(excel_frm, textvariable=self.excel_path, width=52).pack(padx=(15, 0), pady=0, side='left')
        self.btn_select = tk.Button(excel_frm, text="文件选择", command=selectPath, bg="Honeydew", fg="black")
        self.btn_select.pack(padx=(2, 2), pady=5, side='left', fill='y')
        self.btn_ok = ttk.Button(excel_frm, text='一键绘图', style='TButton')
        self.btn_ok.pack(padx=(16, 2), pady=(2, 2), side='left')

        self.text_out1 = tk.Text(frame5, height=10)

        self.text_out1.pack(padx=0, pady=5, fill='y')

        # 开始定义钢筋图输入
        # ********************************************************************************************开始frame6界面定义
        def insert_6(win, text1, duiying, row, col):
            tk.Label(win, text=text1, font=('', 8), padx=2, anchor='w').grid(row=row, column=col, padx=(5, 0), pady=5)
            tk.Entry(win, textvariable=duiying, width=6).grid(row=row, column=col + 1, padx=(0, 55), pady=5)

        frm5_1 = tk.Frame(frame6, relief='ridge', borderwidth=1)
        frm5_1.grid(row=0, column=0, columnspan=6, rowspan=6, padx=2, pady=(4, 4))

        dy4 = [self.N1_1, self.N1_2, self.N2, self.N3, self.N4, self.N5, self.N6, self.N7, self.Nmid_7, self.N16,
               self.N17,
               self.N12, self.jianju_N12, self.N8_9, self.N_jichu, self.baohu]
        items4 = ['底板顶钢筋N1-1', '顶板底钢筋N1-2', '底板弯起钢筋N2', '顶板弯起钢筋N3', '顶底板外钢筋N4', '腹板外侧钢筋N5',
                  '腹板外侧钢筋N6', '腹板内侧钢筋N7', '*中腹板钢筋N7', '*底板加强筋N16', '底板加强筋N17', '分布钢筋N12',
                  '分布筋间距（cm）', '倒角钢筋N8、N9', '基础钢筋', '净保护层厚度（cm）']
        for i in range(len(items4)):
            if i < 11:
                insert_6(frm5_1, items4[i], dy4[i], i, 0)
            else:
                insert_6(frm5_1, items4[i], dy4[i], i - 11, 2)
        texts = ['顶板架立筋N10：HPB300φ16', '底板架立筋N11：HPB300φ16', '腹拉筋N13：HPB300φ10', '顶板拉筋N14：HPB300φ10',
                 '底板拉筋N15：HPB300φ10']
        for i in range(len(texts)):
            tk.Label(frm5_1, text=texts[i], font=('', 8)).grid(row=5 + i, column=2, padx=(0, 55), pady=5, columnspan=2)
        frm5_2 = tk.Frame(frame6, relief='ridge', borderwidth=1)
        frm5_2.grid(row=0, column=7, columnspan=3, rowspan=6, padx=2, pady=(4, 4))
        image_gangjing = tk.Label(frm5_2, image=self.photo2).grid(row=0, column=4, columnspan=2, rowspan=5,
                                                                  padx=(0, 0), pady=0)
        self.btn_pj = ttk.Button(frm5_2, text='生成数据', style='TButton')
        self.btn_pj.grid(row=5, column=4, columnspan=2, rowspan=3, padx=(0, 0))
        # 开始数量表提取 ************************************************************************************************
        frm7_1 = tk.Frame(frame7, relief='ridge', borderwidth=1)
        frm7_1.pack(padx=0, pady=0, fill=tk.X)
        frm7_2 = tk.Frame(frame7, relief='ridge', borderwidth=1, height=350)
        frm7_2.pack(padx=0, pady=0, fill=tk.X)
        self.btn_shuliang = ttk.Button(frm7_1, text='生成本工点数量表', style='TButton')
        self.btn_shuliang.pack(padx=(0, 0), pady=20)
        # self.btn_ok4 = ttk.Button(frm7_1, text='生成全线工点数量表', style='TButton')
        # self.btn_ok4.pack(padx=(0, 0), pady=20)
        tk.Label(frm7_2, text='后续升级预留\n\n\n\n\n\n', font=('', 8)).pack(pady=20)

        # ***********************************************************************开始定义初始值
        def set_ini(vir, value):
            vir.set(value)

        vars = [self.N1_1, self.N1_2, self.N2, self.N3, self.N4, self.N5, self.N6, self.N7, self.Nmid_7, self.N16,
                self.N17, self.N12, self.jianju_N12, self.N8_9, self.N_jichu, self.baohu]
        values = [22, 25, 28, 25, 16, 25, 25, 16, 25, 22, 20, 12, 20, 16, 12, 3.5]
        for i in range(len(vars)):
            set_ini(vars[i], str(values[i]))


if __name__ == '__main__':
    lg = Login()
    lg.title('箱形桥智能设计成图软件')
    lg.welcome.set('- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -')
    # lg.welcome.set('-------------------------------------------------------------------------------')

