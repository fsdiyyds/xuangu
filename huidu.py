from PIL import Image
import numpy as np
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
from win32api import MessageBox
from win32con import MB_OK


class ui(tk.Tk):
    """窗口UI"""

    def __init__(self):
        """初始化"""
        super().__init__()  # 有点相当于tk.Tk()

        self.file_dir = []
        self.file = tk.StringVar()
        self.WIDTH = 640
        self.HEIGHT = 500
        self.excel_path = tk.StringVar()
        self.excel_real_path = []

        self.btn_import = ttk.Button()
        self.btn_ok = ttk.Button()
        self.run()

    def import_file(self):
        file_op = filedialog.asksaveasfilename(filetypes=[('picture', 'jpeg'),
                                                          ('picture', 'jpg'),
                                                          ('picture', 'png')],
                                               initialdir=self.file_dir)

    def run(self):
        self.title('照片转手绘工具')
        self.option_add('*Font', ('', 12))
        ttk.Style().configure(".", font=('', 12))

        frame2 = tk.Frame(self)
        frame2.pack(fill=tk.Y, pady=0)

        self.ws = frame2.winfo_screenwidth()
        self.hs = frame2.winfo_screenheight()
        x = (self.ws / 2) - (self.WIDTH / 2)
        y = (self.hs / 2) - (self.HEIGHT / 2)
        self.config(width=self.WIDTH)

        def selectPath():
            self.excel_real_path = []
            path_ = filedialog.askopenfilenames(filetypes=[('picture', 'jpeg'),
                                                           ('picture', 'jpg'),
                                                           ('picture', 'png')],
                                                initialdir=self.file_dir)
            self.excel_path.set(path_)

            mid = self.excel_path.get().replace('(', '').replace(')', '')
            for i in mid.split(','):
                self.excel_real_path.append(i)
            if '' in self.excel_real_path:
                self.excel_real_path.remove('')
            print("文件名字序列",self.excel_real_path)

        def convert():
            for pic in self.excel_real_path:
                pic1 = ''.join(i for i in pic if i not in '"')
                pic1 = eval(pic1)
                print('图片名字为',pic1, type(pic1))
                a = np.asarray(Image.open(pic1).convert('L')).astype('float')
                depth = 10.  # (0-100)
                grad = np.gradient(a)  # 取图像灰度的梯度值
                grad_x, grad_y = grad  # 分别取横纵图像梯度值
                grad_x = grad_x * depth / 100.
                grad_y = grad_y * depth / 100.
                A = np.sqrt(grad_x ** 2 + grad_y ** 2 + 1.)
                uni_x = grad_x / A
                uni_y = grad_y / A
                uni_z = 1. / A

                vec_el = np.pi / 2.2  # 光源的俯视角度，弧度值
                vec_az = np.pi / 4.  # 光源的方位角度，弧度值
                dx = np.cos(vec_el) * np.cos(vec_az)  # 光源对x 轴的影响
                dy = np.cos(vec_el) * np.sin(vec_az)  # 光源对y 轴的影响
                dz = np.sin(vec_el)  # 光源对z 轴的影响

                b = 255 * (dx * uni_x + dy * uni_y + dz * uni_z)  # 光源归一化
                b = b.clip(0, 255)

                im = Image.fromarray(b.astype('uint8'))  # 重构图像
                file_type = ['.jpg', '.jpeg', '.png']
                for filetp in file_type:
                    if len(pic1.split(filetp)) > 1:
                        path_out = pic1.split(filetp)[0] + '转化后' + filetp
                        print("输出的文件名字为",path_out)
                        print(pic1.split('.'))
                        im.save(path_out)
            MessageBox(0, "转化成功！谢谢使用！", "提醒", MB_OK)

        excel_frm = tk.Frame(frame2, relief='ridge', borderwidth=1)
        excel_frm.pack(padx=0, pady=5, anchor='nw', fill=tk.X)
        tk.Label(excel_frm, text='可以选择一张或者同时选择多张(文件名不得带括号)',
                 font=('', 10)).pack(padx=(2, 2), pady=5, fill='x')

        self.btn_select = tk.Button(excel_frm, text="选择照片", command=selectPath,
                                    bg="Honeydew", fg="black")
        self.btn_select.pack(padx=(2, 2), pady=5)

        Entyr_dwg = tk.Entry(frame2, textvariable=self.excel_path,
                             width=52).pack(padx=(15, 0), pady=(5, 20), side='left')
        self.btn_ok = ttk.Button(frame2, text='一键转化', style='TButton')
        self.btn_ok.configure(command=convert)
        self.btn_ok.pack(padx=(10, 10), pady=(5, 20), side='left')


if __name__ == '__main__':
    lg = ui()
    lg.title('照片转手绘工具')
    lg.mainloop()