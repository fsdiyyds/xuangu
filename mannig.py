import tkinter as tk
from tkinter import ttk
import math
from win32api import MessageBox
from win32con import MB_OK


class FlowCalculator(tk.Frame):
    def __init__(self, master=None):
        super().__init__(master)
        self.master = master
        self.master.title("明渠流量计算器")
        self.setup_widgets()

    def setup_widgets(self):
        # 设置输入窗口
        self.input_frame = ttk.Frame(self.master, padding=10)
        self.input_frame.grid(row=0, column=0, sticky="nsew")
        self.input_variables = []

        # 设置截面类型
        self.section_label = ttk.Label(self.input_frame, text="截面类型：")
        self.section_label.grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.section_values = ["圆形", "矩形", "U型", "马蹄形"]
        self.section_var = tk.StringVar(value=self.section_values[0])
        self.section_combobox = ttk.Combobox(self.input_frame, textvariable=self.section_var,
                                             values=self.section_values, state="readonly")
        self.section_combobox.grid(row=0, column=1, padx=5, pady=5, sticky="we")
        self.section_combobox.bind("<<ComboboxSelected>>", self.on_section_selected)
        self.parameter_labels = []
        self.parameter_entries = []
        self.parameter_vars = []

        # 设置截面参数
        self.width_label = ttk.Label(self.input_frame, text="宽或直径：")
        self.width_label.grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.width_var = tk.StringVar()
        self.width_entry = ttk.Entry(self.input_frame, textvariable=self.width_var)
        self.width_entry.grid(row=1, column=1, padx=5, pady=5, sticky="we")
        self.input_variables.append(self.width_entry)
        self.parameter_entries.append(self.width_entry)
        self.parameter_vars.append(self.width_var)

        self.height_label = ttk.Label(self.input_frame, text="高：")
        self.height_label.grid(row=2, column=0, padx=5, pady=5, sticky="w")
        self.height_var = tk.StringVar()
        self.height_entry = ttk.Entry(self.input_frame, textvariable=self.height_var)
        self.height_entry.grid(row=2, column=1, padx=5, pady=5, sticky="we")
        self.input_variables.append(self.height_entry)
        self.parameter_entries.append(self.height_entry)
        self.parameter_vars.append(self.height_var)

        self.parameter_values = {
            "矩形": ["宽", "高"],
            "U型": ["上底宽", "高", "腹板厚度", "弧形截面高度"],
            "马蹄形": ["上宽", "下宽", "腹板厚度", "弧形截面高度"]
        }

        row = 3
        for param_name in self.parameter_values.get(self.section_var.get(), []):
            label = ttk.Label(self.input_frame, text=param_name + "：")
            label.grid(row=row, column=0, padx=5, pady=5, sticky="w")
            self.parameter_labels.append(label)

            var = tk.StringVar()
            entry = ttk.Entry(self.input_frame, textvariable=var)
            entry.grid(row=row, column=1, padx=5, pady=5, sticky="we")
            self.input_variables.append(entry)
            self.parameter_entries.append(entry)
            self.parameter_vars.append(var)

            row += 1

        # 设置计算按钮
        self.calculate_button = ttk.Button(self.master, text="计算", command=self.calculate_flow)
        self.calculate_button.grid(row=1, column=0, padx=10, pady=10)

        # 设置输出窗口
        self.output_frame = ttk.Frame(self.master, padding=10)
        self.output_frame.grid(row=0, column=1, rowspan=2, sticky="nsew")

        self.R_label = ttk.Label(self.output_frame, text="湿周半径：")
        self.R_label.grid(row=0, column=0, padx=5, pady=5, sticky="w")

        self.A_label = ttk.Label(self.output_frame, text="横截面积：")
        self.A_label.grid(row=1, column=0, padx=5, pady=5, sticky="w")

        self.Q_label = ttk.Label(self.output_frame, text="明渠流量：")
        self.Q_label.grid(row=2, column=0, padx=5, pady=5, sticky="w")

        self.R = ttk.Label(self.output_frame, text="", font="-weight bold")
        self.R.grid(row=0, column=1, padx=5, pady=5, sticky="e")

        self.A = ttk.Label(self.output_frame, text="", font="-weight bold")
        self.A.grid(row=1, column=1, padx=5, pady=5, sticky="e")

        self.Q = ttk.Label(self.output_frame, text="", font="-weight bold")
        self.Q.grid(row=2, column=1, padx=5, pady=5, sticky="e")

        self.slope_label = ttk.Label(self.input_frame, text="底面坡度（‰）：")
        self.slope_label.grid(row=row + 1, column=0, padx=5, pady=5, sticky="w")
        self.slope_var = tk.StringVar(value="10")
        self.slope_entry = ttk.Entry(self.input_frame, textvariable=self.slope_var)
        self.slope_entry.grid(row=row + 1, column=1, padx=5, pady=5, sticky="we")
        self.input_variables.append(self.slope_entry)

        self.friction_label = ttk.Label(self.input_frame, text="Manning粗糙系数：")
        self.friction_label.grid(row=row + 2, column=0, padx=5, pady=5, sticky="w")
        self.friction_var = tk.StringVar(value="0.03")
        self.friction_entry = ttk.Entry(self.input_frame, textvariable=self.friction_var)
        self.friction_entry.grid(row=row + 2, column=1, padx=5, pady=5, sticky="we")
        self.input_variables.append(self.friction_entry)

        self.shear_label = ttk.Label(self.input_frame, text="装置系数（‰）：")
        self.shear_label.grid(row=row + 3, column=0, padx=5, pady=5, sticky="w")
        self.shear_var = tk.StringVar(value="0")
        self.shear_entry = ttk.Entry(self.input_frame, textvariable=self.shear_var)
        self.shear_entry.grid(row=row + 3, column=1, padx=5, pady=5, sticky="we")
        self.input_variables.append(self.shear_entry)

        self.calculate_button = ttk.Button(self.input_frame, text="计算流量", command=self.calculate_flow)
        self.calculate_button.grid(row=row + 4, column=0, padx=5, pady=5, sticky="w")

        # 设置输出结果窗口
        self.output_frame = ttk.Frame(self.master, padding=10)
        self.output_frame.grid(row=0, column=1, sticky="nsew")

        self.R_label = ttk.Label(self.output_frame, text="水槽水力半径R：")
        self.R_label.grid(row=0, column=0, padx=5, pady=5, sticky="w")

        self.A_label = ttk.Label(self.output_frame, text="过水面积A：")
        self.A_label.grid(row=1, column=0, padx=5, pady=5, sticky="w")

        self.Q_label = ttk.Label(self.output_frame, text="流量Q：")
        self.Q_label.grid(row=2, column=0, padx=5, pady=5, sticky="w")

        self.R = ttk.Label(self.output_frame, text="", font="-weight bold")
        self.R.grid(row=0, column=1, padx=5, pady=5, sticky="w")

        self.A = ttk.Label(self.output_frame, text="", font="-weight bold")
        self.A.grid(row=1, column=1, padx=5, pady=5, sticky="w")

        self.Q = ttk.Label(self.output_frame, text="", font="-weight bold")
        self.Q.grid(row=2, column=1, padx=5, pady=5, sticky="w")

    def on_section_selected(self, event):
        """当截面类型选择框中的值改变时，更新界面"""
        section = self.section_var.get()
        width_label = "宽或直径" if section == "圆形" else "上底宽"
        self.width_label.config(text=width_label + "：")

        # 删除原有参数输入框
        for label in self.parameter_labels:
            label.destroy()
        for entry in self.parameter_entries:
            entry.destroy()

        # 添加新的参数输入框
        self.parameter_labels = []
        self.parameter_entries = []
        self.parameter_vars = []

        row = 3
        for param_name in self.parameter_values.get(section, []):
            label = ttk.Label(self.input_frame, text=param_name + "：")
            label.grid(row=row, column=0, padx=5, pady=5, sticky="w")
            self.parameter_labels.append(label)

            var = tk.StringVar()
            entry = ttk.Entry(self.input_frame, textvariable=var)
            entry.grid(row=row, column=1, padx=5, pady=5, sticky="we")
            self.input_variables.append(entry)
            self.parameter_entries.append(entry)
            self.parameter_vars.append(var)

            row += 1

        # 删除原有其他参数输入框
        self.slope_label.destroy()
        self.slope_entry.destroy()

        self.friction_label.destroy()
        self.friction_entry.destroy()

        self.shear_label.destroy()
        self.shear_entry.destroy()

        # 添加新的其他参数输入框
        self.slope_label = ttk.Label(self.input_frame, text="底坡（千分之一）：")
        self.slope_label.grid(row=row, column=0, padx=5, pady=5, sticky="w")
        self.slope_var = tk.StringVar(value="0")
        self.slope_entry = ttk.Entry(self.input_frame, textvariable=self.slope_var)
        self.slope_entry.grid(row=row, column=1, padx=5, pady=5, sticky="we")
        self.input_variables.append(self.slope_entry)

        row += 1

        self.friction_label = ttk.Label(self.input_frame, text="Manning粗糙系数：")
        self.friction_label.grid(row=row, column=0, padx=5, pady=5, sticky="w")
        self.friction_var = tk.StringVar(value="0.022")
        self.friction_entry = ttk.Entry(self.input_frame, textvariable=self.friction_var)
        self.friction_entry.grid(row=row, column=1, padx=5, pady=5, sticky="we")
        self.input_variables.append(self.friction_entry)

        row += 1

        self.shear_label = ttk.Label(self.input_frame, text="装置系数（千分之一）：")
        self.shear_label.grid(row=row, column=0, padx=5, pady=5, sticky="w")
        self.shear_var = tk.StringVar(value="0")
        self.shear_entry = ttk.Entry(self.input_frame, textvariable=self.shear_var)
        self.shear_entry.grid(row=row, column=1, padx=5, pady=5, sticky="we")
        self.input_variables.append(self.shear_entry)

        self.calculate_button = ttk.Button(self.input_frame, text="计算流量", command=self.calculate_flow)
        self.calculate_button.grid(row=row + 1, column=0, padx=5, pady=5, sticky="w")

        # 设置输出结果窗口
        self.output_frame = ttk.Frame(self.master, padding=10)
        self.output_frame.grid(row=0, column=1, sticky="nsew")

        self.R_label = ttk.Label(self.output_frame, text="水槽水力半径R：")
        self.R_label.grid(row=0, column=0, padx=5, pady=5, sticky="w")

        self.A_label = ttk.Label(self.output_frame, text="过水面积A：")
        self.A_label.grid(row=1, column=0, padx=5, pady=5, sticky="w")

        self.Q_label = ttk.Label(self.output_frame, text="流量Q：")
        self.Q_label.grid(row=2, column=0, padx=5, pady=5, sticky="w")

        self.R = ttk.Label(self.output_frame, text="", font="-weight bold")
        self.R.grid(row=0, column=1, padx=5, pady=5, sticky="w")

        self.A = ttk.Label(self.output_frame, text="", font="-weight bold")
        self.A.grid(row=1, column=1, padx=5, pady=5, sticky="w")

        self.Q = ttk.Label(self.output_frame, text="", font="-weight bold")
        self.Q.grid(row=2, column=1, padx=5, pady=5, sticky="w")

    def calculate_flow(self):
        """计算流量并显示结果"""
        try:
            section = self.section_var.get()
            parameters = [float(var.get()) for var in self.parameter_vars]
            slope = float(self.slope_var.get())
            friction = float(self.friction_var.get())
            shear = float(self.shear_var.get())

            width = parameters[0]
            height = parameters[1] if section == "矩形" else None
            upper_width = parameters[0] if section == "U型" or section == "马蹄形" else None
            lower_width = parameters[1] if section == "马蹄形" else None
            belly_thickness = parameters[2] if section == "U型" or section == "马蹄形" else None
            arc_height = parameters[3] if section == "U型" or section == "马蹄形" else None

            # 计算

            if section == "圆形":
                R = width / 2
                A = math.pi * R ** 2
                P = 2 * math.pi * R
                T = P * math.sqrt(1 + slope ** 2)
            elif section == "矩形":
                A = width * height
                R = A / (2 * width + 2 * height)
                P = 2 * width + 2 * height
                T = P * math.sqrt(1 + slope ** 2)
            elif section == "U型":
                A = width * arc_height + (height - arc_height) * belly_thickness
                R = A / (width + 2 * arc_height + math.sqrt(width * height))
                P = width + 2 * arc_height + math.sqrt(width * height) + 2 * belly_thickness
                T = P * math.sqrt(1 + slope ** 2)
            elif section == "马蹄形":
                A = (upper_width + lower_width) * arc_height / 2 + \
                    (height - arc_height) * belly_thickness
                R = A / ((upper_width + lower_width) / 2 + arc_height - (upper_width - lower_width) ** 2 / (8 * arc_height))
                P = (upper_width + lower_width) / math.cos(math.atan((upper_width - lower_width) / (2 * arc_height))) + \
                    2 * math.sqrt(arc_height ** 2 + (upper_width - lower_width) ** 2 / 4)
                T = P * math.sqrt(1 + slope ** 2)
            else:
                raise ValueError("不支持的截面类型")

            v = friction / n * R ** (2 / 3) * abs(slope) ** (1 / 2)
            Q = A * v

            # 增加反映安装方式和位置的系数
            Q *= (1 + shear / 1000)

            # 显示结果
            self.R.config(text="{:.2f} m".format(R))
            self.A.config(text="{:.2f} m²".format(A))
            self.Q.config(text="{:.2f} m³/s".format(Q))
        except (ValueError, ZeroDivisionError) as e:
            print("errro")
        except Exception as e:
            print("errro")

    def clear_input(self):
        """清除所有输入"""
        for var in self.input_variables:
            var.set("")

    def clear_output(self):

        """清除所有输出"""
        self.R.config(text="")
        self.A.config(text="")
        self.Q.config(text="")

    def create_menu(self):
        """创建菜单"""
        menu = tk.Menu(self.master)
        self.master.config(menu=menu)

        file_menu = tk.Menu(menu, tearoff=0)
        menu.add_cascade(label="文件", menu=file_menu)
        file_menu.add_command(label="清除输入", command=self.clear_input)
        file_menu.add_command(label="清除输出", command=self.clear_output)
        file_menu.add_separator()
        file_menu.add_command(label="退出", command=self.master.quit)

    def run(self):
        self.master.mainloop()


if __name__ == "__main__":
    root = tk.Tk()
    app = FlowCalculator(master=root)
    app.create_menu()
    app.run()