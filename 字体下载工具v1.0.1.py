import os
import sys
import re
import psutil
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from PIL import Image, ImageDraw, ImageFont, ImageTk
import shutil
from appdirs import user_data_dir

class FontPreviewApp:
    def __init__(self, root):
        self.root = root
        self.root.withdraw()  # 隐藏主窗口，直到授权成功
        self.root.title("沧海字体下载工具v1.0.1 by:沧海")
        self.root.geometry("1200x900")
        self.root.resizable(False, False)

        self.resource_path = self.get_resource_path()

        self.icon_photo = ImageTk.PhotoImage(file=os.path.join(self.resource_path, 'icon.png'))
        self.root.iconphoto(True, self.icon_photo)  # 设置窗口图标

        self.fonts_dir = ""
        self.font_list = []
        self.filtered_font_list = []

        self.check_authorization()

    def get_resource_path(self):
        """ Get the absolute path to the resource, works for dev and for PyInstaller """
        if hasattr(sys, '_MEIPASS'):
            return sys._MEIPASS
        return os.path.abspath(".")

    def get_config_file_path(self, filename):
        """ 获取配置文件路径 """
        data_dir = user_data_dir('Canghai0.4', 'Canghai')
        os.makedirs(data_dir, exist_ok=True)
        return os.path.join(data_dir, filename)

    def get_mac_address(self):
        """ 获取电脑MAC地址 """
        mac_addresses = []
        for interface, addrs in psutil.net_if_addrs().items():
            for addr in addrs:
                if addr.family == psutil.AF_LINK:
                    mac = addr.address
                    if re.match(r'^([0-9A-Fa-f]{2}[:-]){5}([0-9A-Fa-f]{2})$', mac):
                        mac_addresses.append(mac)
        return mac_addresses

    def get_auth_file_path(self):
        """ 生成授权码文件路径 """
        return self.get_config_file_path('canghaiziti.txt')

    def check_authorization(self):
        """ 检查授权码 """
        auth_file = self.get_auth_file_path()
        mac_addresses = self.get_mac_address()

        if os.path.exists(auth_file):
            with open(auth_file, 'r') as file:
                stored_mac = file.read().strip()
                if stored_mac in mac_addresses:
                    self.root.deiconify()
                    self.create_widgets()
                    return True

        self.show_auth_window(mac_addresses[0])  # 默认显示第一个MAC地址用于授权
        return False

    def show_auth_window(self, mac_address):
        """ 显示授权码输入窗口 """
        auth_window = tk.Toplevel(self.root)
        auth_window.title("授权码")
        auth_window.geometry("330x380")
        auth_window.resizable(False, False)
        auth_window.iconphoto(False, self.icon_photo)

        def submit_auth_code():
            auth_code = auth_entry.get()
            if auth_code == "ziti8":
                with open(self.get_auth_file_path(), 'w') as file:
                    file.write(mac_address)
                auth_window.destroy()
                self.root.deiconify()
                self.create_widgets()
            else:
                error_label.config(text="授权码不正确，请微信扫描二维码获取")

        tk.Label(auth_window, text="请输入授权码：").pack(pady=5)
        auth_entry = tk.Entry(auth_window)
        auth_entry.pack(pady=5)
        tk.Button(auth_window, text="确定", width=10, command=submit_auth_code).pack(pady=10)

        # 增加一个显示错误消息的标签
        error_label = tk.Label(auth_window, text="", fg="red")
        error_label.pack(pady=5)

        qr_image = Image.open(os.path.join(self.resource_path, "qrcode.jpg"))
        qr_image.thumbnail((150, 150))
        qr_image_tk = ImageTk.PhotoImage(qr_image)
        qr_label = tk.Label(auth_window, image=qr_image_tk)
        qr_label.image = qr_image_tk
        qr_label.pack(pady=5)

        tk.Label(auth_window, text="请微信扫描二维码关注公众号回复\n'字体下载'获取授权码").pack(pady=5)

        self.root.wait_window(auth_window)

    def create_widgets(self):
        # 创建菜单
        menubar = tk.Menu(self.root)
        about_menu = tk.Menu(menubar, tearoff=0)
        about_menu.add_command(label="软件版本：v1.0.1")
        about_menu.add_command(label="联系网站", command=self.open_website)
        about_menu.add_command(label="微信公众号", command=self.show_qrcode)
        menubar.add_cascade(label="关于我们", menu=about_menu)
        self.root.config(menu=menubar)

        # 添加选择字体文件夹按钮
        self.select_folder_button = ttk.Button(self.root, text="选择字体文件夹", command=self.select_font_folder)
        self.select_folder_button.grid(row=0, column=0, padx=10, pady=5, sticky="w")

        # 添加显示选中文件夹路径的标签
        self.folder_path_label = ttk.Label(self.root, text="未选择文件夹")
        self.folder_path_label.grid(row=1, column=0, padx=10, pady=5, sticky="w")

        # 搜索框
        self.search_label = ttk.Label(self.root, text="搜索字体", font=('Heiti', 14, 'bold'))
        self.search_label.grid(row=2, column=0, padx=10, pady=5, sticky="w")

        self.search_entry = ttk.Entry(self.root, width=40)
        self.search_entry.grid(row=3, column=0, padx=10, pady=5, sticky="w")
        self.search_entry.bind("<KeyRelease>", self.search_fonts)

        # 添加选择字体标签
        self.label_font_list = ttk.Label(self.root, text="选择字体", font=('Heiti', 14, 'bold'))
        self.label_font_list.grid(row=4, column=0, padx=10, pady=5, sticky="w")

        # 添加滚动条
        self.font_listbox_frame = tk.Frame(self.root)
        self.font_listbox_frame.grid(row=5, column=0, padx=10, pady=10, sticky="w")

        self.font_listbox_scrollbar = tk.Scrollbar(self.font_listbox_frame, orient=tk.VERTICAL)
        self.font_listbox = tk.Listbox(self.font_listbox_frame, height=30, width=40, yscrollcommand=self.font_listbox_scrollbar.set)
        self.font_listbox_scrollbar.config(command=self.font_listbox.yview)

        self.font_listbox.pack(side=tk.LEFT, fill=tk.BOTH)
        self.font_listbox_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.font_listbox.bind("<<ListboxSelect>>", self.on_font_select)

        # 添加自定义文本输入提示标签
        self.label_custom_text = ttk.Label(self.root, text="在下方输入你想要预览的文字（最多50字）", font=('Heiti', 12))
        self.label_custom_text.grid(row=6, column=0, padx=10, pady=5, sticky="w")

        self.custom_text_entry = tk.Text(self.root, height=6, width=40, wrap=tk.WORD)
        self.custom_text_entry.grid(row=7, column=0, padx=10, pady=10, sticky="w")
        self.custom_text_entry.bind("<KeyRelease>", self.update_preview)

        # 替换预览字体信息处的图片
        self.banner_image = ImageTk.PhotoImage(file=os.path.join(self.resource_path, "hengfu.png"))  # 替换为你的图片路径
        self.banner_label = tk.Label(self.root, image=self.banner_image)
        self.banner_label.place(x=338, y=10)  # 自定义x和y轴
        #预览框大小
        self.preview_canvas = tk.Canvas(self.root, width=800, height=690, bg="white", highlightthickness=1, highlightbackground="black")
        self.preview_canvas.grid(row=1, column=1, rowspan=7, padx=10, pady=10, sticky="w")

        # 初始状态下在预览画布中显示提示文字
        self.preview_canvas.create_text(400, 350, text="请左上角选择字体文件夹", fill="grey", font=("Heiti", 48), anchor="center")

        self.font_count_label = ttk.Label(self.root, text="")
        self.font_count_label.grid(row=8, column=0, padx=10, pady=5, sticky="w")

        # 自定义按钮图片
        button_image = Image.open(os.path.join(self.resource_path, "anniu.png"))  # 替换为你的按钮图片路径
        button_photo = ImageTk.PhotoImage(button_image)

        self.install_button = tk.Button(self.root, image=button_photo, command=self.install_font, borderwidth=0)
        self.install_button.image = button_photo  # 保持对图片的引用
        self.install_button.place(x=548, y=840)  # 自定义x和y轴

        # 添加文字介绍
        self.label_info = ttk.Label(self.root, text="提示：\n此软件为本人原创软件，未经允许不得擅自篡改软件、售卖等。软件支持TTF、otf、ttf、otf、woff、woff2、eot、svg格式\n可自行网上下载字体统一放入字体文件夹，注意字体文件夹内不要有文件夹", font=('Heiti', 11), foreground='red', anchor="center")
        self.label_info.place(x=310, y=790, width=850, height=48)  # 自定义x、y轴和宽高

    def open_website(self):
        import webbrowser
        webbrowser.open("https://www.wangyuntian.com/")

    def show_qrcode(self):
        top = tk.Toplevel(self.root)
        top.title("微信公众号")
        qrcode_image = ImageTk.PhotoImage(file=os.path.join(self.resource_path, "qrcode.jpg"))  # 替换为你的二维码图片路径
        qrcode_label = tk.Label(top, image=qrcode_image)
        qrcode_label.image = qrcode_image  # 保持对图片的引用
        qrcode_label.pack(padx=10, pady=10)

    def select_font_folder(self):
        folder_selected = filedialog.askdirectory()
        if folder_selected:
            self.fonts_dir = folder_selected
            self.folder_path_label.config(text=self.fonts_dir)
            self.load_fonts()

    def load_fonts(self):
        def find_fonts(directory):
            font_names = set()
            fonts = []
            for root, _, files in os.walk(directory):
                for filename in files:
                    if any(filename.lower().endswith(ext) for ext in supported_extensions):
                        font_name = filename.split(".")[0]
                        if font_name not in font_names:
                            font_names.add(font_name)
                            fonts.append(os.path.join(root, filename))  # 存储完整路径
            return fonts

        self.font_listbox.delete(0, tk.END)
        self.font_list = []
        supported_extensions = [".ttf", ".otf", ".woff", ".woff2", ".eot", ".svg"]  # 增加支持的字体格式后缀
        fonts = find_fonts(self.fonts_dir)
        for idx, font in enumerate(fonts):
            font_filename = os.path.basename(font)
            self.font_list.append(font)
            self.font_listbox.insert(tk.END, f"{idx+1}. {font_filename}")

        if self.font_list:
            self.font_count_label.config(text=f"目前共有 {len(self.font_list)} 个字体可安装")
        else:
            self.font_count_label.config(text="未找到字体文件")
        self.filtered_font_list = self.font_list.copy()

    def search_fonts(self, event):
        search_term = self.search_entry.get().lower()
        self.font_listbox.delete(0, tk.END)
        self.filtered_font_list = [font for font in self.font_list if search_term in font.lower()]
        for idx, font in enumerate(self.filtered_font_list):
            font_filename = os.path.basename(font)
            self.font_listbox.insert(tk.END, f"{idx+1}. {font_filename}")
        self.font_count_label.config(text=f"找到 {len(self.filtered_font_list)} 个匹配字体")

    def on_font_select(self, event):
        if self.font_listbox.curselection():
            selected_index = self.font_listbox.curselection()[0]
            font_path = self.font_list[selected_index]
            self.preview_font(font_path)

    def update_preview(self, event):
        text = self.custom_text_entry.get("1.0", tk.END).strip()
        if len(text) > 50:
            self.custom_text_entry.delete("1.0", tk.END)
            self.custom_text_entry.insert(tk.END, text[:50])
            text = text[:50]
        if self.font_listbox.curselection():
            selected_index = self.font_listbox.curselection()[0]
            font_path = self.font_list[selected_index]
            self.preview_font(font_path)

    def preview_font(self, font_path):
        sizes = [12, 24, 36, 48, 60]
        default_preview_text = "0123456789\nABCDEFGHIJKLMNOPQRSTUVWXYZ\nabcdefghijklmnopqrstuvwxyz\nCANGHAI\n沧海字体下载工具-方便实用"
        custom_text = self.custom_text_entry.get("1.0", tk.END).strip()

        preview_text = custom_text if custom_text else default_preview_text

        image = Image.new("RGB", (800, 690), "white")
        draw = ImageDraw.Draw(image)

        # 获取实际字体名称
        font = ImageFont.truetype(font_path, 12)
        actual_font_name = font.getname()[0]

        # 使用微软雅黑字体显示实际字体名称
        font_yahei = ImageFont.truetype("msyh.ttc", 16)  # 这里假设微软雅黑字体文件为 msyh.ttc，请确保文件路径正确

        y_position = 0
        draw.text((10, y_position), f"字体名称: {actual_font_name}", fill="black", font=font_yahei)
        y_position += draw.textsize(f"字体名称: {actual_font_name}", font=font_yahei)[1] + 10

        for size in sizes:
            font = ImageFont.truetype(font_path, size)
            draw.text((10, y_position), f"{size}:", fill="black", font=font)
            y_position += draw.textsize(f"{size}:", font=font)[1] + 5
            draw.text((10, y_position), preview_text, fill="black", font=font)
            text_height = draw.textsize(preview_text, font=font)[1]
            y_position += text_height + 10

        self.preview_image = ImageTk.PhotoImage(image)
        self.preview_canvas.delete("all")  # 删除所有内容
        self.preview_canvas.create_image(0, 0, anchor=tk.NW, image=self.preview_image)

    def install_font(self):
        if self.font_listbox.curselection():
            selected_index = self.font_listbox.curselection()[0]
            font_path = self.font_list[selected_index]
            dest_path = os.path.join("C:\\Windows\\Fonts", os.path.basename(font_path))
            try:
                shutil.copy(font_path, dest_path)
                # Register the font in the system
                import ctypes
                ctypes.windll.gdi32.AddFontResourceW(dest_path)
                messagebox.showinfo("成功", f"{os.path.basename(font_path)} 已成功安装")
            except Exception as e:
                messagebox.showerror("错误", str(e))

if __name__ == "__main__":
    root = tk.Tk()
    style = ttk.Style()
    style.configure("Accent.TButton", foreground="white", background="#0078D4", font=('Helvetica', 12))
    app = FontPreviewApp(root)
    root.mainloop()
