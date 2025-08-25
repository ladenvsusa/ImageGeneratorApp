import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from PIL import Image, ImageOps
import os
import xlrd
import threading
import random

class MultiImageGeneratorApp:
    def __init__(self, master):
        self.master = master
        master.title("多源随机图片生成器 (v4.0 - Interactive XLS)")
        master.geometry("700x750") # 稍微增加窗口高度以容纳新控件

        # --- 初始化变量 ---
        self.input_image_paths = []
        self.output_dir = tk.StringVar()
        self.naming_file_path = tk.StringVar()
        self.num_to_generate = tk.IntVar(value=10)
        self.max_images = tk.IntVar(value=0)
        self.naming_mode = tk.StringVar(value="sequential")
        self.manual_name_count = tk.IntVar(value=0) # 新增: 存储手动命名数量

        # --- UI 布局 ---
        main_frame = ttk.Frame(master, padding="10")
        main_frame.pack(fill="both", expand=True)

        # 1. 输入图片选择 (与v3.0相同)
        ttk.Label(main_frame, text="1. 添加源图片 (最多30张):").grid(row=0, column=0, sticky="w", pady=5)
        list_frame = ttk.Frame(main_frame)
        list_frame.grid(row=1, column=0, columnspan=3, sticky="ew")
        self.image_listbox = tk.Listbox(list_frame, height=10)
        self.image_listbox.pack(side="left", fill="both", expand=True)
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=self.image_listbox.yview)
        scrollbar.pack(side="right", fill="y")
        self.image_listbox.config(yscrollcommand=scrollbar.set)
        btn_frame = ttk.Frame(main_frame)
        btn_frame.grid(row=2, column=0, columnspan=3, sticky="w", pady=5)
        ttk.Button(btn_frame, text="添加图片...", command=self.add_input_images).pack(side="left")
        ttk.Button(btn_frame, text="清空列表", command=self.clear_input_list).pack(side="left", padx=10)

        ttk.Label(main_frame, text="根据当前源图片，最多可生成唯一图片总数:").grid(row=3, column=0, columnspan=2, sticky="w", pady=(10,0))
        self.max_label = ttk.Label(main_frame, textvariable=self.max_images, font=("Helvetica", 10, "bold"))
        self.max_label.grid(row=4, column=0, columnspan=2, sticky="w")
        
        # 2. 输出位置 (与v3.0相同)
        ttk.Label(main_frame, text="2. 选择图片保存位置:").grid(row=5, column=0, sticky="w", pady=5)
        self.output_label = ttk.Label(main_frame, text="尚未选择位置", wraplength=550)
        self.output_label.grid(row=6, column=0, columnspan=2, sticky="w")
        ttk.Button(main_frame, text="浏览...", command=self.select_output_dir).grid(row=6, column=2)

        # 3. 命名模式选择 (布局更新)
        ttk.Label(main_frame, text="3. 选择命名模式:").grid(row=7, column=0, sticky="w", pady=(15, 5))
        mode_frame = ttk.Frame(main_frame)
        mode_frame.grid(row=8, column=0, columnspan=3, sticky="w")
        self.seq_radio = ttk.Radiobutton(mode_frame, text="顺序命名", variable=self.naming_mode, value="sequential", command=self.toggle_naming_mode)
        self.seq_radio.pack(side="left")
        self.manual_radio = ttk.Radiobutton(mode_frame, text="手动命名 (来自XLS文件)", variable=self.naming_mode, value="manual", command=self.toggle_naming_mode)
        self.manual_radio.pack(side="left", padx=20)

        # -- 顺序命名相关控件 --
        self.num_label = ttk.Label(main_frame, text="生成图片张数 (x):")
        self.num_label.grid(row=9, column=0, sticky="w", pady=5, padx=20)
        self.num_entry = ttk.Entry(main_frame, textvariable=self.num_to_generate, width=10)
        self.num_entry.grid(row=9, column=1, sticky="w")

        # -- 手动命名相关控件 --
        self.manual_naming_frame = ttk.Frame(main_frame)
        self.manual_naming_frame.grid(row=10, column=0, columnspan=3, sticky="w", padx=20)
        self.naming_label = ttk.Label(self.manual_naming_frame, text="上传命名序列 (.xls格式):", wraplength=400)
        self.naming_label.pack(side="left", fill="x", expand=True)

        manual_btn_frame = ttk.Frame(main_frame)
        manual_btn_frame.grid(row=10, column=2, sticky="e")
        self.naming_button = ttk.Button(manual_btn_frame, text="上传...", command=self.select_naming_file)
        self.naming_button.pack(side="left")
        self.clear_naming_button = ttk.Button(manual_btn_frame, text="清除", command=self.clear_naming_file)
        self.clear_naming_button.pack(side="left", padx=5)

        self.name_count_label = ttk.Label(main_frame, text="读取到的名称数量: 0")
        self.name_count_label.grid(row=11, column=0, columnspan=3, sticky="w", padx=20)

        # --- 分割线 ---
        ttk.Separator(main_frame, orient='horizontal').grid(row=12, column=0, columnspan=3, sticky='ew', pady=20)

        # 4. 执行按钮和进度条
        self.generate_button = ttk.Button(main_frame, text="开始生成", command=self.start_generation_thread)
        self.generate_button.grid(row=13, column=0, columnspan=3, pady=10)
        self.progress_bar = ttk.Progressbar(main_frame, orient='horizontal', length=500, mode='determinate')
        self.progress_bar.grid(row=14, column=0, columnspan=3, pady=5)
        self.status_label = ttk.Label(main_frame, text="等待操作...")
        self.status_label.grid(row=15, column=0, columnspan=3, sticky="w")
        
        self.toggle_naming_mode()

    def toggle_naming_mode(self):
        mode = self.naming_mode.get()
        is_seq = mode == "sequential"
        
        self.num_entry.config(state="normal" if is_seq else "disabled")
        self.num_label.config(state="normal" if is_seq else "disabled")
        
        self.naming_button.config(state="disabled" if is_seq else "normal")
        self.clear_naming_button.config(state="disabled" if is_seq else "normal")
        self.naming_label.config(state="normal" if not is_seq else "disabled")
        self.name_count_label.config(state="normal" if not is_seq else "disabled")

    def add_input_images(self):
        # (此函数与v3.0相同)
        paths = filedialog.askopenfilenames(title="请选择图片 (可多选)", filetypes=[("Image Files", "*.jpg *.jpeg *.png *.bmp")])
        if not paths: return
        current_count = len(self.input_image_paths)
        if current_count + len(paths) > 30:
            messagebox.showwarning("数量超限", f"最多只能选择30张图片，您已选择{current_count}张，本次尝试添加{len(paths)}张。")
            return
        for path in paths:
            if path not in self.input_image_paths: self.input_image_paths.append(path); self.image_listbox.insert(tk.END, os.path.basename(path))
        self.update_max_images()

    def clear_input_list(self):
        # (此函数与v3.0相同)
        self.input_image_paths.clear(); self.image_listbox.delete(0, tk.END); self.update_max_images()

    def update_max_images(self):
        # (此函数与v3.0相同)
        total_max = 0
        for path in self.input_image_paths:
            try:
                with Image.open(path) as img:
                    img = ImageOps.exif_transpose(img)
                    total_max += (img.width - int(img.width * 0.9) + 1) * (img.height - int(img.height * 0.9) + 1)
            except Exception:
                messagebox.showerror("错误", f"无法读取图片: {os.path.basename(path)}，已从列表中移除。")
                try: idx = self.input_image_paths.index(path); self.image_listbox.delete(idx); self.input_image_paths.pop(idx)
                except (ValueError, IndexError): pass
        self.max_images.set(total_max)

    def select_output_dir(self):
        # (此函数与v3.0相同)
        path = filedialog.askdirectory(title="请选择保存位置");
        if path: self.output_dir.set(path); self.output_label.config(text=path)

    def _read_xls_file(self, file_path):
        """辅助函数: 读取XLS文件，返回名称列表或None"""
        try:
            name_list = []
            workbook = xlrd.open_workbook(file_path)
            sheet = workbook.sheet_by_index(0)
            for row_idx in range(sheet.nrows):
                name = sheet.cell_value(row_idx, 0)
                if name: name_list.append(str(name))
            if not name_list:
                messagebox.showerror("读取错误", "XLS文件第一列为空或无效，无法获取任何名称。")
                return None
            return name_list
        except Exception as e:
            messagebox.showerror("文件错误", f"读取Excel文件失败: {e}")
            return None

    def select_naming_file(self):
        """核心更新: 选择文件后立即读取并反馈"""
        path = filedialog.askopenfilename(title="请选择XLS命名文件", filetypes=[("Excel 97-2003", "*.xls")])
        if not path:
            return
        
        name_list = self._read_xls_file(path)
        
        if name_list:
            count = len(name_list)
            self.naming_file_path.set(path)
            self.manual_name_count.set(count)
            self.naming_label.config(text=f"已选文件: {os.path.basename(path)}")
            self.name_count_label.config(text=f"读取到的名称数量: {count}")
            messagebox.showinfo("读取成功", f"成功从文件中读取到 {count} 个名称。")
        else:
            # 如果读取失败，清空状态
            self.clear_naming_file()

    def clear_naming_file(self):
        """新增: 清除已选的命名文件和数量"""
        self.naming_file_path.set("")
        self.manual_name_count.set(0)
        self.naming_label.config(text="请上传命名序列 (.xls格式):")
        self.name_count_label.config(text="读取到的名称数量: 0")

    def start_generation_thread(self):
        # (此函数与v3.0相同)
        thread = threading.Thread(target=self.generate_images); thread.daemon = True; thread.start()

    def generate_images(self):
        self.generate_button.config(state="disabled")
        self.status_label.config(text="开始校验输入数据...")

        if not self.input_image_paths or not self.output_dir.get():
            messagebox.showerror("错误", "请确保已添加源图片并选择了保存位置！")
            self.generate_button.config(state="normal")
            return
        
        mode = self.naming_mode.get()
        name_list = []
        x = 0

        if mode == "sequential":
            try:
                x = self.num_to_generate.get()
                if x <= 0: raise ValueError
            except (tk.TclError, ValueError):
                messagebox.showerror("错误", "顺序命名模式下，生成张数必须是一个大于0的整数！"); self.generate_button.config(state="normal"); return
            name_list = [f"{i+1}" for i in range(x)]
        else: # mode == "manual"
            if not self.naming_file_path.get():
                messagebox.showerror("错误", "手动命名模式下，请先上传一个有效的XLS命名文件。"); self.generate_button.config(state="normal"); return
            x = self.manual_name_count.get()
            # 在生成前最后确认一次文件内容
            name_list = self._read_xls_file(self.naming_file_path.get())
            if not name_list or len(name_list) != x:
                messagebox.showerror("文件校验失败", "XLS文件状态已改变或内容无效，请重新选择文件。"); self.clear_naming_file(); self.generate_button.config(state="normal"); return
            
        if x > self.max_images.get():
            messagebox.showwarning("警告", f"请求生成 {x} 张图片，但所有源图片合计最多只能生成 {self.max_images.get()} 张唯一的图片。"); self.generate_button.config(state="normal"); return
        
        # --- 生成过程 (与v3.0完全相同) ---
        try:
            self.status_label.config(text="正在准备图片数据...")
            shuffled_coords_per_image = {}
            for path in self.input_image_paths:
                with Image.open(path) as img:
                    img = ImageOps.exif_transpose(img)
                    w, h = img.size; cw, ch = int(w*0.9), int(h*0.9)
                    coords = [(l, t, l+cw, t+ch) for t in range(h-ch+1) for l in range(w-cw+1)]
                    random.shuffle(coords); shuffled_coords_per_image[path] = coords
            
            source_assignment_list = []; num_sources = len(self.input_image_paths)
            base_count, remainder = x // num_sources, x % num_sources
            for i in range(num_sources): source_assignment_list.extend([self.input_image_paths[i]] * (base_count + (1 if i < remainder else 0)))
            random.shuffle(source_assignment_list)

            self.progress_bar['maximum'] = x
            coord_counters = {path: 0 for path in self.input_image_paths}

            for i in range(x):
                source_path = source_assignment_list[i]
                coord_index = coord_counters[source_path]
                box = shuffled_coords_per_image[source_path][coord_index]
                coord_counters[source_path] += 1
                
                with Image.open(source_path) as img:
                    img = ImageOps.exif_transpose(img); cropped_image = img.crop(box)

                clean_name = "".join(c for c in name_list[i] if c.isalnum() or c in (' ', '.', '_')).rstrip()
                output_path = os.path.join(self.output_dir.get(), f"{clean_name}.jpg")
                cropped_image.convert('RGB').save(output_path, 'jpeg', quality=95)

                self.progress_bar['value'] = i + 1
                self.status_label.config(text=f"生成: {os.path.basename(output_path)} (源: {os.path.basename(source_path)}) ({i+1}/{x})")
                self.master.update_idletasks()

            messagebox.showinfo("完成", f"成功生成 {x} 张图片！")
        except Exception as e:
            messagebox.showerror("生成失败", f"在生成过程中发生错误: {e}")
        finally:
            self.progress_bar['value'] = 0; self.status_label.config(text="操作完成。"); self.generate_button.config(state="normal")


if __name__ == "__main__":
    root = tk.Tk()
    app = MultiImageGeneratorApp(root)
    root.mainloop()