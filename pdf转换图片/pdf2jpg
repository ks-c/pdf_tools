# -*- coding: utf-8 -*-
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import fitz  # PyMuPDF
import os
import threading
from PIL import Image

class PdfConverterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF 转换图片工具")
        self.root.geometry("550x500") # 窗口加长

        # --- 1. 样式：全部使用宋体 ---
        style = ttk.Style()
        # 注意: 请确保您的操作系统中安装了“宋体”字体
        font_family = "SimSun"

        
        print("当前字体为：",font_family)

        style.configure(".", font=(font_family, 10)) # 为所有ttk控件设置默认字体
        style.configure("TLabel", padding=5)
        style.configure("TButton", padding=5)
        style.configure("TEntry", padding=5)
        style.configure("TCombobox", padding=5)
        style.configure("TRadiobutton", padding=5)
        style.configure("TLabelframe.Label", font=(font_family, 11, 'bold')) # 组标题加粗


        # --- 主框架 ---
        main_frame = ttk.Frame(root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # --- 文件选择 ---
        file_frame = ttk.LabelFrame(main_frame, text="选择文件和路径", padding="10")
        file_frame.pack(fill=tk.X, pady=5)

        self.pdf_path_label = ttk.Label(file_frame, text="PDF 文件:")
        self.pdf_path_label.grid(row=0, column=0, sticky=tk.W)
        self.pdf_path_var = tk.StringVar()
        self.pdf_path_entry = ttk.Entry(file_frame, textvariable=self.pdf_path_var, state="readonly", width=50)
        self.pdf_path_entry.grid(row=0, column=1, padx=5, sticky=tk.EW)
        self.browse_pdf_btn = ttk.Button(file_frame, text="浏览...", command=self.select_pdf_file)
        self.browse_pdf_btn.grid(row=0, column=2)

        self.output_folder_label = ttk.Label(file_frame, text="输出文件夹:")
        self.output_folder_label.grid(row=1, column=0, sticky=tk.W, pady=5)
        self.output_folder_var = tk.StringVar()
        self.output_folder_entry = ttk.Entry(file_frame, textvariable=self.output_folder_var, state="readonly", width=50)
        self.output_folder_entry.grid(row=1, column=1, padx=5, sticky=tk.EW)
        self.browse_output_btn = ttk.Button(file_frame, text="浏览...", command=self.select_output_folder)
        self.browse_output_btn.grid(row=1, column=2)

        file_frame.grid_columnconfigure(1, weight=1)

        # --- 3. 执行按钮提前 ---
        execute_frame = ttk.Frame(main_frame)
        execute_frame.pack(fill=tk.X, pady=10)
        self.convert_btn = ttk.Button(execute_frame, text="开始转换", command=self.start_conversion_thread, style="Accent.TButton")
        style.configure("Accent.TButton", font=(font_family, 12, 'bold')) # 给主按钮一个突出的样式
        self.convert_btn.pack(ipady=8, fill=tk.X)


        # --- 转换选项 ---
        options_frame = ttk.LabelFrame(main_frame, text="调整转换选项", padding="10")
        options_frame.pack(fill=tk.X, pady=5)

        # 页码
        self.pages_label = ttk.Label(options_frame, text="转换页码:")
        self.pages_label.grid(row=0, column=0, sticky=tk.W)
        self.pages_var = tk.StringVar(value="全部")
        self.pages_entry = ttk.Entry(options_frame, textvariable=self.pages_var)
        self.pages_entry.grid(row=0, column=1, padx=5, sticky=tk.EW)
        self.pages_help_label = ttk.Label(options_frame, text="(例: 1,3,5-8)")
        self.pages_help_label.grid(row=0, column=2, sticky=tk.W)
        
        # 图片格式
        self.format_label = ttk.Label(options_frame, text="图片格式:")
        self.format_label.grid(row=1, column=0, sticky=tk.W, pady=5)
        self.format_var = tk.StringVar(value="jpg")
        self.format_combo = ttk.Combobox(options_frame, textvariable=self.format_var, values=["jpg", "png", "bmp", "tiff"], state="readonly")
        self.format_combo.grid(row=1, column=1, padx=5, sticky=tk.EW)

        # --- 2. DPI (清晰度) 改为下拉选择 ---
        self.dpi_label = ttk.Label(options_frame, text="分辨率 (DPI):")
        self.dpi_label.grid(row=2, column=0, sticky=tk.W, pady=5)
        self.dpi_var = tk.IntVar(value=200)
        dpi_values = [72, 100, 150, 200, 300, 600]
        self.dpi_combo = ttk.Combobox(options_frame, textvariable=self.dpi_var, values=dpi_values, state="readonly")
        self.dpi_combo.grid(row=2, column=1, padx=5, sticky=tk.EW)
        self.dpi_help_label = ttk.Label(options_frame, text="(越高越清晰)")
        self.dpi_help_label.grid(row=2, column=2, sticky=tk.W)

        # 输出模式
        self.output_mode_label = ttk.Label(options_frame, text="输出模式:")
        self.output_mode_label.grid(row=3, column=0, sticky=tk.W, pady=5)
        self.output_mode_var = tk.StringVar(value="per_page")
        self.radio_per_page = ttk.Radiobutton(options_frame, text="每页一张图", variable=self.output_mode_var, value="per_page")
        self.radio_per_page.grid(row=3, column=1, sticky=tk.W, padx=5)
        self.radio_single_image = ttk.Radiobutton(options_frame, text="合并为一张长图", variable=self.output_mode_var, value="single_image")
        self.radio_single_image.grid(row=3, column=2, sticky=tk.W)

        options_frame.grid_columnconfigure(1, weight=1)

        # --- 进度 ---
        status_frame = ttk.Frame(main_frame, padding="10")
        status_frame.pack(fill=tk.BOTH, expand=True)

        self.progress_bar = ttk.Progressbar(status_frame, orient="horizontal", mode="determinate")
        self.progress_bar.pack(fill=tk.X, pady=5, side=tk.BOTTOM)

        self.status_var = tk.StringVar(value="准备就绪")
        self.status_label = ttk.Label(status_frame, textvariable=self.status_var, anchor=tk.W)
        self.status_label.pack(fill=tk.X, side=tk.BOTTOM)

    def select_pdf_file(self):
        file_path = filedialog.askopenfilename(
            title="选择 PDF 文件",
            filetypes=[("PDF Files", "*.pdf")]
        )
        if file_path:
            self.pdf_path_var.set(file_path)
            if not self.output_folder_var.get():
                output_dir = os.path.dirname(file_path)
                self.output_folder_var.set(output_dir)

    def select_output_folder(self):
        folder_path = filedialog.askdirectory(title="选择输出文件夹")
        if folder_path:
            self.output_folder_var.set(folder_path)
    
    def update_status(self, message):
        self.status_var.set(message)
    
    def update_progress(self, value):
        self.progress_bar['value'] = value

    def start_conversion_thread(self):
        pdf_path = self.pdf_path_var.get()
        output_folder = self.output_folder_var.get()

        if not pdf_path or not output_folder:
            messagebox.showerror("错误", "请先选择 PDF 文件和输出文件夹！")
            return

        self.convert_btn.config(state="disabled")
        self.status_var.set("正在准备转换...")
        self.progress_bar['value'] = 0

        conversion_thread = threading.Thread(target=self.convert_pdf, daemon=True)
        conversion_thread.start()

    def parse_page_numbers(self, page_str, total_pages):
        if page_str.strip().lower() in ["", "all", "全部"]:
            return list(range(total_pages))
        
        pages = set()
        try:
            parts = page_str.split(',')
            for part in parts:
                part = part.strip()
                if '-' in part:
                    start, end = map(int, part.split('-'))
                    if start < 1 or end > total_pages or start > end: raise ValueError
                    for i in range(start, end + 1): pages.add(i - 1)
                else:
                    page_num = int(part)
                    if page_num < 1 or page_num > total_pages: raise ValueError
                    pages.add(page_num - 1)
        except ValueError:
            return None
        
        return sorted(list(pages))

    def convert_pdf(self):
        try:
            pdf_path = self.pdf_path_var.get()
            output_folder = self.output_folder_var.get()
            page_input = self.pages_var.get()
            img_format = self.format_var.get()
            dpi = self.dpi_var.get()
            output_mode = self.output_mode_var.get()

            doc = fitz.open(pdf_path)
            
            pages_to_convert = self.parse_page_numbers(page_input, doc.page_count)
            if pages_to_convert is None:
                self.root.after(0, messagebox.showerror, "错误", f"页码范围 '{page_input}' 无效。请输入正确的页码，例如 '1,3,5-8'。该PDF总页数为 {doc.page_count}。")
                self.root.after(0, self.reset_ui)
                return

            if not pages_to_convert:
                self.root.after(0, messagebox.showinfo, "提示", "没有需要转换的页码。")
                self.root.after(0, self.reset_ui)
                return
            
            self.progress_bar['maximum'] = len(pages_to_convert)
            base_filename = os.path.splitext(os.path.basename(pdf_path))[0]
            zoom = dpi / 72
            mat = fitz.Matrix(zoom, zoom)
            images_for_merge = []

            for i, page_num in enumerate(pages_to_convert):
                self.root.after(0, self.update_status, f"正在转换第 {page_num + 1} 页...")
                page = doc.load_page(page_num)
                pix = page.get_pixmap(matrix=mat)
                
                if output_mode == "per_page":
                    output_path = os.path.join(output_folder, f"{base_filename}_page_{page_num + 1}.{img_format}")
                    pix.save(output_path)
                else:
                    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                    images_for_merge.append(img)
                self.root.after(0, self.update_progress, i + 1)
            
            if output_mode == "single_image" and images_for_merge:
                self.root.after(0, self.update_status, "正在合并图片...")
                widths, heights = zip(*(i.size for i in images_for_merge))
                total_height = sum(heights)
                max_width = max(widths)
                merged_image = Image.new('RGB', (max_width, total_height), color='white')
                y_offset = 0
                for img in images_for_merge:
                    merged_image.paste(img, (0, y_offset))
                    y_offset += img.height
                output_path = os.path.join(output_folder, f"{base_filename}_merged.{img_format}")
                merged_image.save(output_path)

            doc.close()
            self.root.after(0, self.update_status, "转换完成！")
            self.root.after(0, messagebox.showinfo, "成功", f"所有页面已成功转换并保存到:\n{output_folder}")

        except Exception as e:
            self.root.after(0, self.update_status, "发生错误！")
            self.root.after(0, messagebox.showerror, "错误", f"转换过程中发生错误: \n{str(e)}")
        finally:
            self.root.after(0, self.reset_ui)
    
    def reset_ui(self):
        self.convert_btn.config(state="normal")
        self.status_var.set("准备就绪")
        self.progress_bar['value'] = 0


if __name__ == "__main__":
    root = tk.Tk()
    app = PdfConverterApp(root)
    root.mainloop()
