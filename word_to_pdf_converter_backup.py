# word_to_pdf_converter.py

# This script converts all Word files in a specified directory and its subdirectories to PDF format.
# It maintains the original folder structure and saves the PDF files in a specified output directory.
# It displays the conversion progress and any errors that may occur.

# Import necessary libraries
import os
import tkinter as tk
from tkinter import filedialog, messagebox
from pathlib import Path
import comtypes.client
import time

class WordToPdfConverter:
    def __init__(self):
        self.window = tk.Tk()
        self.window.title("Word转PDF转换器")
        self.window.geometry("600x400")
        
        # 设置窗口样式
        self.window.configure(bg='#f0f0f0')
        
        # 创建主框架
        self.main_frame = tk.Frame(self.window, bg='#f0f0f0', padx=20, pady=20)
        self.main_frame.pack(expand=True, fill='both')
        
        # 创建标题标签
        title_label = tk.Label(
            self.main_frame,
            text="Word转PDF转换器",
            font=("Arial", 16, "bold"),
            bg='#f0f0f0'
        )
        title_label.pack(pady=20)
        
        # 创建说明标签
        instruction_label = tk.Label(
            self.main_frame,
            text="请选择包含Word文件的文件夹",
            font=("Arial", 10),
            bg='#f0f0f0'
        )
        instruction_label.pack(pady=10)
        
        # 创建选择文件夹按钮
        self.select_button = tk.Button(
            self.main_frame,
            text="选择文件夹",
            command=self.select_folder,
            font=("Arial", 10),
            bg='#4CAF50',
            fg='white',
            padx=20,
            pady=10
        )
        self.select_button.pack(pady=20)
        
        # 创建状态标签
        self.status_label = tk.Label(
            self.main_frame,
            text="",
            font=("Arial", 10),
            bg='#f0f0f0',
            wraplength=500
        )
        self.status_label.pack(pady=10)
        
        # 创建进度标签
        self.progress_label = tk.Label(
            self.main_frame,
            text="",
            font=("Arial", 10),
            bg='#f0f0f0'
        )
        self.progress_label.pack(pady=10)

    def select_folder(self):
        folder_path = filedialog.askdirectory(title="选择包含Word文件的文件夹")
        if folder_path:
            self.convert_files(folder_path)

    def convert_docx_to_pdf(self, word_path, pdf_path):
        try:
            # 获取绝对路径
            word_path = os.path.abspath(word_path)
            pdf_path = os.path.abspath(pdf_path)
            
            # 创建Word应用实例
            word = comtypes.client.CreateObject('Word.Application')
            word.Visible = False
            
            try:
                # 打开Word文档
                doc = word.Documents.Open(word_path)
                
                # 设置PDF输出选项
                wdFormatPDF = 17
                
                # 转换为PDF
                doc.SaveAs(pdf_path, FileFormat=wdFormatPDF)
                
                # 关闭文档
                doc.Close()
                return True
                
            except Exception as e:
                print(f"转换失败之后呢? {word_path}: {str(e)}")
                return False
                
            finally:
                # 退出Word应用
                word.Quit()
                
        except Exception as e:
            print(f"转换失败 {word_path}: {str(e)}")
            return False

    def convert_files(self, input_folder):
        try:
            # 创建输出文件夹
            input_folder_path = Path(input_folder)
            output_folder = str(input_folder_path.parent / f"{input_folder_path.name}_pdf")
            
            if not os.path.exists(output_folder):
                os.makedirs(output_folder)
            
            # 获取所有Word文件
            word_files = []
            for root, dirs, files in os.walk(input_folder):
                for file in files:
                    if file.endswith(('.docx', '.doc')):
                        word_files.append(os.path.join(root, file))
            
            if not word_files:
                messagebox.showinfo("提示", "所选文件夹中没有找到Word文件！")
                return
            
            # 更新状态
            self.status_label.config(text=f"开始转换文件...\n共找到 {len(word_files)} 个Word文件")
            self.window.update()
            
            # 转换文件
            success_count = 0
            for i, word_file in enumerate(word_files, 1):
                try:
                    # 获取相对路径
                    rel_path = os.path.relpath(word_file, input_folder)
                    # 创建对应的输出文件夹
                    output_subfolder = os.path.join(output_folder, os.path.dirname(rel_path))
                    if not os.path.exists(output_subfolder):
                        os.makedirs(output_subfolder)
                    
                    # 生成输出PDF文件路径
                    pdf_file = os.path.join(output_subfolder, 
                                          os.path.splitext(os.path.basename(word_file))[0] + '.pdf')
                    
                    # 转换文件
                    if self.convert_docx_to_pdf(word_file, pdf_file):
                        success_count += 1
                        # 等待一小段时间确保文件写入完成
                        time.sleep(1)
                    
                    # 更新进度
                    self.progress_label.config(
                        text=f"正在处理: {i}/{len(word_files)}\n"
                             f"当前文件: {os.path.basename(word_file)}"
                    )
                    self.window.update()
                    
                except Exception as e:
                    print(f"转换失败 {word_file}: {str(e)}")
            
            # 显示完成消息
            messagebox.showinfo(
                "完成",
                f"转换完成！\n"
                f"成功转换: {success_count} 个文件\n"
                f"失败: {len(word_files) - success_count} 个文件\n"
                f"PDF文件保存在: {output_folder}"
            )
            
            # 重置状态
            self.status_label.config(text="")
            self.progress_label.config(text="")
            
        except Exception as e:
            messagebox.showerror("错误", f"发生错误：{str(e)}")

    def run(self):
        self.window.mainloop()

if __name__ == "__main__":
    app = WordToPdfConverter()
    app.run()