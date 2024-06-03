import os
from comtypes.client import CreateObject
from tkinter import Tk, filedialog, Button, Label, StringVar, Entry, messagebox
import tkinter as tk
from tkinter import messagebox
 
 
def browse_file():
    global ppt_path
    file_path = filedialog.askopenfilename(filetypes=[("PowerPoint Files", "*.pptx;*.ppt")])
    normalized_path = os.path.normpath(file_path)  # 规范化路径
    ppt_path.set(normalized_path)
   
def browse_folder():
    global images_dir
    folder_path = filedialog.askdirectory()
    normalized_path = os.path.normpath(folder_path)  # 规范化路径
    images_dir.set(normalized_path)
 
def convert_to_images():
    try:
        ppt_to_images(ppt_path.get(), images_dir.get())
        messagebox.showinfo("完成", "PPT已成功转换为图片")
    except Exception as e:
        messagebox.showerror("错误", str(e))
 
def ppt_to_images(ppt_path, images_dir):
    # 确保输出目录存在
    if not os.path.exists(images_dir):
        os.makedirs(images_dir)
 
    # 初始化PowerPoint应用
    powerpoint = CreateObject("Powerpoint.Application")
    powerpoint.Visible = 1
 
    # 打开PPT文件
    ppt = powerpoint.Presentations.Open(ppt_path)
 
    # 遍历每个幻灯片并保存为图片
    for i, slide in enumerate(ppt.Slides):
        image_path = os.path.join(images_dir, f"slide_{i + 1}.jpg")
        slide.Export(image_path, "JPG")
  
    # 关闭PPT文件和PowerPoint应用
    ppt.Close()
    powerpoint.Quit()
 
if __name__ == "__main__":
    root = Tk()
    root.title("PPT转图片工具")
  
    ppt_path = StringVar()
    images_dir = StringVar()
   
    Label(root, text="选择PPT文件:").pack()
    Entry(root, textvariable=ppt_path, width=50).pack()
    Button(root, text="浏览", command=browse_file).pack()
  
    Label(root, text="图片输出目录:").pack()
    Entry(root, textvariable=images_dir, width=50).pack()
    Button(root, text="浏览", command=browse_folder).pack()
 
    Button(root, text="转换", command=convert_to_images).pack()
  
    root.mainloop()
