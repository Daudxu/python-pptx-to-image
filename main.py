import os  
import subprocess  
from pdf2image import convert_from_path  
from PIL import Image  

def pptx_to_images(pptx_path, output_dir, libreoffice_path):  
    # 创建输出目录（如果不存在）  
    os.makedirs(output_dir, exist_ok=True)  

    # 将 pptx 转换为 pdf（使用 LibreOffice）  
    pdf_path = os.path.join(output_dir, os.path.splitext(os.path.basename(pptx_path))[0] + '.pdf')  
    libreoffice_cmd = [libreoffice_path, '--headless', '--convert-to', 'pdf', '--outdir', output_dir, pptx_path]  
    try:  
        subprocess.run(libreoffice_cmd, check=True)  
        # 检查 PDF 文件是否存在  
        if os.path.isfile(pdf_path):  
            print(f"转换成功，PDF 文件已保存到：{pdf_path}")  
            # 读取PDF文件并转换为图片列表  
            images = convert_from_path(pdf_path)  
            
            # 计算总高度  
            total_height = sum(img.size[1] for img in images)  
            # 假设所有图片宽度都相同（如果不是，您可能需要调整这个逻辑）  
            max_width = images[0].size[0]  
            
            # 创建一个新的空白图片，大小为所有图片宽度中的最大值，高度为所有图片高度之和  
            new_img = Image.new('RGB', (max_width, total_height))  
            
            # 粘贴每张图片到新的图片上，从上到下依次排列  
            y_offset = 0  
            for img in images:  
                new_img.paste(img, (0, y_offset))  
                y_offset += img.size[1]  # 增加偏移量以放置下一张图片  
            
            # 获取当前目录
            current_directory = os.getcwd()
            # 定义要保存合并后图片的路径
            merged_image_path = os.path.join(current_directory, 'image', os.path.splitext(os.path.basename(pdf_path))[0] + '.png')

            # 保存合并后的图片  
            new_img.save(merged_image_path, "PNG")  
        else:  
            print("转换似乎成功了，但 PDF 文件未找到。")  
    except subprocess.CalledProcessError as e:  
        print(f"转换失败：{e}")  
  
# 使用示例，确保将 libreoffice_path 替换为你的 LibreOffice soffice.exe 的实际路径  
libreoffice_path = 'C:\\Program Files\\LibreOffice\\program\\soffice.exe'  # 修改为你的 LibreOffice 路径  
pptx_to_images('test1.pptx', 'output_images', libreoffice_path)
