import io
import openpyxl
import numpy as np
import gradio as gr
from table_extractor import rec_table  # 假设你的table_extractor.py文件在同一目录下
from PIL import Image

import concurrent.futures

# 创建一个线程池，限制并发数量
executor = concurrent.futures.ThreadPoolExecutor(max_workers=5)

def process_image(file_path):
    try:

        # 假设rec_table函数可以直接处理图片对象
        excel_content = rec_table(file_path, output="./", pre=False, logic_info=False, box=False)

        return excel_content
    except Exception as e:
        return f"处理文件时出错: {str(e)}"

iface = gr.Interface(
    fn=process_image, 
    inputs="file", 
    outputs="file", 
    title="表格识别与导出工具",
    description="上传包含表格的文件（可上传格式：pdf、jpg、jpeg、png、tiff、 tif、bmp），系统将自动识别表格并导出为Excel文件。",
)
iface.launch(server_name="0.0.0.0")
