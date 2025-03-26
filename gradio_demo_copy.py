import io
import openpyxl
import numpy as np
import gradio as gr
from table_extractor import rec_table  # 假设你的table_extractor.py文件在同一目录下
from PIL import Image

import concurrent.futures

# 创建一个线程池，限制并发数量
executor = concurrent.futures.ThreadPoolExecutor(max_workers=5)

def process_image(file_paths):
    try:
        # 处理多个文件
        excel_contents = []
        for file_path in file_paths:
            excel_content = rec_table(file_path.name, output="./", pre=False, logic_info=False, box=False)
            excel_contents.append(excel_content)
        return excel_contents
    except Exception as e:
        return f"处理文件时出错: {str(e)}"


with gr.Blocks() as demo:
    gr.Markdown("""<div style="text-align: center;"><h1>表格识别与导出工具</h1></div>""")
    gr.Markdown("#### 上传包含表格的文件（可上传格式：pdf、jpg、jpeg、png、tiff、 tif、bmp），系统将自动识别表格并导出为Excel文件。")
    with gr.Row():
        file_input = gr.Files(label="上传多个文件", file_count="multiple")
        file_output = gr.Files(label="处理后的文件")
    submit_btn = gr.Button("开始处理")
    submit_btn.click(process_image, inputs=file_input, outputs=file_output)

    # 启用队列功能
    demo.queue()

# 启动 Gradio 应用
demo.launch(server_name="0.0.0.0")
