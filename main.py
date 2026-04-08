import os
import io
from pptx import Presentation
from PIL import Image
from concurrent.futures import ThreadPoolExecutor, as_completed
from tqdm import tqdm
import threading

def process_slide(slide, slide_idx, max_pixel=1920, quality=90):
    """处理单个幻灯片中的矢量图"""
    replaced = 0
    shapes = list(slide.shapes)
    
    for shape in reversed(shapes):
        if shape.shape_type == 13:  # 图片
            img = shape.image
            ext = img.ext.lower()
            
            if ext in ["emf", "wmf", "svg"]:
                try:
                    # 提取图片字节
                    img_bytes = img.blob
                    img_pil = Image.open(io.BytesIO(img_bytes)).convert("RGB")
                    
                    # 等比压缩
                    w, h = img_pil.size
                    if max(w, h) > max_pixel:
                        scale = max_pixel / max(w, h)
                        new_w = int(w * scale)
                        new_h = int(h * scale)
                        img_pil = img_pil.resize((new_w, new_h), Image.Resampling.LANCZOS)
                    
                    # 保存为压缩后的 PNG
                    output_io = io.BytesIO()
                    img_pil.save(output_io, format="PNG", optimize=True, quality=quality)
                    output_io.seek(0)
                    
                    # 保留位置与大小
                    left = shape.left
                    top = shape.top
                    width = shape.width
                    height = shape.height
                    
                    # 删除原矢量图
                    slide.shapes._spTree.remove(shape._element)
                    
                    # 插入新 PNG
                    slide.shapes.add_picture(
                        output_io,
                        left, top, width, height
                    )
                    replaced += 1
                    
                except Exception as e:
                    continue
    
    return slide_idx, replaced


def convert_emf_to_png_in_pptx(input_pptx, output_pptx, max_pixel=1920, quality=90, max_workers=4):
    prs = Presentation(input_pptx)
    total_replaced = 0
    
    # 准备任务列表
    tasks = [(slide, idx) for idx, slide in enumerate(prs.slides, 1)]
    total_slides = len(tasks)
    
    # 使用线程池处理
    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        # 提交所有任务
        future_to_slide = {
            executor.submit(process_slide, slide, idx, max_pixel, quality): idx
            for slide, idx in tasks
        }
        
        # 使用进度条显示处理进度
        with tqdm(total=total_slides, desc="处理进度", unit="页") as pbar:
            for future in as_completed(future_to_slide):
                slide_idx, replaced = future.result()
                total_replaced += replaced
                pbar.set_postfix({"已替换": total_replaced})
                pbar.update(1)
    
    # 保存结果
    os.makedirs(os.path.dirname(output_pptx), exist_ok=True)
    prs.save(output_pptx)
    
    return total_replaced

if __name__ == "__main__":
    INPUT = "./data/电力电子技术冲刺版课件7.pptx"
    OUTPUT = "./outdata/电力电子技术冲刺版课件7.pptx"
    MAX_WORKERS = 4
    
    print(f"开始处理，使用 {MAX_WORKERS} 个线程...")
    replaced_count = convert_emf_to_png_in_pptx(INPUT, OUTPUT, max_workers=MAX_WORKERS)
    
    print("\n处理完成")
    print(f"输入文件: {INPUT}")
    print(f"输出文件: {OUTPUT}")
    print(f"替换矢量图数量: {replaced_count}")