import os
import io
from enum import Enum
from pptx import Presentation
from PIL import Image
from concurrent.futures import ThreadPoolExecutor, as_completed
from tqdm import tqdm


# 处理结果状态枚举
class ProcessStatus(Enum):
    SUCCESS = "正常处理完成"
    ERROR_INTERRUPTED = "异常中断"
    FILE_NOT_FOUND = "文件不存在"


def process_slide(slide, slide_idx, max_pixel=1920, quality=90):
    """处理单个幻灯片中的矢量图"""
    replaced = 0
    shapes = list(slide.shapes)
    
    for shape in reversed(shapes):
        if shape.shape_type == 13:  # 图片
            try:
                img = shape.image
                ext = img.ext.lower()
            except Exception:
                # 无法读取的图片格式（如 webp/mpo 等 python-pptx 不支持的格式）直接跳过
                continue

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
    """
    将PPTX中的矢量图转换为PNG
    
    Args:
        input_pptx: 输入PPTX文件路径
        output_pptx: 输出PPTX文件路径
        max_pixel: 最大像素尺寸
        quality: PNG压缩质量
        max_workers: 线程数量
    
    Returns:
        tuple: (ProcessStatus状态, 替换数量或错误信息)
    """
    # 检查输入文件是否存在
    if not os.path.exists(input_pptx):
        return ProcessStatus.FILE_NOT_FOUND, f"输入文件不存在: {input_pptx}"
    
    # 递归创建输出目录
    output_dir = os.path.dirname(output_pptx)
    if output_dir:
        try:
            os.makedirs(output_dir, exist_ok=True)
        except Exception as e:
            return ProcessStatus.ERROR_INTERRUPTED, f"创建输出目录失败: {str(e)}"
    
    try:
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
        prs.save(output_pptx)
        
        return ProcessStatus.SUCCESS, total_replaced
        
    except Exception as e:
        return ProcessStatus.ERROR_INTERRUPTED, f"处理过程中发生错误: {str(e)}"
