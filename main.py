import os
import io
from pptx import Presentation
from PIL import Image

def convert_emf_to_png_in_pptx(input_pptx, output_pptx, max_pixel=1920, quality=90):
    prs = Presentation(input_pptx)
    replaced = 0

    for slide_idx, slide in enumerate(prs.slides, 1):
        print(f"正在处理第 {slide_idx} 页")

        shapes = list(slide.shapes)
        for shape in reversed(shapes):
            if shape.shape_type == 13:  # 图片
                img = shape.image
                ext = img.ext.lower()

                if ext in ["emf", "wmf", "svg"]:
                    print(f"  发现矢量图 {ext}，开始转换...")

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
                        print(f"  转换失败: {e}")
                        continue

    # 保存
    os.makedirs(os.path.dirname(output_pptx), exist_ok=True)
    prs.save(output_pptx)

    print("\n处理完成")
    print(f"输入文件: {input_pptx}")
    print(f"输出文件: {output_pptx}")
    print(f"替换矢量图数量: {replaced}")

if __name__ == "__main__":
    INPUT = "./data/电力电子技术冲刺版课件7.pptx"
    OUTPUT = "./outdata/优化后_矢量转位图.pptx"
    convert_emf_to_png_in_pptx(INPUT, OUTPUT)