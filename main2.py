"""
PPT矢量图批量转换工具

功能：自动扫描input文件夹下的所有PPT文件，将其中的矢量图（EMF/WMF/SVG）转换为PNG格式，
      输出到outdata文件夹。如果input文件夹不存在则自动创建。

使用方法：
    1. 将需要处理的PPT文件放入input文件夹
    2. 运行 python main2.py
    3. 处理后的文件保存在outdata文件夹
"""
import os
import json
from pathlib import Path
from ppt_optimizer import call_module

def process_all_ppt_files():
    """处理input文件夹下所有PPT文件"""
    # 定义输入输出目录
    input_dir = Path("./input")
    output_dir = Path("./outdata")
    
    # 创建目录（如果不存在）
    input_dir.mkdir(exist_ok=True)
    output_dir.mkdir(exist_ok=True)
    
    # 获取所有PPT文件
    ppt_files = list(input_dir.glob("*.pptx")) + list(input_dir.glob("*.ppt"))
    
    if not ppt_files:
        print(f"警告: {input_dir} 目录下没有找到PPT文件")
        return
    
    print(f"找到 {len(ppt_files)} 个PPT文件，开始处理...\n")
    
    # 处理每个文件
    for idx, ppt_file in enumerate(ppt_files, 1):
        print(f"[{idx}/{len(ppt_files)}] 处理: {ppt_file.name}")
        
        # 构造调用参数
        call_params = {
            "function": "convert_emf_to_png_in_pptx",
            "input_pptx": str(ppt_file),
            "output_pptx": str(output_dir / ppt_file.name),
            "max_workers": 4
        }
        
        # 转换为JSON字符串
        params_json = json.dumps(call_params)
        
        # 调用模块
        success, result = call_module("vector_image", params_json)
        
        if success:
            status, data = result
            if status.name == "SUCCESS":
                print(f"  ✓ 完成，替换矢量图数量: {data}\n")
            else:
                print(f"  ✗ 错误: {data}\n")
        else:
            print(f"  ✗ 调用失败: {result}\n")
    
    print("\n所有文件处理完成！")

if __name__ == "__main__":
    process_all_ppt_files()
