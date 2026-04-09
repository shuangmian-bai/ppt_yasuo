"""
PPT矢量图批量转换工具

功能：自动扫描input文件夹下的所有PPT文件，将其中的矢量图（EMF/WMF/SVG）转换为PNG格式，
      输出到outdata文件夹。如果input文件夹不存在则自动创建。
      支持多线程并行处理多个PPT文件，每个文件独立显示进度条。

使用方法：
    1. 将需要处理的PPT文件放入input文件夹
    2. 运行 python main2.py
    3. 处理后的文件保存在outdata文件夹
"""
import os
import json
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor, as_completed
from ppt_optimizer import call_module

def process_single_ppt(ppt_file, output_dir, file_idx, total_files):
    """处理单个PPT文件"""
    print(f"\n[{file_idx}/{total_files}] 开始处理: {ppt_file.name}")
    
    # 构造调用参数
    call_params = {
        "function": "convert_emf_to_png_in_pptx",
        "input_pptx": str(ppt_file),
        "output_pptx": str(output_dir / ppt_file.name),
        "max_workers": 2  # 单文件内使用较少线程，避免资源竞争
    }
    
    # 转换为JSON字符串
    params_json = json.dumps(call_params)
    
    # 调用模块
    success, result = call_module("vector_image", params_json)
    
    if success:
        status, data = result
        if status.name == "SUCCESS":
            print(f"[{file_idx}/{total_files}] ✓ {ppt_file.name} 完成，替换矢量图数量: {data}")
            return True, ppt_file.name, data
        else:
            print(f"[{file_idx}/{total_files}] ✗ {ppt_file.name} 错误: {data}")
            return False, ppt_file.name, data
    else:
        print(f"[{file_idx}/{total_files}] ✗ {ppt_file.name} 调用失败: {result}")
        return False, ppt_file.name, result


def process_all_ppt_files(max_concurrent=3):
    """
    多线程处理input文件夹下所有PPT文件
    
    Args:
        max_concurrent: 同时处理的文件数量（默认3个）
    """
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
    
    print(f"找到 {len(ppt_files)} 个PPT文件，使用 {max_concurrent} 个线程并行处理...\n")
    
    # 统计结果
    success_count = 0
    fail_count = 0
    total_replaced = 0
    
    # 使用线程池并行处理
    with ThreadPoolExecutor(max_workers=max_concurrent) as executor:
        # 提交所有任务
        future_to_file = {
            executor.submit(process_single_ppt, ppt_file, output_dir, idx, len(ppt_files)): ppt_file
            for idx, ppt_file in enumerate(ppt_files, 1)
        }
        
        # 等待所有任务完成
        for future in as_completed(future_to_file):
            success, filename, data = future.result()
            if success:
                success_count += 1
                if isinstance(data, int):
                    total_replaced += data
            else:
                fail_count += 1
    
    # 打印总结
    print(f"\n{'='*50}")
    print(f"处理完成！")
    print(f"成功: {success_count} 个文件")
    print(f"失败: {fail_count} 个文件")
    print(f"总共替换矢量图: {total_replaced} 个")
    print(f"{'='*50}")

if __name__ == "__main__":
    process_all_ppt_files()
