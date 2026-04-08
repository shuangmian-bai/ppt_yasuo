from pprint import pprint

from ppt_optimizer import query_methods, call_module
import json

if __name__ == "__main__":
    # 1. 查询所有可用方法
    print("=== 可用的处理方法 ===")
    methods = query_methods()
    pprint(methods)

    # 2. 调用模块示例
    print("\n=== 调用模块处理PPT ===")
    
    # 构造调用参数
    call_params = {
        "function": "convert_emf_to_png_in_pptx",
        "input_pptx": "./data/电力电子技术冲刺版课件7.pptx",
        "output_pptx": "./outdata/电力电子技术冲刺版课件7.pptx",
        "max_workers": 4
    }
    
    # 转换为JSON字符串
    params_json = json.dumps(call_params)
    
    # 调用模块
    success, result = call_module("vector_image", params_json)
    
    if success:
        status, data = result
        print(f"\n处理状态: {status.value}")
        if status.name == "SUCCESS":
            print(f"替换矢量图数量: {data}")
        else:
            print(f"错误信息: {data}")
    else:
        print(f"\n调用失败: {result}")
