# PPT优化器 - 反射器使用指南

## 核心功能

### 1. 查询所有可用方法
```python
from ppt_optimizer import query_methods

methods = query_methods()
for method in methods:
    print(f"模块: {method['module']}")
    print(f"函数: {method['function']}")
    print(f"参数: {method['params']}")
```

### 2. 调用模块方法
```python
from ppt_optimizer import call_module
import json

# 构造参数
params = {
    "function": "convert_emf_to_png_in_pptx",
    "input_pptx": "./input.pptx",
    "output_pptx": "./output.pptx",
    "max_workers": 4
}

# 调用（模块名 + JSON参数）
success, result = call_module("vector_image", json.dumps(params))

if success:
    status, data = result
    print(f"状态: {status.value}, 数据: {data}")
else:
    print(f"错误: {result}")
```

## 扩展新模块（3步完成）

### 步骤1: 创建模块文件夹
在 `ppt_optimizer/modules/` 下创建新文件夹，如 `my_module/`

### 步骤2: 实现处理函数
创建 `my_module/handler.py`：
```python
def my_function(input_pptx, output_pptx, param1=100):
    # 你的处理逻辑
    return ProcessStatus.SUCCESS, result_count
```

### 步骤3: 注册模块元数据
创建 `my_module/__init__.py`：
```python
from .handler import my_function

MODULE_METADATA = {
    'name': 'my_module',
    'description': '模块描述',
    'functions': [
        {
            'name': 'my_function',
            'description': '函数描述',
            'params': {
                'input_pptx': {'type': 'str', 'required': True, 'description': '输入路径'},
                'output_pptx': {'type': 'str', 'required': True, 'description': '输出路径'},
                'param1': {'type': 'int', 'required': False, 'default': 100, 'description': '参数说明'}
            }
        }
    ]
}

__all__ = ['my_function', 'MODULE_METADATA']
```

**完成！** 反射器会自动发现新模块，无需修改任何现有代码。

## 异常处理

反射器已内置完善的异常处理：
- JSON格式错误
- 模块不存在
- 函数不存在
- 参数类型错误
- 运行时异常

所有错误都会返回 `(False, 错误信息)` 格式。
