"""图片转 WebP 模块"""

from .webp_handler import convert_images_to_webp_in_pptx, ProcessStatus

# 模块元数据 - 用于反射器自动发现
MODULE_METADATA = {
    'name': 'image_to_webp',
    'description': '将PPTX中全部图片转为WebP并等比压缩，减小体积与卡顿',
    'functions': [
        {
            'name': 'convert_images_to_webp_in_pptx',
            'description': '转换全部图片为WebP',
            'params': {
                'input_pptx': {'type': 'str', 'required': True, 'description': '输入PPTX文件路径'},
                'output_pptx': {'type': 'str', 'required': True, 'description': '输出PPTX文件路径'},
                'max_pixel': {'type': 'int', 'required': False, 'default': 1920, 'description': '长边最大像素尺寸'},
                'quality': {'type': 'int', 'required': False, 'default': 80, 'description': 'WebP有损压缩质量(1-100)'},
                'max_workers': {'type': 'int', 'required': False, 'default': 2, 'description': '线程数量'}
            }
        }
    ]
}

__all__ = ['convert_images_to_webp_in_pptx', 'ProcessStatus', 'MODULE_METADATA']
