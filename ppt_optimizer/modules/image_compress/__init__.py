"""图片压缩模块"""

from .compress_handler import compress_images_in_pptx, ProcessStatus

# 模块元数据 - 用于反射器自动发现
MODULE_METADATA = {
    'name': 'image_compress',
    'description': '压缩PPTX中的所有图片以减小文件体积',
    'functions': [
        {
            'name': 'compress_images_in_pptx',
            'description': '压缩PPTX中的位图图片',
            'params': {
                'input_pptx': {'type': 'str', 'required': True, 'description': '输入PPTX文件路径'},
                'output_pptx': {'type': 'str', 'required': True, 'description': '输出PPTX文件路径'},
                'max_pixel': {'type': 'int', 'required': False, 'default': 1920, 'description': '最大像素尺寸'},
                'quality': {'type': 'int', 'required': False, 'default': 85, 'description': 'JPEG压缩质量(1-100)'},
                'max_workers': {'type': 'int', 'required': False, 'default': 4, 'description': '线程数量'}
            }
        }
    ]
}

__all__ = ['compress_images_in_pptx', 'ProcessStatus', 'MODULE_METADATA']
