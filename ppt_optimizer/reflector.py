"""反射器 - 自动扫描和调用模块"""

import os
import json
import importlib
from pathlib import Path
from typing import Dict, List, Any, Tuple


class ModuleReflector:
    """模块反射器 - 自动发现和调用处理方法"""
    
    def __init__(self, modules_dir: str = None):
        """
        初始化反射器
        
        Args:
            modules_dir: 模块目录路径，默认为当前文件的modules目录
        """
        if modules_dir is None:
            modules_dir = Path(__file__).parent / 'modules'
        
        self.modules_dir = Path(modules_dir)
        self._modules_cache = {}
        self._scan_modules()
    
    def _scan_modules(self):
        """扫描并加载所有可用模块"""
        self._modules_cache.clear()
        
        if not self.modules_dir.exists():
            return
        
        # 遍历modules目录下的所有子文件夹
        for item in self.modules_dir.iterdir():
            if item.is_dir() and not item.name.startswith('_'):
                init_file = item / '__init__.py'
                if init_file.exists():
                    try:
                        # 动态导入模块
                        module = importlib.import_module(f'.{item.name}', package='ppt_optimizer.modules')
                        
                        # 检查是否有MODULE_METADATA
                        if hasattr(module, 'MODULE_METADATA'):
                            metadata = module.MODULE_METADATA
                            module_name = metadata.get('name', item.name)
                            
                            # 缓存模块信息和函数引用
                            self._modules_cache[module_name] = {
                                'metadata': metadata,
                                'module': module,
                                'path': str(item)
                            }
                    except Exception as e:
                        # 记录错误但继续加载其他模块
                        print(f"警告: 加载模块 {item.name} 失败: {str(e)}")
                        continue
    
    def reload_modules(self):
        """重新扫描模块（用于热加载新模块）"""
        self._scan_modules()
    
    def get_available_methods(self) -> List[Dict[str, Any]]:
        """
        获取所有可用的处理方法及其参数信息
        
        Returns:
            list: 包含所有模块和方法信息的列表
        """
        methods = []
        
        for module_name, module_info in self._modules_cache.items():
            metadata = module_info['metadata']
            
            for func_info in metadata.get('functions', []):
                method_detail = {
                    'module': module_name,
                    'function': func_info['name'],
                    'description': func_info.get('description', ''),
                    'params': func_info.get('params', {}),
                    'module_description': metadata.get('description', '')
                }
                methods.append(method_detail)
        
        return methods
    
    def call_method(self, module_name: str, params_json: str) -> Tuple[bool, Any]:
        """
        统一调用方法
        
        Args:
            module_name: 模块名称
            params_json: JSON字符串，包含函数名和参数
            
        Returns:
            tuple: (成功标志, 结果或错误信息)
                   成功时返回 (True, 函数返回值)
                   失败时返回 (False, 错误信息)
        """
        try:
            # 解析JSON参数
            try:
                params = json.loads(params_json)
            except json.JSONDecodeError as e:
                return False, f"参数JSON格式错误: {str(e)}"
            
            # 检查模块是否存在
            if module_name not in self._modules_cache:
                available = list(self._modules_cache.keys())
                return False, f"模块 '{module_name}' 不存在。可用模块: {available}"
            
            module_info = self._modules_cache[module_name]
            module = module_info['module']
            
            # 获取函数名
            func_name = params.pop('function', None)
            if not func_name:
                return False, "参数中缺少 'function' 字段"
            
            # 检查函数是否存在
            if not hasattr(module, func_name):
                return False, f"模块 '{module_name}' 中不存在函数 '{func_name}'"
            
            func = getattr(module, func_name)
            
            # 调用函数
            result = func(**params)
            
            return True, result
            
        except TypeError as e:
            return False, f"参数错误: {str(e)}"
        except Exception as e:
            return False, f"调用失败: {str(e)}"
    
    def get_module_info(self, module_name: str) -> Dict[str, Any]:
        """
        获取指定模块的详细信息
        
        Args:
            module_name: 模块名称
            
        Returns:
            dict: 模块信息，如果模块不存在返回None
        """
        if module_name not in self._modules_cache:
            return None
        
        return self._modules_cache[module_name]['metadata']
    
    def list_modules(self) -> List[str]:
        """
        列出所有已加载的模块名称
        
        Returns:
            list: 模块名称列表
        """
        return list(self._modules_cache.keys())


# 创建全局反射器实例
_reflector = None


def get_reflector(modules_dir: str = None) -> ModuleReflector:
    """
    获取反射器实例（单例模式）
    
    Args:
        modules_dir: 模块目录路径
        
    Returns:
        ModuleReflector: 反射器实例
    """
    global _reflector
    if _reflector is None or modules_dir is not None:
        _reflector = ModuleReflector(modules_dir)
    return _reflector


def query_methods() -> List[Dict[str, Any]]:
    """
    查询所有可用方法（便捷函数）
    
    Returns:
        list: 方法信息列表
    """
    return get_reflector().get_available_methods()


def call_module(module_name: str, params_json: str) -> Tuple[bool, Any]:
    """
    调用模块方法（便捷函数）
    
    Args:
        module_name: 模块名称
        params_json: JSON参数字符串
        
    Returns:
        tuple: (成功标志, 结果或错误信息)
    """
    return get_reflector().call_method(module_name, params_json)
