import os
import re
import json
import openpyxl
import PyPDF2
from datetime import datetime
from pathlib import Path
from openai import OpenAI
from typing import List, Dict, Any, Optional, Tuple

'''
定义了项目的配置类 Config ，用于管理项目的各种配置信息，
包括阿里云 API 的密钥、地址、使用的模型，以及项目的输入输出目录路径等。
同时，它还负责创建必要的项目目录。
'''

class Config:
    def __init__(self):
        self.BASE_DIR = Path(r"D:\codes\​AutoTestCase_WrittenByAI")  # 项目根目录
        self.API_KEY = "sk-ad87abf37f884cd08455d87fe2dabbf8"  # 阿里云API Key
        self.BASE_URL = "https://dashscope.aliyuncs.com/compatible-mode/v1"  # 阿里云API地址
        self.MODEL = "qwen-plus"  # 使用的模型
        
        # 路径配置
        self.INPUTS_DIR = self.BASE_DIR / "inputs"
        self.OUTPUTS_DIR = self.BASE_DIR / "outputs"
        self.CONFIG_PATH = self.BASE_DIR / "config"
        
        # 创建必要目录
        for dir_path in [self.OUTPUTS_DIR, self.INPUTS_DIR, self.CONFIG_PATH]:
            dir_path.mkdir(exist_ok=True)