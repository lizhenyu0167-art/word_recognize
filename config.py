#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
项目配置文件
统一管理文件路径、默认设置等配置项
"""

import os

class Config:
    """项目配置类"""
    
    # 文件路径配置
    TEMPLATE_FILE = "格式模板.docx"
    TEST_DOCUMENT = "测试文档.docx"
    OUTPUT_DIR = "output"
    
    # 格式信息文件
    DYNAMIC_FORMAT_INFO = os.path.join(OUTPUT_DIR, "dynamic_format_info.json")
    VALIDATION_REPORT = os.path.join(OUTPUT_DIR, "format_validation_report.json")
    ENHANCED_ANALYSIS = os.path.join(OUTPUT_DIR, "enhanced_format_analysis.json")
    ARCHITECTURE_TEST_REPORT = os.path.join(OUTPUT_DIR, "architecture_test_report.json")
    
    # 默认设置
    DEFAULT_FONT = "宋体"
    DEFAULT_FONT_SIZE = "10.5pt"
    
    # 文档生成设置
    FORMATTED_DOC_PREFIX = "格式化后的测试文档_"
    
    @classmethod
    def ensure_output_dir(cls):
        """确保输出目录存在"""
        if not os.path.exists(cls.OUTPUT_DIR):
            os.makedirs(cls.OUTPUT_DIR)
            print(f"创建输出目录: {cls.OUTPUT_DIR}")
    
    @classmethod
    def get_formatted_doc_path(cls, timestamp=None):
        """获取格式化文档路径"""
        # 使用带时间戳的文件名，避免文件被占用
        from datetime import datetime
        if timestamp is None:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        return os.path.join(cls.OUTPUT_DIR, f"格式化后的测试文档_{timestamp}.docx")
        
    @classmethod
    def get_fixed_formatted_doc_path(cls):
        """获取固定名称的格式化文档路径"""
        return os.path.join(cls.OUTPUT_DIR, "格式化后的测试文档.docx")
    
    @classmethod
    def validate_required_files(cls):
        """验证必需文件是否存在"""
        missing_files = []
        
        if not os.path.exists(cls.TEMPLATE_FILE):
            missing_files.append(cls.TEMPLATE_FILE)
        
        if not os.path.exists(cls.TEST_DOCUMENT):
            missing_files.append(cls.TEST_DOCUMENT)
        
        return missing_files
    
    @classmethod
    def get_latest_formatted_doc(cls):
        """获取最新的格式化文档路径"""
        import glob
        
        # 在当前目录查找
        pattern1 = f"{cls.FORMATTED_DOC_PREFIX}*.docx"
        files1 = glob.glob(pattern1)
        
        # 在输出目录查找
        pattern2 = os.path.join(cls.OUTPUT_DIR, f"{cls.FORMATTED_DOC_PREFIX}*.docx")
        files2 = glob.glob(pattern2)
        
        all_files = files1 + files2
        
        if all_files:
            return max(all_files, key=os.path.getctime)
        else:
            return None

# 创建全局配置实例
config = Config()