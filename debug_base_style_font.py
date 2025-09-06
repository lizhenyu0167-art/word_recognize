#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
调试style_1基础样式的字体设置
"""

from docx import Document
from docx.oxml.ns import qn
import xml.etree.ElementTree as ET

def debug_base_style_font():
    """调试style_1基础样式的字体设置"""
    
    template_path = "格式模板.docx"
    doc = Document(template_path)
    
    # 查找Body Text First Indent 2样式
    base_style = None
    for style in doc.styles:
        if style.name == "Body Text First Indent 2":
            base_style = style
            break
    
    if base_style is None:
        print("未找到Body Text First Indent 2样式")
        return
    
    print(f"=== Body Text First Indent 2样式调试信息 ===")
    print(f"样式名称: {base_style.name}")
    print(f"样式类型: {base_style.type}")
    
    # 检查基本字体属性
    if hasattr(base_style, 'font'):
        font = base_style.font
        print(f"\n=== 基本字体属性 ===")
        print(f"font.name: {font.name}")
        print(f"font.size: {font.size}")
        print(f"font.bold: {font.bold}")
        print(f"font.italic: {font.italic}")
    
    # 检查XML结构
    if hasattr(base_style, '_element'):
        style_element = base_style._element
        print(f"\n=== XML结构分析 ===")
        
        # 打印完整的XML结构
        xml_str = ET.tostring(style_element, encoding='unicode')
        print(f"完整XML: {xml_str}")
        
        # 查找rPr元素
        rpr = style_element.find(qn('w:rPr'))
        if rpr is not None:
            print(f"\n找到rPr元素")
            
            # 查找rFonts元素
            rfonts = rpr.find(qn('w:rFonts'))
            if rfonts is not None:
                print(f"找到rFonts元素")
                print(f"rFonts属性: {rfonts.attrib}")
                
                # 检查各种字体属性
                ascii_font = rfonts.get(qn('w:ascii'))
                hansi_font = rfonts.get(qn('w:hAnsi'))
                eastasia_font = rfonts.get(qn('w:eastAsia'))
                cs_font = rfonts.get(qn('w:cs'))
                
                print(f"\n=== 字体分离详情 ===")
                print(f"ascii字体: {ascii_font if ascii_font else '未设置'}")
                print(f"hAnsi字体: {hansi_font if hansi_font else '未设置'}")
                print(f"eastAsia字体: {eastasia_font if eastasia_font else '未设置'}")
                print(f"cs字体: {cs_font if cs_font else '未设置'}")
            else:
                print("未找到rFonts元素")
        else:
            print("未找到rPr元素")
    
    # 检查是否继承自其他样式
    if hasattr(base_style, 'base_style') and base_style.base_style:
        print(f"\n=== 基础样式信息 ===")
        print(f"基础样式: {base_style.base_style.name}")
    else:
        print(f"\n=== 没有基础样式，继承自Normal或文档默认 ===")
    
    # 检查Normal样式的字体设置
    normal_style = None
    for style in doc.styles:
        if style.name == "Normal":
            normal_style = style
            break
    
    if normal_style:
        print(f"\n=== Normal样式字体设置 ===")
        if hasattr(normal_style, 'font'):
            normal_font = normal_style.font
            print(f"Normal字体名称: {normal_font.name}")
            print(f"Normal字体大小: {normal_font.size}")
        
        # 检查Normal样式的XML字体分离
        if hasattr(normal_style, '_element'):
            normal_element = normal_style._element
            normal_rpr = normal_element.find(qn('w:rPr'))
            if normal_rpr is not None:
                normal_rfonts = normal_rpr.find(qn('w:rFonts'))
                if normal_rfonts is not None:
                    print(f"Normal样式字体分离: {normal_rfonts.attrib}")
                else:
                    print("Normal样式没有字体分离设置")
            else:
                print("Normal样式没有rPr元素")

if __name__ == "__main__":
    debug_base_style_font()