#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
调试style_1样式的字体设置
"""

from docx import Document
from docx.oxml.ns import qn
import xml.etree.ElementTree as ET

def debug_style1_font():
    """调试style_1样式的字体设置"""
    
    template_path = "格式模板.docx"
    doc = Document(template_path)
    
    # 查找style_1样式
    style_1 = None
    for style in doc.styles:
        if style.name == "style_1":
            style_1 = style
            break
    
    if style_1 is None:
        print("未找到style_1样式")
        return
    
    print(f"=== style_1样式调试信息 ===")
    print(f"样式名称: {style_1.name}")
    print(f"样式类型: {style_1.type}")
    
    # 检查基本字体属性
    if hasattr(style_1, 'font'):
        font = style_1.font
        print(f"\n=== 基本字体属性 ===")
        print(f"font.name: {font.name}")
        print(f"font.size: {font.size}")
        print(f"font.bold: {font.bold}")
        print(f"font.italic: {font.italic}")
    
    # 检查XML结构
    if hasattr(style_1, '_element'):
        style_element = style_1._element
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
                
                # 检查是否有其他字体相关的元素
                for child in rpr:
                    tag_name = child.tag.split('}')[-1] if '}' in child.tag else child.tag
                    if 'font' in tag_name.lower() or tag_name in ['rFonts', 'sz', 'szCs']:
                        print(f"字体相关元素: {tag_name} = {child.attrib}")
            else:
                print("未找到rFonts元素")
        else:
            print("未找到rPr元素")
    
    # 检查是否继承自其他样式
    if hasattr(style_1, 'base_style') and style_1.base_style:
        print(f"\n=== 基础样式信息 ===")
        print(f"基础样式: {style_1.base_style.name}")
        
        # 检查基础样式的字体设置
        base_style = style_1.base_style
        if hasattr(base_style, 'font'):
            base_font = base_style.font
            print(f"基础样式字体名称: {base_font.name}")
            print(f"基础样式字体大小: {base_font.size}")

if __name__ == "__main__":
    debug_style1_font()