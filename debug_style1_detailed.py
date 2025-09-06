#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
详细调试style_1样式的字体设置
"""

from docx import Document
from docx.oxml.ns import qn

def debug_style1_fonts():
    """
    详细分析style_1样式的字体设置
    """
    print("=== 详细调试style_1样式字体设置 ===")
    
    # 加载格式模板文档
    doc = Document('格式模板.docx')
    
    # 查找style_1样式
    style_1 = None
    for style in doc.styles:
        if style.name == 'style_1':
            style_1 = style
            break
    
    if not style_1:
        print("未找到style_1样式")
        return
    
    print(f"样式名称: {style_1.name}")
    print(f"样式类型: {style_1.type}")
    
    # 检查基础样式链
    print("\n=== 样式继承链 ===")
    current_style = style_1
    level = 0
    while current_style:
        indent = "  " * level
        print(f"{indent}{current_style.name}")
        
        # 显示当前样式的直接字体设置
        if hasattr(current_style, '_element'):
            style_element = current_style._element
            rpr = style_element.find(qn('w:rPr'))
            if rpr is not None:
                rfonts = rpr.find(qn('w:rFonts'))
                if rfonts is not None:
                    print(f"{indent}  直接字体设置:")
                    for attr in ['ascii', 'hAnsi', 'eastAsia', 'cs']:
                        font_val = rfonts.get(qn(f'w:{attr}'))
                        if font_val:
                            print(f"{indent}    {attr}: {font_val}")
                else:
                    print(f"{indent}  无直接字体设置")
            else:
                print(f"{indent}  无rPr元素")
        
        # 移动到基础样式
        if hasattr(current_style, 'base_style') and current_style.base_style:
            current_style = current_style.base_style
            level += 1
        else:
            break
    
    # 检查文档默认设置
    print("\n=== 文档默认设置 ===")
    try:
        doc_defaults = doc.settings.element.find(qn('w:defaultRunProperties'))
        if doc_defaults is not None:
            default_rfonts = doc_defaults.find(qn('w:rFonts'))
            if default_rfonts is not None:
                print("文档默认字体:")
                for attr in ['ascii', 'hAnsi', 'eastAsia', 'cs']:
                    font_val = default_rfonts.get(qn(f'w:{attr}'))
                    if font_val:
                        print(f"  {attr}: {font_val}")
            else:
                print("无默认字体设置")
        else:
            print("无文档默认运行属性")
    except Exception as e:
        print(f"检查文档默认设置时出错: {e}")
    
    # 检查主题字体
    print("\n=== 主题字体设置 ===")
    try:
        theme_part = doc.part.package.part_related_by("http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme")
        if theme_part:
            print("找到主题文件")
            # 这里可以进一步解析主题字体，但比较复杂
        else:
            print("未找到主题文件")
    except Exception as e:
        print(f"检查主题字体时出错: {e}")

if __name__ == "__main__":
    debug_style1_fonts()