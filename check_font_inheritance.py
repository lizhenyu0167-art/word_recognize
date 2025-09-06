#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
检查字体继承链条
"""

from docx import Document
from docx.oxml.ns import qn

def check_font_inheritance():
    """检查字体继承链条"""
    
    template_path = "格式模板.docx"
    doc = Document(template_path)
    
    # 查找Heading 4样式
    heading4_style = None
    for style in doc.styles:
        if style.name == "Heading 4":
            heading4_style = style
            break
    
    if heading4_style is None:
        print("未找到Heading 4样式")
        return
    
    print(f"=== Heading 4字体继承链条分析 ===")
    
    # 遍历继承链条
    current_style = heading4_style
    level = 0
    
    while current_style and level < 10:  # 防止无限循环
        indent = "  " * level
        print(f"\n{indent}样式: {current_style.name}")
        
        # 检查当前样式的字号设置
        if hasattr(current_style, 'font'):
            font = current_style.font
            print(f"{indent}  font.size: {font.size}")
            if font.size:
                print(f"{indent}  font.size.pt: {font.size.pt}pt")
        
        # 检查XML级别的字号
        if hasattr(current_style, '_element'):
            style_element = current_style._element
            rpr = style_element.find(qn('w:rPr'))
            if rpr is not None:
                sz = rpr.find(qn('w:sz'))
                if sz is not None:
                    sz_val = sz.get(qn('w:val'))
                    print(f"{indent}  XML字号: {int(sz_val)/2}pt")
                else:
                    print(f"{indent}  XML字号: 未设置")
            else:
                print(f"{indent}  无rPr元素")
        
        # 移动到基础样式
        if hasattr(current_style, 'base_style') and current_style.base_style:
            current_style = current_style.base_style
            level += 1
        else:
            break
    
    # 检查Normal样式
    print(f"\n=== Normal样式检查 ===")
    normal_style = None
    for style in doc.styles:
        if style.name == "Normal":
            normal_style = style
            break
    
    if normal_style:
        print(f"Normal样式存在")
        if hasattr(normal_style, 'font'):
            font = normal_style.font
            print(f"  font.size: {font.size}")
            if font.size:
                print(f"  font.size.pt: {font.size.pt}pt")
        
        # 检查Normal样式的XML字号
        if hasattr(normal_style, '_element'):
            style_element = normal_style._element
            rpr = style_element.find(qn('w:rPr'))
            if rpr is not None:
                sz = rpr.find(qn('w:sz'))
                if sz is not None:
                    sz_val = sz.get(qn('w:val'))
                    print(f"  Normal XML字号: {int(sz_val)/2}pt")
                else:
                    print(f"  Normal XML字号: 未设置")
    
    # 检查文档默认字号
    print(f"\n=== 文档默认字号检查 ===")
    try:
        styles_element = doc.styles.element
        doc_defaults = styles_element.find(qn('w:docDefaults'))
        if doc_defaults is not None:
            rpr_default = doc_defaults.find(qn('w:rPrDefault'))
            if rpr_default is not None:
                rpr = rpr_default.find(qn('w:rPr'))
                if rpr is not None:
                    sz = rpr.find(qn('w:sz'))
                    if sz is not None:
                        sz_val = sz.get(qn('w:val'))
                        print(f"文档默认字号: {int(sz_val)/2}pt")
                    else:
                        print("文档默认字号: 未设置")
                else:
                    print("文档默认字号: 未设置")
            else:
                print("文档默认字号: 未设置")
        else:
            print("文档默认字号: 未设置")
    except Exception as e:
        print(f"检查文档默认字号时出错: {e}")
    
    # 检查Word默认字号（通常是12pt）
    print(f"\n=== Word应用程序默认字号 ===")
    print("Word应用程序默认字号通常为12pt")
    print("当样式链条中都没有明确设置字号时，应该使用12pt作为默认值")

if __name__ == "__main__":
    check_font_inheritance()