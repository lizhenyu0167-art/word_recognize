#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
检查Heading 4样式的字号设置
"""

from docx import Document
from docx.oxml.ns import qn

def check_heading4_font_size():
    """检查Heading 4样式的字号设置"""
    
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
    
    print(f"=== Heading 4样式字号检查 ===")
    print(f"样式名称: {heading4_style.name}")
    
    # 检查基本字体属性
    if hasattr(heading4_style, 'font'):
        font = heading4_style.font
        print(f"\n=== 基本字体属性 ===")
        print(f"font.name: {font.name}")
        print(f"font.size: {font.size}")
        if font.size:
            print(f"font.size.pt: {font.size.pt}")
        print(f"font.bold: {font.bold}")
        print(f"font.italic: {font.italic}")
    
    # 检查XML级别的字号设置
    if hasattr(heading4_style, '_element'):
        style_element = heading4_style._element
        print(f"\n=== XML级别字号检查 ===")
        
        # 查找rPr元素
        rpr = style_element.find(qn('w:rPr'))
        if rpr is not None:
            print("找到rPr元素")
            
            # 查找sz元素（字号设置）
            sz = rpr.find(qn('w:sz'))
            if sz is not None:
                sz_val = sz.get(qn('w:val'))
                print(f"找到sz元素，值: {sz_val} (半点，实际字号: {int(sz_val)/2}pt)")
            else:
                print("未找到sz元素（字号未设置）")
            
            # 查找szCs元素（复杂脚本字号设置）
            szCs = rpr.find(qn('w:szCs'))
            if szCs is not None:
                szCs_val = szCs.get(qn('w:val'))
                print(f"找到szCs元素，值: {szCs_val} (半点，实际字号: {int(szCs_val)/2}pt)")
            else:
                print("未找到szCs元素（复杂脚本字号未设置）")
        else:
            print("未找到rPr元素")
    
    # 检查基础样式继承
    if hasattr(heading4_style, 'base_style') and heading4_style.base_style:
        print(f"\n=== 基础样式继承 ===")
        base_style = heading4_style.base_style
        print(f"基础样式: {base_style.name}")
        
        if hasattr(base_style, 'font') and base_style.font.size:
            print(f"基础样式字号: {base_style.font.size.pt}pt")
        else:
            print("基础样式字号: 未设置")
    else:
        print("\n=== 无基础样式继承 ===")
    
    # 检查文档默认字号
    print(f"\n=== 文档默认字号检查 ===")
    try:
        # 查找文档默认设置
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

if __name__ == "__main__":
    check_heading4_font_size()