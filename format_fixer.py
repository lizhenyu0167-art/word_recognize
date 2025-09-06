#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
格式修复器
修复格式化文档中的样式问题
"""

import json
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn
from config import config

def fix_heading3_font_size(doc_path, output_path):
    """修复Heading 3样式的字号问题"""
    try:
        doc = Document(doc_path)
        
        # 查找Heading 3样式
        heading3_style = None
        for style in doc.styles:
            if style.name == 'Heading 3':
                heading3_style = style
                break
        
        if not heading3_style:
            print("未找到Heading 3样式")
            return False
        
        print(f"修复前 Heading 3 字号: {heading3_style.font.size}")
        
        # 清除字号设置，让它继承默认值
        if hasattr(heading3_style, '_element'):
            style_element = heading3_style._element
            
            # 查找rPr元素
            rpr = style_element.find(qn('w:rPr'))
            if rpr is not None:
                # 查找并删除sz元素（字号设置）
                sz = rpr.find(qn('w:sz'))
                if sz is not None:
                    rpr.remove(sz)
                    print("已删除Heading 3的字号设置")
                
                # 查找并删除szCs元素（复杂脚本字号设置）
                szCs = rpr.find(qn('w:szCs'))
                if szCs is not None:
                    rpr.remove(szCs)
                    print("已删除Heading 3的复杂脚本字号设置")
        
        # 同时清除font.size属性
        if hasattr(heading3_style, 'font'):
            heading3_style.font.size = None
        
        print(f"修复后 Heading 3 字号: {heading3_style.font.size}")
        
        # 保存修复后的文档
        doc.save(output_path)
        print(f"修复完成，文档已保存为: {output_path}")
        return True
        
    except Exception as e:
        print(f"修复Heading 3字号时出错: {e}")
        return False

def fix_normal_font_settings(doc_path, output_path):
    """修复Normal样式的字体设置问题"""
    try:
        doc = Document(doc_path)
        
        # 查找Normal样式
        normal_style = None
        for style in doc.styles:
            if style.name == 'Normal':
                normal_style = style
                break
        
        if not normal_style:
            print("未找到Normal样式")
            return False
        
        print(f"修复前 Normal 字体: {normal_style.font.name}")
        
        # 设置正确的字体
        if hasattr(normal_style, 'font'):
            normal_style.font.name = '宋体'
        
        # 修复XML级别的字体设置
        if hasattr(normal_style, '_element'):
            style_element = normal_style._element
            
            # 查找或创建rPr元素
            rpr = style_element.find(qn('w:rPr'))
            if rpr is None:
                rpr = style_element.makeelement(qn('w:rPr'))
                style_element.insert(0, rpr)
            
            # 查找或创建rFonts元素
            rfonts = rpr.find(qn('w:rFonts'))
            if rfonts is None:
                rfonts = rpr.makeelement(qn('w:rFonts'))
                rpr.insert(0, rfonts)
            
            # 清除错误的字体设置
            if rfonts.get(qn('w:ascii')):
                rfonts.attrib.pop(qn('w:ascii'), None)
            if rfonts.get(qn('w:eastAsia')):
                rfonts.attrib.pop(qn('w:eastAsia'), None)
            
            print("已清除Normal样式的XML字体设置")
        
        print(f"修复后 Normal 字体: {normal_style.font.name}")
        
        # 保存修复后的文档
        doc.save(output_path)
        print(f"修复完成，文档已保存为: {output_path}")
        return True
        
    except Exception as e:
        print(f"修复Normal字体时出错: {e}")
        return False

def main():
    """主函数"""
    print("开始修复格式问题...")
    
    # 确保输出目录存在
    config.ensure_output_dir()
    
    # 输入和输出路径
    input_doc = "output\\格式化后的测试文档_1756862137.docx"
    temp_doc = "output\\格式化后的测试文档_修复中.docx"
    output_doc = "output\\格式化后的测试文档_已修复.docx"
    
    print(f"输入文档: {input_doc}")
    print(f"输出文档: {output_doc}")
    
    # 步骤1：修复Heading 3字号问题
    print("\n=== 修复Heading 3字号问题 ===")
    if fix_heading3_font_size(input_doc, temp_doc):
        print("Heading 3字号修复成功")
    else:
        print("Heading 3字号修复失败")
        return
    
    # 步骤2：修复Normal字体问题
    print("\n=== 修复Normal字体问题 ===")
    if fix_normal_font_settings(temp_doc, output_doc):
        print("Normal字体修复成功")
    else:
        print("Normal字体修复失败")
        return
    
    print("\n=== 格式修复完成 ===")
    print(f"修复后的文档已保存为: {output_doc}")
    
    # 删除临时文件
    import os
    try:
        os.remove(temp_doc)
        print("已删除临时文件")
    except:
        pass

if __name__ == "__main__":
    main()