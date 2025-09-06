#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
样式格式检查器
检查格式化后文档中样式的实际格式设置
"""

import json
from docx import Document
from docx.shared import Pt
from config import config

def get_font_info(font):
    """获取字体信息"""
    if not font:
        return None
    
    return {
        'name': font.name,
        'size': str(font.size) if font.size else None,
        'bold': font.bold,
        'italic': font.italic,
        'underline': str(font.underline) if font.underline else None,
        'color': str(font.color.rgb) if font.color and font.color.rgb else None
    }

def get_paragraph_format_info(paragraph_format):
    """获取段落格式信息"""
    if not paragraph_format:
        return None
    
    return {
        'alignment': str(paragraph_format.alignment) if paragraph_format.alignment else None,
        'first_line_indent': str(paragraph_format.first_line_indent) if paragraph_format.first_line_indent else None,
        'left_indent': str(paragraph_format.left_indent) if paragraph_format.left_indent else None,
        'right_indent': str(paragraph_format.right_indent) if paragraph_format.right_indent else None,
        'space_before': str(paragraph_format.space_before) if paragraph_format.space_before else None,
        'space_after': str(paragraph_format.space_after) if paragraph_format.space_after else None,
        'line_spacing': str(paragraph_format.line_spacing) if paragraph_format.line_spacing else None
    }

def analyze_document_styles(doc_path):
    """分析文档中的样式格式"""
    doc = Document(doc_path)
    styles_analysis = {}
    
    # 分析所有样式
    for style in doc.styles:
        if hasattr(style, 'font') and hasattr(style, 'paragraph_format'):
            style_info = {
                'style_name': style.name,
                'style_type': str(style.type),
                'font_info': get_font_info(style.font),
                'paragraph_format': get_paragraph_format_info(style.paragraph_format)
            }
            
            # 检查XML级别的字体设置
            if hasattr(style, '_element') and style._element is not None:
                try:
                    # 查找rPr元素
                    rPr = style._element.find('.//w:rPr', style._element.nsmap)
                    if rPr is not None:
                        # 查找字体设置
                        rFonts = rPr.find('.//w:rFonts', rPr.nsmap)
                        if rFonts is not None:
                            style_info['xml_font_info'] = {
                                'ascii': rFonts.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}ascii'),
                                'hAnsi': rFonts.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}hAnsi'),
                                'eastAsia': rFonts.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}eastAsia'),
                                'cs': rFonts.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}cs')
                            }
                        
                        # 查找字号设置
                        sz = rPr.find('.//w:sz', rPr.nsmap)
                        if sz is not None:
                            style_info['xml_size'] = sz.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val')
                        
                        # 查找粗体设置
                        b = rPr.find('.//w:b', rPr.nsmap)
                        if b is not None:
                            style_info['xml_bold'] = b.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', 'true')
                except Exception as e:
                    style_info['xml_error'] = str(e)
            
            styles_analysis[style.name] = style_info
    
    return styles_analysis

def main():
    """主函数"""
    print("开始分析样式格式...")
    
    # 确保输出目录存在
    config.ensure_output_dir()
    
    # 分析格式化后的文档
    formatted_doc_path = "output\\格式化后的测试文档_1756862137.docx"
    print(f"分析文档: {formatted_doc_path}")
    
    # 执行分析
    styles_analysis = analyze_document_styles(formatted_doc_path)
    
    # 生成摘要
    summary = {
        'document_path': formatted_doc_path,
        'total_styles': len(styles_analysis),
        'styles_analysis': styles_analysis
    }
    
    # 保存结果
    output_path = config.OUTPUT_DIR + "/style_format_analysis.json"
    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(summary, f, ensure_ascii=False, indent=2)
    
    print(f"\n=== 样式格式分析摘要 ===")
    print(f"文档路径: {formatted_doc_path}")
    print(f"总样式数: {len(styles_analysis)}")
    print(f"\n分析报告已保存到: {output_path}")
    
    # 显示关键样式信息
    print("\n=== 关键样式格式信息 ===")
    key_styles = ['Heading 1', 'Heading 2', 'Heading 3', 'Normal']
    
    for style_name in key_styles:
        if style_name in styles_analysis:
            style_info = styles_analysis[style_name]
            print(f"\n样式: {style_name}")
            
            # 显示字体信息
            if style_info.get('font_info'):
                font_info = style_info['font_info']
                print(f"  字体名称: {font_info.get('name', '未设置')}")
                print(f"  字体大小: {font_info.get('size', '未设置')}")
                print(f"  粗体: {font_info.get('bold', '未设置')}")
            
            # 显示XML字体信息
            if style_info.get('xml_font_info'):
                xml_font = style_info['xml_font_info']
                print(f"  XML字体 - ASCII: {xml_font.get('ascii', '未设置')}")
                print(f"  XML字体 - 中文: {xml_font.get('eastAsia', '未设置')}")
            
            # 显示XML字号信息
            if style_info.get('xml_size'):
                print(f"  XML字号: {style_info['xml_size']}")
            
            # 显示XML粗体信息
            if style_info.get('xml_bold'):
                print(f"  XML粗体: {style_info['xml_bold']}")

if __name__ == "__main__":
    main()