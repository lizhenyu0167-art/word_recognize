#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
模板样式检查器
检查模板文档中样式的格式设置
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

def analyze_template_styles(doc_path):
    """分析模板文档中的样式格式"""
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
                        
                        # 查找字号设置（复杂脚本）
                        szCs = rPr.find('.//w:szCs', rPr.nsmap)
                        if szCs is not None:
                            style_info['xml_size_cs'] = szCs.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val')
                        
                        # 查找粗体设置
                        b = rPr.find('.//w:b', rPr.nsmap)
                        if b is not None:
                            style_info['xml_bold'] = b.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', 'true')
                        
                        # 查找粗体设置（复杂脚本）
                        bCs = rPr.find('.//w:bCs', rPr.nsmap)
                        if bCs is not None:
                            style_info['xml_bold_cs'] = bCs.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', 'true')
                except Exception as e:
                    style_info['xml_error'] = str(e)
            
            styles_analysis[style.name] = style_info
    
    return styles_analysis

def main():
    """主函数"""
    print("开始分析模板样式格式...")
    
    # 确保输出目录存在
    config.ensure_output_dir()
    
    # 分析模板文档
    template_doc_path = config.TEMPLATE_FILE
    print(f"分析模板文档: {template_doc_path}")
    
    # 执行分析
    template_styles = analyze_template_styles(template_doc_path)
    
    # 分析格式化后的文档
    formatted_doc_path = "output\\格式化后的测试文档_已修复.docx"
    print(f"分析格式化文档: {formatted_doc_path}")
    formatted_styles = analyze_template_styles(formatted_doc_path)
    
    # 生成对比摘要
    comparison = {
        'template_document': template_doc_path,
        'formatted_document': formatted_doc_path,
        'template_styles': template_styles,
        'formatted_styles': formatted_styles,
        'comparison_summary': {}
    }
    
    # 对比关键样式
    key_styles = ['Heading 1', 'Heading 2', 'Heading 3', 'Normal']
    
    for style_name in key_styles:
        if style_name in template_styles and style_name in formatted_styles:
            template_style = template_styles[style_name]
            formatted_style = formatted_styles[style_name]
            
            comparison_result = {
                'template_size': template_style.get('xml_size'),
                'formatted_size': formatted_style.get('xml_size'),
                'template_font_ascii': template_style.get('xml_font_info', {}).get('ascii'),
                'formatted_font_ascii': formatted_style.get('xml_font_info', {}).get('ascii'),
                'template_font_eastasia': template_style.get('xml_font_info', {}).get('eastAsia'),
                'formatted_font_eastasia': formatted_style.get('xml_font_info', {}).get('eastAsia'),
                'size_match': template_style.get('xml_size') == formatted_style.get('xml_size'),
                'font_match': (
                    template_style.get('xml_font_info', {}).get('ascii') == formatted_style.get('xml_font_info', {}).get('ascii') and
                    template_style.get('xml_font_info', {}).get('eastAsia') == formatted_style.get('xml_font_info', {}).get('eastAsia')
                )
            }
            
            comparison['comparison_summary'][style_name] = comparison_result
    
    # 保存结果
    output_path = config.OUTPUT_DIR + "/template_style_comparison.json"
    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(comparison, f, ensure_ascii=False, indent=2)
    
    print(f"\n=== 模板与格式化文档样式对比 ===")
    print(f"模板文档: {template_doc_path}")
    print(f"格式化文档: {formatted_doc_path}")
    print(f"\n对比报告已保存到: {output_path}")
    
    # 显示对比结果
    print("\n=== 样式对比结果 ===")
    for style_name in key_styles:
        if style_name in comparison['comparison_summary']:
            comp = comparison['comparison_summary'][style_name]
            print(f"\n样式: {style_name}")
            print(f"  模板字号: {comp['template_size']} -> 格式化字号: {comp['formatted_size']} {'✅' if comp['size_match'] else '❌'}")
            print(f"  模板ASCII字体: {comp['template_font_ascii']} -> 格式化ASCII字体: {comp['formatted_font_ascii']}")
            print(f"  模板中文字体: {comp['template_font_eastasia']} -> 格式化中文字体: {comp['formatted_font_eastasia']}")
            print(f"  字体匹配: {'✅' if comp['font_match'] else '❌'}")

if __name__ == "__main__":
    main()