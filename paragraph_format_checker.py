#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
段落格式检查器
检查格式化后文档中每个段落的实际格式应用情况
"""

import json
from docx import Document
from docx.shared import Pt
from config import config

def analyze_paragraph_formats(doc_path):
    """分析文档中每个段落的格式"""
    doc = Document(doc_path)
    paragraph_analysis = []
    
    for i, paragraph in enumerate(doc.paragraphs):
        para_info = {
            'paragraph_index': i,
            'text_preview': paragraph.text[:50] + '...' if len(paragraph.text) > 50 else paragraph.text,
            'style_name': paragraph.style.name if paragraph.style else 'None',
            'runs_count': len(paragraph.runs),
            'runs_analysis': []
        }
        
        # 分析段落级别格式
        if paragraph.paragraph_format:
            pf = paragraph.paragraph_format
            para_info['paragraph_format'] = {
                'alignment': str(pf.alignment) if pf.alignment else None,
                'first_line_indent': str(pf.first_line_indent) if pf.first_line_indent else None,
                'left_indent': str(pf.left_indent) if pf.left_indent else None,
                'right_indent': str(pf.right_indent) if pf.right_indent else None,
                'space_before': str(pf.space_before) if pf.space_before else None,
                'space_after': str(pf.space_after) if pf.space_after else None,
                'line_spacing': str(pf.line_spacing) if pf.line_spacing else None
            }
        
        # 分析每个run的格式
        for j, run in enumerate(paragraph.runs):
            run_info = {
                'run_index': j,
                'text': run.text,
                'font_info': {}
            }
            
            if run.font:
                font = run.font
                run_info['font_info'] = {
                    'name': font.name,
                    'size': str(font.size) if font.size else None,
                    'bold': font.bold,
                    'italic': font.italic,
                    'underline': str(font.underline) if font.underline else None,
                    'color': str(font.color.rgb) if font.color and font.color.rgb else None
                }
                
                # 检查XML级别的字体设置
                if hasattr(run, '_element') and run._element is not None:
                    rPr = run._element.find('.//w:rPr', run._element.nsmap)
                    if rPr is not None:
                        rFonts = rPr.find('.//w:rFonts', rPr.nsmap)
                        if rFonts is not None:
                            run_info['xml_font_info'] = {
                                'ascii': rFonts.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}ascii'),
                                'hAnsi': rFonts.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}hAnsi'),
                                'eastAsia': rFonts.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}eastAsia'),
                                'cs': rFonts.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}cs')
                            }
            
            para_info['runs_analysis'].append(run_info)
        
        paragraph_analysis.append(para_info)
    
    return paragraph_analysis

def main():
    """主函数"""
    print("开始分析段落格式...")
    
    # 确保输出目录存在
    config.ensure_output_dir()
    
    # 分析最新的格式化文档
    formatted_doc_path = "output\\格式化后的测试文档_1756862137.docx"
    print(f"分析文档: {formatted_doc_path}")
    
    # 执行分析
    analysis_result = analyze_paragraph_formats(formatted_doc_path)
    
    # 生成摘要
    total_paragraphs = len(analysis_result)
    paragraphs_with_runs = sum(1 for p in analysis_result if p['runs_count'] > 0)
    
    summary = {
        'document_path': formatted_doc_path,
        'total_paragraphs': total_paragraphs,
        'paragraphs_with_runs': paragraphs_with_runs,
        'analysis_details': analysis_result
    }
    
    # 保存结果
    output_path = config.OUTPUT_DIR + "/paragraph_format_analysis.json"
    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(summary, f, ensure_ascii=False, indent=2)
    
    print(f"\n=== 段落格式分析摘要 ===")
    print(f"文档路径: {formatted_doc_path}")
    print(f"总段落数: {total_paragraphs}")
    print(f"有内容段落数: {paragraphs_with_runs}")
    print(f"\n分析报告已保存到: {output_path}")
    
    # 显示关键段落信息
    print("\n=== 关键段落格式信息 ===")
    for para in analysis_result:
        if para['text_preview'].strip() and ('标题' in para['text_preview'] or '正文' in para['text_preview'] or len(para['text_preview']) > 10):
            print(f"段落 {para['paragraph_index']}: {para['style_name']}")
            print(f"  内容: {para['text_preview']}")
            if para['runs_analysis']:
                for run in para['runs_analysis']:
                    if run['font_info'].get('name'):
                        print(f"    字体: {run['font_info']['name']}, 大小: {run['font_info']['size']}")
            print()

if __name__ == "__main__":
    main()