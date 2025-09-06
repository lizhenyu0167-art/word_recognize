#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
测试文档分析器
分析测试文档的原始格式，检查是否存在run级别的格式设置
"""

import os
import json
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn
from config import config

class TestDocumentAnalyzer:
    def __init__(self):
        pass
    
    def analyze_test_document(self):
        """
        分析测试文档的原始格式
        """
        try:
            doc = Document(config.TEST_DOCUMENT)
            analysis = {
                'document_name': '测试文档（原始）',
                'paragraphs': []
            }
            
            print(f"=== 分析测试文档: {config.TEST_DOCUMENT} ===")
            
            for para_idx, paragraph in enumerate(doc.paragraphs):
                if paragraph.text.strip():  # 只分析有内容的段落
                    para_analysis = {
                        'paragraph_index': para_idx + 1,
                        'text_preview': paragraph.text[:50] + '...' if len(paragraph.text) > 50 else paragraph.text,
                        'style_name': paragraph.style.name,
                        'runs': [],
                        'has_run_formatting': False
                    }
                    
                    # 分析每个run的格式
                    for run_idx, run in enumerate(paragraph.runs):
                        if run.text.strip():  # 只分析有内容的run
                            run_analysis = self._analyze_run_format(run, run_idx)
                            para_analysis['runs'].append(run_analysis)
                            
                            # 检查是否有run级别的格式设置
                            if (run_analysis['font_info'].get('name') or 
                                run_analysis['xml_font_info'] or
                                run_analysis['font_info'].get('size') or
                                run_analysis['font_info'].get('bold') is not None):
                                para_analysis['has_run_formatting'] = True
                    
                    analysis['paragraphs'].append(para_analysis)
                    
                    # 打印有run格式的段落
                    if para_analysis['has_run_formatting']:
                        print(f"段落{para_idx + 1} ({para_analysis['style_name']}): {para_analysis['text_preview']}")
                        for run in para_analysis['runs']:
                            if (run['font_info'].get('name') or run['xml_font_info']):
                                print(f"  Run{run['run_index']}: 字体={run['font_info'].get('name')}, XML字体={run['xml_font_info']}")
            
            # 保存分析结果
            config.ensure_output_dir()
            output_path = os.path.join(config.OUTPUT_DIR, "test_document_analysis.json")
            with open(output_path, 'w', encoding='utf-8') as f:
                json.dump(analysis, f, ensure_ascii=False, indent=2)
            
            print(f"\n测试文档分析报告已保存到: {output_path}")
            
            # 统计摘要
            total_paras = len(analysis['paragraphs'])
            paras_with_run_formatting = sum(1 for p in analysis['paragraphs'] if p['has_run_formatting'])
            
            print(f"\n=== 分析摘要 ===")
            print(f"总段落数: {total_paras}")
            print(f"有run级别格式的段落数: {paras_with_run_formatting}")
            print(f"run格式比例: {paras_with_run_formatting/total_paras*100:.1f}%")
            
            return analysis
            
        except Exception as e:
            print(f"分析测试文档时出错: {e}")
            return None
    
    def _analyze_run_format(self, run, run_idx):
        """
        分析单个run的格式信息
        """
        run_info = {
            'run_index': run_idx + 1,
            'text_preview': run.text[:30] + '...' if len(run.text) > 30 else run.text,
            'font_info': {},
            'xml_font_info': {}
        }
        
        # 分析基本字体信息
        try:
            if hasattr(run, 'font'):
                font = run.font
                run_info['font_info'] = {
                    'name': font.name,
                    'size': f"{font.size.pt}pt" if font.size else None,
                    'bold': font.bold,
                    'italic': font.italic,
                    'underline': str(font.underline) if font.underline else None
                }
        except Exception as e:
            run_info['font_info']['error'] = str(e)
        
        # 分析XML级别的字体信息
        try:
            if hasattr(run, '_element'):
                run_element = run._element
                rpr = run_element.find(qn('w:rPr'))
                if rpr is not None:
                    rfonts = rpr.find(qn('w:rFonts'))
                    if rfonts is not None:
                        run_info['xml_font_info'] = {
                            'ascii': rfonts.get(qn('w:ascii')),
                            'hAnsi': rfonts.get(qn('w:hAnsi')),
                            'eastAsia': rfonts.get(qn('w:eastAsia')),
                            'cs': rfonts.get(qn('w:cs'))
                        }
                        # 移除None值
                        run_info['xml_font_info'] = {k: v for k, v in run_info['xml_font_info'].items() if v is not None}
        except Exception as e:
            run_info['xml_font_info']['error'] = str(e)
        
        return run_info

def main():
    """
    主函数
    """
    # 验证测试文档是否存在
    if not os.path.exists(config.TEST_DOCUMENT):
        print(f"错误：找不到测试文档 {config.TEST_DOCUMENT}")
        return
    
    analyzer = TestDocumentAnalyzer()
    analyzer.analyze_test_document()

if __name__ == "__main__":
    main()