#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Run级别格式分析器
分析Word文档中段落内容的run级别格式设置
用于排查样式级别格式与run级别格式的冲突问题
"""

import os
import json
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn
from config import config

class RunFormatAnalyzer:
    def __init__(self):
        self.analysis_results = {}
    
    def analyze_document_runs(self, doc_path, doc_name):
        """
        分析文档中所有段落的run级别格式
        """
        try:
            doc = Document(doc_path)
            doc_analysis = {
                'document_name': doc_name,
                'paragraphs': []
            }
            
            for para_idx, paragraph in enumerate(doc.paragraphs):
                if paragraph.text.strip():  # 只分析有内容的段落
                    para_analysis = {
                        'paragraph_index': para_idx + 1,
                        'text_preview': paragraph.text[:50] + '...' if len(paragraph.text) > 50 else paragraph.text,
                        'style_name': paragraph.style.name,
                        'runs': []
                    }
                    
                    # 分析每个run的格式
                    for run_idx, run in enumerate(paragraph.runs):
                        if run.text.strip():  # 只分析有内容的run
                            run_analysis = self._analyze_run_format(run, run_idx)
                            para_analysis['runs'].append(run_analysis)
                    
                    doc_analysis['paragraphs'].append(para_analysis)
            
            return doc_analysis
            
        except Exception as e:
            print(f"分析文档 {doc_name} 的run格式时出错: {e}")
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
    
    def compare_documents(self, template_path, formatted_path):
        """
        比较模板文档和格式化文档的run级别格式
        """
        print("=== Run级别格式分析 ===")
        
        # 分析模板文档
        print("\n分析格式模板文档...")
        template_analysis = self.analyze_document_runs(template_path, "格式模板")
        
        # 分析格式化文档
        print("分析格式化文档...")
        formatted_analysis = self.analyze_document_runs(formatted_path, "格式化文档")
        
        if not template_analysis or not formatted_analysis:
            print("文档分析失败")
            return
        
        # 生成比较报告
        comparison_report = {
            'template_analysis': template_analysis,
            'formatted_analysis': formatted_analysis,
            'comparison_summary': self._generate_comparison_summary(template_analysis, formatted_analysis)
        }
        
        # 保存分析结果
        output_path = os.path.join(config.OUTPUT_DIR, "run_format_analysis.json")
        config.ensure_output_dir()
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(comparison_report, f, ensure_ascii=False, indent=2)
        
        print(f"\nRun格式分析报告已保存到: {output_path}")
        
        # 显示摘要
        self._print_analysis_summary(comparison_report['comparison_summary'])
        
        return comparison_report
    
    def _generate_comparison_summary(self, template_analysis, formatted_analysis):
        """
        生成比较摘要
        """
        summary = {
            'template_paragraphs': len(template_analysis['paragraphs']),
            'formatted_paragraphs': len(formatted_analysis['paragraphs']),
            'run_format_issues': [],
            'font_separation_issues': []
        }
        
        # 检查段落数量是否匹配
        if summary['template_paragraphs'] != summary['formatted_paragraphs']:
            summary['run_format_issues'].append(f"段落数量不匹配: 模板{summary['template_paragraphs']}个，格式化文档{summary['formatted_paragraphs']}个")
        
        # 逐段落比较run格式
        min_paras = min(len(template_analysis['paragraphs']), len(formatted_analysis['paragraphs']))
        
        for i in range(min_paras):
            template_para = template_analysis['paragraphs'][i]
            formatted_para = formatted_analysis['paragraphs'][i]
            
            # 检查样式名称
            if template_para['style_name'] != formatted_para['style_name']:
                summary['run_format_issues'].append(
                    f"段落{i+1}样式不匹配: 模板'{template_para['style_name']}' vs 格式化'{formatted_para['style_name']}'"
                )
            
            # 检查run级别的字体设置
            self._compare_paragraph_runs(template_para, formatted_para, i+1, summary)
        
        return summary
    
    def _compare_paragraph_runs(self, template_para, formatted_para, para_num, summary):
        """
        比较段落中run的格式
        """
        template_runs = template_para.get('runs', [])
        formatted_runs = formatted_para.get('runs', [])
        
        # 检查是否存在run级别的格式覆盖
        for run in formatted_runs:
            font_info = run.get('font_info', {})
            xml_font_info = run.get('xml_font_info', {})
            
            # 检查是否有明确的字体设置（可能覆盖样式）
            if font_info.get('name') or xml_font_info:
                issue_desc = f"段落{para_num}存在run级别字体设置: "
                if font_info.get('name'):
                    issue_desc += f"字体名称={font_info['name']} "
                if xml_font_info:
                    issue_desc += f"XML字体分离={xml_font_info}"
                summary['font_separation_issues'].append(issue_desc)
    
    def _print_analysis_summary(self, summary):
        """
        打印分析摘要
        """
        print("\n=== Run格式分析摘要 ===")
        print(f"模板文档段落数: {summary['template_paragraphs']}")
        print(f"格式化文档段落数: {summary['formatted_paragraphs']}")
        
        if summary['run_format_issues']:
            print("\n发现的Run格式问题:")
            for issue in summary['run_format_issues']:
                print(f"  - {issue}")
        
        if summary['font_separation_issues']:
            print("\n发现的字体分离问题:")
            for issue in summary['font_separation_issues']:
                print(f"  - {issue}")
        
        if not summary['run_format_issues'] and not summary['font_separation_issues']:
            print("\n✓ 未发现明显的run级别格式问题")

def main():
    """
    主函数
    """
    # 验证必需文件
    missing_files = config.validate_required_files()
    if missing_files:
        print(f"错误：找不到以下必需文件: {', '.join(missing_files)}")
        return
    
    # 获取最新的格式化文档
    latest_formatted_doc = config.get_latest_formatted_doc()
    if not latest_formatted_doc:
        print("错误：找不到格式化文档")
        return
    
    analyzer = RunFormatAnalyzer()
    
    # 执行run格式分析
    analyzer.compare_documents(
        config.TEMPLATE_FILE,
        latest_formatted_doc
    )

if __name__ == "__main__":
    main()