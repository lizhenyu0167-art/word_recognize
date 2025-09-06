#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
字体和行距分析说明工具
提供字体格式和行距调整的详细分析和说明功能
"""

import json
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.style import WD_STYLE_TYPE
from datetime import datetime
from config import config

class FontAndSpacingAnalyzer:
    def __init__(self):
        self.line_spacing_types = {
            WD_LINE_SPACING.SINGLE: '单倍行距',
            WD_LINE_SPACING.ONE_POINT_FIVE: '1.5倍行距', 
            WD_LINE_SPACING.DOUBLE: '双倍行距',
            WD_LINE_SPACING.AT_LEAST: '最小值',
            WD_LINE_SPACING.EXACTLY: '固定值',
            WD_LINE_SPACING.MULTIPLE: '多倍行距'
        }
        
        self.alignment_types = {
            WD_ALIGN_PARAGRAPH.LEFT: '左对齐',
            WD_ALIGN_PARAGRAPH.CENTER: '居中对齐',
            WD_ALIGN_PARAGRAPH.RIGHT: '右对齐',
            WD_ALIGN_PARAGRAPH.JUSTIFY: '两端对齐',
            WD_ALIGN_PARAGRAPH.DISTRIBUTE: '分散对齐'
        }
    
    def analyze_document_spacing(self, doc_path):
        """
        分析文档中的行距和段落格式
        """
        try:
            doc = Document(doc_path)
            analysis_result = {
                'document_path': doc_path,
                'styles_analysis': {},
                'paragraphs_analysis': [],
                'spacing_summary': {}
            }
            
            # 分析样式中的行距设置
            for style in doc.styles:
                if style.type == WD_STYLE_TYPE.PARAGRAPH:
                    style_info = self._analyze_style_spacing(style)
                    if style_info:
                        analysis_result['styles_analysis'][style.name] = style_info
            
            # 分析实际段落的行距应用
            for i, paragraph in enumerate(doc.paragraphs):
                if paragraph.text.strip():  # 只分析有内容的段落
                    para_info = self._analyze_paragraph_spacing(paragraph, i)
                    analysis_result['paragraphs_analysis'].append(para_info)
            
            # 生成行距使用摘要
            analysis_result['spacing_summary'] = self._generate_spacing_summary(analysis_result)
            
            return analysis_result
            
        except Exception as e:
            print(f"分析文档行距时出错: {e}")
            return None
    
    def _analyze_style_spacing(self, style):
        """
        分析样式的行距设置
        """
        try:
            if not hasattr(style, 'paragraph_format'):
                return None
                
            pf = style.paragraph_format
            style_info = {
                'style_name': style.name,
                'font_info': {},
                'spacing_info': {}
            }
            
            # 字体信息
            if hasattr(style, 'font'):
                font = style.font
                style_info['font_info'] = {
                    'name': font.name,
                    'size': f"{font.size.pt}pt" if font.size else None,
                    'bold': font.bold,
                    'italic': font.italic
                }
            
            # 行距信息
            spacing_info = {}
            
            if pf.line_spacing is not None:
                spacing_info['line_spacing'] = {
                    'value': str(pf.line_spacing),
                    'description': self._describe_line_spacing(pf.line_spacing, pf.line_spacing_rule)
                }
            
            if pf.space_before is not None:
                spacing_info['space_before'] = f"{pf.space_before.pt}pt"
            
            if pf.space_after is not None:
                spacing_info['space_after'] = f"{pf.space_after.pt}pt"
            
            if pf.first_line_indent is not None:
                spacing_info['first_line_indent'] = f"{pf.first_line_indent.pt}pt"
            
            if pf.left_indent is not None:
                spacing_info['left_indent'] = f"{pf.left_indent.pt}pt"
            
            if pf.alignment is not None:
                spacing_info['alignment'] = self.alignment_types.get(pf.alignment, '未知对齐')
            
            style_info['spacing_info'] = spacing_info
            
            return style_info if spacing_info else None
            
        except Exception as e:
            print(f"分析样式 {style.name} 的行距时出错: {e}")
            return None
    
    def _analyze_paragraph_spacing(self, paragraph, index):
        """
        分析段落的实际行距应用
        """
        para_info = {
            'paragraph_index': index + 1,
            'text_preview': paragraph.text[:50] + '...' if len(paragraph.text) > 50 else paragraph.text,
            'style_name': paragraph.style.name,
            'actual_spacing': {}
        }
        
        try:
            pf = paragraph.paragraph_format
            
            if pf.line_spacing is not None:
                para_info['actual_spacing']['line_spacing'] = {
                    'value': str(pf.line_spacing),
                    'description': self._describe_line_spacing(pf.line_spacing, pf.line_spacing_rule)
                }
            
            if pf.space_before is not None:
                para_info['actual_spacing']['space_before'] = f"{pf.space_before.pt}pt"
            
            if pf.space_after is not None:
                para_info['actual_spacing']['space_after'] = f"{pf.space_after.pt}pt"
            
            if pf.alignment is not None:
                para_info['actual_spacing']['alignment'] = self.alignment_types.get(pf.alignment, '未知对齐')
                
        except Exception as e:
            print(f"分析段落 {index+1} 的行距时出错: {e}")
        
        return para_info
    
    def _describe_line_spacing(self, line_spacing, line_spacing_rule=None):
        """
        描述行距设置
        """
        try:
            if line_spacing_rule is not None:
                rule_desc = self.line_spacing_types.get(line_spacing_rule, '未知规则')
                return f"{rule_desc} ({line_spacing})"
            else:
                # 根据数值判断行距类型
                spacing_value = float(line_spacing)
                if spacing_value == 1.0:
                    return "单倍行距 (1.0)"
                elif spacing_value == 1.5:
                    return "1.5倍行距 (1.5)"
                elif spacing_value == 2.0:
                    return "双倍行距 (2.0)"
                else:
                    return f"自定义行距 ({spacing_value})"
        except:
            return f"行距值: {line_spacing}"
    
    def _generate_spacing_summary(self, analysis_result):
        """
        生成行距使用摘要
        """
        summary = {
            'total_styles': len(analysis_result['styles_analysis']),
            'total_paragraphs': len(analysis_result['paragraphs_analysis']),
            'line_spacing_usage': {},
            'alignment_usage': {},
            'spacing_issues': []
        }
        
        # 统计行距使用情况
        line_spacing_count = {}
        alignment_count = {}
        
        # 从样式统计
        for style_name, style_info in analysis_result['styles_analysis'].items():
            spacing_info = style_info.get('spacing_info', {})
            
            if 'line_spacing' in spacing_info:
                desc = spacing_info['line_spacing']['description']
                line_spacing_count[desc] = line_spacing_count.get(desc, 0) + 1
            
            if 'alignment' in spacing_info:
                align = spacing_info['alignment']
                alignment_count[align] = alignment_count.get(align, 0) + 1
        
        # 从段落统计
        for para_info in analysis_result['paragraphs_analysis']:
            actual_spacing = para_info.get('actual_spacing', {})
            
            if 'line_spacing' in actual_spacing:
                desc = actual_spacing['line_spacing']['description']
                line_spacing_count[desc] = line_spacing_count.get(desc, 0) + 1
        
        summary['line_spacing_usage'] = line_spacing_count
        summary['alignment_usage'] = alignment_count
        
        return summary
    
    def generate_spacing_adjustment_guide(self, analysis_result):
        """
        生成行距调整指南
        """
        guide = {
            'document_analysis': analysis_result['document_path'],
            'recommendations': [],
            'adjustment_methods': {},
            'common_issues': []
        }
        
        # 分析常见问题
        spacing_usage = analysis_result['spacing_summary']['line_spacing_usage']
        
        if len(spacing_usage) > 3:
            guide['common_issues'].append("文档中使用了过多不同的行距设置，建议统一行距规范")
        
        # 生成调整建议
        guide['recommendations'] = [
            "标题样式建议使用1.0-1.2倍行距，提高标题的紧凑性",
            "正文样式建议使用1.5倍行距，提高可读性",
            "段前距和段后距应保持一致，建议正文段后距6pt",
            "标题段前距建议12-18pt，段后距建议6-12pt"
        ]
        
        # 调整方法说明
        guide['adjustment_methods'] = {
            '单倍行距': '设置 paragraph_format.line_spacing = 1.0',
            '1.5倍行距': 'paragraph_format.line_spacing = 1.5',
            '双倍行距': 'paragraph_format.line_spacing = 2.0',
            '固定行距': 'paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY; paragraph_format.line_spacing = Pt(值)',
            '段前距': 'paragraph_format.space_before = Pt(值)',
            '段后距': 'paragraph_format.space_after = Pt(值)',
            '首行缩进': 'paragraph_format.first_line_indent = Pt(值)'
        }
        
        return guide
    
    def save_analysis_report(self, analysis_result, output_path=None):
        """
        保存分析报告
        """
        if output_path is None:
            output_path = config.OUTPUT_DIR + "/font_spacing_analysis_report.json"
        
        # 生成调整指南
        adjustment_guide = self.generate_spacing_adjustment_guide(analysis_result)
        
        # 合并报告
        full_report = {
            'analysis_result': analysis_result,
            'adjustment_guide': adjustment_guide,
            'generated_at': str(datetime.now())
        }
        
        try:
            with open(output_path, 'w', encoding='utf-8') as f:
                json.dump(full_report, f, ensure_ascii=False, indent=2)
            
            print(f"行距分析报告已保存到: {output_path}")
            return output_path
            
        except Exception as e:
            print(f"保存分析报告时出错: {e}")
            return None

def main():
    """
    主函数 - 分析文档的字体和行距
    """
    print("=== 字体和行距分析工具 ===")
    
    # 确保输出目录存在
    config.ensure_output_dir()
    
    analyzer = FontAndSpacingAnalyzer()
    
    # 分析格式模板
    template_path = "格式模板.docx"
    print(f"\n分析格式模板: {template_path}")
    template_analysis = analyzer.analyze_document_spacing(template_path)
    
    if template_analysis:
        print(f"模板样式数量: {template_analysis['spacing_summary']['total_styles']}")
        print(f"模板段落数量: {template_analysis['spacing_summary']['total_paragraphs']}")
        
        # 显示行距使用情况
        print("\n=== 模板行距使用情况 ===")
        for spacing_type, count in template_analysis['spacing_summary']['line_spacing_usage'].items():
            print(f"{spacing_type}: {count}次")
    
    # 分析格式化后的文档
    formatted_path = "output\\格式化后的测试文档.docx"
    print(f"\n分析格式化文档: {formatted_path}")
    formatted_analysis = analyzer.analyze_document_spacing(formatted_path)
    
    if formatted_analysis:
        print(f"格式化文档段落数量: {formatted_analysis['spacing_summary']['total_paragraphs']}")
        
        # 显示行距使用情况
        print("\n=== 格式化文档行距使用情况 ===")
        for spacing_type, count in formatted_analysis['spacing_summary']['line_spacing_usage'].items():
            print(f"{spacing_type}: {count}次")
    
    # 保存分析报告
    if template_analysis:
        template_report_path = analyzer.save_analysis_report(template_analysis, 
                                                           config.OUTPUT_DIR + "/template_spacing_analysis.json")
    
    if formatted_analysis:
        formatted_report_path = analyzer.save_analysis_report(formatted_analysis,
                                                            config.OUTPUT_DIR + "/formatted_spacing_analysis.json")
    
    print("\n=== 行距调整建议 ===")
    if formatted_analysis:
        guide = analyzer.generate_spacing_adjustment_guide(formatted_analysis)
        for recommendation in guide['recommendations']:
            print(f"• {recommendation}")
        
        if guide['common_issues']:
            print("\n=== 发现的问题 ===")
            for issue in guide['common_issues']:
                print(f"⚠ {issue}")

if __name__ == "__main__":
    main()