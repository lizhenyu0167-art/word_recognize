#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
动态格式提取器
功能：完整提取格式模板中各标题及正文的排版格式信息
包括中英文字体分离、字号、行间距、段落间距等所有格式属性
以及页眉页脚格式
"""

import os
import json
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml.ns import qn
from datetime import datetime
from config import config

class DynamicFormatExtractor:
    def __init__(self, template_path=None):
        self.template_path = template_path or config.TEMPLATE_FILE
        self.format_info = {
            'extraction_time': None,
            'template_file': self.template_path,
            'document_defaults': {},
            'styles': {},
            'section_settings': {}
        }
    
    def extract_template_formats(self, template_path=None):
        """
        动态提取格式模板中的所有格式信息
        """
        if template_path is None:
            template_path = self.template_path
            
        print(f"正在动态提取格式模板: {template_path}")
        
        self.format_info['extraction_time'] = datetime.now().isoformat()
        self.format_info['template_file'] = os.path.basename(template_path)
        self.format_info['headers'] = {}
        self.format_info['footers'] = {}
        
        try:
            doc = Document(template_path)
            
            # 1. 提取文档默认设置
            self._extract_document_defaults(doc)
            
            # 2. 提取所有段落样式的完整格式信息
            print("\n=== 提取样式格式信息 ===")
            for style in doc.styles:
                if style.type == WD_STYLE_TYPE.PARAGRAPH:
                    style_info = self._extract_complete_style_info(style)
                    self.format_info['styles'][style.name] = style_info
                    print(f"提取样式: {style.name}")
                    
                    # 显示基本字体属性
                    font_attrs = []
                    if 'font_size' in style_info:
                        font_attrs.append(f"字号={style_info['font_size']}")
                    if 'bold' in style_info:
                        font_attrs.append(f"加粗={style_info['bold']}")
                    if 'italic' in style_info:
                        font_attrs.append(f"斜体={style_info['italic']}")
                    if 'color' in style_info:
                        font_attrs.append(f"颜色={style_info['color']}")
                    if 'underline' in style_info:
                        font_attrs.append(f"下划线={style_info['underline']}")
                    
                    if font_attrs:
                        print(f"  字体属性: {', '.join(font_attrs)}")
                    
                    # 显示字体分离信息
                    if 'font_separation' in style_info:
                        font_size = style_info.get('font_size', '未设置')
                        # 优先显示ascii字体，如果未设置则显示hAnsi字体
                        english_font = style_info['font_separation'].get('ascii') or style_info['font_separation'].get('hAnsi', '未设置')
                        print(f"  字体分离: 英文={english_font}, 中文={style_info['font_separation'].get('eastAsia', '未设置')}, 字号={font_size}")
            
            # 3. 提取页眉页脚格式信息
            print("\n=== 提取页眉页脚格式信息 ===")
            self._extract_header_footer_formats(doc)
            
            # 4. 保存格式信息到文件
            self._save_format_info()
            
            print(f"\n格式信息提取完成！")
            print(f"共提取 {len(self.format_info['styles'])} 个样式，页眉 {len(self.format_info['headers'])} 个，页脚 {len(self.format_info['footers'])} 个")
            
            return self.format_info
            
        except Exception as e:
            print(f"提取格式信息时出错: {e}")
            return None
    
    def _extract_document_defaults(self, doc):
        """
        提取文档默认设置，包括XML级别的默认字体分离设置
        """
        print("提取文档默认设置...")
        
        try:
            # 1. 先尝试从样式部件中提取文档默认设置
            styles_part = None
            for rel_id, part in doc.part.related_parts.items():
                if 'StylesPart' in str(type(part)):
                    styles_part = part
                    break
            
            if styles_part:
                styles_element = styles_part.element
                doc_defaults = styles_element.find(qn('w:docDefaults'))
                
                if doc_defaults is not None:
                    rpr_default = doc_defaults.find(qn('w:rPrDefault'))
                    if rpr_default is not None:
                        rpr = rpr_default.find(qn('w:rPr'))
                        if rpr is not None:
                            # 提取默认字体分离设置
                            rfonts = rpr.find(qn('w:rFonts'))
                            if rfonts is not None:
                                eastAsia_font = rfonts.get(qn('w:eastAsia'))
                                if eastAsia_font:
                                    self.format_info['document_defaults']['default_font'] = eastAsia_font
                                    print(f"  文档默认中文字体: {eastAsia_font}")
                                else:
                                    ascii_font = rfonts.get(qn('w:ascii'))
                                    if ascii_font:
                                        self.format_info['document_defaults']['default_font'] = ascii_font
                                        print(f"  文档默认英文字体: {ascii_font}")
                            
                            # 提取默认字号
                            sz = rpr.find(qn('w:sz'))
                            if sz is not None:
                                font_size = sz.get(qn('w:val'))
                                if font_size:
                                    # Word中字号是半点单位，需要除以2
                                    self.format_info['document_defaults']['default_font_size'] = str(int(font_size) / 2) + 'pt'
                                    print(f"  文档默认字号: {int(font_size) / 2}pt")
            
            # 2. 如果没有从XML中提取到，则使用Normal样式作为备选
            if 'default_font' not in self.format_info['document_defaults']:
                normal_style = None
                for style in doc.styles:
                    if style.name == 'Normal':
                        normal_style = style
                        break
                
                if normal_style and hasattr(normal_style, 'font'):
                    font = normal_style.font
                    if font.name:
                        self.format_info['document_defaults']['default_font'] = font.name
                        print(f"  从Normal样式获取默认字体: {font.name}")
                    
                    if font.size and 'default_font_size' not in self.format_info['document_defaults']:
                        self.format_info['document_defaults']['default_font_size'] = str(font.size.pt) + 'pt'
                        print(f"  从Normal样式获取默认字号: {font.size.pt}pt")
            
            # 3. 最后的备选方案
            if 'default_font' not in self.format_info['document_defaults']:
                self.format_info['document_defaults']['default_font'] = '宋体'
                print(f"  使用备选默认字体: 宋体")
            
            if 'default_font_size' not in self.format_info['document_defaults']:
                self.format_info['document_defaults']['default_font_size'] = '10.0pt'
                print(f"  使用备选默认字号: 10.0pt")
                
        except Exception as e:
            print(f"提取文档默认设置时出错: {e}")
            # 设置备选值
            self.format_info['document_defaults']['default_font'] = '宋体'
            self.format_info['document_defaults']['default_font_size'] = '10.0pt'
    
    def _extract_complete_style_info(self, style):
        """
        提取样式的完整格式信息
        """
        style_info = {
            'name': style.name,
            'type': 'paragraph'
        }
        
        try:
            # 1. 提取字体信息（包括字体分离）
            font_info = self._extract_font_info(style)
            if font_info:
                style_info.update(font_info)
            
            # 2. 提取段落格式信息
            paragraph_info = self._extract_paragraph_info(style)
            if paragraph_info:
                style_info.update(paragraph_info)
            
            return style_info
            
        except Exception as e:
            print(f"提取样式 {style.name} 信息时出错: {e}")
            return style_info
    
    def _extract_font_info(self, style):
        """
        提取字体信息，包括完整的字体分离设置
        """
        font_info = {}
        
        try:
            if hasattr(style, 'font'):
                font = style.font
                
                # 基本字体信息
                if font.name is not None:
                    font_info['font_name'] = font.name
                
                if font.size is not None:
                    font_info['font_size'] = str(font.size.pt) + 'pt'
                else:
                    # 当字号未设置时，尝试从继承链条或使用默认值
                    inherited_size = self._get_inherited_font_size(style)
                    if inherited_size:
                        font_info['font_size'] = inherited_size
                
                # 加粗属性 - 总是包含，确保能覆盖原文档设置
                font_info['bold'] = font.bold if font.bold is not None else False
                
                # 斜体属性 - 总是包含，确保能覆盖原文档设置
                font_info['italic'] = font.italic if font.italic is not None else False
                
                if font.underline is not None:
                    font_info['underline'] = str(font.underline)
                
                # 添加字体颜色提取
                if font.color is not None and font.color.rgb is not None:
                    font_info['color'] = str(font.color.rgb)
                
                # 提取字体分离设置（XML级别）
                font_separation = self._extract_font_separation(style)
                if font_separation:
                    font_info['font_separation'] = font_separation
                elif style.name == 'Normal':
                    # 对于Normal样式，如果没有明确的字体分离设置，使用文档默认字体
                    default_font = self.format_info['document_defaults'].get('default_font', '宋体')
                    font_info['font_separation'] = {
                        'eastAsia': default_font
                    }
                    print(f"  Normal样式使用文档默认中文字体: {default_font}")
            
            return font_info
            
        except Exception as e:
            print(f"提取字体信息时出错: {e}")
            return {}
    
    def _get_inherited_font_size(self, style):
        """
        获取继承的字号，如果整个继承链条都没有设置，则使用Word默认值12pt
        """
        try:
            # 遍历继承链条查找字号设置
            current_style = style
            level = 0
            
            while current_style and level < 10:  # 防止无限循环
                # 检查当前样式的字号
                if hasattr(current_style, 'font') and current_style.font.size:
                    return str(current_style.font.size.pt) + 'pt'
                
                # 检查XML级别的字号设置
                if hasattr(current_style, '_element'):
                    style_element = current_style._element
                    rpr = style_element.find(qn('w:rPr'))
                    if rpr is not None:
                        sz = rpr.find(qn('w:sz'))
                        if sz is not None:
                            sz_val = sz.get(qn('w:val'))
                            if sz_val:
                                return str(int(sz_val)/2) + 'pt'
                
                # 移动到基础样式
                if hasattr(current_style, 'base_style') and current_style.base_style:
                    current_style = current_style.base_style
                    level += 1
                else:
                    break
            
            # 检查文档默认字号
            try:
                if hasattr(self, 'format_info') and 'document_defaults' in self.format_info:
                    default_size = self.format_info['document_defaults'].get('default_font_size')
                    if default_size:
                        return default_size
            except:
                pass
            
            # 使用Word默认字号12pt
            return '12.0pt'
            
        except Exception as e:
            print(f"获取继承字号时出错: {e}")
            return '12.0pt'  # 出错时返回默认值
    
    def _extract_font_separation(self, style):
        """
        提取XML级别的字体分离设置，包括继承的字体和主题字体
        """
        try:
            separation = {}
            theme_fonts = {}
            
            if hasattr(style, '_element'):
                style_element = style._element
                
                # 查找rPr元素
                rpr = style_element.find(qn('w:rPr'))
                if rpr is not None:
                    # 查找rFonts元素
                    rfonts = rpr.find(qn('w:rFonts'))
                    if rfonts is not None:
                        # 提取直接字体设置
                        ascii_font = rfonts.get(qn('w:ascii'))
                        if ascii_font:
                            separation['ascii'] = ascii_font
                        
                        hAnsi_font = rfonts.get(qn('w:hAnsi'))
                        if hAnsi_font:
                            separation['hAnsi'] = hAnsi_font
                        
                        eastAsia_font = rfonts.get(qn('w:eastAsia'))
                        if eastAsia_font:
                            separation['eastAsia'] = eastAsia_font
                        
                        cs_font = rfonts.get(qn('w:cs'))
                        if cs_font:
                            separation['cs'] = cs_font
                        
                        # 提取主题字体设置
                        ascii_theme = rfonts.get(qn('w:asciiTheme'))
                        if ascii_theme:
                            theme_fonts['ascii'] = ascii_theme
                        
                        hAnsi_theme = rfonts.get(qn('w:hAnsiTheme'))
                        if hAnsi_theme:
                            theme_fonts['hAnsi'] = hAnsi_theme
                        
                        eastAsia_theme = rfonts.get(qn('w:eastAsiaTheme'))
                        if eastAsia_theme:
                            theme_fonts['eastAsia'] = eastAsia_theme
                        
                        cs_theme = rfonts.get(qn('w:cstheme'))
                        if cs_theme:
                            theme_fonts['cs'] = cs_theme
            
            # 解析主题字体为实际字体名称
            theme_font_map = {
                'majorHAnsi': 'Calibri Light',
                'minorHAnsi': 'Calibri',
                'majorBidi': 'Times New Roman',
                'minorBidi': 'Arial'
            }
            
            for font_type, theme_name in theme_fonts.items():
                if font_type not in separation:  # 只有在没有直接设置时才使用主题字体
                    actual_font = theme_font_map.get(theme_name, theme_name)
                    separation[font_type] = actual_font
            
            # 如果当前样式没有设置某些字体，尝试从基础样式继承
            if hasattr(style, 'base_style') and style.base_style:
                base_separation = self._extract_font_separation(style.base_style)
                if base_separation:
                    # 只继承当前样式没有设置的字体
                    for font_type in ['ascii', 'hAnsi', 'eastAsia', 'cs']:
                        if font_type not in separation and font_type in base_separation:
                            separation[font_type] = base_separation[font_type]
            
            # 如果仍然没有设置，且不是Normal样式，则检查是否需要使用文档默认字体
            if style.name != 'Normal' and separation:
                default_font = self.format_info['document_defaults'].get('default_font', '宋体')
                
                # 只对eastAsia字体使用默认字体
                if 'eastAsia' not in separation:
                    separation['eastAsia'] = default_font
            
            # 处理cs字体作为英文字体的情况
            # 当cs字体设置了常见英文字体时，将其作为英文字体（优先级高于继承的字体）
            if separation and 'cs' in separation:
                cs_font = separation['cs']
                # 如果cs字体是常见的英文字体，则将其作为英文字体
                common_english_fonts = ['Times New Roman', 'Arial', 'Calibri', 'Calibri Light', 'Verdana', 'Tahoma']
                if cs_font in common_english_fonts:
                    # 如果ascii和hAnsi都未直接设置（即来自继承），则用cs字体覆盖
                    # 检查是否是直接设置的字体（通过重新解析当前样式的rFonts）
                    direct_fonts = self._get_direct_fonts(style)
                    if 'ascii' not in direct_fonts and 'hAnsi' not in direct_fonts:
                        separation['ascii'] = cs_font
                        separation['hAnsi'] = cs_font
            
            return separation if separation else None
            
        except Exception as e:
            print(f"提取字体分离设置时出错: {e}")
            return None
    
    def _get_direct_fonts(self, style):
        """
        获取样式中直接设置的字体（不包括继承的字体）
        """
        direct_fonts = {}
        try:
            if hasattr(style, '_element'):
                style_element = style._element
                rpr = style_element.find(qn('w:rPr'))
                if rpr is not None:
                    rfonts = rpr.find(qn('w:rFonts'))
                    if rfonts is not None:
                        ascii_font = rfonts.get(qn('w:ascii'))
                        if ascii_font:
                            direct_fonts['ascii'] = ascii_font
                        
                        hAnsi_font = rfonts.get(qn('w:hAnsi'))
                        if hAnsi_font:
                            direct_fonts['hAnsi'] = hAnsi_font
                        
                        eastAsia_font = rfonts.get(qn('w:eastAsia'))
                        if eastAsia_font:
                            direct_fonts['eastAsia'] = eastAsia_font
                        
                        cs_font = rfonts.get(qn('w:cs'))
                        if cs_font:
                            direct_fonts['cs'] = cs_font
            
            return direct_fonts
        except Exception as e:
            print(f"获取直接字体设置时出错: {e}")
            return {}
    
    def _extract_paragraph_info(self, style):
        """
        提取段落格式信息
        """
        paragraph_info = {}
        
        try:
            if hasattr(style, 'paragraph_format'):
                pf = style.paragraph_format
                
                # 对齐方式
                if pf.alignment is not None:
                    alignment_map = {
                        WD_ALIGN_PARAGRAPH.LEFT: '左对齐',
                        WD_ALIGN_PARAGRAPH.CENTER: '居中',
                        WD_ALIGN_PARAGRAPH.RIGHT: '右对齐',
                        WD_ALIGN_PARAGRAPH.JUSTIFY: '两端对齐',
                        WD_ALIGN_PARAGRAPH.DISTRIBUTE: '分散对齐'
                    }
                    paragraph_info['alignment'] = alignment_map.get(pf.alignment, '未知对齐')
                
                # 行间距
                if pf.line_spacing is not None:
                    paragraph_info['line_spacing'] = str(pf.line_spacing)
                
                # 段前距
                if pf.space_before is not None:
                    paragraph_info['space_before'] = str(pf.space_before.pt) + 'pt'
                
                # 段后距
                if pf.space_after is not None:
                    paragraph_info['space_after'] = str(pf.space_after.pt) + 'pt'
                
                # 首行缩进
                if pf.first_line_indent is not None:
                    paragraph_info['first_line_indent'] = str(pf.first_line_indent.pt) + 'pt'
                
                # 左缩进
                if pf.left_indent is not None:
                    paragraph_info['left_indent'] = str(pf.left_indent.pt) + 'pt'
                
                # 右缩进
                if pf.right_indent is not None:
                    paragraph_info['right_indent'] = str(pf.right_indent.pt) + 'pt'
            
            return paragraph_info
            
        except Exception as e:
            print(f"提取段落格式信息时出错: {e}")
            return {}
    
    def _save_format_info(self, output_path=None):
        """
        保存格式信息到JSON文件
        """
        try:
            if output_path is None:
                output_path = config.DYNAMIC_FORMAT_INFO
            
            # 确保输出目录存在
            config.ensure_output_dir()
            
            # 保存到JSON文件
            with open(output_path, 'w', encoding='utf-8') as f:
                json.dump(self.format_info, f, ensure_ascii=False, indent=2)
            
            print(f"格式信息已保存到: {output_path}")
            
        except Exception as e:
            print(f"保存格式信息时出错: {e}")
            
    def _extract_header_footer_formats(self, doc):
        """
        提取页眉页脚的格式信息
        """
        try:
            # 遍历文档的节
            for i, section in enumerate(doc.sections):
                section_id = f"section_{i+1}"
                
                # 提取页眉顶端距离和页脚底端距离
                if hasattr(section, 'header_distance') and section.header_distance:
                    header_distance_pt = section.header_distance.pt
                    self.format_info['section_settings'][section_id] = self.format_info['section_settings'].get(section_id, {})
                    self.format_info['section_settings'][section_id]['header_distance'] = f"{header_distance_pt}pt"
                    print(f"提取页眉顶端距离: 第{i+1}节 - {header_distance_pt}pt")
                
                if hasattr(section, 'footer_distance') and section.footer_distance:
                    footer_distance_pt = section.footer_distance.pt
                    self.format_info['section_settings'][section_id] = self.format_info['section_settings'].get(section_id, {})
                    self.format_info['section_settings'][section_id]['footer_distance'] = f"{footer_distance_pt}pt"
                    print(f"提取页脚底端距离: 第{i+1}节 - {footer_distance_pt}pt")
                
                # 提取页眉格式
                if section.header.is_linked_to_previous == False:
                    header_info = self._extract_header_footer_content(section.header)
                    if header_info:
                        self.format_info['headers'][section_id] = header_info
                        print(f"提取页眉格式: 第{i+1}节")
                
                # 提取页脚格式
                if section.footer.is_linked_to_previous == False:
                    footer_info = self._extract_header_footer_content(section.footer)
                    if footer_info:
                        self.format_info['footers'][section_id] = footer_info
                        print(f"提取页脚格式: 第{i+1}节")
                        
        except Exception as e:
            print(f"提取页眉页脚格式时出错: {e}")
    
    def _extract_header_footer_content(self, header_footer):
        """
        提取页眉或页脚的内容和格式
        """
        try:
            content_info = {
                'paragraphs': []
            }
            
            for para in header_footer.paragraphs:
                para_info = {
                    'text': para.text,
                    'alignment': str(para.alignment) if para.alignment else 'LEFT',
                    'runs': []
                }
                
                # 提取每个文本块的格式
                for run in para.runs:
                    run_info = {
                        'text': run.text
                    }
                    
                    # 提取字体信息
                    if run.font:
                        if run.font.name:
                            run_info['font_name'] = run.font.name
                        if run.font.size:
                            run_info['font_size'] = str(run.font.size.pt) + 'pt'
                        if run.font.bold is not None:
                            run_info['bold'] = run.font.bold
                        if run.font.italic is not None:
                            run_info['italic'] = run.font.italic
                        if run.font.underline is not None:
                            run_info['underline'] = str(run.font.underline)
                        if run.font.color and run.font.color.rgb:
                            run_info['color'] = str(run.font.color.rgb)
                    
                    para_info['runs'].append(run_info)
                
                content_info['paragraphs'].append(para_info)
            
            return content_info
            
        except Exception as e:
            print(f"提取页眉页脚内容时出错: {e}")
            return None
    
    def load_format_info(self, format_file=None):
        """
        加载已保存的格式信息
        """
        if format_file is None:
            format_file = config.DYNAMIC_FORMAT_INFO
        
        try:
            if os.path.exists(format_file):
                with open(format_file, 'r', encoding='utf-8') as f:
                    self.format_info = json.load(f)
                print(f"已加载格式信息: {format_file}")
                return self.format_info
            else:
                print(f"格式信息文件不存在: {format_file}")
                return None
                
        except Exception as e:
            print(f"加载格式信息时出错: {e}")
            return None

def main():
    """
    主函数：提取格式模板的格式信息
    """
    # 验证必需文件
    missing_files = config.validate_required_files()
    if missing_files:
        print(f"错误：找不到以下必需文件: {', '.join(missing_files)}")
        return
    
    extractor = DynamicFormatExtractor()
    
    # 提取格式模板的格式信息
    format_info = extractor.extract_template_formats()
    if format_info:
        print("\n=== 格式信息提取摘要 ===")
        print(f"提取时间: {format_info['extraction_time']}")
        print(f"模板文件: {format_info['template_file']}")
        print(f"文档默认字体: {format_info['document_defaults'].get('default_font', '未设置')}")
        print(f"样式数量: {len(format_info['styles'])}")

if __name__ == "__main__":
    main()