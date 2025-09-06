#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
检查Word文档中英文字体的实际设置
"""

from docx import Document
from docx.oxml.ns import qn
import xml.etree.ElementTree as ET

def check_english_fonts():
    """
    检查Word文档中英文字体的设置情况
    """
    print("=== 检查Word文档中的英文字体设置 ===")
    
    # 加载格式模板文档
    doc = Document('格式模板.docx')
    
    # 检查文档级别的字体设置
    print("\n=== 文档级别字体设置 ===")
    try:
        # 检查文档默认设置
        settings = doc.settings
        if hasattr(settings, 'element'):
            doc_defaults = settings.element.find(qn('w:docDefaults'))
            if doc_defaults is not None:
                run_defaults = doc_defaults.find(qn('w:rPrDefault'))
                if run_defaults is not None:
                    rpr = run_defaults.find(qn('w:rPr'))
                    if rpr is not None:
                        rfonts = rpr.find(qn('w:rFonts'))
                        if rfonts is not None:
                            print("文档默认字体:")
                            for attr in ['ascii', 'hAnsi', 'eastAsia', 'cs']:
                                font_val = rfonts.get(qn(f'w:{attr}'))
                                if font_val:
                                    print(f"  {attr}: {font_val}")
                            
                            # 检查主题字体
                            for attr in ['asciiTheme', 'hAnsiTheme', 'eastAsiaTheme', 'cstheme']:
                                theme_val = rfonts.get(qn(f'w:{attr}'))
                                if theme_val:
                                    print(f"  {attr}: {theme_val}")
                        else:
                            print("无文档默认字体设置")
                    else:
                        print("无文档默认运行属性")
                else:
                    print("无运行属性默认值")
            else:
                print("无文档默认设置")
    except Exception as e:
        print(f"检查文档默认设置时出错: {e}")
    
    # 检查主题文件
    print("\n=== 主题字体设置 ===")
    try:
        # 尝试获取主题部分
        theme_part = None
        for rel in doc.part.rels.values():
            if 'theme' in rel.reltype:
                theme_part = rel.target_part
                break
        
        if theme_part:
            print("找到主题文件，解析主题字体...")
            theme_xml = theme_part.blob
            root = ET.fromstring(theme_xml)
            
            # 查找字体方案
            ns = {'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'}
            font_scheme = root.find('.//a:fontScheme', ns)
            if font_scheme is not None:
                # 主要字体（通常用于标题）
                major_font = font_scheme.find('a:majorFont', ns)
                if major_font is not None:
                    latin = major_font.find('a:latin', ns)
                    if latin is not None:
                        print(f"主要拉丁字体: {latin.get('typeface')}")
                    ea = major_font.find('a:ea', ns)
                    if ea is not None:
                        print(f"主要东亚字体: {ea.get('typeface')}")
                
                # 次要字体（通常用于正文）
                minor_font = font_scheme.find('a:minorFont', ns)
                if minor_font is not None:
                    latin = minor_font.find('a:latin', ns)
                    if latin is not None:
                        print(f"次要拉丁字体: {latin.get('typeface')}")
                    ea = minor_font.find('a:ea', ns)
                    if ea is not None:
                        print(f"次要东亚字体: {ea.get('typeface')}")
            else:
                print("未找到字体方案")
        else:
            print("未找到主题文件")
    except Exception as e:
        print(f"检查主题字体时出错: {e}")
    
    # 检查特定样式的字体设置
    print("\n=== 样式字体详细分析 ===")
    styles_to_check = ['Normal', 'Heading 1', 'Heading 3', 'style_1']
    
    for style_name in styles_to_check:
        style = None
        for s in doc.styles:
            if s.name == style_name:
                style = s
                break
        
        if style:
            print(f"\n--- {style_name} ---")
            if hasattr(style, '_element'):
                style_element = style._element
                rpr = style_element.find(qn('w:rPr'))
                if rpr is not None:
                    rfonts = rpr.find(qn('w:rFonts'))
                    if rfonts is not None:
                        print("直接字体设置:")
                        for attr in ['ascii', 'hAnsi', 'eastAsia', 'cs']:
                            font_val = rfonts.get(qn(f'w:{attr}'))
                            if font_val:
                                print(f"  {attr}: {font_val}")
                        
                        print("主题字体设置:")
                        for attr in ['asciiTheme', 'hAnsiTheme', 'eastAsiaTheme', 'cstheme']:
                            theme_val = rfonts.get(qn(f'w:{attr}'))
                            if theme_val:
                                print(f"  {attr}: {theme_val}")
                    else:
                        print("无字体设置")
                else:
                    print("无运行属性")
            else:
                print("无样式元素")
        else:
            print(f"未找到样式: {style_name}")

if __name__ == "__main__":
    check_english_fonts()