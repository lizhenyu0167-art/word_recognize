#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
检查文档中页眉顶端距离和页脚底端距离设置
"""

import os
from docx import Document
from docx.shared import Pt
from config import config

def check_header_footer_distance(doc_path):
    """
    检查文档中页眉顶端距离和页脚底端距离设置
    """
    print(f"\n检查文档: {doc_path}")
    
    try:
        doc = Document(doc_path)
        
        # 遍历文档的节
        for i, section in enumerate(doc.sections):
            print(f"\n第{i+1}节:")
            
            # 检查页眉顶端距离
            if hasattr(section, 'header_distance') and section.header_distance:
                header_distance_pt = section.header_distance.pt
                print(f"  页眉顶端距离: {header_distance_pt}pt")
            else:
                print("  页眉顶端距离: 未设置")
            
            # 检查页脚底端距离
            if hasattr(section, 'footer_distance') and section.footer_distance:
                footer_distance_pt = section.footer_distance.pt
                print(f"  页脚底端距离: {footer_distance_pt}pt")
            else:
                print("  页脚底端距离: 未设置")
                
    except Exception as e:
        print(f"检查文档时出错: {e}")

def main():
    # 检查格式化后的文档
    formatted_doc_path = os.path.join(config.OUTPUT_DIR, "格式化后的测试文档.docx")
    if os.path.exists(formatted_doc_path):
        check_header_footer_distance(formatted_doc_path)
    else:
        print(f"文档不存在: {formatted_doc_path}")
    
    # 检查格式模板文档
    template_path = config.TEMPLATE_FILE
    if os.path.exists(template_path):
        check_header_footer_distance(template_path)
    else:
        print(f"文档不存在: {template_path}")

if __name__ == "__main__":
    main()