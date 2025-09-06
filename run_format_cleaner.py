#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Run格式清理器
清理Word文档中run级别的格式设置，让样式级别的格式能够正常生效
"""

import os
import shutil
from docx import Document
from docx.oxml.ns import qn
from config import config

class RunFormatCleaner:
    def __init__(self):
        self.cleaned_runs = 0
        self.total_runs = 0
    
    def clean_document_runs(self, input_path, output_path):
        """
        清理文档中所有run级别的格式设置
        """
        try:
            print(f"=== 清理文档run格式: {input_path} ===")
            
            # 加载文档
            doc = Document(input_path)
            
            # 重置计数器
            self.cleaned_runs = 0
            self.total_runs = 0
            
            # 遍历所有段落
            for para_idx, paragraph in enumerate(doc.paragraphs):
                if paragraph.text.strip():  # 只处理有内容的段落
                    print(f"处理段落{para_idx + 1}: {paragraph.text[:50]}...")
                    
                    # 清理段落中的所有run
                    for run_idx, run in enumerate(paragraph.runs):
                        if run.text.strip():  # 只处理有内容的run
                            self.total_runs += 1
                            if self._clean_run_format(run):
                                self.cleaned_runs += 1
            
            # 保存清理后的文档
            doc.save(output_path)
            
            print(f"\n=== 清理完成 ===")
            print(f"总run数: {self.total_runs}")
            print(f"清理的run数: {self.cleaned_runs}")
            print(f"清理后的文档已保存到: {output_path}")
            
            return True
            
        except Exception as e:
            print(f"清理文档run格式时出错: {e}")
            return False
    
    def _clean_run_format(self, run):
        """
        清理单个run的格式设置
        """
        cleaned = False
        
        try:
            # 清理基本字体格式
            if hasattr(run, 'font'):
                font = run.font
                
                # 清理字体名称
                if font.name is not None:
                    font.name = None
                    cleaned = True
                
                # 清理字体大小
                if font.size is not None:
                    font.size = None
                    cleaned = True
                
                # 清理粗体设置
                if font.bold is not None:
                    font.bold = None
                    cleaned = True
                
                # 清理斜体设置
                if font.italic is not None:
                    font.italic = None
                    cleaned = True
                
                # 清理下划线设置
                if font.underline is not None:
                    font.underline = None
                    cleaned = True
            
            # 清理XML级别的字体设置
            if hasattr(run, '_element'):
                run_element = run._element
                rpr = run_element.find(qn('w:rPr'))
                if rpr is not None:
                    # 查找并移除rFonts元素
                    rfonts = rpr.find(qn('w:rFonts'))
                    if rfonts is not None:
                        rpr.remove(rfonts)
                        cleaned = True
                    
                    # 查找并移除字体大小设置
                    sz = rpr.find(qn('w:sz'))
                    if sz is not None:
                        rpr.remove(sz)
                        cleaned = True
                    
                    # 查找并移除复杂字体大小设置
                    szcs = rpr.find(qn('w:szCs'))
                    if szcs is not None:
                        rpr.remove(szcs)
                        cleaned = True
                    
                    # 查找并移除粗体设置
                    b = rpr.find(qn('w:b'))
                    if b is not None:
                        rpr.remove(b)
                        cleaned = True
                    
                    # 查找并移除复杂粗体设置
                    bcs = rpr.find(qn('w:bCs'))
                    if bcs is not None:
                        rpr.remove(bcs)
                        cleaned = True
                    
                    # 查找并移除斜体设置
                    i = rpr.find(qn('w:i'))
                    if i is not None:
                        rpr.remove(i)
                        cleaned = True
                    
                    # 查找并移除复杂斜体设置
                    ics = rpr.find(qn('w:iCs'))
                    if ics is not None:
                        rpr.remove(ics)
                        cleaned = True
                    
                    # 如果rPr元素为空，则移除它
                    if len(rpr) == 0:
                        run_element.remove(rpr)
        
        except Exception as e:
            print(f"清理run格式时出错: {e}")
        
        return cleaned
    
    def create_clean_test_document(self):
        """
        创建清理后的测试文档
        """
        # 验证测试文档是否存在
        if not os.path.exists(config.TEST_DOCUMENT):
            print(f"错误：找不到测试文档 {config.TEST_DOCUMENT}")
            return False
        
        # 确保输出目录存在
        config.ensure_output_dir()
        
        # 生成清理后的文档路径
        clean_doc_path = os.path.join(config.OUTPUT_DIR, "测试文档_清理后.docx")
        
        # 清理文档
        success = self.clean_document_runs(config.TEST_DOCUMENT, clean_doc_path)
        
        if success:
            print(f"\n清理后的测试文档已创建: {clean_doc_path}")
            return clean_doc_path
        else:
            return None

def main():
    """
    主函数
    """
    cleaner = RunFormatCleaner()
    
    # 创建清理后的测试文档
    clean_doc_path = cleaner.create_clean_test_document()
    
    if clean_doc_path:
        print("\n建议：")
        print("1. 使用清理后的测试文档重新运行格式应用")
        print("2. 或者将清理后的文档替换原始测试文档")
        print("3. 然后重新运行 dynamic_format_applier.py")

if __name__ == "__main__":
    main()