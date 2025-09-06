from docx import Document
from docx.oxml.ns import qn
import xml.etree.ElementTree as ET

# 加载格式模板
doc = Document('格式模板.docx')
normal_style = doc.styles['Normal']

print('文档结构检查:')
print(f'文档部件数量: {len(doc.part.related_parts)}')
print(f'样式数量: {len(doc.styles)}')

# 检查文档的XML结构
print('\n文档XML结构:')
document_element = doc._element
print(f'文档根元素: {document_element.tag}')
for child in document_element:
    print(f'  子元素: {child.tag}')

print('Normal样式字体信息:')
print(f'font.name: {normal_style.font.name}')
print(f'font.size: {normal_style.font.size}')

# 检查XML级别的字体分离设置
rpr = normal_style._element.find(qn('w:rPr'))
rfonts = rpr.find(qn('w:rFonts')) if rpr is not None else None

print(f'rFonts存在: {rfonts is not None}')

if rfonts is not None:
    print(f'ascii: {rfonts.get(qn("w:ascii"))}')
    print(f'hAnsi: {rfonts.get(qn("w:hAnsi"))}')
    print(f'eastAsia: {rfonts.get(qn("w:eastAsia"))}')
    print(f'cs: {rfonts.get(qn("w:cs"))}')
else:
    print('Normal样式没有rFonts设置')

# 检查样式部件中的默认字体设置
print('\n通过样式部件检查默认字体设置:')
# 找到StylesPart
styles_part = None
for rel_id, part in doc.part.related_parts.items():
    if 'StylesPart' in str(type(part)):
        styles_part = part
        print(f'找到样式部件: {rel_id}')
        break

if styles_part:
    styles_element = styles_part.element
    print(f'样式元素: {styles_element.tag}')
    
    doc_defaults = styles_element.find(qn('w:docDefaults'))
    print(f'docDefaults存在: {doc_defaults is not None}')
    
    if doc_defaults is not None:
        rpr_default = doc_defaults.find(qn('w:rPrDefault'))
        print(f'rPrDefault存在: {rpr_default is not None}')
        
        if rpr_default is not None:
            rpr = rpr_default.find(qn('w:rPr'))
            print(f'rPr存在: {rpr is not None}')
            
            if rpr is not None:
                rfonts = rpr.find(qn('w:rFonts'))
                print(f'默认rFonts存在: {rfonts is not None}')
                
                if rfonts is not None:
                    print(f'默认ascii: {rfonts.get(qn("w:ascii"))}')
                    print(f'默认hAnsi: {rfonts.get(qn("w:hAnsi"))}')
                    print(f'默认eastAsia: {rfonts.get(qn("w:eastAsia"))}')
                    print(f'默认cs: {rfonts.get(qn("w:cs"))}')
                    
                # 检查默认字号
                sz = rpr.find(qn('w:sz'))
                if sz is not None:
                    print(f'默认字号: {sz.get(qn("w:val"))}')
else:
    print('未找到样式部件')