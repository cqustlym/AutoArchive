from docx import Document
from docx.oxml.ns import qn
import docx
from docx.oxml.ns import qn
from docx.shared import Pt,RGBColor
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT

def modify_heading1_style(documentPath):
# 打开一个已存在的Word文档
    doc = Document(documentPath)
    # 遍历文档中的所有段落
    for paragraph in doc.paragraphs:
        # 访问段落的XML结构
        pPr = paragraph._element.get_or_add_pPr()
        outline_lvl = pPr.find(qn('w:outlineLvl'))
        # 检查大纲级别是否为1
        if outline_lvl is not None and outline_lvl.get(qn('w:val')) == '0':
            # 将段落样式设置为'Heading 1'
            paragraph.style = doc.styles['Heading 1']
            if paragraph.style.name.startswith('Heading 1'): # 检查当前段落的样式是否以 "Heading 1" 开头
                # 修改一级标题字体样式
                paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
                # 设置段前间距为0.5行（大约6磅）
                paragraph.paragraph_format.space_before = Pt(6)
                # 设置段后间距为0.5行（大约6磅）
                paragraph.paragraph_format.space_after = Pt(6)
                run = paragraph.runs
                for i in run:
                    i.font.bold = False
                    i.font.size = Pt(18)  # 设置为小二号字
                    i.font.color.rgb = RGBColor(0,0,0) # 设置字体颜色为黑色
                    i.font.name = '黑体'
                    # 由于python-docx处理中文时可能会有问题，因此需要额外设置字体的中文部分
                    i._element.rPr.rFonts.set(qn('w:eastAsia'), '黑体')
                    # 修改段落居中

    # 保存修改后的文档
    modified_doc_path = 'modified_document.docx'
    doc.save(modified_doc_path)
    print(f"所有一级标题已修改完成，已保存到 {modified_doc_path}!")

modify_heading1_style('test.docx')




