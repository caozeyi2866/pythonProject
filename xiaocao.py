#快速排版工具
from docx import Document
from docx.shared import RGBColor,Pt,Cm,Inches
from docx.oxml.ns import qn
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
document = Document("花园村人大代表联络站简介.docx")
#页面边距设置
sections = document.sections
for section in sections:
    section.top_margin = Cm(3.7)
    section.bottom_margin = Cm(3.5)
    section.left_margin = Cm(2.8)
    section.right_margin = Cm(2.6)

#读取段落
all_paragraphs = document.paragraphs
#标题设置为方正小标宋简体，二号
all_paragraphs[0].alignment=WD_PARAGRAPH_ALIGNMENT.CENTER#设置为居中
all_paragraphs[0].paragraph_format.space_before=Pt(0)#设置段前 0 磅
all_paragraphs[0].paragraph_format.space_after=Pt(0) #设置段后 0 磅
all_paragraphs[0].paragraph_format.line_spacing=Pt(0) #设置行间距为 1.5
all_paragraphs[0].paragraph_format.left_indent=Inches(0)#设置左缩进
all_paragraphs[0].paragraph_format.right_indent=Inches(0)#设置右缩进
for run in all_paragraphs[0].runs:
    run.font.name=u'方正小标宋简体"'    #设置为宋体
    run._element.rPr.rFonts.set(qn('w:eastAsia'), u'方正小标宋简体"')#设置为宋体，和上边的一起使用
    run.font.size=Pt(22)#设置1级标题文字的大小
    run.font.color.rgb=RGBColor(0,0,0)#设置颜色为黑色
# 删除空白段落
for paragraph in all_paragraphs:  # 读取文档段落
    if len(paragraph.text) == 0 and len(paragraph.runs) <= 1:
        p = paragraph._element
        p.getparent().remove(p)
        p._p = p._element = None
#设置正文字体的样式，仿宋GB_2312，三号
for paragraph in all_paragraphs[1:]:
    paragraph.paragraph_format.line_spacing = Pt(25) #行间距
    paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY  # 设置为两端对齐
    paragraph.paragraph_format.space_before = Pt(0)  # 设置段前 0 磅
    paragraph.paragraph_format.space_after = Pt(0)  # 设置段后 0 磅
    paragraph.paragraph_format.left_indent = Inches(0)  # 设置左缩进
    paragraph.paragraph_format.right_indent = Inches(0)  # 设置右缩进
    paragraph.paragraph_format.first_line_indent = 406400  # 设置首行缩进
    for run in paragraph.runs:
        run.font.size = Pt(16)
        run.font.underline=False
        run.font.strike=False
        run.font.italic=False
        run.font.bold=False
        run.font.color.rgb=RGBColor(0,0,0)
        run.font.name="仿宋_GB2312"
        r=run._element.rPr.rFonts
        r.set(qn('w:eastAsia'),"仿宋_GB2312")
document.save("花园村人大代表联络站简介.docx")




