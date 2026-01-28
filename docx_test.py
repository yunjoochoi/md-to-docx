from docx import Document
from docx.enum.style import WD_STYLE_TYPE

document = Document("/home/shaush/md-to-docx/[Word템플릿]A4.docx")
styles=document.styles
paragraph_styles = [
    s for s in styles if s.type == WD_STYLE_TYPE.PARAGRAPH
]
print(styles)
for style in paragraph_styles:
    print(style.name)

paragraph = document.add_paragraph()
print(paragraph.style.name)

print(document.styles['Heading 1'])