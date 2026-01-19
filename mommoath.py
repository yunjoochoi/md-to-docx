import mammoth

def docx_to_markdown_mammoth(docx_path):
    with open(docx_path, "rb") as docx_file:
        # 1. HTML로 먼저 변환 (Mammoth는 HTML 변환에 최적화됨)
        result = mammoth.convert_to_html(docx_file)
        html = result.value
        messages = result.messages # 경고 메시지 확인 가능

    # 2. HTML을 Markdown으로 변환 (markdownify 라이브러리 활용)
    # pip install markdownify
    from markdownify import markdownify
    md_output = markdownify(html, heading_style="ATX") # # 제목 스타일
    
    return md_output

# 실행
md = docx_to_markdown_mammoth("input.docx")
print(md)