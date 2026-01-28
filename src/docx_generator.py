
"""
DOCX 생성기 (개선된 버전)

- 템플릿 기반 문서 생성
- 매핑된 스타일 적용
- 섹션(Section) 및 배경 디자인 보존 로직 추가
- 표지/본문 디자인이 다른 템플릿 완벽 지원
"""

from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn, nsmap
from pathlib import Path
from typing import List, Optional, Dict, Any

from .template_analyzer import TemplateStructure, DocxTemplateAnalyzer
from .markdown_parser import DocumentStructure
from .style_mapper import StyleMapper, MappedStyle, PageContent, ContentBlock

class DocxGenerator:
    """DOCX 문서 생성기"""

    def __init__(self, template_path: Optional[str] = None):
        self.template_path = template_path
        self.template_structure: Optional[TemplateStructure] = None
        self.md_base_path: Optional[Path] = None
        
        # 보존된 섹션 브레이크 문단들 (삽입 위치 지표)
        self.preserved_section_breaks = []

        if template_path and Path(template_path).exists():
            analyzer = DocxTemplateAnalyzer(template_path)
            self.template_structure = analyzer.analyze()

    def generate(self, pages: List[PageContent], output_path: str) -> str:
        """페이지 콘텐츠로부터 DOCX 생성"""
        if self.template_path and Path(self.template_path).exists():
            doc = Document(self.template_path)
            self._clear_body_content_smart(doc)
        else:
            doc = Document()
            self.preserved_section_breaks = []

        # 각 페이지 처리
        for i, page in enumerate(pages):
            # 삽입 위치 결정 (Target Paragraph)
            target = None
            if page.page_type == 'cover' and self.preserved_section_breaks:
                # 표지는 첫 번째 섹션 브레이크(표지 구역 끝) 앞에 삽입
                target = self.preserved_section_breaks[0]
            
            # 페이지 생성
            if page.page_type == 'cover':
                self._generate_cover_page(doc, page, target_para=target)
            elif page.page_type == 'section':
                self._generate_section_page(doc, page)
            else:
                self._generate_body_content(doc, page)

        doc.save(output_path)
        return output_path

    def _clear_body_content_smart(self, doc: Document):
        """본문 내용을 삭제하되, 섹션 구조와 배경 이미지는 보존"""
        body = doc.element.body
        self.preserved_section_breaks = []
        
        to_delete = []
        
        for child in body:
            # 1. 문서의 마지막 섹션 설정(sectPr)은 건드리지 않음 (태그 이름 체크)
            if child.tag.endswith('sectPr'):
                continue
                
            # 2. 문단(<w:p>)인 경우 정밀 검사
            if child.tag.endswith('p'):
                # 2-1. 섹션 브레이크(<w:sectPr>)가 포함된 문단은 보존 (디자인 구역 유지)
                if child.find(f'.//{{{nsmap["w"]}}}sectPr') is not None:
                    self.preserved_section_breaks.append(child)
                    continue
                
                # 2-2. 배경 이미지/도형 보존
                # 텍스트는 없는데 drawing(이미지)이나 pict(도형)만 있는 경우 유지
                has_text = len(''.join(child.itertext()).strip()) > 0
                has_drawing = child.find(f'.//{{{nsmap["w"]}}}drawing') is not None
                has_pict = child.find(f'.//{{{nsmap["w"]}}}pict') is not None
                
                if (has_drawing or has_pict) and not has_text:
                    continue

            # 그 외(일반 텍스트, 표 등)는 삭제
            to_delete.append(child)
            
        for child in to_delete:
            body.remove(child)

    def generate_from_file(self, md_file: str, output_path: str) -> str:
        """마크다운 파일에서 DOCX 생성"""
        from .markdown_parser import MarkdownParser
        
        md_path = Path(md_file)
        self.md_base_path = md_path.parent

        with open(md_file, 'r', encoding='utf-8') as f:
            md_content = f.read()

        parser = MarkdownParser()
        doc_structure = parser.parse(md_content)

        mapper = StyleMapper(self.template_structure)
        pages = mapper.map_document(doc_structure)

        return self.generate(pages, output_path)

    def _generate_cover_page(self, doc: Document, page: PageContent, target_para=None):
        """표지 페이지 생성"""
        for mapped_block in page.blocks:
            block = mapped_block.original
            style = mapped_block.style

            if block.block_type == 'heading':
                para = self._add_paragraph_with_style(doc, block.content, style, target_para=target_para)
                if style.alignment is None:
                    para.alignment = WD_ALIGN_PARAGRAPH.CENTER

        # 섹션 브레이크를 재사용했다면 별도의 페이지 나누기가 필요 없음
        if not target_para:
            self._add_page_break(doc)

    def _generate_section_page(self, doc: Document, page: PageContent):
        """섹션 구분 페이지 생성"""
        for mapped_block in page.blocks:
            para = self._add_paragraph_with_style(doc, mapped_block.original.content, mapped_block.style)
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        self._add_page_break(doc)

    def _generate_body_content(self, doc: Document, page: PageContent):
        """본문 콘텐츠 생성"""
        for mapped_block in page.blocks:
            block = mapped_block.original
            style = mapped_block.style
            
            # (기존 코드와 동일, 인자만 전달)
            if block.block_type == 'heading':
                self._add_paragraph_with_style(doc, block.content, style, inline_formats=block.inline_formats)
            elif block.block_type == 'paragraph':
                self._add_paragraph_with_style(doc, block.content, style, inline_formats=block.inline_formats)
            elif block.block_type == 'list':
                self._add_list(doc, block, style)
            elif block.block_type == 'list_item':
                self._add_paragraph_with_style(doc, block.content, style, inline_formats=block.inline_formats)
            elif block.block_type == 'blockquote':
                self._add_paragraph_with_style(doc, block.content, style)
            elif block.block_type == 'code':
                self._add_code_block(doc, block)
            elif block.block_type == 'table':
                self._add_table(doc, block)
            elif block.block_type == 'image':
                self._add_image_placeholder(doc, block)
            elif block.block_type == 'horizontal_rule':
                self._add_horizontal_rule(doc)

    def _add_paragraph_with_style(self, doc: Document, text: str, style: MappedStyle, 
                                  inline_formats: List[Dict] = None, target_para=None) -> Any:
        """스타일이 적용된 문단 추가 (삽입 위치 지정 가능)"""
        
        # 1. 문단 생성 (위치에 따라)
        if target_para:
            # target_para는 OXML 요소이므로 이를 감싸는 Paragraph 객체를 만들어야 함
            # 하지만 python-docx API로는 insert_paragraph_before가 텍스트/스타일만 받음
            # 여기서는 약간의 트릭 사용:
            # target_para(XML)를 가진 Paragraph 객체를 찾거나 새로 생성하는 대신
            # doc.add_paragraph() 후 이동시키는 건 복잡함.
            # 가장 쉬운 방법: target_para(XML 요소) 앞에 insert
            
            # python-docx의 프록시 객체 생성 없이 low-level insert
            # 하지만 스타일 적용 편의를 위해 insert_paragraph_before API 사용 시도
            # target_para가 doc.paragraphs 리스트에 있는 객체여야 함
            
            # OXML Element를 Paragraph 객체로 래핑하여 찾기 (비효율적이지만 확실)
            found_p = None
            for p in doc.paragraphs:
                if p._element == target_para:
                    found_p = p
                    break
            
            if found_p:
                para = found_p.insert_paragraph_before(style=style.style_id if style.style_id else None)
            else:
                # 못 찾으면 그냥 append
                para = doc.add_paragraph(style=style.style_id if style.style_id else None)
        else:
            # 기본 append
            try:
                para = doc.add_paragraph(style=style.style_name)
            except:
                try:
                    para = doc.add_paragraph(style=style.style_id)
                except:
                    para = doc.add_paragraph()

        # 2. 내용 채우기 (기존 코드와 동일)
        if inline_formats:
            self._apply_inline_formats(para, text, inline_formats, style)
        else:
            run = para.add_run(text)
            self._apply_run_style(run, style)

        # 정렬 적용
        if style.alignment:
            alignment_map = {
                'left': WD_ALIGN_PARAGRAPH.LEFT,
                'center': WD_ALIGN_PARAGRAPH.CENTER,
                'right': WD_ALIGN_PARAGRAPH.RIGHT,
                'both': WD_ALIGN_PARAGRAPH.JUSTIFY,
            }
            if style.alignment in alignment_map:
                para.alignment = alignment_map[style.alignment]

        return para

    # --- 이하 헬퍼 메서드들은 기존 코드와 동일 (생략 없이 사용하세요) ---
    def _apply_run_style(self, run, style: MappedStyle):
        if not style.apply_direct:
            if style.color_rgb: self._set_run_color(run, style.color_rgb)
            return
        if style.font_name:
            run.font.name = style.font_name
            run._element.rPr.rFonts.set(qn('w:eastAsia'), style.font_name)
        if style.font_size_pt: run.font.size = Pt(style.font_size_pt)
        if style.bold: run.bold = True
        if style.italic: run.italic = True
        if style.color_rgb: self._set_run_color(run, style.color_rgb)

    def _set_run_color(self, run, color_rgb: str):
        if color_rgb and color_rgb != 'auto':
            try:
                color_rgb = color_rgb.lstrip('#')
                if len(color_rgb) == 6:
                    run.font.color.rgb = RGBColor(int(color_rgb[:2], 16), int(color_rgb[2:4], 16), int(color_rgb[4:], 16))
            except: pass

    def _apply_inline_formats(self, para, text: str, formats: List[Dict], base_style: MappedStyle):
        if not formats:
            run = para.add_run(text)
            self._apply_run_style(run, base_style)
            return
        formats = sorted(formats, key=lambda f: f['start'])
        last_end = 0
        for fmt in formats:
            start, end, format_type = fmt['start'], fmt['end'], fmt['format']
            if start > last_end:
                run = para.add_run(text[last_end:start])
                self._apply_run_style(run, base_style)
            run = para.add_run(text[start:end])
            self._apply_run_style(run, base_style)
            if format_type == 'bold': run.bold = True
            elif format_type == 'italic': run.italic = True
            elif format_type == 'strike': run.font.strike = True
            elif format_type == 'code': 
                run.font.name = 'Consolas'; run.font.size = Pt(10)
            last_end = end
        if last_end < len(text):
            run = para.add_run(text[last_end:])
            self._apply_run_style(run, base_style)

    def _add_list(self, doc, block, style):
        for item in block.children:
            self._add_paragraph_with_style(doc, item.content, style, inline_formats=item.inline_formats)

    def _add_code_block(self, doc, block):
        para = doc.add_paragraph()
        run = para.add_run(block.content)
        run.font.name = 'Consolas'; run.font.size = Pt(9)
        para.paragraph_format.left_indent = Inches(0.3)

    def _add_table(self, doc, block):
        rows_data = block.attributes.get('rows', [])
        if not rows_data: return
        num_cols = max(len(r['cells']) for r in rows_data)
        table = doc.add_table(rows=len(rows_data), cols=num_cols)
        try: table.style = 'Table Grid'
        except: pass
        for r_idx, r_data in enumerate(rows_data):
            for c_idx, text in enumerate(r_data['cells']):
                if c_idx < num_cols:
                    cell = table.rows[r_idx].cells[c_idx]
                    cell.text = text
                    if r_data.get('is_header'):
                        for p in cell.paragraphs:
                            for r in p.runs: r.bold = True

    def _add_image_placeholder(self, doc, block):
        from urllib.parse import unquote
        src = unquote(block.attributes.get('src', ''))
        para = doc.add_paragraph()
        image_path = None
        if src:
            p = Path(src)
            if p.is_absolute() and p.exists(): image_path = p
            elif self.md_base_path and (self.md_base_path / src).exists(): image_path = self.md_base_path / src
            elif p.exists(): image_path = p
        
        if image_path:
            try:
                para.add_run().add_picture(str(image_path), width=Inches(5))
                return
            except: pass
        para.add_run(f"![Image]({src})").italic = True

    def _add_horizontal_rule(self, doc):
        p = doc.add_paragraph(); p.add_run('─' * 50); p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    def _add_page_break(self, doc):
        doc.add_page_break()