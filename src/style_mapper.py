"""
스타일 매퍼

- 마크다운 문법 → DOCX 스타일 매핑
- 템플릿의 스타일 정보를 기반으로 동적 매핑
- 마크다운 특수문자만을 지표로 사용 (하드코딩 텍스트 없음)
"""

from dataclasses import dataclass, field
from typing import Dict, List, Optional, Any
from .template_analyzer import TemplateStructure, StyleInfo
from .markdown_parser import ContentBlock, DocumentStructure, Section


@dataclass
class MappedStyle:
    """매핑된 스타일 정보"""
    style_name: str  # DOCX 스타일 이름 (예: 'heading 1', 'Body Text')
    style_id: str = ''  # DOCX 스타일 ID
    font_name: Optional[str] = None
    font_size_pt: Optional[float] = None
    bold: Optional[bool] = None
    italic: Optional[bool] = None
    color_rgb: Optional[str] = None
    alignment: Optional[str] = None

    # 스타일을 직접 적용해야 하는지 (템플릿에 없는 경우)
    apply_direct: bool = False


@dataclass
class PageContent:
    """페이지별 콘텐츠"""
    page_type: str  # 'cover', 'section', 'body'
    blocks: List['MappedBlock'] = field(default_factory=list)


@dataclass
class MappedBlock:
    """매핑된 콘텐츠 블록"""
    original: ContentBlock
    style: MappedStyle
    page_type: str = 'body'  # 이 블록이 속할 페이지 유형


class StyleMapper:
    """마크다운 → DOCX 스타일 매핑

    우선순위:
    1. outlineLvl 기반 매핑 (가장 확실함)
    2. 스타일 이름 기반 매핑 (폴백)
    """

    # 마크다운 문법 → DOCX 스타일 폴백 매핑 (outlineLvl이 없을 경우에만 사용)
    DEFAULT_MAPPING = {
        # Paragraph
        'paragraph': ['Body Text', 'Normal'],

        # Lists (-, *, +)
        'bullet_list': ['List Bullet', 'List Bullet 2'],
        'ordered_list': ['List Number', 'List Number 2'],

        # Others
        'blockquote': ['Quote', 'Intense Quote', 'Block Text'],
        'code': ['Normal'],  # 코드는 보통 별도 스타일 없음
        'table': ['Table Grid'],
    }

    # 페이지 유형 감지를 위한 패턴
    SECTION_PATTERNS = [
        # 숫자만 있는 헤딩 (01, 02, 1, 2 등)
        r'^\d{1,2}$',
    ]

    def __init__(self, template_structure: Optional[TemplateStructure] = None):
        self.template = template_structure
        self.style_cache: Dict[str, MappedStyle] = {}

        # 템플릿에서 사용 가능한 스타일 목록 구축
        self.available_styles: Dict[str, StyleInfo] = {}
        if template_structure:
            self.available_styles = template_structure.styles

    def map_document(self, doc: DocumentStructure) -> List[PageContent]:
        """전체 문서 매핑 - 페이지 유형별 콘텐츠 분리"""
        pages = []

        # 템플릿에 Cover 스타일(Title, Subtitle)이 있는지 확인
        has_cover_styles = self._has_cover_styles()

        # 타이틀/서브타이틀 추출 (cover 생성 및 중복 방지용)
        title_text = ''
        subtitle_text = ''

        # 1. Cover 페이지 (템플릿에 Cover 스타일이 있을 때만 생성)
        if has_cover_styles:
            title_text = self._extract_title(doc)
            subtitle_text = self._extract_subtitle(doc)

            cover = PageContent(page_type='cover', blocks=[])

            if title_text:
                title_block = ContentBlock(block_type='heading', content=title_text, level=1)
                cover.blocks.append(MappedBlock(
                    original=title_block,
                    style=self._get_style('title'),
                    page_type='cover'
                ))

            if subtitle_text:
                subtitle_block = ContentBlock(block_type='heading', content=subtitle_text, level=2)
                cover.blocks.append(MappedBlock(
                    original=subtitle_block,
                    style=self._get_style('subtitle'),
                    page_type='cover'
                ))

            if cover.blocks:
                pages.append(cover)

        # 2. 본문 페이지들 - 엄격한 필터링
        current_body = PageContent(page_type='body', blocks=[])
        skip_next_heading = False  # 섹션 번호 다음 제목 스킵용

        for i, block in enumerate(doc.raw_blocks):
            # Cover에서 이미 사용된 콘텐츠 스킵 (Cover가 있을 때만)
            if has_cover_styles and self._is_title_block(block, title_text, subtitle_text):
                continue

            # 이미지 블록은 본문에서 스킵 (첫 이미지는 cover 판단에만 사용)
            if block.block_type == 'image' and i < 3:
                continue

            # 섹션 헤드라인 감지 (숫자만 있는 문단)
            section_number = self._extract_section_number(block)
            if section_number:
                # 현재 body 저장
                if current_body.blocks:
                    pages.append(current_body)
                    current_body = PageContent(page_type='body', blocks=[])

                # 섹션 페이지 생성
                section_page = PageContent(page_type='section', blocks=[])

                # 섹션 번호
                section_page.blocks.append(MappedBlock(
                    original=block,
                    style=self._get_style('section_number'),
                    page_type='section'
                ))

                # 다음 블록이 섹션 제목인지 확인
                next_title = self._get_section_title(doc.raw_blocks, i)
                if next_title:
                    title_block = ContentBlock(
                        block_type='heading',
                        content=next_title,
                        level=2
                    )
                    section_page.blocks.append(MappedBlock(
                        original=title_block,
                        style=self._get_style('section_title'),
                        page_type='section'
                    ))
                    skip_next_heading = True

                pages.append(section_page)
                continue

            # 섹션 제목으로 이미 처리된 경우 스킵
            if skip_next_heading and block.block_type in ('heading', 'paragraph'):
                if self._is_section_title_candidate(block):
                    skip_next_heading = False
                    continue

            # 일반 블록 매핑
            mapped = self._map_block(block)
            if mapped:
                current_body.blocks.append(mapped)

        # 마지막 body 페이지 추가
        if current_body.blocks:
            pages.append(current_body)

        return pages

    def _extract_title(self, doc: DocumentStructure) -> str:
        """타이틀 추출 - 이미지 경로에서 폴더명 우선, 없으면 첫 heading"""
        # 1. 첫 번째 이미지 경로에서 폴더명 추출 (우선순위 높음)
        if doc.first_image_path:
            parts = doc.first_image_path.split('/')
            if parts:
                # 의미있는 첫 폴더명
                for part in parts:
                    if part and not part.startswith('page_') and not part.startswith('pictures'):
                        return part

        # 2. 첫 heading 1
        for block in doc.raw_blocks:
            if block.block_type == 'heading' and block.level == 1:
                return block.content

        # 3. doc.title 사용
        return doc.title or ''

    def _extract_subtitle(self, doc: DocumentStructure) -> str:
        """서브타이틀 추출 - 첫 heading (타이틀이 이미지 경로인 경우)"""
        title = self._extract_title(doc)

        # 이미지 경로에서 타이틀을 추출한 경우, 첫 heading을 subtitle로 사용
        if doc.first_image_path and title:
            # 첫 heading 2를 subtitle로
            for block in doc.raw_blocks:
                if block.block_type == 'heading' and block.level == 2:
                    return block.content

        return doc.subtitle or ''

    def _is_title_block(self, block: ContentBlock, title: str, subtitle: str) -> bool:
        """이 블록이 타이틀/서브타이틀인지 확인"""
        if block.block_type == 'heading':
            if block.level == 1 and block.content == title:
                return True
            if block.level == 2 and block.content == subtitle:
                return True
        return False

    def _extract_section_number(self, block: ContentBlock) -> Optional[str]:
        """섹션 번호 추출 (숫자만 있는 문단)"""
        import re
        if block.block_type != 'paragraph':
            return None

        text = block.content.strip()

        # 1자리 또는 2자리 숫자만
        if re.match(r'^0?\d$', text):
            return text

        return None

    def _get_section_title(self, blocks: List[ContentBlock], current_idx: int) -> Optional[str]:
        """섹션 번호 다음의 섹션 제목 추출"""
        # 다음 몇 개 블록에서 제목 찾기
        for j in range(current_idx + 1, min(current_idx + 4, len(blocks))):
            next_block = blocks[j]
            if next_block.block_type == 'image':
                continue
            if next_block.block_type in ('heading', 'paragraph'):
                text = next_block.content.strip()
                # 짧은 텍스트면 섹션 제목
                if text and len(text) < 50:
                    return text
            break
        return None

    def _has_cover_styles(self) -> bool:
        """템플릿에 Cover 스타일(Title, Subtitle)이 있는지 확인"""
        if not self.available_styles:
            return False

        cover_style_names = {'Title', 'Subtitle', 'af0', 'ae'}
        for style_name in self.available_styles:
            if style_name in cover_style_names:
                return True
        return False

    def _is_section_title_candidate(self, block: ContentBlock) -> bool:
        """섹션 제목 후보인지 확인"""
        text = block.content.strip()
        return len(text) < 50 and len(text) > 0

    def _is_section_headline(self, block: ContentBlock) -> bool:
        """섹션 헤드라인인지 판단"""
        if block.block_type != 'paragraph':
            return False

        import re
        text = block.content.strip()

        # 숫자만 있는 짧은 텍스트
        if re.match(r'^\d{1,2}$', text):
            return True

        return False

    def _map_block(self, block: ContentBlock) -> Optional[MappedBlock]:
        """단일 블록 매핑"""
        style = None

        if block.block_type == 'heading':
            style = self._get_style(f'heading_{block.level}')

        elif block.block_type == 'paragraph':
            style = self._get_style('paragraph')

        elif block.block_type == 'list':
            # 리스트는 children을 개별 매핑해야 함
            # 여기서는 컨테이너만 처리
            style = self._get_style(f'{block.list_type}_list')

        elif block.block_type == 'list_item':
            style = self._get_style(f'{block.list_type}_list')

        elif block.block_type == 'blockquote':
            style = self._get_style('blockquote')

        elif block.block_type == 'code':
            style = self._get_style('code')

        elif block.block_type == 'table':
            style = self._get_style('table')

        elif block.block_type == 'image':
            style = self._get_style('paragraph')  # 이미지는 문단 스타일

        elif block.block_type == 'horizontal_rule':
            style = self._get_style('paragraph')

        if style is None:
            style = self._get_style('paragraph')  # 기본값

        return MappedBlock(original=block, style=style, page_type='body')

    def _get_style(self, style_key: str) -> MappedStyle:
        """스타일 키로 MappedStyle 가져오기

        우선순위:
        1. heading_N -> outlineLvl 기반 매핑 (절대적 기준)
        2. 특수 스타일(title, subtitle) -> 이름 기반 폴백
        3. 일반 스타일 -> DEFAULT_MAPPING 폴백
        """
        if style_key in self.style_cache:
            return self.style_cache[style_key]

        style = None

        # 1. Heading 스타일: outlineLvl 기반 매핑 (가장 확실함)
        if style_key.startswith('heading_'):
            md_level = int(style_key.split('_')[1])  # 'heading_1' -> 1
            outline_level = md_level - 1  # Markdown # -> outlineLvl 0

            # outlineLvl로 정확히 찾기
            if self.template:
                style_info = self.template.get_style_by_outline_level(outline_level)
                if style_info:
                    style = self._style_info_to_mapped(style_info)

            # 폴백: 이름 기반 탐색 (만약 outlineLvl이 없다면)
            if style is None:
                fallback_names = [f'heading {md_level}', f'Heading {md_level}', f'Heading{md_level}']
                style = self._find_best_style(fallback_names)

        # 2. 페이지 유형별 특수 스타일 처리
        elif style_key == 'title':
            style = self._find_best_style(['Title', 'af0'])
        elif style_key == 'subtitle':
            style = self._find_best_style(['Subtitle', 'ae'])
        elif style_key == 'section_number':
            # 섹션 번호 - heading 1과 동일 (outlineLvl 0)
            if self.template:
                style_info = self.template.get_style_by_outline_level(0)
                if style_info:
                    style = self._style_info_to_mapped(style_info)
            if style is None:
                style = self._find_best_style(['heading 1', 'Heading 1'])
        elif style_key == 'section_title':
            # 섹션 제목 - heading 2와 동일 (outlineLvl 1)
            if self.template:
                style_info = self.template.get_style_by_outline_level(1)
                if style_info:
                    style = self._style_info_to_mapped(style_info)
            if style is None:
                style = self._find_best_style(['Subtitle', 'ae', 'heading 2'])
        elif style_key == 'section_headline':
            # 섹션 헤드라인 - heading 1과 동일
            if self.template:
                style_info = self.template.get_style_by_outline_level(0)
                if style_info:
                    style = self._style_info_to_mapped(style_info)
            if style is None:
                style = self._find_best_style(['heading 1', 'Heading 1'])

        # 3. 일반 스타일 (paragraph, list, blockquote 등)
        else:
            candidates = self.DEFAULT_MAPPING.get(style_key, ['Normal'])
            style = self._find_best_style(candidates)

        # 최종 폴백
        if style is None:
            style = MappedStyle(style_name='Normal', apply_direct=True)

        self.style_cache[style_key] = style
        return style

    def _find_best_style(self, candidates: List[str]) -> Optional[MappedStyle]:
        """후보 중에서 템플릿에 있는 스타일 찾기"""
        for candidate in candidates:
            # style_id로 먼저 찾기
            if candidate in self.available_styles:
                style_info = self.available_styles[candidate]
                return self._style_info_to_mapped(style_info)

            # name으로 찾기
            for style_id, style_info in self.available_styles.items():
                if style_info.name.lower() == candidate.lower():
                    return self._style_info_to_mapped(style_info)

        # 템플릿에 없으면 기본 스타일 반환
        return MappedStyle(
            style_name=candidates[0] if candidates else 'Normal',
            apply_direct=True
        )

    def _style_info_to_mapped(self, style_info: StyleInfo) -> MappedStyle:
        """StyleInfo를 MappedStyle로 변환"""
        return MappedStyle(
            style_name=style_info.name,
            style_id=style_info.style_id,
            font_name=style_info.font_name,
            font_size_pt=style_info.font_size_pt,
            bold=style_info.bold,
            italic=style_info.italic,
            color_rgb=style_info.color_rgb,
            alignment=style_info.alignment,
            apply_direct=False
        )

    def get_inline_format_style(self, format_type: str) -> Dict[str, Any]:
        """인라인 서식 스타일 (bold, italic, code 등)"""
        styles = {
            'bold': {'bold': True},
            'italic': {'italic': True},
            'strike': {'strike': True},
            'code': {
                'font_name': 'Consolas',
                'font_size_pt': 10,
            },
        }
        return styles.get(format_type, {})


if __name__ == '__main__':
    from .template_analyzer import DocxTemplateAnalyzer
    from .markdown_parser import MarkdownParser

    # 테스트
    template_path = '/home/shaush/md-to-docx/docx_only/[Word 템플릿] A4.docx'
    md_path = '/home/shaush/work/parsed-outputs/sample.md'

    # 템플릿 분석
    analyzer = DocxTemplateAnalyzer(template_path)
    template = analyzer.analyze()

    # 마크다운 파싱
    parser = MarkdownParser()
    doc = parser.parse_file(md_path)

    # 스타일 매핑
    mapper = StyleMapper(template)
    pages = mapper.map_document(doc)

    print(f"📄 Pages: {len(pages)}")
    for i, page in enumerate(pages):
        print(f"\n  Page {i+1} [{page.page_type}]:")
        print(f"    Blocks: {len(page.blocks)}")
        for block in page.blocks[:3]:
            print(f"      - [{block.original.block_type}] "
                  f"style: {block.style.style_name}, "
                  f"content: {block.original.content[:30]}...")