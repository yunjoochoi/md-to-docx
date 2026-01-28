"""
템플릿 페이지 구조 분석기 (범용 호환성 버전)

- 하드코딩된 스타일 ID 제거 (af0, ae 등 삭제)
- 표준 스타일 이름(Title, Heading 1 등)을 기반으로 ID 동적 매핑
- 한글/영문 워드 호환성 강화 ('Heading 1' vs '제목 1')
- 스타일 이름이 달라도 텍스트 분량 등으로 표지를 추측하는 휴리스틱 추가
"""

from docx import Document
from pathlib import Path
from dataclasses import dataclass, field
from typing import List, Dict, Optional, Set

@dataclass
class PageInfo:
    """페이지 정보"""
    page_number: int
    page_type: str  # 'cover', 'toc', 'section', 'body'
    styles_used: List[str] = field(default_factory=list)
    paragraph_count: int = 0
    content_preview: str = ''

@dataclass
class TemplatePageStructure:
    """템플릿 페이지 구조"""
    pages: List[PageInfo] = field(default_factory=list)
    cover_page: Optional[int] = None
    toc_page: Optional[int] = None
    section_pages: List[int] = field(default_factory=list)
    body_start_page: Optional[int] = None


class TemplatePageAnalyzer:
    """템플릿 페이지 구조 분석 (동적 매핑 적용)"""

    def __init__(self, docx_path: str):
        self.docx_path = Path(docx_path)
        self.doc = Document(docx_path)
        self.structure = TemplatePageStructure()

        # 동적으로 채워질 스타일 ID 집합 (초기엔 비어있음)
        self.cover_styles: Set[str] = set()
        self.toc_styles: Set[str] = set()
        self.section_styles: Set[str] = set()
        self.body_styles: Set[str] = set()

        # 1. 문서를 스캔하여 스타일 ID들을 동적으로 등록
        self._discover_style_ids()

    def _discover_style_ids(self):
        """
        문서 내의 모든 스타일을 검사하여, 이름(Name)을 기반으로 ID를 분류합니다.
        어떤 템플릿이 들어와도 이름이 표준(Title, 제목 등)을 따른다면 작동합니다.
        """
        # 찾고자 하는 표준 이름 패턴 (소문자 기준)
        target_patterns = {
            'cover': ['title', 'subtitle', 'cover title', '제목', '부제목'],
            'toc': ['toc', 'table of contents', '목차', '차례'],
            'section': ['heading 1', '제목 1', 'part title', 'section'],
            'body': ['normal', 'body text', '본문', '바탕글', 'list paragraph']
        }

        # docx의 styles 속성을 순회
        for style in self.doc.styles:
            # 스타일 이름과 ID 확보
            if not style.name or not style.style_id:
                continue
            
            name_lower = style.name.lower()
            style_id = style.style_id

            # 1. Cover 스타일 감지
            if any(p == name_lower for p in target_patterns['cover']):
                self.cover_styles.add(style_id)

            # 2. TOC 스타일 감지 (이름에 포함되어 있으면 인정)
            if any(p in name_lower for p in target_patterns['toc']):
                self.toc_styles.add(style_id)

            # 3. Section 스타일 감지
            if any(p == name_lower for p in target_patterns['section']):
                self.section_styles.add(style_id)

            # 4. Body 스타일 감지
            if any(p == name_lower for p in target_patterns['body']):
                self.body_styles.add(style_id)
        
        # 디버깅: 감지된 스타일 출력
        # print(f"[DEBUG] 감지된 Cover Styles: {self.cover_styles}")
        # print(f"[DEBUG] 감지된 Section Styles: {self.section_styles}")

    def analyze(self) -> TemplatePageStructure:
        """페이지 구조 분석 실행"""
        paragraphs_by_page = self._split_by_page_breaks()

        for page_num, paragraphs in enumerate(paragraphs_by_page):
            page_info = self._analyze_page(page_num, paragraphs)
            self.structure.pages.append(page_info)

            # 인덱싱 (가장 먼저 발견된 페이지를 해당 유형의 대표로 설정)
            if page_info.page_type == 'cover' and self.structure.cover_page is None:
                self.structure.cover_page = page_num
            elif page_info.page_type == 'toc' and self.structure.toc_page is None:
                self.structure.toc_page = page_num
            elif page_info.page_type == 'section':
                self.structure.section_pages.append(page_num)
            elif page_info.page_type == 'body' and self.structure.body_start_page is None:
                self.structure.body_start_page = page_num

        return self.structure

    def _split_by_page_breaks(self) -> List[List]:
        """페이지(Hard Break) 기준으로 문단 분리"""
        pages = []
        current_page = []

        for para in self.doc.paragraphs:
            current_page.append(para)
            if self._has_page_break(para):
                pages.append(current_page)
                current_page = []
        
        if current_page:
            pages.append(current_page)
        return pages

    def _has_page_break(self, para) -> bool:
        """<w:br type="page"> 태그 확인"""
        if para._element is None: return False
        xml = para._element.xml
        return 'w:br' in xml and 'type="page"' in xml

    def _analyze_page(self, page_num: int, paragraphs: List) -> PageInfo:
        """페이지 유형 판단 (스타일 기반 + 휴리스틱)"""
        styles = []
        texts = []
        for p in paragraphs:
            if p.style and p.style.style_id:
                styles.append(p.style.style_id)
            if p.text.strip():
                texts.append(p.text.strip()[:30])

        style_set = set(styles)

        # 1. Cover 판단: 첫 페이지이고, Cover 스타일이 있거나 글자 수가 매우 적을 때
        is_cover = False
        if page_num == 0:
            if style_set & self.cover_styles:
                is_cover = True
            # 스타일을 못 찾았더라도, 텍스트가 매우 적고(5줄 이하) 첫 페이지면 표지로 간주 (휴리스틱)
            elif len(paragraphs) <= 5 and len(texts) > 0:
                is_cover = True
        
        if is_cover:
            return PageInfo(page_num, 'cover', list(style_set), len(paragraphs), str(texts))

        # 2. TOC 판단
        if style_set & self.toc_styles:
            return PageInfo(page_num, 'toc', list(style_set), len(paragraphs), "목차")

        # 3. Section 판단: Section 스타일(H1)이 존재하고, 본문 스타일 빈도가 낮을 때
        has_section_header = bool(style_set & self.section_styles)
        body_style_count = len([s for s in styles if s in self.body_styles])
        
        # 제목만 덩그러니 있거나(글자 수 적음), 섹션 스타일이 명확할 때
        if has_section_header and body_style_count <= 2 and len(texts) < 5:
            return PageInfo(page_num, 'section', list(style_set), len(paragraphs), str(texts))

        # 4. 나머지는 모두 Body
        return PageInfo(page_num, 'body', list(style_set), len(paragraphs), str(texts))

    def get_page_mapping_rules(self) -> Dict:
        """매핑 규칙 반환 (동적으로 찾은 ID 중 하나를 대표로 반환)"""
        def get_best_id(style_set, fallback_name):
            # 찾은 ID가 있으면 그거 쓰고, 없으면 표준 이름(fallback)을 그대로 씀
            return next(iter(style_set), fallback_name)

        return {
            'cover': {
                'styles': {'title': get_best_id(self.cover_styles, 'Title')}
            },
            'section': {
                'styles': {'section_title': get_best_id(self.section_styles, 'Heading 1')}
            },
            'body': {
                'styles': {'paragraph': get_best_id(self.body_styles, 'Normal')}
            }
        }

if __name__ == '__main__':
    import sys
    # 테스트: 인자로 받은 파일 분석 (없으면 기본값)
    path = sys.argv[1] if len(sys.argv) > 1 else "template.docx"
    
    if Path(path).exists():
        analyzer = TemplatePageAnalyzer(path)
        res = analyzer.analyze()
        print(f"File: {path}")
        print(f"Cover Page Index: {res.cover_page}")
        print(f"Section Pages: {res.section_pages}")
        print(f"Body Start Index: {res.body_start_page}")
        
        # 동적 감지된 스타일 확인
        print("\n[Detected Styles]")
        print(f"Cover IDs: {analyzer.cover_styles}")
        print(f"Section IDs: {analyzer.section_styles}")
    else:
        print(f"File not found: {path}")