"""
마크다운 파서

- markdown-it 기반 토큰 파싱
- `- Page N -`, `- Slide N -` 패턴 무시 (DOCX는 동적 페이지)
- 문서 구조 추출 (타이틀, 섹션, 본문)
- 이미지 경로에서 타이틀 추출
"""

from markdown_it import MarkdownIt
from markdown_it.token import Token
from dataclasses import dataclass, field
from typing import List, Optional, Dict, Any
from pathlib import Path
import re


@dataclass
class ContentBlock:
    """콘텐츠 블록"""
    block_type: str  # 'title', 'subtitle', 'heading', 'paragraph', 'list_item',
                     # 'table', 'code', 'image', 'horizontal_rule', 'blockquote'
    content: str = ''
    level: int = 0  # heading level (1-6), list depth
    list_type: str = ''  # 'bullet', 'ordered'
    children: List['ContentBlock'] = field(default_factory=list)
    attributes: Dict[str, Any] = field(default_factory=dict)

    # 인라인 서식 정보
    inline_formats: List[Dict] = field(default_factory=list)  # [{start, end, format}]


@dataclass
class DocumentStructure:
    """파싱된 문서 구조"""
    title: str = ''
    subtitle: str = ''
    first_image_path: str = ''  # 첫 번째 이미지 경로 (타이틀 추출용)
    sections: List['Section'] = field(default_factory=list)
    raw_blocks: List[ContentBlock] = field(default_factory=list)


@dataclass
class Section:
    """문서 섹션"""
    heading: str = ''
    level: int = 1
    blocks: List[ContentBlock] = field(default_factory=list)


class MarkdownParser:
    """마크다운을 구조화된 데이터로 파싱"""

    # 무시할 패턴 (페이지 구분자)
    IGNORE_PATTERNS = [
        r'^-\s*Page\s*\d+\s*-$',
        r'^-\s*Slide\s*\d+\s*-$',
        r'^-\s*페이지\s*\d+\s*-$',
    ]

    def __init__(self):
        self.md = MarkdownIt('commonmark', {'breaks': True, 'html': True})
        self.md.enable('table')

        self._ignore_regex = [re.compile(p, re.IGNORECASE) for p in self.IGNORE_PATTERNS]

    def parse(self, md_content: str) -> DocumentStructure:
        """마크다운 파싱"""
        # 전처리: 무시할 패턴 제거
        cleaned_content = self._preprocess(md_content)

        # 토큰 파싱
        tokens = self.md.parse(cleaned_content)

        # 구조화
        doc = DocumentStructure()
        doc.raw_blocks = self._tokens_to_blocks(tokens)

        # 타이틀/서브타이틀 추출
        self._extract_title_subtitle(doc)

        # 섹션 구조화
        doc.sections = self._organize_sections(doc.raw_blocks)

        return doc

    def parse_file(self, file_path: str) -> DocumentStructure:
        """파일에서 파싱"""
        with open(file_path, 'r', encoding='utf-8') as f:
            content = f.read()
        return self.parse(content)

    def _preprocess(self, content: str) -> str:
        """전처리 - 무시할 패턴 제거 및 이미지 경로 공백 인코딩"""
        lines = content.split('\n')
        result_lines = []

        # 이미지 패턴: ![...](path)
        image_pattern = re.compile(r'!\[([^\]]*)\]\(([^)]+)\)')

        for line in lines:
            stripped = line.strip()

            # 빈 줄은 유지
            if not stripped:
                result_lines.append(line)
                continue

            # 무시 패턴 체크
            should_ignore = any(regex.match(stripped) for regex in self._ignore_regex)
            if should_ignore:
                continue

            # 이미지 경로의 공백을 %20으로 인코딩 (markdown-it 파싱 호환성)
            def encode_image_path(match):
                alt = match.group(1)
                path = match.group(2)
                # 공백만 인코딩 (한글 등은 그대로)
                encoded_path = path.replace(' ', '%20')
                return f'![{alt}]({encoded_path})'

            line = image_pattern.sub(encode_image_path, line)
            result_lines.append(line)

        return '\n'.join(result_lines)

    def _tokens_to_blocks(self, tokens: List[Token]) -> List[ContentBlock]:
        """토큰을 ContentBlock 리스트로 변환"""
        blocks = []
        i = 0

        while i < len(tokens):
            token = tokens[i]
            block, consumed = self._process_token(tokens, i)

            if block:
                blocks.append(block)

            i += consumed

        return blocks

    def _process_token(self, tokens: List[Token], index: int) -> tuple:
        """단일 토큰 처리, (ContentBlock, consumed_count) 반환"""
        token = tokens[index]

        # 헤딩
        if token.type == 'heading_open':
            return self._process_heading(tokens, index)

        # 문단
        elif token.type == 'paragraph_open':
            return self._process_paragraph(tokens, index)

        # 불릿 리스트
        elif token.type == 'bullet_list_open':
            return self._process_list(tokens, index, 'bullet')

        # 숫자 리스트
        elif token.type == 'ordered_list_open':
            return self._process_list(tokens, index, 'ordered')

        # 인용문
        elif token.type == 'blockquote_open':
            return self._process_blockquote(tokens, index)

        # 코드 블록
        elif token.type == 'fence' or token.type == 'code_block':
            block = ContentBlock(
                block_type='code',
                content=token.content,
                attributes={'language': token.info or ''}
            )
            return block, 1

        # 테이블
        elif token.type == 'table_open':
            return self._process_table(tokens, index)

        # 수평선
        elif token.type == 'hr':
            return ContentBlock(block_type='horizontal_rule'), 1

        # 기타 (무시)
        return None, 1

    def _process_heading(self, tokens: List[Token], start: int) -> tuple:
        """헤딩 처리"""
        open_token = tokens[start]
        level = int(open_token.tag[1])  # h1 -> 1

        inline_token = tokens[start + 1] if start + 1 < len(tokens) else None
        text, formats = self._extract_inline_content(inline_token)

        block = ContentBlock(
            block_type='heading',
            content=text,
            level=level,
            inline_formats=formats
        )

        return block, 3  # open, inline, close

    def _process_paragraph(self, tokens: List[Token], start: int) -> tuple:
        """문단 처리"""
        inline_token = tokens[start + 1] if start + 1 < len(tokens) else None

        if inline_token and inline_token.type != 'inline':
            return None, 1

        # 이미지 체크
        if inline_token and inline_token.children:
            images = [c for c in inline_token.children if c.type == 'image']
            if images:
                img = images[0]
                block = ContentBlock(
                    block_type='image',
                    content=img.attrGet('alt') or '',
                    attributes={
                        'src': img.attrGet('src') or '',
                        'title': img.attrGet('title') or '',
                    }
                )
                return block, 3

        text, formats = self._extract_inline_content(inline_token)

        # 빈 문단은 건너뛰기
        if not text.strip():
            return None, 3

        block = ContentBlock(
            block_type='paragraph',
            content=text,
            inline_formats=formats
        )

        return block, 3

    def _process_list(self, tokens: List[Token], start: int, list_type: str) -> tuple:
        """리스트 처리"""
        close_tag = f'{list_type}_list_close'
        items = []
        depth = 1
        i = start + 1

        while i < len(tokens):
            token = tokens[i]

            if token.type == f'{list_type}_list_close':
                depth -= 1
                if depth == 0:
                    break

            elif token.type == f'{list_type}_list_open':
                depth += 1

            elif token.type == 'paragraph_open':
                if i + 1 < len(tokens) and tokens[i + 1].type == 'inline':
                    inline_token = tokens[i + 1]
                    text, formats = self._extract_inline_content(inline_token)
                    items.append(ContentBlock(
                        block_type='list_item',
                        content=text,
                        list_type=list_type,
                        inline_formats=formats
                    ))
                    i += 2
                    continue

            i += 1

        # 리스트 컨테이너 블록 생성
        block = ContentBlock(
            block_type='list',
            list_type=list_type,
            children=items
        )

        return block, i - start + 1

    def _process_blockquote(self, tokens: List[Token], start: int) -> tuple:
        """인용문 처리"""
        depth = 1
        content_parts = []
        i = start + 1

        while i < len(tokens):
            token = tokens[i]

            if token.type == 'blockquote_close':
                depth -= 1
                if depth == 0:
                    break

            elif token.type == 'blockquote_open':
                depth += 1

            elif token.type == 'paragraph_open':
                if i + 1 < len(tokens) and tokens[i + 1].type == 'inline':
                    text, _ = self._extract_inline_content(tokens[i + 1])
                    content_parts.append(text)
                    i += 2
                    continue

            i += 1

        block = ContentBlock(
            block_type='blockquote',
            content='\n'.join(content_parts)
        )

        return block, i - start + 1

    def _process_table(self, tokens: List[Token], start: int) -> tuple:
        """테이블 처리"""
        rows = []
        current_row = []
        is_header = False
        i = start + 1

        while i < len(tokens):
            token = tokens[i]

            if token.type == 'table_close':
                break

            elif token.type == 'thead_open':
                is_header = True

            elif token.type == 'thead_close':
                is_header = False

            elif token.type == 'tr_open':
                current_row = []

            elif token.type == 'tr_close':
                rows.append({'cells': current_row, 'is_header': is_header})

            elif token.type == 'inline':
                text, _ = self._extract_inline_content(token)
                current_row.append(text)

            i += 1

        block = ContentBlock(
            block_type='table',
            attributes={'rows': rows}
        )

        return block, i - start + 1

    def _extract_inline_content(self, inline_token: Optional[Token]) -> tuple:
        """인라인 토큰에서 텍스트와 서식 정보 추출"""
        if not inline_token:
            return '', []

        if not inline_token.children:
            return inline_token.content or '', []

        text_parts = []
        formats = []
        current_pos = 0
        format_stack = []

        for child in inline_token.children:
            if child.type == 'text':
                text_parts.append(child.content)
                current_pos += len(child.content)

            elif child.type == 'code_inline':
                start = current_pos
                text_parts.append(child.content)
                current_pos += len(child.content)
                formats.append({'start': start, 'end': current_pos, 'format': 'code'})

            elif child.type == 'softbreak':
                text_parts.append(' ')
                current_pos += 1

            elif child.type == 'hardbreak':
                text_parts.append('\n')
                current_pos += 1

            elif child.type == 'strong_open':
                format_stack.append(('bold', current_pos))

            elif child.type == 'strong_close':
                if format_stack:
                    fmt, start = format_stack.pop()
                    if fmt == 'bold':
                        formats.append({'start': start, 'end': current_pos, 'format': 'bold'})

            elif child.type == 'em_open':
                format_stack.append(('italic', current_pos))

            elif child.type == 'em_close':
                if format_stack:
                    fmt, start = format_stack.pop()
                    if fmt == 'italic':
                        formats.append({'start': start, 'end': current_pos, 'format': 'italic'})

            elif child.type == 's_open':
                format_stack.append(('strike', current_pos))

            elif child.type == 's_close':
                if format_stack:
                    fmt, start = format_stack.pop()
                    if fmt == 'strike':
                        formats.append({'start': start, 'end': current_pos, 'format': 'strike'})

        return ''.join(text_parts), formats

    def _extract_title_subtitle(self, doc: DocumentStructure):
        """타이틀/서브타이틀 추출"""
        from urllib.parse import unquote

        # 1. 첫 번째 이미지 경로에서 파일명 추출 (타이틀 후보)
        for block in doc.raw_blocks:
            if block.block_type == 'image':
                src = block.attributes.get('src', '')
                # URL 디코딩 (%20 -> 공백 등)
                doc.first_image_path = unquote(src)
                # 이미지 경로에서 폴더명 추출 (예: sample/page_0001/... -> sample)
                if doc.first_image_path:
                    parts = doc.first_image_path.split('/')
                    if len(parts) > 1:
                        # 첫 번째 의미있는 폴더명을 타이틀 후보로
                        doc.title = parts[0]
                break

        # 2. 첫 번째 heading level 1을 타이틀로
        for block in doc.raw_blocks:
            if block.block_type == 'heading':
                if block.level == 1 and not doc.title:
                    doc.title = block.content
                elif block.level == 2 and not doc.subtitle:
                    doc.subtitle = block.content
                    break

        # 3. heading이 없으면 첫 번째 굵은 텍스트(**...**)를 타이틀로
        if not doc.title:
            for block in doc.raw_blocks:
                if block.block_type == 'paragraph':
                    # 볼드 서식이 전체인 경우
                    bold_formats = [f for f in block.inline_formats if f['format'] == 'bold']
                    if bold_formats:
                        first_bold = bold_formats[0]
                        if first_bold['start'] == 0 and first_bold['end'] == len(block.content):
                            doc.title = block.content
                            break

    def _organize_sections(self, blocks: List[ContentBlock]) -> List[Section]:
        """블록들을 섹션으로 구조화"""
        sections = []
        current_section = None

        for block in blocks:
            if block.block_type == 'heading' and block.level <= 2:
                # 새 섹션 시작
                if current_section:
                    sections.append(current_section)
                current_section = Section(
                    heading=block.content,
                    level=block.level,
                    blocks=[]
                )
            else:
                if current_section is None:
                    current_section = Section(heading='', level=0, blocks=[])
                current_section.blocks.append(block)

        if current_section:
            sections.append(current_section)

        return sections


if __name__ == '__main__':
    import sys
    import json

    test_file = '/home/shaush/work/parsed-outputs/sample.md'
    if len(sys.argv) > 1:
        test_file = sys.argv[1]

    parser = MarkdownParser()
    doc = parser.parse_file(test_file)

    print(f"📄 Title: {doc.title}")
    print(f"📝 Subtitle: {doc.subtitle}")
    print(f"🖼️ First Image: {doc.first_image_path}")
    print(f"\n📚 Sections: {len(doc.sections)}")

    for i, section in enumerate(doc.sections[:5]):
        print(f"\n  Section {i+1}: {section.heading or '(no heading)'}")
        print(f"    Blocks: {len(section.blocks)}")
        for block in section.blocks[:3]:
            preview = block.content[:50] + '...' if len(block.content) > 50 else block.content
            print(f"      - [{block.block_type}] {preview}")

    print(f"\n📊 Total raw blocks: {len(doc.raw_blocks)}")
    block_types = {}
    for b in doc.raw_blocks:
        block_types[b.block_type] = block_types.get(b.block_type, 0) + 1
    for bt, count in sorted(block_types.items(), key=lambda x: -x[1]):
        print(f"    {bt}: {count}")