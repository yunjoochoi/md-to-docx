"""
DOCX 템플릿 분석기

- 템플릿의 페이지 구조 분석 (Cover, Section, Body 등)
- 스타일 정보 추출 (폰트, 크기, 색상, 상속 관계)
- 배경 이미지 추출 및 분류
- 헤더/푸터 구조 분석
"""

from docx import Document
from docx.shared import Pt, Inches, Emu, Twips
from docx.oxml.ns import qn
from docx.enum.style import WD_STYLE_TYPE
from lxml import etree
import zipfile
import os
import json
from pathlib import Path
from dataclasses import dataclass, field, asdict
from typing import Optional, Dict, List, Any
import re
from .template_page_analyzer import TemplatePageAnalyzer

# XML namespaces
NS = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'pic': 'http://schemas.openxmlformats.org/drawingml/2006/picture',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'rel': 'http://schemas.openxmlformats.org/package/2006/relationships',
}


@dataclass
class StyleInfo:
    """스타일 정보"""
    style_id: str
    name: str
    style_type: str  # paragraph, character, table
    base_style: Optional[str] = None
    font_name: Optional[str] = None
    font_size_pt: Optional[float] = None
    bold: Optional[bool] = None
    italic: Optional[bool] = None
    color_rgb: Optional[str] = None  # 예: "404040"
    alignment: Optional[str] = None
    space_before_pt: Optional[float] = None
    space_after_pt: Optional[float] = None
    line_spacing: Optional[float] = None
    left_indent_pt: Optional[float] = None
    outline_level: Optional[int] = None  # 개요 수준 (0=heading 1, 1=heading 2, ...)

    def to_dict(self) -> dict:
        return {k: v for k, v in asdict(self).items() if v is not None}


@dataclass
class ImageInfo:
    """이미지 정보"""
    name: str
    original_path: str
    extracted_path: str = ''
    width_emu: int = 0
    height_emu: int = 0
    position: str = ''  # 'header', 'footer', 'body'
    is_background: bool = False
    rel_id: str = ''
    page_type: str = ''  # 'cover', 'section', 'body'

    @property
    def width_inches(self) -> float:
        return self.width_emu / 914400 if self.width_emu else 0

    @property
    def height_inches(self) -> float:
        return self.height_emu / 914400 if self.height_emu else 0


@dataclass
class PageTemplate:
    """페이지 템플릿 유형"""
    page_type: str  # 'cover', 'section', 'body', 'toc'
    background_image: Optional[str] = None
    header_image: Optional[str] = None
    footer_image: Optional[str] = None
    styles_used: List[str] = field(default_factory=list)
    description: str = ''


@dataclass
class TemplateStructure:
    """템플릿 전체 구조"""
    file_path: str
    page_width_inches: float = 8.27
    page_height_inches: float = 11.69
    margins: Dict[str, float] = field(default_factory=dict)

    # 스타일 정보 (style_id -> StyleInfo)
    styles: Dict[str, StyleInfo] = field(default_factory=dict)

    # 이미지 정보
    images: List[ImageInfo] = field(default_factory=list)

    # 페이지 유형별 템플릿
    page_templates: Dict[str, PageTemplate] = field(default_factory=dict)
    page_structure: Optional[object] = None
    # 기본 폰트 정보
    default_font: Optional[str] = None
    default_font_size_pt: Optional[float] = None

    # 테마 색상
    theme_colors: Dict[str, str] = field(default_factory=dict)

    def get_style_by_outline_level(self, level: int) -> Optional[StyleInfo]:
        """개요 수준(outlineLvl)으로 스타일 찾기 - 가장 확실한 방법"""
        for style in self.styles.values():
            if style.outline_level == level:
                return style
        return None


class DocxTemplateAnalyzer:
    """DOCX 템플릿을 분석하여 구조 정보 추출"""

    def __init__(self, docx_path: str):
        self.docx_path = Path(docx_path)
        self.output_dir = self.docx_path.parent / f"{self.docx_path.stem}_assets"
        self.doc = Document(docx_path)
        self.structure = TemplateStructure(file_path=str(docx_path))

        # XML 트리 캐시
        self._styles_xml = None
        self._document_xml = None
        self._theme_xml = None

    def analyze(self) -> TemplateStructure:
        """전체 분석 실행"""
        self._load_xml_trees()
        self._analyze_page_setup()
        self._analyze_default_fonts()
        self._analyze_styles_from_xml()
        self._analyze_theme_colors()
        self._extract_images()
        
        self._run_page_analysis() 
        
        return self.structure

    def _run_page_analysis(self):
        """페이지 구조 분석기 실행 및 결과 통합"""
        # 1. 페이지 분석기 인스턴스 생성 (범용 버전 사용)
        page_analyzer = TemplatePageAnalyzer(str(self.docx_path))
        
        # 2. 분석 실행 (동적 스타일 감지 포함)
        page_result = page_analyzer.analyze()
        
        # 3. 결과를 메인 구조체에 저장
        self.structure.page_structure = page_result
        
        # 4. 분석된 페이지 정보를 바탕으로 기존 page_templates 필드 채우기 (호환성 유지)
        # Cover가 감지되었으면 템플릿 정보 업데이트
        if page_result.cover_page is not None:
            cover_info = page_result.pages[page_result.cover_page]
            self.structure.page_templates['cover'] = PageTemplate(
                page_type='cover',
                styles_used=cover_info.styles_used,
                description=f"페이지 {page_result.cover_page}에서 감지된 표지 양식"
            )
        
        # Section이 감지되었으면 업데이트
        if page_result.section_pages:
            # 첫 번째 발견된 섹션 페이지를 템플릿으로 사용
            section_idx = page_result.section_pages[0]
            section_info = page_result.pages[section_idx]
            self.structure.page_templates['section'] = PageTemplate(
                page_type='section',
                styles_used=section_info.styles_used,
                description=f"페이지 {section_idx}에서 감지된 섹션 양식"
            )

        # Body는 기본적으로 body_start_page 사용
        if page_result.body_start_page is not None:
             body_idx = page_result.body_start_page
             body_info = page_result.pages[body_idx]
             self.structure.page_templates['body'] = PageTemplate(
                page_type='body',
                styles_used=body_info.styles_used,
                description="본문 표준 양식"
            )
             
    def _load_xml_trees(self):
        """XML 파일들을 메모리에 로드 (OPC 관계 추적 방식)"""
        with zipfile.ZipFile(self.docx_path, 'r') as zf:
            # styles.xml: .rels 파일을 통해 경로 추적
            styles_path = self._find_styles_xml_via_rels(zf)
            if styles_path:
                with zf.open(styles_path) as f:
                    self._styles_xml = etree.parse(f)
            else:
                # 폴백: 표준 경로 시도
                for fallback in ['word/styles.xml', 'word/styles2.xml']:
                    if fallback in zf.namelist():
                        with zf.open(fallback) as f:
                            self._styles_xml = etree.parse(f)
                        break

            # document.xml
            if 'word/document.xml' in zf.namelist():
                with zf.open('word/document.xml') as f:
                    self._document_xml = etree.parse(f)

            # theme1.xml
            if 'word/theme/theme1.xml' in zf.namelist():
                with zf.open('word/theme/theme1.xml') as f:
                    self._theme_xml = etree.parse(f)

    def _find_styles_xml_via_rels(self, zf: zipfile.ZipFile) -> Optional[str]:
        """OPC 관계 파일을 통해 styles.xml 경로 추적"""
        try:
            # word/document.xml의 관계 파일 확인
            rels_path = 'word/_rels/document.xml.rels'
            if rels_path not in zf.namelist():
                return None

            with zf.open(rels_path) as f:
                rels_tree = etree.parse(f)
                rels_ns = {'rel': 'http://schemas.openxmlformats.org/package/2006/relationships'}

                # Type이 styles인 Relationship 찾기
                for rel in rels_tree.findall('.//rel:Relationship', rels_ns):
                    rel_type = rel.get('Type', '')
                    if 'styles' in rel_type.lower():
                        target = rel.get('Target', '')
                        if target:
                            # Target 경로 정규화:
                            # "/word/styles2.xml" -> "word/styles2.xml"
                            # "styles.xml" -> "word/styles.xml"
                            # "../styles.xml" -> "word/styles.xml"
                            target = target.lstrip('/')  # 절대 경로 표시 제거
                            if not target.startswith('word/'):
                                # 상대 경로인 경우 word/ 기준으로 해석
                                target = target.lstrip('../')
                                styles_path = f"word/{target}"
                            else:
                                styles_path = target

                            if styles_path in zf.namelist():
                                return styles_path
        except Exception:
            pass

        return None

    def _analyze_page_setup(self):
        """페이지 설정 분석"""
        section = self.doc.sections[0]
        self.structure.page_width_inches = section.page_width.inches
        self.structure.page_height_inches = section.page_height.inches
        self.structure.margins = {
            'left': section.left_margin.inches,
            'right': section.right_margin.inches,
            'top': section.top_margin.inches,
            'bottom': section.bottom_margin.inches,
        }

    def _analyze_default_fonts(self):
        """기본 폰트 설정 분석"""
        if self._styles_xml is None:
            return

        root = self._styles_xml.getroot()

        # docDefaults에서 기본 폰트 추출
        doc_defaults = root.find('.//w:docDefaults', NS)
        if doc_defaults is not None:
            rpr_default = doc_defaults.find('.//w:rPrDefault/w:rPr', NS)
            if rpr_default is not None:
                # 폰트
                fonts = rpr_default.find('w:rFonts', NS)
                if fonts is not None:
                    self.structure.default_font = (
                        fonts.get(f'{{{NS["w"]}}}ascii') or
                        fonts.get(f'{{{NS["w"]}}}hAnsi') or
                        fonts.get(f'{{{NS["w"]}}}eastAsia')
                    )

                # 크기
                sz = rpr_default.find('w:sz', NS)
                if sz is not None:
                    val = sz.get(f'{{{NS["w"]}}}val')
                    if val:
                        self.structure.default_font_size_pt = int(val) / 2

    def _analyze_styles_from_xml(self):
        """XML에서 스타일 정보 추출 (상속 관계 포함)"""
        if self._styles_xml is None:
            return

        root = self._styles_xml.getroot()

        for style_elem in root.findall('.//w:style', NS):
            style_id = style_elem.get(f'{{{NS["w"]}}}styleId')
            style_type = style_elem.get(f'{{{NS["w"]}}}type')

            if not style_id:
                continue

            # 이름
            name_elem = style_elem.find('w:name', NS)
            name = name_elem.get(f'{{{NS["w"]}}}val') if name_elem is not None else style_id

            # 기본 스타일
            based_on = style_elem.find('w:basedOn', NS)
            base_style = based_on.get(f'{{{NS["w"]}}}val') if based_on is not None else None

            style_info = StyleInfo(
                style_id=style_id,
                name=name,
                style_type=style_type or 'paragraph',
                base_style=base_style,
            )

            # 런 속성 (rPr) - 폰트, 크기, 색상 등
            rpr = style_elem.find('.//w:rPr', NS)
            if rpr is not None:
                self._parse_run_properties(rpr, style_info)

            # 문단 속성 (pPr) - 정렬, 간격 등
            ppr = style_elem.find('.//w:pPr', NS)
            if ppr is not None:
                self._parse_paragraph_properties(ppr, style_info)

            self.structure.styles[style_id] = style_info

        # 스타일 상속 해결 (부모 스타일에서 값 상속)
        self._resolve_style_inheritance()

    def _parse_run_properties(self, rpr: etree._Element, style_info: StyleInfo):
        """런 속성 파싱"""
        # 폰트
        fonts = rpr.find('w:rFonts', NS)
        if fonts is not None:
            style_info.font_name = (
                fonts.get(f'{{{NS["w"]}}}ascii') or
                fonts.get(f'{{{NS["w"]}}}hAnsi') or
                fonts.get(f'{{{NS["w"]}}}eastAsia')
            )

        # 크기
        sz = rpr.find('w:sz', NS)
        if sz is not None:
            val = sz.get(f'{{{NS["w"]}}}val')
            if val:
                style_info.font_size_pt = int(val) / 2

        # 볼드
        bold = rpr.find('w:b', NS)
        if bold is not None:
            val = bold.get(f'{{{NS["w"]}}}val')
            style_info.bold = val != '0' if val else True

        # 이탤릭
        italic = rpr.find('w:i', NS)
        if italic is not None:
            val = italic.get(f'{{{NS["w"]}}}val')
            style_info.italic = val != '0' if val else True

        # 색상
        color = rpr.find('w:color', NS)
        if color is not None:
            style_info.color_rgb = color.get(f'{{{NS["w"]}}}val')

    def _parse_paragraph_properties(self, ppr: etree._Element, style_info: StyleInfo):
        """문단 속성 파싱"""
        # 정렬
        jc = ppr.find('w:jc', NS)
        if jc is not None:
            style_info.alignment = jc.get(f'{{{NS["w"]}}}val')

        # 간격
        spacing = ppr.find('w:spacing', NS)
        if spacing is not None:
            before = spacing.get(f'{{{NS["w"]}}}before')
            if before:
                style_info.space_before_pt = int(before) / 20  # twips to pt

            after = spacing.get(f'{{{NS["w"]}}}after')
            if after:
                style_info.space_after_pt = int(after) / 20

            line = spacing.get(f'{{{NS["w"]}}}line')
            if line:
                style_info.line_spacing = int(line) / 240  # to lines

        # 들여쓰기
        ind = ppr.find('w:ind', NS)
        if ind is not None:
            left = ind.get(f'{{{NS["w"]}}}left')
            if left:
                style_info.left_indent_pt = int(left) / 20

        # 개요 수준 (outlineLvl) - CRITICAL: 이것이 진짜 헤딩 레벨 판별 기준
        outline_lvl = ppr.find('w:outlineLvl', NS)
        if outline_lvl is not None:
            val = outline_lvl.get(f'{{{NS["w"]}}}val')
            if val:
                style_info.outline_level = int(val)

    def _resolve_style_inheritance(self):
        """스타일 상속 해결"""
        def get_inherited_value(style_id: str, attr: str, visited: set = None):
            if visited is None:
                visited = set()

            if style_id in visited:
                return None
            visited.add(style_id)

            style = self.structure.styles.get(style_id)
            if not style:
                return None

            value = getattr(style, attr, None)
            if value is not None:
                return value

            if style.base_style:
                return get_inherited_value(style.base_style, attr, visited)

            return None

        # 각 스타일에 대해 상속값 적용
        attrs_to_inherit = [
            'font_name', 'font_size_pt', 'bold', 'italic', 'color_rgb',
            'alignment', 'space_before_pt', 'space_after_pt', 'line_spacing'
        ]

        for style_id, style in self.structure.styles.items():
            for attr in attrs_to_inherit:
                if getattr(style, attr) is None and style.base_style:
                    inherited = get_inherited_value(style.base_style, attr)
                    if inherited is not None:
                        setattr(style, attr, inherited)

    def _analyze_theme_colors(self):
        """테마 색상 분석"""
        if self._theme_xml is None:
            return

        root = self._theme_xml.getroot()
        theme_ns = {'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'}

        # 주요 색상 추출
        color_scheme = root.find('.//a:clrScheme', theme_ns)
        if color_scheme is not None:
            for color_elem in color_scheme:
                color_name = color_elem.tag.split('}')[-1]

                # srgbClr 또는 sysClr에서 색상값 추출
                srgb = color_elem.find('.//a:srgbClr', theme_ns)
                if srgb is not None:
                    self.structure.theme_colors[color_name] = srgb.get('val', '')

                sys_clr = color_elem.find('.//a:sysClr', theme_ns)
                if sys_clr is not None:
                    self.structure.theme_colors[color_name] = sys_clr.get('lastClr', '')

    def _extract_images(self):
        """이미지 추출"""
        self.output_dir.mkdir(parents=True, exist_ok=True)

        with zipfile.ZipFile(self.docx_path, 'r') as zf:
            # 이미지 파일 추출
            image_paths = {}
            for name in zf.namelist():
                if name.startswith('word/media/'):
                    image_name = os.path.basename(name)
                    output_path = self.output_dir / image_name
                    with zf.open(name) as src:
                        with open(output_path, 'wb') as dst:
                            dst.write(src.read())
                    image_paths[name] = str(output_path)

            # 헤더에서 이미지 정보 추출
            self._parse_header_footer_images(zf, 'header', image_paths)

            # 푸터에서 이미지 정보 추출
            self._parse_header_footer_images(zf, 'footer', image_paths)

    def _parse_header_footer_images(self, zf: zipfile.ZipFile, part_type: str, image_paths: dict):
        """헤더/푸터에서 이미지 정보 추출"""
        # 관련 파일 찾기
        for name in zf.namelist():
            if name.startswith(f'word/{part_type}') and name.endswith('.xml'):
                rels_path = name.replace('.xml', '.xml.rels').replace('word/', 'word/_rels/')

                # 관계 파일에서 이미지 매핑
                rels_map = {}
                if rels_path in zf.namelist():
                    with zf.open(rels_path) as f:
                        rels_tree = etree.parse(f)
                        for rel in rels_tree.getroot():
                            rel_id = rel.get('Id')
                            target = rel.get('Target')
                            if target and 'image' in target.lower():
                                full_path = f"word/{target.lstrip('../')}"
                                rels_map[rel_id] = full_path

                # XML 파싱하여 이미지 정보 추출
                with zf.open(name) as f:
                    tree = etree.parse(f)
                    root = tree.getroot()

                    for drawing in root.iter(f'{{{NS["w"]}}}drawing'):
                        self._parse_drawing(drawing, part_type, rels_map, image_paths)

    def _parse_drawing(self, drawing: etree._Element, position: str,
                       rels_map: dict, image_paths: dict):
        """drawing 요소에서 이미지 정보 추출"""
        # anchor 또는 inline 찾기
        anchor = drawing.find(f'.//{{{NS["wp"]}}}anchor')
        if anchor is None:
            anchor = drawing.find(f'.//{{{NS["wp"]}}}inline')

        if anchor is None:
            return

        # 크기 정보
        extent = anchor.find(f'{{{NS["wp"]}}}extent')
        width_emu = int(extent.get('cx', 0)) if extent is not None else 0
        height_emu = int(extent.get('cy', 0)) if extent is not None else 0

        # behindDoc 속성 (배경 여부)
        is_background = anchor.get('behindDoc') == '1'

        # 이미지 참조 ID
        blip = anchor.find(f'.//{{{NS["a"]}}}blip')
        if blip is not None:
            embed_id = blip.get(f'{{{NS["r"]}}}embed')
            if embed_id and embed_id in rels_map:
                original_path = rels_map[embed_id]
                extracted_path = image_paths.get(original_path, '')

                # 페이지 유형 추정
                is_full_page = width_emu > 7000000 and height_emu > 10000000
                if is_full_page:
                    page_type = 'cover'
                elif position == 'footer':
                    page_type = 'body'
                else:
                    page_type = 'header'

                img_info = ImageInfo(
                    name=os.path.basename(original_path),
                    original_path=original_path,
                    extracted_path=extracted_path,
                    width_emu=width_emu,
                    height_emu=height_emu,
                    position=position,
                    is_background=is_background,
                    rel_id=embed_id,
                    page_type=page_type if is_background else '',
                )

                self.structure.images.append(img_info)

    def _detect_page_templates(self):
        """페이지 템플릿 유형 감지"""
        # 문서 텍스트 분석으로 페이지 유형 감지
        doc_text = []
        for para in self.doc.paragraphs:
            doc_text.append({
                'text': para.text,
                'style': para.style.name if para.style else 'Normal'
            })

        # Cover 페이지 템플릿
        cover_bg = next((img for img in self.structure.images
                        if img.is_background and img.page_type == 'cover'), None)

        self.structure.page_templates['cover'] = PageTemplate(
            page_type='cover',
            background_image=cover_bg.extracted_path if cover_bg else None,
            styles_used=['Title', 'Subtitle'],
            description='표지 페이지 - 타이틀과 서브타이틀만 표시'
        )

        # Section 페이지 템플릿
        self.structure.page_templates['section'] = PageTemplate(
            page_type='section',
            styles_used=['heading 1'],
            description='섹션 구분 페이지 - 섹션 번호와 제목만 표시'
        )

        # Body 페이지 템플릿
        footer_img = next((img for img in self.structure.images
                          if img.position == 'footer'), None)
        header_img = next((img for img in self.structure.images
                          if img.position == 'header' and not img.is_background), None)

        self.structure.page_templates['body'] = PageTemplate(
            page_type='body',
            header_image=header_img.extracted_path if header_img else None,
            footer_image=footer_img.extracted_path if footer_img else None,
            styles_used=['heading 1', 'heading 2', 'heading 3', 'Body Text',
                        'List Bullet', 'List Number'],
            description='본문 페이지 - 헤딩, 본문, 리스트 등'
        )

    def get_style_by_name(self, name: str) -> Optional[StyleInfo]:
        """이름으로 스타일 찾기"""
        for style in self.structure.styles.values():
            if style.name.lower() == name.lower():
                return style
        return None

    def save_structure(self, output_path: Optional[str] = None) -> str:
        """구조 정보를 JSON으로 저장"""
        if output_path is None:
            output_path = self.output_dir / 'template_structure.json'

        data = {
            'file_path': self.structure.file_path,
            'page_width_inches': self.structure.page_width_inches,
            'page_height_inches': self.structure.page_height_inches,
            'margins': self.structure.margins,
            'default_font': self.structure.default_font,
            'default_font_size_pt': self.structure.default_font_size_pt,
            'theme_colors': self.structure.theme_colors,
            'styles': {k: v.to_dict() for k, v in self.structure.styles.items()},
            'images': [asdict(img) for img in self.structure.images],
            'page_templates': {k: asdict(v) for k, v in self.structure.page_templates.items()},
        }

        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)

        return str(output_path)

    def print_summary(self):
        """분석 결과 요약 출력"""
        s = self.structure
        print(f"\n{'='*60}")
        print(f"템플릿 분석: {Path(s.file_path).name}")
        print('='*60)

        print(f"\n📐 페이지 설정:")
        print(f"   크기: {s.page_width_inches:.2f}\" x {s.page_height_inches:.2f}\"")
        print(f"   여백: L={s.margins.get('left', 0):.2f}\", R={s.margins.get('right', 0):.2f}\", "
              f"T={s.margins.get('top', 0):.2f}\", B={s.margins.get('bottom', 0):.2f}\"")

        print(f"\n🔤 기본 폰트:")
        print(f"   {s.default_font or 'Theme Font'}, {s.default_font_size_pt or 11}pt")

        print(f"\n📝 주요 스타일 ({len(s.styles)}개):")
        key_styles = ['Title', 'Subtitle', 'heading 1', 'heading 2', 'heading 3',
                     'Normal', 'Body Text', 'List Bullet', 'List Number']
        for name in key_styles:
            style = self.get_style_by_name(name)
            if style:
                font = style.font_name or s.default_font or 'Default'
                size = style.font_size_pt or s.default_font_size_pt or 11
                color = f", color: #{style.color_rgb}" if style.color_rgb else ""
                bold = ", Bold" if style.bold else ""
                print(f"   [{name}] {font}, {size}pt{bold}{color}")

        print(f"\n🖼️ 이미지 ({len(s.images)}개):")
        for img in s.images:
            bg = " [배경]" if img.is_background else ""
            size = f"{img.width_inches:.2f}\" x {img.height_inches:.2f}\""
            print(f"   [{img.position}] {img.name} ({size}){bg}")

        print(f"\n📄 페이지 템플릿:")
        for name, tmpl in s.page_templates.items():
            print(f"   [{name}] {tmpl.description}")
            if tmpl.background_image:
                print(f"      배경: {Path(tmpl.background_image).name}")


if __name__ == '__main__':
    import sys

    template_path = '/home/shaush/md-to-docx/docx_only/[Word 템플릿] A4.docx'
    if len(sys.argv) > 1:
        template_path = sys.argv[1]

    analyzer = DocxTemplateAnalyzer(template_path)
    analyzer.analyze()
    analyzer.print_summary()
    output = analyzer.save_structure()
    print(f"\n✅ 저장: {output}")