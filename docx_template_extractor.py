"""
DOCX 템플릿 추출기 (배경이미지, 스타일 정보)

- 헤더/푸터의 배경 이미지 추출
- 표지용/본문용 배경 구분
- 스타일 정보 (폰트, 크기, 색상) 추출
- 새 문서 생성 시 템플릿 정보 적용
"""

from docx import Document
from docx.shared import Pt, Inches, Emu
from docx.oxml.ns import qn, nsmap
from docx.enum.style import WD_STYLE_TYPE
from lxml import etree
import zipfile
import os
import json
import shutil
from pathlib import Path
from dataclasses import dataclass, field, asdict
from typing import Optional, Dict, List


# XML namespaces
NAMESPACES = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'pic': 'http://schemas.openxmlformats.org/drawingml/2006/picture',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
}


@dataclass
class ImageInfo:
    """이미지 정보"""
    name: str
    original_path: str  # word/media/image1.png
    width_emu: int = 0
    height_emu: int = 0
    position: str = ''  # 'header', 'footer', 'body'
    is_background: bool = False  # behindDoc="1" 여부
    rel_id: str = ''  # rId

    @property
    def width_inches(self):
        return self.width_emu / 914400 if self.width_emu else 0

    @property
    def height_inches(self):
        return self.height_emu / 914400 if self.height_emu else 0


@dataclass
class StyleInfo:
    """스타일 정보"""
    name: str
    font_name: Optional[str] = None
    font_size_pt: Optional[float] = None
    bold: Optional[bool] = None
    italic: Optional[bool] = None
    color_rgb: Optional[str] = None
    alignment: Optional[str] = None
    space_before_pt: Optional[float] = None
    space_after_pt: Optional[float] = None


@dataclass
class TemplateInfo:
    """템플릿 전체 정보"""
    file_path: str
    page_width_inches: float = 8.27  # A4 default
    page_height_inches: float = 11.69
    margins: Dict[str, float] = field(default_factory=dict)
    styles: Dict[str, StyleInfo] = field(default_factory=dict)
    images: List[ImageInfo] = field(default_factory=list)
    background_image: Optional[str] = None  # 메인 배경 이미지 경로
    header_logo: Optional[str] = None  # 헤더 로고 이미지 경로
    footer_image: Optional[str] = None  # 푸터 이미지 경로


class DocxTemplateExtractor:
    """DOCX 템플릿에서 스타일과 배경 이미지를 추출"""

    def __init__(self, docx_path: str):
        self.docx_path = Path(docx_path)
        self.output_dir = self.docx_path.parent / f"{self.docx_path.stem}_assets"
        self.doc = Document(docx_path)
        self.template_info = TemplateInfo(file_path=str(docx_path))

    def extract_all(self) -> TemplateInfo:
        """모든 정보 추출"""
        self._extract_page_setup()
        self._extract_styles()
        self._extract_images_with_context()
        return self.template_info

    def _extract_page_setup(self):
        """페이지 설정 추출"""
        section = self.doc.sections[0]
        self.template_info.page_width_inches = section.page_width.inches
        self.template_info.page_height_inches = section.page_height.inches
        self.template_info.margins = {
            'left': section.left_margin.inches,
            'right': section.right_margin.inches,
            'top': section.top_margin.inches,
            'bottom': section.bottom_margin.inches,
        }

    def _extract_styles(self):
        """스타일 정보 추출"""
        key_styles = ['Title', 'Heading 1', 'Heading 2', 'Heading 3', 'Normal', 'Body Text', 'List Bullet', 'List Number']

        for style in self.doc.styles:
            if style.type == WD_STYLE_TYPE.PARAGRAPH and style.name in key_styles:
                style_info = StyleInfo(name=style.name)

                if style.font:
                    if style.font.name:
                        style_info.font_name = style.font.name
                    if style.font.size:
                        style_info.font_size_pt = style.font.size.pt
                    if style.font.bold is not None:
                        style_info.bold = style.font.bold
                    if style.font.italic is not None:
                        style_info.italic = style.font.italic
                    if style.font.color and style.font.color.rgb:
                        style_info.color_rgb = str(style.font.color.rgb)

                if style.paragraph_format:
                    pf = style.paragraph_format
                    if pf.alignment:
                        style_info.alignment = str(pf.alignment)
                    if pf.space_before:
                        style_info.space_before_pt = pf.space_before.pt
                    if pf.space_after:
                        style_info.space_after_pt = pf.space_after.pt

                self.template_info.styles[style.name] = style_info

    def _extract_images_with_context(self):
        """이미지와 위치 정보 추출 (XML 파싱)"""
        self.output_dir.mkdir(parents=True, exist_ok=True)

        with zipfile.ZipFile(self.docx_path, 'r') as zf:
            # 이미지 파일 추출
            image_files = {}
            for name in zf.namelist():
                if name.startswith('word/media/'):
                    image_name = os.path.basename(name)
                    output_path = self.output_dir / image_name
                    with zf.open(name) as src:
                        with open(output_path, 'wb') as dst:
                            dst.write(src.read())
                    image_files[name] = str(output_path)

            # 헤더에서 이미지 정보 추출
            self._parse_header_images(zf, image_files)

            # 푸터에서 이미지 정보 추출
            self._parse_footer_images(zf, image_files)

    def _parse_header_images(self, zf: zipfile.ZipFile, image_files: dict):
        """헤더 XML에서 이미지 정보 파싱"""
        # header1.xml.rels에서 이미지 매핑 가져오기
        rels_map = {}
        if 'word/_rels/header1.xml.rels' in zf.namelist():
            with zf.open('word/_rels/header1.xml.rels') as f:
                rels_tree = etree.parse(f)
                for rel in rels_tree.getroot():
                    rel_id = rel.get('Id')
                    target = rel.get('Target')
                    if target and 'image' in target.lower():
                        rels_map[rel_id] = f"word/{target.lstrip('../')}"

        # header1.xml 파싱
        if 'word/header1.xml' in zf.namelist():
            with zf.open('word/header1.xml') as f:
                tree = etree.parse(f)
                root = tree.getroot()

                # 모든 drawing 요소 찾기
                for drawing in root.iter('{%s}drawing' % NAMESPACES['w']):
                    self._parse_drawing_element(drawing, 'header', rels_map, image_files)

    def _parse_footer_images(self, zf: zipfile.ZipFile, image_files: dict):
        """푸터 XML에서 이미지 정보 파싱"""
        rels_map = {}
        if 'word/_rels/footer2.xml.rels' in zf.namelist():
            with zf.open('word/_rels/footer2.xml.rels') as f:
                rels_tree = etree.parse(f)
                for rel in rels_tree.getroot():
                    rel_id = rel.get('Id')
                    target = rel.get('Target')
                    if target and 'image' in target.lower():
                        rels_map[rel_id] = f"word/{target.lstrip('../')}"

        if 'word/footer2.xml' in zf.namelist():
            with zf.open('word/footer2.xml') as f:
                tree = etree.parse(f)
                root = tree.getroot()

                for drawing in root.iter('{%s}drawing' % NAMESPACES['w']):
                    self._parse_drawing_element(drawing, 'footer', rels_map, image_files)

    def _parse_drawing_element(self, drawing, position: str, rels_map: dict, image_files: dict):
        """drawing 요소에서 이미지 정보 추출"""
        # anchor 또는 inline 찾기
        anchor = drawing.find('.//{%s}anchor' % NAMESPACES['wp'])
        if anchor is None:
            anchor = drawing.find('.//{%s}inline' % NAMESPACES['wp'])

        if anchor is None:
            return

        # 크기 정보
        extent = anchor.find('{%s}extent' % NAMESPACES['wp'])
        width_emu = int(extent.get('cx', 0)) if extent is not None else 0
        height_emu = int(extent.get('cy', 0)) if extent is not None else 0

        # behindDoc 속성 (배경 여부)
        is_background = anchor.get('behindDoc') == '1'

        # 이미지 참조 ID 찾기
        blip = anchor.find('.//{%s}blip' % NAMESPACES['a'])
        if blip is not None:
            embed_id = blip.get('{%s}embed' % NAMESPACES['r'])
            if embed_id and embed_id in rels_map:
                original_path = rels_map[embed_id]
                extracted_path = image_files.get(original_path, '')

                img_info = ImageInfo(
                    name=os.path.basename(original_path),
                    original_path=original_path,
                    width_emu=width_emu,
                    height_emu=height_emu,
                    position=position,
                    is_background=is_background,
                    rel_id=embed_id
                )

                self.template_info.images.append(img_info)

                # 특별 분류 (크기 기준으로 배경 이미지 판별)
                # A4 전체 배경: 약 8.27" x 11.69" = 7562088 x 10689336 EMU
                is_full_page = width_emu > 7000000 and height_emu > 10000000

                if is_background and is_full_page:
                    # 전체 페이지 배경 이미지
                    self.template_info.background_image = extracted_path
                elif position == 'header' and not is_background:
                    self.template_info.header_logo = extracted_path
                elif position == 'footer' and is_background:
                    self.template_info.footer_image = extracted_path

    def save_template_info(self, output_path: Optional[str] = None):
        """템플릿 정보를 JSON으로 저장"""
        if output_path is None:
            output_path = self.output_dir / 'template_info.json'

        # dataclass를 dict로 변환
        data = {
            'file_path': self.template_info.file_path,
            'page_width_inches': self.template_info.page_width_inches,
            'page_height_inches': self.template_info.page_height_inches,
            'margins': self.template_info.margins,
            'styles': {k: asdict(v) for k, v in self.template_info.styles.items()},
            'images': [asdict(img) for img in self.template_info.images],
            'background_image': self.template_info.background_image,
            'header_logo': self.template_info.header_logo,
            'footer_image': self.template_info.footer_image,
        }

        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)

        return output_path

    def print_summary(self):
        """추출 결과 요약 출력"""
        t = self.template_info
        print(f"\n{'='*60}")
        print(f"템플릿 분석 결과: {t.file_path}")
        print('='*60)

        print(f"\n📐 페이지 설정:")
        print(f"   크기: {t.page_width_inches:.2f}\" x {t.page_height_inches:.2f}\"")
        print(f"   여백: L={t.margins.get('left', 0):.2f}\", R={t.margins.get('right', 0):.2f}\", "
              f"T={t.margins.get('top', 0):.2f}\", B={t.margins.get('bottom', 0):.2f}\"")

        print(f"\n📝 스타일 정보:")
        for name, style in t.styles.items():
            font_info = f"{style.font_name or 'N/A'}, {style.font_size_pt or 'N/A'}pt"
            bold_info = ", Bold" if style.bold else ""
            print(f"   [{name}] {font_info}{bold_info}")

        print(f"\n🖼️ 이미지 정보:")
        for img in t.images:
            size = f"{img.width_inches:.2f}\" x {img.height_inches:.2f}\""
            bg = " [배경]" if img.is_background else ""
            print(f"   [{img.position}] {img.name} ({size}){bg}")

        print(f"\n🎨 분류된 이미지:")
        print(f"   배경 이미지: {t.background_image or 'None'}")
        print(f"   헤더 로고: {t.header_logo or 'None'}")
        print(f"   푸터 이미지: {t.footer_image or 'None'}")


def main():
    import sys

    template_path = '/home/shaush/md-to-docx/docx_only/[Word 템플릿] A4.docx'
    if len(sys.argv) > 1:
        template_path = sys.argv[1]

    extractor = DocxTemplateExtractor(template_path)
    extractor.extract_all()
    extractor.print_summary()

    output_json = extractor.save_template_info()
    print(f"\n✅ 템플릿 정보 저장: {output_json}")


if __name__ == '__main__':
    main()
