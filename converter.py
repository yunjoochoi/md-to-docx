"""
통합 MD → DOCX 변환기

사용법:
    # 기본 변환 (템플릿 없이)
    python converter.py input.md output.docx

    # 템플릿 사용
    python converter.py input.md output.docx --template template.docx

    # 템플릿 분석만
    uv run python converter.py --analyze "/home/shaush/md-to-docx/[Word템플릿]A4.docx"
    
    
    # 디렉토리 일괄 변환
    python converter.py input_dir/ output_dir/ --template /home/shaush/md-to-docx/[Word템플릿]A4.docx
"""

import argparse
from pathlib import Path
from typing import Optional, List
import json

from docx_template_extractor import DocxTemplateExtractor, TemplateInfo
from md_to_docx_converter import MarkdownToDocxConverter


class IntegratedConverter:
    """통합 변환기"""

    def __init__(self, template_path: Optional[str] = None):
        self.template_path = template_path
        self.template_info: Optional[TemplateInfo] = None

        if template_path and Path(template_path).exists():
            extractor = DocxTemplateExtractor(template_path)
            self.template_info = extractor.extract_all()

    def convert_file(self, md_path: str, output_path: str) -> str:
        """단일 파일 변환"""
        converter = MarkdownToDocxConverter(self.template_path)
        return converter.convert_file(md_path, output_path)

    def convert_directory(self, input_dir: str, output_dir: str) -> List[str]:
        """디렉토리 내 모든 .md 파일 변환"""
        input_path = Path(input_dir)
        output_path = Path(output_dir)
        output_path.mkdir(parents=True, exist_ok=True)

        results = []
        md_files = list(input_path.glob('*.md'))

        for md_file in md_files:
            output_file = output_path / f"{md_file.stem}.docx"
            try:
                result = self.convert_file(str(md_file), str(output_file))
                results.append(result)
                print(f"✅ {md_file.name} → {output_file.name}")
            except Exception as e:
                print(f"❌ {md_file.name}: {e}")

        return results

    def get_template_summary(self) -> dict:
        """템플릿 정보 요약"""
        if not self.template_info:
            return {}

        return {
            'page_size': f"{self.template_info.page_width_inches:.2f}\" x {self.template_info.page_height_inches:.2f}\"",
            'margins': self.template_info.margins,
            'styles': list(self.template_info.styles.keys()),
            'background_image': self.template_info.background_image,
            'header_logo': self.template_info.header_logo,
            'footer_image': self.template_info.footer_image,
        }


def analyze_template(template_path: str):
    """템플릿 분석 및 출력"""
    extractor = DocxTemplateExtractor(template_path)
    extractor.extract_all()
    extractor.print_summary()
    output_json = extractor.save_template_info()
    print(f"\n📁 에셋 저장 위치: {extractor.output_dir}")


def main():
    parser = argparse.ArgumentParser(
        description='마크다운을 DOCX로 변환',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
예시:
  # 단일 파일 변환
  python converter.py input.md output.docx

  # 템플릿 사용
  python converter.py input.md output.docx -t template.docx

  # 디렉토리 일괄 변환
  python converter.py input_dir/ output_dir/ -t template.docx

  # 템플릿 분석
  python converter.py --analyze template.docx
        """
    )

    parser.add_argument('input', nargs='?', help='입력 마크다운 파일 또는 디렉토리')
    parser.add_argument('output', nargs='?', help='출력 DOCX 파일 또는 디렉토리')
    parser.add_argument('-t', '--template', help='DOCX 템플릿 파일')
    parser.add_argument('--analyze', metavar='DOCX', help='템플릿 분석 모드')

    args = parser.parse_args()

    # 템플릿 분석 모드
    if args.analyze:
        analyze_template(args.analyze)
        return

    # 변환 모드
    if not args.input:
        parser.print_help()
        return

    input_path = Path(args.input)
    output_path = args.output

    converter = IntegratedConverter(args.template)

    if input_path.is_dir():
        # 디렉토리 일괄 변환
        if not output_path:
            output_path = str(input_path) + '_converted'
        results = converter.convert_directory(str(input_path), output_path)
        print(f"\n📊 변환 완료: {len(results)}개 파일")
    else:
        # 단일 파일 변환
        if not output_path:
            output_path = input_path.stem + '.docx'
        result = converter.convert_file(str(input_path), output_path)
        print(f"✅ 변환 완료: {result}")


if __name__ == '__main__':
    main()
