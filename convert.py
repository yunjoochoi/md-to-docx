#!/usr/bin/env python3
"""
마크다운 → DOCX 변환기 (outlineLvl 기반)

사용법:
    python convert.py <마크다운파일> <출력DOCX> [템플릿DOCX]

예제:
    python convert.py input.md output.docx
    python convert.py "/home/shaush/work/output_md/20251128_company_554088000.md" output.docx "/home/shaush/md-to-docx/[Word템플릿]A4.docx"
"""

import sys
from pathlib import Path
from src.docx_generator import DocxGenerator


def main():
    if len(sys.argv) < 3:
        print(__doc__)
        print("\n오류: 인자가 부족합니다.")
        print("사용법: python convert.py <마크다운파일> <출력DOCX> [템플릿DOCX]")
        sys.exit(1)

    md_file = sys.argv[1]
    output_path = sys.argv[2]
    template_path = sys.argv[3] if len(sys.argv) > 3 else None

    # 입력 파일 확인
    if not Path(md_file).exists():
        print(f"오류: 마크다운 파일을 찾을 수 없습니다: {md_file}")
        sys.exit(1)

    if template_path and not Path(template_path).exists():
        print(f"오류: 템플릿 파일을 찾을 수 없습니다: {template_path}")
        sys.exit(1)

    # 변환 실행
    print(f"마크다운: {md_file}")
    if template_path:
        print(f"템플릿: {template_path}")
    print(f"출력: {output_path}")
    print()

    generator = DocxGenerator(template_path)
    result = generator.generate_from_file(md_file, output_path)

    print(f"\n변환 완료: {result}")


if __name__ == '__main__':
    main()
