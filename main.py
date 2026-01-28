#!/usr/bin/env python3
"""
MD → DOCX 변환기

사용법:
    # 기본 변환 (기존 방식)
    uv run python main.py input.md output.docx -t template.docx

    # 플레이스홀더 기반 변환 (신규)
    uv run python main.py --pipeline input.md -t template_with_placeholders.docx -o output.docx

    # LLM 기반 매핑 사용
    uv run python main.py --pipeline input.md -t template.docx -o output.docx --llm

    # 디렉토리 일괄 변환
    uv run python main.py input_dir/ output_dir/ -t template.docx

    # 템플릿 분석 (플레이스홀더 확인)
    uv run python main.py --analyze template.docx

    # 마크다운 분석
    uv run python main.py --parse input.md

"""

import argparse
import sys
from pathlib import Path

# src 경로 추가
sys.path.insert(0, str(Path(__file__).parent))

from src.template_analyzer import DocxTemplateAnalyzer
from src.markdown_parser import MarkdownParser
from src.docx_generator import DocxGenerator


def analyze_template(template_path: str, show_placeholders: bool = True):
    """템플릿 분석"""
    analyzer = DocxTemplateAnalyzer(template_path)
    analyzer.analyze()
    analyzer.print_summary()
    output = analyzer.save_structure()
    print(f"\n📁 에셋 저장: {analyzer.output_dir}")

    # 플레이스홀더 분석 추가
    if show_placeholders:
        try:
            from src.template_parser import TemplateParser
            parser = TemplateParser(template_path)
            result = parser.parse()

            print(f"\n📌 플레이스홀더 분석:")
            if result.placeholders:
                for p in result.placeholders:
                    print(f"   - {p.id} ({p.placeholder_type}) @ 문단 {p.paragraph_index}")
            else:
                print("   플레이스홀더 없음. {{TITLE}}, {{BODY}} 등을 템플릿에 추가하세요.")
        except Exception as e:
            print(f"   플레이스홀더 분석 실패: {e}")


def parse_markdown(md_path: str):
    """마크다운 파싱 분석"""
    parser = MarkdownParser()
    doc = parser.parse_file(md_path)

    print(f"\n📄 Title: {doc.title}")
    print(f"📝 Subtitle: {doc.subtitle}")
    print(f"🖼️ First Image: {doc.first_image_path}")
    print(f"\n📚 Sections: {len(doc.sections)}")

    for i, section in enumerate(doc.sections[:5]):
        print(f"\n  Section {i+1}: {section.heading or '(no heading)'}")
        print(f"    Blocks: {len(section.blocks)}")

    print(f"\n📊 Total blocks: {len(doc.raw_blocks)}")
    block_types = {}
    for b in doc.raw_blocks:
        block_types[b.block_type] = block_types.get(b.block_type, 0) + 1
    for bt, count in sorted(block_types.items(), key=lambda x: -x[1]):
        print(f"    {bt}: {count}")


def convert_file(md_path: str, output_path: str, template_path: str = None):
    """단일 파일 변환"""
    generator = DocxGenerator(template_path)
    result = generator.generate_from_file(md_path, output_path)
    print(f"✅ {Path(md_path).name} → {Path(output_path).name}")
    return result


def convert_directory(input_dir: str, output_dir: str, template_path: str = None):
    """디렉토리 일괄 변환"""
    input_path = Path(input_dir)
    output_path = Path(output_dir)
    output_path.mkdir(parents=True, exist_ok=True)

    md_files = list(input_path.glob('*.md'))
    results = []

    for md_file in md_files:
        output_file = output_path / f"{md_file.stem}.docx"
        try:
            result = convert_file(str(md_file), str(output_file), template_path)
            results.append(result)
        except Exception as e:
            print(f"❌ {md_file.name}: {e}")

    print(f"\n📊 변환 완료: {len(results)}/{len(md_files)}개")
    return results


def run_pipeline_mode(args):
    """플레이스홀더 기반 파이프라인 모드"""
    from pipeline import run_pipeline

    if not args.template:
        print("❌ 파이프라인 모드에서는 --template (-t) 옵션이 필수입니다.")
        return

    output_path = args.output or (Path(args.input).stem + "_output.docx")

    print(f"\n🔄 파이프라인 모드 실행")
    print(f"   마크다운: {args.input}")
    print(f"   템플릿: {args.template}")
    print(f"   출력: {output_path}")
    print(f"   LLM 사용: {args.llm}")
    print()

    try:
        result = run_pipeline(
            markdown_path=args.input,
            template_path=args.template,
            output_path=output_path,
            use_llm=args.llm,
            vllm_base_url=args.vllm_url,
            vllm_model=args.model,
        )
        print(f"\n✅ 생성 완료: {result}")
    except Exception as e:
        print(f"\n❌ 오류: {e}")


def main():
    import time
    s = time.perf_counter()
    parser = argparse.ArgumentParser(
        description='마크다운 → DOCX 변환기',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
예시:
  # 기본 변환
  uv run python main.py input.md output.docx -t template.docx

  # 플레이스홀더 기반 파이프라인 (신규)
  uv run python main.py --pipeline input.md -t template.docx -o output.docx

  # LLM 매핑 사용
  uv run python main.py --pipeline input.md -t template.docx --llm

  # 디렉토리 일괄 변환
  uv run python main.py input_dir/ output_dir/ -t template.docx

  # 템플릿 분석
  uv run python main.py --analyze template.docx

  # 마크다운 분석
  uv run python main.py --parse input.md
        """
    )

    parser.add_argument('input', nargs='?', help='입력 마크다운 파일 또는 디렉토리')
    parser.add_argument('output', nargs='?', help='출력 DOCX 파일 또는 디렉토리')
    parser.add_argument('-t', '--template', help='DOCX 템플릿 파일')
    parser.add_argument('-o', '--out', dest='output_alt', help='출력 파일 경로 (--pipeline 모드용)')
    parser.add_argument('--analyze', metavar='DOCX', help='템플릿 분석 모드')
    parser.add_argument('--parse', metavar='MD', help='마크다운 분석 모드')

    # 파이프라인 모드 옵션
    parser.add_argument('--pipeline', action='store_true', help='플레이스홀더 기반 파이프라인 모드')
    parser.add_argument('--llm', action='store_true', help='LLM 매핑 사용 (vLLM 서버 필요)')
    parser.add_argument('--vllm-url', default='http://localhost:8000/v1', help='vLLM 서버 URL')
    parser.add_argument('--model', default='Qwen/Qwen2.5-7B-Instruct', help='LLM 모델')

    args = parser.parse_args()

    # output 우선순위: output > output_alt
    if args.output_alt and not args.output:
        args.output = args.output_alt

    # 템플릿 분석 모드
    if args.analyze:
        analyze_template(args.analyze)
        print(f"\n⏱️ 소요 시간: {time.perf_counter()-s:.2f}s")
        return

    # 마크다운 분석 모드
    if args.parse:
        parse_markdown(args.parse)
        print(f"\n⏱️ 소요 시간: {time.perf_counter()-s:.2f}s")
        return

    # 파이프라인 모드
    if args.pipeline:
        if not args.input:
            print("❌ 파이프라인 모드에서는 입력 파일이 필요합니다.")
            parser.print_help()
            return
        run_pipeline_mode(args)
        print(f"\n⏱️ 소요 시간: {time.perf_counter()-s:.2f}s")
        return

    # 기본 변환 모드
    if not args.input:
        parser.print_help()
        return

    input_path = Path(args.input)
    output_path = args.output

    if input_path.is_dir():
        if not output_path:
            output_path = str(input_path) + '_converted'
        convert_directory(str(input_path), output_path, args.template)
    else:
        if not output_path:
            output_path = input_path.stem + '.docx'
        convert_file(str(input_path), output_path, args.template)

    print(f"\n⏱️ 소요 시간: {time.perf_counter()-s:.2f}s")


if __name__ == '__main__':
    main()
