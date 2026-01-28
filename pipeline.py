"""
문서 자동화 파이프라인

Markdown + DOCX Template → Final DOCX

파이프라인 흐름:
1. 템플릿 파싱 (플레이스홀더 추출)
2. 마크다운 파싱 (콘텐츠 블록 추출)
3. 콘텐츠 매핑 (LLM 또는 규칙 기반)
4. DOCX 조립 (플레이스홀더 교체)
"""

import asyncio
from pathlib import Path
from typing import Optional
from dataclasses import dataclass

from src.template_parser import TemplateParser
from src.markdown_parser import MarkdownParser, DocumentStructure
from src.llm_content_mapper import LLMContentMapper, ContentMapperSync
from src.docx_composer import DocxComposer
from src.models import ParsedTemplate, ContentMappingPlan


@dataclass
class PipelineConfig:
    """파이프라인 설정"""
    use_llm: bool = False
    vllm_base_url: str = "http://localhost:8000/v1"
    vllm_model: str = "Qwen/Qwen2.5-7B-Instruct"
    placeholder_pattern: str = "default"  # default, bracket, angle, underscore


@dataclass
class PipelineResult:
    """파이프라인 결과"""
    output_path: str
    template_info: ParsedTemplate
    content_info: DocumentStructure
    mapping_plan: ContentMappingPlan
    success: bool = True
    error: Optional[str] = None


class DocumentAutomationPipeline:
    """
    문서 자동화 파이프라인

    마크다운 콘텐츠와 DOCX 템플릿을 입력받아
    최종 DOCX 파일을 생성합니다.
    """

    def __init__(self, config: Optional[PipelineConfig] = None):
        """
        Args:
            config: 파이프라인 설정
        """
        self.config = config or PipelineConfig()

    async def process_async(
        self,
        markdown_path: str,
        template_path: str,
        output_path: Optional[str] = None,
    ) -> PipelineResult:
        """
        비동기 처리 (LLM 사용 시)

        Args:
            markdown_path: 마크다운 파일 경로
            template_path: 템플릿 DOCX 파일 경로
            output_path: 출력 파일 경로 (선택)

        Returns:
            PipelineResult 객체
        """
        try:
            # 1. 템플릿 파싱
            template_parser = TemplateParser(
                template_path,
                placeholder_pattern=self.config.placeholder_pattern
            )
            template_info = template_parser.parse()

            print(f"[1/4] 템플릿 파싱 완료: {len(template_info.placeholders)}개 플레이스홀더 발견")

            # 2. 마크다운 파싱
            md_parser = MarkdownParser()
            content_info = md_parser.parse_file(markdown_path)

            print(f"[2/4] 마크다운 파싱 완료: {len(content_info.raw_blocks)}개 콘텐츠 블록")

            # 3. 콘텐츠 매핑
            if self.config.use_llm:
                async with LLMContentMapper(
                    base_url=self.config.vllm_base_url,
                    model=self.config.vllm_model,
                    use_llm=True,
                ) as mapper:
                    mapping_plan = await mapper.create_mapping_plan(template_info, content_info)
            else:
                mapper = ContentMapperSync()
                mapping_plan = mapper.create_mapping_plan(template_info, content_info)

            print(f"[3/4] 매핑 완료: {len(mapping_plan.mappings)}개 매핑, 신뢰도: {mapping_plan.confidence:.2f}")

            # 4. DOCX 조립
            output_dir = Path(output_path).parent if output_path else None
            composer = DocxComposer(
                template_path,
                output_dir=str(output_dir) if output_dir else None
            )

            final_output = composer.compose(
                mapping_plan=mapping_plan,
                content=content_info,
                output_filename=Path(output_path).name if output_path else None,
            )

            print(f"[4/4] DOCX 생성 완료: {final_output}")

            return PipelineResult(
                output_path=final_output,
                template_info=template_info,
                content_info=content_info,
                mapping_plan=mapping_plan,
                success=True,
            )

        except Exception as e:
            return PipelineResult(
                output_path="",
                template_info=ParsedTemplate(file_path=""),
                content_info=DocumentStructure(),
                mapping_plan=ContentMappingPlan(),
                success=False,
                error=str(e),
            )

    def process(
        self,
        markdown_path: str,
        template_path: str,
        output_path: Optional[str] = None,
    ) -> PipelineResult:
        """
        동기 처리

        Args:
            markdown_path: 마크다운 파일 경로
            template_path: 템플릿 DOCX 파일 경로
            output_path: 출력 파일 경로 (선택)

        Returns:
            PipelineResult 객체
        """
        return asyncio.run(self.process_async(markdown_path, template_path, output_path))


def run_pipeline(
    markdown_path: str,
    template_path: str,
    output_path: Optional[str] = None,
    use_llm: bool = False,
    vllm_base_url: Optional[str] = None,
    vllm_model: Optional[str] = None,
) -> str:
    """
    편의 함수: 파이프라인 실행

    Args:
        markdown_path: 마크다운 파일 경로
        template_path: 템플릿 DOCX 파일 경로
        output_path: 출력 파일 경로
        use_llm: LLM 사용 여부
        vllm_base_url: vLLM 서버 URL
        vllm_model: 모델 이름

    Returns:
        생성된 파일 경로
    """
    config = PipelineConfig(
        use_llm=use_llm,
        vllm_base_url=vllm_base_url or "http://localhost:8000/v1",
        vllm_model=vllm_model or "Qwen/Qwen2.5-7B-Instruct",
    )

    pipeline = DocumentAutomationPipeline(config)
    result = pipeline.process(markdown_path, template_path, output_path)

    if result.success:
        return result.output_path
    else:
        raise RuntimeError(f"Pipeline failed: {result.error}")


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(
        description="문서 자동화 파이프라인: Markdown + Template → DOCX"
    )
    parser.add_argument("markdown", help="마크다운 파일 경로")
    parser.add_argument("template", help="템플릿 DOCX 파일 경로")
    parser.add_argument("-o", "--output", help="출력 파일 경로")
    parser.add_argument("--llm", action="store_true", help="LLM 사용 (vLLM 서버 필요)")
    parser.add_argument("--vllm-url", default="http://localhost:8000/v1", help="vLLM 서버 URL")
    parser.add_argument("--model", default="Qwen/Qwen2.5-7B-Instruct", help="모델 이름")

    args = parser.parse_args()

    try:
        result = run_pipeline(
            markdown_path=args.markdown,
            template_path=args.template,
            output_path=args.output,
            use_llm=args.llm,
            vllm_base_url=args.vllm_url,
            vllm_model=args.model,
        )
        print(f"\n✅ 생성 완료: {result}")
    except Exception as e:
        print(f"\n❌ 오류 발생: {e}")
        exit(1)
