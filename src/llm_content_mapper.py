"""
LLM 기반 콘텐츠 매핑

마크다운 콘텐츠를 템플릿 플레이스홀더에 자동으로 매핑
- LLM 모드: vLLM/Qwen을 사용한 지능형 매핑
- 자동 모드: 규칙 기반 매핑 (LLM 없이)
"""

import re
from typing import List, Optional, Dict, Any
from dataclasses import asdict

from .models import (
    ContentMappingPlan, ContentMapping, ParsedTemplate,
    PlaceholderType, Placeholder
)
from .markdown_parser import ContentBlock, DocumentStructure

import sys
sys.path.insert(0, str(__file__).rsplit('/src/', 1)[0])
from llm.vllm_client import VLLMClient
from llm.prompts import (
    CONTENT_MAPPING_SYSTEM_PROMPT,
    build_mapping_prompt,
    get_auto_mapping_rule,
)


class LLMContentMapper:
    """
    LLM 기반 콘텐츠 매핑 생성기

    플레이스홀더와 마크다운 콘텐츠를 분석하여
    최적의 매핑 계획을 생성합니다.
    """

    def __init__(
        self,
        base_url: Optional[str] = None,
        model: Optional[str] = None,
        use_llm: bool = True,
    ):
        """
        Args:
            base_url: vLLM 서버 URL
            model: 모델 이름
            use_llm: LLM 사용 여부 (False면 규칙 기반 매핑)
        """
        self.use_llm = use_llm
        self._client: Optional[VLLMClient] = None

        if use_llm:
            self._client = VLLMClient(base_url=base_url, model=model)

    async def __aenter__(self):
        return self

    async def __aexit__(self, exc_type, exc_val, exc_tb):
        if self._client:
            await self._client.close()

    async def create_mapping_plan(
        self,
        template: ParsedTemplate,
        content: DocumentStructure,
    ) -> ContentMappingPlan:
        """
        매핑 계획 생성

        Args:
            template: 파싱된 템플릿 (플레이스홀더 포함)
            content: 파싱된 마크다운 문서

        Returns:
            ContentMappingPlan 객체
        """
        if not template.placeholders:
            return ContentMappingPlan(
                warnings=["No placeholders found in template"],
                confidence=0.0
            )

        if not content.raw_blocks:
            return ContentMappingPlan(
                warnings=["No content blocks found in markdown"],
                confidence=0.0
            )

        if self.use_llm and self._client:
            return await self._create_mapping_with_llm(template, content)
        else:
            return self._create_mapping_auto(template, content)

    async def _create_mapping_with_llm(
        self,
        template: ParsedTemplate,
        content: DocumentStructure,
    ) -> ContentMappingPlan:
        """LLM을 사용한 매핑 생성"""
        # 플레이스홀더 정보 준비
        placeholders_data = [
            {
                "id": p.id,
                "type": p.placeholder_type,
                "section": p.section_type,
                "style": p.style_name,
            }
            for p in template.placeholders
        ]

        # 콘텐츠 블록 정보 준비
        content_data = [
            {
                "block_type": block.block_type,
                "content": block.content,
                "level": block.level,
                "list_type": block.list_type,
            }
            for block in content.raw_blocks
        ]

        # 프롬프트 생성
        user_prompt = build_mapping_prompt(placeholders_data, content_data)

        # LLM 호출
        try:
            response = await self._client.get_content_mapping(
                system_prompt=CONTENT_MAPPING_SYSTEM_PROMPT,
                user_prompt=user_prompt,
            )

            return self._parse_llm_response(response, template, content)

        except Exception as e:
            # LLM 실패 시 자동 매핑으로 폴백
            plan = self._create_mapping_auto(template, content)
            plan.warnings.append(f"LLM mapping failed, using auto-mapping: {str(e)}")
            return plan

    def _parse_llm_response(
        self,
        response: Dict[str, Any],
        template: ParsedTemplate,
        content: DocumentStructure,
    ) -> ContentMappingPlan:
        """LLM 응답 파싱"""
        if "error" in response:
            # 파싱 실패 - 자동 매핑으로 폴백
            plan = self._create_mapping_auto(template, content)
            plan.warnings.append(f"LLM response parsing failed: {response.get('error')}")
            return plan

        mappings = []
        for m in response.get("mappings", []):
            mappings.append(ContentMapping(
                placeholder_id=m.get("placeholder_id", ""),
                content_block_indices=m.get("content_block_indices", []),
                transformation=m.get("transformation", "none"),
            ))

        return ContentMappingPlan(
            mappings=mappings,
            unmapped_content=response.get("unmapped_content", []),
            warnings=response.get("warnings", []),
            confidence=response.get("confidence", 0.8),
        )

    def _create_mapping_auto(
        self,
        template: ParsedTemplate,
        content: DocumentStructure,
    ) -> ContentMappingPlan:
        """규칙 기반 자동 매핑"""
        mappings = []
        used_indices = set()
        warnings = []

        # 블록 인덱스 준비
        blocks = content.raw_blocks
        total_blocks = len(blocks)

        # 플레이스홀더별 매핑
        for placeholder in template.placeholders:
            mapping = self._auto_map_placeholder(
                placeholder, blocks, used_indices
            )
            if mapping:
                mappings.append(mapping)
                used_indices.update(mapping.content_block_indices)

        # 매핑되지 않은 콘텐츠 찾기
        unmapped = [i for i in range(total_blocks) if i not in used_indices]

        # BODY 플레이스홀더가 있으면 남은 콘텐츠 모두 추가
        body_mapping = next(
            (m for m in mappings if "BODY" in m.placeholder_id.upper()),
            None
        )
        if body_mapping and unmapped:
            body_mapping.content_block_indices.extend(unmapped)
            body_mapping.content_block_indices.sort()
            unmapped = []

        return ContentMappingPlan(
            mappings=mappings,
            unmapped_content=unmapped,
            warnings=warnings,
            confidence=0.7 if mappings else 0.0,
        )

    def _auto_map_placeholder(
        self,
        placeholder: Placeholder,
        blocks: List[ContentBlock],
        used_indices: set,
    ) -> Optional[ContentMapping]:
        """단일 플레이스홀더 자동 매핑"""
        ptype = placeholder.placeholder_type
        indices = []

        if ptype == PlaceholderType.TITLE:
            # 첫 번째 H1 헤딩 찾기
            for i, block in enumerate(blocks):
                if i in used_indices:
                    continue
                if block.block_type == "heading" and block.level == 1:
                    indices.append(i)
                    break

        elif ptype == PlaceholderType.SUBTITLE:
            # 첫 번째 H2 또는 타이틀 다음 문단 찾기
            for i, block in enumerate(blocks):
                if i in used_indices:
                    continue
                if block.block_type == "heading" and block.level == 2:
                    indices.append(i)
                    break
                # 첫 번째 일반 문단도 고려
                if block.block_type == "paragraph" and not indices:
                    indices.append(i)
                    break

        elif ptype == PlaceholderType.BODY:
            # 타이틀/서브타이틀 이후 모든 콘텐츠
            skip_first_heading = True
            for i, block in enumerate(blocks):
                if i in used_indices:
                    continue
                # 첫 H1, H2는 건너뛰기 (타이틀/서브타이틀용)
                if skip_first_heading and block.block_type == "heading" and block.level <= 2:
                    skip_first_heading = False
                    continue
                indices.append(i)

        elif ptype == PlaceholderType.SECTION:
            # 특정 섹션 번호에 해당하는 콘텐츠
            section_num = placeholder.section_number or 1
            current_section = 0

            in_section = False
            for i, block in enumerate(blocks):
                if i in used_indices:
                    continue

                # H2로 섹션 구분
                if block.block_type == "heading" and block.level == 2:
                    current_section += 1
                    in_section = (current_section == section_num)
                    if in_section:
                        indices.append(i)
                    elif current_section > section_num:
                        break
                elif in_section:
                    indices.append(i)

        elif ptype == PlaceholderType.TOC:
            # 목차는 자동 생성으로 처리 (콘텐츠 매핑 없음)
            return ContentMapping(
                placeholder_id=placeholder.id,
                content_block_indices=[],
                transformation="none",
            )

        if indices:
            return ContentMapping(
                placeholder_id=placeholder.id,
                content_block_indices=indices,
                transformation="none",
            )

        return None


# 동기 버전 (테스트/간단한 사용)
class ContentMapperSync:
    """동기 버전 콘텐츠 매퍼"""

    def create_mapping_plan(
        self,
        template: ParsedTemplate,
        content: DocumentStructure,
    ) -> ContentMappingPlan:
        """규칙 기반 매핑 생성 (LLM 없이)"""
        mapper = LLMContentMapper(use_llm=False)
        return mapper._create_mapping_auto(template, content)


if __name__ == "__main__":
    import asyncio
    from .template_parser import TemplateParser
    from .markdown_parser import MarkdownParser

    async def test():
        # 테스트용 마크다운
        test_md = """
# 프로젝트 제안서

## 개요

이 프로젝트는 문서 자동화 시스템입니다.

## 주요 기능

- 마크다운 파싱
- 템플릿 분석
- DOCX 생성

## 결론

문서 자동화로 업무 효율을 높입니다.
"""

        # 마크다운 파싱
        md_parser = MarkdownParser()
        content = md_parser.parse(test_md)

        print(f"콘텐츠 블록 수: {len(content.raw_blocks)}")
        for i, block in enumerate(content.raw_blocks):
            preview = block.content[:30] + "..." if len(block.content) > 30 else block.content
            print(f"  [{i}] {block.block_type} (level={block.level}): {preview}")

        # 더미 템플릿 생성 (실제로는 TemplateParser 사용)
        from .models import ParsedTemplate, Placeholder, PlaceholderType, SectionType

        template = ParsedTemplate(
            file_path="test.docx",
            placeholders=[
                Placeholder(
                    id="{{TITLE}}",
                    placeholder_type=PlaceholderType.TITLE,
                    section_type=SectionType.COVER,
                    paragraph_index=0,
                ),
                Placeholder(
                    id="{{BODY}}",
                    placeholder_type=PlaceholderType.BODY,
                    section_type=SectionType.BODY,
                    paragraph_index=5,
                ),
            ]
        )

        # 자동 매핑 테스트
        mapper = LLMContentMapper(use_llm=False)
        plan = mapper._create_mapping_auto(template, content)

        print(f"\n매핑 결과:")
        for m in plan.mappings:
            print(f"  {m.placeholder_id} -> 블록 {m.content_block_indices}")

        print(f"\n매핑되지 않은 블록: {plan.unmapped_content}")
        print(f"신뢰도: {plan.confidence}")

    asyncio.run(test())
