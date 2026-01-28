"""
문서 자동화 시스템 - 데이터 모델

Pydantic 모델들:
- Placeholder: 템플릿의 플레이스홀더 정보
- ContentMapping: 콘텐츠-플레이스홀더 매핑
- ContentMappingPlan: LLM이 생성한 전체 매핑 계획
"""

from pydantic import BaseModel, Field
from typing import Optional, List, Dict, Any, Literal
from enum import Enum


class PlaceholderType(str, Enum):
    """플레이스홀더 유형"""
    TITLE = "title"           # {{TITLE}}
    SUBTITLE = "subtitle"     # {{SUBTITLE}}
    DATE = "date"             # {{DATE}}
    AUTHOR = "author"         # {{AUTHOR}}
    TOC = "toc"               # {{TOC}} - 목차
    BODY = "body"             # {{BODY}} - 본문 전체
    SECTION = "section"       # {{SECTION_N}} - N번째 섹션
    IMAGE = "image"           # {{IMAGE}}
    CUSTOM = "custom"         # 기타 사용자 정의


class SectionType(str, Enum):
    """템플릿 섹션 유형"""
    COVER = "cover"           # 표지
    TOC = "toc"               # 목차
    SECTION_BREAK = "section" # 섹션 구분 페이지
    BODY = "body"             # 본문


class Placeholder(BaseModel):
    """
    템플릿에서 발견된 플레이스홀더

    예: {{TITLE}}, {{BODY}}, {{SECTION_1}}
    """
    id: str                             # "{{TITLE}}", "{{BODY}}"
    placeholder_type: PlaceholderType   # title, body, section 등
    section_type: SectionType = SectionType.BODY  # cover, toc, body
    paragraph_index: int                # 원본 문단 위치
    run_index: int = 0                  # 문단 내 run 위치
    style_id: Optional[str] = None      # 적용된 스타일 ID
    style_name: Optional[str] = None    # 스타일 이름 (예: "Heading 1")

    # 섹션 번호 ({{SECTION_N}}인 경우)
    section_number: Optional[int] = None

    # 원본 텍스트 (플레이스홀더 전후 텍스트 포함)
    original_text: str = ""

    class Config:
        use_enum_values = True


class ParsedTemplate(BaseModel):
    """
    파싱된 템플릿 정보

    플레이스홀더 목록과 섹션 구조를 포함
    """
    file_path: str
    placeholders: List[Placeholder] = Field(default_factory=list)

    # 섹션별 플레이스홀더 그룹
    cover_placeholders: List[Placeholder] = Field(default_factory=list)
    toc_placeholders: List[Placeholder] = Field(default_factory=list)
    body_placeholders: List[Placeholder] = Field(default_factory=list)

    # 총 문단 수
    total_paragraphs: int = 0

    def get_placeholder_by_id(self, placeholder_id: str) -> Optional[Placeholder]:
        """ID로 플레이스홀더 찾기"""
        for p in self.placeholders:
            if p.id == placeholder_id:
                return p
        return None

    def get_placeholders_by_type(self, ptype: PlaceholderType) -> List[Placeholder]:
        """타입으로 플레이스홀더 필터링"""
        return [p for p in self.placeholders if p.placeholder_type == ptype]


class ContentMapping(BaseModel):
    """
    단일 콘텐츠-플레이스홀더 매핑

    마크다운의 어떤 블록들이 어떤 플레이스홀더에 들어가는지 지정
    """
    placeholder_id: str                      # "{{TITLE}}", "{{BODY}}"
    content_block_indices: List[int]         # 마크다운 블록 인덱스들

    # 변환 옵션
    transformation: Optional[Literal["none", "summarize", "extract_first"]] = "none"

    # 스타일 오버라이드 (None이면 플레이스홀더 스타일 사용)
    style_override: Optional[str] = None


class ContentMappingPlan(BaseModel):
    """
    LLM이 생성한 전체 매핑 계획

    모든 플레이스홀더에 대한 콘텐츠 매핑 정보
    """
    mappings: List[ContentMapping] = Field(default_factory=list)

    # 매핑되지 않은 콘텐츠 블록 인덱스
    unmapped_content: List[int] = Field(default_factory=list)

    # LLM의 추가 메시지/경고
    warnings: List[str] = Field(default_factory=list)

    # 매핑 신뢰도 (0.0 ~ 1.0)
    confidence: float = 0.0

    def get_mapping_for_placeholder(self, placeholder_id: str) -> Optional[ContentMapping]:
        """플레이스홀더 ID로 매핑 찾기"""
        for m in self.mappings:
            if m.placeholder_id == placeholder_id:
                return m
        return None

    def get_content_indices_for_placeholder(self, placeholder_id: str) -> List[int]:
        """플레이스홀더에 매핑된 콘텐츠 인덱스 반환"""
        mapping = self.get_mapping_for_placeholder(placeholder_id)
        return mapping.content_block_indices if mapping else []


class LLMRequest(BaseModel):
    """LLM 요청 데이터"""
    template_placeholders: List[Dict[str, Any]]  # 플레이스홀더 정보
    content_blocks: List[Dict[str, Any]]         # 마크다운 블록 정보
    system_prompt: str = ""
    user_prompt: str = ""


class LLMResponse(BaseModel):
    """LLM 응답 데이터"""
    mapping_plan: ContentMappingPlan
    raw_response: str = ""
    model_name: str = ""
    tokens_used: int = 0


# 플레이스홀더 패턴 상수
PLACEHOLDER_PATTERNS = {
    "default": r"\{\{([A-Z_]+(?:_\d+)?)\}\}",     # {{TITLE}}, {{SECTION_1}}
    "bracket": r"\[\[([A-Z_]+(?:_\d+)?)\]\]",     # [[TITLE]]
    "angle": r"<<([A-Z_]+(?:_\d+)?)>>",           # <<TITLE>>
    "underscore": r"___([A-Z_]+(?:_\d+)?)___",    # ___TITLE___
}

# 플레이스홀더 ID -> 타입 매핑
PLACEHOLDER_TYPE_MAP = {
    "TITLE": PlaceholderType.TITLE,
    "SUBTITLE": PlaceholderType.SUBTITLE,
    "DATE": PlaceholderType.DATE,
    "AUTHOR": PlaceholderType.AUTHOR,
    "TOC": PlaceholderType.TOC,
    "BODY": PlaceholderType.BODY,
    "IMAGE": PlaceholderType.IMAGE,
    "COVER_IMAGE": PlaceholderType.IMAGE,
}


def parse_placeholder_id(placeholder_id: str) -> tuple[PlaceholderType, Optional[int]]:
    """
    플레이스홀더 ID를 파싱하여 타입과 섹션 번호 반환

    예:
        "{{TITLE}}" -> (PlaceholderType.TITLE, None)
        "{{SECTION_1}}" -> (PlaceholderType.SECTION, 1)
    """
    import re

    # {{ }} 제거
    clean_id = placeholder_id.strip("{}")

    # SECTION_N 패턴 체크
    section_match = re.match(r"SECTION_(\d+)", clean_id)
    if section_match:
        return PlaceholderType.SECTION, int(section_match.group(1))

    # 일반 타입 매핑
    ptype = PLACEHOLDER_TYPE_MAP.get(clean_id, PlaceholderType.CUSTOM)
    return ptype, None
