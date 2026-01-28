"""
콘텐츠 매핑을 위한 LLM 프롬프트 템플릿
"""

import json
from typing import List, Dict, Any


# 시스템 프롬프트
CONTENT_MAPPING_SYSTEM_PROMPT = """당신은 문서 레이아웃 전문가입니다.
마크다운 콘텐츠를 DOCX 템플릿의 플레이스홀더에 매핑하는 작업을 수행합니다.

## 규칙:
1. 각 플레이스홀더에 가장 적합한 콘텐츠 블록을 매핑하세요
2. TITLE 플레이스홀더에는 제목 (heading level 1)을 매핑
3. SUBTITLE 플레이스홀더에는 부제목 또는 요약을 매핑
4. BODY 플레이스홀더에는 본문 콘텐츠를 매핑
5. SECTION_N 플레이스홀더에는 N번째 섹션의 콘텐츠를 매핑
6. 매핑되지 않은 콘텐츠는 unmapped_content에 인덱스를 기록

## 출력 형식:
반드시 아래 JSON 형식으로 응답하세요:

```json
{
    "mappings": [
        {
            "placeholder_id": "{{TITLE}}",
            "content_block_indices": [0],
            "transformation": "none"
        },
        {
            "placeholder_id": "{{BODY}}",
            "content_block_indices": [1, 2, 3, 4],
            "transformation": "none"
        }
    ],
    "unmapped_content": [],
    "warnings": [],
    "confidence": 0.95
}
```

transformation 옵션:
- "none": 그대로 사용
- "summarize": 요약하여 사용
- "extract_first": 첫 번째 요소만 사용
"""


# 사용자 프롬프트 템플릿
CONTENT_MAPPING_USER_PROMPT = """## 템플릿 플레이스홀더:
{placeholders_json}

## 마크다운 콘텐츠 블록:
{content_json}

위 플레이스홀더와 콘텐츠 블록을 분석하여 최적의 매핑을 생성하세요.
JSON 형식으로만 응답하세요.
"""


def build_mapping_prompt(
    placeholders: List[Dict[str, Any]],
    content_blocks: List[Dict[str, Any]],
) -> str:
    """
    콘텐츠 매핑 프롬프트 생성

    Args:
        placeholders: 플레이스홀더 정보 리스트
        content_blocks: 마크다운 콘텐츠 블록 리스트

    Returns:
        완성된 사용자 프롬프트
    """
    # 플레이스홀더 요약
    ph_summary = []
    for p in placeholders:
        ph_summary.append({
            "id": p.get("id"),
            "type": p.get("type"),
            "style": p.get("style"),
        })

    # 콘텐츠 블록 요약 (긴 텍스트는 잘라서)
    content_summary = []
    for idx, block in enumerate(content_blocks):
        content = block.get("content", "")
        if len(content) > 200:
            content = content[:200] + "..."

        content_summary.append({
            "index": idx,
            "type": block.get("block_type"),
            "level": block.get("level", 0),
            "content_preview": content,
        })

    return CONTENT_MAPPING_USER_PROMPT.format(
        placeholders_json=json.dumps(ph_summary, ensure_ascii=False, indent=2),
        content_json=json.dumps(content_summary, ensure_ascii=False, indent=2)
    )


def build_simple_mapping_prompt(
    placeholders: List[str],
    content_blocks: List[Dict[str, Any]],
) -> str:
    """
    간단한 매핑 프롬프트 (플레이스홀더 ID 리스트만)

    Args:
        placeholders: 플레이스홀더 ID 리스트 (예: ["{{TITLE}}", "{{BODY}}"])
        content_blocks: 마크다운 콘텐츠 블록 리스트

    Returns:
        완성된 사용자 프롬프트
    """
    # 콘텐츠 블록 요약
    content_summary = []
    for idx, block in enumerate(content_blocks):
        content = block.get("content", "")
        if len(content) > 150:
            content = content[:150] + "..."

        content_summary.append({
            "index": idx,
            "type": block.get("block_type"),
            "level": block.get("level", 0),
            "preview": content,
        })

    prompt = f"""## 플레이스홀더 목록:
{json.dumps(placeholders, ensure_ascii=False)}

## 콘텐츠 블록:
{json.dumps(content_summary, ensure_ascii=False, indent=2)}

각 플레이스홀더에 적합한 콘텐츠 블록 인덱스를 매핑하세요.
JSON 형식으로 응답하세요.
"""
    return prompt


# 자동 매핑용 규칙 기반 프롬프트 (LLM 없이 사용)
AUTO_MAPPING_RULES = {
    "{{TITLE}}": {
        "block_types": ["heading"],
        "levels": [1],
        "max_count": 1,
    },
    "{{SUBTITLE}}": {
        "block_types": ["heading", "paragraph"],
        "levels": [2],
        "max_count": 1,
    },
    "{{BODY}}": {
        "block_types": ["heading", "paragraph", "list", "list_item", "table", "blockquote", "code"],
        "levels": [2, 3, 4, 5, 6, 0],
        "max_count": None,  # 제한 없음
    },
    "{{DATE}}": {
        "block_types": ["paragraph"],
        "patterns": [r"\d{4}[-/]\d{1,2}[-/]\d{1,2}"],
        "max_count": 1,
    },
}


def get_auto_mapping_rule(placeholder_id: str) -> Dict[str, Any]:
    """자동 매핑 규칙 반환"""
    return AUTO_MAPPING_RULES.get(placeholder_id, AUTO_MAPPING_RULES["{{BODY}}"])
