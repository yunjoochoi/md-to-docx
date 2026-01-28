생성된 파일들
파일	설명
src/models.py	Pydantic 데이터 모델 (Placeholder, ContentMapping 등)
src/template_parser.py	플레이스홀더 추출 ({{TITLE}}, {{BODY}} 등)
llm/vllm_client.py	vLLM API 클라이언트 (Qwen 모델 연동)
llm/prompts.py	LLM 프롬프트 템플릿
src/llm_content_mapper.py	콘텐츠 매핑 로직 (LLM/규칙 기반)
src/docx_composer.py	DOCX 조립기 (플레이스홀더 교체)
pipeline.py	전체 파이프라인 통합
사용법

# 1. 템플릿 분석 (플레이스홀더 확인)
uv run python main.py --analyze template.docx

# 2. 플레이스홀더 기반 변환 (규칙 기반 매핑)
uv run python main.py --pipeline content.md -t template.docx -o output.docx

# 3. LLM 기반 매핑 사용 (vLLM 서버 필요)
uv run python main.py --pipeline content.md -t template.docx -o output.docx --llm

# 4. vLLM 서버 URL 지정
uv run python main.py --pipeline content.md -t template.docx --llm --vllm-url http://your-server:8000/v1
플레이스홀더 규칙
템플릿에 다음 플레이스홀더를 추가하세요:

플레이스홀더	용도
{{TITLE}}	문서 제목 (H1 헤딩)
{{SUBTITLE}}	부제목
{{BODY}}	본문 전체
{{SECTION_N}}	N번째 섹션
{{DATE}}	날짜
{{TOC}}	목차
vLLM 서버 설정 (LLM 모드 사용 시)

# Qwen 모델로 vLLM 서버 실행
python -m vllm.entrypoints.openai.api_server \
    --model Qwen/Qwen2.5-7B-Instruct \
    --port 8000
