"""
vLLM API 클라이언트

OpenAI-compatible API를 통해 vLLM 서버와 통신
"""

import os
import json
import httpx
from typing import Optional, Dict, Any, List
from dataclasses import dataclass


@dataclass
class VLLMConfig:
    """vLLM 클라이언트 설정"""
    base_url: str = "http://localhost:8000/v1"
    model: str = "Qwen/Qwen2.5-7B-Instruct"
    api_key: str = "EMPTY"  # vLLM은 보통 API 키 불필요
    timeout: float = 120.0
    max_tokens: int = 4096
    temperature: float = 0.1  # 낮은 온도로 일관된 JSON 출력


class VLLMClient:
    """
    vLLM 서버 API 클라이언트

    OpenAI-compatible API 사용
    """

    def __init__(
        self,
        base_url: Optional[str] = None,
        model: Optional[str] = None,
        api_key: Optional[str] = None,
        timeout: float = 120.0,
    ):
        """
        Args:
            base_url: vLLM 서버 URL (기본: 환경변수 VLLM_BASE_URL 또는 localhost:8000)
            model: 모델 이름 (기본: 환경변수 VLLM_MODEL)
            api_key: API 키 (보통 불필요)
            timeout: 요청 타임아웃
        """
        self.config = VLLMConfig(
            base_url=base_url or os.getenv("VLLM_BASE_URL", "http://localhost:8000/v1"),
            model=model or os.getenv("VLLM_MODEL", "Qwen/Qwen2.5-7B-Instruct"),
            api_key=api_key or os.getenv("VLLM_API_KEY", "EMPTY"),
            timeout=timeout,
        )

        self._client: Optional[httpx.AsyncClient] = None

    @property
    def client(self) -> httpx.AsyncClient:
        """HTTP 클라이언트 (lazy initialization)"""
        if self._client is None:
            self._client = httpx.AsyncClient(
                base_url=self.config.base_url,
                headers={
                    "Authorization": f"Bearer {self.config.api_key}",
                    "Content-Type": "application/json",
                },
                timeout=httpx.Timeout(self.config.timeout),
            )
        return self._client

    async def close(self):
        """클라이언트 종료"""
        if self._client:
            await self._client.aclose()
            self._client = None

    async def __aenter__(self):
        return self

    async def __aexit__(self, exc_type, exc_val, exc_tb):
        await self.close()

    async def health_check(self) -> bool:
        """서버 상태 확인"""
        try:
            response = await self.client.get("/models")
            return response.status_code == 200
        except Exception:
            return False

    async def chat_completion(
        self,
        messages: List[Dict[str, str]],
        model: Optional[str] = None,
        max_tokens: Optional[int] = None,
        temperature: Optional[float] = None,
        response_format: Optional[Dict[str, str]] = None,
    ) -> Dict[str, Any]:
        """
        Chat completion API 호출

        Args:
            messages: 대화 메시지 리스트
            model: 모델 이름 (기본: 설정값)
            max_tokens: 최대 토큰 수
            temperature: 샘플링 온도
            response_format: 응답 형식 (예: {"type": "json_object"})

        Returns:
            API 응답 딕셔너리
        """
        payload = {
            "model": model or self.config.model,
            "messages": messages,
            "max_tokens": max_tokens or self.config.max_tokens,
            "temperature": temperature if temperature is not None else self.config.temperature,
        }

        if response_format:
            payload["response_format"] = response_format

        response = await self.client.post("/chat/completions", json=payload)
        response.raise_for_status()

        return response.json()

    async def generate_json(
        self,
        system_prompt: str,
        user_prompt: str,
        model: Optional[str] = None,
    ) -> Dict[str, Any]:
        """
        JSON 응답 생성

        Args:
            system_prompt: 시스템 프롬프트
            user_prompt: 사용자 프롬프트
            model: 모델 이름

        Returns:
            파싱된 JSON 딕셔너리
        """
        messages = [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt},
        ]

        # JSON 모드 시도 (지원하는 경우)
        try:
            response = await self.chat_completion(
                messages=messages,
                model=model,
                response_format={"type": "json_object"},
            )
        except Exception:
            # JSON 모드 미지원 시 일반 요청
            response = await self.chat_completion(messages=messages, model=model)

        content = response["choices"][0]["message"]["content"]

        # JSON 파싱
        return self._extract_json(content)

    def _extract_json(self, text: str) -> Dict[str, Any]:
        """텍스트에서 JSON 추출"""
        # 직접 파싱 시도
        try:
            return json.loads(text)
        except json.JSONDecodeError:
            pass

        # 코드 블록에서 추출
        import re
        json_match = re.search(r"```(?:json)?\s*([\s\S]*?)\s*```", text)
        if json_match:
            try:
                return json.loads(json_match.group(1))
            except json.JSONDecodeError:
                pass

        # { } 사이 추출
        brace_match = re.search(r"\{[\s\S]*\}", text)
        if brace_match:
            try:
                return json.loads(brace_match.group(0))
            except json.JSONDecodeError:
                pass

        # 실패 시 빈 딕셔너리
        return {"error": "Failed to parse JSON", "raw_response": text}

    async def get_content_mapping(
        self,
        system_prompt: str,
        user_prompt: str,
    ) -> Dict[str, Any]:
        """
        콘텐츠 매핑 생성

        Args:
            system_prompt: 시스템 프롬프트
            user_prompt: 플레이스홀더와 콘텐츠 정보를 담은 프롬프트

        Returns:
            매핑 계획 딕셔너리
        """
        return await self.generate_json(system_prompt, user_prompt)


# 동기 래퍼 (간단한 테스트용)
class VLLMClientSync:
    """동기 버전 vLLM 클라이언트"""

    def __init__(self, **kwargs):
        self.config = VLLMConfig(
            base_url=kwargs.get("base_url") or os.getenv("VLLM_BASE_URL", "http://localhost:8000/v1"),
            model=kwargs.get("model") or os.getenv("VLLM_MODEL", "Qwen/Qwen2.5-7B-Instruct"),
            api_key=kwargs.get("api_key") or os.getenv("VLLM_API_KEY", "EMPTY"),
            timeout=kwargs.get("timeout", 120.0),
        )

    def chat_completion(
        self,
        messages: List[Dict[str, str]],
        **kwargs
    ) -> Dict[str, Any]:
        """동기 Chat completion"""
        with httpx.Client(
            base_url=self.config.base_url,
            headers={
                "Authorization": f"Bearer {self.config.api_key}",
                "Content-Type": "application/json",
            },
            timeout=httpx.Timeout(self.config.timeout),
        ) as client:
            payload = {
                "model": kwargs.get("model", self.config.model),
                "messages": messages,
                "max_tokens": kwargs.get("max_tokens", self.config.max_tokens),
                "temperature": kwargs.get("temperature", self.config.temperature),
            }

            response = client.post("/chat/completions", json=payload)
            response.raise_for_status()
            return response.json()


if __name__ == "__main__":
    import asyncio

    async def test():
        async with VLLMClient() as client:
            # 서버 상태 확인
            healthy = await client.health_check()
            print(f"vLLM Server healthy: {healthy}")

            if healthy:
                # 간단한 테스트
                result = await client.generate_json(
                    system_prompt="You are a helpful assistant that responds in JSON format.",
                    user_prompt='Return a JSON object with a "greeting" key and a friendly message.'
                )
                print(f"Response: {json.dumps(result, indent=2)}")

    asyncio.run(test())
