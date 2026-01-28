

import pypandoc
import os

def run_pypandoc_conversion(md_path, template_path, output_path):
    """
    pypandoc을 사용하여 마크다운을 템플릿 스타일이 적용된 docx로 변환합니다.
    """
    
    # 1. Pandoc 바이너리 존재 확인 및 다운로드 (필요시)
    try:
        # Pandoc이 설치되어 있는지 확인
        version = pypandoc.get_pandoc_version()
        print(f"Pandoc version found: {version}")
    except OSError:
        print("Pandoc 바이너리가 시스템에 없습니다. 다운로드를 시도합니다...")
        try:
            # pypandoc이 제공하는 유틸리티로 pandoc 바이너리 다운로드
            pypandoc.download_pandoc()
            print("Pandoc 다운로드 완료.")
        except Exception as e:
            print(f"Pandoc 다운로드 실패. 직접 설치해주세요: {e}")
            return

    # 2. 입력 파일 존재 확인
    if not os.path.exists(md_path):
        print(f"오류: 마크다운 파일({md_path})을 찾을 수 없습니다.")
        return
    if not os.path.exists(template_path):
        print(f"오류: 템플릿 파일({template_path})을 찾을 수 없습니다.")
        return

    # 3. 변환 실행 (핵심 로직)
    try:
        print(f"변환 시작: {md_path} -> {output_path} (Template: {template_path})")
        
        # convert_file 함수 사용
        # extra_args=['--reference-doc=...'] 부분이 핵심입니다.
        pypandoc.convert_file(
            md_path,
            to='docx',
            outputfile=output_path,
            extra_args=[f'--reference-doc={template_path}'] # 템플릿 지정 옵션
        )
        
        print(f"✅ 변환 성공! 결과 파일: {output_path}")
        
    except RuntimeError as e:
        print(f"❌ 변환 중 오류 발생 (Pandoc 에러): {e}")
    except Exception as e:
        print(f"❌ 알 수 없는 오류 발생: {e}")

# --- 실행 예시 ---
if __name__ == "__main__":
    # 테스트할 파일 경로를 지정하세요.
    # 1. 변환할 마크다운 파일
    my_markdown = "/home/shaush/work/output_md/20251128_company_554088000.md"  
    
    # 2. 스타일이 정의된 템플릿 워드 파일
    my_template = "/home/shaush/md-to-docx/[Word템플릿]A4.docx" 
    
    # 3. 저장될 결과 파일 이름
    my_output = "final_result_pandoc.docx"

    # (테스트용 더미 파일 생성 - 실제 파일이 있다면 이 부분은 주석 처리하세요)
    if not os.path.exists(my_markdown):
        with open(my_markdown, "w", encoding="utf-8") as f:
            f.write("# 테스트 제목\n\n이것은 본문입니다.\n\n## 소제목\n- 리스트 1\n- 리스트 2")
            
    if not os.path.exists(my_template):
        print("⚠️ 주의: 템플릿 파일이 없습니다. 코드가 템플릿 없이 기본 변환을 수행할 수 있지만, 스타일 적용을 테스트하려면 실제 docx 파일 경로를 넣어주세요.")

    # 함수 실행
    run_pypandoc_conversion(my_markdown, my_template, my_output)