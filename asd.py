import zipfile
import xml.etree.ElementTree as ET
import re

# 1. 네임스페이스 정의 (필수)
NS = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
}

class DocxSegmenter:
    def __init__(self, docx_path):
        self.docx_path = docx_path
        self.xml_content = self._extract_document_xml()
        self.tree = ET.fromstring(self.xml_content)
        self.body = self.tree.find('w:body', NS)
        
        # 결과 저장소
        self.segments = {
            "cover": [],
            "toc": [],
            "body": []
        }

    def _extract_document_xml(self):
        """docx 압축을 풀고 document.xml을 읽어옵니다."""
        with zipfile.ZipFile(self.docx_path, 'r') as zf:
            return zf.read('word/document.xml').decode('utf-8')

    def _is_toc(self, element):
        """
        [단서 1] TOC(목차)인지 판별하는 결정적 로직
        1. <w:sdt> 태그 안에 'Table of Contents' 갤러리가 있는가?
        2. <w:instrText> 태그 안에 'TOC' 명령어가 있는가?
        """
        # Case A: SDT (Structured Document Tag) 확인
        if element.tag.endswith('sdt'):
            gallery = element.find('.//w:docPartGallery', NS)
            if gallery is not None and 'Table of Contents' in gallery.get(f"{{{NS['w']}}}val", ''):
                return True
        
        # Case B: Field Code 확인 (구버전 호환)
        for instr in element.findall('.//w:instrText', NS):
            if instr.text and 'TOC' in instr.text.strip():
                return True
        return False

    def _is_heading_1(self, element):
        """
        [단서 2] 본문의 시작점인 Heading 1인지 판별
        (styles.xml을 참조하는 게 정석이지만, 여기서는 스타일 ID로 약식 추론)
        """
        pPr = element.find('w:pPr', NS)
        if pPr is not None:
            pStyle = pPr.find('w:pStyle', NS)
            if pStyle is not None:
                val = pStyle.get(f"{{{NS['w']}}}val", "")
                # 스타일 ID가 '1', 'Heading1', '제목1' 인 경우 (템플릿마다 다를 수 있음)
                if val in ['1', 'Heading1', '제목1']: 
                    return True
                
                # 더 정확하게는 여기서 outlineLvl을 체크해야 함 (extracted_xml.txt 참조)
                # 하지만 document.xml에는 outlineLvl이 없고 styles.xml에만 있는 경우가 많음.
                # 따라서 여기서는 Heuristic하게 스타일 ID로 판단.
        return False

    def _has_section_break(self, element):
        """[단서 3] 섹션 나누기(sectPr)가 포함되어 있는가?"""
        if element.find('.//w:sectPr', NS) is not None:
            return True
        # 문단 자체가 sectPr인 경우 (마지막 섹션)
        if element.tag.endswith('sectPr'):
            return True
        return False

    def segment_document(self):
        """
        ★ 핵심 로직: 상태 머신 (State Machine)
        문서를 위에서 아래로 훑으며 상태를 변경함: [Unknown/Cover] -> [TOC] -> [Body]
        """
        current_state = "cover"  # 기본값: 문서는 표지로 시작한다고 가정
        
        for element in self.body:
            # 태그 이름 정리 (네임스페이스 제거하고 보기 편하게)
            tag_name = element.tag.split('}')[-1]

            # --- [상태 전환 로직] ---

            # 1. 목차(TOC) 감지 -> 상태를 무조건 TOC로 변경
            if self._is_toc(element):
                current_state = "toc"
            
            # 2. 본문(Heading 1) 감지 -> TOC나 Cover 상태였다면 Body로 확정
            elif current_state in ["cover", "toc"] and self._is_heading_1(element):
                current_state = "body"

            # 3. 섹션 브레이크 감지 -> 표지 섹션이 끝났다면 다음은 본문일 확률 높음
            elif current_state == "cover" and self._has_section_break(element):
                # 현재 요소까지는 커버로 넣고, 다음부터 본문으로 전환하기 위해 플래그 설정
                self.segments[current_state].append(ET.tostring(element, encoding='unicode'))
                current_state = "body" 
                continue 

            # 4. 목차가 끝났는지 감지 (TOC 상태인데 일반 문단이 나오기 시작하면 본문으로)
            # (이 로직은 TOC가 <sdt>로 감싸져 있지 않은 경우를 대비함)
            elif current_state == "toc" and not self._is_toc(element) and tag_name == 'p':
                 # TOC 내의 문단이 아니라면 본문으로 전환 (이 부분은 정밀 튜닝 필요)
                 # <sdt> 태그를 쓴다면 이 로직 없이 <sdt> 닫히면 바로 본문임.
                 pass

            # --- [데이터 분류 저장] ---
            self.segments[current_state].append(ET.tostring(element, encoding='unicode'))

        return self.segments

# --- 실행부 ---
if __name__ == "__main__":
    # 테스트할 파일 경로
    input_file = "/home/shaush/md-to-docx/docx_only/[Word 템플릿] A4.docx" 
    
    try:
        segmenter = DocxSegmenter(input_file)
        results = segmenter.segment_document()

        print(f"분석 완료")
        print(f"표지(Cover) 요소 개수: {len(results['cover'])}")
        print(f"목차(TOC) 요소 개수:   {len(results['toc'])}")
        print(f"본문(Body) 요소 개수:   {len(results['body'])}")
        print(results['cover'])
        if results['body']:
            print("\n[본문 첫 요소 미리보기]:")
            print(results['body'][0][:500])

    except Exception as e:
        print(f"오류 발생 (파일이 없거나 포맷이 다름): {e}")
