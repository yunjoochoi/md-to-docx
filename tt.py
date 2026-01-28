import zipfile
import xml.etree.ElementTree as ET
import hashlib

NS = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}

def get_section_fingerprint(sectPr_element):
    """
    섹션 설정(sectPr)에서 시각적으로 중요한 요소만 뽑아서
    고유한 해시값(지문)을 만듭니다.
    """
    if sectPr_element is None:
        return "default_section"
    
    # 비교할 핵심 속성들 (이것들이 같으면 같은 모양의 섹션임)
    fingerprint_keys = []
    
    # 1. 페이지 크기 (가로/세로 포함)
    pgSz = sectPr_element.find('w:pgSz', NS)
    if pgSz is not None:
        fingerprint_keys.append(f"Sz:{pgSz.get(f'{{{NS['w']}}}w')}-{pgSz.get(f'{{{NS['w']}}}h')}-{pgSz.get(f'{{{NS['w']}}}orient')}")
        
    # 2. 여백
    pgMar = sectPr_element.find('w:pgMar', NS)
    if pgMar is not None:
        # 상하좌우 여백 수치를 문자열로 연결
        attrs = ['top', 'bottom', 'left', 'right', 'header', 'footer']
        mar_str = "-".join([str(pgMar.get(f'{{{NS['w']}}}{a}')) for a in attrs])
        fingerprint_keys.append(f"Mar:{mar_str}")
        
    # 3. 헤더/푸터 참조 (매우 중요: 배경 이미지가 여기서 결정됨)
    # r:id는 파일 내부 ID이므로, 실제로는 연결된 header.xml 파일의 내용을 비교해야 완벽하지만,
    # 보통 같은 템플릿 내에서는 r:id 참조 구성이 같으면 같은 스타일로 봐도 무방함.
    headers = sectPr_element.findall('w:headerReference', NS)
    for h in headers:
        fingerprint_keys.append(f"H:{h.get(f'{{{NS['w']}}}type')}") # type=first/default 등

    # 4. 해시 생성
    fingerprint_string = "|".join(fingerprint_keys)
    return hashlib.md5(fingerprint_string.encode()).hexdigest()

def extract_unique_layouts(docx_path):
    with zipfile.ZipFile(docx_path, 'r') as zf:
        xml_content = zf.read('word/document.xml')
        tree = ET.fromstring(xml_content)
        body = tree.find('w:body', NS)
        
        unique_layouts = []
        seen_hashes = set()
        
        # 문서 중간중간의 섹션 브레이크 (<w:p> 내부의 <w:sectPr>)
        for p in body.findall('w:p', NS):
            pPr = p.find('w:pPr', NS)
            if pPr is not None:
                sectPr = pPr.find('w:sectPr', NS)
                if sectPr is not None:
                    fp = get_section_fingerprint(sectPr)
                    if fp not in seen_hashes:
                        seen_hashes.add(fp)
                        unique_layouts.append({
                            "type": "intermediate",
                            "hash": fp,
                            "xml": ET.tostring(sectPr, encoding='unicode')
                        })

        # 문서 맨 마지막 섹션 설정 (body 직계 자식 <w:sectPr>) -> 보통 이게 '본문' 스타일
        last_sectPr = body.find('w:sectPr', NS)
        if last_sectPr is not None:
            fp = get_section_fingerprint(last_sectPr)
            if fp not in seen_hashes:
                seen_hashes.add(fp)
                unique_layouts.append({
                    "type": "final_body", # 가장 중요한 본문 레이아웃
                    "hash": fp,
                    "xml": ET.tostring(last_sectPr, encoding='unicode')
                })
                
        return unique_layouts

# 실행 예시
layouts = extract_unique_layouts("/home/shaush/md-to-docx/[Word템플릿]A4.docx")
print(f"추출된 유니크 레이아웃 개수: {len(layouts)}")
# 결과가 2개라면 -> [0]: 표지용, [1]: 본문용 으로 매핑하면 됨.
