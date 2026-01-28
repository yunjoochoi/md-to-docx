import zipfile
import xml.dom.minidom
import xml.etree.ElementTree as ET
import os

def get_real_xml_paths(zf):
    """
    _rels/.rels와 document.xml.rels를 추적하여
    실제 document와 styles 파일의 경로를 동적으로 찾아냅니다.
    """
    # OPC 네임스페이스 표준 상수
    # (XML 파싱 시 네임스페이스 처리를 위해 필요하지만, 여기서는 Type 속성 값 비교에 사용)
    TYPE_OFFICE_DOCUMENT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
    TYPE_STYLES = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"

    real_paths = {
        'document': None,
        'styles': None
    }

    # 1단계: 루트 관계 파일(_rels/.rels)에서 'document' 위치 찾기
    if '_rels/.rels' in zf.namelist():
        xml_data = zf.read('_rels/.rels')
        tree = ET.fromstring(xml_data)
        
        # 네임스페이스 무시하고 모든 Relationship 태그 검색
        # (lxml 대신 기본 xml 라이브러리 사용 시 네임스페이스 처리가 번거로울 수 있어 태그명만으로 검색)
        for rel in tree.findall(f'.//{{http://schemas.openxmlformats.org/package/2006/relationships}}Relationship'):
            if rel.get('Type') == TYPE_OFFICE_DOCUMENT:
                target = rel.get('Target')
                # Target이 "/word/document.xml" 처럼 절대경로일 수도, "word/document.xml" 상대경로일 수도 있음
                real_paths['document'] = target.lstrip('/')
                break
    
    # 만약 루트 관계 파일에서 문서를 못 찾았다면 기본값(폴백) 설정
    if not real_paths['document']:
        real_paths['document'] = 'word/document.xml'

    # 2단계: document 관계 파일(word/_rels/document.xml.rels)에서 'styles' 위치 찾기
    doc_path = real_paths['document']
    doc_dir = os.path.dirname(doc_path)
    doc_name = os.path.basename(doc_path)
    
    # document.xml -> _rels/document.xml.rels 경로 계산
    # 예: word/document.xml -> word/_rels/document.xml.rels
    rels_path = f"{doc_dir}/_rels/{doc_name}.rels"
    
    # 일부 문서는 _rels 폴더가 없을 수도 있으므로 체크
    if rels_path in zf.namelist():
        xml_data = zf.read(rels_path)
        tree = ET.fromstring(xml_data)
        
        for rel in tree.findall(f'.//{{http://schemas.openxmlformats.org/package/2006/relationships}}Relationship'):
            if rel.get('Type') == TYPE_STYLES:
                target = rel.get('Target')
                # Target이 "styles.xml" 처럼 파일명만 있으면 document가 있는 폴더와 합침
                if not target.startswith('/'):
                    # 예: word/ + styles.xml
                    real_paths['styles'] = f"{doc_dir}/{target}"
                    # 경로 구분자 정리 (혹시 모를 // 방지)
                    real_paths['styles'] = real_paths['styles'].replace('//', '/').lstrip('/')
                else:
                    real_paths['styles'] = target.lstrip('/')
                break
    
    return real_paths

def extract_docx_xml_to_text_opc(docx_path, output_txt_path=None):
    """
    OPC 표준에 따라 경로를 동적으로 찾아서 XML을 추출합니다.
    """
    extracted_text = ""
    
    try:
        if not os.path.exists(docx_path):
             return f"❌ Error: File not found - {docx_path}"

        with zipfile.ZipFile(docx_path, 'r') as zf:
            # ★ 핵심: 하드코딩 대신 진짜 경로 찾아오기
            paths = get_real_xml_paths(zf)
            
            print(f"🔎 [Path Discovery] Document Path: {paths['document']}")
            print(f"🔎 [Path Discovery] Styles Path:   {paths['styles']}")

            # 추출할 타겟 설정 (찾아낸 경로 사용)
            targets = []
            if paths['document']: 
                targets.append((paths['document'], 'Document XML (Main Content)'))
            
            if paths['styles']:   
                targets.append((paths['styles'], 'Styles XML (Formatting)'))
            else:
                extracted_text += "[WARNING] Styles file path could not be determined via relationships.\n\n"

            for xml_filename, desc in targets:
                if xml_filename in zf.namelist():
                    xml_bytes = zf.read(xml_filename)
                    # 보기 좋게 포맷팅 (Pretty Print)
                    try:
                        parsed_xml = xml.dom.minidom.parseString(xml_bytes.decode('utf-8'))
                        pretty_xml = parsed_xml.toprettyxml(indent="  ")
                    except Exception as parse_err:
                        pretty_xml = f"(XML Parsing Failed: {str(parse_err)})\n" + xml_bytes.decode('utf-8')
                    
                    extracted_text += f"{'='*30}\nFILE: {xml_filename} ({desc})\n{'='*30}\n"
                    extracted_text += pretty_xml + "\n\n"
                else:
                    extracted_text += f"[WARNING] Path found ({xml_filename}) via relationships, but file is missing in zip.\n\n"
        
        if output_txt_path:
            with open(output_txt_path, 'w', encoding='utf-8') as f:
                f.write(extracted_text)
            print(f"✅ 추출 완료 및 저장됨: {output_txt_path}")
            
        return extracted_text

    except Exception as e:
        import traceback
        return f"❌ Error processing DOCX file: {str(e)}\n{traceback.format_exc()}"

# --- 실행 예시 ---
if __name__ == "__main__":
    # 1. 테스트할 docx 파일 경로 지정
    input_docx = "/home/shaush/md-to-docx/[Word템플릿]A4.docx"
    output_txt = "extracted_xml.txt"

    # 실행
    result = extract_docx_xml_to_text_opc(input_docx, output_txt)
    
    # 에러 발생 시 출력
    if result.startswith("❌"):
        print(result)