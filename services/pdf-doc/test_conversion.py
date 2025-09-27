import requests
import os
import time

def test_pdf_conversion():
    """PDF 변환 테스트 함수"""
    base_url = "http://127.0.0.1:5000"
    test_file = "test_korean_document.pdf"
    
    if not os.path.exists(test_file):
        print(f"테스트 파일 {test_file}이 존재하지 않습니다.")
        return
    
    print("=== PDF 변환 테스트 시작 ===")
    print(f"테스트 파일: {test_file}")
    
    # 1. 기본 변환 모드 테스트
    print("\n1. 기본 변환 모드 테스트")
    try:
        with open(test_file, 'rb') as f:
            files = {'file': f}
            data = {
                'conversion_mode': 'basic',
                'preserve_images': 'true',
                'ocr_enabled': 'true'
            }
            
            response = requests.post(f"{base_url}/convert", files=files, data=data)
            
            if response.status_code == 200:
                print("✓ 기본 변환 성공")
                # 응답 헤더에서 파일명 확인
                content_disposition = response.headers.get('content-disposition', '')
                print(f"응답 헤더: {content_disposition}")
                print(f"응답 크기: {len(response.content)} bytes")
            else:
                print(f"✗ 기본 변환 실패: {response.status_code}")
                print(f"오류 메시지: {response.text}")
                
    except Exception as e:
        print(f"✗ 기본 변환 중 오류: {e}")
    
    time.sleep(2)
    
    # 2. 이미지 변환 모드 테스트
    print("\n2. 이미지 변환 모드 테스트")
    try:
        with open(test_file, 'rb') as f:
            files = {'file': f}
            data = {
                'conversion_mode': 'image_mode',
                'preserve_images': 'true',
                'ocr_enabled': 'true',
                'image_quality': 'high'
            }
            
            response = requests.post(f"{base_url}/convert", files=files, data=data)
            
            if response.status_code == 200:
                print("✓ 이미지 변환 모드 성공")
                content_disposition = response.headers.get('content-disposition', '')
                print(f"응답 헤더: {content_disposition}")
                print(f"응답 크기: {len(response.content)} bytes")
            else:
                print(f"✗ 이미지 변환 모드 실패: {response.status_code}")
                print(f"오류 메시지: {response.text}")
                
    except Exception as e:
        print(f"✗ 이미지 변환 모드 중 오류: {e}")
    
    time.sleep(2)
    
    # 3. 고급 변환 모드 테스트
    print("\n3. 고급 변환 모드 테스트")
    try:
        with open(test_file, 'rb') as f:
            files = {'file': f}
            data = {
                'conversion_mode': 'advanced',
                'preserve_images': 'true',
                'ocr_enabled': 'true',
                'image_quality': 'high',
                'layout_preservation': 'true'
            }
            
            response = requests.post(f"{base_url}/convert", files=files, data=data)
            
            if response.status_code == 200:
                print("✓ 고급 변환 모드 성공")
                content_disposition = response.headers.get('content-disposition', '')
                print(f"응답 헤더: {content_disposition}")
                print(f"응답 크기: {len(response.content)} bytes")
                
                # 결과 파일 저장
                output_filename = "test_advanced_conversion.docx"
                with open(output_filename, 'wb') as output_file:
                    output_file.write(response.content)
                print(f"✓ 결과 파일 저장: {output_filename}")
                
            else:
                print(f"✗ 고급 변환 모드 실패: {response.status_code}")
                print(f"오류 메시지: {response.text}")
                
    except Exception as e:
        print(f"✗ 고급 변환 모드 중 오류: {e}")
    
    print("\n=== 테스트 완료 ===")

def check_server_status():
    """서버 상태 확인"""
    try:
        response = requests.get("http://127.0.0.1:5000")
        if response.status_code == 200:
            print("✓ 서버가 정상적으로 실행 중입니다.")
            return True
        else:
            print(f"✗ 서버 응답 오류: {response.status_code}")
            return False
    except Exception as e:
        print(f"✗ 서버 연결 실패: {e}")
        return False

if __name__ == "__main__":
    print("=== 개선된 이미지 처리 시스템 테스트 ===")
    
    # 서버 상태 확인
    if check_server_status():
        # 변환 테스트 실행
        test_pdf_conversion()
    else:
        print("서버가 실행되지 않았습니다. 먼저 서버를 시작해주세요.")