import pytesseract
from PIL import Image
import os

print("=== Tesseract OCR 설치 확인 테스트 ===")

try:
    # 1. 기본 설정 확인
    print("\n1. Tesseract 버전 확인:")
    version = pytesseract.get_tesseract_version()
    print(f"   버전: {version}")
    
    # 2. 사용 가능한 언어 확인
    print("\n2. 설치된 언어팩 확인:")
    languages = pytesseract.get_languages()
    print(f"   언어 목록: {languages}")
    
    # 3. 한국어 지원 확인
    if 'kor' in languages:
        print("   ✅ 한국어 언어팩 설치 완료!")
    else:
        print("   ❌ 한국어 언어팩이 설치되지 않았습니다.")
    
    # 4. 영어 지원 확인
    if 'eng' in languages:
        print("   ✅ 영어 언어팩 설치 완료!")
    else:
        print("   ❌ 영어 언어팩이 설치되지 않았습니다.")
    
    print("\n🎉 Tesseract OCR 설치 및 설정이 완료되었습니다!")
    print("🚀 이제 PDF ↔ PPTX 변환기에서 OCR 기능을 사용할 수 있습니다!")
    
except Exception as e:
    print(f"\n❌ 오류 발생: {e}")
    print("\n🔧 해결 방법:")
    print("1. Tesseract가 올바르게 설치되었는지 확인")
    print("2. PATH 환경변수에 Tesseract가 추가되었는지 확인")
    print("3. 새 터미널에서 다시 시도")
    
    # 수동 경로 설정 시도
    print("\n🔄 수동 경로 설정 시도...")
    pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
    try:
        languages = pytesseract.get_languages()
        print(f"   수동 설정 성공! 언어 목록: {languages}")
    except Exception as e2:
        print(f"   수동 설정도 실패: {e2}")