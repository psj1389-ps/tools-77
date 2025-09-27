import pytesseract
from PIL import Image
from pdf2image import convert_from_path
import os
import logging

# OCR 설정 및 오류 처리
try:
    # Tesseract 경로 설정 (Render 환경에서 자동 감지)
    if os.path.exists('/usr/bin/tesseract'):
        pytesseract.pytesseract.tesseract_cmd = '/usr/bin/tesseract'
    elif os.path.exists('/opt/render/project/src/.apt/usr/bin/tesseract'):
        pytesseract.pytesseract.tesseract_cmd = '/opt/render/project/src/.apt/usr/bin/tesseract'
    
    # OCR 사용 가능 여부 확인
    OCR_AVAILABLE = True
    pytesseract.get_tesseract_version()
except Exception as e:
    OCR_AVAILABLE = False
    logging.warning(f"OCR 엔진을 사용할 수 없습니다: {e}")

def extract_text_with_ocr(pdf_path, lang='kor+eng'):
    """
    PDF에서 OCR을 사용하여 텍스트 추출
    
    Args:
        pdf_path (str): PDF 파일 경로
        lang (str): OCR 언어 설정 (기본값: 한국어+영어)
    
    Returns:
        list: 각 페이지별 추출된 텍스트 리스트
    """
    if not OCR_AVAILABLE:
        print("OCR 엔진을 사용할 수 없습니다. 빈 결과를 반환합니다.")
        return []
    
    try:
        # PDF를 이미지로 변환
        images = convert_from_path(pdf_path, dpi=300)
        
        extracted_texts = []
        for i, image in enumerate(images):
            print(f"페이지 {i+1} OCR 처리 중...")
            
            try:
                # OCR로 텍스트 추출
                text = pytesseract.image_to_string(image, lang=lang)
                extracted_texts.append(text)
                print(f"페이지 {i+1} 완료 - {len(text)} 글자 추출")
            except Exception as ocr_error:
                print(f"페이지 {i+1} OCR 실패: {ocr_error}")
                # Fallback: 영어만으로 재시도
                try:
                    text = pytesseract.image_to_string(image, lang='eng')
                    extracted_texts.append(text)
                    print(f"페이지 {i+1} 영어 OCR로 완료 - {len(text)} 글자 추출")
                except Exception:
                    extracted_texts.append("")
                    print(f"페이지 {i+1} OCR 완전 실패")
        
        return extracted_texts
        
    except Exception as e:
        print(f"OCR 처리 오류: {e}")
        return []

def test_ocr_with_sample():
    """
    샘플 이미지로 OCR 테스트
    """
    if not OCR_AVAILABLE:
        print("OCR 엔진을 사용할 수 없어 테스트를 건너뜁니다.")
        return None
    
    try:
        # 간단한 테스트 이미지 생성
        from PIL import Image, ImageDraw, ImageFont
        
        # 테스트 이미지 생성
        img = Image.new('RGB', (400, 100), color='white')
        draw = ImageDraw.Draw(img)
        
        # 텍스트 추가 (기본 폰트 사용)
        draw.text((10, 10), "Hello World", fill='black')
        draw.text((10, 40), "안녕하세요 테스트입니다", fill='black')
        
        # 임시 파일로 저장
        test_image_path = "test_ocr.png"
        img.save(test_image_path)
        
        try:
            # OCR 테스트
            text = pytesseract.image_to_string(Image.open(test_image_path), lang='kor+eng')
            print(f"OCR 결과: {text}")
        except Exception as ocr_error:
            print(f"한국어+영어 OCR 실패: {ocr_error}")
            # Fallback: 영어만으로 재시도
            try:
                text = pytesseract.image_to_string(Image.open(test_image_path), lang='eng')
                print(f"영어 OCR 결과: {text}")
            except Exception:
                text = "OCR 완전 실패"
                print("OCR 테스트 완전 실패")
        
        # 임시 파일 삭제
        if os.path.exists(test_image_path):
            os.remove(test_image_path)
        
        return text
        
    except Exception as e:
        print(f"OCR 테스트 오류: {e}")
        return None

if __name__ == "__main__":
    print("=== OCR 기능 테스트 ===")
    test_ocr_with_sample()