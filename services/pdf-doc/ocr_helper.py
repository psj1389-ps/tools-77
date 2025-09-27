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

def convert_pdf_to_images_and_extract_text(pdf_path, lang='kor+eng'):
    """PDF를 이미지로 변환하고 OCR로 텍스트를 추출합니다."""
    try:
        # Tesseract 가용성 확인
        if not OCR_AVAILABLE:
            print("Tesseract OCR을 사용할 수 없습니다.")
            return []
        
        # Tesseract 경로 설정 (Render 환경 고려)
        try:
            pytesseract.get_tesseract_version()
        except Exception:
            # Render 환경에서 Tesseract 경로 설정
            pytesseract.pytesseract.tesseract_cmd = '/usr/bin/tesseract'
        
        # PDF를 이미지로 변환 (메모리 최적화 - Render 환경 고려)
        images = convert_from_path(
            pdf_path, 
            dpi=150,  # DPI 낮춤으로 메모리 사용량 감소
            fmt='PNG',
            thread_count=1,  # 단일 스레드로 메모리 사용량 제어
            use_pdftocairo=False  # 메모리 효율적인 변환 방식 사용
        )
        
        extracted_texts = []
        for i, image in enumerate(images):
            print(f"페이지 {i+1} OCR 처리 중...")
            
            try:
                # 이미지 크기가 너무 크면 리사이즈 (메모리 절약 - Render 환경 고려)
                max_size = 1500  # Render 환경에서 메모리 제한 고려하여 축소
                if image.width > max_size or image.height > max_size:
                    image.thumbnail((max_size, max_size), Image.Resampling.LANCZOS)
                
                # 이미지 모드 최적화 (메모리 사용량 감소)
                if image.mode not in ('RGB', 'L'):
                    image = image.convert('RGB')
                
                # OCR 설정 (Render 환경 최적화)
                custom_config = r'--oem 3 --psm 6 -c tessedit_do_invert=0'
                
                # OCR 실행 (타임아웃 설정)
                text = pytesseract.image_to_string(image, lang=lang, config=custom_config, timeout=30)
                extracted_texts.append({
                    'page': i + 1,
                    'text': text.strip()
                })
                
            except Exception as e:
                print(f"페이지 {i+1} OCR 처리 중 오류: {e}")
                extracted_texts.append({
                    'page': i + 1,
                    'text': ''
                })
            finally:
                # 메모리 정리
                if image:
                    image.close()
        
        return extracted_texts
        
    except Exception as e:
        print(f"PDF OCR 처리 중 오류 발생: {e}")
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

def extract_text_from_image_with_ocr(image_path, lang='kor+eng'):
    """이미지에서 OCR을 사용하여 텍스트를 추출합니다."""
    try:
        # Tesseract 가용성 확인
        if not OCR_AVAILABLE:
            print("Tesseract OCR을 사용할 수 없습니다.")
            return ""
            
        # Tesseract 경로 설정 (Render 환경 고려)
        try:
            pytesseract.get_tesseract_version()
        except Exception:
            # Render 환경에서 Tesseract 경로 설정
            pytesseract.pytesseract.tesseract_cmd = '/usr/bin/tesseract'
        
        # 이미지 로드 및 전처리 (메모리 최적화)
        image = None
        try:
            image = Image.open(image_path)
            
            # 이미지 크기가 너무 크면 리사이즈 (메모리 절약 - Render 환경 고려)
            max_size = 1500  # Render 환경에서 메모리 제한 고려하여 축소
            if image.width > max_size or image.height > max_size:
                image.thumbnail((max_size, max_size), Image.Resampling.LANCZOS)
            
            # 이미지 모드 최적화 (메모리 사용량 감소)
            if image.mode not in ('RGB', 'L'):
                image = image.convert('RGB')
            
            # OCR 설정 (Render 환경 최적화)
            custom_config = r'--oem 3 --psm 6 -c tessedit_do_invert=0'
            
            # OCR 실행 (타임아웃 설정)
            text = pytesseract.image_to_string(image, lang=lang, config=custom_config, timeout=30)
            
            return text.strip()
            
        finally:
            # 메모리 정리
            if image:
                image.close()
        
    except pytesseract.TesseractError as e:
        print(f"Tesseract OCR 오류: {e}")
        return ""
    except Exception as e:
        print(f"OCR 처리 중 오류 발생: {e}")
        return ""

if __name__ == "__main__":
    print("=== OCR 기능 테스트 ===")
    test_ocr_with_sample()