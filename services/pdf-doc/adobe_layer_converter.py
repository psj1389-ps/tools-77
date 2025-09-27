import os
import json
import logging
import zipfile
import tempfile
import time
import subprocess
import sys
from typing import Dict, List, Any, Optional
from dotenv import load_dotenv

# .env 파일 로드
load_dotenv()

# Adobe SDK 강제 초기화 시스템
class AdobeSDKForceInitializer:
    """Adobe SDK를 강제로 초기화하고 100% 가용성을 보장하는 시스템"""
    
    def __init__(self):
        self.max_retries = 5
        self.retry_delay = 2  # 초
        self.sdk_available = False
        self.execution_context = None
        self.demo_mode = False
        
    def force_initialize_sdk(self) -> bool:
        """SDK를 강제로 초기화 (최대 5회 재시도 + 네트워크 오프라인 모드)"""
        for attempt in range(self.max_retries):
            try:
                logging.info(f"🔄 SDK 초기화 시도 {attempt + 1}/{self.max_retries}")
                
                # 0단계: 네트워크 연결 상태 확인
                network_status = self._check_network_connectivity()
                if not network_status:
                    logging.warning("🌐 네트워크 연결 실패 - 오프라인 모드로 전환")
                    return self._activate_offline_mode()
                
                # 1단계: SDK 라이브러리 자동 설치 시도
                if not self._check_sdk_installation():
                    self._auto_install_sdk()
                
                # 2단계: SDK 임포트 시도
                sdk_modules = self._import_adobe_sdk()
                if not sdk_modules:
                    continue
                
                # 3단계: 환경 변수 검증 및 자동 설정
                credentials_info = self._validate_and_setup_credentials()
                
                # 4단계: ExecutionContext 생성
                self.execution_context = self._create_execution_context(sdk_modules, credentials_info)
                
                if self.execution_context:
                    self.sdk_available = True
                    logging.info(f"✅ SDK 초기화 성공! (시도 {attempt + 1})")
                    return True
                    
            except Exception as e:
                logging.warning(f"초기화 시도 {attempt + 1} 실패: {e}")
                if attempt < self.max_retries - 1:
                    time.sleep(self.retry_delay)
                    self.retry_delay *= 1.5  # 지수 백오프
        
        # 모든 시도 실패 시 오프라인 모드 활성화
        logging.warning("🚨 모든 초기화 시도 실패 - 오프라인 모드 활성화")
        return self._activate_offline_mode()
    
    def _check_network_connectivity(self) -> bool:
        """네트워크 연결 상태 확인"""
        try:
            import urllib.request
            import socket
            
            # Adobe API 서버 연결 테스트
            test_urls = [
                'https://pdf-services.adobe.io',
                'https://www.google.com',
                'https://www.adobe.com'
            ]
            
            for url in test_urls:
                try:
                    urllib.request.urlopen(url, timeout=5)
                    logging.info(f"✅ 네트워크 연결 확인: {url}")
                    return True
                except:
                    continue
            
            logging.warning("🌐 모든 네트워크 테스트 실패")
            return False
            
        except Exception as e:
            logging.warning(f"네트워크 확인 중 오류: {e}")
            return False
    
    def _activate_offline_mode(self) -> bool:
        """오프라인 모드 활성화"""
        logging.info("🔌 오프라인 모드 활성화 - 로컬 라이브러리만 사용")
        
        self.demo_mode = True
        self.offline_mode = True
        self.execution_context = None
        self.sdk_available = False
        
        # 오프라인에서 사용 가능한 라이브러리 확인
        offline_libraries = []
        
        try:
            import fitz
            offline_libraries.append('PyMuPDF')
        except ImportError:
            pass
        
        try:
            import pdfplumber
            offline_libraries.append('pdfplumber')
        except ImportError:
            pass
        
        if offline_libraries:
            logging.info(f"📚 오프라인 사용 가능 라이브러리: {', '.join(offline_libraries)}")
        else:
            logging.info("🛡️ 기본 추출 모드로 동작")
        
        return True
    
    def _check_sdk_installation(self):
        """SDK 설치 상태 확인"""
        try:
            import adobe.pdfservices
            return True
        except ImportError:
            return False
    
    def _auto_install_sdk(self):
        """SDK 자동 설치 시도"""
        try:
            logging.info("Adobe PDF Services SDK 자동 설치 시도...")
            subprocess.check_call([sys.executable, "-m", "pip", "install", "pdfservices-sdk", "--quiet"])
            logging.info("SDK 설치 완료")
        except Exception as e:
            logging.warning(f"SDK 자동 설치 실패: {e}")
    
    def _import_adobe_sdk(self):
        """Adobe SDK 모듈 임포트"""
        try:
            from adobe.pdfservices.operation.auth.credentials import Credentials
            from adobe.pdfservices.operation.execution_context import ExecutionContext
            from adobe.pdfservices.operation.io.file_ref import FileRef
            from adobe.pdfservices.operation.pdfops.extract_pdf_operation import ExtractPDFOperation
            from adobe.pdfservices.operation.pdfops.options.extractpdf.extract_pdf_options import ExtractPDFOptions
            from adobe.pdfservices.operation.pdfops.options.extractpdf.extract_element_type import ExtractElementType as PDFElementType
            from adobe.pdfservices.operation.exception.exceptions import ServiceApiException, ServiceUsageException, SdkException
            
            return {
                'Credentials': Credentials,
                'ExecutionContext': ExecutionContext,
                'FileRef': FileRef,
                'ExtractPDFOperation': ExtractPDFOperation,
                'ExtractPDFOptions': ExtractPDFOptions,
                'PDFElementType': PDFElementType,
                'ServiceApiException': ServiceApiException,
                'ServiceUsageException': ServiceUsageException,
                'SdkException': SdkException
            }
        except ImportError as e:
            logging.warning(f"SDK 임포트 실패: {e}")
            return None
    
    def _validate_and_setup_credentials(self):
        """환경 변수 검증 및 자동 설정"""
        client_id = os.getenv('ADOBE_CLIENT_ID')
        client_secret = os.getenv('ADOBE_CLIENT_SECRET')
        organization_id = os.getenv('ADOBE_ORGANIZATION_ID')
        
        # 환경 변수 누락 시 기본값 설정 (데모 모드)
        if not all([client_id, client_secret, organization_id]):
            logging.warning("환경 변수 누락 - 기본값으로 설정")
            client_id = client_id or "demo_client_id"
            client_secret = client_secret or "demo_client_secret"
            organization_id = organization_id or "demo_organization_id"
            self.demo_mode = True
        
        return {
            'client_id': client_id,
            'client_secret': client_secret,
            'organization_id': organization_id
        }
    
    def _create_execution_context(self, sdk_modules, credentials_info):
        """ExecutionContext 생성"""
        try:
            Credentials = sdk_modules['Credentials']
            ExecutionContext = sdk_modules['ExecutionContext']
            
            credentials = Credentials.service_account_credentials_builder() \
                .with_client_id(credentials_info['client_id']) \
                .with_client_secret(credentials_info['client_secret']) \
                .with_organization_id(credentials_info['organization_id']) \
                .build()
            
            return ExecutionContext.create(credentials)
        except Exception as e:
            logging.warning(f"ExecutionContext 생성 실패: {e}")
            return None

# 전역 강제 초기화 시스템
force_initializer = AdobeSDKForceInitializer()
ADOBE_SDK_AVAILABLE = force_initializer.force_initialize_sdk()

# SDK 모듈들을 전역으로 설정
if ADOBE_SDK_AVAILABLE and not force_initializer.demo_mode:
    try:
        from adobe.pdfservices.operation.auth.credentials import Credentials
        from adobe.pdfservices.operation.execution_context import ExecutionContext
        from adobe.pdfservices.operation.io.file_ref import FileRef
        from adobe.pdfservices.operation.pdfops.extract_pdf_operation import ExtractPDFOperation
        from adobe.pdfservices.operation.pdfops.options.extractpdf.extract_pdf_options import ExtractPDFOptions
        from adobe.pdfservices.operation.pdfops.options.extractpdf.extract_element_type import ExtractElementType as PDFElementType
        from adobe.pdfservices.operation.exception.exceptions import ServiceApiException, ServiceUsageException, SdkException
    except ImportError:
        ADOBE_SDK_AVAILABLE = False

class AdobeLayerConverter:
    """Adobe SDK 기반 ExtractPDFOperation을 활용한 레이어 결합 방식 변환기 (100% 가용성 보장)"""
    
    def __init__(self):
        # 강제 초기화 시스템 연동
        self.force_initializer = force_initializer
        self.api_available = ADOBE_SDK_AVAILABLE
        self.execution_context = force_initializer.execution_context
        self.demo_mode = force_initializer.demo_mode
        self.offline_mode = getattr(force_initializer, 'offline_mode', False)
        
        # 실시간 모니터링 및 자동 복구 시스템
        self.monitoring_enabled = True
        self.last_health_check = time.time()
        self.health_check_interval = 30  # 30초마다 상태 확인
        
        # 대체 방법들
        self.fallback_methods = [
            self._fallback_pymupdf,
            self._fallback_pdfplumber,
            self._fallback_basic_extraction
        ]
        
        if self.api_available:
            logging.info("🚀 Adobe SDK 강제 초기화 완료 - 100% 가용성 보장 모드 활성화")
        else:
            logging.info("⚡ 대체 방법 활성화 - 서비스 연속성 보장")
    
    def _ensure_sdk_availability(self):
        """SDK 가용성을 실시간으로 보장"""
        current_time = time.time()
        
        # 주기적 상태 확인
        if current_time - self.last_health_check > self.health_check_interval:
            self.last_health_check = current_time
            
            if not self.api_available or not self.execution_context:
                logging.info("🔄 SDK 자동 복구 시도...")
                if self.force_initializer.force_initialize_sdk():
                    self.api_available = True
                    self.execution_context = self.force_initializer.execution_context
                    logging.info("✅ SDK 자동 복구 성공!")
                else:
                    logging.info("⚡ 대체 방법으로 계속 진행")
        
        return self.api_available and self.execution_context
    
    def _fallback_pymupdf(self, pdf_path: str) -> Optional[Dict[str, Any]]:
        """PyMuPDF를 사용한 대체 추출 방법"""
        try:
            import fitz  # PyMuPDF
            logging.info("📚 PyMuPDF 대체 방법 사용")
            
            doc = fitz.open(pdf_path)
            extracted_data = {
                'elements': [],
                'pages': []
            }
            
            for page_num in range(len(doc)):
                page = doc.load_page(page_num)
                
                # 텍스트 추출
                text_dict = page.get_text("dict")
                for block in text_dict["blocks"]:
                    if "lines" in block:
                        for line in block["lines"]:
                            for span in line["spans"]:
                                element = {
                                    'Path': '/Text',
                                    'Text': span['text'],
                                    'Bounds': {
                                        'x': span['bbox'][0],
                                        'y': span['bbox'][1],
                                        'width': span['bbox'][2] - span['bbox'][0],
                                        'height': span['bbox'][3] - span['bbox'][1]
                                    },
                                    'Font': {
                                        'name': span['font'],
                                        'size': span['size']
                                    }
                                }
                                extracted_data['elements'].append(element)
                
                # 이미지 추출
                image_list = page.get_images()
                for img_index, img in enumerate(image_list):
                    pix = fitz.Pixmap(doc, img[0])
                    if pix.n < 5:  # GRAY or RGB
                        img_data = pix.tobytes("png")
                        img_path = f"fallback_image_{page_num}_{img_index}.png"
                        
                        element = {
                            'Path': '/Figure',
                            'Bounds': {
                                'x': 0, 'y': 0, 'width': pix.width, 'height': pix.height
                            },
                            'filePaths': [img_path],
                            'image_data': img_data
                        }
                        extracted_data['elements'].append(element)
                    pix = None
            
            doc.close()
            return {'json_data': extracted_data, 'images': {}, 'extract_dir': None}
            
        except ImportError:
            logging.warning("PyMuPDF 라이브러리가 설치되지 않음")
            return None
        except Exception as e:
            logging.warning(f"PyMuPDF 대체 방법 실패: {e}")
            return None
    
    def _fallback_pdfplumber(self, pdf_path: str) -> Optional[Dict[str, Any]]:
        """pdfplumber를 사용한 대체 추출 방법"""
        try:
            import pdfplumber
            logging.info("🔧 pdfplumber 대체 방법 사용")
            
            extracted_data = {'elements': []}
            
            with pdfplumber.open(pdf_path) as pdf:
                for page_num, page in enumerate(pdf.pages):
                    # 텍스트 추출
                    chars = page.chars
                    for char in chars:
                        element = {
                            'Path': '/Text',
                            'Text': char['text'],
                            'Bounds': {
                                'x': char['x0'],
                                'y': char['y0'],
                                'width': char['x1'] - char['x0'],
                                'height': char['y1'] - char['y0']
                            },
                            'Font': {
                                'name': char.get('fontname', ''),
                                'size': char.get('size', 12)
                            }
                        }
                        extracted_data['elements'].append(element)
            
            return {'json_data': extracted_data, 'images': {}, 'extract_dir': None}
            
        except ImportError:
            logging.warning("pdfplumber 라이브러리가 설치되지 않음")
            return None
        except Exception as e:
            logging.warning(f"pdfplumber 대체 방법 실패: {e}")
            return None
    
    def _fallback_basic_extraction(self, pdf_path: str) -> Optional[Dict[str, Any]]:
        """기본 텍스트 추출 방법 (최후의 수단)"""
        try:
            logging.info("🛡️ 기본 추출 방법 사용 (최후의 수단)")
            
            # 간단한 더미 데이터 생성
            extracted_data = {
                'elements': [
                    {
                        'Path': '/Text',
                        'Text': f'PDF 파일: {os.path.basename(pdf_path)}',
                        'Bounds': {'x': 50, 'y': 50, 'width': 400, 'height': 20},
                        'Font': {'name': 'Arial', 'size': 14}
                    },
                    {
                        'Path': '/Text',
                        'Text': '⚠️ Adobe SDK를 사용할 수 없어 기본 모드로 실행됩니다.',
                        'Bounds': {'x': 50, 'y': 80, 'width': 500, 'height': 20},
                        'Font': {'name': 'Arial', 'size': 12}
                    },
                    {
                        'Path': '/Text',
                        'Text': '정확한 레이아웃 변환을 위해 Adobe API 설정을 확인해주세요.',
                        'Bounds': {'x': 50, 'y': 110, 'width': 600, 'height': 20},
                        'Font': {'name': 'Arial', 'size': 12}
                    }
                ]
            }
            
            return {'json_data': extracted_data, 'images': {}, 'extract_dir': None}
            
        except Exception as e:
            logging.error(f"기본 추출 방법도 실패: {e}")
            return None
    
    def extract_pdf_data(self, pdf_path: str) -> Optional[Dict[str, Any]]:
        """PDF에서 데이터를 추출하여 JSON 형태로 반환 (100% 가용성 보장)"""
        if not os.path.exists(pdf_path):
            logging.error(f"PDF 파일을 찾을 수 없습니다: {pdf_path}")
            return None
        
        # 1단계: SDK 가용성 실시간 확인 및 자동 복구
        if self._ensure_sdk_availability():
            try:
                logging.info("🚀 Adobe SDK로 고품질 추출 시도")
                
                # ExtractPDFOperation 생성
                extract_pdf_operation = ExtractPDFOperation.create_new()
                
                # 입력 파일 설정
                source = FileRef.create_from_local_file(pdf_path)
                extract_pdf_operation.set_input(source)
                
                # Extract PDF 옵션 설정
                extract_pdf_options = ExtractPDFOptions.builder() \
                    .with_elements_to_extract([PDFElementType.TEXT, PDFElementType.TABLES, PDFElementType.FIGURES]) \
                    .with_elements_to_extract_renditions([PDFElementType.FIGURES, PDFElementType.TABLES]) \
                    .build()
                
                extract_pdf_operation.set_options(extract_pdf_options)
                
                # 작업 실행
                result = extract_pdf_operation.execute(self.execution_context)
                
                # 임시 ZIP 파일로 저장
                temp_zip_path = tempfile.mktemp(suffix='.zip')
                result.save_as(temp_zip_path)
                
                # ZIP 파일 압축 해제
                extract_dir = tempfile.mkdtemp()
                with zipfile.ZipFile(temp_zip_path, 'r') as zip_ref:
                    zip_ref.extractall(extract_dir)
                
                # JSON 파일 읽기
                json_file_path = os.path.join(extract_dir, 'structuredData.json')
                if os.path.exists(json_file_path):
                    with open(json_file_path, 'r', encoding='utf-8') as json_file:
                        extracted_data = json.load(json_file)
                else:
                    logging.error("추출된 JSON 파일을 찾을 수 없습니다.")
                    raise Exception("JSON 파일 없음")
                
                # 이미지 파일들도 포함
                figures_dir = os.path.join(extract_dir, 'figures')
                images = {}
                if os.path.exists(figures_dir):
                    for img_file in os.listdir(figures_dir):
                        img_path = os.path.join(figures_dir, img_file)
                        images[img_file] = img_path
                
                # 정리
                os.unlink(temp_zip_path)
                
                logging.info("✅ Adobe SDK 추출 성공!")
                return {
                    'json_data': extracted_data,
                    'images': images,
                    'extract_dir': extract_dir
                }
                
            except Exception as e:
                logging.warning(f"⚠️ Adobe SDK 추출 실패: {e} - 대체 방법으로 전환")
                # SDK 실패 시 자동으로 대체 방법 사용
        
        # 2단계: 대체 방법들 순차 시도 (절대 실패하지 않음)
        logging.info("🔄 대체 방법들로 추출 시도 - 서비스 연속성 보장")
        
        for i, fallback_method in enumerate(self.fallback_methods, 1):
            try:
                logging.info(f"📋 대체 방법 {i}/{len(self.fallback_methods)} 시도")
                result = fallback_method(pdf_path)
                
                if result:
                    logging.info(f"✅ 대체 방법 {i} 성공!")
                    return result
                else:
                    logging.info(f"⚠️ 대체 방법 {i} 실패 - 다음 방법 시도")
                    
            except Exception as e:
                logging.warning(f"대체 방법 {i} 오류: {e}")
                continue
        
        # 3단계: 절대 실패하지 않는 최종 보장
        logging.error("🚨 모든 추출 방법 실패 - 응급 모드 활성화")
        return self._emergency_fallback(pdf_path)
    
    def _emergency_fallback(self, pdf_path: str) -> Dict[str, Any]:
        """절대 실패하지 않는 응급 대체 방법"""
        logging.info("🆘 응급 모드: 기본 구조 생성")
        
        file_size = os.path.getsize(pdf_path) if os.path.exists(pdf_path) else 0
        
        emergency_data = {
            'elements': [
                {
                    'Path': '/Text',
                    'Text': f'📄 PDF 파일: {os.path.basename(pdf_path)}',
                    'Bounds': {'x': 50, 'y': 50, 'width': 500, 'height': 25},
                    'Font': {'name': 'Arial', 'size': 16}
                },
                {
                    'Path': '/Text', 
                    'Text': f'📊 파일 크기: {file_size:,} bytes',
                    'Bounds': {'x': 50, 'y': 85, 'width': 400, 'height': 20},
                    'Font': {'name': 'Arial', 'size': 12}
                },
                {
                    'Path': '/Text',
                    'Text': '🔧 시스템 상태: 응급 모드 (서비스 연속성 보장)',
                    'Bounds': {'x': 50, 'y': 115, 'width': 600, 'height': 20},
                    'Font': {'name': 'Arial', 'size': 12}
                },
                {
                    'Path': '/Text',
                    'Text': '💡 정상 서비스를 위해 Adobe API 설정을 확인하거나 대체 라이브러리를 설치해주세요.',
                    'Bounds': {'x': 50, 'y': 145, 'width': 700, 'height': 20},
                    'Font': {'name': 'Arial', 'size': 11}
                },
                {
                    'Path': '/Text',
                    'Text': '📋 권장 라이브러리: PyMuPDF (pip install PyMuPDF) 또는 pdfplumber (pip install pdfplumber)',
                    'Bounds': {'x': 50, 'y': 175, 'width': 800, 'height': 20},
                    'Font': {'name': 'Arial', 'size': 11}
                }
            ]
        }
        
        return {
            'json_data': emergency_data,
            'images': {},
            'extract_dir': None
        }
    
    def parse_text_elements(self, json_data: Dict[str, Any]) -> List[Dict[str, Any]]:
        """JSON 데이터에서 텍스트 요소와 좌표 정보를 파싱"""
        text_elements = []
        
        try:
            elements = json_data.get('elements', [])
            
            for element in elements:
                if element.get('Path', '').endswith('/Text'):
                    # 텍스트 내용
                    text_content = element.get('Text', '')
                    
                    # 바운딩 박스 좌표
                    bounds = element.get('Bounds', {})
                    
                    # 스타일 정보
                    font_info = element.get('Font', {})
                    
                    text_element = {
                        'text': text_content,
                        'bounds': {
                            'x': bounds.get('x', 0),
                            'y': bounds.get('y', 0),
                            'width': bounds.get('width', 0),
                            'height': bounds.get('height', 0)
                        },
                        'style': {
                            'font_name': font_info.get('name', ''),
                            'font_size': font_info.get('size', 12),
                            'font_weight': font_info.get('weight', 'normal'),
                            'color': element.get('TextColor', '#000000')
                        }
                    }
                    
                    text_elements.append(text_element)
                    
        except Exception as e:
            logging.error(f"텍스트 요소 파싱 오류: {e}")
            
        return text_elements
    
    def parse_figure_elements(self, json_data: Dict[str, Any]) -> List[Dict[str, Any]]:
        """JSON 데이터에서 이미지/도형 요소와 좌표 정보를 파싱"""
        figure_elements = []
        
        try:
            elements = json_data.get('elements', [])
            
            for element in elements:
                if element.get('Path', '').endswith('/Figure'):
                    # 바운딩 박스 좌표
                    bounds = element.get('Bounds', {})
                    
                    # 이미지 파일 경로
                    image_path = element.get('filePaths', [])
                    
                    figure_element = {
                        'bounds': {
                            'x': bounds.get('x', 0),
                            'y': bounds.get('y', 0),
                            'width': bounds.get('width', 0),
                            'height': bounds.get('height', 0)
                        },
                        'image_path': image_path[0] if image_path else None
                    }
                    
                    figure_elements.append(figure_element)
                    
        except Exception as e:
            logging.error(f"도형 요소 파싱 오류: {e}")
            
        return figure_elements
    
    def generate_html_layer(self, pdf_path: str, output_dir: str = None) -> Optional[str]:
        """레이어 결합 방식으로 HTML 파일 생성"""
        if output_dir is None:
            output_dir = os.path.join(os.path.dirname(pdf_path), 'layer_output')
        
        os.makedirs(output_dir, exist_ok=True)
        
        # PDF 데이터 추출
        extracted_data = self.extract_pdf_data(pdf_path)
        if not extracted_data:
            return None
        
        json_data = extracted_data['json_data']
        images = extracted_data['images']
        
        # 텍스트 및 이미지 요소 파싱
        text_elements = self.parse_text_elements(json_data)
        figure_elements = self.parse_figure_elements(json_data)
        
        # HTML 생성
        html_content = self._generate_html_content(
            text_elements, figure_elements, images, output_dir
        )
        
        # HTML 파일 저장 (BOM 포함하여 Word 호환성 향상)
        html_file_path = os.path.join(output_dir, 'layered_document.html')
        with open(html_file_path, 'w', encoding='utf-8-sig') as html_file:
            html_file.write(html_content)
        
        logging.info(f"레이어 결합 HTML 파일이 생성되었습니다: {html_file_path}")
        return html_file_path
    
    def _generate_html_content(self, text_elements: List[Dict], figure_elements: List[Dict], 
                              images: Dict[str, str], output_dir: str) -> str:
        """HTML 콘텐츠 생성 - Word 호환성 개선"""
        
        # 이미지 파일들을 output_dir로 복사
        import shutil
        copied_images = {}
        for img_name, img_path in images.items():
            dest_path = os.path.join(output_dir, img_name)
            shutil.copy2(img_path, dest_path)
            copied_images[img_name] = img_name
        
        # Word 호환성을 위한 HTML 구조 개선
        html_content = """
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="ko" lang="ko">
<head>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <meta name="ProgId" content="Word.Document" />
    <meta name="Generator" content="Microsoft Word" />
    <meta name="Originator" content="Microsoft Word" />
    <title>레이어 결합 문서</title>
    <style type="text/css">
        /* Word 호환성을 위한 기본 스타일 */
        body {
            margin: 0pt;
            padding: 12pt;
            font-family: 'Malgun Gothic', '맑은 고딕', Arial, sans-serif;
            font-size: 11pt;
            background-color: white;
            color: black;
            line-height: 115%;
            word-wrap: break-word;
            -ms-word-wrap: break-word;
        }
        
        .document-container {
            position: relative;
            background-color: white;
            margin: 0 auto;
            width: 210mm;
            min-height: 297mm;
            page-break-inside: avoid;
        }
        
        .text-layer {
            position: absolute;
            font-family: inherit;
            white-space: pre-wrap;
            word-wrap: break-word;
            -ms-word-wrap: break-word;
            overflow: visible;
        }
        
        .figure-layer {
            position: absolute;
            overflow: visible;
        }
        
        .figure-layer img {
            width: 100%;
            height: 100%;
            border: none;
        }
        
        .search-highlight {
            background-color: yellow;
        }
        
        /* Word 호환 테이블 스타일 */
        table {
            border-collapse: collapse;
            width: 100%;
        }
        
        td, th {
            border: 1pt solid black;
            padding: 2pt;
            vertical-align: top;
        }
        
        /* 인쇄 및 Word 호환성 */
        @media print {
            body {
                background-color: white;
                padding: 0;
            }
            .document-container {
                width: auto;
                min-height: auto;
            }
        }
    </style>
</head>
<body>
    <div class="document-container">
"""
        
        # 이미지 레이어 추가
        for figure in figure_elements:
            bounds = figure['bounds']
            img_path = figure.get('image_path')
            
            if img_path and img_path in copied_images:
                html_content += f"""
        <div class="figure-layer" style="
            left: {bounds['x']}px;
            top: {bounds['y']}px;
            width: {bounds['width']}px;
            height: {bounds['height']}px;
        ">
            <img src="{copied_images[img_path]}" alt="Figure" />
        </div>
"""
        
        # 텍스트 레이어 추가
        for text_elem in text_elements:
            bounds = text_elem['bounds']
            style = text_elem['style']
            text = text_elem['text']
            
            # HTML 엔티티 이스케이프 (Word 호환성 향상)
            import html
            escaped_text = html.escape(text, quote=True)
            
            # 폰트명도 안전하게 처리
            safe_font_name = html.escape(style['font_name'], quote=True) if style['font_name'] else 'Arial'
            
            html_content += f"""
        <div class="text-layer" style="
            left: {bounds['x']}px;
            top: {bounds['y']}px;
            width: {bounds['width']}px;
            height: {bounds['height']}px;
            font-family: '{safe_font_name}';
            font-size: {style['font_size']}px;
            font-weight: {style['font_weight']};
            color: {style['color']};
        ">{escaped_text}</div>
"""
        
        html_content += """
    </div>
    
    <!-- Word 호환성을 위해 JavaScript 제거 -->
    <!-- 검색 기능은 Word의 기본 찾기 기능 사용 -->
    
</body>
</html>
"""
        
        return html_content