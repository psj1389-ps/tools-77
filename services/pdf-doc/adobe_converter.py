import os
import logging
import re
from urllib.parse import unquote
from dotenv import load_dotenv

# .env 파일 로드
load_dotenv()
try:
    from adobe.pdfservices.operation.auth.service_principal_credentials import ServicePrincipalCredentials
    from adobe.pdfservices.operation.pdf_services import PDFServices
    from adobe.pdfservices.operation.pdfjobs.jobs.export_pdf_job import ExportPDFJob
    from adobe.pdfservices.operation.pdfjobs.params.export_pdf.export_pdf_params import ExportPDFParams
    from adobe.pdfservices.operation.pdfjobs.params.export_pdf.export_pdf_target_format import ExportPDFTargetFormat
    from adobe.pdfservices.operation.io.cloud_asset import CloudAsset
    from adobe.pdfservices.operation.io.stream_asset import StreamAsset
    ADOBE_SDK_AVAILABLE = True
except ImportError:
    ADOBE_SDK_AVAILABLE = False
    logging.warning("Adobe PDF Services SDK를 사용할 수 없습니다. 폴백 방식을 사용합니다.")

class AdobePDFConverter:
    def __init__(self):
        self.api_available = ADOBE_SDK_AVAILABLE
        self.execution_context = None
        if not self.api_available:
            logging.warning("Adobe PDF Services SDK를 사용할 수 없습니다.")
            return
            
        try:
            # 자격 증명 설정 (SDK 4.2 호환)
            credentials = ServicePrincipalCredentials(
                client_id=os.getenv('ADOBE_CLIENT_ID'),
                client_secret=os.getenv('ADOBE_CLIENT_SECRET')
            )
            
            # PDF Services API 초기화
            self.pdf_services_api = PDFServices(credentials=credentials)
            self.execution_context = self.pdf_services_api  # execution_context 속성 추가
            logging.info("Adobe PDF Services API가 성공적으로 초기화되었습니다.")

        except Exception as e:
            self.api_available = False
            self.execution_context = None
            logging.error(f"Adobe PDF Services API 초기화 실패: {e}")

    def _get_safe_filename(self, pdf_path):
        """원본 파일명에서 안전한 파일명 추출 (확장자 제거, 특수문자 처리)"""
        try:
            # 파일명 추출
            basename = os.path.basename(pdf_path)
            # 확장자 제거
            name_without_ext = os.path.splitext(basename)[0]
            
            # URL 디코딩 (한글 파일명 처리)
            try:
                name_without_ext = unquote(name_without_ext, encoding='utf-8')
            except:
                pass
            
            # 파일명에 사용할 수 없는 문자 제거/대체
            safe_chars = re.sub(r'[<>:"/\\|?*]', '_', name_without_ext)
            # 연속된 언더스코어 정리
            safe_chars = re.sub(r'_+', '_', safe_chars)
            # 앞뒤 공백 및 언더스코어 제거
            safe_chars = safe_chars.strip('_ ')
            
            # 빈 문자열이면 기본값 사용
            if not safe_chars:
                safe_chars = 'converted_document'
                
            return safe_chars
        except Exception as e:
            logging.warning(f"파일명 처리 중 오류 발생: {e}")
            return 'converted_document'

    def _get_unique_filename(self, file_path):
        """중복 파일명 처리 - 파일이 존재하면 번호를 추가"""
        if not os.path.exists(file_path):
            return file_path
        
        directory = os.path.dirname(file_path)
        filename = os.path.basename(file_path)
        name, ext = os.path.splitext(filename)
        
        counter = 1
        while True:
            new_filename = f"{name}_{counter}{ext}"
            new_path = os.path.join(directory, new_filename)
            if not os.path.exists(new_path):
                return new_path
            counter += 1
            
            # 무한 루프 방지
            if counter > 1000:
                import time
                timestamp = int(time.time())
                new_filename = f"{name}_{timestamp}{ext}"
                return os.path.join(directory, new_filename)

    def convert_to_docx(self, pdf_path):
        """Adobe API로 완벽한 PDF→DOCX 변환"""
        if not self.api_available:
            logging.error("Adobe API가 초기화되지 않아 변환을 진행할 수 없습니다.")
            return None

        try:
            # 원본 파일명을 유지하여 출력 파일명 생성
            original_name = self._get_safe_filename(pdf_path)
            filename = f"{original_name}.docx"
            outputs_dir = os.path.join(os.path.dirname(pdf_path), '..', 'outputs')
            os.makedirs(outputs_dir, exist_ok=True)
            output_path = self._get_unique_filename(os.path.join(outputs_dir, filename))
            
            # PDF to Word 변환 실행
            with open(pdf_path, 'rb') as input_stream:
                # StreamAsset 생성
                input_asset = StreamAsset(input_stream, "application/pdf")
                
                # Export PDF 파라미터 설정
                export_pdf_params = ExportPDFParams(target_format=ExportPDFTargetFormat.DOCX)
                
                # Export PDF Job 생성
                export_pdf_job = ExportPDFJob(input_asset=input_asset, export_pdf_params=export_pdf_params)
                
                print(">>> [DEBUG 1] Adobe 변환 함수 진입")
                try:
                    print(">>> [DEBUG 2] try 블록 진입, execute() 호출 직전")
                    
                    # Job 실행 - 실제 Adobe API 실행 지점
                    location = self.pdf_services_api.submit(export_pdf_job)
                    pdf_services_response = self.pdf_services_api.get_job_result(location, CloudAsset)
                    
                    print(">>> [DEBUG 3] execute() 호출 성공")
                    conversion_success = True  # 성공했음을 표시
                    
                except ServiceApiException as e:
                    # Adobe API 관련 에러 (가장 흔함)
                    print("!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!")
                    print(f"❌ Adobe ServiceApiException 발생: {e}")
                    print(f"    - Status Code: {e.status_code if hasattr(e, 'status_code') else 'N/A'}")
                    print(f"    - Error Code: {e.error_code if hasattr(e, 'error_code') else 'N/A'}")
                    print(f"    - Error Message: {e.message if hasattr(e, 'message') else str(e)}")
                    print("!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!")
                    conversion_success = False
                    raise  # 기존 예외 처리로 전달
                    
                except Exception as e:
                    # 그 외 모든 예상치 못한 에러
                    print("!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!")
                    print(f"❌ 변환 중 알 수 없는 예외 발생: {str(e)}")
                    import traceback
                    traceback.print_exc()
                    print("!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!")
                    conversion_success = False
                    raise  # 기존 예외 처리로 전달
                    
                print(">>> [DEBUG 4] Adobe 변환 함수 종료")
                
                # 결과 다운로드
                result_asset = pdf_services_response.get_result().get_download_uri()
                
                # 결과를 파일로 저장
                import requests
                response = requests.get(result_asset)
                with open(output_path, 'wb') as output_stream:
                    output_stream.write(response.content)
            
            logging.info(f"Adobe API를 사용하여 PDF를 DOCX로 성공적으로 변환했습니다: {output_path}")
            return output_path
            
        except ServiceApiException as e:
            logging.error(f"Adobe API ServiceApiException 발생:")
            logging.error(f"  - 오류 메시지: {str(e)}")
            logging.error(f"  - 오류 타입: {type(e).__name__}")
            if hasattr(e, 'status_code'):
                logging.error(f"  - HTTP 상태 코드: {e.status_code}")
            if hasattr(e, 'error_code'):
                logging.error(f"  - Adobe 오류 코드: {e.error_code}")
            if hasattr(e, 'message'):
                logging.error(f"  - 상세 메시지: {e.message}")
            
            # HTTP 400 오류에 대한 특별 처리
            if hasattr(e, 'status_code') and e.status_code == 400:
                logging.error("  - HTTP 400 오류: 요청 매개변수나 파일 형식을 확인하세요")
            
            return None
        except Exception as e:
            logging.error(f"Adobe API 변환 오류: {e}")
            # 폴백 처리
            return None
    
    def convert_to_docx_optimized(self, pdf_path, analysis_result):
        """방향별 최적화가 적용된 Adobe API PDF→DOCX 변환"""
        if not self.api_available:
            logging.error("Adobe API가 초기화되지 않아 변환을 진행할 수 없습니다.")
            return None

        try:
            orientation_info = analysis_result.get("orientation", {})
            orientation = orientation_info.get("orientation", "portrait")
            
            # 원본 파일명을 유지하여 출력 파일명 생성
            original_name = self._get_safe_filename(pdf_path)
            filename = f"{original_name}_{orientation}_adobe.docx"
            outputs_dir = os.path.join(os.path.dirname(pdf_path), '..', 'outputs')
            os.makedirs(outputs_dir, exist_ok=True)
            output_path = self._get_unique_filename(os.path.join(outputs_dir, filename))
            
            # PDF to Word 변환 실행
            with open(pdf_path, 'rb') as input_stream:
                # StreamAsset 생성
                input_asset = StreamAsset(input_stream, "application/pdf")
                
                # Export PDF 파라미터 설정 (방향별 최적화)
                export_pdf_params = ExportPDFParams(target_format=ExportPDFTargetFormat.DOCX)
                
                # Export PDF Job 생성
                export_pdf_job = ExportPDFJob(input_asset=input_asset, export_pdf_params=export_pdf_params)
                
                print(">>> [DEBUG 1] Adobe 최적화 변환 함수 진입")
                try:
                    print(">>> [DEBUG 2] try 블록 진입, execute() 호출 직전")
                    
                    # Job 실행 - 실제 Adobe API 실행 지점
                    location = self.pdf_services_api.submit(export_pdf_job)
                    pdf_services_response = self.pdf_services_api.get_job_result(location, CloudAsset)
                    
                    print(">>> [DEBUG 3] execute() 호출 성공")
                    conversion_success = True  # 성공했음을 표시
                    
                except ServiceApiException as e:
                    # Adobe API 관련 에러 (가장 흔함)
                    print("!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!")
                    print(f"❌ Adobe ServiceApiException 발생: {e}")
                    print(f"    - Status Code: {e.status_code if hasattr(e, 'status_code') else 'N/A'}")
                    print(f"    - Error Code: {e.error_code if hasattr(e, 'error_code') else 'N/A'}")
                    print(f"    - Error Message: {e.message if hasattr(e, 'message') else str(e)}")
                    print("!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!")
                    conversion_success = False
                    raise  # 기존 예외 처리로 전달
                    
                except Exception as e:
                    # 그 외 모든 예상치 못한 에러
                    print("!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!")
                    print(f"❌ 변환 중 알 수 없는 예외 발생: {str(e)}")
                    import traceback
                    traceback.print_exc()
                    print("!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!")
                    conversion_success = False
                    raise  # 기존 예외 처리로 전달
                    
                print(">>> [DEBUG 4] Adobe 최적화 변환 함수 종료")
                
                # 결과 다운로드
                result_asset = pdf_services_response.get_result().get_download_uri()
                
                # 결과를 파일로 저장
                import requests
                response = requests.get(result_asset)
                with open(output_path, 'wb') as output_stream:
                    output_stream.write(response.content)
            
            logging.info(f"Adobe API를 사용하여 {orientation} PDF를 DOCX로 성공적으로 변환했습니다: {output_path}")
            return output_path
            
        except ServiceApiException as e:
            logging.error(f"Adobe API 최적화 변환 ServiceApiException 발생:")
            logging.error(f"  - 오류 메시지: {str(e)}")
            logging.error(f"  - 오류 타입: {type(e).__name__}")
            if hasattr(e, 'status_code'):
                logging.error(f"  - HTTP 상태 코드: {e.status_code}")
            if hasattr(e, 'error_code'):
                logging.error(f"  - Adobe 오류 코드: {e.error_code}")
            if hasattr(e, 'message'):
                logging.error(f"  - 상세 메시지: {e.message}")
            
            # HTTP 400 오류에 대한 특별 처리
            if hasattr(e, 'status_code') and e.status_code == 400:
                logging.error("  - HTTP 400 오류: 요청 매개변수나 파일 형식을 확인하세요")
            
            return None
        except Exception as e:
            logging.error(f"Adobe API 최적화 변환 오류: {e}")
            return None
    
    def convert_to_docx_official(self, pdf_path, analysis_result):
        """공문서 특화 Adobe API PDF→DOCX 변환"""
        if not self.api_available:
            logging.error("Adobe API가 초기화되지 않아 변환을 진행할 수 없습니다.")
            return None

        try:
            official_info = analysis_result.get("official_document", {})
            confidence = official_info.get("confidence", 0)
            
            # 원본 파일명을 유지하여 출력 파일명 생성
            original_name = self._get_safe_filename(pdf_path)
            filename = f"{original_name}_공문서_adobe.docx"
            outputs_dir = os.path.join(os.path.dirname(pdf_path), '..', 'outputs')
            os.makedirs(outputs_dir, exist_ok=True)
            output_path = self._get_unique_filename(os.path.join(outputs_dir, filename))
            
            # PDF to Word 변환 실행 (공문서 최적화)
            with open(pdf_path, 'rb') as input_stream:
                # StreamAsset 생성
                input_asset = StreamAsset(input_stream, "application/pdf")
                
                # Export PDF 파라미터 설정 (공문서 레이아웃 보존 최우선)
                export_pdf_params = ExportPDFParams(target_format=ExportPDFTargetFormat.DOCX)
                
                # Export PDF Job 생성
                export_pdf_job = ExportPDFJob(input_asset=input_asset, export_pdf_params=export_pdf_params)
                
                print(">>> [DEBUG 1] Adobe 공식 변환 함수 진입")
                try:
                    print(">>> [DEBUG 2] try 블록 진입, execute() 호출 직전")
                    
                    # Job 실행 - 실제 Adobe API 실행 지점
                    location = self.pdf_services_api.submit(export_pdf_job)
                    pdf_services_response = self.pdf_services_api.get_job_result(location, CloudAsset)
                    
                    print(">>> [DEBUG 3] execute() 호출 성공")
                    conversion_success = True  # 성공했음을 표시
                    
                except ServiceApiException as e:
                    # Adobe API 관련 에러 (가장 흔함)
                    print("!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!")
                    print(f"❌ Adobe ServiceApiException 발생: {e}")
                    print(f"    - Status Code: {e.status_code if hasattr(e, 'status_code') else 'N/A'}")
                    print(f"    - Error Code: {e.error_code if hasattr(e, 'error_code') else 'N/A'}")
                    print(f"    - Error Message: {e.message if hasattr(e, 'message') else str(e)}")
                    print("!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!")
                    conversion_success = False
                    raise  # 기존 예외 처리로 전달
                    
                except Exception as e:
                    # 그 외 모든 예상치 못한 에러
                    print("!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!")
                    print(f"❌ 변환 중 알 수 없는 예외 발생: {str(e)}")
                    import traceback
                    traceback.print_exc()
                    print("!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!")
                    conversion_success = False
                    raise  # 기존 예외 처리로 전달
                    
                print(">>> [DEBUG 4] Adobe 공식 변환 함수 종료")
                
                # 결과 다운로드
                result_asset = pdf_services_response.get_result().get_download_uri()
                
                # 결과를 파일로 저장
                import requests
                response = requests.get(result_asset)
                with open(output_path, 'wb') as output_stream:
                    output_stream.write(response.content)
            
            logging.info(f"Adobe API를 사용하여 공문서(신뢰도: {confidence:.2f})를 DOCX로 성공적으로 변환했습니다: {output_path}")
            return output_path
            
        except ServiceApiException as e:
            logging.error(f"Adobe API 공문서 변환 ServiceApiException 발생:")
            logging.error(f"  - 오류 메시지: {str(e)}")
            logging.error(f"  - 오류 타입: {type(e).__name__}")
            if hasattr(e, 'status_code'):
                logging.error(f"  - HTTP 상태 코드: {e.status_code}")
            if hasattr(e, 'error_code'):
                logging.error(f"  - Adobe 오류 코드: {e.error_code}")
            if hasattr(e, 'message'):
                logging.error(f"  - 상세 메시지: {e.message}")
            
            # HTTP 400 오류에 대한 특별 처리
            if hasattr(e, 'status_code') and e.status_code == 400:
                logging.error("  - HTTP 400 오류: 요청 매개변수나 파일 형식을 확인하세요")
            
            return None
        except Exception as e:
            logging.error(f"Adobe API 공문서 변환 오류: {e}")
            return None