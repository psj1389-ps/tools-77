from adobe.pdfservices.operation.auth.service_principal_credentials import ServicePrincipalCredentials
from adobe.pdfservices.operation.pdf_services import PDFServices
from adobe.pdfservices.operation.pdf_services_media_type import PDFServicesMediaType
from adobe.pdfservices.operation.pdfjobs.jobs.export_pdf_job import ExportPDFJob
from adobe.pdfservices.operation.pdfjobs.params.export_pdf.export_pdf_params import ExportPDFParams
from adobe.pdfservices.operation.pdfjobs.params.export_pdf.export_pdf_target_format import ExportPDFTargetFormat
from adobe.pdfservices.operation.pdfjobs.result.export_pdf_result import ExportPDFResult
import os
import traceback
import pdfplumber
import fitz
import pypdfium2 as pdfium
from PyPDF2 import PdfReader, PdfWriter


def adobe_pdf_to_docx(input_path: str, output_path: str) -> tuple[bool, dict]:
    """Adobe API를 사용하여 PDF를 DOCX로 변환
    
    Args:
        input_path: 입력 PDF 파일 경로
        output_path: 출력 DOCX 파일 경로
        
    Returns:
        tuple[bool, dict]: (성공 여부, 에러 정보)
    """
    print(">>> [DEBUG] adobe_pdf_to_docx: start", flush=True)
    info = {}
    
    try:
        # Adobe API 자격 증명 설정
        creds = ServicePrincipalCredentials(
            client_id=os.environ["ADOBE_CLIENT_ID"],
            client_secret=os.environ["ADOBE_CLIENT_SECRET"]
        )
        pdf_services = PDFServices(credentials=creds)

        # PDF 파일 읽기
        with open(input_path, "rb") as f:
            input_bytes = f.read()
        
        print(">>> [DEBUG] upload", flush=True)
        asset = pdf_services.upload(input_bytes, PDFServicesMediaType.PDF)

        # Export 작업 설정
        params = ExportPDFParams(ExportPDFTargetFormat.DOCX)
        job = ExportPDFJob(asset, params)
        
        print(">>> [DEBUG] submit", flush=True)
        location = pdf_services.submit(job)

        print(">>> [DEBUG] get_job_result", flush=True)
        result_asset = pdf_services.get_job_result(location, ExportPDFResult)

        print(">>> [DEBUG] get_content", flush=True)
        content = pdf_services.get_content(result_asset)

        # 결과 파일 저장
        with open(output_path, "wb") as f:
            f.write(content)

        print(">>> [DEBUG] success", flush=True)
        return True, info

    except Exception as e:
        # 모든 예외를 포괄해서 상세 출력
        info = {
            "type": type(e).__name__,
            "status": getattr(e, "status_code", None),
            "error_code": getattr(e, "error_code", None),
            "message": getattr(e, "message", str(e)),
            "request_id": getattr(e, "request_id", None),
            "error_report": getattr(e, "error_report", None),
        }
        print("!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!", flush=True)
        print("❌ Adobe call failed", flush=True)
        for k, v in info.items():
            print(f"{k}: {v}", flush=True)
        traceback.print_exc()
        print("!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!", flush=True)
        return False, info

def is_image_only_pdf(path: str, min_chars=10) -> bool:
    """PDF가 이미지만 포함하는지 확인
    
    Args:
        path: PDF 파일 경로
        min_chars: 텍스트로 간주할 최소 문자 수
        
    Returns:
        bool: 이미지만 포함하면 True, 텍스트가 있으면 False
    """
    try:
        with pdfplumber.open(path) as pdf:
            total_chars = 0
            for page in pdf.pages:
                text = (page.extract_text() or "").strip()
                total_chars += len(text)
                if total_chars >= min_chars:
                    return False
        return True
    except Exception:
        return False  # 문제가 있으면 보수적으로 False


def is_encrypted_pdf(path: str) -> bool:
    """PDF가 암호화되어 있는지 확인
    
    Args:
        path: PDF 파일 경로
        
    Returns:
        bool: 암호화되어 있으면 True
    """
    try:
        with fitz.open(path) as doc:
            return doc.isEncrypted
    except Exception:
        return False


def normalize_pdf(src: str, dst: str) -> bool:
    """PDF를 정규화하여 다시 저장
    
    Args:
        src: 원본 PDF 파일 경로
        dst: 정규화된 PDF 파일 경로
        
    Returns:
        bool: 성공 여부
    """
    try:
        # PyPDF2를 사용한 PDF 정규화
        reader = PdfReader(src)
        writer = PdfWriter()
        
        for page in reader.pages:
            writer.add_page(page)
            
        with open(dst, "wb") as f:
            writer.write(f)
            
        return True
    except Exception as e:
        print(f"PDF 정규화 실패: {e}", flush=True)
        return False