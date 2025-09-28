import os
import json
from typing import Optional, Dict, Any
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# Google Drive API 스코프 - 읽기 전용 접근
SCOPES = ['https://www.googleapis.com/auth/drive.readonly']

class GoogleDriveAuth:
    """Google Drive API OAuth 2.0 인증 관리 클래스"""
    
    def __init__(self, client_secret_path: str = 'client_secret.json', token_path: str = 'token.json'):
        self.client_secret_path = client_secret_path
        self.token_path = token_path
        self.credentials: Optional[Credentials] = None
        self.service = None
        
    def get_credentials(self) -> Optional[Credentials]:
        """OAuth 2.0 자격 증명을 가져오거나 생성합니다."""
        creds = None
        
        # 기존 토큰 파일이 있으면 로드
        if os.path.exists(self.token_path):
            try:
                creds = Credentials.from_authorized_user_file(self.token_path, SCOPES)
                print(f"기존 토큰 파일에서 자격 증명을 로드했습니다: {self.token_path}")
            except Exception as e:
                print(f"토큰 파일 로드 중 오류 발생: {e}")
                creds = None
        
        # 유효한 자격 증명이 없거나 만료된 경우
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                try:
                    print("만료된 토큰을 새로고침합니다...")
                    creds.refresh(Request())
                    print("토큰 새로고침 완료")
                except Exception as e:
                    print(f"토큰 새로고침 실패: {e}")
                    creds = None
            
            # 새로운 인증이 필요한 경우
            if not creds or not creds.valid:
                if not os.path.exists(self.client_secret_path):
                    raise FileNotFoundError(
                        f"Google API client_secret.json 파일을 찾을 수 없습니다: {self.client_secret_path}\n"
                        "Google Cloud Console에서 OAuth 2.0 클라이언트 ID를 생성하고 client_secret.json 파일을 다운로드하세요."
                    )
                
                try:
                    print("새로운 OAuth 2.0 인증을 시작합니다...")
                    flow = InstalledAppFlow.from_client_secrets_file(
                        self.client_secret_path, SCOPES
                    )
                    # 로컬 서버를 사용하여 인증 플로우 실행
                    creds = flow.run_local_server(port=0)
                    print("OAuth 2.0 인증 완료")
                except Exception as e:
                    print(f"OAuth 2.0 인증 실패: {e}")
                    raise
            
            # 자격 증명을 파일에 저장
            if creds and creds.valid:
                try:
                    with open(self.token_path, 'w', encoding='utf-8') as token_file:
                        token_file.write(creds.to_json())
                    print(f"자격 증명을 저장했습니다: {self.token_path}")
                except Exception as e:
                    print(f"토큰 저장 중 오류 발생: {e}")
        
        self.credentials = creds
        return creds
    
    def get_service(self):
        """Google Drive API 서비스 객체를 반환합니다."""
        if not self.service:
            creds = self.get_credentials()
            if not creds:
                raise Exception("Google Drive API 자격 증명을 가져올 수 없습니다.")
            
            try:
                self.service = build('drive', 'v3', credentials=creds)
                print("Google Drive API 서비스가 성공적으로 초기화되었습니다.")
            except Exception as e:
                print(f"Google Drive API 서비스 초기화 실패: {e}")
                raise
        
        return self.service
    
    def test_connection(self) -> Dict[str, Any]:
        """Google Drive API 연결을 테스트합니다."""
        try:
            service = self.get_service()
            
            # 파일 목록 가져오기 (최대 10개)
            results = service.files().list(
                pageSize=10, 
                fields="nextPageToken, files(id, name, mimeType, size, modifiedTime)"
            ).execute()
            
            items = results.get('files', [])
            
            return {
                'success': True,
                'message': 'Google Drive API 연결 성공',
                'file_count': len(items),
                'files': items[:5]  # 처음 5개 파일만 반환
            }
            
        except HttpError as error:
            return {