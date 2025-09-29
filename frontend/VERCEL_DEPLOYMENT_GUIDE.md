# Vercel 배포 가이드

## 📋 배포 준비 완료 사항

✅ **vercel.json 설정 최적화**
- Vite 프레임워크 지정
- 보안 헤더 추가
- API 프록시 설정 (pdf-doc 서비스)
- SPA 라우팅 지원

✅ **빌드 설정 최적화**
- 청크 분할 설정 (vendor, router, icons)
- 소스맵 숨김 처리
- 빌드 크기 경고 임계값 설정

✅ **TypeScript 설정 완료**
- tsconfig.json, tsconfig.app.json, tsconfig.node.json 생성
- 경로 매핑 설정 (@/* → ./src/*)

✅ **환경변수 템플릿 생성**
- .env.example 파일 생성

## 🚀 Vercel 배포 단계

### 1. Vercel 계정 준비
1. [Vercel](https://vercel.com)에 가입/로그인
2. GitHub 계정 연결

### 2. 프로젝트 배포

#### 방법 1: Vercel CLI 사용
```bash
# Vercel CLI 설치
npm i -g vercel

# 프론트엔드 디렉토리에서 배포
cd frontend
vercel

# 프로덕션 배포
vercel --prod
```

#### 방법 2: GitHub 연동 (권장)
1. GitHub에 프로젝트 푸시
2. Vercel 대시보드에서 "New Project" 클릭
3. GitHub 저장소 선택
4. 프로젝트 설정:
   - **Framework Preset**: Vite
   - **Root Directory**: `frontend`
   - **Build Command**: `npm ci && npm run build`
   - **Output Directory**: `dist`
   - **Install Command**: `npm ci`

### 3. 환경변수 설정 (선택사항)
Vercel 대시보드 → Settings → Environment Variables에서 설정:

```
VITE_API_BASE_URL=https://tools-77.vercel.app
VITE_PDF_SERVICE_URL=https://pdf-doc-306w.onrender.com
VITE_DEV_MODE=false
VITE_APP_VERSION=1.0.0
```

### 4. 도메인 설정
1. Vercel 대시보드 → Settings → Domains
2. 커스텀 도메인 추가 (선택사항)
3. 자동 HTTPS 인증서 적용

## 🔧 배포 후 확인사항

### 1. 기본 기능 테스트
- [ ] 홈페이지 로딩
- [ ] 라우팅 동작 (페이지 이동)
- [ ] 반응형 디자인
- [ ] 콘솔 에러 확인

### 2. API 연결 테스트
- [ ] PDF 변환 서비스 연결 확인
- [ ] 네트워크 탭에서 API 호출 상태 확인

### 3. 성능 확인
- [ ] Lighthouse 점수 확인
- [ ] 로딩 속도 테스트
- [ ] 번들 크기 확인

## 🛠️ 문제 해결

### 빌드 실패 시
```bash
# 로컬에서 빌드 테스트
npm run build

# TypeScript 에러 확인
npm run check
```

### 라우팅 문제 시
- vercel.json의 rewrites 설정 확인
- SPA 라우팅이 올바르게 설정되었는지 확인

### API 연결 문제 시
- CORS 설정 확인
- 백엔드 서버 상태 확인
- 환경변수 설정 확인

## 📊 배포 정보

**현재 설정:**
- 프레임워크: Vite + React + TypeScript
- 빌드 도구: Vite
- 라우터: React Router DOM
- 상태 관리: Zustand
- 스타일링: Tailwind CSS
- 아이콘: Lucide React

**최적화 적용:**
- 코드 분할 (vendor, router, icons)
- 트리 쉐이킹
- 압축 및 최소화
- 보안 헤더 적용

## 🔄 지속적 배포

GitHub 연동 시 자동 배포:
- `main` 브랜치 푸시 → 프로덕션 배포
- 다른 브랜치 푸시 → 프리뷰 배포
- Pull Request → 프리뷰 URL 자동 생성

---

**배포 완료 후 URL을 확인하고 모든 기능이 정상 작동하는지 테스트해주세요!** 🎉