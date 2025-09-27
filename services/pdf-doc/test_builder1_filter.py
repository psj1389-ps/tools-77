from advanced_text_filter import filter_text_blocks

def test_builder1_content():
    """Builder1 문서 내용 필터링 테스트"""
    
    # Builder1 문서에서 발견된 실제 노이즈 패턴
    builder1_sample = """
    변환 방식: 표준 변환 (빠름)
    ### HTML 템플릿 업데이트: ```html
    변환 방식: 표준 변환 (빠름)
    ### HTML 템플릿 업데이트: ```html
    변환 방식: 표준 변환 (빠름)
    ## 🎯 4. 웹 인터페이스 개선: ```html
    
    # 해결: 웹 인터페이스 환경변수 관련 업데이트 및 개선
    `index.html` 파일 로딩 div 요소 제거
    
    # 해결: 웹 인터페이스 환경변수 관련 업데이트 및 개선
    로딩 애니메이션 제거
    
    실제 문서의 중요한 내용입니다.
    회사 공지사항
    수신: 전 직원
    제목: 중요 안내사항
    담당자: 김철수
    연락처: 02-1234-5678
    
    PDF 파일을 업로드하면 PPTX로 변환하는 파일 변환기를 개발해 달라고 합니다.
    현재 작업 디렉토리 구조를 확인해보겠습니다.
    
    변환 속도: 기존 대비 60-80% 향상
    파일 크기: 기존 대비로 더 작은 다운로드
    메모리 사용량: 현저히 감소
    처리 시간: 대용량 PDF에 대한 빠른 처리
    """
    
    print("=== Builder1 원본 내용 ===")
    print(builder1_sample)
    
    print("\n=== 강화된 필터링 적용 (디버그 모드) ===")
    filtered = filter_text_blocks(builder1_sample, debug=True)
    
    print("\n=== 최종 필터링 결과 ===")
    print(filtered)
    
    # 결과 분석
    original_lines = [l.strip() for l in builder1_sample.splitlines() if l.strip()]
    filtered_lines = [l.strip() for l in filtered.splitlines() if l.strip()]
    
    print(f"\n📊 필터링 통계:")
    print(f"원본 라인 수: {len(original_lines)}")
    print(f"필터링 후: {len(filtered_lines)}")
    print(f"제거율: {((len(original_lines) - len(filtered_lines)) / len(original_lines) * 100):.1f}%")
    
    print("\n❌ 제거된 노이즈:")
    removed_lines = []
    for line in original_lines:
        if line not in filtered_lines:
            removed_lines.append(line)
    
    for line in removed_lines[:10]:  # 처음 10개만 표시
        print(f"  - {line}")
    
    print("\n✅ 보존된 중요 내용:")
    for line in filtered_lines[:10]:  # 처음 10개만 표시
        print(f"  + {line}")

if __name__ == "__main__":
    test_builder1_content()