from advanced_text_filter import filter_text_blocks

def test_preview_content():
    """프리뷰에서 보이는 내용 필터링 테스트"""
    
    # 사용자가 보여준 프리뷰 내용
    preview_text = """
    변환 방식: 표준 변환 (빠름)
    ### HTML 템플릿 업데이트: "html
    변환 방식: 표준 변환 (빠름)
    ### HTML 템플릿 업데이트: "html
    변환 방식: 표준 변환 (빠름)
    ## 🎯 4. 웹 인터페이스 개선: "html
    
    실제 문서 내용입니다.
    이것은 보존되어야 할 중요한 정보입니다.
    회사 공지사항
    담당자: 홍길동
    연락처: 02-1234-5678
    """
    
    print("=== 원본 텍스트 ===")
    print(preview_text)
    
    print("\n=== 필터링 결과 (디버그 모드) ===")
    filtered = filter_text_blocks(preview_text, debug=True)
    
    print("\n=== 최종 결과 ===")
    print(filtered)
    
    print("\n=== 제거된 내용 분석 ===")
    original_lines = [l.strip() for l in preview_text.splitlines() if l.strip()]
    filtered_lines = [l.strip() for l in filtered.splitlines() if l.strip()]
    
    removed_lines = []
    for line in original_lines:
        if line not in filtered_lines:
            removed_lines.append(line)
    
    if removed_lines:
        print("제거된 라인들:")
        for line in removed_lines:
            print(f"  ❌ {line}")
    else:
        print("제거된 라인이 없습니다.")
    
    if filtered_lines:
        print("\n보존된 라인들:")
        for line in filtered_lines:
            print(f"  ✅ {line}")

if __name__ == "__main__":
    test_preview_content()