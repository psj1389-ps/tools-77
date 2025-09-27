from advanced_text_filter import filter_text_blocks, is_nuke_line, block_remove

def test_nuke_mode():
    """NUKE 모드 테스트"""
    print("=== NUKE 모드 테스트 ===")
    
    nuke_test_lines = [
        "변환 방식:",
        "표준 변환 (빠름)",
        "### HTML 템플릿 업데이트:",
        "## 🎯 4. 웹 인터페이스 개선:",
        "```html",
        "실제 내용입니다",
        "공지사항"
    ]
    
    for line in nuke_test_lines:
        result = is_nuke_line(line)
        status = "❌ 제거" if result else "✅ 보존"
        print(f"{status}: {line}")

def test_block_removal():
    """블록 제거 테스트"""
    print("\n=== 블록 제거 테스트 ===")
    
    test_lines = [
        "변환 방식:",
        "표준 변환 (빠름)",
        "고급 변환 (정확)",
        "### HTML 템플릿 업데이트:",
        "```html",
        "<div>내용</div>",
        "```",
        "실제 문서 내용",
        "중요한 정보"
    ]
    
    print("원본 라인들:")
    for i, line in enumerate(test_lines):
        print(f"{i}: {line}")
    
    filtered_lines = block_remove(test_lines)
    
    print("\n필터링 후:")
    for i, line in enumerate(filtered_lines):
        print(f"{i}: {line}")

def test_full_integration():
    """전체 통합 테스트"""
    print("\n=== 전체 통합 테스트 ===")
    
    complex_text = """
    변환 방식: 표준 변환 (빠름)
    ### HTML 템플릿 업데이트: "html
    변환 방식: 표준 변환 (빠름)
    ### HTML 템플릿 업데이트: "html
    변환 방식: 표준 변환 (빠름)
    ## 🎯 4. 웹 인터페이스 개선: "html
    
    ```html
    <div class="upload-area">
        <p>파일을 선택하세요</p>
    </div>
    ```
    
    실제 문서의 중요한 내용입니다.
    회사 공지사항
    수신: 전 직원
    제목: 중요 안내사항
    담당자: 김철수
    연락처: 02-1234-5678
    
    이 내용은 반드시 보존되어야 합니다.
    """
    
    print("원본:")
    print(complex_text)
    
    print("\n강화된 필터링 결과:")
    result = filter_text_blocks(complex_text, debug=True)
    print("\n최종 결과:")
    print(result)
    
    # 결과 분석
    original_lines = [l.strip() for l in complex_text.splitlines() if l.strip()]
    result_lines = [l.strip() for l in result.splitlines() if l.strip()]
    
    print(f"\n📊 통계:")
    print(f"원본 라인 수: {len(original_lines)}")
    print(f"필터링 후: {len(result_lines)}")
    print(f"제거율: {((len(original_lines) - len(result_lines)) / len(original_lines) * 100):.1f}%")

if __name__ == "__main__":
    test_nuke_mode()
    test_block_removal()
    test_full_integration()