# c:\Users\박\Desktop\앱개# c:\Users\박\Desktop\앱개발\문서변환기모음\PDF_DOC\final_server.py
# 완전 동일 PDF 생성 함수

def create_identical_pdf_advanced(document_structure, output_path):
    """원본과 완전 동일한 PDF 생성"""
    try:
        # Add imports and main logic here
    except Exception as e:
        print(f"Error creating PDF: {e}")
        raise
        print(f"Error creating PDF: {e}")
        raise
        print(f"Error creating PDF: {e}")
        raise
        print(f"Error creating PDF: {e}")
        raise
        from reportlab.pdfgen import canvas
        from reportlab.lib.pagesizes import A4, landscape, portrait
        from reportlab.lib.utils import ImageReader
        from reportlab.lib.colors import Color, black, white
        from reportlab.lib.units import inch, cm, mm
        
        print("📄 완전 동일 PDF 생성 시작...")
        
        # 섹션 정보 가져오기
        section_info = document_structure['sections'][0] if document_structure['sections'] else {
            'orientation': 'portrait',
            'page_width': 8.5,
            'page_height': 11.0,
            'top_margin': 1.0,
            'bottom_margin': 1.0,
            'left_margin': 1.0,
            'right_margin': 1.0
        }
        
        # 정확한 페이지 크기 설정
        page_width = section_info['page_width'] * inch
        page_height = section_info['page_height'] * inch
        
        if section_info['orientation'] == 'landscape':
            page_size = (page_height, page_width)  # 가로세로 바꿈
            print(f"📐 가로 방향: {page_height/inch:.1f}×{page_width/inch:.1f} 인치")
        else:
            page_size = (page_width, page_height)
            print(f"📐 세로 방향: {page_width/inch:.1f}×{page_height/inch:.1f} 인치")
        
        # PDF 생성
        c = canvas.Canvas(output_path, pagesize=page_size)
        width, height = page_size
        
        # 정확한 여백 설정 (원본 기준)
        margin_left = section_info['left_margin'] * inch
        margin_right = width - (section_info['right_margin'] * inch)
        margin_top = height - (section_info['top_margin'] * inch)
        margin_bottom = section_info['bottom_margin'] * inch
        
        print(f"📏 여백 설정: L{section_info['left_margin']:.1f} R{section_info['right_margin']:.1f} T{section_info['top_margin']:.1f} B{section_info['bottom_margin']:.1f}")
        
        # 현재 위치 추적
        current_y = margin_top
        current_page = 1
        
        # 문서 메타데이터 설정
        metadata = document_structure['metadata']
        if metadata.get('title'):
            c.setTitle(metadata['title'])
        if metadata.get('author'):
            c.setAuthor(metadata['author'])
        if metadata.get('subject'):
            c.setSubject(metadata['subject'])
        
        # 이미지 인덱스 추적
        image_index = 0
        images = document_structure['images']
        
        # 문단별 정확한 재현
        for para_data in document_structure['paragraphs']:
            # 페이지 넘김 확인
            estimated_height = 50  # 기본 문단 높이 추정
            if current_y - estimated_height < margin_bottom:
                c.showPage()
                current_y = margin_top
                current_page += 1
                print(f"📄 페이지 {current_page} 시작")
            
            # 문단 전 간격 적용
            if para_data['space_before'] > 0:
                space_before = (para_data['space_before'] / 72) * inch  # pt를 inch로 변환
                current_y -= space_before
            
            # 빈 문단 처리 (레이아웃 보존)
            if para_data['is_empty']:
                line_height = 12  # 기본 줄 높이
                current_y -= line_height
                continue
            
            # 스타일 기반 폰트 설정
            style_name = para_data['style_name']
            base_font_size = 11
            font_name = 'Helvetica'
            
            # 스타일별 폰트 크기 및 이름 설정
            if 'Heading 1' in style_name:
                base_font_size = 18
                font_name = 'Helvetica-Bold'
            elif 'Heading 2' in style_name:
                base_font_size = 16
                font_name = 'Helvetica-Bold'
            elif 'Heading 3' in style_name:
                base_font_size = 14
                font_name = 'Helvetica-Bold'
            elif 'Title' in style_name:
                base_font_size = 20
                font_name = 'Helvetica-Bold'
            
            # 정렬 설정
            alignment = para_data['alignment']
            
            # 들여쓰기 적용
            para_left_margin = margin_left
            if para_data['left_indent'] > 0:
                para_left_margin += (para_data['left_indent'] / 72) * inch
            if para_data['first_line_indent'] > 0:
                para_left_margin += (para_data['first_line_indent'] / 72) * inch
            
            # Run별 텍스트 처리 (완전한 서식 보존)
            line_content = []
            current_line_width = 0
            max_line_width = margin_right - para_left_margin
            
            for run_data in para_data['runs']:
                if not run_data['text']:
                    continue
                
                # Run별 폰트 설정
                run_font_size = run_data['font_size'] if run_data['font_size'] > 0 else base_font_size
                run_font_name = run_data['font_name'] if run_data['font_name'] != 'Calibri' else font_name
                
                # 볼드, 이탤릭 적용
                if run_data['bold'] and 'Bold' not in run_font_name:
                    run_font_name = run_font_name.replace('Helvetica', 'Helvetica-Bold')
                if run_data['italic'] and 'Oblique' not in run_font_name:
                    run_font_name = run_font_name.replace('Helvetica', 'Helvetica-Oblique')
                    if 'Bold' in run_font_name:
                        run_font_name = 'Helvetica-BoldOblique'
                
                # 텍스트 색상 설정
                text_color = black
                if run_data['font_color'] != '000000':
                    try:
                        color_hex = run_data['font_color']
                        if len(color_hex) == 6:
                            r = int(color_hex[0:2], 16) / 255.0
                            g = int(color_hex[2:4], 16) / 255.0
                            b = int(color_hex[4:6], 16) / 255.0
                            text_color = Color(r, g, b)
                    except:
                        text_color = black
                
                # 텍스트 출력
                text = run_data['text']
                
                # 한글 텍스트 처리
                if any('\uac00' <= char <= '\ud7af' for char in text):
                    # 한글 폰트 사용 시도
                    if KOREAN_FONT_AVAILABLE and KOREAN_FONT != 'Helvetica':
                        try:
                            c.setFont(KOREAN_FONT, run_font_size)
                            c.setFillColor(text_color)
                            
                            # 정렬에 따른 X 위치 계산
                            if alignment == 'CENTER':
                                text_width = c.stringWidth(text, KOREAN_FONT, run_font_size)
                                x_pos = para_left_margin + (max_line_width - text_width) / 2
                            elif alignment == 'RIGHT':
                                text_width = c.stringWidth(text, KOREAN_FONT, run_font_size)
                                x_pos = margin_right - text_width
                            else:
                                x_pos = para_left_margin
                            
                            c.drawString(x_pos, current_y, text)
                            print(f"✅ 한글 텍스트 출력: {text[:20]}...")
                            
                        except Exception as e:
                            print(f"한글 폰트 오류: {e}")
                            # 대체 처리
                            c.setFont('Helvetica', run_font_size)
                            c.setFillColor(text_color)
                            korean_count = len([c for c in text if '\uac00' <= c <= '\ud7af'])
                            english_part = ''.join([c for c in text if ord(c) < 128])
                            
                            if english_part:
                                display_text = f"{english_part} [한글{korean_count}자]"
                            else:
                                display_text = f"[한글 텍스트 {korean_count}자]"
                            
                            c.drawString(para_left_margin, current_y, display_text)
                    else:
                        # 한글 폰트 없을 때 대체 처리
                        c.setFont('Helvetica', run_font_size)
                        c.setFillColor(text_color)
                        korean_count = len([c for c in text if '\uac00' <= c <= '\ud7af'])
                        english_part = ''.join([c for c in text if ord(c) < 128])
                        
                        if english_part:
                            display_text = f"{english_part}"  # Display English part of text

def create_identical_pdf_advanced(document_structure, output_path):
    """원본과 완전 동일한 PDF 생성"""
    try:
        # Main logic will be here
    except Exception as e:
        print(f"Error creating PDF: {e}")
        raise
        print(f"Error creating PDF: {e}")
        raise
        from reportlab.pdfgen import canvas
        from reportlab.lib.pagesizes import A4, landscape, portrait
        from reportlab.lib.utils import ImageReader
        from reportlab.lib.colors import Color, black, white
        from reportlab.lib.units import inch, cm, mm
        
        print("📄 완전 동일 PDF 생성 시작...")
        
        # 섹션 정보 가져오기
        section_info = document_structure['sections'][0] if document_structure['sections'] else {
            'orientation': 'portrait',
            'page_width': 8.5,
            'page_height': 11.0,
            'top_margin': 1.0,
            'bottom_margin': 1.0,
            'left_margin': 1.0,
            'right_margin': 1.0
        }
        
        # 정확한 페이지 크기 설정
        page_width = section_info['page_width'] * inch
        page_height = section_info['page_height'] * inch
        
        if section_info['orientation'] == 'landscape':
            page_size = (page_height, page_width)  # 가로세로 바꿈
            print(f"📐 가로 방향: {page_height/inch:.1f}×{page_width/inch:.1f} 인치")
        else:
            page_size = (page_width, page_height)
            print(f"📐 세로 방향: {page_width/inch:.1f}×{page_height/inch:.1f} 인치")
        
        # PDF 생성
        c = canvas.Canvas(output_path, pagesize=page_size)
        width, height = page_size
        
        # 정확한 여백 설정 (원본 기준)
        margin_left = section_info['left_margin'] * inch
        margin_right = width - (section_info['right_margin'] * inch)
        margin_top = height - (section_info['top_margin'] * inch)
        margin_bottom = section_info['bottom_margin'] * inch
        
        print(f"📏 여백 설정: L{section_info['left_margin']:.1f} R{section_info['right_margin']:.1f} T{section_info['top_margin']:.1f} B{section_info['bottom_margin']:.1f}")
        
        # 현재 위치 추적
        current_y = margin_top
        current_page = 1
        
        # 문서 메타데이터 설정
        metadata = document_structure['metadata']
        if metadata.get('title'):
            c.setTitle(metadata['title'])
        if metadata.get('author'):
            c.setAuthor(metadata['author'])
        if metadata.get('subject'):
            c.setSubject(metadata['subject'])
        
        # 이미지 인덱스 추적
        image_index = 0
        images = document_structure['images']
        
        # 문단별 정확한 재현
        for para_data in document_structure['paragraphs']:
            # 페이지 넘김 확인
            estimated_height = 50  # 기본 문단 높이 추정
            if current_y - estimated_height < margin_bottom:
                c.showPage()
                current_y = margin_top
                current_page += 1
                print(f"📄 페이지 {current_page} 시작")
            
            # 문단 전 간격 적용
            if para_data['space_before'] > 0:
                space_before = (para_data['space_before'] / 72) * inch  # pt를 inch로 변환
                current_y -= space_before
            
            # 빈 문단 처리 (레이아웃 보존)
            if para_data['is_empty']:
                line_height = 12  # 기본 줄 높이
                current_y -= line_height
                continue
            
            # 스타일 기반 폰트 설정
            style_name = para_data['style_name']
            base_font_size = 11
            font_name = 'Helvetica'
            
            # 스타일별 폰트 크기 및 이름 설정
            if 'Heading 1' in style_name:
                base_font_size = 18
                font_name = 'Helvetica-Bold'
            elif 'Heading 2' in style_name:
                base_font_size = 16
                font_name = 'Helvetica-Bold'
            elif 'Heading 3' in style_name:
                base_font_size = 14
                font_name = 'Helvetica-Bold'
            elif 'Title' in style_name:
                base_font_size = 20
                font_name = 'Helvetica-Bold'
            
            # 정렬 설정
            alignment = para_data['alignment']
            
            # 들여쓰기 적용
            para_left_margin = margin_left
            if para_data['left_indent'] > 0:
                para_left_margin += (para_data['left_indent'] / 72) * inch
            if para_data['first_line_indent'] > 0:
                para_left_margin += (para_data['first_line_indent'] / 72) * inch
            
            # Run별 텍스트 처리 (완전한 서식 보존)
            line_content = []
            current_line_width = 0
            max_line_width = margin_right - para_left_margin
            
            for run_data in para_data['runs']:
                if not run_data['text']:
                    continue
                
                # Run별 폰트 설정
                run_font_size = run_data['font_size'] if run_data['font_size'] > 0 else base_font_size
                run_font_name = run_data['font_name'] if run_data['font_name'] != 'Calibri' else font_name
                
                # 볼드, 이탤릭 적용
                if run_data['bold'] and 'Bold' not in run_font_name:
                    run_font_name = run_font_name.replace('Helvetica', 'Helvetica-Bold')
                if run_data['italic'] and 'Oblique' not in run_font_name:
                    run_font_name = run_font_name.replace('Helvetica', 'Helvetica-Oblique')
                    if 'Bold' in run_font_name:
                        run_font_name = 'Helvetica-BoldOblique'
                
                # 텍스트 색상 설정
                text_color = black
                if run_data['font_color'] != '000000':
                    try:
                        color_hex = run_data['font_color']
                        if len(color_hex) == 6:
                            r = int(color_hex[0:2], 16) / 255.0
                            g = int(color_hex[2:4], 16) / 255.0
                            b = int(color_hex[4:6], 16) / 255.0
                            text_color = Color(r, g, b)
                    except:
                        text_color = black
                
                # 텍스트 출력
                text = run_data['text']
                
                # 한글 텍스트 처리
                if any('\uac00' <= char <= '\ud7af' for char in text):
                    # 한글 폰트 사용 시도
                    if KOREAN_FONT_AVAILABLE and KOREAN_FONT != 'Helvetica':
                        try:
                            c.setFont(KOREAN_FONT, run_font_size)
                            c.setFillColor(text_color)
                            
                            # 정렬에 따른 X 위치 계산
                            if alignment == 'CENTER':
                                text_width = c.stringWidth(text, KOREAN_FONT, run_font_size)
                                x_pos = para_left_margin + (max_line_width - text_width) / 2
                            elif alignment == 'RIGHT':
                                text_width = c.stringWidth(text, KOREAN_FONT, run_font_size)
                                x_pos = margin_right - text_width
                            else:
                                x_pos = para_left_margin
                            
                            c.drawString(x_pos, current_y, text)
                            print(f"✅ 한글 텍스트 출력: {text[:20]}...")
                            
                        except Exception as e:
                            print(f"한글 폰트 오류: {e}")
                            # 대체 처리
                            c.setFont('Helvetica', run_font_size)
                            c.setFillColor(text_color)
                            korean_count = len([c for c in text if '\uac00' <= c <= '\ud7af'])
                            english_part = ''.join([c for c in text if ord(c) < 128])
                            
                            if english_part:
                                display_text = f"{english_part} [한글{korean_count}자]"
                            else:
                                display_text = f"[한글 텍스트 {korean_count}자]"
                            
                            c.drawString(para_left_margin, current_y, display_text)
                    else:
                        # 한글 폰트 없을 때 대체 처리
                        c.setFont('Helvetica', run_font_size)
                        c.setFillColor(text_color)
                        korean_count = len([c for c in text if '\uac00' <= c <= '\ud7af'])
                        english_part = ''.join([c for c in text if ord(c) < 128])
                        
                        if english_part:
                            display_text = f"{english_part}"

def create_identical_pdf_advanced(document_structure, output_path):
    """원본과 완전 동일한 PDF 생성"""
    try:
        from reportlab.pdfgen import canvas
        from reportlab.lib.pagesizes import A4, landscape, portrait
        from reportlab.lib.utils import ImageReader
        from reportlab.lib.colors import Color, black, white
        from reportlab.lib.units import inch, cm, mm
        
        print("📄 완전 동일 PDF 생성 시작...")
        
        # 섹션 정보 가져오기
        section_info = document_structure['sections'][0] if document_structure['sections'] else {
            'orientation': 'portrait',
            'page_width': 8.5,
            'page_height': 11.0,
            'top_margin': 1.0,
            'bottom_margin': 1.0,
            'left_margin': 1.0,
            'right_margin': 1.0
        }
        
        # 정확한 페이지 크기 설정
        page_width = section_info['page_width'] * inch
        page_height = section_info['page_height'] * inch
        
        if section_info['orientation'] == 'landscape':
            page_size = (page_height, page_width)  # 가로세로 바꿈
            print(f"📐 가로 방향: {page_height/inch:.1f}×{page_width/inch:.1f} 인치")
        else:
            page_size = (page_width, page_height)
            print(f"📐 세로 방향: {page_width/inch:.1f}×{page_height/inch:.1f} 인치")
        
        # PDF 생성
        c = canvas.Canvas(output_path, pagesize=page_size)
        width, height = page_size
        
        # 정확한 여백 설정 (원본 기준)
        margin_left = section_info['left_margin'] * inch
        margin_right = width - (section_info['right_margin'] * inch)
        margin_top = height - (section_info['top_margin'] * inch)
        margin_bottom = section_info['bottom_margin'] * inch
        
        print(f"📏 여백 설정: L{section_info['left_margin']:.1f} R{section_info['right_margin']:.1f} T{section_info['top_margin']:.1f} B{section_info['bottom_margin']:.1f}")
        
        # 현재 위치 추적
        current_y = margin_top
        current_page = 1
        
        # 문서 메타데이터 설정
        metadata = document_structure['metadata']
        if metadata.get('title'):
            c.setTitle(metadata['title'])
        if metadata.get('author'):
            c.setAuthor(metadata['author'])
        if metadata.get('subject'):
            c.setSubject(metadata['subject'])
        
        # 이미지 인덱스 추적
        image_index = 0
        images = document_structure['images']
        
        # 문단별 정확한 재현
        for para_data in document_structure['paragraphs']:
            # 페이지 넘김 확인
            estimated_height = 50  # 기본 문단 높이 추정
            if current_y - estimated_height < margin_bottom:
                c.showPage()
                current_y = margin_top
                current_page += 1
                print(f"📄 페이지 {current_page} 시작")
            
            # 문단 전 간격 적용
            if para_data['space_before'] > 0:
                space_before = (para_data['space_before'] / 72) * inch  # pt를 inch로 변환
                current_y -= space_before
            
            # 빈 문단 처리 (레이아웃 보존)
            if para_data['is_empty']:
                line_height = 12  # 기본 줄 높이
                current_y -= line_height
                continue
            
            # 스타일 기반 폰트 설정
            style_name = para_data['style_name']
            base_font_size = 11
            font_name = 'Helvetica'
            
            # 스타일별 폰트 크기 및 이름 설정
            if 'Heading 1' in style_name:
                base_font_size = 18
                font_name = 'Helvetica-Bold'
            elif 'Heading 2' in style_name:
                base_font_size = 16
                font_name = 'Helvetica-Bold'
            elif 'Heading 3' in style_name:
                base_font_size = 14
                font_name = 'Helvetica-Bold'
            elif 'Title' in style_name:
                base_font_size = 20
                font_name = 'Helvetica-Bold'
            
            # 정렬 설정
            alignment = para_data['alignment']
            
            # 들여쓰기 적용
            para_left_margin = margin_left
            if para_data['left_indent'] > 0:
                para_left_margin += (para_data['left_indent'] / 72) * inch
            if para_data['first_line_indent'] > 0:
                para_left_margin += (para_data['first_line_indent'] / 72) * inch
            
            # Run별 텍스트 처리 (완전한 서식 보존)
            line_content = []
            current_line_width = 0
            max_line_width = margin_right - para_left_margin
            
            for run_data in para_data['runs']:
                if not run_data['text']:
                    continue
                
                # Run별 폰트 설정
                run_font_size = run_data['font_size'] if run_data['font_size'] > 0 else base_font_size
                run_font_name = run_data['font_name'] if run_data['font_name'] != 'Calibri' else font_name
                
                # 볼드, 이탤릭 적용
                if run_data['bold'] and 'Bold' not in run_font_name:
                    run_font_name = run_font_name.replace('Helvetica', 'Helvetica-Bold')
                if run_data['italic'] and 'Oblique' not in run_font_name:
                    run_font_name = run_font_name.replace('Helvetica', 'Helvetica-Oblique')
                    if 'Bold' in run_font_name:
                        run_font_name = 'Helvetica-BoldOblique'
                
                # 텍스트 색상 설정
                text_color = black
                if run_data['font_color'] != '000000':
                    try:
                        color_hex = run_data['font_color']
                        if len(color_hex) == 6:
                            r = int(color_hex[0:2], 16) / 255.0
                            g = int(color_hex[2:4], 16) / 255.0
                            b = int(color_hex[4:6], 16) / 255.0
                            text_color = Color(r, g, b)
                    except:
                        text_color = black
                
                # 텍스트 출력
                text = run_data['text']
                
                # 한글 텍스트 처리
                if any('\uac00' <= char <= '\ud7af' for char in text):
                    # 한글 폰트 사용 시도
                    if KOREAN_FONT_AVAILABLE and KOREAN_FONT != 'Helvetica':
                        try:
                            c.setFont(KOREAN_FONT, run_font_size)
                            c.setFillColor(text_color)
                            
                            # 정렬에 따른 X 위치 계산
                            if alignment == 'CENTER':
                                text_width = c.stringWidth(text, KOREAN_FONT, run_font_size)
                                x_pos = para_left_margin + (max_line_width - text_width) / 2
                            elif alignment == 'RIGHT':
                                text_width = c.stringWidth(text, KOREAN_FONT, run_font_size)
                                x_pos = margin_right - text_width
                            else:
                                x_pos = para_left_margin
                            
                            c.drawString(x_pos, current_y, text)
                            print(f"✅ 한글 텍스트 출력: {text[:20]}...")
                            
                        except Exception as e:
                            print(f"한글 폰트 오류: {e}")
                            # 대체 처리
                            c.setFont('Helvetica', run_font_size)
                            c.setFillColor(text_color)
                            korean_count = len([c for c in text if '\uac00' <= c <= '\ud7af'])
                            english_part = ''.join([c for c in text if ord(c) < 128])
                            
                            if english_part:
                                display_text = f"{english_part} [한글{korean_count}자]"
                            else:
                                display_text = f"[한글 텍스트 {korean_count}자]"
                            
                            c.drawString(para_left_margin, current_y, display_text)
                    else:
                        # 한글 폰트 없을 때 대체 처리
                        c.setFont('Helvetica', run_font_size)
                        c.setFillColor(text_color)
                        korean_count = len([c for c in text if '\uac00' <= c <= '\ud7af'])
                        english_part = ''.join([c for c in text if ord(c) < 128])
                        
                        if english_part:
                            display_text = f"{english_part}