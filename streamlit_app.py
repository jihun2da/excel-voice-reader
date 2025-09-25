import streamlit as st
import openpyxl
import tempfile
import os
import uuid
import io
import re
import pandas as pd
import pyttsx3
import threading
import time

# 페이지 설정
st.set_page_config(
    page_title="🎵 엑셀 음성 리더",
    page_icon="🎵",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 사이즈 변환 딕셔너리
SIZE_DICT = {
    "XS": "엑스스몰", "S": "스몰", "M": "미디움", "L": "라지", "FREE": "프리",
    "XL": "엑스라지", "XXL": "투엑스라지",
    "JS": "주니어 스몰", "JM": "주니어 미디움", "JL": "주니어 라지"
}

# 한국어 숫자 변환
KOREAN_NUMBER_MAP = {
    1: "한개", 2: "두개", 3: "세개", 4: "네개",
    5: "다섯개", 6: "여섯개", 7: "일곱개",
    8: "여덟개", 9: "아홉개", 10: "열개"
}

def convert_size(size_code):
    return SIZE_DICT.get(str(size_code).upper(), size_code)

def convert_quantity(n):
    if isinstance(n, int) and n >= 1:
        return KOREAN_NUMBER_MAP.get(n, f"{n}개")
    return ""

def clean_g_value(text):
    return re.sub(r"\(.*?\)", "", text).strip()

# 간단한 TTS 함수
def speak_text(text):
    try:
        engine = pyttsx3.init()
        engine.setProperty('rate', 200)  # 속도
        engine.setProperty('volume', 1.0)  # 볼륨
        engine.say(text)
        engine.runAndWait()
        return True
    except Exception as e:
        st.error(f"음성 재생 실패: {str(e)}")
        return False

# 세션 상태 초기화
if 'file_data' not in st.session_state:
    st.session_state.file_data = None
if 'current_row' not in st.session_state:
    st.session_state.current_row = 2
if 'reading' not in st.session_state:
    st.session_state.reading = False
if 'prev_g_value' not in st.session_state:
    st.session_state.prev_g_value = None

# 메인 헤더
st.title("🎵 엑셀 음성 리더 - 웹버전")
st.markdown("**간단하고 안정적인 TTS 시스템**")

# 사이드바 설정
with st.sidebar:
    st.header("⚙️ 설정")
    
    # 속도 설정
    speed = st.selectbox(
        "음성 속도",
        ["1", "2", "3", "4", "5"],
        index=2,
        format_func=lambda x: {
            "1": "매우 느림", "2": "느림", "3": "보통", 
            "4": "빠름", "5": "매우 빠름"
        }[x]
    )
    
    # 시작 행 설정
    start_row = st.number_input("시작 행", min_value=2, value=2, step=1)
    
    # 옵션들
    st.subheader("📋 옵션")
    auto_advance = st.checkbox("자동 진행", help="자동으로 다음 행으로 이동합니다.")
    announce_group = st.checkbox("브랜드/묶음 변화 알림", value=True, help="G열 값이 바뀔 때 알림합니다.")
    
    if auto_advance:
        auto_interval = st.slider("자동 진행 간격 (초)", 0.1, 5.0, 0.3, 0.1)

# 메인 컨텐츠
col1, col2 = st.columns([2, 1])

with col1:
    # 파일 업로드
    st.subheader("📂 파일 업로드")
    uploaded_file = st.file_uploader(
        "엑셀 파일을 선택하세요",
        type=['xlsx', 'xls'],
        help="엑셀 파일(.xlsx, .xls)을 업로드하세요"
    )
    
    if uploaded_file is not None:
        try:
            # 엑셀 파일 읽기
            wb = openpyxl.load_workbook(io.BytesIO(uploaded_file.read()), data_only=True)
            ws = wb.active
            
            # 데이터 추출
            data = []
            max_row = ws.max_row
            
            for row in range(2, max_row + 1):  # 헤더 제외
                row_data = {
                    'row': row,
                    'g': str(ws.cell(row=row, column=7).value or "").strip(),
                    'h': str(ws.cell(row=row, column=8).value or "").strip(),
                    'i': str(ws.cell(row=row, column=9).value or "").strip(),
                    'j': str(ws.cell(row=row, column=10).value or "").strip(),
                    'k': str(ws.cell(row=row, column=11).value or "").strip(),
                    'l': ws.cell(row=row, column=12).value
                }
                data.append(row_data)
            
            st.session_state.file_data = {
                'data': data,
                'max_row': max_row,
                'filename': uploaded_file.name
            }
            
            st.success(f"✅ {uploaded_file.name} 파일이 성공적으로 로드되었습니다! ({len(data)}개 행)")
            
        except Exception as e:
            st.error(f"❌ 파일 읽기 실패: {str(e)}")
    
    # 현재 행 표시
    if st.session_state.file_data:
        st.subheader("📊 현재 상태")
        
        # 진행률 표시
        progress = st.session_state.current_row / st.session_state.file_data['max_row']
        st.progress(progress, text=f"진행률: {st.session_state.current_row}/{st.session_state.file_data['max_row']} 행")
        
        # 현재 행 데이터 표시
        if st.session_state.current_row <= len(st.session_state.file_data['data']):
            current_data = st.session_state.file_data['data'][st.session_state.current_row - 2]
            
            st.markdown("**현재 행 데이터:**")
            col_a, col_b, col_c, col_d = st.columns(4)
            
            with col_a:
                st.metric("G열 (브랜드)", current_data['g'] or "없음")
            with col_b:
                st.metric("H열", current_data['h'] or "없음")
            with col_c:
                st.metric("I열 (상품명)", current_data['i'] or "없음")
            with col_d:
                st.metric("J열 (색상)", current_data['j'] or "없음")
            
            col_e, col_f = st.columns(2)
            with col_e:
                st.metric("K열 (사이즈)", current_data['k'] or "없음")
            with col_f:
                st.metric("L열 (수량)", current_data['l'] or "없음")

with col2:
    # 제어 버튼들
    st.subheader("🎮 제어")
    
    col_btn1, col_btn2 = st.columns(2)
    
    with col_btn1:
        if st.button("▶️ 시작", type="primary", use_container_width=True):
            if st.session_state.file_data:
                st.session_state.reading = True
                st.session_state.current_row = start_row
                st.session_state.prev_g_value = None
                st.rerun()
            else:
                st.warning("먼저 파일을 업로드하세요.")
    
    with col_btn2:
        if st.button("⏹️ 중지", type="secondary", use_container_width=True):
            st.session_state.reading = False
            st.rerun()
    
    # 단일 행 읽기
    if st.button("🔊 현재 행 읽기", use_container_width=True):
        if st.session_state.file_data and st.session_state.current_row <= len(st.session_state.file_data['data']):
            # 현재 행 읽기 로직
            current_data = st.session_state.file_data['data'][st.session_state.current_row - 2]
            
            # 읽을 텍스트 구성
            text_parts = []
            
            # G열 (브랜드/묶음) 처리
            g_value = clean_g_value(current_data['g'])
            if announce_group and g_value and g_value != st.session_state.prev_g_value:
                text_parts.append(g_value)
            
            st.session_state.prev_g_value = g_value
            
            # 상품명, 색상, 사이즈, 수량
            if current_data['i']:
                text_parts.append(current_data['i'])
            if current_data['j']:
                text_parts.append(current_data['j'])
            if current_data['k']:
                text_parts.append(convert_size(current_data['k']))
            if current_data['l'] and isinstance(current_data['l'], (int, float)) and current_data['l'] >= 2:
                text_parts.append(convert_quantity(int(current_data['l'])))
            
            combined_text = " ".join(text_parts)
            
            if combined_text.strip():
                st.info(f"🔊 읽을 내용: {combined_text}")
                
                # TTS 실행
                with st.spinner("음성을 재생하고 있습니다..."):
                    success = speak_text(combined_text)
                    if success:
                        st.success("✅ 음성 재생 완료!")
                    else:
                        st.error("❌ 음성 재생 실패")
            else:
                st.warning("읽을 내용이 없습니다.")
        else:
            st.warning("읽을 데이터가 없습니다.")
    
    # 다음 행
    if st.button("➡️ 다음 행", use_container_width=True):
        if st.session_state.file_data:
            st.session_state.current_row += 1
            if st.session_state.current_row > st.session_state.file_data['max_row']:
                st.session_state.current_row = st.session_state.file_data['max_row']
                st.info("마지막 행입니다.")
            st.rerun()
    
    # 이전 행
    if st.button("⬅️ 이전 행", use_container_width=True):
        if st.session_state.current_row > 2:
            st.session_state.current_row -= 1
            st.rerun()
    
    # 네이버 검색
    if st.button("🔍 네이버 검색", use_container_width=True):
        if st.session_state.file_data and st.session_state.current_row <= len(st.session_state.file_data['data']):
            current_data = st.session_state.file_data['data'][st.session_state.current_row - 2]
            h_value = current_data['h']
            i_value = current_data['i']
            
            if h_value and i_value:
                query = f"{h_value} {i_value}"
                naver_url = f"https://search.naver.com/search.naver?query={query}"
                
                # JavaScript로 새창 열기
                st.markdown(f"""
                <script>
                window.open('{naver_url}', '_blank');
                </script>
                """, unsafe_allow_html=True)
                
                st.success(f"🔍 '{query}' 검색을 새창에서 열었습니다.")
            else:
                st.warning("검색할 내용이 없습니다.")
    
    # 키보드 단축키 버튼들
    st.subheader("⌨️ 키보드 단축키")
    
    col_kb1, col_kb2, col_kb3 = st.columns(3)
    
    with col_kb1:
        if st.button("1️⃣ 다음 행", use_container_width=True, key="kb_1"):
            if st.session_state.file_data:
                st.session_state.current_row += 1
                if st.session_state.current_row > st.session_state.file_data['max_row']:
                    st.session_state.current_row = st.session_state.file_data['max_row']
                    st.info("마지막 행입니다.")
                st.rerun()
    
    with col_kb2:
        if st.button("2️⃣ 다시 읽기", use_container_width=True, key="kb_2"):
            if st.session_state.file_data:
                # 현재 행 읽기 로직
                current_data = st.session_state.file_data['data'][st.session_state.current_row - 2]
                
                # 읽을 텍스트 구성
                text_parts = []
                
                # G열 (브랜드/묶음) 처리
                g_value = clean_g_value(current_data['g'])
                if announce_group and g_value and g_value != st.session_state.prev_g_value:
                    text_parts.append(g_value)
                
                st.session_state.prev_g_value = g_value
                
                # 상품명, 색상, 사이즈, 수량
                if current_data['i']:
                    text_parts.append(current_data['i'])
                if current_data['j']:
                    text_parts.append(current_data['j'])
                if current_data['k']:
                    text_parts.append(convert_size(current_data['k']))
                if current_data['l'] and isinstance(current_data['l'], (int, float)) and current_data['l'] >= 2:
                    text_parts.append(convert_quantity(int(current_data['l'])))
                
                combined_text = " ".join(text_parts)
                
                if combined_text.strip():
                    st.info(f"🔊 읽을 내용: {combined_text}")
                    
                    # TTS 실행
                    with st.spinner("음성을 재생하고 있습니다..."):
                        success = speak_text(combined_text)
                        if success:
                            st.success("✅ 음성 재생 완료!")
                        else:
                            st.error("❌ 음성 재생 실패")
                else:
                    st.warning("읽을 내용이 없습니다.")
    
    with col_kb3:
        if st.button("3️⃣ 네이버 검색", use_container_width=True, key="kb_3"):
            if st.session_state.file_data and st.session_state.current_row <= len(st.session_state.file_data['data']):
                current_data = st.session_state.file_data['data'][st.session_state.current_row - 2]
                h_value = current_data['h']
                i_value = current_data['i']
                
                if h_value and i_value:
                    query = f"{h_value} {i_value}"
                    naver_url = f"https://search.naver.com/search.naver?query={query}"
                    
                    # JavaScript로 새창 열기
                    st.markdown(f"""
                    <script>
                    window.open('{naver_url}', '_blank');
                    </script>
                    """, unsafe_allow_html=True)
                    
                    st.success(f"🔍 '{query}' 검색을 새창에서 열었습니다.")
                else:
                    st.warning("검색할 내용이 없습니다.")

# 엑셀 데이터 미리보기
if st.session_state.file_data:
    st.subheader("📋 엑셀 데이터 미리보기")
    
    # 데이터프레임 생성
    df_data = []
    for item in st.session_state.file_data['data']:
        df_data.append({
            '행': item['row'],
            'G열(브랜드)': item['g'],
            'H열': item['h'],
            'I열(상품명)': item['i'],
            'J열(색상)': item['j'],
            'K열(사이즈)': item['k'],
            'L열(수량)': item['l']
        })
    
    df = pd.DataFrame(df_data)
    
    # 현재 행 하이라이트를 위한 스타일링
    def highlight_current_row(row):
        if row['행'] == st.session_state.current_row:
            return ['background-color: #ffeb3b'] * len(row)
        return [''] * len(row)
    
    # 스타일 적용
    styled_df = df.style.apply(highlight_current_row, axis=1)
    
    # 데이터프레임 표시
    st.dataframe(
        styled_df,
        use_container_width=True,
        height=400,
        hide_index=True
    )
    
    # 현재 행 정보
    st.markdown(f"**현재 선택된 행: {st.session_state.current_row}**")
    
    # 행 이동 버튼들
    col_nav1, col_nav2, col_nav3, col_nav4 = st.columns(4)
    
    with col_nav1:
        if st.button("⏮️ 처음으로", use_container_width=True):
            st.session_state.current_row = 2
            st.rerun()
    
    with col_nav2:
        if st.button("⬅️ 이전 행", use_container_width=True):
            if st.session_state.current_row > 2:
                st.session_state.current_row -= 1
                st.rerun()
    
    with col_nav3:
        if st.button("➡️ 다음 행", use_container_width=True):
            if st.session_state.current_row < st.session_state.file_data['max_row']:
                st.session_state.current_row += 1
                st.rerun()
    
    with col_nav4:
        if st.button("⏭️ 마지막으로", use_container_width=True):
            st.session_state.current_row = st.session_state.file_data['max_row']
            st.rerun()

# 자동 진행 처리
if st.session_state.reading and auto_advance:
    if st.session_state.current_row < st.session_state.file_data['max_row']:
        st.session_state.current_row += 1
        # 자동 진행 시에는 음성 재생하지 않고 행만 이동
        st.rerun()
    else:
        st.session_state.reading = False
        st.success("모든 행을 읽었습니다!")

# 사용법 안내
st.markdown("---")
st.markdown("### 🎯 사용법")
st.markdown("""
1. **파일 업로드**: 엑셀 파일을 업로드하세요
2. **설정 조정**: 사이드바에서 음성 속도를 조정하세요
3. **읽기 시작**: '현재 행 읽기' 버튼을 클릭하세요
4. **자동 진행**: 필요시 자동 진행을 활성화하세요
5. **키보드 단축키**: 1(다음 행), 2(다시 읽기), 3(네이버 검색) 버튼을 사용하세요
""")

# 테스트 버튼
st.markdown("### 🧪 음성 테스트")
if st.button("🔊 음성 테스트", use_container_width=True):
    test_text = "안녕하세요. 음성 테스트입니다."
    st.info(f"테스트 텍스트: {test_text}")
    
    with st.spinner("음성을 재생하고 있습니다..."):
        success = speak_text(test_text)
        if success:
            st.success("✅ 음성 테스트 성공!")
        else:
            st.error("❌ 음성 테스트 실패")