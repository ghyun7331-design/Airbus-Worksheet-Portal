import streamlit as st
import pandas as pd
import re

# ==========================================
# 1. 페이지 기본 설정 (모바일 및 데스크톱 폭 최적화)
# ==========================================
st.set_page_config(
    page_title="AIRBUS WORKSHEET PORTAL", 
    layout="wide", 
    initial_sidebar_state="collapsed"
)

# 고유 UI 디자인 및 뱃지 스타일 정의 (CSS 주입)
st.markdown("""
    <style>
    .data-card { 
        padding: 16px 20px; 
        border: 1px solid #cfd8dc; 
        border-radius: 6px; 
        margin-bottom: 12px; 
        background: #fff; 
        border-left: 5px solid #00b5e2; 
        word-wrap: break-word;
        overflow-wrap: break-word;
    }
    .gii-badge { 
        padding: 5px 10px; 
        border-radius: 3px; 
        font-size: 12px; 
        font-weight: bold; 
        margin-right: 6px; 
        display: inline-block; 
        margin-bottom: 6px; 
    }
    .fleet { background: #00205B; color: #fff; }
    .ata { background: #e2e8f0; color: #005587; border: 1px solid #cbd5e1; }
    .type { background: #dcfce7; color: #166534; border: 1px solid #bbf7d0; }
    .girii { background: #fee2e2; color: #991b1b; border: 1px solid #f87171; }
    .code { background: #fef08a; color: #854d0e; border: 1px solid #fde047; }
    
    .main-text { 
        font-size: 15px; 
        color: #00205B; 
        font-weight: bold; 
        margin-top: 6px; 
        line-height: 1.45; 
        word-break: keep-all;
    }
    .sub-text { 
        margin-top: 10px; 
        padding-top: 10px; 
        border-top: 1px dashed #e2e8f0; 
        font-size: 13px; 
        color: #475569; 
        white-space: pre-wrap; 
        word-break: break-word;
        background-color: #f8fafc;
        padding: 10px 14px;
        border-radius: 4px;
    }
    .mark { 
        background-color: #FFF9C4; 
        color: #D22730; 
        font-weight: bold; 
        padding: 0 2px; 
        border-radius: 2px;
    }
    .intro-box {
        text-align: center;
        padding: 60px 20px;
        background: linear-gradient(135deg, #f0f7ff 0%, #e0f2fe 100%);
        border: 2px dashed #bae6fd;
        border-radius: 12px;
        margin-top: 20px;
    }
    .intro-title {
        font-size: 28px;
        color: #005587;
        font-weight: 900;
        margin-bottom: 15px;
        letter-spacing: -0.5px;
    }
    .intro-emoji {
        font-size: 60px;
        margin-bottom: 20px;
        display: block;
    }
    .intro-desc {
        color: #475569;
        font-size: 16px;
        line-height: 1.6;
    }
    </style>
""", unsafe_allow_html=True)

# ==========================================
# 2. Session State 초기화 (시트 변경 시 검색창 리셋)
# ==========================================
if 'prev_sheet' not in st.session_state:
    st.session_state.prev_sheet = ""

if 'search_query_1' not in st.session_state:
    st.session_state.search_query_1 = ""

if 'search_query_2' not in st.session_state:
    st.session_state.search_query_2 = ""

# ==========================================
# 3. 데이터 로드 로직
# ==========================================
@st.cache_data
def load_data():
    file_path = 'data/Worksheet Asist Tool_V2.xlsx' 
    try:
        sheets = {
            "jascRaw": pd.read_excel(file_path, sheet_name='JASC_CODE', dtype=str).fillna(''),
            "cabinRaw": pd.read_excel(file_path, sheet_name='CABIN_CODE', dtype=str).fillna(''),
            "fml": pd.read_excel(file_path, sheet_name='FML_LIST', dtype=str).fillna(''),
            "gii": pd.read_excel(file_path, sheet_name='GII_RII_LIST', dtype=str).fillna(''),
            "nef": pd.read_excel(file_path, sheet_name='NEF_LIST', dtype=str).fillna('')
        }
        for key in sheets:
            sheets[key].columns = sheets[key].columns.astype(str).str.strip()
        return sheets
    except Exception as e:
        st.error(f"🚨 [데이터 로드 오류] {e}")
        return {k: pd.DataFrame() for k in ["jascRaw", "cabinRaw", "fml", "gii", "nef"]}

db = load_data()

# ==========================================
# 4. 모달 규정 내용 로직 (실제 엑셀 시트 연동)
# ==========================================
def get_gii_modal_content():
    try:
        df_gii_desc = pd.read_excel('data/Worksheet Asist Tool_V2.xlsx', sheet_name='GII_DESC', dtype=str).fillna('')
        content = ""
        for _, row in df_gii_desc.iterrows():
            if str(row.iloc[0]).strip():
                content += f"#### {row.iloc[0]}\n"
                if len(row) > 1 and str(row.iloc[1]).strip():
                    content += f"{row.iloc[1]}\n\n"
        return content if content else "GII 규정 상세 내용이 시트에 존재하지 않습니다."
    except:
        return "GII_DESC 시트를 읽어올 수 없습니다. 엑셀 파일에 시트가 있는지 확인해 주세요."

def get_nef_modal_content():
    try:
        df_nef_desc = pd.read_excel('data/Worksheet Asist Tool_V2.xlsx', sheet_name='NEF_DESC', dtype=str).fillna('')
        content = ""
        for _, row in df_nef_desc.iterrows():
            if str(row.iloc[0]).strip():
                content += f"#### {row.iloc[0]}\n"
                if len(row) > 1 and str(row.iloc[1]).strip():
                    content += f"{row.iloc[1]}\n\n"
        return content if content else "NEF 규정 상세 내용이 시트에 존재하지 않습니다."
    except:
        return "NEF_DESC 시트를 읽어올 수 없습니다. 엑셀 파일에 시트가 있는지 확인해 주세요."

@st.dialog("📖 일반검사항목(GII) 규정")
def gii_modal():
    st.warning("⚠️ 본 정보는 참고용 가이드라인입니다. 실제 정비 판정은 최신 AMM 마스터 본을 기준하십시오.")
    st.markdown(get_gii_modal_content())

@st.dialog("📖 NEF Description 규정")
def nef_modal():
    st.error("⚠️ [안전 주의] 최종 작업 결정은 인가된 최신 정비교범(MEL/CDL) 문서에 의거해야 합니다.")
    st.markdown(get_nef_modal_content())

# ==========================================
# 5. 헤더 및 UI 기본 설정
# ==========================================
st.markdown("## ✈️ WORKSHEET ASSIST PORTAL")
st.caption("Directed by 그런고래 with 휘년 | Python Streamlit Robust Search Edition")

sheet_options = {
    "": "====== 시트를 선택하세요 ======",
    "jascRaw": "📋 JASC CODE (ATA / Desc)",
    "cabinRaw": "💺 CABIN CODE (Code / Meaning)",
    "fml": "📝 FML LOG (Task / Desc)",
    "gii": "⚠️ GII/RII CRITERIA (Fleet/ATA/Req)",
    "nef": "🔧 NEF CRITERIA (Cat/Num/Desc)"
}

selected_key = st.selectbox(
    "1. 데이터 시트 선택", 
    options=list(sheet_options.keys()), 
    format_func=lambda x: sheet_options[x]
)

# 시트가 변경되었을 때 검색창 세션 리셋 로직
if selected_key != st.session_state.prev_sheet:
    st.session_state.search_query_1 = ""
    st.session_state.search_query_2 = ""
    st.session_state.prev_sheet = selected_key
    st.rerun()

# 모달 트리거 버튼
if selected_key == "gii":
    if st.button("📖 GII 규정 상세 내용 보기 ⤢", use_container_width=True):
        gii_modal()
elif selected_key == "nef":
    if st.button("📖 NEF 규정 상세 내용 보기 ⤢", use_container_width=True):
        nef_modal()

# ==========================================
# 6. 검색 엔진 (세션 스테이트 연동)
# ==========================================
if selected_key:
    col1, col2 = st.columns(2)
    with col1:
        primary_search = st.text_input("2. 1차 전체 범위 검색", key="search_query_1", placeholder="검색어 입력 (예: 0506 17, 754, Seat)...").strip().lower()
    with col2:
        secondary_search = st.text_input("3. 2차 결과 내 재검색", key="search_query_2", placeholder="결과 축소 및 좁히기...").strip().lower()

    df = db[selected_key]
    
    # 검색어가 하나라도 입력된 경우에만 필터링 및 결과 렌더링
    if primary_search or secondary_search:
        if primary_search and not df.empty:
            p_clean = primary_search.replace(' ', '').replace('-', '')
            df = df[df.apply(lambda row: p_clean in ''.join(row.astype(str)).lower().replace(' ', '').replace('-', ''), axis=1)]
            
        if secondary_search and not df.empty:
            s_clean = secondary_search.replace(' ', '').replace('-', '')
            df = df[df.apply(lambda row: s_clean in ''.join(row.astype(str)).lower().replace(' ', '').replace('-', ''), axis=1)]

        st.markdown(f"<div style='text-align:right; color:#005587; font-weight:bold;'>검색 결과: {len(df)}건</div>", unsafe_allow_html=True)

        # ==========================================
        # 7. 정비 항목 렌더링 카드 생성
        # ==========================================
        if len(df) == 0:
            st.info("조건에 일치하는 정비 데이터가 존재하지 않습니다.")
        else:
            for _, row in df.iterrows():
                badges = []
                main_text = ""
                sub_text = ""
                
                vals = [str(v).strip() for v in row.values if str(v).strip() and str(v).lower() != 'nan']

                if selected_key == 'jascRaw':
                    ata = row.get('ATA', vals[0] if len(vals) > 0 else '')
                    sec = row.get('SEC', vals[1] if len(vals) > 1 else '')
                    sub = row.get('SUB', vals[2] if len(vals) > 2 else '')
                    
                    if str(ata).strip(): badges.append(f"<span class='gii-badge fleet'>ATA {ata}</span>")
                    if str(sec).strip(): badges.append(f"<span class='gii-badge ata'>SEC {sec}</span>")
                    if str(sub).strip(): badges.append(f"<span class='gii-badge type'>SUB {sub}</span>")
                    
                    if str(ata).strip() and str(sub).strip():
                        combined_code = f"{ata}{sub}" if not str(sub).startswith(str(ata)) else sub
                        main_text = f"{ata} &nbsp;|&nbsp; {sec} &nbsp;|&nbsp; <span style='font-family: sans-serif; font-weight: 900; color: #D22730; font-size: 1.25em;'>{combined_code}</span>"
                    else:
                        main_text = " | ".join(vals[:min(3, len(vals))])
                    
                    desc_elements = vals[3:] if len(vals) > 3 else []
                    if desc_elements:
                        sub_text = " &nbsp;|&nbsp; ".join(desc_elements)

                elif selected_key == 'cabinRaw':
                    code = row.get('Code', vals[0] if len(vals) > 0 else '')
                    meaning = row.get('Meaning', vals[1] if len(vals) > 1 else '')
                    if str(code).strip(): badges.append(f"<span class='gii-badge code'>{code}</span>")
                    main_text = str(meaning) if str(meaning).strip() else "상세 내용 없음"
                    if len(vals) > 2: sub_text = " | ".join(vals[2:])

                elif selected_key == 'fml':
                    task = row.get('Task', vals[0] if len(vals) > 0 else '')
                    desc = row.get('Desc', vals[1] if len(vals) > 1 else '')
                    log = row.get('Log', vals[2] if len(vals) > 2 else '-')
                    note = row.get('Note', vals[3] if len(vals) > 3 else '')
                    is_log = str(log).upper() in ["O", "YES"]
                    badges.append(f"<span class='gii-badge {'fleet' if is_log else 'girii'}'>LOG: {log}</span>")
                    main_text = f"{task} : {desc}" if task or desc else "FML 기록 정보"
                    sub_text = f"Note: {note}" if str(note).strip() else ""
                    if len(vals) > 4: sub_text += "\n" + " | ".join(vals[4:])

                elif selected_key == 'gii':
                    fleet = row.get('Fleet', vals[0] if len(vals) > 0 else '')
                    ata_val = row.get('ATA', vals[1] if len(vals) > 1 else '')
                    gii_type = row.get('Type', vals[2] if len(vals) > 2 else '')
                    subject = row.get('Subject', vals[3] if len(vals) > 3 else '')
                    note = row.get('Note', vals[4] if len(vals) > 4 else '')
                    gii_rii = row.get('GiiRii', vals[5] if len(vals) > 5 else '')
                    if str(fleet).strip(): badges.append(f"<span class='gii-badge fleet'>{fleet}</span>")
                    if str(ata_val).strip(): badges.append(f"<span class='gii-badge ata'>ATA {ata_val}</span>")
                    if str(gii_type).strip(): badges.append(f"<span class='gii-badge type'>{gii_type}</span>")
                    if str(gii_rii).strip(): badges.append(f"<span class='gii-badge girii'>{gii_rii}</span>")
                    main_text = str(subject) if str(subject).strip() else "GII 정비 요구사항"
                    sub_text = f"Note: {note}" if str(note).strip() else ""
                    if len(vals) > 6: sub_text += "\n" + " | ".join(vals[6:])

                elif selected_key == 'nef':
                    main_sub = row.get('MainSub', vals[0] if len(vals) > 0 else '')
                    num = row.get('Number', vals[1] if len(vals) > 1 else '')
                    disc = row.get('Discrepancy', vals[2] if len(vals) > 2 else '')
                    note = row.get('Note', vals[3] if len(vals) > 3 else '')
                    if str(main_sub).strip(): badges.append(f"<span class='gii-badge fleet'>{main_sub}</span>")
                    if str(num).strip(): badges.append(f"<span class='gii-badge ata'>{num}</span>")
                    main_text = str(disc) if str(disc).strip() else "NEF 결함 항목"
                    sub_text = f"Note / Procedure:\n{note}" if str(note).strip() else ""
                    if len(vals) > 4: sub_text += "\n" + " | ".join(vals[4:])

                def highlight(text, term):
                    if not term or pd.isna(text): return str(text)
                    return re.sub(f"({re.escape(term)})", r"<span class='mark'>\1</span>", str(text), flags=re.IGNORECASE)

                main_text = highlight(highlight(main_text, primary_search), secondary_search)
                sub_text = highlight(highlight(sub_text, primary_search), secondary_search)

                sub_html = f"<div class='sub-text'>{sub_text}</div>" if sub_text.strip() else ""
                card_html = f"""
                <div class="data-card">
                    <div>{''.join(badges)}</div>
                    <div class="main-text">{main_text}</div>
                    {sub_html}
                </div>
                """
                st.markdown(card_html, unsafe_allow_html=True)
    
    # 검색어가 없는 초기 상태 (고래 인트로 화면 출력)
    else:
        st.markdown(f"""
            <div class="intro-box">
                <span class="intro-emoji">🐳</span>
                <div class="intro-title">그런고래 검색 엔진 준비 완료!</div>
                <div class="intro-desc">
                    위 검색창에 원하시는 작업 번호나 키워드를 입력해 주세요.<br>
                    <strong>[{sheet_options[selected_key].split('(')[0].strip()}]</strong> 시트의 데이터를 빠르게 찾아드립니다.
                </div>
            </div>
        """, unsafe_allow_html=True)

else:
    st.markdown("""
        <div style="padding: 40px 10px; text-align: center; color: #999; font-weight: bold; font-size: 15px; border: 1px solid #cfd8dc; border-radius: 6px; background: #fff;">
            상단의 콤보박스에서 데이터 시트를 먼저 지정해 주십시오.
        </div>
    """, unsafe_allow_html=True)