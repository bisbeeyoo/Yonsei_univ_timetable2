import streamlit as st
import pandas as pd
import os
import re

# ==========================================
#  🦅 CORE: 레포지토리에 함께 올린 내장 엑셀 파일을 자동 파싱하는 엔진
# ==========================================
@st.cache_data
def load_and_parse_yonsei_excel():
    """
    깃허브 레포지토리에 app.py와 함께 업로드한 'time_table1(2025-2).xls' 파일을
    외부 네트워크 통신 없이 내부 서버 경로에서 직접 안전하게 읽어와 파싱합니다.
    """
    # 깃허브 폴더 내에 함께 존재하는 시간표 파일 이름 지정
    local_filename = "time_table1(2025-2).xls"
    
    if not os.path.exists(local_filename):
        st.error(f"❌ '{local_filename}' 파일을 찾을 수 없습니다. GitHub 레포지토리에 app.py와 함께 해당 엑셀 파일을 업로드했는지 확인해 주세요!")
        return pd.DataFrame()
        
    lines = []
    try:
        # 1차 시도: 표준 엑셀 서식으로 로컬 파일 읽기
        excel_df = pd.read_excel(local_filename, header=None)
        for idx, row in excel_df.iterrows():
            line_str = ",".join([str(val).strip() if pd.notna(val) else "" for val in row])
            lines.append(line_str + "\n")
            
    except Exception:
        try:
            # 2차 시도: 텍스트 기반 CSV 형태로 인코딩 예외 처리하며 읽기
            with open(local_filename, "rb") as f:
                file_bytes = f.read()
            try:
                text_content = file_bytes.decode("utf-8")
            except UnicodeDecodeError:
                text_content = file_bytes.decode("cp949")
                
            lines = text_content.splitlines(keepends=True)
        except Exception as e:
            st.error(f"❌ 내부 시간표 파일을 읽는 중 오류가 발생했습니다: {e}")
            return pd.DataFrame()
        
    parsed_courses = []
    current_major = "공통/교직"
    
    idx = 0
    while idx < len(lines):
        line = lines[idx].strip()
        if line and not line.startswith(",") and ",,,," in line:
            current_major = line.split(",")[0].strip()
            idx += 1
            continue
            
        if re.search(r"구\s*분", line):
            header_parts = [p.strip() for p in line.split(",")]
            days_in_block = [p for p in header_parts if p in ["월", "화", "수", "목", "금"]]
            
            idx += 1
            while idx < len(lines) and not re.search(r"구\s*분", lines[idx]) and ",,,," not in lines[idx]:
                block_line = lines[idx].strip()
                if "1,2교시" in block_line: current_period = "1,2"
                elif "3,4교시" in block_line: current_period = "3,4"
                    
                if "과 목 종 별" in block_line:
                    try:
                        types = lines[idx].strip().split(",")[2:]
                        codes = lines[idx+1].strip().split(",")[2:]
                        names = lines[idx+2].strip().split(",")[2:]
                        profs = lines[idx+3].strip().split(",")[2:]
                        rooms = lines[idx+4].strip().split(",")[2:]
                        
                        for col_idx, day in enumerate(days_in_block):
                            if col_idx < len(names) and names[col_idx].strip():
                                c_name = names[col_idx].strip()
                                if "(영어)" in codes[col_idx]: c_name += " (영어)"
                                h_code = codes[col_idx].split("(")[0].strip()
                                credit = 2 if "SPT" in h_code else 3
                                
                                final_major = current_major
                                if "교직" in current_major:
                                    if "SPT" in h_code: final_major = "교직(자격증)"
                                    elif "SPL" in h_code: final_major = "평생교육사"
                                    else: final_major = "교직(공통)"
                                
                                parsed_courses.append({
                                    "전공": final_major, "요일": day, "교시": current_period,
                                    "과목종별": types[col_idx].strip() if col_idx < len(types) else "전공",
                                    "학정번호": h_code, "과목명": c_name,
                                    "교수명": profs[col_idx].strip() if col_idx < len(profs) else "미지정",
                                    "강의실": rooms[col_idx].strip() if col_idx < len(rooms) else "미지정", "학점": credit
                                })
                    except Exception: pass
                    idx += 5
                    continue
                idx += 1
            continue
        idx += 1

    df = pd.DataFrame(parsed_courses).drop_duplicates(subset=['학정번호', '요일', '교시'])
    if not df.empty:
        df['time_slots_set'] = df.apply(lambda r: set((r['요일'], int(p)) for p in r['교시'].split(',')), axis=1)
    return df


# --- [UI/UX] 에브리타임 감성의 깔끔한 Pretendard 폰트 및 연세 블루 스타일 세팅 ---
st.set_page_config(page_title="YONSEI GS-ED Timetable", layout="wide", initial_sidebar_state="expanded")

st.markdown("""
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Pretendard:wght@400;600;800&display=swap');
        * { font-family: 'Pretendard', sans-serif !important; }
        .main-title { font-size: 2.2rem; font-weight: 800; color: #112F6F; margin-bottom: 5px; }
        .sub-title { font-size: 1rem; color: #64748B; margin-bottom: 25px; }
        .card { background-color: #F8FAFC; padding: 18px; border-radius: 12px; border: 1px solid #E2E8F0; margin-bottom: 12px; }
        .stTabs [data-baseweb="tab"] { font-weight: 600; color: #64748B; font-size: 15px; }
        .stTabs [data-baseweb="tab"][aria-selected="true"] { color: #112F6F !important; border-bottom-color: #112F6F !important; }
    </style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-title">🦅 YONSEI GS-ED</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-title">연세대학교 교육대학원 수강신청 시간표 도우미 (공유 & 저장 지원)</div>', unsafe_allow_html=True)

PREDEFINED_COLORS = ["#E2EFFE", "#FEE2E2", "#FEF3C7", "#E0F2FE", "#ECEFEE", "#F3E8FF", "#ECFDF5", "#FFF1F2", "#F0FDFA", "#EFF6FF"]

# 자체 시스템 내부에서 엑셀 자동 연동 처리
master_df = load_and_parse_yonsei_excel()

if master_df.empty:
    st.stop()

# 수강 신청 및 컬러 매핑 상태 관리 초기화
if 'my_courses' not in st.session_state: st.session_state.my_courses = []
if 'color_map' not in st.session_state: st.session_state.color_map = {}

# --- 🔗 공유된 링크(Query Parameter) 파싱 및 복원 시스템 ---
query_dict = st.query_params.to_dict()
if "courses" in query_dict and not st.session_state.my_courses:
    try:
        courses_str = query_dict["courses"]
        if courses_str:
            shared_courses = [c for c in courses_str.split(',') if not master_df[master_df['학정번호'] == c].empty]
            if shared_courses:
                st.session_state.my_courses = shared_courses
                for h_no in shared_courses:
                    name = master_df[master_df['학정번호'] == h_no].iloc[0]['과목명']
                    if name not in st.session_state.color_map:
                        st.session_state.color_map[name] = PREDEFINED_COLORS[len(st.session_state.color_map) % len(PREDEFINED_COLORS)]
                st.rerun()
    except Exception:
        for key in list(st.query_params.keys()):
            del st.query_params[key]

# 중복 및 공강 시간 자동 필터링 엔진
def get_available_courses(df, selected_ids):
    if not selected_ids: return df
    available_df = df[~df['학정번호'].isin(selected_ids)]
    my_busy_slots = set().union(*df[df['학정번호'].isin(selected_ids)]['time_slots_set'])
    return available_df[available_df['time_slots_set'].apply(lambda s: s.isdisjoint(my_busy_slots))]

available_df = get_available_courses(master_df, st.session_state.my_courses)


# ==========================================
#  LAYOUT SIDEBAR: 에타 감성의 통합 검색창 & 필터
# ==========================================
with st.sidebar:
    st.markdown("### 🛠️ 강좌 검색 및 필터 패널")
    search_query = st.text_input("🔎 과목명 또는 교수명 검색", placeholder="예: 학습과학 또는 이희승...").strip().lower()
    st.write("---")
    
    tab_m, tab_k = st.tabs(["🎓 전공 강좌 조회", "🍎 교직/공통 조회"])
    
    with tab_m:
        major_list = sorted([m for m in master_df['전공'].unique() if "교직" not in m and "평생" not in m])
        if major_list:
            selected_major = st.selectbox("소속 전공 선택", major_list)
            filtered_major = available_df[available_df['전공'] == selected_major]
            if search_query:
                filtered_major = filtered_major[filtered_major['과목명'].str.lower().str.contains(search_query) | filtered_major['교수명'].str.lower().str.contains(search_query)]
                
            if not filtered_major.empty:
                sel_idx = st.selectbox("과목 선택", options=filtered_major.index, 
                                       format_func=lambda idx: f"[{filtered_major.loc[idx]['요일']}] {filtered_major.loc[idx]['과목명']} - {filtered_major.loc[idx]['교수명']}")
                if st.button("➕ 시간표에 전공 추가", use_container_width=True, type="primary"):
                    row = filtered_major.loc[sel_idx]
                    st.session_state.my_courses.append(row['학정번호'])
                    st.session_state.color_map[row['과목명']] = PREDEFINED_COLORS[len(st.session_state.color_map) % len(PREDEFINED_COLORS)]
                    st.query_params["courses"] = ",".join(st.session_state.my_courses)
                    st.rerun()
            else:
                st.caption("현재 추가 가능한 전공 과목이 없습니다.")
        else:
            st.caption("조회된 전공이 없습니다.")

    with tab_k:
        kyojik_type = st.selectbox("교직/공통 분류", ["전체", "교직(공통)", "교직(자격증)", "평생교육사"])
        filtered_k = available_df[available_df['전공'].str.contains("교직|평생")]
        if kyojik_type != "전체":
            filtered_k = filtered_k[filtered_k['전공'] == kyojik_type]
        if search_query:
            filtered_k = filtered_k[filtered_k['과목명'].str.lower().str.contains(search_query) | filtered_k['교수명'].str.lower().str.contains(search_query)]
            
        if not filtered_k.empty:
            sel_idx_k = st.selectbox("과목 선택 ", options=filtered_k.index, 
                                     format_func=lambda idx: f"[{filtered_k.loc[idx]['요일']}] {filtered_k.loc[idx]['과목명']} - {filtered_k.loc[idx]['교수명']}")
            if st.button("➕ 시간표에 교직 추가", use_container_width=True, type="primary"):
                row = filtered_k.loc[sel_idx_k]
                st.session_state.my_courses.append(row['학정번호'])
                st.session_state.color_map[row['과목명']] = PREDEFINED_COLORS[len(st.session_state.color_map) % len(PREDEFINED_COLORS)]
                st.query_params["courses"] = ",".join(st.session_state.my_courses)
                st.rerun()
        else:
            st.caption("현재 추가 가능한 교직 과목이 없습니다.")


# ==========================================
#  LAYOUT MAIN: 시각화 시간표 보드 & 장바구니 리스트
# ==========================================
if not st.session_state.my_courses:
    st.info("💡 왼쪽 사이드바에서 소속 전공이나 교직 과목을 선택하시면, 실시간 강좌 리스트가 나타납니다!")
else:
    my_df = master_df[master_df['학정번호'].isin(st.session_state.my_courses)].drop_duplicates(subset=['학정번호'])
    total_credits = my_df['학점'].sum()
    
    col_a, col_b = st.columns([0.75, 0.25])
    with col_a:
        st.markdown(f"### 🗓️ MY TIMETABLE `[ 총 {len(my_df)} 과목 / {total_credits} 학점 이수 ]`")
    with col_b:
        if st.button("🗑️ 전체 초기화", use_container_width=True):
            st.session_state.my_courses, st.session_state.color_map = [], {}
            for key in list(st.query_params.keys()):
                del st.query_params[key]
            st.rerun()

    days = ['월', '화', '수', '목', '금']
    periods = [1, 2, 3, 4]
    time_labels = {1: "1교시<br><small>18:20-19:10</small>", 2: "2교시<br><small>19:15-20:05</small>", 3: "3교시<br><small>20:10-21:00</small>", 4: "4교시<br><small>21:05-21:55</small>"}
    grid = {(p, d): {"text": "", "color": "#FFFFFF", "span": 1, "visible": True} for p in periods for d in days}
    
    for _, row in master_df[master_df['학정번호'].isin(st.session_state.my_courses)].iterrows():
        color = st.session_state.color_map.get(row['과목명'], "#FFFFFF")
        p_list = sorted([int(p) for p in row['교시'].split(',')])
        if len(p_list) == 2 and p_list[1] == p_list[0] + 1:
            grid[(p_list[0], row['요일'])] = {
                "text": f"<div style='font-weight:700; color:#1E293B; font-size:13px;'>{row['과목명']}</div><div style='font-size:11px; margin-top:4px; color:#64748B;'>{row['교수명']} · {row['강의실']}</div>",
                "color": color, "span": 2, "visible": True
            }
            grid[(p_list[1], row['요일'])]["visible"] = False
        else:
            for p in p_list:
                grid[(p, row['요일'])] = {
                    "text": f"<div style='font-weight:700; color:#1E293B; font-size:13px;'>{row['과목명']}</div><div style='font-size:11px; margin-top:4px; color:#64748B;'>{row['교수명']} · {row['강의실']}</div>",
                    "color": color, "span": 1, "visible": True
                }

    table_html = """
    <div id="capture-area" style="padding: 16px; background: #ffffff; border-radius: 16px; border: 1px solid #E2E8F0;">
    <table style="width:100%; border-collapse:separate; border-spacing: 6px; text-align:center; table-layout:fixed;">
        <thead>
            <tr style="height:40px; background-color:#F1F5F9;">
                <th style="border-radius:8px; color:#475569; font-size:12px; font-weight:600; width:13%;">TIME</th>
    """
    for d in days:
        table_html += f'<th style="border-radius:8px; color:#475569; font-size:13px; font-weight:600;">{d}요일</th>'
    table_html += '</tr></thead><tbody>'
    
    for p in periods:
        table_html += f'<tr style="height:85px;"><td style="background-color:#F8FAFC; border-radius:8px; color:#64748B; font-size:11px; font-weight:600; padding:5px; line-height:1.4;">{time_labels[p]}</td>'
        for d in days:
            cell = grid[(p, d)]
            if cell["visible"]:
                bg = cell["color"]
                border_radius = "border-radius: 10px;" if cell["text"] else "border-radius: 10px; background-color:#F8FAFC; border: 1px dashed #E2E8F0;"
                table_html += f'<td rowspan="{cell["span"]}" style="{border_radius} background-color:{bg}; padding:10px;">{cell["text"]}</td>'
        table_html += '</tr>'
    table_html += '</tbody></table></div>'

    js_downloader = """
    <script src="https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js"></script>
    <div style="display: flex; gap: 10px; margin-top: 15px;">
        <button id="download-btn" style="flex: 1; padding: 12px; background-color: #112F6F; color: white; border: none; border-radius: 10px; cursor: pointer; font-weight: 600; font-size: 14px;">✨ 감성 시간표 이미지 저장하기</button>
    </div>
    <div id="status-log" style="margin-top:8px; font-size:12px; text-align:center; font-weight:600;"></div>
    <script>
        document.getElementById('download-btn').onclick = function() {
            const area = document.getElementById("capture-area");
            const log = document.getElementById('status-log');
            log.innerText = '📸 예쁜 고화질 시간표 이미지 제작 중...'; log.style.color = '#112F6F';
            html2canvas(area, {scale: 3, backgroundColor: '#ffffff', borderRadius: 16}).then(canvas => {
                const a = document.createElement("a");
                a.href = canvas.toDataURL("image/png");
                a.download = "YONSEI_TIMETABLE.png";
                document.body.appendChild(a); a.click(); document.body.removeChild(a);
                log.innerText = '✅ 다운로드 폴더에 시간표 이미지가 안전하게 저장되었습니다!'; log.style.color = '#059669';
            }).catch(e => { log.innerText = '❌ 에러가 발생했습니다: ' + e; log.style.color = '#DC2626'; });
        };
    </script>
    """
    st.components.v1.html(table_html + js_downloader, height=500)
    
    # --- 🔗 공유 시스템용 링크 구성 ---
    st.write(" ")
    share_link = f"https://yonseitimetable.streamlit.app/?courses={','.join(st.session_state.my_courses)}"
    
    st.markdown("#### 🔗 내 시간표 친구에게 링크로 공유하기")
    col_link, col_btn = st.columns([0.8, 0.2])
    with col_link:
        st.text_input("공유용 링크", value=share_link, readonly=True, label_visibility="collapsed")
    with col_btn:
        if st.button("📋 링크 복사", use_container_width=True):
            st.toast("링크가 복사되었습니다! 친구에게 공유해 보세요. 🦅")
            
    st.write("---")
    st.markdown("#### 📝 확정된 장바구니 강좌 상세 내역")
    
    for idx, row in my_df.iterrows():
        st.markdown(f"""
        <div class="card">
            <div style="display: flex; justify-content: space-between; align-items: center;">
                <div>
                    <span style="background-color:#E0F2FE; color:#0369A1; padding:3px 8px; border-radius:6px; font-size:11px; font-weight:600; margin-right:8px;">{row['과목종별']}</span>
                    <strong style="font-size:15px; color:#1E293B;">{row['과목명']}</strong>
                    <span style="font-size:13px; color:#64748B; margin-left:10px;">| {row['교수명']} 교수님 · {row['강의실']} ({row['요일']}요일 {row['교시']}교시)</span>
                </div>
                <div style="font-size:12px; color:#94A3B8;">학정번호: {row['학정번호']} ({row['학점']}학점)</div>
            </div>
        </div>
        """, unsafe_allow_html=True)
        
        if st.button("과목 제외", key=f"del-{row['학정번호']}", type="secondary"):
            st.session_state.my_courses.remove(row['학정번호'])
            if st.session_state.my_courses:
                st.query_params["courses"] = ",".join(st.session_state.my_courses)
            else:
                for key in list(st.query_params.keys()):
                    del st.query_params[key]
            st.rerun()
