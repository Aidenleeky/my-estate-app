import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import re
import os
import time

# ==============================================================================
# 1. 기본 설정 및 데이터 로드/저장 함수
# ==============================================================================
st.set_page_config(page_title="원주민부동산 매물 평가", layout="wide")
st.title("🏢 [원주민부동산중개법인] 고부가 매물 종합 평가 AI")

WEIGHTS = {
    "계약확율": 10, "중개보수지급성향": 10, "중개팀_집중역량": 10,
    "광고협조와_효율": 10, "시투입_가성비": 10, "시장선호도": 10,
    "가격적정성": 7, "조건유동성": 7, "안전성": 7,
    "방문용이성": 7, "성장성": 7, "자율1": 2.5, "자율2": 2.5,
}

DB_FILE = "매물_순위_평가_최신DB.xlsx"

def load_db():
    if os.path.exists(DB_FILE):
        return pd.read_excel(DB_FILE)
    return pd.DataFrame()

def save_db(df):
    df.to_excel(DB_FILE, index=False)

# 엑셀의 빈칸(nan)을 완벽하게 처리해주는 함수
def get_val(r, key):
    val = r.get(key, "")
    if pd.isna(val) or str(val).strip().lower() == 'nan':
        return ""
    return str(val).strip()

# 🔗 주소에 http/https를 자동으로 붙여주는 함수
def clean_url(url):
    url = str(url).strip()
    if not url or url.lower() == 'nan':
        return ""
    if not url.startswith(("http://", "https://")):
        return f"https://{url}"
    return url

if 'db' not in st.session_state:
    st.session_state.db = load_db()
if 'edit_target' not in st.session_state:
    st.session_state.edit_target = None

# ==============================================================================
# 2. 사이드바 (제어판)
# ==============================================================================
with st.sidebar:
    st.header("⚙️ 제어판")
    
    if st.button("📥 최신 DB 강제불러오기"):
        st.session_state.db = load_db()
        st.success(f"✅ 총 {len(st.session_state.db)}개의 매물을 불러왔습니다!")
        time.sleep(1)
        st.rerun()

    st.subheader("🔍 검색 및 필터")
    search_kwd = st.text_input("매물명 실시간 검색")

    options = ["-"] + st.session_state.db['매물명'].astype(str).tolist() if not st.session_state.db.empty else ["-"]
    
    default_idx = 0
    if st.session_state.edit_target in options:
        default_idx = options.index(st.session_state.edit_target)
        
    edit_sel = st.selectbox("🔄 수정/삭제할 매물 선택", options, index=default_idx)

    if edit_sel != "-" and st.session_state.edit_target != edit_sel:
        st.session_state.edit_target = edit_sel
        st.rerun()
    elif edit_sel == "-" and st.session_state.edit_target is not None:
        st.session_state.edit_target = None
        st.rerun()

    if st.button("❌ 선택 매물 삭제", type="primary"):
        if edit_sel != "-":
            st.session_state.db = st.session_state.db[st.session_state.db['매물명'].astype(str) != edit_sel]
            save_db(st.session_state.db)
            st.session_state.edit_target = None
            st.success("삭제되었습니다.")
            st.rerun()

    if st.button("🆕 새 매물 (폼 초기화)"):
        st.session_state.edit_target = None
        st.rerun()

# ==============================================================================
# 3. 데이터 로드 (수정 모드일 때)
# ==============================================================================
row = {}
if st.session_state.edit_target and not st.session_state.db.empty:
    target_df = st.session_state.db[st.session_state.db['매물명'].astype(str) == st.session_state.edit_target]
    if not target_df.empty:
        row = target_df.iloc[0].to_dict()

# 화면 새로고침 시 입력칸을 완벽하게 새로 그리기 위한 고유 이름표
suffix = str(st.session_state.edit_target) if st.session_state.edit_target else "new"

# ==============================================================================
# 4. 메인 화면 (입력 폼)
# ==============================================================================
with st.form("property_form"):
    tab1, tab2, tab3 = st.tabs(["📋 매물 정보", "👤 가망 고객 (1~6)", "🎯 13대 핵심 지표"])

    with tab1:
        col1, col2 = st.columns(2)
        with col1:
            title = st.text_input("매물명 (필수)", value=get_val(row, "매물명"), key=f"title_{suffix}")
            own_name = st.text_input("소유자 존함", value=get_val(row, "소유자_존함"), key=f"own_name_{suffix}")
            own_phone = st.text_input("소유자 연락처", value=get_val(row, "소유자_연락처") or "010", key=f"own_phone_{suffix}")
            agt_name = st.text_input("대리인 존함", value=get_val(row, "대리인_존함"), key=f"agt_name_{suffix}")
            agt_phone = st.text_input("대리인 연락처", value=get_val(row, "대리인_연락처") or "010", key=f"agt_phone_{suffix}")

        with col2:
            core_analysis = st.text_area("💡🎯 킬러앱 (고객 핵심 분석)", value=get_val(row, "핵심쟁점_분석"), height=90, key=f"core_{suffix}")
            memo = st.text_area("주요 메모 사항", value=get_val(row, "주요_메모_사항"), height=90, key=f"memo_{suffix}")

            st.markdown("**📅 연락 주기 설정**")
            c1, c2 = st.columns(2)
            last_contact_val = get_val(row, "매물측_최종연락일시")
            last_contact = c1.date_input("최종연락일", value=pd.to_datetime(last_contact_val).date() if last_contact_val else datetime.now().date(), key=f"last_contact_{suffix}")
            
            cycle_opts = ["선택안함", "매일(1일)", "3일", "매주(7일)", "10일", "격주(14일)", "매월(30일)", "격월(60일)", "3개월(90일)", "6개월(180일)", "1년(365일)", "2년(730일)"]
            current_cycle = get_val(row, "연락_주기")
            if current_cycle not in cycle_opts:
                current_cycle = "선택안함"
            cycle_opt = c2.selectbox("연락 주기", cycle_opts, index=cycle_opts.index(current_cycle), key=f"cycle_opt_{suffix}")

        st.markdown("**🌐 매물 URL 및 메모 (최대 5개)**")
        url_cols = st.columns(5)
        urls = {}
        for i in range(1, 6):
            with url_cols[i-1]:
                urls[f"url_{i}"] = st.text_input(f"URL {i}", value=get_val(row, f"매물_URL{i}"), key=f"u_{i}_{suffix}")
                # 🔗 링크 열기 버튼 추가
                link_target = clean_url(urls[f"url_{i}"])
                if link_target:
                    st.link_button("🔗 링크 열기", link_target, use_container_width=True)
                
                urls[f"memo_{i}"] = st.text_input(f"URL {i} 메모", value=get_val(row, f"매물_URL{i}_메모"), key=f"um_{i}_{suffix}")

    with tab2:
        st.subheader("👤 가망 고객 (최대 6명)")
        buyers = {}
        for i in range(1, 7):
            with st.expander(f"가망 고객 {i}", expanded=(i==1)):
                c1, c2, c3, c4 = st.columns(4)
                buyers[f"name_{i}"] = c1.text_input(f"존함", value=get_val(row, f"가망{i}_존함"), key=f"bn_{i}_{suffix}")
                buyers[f"phone_{i}"] = c2.text_input(f"연락처", value=get_val(row, f"가망{i}_연락처") or "010", key=f"bp_{i}_{suffix}")
                buyers[f"memo_{i}"] = c3.text_input(f"메모/요청", value=get_val(row, f"가망{i}_메모"), key=f"bm_{i}_{suffix}")
                
                # 가망 고객 참조 URL 및 링크 버튼
                buyers[f"url_{i}"] = c4.text_input(f"참조 URL", value=get_val(row, f"가망{i}_URL"), key=f"bu_{i}_{suffix}")
                link_target_buyer = clean_url(buyers[f"url_{i}"])
                if link_target_buyer:
                    with c4:
                        st.link_button("🔗 참조 링크 열기", link_target_buyer, use_container_width=True)
                
                cdate_val = get_val(row, f"가망{i}_최종컨택일")
                buyers[f"cdate_{i}"] = st.date_input(f"최종컨택일", value=pd.to_datetime(cdate_val).date() if cdate_val else None, key=f"bd_{i}_{suffix}")

    with tab3:
        st.subheader("🎯 13대 핵심 지표 평가 (0~10점)")
        scores = {}
        custom_names = {}
        cols = st.columns(2)
        idx = 0
        for k, weight in WEIGHTS.items():
            col = cols[idx % 2]
            val_score = pd.to_numeric(row.get(f"점수_{k}", 5), errors='coerce')
            val_score = int(val_score) if pd.notna(val_score) else 5
            val_score = max(0, min(10, val_score))

            if k in ["자율1", "자율2"]:
                num = "1" if k == "자율1" else "2"
                default_cn = get_val(row, f"자율지표{num}_이름") or f"자율지표{num}"
                custom_names[k] = col.text_input(f"자율지표 {num} 이름", value=default_cn, key=f"cn_{k}_{suffix}")
                scores[k] = col.slider(f"{custom_names[k]} (가중치 {weight})", 0, 10, val_score, key=f"s_{k}_{suffix}")
            else:
                scores[k] = col.slider(f"{k} (가중치 {weight})", 0, 10, val_score, key=f"s_{k}_{suffix}")
            idx += 1

    st.markdown("<br>", unsafe_allow_html=True)
    submitted = st.form_submit_button("🌟 매물 등록 및 업데이트 (즉시 저장) 🌟", type="primary", use_container_width=True)

# ==============================================================================
# 5. 저장 로직
# ==============================================================================
if submitted:
    if not title.strip():
        st.error("⚠️ 매물명을 입력해주세요.")
    else:
        total_weight = sum(WEIGHTS.values())
        final_score = sum(scores[k] * (WEIGHTS[k] / total_weight) for k in WEIGHTS)
        raw_sum = sum(scores.values())
        simple_avg = round(raw_sum / len(WEIGHTS), 2)
        
        new_data = {
            "매물명": title, "종합_성사_점수": round(final_score, 2), "점수합계": raw_sum, "단순평균": simple_avg,
            "소유자_존함": own_name, "소유자_연락처": own_phone,
            "대리인_존함": agt_name, "대리인_연락처": agt_phone,
            "핵심쟁점_분석": core_analysis, "주요_메모_사항": memo,
            "매물측_최종연락일시": str(last_contact) if last_contact else "",
            "연락_주기": cycle_opt
        }
        for k in WEIGHTS.keys():
            new_data[f"점수_{k}"] = scores[k]
            if k in ["자율1", "자율2"]:
                num = "1" if k == "자율1" else "2"
                new_data[f"자율지표{num}_이름"] = custom_names[k]
            
        for i in range(1, 6):
            new_data[f"매물_URL{i}"] = urls[f"url_{i}"]
            new_data[f"매물_URL{i}_메모"] = urls[f"memo_{i}"]
            
        for i in range(1, 7):
            new_data[f"가망{i}_존함"] = buyers[f"name_{i}"]
            new_data[f"가망{i}_연락처"] = buyers[f"phone_{i}"]
            new_data[f"가망{i}_메모"] = buyers[f"memo_{i}"]
            new_data[f"가망{i}_URL"] = buyers[f"url_{i}"]
            new_data[f"가망{i}_최종컨택일"] = str(buyers[f"cdate_{i}"]) if buyers[f"cdate_{i}"] else ""

        db = st.session_state.db
        if st.session_state.edit_target:
            db = db[db['매물명'].astype(str) != st.session_state.edit_target]
        db = db[db['매물명'].astype(str) != title]
        
        db = pd.concat([db, pd.DataFrame([new_data])], ignore_index=True)
        st.session_state.db = db
        save_db(db)
        
        st.success(f"✅ '{title}' 매물이 성공적으로 저장되었습니다! (종합 점수: {final_score:.2f}점)")
        st.session_state.edit_target = None
        time.sleep(1)
        st.rerun()

# ==============================================================================
# 6. 하단 대시보드 (순위 표)
# ==============================================================================
st.markdown("---")
st.subheader("📊 실시간 매물 순위 대시보드 (매물명을 클릭하면 위 폼으로 불러옵니다)")

if not st.session_state.db.empty:
    view_df = st.session_state.db.copy()
    
    if '종합_성사_점수' in view_df.columns:
        view_df['종합_성사_점수'] = pd.to_numeric(view_df['종합_성사_점수'], errors='coerce').fillna(0)
        view_df = view_df.sort_values(by="종합_성사_점수", ascending=False).reset_index(drop=True)
    
    if search_kwd:
        view_df = view_df[view_df.astype(str).apply(lambda col: col.str.contains(search_kwd, case=False, na=False)).any(axis=1)]

    if not view_df.empty:
        display_cols = ["매물명", "종합_성사_점수", "점수합계", "소유자_존함", "소유자_연락처", "대리인_연락처", "매물측_최종연락일시", "연락_주기"]
        display_cols = [c for c in display_cols if c in view_df.columns]
        
        def make_tel_link(phone):
            if pd.isna(phone) or str(phone).strip() in ["", "010", "nan"]: return None
            clean = re.sub(r'[^0-9]', '', str(phone))
            return f"tel:{clean}"
            
        for col in ["소유자_연락처", "대리인_연락처"]:
            if col in view_df.columns:
                view_df[col] = view_df[col].apply(make_tel_link)

        event = st.dataframe(
            view_df[display_cols], 
            use_container_width=True, 
            hide_index=True,
            on_select="rerun",
            selection_mode="single-row",
            column_config={
                "매물명": st.column_config.TextColumn("매물명 (클릭하여 불러오기)", width="medium"),
                "종합_성사_점수": st.column_config.NumberColumn("종합 점수", format="%.2f"),
                "소유자_연락처": st.column_config.LinkColumn("소유자 연락처", display_text="📞 전화걸기"),
                "대리인_연락처": st.column_config.LinkColumn("대리인 연락처", display_text="📞 전화걸기")
            }
        )
        
        if len(event.selection.rows) > 0:
            selected_idx = event.selection.rows[0]
            selected_name = view_df.iloc[selected_idx]['매물명']
            if st.session_state.edit_target != selected_name:
                st.session_state.edit_target = selected_name
                st.rerun()
    else:
        st.info("검색 결과가 없습니다.")
else:
    st.info("아직 등록된 매물이 없습니다. 위에서 첫 매물을 등록해보세요!")
