import streamlit as st
import gspread
import json
import base64
from oauth2client.service_account import ServiceAccountCredentials
import pandas as pd
from datetime import datetime

# --- 1. 앱 설정 ---
st.set_page_config(page_title="실험실 재고 관리기 v54", layout="wide")
st.title("🔬 실험실 재고 관리기 v54")

# --- 2. 구글 시트 연결 설정 ---
REAGENT_DB_NAME = "Reagent_DB"
REAGENT_DB_TAB = "Master"       
USAGE_LOG_NAME = "Usage_Log"    
USAGE_LOG_TAB = "Log"

@st.cache_resource
def get_gspread_client():
    try:
        scope = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
        if 'gcp_json_base64' in st.secrets:
            creds = ServiceAccountCredentials.from_json_keyfile_dict(
                json.loads(base64.b64decode(st.secrets["gcp_json_base64"]).decode("utf-8")), scope)
        else:
            creds = ServiceAccountCredentials.from_service_account_file('.streamlit/secrets.toml', scope)
        return gspread.authorize(creds), None
    except Exception as e:
        return None, f"인증 오류: {e}"

def load_data(client):
    # Master DB 로드
    sh_db = client.open(REAGENT_DB_NAME)
    ws_db = sh_db.worksheet(REAGENT_DB_TAB)
    data_db = ws_db.get_all_records()
    expected_cols = ["제품명", "상세 특징", "Cat. No.", "규격(용량)", "단위", "제조사", "포장단위", "보관 위치", "알림 기준 수량", "등록일", "등록자"]
    
    if not data_db:
        df_master = pd.DataFrame(columns=expected_cols)
    else:
        df_master = pd.DataFrame(data_db)
        for col in expected_cols:
            if col not in df_master.columns:
                df_master[col] = ""

    # Log 로드
    sh_log = client.open(USAGE_LOG_NAME)
    ws_log = sh_log.worksheet(USAGE_LOG_TAB)
    data_log = ws_log.get_all_records()
    df_log = pd.DataFrame(data_log)
    
    return df_master, df_log, ws_db, ws_log

# --- 3. 앱 실행 ---
client, err = get_gspread_client()
if err: st.error(err); st.stop()

tab1, tab2, tab3 = st.tabs(["📂 BOM(품목) 관리", "📦 입고/사용 등록", "📊 실시간 재고 현황"])

# ==============================================================================
# [Tab 1] BOM(품목) 관리
# ==============================================================================
with tab1:
    st.header("📂 BOM 마스터 데이터 관리")
    with st.expander("엑셀 BOM 업로드 (기준 정보 갱신)", expanded=True):
        uploaded_file = st.file_uploader("정리된 BOM 엑셀 파일 (.xlsx)", type=['xlsx'])
        if uploaded_file:
            df_upload = pd.read_excel(uploaded_file)
            df_upload.columns = df_upload.columns.str.strip()
            df_upload = df_upload.fillna("")
            
            COL_MAP = {
                "제품명": "품목명", "상세 특징": "상세 특징", "Cat. No.": "Cat. No.",
                "규격(용량)": "용량", "단위": "단위", "제조사": "제조사", 
                "포장단위": "포장", "보관 위치": "보관 장소", "알림 기준 수량": "안전재고"
            }
            
            if st.button("🚀 기준 정보 덮어쓰기"):
                try:
                    sh = client.open(REAGENT_DB_NAME)
                    ws = sh.worksheet(REAGENT_DB_TAB)
                    processed = []
                    header = ["제품명", "상세 특징", "Cat. No.", "규격(용량)", "단위", "제조사", "포장단위", "보관 위치", "알림 기준 수량", "등록일", "등록자"]
                    processed.append(header)
                    
                    for _, row in df_upload.iterrows():
                        p_name = str(row.get(COL_MAP["제품명"], "")).strip()
                        if not p_name: continue
                        
                        try: safe_stock = float(str(row.get(COL_MAP["알림 기준 수량"], 0)).replace("-","0").replace(",",""))
                        except: safe_stock = 0.0

                        processed.append([
                            p_name,
                            str(row.get(COL_MAP["상세 특징"], "-")),
                            str(row.get(COL_MAP["Cat. No."], "-")),
                            str(row.get(COL_MAP["규격(용량)"], "-")),
                            str(row.get(COL_MAP["단위"], "ea")),
                            str(row.get(COL_MAP["제조사"], "-")),
                            str(row.get(COL_MAP["포장단위"], "-")),
                            str(row.get(COL_MAP["보관 위치"], "-")),
                            safe_stock,
                            datetime.now().strftime("%Y-%m-%d"),
                            "관리자(일괄)"
                        ])
                    
                    ws.clear()
                    ws.update(processed)
                    st.success(f"✅ 기준 정보 {len(processed)-1}건 등록 완료!")
                    st.cache_data.clear()
                except Exception as e:
                    st.error(f"업로드 실패: {e}")

# ==============================================================================
# [Tab 2] 입고 및 사용 등록 (수정됨: 체크박스 위치 변경)
# ==============================================================================
with tab2:
    st.header("📦 자재 수불 관리")
    
    df_master, df_log, ws_db, ws_log = load_data(client)
    
    # 1. 작업 유형 선택
    col_type, col_check = st.columns([1, 2])
    action_type = col_type.radio("작업 유형", ["🔵 입고 (구매/채워넣기)", "🔴 사용 (소진/출고)"])
    
    # [핵심 수정] 폼 밖으로 체크박스를 꺼냈습니다. 이제 즉시 반응합니다!
    is_new_product = False
    if "입고" in action_type:
        is_new_product = col_check.checkbox("🆕 신규 품목 등록 (목록에 없으면 체크하세요)")
    
    st.divider()

    # --- 입력 폼 로직 시작 ---
    with st.form("action_form", clear_on_submit=True):
        
        # [CASE A] 입고 작업일 때
        if "입고" in action_type:
            
            if is_new_product:
                # 여기가 신규 등록 화면
                st.markdown("##### 📝 신규 품목 정보 입력")
                c1, c2, c3 = st.columns(3)
                new_p_name = c1.text_input("제품명 (필수)*")
                new_cat_no = c2.text_input("Cat. No.")
                new_maker = c3.text_input("제조사")
                
                c4, c5, c6 = st.columns(3)
                new_spec = c4.text_input("상세 특징 (Spec)")
                new_cap = c5.text_input("용량 (규격)")
                new_unit = c6.selectbox("단위", ["ea", "box", "ml", "L", "g", "kg", "kit"])
                
                c7, c8 = st.columns(2)
                new_pkg = c7.text_input("포장단위 (예: 10ea/box)")
                new_alert = c8.number_input("안전재고(알림 기준)", value=5, step=1)
                
                st.markdown("---")
                st.markdown("##### 🔽 입고 수량 입력")
                lc1, lc2, lc3 = st.columns(3)
                qty = lc1.number_input("입고 수량 (정수)", min_value=1, step=1, format="%d")
                lot_input = lc2.text_input("Lot 번호", value=datetime.now().strftime("%Y%m%d"))
                expiry_input = lc3.date_input("유효기간").strftime("%Y-%m-%d")
                
                selected_product = new_p_name # 로직 연결용
                
            else: # 기존 품목 입고 화면
                if df_master.empty:
                    st.warning("등록된 품목이 없습니다. 위 체크박스를 눌러 신규 등록하세요.")
                    st.stop()
                    
                selected_product = st.selectbox("품목 선택", sorted(df_master['제품명'].unique()))
                # 품목 정보 표시
                if selected_product:
                    info = df_master[df_master['제품명'] == selected_product].iloc[0]
                    st.info(f"ℹ️ 선택됨: **{selected_product}** (Spec: {info['상세 특징']} | Cat: {info['Cat. No.']})")
                
                lc1, lc2, lc3 = st.columns(3)
                qty = lc1.number_input("입고 수량 (정수)", min_value=1, step=1, format="%d")
                lot_input = lc2.text_input("Lot 번호", value=datetime.now().strftime("%Y%m%d"))
                expiry_input = lc3.date_input("유효기간").strftime("%Y-%m-%d")

        # [CASE B] 사용(출고) 작업일 때
        else:
            if df_master.empty:
                st.warning("등록된 품목이 없습니다.")
                st.stop()
            
            selected_product = st.selectbox("품목 선택", sorted(df_master['제품명'].unique()))
            
            # Lot 선택 로직
            existing_lots = ["Initial"]
            if not df_log.empty and selected_product:
                log_filtered = df_log[df_log['제품명'] == selected_product]
                if not log_filtered.empty:
                    found = log_filtered['Lot 번호'].unique().tolist()
                    if found: existing_lots = sorted(found)
            
            lc1, lc2 = st.columns(2)
            qty = lc1.number_input("사용 수량 (정수)", min_value=1, step=1, format="%d")
            lot_input = lc2.selectbox("Lot 번호 (사용 제품)", existing_lots)
            expiry_input = "-" 

        # 공통 입력 (담당자, 비고)
        uc1, uc2 = st.columns(2)
        user = uc1.text_input("담당자", value="관리자")
        note = uc2.text_input("비고")
        
        # --- 저장 버튼 ---
        if st.form_submit_button("저장하기"):
            if not selected_product:
                st.error("제품명을 입력해주세요.")
            else:
                # 1. 신규 품목이면 마스터 DB에 먼저 등록
                if "입고" in action_type and is_new_product:
                    new_row = [
                        new_p_name, new_spec, new_cat_no, new_cap, new_unit, 
                        new_maker, new_pkg, "-", new_alert, 
                        datetime.now().strftime("%Y-%m-%d"), user
                    ]
                    ws_db.append_row(new_row)
                    st.toast(f"✨ 신규 품목 '{new_p_name}' 등록 완료!")

                # 2. 로그 저장
                final_qty = qty if "입고" in action_type else -qty
                action_code = "IN" if "입고" in action_type else "OUT"
                
                log_row = [
                    datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    action_code, selected_product, lot_input, final_qty,
                    expiry_input, user, note
                ]
                ws_log.append_row(log_row)
                
                st.success(f"✅ {selected_product} 처리 완료!")
                st.cache_data.clear()

# ==============================================================================
# [Tab 3] 실시간 재고 현황
# ==============================================================================
with tab3:
    st.header("📊 실시간 재고 현황")
    if st.button("🔄 새로고침"): st.rerun()
    
    df_master, df_log, _, _ = load_data(client)
    
    if not df_master.empty:
        if df_log.empty:
            df_stock = df_master.copy()
            df_stock['현재고'] = 0
        else:
            df_log['수량'] = pd.to_numeric(df_log['수량'], errors='coerce').fillna(0)
            stock_grp = df_log.groupby('제품명')['수량'].sum().reset_index()
            stock_grp.rename(columns={'수량': '현재고'}, inplace=True)
            df_stock = pd.merge(df_master, stock_grp, on='제품명', how='left')
            df_stock['현재고'] = df_stock['현재고'].fillna(0).astype(int)

        try: df_stock['알림 기준 수량'] = pd.to_numeric(df_stock['알림 기준 수량'], errors='coerce').fillna(0)
        except: pass
        
        low_stock = df_stock[df_stock['현재고'] <= df_stock['알림 기준 수량']]
        
        if not low_stock.empty:
            st.error(f"🚨 재고 부족 ({len(low_stock)}건)")
            st.dataframe(low_stock[['제품명', '현재고', '알림 기준 수량', '보관 위치']], hide_index=True)
        
        st.subheader("📦 전체 재고 리스트")
        disp_cols = ["제품명", "현재고", "단위", "규격(용량)", "제조사", "Cat. No.", "상세 특징", "보관 위치"]
        valid_cols = [c for c in disp_cols if c in df_stock.columns]
        st.dataframe(df_stock[valid_cols], use_container_width=True, hide_index=True)
