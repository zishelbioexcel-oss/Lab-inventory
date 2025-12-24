import streamlit as st
import gspread
import json
import base64
from oauth2client.service_account import ServiceAccountCredentials
import pandas as pd
from datetime import datetime
import re

# --- 1. 앱 설정 ---
st.set_page_config(page_title="실험실 재고 관리기 v57", layout="wide")
st.title("🔬 실험실 재고 관리기 v57 (Dual Lot Tracking)")

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
    sh_db = client.open(REAGENT_DB_NAME)
    ws_db = sh_db.worksheet(REAGENT_DB_TAB)
    data_db = ws_db.get_all_records()
    expected_cols = ["제품명", "식별코드", "상세 특징", "Cat. No.", "규격(용량)", "단위", "제조사", "포장단위", "보관 위치", "알림 기준 수량", "등록일", "등록자"]
    
    if not data_db:
        df_master = pd.DataFrame(columns=expected_cols)
    else:
        df_master = pd.DataFrame(data_db)
        for col in expected_cols:
            if col not in df_master.columns:
                df_master[col] = ""

    sh_log = client.open(USAGE_LOG_NAME)
    ws_log = sh_log.worksheet(USAGE_LOG_TAB)
    data_log = ws_log.get_all_records()
    df_log = pd.DataFrame(data_log)
    
    return df_master, df_log, ws_db, ws_log

def make_smart_abbr(name):
    if not name: return "UNK"
    clean_name = re.sub(r'[^\w\s]', '', str(name)).upper()
    words = clean_name.split()
    if len(words) >= 3:
        return (words[0][0] + words[1][0] + words[2][0])
    elif len(words) == 2:
        return (words[0][0] + words[1][0] + words[1][-1]) 
    else:
        return clean_name[:3]

def generate_internal_lot(abbr, df_log):
    if not abbr: return "UNKNOWN-0000-00"
    today_ym = datetime.now().strftime("%y%m")
    prefix = f"{abbr}-{today_ym}"
    
    if df_log.empty:
        seq = 1
    else:
        mask = df_log['Lot 번호'].astype(str).str.startswith(prefix)
        count = mask.sum()
        seq = count + 1
        
    return f"{prefix}-{seq:02d}"

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
                "제품명": "품목명", "식별코드": "식별코드", "상세 특징": "상세 특징", 
                "Cat. No.": "Cat. No.", "규격(용량)": "용량", "단위": "단위", 
                "제조사": "제조사", "포장단위": "포장", "보관 위치": "보관 장소", 
                "알림 기준 수량": "안전재고"
            }
            
            if st.button("🚀 기준 정보 덮어쓰기"):
                try:
                    sh = client.open(REAGENT_DB_NAME)
                    ws = sh.worksheet(REAGENT_DB_TAB)
                    processed = []
                    header = ["제품명", "식별코드", "상세 특징", "Cat. No.", "규격(용량)", "단위", "제조사", "포장단위", "보관 위치", "알림 기준 수량", "등록일", "등록자"]
                    processed.append(header)
                    
                    for _, row in df_upload.iterrows():
                        p_name = str(row.get(COL_MAP["제품명"], "")).strip()
                        if not p_name: continue
                        
                        raw_abbr = str(row.get(COL_MAP["식별코드"], "")).strip()
                        if not raw_abbr: final_abbr = make_smart_abbr(p_name)
                        else: final_abbr = raw_abbr.upper()

                        try: safe_stock = float(str(row.get(COL_MAP["알림 기준 수량"], 0)).replace("-","0").replace(",",""))
                        except: safe_stock = 0.0

                        processed.append([
                            p_name, final_abbr,
                            str(row.get(COL_MAP["상세 특징"], "-")), str(row.get(COL_MAP["Cat. No."], "-")),
                            str(row.get(COL_MAP["규격(용량)"], "-")), str(row.get(COL_MAP["단위"], "ea")),
                            str(row.get(COL_MAP["제조사"], "-")), str(row.get(COL_MAP["포장단위"], "-")),
                            str(row.get(COL_MAP["보관 위치"], "-")), safe_stock,
                            datetime.now().strftime("%Y-%m-%d"), "관리자(일괄)"
                        ])
                    ws.clear()
                    ws.update(processed)
                    st.success(f"✅ 기준 정보 등록 완료!")
                    st.cache_data.clear()
                except Exception as e:
                    st.error(f"업로드 실패: {e}")

# ==============================================================================
# [Tab 2] 입고 및 사용 등록
# ==============================================================================
with tab2:
    st.header("📦 자재 수불 관리")
    
    df_master, df_log, ws_db, ws_log = load_data(client)
    
    col_type, col_check = st.columns([1, 2])
    action_type = col_type.radio("작업 유형", ["🔵 입고 (구매/채워넣기)", "🔴 사용 (소진/출고)"])
    
    is_new_product = False
    if "입고" in action_type:
        is_new_product = col_check.checkbox("🆕 신규 품목 등록")
    
    st.divider()

    with st.form("action_form", clear_on_submit=True):
        
        # [CASE A] 입고
        if "입고" in action_type:
            
            if is_new_product:
                st.markdown("##### 📝 신규 품목 정보")
                c1, c2, c3 = st.columns(3)
                new_p_name = c1.text_input("제품명 (필수)*")
                new_abbr_input = c2.text_input("식별코드 (3글자, 내부용)", max_chars=3)
                new_cat_no = c3.text_input("Cat. No.")
                
                c4, c5, c6 = st.columns(3)
                new_spec = c4.text_input("상세 특징")
                new_cap = c5.text_input("용량 (규격)")
                new_unit = c6.selectbox("단위", ["ea", "box", "ml", "L", "g", "kg", "kit"])
                new_pkg = st.text_input("포장단위")
                
                st.markdown("---")
                st.markdown("##### 🔽 입고 세부 정보")
                lc1, lc2, lc3 = st.columns(3)
                qty = lc1.number_input("입고 수량", min_value=1, step=1, format="%d")
                
                # [수정] 제조사 Lot 입력칸 부활
                mfg_lot = lc2.text_input("제조사 Lot 번호 (실제 병에 적힌 것)")
                expiry_input = lc3.date_input("유효기간").strftime("%Y-%m-%d")
                
                selected_product = new_p_name
                lot_to_save = "AUTO" 

            else: # 기존 품목 입고
                if df_master.empty: st.stop()
                selected_product = st.selectbox("품목 선택", sorted(df_master['제품명'].unique()))
                
                mfg_lot = "" # 초기화
                
                if selected_product:
                    info = df_master[df_master['제품명'] == selected_product].iloc[0]
                    current_abbr = str(info.get('식별코드', ''))
                    if not current_abbr: current_abbr = make_smart_abbr(selected_product)
                    
                    st.info(f"ℹ️ Spec: **{info['상세 특징']}** | Cat: **{info['Cat. No.']}**")
                    
                    # 1. 식별코드(내부용) 확인/수정
                    c_code1, c_code2 = st.columns([1, 2])
                    final_abbr = c_code1.text_input("식별코드 (내부 번호 생성용)", value=current_abbr, max_chars=3)
                    
                    # 내부 번호 미리보기
                    auto_lot = generate_internal_lot(final_abbr, df_log)
                    c_code2.success(f"🎫 생성될 내부 관리번호: **{auto_lot}** (라벨에 부착)")
                    
                    lot_to_save = auto_lot
                
                st.markdown("---")
                lc1, lc2, lc3 = st.columns(3)
                qty = lc1.number_input("입고 수량", min_value=1, step=1, format="%d")
                
                # [수정] 제조사 Lot 입력칸 부활
                mfg_lot = lc2.text_input("제조사 Lot 번호 (필수 입력)", help="시약병에 적혀있는 Lot No.를 적으세요")
                
                expiry_input = lc3.date_input("유효기간").strftime("%Y-%m-%d")

        # [CASE B] 사용
        else:
            if df_master.empty: st.stop()
            selected_product = st.selectbox("품목 선택", sorted(df_master['제품명'].unique()))
            mfg_lot = "-" # 사용시는 제조사 Lot 입력 불필요 (내부 번호로 찾음)
            
            existing_lots = ["Initial"]
            if not df_log.empty and selected_product:
                log_filtered = df_log[df_log['제품명'] == selected_product]
                if not log_filtered.empty:
                    found = log_filtered['Lot 번호'].unique().tolist()
                    if found: existing_lots = sorted(found)
            
            lc1, lc2 = st.columns(2)
            qty = lc1.number_input("사용 수량", min_value=1, step=1, format="%d")
            lot_to_save = lc2.selectbox("Lot 번호 (내부 관리번호)", existing_lots)
            expiry_input = "-" 

        uc1, uc2 = st.columns(2)
        user = uc1.text_input("담당자", value="관리자")
        note = uc2.text_input("비고 (실험명 등)")
        
        if st.form_submit_button("저장하기"):
            if not selected_product:
                st.error("제품명을 입력해주세요.")
            else:
                # 1. 신규 품목 등록
                if "입고" in action_type and is_new_product:
                    final_abbr = new_abbr_input.upper() if new_abbr_input else make_smart_abbr(new_p_name)
                    new_row = [
                        new_p_name, final_abbr, new_spec, new_cat_no, new_cap, new_unit, 
                        new_maker, new_pkg, "-", 0, 
                        datetime.now().strftime("%Y-%m-%d"), user
                    ]
                    ws_db.append_row(new_row)
                    st.toast(f"✨ 신규 품목 등록 완료!")
                    lot_to_save = generate_internal_lot(final_abbr, df_log)

                # 2. 비고란에 제조사 Lot 정보 합치기
                final_note = note
                if "입고" in action_type and mfg_lot:
                    final_note = f"(Mfg: {mfg_lot}) {note}"

                # 3. 로그 저장
                final_qty = qty if "입고" in action_type else -qty
                action_code = "IN" if "입고" in action_type else "OUT"
                
                log_row = [
                    datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    action_code, selected_product, lot_to_save, final_qty,
                    expiry_input, user, final_note
                ]
                ws_log.append_row(log_row)
                
                st.success(f"✅ {selected_product} 입고 완료! (내부번호: {lot_to_save})")
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
        disp_cols = ["제품명", "식별코드", "현재고", "단위", "규격(용량)", "제조사", "Cat. No.", "보관 위치"]
        valid_cols = [c for c in disp_cols if c in df_stock.columns]
        st.dataframe(df_stock[valid_cols], use_container_width=True, hide_index=True)
