import streamlit as st
import gspread
import json
import base64
from oauth2client.service_account import ServiceAccountCredentials
import pandas as pd
from datetime import datetime
import re
import time

# --- 1. 앱 설정 ---
st.set_page_config(page_title="실험실 재고 관리기 v62", layout="wide")
st.title("🔬 실험실 재고 관리기 v62 (Fix: NameError)")

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

@st.cache_data(ttl=5)
def load_data_only(_client):
    try:
        # Master DB
        sh_db = _client.open(REAGENT_DB_NAME)
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

        # Log DB
        sh_log = _client.open(USAGE_LOG_NAME)
        ws_log = sh_log.worksheet(USAGE_LOG_TAB)
        data_log = ws_log.get_all_records()
        
        log_cols = ["일시", "구분", "제품명", "관리번호", "제조사 Lot", "수량", "유효기간", "담당자", "비고"]
        if not data_log:
            df_log = pd.DataFrame(columns=log_cols)
        else:
            df_log = pd.DataFrame(data_log)
            for col in log_cols:
                if col not in df_log.columns:
                    df_log[col] = ""
        
        return df_master, df_log, None
    except Exception as e:
        return pd.DataFrame(), pd.DataFrame(), str(e)

def get_worksheets(client):
    try:
        sh_db = client.open(REAGENT_DB_NAME)
        ws_db = sh_db.worksheet(REAGENT_DB_TAB)
        
        sh_log = client.open(USAGE_LOG_NAME)
        ws_log = sh_log.worksheet(USAGE_LOG_TAB)
        return ws_db, ws_log
    except Exception as e:
        st.error(f"시트 연결 실패: {e}")
        return None, None

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
    if not abbr: return "UNK-0000-00"
    today_ym = datetime.now().strftime("%y%m")
    prefix = f"{abbr}-{today_ym}"
    
    if df_log.empty:
        seq = 1
    else:
        if '관리번호' in df_log.columns:
            mask = df_log['관리번호'].astype(str).str.startswith(prefix)
            count = mask.sum()
            seq = count + 1
        else:
            seq = 1
    return f"{prefix}-{seq:02d}"

def update_master_abbr(client, product_name, new_abbr):
    try:
        sh = client.open(REAGENT_DB_NAME)
        ws = sh.worksheet(REAGENT_DB_TAB)
        cell = ws.find(product_name, in_column=1)
        if cell:
            ws.update_cell(cell.row, 2, new_abbr)
            return True
        return False
    except:
        return False

# --- 3. 앱 실행 ---
client, err = get_gspread_client()
if err: st.error(err); st.stop()

df_master, df_log, load_err = load_data_only(client)
if load_err:
    st.error(f"데이터 로딩 실패: {load_err}")
    st.stop()

tab1, tab2, tab3, tab4 = st.tabs(["📂 BOM 업로드", "📦 입고/사용", "📊 재고 현황", "⚙️ 품목/코드 관리"])

# ==============================================================================
# [Tab 1] BOM(품목) 관리
# ==============================================================================
with tab1:
    st.header("📂 BOM 마스터 데이터 관리")
    with st.expander("엑셀 BOM 업로드", expanded=True):
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
                ws_db, _ = get_worksheets(client)
                if ws_db:
                    try:
                        processed = []
                        header = ["제품명", "식별코드", "상세 특징", "Cat. No.", "규격(용량)", "단위", "제조사", "포장단위", "보관 위치", "알림 기준 수량", "등록일", "등록자"]
                        processed.append(header)
                        
                        for _, row in df_upload.iterrows():
                            p_name = str(row.get(COL_MAP["제품명"], "")).strip()
                            if not p_name: continue
                            
                            raw_abbr = str(row.get(COL_MAP["식별코드"], "")).strip()
                            final_abbr = raw_abbr.upper() if raw_abbr else make_smart_abbr(p_name)
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
                        ws_db.clear()
                        ws_db.update(processed)
                        st.success(f"✅ 기준 정보 등록 완료!")
                        st.cache_data.clear()
                    except Exception as e:
                        st.error(f"업로드 실패: {e}")

# ==============================================================================
# [Tab 2] 입고 및 사용 등록
# ==============================================================================
with tab2:
    st.header("📦 자재 수불 관리")
    
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
                new_abbr_input = c2.text_input("식별코드 (3글자, 예: DME)", max_chars=3)
                new_cat_no = c3.text_input("Cat. No.")
                
                c4, c5, c6 = st.columns(3)
                new_spec = c4.text_input("상세 특징")
                new_cap = c5.text_input("용량 (규격)")
                new_unit = c6.selectbox("단위", ["ea", "box", "ml", "L", "g", "kg", "kit"])
                
                # [버그 수정] 제조사(new_maker) 입력칸 복구!
                c7, c8 = st.columns(2)
                new_pkg = c7.text_input("포장단위 (예: 10ea/box)")
                new_maker = c8.text_input("제조사 (Maker)") 
                
                st.markdown("---")
                lc1, lc2, lc3 = st.columns(3)
                qty = lc1.number_input("입고 수량", min_value=1, step=1, format="%d")
                mfg_lot = lc2.text_input("제조사 Lot 번호 (필수)")
                expiry_input = lc3.date_input("유효기간").strftime("%Y-%m-%d")
                
                selected_product = new_p_name
                lot_to_save = "AUTO" 
            else: 
                if df_master.empty: st.stop()
                selected_product = st.selectbox("품목 선택", sorted(df_master['제품명'].unique()))
                mfg_lot = "" 
                
                if selected_product:
                    info = df_master[df_master['제품명'] == selected_product].iloc[0]
                    current_abbr_master = str(info.get('식별코드', '')).strip()
                    if not current_abbr_master: current_abbr_master = make_smart_abbr(selected_product)
                    
                    st.info(f"ℹ️ Spec: {info['상세 특징']} | Code: **{current_abbr_master}**")
                    auto_lot = generate_internal_lot(current_abbr_master, df_log)
                    st.success(f"🎫 생성된 관리번호: **{auto_lot}**")
                    lot_to_save = auto_lot
                
                st.markdown("---")
                lc1, lc2, lc3 = st.columns(3)
                qty = lc1.number_input("입고 수량", min_value=1, step=1, format="%d")
                mfg_lot = lc2.text_input("제조사 Lot 번호 (필수)", help="시약병에 적힌 번호")
                expiry_input = lc3.date_input("유효기간").strftime("%Y-%m-%d")

        # [CASE B] 사용
        else:
            if df_master.empty: st.stop()
            selected_product = st.selectbox("품목 선택", sorted(df_master['제품명'].unique()))
            mfg_lot = "-" 
            existing_lots = ["Initial"]
            lot_map = {} 

            if not df_log.empty and selected_product:
                log_in = df_log[(df_log['제품명'] == selected_product) & (df_log['구분'] == 'IN')]
                if not log_in.empty:
                    for _, row in log_in.iterrows():
                        internal = str(row.get('관리번호', ''))
                        mfg = str(row.get('제조사 Lot', ''))
                        if internal: lot_map[internal] = mfg
                    found = sorted(list(lot_map.keys()))
                    if found: existing_lots = found
            
            lc1, lc2 = st.columns(2)
            qty = lc1.number_input("사용 수량", min_value=1, step=1, format="%d")
            lot_to_save = lc2.selectbox("사용할 관리번호", existing_lots)
            
            matched_mfg_lot = lot_map.get(lot_to_save, "Unknown")
            if lot_to_save != "Initial":
                st.info(f"🔎 추적된 제조사 Lot: **[{matched_mfg_lot}]**")
                mfg_lot = matched_mfg_lot
            expiry_input = "-" 

        uc1, uc2 = st.columns(2)
        user = uc1.text_input("담당자", value="관리자")
        note = uc2.text_input("비고")
        
        if st.form_submit_button("저장하기"):
            ws_db, ws_log = get_worksheets(client)
            if not ws_log: st.stop()
            
            if not selected_product:
                st.error("제품명을 입력해주세요.")
            else:
                if "입고" in action_type and is_new_product:
                    final_abbr = new_abbr_input.upper() if new_abbr_input else make_smart_abbr(new_p_name)
                    # [버그 수정] new_maker 변수가 이제 존재하므로 에러 없음
                    new_row = [
                        new_p_name, final_abbr, new_spec, new_cat_no, new_cap, new_unit, 
                        new_maker, new_pkg, "-", 0, 
                        datetime.now().strftime("%Y-%m-%d"), user
                    ]
                    ws_db.append_row(new_row)
                    st.toast(f"✨ 신규 품목 등록 완료!")
                    lot_to_save = generate_internal_lot(final_abbr, df_log)

                final_qty = qty if "입고" in action_type else -qty
                action_code = "IN" if "입고" in action_type else "OUT"
                
                log_row = [
                    datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    action_code, selected_product, lot_to_save, mfg_lot,
                    final_qty, expiry_input, user, note
                ]
                ws_log.append_row(log_row)
                
                st.success(f"✅ 저장 완료! ({lot_to_save})")
                st.cache_data.clear()

# ==============================================================================
# [Tab 3] 실시간 재고 현황
# ==============================================================================
with tab3:
    st.header("📊 실시간 재고 현황")
    if st.button("🔄 새로고침"):
        st.cache_data.clear()
        st.rerun()
    
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
        
        st.subheader("📦 품목별 재고 현황 (Lot 추적)")
        if not df_log.empty:
            df_log['관리번호'] = df_log['관리번호'].replace("", "Old-Data")
            lot_stock = df_log.groupby(['제품명', '관리번호', '제조사 Lot'])['수량'].sum().reset_index()
            active_lots = lot_stock[lot_stock['수량'] > 0].sort_values('제품명')
            st.dataframe(active_lots, use_container_width=True)

# ==============================================================================
# [Tab 4] 품목/코드 관리
# ==============================================================================
with tab4:
    st.header("⚙️ 기준 정보 및 식별코드 관리")
    
    if df_master.empty:
        st.warning("등록된 품목이 없습니다.")
    else:
        target_product = st.selectbox("수정할 품목 선택", sorted(df_master['제품명'].unique()))
        
        if target_product:
            info = df_master[df_master['제품명'] == target_product].iloc[0]
            current_abbr = str(info.get('식별코드', ''))
            
            col1, col2 = st.columns(2)
            with col1:
                st.write(f"**현재 식별코드:** `{current_abbr}`")
                st.write(f"**현재 생성 예시:** `{current_abbr}-2512-01`")
            
            with col2:
                new_abbr_edit = st.text_input("새로운 식별코드 (3글자 영문)", value=current_abbr, max_chars=3)
                
                if st.button("💾 식별코드 변경사항 저장"):
                    if new_abbr_edit and new_abbr_edit != current_abbr:
                        success = update_master_abbr(client, target_product, new_abbr_edit.upper())
                        if success:
                            st.success(f"✅ 변경 완료! ({current_abbr} → {new_abbr_edit.upper()})")
                            st.cache_data.clear()
                        else:
                            st.error("DB 업데이트 실패")
                    else:
                        st.warning("변경 사항이 없거나 코드가 비어있습니다.")

