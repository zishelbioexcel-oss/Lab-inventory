import streamlit as st
import gspread 
import json 
import base64 
from oauth2client.service_account import ServiceAccountCredentials 
import pandas as pd 
from datetime import datetime

# --- 1. 앱의 기본 설정 ---
st.set_page_config(page_title="실험실 재고 관리기 v50", layout="wide")
st.title("🔬 실험실 재고 관리기 v50")
st.write("새 품목을 등록하고, 사용량을 기록하며, 재고 현황을 확인합니다.")

# --- 2. Google Sheets 인증 및 설정 ---
# (v49와 동일)
REAGENT_DB_NAME = "Reagent_DB"  
REAGENT_DB_TAB = "Master"       
USAGE_LOG_NAME = "Usage_Log"    
USAGE_LOG_TAB = "Log"           

# (1) 인증된 '클라이언트' 생성 (v49와 동일)
@st.cache_resource(ttl=600)
def get_gspread_client():
    try:
        scope = [
            'https://www.googleapis.com/auth/spreadsheets',
            'https://www.googleapis.com/auth/drive'
        ]
        if 'gcp_json_base64' in st.secrets:
            base64_string = st.secrets["gcp_json_base64"]
            json_string = base64.b64decode(base64_string).decode("utf-8")
            creds_dict = json.loads(json_string) 
            creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
        else:
            creds = ServiceAccountCredentials.from_service_account_file('.streamlit/secrets.toml', scope)
        client = gspread.authorize(creds)
        return client, None
    except FileNotFoundError:
        return None, "로컬 Secrets 파일('.streamlit/secrets.toml')을 찾을 수 없습니다."
    except Exception as e:
        return None, f"Google 인증 실패: {e}"

# (2) 마스터 DB 로드 함수 (v49와 동일)
@st.cache_data(ttl=60) 
def load_reagent_db(_client):
    try:
        sh = _client.open(REAGENT_DB_NAME)
        sheet = sh.worksheet(REAGENT_DB_TAB)
        data = sheet.get_all_records()
        if not data:
            st.warning("마스터 시트(Reagent_DB)가 비어있습니다...")
            return pd.DataFrame(columns=["제품명", "제조사", "Cat. No.", "Lot 번호", "최초 수량", "단위", "유통기한", "알림 기준 수량", "알림 무시"])
        
        df = pd.DataFrame(data)
        
        required_cols = ["제품명", "제조사", "Cat. No.", "Lot 번호", "최초 수량", "단위", "유통기한", "보관 위치", "등록 날짜", "등록자", "알림 기준 수량", "알림 무시"]
        if not all(col in df.columns for col in required_cols):
             st.error(f"Reagent_DB 'Master' 탭에 {required_cols} 컬럼이 모두 필요합니다. (A~L열 순서 확인)")
             return pd.DataFrame(columns=required_cols)
        
        df['제품명'] = df['제품명'].astype(str)
        df['제조사'] = df['제조사'].astype(str) 
        df['Cat. No.'] = df['Cat. No.'].astype(str)
        df['Lot 번호'] = df['Lot 번호'].astype(str)
        df['최초 수량'] = pd.to_numeric(df['최초 수량'], errors='coerce').fillna(0)
        df['알림 기준 수량'] = pd.to_numeric(df['알림 기준 수량'], errors='coerce').fillna(0) 
        df['유통기한'] = pd.to_datetime(df['유통기한'], errors='coerce') 
        df['단위'] = df['단위'].astype(str)
        df['보관 위치'] = df['보관 위치'].astype(str)
        df['등록 날짜'] = pd.to_datetime(df['등록 날짜'], errors='coerce') 
        df['등록자'] = df['등록자'].astype(str)
        df['알림 무시'] = df['알림 무시'].astype(str).fillna("아니요") 
        
        df = df.sort_values(by='등록 날짜')
        
        df_agg = df.groupby(['제품명', 'Cat. No.', 'Lot 번호'], as_index=False).agg(
            agg_qty=('최초 수량', 'sum'),       
            agg_alert_qty=('알림 기준 수량', 'last'), 
            agg_unit=('단위', 'last'),          
            agg_location=('보관 위치', 'last'),     
            agg_expiry=('유통기한', 'last'),   
            agg_reg_date=('등록 날짜', 'last'),   
            agg_registrant=('등록자', 'last'),
            agg_mute=('알림 무시', 'last'),
            agg_manufacturer=('제조사', 'last') 
        )
        
        df_agg = df_agg.rename(columns={
            'agg_qty': '최초 수량',
            'agg_alert_qty': '알림 기준 수량', 
            'agg_unit': '단위',
            'agg_location': '보관 위치',
            'agg_expiry': '유통기한',
            'agg_reg_date': '등록 날짜',
            'agg_registrant': '등록자',
            'agg_mute': '알림 무시',
            'agg_manufacturer': '제조사' 
        })
        
        df_agg['등록 날짜'] = df_agg['등록 날짜'].dt.strftime('%Y-%m-%d %H:%M:%S')
             
        return df_agg 
    
    except Exception as e:
        st.error(f"Reagent_DB 로드 실패: {e}")
        return pd.DataFrame(columns=["제품명", "제조사", "Cat. No.", "Lot 번호", "최초 수량", "단위", "유통기한", "알림 기준 수량", "알림 무시"])

# (3) 사용 기록(Log) 로드 함수 (v49와 동일)
@st.cache_data(ttl=60)
def load_usage_log(_client):
    try:
        sh = _client.open(USAGE_LOG_NAME)
        sheet = sh.worksheet(USAGE_LOG_TAB)
        data = sheet.get_all_records()
        if not data:
            return pd.DataFrame(columns=["제품명", "Lot 번호", "사용량", "Timestamp"]) 
        
        df = pd.DataFrame(data)
        
        required_cols = ["제품명", "Lot 번호", "사용량", "Timestamp", "사용자", "비고"]
        if not all(col in df.columns for col in required_cols):
             st.error("Usage_Log 'Log' 탭에 '제품명', 'Lot 번호', '사용량' 컬럼이 없습니다. (1행 헤더 확인)")
             return pd.DataFrame(columns=required_cols)
        
        df['제품명'] = df['제품명'].astype(str)
        df['Lot 번호'] = df['Lot 번호'].astype(str)
        df['사용량'] = pd.to_numeric(df['사용량'], errors='coerce').fillna(0)
        df['Timestamp'] = pd.to_datetime(df['Timestamp'], errors='coerce') 
             
        return df
    except Exception as e:
        st.error(f"Usage_Log 로드 실패: {e}")
        return pd.DataFrame(columns=["제품명", "Lot 번호", "사용량", "Timestamp"])

# --- 3. 앱 실행 ---
client, auth_error_msg = get_gspread_client()

if auth_error_msg:
    st.error(auth_error_msg)
    st.warning("Secrets 설정, API 권한, 봇 초대를 확인하세요.")
    st.stop() 

tab1, tab2, tab3 = st.tabs(["📝 새 품목 등록", "📉 시약 사용", "📊 대시보드 (재고 현황)"])


# --- 4. 탭 1: 새 품목 등록 (v49와 동일) ---
with tab1:
    st.header("📝 새 시약/소모품 등록")
    # ... (v49 탭1 코드 전체 생략 - 동일) ...
    st.write(f"이 폼을 제출하면 **'{REAGENT_DB_NAME}'** 시트의 **'{REAGENT_DB_TAB}'** 탭에 저장됩니다.")
    df_db_copy = load_reagent_db(client) 
    copied_data = {}
    unit_options = ["개", "box", "kit", "mL", "L", "g", "kg"]
    if not df_db_copy.empty:
        if st.checkbox("🖨️ 기존 품목 정보 복사하기 (Cat.No., 제조사, 단위, 위치, 알림 기준)"): 
            all_products = sorted(df_db_copy['제품명'].dropna().unique())
            if 'product_to_copy' not in st.session_state:
                st.session_state.product_to_copy = all_products[0]
            selected_product_to_copy = st.selectbox(
                "복사할 제품명 선택:", 
                options=all_products, 
                key="product_to_copy"
            )
            if selected_product_to_copy:
                item_info = df_db_copy[
                    df_db_copy['제품명'] == selected_product_to_copy
                ].iloc[-1] 
                copied_data['product_name'] = item_info.get('제품명', '')
                copied_data['cat_no'] = item_info.get('Cat. No.', '')
                copied_data['manufacturer'] = item_info.get('제조사', '') 
                copied_data['unit'] = item_info.get('단위', '개')
                copied_data['location'] = item_info.get('보관 위치', '')
                copied_data['alert_qty'] = item_info.get('알림 기준 수량', 10) 
    st.divider()
    with st.form(key="new_item_form", clear_on_submit=True): 
        col1, col2 = st.columns(2)
        with col1:
            st.write("**필수 정보**")
            product_name = st.text_input("제품명*", value=copied_data.get('product_name', ''), help="예: DMEM, 10% FBS")
            manufacturer = st.text_input("제조사*", 
                                         value=copied_data.get('manufacturer', ''), 
                                         help="예: Thermo Fisher, Gibco, Merck")
            cat_no = st.text_input("Cat. No.*", value=copied_data.get('cat_no', ''), help="카탈로그 번호 (예: 11995-065)")
            lot_no = st.text_input("Lot 번호*", help="새로 등록할 Lot 번호를 입력하세요.")
        with col2:
            st.write("**수량 및 알림**")
            initial_qty = st.number_input("최초 수량*", min_value=0.0, step=1.0, format="%.2f")
            unit_index = unit_options.index(copied_data.get('unit')) if copied_data.get('unit') in unit_options else 0
            unit = st.selectbox("단위*", options=unit_options, index=unit_index) 
            alert_qty = st.number_input(
                "알림 기준 수량*", 
                min_value=0.0, 
                value=copied_data.get('alert_qty', 10.0), 
                step=1.0, 
                format="%.2f",
                help="이 수량 '이하'로 재고가 남으면 알림이 뜹니다."
            )
        st.divider()
        st.write("**기타 정보**")
        location = st.text_input("보관 위치", value=copied_data.get('location', ''), help="예: 4도 냉장고 A-1 선반...")
        expiry_date = st.date_input("유통기한", datetime.now() + pd.DateOffset(years=1))
        registrant = st.text_input("등록자 이름*")
        submit_button = st.form_submit_button(label="✅ 신규 등록하기")
    if "form1_status" in st.session_state:
        if st.session_state.form1_status == "success": st.success(st.session_state.form1_message)
        else: st.error(st.session_state.form1_message)
        del st.session_state.form1_status
        del st.session_state.form1_message
    if submit_button:
        if not all([product_name, cat_no, lot_no, manufacturer, initial_qty > 0, registrant, alert_qty >= 0]):
            st.session_state.form1_status = "error"
            st.session_state.form1_message = "필수 항목(*)을 모두 입력해야 합니다. (최초 수량 > 0, 알림 기준 >= 0)"
        else:
            try:
                sh = client.open(REAGENT_DB_NAME)
                sheet = sh.worksheet(REAGENT_DB_TAB)
                log_data_list = [
                    product_name,   # A
                    manufacturer,   # B
                    cat_no,         # C
                    lot_no,         # D
                    float(initial_qty), # E
                    unit,           # F
                    expiry_date.strftime("%Y-%m-%d"), # G
                    location,       # H
                    datetime.now().strftime("%Y-%m-%d %H:%M:%S"), # I
                    registrant,     # J
                    float(alert_qty), # K
                    "아니요"         # L
                ]
                sheet.append_row(log_data_list)
                st.session_state.form1_status = "success"
                st.session_state.form1_message = f"✅ **{product_name} (Lot: {lot_no})**가 마스터 시트에 성공적으로 등록되었습니다!"
                st.cache_data.clear() 
            except Exception as e:
                st.session_state.form1_status = "error"
                st.session_state.form1_message = f"Google Sheet 저장 실패: {e}"
        st.rerun()


# --- 5. 탭 2: 시약 사용 (v49와 동일) ---
with tab2:
    st.header("📉 시약 사용 기록")
    # ... (v49 탭2 코드 전체 생략 - 동일) ...
    st.write(f"이 폼을 제출하면 **'{USAGE_LOG_NAME}'** 시트의 **'{USAGE_LOG_TAB}'** 탭에 저장됩니다.")
    st.divider()
    df_db = load_reagent_db(client) 
    df_log = load_usage_log(client) 
    if df_db.empty:
        st.error("마스터 DB(Reagent_DB)에 등록된 품목이 없습니다. '새 품목 등록' 탭에서 먼저 품목을 등록하세요.")
    else:
        st.subheader("1. 사용할 품목 선택")
        all_products = sorted(df_db['제품명'].dropna().unique())
        selected_product = st.selectbox("사용한 제품명*", options=all_products)
        if selected_product:
            available_lots = sorted(
                df_db[df_db['제품명'] == selected_product]['Lot 번호'].dropna().unique()
            )
            selected_lot = st.selectbox("Lot 번호*", options=available_lots)
        else:
            selected_lot = st.selectbox("Lot 번호*", options=["제품명을 먼저 선택하세요"])
        current_stock = 0.0 
        unit = ""
        alert_level = 0.0 
        if selected_product and selected_lot:
            try:
                item_info = df_db[
                    (df_db['제품명'] == selected_product) & 
                    (df_db['Lot 번호'] == selected_lot)
                ].iloc[0] 
                initial_stock = item_info['최초 수량'] 
                unit = item_info['단위']
                alert_level = item_info['알림 기준 수량'] 
                usage_df = df_log[
                    (df_log['제품명'] == selected_product) & 
                    (df_log['Lot 번호'] == selected_lot)
                ]
                total_usage = usage_df['사용량'].sum()
                current_stock = initial_stock - total_usage
                st.info(f"**현재 남은 재고:** {current_stock:.2f} {unit} (총 입고: {initial_stock:.2f} {unit} / 알림 기준: {alert_level:.2f} {unit})")
            except (IndexError, TypeError, KeyError):
                st.warning("재고를 계산할 수 없습니다. (마스터DB/로그 확인)")
        st.divider()
        st.subheader("2. 사용 정보 입력")
        if "usage_qty_input" not in st.session_state:
            st.session_state.usage_qty_input = 0.0
        if "usage_user" not in st.session_state:
            st.session_state.usage_user = ""
        if "usage_notes" not in st.session_state:
            st.session_state.usage_notes = ""
        def submit_usage_callback(product, lot, qty, user, notes, date, stock, unit_str):
            if not all([product, lot, qty > 0, user]):
                st.session_state.form2_status = "error"
                st.session_state.form2_message = "필수 항목(*)을 모두 입력해야 합니다. (사용량은 0보다 커야 함)"
            elif float(qty) > stock:
                shortage = float(qty) - stock
                st.session_state.form2_status = "error"
                st.session_state.form2_message = f"⚠️ 재고 부족! 현재 재고({stock:.2f} {unit_str})보다 {shortage:.2f} {unit_str} 만큼 더 많이 입력했습니다."
            else:
                try:
                    sh_log = client.open(USAGE_LOG_NAME)
                    sheet_log = sh_log.worksheet(USAGE_LOG_TAB)
                    log_timestamp = datetime.combine(date, datetime.now().time())
                    log_data_list = [
                        log_timestamp.strftime("%Y-%m-%d %H:%M:%S"), 
                        str(product), 
                        str(lot),     
                        float(qty),      
                        user,
                        notes
                    ]
                    sheet_log.append_row(log_data_list)
                    st.session_state.form2_status = "success"
                    st.session_state.form2_message = f"✅ **{product} (Lot: {lot})** 사용 기록이 저장되었습니다!"
                    st.cache_data.clear() 
                    st.session_state.usage_qty_input = 0.0
                except Exception as e:
                    st.session_state.form2_status = "error"
                    st.session_state.form2_message = f"Google Sheet 저장 실패: {e}"
        with st.form(key="usage_form"):
            usage_qty = st.number_input("사용한 양*", min_value=0.0, step=1.0, format="%.2f", key="usage_qty_input")
            user = st.text_input("사용자 이름*", key="usage_user") 
            usage_date = st.date_input("사용 일자", value=datetime.now().date())
            notes = st.text_area("비고 (실험명 등)", key="usage_notes")
            submit_usage_button = st.form_submit_button(
                label="📉 사용 기록하기",
                on_click=submit_usage_callback,
                args=(
                    selected_product,
                    selected_lot,
                    st.session_state.usage_qty_input,
                    st.session_state.usage_user,
                    st.session_state.usage_notes,
                    usage_date, 
                    current_stock,
                    unit
                )
            )
        if "form2_status" in st.session_state:
            if st.session_state.form2_status == "success": st.success(st.session_state.form2_message)
            else: st.error(st.session_state.form2_message)
            del st.session_state.form2_status
            del st.session_state.form2_message


# --- 6. 탭 3: 대시보드 (재고 현황) (v50 수정됨) ---
with tab3:
    st.header("📊 대시보드 (재고 현황)")

    if st.button("새로고침 (Refresh Data)"):
        st.cache_data.clear() 
        st.rerun()

    # 1. 데이터 로드 (v49와 동일)
    df_db = load_reagent_db(client)
    df_log = load_usage_log(client)

    if df_db.empty:
        st.warning("마스터 DB(Reagent_DB)에 등록된 품목이 없습니다.")
    else:
        # 2. 총 사용량 계산 (v49와 동일)
        if not df_log.empty:
            usage_summary = df_log.groupby(['제품명', 'Lot 번호'])['사용량'].sum().reset_index()
            usage_summary = usage_summary.rename(columns={'사용량': '총 사용량'})
            df_inventory = pd.merge(df_db, usage_summary, on=['제품명', 'Lot 번호'], how='left')
            df_inventory['총 사용량'] = df_inventory['총 사용량'].fillna(0) 
        else:
            df_inventory = df_db.copy()
            df_inventory['총 사용량'] = 0.0

        # (v49 방식: 컬럼 통합)
        df_inventory['현재 재고'] = df_inventory['최초 수량'] - df_inventory['총 사용량']
        df_inventory['재고 비율 (%)'] = df_inventory.apply(
            lambda row: (row['현재 재고'] / row['최초 수량']) * 100 if row['최초 수량'] > 0 else 0,
            axis=1
        )
        df_inventory['재고 비율 (%)'] = df_inventory['재고 비율 (%)'].clip(0, 100)
        
        # 5. 자동 알림 (v49와 동일)
        st.subheader("🚨 자동 알림")
        expiry_threshold_days = 30
        today = pd.to_datetime(datetime.now().date()) 
        df_inventory['유통기한'] = df_inventory['유통기한'].fillna(pd.NaT) 
        
        expiring_soon = df_inventory[
            (df_inventory['유통기한'] >= today) &
            (df_inventory['유통기한'] <= (today + pd.DateOffset(days=expiry_threshold_days))) &
            (df_inventory['현재 재고'] > 0) &
            (df_inventory['알림 무시'] != "예") 
        ]
        expired = df_inventory[
            (df_inventory['유통기한'] < today) &
            (df_inventory['현재 재고'] > 0) &
            (df_inventory['알림 무시'] != "예") 
        ]
        if not expiring_soon.empty:
            st.warning(f"**유통기한 {expiry_threshold_days}일 이내 임박** (재고 있음)")
            expiring_display = expiring_soon.copy()
            expiring_display['유통기한'] = expiring_display['유통기한'].dt.strftime('%Y-%m-%d')
            st.dataframe(expiring_display[['제품명', 'Lot 번호', '유통기한', '보관 위치', '현재 재고']], use_container_width=True)
        if not expired.empty:
            st.error(f"**유통기한 만료** (재고 있음)")
            expired_display = expired.copy()
            expired_display['유통기한'] = expired_display['유통기한'].dt.strftime('%Y-%m-%d')
            st.dataframe(expired_display[['제품명', 'Lot 번호', '유통기한', '보관 위치', '현재 재고']], use_container_width=True)
        
        low_stock = df_inventory[
            (df_inventory['현재 재고'] <= df_inventory['알림 기준 수량']) &
            (df_inventory['현재 재고'] > 0) &
            (df_inventory['알림 무시'] != "예") 
        ]
        out_of_stock = df_inventory[
            (df_inventory['현재 재고'] <= 0) &
            (df_inventory['알림 무시'] != "예") 
        ]
        if not low_stock.empty:
            st.warning(f"**재고 부족 (알림 기준 수량 이하)**")
            st.dataframe(low_stock[['제품명', 'Lot 번호', '현재 재고', '단위', '알림 기준 수량']], use_container_width=True)
        if not out_of_stock.empty:
            st.error(f"**재고 소진 (0 이하)**")
            st.dataframe(out_of_stock[['제품명', 'Lot 번호', '현재 재고', '단위']], use_container_width=True)
            
        if expiring_soon.empty and expired.empty and low_stock.empty and out_of_stock.empty:
            st.success("✅ 모든 재고가 양호합니다!")
        
        # (v49의 알림 해제 섹션)
        st.divider()
        st.subheader("🗃️ 품목 보관 (알림 해제)")
        
        if not out_of_stock.empty:
            mute_options = [
                f"{row['제품명']} / Lot: {row['Lot 번호']}" for index, row in out_of_stock.iterrows()
            ]
            mute_options.insert(0, "알림을 해제할 품목을 선택하세요...") 
            
            selected_item_to_mute = st.selectbox("재고 소진 품목 알림 해제:", options=mute_options)
            
            if st.button("➡️ 이 품목 알림 해제하기"):
                if selected_item_to_mute == mute_options[0]:
                    st.warning("알림을 해제할 품목을 선택하세요.")
                else:
                    try:
                        product_to_mute, lot_to_mute = selected_item_to_mute.split(" / Lot: ")
                        
                        sh_db = client.open(REAGENT_DB_NAME)
                        sheet_db = sh_db.worksheet(REAGENT_DB_TAB)
                        
                        all_data = sheet_db.get_all_records()
                        target_rows = []
                        for i, record in enumerate(all_data):
                            if (str(record['제품명']) == product_to_mute and 
                                str(record['Lot 번호']) == lot_to_mute):
                                target_rows.append(i + 2) 
                        
                        if not target_rows:
                            st.error(f"시트에서 '{selected_item_to_mute}'을(를) 찾지 못했습니다. (데이터 확인 필요)")
                        else:
                            # (v49: L열(12)로 '알림 무시' 컬럼 위치 변경)
                            for row_index in target_rows:
                                sheet_db.update_cell(row_index, 12, "예") # 12 = L열
                            
                            st.success(f"✅ '{product_to_mute}' (Lot: {lot_to_mute}) 품목이 알림에서 해제되었습니다.")
                            st.cache_data.clear()
                            st.rerun()

                    except Exception as e:
                        st.error(f"알림 해제 중 오류 발생: {e}")
        else:
            st.info("현재 알림을 해제할 '재고 소진' 품목이 없습니다.")
            
        st.divider()

        # --- 6. 전체 재고 현황 (v50 수정됨) ---
        st.subheader("전체 재고 현황")
        
        # ▼▼▼ [신규] v50: 고급 필터 (v48) ▼▼▼
        st.write("**고급 필터**")
        col1, col2 = st.columns(2)
        
        # (제조사 필터)
        all_manufacturers = sorted(df_inventory['제조사'].dropna().unique())
        selected_manufacturers = col1.multiselect(
            "제조사 필터:",
            options=all_manufacturers,
            default=all_manufacturers
        )
        
        # (보관 위치 필터)
        all_locations = sorted(df_inventory['보관 위치'].dropna().unique())
        selected_locations = col2.multiselect(
            "보관 위치 필터:",
            options=all_locations,
            default=all_locations
        )
        
        # (검색창)
        search_query = st.text_input(
            "🔎 빠른 검색 (제품명, Cat. No., Lot 번호)", 
            placeholder="DMEM, 1111, 2222dd 등으로 검색..."
        )
        # ▲▲▲ [신규] v50 ▲▲▲

        
        # (v49 방식: 컬럼 통합)
        display_columns = [
            "제품명", "제조사", "Cat. No.", "Lot 번호", 
            "현재 재고", "단위", "최초 수량", "총 사용량",
            "재고 비율 (%)", 
            "알림 기준 수량", "알림 무시", 
            "유통기한", "보관 위치", "등록자", "등록 날짜"
        ]
        
        available_columns = [col for col in display_columns if col in df_inventory.columns]
        
        if '유통기한' in available_columns:
            df_inventory['유통기한 (YYYY-MM-DD)'] = df_inventory['유통기한'].dt.strftime('%Y-%m-%d')
            available_columns[available_columns.index('유통기한')] = '유통기한 (YYYY-MM-DD)'
            
        # ▼▼▼ [수정됨] v50: 필터 로직 적용 ▼▼▼
        df_display = df_inventory[available_columns] 
        
        # (1. 고급 필터 적용)
        df_display = df_display[
            df_display['제조사'].isin(selected_manufacturers) &
            df_display['보관 위치'].isin(selected_locations)
        ]
        
        # (2. 빠른 검색 적용)
        if search_query:
            query = search_query.lower() 
            mask = (
                df_display['제품명'].astype(str).str.lower().str.contains(query) |
                df_display['제조사'].astype(str).str.lower().str.contains(query) | 
                df_display['Cat. No.'].astype(str).str.lower().str.contains(query) |
                df_display['Lot 번호'].astype(str).str.lower().str.contains(query)
            )
            df_display = df_display[mask]
        # ▲▲▲ [수정됨] v50 ▲▲▲
            
        # (v49 방식: data_editor + column_config)
        st.data_editor( 
            df_display,
            use_container_width=True,
            disabled=True, 
            
            column_config={
                "재고 비율 (%)": st.column_config.ProgressColumn(
                    "재고 비율 (%)",  
                    format="%.1f%%", # (소수점 첫째 자리 %)
                    min_value=0,
                    max_value=100,
                ),
                "현재 재고": st.column_config.NumberColumn(
                    "현재 재고",
                    format="%.2f", 
                ),
                "총 사용량": st.column_config.NumberColumn(
                    "총 사용량",
                    format="%.0f", 
                ),
                "알림 기준 수량": st.column_config.NumberColumn(
                    "알림 기준",
                    format="%.2f",
                ),
                "알림 무시": st.column_config.TextColumn(
                    "알림 무시"
                ),
                "제조사": st.column_config.TextColumn( 
                    "제조사"
                ),
            }
        )
        
        # (v49의 상세 사용 이력 섹션)
        st.divider()
        st.subheader("📈 상세 사용 이력 (필터링된 품목)")
        
        if df_display.empty:
            if search_query or (len(selected_manufacturers) < len(all_manufacturers)) or (len(selected_locations) < len(all_locations)):
                st.warning("선택된 필터/검색어에 해당하는 품목이 없습니다.")
            else:
                st.info("상세 이력을 보려면 위 검색창에서 품목을 검색하세요.")
        else:
            products_to_show = df_display['제품명'].unique()
            lots_to_show = df_display['Lot 번호'].unique()
            
            log_mask = (
                df_log['제품명'].isin(products_to_show) &
                df_log['Lot 번호'].isin(lots_to_show)
            )
            df_log_filtered = df_log[log_mask]
        
            if df_log_filtered.empty:
                st.info("선택된 품목에 대한 사용 기록(Usage Log)이 없습니다.")
            else:
                df_log_filtered = df_log_filtered.sort_values(by="Timestamp", ascending=False)
                df_log_filtered['Timestamp (YYYY-MM-DD)'] = df_log_filtered['Timestamp'].dt.strftime('%Y-%m-%d %H:%M')
                st.dataframe(
                    df_log_filtered[['Timestamp (YYYY-MM-DD)', '제품명', 'Lot 번호', '사용자', '사용량', '비고']], 
                    use_container_width=True
                )
