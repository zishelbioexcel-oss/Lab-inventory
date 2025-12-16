import streamlit as st
import gspread
import json
import base64
from oauth2client.service_account import ServiceAccountCredentials
import pandas as pd
from datetime import datetime

# --- 1. 앱의 기본 설정 ---
st.set_page_config(page_title="실험실 재고 관리기 v51", layout="wide")
st.title("🔬 실험실 재고 관리기 v51 (Excel BOM 연동형)")
st.write("엑셀 BOM 파일을 업로드하여 초기 재고를 셋팅하고, 사용량을 관리합니다.")

# --- 2. Google Sheets 인증 및 설정 ---
REAGENT_DB_NAME = "Reagent_DB"  
REAGENT_DB_TAB = "Master"       
USAGE_LOG_NAME = "Usage_Log"    
USAGE_LOG_TAB = "Log"           

# (1) 인증 클라이언트 생성
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
            # 로컬 실행 시 secrets.toml 경로 확인 필요
            creds = ServiceAccountCredentials.from_service_account_file('.streamlit/secrets.toml', scope)
        client = gspread.authorize(creds)
        return client, None
    except FileNotFoundError:
        return None, "로컬 Secrets 파일('.streamlit/secrets.toml')을 찾을 수 없습니다."
    except Exception as e:
        return None, f"Google 인증 실패: {e}"

# (2) 마스터 DB 로드
@st.cache_data(ttl=60) 
def load_reagent_db(_client):
    try:
        sh = _client.open(REAGENT_DB_NAME)
        sheet = sh.worksheet(REAGENT_DB_TAB)
        data = sheet.get_all_records()
        
        # 필수 컬럼 정의
        required_cols = ["제품명", "제조사", "Cat. No.", "Lot 번호", "최초 수량", "단위", "유통기한", "보관 위치", "등록 날짜", "등록자", "알림 기준 수량", "알림 무시"]
        
        if not data:
            return pd.DataFrame(columns=required_cols)
        
        df = pd.DataFrame(data)
        
        # 누락된 컬럼이 있다면 빈 값으로 생성
        for col in required_cols:
            if col not in df.columns:
                df[col] = ""
        
        # 데이터 타입 정리
        df['제품명'] = df['제품명'].astype(str)
        df['최초 수량'] = pd.to_numeric(df['최초 수량'], errors='coerce').fillna(0)
        df['알림 기준 수량'] = pd.to_numeric(df['알림 기준 수량'], errors='coerce').fillna(0) 
        df['유통기한'] = pd.to_datetime(df['유통기한'], errors='coerce') 
        df['등록 날짜'] = pd.to_datetime(df['등록 날짜'], errors='coerce') 
        
        return df    
    except Exception as e:
        st.error(f"Reagent_DB 로드 실패: {e}")
        return pd.DataFrame()

# (3) 사용 기록(Log) 로드
@st.cache_data(ttl=60)
def load_usage_log(_client):
    try:
        sh = _client.open(USAGE_LOG_NAME)
        sheet = sh.worksheet(USAGE_LOG_TAB)
        data = sheet.get_all_records()
        if not data:
            return pd.DataFrame(columns=["제품명", "Lot 번호", "사용량", "Timestamp", "사용자", "비고"]) 
        
        df = pd.DataFrame(data)
        df['사용량'] = pd.to_numeric(df['사용량'], errors='coerce').fillna(0)
        df['Timestamp'] = pd.to_datetime(df['Timestamp'], errors='coerce') 
        return df
    except Exception as e:
        st.error(f"Usage_Log 로드 실패: {e}")
        return pd.DataFrame()

# --- 3. 앱 실행 ---
client, auth_error_msg = get_gspread_client()

if auth_error_msg:
    st.error(auth_error_msg)
    st.stop() 

tab1, tab2, tab3 = st.tabs(["📝 새 품목/엑셀 등록", "📉 시약 사용", "📊 재고 대시보드"])

# ==============================================================================
# [Tab 1] 새 품목 등록 & 엑셀 업로드 (핵심 수정됨)
# ==============================================================================
with tab1:
    st.header("📝 품목 등록 관리")

    # --- 1. 엑셀 일괄 업로드 섹션 ---
    with st.expander("📂 엑셀 BOM 파일로 일괄 업로드 (추천)", expanded=True):
        st.info("💡 작성하신 BOM 엑셀 파일을 업로드하면 DB에 한 번에 등록됩니다.")
        uploaded_file = st.file_uploader("엑셀 파일 선택 (.xlsx)", type=['xlsx'])
        
        if uploaded_file:
            try:
                # 엑셀 읽기
                df_upload = pd.read_excel(uploaded_file)
                # 1. 모든 컬럼에서 '-' 문자를 0으로 치환 (숫자 컬럼 계산을 위해)
                df_upload = df_upload.replace('-', 0)
                # 2. 혹시 몰라 NaN(빈칸)도 0으로 채우기 (선택사항)
                df_upload = df_upload.fillna(0)
                
                st.write("엑셀 데이터 미리보기:", df_upload.head(3))
                
                # 매핑 가이드 (사용자 엑셀 -> 앱 DB)
                # 좌측: 앱에서 쓰는 이름 / 우측: 엑셀 파일의 헤더 이름
                COLUMN_MAPPING = {
                    "제품명": "품목명",       # 엑셀의 '품목명' -> 앱의 '제품명'
                    "제조사": "제조사",
                    "Cat. No.": "Cat.No.",   # 엑셀의 'Cat.No.'
                    "최초 수량": "용량",      # 엑셀의 '용량' 수치
                    "단위": "단위",
                    "유통기한": "유효기간",
                    "보관 위치": "보관 장소",  # 엑셀의 '보관 장소'
                    "알림 기준 수량": "안전재고" # 엑셀의 '안전재고'
                }
                
                col1, col2 = st.columns(2)
                registrant_name = col1.text_input("등록자 이름 (일괄 적용)", value="관리자")
                
                if st.button("🚀 Google Sheets에 덮어쓰기 (기존 데이터 초기화)"):
                    sh = client.open(REAGENT_DB_NAME)
                    sheet = sh.worksheet(REAGENT_DB_TAB)
                    
                    # 1. 데이터 전처리
                    processed_data = []
                    for index, row in df_upload.iterrows():
                        # 필수값 매핑 (없으면 기본값)
                        product_name = str(row.get(COLUMN_MAPPING["제품명"], ""))
                        if not product_name or product_name == "nan": continue # 품목명 없으면 스킵
                        
                        # Lot 번호 처리 (엑셀에 없으면 Initial로 지정)
                        lot_no = str(row.get("Lot No", "Initial"))
                        if lot_no == "nan": lot_no = "Initial"

                        # 날짜 변환 (YYYY-MM-DD)
                        expiry_raw = row.get(COLUMN_MAPPING["유통기한"], "")
                        expiry_str = ""
                        if pd.notnull(expiry_raw) and expiry_raw != "":
                            try:
                                expiry_str = pd.to_datetime(expiry_raw).strftime("%Y-%m-%d")
                            except:
                                expiry_str = str(expiry_raw) # 변환 실패시 텍스트 그대로

                        item = {
                            "제품명": product_name,
                            "제조사": str(row.get(COLUMN_MAPPING["제조사"], "-")),
                            "Cat. No.": str(row.get(COLUMN_MAPPING["Cat. No."], "-")),
                            "Lot 번호": lot_no,
                            "최초 수량": float(row.get(COLUMN_MAPPING["최초 수량"], 0)),
                            "단위": str(row.get(COLUMN_MAPPING["단위"], "ea")),
                            "유통기한": expiry_str,
                            "보관 위치": str(row.get(COLUMN_MAPPING["보관 위치"], "-")),
                            "등록 날짜": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                            "등록자": registrant_name,
                            "알림 기준 수량": float(row.get(COLUMN_MAPPING["알림 기준 수량"], 0)),
                            "알림 무시": "아니요"
                        }
                        # 리스트 순서 보장 (DB 컬럼 순서대로)
                        row_list = [
                            item["제품명"], item["제조사"], item["Cat. No."], item["Lot 번호"],
                            item["최초 수량"], item["단위"], item["유통기한"], item["보관 위치"],
                            item["등록 날짜"], item["등록자"], item["알림 기준 수량"], item["알림 무시"]
                        ]
                        processed_data.append(row_list)
                    
                    # 2. 구글 시트 초기화 및 업데이트
                    sheet.clear()
                    # 헤더 추가
                    header = ["제품명", "제조사", "Cat. No.", "Lot 번호", "최초 수량", "단위", "유통기한", "보관 위치", "등록 날짜", "등록자", "알림 기준 수량", "알림 무시"]
                    processed_data.insert(0, header)
                    
                    sheet.update(processed_data)
                    st.success(f"✅ 총 {len(processed_data)-1}건의 데이터가 성공적으로 등록되었습니다!")
                    st.cache_data.clear()
                    st.rerun()

            except Exception as e:
                st.error(f"엑셀 처리 중 오류 발생: {e}")

    # --- 2. 개별 직접 등록 섹션 (기존 기능) ---
    st.divider()
    with st.expander("➕ 개별 품목 직접 등록 (추가)"):
        with st.form(key="new_item_form", clear_on_submit=True):
            col1, col2 = st.columns(2)
            with col1:
                product_name = st.text_input("제품명*")
                manufacturer = st.text_input("제조사")
                cat_no = st.text_input("Cat. No.")
                lot_no = st.text_input("Lot 번호*", value="Initial")
            with col2:
                initial_qty = st.number_input("최초 수량*", min_value=0.0, step=1.0)
                unit = st.selectbox("단위*", ["mL", "L", "g", "kg", "box", "kit", "ea"])
                alert_qty = st.number_input("알림 기준 수량 (안전재고)", value=5.0)
            
            location = st.text_input("보관 위치")
            expiry_date = st.date_input("유통기한", datetime.now() + pd.DateOffset(years=1))
            registrant = st.text_input("등록자 이름*")
            
            if st.form_submit_button("등록하기"):
                if not product_name or initial_qty <= 0:
                    st.error("제품명과 수량은 필수입니다.")
                else:
                    sh = client.open(REAGENT_DB_NAME)
                    sheet = sh.worksheet(REAGENT_DB_TAB)
                    row_data = [
                        product_name, manufacturer, cat_no, lot_no, float(initial_qty), unit,
                        expiry_date.strftime("%Y-%m-%d"), location, 
                        datetime.now().strftime("%Y-%m-%d %H:%M:%S"), registrant, float(alert_qty), "아니요"
                    ]
                    sheet.append_row(row_data)
                    st.success(f"{product_name} 등록 완료!")
                    st.cache_data.clear()

# ==============================================================================
# [Tab 2] 시약 사용 기록
# ==============================================================================
with tab2:
    st.header("📉 시약 사용 기록")
    df_db = load_reagent_db(client)
    df_log = load_usage_log(client)

    if df_db.empty:
        st.warning("등록된 품목이 없습니다. Tab 1에서 엑셀 업로드 또는 등록을 먼저 해주세요.")
    else:
        # 제품 선택
        all_products = sorted(df_db['제품명'].unique())
        selected_product = st.selectbox("사용할 제품 선택", all_products)
        
        # Lot 선택
        if selected_product:
            lots = df_db[df_db['제품명'] == selected_product]['Lot 번호'].unique()
            selected_lot = st.selectbox("Lot 번호 선택", lots)
            
            # 현재 재고 계산
            item_info = df_db[(df_db['제품명'] == selected_product) & (df_db['Lot 번호'] == selected_lot)].iloc[0]
            initial_stock = item_info['최초 수량']
            
            used_stock = 0
            if not df_log.empty:
                used_stock = df_log[(df_log['제품명'] == selected_product) & (df_log['Lot 번호'] == selected_lot)]['사용량'].sum()
            
            current_stock = initial_stock - used_stock
            unit = item_info['단위']
            
            st.info(f"📊 **현재 재고: {current_stock} {unit}** (최초: {initial_stock} - 사용: {used_stock})")
            
            # 사용 입력 폼
            with st.form("usage_form", clear_on_submit=True):
                use_qty = st.number_input("사용량", min_value=0.0, step=0.1, max_value=float(current_stock))
                user_name = st.text_input("사용자")
                note = st.text_input("비고 (실험명 등)")
                
                if st.form_submit_button("사용 기록 저장"):
                    if use_qty > 0 and user_name:
                        sh_log = client.open(USAGE_LOG_NAME)
                        sheet_log = sh_log.worksheet(USAGE_LOG_TAB)
                        sheet_log.append_row([
                            datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                            selected_product, selected_lot, use_qty, user_name, note
                        ])
                        st.success("저장되었습니다.")
                        st.cache_data.clear()
                        st.rerun()
                    else:
                        st.error("사용량과 사용자 이름을 확인하세요.")

# ==============================================================================
# [Tab 3] 재고 대시보드 (자동 알림)
# ==============================================================================
with tab3:
    st.header("📊 재고 현황 대시보드")
    
    if st.button("🔄 데이터 새로고침"):
        st.cache_data.clear()
        st.rerun()

    df_db = load_reagent_db(client)
    df_log = load_usage_log(client)
    
    if not df_db.empty:
        # 재고 계산 로직
        if not df_log.empty:
            usage_grp = df_log.groupby(['제품명', 'Lot 번호'])['사용량'].sum().reset_index()
            df_final = pd.merge(df_db, usage_grp, on=['제품명', 'Lot 번호'], how='left')
            df_final['사용량'] = df_final['사용량'].fillna(0)
        else:
            df_final = df_db.copy()
            df_final['사용량'] = 0
            
        df_final['현재 재고'] = df_final['최초 수량'] - df_final['사용량']
        
        # --- 🚨 자동 알림 로직 ---
        st.subheader("🚨 알림 센터")
        col1, col2 = st.columns(2)
        
        # 1. 재고 부족 알림
        low_stock = df_final[df_final['현재 재고'] <= df_final['알림 기준 수량']]
        if not low_stock.empty:
            col1.error(f"⚠️ **재고 부족 ({len(low_stock)}건)**")
            col1.dataframe(low_stock[['제품명', '현재 재고', '알림 기준 수량', '보관 위치']], hide_index=True)
        else:
            col1.success("재고 수량 양호 ✅")

        # 2. 유효기간 임박 알림 (30일)
        today = pd.to_datetime(datetime.now().date())
        df_final['유통기한'] = pd.to_datetime(df_final['유통기한'])
        
        expiring = df_final[
            (df_final['유통기한'] <= today + pd.DateOffset(days=30)) & 
            (df_final['현재 재고'] > 0)
        ]
        
        if not expiring.empty:
            col2.warning(f"⏳ **유효기간 임박/만료 ({len(expiring)}건)**")
            col2.dataframe(expiring[['제품명', '유통기한', '현재 재고', '보관 위치']], hide_index=True)
        else:
            col2.success("유효기간 양호 ✅")

        st.divider()
        
        # --- 전체 재고 테이블 (필터링) ---
        st.subheader("📦 전체 재고 목록")
        
        # 필터
        f_col1, f_col2 = st.columns(2)
        manufacturers = ["전체"] + list(df_final['제조사'].unique())
        selected_mfg = f_col1.selectbox("제조사 필터", manufacturers)
        
        search_txt = f_col2.text_input("🔍 품목명 검색")
        
        view_df = df_final.copy()
        if selected_mfg != "전체":
            view_df = view_df[view_df['제조사'] == selected_mfg]
        if search_txt:
            view_df = view_df[view_df['제품명'].str.contains(search_txt, case=False)]
            
        # 보여줄 컬럼 정리
        view_cols = ["제품명", "제조사", "Cat. No.", "Lot 번호", "현재 재고", "단위", "유통기한", "보관 위치"]
        st.dataframe(view_df[view_cols], use_container_width=True, hide_index=True)

