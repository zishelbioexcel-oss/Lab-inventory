import streamlit as st
import gspread
import json
import base64
from oauth2client.service_account import ServiceAccountCredentials
import pandas as pd
from datetime import datetime

# --- 1. 앱 설정 ---
st.set_page_config(page_title="실험실 재고 관리기 v52", layout="wide")
st.title("🔬 실험실 재고 관리기 v52 (Master/History 분리형)")
st.caption("BOM은 기준 정보만 관리하고, 재고 수량은 입출고 내역을 통해 자동 계산합니다.")

# --- 2. 구글 시트 연결 설정 ---
REAGENT_DB_NAME = "Reagent_DB"
REAGENT_DB_TAB = "Master"       # BOM (기준정보)
USAGE_LOG_NAME = "Usage_Log"    # 입출고 내역 (변동정보)
USAGE_LOG_TAB = "Log"

# (1) 인증 함수
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

# (2) 데이터 로드 함수 (수정됨: 빈 시트 방어 로직 추가)
def load_data(client):
    # Master DB 로드
    sh_db = client.open(REAGENT_DB_NAME)
    ws_db = sh_db.worksheet(REAGENT_DB_TAB)
    data_db = ws_db.get_all_records()
    df_master = pd.DataFrame(data_db)
    
    # Log 로드
    sh_log = client.open(USAGE_LOG_NAME)
    ws_log = sh_log.worksheet(USAGE_LOG_TAB)
    data_log = ws_log.get_all_records()
    df_log = pd.DataFrame(data_log)
    
    # [수정] Log 시트가 비어있거나 '수량' 컬럼이 없으면 빈 껍데기 생성 (에러 방지)
    required_cols = ['날짜', '구분', '제품명', 'Lot 번호', '수량', '유효기간', '담당자', '비고']
    
    # 데이터프레임이 비어있거나 필수 컬럼이 하나라도 없으면 재정의
    if df_log.empty or not set(['수량', '제품명']).issubset(df_log.columns):
        df_log = pd.DataFrame(columns=required_cols)
    
    return df_master, df_log, ws_db, ws_log

# --- 3. 앱 실행 로직 ---
client, err = get_gspread_client()
if err: st.error(err); st.stop()

# 탭 구성: 이젠 '입고'와 '사용'이 같은 레벨의 액션입니다.
tab1, tab2, tab3 = st.tabs(["📂 BOM(품목) 관리", "📦 입고/사용 등록", "📊 실시간 재고 현황"])

# ==============================================================================
# [Tab 1] BOM(품목) 관리 - 기준 정보 업로드
# ==============================================================================
with tab1:
    st.header("📂 BOM 마스터 데이터 관리")
    st.info("여기서는 '재고 수량'을 입력하지 않습니다. 품목의 **정의(이름, 제조사, 위치 등)**만 등록하세요.")
    
    with st.expander("엑셀 BOM 업로드 (기준 정보 갱신)", expanded=True):
        uploaded_file = st.file_uploader("정리된 BOM 엑셀 파일 (.xlsx)", type=['xlsx'])
        if uploaded_file:
            df_upload = pd.read_excel(uploaded_file)
            df_upload.columns = df_upload.columns.str.strip() # 헤더 공백 제거
            
            st.write("미리보기:", df_upload.head(3))
            
            # 매핑: 수량/날짜 컬럼이 사라졌으므로 매핑도 단순해집니다.
            COL_MAP = {
                "제품명": "품목명", "제조사": "제조사", "Cat. No.": "Cat. No.",
                "단위": "단위", "보관 위치": "보관 장소", "알림 기준 수량": "안전재고"
            }
            
            if st.button("🚀 기준 정보 덮어쓰기 (DB 초기화)"):
                try:
                    sh = client.open(REAGENT_DB_NAME)
                    ws = sh.worksheet(REAGENT_DB_TAB)
                    
                    processed = []
                    header = ["제품명", "제조사", "Cat. No.", "단위", "보관 위치", "알림 기준 수량", "등록일", "등록자"]
                    processed.append(header)
                    
                    for _, row in df_upload.iterrows():
                        p_name = str(row.get(COL_MAP["제품명"], "")).strip()
                        if not p_name or p_name == "nan": continue
                        
                        # 안전재고 숫자 처리
                        try: safe_stock = float(str(row.get(COL_MAP["알림 기준 수량"], 0)).replace("-","0"))
                        except: safe_stock = 0.0

                        processed.append([
                            p_name,
                            str(row.get(COL_MAP["제조사"], "-")),
                            str(row.get(COL_MAP["Cat. No."], "-")),
                            str(row.get(COL_MAP["단위"], "ea")),
                            str(row.get(COL_MAP["보관 위치"], "-")),
                            safe_stock,
                            datetime.now().strftime("%Y-%m-%d"),
                            "관리자(일괄)"
                        ])
                    
                    ws.clear()
                    ws.update(processed)
                    st.success(f"✅ 기준 정보 {len(processed)-1}건 등록 완료! 이제 '입고' 탭에서 수량을 채워넣으세요.")
                    st.cache_data.clear()
                except Exception as e:
                    st.error(f"업로드 실패: {e}")

# ==============================================================================
# [Tab 2] 입고 및 사용 등록 (핵심 기능)
# ==============================================================================
with tab2:
    st.header("📦 자재 수불 관리 (Transaction)")
    
    df_master, df_log, _, ws_log = load_data(client)
    
    if df_master.empty:
        st.warning("먼저 BOM을 등록해주세요.")
    else:
        products = sorted(df_master['제품명'].unique())
        
        # 1. 작업 유형 선택
        col1, col2 = st.columns([1, 2])
        action_type = col1.radio("작업 유형 선택", ["🔵 입고 (구매/채워넣기)", "🔴 사용 (소진/출고)"])
        
        # 2. 품목 선택
        selected_product = col2.selectbox("품목 선택", products)
        
        # 품목 상세 정보 가져오기
        item_info = df_master[df_master['제품명'] == selected_product].iloc[0]
        unit = item_info['단위']
        
        st.divider()
        
        # 3. 입력 폼 (입고냐 사용이냐에 따라 다르게 보임)
        with st.form("action_form", clear_on_submit=True):
            c1, c2, c3 = st.columns(3)
            
            qty = c1.number_input(f"수량 ({unit})", min_value=0.1, step=1.0)
            user = c2.text_input("담당자", value="관리자")
            
            # 입고일 때는 Lot와 유효기간이 필수 / 사용일 때는 Lot 선택
            lot_input = "Initial"
            expiry_input = "-"
            
            if "입고" in action_type:
                lot_input = c3.text_input("Lot 번호 (새로 입력)", value=datetime.now().strftime("%Y%m%d"))
                expiry_input = st.date_input("유효기간 설정", datetime.now()).strftime("%Y-%m-%d")
                note_label = "비고 (구매처 등)"
            else:
                # 사용 시에는 기존 Lot 중에서 선택 (구현 간소화를 위해 텍스트 입력 혹은 추후 고도화)
                # 현재는 단순화를 위해 Lot 직접 입력 혹은 'Any'
                lot_input = c3.text_input("Lot 번호 (사용할 제품)", value="Initial")
                note_label = "비고 (실험명 등)"
            
            note = st.text_input(note_label)
            
            # 저장 버튼
            if st.form_submit_button(f"{action_type} 저장하기"):
                # 사용(출고)일 경우 수량을 음수로 저장
                final_qty = qty if "입고" in action_type else -qty
                action_code = "IN" if "입고" in action_type else "OUT"
                
                # 로그 저장 순서: [날짜, 구분, 제품명, Lot, 수량, 유효기간, 담당자, 비고]
                row_data = [
                    datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    action_code,
                    selected_product,
                    lot_input,
                    final_qty,  # 핵심: 입고는 +, 사용은 -
                    expiry_input,
                    user,
                    note
                ]
                
                ws_log.append_row(row_data)
                st.success(f"✅ {selected_product} {qty}{unit} {action_type} 처리되었습니다.")
                st.cache_data.clear() # 데이터 갱신

# ==============================================================================
# [Tab 3] 실시간 재고 현황 (Dashboard)
# ==============================================================================
with tab3:
    st.header("📊 실시간 재고 현황")
    if st.button("🔄 새로고침"): st.rerun()
    
    df_master, df_log, _, _ = load_data(client)
    
    if df_master.empty:
        st.info("데이터가 없습니다.")
    else:
        # --- 재고 계산 로직 (핵심) ---
        # 1. 로그 데이터가 없으면 재고는 0
        if df_log.empty:
            df_stock = df_master.copy()
            df_stock['현재고'] = 0.0
        else:
            # 2. 제품별 수량 합계 계산 (입고는 +, 출고는 - 이므로 그냥 sum하면 됨)
            # 수치형 변환
            df_log['수량'] = pd.to_numeric(df_log['수량'], errors='coerce').fillna(0)
            
            # 제품별 GroupBy
            stock_grp = df_log.groupby('제품명')['수량'].sum().reset_index()
            stock_grp.rename(columns={'수량': '현재고'}, inplace=True)
            
            # 3. Master와 결합 (Left Join)
            df_stock = pd.merge(df_master, stock_grp, on='제품명', how='left')
            df_stock['현재고'] = df_stock['현재고'].fillna(0) # 거래 없는 품목은 0 처리
            
        # --- 화면 표시 ---
        # 알림 로직
        df_stock['알림 기준 수량'] = pd.to_numeric(df_stock['알림 기준 수량'], errors='coerce').fillna(0)
        low_stock = df_stock[df_stock['현재고'] <= df_stock['알림 기준 수량']]
        
        c1, c2 = st.columns(2)
        c1.metric("총 등록 품목 수", f"{len(df_stock)}개")
        c2.metric("재고 부족 품목", f"{len(low_stock)}개", delta_color="inverse")
        
        if not low_stock.empty:
            st.error("🚨 재고 부족 알림")
            st.dataframe(low_stock[['제품명', '현재고', '알림 기준 수량', '보관 위치']], hide_index=True)
        
        st.divider()
        st.subheader("📦 전체 재고 리스트")
        
        # 검색 기능
        search = st.text_input("🔍 품목 검색")
        if search:
            df_stock = df_stock[df_stock['제품명'].str.contains(search, case=False)]
            
        # 보기 좋게 컬럼 정리
        display_cols = ["제품명", "현재고", "단위", "보관 위치", "제조사", "Cat. No.", "알림 기준 수량"]
        st.dataframe(df_stock[display_cols], use_container_width=True, hide_index=True)

        # (선택 사항) 입출고 히스토리 보기
        with st.expander("📜 상세 입출고 이력 보기"):
            if not df_log.empty:
                # use 라고 적힌 부분을 use_container_width=True 로 수정했습니다.
                st.dataframe(df_log.sort_values(by=df_log.columns[0], ascending=False), use_container_width=True)
            else:
                st.info("아직 입출고 기록이 없습니다.")


