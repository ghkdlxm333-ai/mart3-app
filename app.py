import streamlit as st
import pandas as pd
import io
from datetime import datetime

# 1. 화면 설정 (Wide 모드 및 상단 바 설정)
st.set_page_config(
    page_title="발주처별 수주 자동화 시스템",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- [날짜 변환 함수: YYYYMMDD 형식] ---
def format_delivery_date(val):
    if pd.isna(val) or str(val).strip() == "":
        return datetime.now().strftime('%Y%m%d')
    try:
        dt = pd.to_datetime(val)
        return dt.strftime('%Y%m%d')
    except:
        clean_val = ''.join(filter(str.isdigit, str(val)))
        return clean_val[:8] if len(clean_val) >= 8 else datetime.now().strftime('%Y%m%d')

@st.cache_data
def load_master_data(file_path):
    """각 채널별 마스터 파일 로드 및 매핑 데이터 생성"""
    try:
        xls = pd.ExcelFile(file_path)
        sheet_map = {s.strip(): s for s in xls.sheet_names}
        
        # 배송코드 시트 로직
        target_sheet = sheet_map.get('배송코드')
        if not target_sheet:
            return None, f"'{file_path}'에 '배송코드' 시트가 없습니다."
        
        df_tmp = pd.read_excel(xls, sheet_name=target_sheet, dtype=str)
        df_center = None
        for i, row in df_tmp.iterrows():
            if '배송코드' in [str(v).strip() for v in row.values]:
                df_center = pd.read_excel(xls, sheet_name=target_sheet, skiprows=i+1, dtype=str)
                df_center.columns = [str(c).strip() for c in df_tmp.iloc[i]]
                break
        
        if df_center is None: df_center = df_tmp
        
        # 매핑 구성
        c_to_b_map = dict(zip(df_center['센터코드'].str.strip(), df_center['배송코드'].str.strip()))
        b_to_name_map = dict(zip(df_center['배송코드'].str.strip(), df_center.iloc[:, 2].str.strip()))
        
        # 제품명 시트 로직
        prod_sheet = sheet_map.get('제품명')
        df_prod = pd.read_excel(xls, sheet_name=prod_sheet, dtype=str) if prod_sheet else None
        p_map = dict(zip(df_prod['상품코드'].str.strip(), df_prod['ME코드'].str.strip())) if df_prod is not None else {}
        n_map = dict(zip(df_prod['상품코드'].str.strip(), df_prod.iloc[:, 1].str.strip())) if df_prod is not None else {}
        
        return {'c_to_b': c_to_b_map, 'b_to_n': b_to_name_map, 'products': p_map, 'names': n_map}, None
    except Exception as e:
        return None, str(e)

# --- 채널 설정 ---
CHANNELS = {
    'TRADERS': {'name': '이마트 트레이더스', 'code': '81011010', 'file': '트레이더스_서식파일_업데이트용.xlsx'},
    'NOBRAND': {'name': '노브랜드', 'code': '81010000', 'file': '노브랜드_서식파일_업데이트용.xlsx'},
    'EMART': {'name': '이마트', 'code': '81010000', 'file': '이마트_서식파일_업데이트용.xlsx'}
}

st.title("🛒 발주처별 통합 수주 자동화")
st.info("점포명에 'TR'이 포함되면 트레이더스, 'NBR'이 포함되면 노브랜드, 나머지는 이마트로 자동 분류됩니다.")

# 마스터 데이터 로드
masters = {}
status_ok = True
for k, v in CHANNELS.items():
    data, err = load_master_data(v['file'])
    if err:
        st.warning(f"⚠️ {v['file']} 로드 실패 (파일 확인 필요)")
        status_ok = False
    else:
        masters[k] = data

if status_ok:
    uploaded_file = st.file_uploader("일반 주문서(ORDERS) 업로드", type=['xlsx'])
    
    if uploaded_file:
        try:
            df_raw = pd.read_excel(uploaded_file)
            date_col = next((c for c in df_raw.columns if '센터입하일자' in str(c)), '센터입하일자')
            
            final_data = []
            for _, row in df_raw.iterrows():
                store_name = str(row['점포명'])
                
                # [분류 로직] 점포명 키워드 기준
                if 'TR' in store_name.upper():
                    ch_key = 'TRADERS'
                elif 'NBR' in store_name.upper():
                    ch_key = 'NOBRAND'
                else:
                    ch_key = 'EMART'
                
                master = masters[ch_key]
                
                # 데이터 추출 및 매핑
                center_code = str(row['센터코드']).strip()
                delivery_code = master['c_to_b'].get(center_code, "")
                delivery_place = master['b_to_n'].get(delivery_code, store_name)
                
                prod_code = str(row['상품코드']).strip()
                me_code = master['products'].get(prod_code, prod_code)
                prod_name = master['names'].get(prod_code, row['상품명'])
                
                final_data.append({
                    '발주처분류': CHANNELS[ch_key]['name'],
                    '수주일자': datetime.now().strftime('%Y%m%d'),
                    '납품일자': format_delivery_date(row[date_col]),
                    '발주처코드': CHANNELS[ch_key]['code'],
                    '배송코드': delivery_code,
                    '배송지': delivery_place,
                    '상품코드': me_code,
                    '상품명': prod_name,
                    'UNIT수량': pd.to_numeric(row['수량'], errors='coerce'),
                    'UNIT단가': pd.to_numeric(row['발주원가'], errors='coerce')
                })
            
            df_result = pd.DataFrame(final_data)
            
            # 합산 처리 (발주처분류 포함)
            group_cols = ['발주처분류', '수주일자', '납품일자', '발주처코드', '배송코드', '배송지', '상품코드', '상품명', 'UNIT단가']
            df_final = df_result.groupby(group_cols, as_index=False)['UNIT수량'].sum()
            df_final['Total Amount'] = df_final['UNIT수량'] * df_final['UNIT단가']
            
            # 화면 표시
            st.subheader("📊 처리 결과 (발주처별)")
            for channel_name in [v['name'] for v in CHANNELS.values()]:
                filtered_df = df_final[df_final['발주처분류'] == channel_name]
                if not filtered_df.empty:
                    st.write(f"### {channel_name}")
                    st.dataframe(filtered_df)
            
            # 다운로드 파일 생성
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                for channel_key, info in CHANNELS.items():
                    temp_df = df_final[df_final['발주처분류'] == info['name']].drop(columns=['발주처분류'])
                    if not temp_df.empty:
                        temp_df.to_excel(writer, index=False, sheet_name=info['name'])
            
            st.download_button(
                "📥 발주처별 시트 분리 파일 다운로드",
                output.getvalue(),
                f"Integrated_Order_{datetime.now().strftime('%m%d')}.xlsx"
            )
            
        except Exception as e:
            st.error(f"오류 발생: {e}")
