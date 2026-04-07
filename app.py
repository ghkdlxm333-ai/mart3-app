import streamlit as st
import pandas as pd
import io
from datetime import datetime

st.set_page_config(page_title="이마트 계열 수주 자동화", layout="wide")

# --- [날짜 변환 함수: 센터입하일자를 YYYYMMDD로 강제 변환] ---
def format_delivery_date(val):
    if pd.isna(val) or str(val).strip() == "" or str(val).strip() == "0":
        return datetime.now().strftime('%Y%m%d') # 값이 없으면 오늘 날짜
    
    try:
        # 1. 엑셀 날짜 형식(datetime)인 경우
        if isinstance(val, datetime):
            return val.strftime('%Y%m%d')
        
        # 2. 문자열이나 숫자인 경우 (하이픈, 슬래시 등 제거)
        str_val = str(val).split(' ')[0] # 시간 정보 포함 시 날짜만 추출
        clean_val = ''.join(filter(str.isdigit, str_val))
        
        # 8자리 이상이면 앞의 8자리만 사용 (YYYYMMDD)
        if len(clean_val) >= 8:
            return clean_val[:8]
        else:
            # 8자리가 안 되면 다시 시도하거나 오늘 날짜
            dt = pd.to_datetime(val)
            return dt.strftime('%Y%m%d')
    except:
        return datetime.now().strftime('%Y%m%d')

@st.cache_data
def load_master_data(file_path):
    try:
        xls = pd.ExcelFile(file_path)
        sheet_map = {s.strip(): s for s in xls.sheet_names}
        
        target_sheet = sheet_map.get('배송코드')
        if not target_sheet: return None, f"'{file_path}'에 '배송코드' 시트가 없습니다."
        
        df_tmp = pd.read_excel(xls, sheet_name=target_sheet, dtype=str)
        df_center = None
        for i, row in df_tmp.iterrows():
            if '배송코드' in [str(v).strip() for v in row.values]:
                df_center = pd.read_excel(xls, sheet_name=target_sheet, skiprows=i+1, dtype=str)
                df_center.columns = [str(c).strip() for c in df_tmp.iloc[i]]
                break
        if df_center is None: df_center = df_tmp

        c_to_b = dict(zip(df_center['센터코드'].str.strip(), df_center['배송코드'].str.strip()))
        # 마스터 파일의 '배송코드' 바로 옆(3번째 컬럼)을 배송지명으로 사용
        b_to_n = dict(zip(df_center['배송코드'].str.strip(), df_center.iloc[:, 2].str.strip())) 
        
        prod_sheet = sheet_map.get('제품명')
        df_prod = pd.read_excel(xls, sheet_name=prod_sheet, dtype=str) if prod_sheet else None
        p_map = dict(zip(df_prod['상품코드'].str.strip(), df_prod['ME코드'].str.strip())) if df_prod is not None else {}
        n_map = dict(zip(df_prod['상품코드'].str.strip(), df_prod.iloc[:, 1].str.strip())) if df_prod is not None else {}
        
        return {'c_to_b': c_to_b, 'b_to_n': b_to_n, 'products': p_map, 'names': n_map}, None
    except Exception as e:
        return None, str(e)

# --- 실행부 ---
st.title("🛒 통합 수주 자동화 (납품일자 형식 고정)")

CHANNELS = {
    'TRADERS': {'name': '이마트 트레이더스', 'code': '81011010', 'file': '트레이더스_서식파일_업데이트용.xlsx'},
    'NOBRAND': {'name': '노브랜드', 'code': '81010000', 'file': '노브랜드_서식파일_업데이트용.xlsx'},
    'EMART': {'name': '이마트', 'code': '81010000', 'file': '이마트_서식파일_업데이트용.xlsx'}
}

masters = {}
status_ok = True
for k, v in CHANNELS.items():
    data, err = load_master_data(v['file'])
    if err:
        st.error(f"❌ {v['file']} 확인 필요: {err}")
        status_ok = False
    else:
        masters[k] = data

if status_ok:
    uploaded_file = st.file_uploader("ORDERS 파일 업로드", type=['xlsx'])
    
    if uploaded_file:
        try:
            df_raw = pd.read_excel(uploaded_file)
            # '센터입하일자' 컬럼 정확히 찾기
            date_col = next((c for c in df_raw.columns if '센터입하일자' in str(c)), None)
            
            final_data = []
            for _, row in df_raw.iterrows():
                store_raw = str(row.get('점포명', ''))
                ch = 'TRADERS' if 'TR' in store_raw.upper() else ('NOBRAND' if 'NBR' in store_raw.upper() else 'EMART')
                
                m = masters[ch]
                # P열(인덱스 15) 센터코드
                c_val = str(row.iloc[15]).strip() if len(row) > 15 else ""
                d_code = m['c_to_b'].get(c_val, "")
                d_place = m['b_to_n'].get(d_code, store_raw)
                
                # F열(인덱스 5) 상품코드
                p_val = str(row.iloc[5]).strip() if len(row) > 5 else ""
                me_code = m['products'].get(p_val, p_val)
                p_name = m['names'].get(p_val, str(row.get('상품명', '')))

                final_data.append({
                    '수주일자': datetime.now().strftime('%Y%m%d'),
                    '납품일자': format_delivery_date(row[date_col]) if date_col else "",
                    '발주처코드': CHANNELS[ch]['code'],
                    '발주처': CHANNELS[ch]['name'],
                    '배송코드': d_code,
                    '배송지': d_place,
                    '상품코드': me_code,
                    '상품명': p_name,
                    'UNIT수량': pd.to_numeric(row.get('수량', 0), errors='coerce'),
                    'UNIT단가': pd.to_numeric(row.get('발주원가', 0), errors='coerce')
                })
            
            df_mid = pd.DataFrame(final_data)
            group_cols = ['수주일자', '납품일자', '발주처코드', '발주처', '배송코드', '배송지', '상품코드', '상품명', 'UNIT단가']
            df_final = df_mid.groupby(group_cols, as_index=False)['UNIT수량'].sum()
            df_final['Total Amount'] = df_final['UNIT수량'] * df_final['UNIT단가']
            
            # 납품일자 컬럼을 문자열(String)로 강제 지정하여 숫자로 변하는 것 방지
            df_final['납품일자'] = df_final['납품일자'].astype(str)

            st.success("✅ 납품일자 형식이 YYYYMMDD로 고정되었습니다.")
            st.dataframe(df_final)
            
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df_final.to_excel(writer, index=False, sheet_name='수주업로드용')
            st.download_button("📥 수정된 파일 다운로드", output.getvalue
