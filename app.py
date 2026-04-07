import streamlit as st
import pandas as pd
import io
from datetime import datetime

# 1. 화면 설정
st.set_page_config(page_title="이마트 계열 수주 자동화", page_icon="🟢", layout="wide")

# --- [날짜 변환 함수] ---
def format_delivery_date(val):
    if pd.isna(val) or str(val).strip() in ["", "0", "nan", "None", "19700101"]:
        return datetime.now().strftime('%Y%m%d')
    try:
        if isinstance(val, datetime):
            return val.strftime('%Y%m%d')
        str_val = str(val).split(' ')[0].split('T')[0]
        clean_val = ''.join(filter(str.isdigit, str_val))
        if len(clean_val) >= 8:
            return clean_val[:8]
        return pd.to_datetime(val).strftime('%Y%m%d')
    except:
        return datetime.now().strftime('%Y%m%d')

@st.cache_data
def load_master_data(file_path):
    try:
        xls = pd.ExcelFile(file_path)
        sheet_map = {s.strip(): s for s in xls.sheet_names}
        
        # 1. 배송코드 시트 처리
        target_sheet = sheet_map.get('배송코드')
        df_tmp = pd.read_excel(xls, sheet_name=target_sheet, dtype=str)
        df_center = None
        for i, row in df_tmp.iterrows():
            if '배송코드' in [str(v).strip() for v in row.values]:
                df_center = pd.read_excel(xls, sheet_name=target_sheet, skiprows=i+1, dtype=str)
                df_center.columns = [str(c).strip() for c in df_tmp.iloc[i]]
                break
        if df_center is None: df_center = df_tmp
        c_to_b = dict(zip(df_center['센터코드'].str.strip(), df_center['배송코드'].str.strip()))
        b_to_n = dict(zip(df_center['배송코드'].str.strip(), df_center.iloc[:, 2].str.strip())) 
        
        # 2. 제품명 시트 처리 (상품명 매칭 로직 강화)
        prod_sheet = sheet_map.get('제품명')
        p_map, n_map = {}, {}
        if prod_sheet:
            df_prod_raw = pd.read_excel(xls, sheet_name=prod_sheet, dtype=str)
            df_prod_raw.columns = [str(c).strip() for c in df_prod_raw.columns]
            
            # 정확한 컬럼명 찾기 (상품코드, ME코드, 상품명)
            col_s = next((c for c in df_prod_raw.columns if '상품코드' in c), None)
            col_m = next((c for c in df_prod_raw.columns if 'ME코드' in c), None)
            col_n = next((c for c in df_prod_raw.columns if '상품명' in c), None)
            
            for _, p_row in df_prod_raw.iterrows():
                s_code = str(p_row.get(col_s, '')).strip()
                if s_code and s_code != 'nan':
                    p_map[s_code] = str(p_row.get(col_m, s_code)).strip()
                    n_map[s_code] = str(p_row.get(col_n, '')).strip()
        
        return {'c_to_b': c_to_b, 'b_to_n': b_to_n, 'products': p_map, 'names': n_map}, None
    except Exception as e:
        return None, str(e)

# --- 메인
