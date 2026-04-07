import streamlit as st
import pandas as pd
import io
from datetime import datetime

# 화면을 넓게 설정
st.set_page_config(page_title="통합 수주 관리 시스템", layout="wide")

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
    try:
        xls = pd.ExcelFile(file_path)
        sheet_map = {s.strip(): s for s in xls.sheet_names}
        
        # 배송코드 매핑
        df_center = pd.read_excel(xls, sheet_name=sheet_map.get('배송코드'), dtype=str)
        # 헤더 위치 자동 보정 (필요시)
        for i, row in df_center.iterrows():
            if '배송코드' in [str(v).strip() for v in row.values]:
                df_center = pd.read_excel(xls, sheet_name=sheet_map.get('배송코드'), skiprows=i+1, dtype=str)
                df_center.columns = [str(c).strip() for c in pd.read_excel(xls, sheet_name=sheet_map.get('배송코드'), nrows=i+1).iloc[i]]
                break
        
        c_to_b = dict(zip(df_center['센터코드'].str.strip(), df_center['배송코드'].str.strip()))
        b_to_n = dict(zip(df_center['배송코드'].str.strip(), df_center.iloc[:, 2].str.strip()))
        
        # 제품명 매핑
        df_prod = pd.read_excel(xls, sheet_name=sheet_map.get('제품명'), dtype=str)
        p_map = dict(zip(df_prod['상품코드'].str.strip(), df_prod['ME코드'].str.strip()))
        n_map = dict(zip(df_prod['상품코드'].str.strip(), df_prod.iloc[:, 1].str.strip()))
        
        return {'c_to_b': c_to_b, 'b_to_n': b_to_name, 'products': p_map, 'names': n_map}, None
    except Exception as e:
        return None, str(e)

# --- 채널 정의 ---
CHANNELS = {
    'TRADERS': {'name': '이마트 트레이더스', 'code': '81011010', 'file': '트레이더스_서식파일_업데이트용.xlsx'},
    'NOBRAND': {'name': '노브랜드', 'code': '81010000', 'file': '노브랜드_서식파일_업데이트용.xlsx'},
    'EMART': {'name': '이마트', 'code': '81010000', 'file': '이마트_서식파일_업데이트용.xlsx'}
}

st.title("🛒 통합 수주 자동화 (발주처 상세 분류)")

# 마스터 로드
masters = {}
for k, v in CHANNELS.items():
    data, err = load_master_data(v['file'])
    if not err: masters[k] = data

uploaded_file = st.file_uploader("ORDERS 파일 업로드", type=['xlsx'])

if uploaded_file:
    df_raw = pd.read_excel(uploaded_file)
    final_data = []

    for _, row in df_raw.iterrows():
        store_raw = str(row['점포명'])
        
        # [핵심] 점포명에 따른 발주처 분류
        if 'TR' in store_raw.upper(): ch = 'TRADERS'
        elif 'NBR' in store_raw.upper(): ch = 'NOBRAND'
        else: ch = 'EMART'
        
        m = masters.get(ch)
        if not m: continue

        # 매핑 로직
        c_code = str(row['센터코드']).strip()
        d_code = m['c_to_b'].get(c_code, "")
        d_place = m['b_to_n'].get(d_code, store_raw)
        
        p_code = str(row['상품코드']).strip()
        me_code = m['products'].get(p_code, p_code)
        p_name = m['names'].get(p_code, row['상품명'])

        final_data.append({
            '수주일자': datetime.now().strftime('%Y%m%d'),
            '납품일자': format_delivery_date(row.get('센터입하일자', '')),
            '발주처코드': CHANNELS[ch]['code'],
            '발주처': CHANNELS[ch]['name'],  # 여기가 '이마트', '노브랜드', '이마트 트레이더스'로 찍힙니다
            '배송코드': d_code,
            '배송지': d_place,
            '상품코드': me_code,
            '상품명': p_name,
            'UNIT수량': pd.to_numeric(row['수량'], errors='coerce'),
            'UNIT단가': pd.to_numeric(row['발주원가'], errors='coerce')
        })

    df_total = pd.DataFrame(final_data)
    
    # 중복 행 합산 처리
    group_cols = ['수주일자', '납품일자', '발주처코드', '발주처', '배송코드', '배송지', '상품코드', '상품명', 'UNIT단가']
    df_final = df_total.groupby(group_cols, as_index=False)['UNIT수량'].sum()
    df_final['Total Amount'] = df_final['UNIT수량'] * df_final['UNIT단가']

    st.success("✅ 발주처별 통합 분류 완료")
    st.dataframe(df_final)  # 하나의 테이블에 발주처가 섞여서 나옴

    # 엑셀 다운로드
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_final.to_excel(writer, index=False, sheet_name='통합수주데이터')
    st.download_button("📥 통합 결과 다운로드", output.getvalue(), f"Integrated_Order_{datetime.now().strftime('%m%d')}.xlsx")
