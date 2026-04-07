import streamlit as st
import pandas as pd
import io
from datetime import datetime

st.set_page_config(page_title="이마트/NB/TR 수주 자동화", layout="wide")

# --- [채널 설정] ---
CHANNELS = {
    'TRADERS': {'name': '이마트 트레이더스', 'code': '81011010', 'file': '트레이더스_서식파일_업데이트용.xlsx'},
    'NOBRAND': {'name': '이마트', 'code': '81010000', 'file': '노브랜드_서식파일_업데이트용.xlsx'},
    'EMART': {'name': '이마트', 'code': '81010000', 'file': '이마트_서식파일_업데이트용.xlsx'}
}

@st.cache_data
def load_master_data(file_path):
    """시트명 공백 제거 및 정확한 매핑 데이터 추출"""
    try:
        xls = pd.ExcelFile(file_path)
        sheet_names = {s.strip(): s for s in xls.sheet_names}
        
        # 1. 센터코드 매핑 (배송코드 추출용)
        df_center = pd.read_excel(xls, sheet_name=sheet_names.get('센터코드'), dtype=str)
        # 2. 제품명 매핑 (ME코드 추출용)
        df_prod = pd.read_excel(xls, sheet_name=sheet_names.get('제품명'), dtype=str)
        
        # 딕셔너리 생성 (Key의 공백 제거 필수)
        center_map = dict(zip(df_center['센터코드'].str.strip(), df_center['배송코드'].str.strip()))
        
        # 상품코드(F열) -> ME코드 매핑
        # 파일에 따라 컬럼명이 '이마트 상품명 '처럼 공백이 있을 수 있으므로 인덱스나 strip 활용
        prod_map = dict(zip(df_prod['상품코드'].str.strip(), df_prod['ME코드'].str.strip()))
        name_map = dict(zip(df_prod['상품코드'].str.strip(), df_prod.iloc[:, 1].str.strip())) # 2번째 컬럼을 상품명으로 가정
        
        return {'centers': center_map, 'products': prod_map, 'names': name_map}, None
    except Exception as e:
        return None, str(e)

def identify_channel(store_name):
    """점포명 기준 채널 분류 로직"""
    name = str(store_name).upper()
    if 'TR' in name: return 'TRADERS'
    if 'NBR' in name: return 'NOBRAND'
    return 'EMART'

st.title("🛒 이마트 계열 수주 자동화 시스템")

# 마스터 데이터 로드
masters = {}
status_ok = True
for key, info in CHANNELS.items():
    data, err = load_master_data(info['file'])
    if err:
        st.error(f"❌ {info['file']} 로드 오류: {err}")
        status_ok = False
    else:
        masters[key] = data

if status_ok:
    uploaded_file = st.file_uploader("이마트 로우 데이터(일반 주문서)를 업로드하세요", type=['xlsx'])
    
    if uploaded_file:
        try:
            df_raw = pd.read_excel(uploaded_file)
            temp_results = []
            
            for _, row in df_raw.iterrows():
                store_name = str(row['점포명'])
                ch = identify_channel(store_name)
                
                # P열 (센터코드, 인덱스 15) 매칭
                p_col_val = str(row.iloc[15]).strip()
                delivery_code = masters[ch]['centers'].get(p_col_val, "")
                
                # F열 (상품코드, 인덱스 5) 매칭
                f_col_val = str(row.iloc[5]).strip()
                me_code = masters[ch]['products'].get(f_col_val, f_col_val)
                prod_name = masters[ch]['names'].get(f_col_val, row['상품명'])
                
                temp_results.append({
                    '수주일자': datetime.now().strftime('%Y%m%d'),
                    '납품일자': str(row['납품일자']).replace('-', '')[:8],
                    '발주처코드': CHANNELS[ch]['code'],
                    '발주처': CHANNELS[ch]['name'],
                    '배송코드': delivery_code,
                    '배송지': store_name,
                    '상품코드': me_code,
                    '상품명': prod_name,
                    'UNIT수량': pd.to_numeric(row['수량'], errors='coerce'),
                    'UNIT단가': pd.to_numeric(row['발주원가'], errors='coerce')
                })
            
            df_processed = pd.DataFrame(temp_results)
            
            # --- [합산 로직] ---
            # 배송코드와 상품코드(ME코드)가 같으면 수량 합산
            group_keys = ['수주일자', '납품일자', '발주처코드', '발주처', '배송코드', '배송지', '상품코드', '상품명', 'UNIT단가']
            df_final = df_processed.groupby(group_keys, as_index=False)['UNIT수량'].sum()
            
            # 금액 계산
            df_final['금액'] = df_final['UNIT수량'] * df_final['UNIT단가']
            df_final['부가세'] = (df_final['금액'] * 0.1).astype(int)
            
            st.success("✅ 수주 데이터 생성 및 합산 완료!")
            st.dataframe(df_final)
            
            # 엑셀 다운로드
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df_final.to_excel(writer, index=False, sheet_name='수주업로드')
            
            st.download_button(
                label="📥 결과 파일 다운로드",
                data=output.getvalue(),
                file_name=f"EMART_ORDER_{datetime.now().strftime('%m%d')}.xlsx"
            )
            
        except Exception as e:
            st.error(f"데이터 처리 중 오류 발생: {e}")
