import streamlit as st
import pandas as pd
import io
from datetime import datetime

st.set_page_config(page_title="이마트 계열 수주 자동화", layout="wide")

# --- [채널 설정] ---
CHANNELS = {
    'TRADERS': {'name': '이마트 트레이더스', 'code': '81011010', 'file': '트레이더스_서식파일_업데이트용.xlsx'},
    'NOBRAND': {'name': '이마트', 'code': '81010000', 'file': '노브랜드_서식파일_업데이트용.xlsx'},
    'EMART': {'name': '이마트', 'code': '81010000', 'file': '이마트_서식파일_업데이트용.xlsx'}
}

@st.cache_data
def load_master_data(file_path):
    """'배송코드' 시트 및 '제품명' 시트 로드 (공백 및 헤더 위치 보정)"""
    try:
        xls = pd.ExcelFile(file_path)
        # 모든 시트 이름의 공백을 제거하여 매핑
        sheet_map = {s.strip(): s for s in xls.sheet_names}
        
        # 1. '배송코드' 시트 로드 (구 '센터코드')
        target_sheet = sheet_map.get('배송코드')
        if not target_sheet:
            return None, f"'{file_path}' 내에 '배송코드' 시트가 없습니다."
        
        # 헤더가 2행(A2)에 있을 경우를 대비해 '배송코드' 글자가 있는 행을 찾음
        df_tmp = pd.read_excel(xls, sheet_name=target_sheet, dtype=str)
        df_center = None
        for i, row in df_tmp.iterrows():
            if '배송코드' in row.values or '센터코드' in row.values:
                df_center = pd.read_excel(xls, sheet_name=target_sheet, skiprows=i+1, dtype=str)
                # 컬럼명 재설정 (skiprows 사용 시 그 다음 행이 컬럼이 됨)
                df_center.columns = df_tmp.iloc[i].str.strip() 
                break
        if df_center is None: df_center = df_tmp # 못 찾으면 기본 로드

        # 2. '제품명' 시트 로드
        prod_sheet = sheet_map.get('제품명')
        df_prod = pd.read_excel(xls, sheet_name=prod_sheet, dtype=str) if prod_sheet else None
        
        # 매핑 딕셔너리 구성
        # 로우 데이터의 P열(센터코드) 값으로 마스터의 '배송코드'를 가져와야 함
        # 주의: 마스터 파일의 컬럼명이 '센터코드'인지 확인 필요
        c_map = dict(zip(df_center['센터코드'].str.strip(), df_center['배송코드'].str.strip()))
        
        p_map = {}
        n_map = {}
        if df_prod is not None:
            p_map = dict(zip(df_prod['상품코드'].str.strip(), df_prod['ME코드'].str.strip()))
            n_map = dict(zip(df_prod['상품코드'].str.strip(), df_prod.iloc[:, 1].str.strip()))
            
        return {'centers': c_map, 'products': p_map, 'names': n_map}, None
    except Exception as e:
        return None, str(e)

def identify_channel(store_name):
    name = str(store_name).upper()
    if 'TR' in name: return 'TRADERS'
    if 'NBR' in name: return 'NOBRAND'
    return 'EMART'

st.title("🛒 통합 수주 자동화 시스템 (시트명: 배송코드)")

# 마스터 로드
masters = {}
ready = True
for key, info in CHANNELS.items():
    data, err = load_master_data(info['file'])
    if err:
        st.error(f"❌ {info['file']} 로드 실패: {err}")
        ready = False
    else:
        masters[key] = data

if ready:
    uploaded_file = st.file_uploader("일반 주문서(이마트 로우 데이터) 업로드", type=['xlsx'])
    
    if uploaded_file:
        try:
            df_raw = pd.read_excel(uploaded_file)
            processed_data = []
            
            for _, row in df_raw.iterrows():
                store_name = str(row['점포명'])
                ch = identify_channel(store_name)
                
                # [매칭 1] P열 (센터코드, index 15) -> 배송코드 추출
                raw_center_val = str(row.iloc[15]).strip()
                delivery_code = masters[ch]['centers'].get(raw_center_val, "")
                
                # [매칭 2] F열 (상품코드, index 5) -> ME코드 및 상품명 추출
                raw_prod_val = str(row.iloc[5]).strip()
                me_code = masters[ch]['products'].get(raw_prod_val, raw_prod_val)
                prod_name = masters[ch]['names'].get(raw_prod_val, row['상품명'])
                
                processed_data.append({
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
            
            df_mid = pd.DataFrame(processed_data)
            
            # --- [수량 합산 로직] ---
            # 동일 배송코드 + 동일 ME코드인 경우 수량 합산
            group_cols = ['수주일자', '납품일자', '발주처코드', '발주처', '배송코드', '배송지', '상품코드', '상품명', 'UNIT단가']
            df_final = df_mid.groupby(group_cols, as_index=False)['UNIT수량'].sum()
            
            # 합계 금액
            df_final['Total Amount'] = df_final['UNIT수량'] * df_final['UNIT단가']
            
            st.success("✅ 매칭 및 수량 합산이 완료되었습니다.")
            st.dataframe(df_final)
            
            # 다운로드
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df_final.to_excel(writer, index=False, sheet_name='수주업로드용')
            
            st.download_button(
                label="📥 결과 엑셀 다운로드",
                data=output.getvalue(),
                file_name=f"Result_{datetime.now().strftime('%m%d')}.xlsx"
            )
            
        except Exception as e:
            st.error(f"처리 오류: {e}")
