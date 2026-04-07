import streamlit as st
import pandas as pd
import io
from datetime import datetime

st.set_page_config(page_title="이마트/NB/TR 수주 시스템", layout="wide")

# --- [채널 설정] ---
CHANNELS = {
    'TRADERS': {'name': '이마트 트레이더스', 'code': '81011010', 'file': '트레이더스_서식파일_업데이트용.xlsx'},
    'NOBRAND': {'name': '이마트', 'code': '81010000', 'file': '노브랜드_서식파일_업데이트용.xlsx'},
    'EMART': {'name': '이마트', 'code': '81010000', 'file': '이마트_서식파일_업데이트용.xlsx'}
}

@st.cache_data
def load_master_data(file_path):
    """시트명 공백 제거 및 유연한 데이터 로드"""
    try:
        xls = pd.ExcelFile(file_path)
        # 모든 시트 이름을 가져와 앞뒤 공백을 제거한 매핑 생성
        actual_sheets = {s.strip(): s for s in xls.sheet_names}
        
        # 1. '센터코드' 시트 로드 (공백 무시)
        target_center = actual_sheets.get('센터코드')
        if not target_center:
            return None, f"'{file_path}'에 '센터코드' 시트가 없습니다. (확인된 시트: {xls.sheet_names})"
        df_center = pd.read_excel(xls, sheet_name=target_center, dtype=str)
        
        # 2. '제품명' 시트 로드 (공백 무시)
        target_prod = actual_sheets.get('제품명')
        if not target_prod:
            return None, f"'{file_path}'에 '제품명' 시트가 없습니다."
        df_prod = pd.read_excel(xls, sheet_name=target_prod, dtype=str)
        
        # 매핑 딕셔너리 생성 (데이터 내부 공백도 제거)
        center_map = dict(zip(df_center['센터코드'].str.strip(), df_center['배송코드'].str.strip()))
        prod_map = dict(zip(df_prod['상품코드'].str.strip(), df_prod['ME코드'].str.strip()))
        # 상품명 컬럼은 파일마다 다를 수 있어 인덱스(보통 2번째 열) 활용 권장
        name_map = dict(zip(df_prod['상품코드'].str.strip(), df_prod.iloc[:, 1].str.strip()))
        
        return {'centers': center_map, 'products': prod_map, 'names': name_map}, None
    except Exception as e:
        return None, str(e)

def get_channel(store_name):
    """점포명 기준 채널 분류"""
    name = str(store_name).upper()
    if 'TR' in name: return 'TRADERS'
    if 'NBR' in name: return 'NOBRAND'
    return 'EMART'

st.title("🛒 이마트 계열 수주 자동화 시스템")

# 마스터 파일 사전 로드
masters = {}
success = True
for key, info in CHANNELS.items():
    data, err = load_master_data(info['file'])
    if err:
        st.error(f"❌ {info['file']} 로드 오류: {err}")
        success = False
    else:
        masters[key] = data

if success:
    st.success("✅ 모든 업데이트용 마스터 파일 로드 완료")
    uploaded_file = st.file_uploader("일반 주문서(이마트 로우 데이터) 업로드", type=['xlsx'])
    
    if uploaded_file:
        try:
            df_raw = pd.read_excel(uploaded_file)
            final_rows = []
            
            for _, row in df_raw.iterrows():
                store_name = str(row['점포명'])
                ch = get_channel(store_name)
                
                # [P열 매칭] 센터코드 (인덱스 15)
                raw_center_code = str(row.iloc[15]).strip()
                delivery_code = masters[ch]['centers'].get(raw_center_code, "")
                
                # [F열 매칭] 상품코드 (인덱스 5)
                raw_prod_code = str(row.iloc[5]).strip()
                me_code = masters[ch]['products'].get(raw_prod_code, raw_prod_code)
                prod_name = masters[ch]['names'].get(raw_prod_code, row['상품명'])
                
                final_rows.append({
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
            
            df_processed = pd.DataFrame(final_rows)
            
            # --- [수량 합산 로직] ---
            # 배송코드와 상품코드가 같으면 하나의 행으로 합침
            group_cols = ['수주일자', '납품일자', '발주처코드', '발주처', '배송코드', '배송지', '상품코드', '상품명', 'UNIT단가']
            df_final = df_processed.groupby(group_cols, as_index=False)['UNIT수량'].sum()
            
            # 합계 금액 계산
            df_final['Total Amount'] = df_final['UNIT수량'] * df_final['UNIT단가']
            
            st.subheader("📋 처리 및 합산 결과")
            st.dataframe(df_final)
            
            # 다운로드 버튼
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df_final.to_excel(writer, index=False, sheet_name='수주업로드용')
            
            st.download_button(
                label="📥 결과 엑셀 다운로드",
                data=output.getvalue(),
                file_name=f"Order_Summary_{datetime.now().strftime('%m%d')}.xlsx"
            )
            
        except Exception as e:
            st.error(f"데이터 처리 중 오류: {e}")
