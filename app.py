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
    """시트명 공백 제거 및 유연한 시트 로드"""
    try:
        xls = pd.ExcelFile(file_path)
        # 모든 시트 이름을 가져와 앞뒤 공백을 제거한 맵 생성
        sheet_map = {s.strip(): s for s in xls.sheet_names}
        
        # 1. '센터코드' 시트 로드
        center_sheet_real_name = sheet_map.get('센터코드')
        if not center_sheet_real_name:
            return None, f"'{file_path}' 내에 '센터코드' 시트가 없습니다. (현재 시트목록: {xls.sheet_names})"
        
        # 데이터가 2번째 줄부터 시작할 수 있으므로 skiprows 고려 (파일 특성에 맞게 조정 가능)
        df_center = pd.read_excel(xls, sheet_name=center_sheet_real_name, dtype=str).dropna(how='all')
        
        # 2. '제품명' 시트 로드
        prod_sheet_real_name = sheet_map.get('제품명')
        if not prod_sheet_real_name:
            return None, f"'{file_path}' 내에 '제품명' 시트가 없습니다."
        
        df_prod = pd.read_excel(xls, sheet_name=prod_sheet_real_name, dtype=str).dropna(how='all')
        
        # 매핑 딕셔너리 생성 (Key/Value 공백 제거)
        # 센터코드 -> 배송코드
        c_map = dict(zip(df_center['센터코드'].str.strip(), df_center['배송코드'].str.strip()))
        
        # 상품코드 -> ME코드 / 상품명
        p_map = dict(zip(df_prod['상품코드'].str.strip(), df_prod['ME코드'].str.strip()))
        # 상품명 컬럼은 '이마트 상품명 ' 처럼 공백이 있을 수 있어 인덱스(1번 열) 활용
        n_map = dict(zip(df_prod['상품코드'].str.strip(), df_prod.iloc[:, 1].str.strip()))
        
        return {'centers': c_map, 'products': p_map, 'names': n_map}, None
    except Exception as e:
        return None, str(e)

def get_channel(store_name):
    """점포명 기준 채널 분류"""
    name = str(store_name).upper()
    if 'TR' in name: return 'TRADERS'
    if 'NBR' in name: return 'NOBRAND'
    return 'EMART'

st.title("🛒 이마트·노브랜드·트레이더스 수주 통합 시스템")

# 마스터 로드 (오류 발생 시 화면에 표시)
masters = {}
is_ready = True
for key, info in CHANNELS.items():
    data, err = load_master_data(info['file'])
    if err:
        st.error(f"⚠️ 파일 로드 실패 [{info['file']}]: {err}")
        is_ready = False
    else:
        masters[key] = data

if is_ready:
    st.info("✅ 모든 마스터 파일이 정상적으로 로드되었습니다.")
    uploaded_file = st.file_uploader("일반 주문서 (이마트 로우 데이터) 업로드", type=['xlsx'])
    
    if uploaded_file:
        try:
            df_raw = pd.read_excel(uploaded_file)
            results = []
            
            for _, row in df_raw.iterrows():
                store_name = str(row['점포명'])
                ch = get_channel(store_name)
                
                # [P열 매칭] 센터코드 추출 (인덱스 15)
                raw_center = str(row.iloc[15]).strip()
                delivery_code = masters[ch]['centers'].get(raw_center, "")
                
                # [F열 매칭] 상품코드 추출 (인덱스 5)
                raw_prod = str(row.iloc[5]).strip()
                me_code = masters[ch]['products'].get(raw_prod, raw_prod)
                prod_name = masters[ch]['names'].get(raw_prod, row['상품명'])
                
                results.append({
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
            
            df_res = pd.DataFrame(results)
            
            # [수량 합산] 같은 배송코드 + 같은 ME코드 기준
            group_cols = ['수주일자', '납품일자', '발주처코드', '발주처', '배송코드', '배송지', '상품코드', '상품명', 'UNIT단가']
            df_final = df_res.groupby(group_cols, as_index=False)['UNIT수량'].sum()
            
            # 금액 계산
            df_final['Total Amount'] = df_final['UNIT수량'] * df_final['UNIT단가']
            
            st.subheader("📊 처리 결과 (합산 완료)")
            st.dataframe(df_final)
            
            # 엑셀 변환 및 다운로드
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df_final.to_excel(writer, index=False, sheet_name='Summary(수주업로드용)')
            
            st.download_button(
                label="📥 통합 주문서 다운로드",
                data=output.getvalue(),
                file_name=f"Integrated_Order_{datetime.now().strftime('%m%d')}.xlsx"
            )
            
        except Exception as e:
            st.error(f"데이터 처리 중 오류가 발생했습니다: {e}")
