import streamlit as st
import pandas as pd
import io
from datetime import datetime

st.set_page_config(page_title="이마트 계열 통합 수주 시스템", layout="wide")

# --- [채널 설정] ---
CHANNELS = {
    'TRADERS': {'name': '이마트 트레이더스', 'code': '81011010', 'file': '트레이더스_서식파일_업데이트용.xlsx'},
    'NOBRAND': {'name': '이마트', 'code': '81010000', 'file': '노브랜드_서식파일_업데이트용.xlsx'},
    'EMART': {'name': '이마트', 'code': '81010000', 'file': '이마트_서식파일_업데이트용.xlsx'}
}

@st.cache_data
def load_master_safe(file_path):
    """시트명 공백 제거 및 데이터 로드 (에러 방지용)"""
    try:
        xls = pd.ExcelFile(file_path)
        sheet_names = xls.sheet_names
        
        # 시트명에서 공백 제거 후 매칭 ('센터코드', '제품명' 찾기)
        target_sheets = {s.strip(): s for s in sheet_names}
        
        center_sheet = target_sheets.get('센터코드')
        prod_sheet = target_sheets.get('제품명')
        
        if not center_sheet or not prod_sheet:
            return None, f"'{file_path}' 내에 '센터코드' 또는 '제품명' 시트가 없습니다."
        
        df_center = pd.read_excel(xls, sheet_name=center_sheet, dtype=str)
        df_prod = pd.read_excel(xls, sheet_name=prod_sheet, dtype=str)
        
        # 매핑 딕셔너리 생성
        center_map = dict(zip(df_center['센터코드'].str.strip(), df_center['배송코드'].str.strip()))
        prod_map = dict(zip(df_prod['상품코드'].str.strip(), df_prod['ME코드'].str.strip()))
        name_map = dict(zip(df_prod['상품코드'].str.strip(), df_prod['이마트 상품명 '].str.strip())) # 엑셀상 컬럼명 주의
        
        return {'centers': center_map, 'products': prod_map, 'names': name_map}, None
    except Exception as e:
        return None, str(e)

def get_channel(store_name):
    """점포명 기준 채널 분류"""
    name = str(store_name).upper()
    if 'TR' in name: return 'TRADERS'
    if 'NBR' in name: return 'NOBRAND'
    return 'EMART' # EM 포함 및 기타

st.title("🛒 이마트 계열 수주 자동화 (통합 버전)")

# 1. 마스터 파일 로드
masters = {}
load_success = True
for key, info in CHANNELS.items():
    data, err = load_master_safe(info['file'])
    if err:
        st.error(f"⚠️ {info['file']} 로드 실패: {err}")
        load_success = False
    else:
        masters[key] = data

# 2. 메인 로직
if load_success:
    uploaded_file = st.file_uploader("일반 주문서(이마트 로우 데이터) 업로드", type=['xlsx'])
    
    if uploaded_file:
        try:
            # 로우 데이터 읽기 (보통 첫 행이 제목이므로 header=0 확인)
            df_raw = pd.read_excel(uploaded_file)
            
            final_data = []
            
            for _, row in df_raw.iterrows():
                # [채널 분류]
                store_name = row['점포명']
                ch = get_channel(store_name)
                
                # [배송코드 추출] - P열 (인덱스 15)
                raw_center_code = str(row.iloc[15]).strip()
                delivery_code = masters[ch]['centers'].get(raw_center_code, "")
                
                # [ME코드/제품명 추출] - F열 (인덱스 5)
                raw_prod_code = str(row.iloc[5]).strip()
                me_code = masters[ch]['products'].get(raw_prod_code, raw_prod_code)
                prod_name = masters[ch]['names'].get(raw_prod_code, row['상품명'])
                
                final_data.append({
                    '수주일자': datetime.now().strftime('%Y%m%d'),
                    '납품일자': str(row['납품일자']).replace('-', '')[:8],
                    '발주처코드': CHANNELS[ch]['code'],
                    '발주처': CHANNELS[ch]['name'],
                    '배송코드': delivery_code,
                    '배송지': store_name,
                    '상품코드': me_code,
                    '상품명': prod_name,
                    'UNIT수량': pd.to_numeric(row['수량'], errors='coerce'),
                    'UNIT단가': pd.to_numeric(row['발주원가'], errors='coerce'),
                    '채널구분': ch # 합산용 임시 필드
                })
            
            df_processed = pd.DataFrame(final_data)
            
            # [합산 로직] 같은 배송코드 + 같은 ME코드인 경우 합산
            group_cols = ['수주일자', '납품일자', '발주처코드', '발주처', '배송코드', '배송지', '상품코드', '상품명', 'UNIT단가']
            df_summary = df_processed.groupby(group_cols, as_index=False)['UNIT수량'].sum()
            
            # 금액 계산
            df_summary['금액'] = df_summary['UNIT수량'] * df_summary['UNIT단가']
            df_summary['부가세'] = (df_summary['금액'] * 0.1).astype(int)
            
            st.success("✅ 매칭 및 수량 합산 완료!")
            st.dataframe(df_summary, use_container_width=True)
            
            # 다운로드
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df_summary.to_excel(writer, index=False, sheet_name='수주업로드용')
            
            st.download_button(
                label="📥 결과 엑셀 다운로드",
                data=output.getvalue(),
                file_name=f"EMART_GROUP_{datetime.now().strftime('%m%d')}.xlsx"
            )
            
        except Exception as e:
            st.error(f"처리 중 오류 발생: {e}")
