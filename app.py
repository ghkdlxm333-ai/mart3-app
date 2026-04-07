import streamlit as st
import pandas as pd
import io
from datetime import datetime

st.set_page_config(page_title="이마트 계열 수주 자동화", page_icon="🟢", layout="wide")

# --- [설정 및 상수] ---
CHANNELS = {
    'TRADERS': {'name': '이마트 트레이더스', 'code': '81011010', 'file': '트레이더스_서식파일_업데이트용.xlsx'},
    'NOBRAND': {'name': '이마트', 'code': '81010000', 'file': '노브랜드_서식파일_업데이트용.xlsx'},
    'EMART': {'name': '이마트', 'code': '81010000', 'file': '이마트_서식파일_업데이트용.xlsx'}
}

@st.cache_data
def load_all_masters():
    """각 채널별 마스터 파일의 센터코드와 제품명 정보를 로드"""
    master_data = {}
    for key, info in CHANNELS.items():
        try:
            # 센터코드 매핑용 (P열 매칭용)
            df_center = pd.read_excel(info['file'], sheet_name='센터코드', dtype=str)
            # 제품명 매핑용 (F열 매칭용)
            df_prod = pd.read_excel(info['file'], sheet_name='제품명', dtype=str)
            
            master_data[key] = {
                'centers': dict(zip(df_center['센터코드'], df_center['배송코드'])),
                'products': dict(zip(df_prod['상품코드'], df_prod['ME코드'])),
                'prod_names': dict(zip(df_prod['상품코드'], df_prod['상품명']))
            }
        except Exception as e:
            st.error(f"{info['file']} 로드 중 오류 발생: {e}")
    return master_data

def identify_channel(store_name):
    """점포명에 따른 채널 구분 로직"""
    store_name = str(store_name).upper()
    if 'TR' in store_name:
        return 'TRADERS'
    elif 'NBR' in store_name:
        return 'NOBRAND'
    else:
        # EM 포함 또는 그 외 나머지 모두 이마트
        return 'EMART'

st.title("🛒🟢 이마트·노브랜드·트레이더스 수주 자동화")

masters = load_all_masters()

uploaded_file = st.file_uploader("일반 주문서(이마트 로우 데이터)를 업로드하세요", type=['xlsx'])

if uploaded_file and masters:
    try:
        # 일반 주문서 로드
        df_raw = pd.read_excel(uploaded_file)
        
        processed_rows = []
        
        for _, row in df_raw.iterrows():
            # 1. 채널 구분 (점포명 기준)
            store_name = str(row.get('점포명', ''))
            channel_key = identify_channel(store_name)
            channel_info = CHANNELS[channel_key]
            master_info = masters[channel_key]
            
            # 2. 배송코드 매칭 (P열 센터코드 기준)
            # 'P열'은 인덱스로 15번입니다 (0부터 시작 시)
            center_code = str(row.iloc[15]).strip()
            delivery_code = master_info['centers'].get(center_code, "")
            
            # 3. ME코드 매칭 (F열 상품코드 기준)
            # 'F열'은 인덱스로 5번입니다
            prod_code = str(row.iloc[5]).strip()
            me_code = master_info['products'].get(prod_code, "")
            prod_name = master_info['prod_names'].get(prod_code, row.get('상품명', ''))
            
            # 4. 데이터 저장
            processed_rows.append({
                '수주일자': datetime.now().strftime('%Y%m%d'),
                '납품일자': str(row.get('납품일자', '')).replace('-', '')[:8],
                '발주처코드': channel_info['code'],
                '발주처': channel_info['name'],
                '배송코드': delivery_code,
                '배송지': store_name,
                '상품코드': me_code if me_code else prod_code,
                '상품명': prod_name,
                'UNIT수량': pd.to_numeric(row.get('수량', 0), errors='coerce'),
                'UNIT단가': pd.to_numeric(row.get('발주원가', 0), errors='coerce'),
                '채널': channel_key
            })
            
        if processed_rows:
            df_temp = pd.DataFrame(processed_rows)
            
            # 5. 합산 로직 (채널별, 배송코드별, ME코드별 수량 합산)
            group_cols = ['수주일자', '납품일자', '발주처코드', '발주처', '배송코드', '배송지', '상품코드', '상품명', 'UNIT단가', '채널']
            df_final = df_temp.groupby(group_cols, as_index=False)['UNIT수량'].sum()
            
            # 금액 계산
            df_final['금액'] = df_final['UNIT수량'] * df_final['UNIT단가']
            df_final['부가세'] = (df_final['금액'] * 0.1).astype(int)
            
            st.success(f"✅ 분석 완료 (총 {len(df_final)} 건)")
            st.dataframe(df_final, use_container_width=True)
            
            # 다운로드 버튼
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df_final.to_excel(writer, index=False, sheet_name='수주합산')
            
            st.download_button(
                label="📥 통합 수주 결과 다운로드",
                data=output.getvalue(),
                file_name=f"Integrated_Order_{datetime.now().strftime('%m%d')}.xlsx"
            )
            
    except Exception as e:
        st.error(f"처리 중 오류가 발생했습니다: {e}")
