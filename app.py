import streamlit as st
import pandas as pd
import io
from datetime import datetime

st.set_page_config(page_title="이마트 계열 수주 자동화", layout="wide")

# --- [날짜 변환 함수] ---
def format_delivery_date(val):
    if pd.isna(val) or str(val).strip() == "":
        return datetime.now().strftime('%Y%m%d')
    try:
        # 엑셀 날짜 형식 및 문자열 모두 처리
        dt = pd.to_datetime(val)
        return dt.strftime('%Y%m%d')
    except:
        # 숫자만 추출하여 8자리 반환
        clean_val = ''.join(filter(str.isdigit, str(val)))
        return clean_val[:8] if len(clean_val) >= 8 else datetime.now().strftime('%Y%m%d')

@st.cache_data
def load_master_data(file_path):
    try:
        xls = pd.ExcelFile(file_path)
        sheet_map = {s.strip(): s for s in xls.sheet_names}
        
        # 1. 배송코드 시트 로드
        target_sheet = sheet_map.get('배송코드')
        if not target_sheet:
            return None, f"'{file_path}'에 '배송코드' 시트가 없습니다."
        
        df_tmp = pd.read_excel(xls, sheet_name=target_sheet, dtype=str)
        df_center = None
        
        for i, row in df_tmp.iterrows():
            row_values = [str(v).strip() for v in row.values]
            if '배송코드' in row_values:
                df_center = pd.read_excel(xls, sheet_name=target_sheet, skiprows=i+1, dtype=str)
                df_center.columns = [str(c).strip() for c in df_tmp.iloc[i]]
                break
        
        if df_center is None: df_center = df_tmp

        c_to_b_map = dict(zip(df_center['센터코드'].str.strip(), df_center['배송코드'].str.strip()))
        b_to_name_map = dict(zip(df_center['배송코드'].str.strip(), df_center.iloc[:, 2].str.strip())) 
        
        # 2. 제품명 시트 로드
        prod_sheet = sheet_map.get('제품명')
        df_prod = pd.read_excel(xls, sheet_name=prod_sheet, dtype=str) if prod_sheet else None
        p_map = dict(zip(df_prod['상품코드'].str.strip(), df_prod['ME코드'].str.strip())) if df_prod is not None else {}
        n_map = dict(zip(df_prod['상품코드'].str.strip(), df_prod.iloc[:, 1].str.strip())) if df_prod is not None else {}
        
        return {'c_to_b': c_to_b_map, 'b_to_n': b_to_name_map, 'products': p_map, 'names': n_map}, None
    except Exception as e:
        return None, str(e)

# --- 메인 실행부 ---
st.title("🛒 최종 수주 자동화 (납품일자=센터입하일자 적용)")

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
    uploaded_file = st.file_uploader("일반 주문서(ORDERS) 업로드", type=['xlsx'])
    
    if uploaded_file:
        try:
            df_raw = pd.read_excel(uploaded_file)
            
            # [수정] 센터입하일자 컬럼 찾기 (오타 및 공백 방지)
            date_col = next((c for c in df_raw.columns if '센터입하일자' in str(c).replace(" ", "")), None)
            
            final_data = []
            for _, row in df_raw.iterrows():
                # 채널 구분
                store_raw = str(row.get('점포명', ''))
                ch = 'EMART'
                if 'TR' in store_raw.upper(): ch = 'TRADERS'
                elif 'NBR' in store_raw.upper(): ch = 'NOBRAND'
                
                # 배송 정보 (P열 = index 15)
                center_code_val = str(row.iloc[15]).strip() if len(row) > 15 else ""
                delivery_code = masters[ch]['c_to_b'].get(center_code_val, "")
                delivery_place = masters[ch]['b_to_n'].get(delivery_code, store_raw) 
                
                # 상품 정보 (F열 = index 5)
                p_code_raw = str(row.iloc[5]).strip() if len(row) > 5 else ""
                me_code = masters[ch]['products'].get(p_code_raw, p_code_raw)
                p_name_raw = row.get('상품명', '')
                prod_name = masters[ch]['names'].get(p_code_raw, p_name_raw)
                
                # [핵심 수정] 납품일자 = 센터입하일자 적용
                raw_date = row[date_col] if date_col else ""
                
                final_data.append({
                    '수주일자': datetime.now().strftime('%Y%m%d'),
                    '납품일자': format_delivery_date(raw_date), # 센터입하일자를 포맷팅하여 삽입
                    '발주처코드': CHANNELS[ch]['code'],
                    '발주처': CHANNELS[ch]['name'],
                    '배송코드': delivery_code,
                    '배송지': delivery_place,
                    '상품코드': me_code,
                    '상품명': prod_name,
                    'UNIT수량': pd.to_numeric(row.get('수량', 0), errors='coerce'),
                    'UNIT단가': pd.to_numeric(row.get('발주원가', 0), errors='coerce')
                })
            
            df_mid = pd.DataFrame(final_data)
            
            # 수량 합산 기준 컬럼
            group_cols = ['수주일자', '납품일자', '발주처코드', '발주처', '배송코드', '배송지', '상품코드', '상품명', 'UNIT단가']
            df_final = df_mid.groupby(group_cols, as_index=False)['UNIT수량'].sum()
            df_final['Total Amount'] = df_final['UNIT수량'] * df_final['UNIT단가']
            
            st.success("✅ '납품일자 = 센터입하일자' 적용 및 데이터 통합 완료")
            st.dataframe(df_final)
            
            # 다운로드
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df_final.to_excel(writer, index=False, sheet_name='수주업로드용')
            st.download_button("📥 최종 파일 다운로드", output.getvalue(), f"Final_Order_{datetime.now().strftime('%m%d')}.xlsx")
            
        except Exception as e:
            st.error(f"처리 중 오류 발생: {e}")
