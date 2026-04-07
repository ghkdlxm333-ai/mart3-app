import streamlit as st
import pandas as pd
import io
from datetime import datetime

# --- [날짜 변환 함수: 센터입하일자 전용] ---
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
    """'배송코드' 시트에서 배송코드와 배송지(센터명) 매핑"""
    try:
        xls = pd.ExcelFile(file_path)
        sheet_map = {s.strip(): s for s in xls.sheet_names}
        
        # 1. 배송코드 시트 로드
        target_sheet = sheet_map.get('배송코드')
        if not target_sheet:
            return None, f"'{file_path}'에 '배송코드' 시트가 없습니다."
        
        df_tmp = pd.read_excel(xls, sheet_name=target_sheet, dtype=str)
        df_center = None
        
        # '배송코드' 헤더 위치 찾기 (A2 등 시작 위치 보정)
        for i, row in df_tmp.iterrows():
            row_values = [str(v).strip() for v in row.values]
            if '배송코드' in row_values:
                df_center = pd.read_excel(xls, sheet_name=target_sheet, skiprows=i+1, dtype=str)
                df_center.columns = [str(c).strip() for c in df_tmp.iloc[i]]
                break
        
        if df_center is None: df_center = df_tmp

        # 매핑 딕셔너리 구성
        # 센터코드(P열값) -> 배송코드 매핑
        c_to_b_map = dict(zip(df_center['센터코드'].str.strip(), df_center['배송코드'].str.strip()))
        
        # 배송코드 -> 배송지(옆에 있는 지명/센터명) 매핑
        # 주의: 마스터 파일 내 '배송코드' 바로 옆 컬럼이 배송지(지명)여야 함
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
st.title("🛒 최종 수주 자동화 (배송지 매핑 보강)")

CHANNELS = {
    'TRADERS': {'name': '이마트 트레이더스', 'code': '81011010', 'file': '트레이더스_서식파일_업데이트용.xlsx'},
    'NOBRAND': {'name': '이마트', 'code': '81010000', 'file': '노브랜드_서식파일_업데이트용.xlsx'},
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
            
            # 센터입하일자 컬럼 확인
            date_col = next((c for c in df_raw.columns if '센터입하일자' in str(c)), '센터입하일자')
            
            final_data = []
            for _, row in df_raw.iterrows():
                # 채널 구분
                store_raw = str(row['점포명'])
                ch = 'EMART'
                if 'TR' in store_raw.upper(): ch = 'TRADERS'
                elif 'NBR' in store_raw.upper(): ch = 'NOBRAND'
                
                # [로직 1] 배송코드 및 배송지 추출
                center_code_val = str(row.iloc[15]).strip() # P열
                delivery_code = masters[ch]['c_to_b'].get(center_code_val, "")
                
                # 마스터의 배송코드 옆에 있는 '지명'을 배송지로 사용
                delivery_place = masters[ch]['b_to_n'].get(delivery_code, store_raw) 
                
                # [로직 2] 상품 정보
                prod_code_val = str(row.iloc[5]).strip() # F열
                me_code = masters[ch]['products'].get(prod_code_val, prod_code_val)
                prod_name = masters[ch]['names'].get(prod_code_val, row['상품명'])
                
                final_data.append({
                    '수주일자': datetime.now().strftime('%Y%m%d'),
                    '납품일자': format_delivery_date(row[date_col]), # YYYYMMDD 변환
                    '발주처코드': CHANNELS[ch]['code'],
                    '발주처': CHANNELS[ch]['name'],
                    '배송코드': delivery_code,
                    '배송지': delivery_place, # 마스터 데이터 기준
                    '상품코드': me_code,
                    '상품명': prod_name,
                    'UNIT수량': pd.to_numeric(row['수량'], errors='coerce'),
                    'UNIT단가': pd.to_numeric(row['발주원가'], errors='coerce')
                })
            
            df_mid = pd.DataFrame(final_data)
            
            # 수량 합산
            group_cols = ['수주일자', '납품일자', '발주처코드', '발주처', '배송코드', '배송지', '상품코드', '상품명', 'UNIT단가']
            df_final = df_mid.groupby(group_cols, as_index=False)['UNIT수량'].sum()
            df_final['Total Amount'] = df_final['UNIT수량'] * df_final['UNIT단가']
            
            st.success("✅ 센터입하일자 및 마스터 배송지 매핑 완료!")
            st.dataframe(df_final)
            
            # 다운로드
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df_final.to_excel(writer, index=False, sheet_name='수주업로드용')
            st.download_button("📥 최종 파일 다운로드", output.getvalue(), f"Final_Order_{datetime.now().strftime('%m%d')}.xlsx")
            
        except Exception as e:
            st.error(f"처리 중 오류: {e}")
