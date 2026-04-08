import streamlit as st
import pandas as pd
import io
from datetime import datetime

# --- [날짜 변환 함수] ---
def format_to_yyyymmdd(val):
    if pd.isna(val) or str(val).strip() == "":
        return datetime.now().strftime('%Y%m%d')
    try:
        # datetime 객체나 하이픈(-) 포함 문자열 처리
        dt = pd.to_datetime(val)
        return dt.strftime('%Y%m%d')
    except:
        # 숫자만 남기고 8자리 추출
        clean_val = ''.join(filter(str.isdigit, str(val)))
        return clean_val[:8] if len(clean_val) >= 8 else datetime.now().strftime('%Y%m%d')

@st.cache_data
def load_master_data(file_path):
    """시트명 '배송코드' 및 '제품명' 로드 (헤더 위치 자동 보정)"""
    try:
        xls = pd.ExcelFile(file_path)
        sheet_map = {s.strip(): s for s in xls.sheet_names}
        
        # 1. 배송코드 시트 로드
        target_sheet = sheet_map.get('배송코드')
        if not target_sheet:
            return None, f"'{file_path}'에 '배송코드' 시트가 없습니다."
        
        df_tmp = pd.read_excel(xls, sheet_name=target_sheet, dtype=str)
        df_center = df_tmp
        # '배송코드'나 '센터코드'라는 글자가 있는 행을 찾아 헤더로 설정
        for i, row in df_tmp.iterrows():
            row_values = [str(v).strip() for v in row.values]
            if '배송코드' in row_values or '센터코드' in row_values:
                df_center = pd.read_excel(xls, sheet_name=target_sheet, skiprows=i+1, dtype=str)
                df_center.columns = [str(c).strip() for c in df_tmp.iloc[i]]
                break
        
        # 2. 제품명 시트 로드
        prod_sheet = sheet_map.get('제품명')
        df_prod = pd.read_excel(xls, sheet_name=prod_sheet, dtype=str) if prod_sheet else None
        
        # 매핑 생성
        c_map = dict(zip(df_center['센터코드'].str.strip(), df_center['배송코드'].str.strip()))
        p_map = dict(zip(df_prod['상품코드'].str.strip(), df_prod['ME코드'].str.strip())) if df_prod is not None else {}
        n_map = dict(zip(df_prod['상품코드'].str.strip(), df_prod.iloc[:, 1].str.strip())) if df_prod is not None else {}
        
        return {'centers': c_map, 'products': p_map, 'names': n_map}, None
    except Exception as e:
        return None, str(e)

# --- 메인 앱 ---
st.title("🛒 통합 수주 자동화 (센터입하일자 기준)")

CHANNELS = {
    'TRADERS': {'name': '이마트 트레이더스', 'code': '81011010', 'file': '트레이더스_서식파일_업데이트용.xlsx'},
    'NOBRAND': {'name': '이마트', 'code': '81010000', 'file': '노브랜드_서식파일_업데이트용.xlsx'},
    'EMART': {'name': '이마트', 'code': '81010000', 'file': '이마트_서식파일_업데이트용.xlsx'}
}

# 마스터 로드
masters = {}
is_ok = True
for k, v in CHANNELS.items():
    data, err = load_master_data(v['file'])
    if err:
        st.error(f"❌ {v['file']} 파일 확인 필요: {err}")
        is_ok = False
    else:
        masters[k] = data

if is_ok:
    uploaded_file = st.file_uploader("일반 주문서(Raw Data) 업로드", type=['xlsx'])
    
    if uploaded_file:
        try:
            df_raw = pd.read_excel(uploaded_file)
            
            # [수정] 납품일자 대용으로 '센터입하일자' 사용
            # 만약 센터입하일자 컬럼도 없으면 점입점일자나 발주일자를 찾음
            date_col = next((c for c in df_raw.columns if '센터입하일자' in str(c) or '납품일자' in str(c)), None)
            if not date_col:
                date_col = next((c for c in df_raw.columns if '점입점일자' in str(c) or '발주일자' in str(c)), df_raw.columns[0])

            final_list = []
            for _, row in df_raw.iterrows():
                store = str(row['점포명'])
                ch = 'EMART'
                if 'TR' in store.upper(): ch = 'TRADERS'
                elif 'NBR' in store.upper(): ch = 'NOBRAND'
                
                # P열(15번 인덱스) -> 센터코드 매칭
                raw_c = str(row.iloc[15]).strip()
                delivery_code = masters[ch]['centers'].get(raw_c, "")
                
                # F열(5번 인덱스) -> 상품코드 매칭
                raw_p = str(row.iloc[5]).strip()
                me_code = masters[ch]['products'].get(raw_p, raw_p)
                p_name = masters[ch]['names'].get(raw_p, row['상품명'])
                
                final_list.append({
                    '수주일자': datetime.now().strftime('%Y%m%d'),
                    '납품일자': format_to_yyyymmdd(row[date_col]),
                    '발주처코드': CHANNELS[ch]['code'],
                    '발주처': CHANNELS[ch]['name'],
                    '배송코드': delivery_code,
                    '배송지': store,
                    '상품코드': me_code,
                    '상품명': p_name,
                    'UNIT수량': pd.to_numeric(row['수량'], errors='coerce'),
                    'UNIT단가': pd.to_numeric(row['발주원가'], errors='coerce')
                })
            
            df_processed = pd.DataFrame(final_list)
            
            # --- [동일 배송지/상품 수량 합산] ---
            group_keys = ['수주일자', '납품일자', '발주처코드', '발주처', '배송코드', '배송지', '상품코드', '상품명', 'UNIT단가']
            df_final = df_processed.groupby(group_keys, as_index=False)['UNIT수량'].sum()
            df_final['Total Amount'] = df_final['UNIT수량'] * df_final['UNIT단가']
            
            st.success(f"✅ 분석 완료 (기준 컬럼: {date_col})")
            st.dataframe(df_final)
            
            # 엑셀 다운로드
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df_final.to_excel(writer, index=False, sheet_name='Summary')
            
            st.download_button("📥 통합 주문서 다운로드", output.getvalue(), f"Order_{datetime.now().strftime('%m%d')}.xlsx")
            
        except Exception as e:
            st.error(f"데이터 처리 중 오류: {e}")
