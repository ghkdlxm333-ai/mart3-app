import streamlit as st
import pandas as pd
import io
from datetime import datetime

# 1. 화면 설정
st.set_page_config(page_title="이마트 계열 수주 자동화", page_icon="🟢", layout="wide")

# --- [날짜 변환 함수: YYYYMMDD 강제 고정] ---
def format_delivery_date(val):
    if pd.isna(val) or str(val).strip() in ["", "0", "nan"]:
        return datetime.now().strftime('%Y%m%d')
    
    try:
        # 엑셀 날짜 객체 처리
        if isinstance(val, datetime):
            return val.strftime('%Y%m%d')
        
        # 문자열 내 숫자만 추출
        str_val = str(val).split(' ')[0]
        clean_val = ''.join(filter(str.isdigit, str_val))
        
        if len(clean_val) >= 8:
            return clean_val[:8]
        else:
            # 8자리가 아닐 경우 판다스 도구로 재파싱
            return pd.to_datetime(val).strftime('%Y%m%d')
    except:
        return datetime.now().strftime('%Y%m%d')

@st.cache_data
def load_master_data(file_path):
    try:
        xls = pd.ExcelFile(file_path)
        sheet_map = {s.strip(): s for s in xls.sheet_names}
        
        target_sheet = sheet_map.get('배송코드')
        if not target_sheet: return None, f"'{file_path}'에 '배송코드' 시트가 없습니다."
        
        df_tmp = pd.read_excel(xls, sheet_name=target_sheet, dtype=str)
        df_center = None
        for i, row in df_tmp.iterrows():
            if '배송코드' in [str(v).strip() for v in row.values]:
                df_center = pd.read_excel(xls, sheet_name=target_sheet, skiprows=i+1, dtype=str)
                df_center.columns = [str(c).strip() for c in df_tmp.iloc[i]]
                break
        
        if df_center is None: df_center = df_tmp
        c_to_b = dict(zip(df_center['센터코드'].str.strip(), df_center['배송코드'].str.strip()))
        b_to_n = dict(zip(df_center['배송코드'].str.strip(), df_center.iloc[:, 2].str.strip())) 
        
        prod_sheet = sheet_map.get('제품명')
        df_prod = pd.read_excel(xls, sheet_name=prod_sheet, dtype=str) if prod_sheet else None
        p_map = dict(zip(df_prod['상품코드'].str.strip(), df_prod['ME코드'].str.strip())) if df_prod is not None else {}
        n_map = dict(zip(df_prod['상품코드'].str.strip(), df_prod.iloc[:, 1].str.strip())) if df_prod is not None else {}
        
        return {'c_to_b': c_to_b, 'b_to_n': b_to_n, 'products': p_map, 'names': n_map}, None
    except Exception as e:
        return None, str(e)

# --- 메인 실행부 ---
st.title("🛒🟢 이마트 계열 수주 자동화 ")

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
    uploaded_file = st.file_uploader("ORDERS 파일 업로드", type=['xlsx'])
    
    if uploaded_file:
        try:
            df_raw = pd.read_excel(uploaded_file)
            # 센터입하일자 컬럼 찾기
            date_col = next((c for c in df_raw.columns if '센터입하일자' in str(c)), None)
            
            final_data = []
            for _, row in df_raw.iterrows():
                store_raw = str(row.get('점포명', ''))
                
                # [채널 분류 로직 수정] NBFC 등 대응을 위해 'NB'로 완화
                if 'TR' in store_raw.upper():
                    ch = 'TRADERS'
                elif 'NB' in store_raw.upper():  # NBR뿐만 아니라 NBFC 등도 포함
                    ch = 'NOBRAND'
                else:
                    ch = 'EMART'
                
                m = masters[ch]
                # P열(15), F열(5) 인덱스 기준 추출
                c_val = str(row.iloc[15]).strip() if len(row) > 15 else ""
                d_code = m['c_to_b'].get(c_val, "")
                d_place = m['b_to_n'].get(d_code, store_raw)
                
                p_val = str(row.iloc[5]).strip() if len(row) > 5 else ""
                me_code = m['products'].get(p_val, p_val)
                p_name = m['names'].get(p_val, str(row.get('상품명', '')))

                final_data.append({
                    '수주일자': datetime.now().strftime('%Y%m%d'),
                    '납품일자': format_delivery_date(row[date_col]) if date_col else datetime.now().strftime('%Y%m%d'),
                    '발주처코드': CHANNELS[ch]['code'],
                    '발주처': CHANNELS[ch]['name'],
                    '배송코드': d_code,
                    '배송지': d_place,
                    '상품코드': me_code,
                    '상품명': p_name,
                    'UNIT수량': pd.to_numeric(row.get('수량', 0), errors='coerce'),
                    'UNIT단가': pd.to_numeric(row.get('발주원가', 0), errors='coerce')
                })
            
            df_mid = pd.DataFrame(final_data)
            group_cols = ['수주일자', '납품일자', '발주처코드', '발주처', '배송코드', '배송지', '상품코드', '상품명', 'UNIT단가']
            df_final = df_mid.groupby(group_cols, as_index=False)['UNIT수량'].sum()
            df_final['Total Amount'] = df_final['UNIT수량'] * df_final['UNIT단가']
            
            # 납품일자 문자열 유지
            df_final['납품일자'] = df_final['납품일자'].astype(str)

            st.success("✅ 노브랜드(NB) 채널 확장 및 납품일자 고정 완료")
            st.dataframe(df_final)
            
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df_final.to_excel(writer, index=False, sheet_name='수주업로드용')
            
            st.download_button(
                label="📥 최종 파일 다운로드",
                data=output.getvalue(),
                file_name=f"Order_Final_{datetime.now().strftime('%m%d')}.xlsx"
            )
            
        except Exception as e:
            st.error(f"오류 발생: {e}")
