import streamlit as st
import pandas as pd
import io
from datetime import datetime

# 1. 화면 설정
st.set_page_config(page_title="이마트 계열 수주 자동화", page_icon="🟢", layout="wide")

# --- [날짜 변환 함수: YYYYMMDD 형식을 확실하게 보장] ---
def format_delivery_date(val):
    if pd.isna(val) or str(val).strip() in ["", "0", "nan", "None", "19700101"]:
        return datetime.now().strftime('%Y%m%d')
    
    try:
        if isinstance(val, datetime):
            return val.strftime('%Y%m%d')
        
        # 문자열에서 날짜 부분만 추출 및 숫자만 남기기
        str_val = str(val).split(' ')[0].split('T')[0]
        clean_val = ''.join(filter(str.isdigit, str_val))
        
        if len(clean_val) >= 8:
            return clean_val[:8]
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
st.title("🛒🟢 이마트 계열 수주 자동화 (DRY센터 로직 추가)")

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
            date_col = next((c for c in df_raw.columns if '센터입하일자' in str(c).replace(" ", "")), None)
            
            final_data = []
            for _, row in df_raw.iterrows():
                store_name = str(row.get('점포명', '')).upper().strip()
                
                # [채널 분류 로직 수정]
                # 1. DRY가 포함되면 무조건 이마트 (NB보다 우선순위 높음)
                if 'DRY' in store_name:
                    ch = 'EMART'
                # 2. NB나 NBR이 포함되면 노브랜드
                elif 'NB' in store_name:
                    ch = 'NOBRAND'
                # 3. TR이 포함되면 트레이더스
                elif 'TR' in store_name:
                    ch = 'TRADERS'
                # 4. 그 외 이마트
                else:
                    ch = 'EMART'
                
                m = masters[ch]
                # P열(15): 센터코드, F열(5): 상품코드
                c_val = str(row.iloc[15]).strip() if len(row) > 15 else ""
                d_code = m['c_to_b'].get(c_val, "")
                # 배송지는 마스터의 지명을 우선하되 없으면 점포명 유지
                d_place = m['b_to_n'].get(d_code, str(row.get('점포명', '')))
                
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
            
            # 납품일자 문자열 고정
            df_final['납품일자'] = df_final['납품일자'].astype(str)

            st.success(f"✅ 분류 완료: DRY센터(이마트), NB/NBR(노브랜드)")
            st.dataframe(df_final, use_container_width=True)
            
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df_final.to_excel(writer, index=False, sheet_name='수주업로드용')
            
            st.download_button(
                label="📥 최종 수정 파일 다운로드",
                data=output.getvalue(),
                file_name=f"Final_Order_DRY_Fixed_{datetime.now().strftime('%m%d')}.xlsx"
            )
            
        except Exception as e:
            st.error(f"오류 발생: {e}")
