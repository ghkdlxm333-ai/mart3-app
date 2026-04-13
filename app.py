import streamlit as st
import pandas as pd
import io
from datetime import datetime

# 1. 화면 설정
st.set_page_config(page_title="이마트 계열 수주 자동화", page_icon="🟢", layout="wide")

# --- [날짜 변환 함수: 납품일자용] ---
def format_delivery_date(val):
    if pd.isna(val) or str(val).strip() in ["", "0", "nan", "None", "19700101"]:
        return datetime.now().strftime('%Y%m%d')
    try:
        if isinstance(val, datetime):
            return val.strftime('%Y%m%d')
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
        
        # 1. 배송코드 시트 처리
        target_sheet = sheet_map.get('배송코드')
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
        
        # 2. 제품명 시트 처리 (상품명(기획) E열 대응 로직)
        prod_sheet = sheet_map.get('제품명')
        p_map, n_map = {}, {}
        if prod_sheet:
            df_prod_raw = pd.read_excel(xls, sheet_name=prod_sheet, dtype=str)
            df_prod_raw.columns = [str(c).strip() for c in df_prod_raw.columns]
            
            col_s = next((c for c in df_prod_raw.columns if '상품코드' in c), None)
            col_m = next((c for c in df_prod_raw.columns if 'ME코드' in c), None)
            # [추가로직] '상품명(기획)' 컬럼을 최우선으로 찾고, 없으면 '상품명' 사용
            col_n = next((c for c in df_prod_raw.columns if '상품명(기획)' in c), 
                         next((c for c in df_prod_raw.columns if '상품명' in c), None))
            
            for _, p_row in df_prod_raw.iterrows():
                s_code = str(p_row.get(col_s, '')).strip()
                if s_code and s_code != 'nan':
                    p_map[s_code] = str(p_row.get(col_m, s_code)).strip()
                    n_map[s_code] = str(p_row.get(col_n, '')).strip()
        
        return {'c_to_b': c_to_b, 'b_to_n': b_to_n, 'products': p_map, 'names': n_map}, None
    except Exception as e:
        return None, str(e)

# --- 메인 실행부 ---
st.title("🛒🟢 이마트 계열 수주 자동화")

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
    # --- [추가로직 1] 안내 문구 구간 ---
    st.markdown("### ※ 업로드 전 확인사항")
    st.info("💡 **엑셀파일 확장자를 .xlsx로 변환 후 업로드해주세요.** (xls, csv 파일은 변환이 필요합니다)")
    
    # -------------------------------

    uploaded_file = st.file_uploader("이마트, 노브랜드, 트레이더스 발주서 취합 파일로 업로드해주세요.", type=['xlsx'])
    
    if uploaded_file:
        try:
            # 1. 파일 읽기
            df_raw = pd.read_excel(uploaded_file)
            
            # [보완 1] 원본 파일에 있는 날짜/일자 관련 컬럼명을 아예 삭제 (인식 차단)
            # A열이 '발주일자'이므로 이를 삭제해야 시스템이 헷갈리지 않습니다.
            raw_cols = df_raw.columns.tolist()
            ignore_list = ['발주일자', '수주일자', '발주 일자', '수주 일자']
            cols_to_drop = [c for c in raw_cols if any(x in str(c).replace(" ", "") for x in ignore_list)]
            df_raw = df_raw.drop(columns=cols_to_drop)
            
            # [보완 2] 납품일자 계산을 위한 센터입하일자 컬럼은 별도로 확보
            # (위에서 삭제되지 않도록 '센터입하일자'라는 정확한 명칭은 보존됨)
            date_col = next((c for c in df_raw.columns if '센터입하일자' in str(c).replace(" ", "")), None)
            
            # [보완 3] 오늘 날짜 정의 (문자열)
            real_today = datetime.now().strftime('%Y%m%d')

            final_data = []
            for _, row in df_raw.iterrows():
                # ... (채널 판별 및 매핑 로직 동일) ...
                store_raw = str(row.get('점포명', ''))
                store_upper = store_raw.upper().strip()
                if 'DRY' in store_upper: ch = 'EMART'
                elif 'NB' in store_upper: ch = 'NOBRAND'
                elif 'TR' in store_upper: ch = 'TRADERS'
                else: ch = 'EMART'
                
                m = masters[ch]
                c_val = str(row.iloc[15]).strip() if len(row) > 15 else ""
                d_code = m['c_to_b'].get(c_val, "")
                d_place = m['b_to_n'].get(d_code, store_raw)
                p_val = str(row.iloc[5]).strip() if len(row) > 5 else ""
                me_code = m['products'].get(p_val, p_val)
                p_name_master = m['names'].get(p_val)
                p_name_final = p_name_master if p_name_master and p_name_master != 'nan' else str(row.get('상품명', ''))

                # [보완 4] 데이터 추가 시점에 real_today를 명시적 문자열로 주입
                final_data.append({
                    '구분': 0,
                    '수주일자': str(real_today), 
                    '납품일자': format_delivery_date(row.get(date_col)) if date_col else real_today,
                    '발주처코드': CHANNELS[ch]['code'],
                    '발주처': CHANNELS[ch]['name'],
                    '배송코드': d_code,
                    '배송지': d_place,
                    '상품코드': me_code,
                    '상품명': p_name_final,
                    'UNIT수량': pd.to_numeric(row.get('수량', 0), errors='coerce'),
                    'UNIT단가': pd.to_numeric(row.get('발주원가', 0), errors='coerce')
                })
            
            df_mid = pd.DataFrame(final_data)
            
            # [보완 5] 그룹화 전 데이터 타입을 문자열로 강제 고정
            df_mid['수주일자'] = df_mid['수주일자'].astype(str)
            
            group_cols = ['구분', '수주일자', '납품일자', '발주처코드', '발주처', '배송코드', '배송지', '상품코드', '상품명', 'UNIT단가']
            df_final = df_mid.groupby(group_cols, as_index=False)['UNIT수량'].sum()
            df_final['Total Amount'] = df_final['UNIT수량'] * df_final['UNIT단가']
            
            column_order = ['구분', '수주일자', '납품일자', '발주처코드', '발주처', '배송코드', '배송지', '상품코드', '상품명', 'UNIT수량', 'UNIT단가', 'Total Amount']
            df_final = df_final[column_order]
            
            # [보완 6] 최종 방어: 모든 처리가 끝난 후 컬럼 전체를 오늘 날짜 문자열로 덮어쓰기
            df_final['수주일자'] = str(real_today)
            df_final['납품일자'] = df_final['납품일자'].astype(str)

            st.success(f"✅ 분석 완료!")
            st.dataframe(df_final, use_container_width=True)
            
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df_final.to_excel(writer, index=False, sheet_name='수주업로드용')
            
            st.download_button(
                label="📥 결과 다운로드",
                data=output.getvalue(),
                file_name=f"Order_Upload_TODAY_{real_today}.xlsx"
            )
            
        except Exception as e:
            st.error(f"오류 발생: {e}")
