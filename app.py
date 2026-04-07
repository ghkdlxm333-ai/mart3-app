# ... (상단 로직 동일)

    for _, row in df_raw.iterrows():
        # 컬럼 존재 여부 확인 (에러 방지)
        store_raw = str(row.get('점포명', '알수없음'))
        
        if 'TR' in store_raw.upper(): ch = 'TRADERS'
        elif 'NBR' in store_raw.upper(): ch = 'NOBRAND'
        else: ch = 'EMART'
        
        m = masters.get(ch)
        if not m: continue

        # 센터입하일자 컬럼을 유연하게 찾기
        date_col = next((c for c in df_raw.columns if '센터입하일자' in str(c)), None)
        delivery_date = format_delivery_date(row[date_col]) if date_col else datetime.now().strftime('%Y%m%d')

        # 데이터 매핑 (딕셔너리 키 이름을 group_cols와 반드시 일치시켜야 함)
        center_code_val = str(row.get('센터코드', '')).strip()
        d_code = m['c_to_b'].get(center_code_val, "")
        d_place = m['b_to_n'].get(d_code, store_raw)
        
        prod_code_val = str(row.get('상품코드', '')).strip()
        me_code = m['products'].get(prod_code_val, prod_code_val)
        p_name = m['names'].get(prod_code_val, row.get('상품명', ''))

        final_data.append({
            '수주일자': datetime.now().strftime('%Y%m%d'),
            '납품일자': delivery_date,
            '발주처코드': CHANNELS[ch]['code'],
            '발주처': CHANNELS[ch]['name'],
            '배송코드': d_code,
            '배송지': d_place,
            '상품코드': me_code,
            '상품명': p_name,
            'UNIT수량': pd.to_numeric(row.get('수량', 0), errors='coerce'),
            'UNIT단가': pd.to_numeric(row.get('발주원가', 0), errors='coerce')
        })

    # 데이터프레임 생성
    df_total = pd.DataFrame(final_data)
    
    # [중요] group_cols 리스트가 df_total의 컬럼명과 100% 일치해야 함
    group_cols = ['수주일자', '납품일자', '발주처코드', '발주처', '배송코드', '배송지', '상품코드', '상품명', 'UNIT단가']
    
    try:
        # 합산 처리
        df_final = df_total.groupby(group_cols, as_index=False)['UNIT수량'].sum()
        df_final['Total Amount'] = df_final['UNIT수량'] * df_final['UNIT단가']
        
        st.success("✅ 발주처별 통합 분류 및 합산 완료")
        st.dataframe(df_final)
    except KeyError as e:
        st.error(f"❌ 그룹화 오류: 데이터프레임에 {e} 컬럼이 없습니다. 데이터 생성 부분을 확인하세요.")
