import streamlit as st
import pandas as pd
from datetime import datetime

# 파일 업로드
uploaded_file = st.file_uploader("엑셀 파일을 업로드하세요", type=["xlsx"])

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)

    # 날짜 형식 변환
    df['공사시작일'] = pd.to_datetime(df['공사시작일'], errors='coerce')
    df['공사종료일'] = pd.to_datetime(df['공사종료일'], errors='coerce')

    # 오늘 날짜
    today = datetime.now()

    # 필터링 조건 적용
    filtered_df = df[
        (df['공사시작일'] < today) & 
        (df['공사종료일'] >= today) &
        df['공사명'].str.contains('굴착|관로|인입', na=False) & 
        df['광케이블 조수현황'].str.contains('144C', na=False)
    ]

    st.write("필터링된 데이터:")
    st.dataframe(filtered_df)

    # 필터링된 데이터를 엑셀 파일로 저장
    output_file_path = 'extracted_data_filtered.xlsx'
    filtered_df.to_excel(output_file_path, index=False)

    # 다운로드 링크 제공
    with open(output_file_path, "rb") as file:
        btn = st.download_button(
            label="필터링된 데이터 다운로드",
            data=file,
            file_name="extracted_data_filtered.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
