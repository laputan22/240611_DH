import streamlit as st
import pandas as pd
import datetime
from io import BytesIO
import matplotlib.pyplot as plt
from matplotlib import font_manager, rc
from pathlib import Path

# 프로젝트 루트 디렉토리에 업로드된 폰트 파일 사용
font_path = Path("malgun.ttf")

# 폰트 설정
if font_path.exists():
    font_name = font_manager.FontProperties(fname=str(font_path)).get_name()
    rc('font', family=font_name)
else:
    st.error("폰트 파일을 사용할 수 없습니다.")

st.title('굴착 정보 필터링 앱')

# 파일 업로드
uploaded_file = st.file_uploader("엑셀 파일을 업로드하세요", type=["xlsx"])

if uploaded_file:
    # 데이터 로드
    @st.cache_data
    def load_data(file):
        df = pd.read_excel(file, sheet_name='KTOA-EOCS 굴착정보_1')
        df['공사시작일'] = pd.to_datetime(df['공사시작일'], errors='coerce')
        df['공사종료일'] = pd.to_datetime(df['공사종료일'], errors='coerce')
        return df

    df = load_data(uploaded_file)

    # 오늘 날짜와 이번 달의 마지막 날짜 구하기
    today = pd.to_datetime(datetime.date.today())
    end_of_month = today + pd.offsets.MonthEnd(0)

    # 공사명 제외 조건 리스트
    exclude_keywords = ['확장', '지진 저항', '주차장', '보강', '파란', '등대', '유지보수', '아스베스트', '단열재',
                        '대체', '철거', '궤도', '승강기', '외벽 수리', '외벽 개선', '외벽 마감', '족', '스프링클러',
                        '로비', '휴게실', '옥상 방수', '옥상', '방수', '이중 창문 교체', '슬레이트 해체', '창문',
                        '차선 페인팅', '지붕 개선', '물 누출', '화장실', '식당', '활성 탄소', '노폐물 포장', '샌드위치',
                        '수도계량기', '교실 바닥', '바닥', '중요 수리', '5030', '플랫폼', '아케이드', '싱크대 수전',
                        '보일러', '그늘 천막', '포장 복원', '캐노피', '유지보수', '기계 장비', '중학교', '대학교']

    # 필터링 조건 적용
    filtered_df = df[
        (~df['공사명'].str.contains('|'.join(exclude_keywords), na=False)) &
        (df['T_지중'].str.contains('144C|288C', na=False)) &
        (df['공사종료일'] >= today) &
        (df['공사종료일'] <= end_of_month)
    ]

    # 공사시작일에 따른 두 개의 데이터프레임 생성
    construction_start_today_or_before = filtered_df[filtered_df['공사시작일'] <= today]
    construction_start_after_today = filtered_df[(filtered_df['공사시작일'] > today) & (filtered_df['공사시작일'] <= end_of_month)]

    # 공사시점주소의 첫 단어를 기준으로 그룹화
    construction_start_today_or_before['권역'] = construction_start_today_or_before['공사시점주소'].apply(lambda x: x.split()[0])
    construction_start_after_today['권역'] = construction_start_after_today['공사시점주소'].apply(lambda x: x.split()[0])

    grouped_start = construction_start_today_or_before.groupby('권역').size().reset_index(name='공사시작')
    grouped_upcoming = construction_start_after_today.groupby('권역').size().reset_index(name='공사예정')

    # 그룹화된 데이터 병합
    grouped = pd.merge(grouped_start, grouped_upcoming, on='권역', how='outer').fillna(0)
    grouped['공사시작'] = grouped['공사시작'].astype(int)
    grouped['공사예정'] = grouped['공사예정'].astype(int)

    # 합계 행 추가
    total_row = pd.DataFrame([['합계', grouped['공사시작'].sum(), grouped['공사예정'].sum()]], columns=['권역', '공사시작', '공사예정'])
    grouped = pd.concat([total_row, grouped])

    # 데이터 프레임을 화면에 표시
    st.subheader('필터링된 결과')
    st.dataframe(grouped)

    # 데이터 시각화
    fig, ax = plt.subplots(figsize=(10, 6))
    bars = grouped.set_index('권역').plot(kind='bar', stacked=True, ax=ax)

    # 각 막대 위에 숫자 표시
    for p in ax.patches:
        height = p.get_height()
        if height > 0:
            ax.annotate(f'{int(height)}', 
                        (p.get_x() + p.get_width() / 2., height), 
                        ha = 'center', 
                        va = 'bottom', 
                        xytext = (0, 5), 
                        textcoords = 'offset points',
                        fontsize=10,
                        color='black',
                        weight='bold')

    plt.title('공사 현황')
    plt.xlabel('권역')
    plt.ylabel('공사 건수')
    plt.xticks(rotation=45)
    st.pyplot(fig)

    # 필터링된 데이터 다운로드
    def convert_df_to_excel(df):
        output = BytesIO()
        writer = pd.ExcelWriter(output, engine='xlsxwriter')
        df.to_excel(writer, index=False, sheet_name='Sheet1')
        writer.close()
        processed_data = output.getvalue()
        return processed_data

    start_excel = convert_df_to_excel(construction_start_today_or_before)
    planned_excel = convert_df_to_excel(construction_start_after_today)

    st.download_button(
        label='공사시작 엑셀 파일 다운로드',
        data=start_excel,
        file_name='공사시작.xlsx',
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    )

    st.download_button(
        label='공사예정 엑셀 파일 다운로드',
        data=planned_excel,
        file_name='공사예정.xlsx',
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    )
