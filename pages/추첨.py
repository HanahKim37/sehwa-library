import os
import sys
import io

import streamlit as st
import pandas as pd


st.title("🎲 추첨 프로그램")
st.write(
    """
업로드한 엑셀 파일에서 **'학번'**, **'이름'** 열을 기준으로  
학년·반을 선택하고, 특정 조건의 학생을 제외한 뒤 무작위로 추첨합니다.
"""
)

# 1. 엑셀 파일 업로드
uploaded_file = st.file_uploader(
    "엑셀 파일을 업로드하세요 (.xlsx, .xls)",
    type=["xlsx", "xls"],
)

if uploaded_file is None:
    st.info("먼저 엑셀 파일을 업로드해주세요.")
    st.stop()

# 2. 엑셀 읽기
try:
    df = pd.read_excel(uploaded_file)
except Exception:
    st.error("엑셀 파일을 읽는 중 오류가 발생했습니다. 파일 형식을 다시 확인해주세요.")
    st.stop()

required_cols = ["학번", "이름"]
optional_grade_class_cols = ["학년", "반"]

# 3. 필수 열 확인
if not all(col in df.columns for col in required_cols):
    st.error("엑셀에 **'학번'**, **'이름'** 열이 모두 존재해야 합니다.")
    st.write("현재 엑셀에 있는 열 목록:", list(df.columns))
    st.stop()

# ---------------------------------------------------------------------
# 4. 학년·반 필터링
# ---------------------------------------------------------------------
df_filtered = df.copy()

st.subheader("1️⃣ 학년·반 필터 설정")

has_grade = "학년" in df.columns
has_class = "반" in df.columns

selected_grade = None
selected_class = None

if has_grade or has_class:
    col1, col2 = st.columns(2)

    # 학년 선택
    if has_grade:
        with col1:
            grades = df["학년"].dropna().unique().tolist()
            grades = sorted(grades)
            grade_options = ["전체"] + [str(g) for g in grades]
            selected_grade = st.selectbox("학년 선택", grade_options, index=0)

            if selected_grade != "전체":
                df_filtered = df_filtered[df_filtered["학년"].astype(str) == selected_grade]

    # 반 선택 (학년이 있으면 선택된 학년 범위 내에서 반 목록 생성)
    if has_class:
        with col2:
            if has_grade and selected_grade not in (None, "전체"):
                class_source = df_filtered
            else:
                class_source = df

            classes = class_source["반"].dropna().unique().tolist()
            classes = sorted(classes)
            class_options = ["전체"] + [str(c) for c in classes]
            selected_class = st.selectbox("반 선택", class_options, index=0)

            if selected_class != "전체":
                df_filtered = df_filtered[df_filtered["반"].astype(str) == selected_class]

else:
    st.info("이 엑셀에는 '학년', '반' 열이 없어 전체 인원을 대상으로 추첨합니다.")

# ---------------------------------------------------------------------
# 5. 제외 기준 열 및 값 선택
# ---------------------------------------------------------------------
st.subheader("2️⃣ 추첨 대상에서 제외할 기준 설정")

# 제외 기준 열 선택
exclude_col_options = ["사용 안 함"] + list(df.columns)
exclude_col = st.selectbox(
    "제외 기준이 적혀 있는 열을 선택하세요 (없으면 '사용 안 함' 선택)",
    options=exclude_col_options,
    index=0,
)

exclude_values = []
if exclude_col != "사용 안 함":
    # 현재 학년·반 필터가 적용된 데이터 기준으로 값 목록 생성
    col_values = df_filtered[exclude_col].dropna().astype(str).unique().tolist()
    col_values = sorted(col_values)

    if len(col_values) == 0:
        st.info(f"선택한 열('{exclude_col}')에 값이 없어 제외 조건을 적용할 수 없습니다.")
    else:
        exclude_values = st.multiselect(
            f"열 '{exclude_col}'에서 제외할 값을 선택하세요.",
            options=col_values,
        )

# 제외 조건 적용
df_final = df_filtered.copy()
if exclude_col != "사용 안 함" and exclude_values:
    df_final = df_final[~df_final[exclude_col].astype(str).isin(exclude_values)]

# ---------------------------------------------------------------------
# 6. 인원 요약 정보
# ---------------------------------------------------------------------
total_count = len(df)
after_grade_class_count = len(df_filtered)
final_count = len(df_final)
excluded_count = after_grade_class_count - final_count

st.subheader("3️⃣ 인원 요약")

col_a, col_b, col_c, col_d = st.columns(4)
col_a.metric("전체 인원", f"{total_count}명")
col_b.metric("학년·반 필터 후", f"{after_grade_class_count}명")
col_c.metric("제외 처리 인원", f"{excluded_count}명")
col_d.metric("최종 추첨 대상", f"{final_count}명")

if final_count == 0:
    st.warning("최종 추첨 대상 인원이 0명입니다. 학년·반 또는 제외 조건을 조정해 주세요.")
    st.stop()



# ---------------------------------------------------------------------
# 7. 추첨 인원 수 입력 및 추첨 실행
# ---------------------------------------------------------------------
st.subheader("4️⃣ 추첨 실행")

num_winners = st.number_input(
    "추첨 인원 수를 입력하세요",
    min_value=1,
    step=1,
)

if st.button("✅ 추첨 시작"):
    if int(num_winners) > final_count:
        st.warning(
            f"추첨 인원({int(num_winners)}명)이 최종 추첨 대상 인원({final_count}명)보다 많습니다. "
            "인원 수를 줄이거나 필터/제외 조건을 조정해 주세요."
        )
        st.stop()

    # 학번·이름 기준으로만 추첨
    result_df = df_final[required_cols].sample(
        n=int(num_winners),
        replace=False,
        random_state=None,  # 실행할 때마다 다른 결과
    ).reset_index(drop=True)

    # 번호 컬럼 추가
    result_df.index = result_df.index + 1
    result_df.index.name = "번호"

    st.success("추첨이 완료되었습니다.")
    st.subheader("🎉 추첨 결과")
    st.dataframe(result_df)

    # 엑셀 다운로드
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        result_df.to_excel(writer, sheet_name="추첨결과")
    output.seek(0)

    st.download_button(
        label="📥 추첨 결과 엑셀 다운로드",
        data=output,
        file_name="추첨결과.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
