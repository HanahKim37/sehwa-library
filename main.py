import streamlit as st

st.set_page_config(page_title="세화고 도서관 프로그램", layout="wide")

st.markdown("""
<style>
/* 전체 폭 제한(너무 넓어 보이지 않게) */
.block-container { max-width: 1100px; padding-top: 2rem; }

/* 카드 스타일 */
.card {
    border: 1px solid rgba(49, 51, 63, 0.15);
    border-radius: 16px;
    padding: 18px 18px 14px 18px;
    background: rgba(255, 255, 255, 0.6);
    box-shadow: 0 6px 18px rgba(0,0,0,0.04);
}

/* 작은 뱃지 */
.badge {
    display: inline-block;
    font-size: 0.85rem;
    padding: 4px 10px;
    border-radius: 999px;
    border: 1px solid rgba(49, 51, 63, 0.15);
    background: rgba(255,255,255,0.7);
}

/* 섹션 타이틀 */
.section-title {
    font-size: 1.1rem;
    font-weight: 700;
    margin: 4px 0 10px 0;
}

/* 미세한 안내 텍스트 */
.muted {
    color: rgba(49, 51, 63, 0.65);
    font-size: 0.95rem;
}
</style>
""", unsafe_allow_html=True)

# =========================================================
# Hero
# =========================================================
st.markdown("""
<div class="card">
  <div class="badge">🏫 세화고등학교 · 도서관</div>
  <h1 style="margin: 10px 0 6px 0;">세화고 도서관 프로그램</h1>
  <p class="muted" style="margin: 0;">
    도서관 업무에 필요한 도구를 한곳에 모았습니다. 왼쪽 사이드바에서 프로그램을 선택해 주세요.
  </p>
</div>
""", unsafe_allow_html=True)

st.write("")
st.write("")

# =========================================================
# Two main cards
# =========================================================
c1, c2 = st.columns(2, gap="large")

with c1:
    st.markdown("""
    <div class="card">
      <div class="section-title">🎲 추첨 프로그램</div>
      <p class="muted" style="margin-top:0;">
        도서관 행사·수업 활동에서 공정한 추첨이 필요할 때 사용합니다.
      </p>
      <ul style="margin: 0 0 8px 18px;">
        <li>명단 기반 추첨</li>
        <li>중복 방지·결과 기록</li>
        <li>간단한 설정으로 빠르게 실행</li>
      </ul>
    </div>
    """, unsafe_allow_html=True)
    st.write("")
    st.info("사이드바에서 **추첨 프로그램**을 선택해 실행하세요.", icon="👉")

with c2:
    st.markdown("""
    <div class="card">
      <div class="section-title">📚 독서</div>
      <p class="muted" style="margin-top:0;">
        독서 관련 자료를 정리하고 확인하는 작업을 돕습니다.
      </p>
      <ul style="margin: 0 0 8px 18px;">
        <li>독서 데이터 확인·정리</li>
        <li>필요 기준 점검</li>
        <li>업무 흐름 단순화</li>
      </ul>
    </div>
    """, unsafe_allow_html=True)
    st.write("")
    st.info("사이드바에서 **독서가**를 선택해 실행하세요.", icon="👉")

st.write("")
st.write("")

# =========================================================
# Notice / Contact
# =========================================================
st.markdown("""
<div class="card">
  <div class="section-title">💬 문의 및 요청</div>
  <p class="muted" style="margin: 0;">
    필요한 기능이나 개선 요청이 있으면 편하게 알려주세요. 운영 흐름을 해치지 않는 범위에서 빠르게 반영하겠습니다.
  </p>
</div>
""", unsafe_allow_html=True)
