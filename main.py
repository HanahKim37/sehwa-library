import streamlit as st

# ─────────────────────────────────────────────────────────────
# 설정
# ─────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="세화고 도서관 프로그램",
    page_icon="🏫",
    layout="wide",
)

DRAW_PAGE_PATH = "pages/추첨.py"
READING_PAGE_PATH = "pages/독서가.py"

# ─────────────────────────────────────────────────────────────
# 스타일 (간단 카드 UI)
# ─────────────────────────────────────────────────────────────
st.markdown(
    """
<style>
.hero {
  padding: 22px 24px;
  border-radius: 18px;
  background: linear-gradient(135deg, rgba(30,136,229,0.10), rgba(0,188,212,0.10));
  border: 1px solid rgba(0,0,0,0.06);
}
.hero h1 { margin: 0 0 6px 0; font-size: 34px; }
.hero p { margin: 0; color: rgba(0,0,0,0.65); font-size: 15px; line-height: 1.5; }

.card {
  padding: 18px 18px;
  border-radius: 16px;
  border: 1px solid rgba(0,0,0,0.08);
  background: rgba(255,255,255,0.60);
}
.card h3 { margin: 0 0 6px 0; font-size: 20px; }
.card .desc { margin: 0 0 10px 0; color: rgba(0,0,0,0.70); line-height: 1.55; }
.badge {
  display: inline-block;
  padding: 4px 10px;
  border-radius: 999px;
  font-size: 12px;
  background: rgba(0,0,0,0.05);
  margin-right: 6px;
}
.hr {
  height: 1px;
  border: 0;
  background: rgba(0,0,0,0.08);
  margin: 14px 0;
}
.small { color: rgba(0,0,0,0.60); font-size: 13px; line-height: 1.55; }
</style>
""",
    unsafe_allow_html=True,
)

# ─────────────────────────────────────────────────────────────
# 상단 히어로
# ─────────────────────────────────────────────────────────────
st.markdown(
    """
<div class="hero">
  <h1>🏫 세화고 도서관 프로그램</h1>
  <p>
    도서관 업무를 효율적으로 처리할 수 있도록 필요한 도구를 한곳에 모았습니다.<br/>
    왼쪽 페이지 메뉴(또는 아래 버튼)에서 원하는 프로그램으로 이동할 수 있습니다.
  </p>
</div>
""",
    unsafe_allow_html=True,
)

st.write("")

# ─────────────────────────────────────────────────────────────
# 페이지 이동 헬퍼
# ─────────────────────────────────────────────────────────────
def go_to(page_path: str):
    # Streamlit 버전에 따라 page_link/switch_page 지원이 다를 수 있어 안전 처리
    try:
        st.switch_page(page_path)
    except Exception:
        st.warning("이 환경에서는 버튼 이동이 제한됩니다. 왼쪽 페이지 메뉴에서 이동해 주세요.")

# ─────────────────────────────────────────────────────────────
# 카드 섹션
# ─────────────────────────────────────────────────────────────
col1, col2 = st.columns(2, gap="large")

with col1:
    st.markdown(
        """
<div class="card">
  <h3>🎲 추첨 프로그램</h3>
  <p class="desc">
    엑셀 파일을 업로드한 뒤 학년·반을 선택하고, 특정 조건의 학생을 제외한 후 무작위로 추첨합니다.
  </p>
  <span class="badge">엑셀 업로드</span>
  <span class="badge">학년·반 필터</span>
  <span class="badge">제외 조건</span>
  <span class="badge">결과 엑셀 다운로드</span>
  <hr class="hr"/>
  <div class="small">
    • 필수 열: <b>학번</b>, <b>이름</b><br/>
    • (선택) 학년/반 열이 있으면 학년·반 필터가 활성화됩니다.
  </div>
</div>
""",
        unsafe_allow_html=True,
    )
    st.write("")
    if st.button("🎲 추첨 프로그램 열기", use_container_width=True):
        go_to(DRAW_PAGE_PATH)

with col2:
    st.markdown(
        """
<div class="card">
  <h3>📚 독서가</h3>
  <p class="desc">
    독서활동상황 엑셀을 업로드하면 학기별 충족 여부와 총 충족 여부를 산출하고, 결과 엑셀을 내려받을 수 있습니다.
  </p>
  <span class="badge">학기별 판정</span>
  <span class="badge">총 충족 계산</span>
  <span class="badge">필독서 판정</span>
  <span class="badge">상세/요약 엑셀</span>
  <hr class="hr"/>
  <div class="small">
    • 학기 체크 해제 시 해당 학기는 <b>‘-’</b>로 처리되며 총 충족 계산에서 제외됩니다.<br/>
    • 체크 해제된 학기는 <b>상세 시트 및 다운로드 엑셀에서도 제외</b>됩니다.
  </div>
</div>
""",
        unsafe_allow_html=True,
    )
    st.write("")
    if st.button("📚 독서가 열기", use_container_width=True):
        go_to(READING_PAGE_PATH)

st.write("")
st.markdown("---")

# ─────────────────────────────────────────────────────────────
# 하단 안내
# ─────────────────────────────────────────────────────────────
st.subheader("사용 안내")
st.markdown(
    """
- 프로그램 실행은 **왼쪽 페이지 메뉴** 또는 **위 버튼**으로 이동할 수 있습니다.  
- 파일 업로드 후 설정값(학년·반/학기 체크 등)을 조정하면, 화면에서 결과가 갱신됩니다.  
- 결과 엑셀 생성 중에는 진행 표시(스피너/진행바)가 나타납니다.
"""
)

st.subheader("문의 및 개선 요청")
st.markdown(
    """
필요한 기능이 있거나 오류가 발견되면, 파일(또는 화면 캡처)과 함께 알려 주시면 빠르게 반영하겠습니다.
"""
)
