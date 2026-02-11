import streamlit as st
import os
import sys

# 경로 인식 문제 방지
current_dir = os.path.dirname(os.path.abspath(__file__))
if current_dir not in sys.path:
    sys.path.append(current_dir)

from utils.sidebar import render_sidebar
from utils.permissions import fetch_pages

# 기본 설정
HOME_PAGE_KEY = "home"
st.set_page_config(page_title="세화고 프로그램 모음", layout="wide")

# =========================================================
# 1) 데이터 로드 (로그인 없이 바로 로드)
# =========================================================
pages_catalog = fetch_pages()

# 로그인 기능을 뺐으므로 기본 사용자 정보를 설정합니다.
# 모든 사용자를 관리자(is_admin=True) 권한으로 설정하여 모든 페이지를 보이게 합니다.
is_admin = True 
display_name = "사용자"

# 모든 페이지 키를 허용 목록에 담습니다.
allowed_keys = {p["page_key"] for p in pages_catalog}
allowed_keys.add(HOME_PAGE_KEY)

# =========================================================
# 2) 사이드바 및 페이지 실행
# =========================================================
with st.sidebar:
    st.markdown(f"### 🏫 세화고등학교")
    st.markdown(f"**{display_name}님 반갑습니다.**")
    st.divider()

# 사이드바 렌더링 (utils/sidebar.py 함수 호출)
render_sidebar(pages_catalog, allowed_keys, is_admin)

# 현재 선택된 페이지 실행 로직
go_key = st.session_state.get("__go_page_key__") or HOME_PAGE_KEY
page_map = {p["page_key"]: p for p in pages_catalog}
p = page_map.get(go_key)

if p:
    if p.get("is_active", True):
        pg = st.navigation([st.Page(p["file_path"], title=p["title"])])
        pg.run()
    else:
        st.warning("현재 비활성화된 페이지입니다.")
else:
    # 홈 페이지 기본 안내
    st.title("🏠 세화고 프로그램 통합 관리 시스템")
    st.info("왼쪽 사이드바에서 사용할 프로그램을 선택해 주세요.")
