import re
import sys
from io import BytesIO
from pathlib import Path
from typing import Dict, List, Optional, Tuple
from difflib import SequenceMatcher

import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font

# =========================================================
# 0) 임계값(요청 반영)
# =========================================================
REQUIRED_TITLE_THRESHOLD = 0.80
REQUIRED_AUTHOR_THRESHOLD = 0.60
DUP_TITLE_THRESHOLD = 0.80
DUP_AUTHOR_THRESHOLD = 0.60

# =========================================================
# 1) 페이지 설정
# =========================================================
PAGE_TITLE = "독서활동상황_충족_여부_판단"

# 앱 페이지를 단독 실행(streamlit run app_pages/...)할 때도 루트 utils를 import할 수 있도록 경로를 보정
PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

st.title(PAGE_TITLE)

 
-# =========================================================
-# 2) 공통 사이드바
-# =========================================================
-from utils.sidebar import render_sidebar
 
-render_sidebar()
+# 앱 페이지를 단독 실행(streamlit run app_pages/...)할 때도 루트 utils를 import할 수 있도록 경로를 보정
+PROJECT_ROOT = Path(__file__).resolve().parents[1]
+if str(PROJECT_ROOT) not in sys.path:
+    sys.path.insert(0, str(PROJECT_ROOT))
 
 st.title(PAGE_TITLE)
 
 st.info(
     "※ 독서활동상황 엑셀은 **한 반의 파일만** 업로드해 주세요.\n\n"
     "✅ 또한 **파일명에 학년-반을 반드시 포함**해 주세요. (예: `1-1`, `3-12`)\n"
     "- 예시 파일명: `독서활동_2-10.xlsx`, `2-10_독서활동상황.xlsx`"
 )
 
 # =========================================================
-# 3) 내장 필독서 경로 (깃허브 저장 파일)
+# 2) 내장 필독서 경로 (깃허브 저장 파일)
 # =========================================================
 REQUIRED_2024_PATH = Path("data/required_books/필독서_2024.xlsx")
 REQUIRED_2025_PATH = Path("data/required_books/필독서_2025.xlsx")
 
 # =========================================================
-# 4) 유틸 함수
+# 3) 유틸 함수
 # =========================================================
 def _normalize_text(s: str) -> str:
     """비교용 정규화(공백 제거/기호 제거/대소문자 무시)."""
     if s is None:
         return ""
     s = str(s).strip().lower()
     s = re.sub(r"\s+", "", s)
     s = re.sub(r"[^\w가-힣]", "", s)
     return s
 
 
 def _title_variants_norm(title: str) -> List[str]:
     """
     제목에서 비교용 변형 키를 여러 개 생성
     - 원문 전체
     - 괄호/대괄호/콜론/슬래시/대시 등 앞부분
     """
     if title is None:
         return []
     t = str(title).strip()
     if not t:
         return []
 
     seps = ["(", "（", "[", "【", ":", "：", "/", "／", " - ", " – ", " — ", "-", "–", "—"]
     candidates = [t]
