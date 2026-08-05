"""
OpenDART + KRX 기업 스크리너 - Streamlit 단일 페이지 앱

실행 방법:
    pip install streamlit pandas requests openpyxl lxml finance-datareader
    streamlit run app.py

배포 방법 (무료로 URL 받기):
    1. 이 파일을 GitHub 저장소에 올린다
    2. share.streamlit.io 에서 저장소 연결 -> 자동 배포
    (API 키는 st.secrets 또는 화면에서 매번 입력하는 방식 사용 - 코드에 하드코딩하지 말 것)
"""

import io
import re
import time
import zipfile
import datetime
import xml.etree.ElementTree as ET
from urllib.parse import quote
from concurrent.futures import ThreadPoolExecutor, as_completed

import pandas as pd
import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry
import streamlit as st

st.set_page_config(page_title="기업 스크리너", layout="wide")


def make_session():
    """일시적인 연결 끊김에 대비해 재시도 로직이 포함된 세션 생성"""
    s = requests.Session()
    retry = Retry(
        total=5,
        backoff_factor=1.5,  # 1.5s, 3s, 4.5s, 6s, 7.5s 간격으로 재시도
        status_forcelist=[429, 500, 502, 503, 504],
        allowed_methods=["GET"],
    )
    adapter = HTTPAdapter(max_retries=retry)
    s.mount("https://", adapter)
    s.mount("http://", adapter)
    return s

SIDO_MAP = {
    '서울특별시': '서울',
