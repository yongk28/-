"""
DART 전체법인 + OpenDART 기업 스크리너 - Streamlit 단일 페이지 앱

실행 방법:
    pip install streamlit pandas requests openpyxl
    streamlit run app.py

배포 방법 (무료로 URL 받기):
    1. 이 파일을 GitHub 저장소에 올린다
    2. share.streamlit.io 에서 저장소 연결 -> 자동 배포
    (API 키는 화면에서 매번 입력하는 방식 - 코드에 하드코딩하지 말 것)
"""

import io
import re
import time
import threading
from urllib.parse import quote
from concurrent.futures import ThreadPoolExecutor, as_completed

import pandas as pd
import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry
import streamlit as st

st.set_page_config(page_title="기업 스크리너", layout="wide")


class RateLimiter:
    """OpenDART가 초당 요청 수를 제한할 가능성에 대비해, 전체 스레드가 공유하는
    최소 호출 간격을 강제하는 간단한 속도 제한기."""

    def __init__(self, min_interval_sec: float):
        self.min_interval = min_interval_sec
        self.lock = threading.Lock()
        self.last_call = 0.0

    def wait(self):
        with self.lock:
            now = time.monotonic()
            elapsed = now - self.last_call
            if elapsed < self.min_interval:
                time.sleep(self.min_interval - elapsed)
            self.last_call = time.monotonic()


RATE_LIMITER = RateLimiter(min_interval_sec=1 / 15)


def make_session(total=8, backoff_factor=1.5):
    """일시적인 연결 끊김에 대비해 재시도 로직이 포함된 세션 생성."""
    s = requests.Session()
    s.headers.update({
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
            "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
        ),
        "Connection": "close",
    })
    retry = Retry(
        total=total,
        backoff_factor=backoff_factor,
        status_forcelist=[429, 500, 502, 503, 504],
        allowed_methods=["GET"],
    )
    adapter = HTTPAdapter(max_retries=retry, pool_connections=50, pool_maxsize=50)
    s.mount("https://", adapter)
    s.mount("http://", adapter)
    return s


SIDO_MAP = {
    '서울특별시': '서울', '부산광역시': '부산', '대구광역시': '대구', '인천광역시': '인천',
    '광주광역시': '광주', '대전광역시': '대전', '울산광역시': '울산', '세종특별자치시': '세종',
    '경기도': '경기', '강원도': '강원', '강원특별자치도': '강원', '충청북도': '충북', '충청남도': '충남',
    '전라북도': '전북', '전북특별자치도': '전북', '전라남도': '전남', '경상북도': '경북',
    '경상남도': '경남', '제주특별자치도': '제주'
}


def simplify_address(addr: str) -> str:
    if not addr or not isinstance(addr, str):
        return ''
    addr = addr.strip()
    tokens = addr.split()
    if not tokens:
        return addr
    sido = SIDO_MAP.get(tokens[0], tokens[0])
    gu = tokens[1] if len(tokens) > 1 else ''
    if gu and re.match(r'.+(구|시|군)$', gu):
        return f"{sido} {gu}"
    return sido


def _download_bytes(session, url, timeout=(20, 60), attempts=3):
    """용량이 큰 파일도 안정적으로 받기 위해 스트리밍 + 재시도로 다운로드."""
    last_err = None
    for attempt in range(attempts):
        try:
            with session.get(url, timeout=timeout, stream=True,
                              headers={"Accept-Encoding": "identity"}) as resp:
                resp.raise_for_status()
                chunks = []
                for chunk in resp.iter_content(chunk_size=64 * 1024):
                    if chunk:
                        chunks.append(chunk)
                return b"".join(chunks)
        except requests.exceptions.ConnectionError as e:
            last_err = e
            time.sleep(1.5 * (attempt + 1))
    raise RuntimeError(f"다운로드 실패: {last_err}")


@st.cache_data(show_spinner=False, ttl=60 * 60 * 24)
def load_dart_full_registry():
    """DART 기업개황(업종별) 화면의 '엑셀일괄' 다운로드를 그대로 받아온다.
    상장/비상장 포함 DART 등록법인 전체(약 11만개+)에 실제 업종명이 붙어 있어,
    KRX 상장법인목록보다 훨씬 정확하고 넓은 범위를 커버한다. API 키 불필요."""
    session = make_session(total=3, backoff_factor=1.5)
    # 서버가 세션(쿠키) 기반으로 동작할 수 있어, 검색 화면을 먼저 방문한 뒤 같은 세션으로 다운로드한다
    try:
        session.get("https://dart.fss.or.kr/dsae001/main.do", timeout=(20, 20))
    except Exception:
        pass

    content = _download_bytes(
        session, "https://dart.fss.or.kr/dsae001/downloadExcel.do", timeout=(20, 90)
    )
    df = pd.read_excel(io.BytesIO(content))

    df['종목코드'] = df['종목코드'].astype(str).str.strip()
    df.loc[df['종목코드'] == '', '종목코드'] = None
    df['회사명'] = df['회사이름'].astype(str).str.strip()
    df['본사_위치'] = df['주소'].apply(simplify_address)

    def _year(x):
        try:
            return int(str(x)[:4])
        except Exception:
            return None
    df['설립연도'] = df['설립일'].apply(_year)

    return df


@st.cache_data(show_spinner=False, ttl=60 * 60 * 24)
def load_corp_map(api_key: str):
    """OpenDART corpCode.xml -> {종목코드: corp_code} 매핑. 매출 조회 시에만 필요."""
    import zipfile
    import xml.etree.ElementTree as ET

    url = f"https://opendart.fss.or.kr/api/corpCode.xml?crtfc_key={api_key}"
    session = make_session(total=3, backoff_factor=1.5)
    content = _download_bytes(session, url, timeout=(20, 60))

    zf = zipfile.ZipFile(io.BytesIO(content))
    xml_bytes = zf.read(zf.namelist()[0])
    root = ET.fromstring(xml_bytes)
    corp_map = {}
    for child in root.findall('list'):
        stock_code = child.findtext('stock_code', default='').strip()
        corp_code = child.findtext('corp_code', default='').strip()
        if stock_code:
            corp_map[stock_code] = corp_code
    return corp_map


def get_financials(session, api_key, corp_code, bsns_year, reprt_code):
    """fnlttSinglAcnt(주요계정) API - 매출액/영업이익/당기순이익만 가볍게 조회"""
    url = "https://opendart.fss.or.kr/api/fnlttSinglAcnt.json"
    params = {"crtfc_key": api_key, "corp_code": corp_code, "bsns_year": bsns_year, "reprt_code": reprt_code}
    try:
        RATE_LIMITER.wait()
        r = session.get(url, params=params, timeout=(20, 15))
        data = r.json()
        if data.get('status') != '000':
            reason = f"OpenDART 응답: {data.get('status')} - {data.get('message', '')}"
            return None, None, None, reason
        revenue = revenue_fallback = op_profit = net_income = None
        for it in data.get('list', []):
            nm = it.get('account_nm', '')
            amt_str = it.get('thstrm_amount', '').replace(',', '')
            try:
                amt = int(amt_str)
            except ValueError:
                continue
            if nm in ('매출액', '수익(매출액)') and revenue is None:
                revenue = amt
            elif nm == '영업수익' and revenue_fallback is None:
                revenue_fallback = amt
            elif nm == '영업이익' and op_profit is None:
                op_profit = amt
            elif nm in ('당기순이익', '당기순이익(손실)') and net_income is None:
                net_income = amt
        if revenue is None and revenue_fallback is not None:
            revenue = revenue_fallback
        if revenue is None:
            return None, op_profit, net_income, (
                "OpenDART 응답은 정상이지만 '매출액'에 해당하는 계정을 못 찾음"
            )
        return revenue, op_profit, net_income, None
    except Exception as e:
        return None, None, None, f"연결 오류: {e}"


# ---------------- UI ----------------

st.title("📊 조건 기반 기업 스크리너")
st.caption("DART 등록법인 전체(상장+비상장) 기준으로 업종/지역/설립연도 등을 필터링합니다. "
           "매출 조회는 상장사에 한해 선택적으로 가능합니다.")

with st.spinner("DART 전체법인 목록 불러오는 중... (최초 1회, 파일이 커서 1분 정도 걸릴 수 있습니다)"):
    try:
        registry_df = load_dart_full_registry()
        load_error = None
    except Exception as e:
        registry_df = None
        load_error = str(e)

if load_error:
    st.error(
        f"DART 전체법인 목록을 못 가져왔습니다: {load_error}\n\n"
        "잠시 후 새로고침해서 다시 시도해보세요."
    )
    st.stop()

industry_options = sorted(registry_df['업종명'].dropna().astype(str).str.strip().unique().tolist())
industry_options = [x for x in industry_options if x]
corp_type_options = sorted(registry_df['법인구분'].dropna().astype(str).unique().tolist())

with st.sidebar:
    st.header("🔎 필터 조건")

    industry_select = st.multiselect(
        "업종 선택 (DART 정밀 업종명, 목록에서 고르기)",
        options=industry_options, default=[],
    )
    industry_keywords = st.text_input("업종 키워드 검색 (콤마로 구분)", "")
    company_name_search = st.text_input(
        "회사명 검색 (Enter로 바로 검색)", "",
        key="company_name_search_input",
        on_change=lambda: st.session_state.update({"enter_pressed_search": True}),
    )
    ceo_name_search = st.text_input("대표자명 검색", "")
    corp_type_filter = st.multiselect(
        "법인구분", options=corp_type_options, default=[],
        help="유가증권시장/코스닥시장/코넥스시장 = 상장사, 기타법인 = 비상장(외감대상)",
    )
    region_filter = st.text_input("본사 지역 (예: 서울)", "")
    min_founding_year = st.number_input("설립연도 (이후 설립된 기업만, 0=필터 없음)", 0, 2100, 0)
    top_n = st.number_input("최대 결과 개수", 1, 2000, 200)

    st.markdown("---")
    st.header("💰 매출 조회 (선택, 상장사만 가능)")
    fetch_revenue = st.checkbox("매출/영업이익/순이익도 같이 조회하기 (API 키 필요, 상장사만 해당)", value=False)
    api_key = ""
    bsns_year = "2025"
    min_rev = max_rev = min_op = max_op = min_ni = max_ni = None
    max_workers = 10
    if fetch_revenue:
        api_key = st.text_input("OpenDART API 키", type="password")
        bsns_year = st.text_input("조회 사업연도", "2025")
        st.markdown("**매출액(억원)**")
        c1, c2 = st.columns(2)
        min_rev = c1.number_input("최소", value=0, step=100, key="min_rev")
        max_rev = c2.number_input("최대", value=10000000, step=100, key="max_rev")
        st.markdown("**영업이익(억원)**")
        c1, c2 = st.columns(2)
        min_op = c1.number_input("최소", value=-10000000, step=100, key="min_op")
        max_op = c2.number_input("최대", value=10000000, step=100, key="max_op")
        st.markdown("**당기순이익(억원)**")
        c1, c2 = st.columns(2)
        min_ni = c1.number_input("최소", value=-10000000, step=100, key="min_ni")
        max_ni = c2.number_input("최대", value=10000000, step=100, key="max_ni")
        max_workers = st.number_input("동시 요청 수", 1, 30, 10)

    run_btn_clicked = st.button("🚀 검색 실행", type="primary", use_container_width=True)
    run_btn = run_btn_clicked or st.session_state.get("enter_pressed_search", False)

COOLDOWN_SEC = 65
if st.session_state.get("still_running", False):
    elapsed = time.time() - st.session_state.get("last_dispatch_ts", 0)
    remaining = COOLDOWN_SEC - elapsed
    if remaining > 0:
        st.warning(f"⏳ 이전 검색이 중간에 중단된 것 같습니다. 약 **{int(remaining)}초** 더 기다려주세요.")
        time.sleep(1)
        st.rerun()
    else:
        st.session_state["still_running"] = False

if run_btn:
    st.session_state["enter_pressed_search"] = False

    if fetch_revenue and not api_key:
        st.warning("매출 조회를 켜셨다면 OpenDART API 키를 입력해주세요.")
        st.stop()

    candidates = registry_df.copy()
    kw_list = [k.strip() for k in industry_keywords.split(',') if k.strip()]

    if company_name_search.strip():
        candidates = candidates[candidates['회사명'].str.contains(company_name_search.strip(), na=False)]
    elif industry_select or kw_list:
        mask = pd.Series(False, index=candidates.index)
        if industry_select:
            mask = mask | candidates['업종명'].isin(industry_select)
        if kw_list:
            pattern = '|'.join(kw_list)
            mask = mask | candidates['업종명'].astype(str).str.contains(pattern, na=False)
        candidates = candidates[mask]

    if ceo_name_search.strip():
        candidates = candidates[candidates['대표자명'].astype(str).str.contains(ceo_name_search.strip(), na=False)]

    if corp_type_filter:
        candidates = candidates[candidates['법인구분'].isin(corp_type_filter)]

    if min_founding_year and min_founding_year > 0:
        candidates = candidates[candidates['설립연도'].isna() | (candidates['설립연도'] >= min_founding_year)]

    if region_filter.strip():
        candidates = candidates[candidates['본사_위치'].str.contains(region_filter.strip(), na=False)]

    candidates = candidates.reset_index(drop=True)
    st.write(f"1차 필터 후 후보 기업 수: **{len(candidates)}**")

    if len(candidates) == 0:
        st.warning("조건에 맞는 후보가 없습니다. 필터를 완화해보세요.")
        st.stop()

    final = candidates.copy()
    final['매출액(억원)'] = None
    final['영업이익(억원)'] = None
    final['당기순이익(억원)'] = None
    final['실패사유'] = None

    if fetch_revenue:
        st.session_state["still_running"] = True
        st.session_state["last_dispatch_ts"] = time.time()
        try:
            corp_map = load_corp_map(api_key)
        except Exception as e:
            st.error(f"OpenDART corp_code 매핑을 못 가져왔습니다: {e}")
            st.session_state["still_running"] = False
            st.stop()

        listed = final[final['종목코드'].notna()].copy()
        unlisted_count = len(final) - len(listed)
        if unlisted_count > 0:
            st.caption(f"비상장(기타법인) {unlisted_count}개사는 매출 조회 대상이 아니라 매출 없이 표시됩니다.")

        session = make_session(total=2, backoff_factor=0.5)
        progress = st.progress(0.0, text="매출/영업이익/순이익 조회 중...")
        results = {}
        executor = ThreadPoolExecutor(max_workers=int(max_workers))
        try:
            future_map = {}
            for idx, row in listed.iterrows():
                corp_code = corp_map.get(row['종목코드'])
                if not corp_code:
                    continue
                fut = executor.submit(get_financials, session, api_key, corp_code, bsns_year, "11011")
                future_map[fut] = idx
            done = 0
            for fut in as_completed(future_map):
                idx = future_map[fut]
                revenue, op_profit, net_income, reason = fut.result()
                results[idx] = (revenue, op_profit, net_income, reason)
                done += 1
                if len(future_map):
                    progress.progress(done / len(future_map),
                                       text=f"매출/영업이익/순이익 조회 중... ({done}/{len(future_map)})")
        finally:
            executor.shutdown(wait=False, cancel_futures=True)
        progress.empty()
        st.session_state["still_running"] = False

        for idx, (revenue, op_profit, net_income, reason) in results.items():
            final.at[idx, '매출액(억원)'] = (revenue / 1e8) if revenue is not None else None
            final.at[idx, '영업이익(억원)'] = (op_profit / 1e8) if op_profit is not None else None
            final.at[idx, '당기순이익(억원)'] = (net_income / 1e8) if net_income is not None else None
            final.at[idx, '실패사유'] = reason

        no_data = final[final['종목코드'].notna() & final['매출액(억원)'].isna()]
        if not no_data.empty:
            with st.expander(f"⚠️ 매출 데이터를 못 가져온 상장사 {len(no_data)}개 (클릭해서 보기)"):
                st.dataframe(no_data[['회사명', '종목코드', '실패사유']], use_container_width=True)

        for col, mn, mx in [('매출액(억원)', min_rev, max_rev),
                             ('영업이익(억원)', min_op, max_op),
                             ('당기순이익(억원)', min_ni, max_ni)]:
            final[col] = pd.to_numeric(final[col], errors='coerce')
            final = final[final[col].isna() | ((final[col] >= mn) & (final[col] <= mx))]

    final = final.sort_values(
        '매출액(억원)', ascending=False, na_position='last'
    ).head(int(top_n)).reset_index(drop=True)
    for col in ['매출액(억원)', '영업이익(억원)', '당기순이익(억원)']:
        final[col] = pd.to_numeric(final[col], errors='coerce').round(0).astype('Int64')

    output_df = final.rename(columns={
        '업종명': '업종', '홈페이지': '홈페이지 주소', '본사_위치': '본사 위치',
    })
    output_df['관련기사'] = output_df['회사명'].apply(
        lambda name: f"https://search.naver.com/search.naver?where=news&query={quote(str(name))}"
    )
    output_df = output_df[['회사명', '관련기사', '업종', '법인구분', '대표자명',
                            '매출액(억원)', '영업이익(억원)', '당기순이익(억원)',
                            '설립연도', '홈페이지 주소', '본사 위치']]

    st.success(f"최종 {len(output_df)}개사")

    def normalize_url(u):
        if not isinstance(u, str) or not u.strip():
            return None
        u = u.strip()
        if not u.startswith('http://') and not u.startswith('https://'):
            u = 'https://' + u
        return u

    output_df['홈페이지 주소'] = output_df['홈페이지 주소'].apply(normalize_url)

    display_df = output_df.copy()
    for col in ['매출액(억원)', '영업이익(억원)', '당기순이익(억원)']:
        display_df[col] = display_df[col].apply(lambda x: f"{int(x):,}" if pd.notna(x) else "")

    st.dataframe(
        display_df,
        use_container_width=True,
        column_config={
            "홈페이지 주소": st.column_config.LinkColumn("홈페이지 주소", display_text="바로가기"),
            "관련기사": st.column_config.LinkColumn("관련기사", display_text="기사보기"),
        },
    )

    buf = io.BytesIO()
    output_df.to_excel(buf, index=False)
    st.download_button(
        "📥 엑셀로 다운로드",
        data=buf.getvalue(),
        file_name="screener_result.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
else:
    st.info("왼쪽에서 조건을 입력하고 '검색 실행' 버튼을 누르세요.")
