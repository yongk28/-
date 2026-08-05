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
    '서울특별시': '서울', '부산광역시': '부산', '대구광역시': '대구', '인천광역시': '인천',
    '광주광역시': '광주', '대전광역시': '대전', '울산광역시': '울산', '세종특별자치시': '세종',
    '경기도': '경기', '강원도': '강원', '강원특별자치도': '강원', '충청북도': '충북', '충청남도': '충남',
    '전라북도': '전북', '전북특별자치도': '전북', '전라남도': '전남', '경상북도': '경북',
    '경상남도': '경남', '제주특별자치도': '제주'
}


def simplify_address(addr: str) -> str:
    if not addr:
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


@st.cache_data(show_spinner=False, ttl=60 * 60 * 24)
def load_kind_listing():
    """KRX KIND 상장법인목록 전체 다운로드 (하루 1회만 다시 받음)"""
    kind_url = "https://kind.krx.co.kr/corpgeneral/corpList.do?method=download&searchType=13"
    df = pd.read_html(kind_url, header=0, encoding='euc-kr')[0]
    df['종목코드'] = df['종목코드'].astype(str).str.zfill(6)
    df = df.rename(columns={'주요제품': '대표상품_브랜드', '홈페이지': '홈페이지_주소'})

    import FinanceDataReader as fdr
    market_map = {}
    for m in ['KOSPI', 'KOSDAQ', 'KONEX']:
        try:
            df_m = fdr.StockListing(m)
            code_col_m = 'Code' if 'Code' in df_m.columns else '종목코드'
            for c in df_m[code_col_m].astype(str).str.zfill(6):
                market_map[c] = m
        except Exception:
            pass
    df['시장구분'] = df['종목코드'].map(market_map).fillna('기타')
    return df


@st.cache_data(show_spinner=False, ttl=60 * 60 * 24)
def load_corp_map(api_key: str):
    """OpenDART corpCode.xml -> {종목코드: corp_code} 매핑 (하루 1회만 다시 받음)"""
    url = f"https://opendart.fss.or.kr/api/corpCode.xml?crtfc_key={api_key}"
    s = make_session()
    try:
        resp = s.get(url, timeout=30)
        resp.raise_for_status()
    except requests.exceptions.ConnectionError as e:
        raise RuntimeError(
            "OpenDART 서버에 연결할 수 없습니다. 네트워크(방화벽/백신/사내망) 문제일 수 있습니다. "
            "브라우저로 같은 주소가 열리는지 먼저 확인해보세요."
        ) from e
    zf = zipfile.ZipFile(io.BytesIO(resp.content))
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
    url = "https://opendart.fss.or.kr/api/fnlttSinglAcntAll.json"
    params = {"crtfc_key": api_key, "corp_code": corp_code, "bsns_year": bsns_year,
              "reprt_code": reprt_code, "fs_div": "CFS"}
    try:
        r = session.get(url, params=params, timeout=10)
        data = r.json()
        if data.get('status') != '000':
            params['fs_div'] = 'OFS'
            r = session.get(url, params=params, timeout=10)
            data = r.json()
        if data.get('status') != '000':
            return None, None, None
        revenue = op_profit = net_income = None
        for it in data.get('list', []):
            nm = it.get('account_nm', '')
            amt_str = it.get('thstrm_amount', '').replace(',', '')
            try:
                amt = int(amt_str)
            except ValueError:
                continue
            if nm in ('매출액', '수익(매출액)') and revenue is None:
                revenue = amt
            elif nm == '영업이익' and op_profit is None:
                op_profit = amt
            elif nm in ('당기순이익', '당기순이익(손실)') and net_income is None:
                net_income = amt
        return revenue, op_profit, net_income
    except Exception:
        return None, None, None


def get_company_detail(session, api_key, corp_code):
    url = "https://opendart.fss.or.kr/api/company.json"
    params = {"crtfc_key": api_key, "corp_code": corp_code}
    try:
        r = session.get(url, params=params, timeout=10)
        data = r.json()
        if data.get('status') == '000':
            return data.get('adres', ''), data.get('est_dt', '')
        return '', ''
    except Exception:
        return '', ''


def get_employee_count(session, api_key, corp_code, bsns_year, reprt_code):
    url = "https://opendart.fss.or.kr/api/empSttus.json"
    params = {"crtfc_key": api_key, "corp_code": corp_code, "bsns_year": bsns_year, "reprt_code": reprt_code}
    try:
        r = session.get(url, params=params, timeout=10)
        data = r.json()
        if data.get('status') != '000':
            return None
        total, found = 0, False
        for it in data.get('list', []):
            for key in ('rgllbr_co', 'cnttk_co'):
                s = (it.get(key, '') or '').replace(',', '').strip()
                if s.isdigit():
                    total += int(s)
                    found = True
        return total if found else None
    except Exception:
        return None


def get_founding_year(est_dt):
    try:
        return int(est_dt[:4])
    except Exception:
        return None


# ---------------- UI ----------------

st.title("📊 조건 기반 기업 스크리너")
st.caption("KRX 상장법인 + OpenDART 공시 데이터를 기반으로 조건에 맞는 기업을 찾습니다.")

# API 키 없이도 되는 KRX 목록은 앱 시작할 때 미리 불러와서, 업종 선택 목록을 만드는 데 씁니다.
# (하루 1회만 실제로 다시 받아오고 그 외엔 캐시를 사용하므로 매번 느려지지 않습니다)
with st.spinner("업종 목록 불러오는 중..."):
    try:
        _kind_df_for_ui = load_kind_listing()
        industry_options = sorted(_kind_df_for_ui['업종'].dropna().astype(str).unique().tolist())
    except Exception:
        _kind_df_for_ui = None
        industry_options = []

with st.sidebar:
    st.header("🔑 API 키")
    api_key = st.text_input("OpenDART API 키", type="password", help="opendart.fss.or.kr 에서 발급받은 키")

    if st.button("🛑 앱 완전히 종료", use_container_width=True):
        st.warning("앱을 종료합니다. 이 브라우저 탭은 이제 닫으셔도 됩니다.")
        import os
        os._exit(0)

    st.markdown("---")
    st.header("🔎 필터 조건")

    industry_select = st.multiselect(
        "업종 선택 (목록에서 고르기)",
        options=industry_options,
        default=[],
        help="KRX에 등록된 실제 업종명 중에서 정확히 골라서 필터링합니다.",
    )
    industry_keywords = st.text_input(
        "업종/주요제품 키워드 검색 (콤마로 구분, 위 목록에 없는 것도 잡고 싶을 때)", ""
    )
    whitelist_names = st.text_input("강제 포함 회사명 (콤마로 구분)", "")
    company_name_search = st.text_input("회사명 검색", "")
    ceo_name_search = st.text_input("대표자명 검색", "")
    market_filter = st.multiselect("시장구분", ["KOSPI", "KOSDAQ", "KONEX"], default=[])
    settlement_month = st.text_input("결산월 (예: 12)", "")

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

    st.markdown("**직원수**")
    fetch_employees = st.checkbox("직원수 조회하기 (API 호출이 추가로 발생, 네트워크 불안정하면 꺼두세요)", value=False)
    c1, c2 = st.columns(2)
    min_emp = c1.number_input("최소", value=0, step=10, key="min_emp", disabled=not fetch_employees)
    max_emp = c2.number_input("최대", value=10000000, step=10, key="max_emp", disabled=not fetch_employees)

    min_founding_year = st.number_input(
        "설립연도 (이 연도 이후에 생긴 기업만, 0이면 필터 없음)", 0, 2100, 0
    )
    region_filter = st.text_input("본사 지역 (예: 서울)", "")

    bsns_year = st.text_input("조회 사업연도", "2025")
    top_n = st.number_input("최대 결과 개수", 1, 500, 100)
    max_workers = st.number_input("동시 요청 수", 1, 30, 5)

    run_btn = st.button("🚀 검색 실행", type="primary", use_container_width=True)

if not api_key:
    st.info("왼쪽 사이드바에 OpenDART API 키를 입력하면 시작할 수 있습니다.")
    st.stop()

if run_btn:
    with st.spinner("KRX 상장법인 목록 로딩 중... (최초 1회, 캐시되어 다음부터는 빠릅니다)"):
        try:
            kind_df = load_kind_listing()
            corp_map = load_corp_map(api_key)
        except RuntimeError as e:
            st.error(str(e))
            st.stop()
        except requests.exceptions.ConnectionError:
            st.error(
                "서버 연결이 중간에 끊겼습니다 (네트워크/방화벽/백신 문제일 가능성). "
                "잠시 후 다시 시도하시거나, 다른 네트워크(휴대폰 핫스팟 등)에서 시도해보세요."
            )
            st.stop()

    candidates = kind_df.copy()
    kw_list = [k.strip() for k in industry_keywords.split(',') if k.strip()]
    wl_list = [k.strip() for k in whitelist_names.split(',') if k.strip()]

    if industry_select or kw_list:
        mask = pd.Series(False, index=candidates.index)
        if industry_select:
            mask = mask | candidates['업종'].isin(industry_select)
        if kw_list:
            pattern = '|'.join(kw_list)
            mask = mask | candidates['업종'].astype(str).str.contains(pattern, na=False)
            mask = mask | candidates['대표상품_브랜드'].astype(str).str.contains(pattern, na=False)
        candidates = candidates[mask]

    if company_name_search.strip():
        candidates = candidates[candidates['회사명'].astype(str).str.contains(company_name_search.strip(), na=False)]

    if ceo_name_search.strip() and '대표자명' in candidates.columns:
        candidates = candidates[candidates['대표자명'].astype(str).str.contains(ceo_name_search.strip(), na=False)]

    if market_filter:
        candidates = candidates[candidates['시장구분'].isin(market_filter)]

    if settlement_month.strip() and '결산월' in candidates.columns:
        candidates = candidates[candidates['결산월'].astype(str).str.contains(settlement_month.strip(), na=False)]

    if wl_list:
        wl_matches = kind_df[kind_df['회사명'].isin(wl_list)]
        candidates = pd.concat([candidates, wl_matches]).drop_duplicates(subset='종목코드')

    candidates = candidates.reset_index(drop=True)
    st.write(f"1차 필터 후 후보 기업 수: **{len(candidates)}**")

    if len(candidates) == 0:
        st.warning("조건에 맞는 후보가 없습니다. 필터를 완화해보세요.")
        st.stop()

    session = make_session()

    progress = st.progress(0.0, text="매출/영업이익/순이익 조회 중...")
    rows = [row for _, row in candidates.iterrows()]
    results = []
    with ThreadPoolExecutor(max_workers=int(max_workers)) as executor:
        future_map = {}
        for row in rows:
            code = row['종목코드']
            corp_code = corp_map.get(code)
            if not corp_code:
                continue
            fut = executor.submit(get_financials, session, api_key, corp_code, bsns_year, "11011")
            future_map[fut] = (row, corp_code)
        done = 0
        for fut in as_completed(future_map):
            row, corp_code = future_map[fut]
            revenue, op_profit, net_income = fut.result()
            results.append({
                '종목코드': row['종목코드'], 'corp_code': corp_code, '기업명': row['회사명'], '업종': row['업종'],
                '대표상품_브랜드': row.get('대표상품_브랜드'), '홈페이지_주소': row.get('홈페이지_주소'),
                '시장구분': row.get('시장구분'), '대표자명': row.get('대표자명'), '결산월': row.get('결산월'),
                '매출액_억': (revenue / 1e8) if revenue is not None else None,
                '영업이익_억': (op_profit / 1e8) if op_profit is not None else None,
                '당기순이익_억': (net_income / 1e8) if net_income is not None else None,
            })
            done += 1
            progress.progress(done / len(future_map), text=f"매출/영업이익/순이익 조회 중... ({done}/{len(future_map)})")
    progress.empty()

    df_result = pd.DataFrame(results)
    if df_result.empty:
        st.warning("매출 데이터를 확보한 회사가 없습니다.")
        st.stop()

    filtered = df_result.dropna(subset=['매출액_억']).copy()
    filtered = filtered[(filtered['매출액_억'] >= min_rev) & (filtered['매출액_억'] <= max_rev)]
    filtered = filtered[filtered['영업이익_억'].isna() | ((filtered['영업이익_억'] >= min_op) & (filtered['영업이익_억'] <= max_op))]
    filtered = filtered[filtered['당기순이익_억'].isna() | ((filtered['당기순이익_억'] >= min_ni) & (filtered['당기순이익_억'] <= max_ni))]
    filtered = filtered.reset_index(drop=True)

    st.write(f"매출/영업이익/순이익 조건 통과: **{len(filtered)}**개사")

    if len(filtered) == 0:
        st.warning("조건에 맞는 회사가 없습니다. 필터를 완화해보세요.")
        st.stop()

    progress2 = st.progress(0.0, text="주소/설립일 조회 중...")
    corp_codes = filtered['corp_code'].tolist()
    details = [None] * len(corp_codes)
    employees = [None] * len(corp_codes)

    with ThreadPoolExecutor(max_workers=int(max_workers)) as executor:
        future_map = {executor.submit(get_company_detail, session, api_key, cc): i for i, cc in enumerate(corp_codes)}
        done = 0
        for fut in as_completed(future_map):
            idx = future_map[fut]
            details[idx] = fut.result()
            done += 1
            progress2.progress(done / len(corp_codes), text=f"주소/설립일 조회 중... ({done}/{len(corp_codes)})")
    progress2.empty()

    if fetch_employees:
        progress3 = st.progress(0.0, text="직원수 조회 중...")
        with ThreadPoolExecutor(max_workers=int(max_workers)) as executor:
            future_map = {executor.submit(get_employee_count, session, api_key, cc, bsns_year, "11011"): i
                          for i, cc in enumerate(corp_codes)}
            done = 0
            for fut in as_completed(future_map):
                idx = future_map[fut]
                employees[idx] = fut.result()
                done += 1
                progress3.progress(done / len(corp_codes), text=f"직원수 조회 중... ({done}/{len(corp_codes)})")
        progress3.empty()
    else:
        st.caption("직원수 조회를 건너뛰었습니다 (체크박스 꺼짐). 직원수 칸은 비어있게 나옵니다.")

    filtered['본사_위치'] = [simplify_address(d[0]) for d in details]
    filtered['설립일'] = [d[1] for d in details]
    filtered['직원수'] = employees
    filtered['설립연도'] = filtered['설립일'].apply(get_founding_year)

    final = filtered.copy()
    final = final[final['직원수'].isna() | ((final['직원수'] >= min_emp) & (final['직원수'] <= max_emp))]
    if min_founding_year and min_founding_year > 0:
        final = final[final['설립연도'].isna() | (final['설립연도'] >= min_founding_year)]

    if region_filter.strip():
        final = final[final['본사_위치'].str.contains(region_filter.strip(), na=False)]

    final = final.sort_values('매출액_억', ascending=False).head(int(top_n)).reset_index(drop=True)
    for col in ['매출액_억', '영업이익_억', '당기순이익_억']:
        final[col] = final[col].round(0).astype('Int64')

    output_df = final.rename(columns={
        '대표상품_브랜드': '대표상품 or 브랜드', '매출액_억': '매출액(억원)', '영업이익_억': '영업이익(억원)',
        '당기순이익_억': '당기순이익(억원)', '홈페이지_주소': '홈페이지 주소', '본사_위치': '본사 위치'
    })
    output_df['관련기사'] = output_df['기업명'].apply(
        lambda name: f"https://search.naver.com/search.naver?where=news&query={quote(str(name))}"
    )
    output_df = output_df[['기업명', '관련기사', '업종', '대표상품 or 브랜드', '시장구분', '매출액(억원)', '영업이익(억원)',
                            '당기순이익(억원)', '직원수', '설립연도', '홈페이지 주소', '본사 위치']]

    st.success(f"최종 {len(output_df)}개사")

    # 홈페이지 주소에 http(s):// 접두어가 없으면 붙여서, 링크가 실제로 동작하게 함
    def normalize_url(u):
        if not isinstance(u, str) or not u.strip():
            return None
        u = u.strip()
        if not u.startswith('http://') and not u.startswith('https://'):
            u = 'https://' + u
        return u

    output_df['홈페이지 주소'] = output_df['홈페이지 주소'].apply(normalize_url)

    # 화면에 보여줄 때만 천 단위 콤마를 넣은 문자열로 변환 (엑셀 파일은 숫자 그대로 유지)
    display_df = output_df.copy()
    for col in ['매출액(억원)', '영업이익(억원)', '당기순이익(억원)']:
        display_df[col] = display_df[col].apply(lambda x: f"{int(x):,}" if pd.notna(x) else "")
    if '직원수' in display_df.columns:
        display_df['직원수'] = display_df['직원수'].apply(lambda x: f"{int(x):,}" if pd.notna(x) else "")

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
