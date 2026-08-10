# -*- coding: utf-8 -*-
"""
=============================================================================
 나라장터(조달청) 입찰정보 자동 수집 & 이메일 발송 프로그램
=============================================================================

 [ 이 프로그램이 하는 일 ]
   1. 조달청 나라장터 Open API에 접속해서
   2. 정해진 기간(기본: 전날 오전 10시 ~ 오늘 오전 10시)에 새로 올라온
      [사업계획(발주계획)] / [사전규격] / [본공고] 를 모두 가져오고
   3. 내가 정한 검색어(KEYWORDS)가 들어간 건만 골라내서
   4. 엑셀 파일(.xlsx) 하나로 예쁘게 정리한 뒤
   5. 내 이메일로 첨부해서 보내줍니다.

 [ 코딩을 몰라도 괜찮습니다 ]
   아래 "★★★ 사용자 설정 영역 ★★★" 안의 값만 바꾸면 됩니다.
   그 아래쪽(=== 여기부터는 프로그램 본체 ===)은 건드리지 않아도 됩니다.
=============================================================================
"""

import os                    # 환경변수(깃허브 Secrets)를 읽기 위한 도구
import re                    # 글자 정리(공백 제거 등)를 위한 도구
import ssl                   # 이메일을 안전하게 보내기 위한 도구
import sys                   # 프로그램 종료 처리를 위한 도구
import time                  # 잠깐 쉬어가기(API 과부하 방지)를 위한 도구
import smtplib               # 이메일 발송(SMTP) 도구
import traceback             # 오류가 났을 때 원인을 자세히 보여주는 도구
from datetime import datetime, timedelta, timezone
from email.message import EmailMessage

import requests              # 인터넷(API)에서 데이터를 받아오는 도구
import pandas as pd          # 표(엑셀) 데이터를 다루는 도구


# =============================================================================
# ★★★ 사용자 설정 영역 ★★★  (여기만 편하게 수정하세요)
# =============================================================================

# ---------------------------------------------------------------------------
# 1) 검색어(키워드) 목록
#    - 공고 제목 / 품명 / 사업내용 안에 아래 단어 중 "하나라도" 들어있으면 수집합니다.
#    - 추가하고 싶으면 줄을 하나 늘리고 '단어', 형식으로 적으면 됩니다.
#    - 빼고 싶으면 그 줄 맨 앞에 # 을 붙이거나 줄을 통째로 지우면 됩니다.
#    - 띄어쓰기는 자동으로 무시됩니다. ('실험 기기' 와 '실험기기' 를 같은 걸로 봅니다)
# ---------------------------------------------------------------------------
KEYWORDS = [
    # --- 핵심 장비 ---
    '실험기기',
    '실험장비',
    '이화학',
    '인큐베이터',
    '배양기',
    'CO2배양기',
    '항온',
    '항습',
    '챔버',
    '오븐',
    '건조기',
    '멸균기',
    '고압증기',
    '초저온',
    '냉동고',
    '디프프리저',
    '원심분리기',
    '진탕기',
    '쉐이커',
    '교반기',
    '현미경',
    '분광광도계',
    '클린벤치',
    '안전캐비닛',
    '흄후드',
    '무균작업대',

    # --- 넓은 범위(연구·바이오·시험) ---
    '연구장비',
    '연구기자재',
    '실험실',
    '실험대',
    '바이오',
    '세포',
    '시험장비',
    '분석장비',
    '분석기기',
    '계측장비',
    '시약',
    '진단',
    '생명공학',
    '유전자',
    '항온항습기',
]

# ---------------------------------------------------------------------------
# 2) 제외 키워드 (선택)
#    - 위 키워드에 걸렸더라도, 아래 단어가 들어있으면 결과에서 빼버립니다.
#    - 예: 청소용역, 경비용역처럼 관련 없는 공고가 자꾸 걸릴 때 사용하세요.
#    - 필요 없으면 아래처럼 비워두면 됩니다:  EXCLUDE_KEYWORDS = []
# ---------------------------------------------------------------------------
EXCLUDE_KEYWORDS = [
    # '청소',
    # '경비용역',
    # '급식',
]

# ---------------------------------------------------------------------------
# 3) 수집할 단계 (True = 수집함 / False = 수집 안 함)
#    - 필요 없는 단계는 False 로 바꾸면 그만큼 빨라집니다.
# ---------------------------------------------------------------------------
COLLECT_STAGES = {
    '사업계획': True,   # 발주계획(연간/분기 발주 예정 사업)
    '사전규격': True,   # 규격 사전공개(본공고 나오기 전 단계) ★영업 골든타임
    '본공고': True,     # 실제 입찰공고
}

# ---------------------------------------------------------------------------
# 4) 수집할 업무 구분
#    - 보통 장비 영업은 '물품'과 '용역'이면 충분합니다.
#    - 공사까지 보고 싶으면 '공사' 줄의 # 을 지우세요.
# ---------------------------------------------------------------------------
BUSINESS_TYPES = [
    '물품',
    '용역',
    # '공사',
]

# ---------------------------------------------------------------------------
# 5) 수집 기준 시각 (매일 몇 시를 기준으로 "하루"를 자를지)
#    - 10 이면 "어제 10:00 ~ 오늘 10:00" 입니다.
#    - 깃허브 액션 실행 시간도 같이 바꿔야 합니다(.yml 파일의 cron 참고).
# ---------------------------------------------------------------------------
DAILY_CUTOFF_HOUR = 10

# ---------------------------------------------------------------------------
# 6) 메일 제목 앞에 붙일 말머리
# ---------------------------------------------------------------------------
MAIL_SUBJECT_PREFIX = '[나라장터 입찰정보]'

# =============================================================================
# ★★★ 사용자 설정 영역 끝 ★★★
#     아래부터는 건드리지 않으셔도 됩니다.
# =============================================================================


# -----------------------------------------------------------------------------
# 기본 상수들
# -----------------------------------------------------------------------------
KST = timezone(timedelta(hours=9))          # 한국 표준시(UTC+9)
API_BASE = 'http://apis.data.go.kr/1230000'  # 조달청 Open API 공통 주소
PAGE_SIZE = 500        # 한 번 요청할 때 가져올 건수 (너무 크면 실패할 수 있어 500)
MAX_PAGES = 30         # 한 종류당 최대 페이지 수 (안전장치: 무한루프 방지)
REQUEST_TIMEOUT = 30   # 응답을 기다리는 최대 시간(초)
SLEEP_BETWEEN_CALLS = 0.3  # API 호출 사이 쉬는 시간(초) - 과부하 방지

SERVICE_KEY = ''       # 공공데이터포털 인증키 (실행할 때 Secrets 에서 채워집니다)

# 단계별 API 주소 목록
# - 앞의 것부터 시도하고, 실패하면 다음 것으로 자동 전환합니다(자동 백업 경로).
API_ENDPOINTS = {
    '본공고': {
        '물품': ['ad/BidPublicInfoService/getBidPblancListInfoThng',
                 'ad/BidPublicInfoService/getBidPblancListInfoThngPPSSrch'],
        '용역': ['ad/BidPublicInfoService/getBidPblancListInfoServc',
                 'ad/BidPublicInfoService/getBidPblancListInfoServcPPSSrch'],
        '공사': ['ad/BidPublicInfoService/getBidPblancListInfoCnstwk',
                 'ad/BidPublicInfoService/getBidPblancListInfoCnstwkPPSSrch'],
    },
    '사전규격': {
        '물품': ['ao/HrcspSsstndrdInfoService/getPublicPrcureThngInfoThng',
                 'ao/HrcspSsstndrdInfoService/getPublicPrcureThngInfoThngPPSSrch'],
        '용역': ['ao/HrcspSsstndrdInfoService/getPublicPrcureThngInfoServc',
                 'ao/HrcspSsstndrdInfoService/getPublicPrcureThngInfoServcPPSSrch'],
        '공사': ['ao/HrcspSsstndrdInfoService/getPublicPrcureThngInfoCnstwk',
                 'ao/HrcspSsstndrdInfoService/getPublicPrcureThngInfoCnstwkPPSSrch'],
    },
    # 발주계획(사업계획)은 조달청에서 서비스 주소가 바뀐 적이 있어
    # 가능성 있는 주소를 여러 개 넣어두고 되는 것을 자동으로 찾아 씁니다.
    '사업계획': {
        '물품': ['at/OrderPlanSttusService/getOrderPlanSttusListThng',
                 'ao/OrderPlanSttusService/getOrderPlanSttusListThng',
                 'at/PubPrcrmntOrderPlanService/getOrderPlanSttusListThng'],
        '용역': ['at/OrderPlanSttusService/getOrderPlanSttusListServc',
                 'ao/OrderPlanSttusService/getOrderPlanSttusListServc',
                 'at/PubPrcrmntOrderPlanService/getOrderPlanSttusListServc'],
        '공사': ['at/OrderPlanSttusService/getOrderPlanSttusListCnstwk',
                 'ao/OrderPlanSttusService/getOrderPlanSttusListCnstwk',
                 'at/PubPrcrmntOrderPlanService/getOrderPlanSttusListCnstwk'],
    },
}

# 엑셀에 넣을 열(컬럼) 순서
COLUMNS = [
    '구분', '업무', '공고명', '수요기관', '공고기관',
    '등록/공고일시', '사업기한(마감일)', '배정예산(원)',
    '사업내용요약', '매칭키워드', '공고번호', '공고링크',
]


# -----------------------------------------------------------------------------
# [도우미 함수] 로그(진행상황) 출력
# -----------------------------------------------------------------------------
def log(message):
    """실행 중 무슨 일이 일어나는지 화면에 찍어줍니다. (깃허브 액션 로그에서 확인 가능)"""
    now = datetime.now(KST).strftime('%H:%M:%S')
    print(f'[{now}] {message}', flush=True)


# -----------------------------------------------------------------------------
# [도우미 함수] 환경변수(깃허브 Secrets) 읽기
# -----------------------------------------------------------------------------
def get_env(name, required=True, default=''):
    """
    깃허브 Secrets 에 등록한 값을 읽어옵니다.
    required=True 인데 값이 없으면 프로그램을 즉시 멈추고 이유를 알려줍니다.
    """
    value = os.environ.get(name, default)
    value = (value or '').strip()
    if required and not value:
        log(f'❌ 필수 설정값 "{name}" 이(가) 비어 있습니다. 깃허브 Secrets 를 확인하세요.')
        sys.exit(1)
    return value


# -----------------------------------------------------------------------------
# [핵심 1] 수집 기간 계산하기
# -----------------------------------------------------------------------------
def calc_period(now=None, lookback_days=0):
    """
    '언제부터 언제까지' 올라온 공고를 가져올지 계산합니다.

    기본 규칙
      - 끝 시각  : 지금(실행 시각)
      - 시작 시각: 어제 오전 10시  (DAILY_CUTOFF_HOUR 값)
      - 월요일이면: 금요일 오전 10시부터 (→ 금·토·일 공고를 모두 포함)

    lookback_days 가 1 이상이면(수동 실행 시 입력) 그 일수만큼 거슬러 올라갑니다.
    """
    if now is None:
        now = datetime.now(KST)

    end_dt = now

    # (1) 수동 실행에서 "최근 N일" 을 지정한 경우 → 그 값을 최우선으로 사용
    if lookback_days and lookback_days > 0:
        start_dt = (now - timedelta(days=lookback_days)).replace(
            hour=DAILY_CUTOFF_HOUR, minute=0, second=0, microsecond=0)
        return start_dt, end_dt

    # (2) 기본 로직: 오늘 기준 며칠 전부터 볼 것인가?
    #     weekday(): 월=0, 화=1, 수=2, 목=3, 금=4, 토=5, 일=6
    weekday = now.weekday()
    if weekday == 0:        # 월요일 → 3일 전(금요일)부터
        days_back = 3
    elif weekday == 6:      # 일요일 → 2일 전(금요일)부터 (주말 수동 실행 대비)
        days_back = 2
    else:                   # 화~토 → 하루 전부터
        days_back = 1

    start_dt = (now - timedelta(days=days_back)).replace(
        hour=DAILY_CUTOFF_HOUR, minute=0, second=0, microsecond=0)

    # 안전장치: 혹시 시작이 끝보다 뒤면(새벽 실행 등) 하루 더 뒤로 밀어줍니다.
    if start_dt >= end_dt:
        start_dt = start_dt - timedelta(days=1)

    return start_dt, end_dt


def fmt_api_datetime(dt):
    """API가 요구하는 날짜 형식(yyyyMMddHHmm)으로 바꿔줍니다. 예) 202608091000"""
    return dt.strftime('%Y%m%d%H%M')


# -----------------------------------------------------------------------------
# [핵심 2] API 호출
# -----------------------------------------------------------------------------
def clean_service_key(raw_key):
    """
    공공데이터포털 인증키는 '일반 인증키(Decoding)' 를 써야 합니다.
    혹시 Encoding 키(%2B, %3D 같은 게 섞인 키)를 넣었다면 자동으로 풀어줍니다.
    """
    from urllib.parse import unquote
    key = raw_key.strip()
    if '%' in key:
        key = unquote(key)
    return key


def call_api(path, params, page_size):
    """
    조달청 API를 1회 호출하고 결과(JSON)를 돌려줍니다.
    실패하면 None 을 돌려주고, 이유를 로그에 남깁니다.
    """
    url = f'{API_BASE}/{path}'
    query = dict(params)
    query['numOfRows'] = page_size
    query['type'] = 'json'

    try:
        response = requests.get(url, params=query, timeout=REQUEST_TIMEOUT)
    except Exception as error:
        log(f'   ⚠️ 통신 실패({path}): {error}')
        return None

    if response.status_code != 200:
        log(f'   ⚠️ 서버 응답코드 {response.status_code} ({path})')
        return None

    # 정상이면 JSON, 문제가 있으면 XML 형태의 오류문이 오는 경우가 있습니다.
    try:
        data = response.json()
    except Exception:
        snippet = response.text[:250].replace('\n', ' ')
        log(f'   ⚠️ JSON이 아닌 응답이 왔습니다 ({path}) → {snippet}')
        return None

    # 응답 안의 결과코드 확인 ('00' 이면 정상)
    header = (data.get('response', {}) or {}).get('header', {}) or {}
    result_code = str(header.get('resultCode', '')).strip()
    if result_code and result_code not in ('00', '0'):
        log(f'   ⚠️ API 오류 [{result_code}] {header.get("resultMsg", "")} ({path})')
        return None

    return data


def extract_items(data):
    """API 응답 덩어리에서 실제 목록(items)만 꺼내옵니다."""
    body = (data.get('response', {}) or {}).get('body', {}) or {}
    items = body.get('items', [])
    if isinstance(items, dict):          # {'item': [...]} 형태인 경우
        items = items.get('item', [])
    if isinstance(items, dict):          # 결과가 1건이면 리스트가 아닐 수 있음
        items = [items]
    if not isinstance(items, list):
        items = []
    total = body.get('totalCount', len(items))
    try:
        total = int(total)
    except Exception:
        total = len(items)
    return items, total


def fetch_all(path, start_dt, end_dt):
    """
    한 종류(예: 본공고-물품)의 데이터를 기간 내 전부 가져옵니다.
    페이지를 넘겨가며(1페이지, 2페이지 ...) 끝까지 수집합니다.
    """
    collected = []
    page_size = PAGE_SIZE

    for page_no in range(1, MAX_PAGES + 1):
        params = {
            'serviceKey': SERVICE_KEY,
            'pageNo': page_no,
            'inqryDiv': '1',                        # 1 = 등록/공고 일시 기준 조회
            'inqryBgnDt': fmt_api_datetime(start_dt),
            'inqryEndDt': fmt_api_datetime(end_dt),
        }

        data = call_api(path, params, page_size)

        # 한 번에 500건 요청이 막히는 서버도 있어, 실패 시 100건으로 재시도합니다.
        if data is None and page_size != 100:
            page_size = 100
            data = call_api(path, params, page_size)

        if data is None:
            break

        items, total = extract_items(data)
        collected.extend(items)

        # 더 가져올 게 없으면 종료
        if len(items) < page_size or len(collected) >= total:
            break

        time.sleep(SLEEP_BETWEEN_CALLS)

    return collected


def fetch_with_fallback(stage, biz_type, start_dt, end_dt):
    """
    한 단계·업무구분에 대해 주소 후보를 차례로 시도해서
    데이터가 나오는 주소를 자동으로 찾아 사용합니다.
    """
    paths = API_ENDPOINTS.get(stage, {}).get(biz_type, [])
    for path in paths:
        items = fetch_all(path, start_dt, end_dt)
        if items:
            log(f'   ✅ {stage}-{biz_type}: {len(items)}건 수신 ({path.split("/")[-1]})')
            return items
    log(f'   · {stage}-{biz_type}: 수신 0건 (해당 기간 데이터 없음 또는 미제공 서비스)')
    return []


# -----------------------------------------------------------------------------
# [핵심 3] 키워드 걸러내기
# -----------------------------------------------------------------------------
def normalize(text):
    """비교하기 좋게 글자를 정리합니다. (공백/기호 제거 + 소문자화)"""
    if text is None:
        return ''
    return re.sub(r'[\s\-_/()\[\]·,.]', '', str(text)).lower()


NORMALIZED_KEYWORDS = [(kw, normalize(kw)) for kw in KEYWORDS]
NORMALIZED_EXCLUDES = [normalize(kw) for kw in EXCLUDE_KEYWORDS if kw.strip()]


def match_keywords(text):
    """
    주어진 글자 안에 내 키워드가 있는지 확인합니다.
    - 걸린 키워드 목록을 돌려주고, 하나도 없으면 빈 목록을 돌려줍니다.
    - 제외 키워드가 걸리면 무조건 빈 목록(=수집 안 함)입니다.
    """
    haystack = normalize(text)
    if not haystack:
        return []

    for exclude in NORMALIZED_EXCLUDES:
        if exclude and exclude in haystack:
            return []

    hits = [original for original, norm in NORMALIZED_KEYWORDS
            if norm and norm in haystack]
    return hits


# -----------------------------------------------------------------------------
# [핵심 4] API 응답을 엑셀용 표 형태로 정리
# -----------------------------------------------------------------------------
def pick(item, candidate_keys, default=''):
    """
    API 응답의 항목 이름이 서비스마다 조금씩 달라서,
    가능한 이름들을 순서대로 찾아보고 값이 있는 것을 씁니다.
    """
    for key in candidate_keys:
        value = item.get(key)
        if value not in (None, '', ' '):
            return str(value).strip()
    return default


def fmt_datetime_text(text):
    """20260810 1000 / 202608101000 같은 값을 '2026-08-10 10:00' 형태로 보기 좋게."""
    digits = re.sub(r'\D', '', str(text or ''))
    if len(digits) >= 12:
        return f'{digits[0:4]}-{digits[4:6]}-{digits[6:8]} {digits[8:10]}:{digits[10:12]}'
    if len(digits) >= 8:
        return f'{digits[0:4]}-{digits[4:6]}-{digits[6:8]}'
    return str(text or '').strip()


def to_amount(text):
    """'1,200,000원' 같은 값을 숫자 1200000 으로 바꿔줍니다. 못 바꾸면 빈칸."""
    digits = re.sub(r'[^\d]', '', str(text or ''))
    if not digits:
        return ''
    try:
        return int(digits)
    except Exception:
        return ''


def build_summary(item, stage):
    """
    '사업내용 요약' 열을 만듭니다.
    나라장터 API는 공고 본문 전체를 주지 않기 때문에,
    API가 제공하는 항목들(품명/규격/계약방법/낙찰방법 등)을 모아 요약으로 씁니다.
    """
    parts = []

    def add(label, keys):
        value = pick(item, keys)
        if value:
            parts.append(f'{label}: {value}')

    add('품명', ['prdctClsfcNoNm', 'prdctIdntNoNm', 'prdctNm', 'bsnsNm'])
    add('규격', ['prdctSpecNm', 'prdctSpec', 'specNm'])
    add('수량', ['prdctQty', 'qty'])
    add('계약방법', ['cntrctCnclsMthdNm', 'cntrctMthdNm', 'cntrctCnclsSttusNm'])
    add('낙찰방법', ['sucsfbidMthdNm', 'sucsfbidLwltRate'])
    add('참가자격', ['bidprcPsblIndstrytyNm', 'prtcptPsblRgnNm', 'indstrytyNm'])
    add('발주예정시기', ['orderPlanDt', 'orderPlanYm', 'cntrctMth', 'ordrPlanMt'])
    add('비고', ['ntceKindNm', 'rgstTyNm', 'bidNtceDtlsNm', 'bsnsDivNm'])

    if stage == '사전규격':
        add('의견등록마감', ['opninRgstClseDt'])

    return ' | '.join(parts)


def build_link(item, stage):
    """공고 상세 링크를 만듭니다. API가 링크를 주면 그대로 쓰고, 없으면 빈칸."""
    url = pick(item, [
        'bidNtceDtlUrl',        # 본공고 상세 URL
        'ntceSpecDocUrl1',      # 규격서 파일 URL
        'specDocUrl',
        'dtlUrl',
        'srvceDtlUrl',
    ])
    return url


def parse_item(item, stage, biz_type):
    """API 응답 한 건(item)을 엑셀 한 줄(dict)로 바꿉니다."""
    title = pick(item, [
        'bidNtceNm',            # 본공고: 공고명
        'prdctClsfcNoNm',       # 사전규격: 품명
        'bsnsNm',               # 발주계획: 사업명
        'ntceNm', 'prdctNm', 'sbmsnRcptNm',
    ])
    summary = build_summary(item, stage)

    # 공고명 + 사업내용요약 을 대상으로 키워드를 찾습니다.
    # (기관명은 일부러 제외합니다. '○○생명공학연구원' 같은 이름 때문에
    #  그 기관의 모든 공고가 딸려오는 것을 막기 위해서입니다.)
    org = pick(item, ['dminsttNm', 'rlDminsttNm', 'demandInsttNm', 'orderInsttNm'])
    hits = match_keywords(f'{title} {summary}')
    if not hits:
        return None     # 키워드에 안 걸리면 버립니다.

    row = {
        '구분': stage,
        '업무': biz_type,
        '공고명': title,
        '수요기관': org,
        '공고기관': pick(item, ['ntceInsttNm', 'orderInsttNm', 'insttNm']),
        '등록/공고일시': fmt_datetime_text(pick(item, [
            'bidNtceDt', 'rgstDt', 'rcptDt', 'bfSpecRgstDt', 'orderPlanRgstDt',
        ])),
        '사업기한(마감일)': fmt_datetime_text(pick(item, [
            'bidClseDt',            # 본공고 입찰마감일시
            'opninRgstClseDt',      # 사전규격 의견등록 마감
            'opengDt',              # 개찰일시
            'bidBeginDt',
            'orderPlanDt', 'ordrPlanMt',
        ])),
        '배정예산(원)': to_amount(pick(item, [
            'asignBdgtAmt',         # 배정예산액
            'presmptPrce',          # 추정가격
            'bdgtAmt', 'sumOfBdgt', 'orderPlanAmt', 'budgetAmount',
        ])),
        '사업내용요약': summary,
        '매칭키워드': ', '.join(hits),
        '공고번호': pick(item, [
            'bidNtceNo', 'bfSpecRgstNo', 'specRgstNo', 'orderPlanNo', 'refNo',
        ]),
        '공고링크': build_link(item, stage),
    }
    return row


# -----------------------------------------------------------------------------
# [핵심 5] 엑셀 파일 만들기
# -----------------------------------------------------------------------------
def make_excel(df, file_path, start_dt, end_dt):
    """정리된 표를 보기 좋은 엑셀 파일로 저장합니다."""
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='입찰정보', startrow=1)

        workbook = writer.book
        sheet = writer.sheets['입찰정보']

        # --- 맨 윗줄: 수집 기간 안내 ---
        sheet.cell(row=1, column=1).value = (
            f'수집기간: {start_dt.strftime("%Y-%m-%d %H:%M")} ~ '
            f'{end_dt.strftime("%Y-%m-%d %H:%M")}  |  총 {len(df)}건'
        )
        sheet.cell(row=1, column=1).font = Font(bold=True, size=11)

        # --- 머리글(2행) 꾸미기 ---
        header_fill = PatternFill('solid', fgColor='1F4E79')
        header_font = Font(bold=True, color='FFFFFF', size=10)
        thin = Side(style='thin', color='BFBFBF')
        border = Border(left=thin, right=thin, top=thin, bottom=thin)

        for col_idx in range(1, len(df.columns) + 1):
            cell = sheet.cell(row=2, column=col_idx)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.border = border

        # --- 열 너비 지정 ---
        widths = {
            '구분': 10, '업무': 8, '공고명': 55, '수요기관': 22, '공고기관': 22,
            '등록/공고일시': 18, '사업기한(마감일)': 18, '배정예산(원)': 16,
            '사업내용요약': 60, '매칭키워드': 20, '공고번호': 18, '공고링크': 35,
        }
        for idx, column_name in enumerate(df.columns, start=1):
            sheet.column_dimensions[get_column_letter(idx)].width = widths.get(column_name, 18)

        # --- 본문 셀 정렬 / 금액 서식 / 링크 처리 ---
        columns = list(df.columns)
        amount_col = columns.index('배정예산(원)') + 1 if '배정예산(원)' in columns else None
        link_col = columns.index('공고링크') + 1 if '공고링크' in columns else None
        stage_col = columns.index('구분') + 1 if '구분' in columns else None

        stage_colors = {
            '사업계획': 'FFF2CC',   # 연노랑
            '사전규격': 'DDEBF7',   # 연파랑
            '본공고': 'E2EFDA',     # 연초록
        }

        for row_idx in range(3, len(df) + 3):
            for col_idx in range(1, len(columns) + 1):
                cell = sheet.cell(row=row_idx, column=col_idx)
                cell.alignment = Alignment(vertical='top', wrap_text=True)
                cell.border = border
                cell.font = Font(size=10)

            if amount_col:
                sheet.cell(row=row_idx, column=amount_col).number_format = '#,##0'
                sheet.cell(row=row_idx, column=amount_col).alignment = Alignment(
                    horizontal='right', vertical='top')

            if stage_col:
                stage_cell = sheet.cell(row=row_idx, column=stage_col)
                color = stage_colors.get(str(stage_cell.value), None)
                if color:
                    stage_cell.fill = PatternFill('solid', fgColor=color)
                stage_cell.alignment = Alignment(horizontal='center', vertical='top')

            if link_col:
                link_cell = sheet.cell(row=row_idx, column=link_col)
                url = str(link_cell.value or '').strip()
                if url.startswith('http'):
                    link_cell.hyperlink = url
                    link_cell.value = '공고 바로가기'
                    link_cell.font = Font(color='0563C1', underline='single', size=10)

        # --- 필터 + 틀 고정 ---
        last_col_letter = get_column_letter(len(columns))
        sheet.auto_filter.ref = f'A2:{last_col_letter}{len(df) + 2}'
        sheet.freeze_panes = 'C3'

    return file_path


# -----------------------------------------------------------------------------
# [핵심 6] 이메일 발송
# -----------------------------------------------------------------------------
def build_mail_body(df, start_dt, end_dt):
    """메일 본문(HTML)을 만듭니다. 엑셀을 열지 않아도 대충 볼 수 있게 해줍니다."""
    period_text = (f'{start_dt.strftime("%Y-%m-%d %H:%M")} ~ '
                   f'{end_dt.strftime("%Y-%m-%d %H:%M")}')

    if df.empty:
        return f"""
        <div style="font-family:'맑은 고딕',sans-serif;font-size:14px;">
          <p>안녕하세요. 나라장터 입찰정보 자동수집 결과입니다.</p>
          <p><b>수집기간:</b> {period_text}</p>
          <p style="color:#c00;"><b>해당 기간에 키워드와 일치하는 신규 공고가 없습니다.</b></p>
        </div>
        """

    # 단계별 건수 요약
    counts = df['구분'].value_counts().to_dict()
    summary_line = ' / '.join(f'{k} {v}건' for k, v in counts.items())

    rows_html = ''
    for _, row in df.head(30).iterrows():
        budget = row['배정예산(원)']
        budget_text = f'{budget:,}' if isinstance(budget, int) else '-'
        link = str(row['공고링크'] or '')
        link_html = f'<a href="{link}">보기</a>' if link.startswith('http') else '-'
        rows_html += f"""
          <tr>
            <td style="border:1px solid #ddd;padding:6px;white-space:nowrap;">{row['구분']}</td>
            <td style="border:1px solid #ddd;padding:6px;">{row['공고명']}</td>
            <td style="border:1px solid #ddd;padding:6px;white-space:nowrap;">{row['사업기한(마감일)'] or '-'}</td>
            <td style="border:1px solid #ddd;padding:6px;text-align:right;white-space:nowrap;">{budget_text}</td>
            <td style="border:1px solid #ddd;padding:6px;text-align:center;">{link_html}</td>
          </tr>"""

    more_note = ('<p style="color:#666;">※ 상위 30건만 표시했습니다. '
                 '전체 내용은 첨부된 엑셀 파일을 확인해 주세요.</p>'
                 if len(df) > 30 else '')

    return f"""
    <div style="font-family:'맑은 고딕',sans-serif;font-size:14px;">
      <p>안녕하세요. 나라장터 입찰정보 자동수집 결과입니다.</p>
      <p><b>수집기간:</b> {period_text}<br>
         <b>총 건수:</b> {len(df)}건 ({summary_line})</p>
      <table style="border-collapse:collapse;font-size:13px;">
        <thead>
          <tr style="background:#1F4E79;color:#fff;">
            <th style="border:1px solid #ddd;padding:6px;">구분</th>
            <th style="border:1px solid #ddd;padding:6px;">공고명</th>
            <th style="border:1px solid #ddd;padding:6px;">사업기한</th>
            <th style="border:1px solid #ddd;padding:6px;">배정예산(원)</th>
            <th style="border:1px solid #ddd;padding:6px;">링크</th>
          </tr>
        </thead>
        <tbody>{rows_html}</tbody>
      </table>
      {more_note}
      <p style="color:#888;font-size:12px;margin-top:20px;">
        본 메일은 GitHub Actions 자동화 스크립트가 발송했습니다.
      </p>
    </div>
    """


def send_email(subject, html_body, attachment_path=None):
    """Gmail SMTP 를 통해 메일을 보냅니다."""
    sender = get_env('GMAIL_USER')                 # 보내는 사람 (내 Gmail 주소)
    password = get_env('GMAIL_APP_PASSWORD')       # Gmail '앱 비밀번호' 16자리
    receivers = get_env('MAIL_TO')                 # 받는 사람 (쉼표로 여러 명 가능)

    receiver_list = [addr.strip() for addr in receivers.replace(';', ',').split(',')
                     if addr.strip()]

    message = EmailMessage()
    message['Subject'] = subject
    message['From'] = sender
    message['To'] = ', '.join(receiver_list)
    message.set_content('HTML 메일입니다. HTML을 지원하는 메일앱에서 확인해 주세요.')
    message.add_alternative(html_body, subtype='html')

    # 엑셀 파일 첨부
    if attachment_path and os.path.exists(attachment_path):
        with open(attachment_path, 'rb') as file:
            message.add_attachment(
                file.read(),
                maintype='application',
                subtype='vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                filename=os.path.basename(attachment_path),
            )

    context = ssl.create_default_context()
    with smtplib.SMTP_SSL('smtp.gmail.com', 465, context=context) as server:
        # Gmail 앱 비밀번호는 공백 없이 16자리입니다. 혹시 공백이 있으면 제거.
        server.login(sender, password.replace(' ', ''))
        server.send_message(message)

    log(f'📧 메일 발송 완료 → {", ".join(receiver_list)}')


# -----------------------------------------------------------------------------
# [메인] 실제 실행 순서
# -----------------------------------------------------------------------------
def main():
    global SERVICE_KEY

    log('=' * 60)
    log('나라장터 입찰정보 자동 수집을 시작합니다.')
    log('=' * 60)

    # (0) 공공데이터포털 인증키 읽기
    SERVICE_KEY = clean_service_key(get_env('SERVICE_KEY'))

    # (1) 수집 기간 계산
    lookback_days = 0
    try:
        lookback_days = int(os.environ.get('LOOKBACK_DAYS', '0') or 0)
    except ValueError:
        lookback_days = 0

    start_dt, end_dt = calc_period(lookback_days=lookback_days)
    log(f'🗓  수집기간: {start_dt.strftime("%Y-%m-%d(%a) %H:%M")} '
        f'~ {end_dt.strftime("%Y-%m-%d(%a) %H:%M")}')
    if end_dt.weekday() == 0 and not lookback_days:
        log('   (월요일이므로 금·토·일 공고까지 함께 수집합니다)')

    # (2) 단계별로 데이터 수집
    rows = []
    raw_total = 0

    for stage, enabled in COLLECT_STAGES.items():
        if not enabled:
            log(f'⏭  {stage}: 설정에서 꺼져 있어 건너뜁니다.')
            continue

        log(f'🔍 [{stage}] 수집 중...')
        for biz_type in BUSINESS_TYPES:
            try:
                items = fetch_with_fallback(stage, biz_type, start_dt, end_dt)
            except Exception as error:
                log(f'   ⚠️ {stage}-{biz_type} 수집 중 오류: {error}')
                items = []

            raw_total += len(items)
            for item in items:
                if not isinstance(item, dict):
                    continue
                row = parse_item(item, stage, biz_type)
                if row:
                    rows.append(row)

    log(f'📊 전체 수신 {raw_total}건 → 키워드 일치 {len(rows)}건')

    # (3) 표로 정리 + 중복 제거 + 정렬
    if rows:
        df = pd.DataFrame(rows, columns=COLUMNS)
        df = df.drop_duplicates(subset=['구분', '공고번호', '공고명'], keep='first')
        # 진행 단계 순서(사업계획 → 사전규격 → 본공고)로 정렬하고,
        # 같은 단계 안에서는 최신 공고가 위로 오게 합니다.
        stage_order = {'사업계획': 0, '사전규격': 1, '본공고': 2}
        df['_순서'] = df['구분'].map(stage_order).fillna(9)
        df = df.sort_values(
            by=['_순서', '등록/공고일시'], ascending=[True, False]
        ).drop(columns=['_순서']).reset_index(drop=True)
    else:
        df = pd.DataFrame(columns=COLUMNS)

    # (4) 엑셀 만들기
    file_name = f'입찰정보_{end_dt.strftime("%Y%m%d_%H%M")}.xlsx'
    attachment = None
    if not df.empty:
        make_excel(df, file_name, start_dt, end_dt)
        attachment = file_name
        log(f'📁 엑셀 생성 완료: {file_name}')
    else:
        log('📁 수집 결과가 없어 엑셀은 만들지 않습니다.')

    # (5) 메일 발송
    subject = (f'{MAIL_SUBJECT_PREFIX} {end_dt.strftime("%m월 %d일")} '
               f'신규 {len(df)}건')
    body = build_mail_body(df, start_dt, end_dt)
    send_email(subject, body, attachment)

    log('✅ 모든 작업이 정상적으로 끝났습니다.')


# 이 파일을 직접 실행했을 때만 main() 이 돌아갑니다.
if __name__ == '__main__':
    try:
        main()
    except SystemExit:
        raise
    except Exception:
        # 예상 못 한 오류가 나면 원인을 로그에 남기고, 오류 알림 메일을 시도합니다.
        log('❌ 예상치 못한 오류가 발생했습니다.')
        traceback.print_exc()
        try:
            send_email(
                f'{MAIL_SUBJECT_PREFIX} ⚠️ 수집 실패 알림',
                f'<pre style="font-size:12px;">{traceback.format_exc()}</pre>',
                None,
            )
        except Exception:
            pass
        sys.exit(1)
