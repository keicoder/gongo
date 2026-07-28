#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
전담기관 사업공고 크롤러 (다중 사이트 지원)

지원 사이트:
    - nipa : 정보통신산업진흥원  https://www.nipa.kr/home/2-2?curPage=1
             (일반 HTML 테이블, requests + BeautifulSoup로 수집)
    - iris : 범부처통합연구지원시스템(IRIS) https://www.iris.go.kr/contents/retrieveBsnsAncmBtinSituListView.do
             (목록은 requests로 수집 가능 확인됨. 상세링크는 onclick 기반 JS라
              정확한 URL 패턴은 로컬 확인 필요, 페이지네이션 파라미터도 미검증)
    - nia  : 한국지능정보사회진흥원(NIA) 입찰공고  https://nia.or.kr/site/nia_kor/ex/bbs/List.do?cbIdx=78336
             (목록/페이지네이션은 pageIndex 파라미터의 실제 URL이라 requests로 수집 가능.
              단, 상세링크는 onclick 기반 JS라 정확한 URL 패턴은 로컬 확인 필요)
    - keit : 한국산업기술기획평가원(KEIT) S-Rome 과제공고
             https://srome.keit.re.kr/srome/biz/perform/opnnPrpsl/retrieveTaskAnncmListView.do
             (목록/페이지네이션은 pageIndex 파라미터의 실제 URL이라 requests로 수집 가능.
              단, 상세링크는 onclick 기반 JS라 정확한 URL 패턴은 로컬 확인 필요)
    - kiat : 한국산업기술진흥원(KIAT) 입찰공고
             https://www.kiat.or.kr/front/board/boardContentsListPage.do?board_id=77
             (목록은 POST boardContentsListAjax.do 로 채워지는 AJAX 방식.
              requests.post로 동일 파라미터를 흉내내어 수집. 상세링크의 정확한
              URL 패턴(boardContentsView.do)은 로컬 확인 필요)
    - kisa : 한국인터넷진흥원(KISA) 입찰공고  https://www.kisa.or.kr/403
             (일반 HTML 테이블 + 실제 상세링크가 그대로 노출되어 requests만으로 수집 가능.
              페이지네이션은 ?page= 파라미터로 추정, 로컬 확인 권장)
    - smtech : 중소기업기술개발사업 종합관리시스템(SMTECH) 사업공고
             https://www.smtech.go.kr/front/ifg/no/notice02_list.do
             (일반 HTML 테이블, requests로 수집 가능. 단, 목록에 SMTECH 자체 공고와
              IRIS 연동 공고가 섞여있고 IRIS 연동 건은 상세링크가 javascript:goMove()라
              실제 URL을 못 얻음 - 작성자 컬럼에 "SMTECH"/"IRIS" 구분 표시)

기능:
    1) 사이트별로 공고 목록 수집 (NIPA: curPage 순회 / IRIS: 페이지 번호 클릭)
    2) 기존에 쌓아둔 DB 파일(csv/xlsx)과 비교해서 새로 올라온 공고만 추출
    3) DB 파일에 신규 공고를 누적 저장 (다음 실행 때 비교 기준이 됨)
    모든 사이트가 동일한 컬럼(기관명 ~ 상세링크)에 맞춰 저장되므로,
    crawling_db.csv 하나에 여러 기관 공고를 함께 누적할 수 있습니다.
    --db를 따로 지정하지 않으면, 실행 위치와 상관없이 항상 crawling.py와
    같은 폴더의 crawling_db.csv에 저장/누적됩니다.
    --site를 지정하지 않으면 기본값 all(=nipa+nia+keit)로 세 기관을 한 번에 수집합니다.

사용법:
    # NIPA: 최초 1회 전체 수집
    python crawling.py --site nipa --start 1 --end 37 --db crawling_db.csv

    # NIPA: 매일 신규 공고만 확인
    python crawling.py --site nipa --start 1 --end 3 --db crawling_db.csv --new-out new_today.csv

    # IRIS: 최초 1회 전체 수집
    python crawling.py --site iris --start 1 --end 3 --db crawling_db.csv

    # NIA: 최초 1회 전체 수집 (전체 895페이지 중 원하는 범위만)
    python crawling.py --site nia --start 1 --end 5 --db crawling_db.csv

    # KEIT: 최초 1회 전체 수집 (전체 72페이지 중 원하는 범위만)
    python crawling.py --site keit --start 1 --end 5 --db crawling_db.csv

    # KIAT: 최초 1회 전체 수집
    python crawling.py --site kiat --start 1 --end 5 --db crawling_db.csv

    # KISA: 최초 1회 전체 수집
    python crawling.py --site kisa --start 1 --end 5 --db crawling_db.csv

    # SMTECH: 게시물이 많아 페이지당 30건씩 가져오며 수집
    python crawling.py --site smtech --start 1 --end 5 --db crawling_db.csv

    # 한 번에 여러 기관: NIPA+NIA+KEIT를 순서대로 수집해서 같은 DB에 누적
    python crawling.py --site all --start 1 --end 3

주의:
    - 사이트 구조가 바뀌면 각 사이트의 parse 함수 안 셀렉터를 수정해야 합니다.
    - IRIS는 이 환경에서 실제 DOM(class/id)을 확인하지 못한 상태로 작성했습니다.
      최초 실행 결과가 비거나 이상하면 아래 "IRIS 셀렉터 확인" 안내를 참고해 로컬에서 조정하세요.
    - 서버 부담을 줄이기 위해 요청 사이에 딜레이를 둡니다.
"""

import argparse
import datetime as _dt
import os
import random
import re
import sys
import time
from dataclasses import dataclass, asdict, fields, replace

import requests
from bs4 import BeautifulSoup

import json
import gspread
import pandas as pd
from google.oauth2.service_account import Credentials

NIPA_BASE_URL = "https://www.nipa.kr/home/2-2"
IRIS_URL = "https://www.iris.go.kr/contents/retrieveBsnsAncmBtinSituListView.do"
NIA_BASE_URL = "https://nia.or.kr/site/nia_kor/ex/bbs/List.do"
NIA_CBIDX = "78336"  # 입찰공고 게시판 ID (다른 게시판을 원하면 cbIdx만 바꾸면 됨)
KEIT_BASE_URL = "https://srome.keit.re.kr/srome/biz/perform/opnnPrpsl/retrieveTaskAnncmListView.do"
KEIT_PRGM_ID = "XPG201040000"
KIAT_LIST_AJAX_URL = "https://www.kiat.or.kr/front/board/boardContentsListAjax.do"
KIAT_VIEW_URL = "https://www.kiat.or.kr/front/board/boardContentsView.do"
KIAT_BOARD_ID = "77"
KIAT_MENU_ID = "1e29209309434ec29095728c6f1356c7"
KISA_URL = "https://www.kisa.or.kr/403"
SMTECH_LIST_URL = "https://www.smtech.go.kr/front/ifg/no/notice02_list.do"
SMTECH_ROWS_PER_PAGE = 30  # 게시 건수가 많아 페이지당 넉넉히 가져와 요청 횟수를 줄임

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    )
}

# 신규 공고 판별에 사용할 고유 키 컬럼 (상세링크가 가장 안전함)
UNIQUE_KEY = "상세링크"

_SESSION_ID_RE = re.compile(r";jsessionid=[^?]*", re.IGNORECASE)


def _strip_session_id(url: str) -> str:
    """URL을 정규화해 같은 공고가 매 실행마다 다른 링크로 취급되지 않게 한다.
    - ;jsessionid=... : 요청마다 바뀌는 세션ID 제거
    - pageIndex=N : 몇 번째 목록 페이지에서 봤는지는 공고 자체와 무관하므로 제거"""
    url = _SESSION_ID_RE.sub("", url)
    url = re.sub(r"([&?])pageIndex=\d+&?", r"\1", url)
    url = re.sub(r"[?&]$", "", url)
    return url

# crawling.py가 있는 폴더를 기준으로 기본 DB 파일 경로를 잡는다.
# (어느 위치에서 실행하든 --db를 따로 지정하지 않으면 항상 이 스크립트와 같은 폴더에 저장됨)
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
DEFAULT_DB_PATH = os.path.join(SCRIPT_DIR, "crawling_db.csv")


@dataclass
class Announcement:
    공고일자: str = ""  # 가장 앞쪽 열
    기관명: str = ""
    # 아래는 크롤링 시점에 제목/비고와 사이트별 접수기간 원문을 분석해 자동으로
    # 채우는 파생 컬럼(enrich_announcement 및 각 사이트 파서 참고). 기관명 바로 뒤에 배치.
    AI관련여부: str = ""
    품질인증관련여부: str = ""
    관련키워드: str = ""
    공고유형: str = ""
    접수시작일: str = ""
    접수종료일: str = ""
    마감Dday: str = ""
    번호: str = ""
    남은기간: str = ""
    제목: str = ""
    비고: str = ""
    작성자: str = ""
    상세링크: str = ""


# ---------------------------------------------------------------------------
# AI품질역량센터 관점 자동 분류 (제목/비고 기반 규칙)
# ---------------------------------------------------------------------------
# 크롤링 직후 한 번만 계산해서 Announcement에 채워 넣는다. 신규 공고 판별/DB
# 누적 로직(find_new_items, update_db)은 상세링크 기준으로 동작하므로, 이미
# DB에 있는 공고는 재계산 없이 그대로 유지되고 새로 발견된 공고만 새로 분류된다.
# 규칙(키워드 사전 등)은 실제 데이터를 보며 계속 다듬어야 정확도가 올라간다.

AI_KEYWORDS = (
    "AI", "인공지능", "데이터", "지능형", "생성형", "생성AI", "머신러닝", "딥러닝",
    "빅데이터", "챗봇", "LLM", "sLM", "피지컬AI", "에이전트", "에이전틱",
    "초거대", "파운데이션모델", "온디바이스AI", "멀티모달",
    "자율주행", "로봇", "로보틱스", "컴퓨터비전", "자연어처리", "NLP", "강화학습", "신경망",
)
QUALITY_KEYWORDS = (
    "품질", "인증", "검증", "검인증", "시험", "신뢰성", "평가", "표준",
    "벤치마크", "적합성평가", "안전성", "성능평가", "실증",
    "국제표준", "규격", "ISO", "ISO/IEC", "보안", "사이버보안", "취약점",
)

_DATE_IN_TEXT_RE = re.compile(r"(\d{4})[./\-](\d{1,2})[./\-](\d{1,2})")

# 공고문 본문처럼 자유 서술형 텍스트에서는 '2024. 4. 9.'처럼 날짜 사이에
# 공백이 섞여 나오는 경우가 많아, 이를 '2024.4.9'로 압축해 _DATE_IN_TEXT_RE로
# 다시 인식할 수 있게 해주는 보조 정규식.
_LOOSE_DATE_RE = re.compile(r"(\d{4})\s*\.\s*(\d{1,2})\s*\.\s*(\d{1,2})\s*\.?")

# '2026년 7월 22일'처럼 한글로 된 날짜, 그리고 범위 표기에서 뒤쪽 날짜의 연도가
# 생략된 '7월 26일'(앞서 나온 연도를 그대로 사용) 형식도 함께 처리한다.
_KOREAN_DATE_FULL_RE = re.compile(r"(\d{4})년\s*(\d{1,2})월\s*(\d{1,2})일")
_KOREAN_DATE_BARE_RE = re.compile(r"(?<!\d)(\d{1,2})월\s*(\d{1,2})일")


def _compact_loose_dates(text: str) -> str:
    text = _LOOSE_DATE_RE.sub(lambda m: f"{m.group(1)}.{m.group(2)}.{m.group(3)}", text)

    last_year = [None]

    def _full_korean(m):
        last_year[0] = m.group(1)
        return f"{m.group(1)}.{m.group(2)}.{m.group(3)}"

    text = _KOREAN_DATE_FULL_RE.sub(_full_korean, text)

    def _bare_korean(m):
        if last_year[0]:
            return f"{last_year[0]}.{m.group(1)}.{m.group(2)}"
        return m.group(0)  # 연도를 알 수 없으면 그대로 둠

    text = _KOREAN_DATE_BARE_RE.sub(_bare_korean, text)

    return text


def _keyword_window(body_text: str, idx: int, max_len: int = 80) -> str:
    """idx 위치부터 다음 줄바꿈 전까지(또는 max_len 중 짧은 쪽)만 잘라온다.
    고정 길이로만 자르면 다음 항목(예: '납품기한')의 날짜까지 같이 딸려 들어와
    엉뚱하게 하나의 기간(범위)으로 묶이는 문제가 있어, 같은 줄 안으로 제한한다."""
    newline_pos = body_text.find("\n", idx)
    end = idx + max_len
    if newline_pos != -1:
        end = min(end, newline_pos)
    return body_text[idx:end]


def _match_keywords(text: str, keywords: tuple) -> list:
    return [k for k in keywords if k in text]


def _classify_공고유형(title: str) -> str:
    if any(k in title for k in ("입찰", "용역", "구매")):
        return "입찰(용역)공고"
    if any(k in title for k in ("채용", "인재영입")):
        return "인력모집"
    if any(k in title for k in ("포상", "시상")):
        return "포상·시상"
    if any(k in title for k in ("설명회", "세미나", "컨퍼런스", "간담회", "박람회")):
        return "설명회·행사"
    if any(k in title for k in ("공고", "모집", "지원", "선정")):
        return "R&D지원사업"
    return "기타"


def _extract_dates(period_text: str):
    """신청기간 텍스트(사이트마다 형식이 달라 YYYY-MM-DD/YYYY.MM.DD 등을
    느슨하게 매칭)에서 시작일/종료일을 뽑는다.

    - 날짜가 2개면: 앞을 시작일, 뒤를 종료일로 사용
    - 날짜가 1개뿐이면: 그 날짜 앞뒤 기호/문구로 시작일인지 종료일인지 추정
      ('~2026.08.15', '2026.08.15까지', '마감 2026.08.15' 등은 종료일 /
       '2026.08.15~', '2026.08.15부터' 등은 시작일). 판단 근거가 없으면
      실무 관행상 더 흔한 '종료일(마감일)'로 간주한다.
    - 날짜가 없으면 둘 다 빈 문자열.
    """
    period_text = period_text or ""
    matches = list(_DATE_IN_TEXT_RE.finditer(period_text))

    parsed = []
    for m in matches:
        y, mo, d = m.groups()
        try:
            parsed.append((m, _dt.date(int(y), int(mo), int(d))))
        except ValueError:
            continue

    if len(parsed) >= 2:
        return parsed[0][1].isoformat(), parsed[1][1].isoformat()

    if len(parsed) == 1:
        m, date_val = parsed[0]
        before = period_text[:m.start()]
        after = period_text[m.end():].lstrip()
        # '2026.08.15~', '2026.08.15부터'처럼 날짜 뒤에 '~'/'부터'가 오면 시작일
        if (after.startswith("~") or after.startswith("부터")) and "~" not in before:
            return date_val.isoformat(), ""
        # 그 외(날짜 앞에 '~', '까지'/'마감' 문구, 혹은 아무 단서 없음)는
        # 실무에서 더 흔한 '종료일(마감일)'로 간주
        return "", date_val.isoformat()

    return "", ""


def _compute_dday(end_date_str: str) -> str:
    if not end_date_str:
        return ""
    try:
        end = _dt.date.fromisoformat(end_date_str)
    except ValueError:
        return ""
    diff = (end - _dt.date.today()).days
    if diff < 0:
        return "마감"
    if diff == 0:
        return "D-day"
    return f"D-{diff}"


_WRITTEN_DATE_RE = re.compile(r"^(\d{4})[.\-/](\d{1,2})[.\-/](\d{1,2})\.?$")


def _normalize_written_date(raw: str) -> str:
    """공고일자를 'YYYY-MM-DD' 형식으로 통일한다.
    '2026.07.23', '2026/07/23' 등은 '2026-07-23'로 변환하고,
    이미 'YYYY-MM-DD' 형식이거나 알 수 없는 형식/빈 값은 최대한 원본을 보존한다."""
    if not raw:
        return raw
    raw = str(raw).strip()
    m = _WRITTEN_DATE_RE.match(raw)
    if not m:
        return raw
    y, mo, d = m.groups()
    try:
        return f"{int(y):04d}-{int(mo):02d}-{int(d):02d}"
    except ValueError:
        return raw


def enrich_announcement(ann: Announcement) -> Announcement:
    """제목/비고를 분석해 AI품질역량센터 관점의 분류 컬럼을 채운 새 Announcement 반환.
    접수시작일/접수종료일은 각 사이트 크롤링 시점에 이미 채워져 들어오므로 여기서는
    마감Dday만 그 값을 바탕으로 계산하고, 접수시작일이 끝내 비어있으면(사이트에서
    못 찾은 경우) 공고일자로 대체한다(모든 사이트 공통 규칙)."""
    text = f"{ann.제목} {ann.비고}"
    # '사전규격공개'는 입찰 공고에서 쓰는 관용구(사업 규격을 미리 공개)로,
    # 품질/표준 인증과는 무관하므로 '규격' 키워드 매칭 대상에서만 제외한다.
    quality_match_text = text.replace("사전규격공개", "")

    ai_hits = _match_keywords(text, AI_KEYWORDS)
    quality_hits = _match_keywords(quality_match_text, QUALITY_KEYWORDS)

    접수시작일 = ann.접수시작일 or ann.공고일자

    return replace(
        ann,
        AI관련여부="Y" if ai_hits else "N",
        품질인증관련여부="Y" if quality_hits else "N",
        관련키워드=", ".join(ai_hits + quality_hits),
        공고유형=_classify_공고유형(ann.제목),
        접수시작일=접수시작일,
        마감Dday=_compute_dday(ann.접수종료일),
    )


def fetch_nipa_page(cur_page: int) -> str:
    """NIPA 지정한 curPage의 HTML을 가져온다."""
    params = {"curPage": cur_page}
    resp = requests.get(NIPA_BASE_URL, params=params, headers=HEADERS, timeout=15)
    resp.raise_for_status()
    resp.encoding = resp.apparent_encoding or "utf-8"
    return resp.text


def parse_nipa_row(tr, org: str) -> Announcement | None:
    """NIPA <tr> 한 행을 Announcement로 변환. 구조가 다르면 여기만 수정하면 됨."""
    tds = tr.find_all("td")
    if len(tds) < 5:
        # 헤더 행이거나 구조가 다른 행은 건너뜀
        return None

    번호 = tds[0].get_text(strip=True)
    남은기간 = tds[1].get_text(strip=True)

    title_cell = tds[2]
    a_tag = title_cell.find("a")
    if a_tag is None:
        return None

    제목 = a_tag.get_text(strip=True)
    href = a_tag.get("href", "")
    상세링크 = href if href.startswith("http") else f"https://www.nipa.kr{href}"

    # title_cell 안의 전체 텍스트에서 제목/신청기간을 제외한 나머지를 비고로 추정
    full_text = title_cell.get_text(separator="\n", strip=True)
    lines = [ln.strip() for ln in full_text.split("\n") if ln.strip()]

    비고 = ""
    신청기간_원문 = ""
    for ln in lines:
        if ln == 제목:
            continue
        m = re.search(r"신청기간\s*:?\s*(.+)", ln)
        if m:
            신청기간_원문 = m.group(1).strip()
        elif not 비고:
            비고 = ln

    접수시작일, 접수종료일 = _extract_dates(신청기간_원문)

    작성자 = tds[3].get_text(strip=True)
    공고일자 = tds[4].get_text(strip=True)

    return Announcement(
        기관명=org,
        번호=번호,
        남은기간=남은기간,
        제목=제목,
        비고=비고,
        접수시작일=접수시작일,
        접수종료일=접수종료일,
        작성자=작성자,
        공고일자=공고일자,
        상세링크=상세링크,
    )


def parse_nipa_list(html: str, org: str) -> list[Announcement]:
    soup = BeautifulSoup(html, "html.parser")
    table = soup.find("table")
    if table is None:
        return []

    rows = table.find_all("tr")
    results = []
    for tr in rows:
        item = parse_nipa_row(tr, org)
        if item is not None:
            results.append(item)
    return results


def crawl_nipa(start_page: int, end_page: int, org: str, delay_range=(0.5, 1.0)) -> list[Announcement]:
    all_items: list[Announcement] = []
    for page in range(start_page, end_page + 1):
        print(f"[NIPA {page}/{end_page}] 수집 중...", file=sys.stderr)
        try:
            html = fetch_nipa_page(page)
        except requests.RequestException as e:
            print(f"  -> 요청 실패: {e}", file=sys.stderr)
            continue

        items = parse_nipa_list(html, org)
        if not items:
            print("  -> 공고 없음 (마지막 페이지이거나 구조 변경 가능성)", file=sys.stderr)
        all_items.extend(items)

        time.sleep(random.uniform(*delay_range))

    return all_items


# ---------------------------------------------------------------------------
# NIA (한국지능정보사회진흥원) 입찰공고
# ---------------------------------------------------------------------------
# 목록/페이지네이션은 pageIndex 파라미터를 쓰는 실제 URL이라 requests로 수집
# 가능합니다. 각 공고의 목록 링크 자체는 href="#view" + onclick 방식의 자바
# 스크립트지만, 실제 상세페이지 URL 패턴은 다른 경로(검색엔진에 노출된 실제
# 공고들)를 통해 다음과 같이 확인했습니다:
#   https://www.nia.or.kr/site/nia_kor/ex/bbs/View.do?cbIdx=78336&bcIdx=<번호>&parentSeq=<번호>
# onclick 안의 게시글 번호(bcIdx)만 정확히 뽑아내면 상세페이지로 바로 이동 가능.
#
# 다만 상세페이지 내용은 공고 유형에 따라 완전히 다릅니다:
# - 나라장터(조달청) 연계형: "입찰개요" 표에 접수기간이 아예 없고 납품기한만 있음
# - 직접 작성형 공고문: "입찰일정" 표가 있지만 날짜가 '11.05(화)~11.19(화)'처럼
#   연도 없이 나오는 경우가 많아 정확한 날짜 변환이 어려움
# 따라서 KISA와 같은 방식(본문에서 키워드 주변 날짜 탐색)으로 최대한 잡되,
# 연도가 명시되지 않은 경우는 못 잡을 수 있다는 한계가 있음.

_NIA_DATE_RE = re.compile(r"^\d{4}\.\d{2}\.\d{2}$")
_NIA_VIEWCOUNT_RE = re.compile(r"조회수\s*(\d+)")
_NIA_ID_IN_ONCLICK_RE = re.compile(r"'(\d+)'")
NIA_VIEW_URL = "https://www.nia.or.kr/site/nia_kor/ex/bbs/View.do"


def fetch_nia_page(page_index: int) -> str:
    params = {"cbIdx": NIA_CBIDX, "pageIndex": page_index}
    resp = requests.get(NIA_BASE_URL, params=params, headers=HEADERS, timeout=15)
    resp.raise_for_status()
    resp.encoding = resp.apparent_encoding or "utf-8"
    return resp.text


def parse_nia_row(li, org: str = "NIA") -> Announcement | None:
    """<li> 한 항목을 Announcement로 변환. 실제 class명 확인 전이라 텍스트 패턴 기반."""
    text = li.get_text(separator="\n", strip=True)
    if not text or "조회수" not in text:
        return None

    lines = [ln.strip() for ln in text.split("\n") if ln.strip()]

    첨부파일 = any("첨부파일" in ln for ln in lines)
    신규 = any(ln.lower() == "new" for ln in lines)
    공고일자 = next((ln for ln in lines if _NIA_DATE_RE.match(ln)), "")

    조회수 = ""
    for ln in lines:
        m = _NIA_VIEWCOUNT_RE.match(ln)
        if m:
            조회수 = m.group(1)
            break

    # 제목/작성자/부서 후보: 첨부파일표시·new·날짜·조회수 라인을 제외한 나머지
    rest = [
        ln for ln in lines
        if "첨부파일" not in ln
        and ln.lower() != "new"
        and not _NIA_DATE_RE.match(ln)
        and not _NIA_VIEWCOUNT_RE.match(ln)
    ]
    if not rest:
        return None

    제목 = rest[0] + (" [첨부파일]" if 첨부파일 else "")
    작성자 = rest[1] if len(rest) > 1 else ""
    부서 = rest[2] if len(rest) > 2 else ""

    a_tag = li.find("a")
    href = a_tag.get("href", "") if a_tag else ""
    onclick = a_tag.get("onclick", "") if a_tag else ""

    onclick_ids = _NIA_ID_IN_ONCLICK_RE.findall(onclick) if onclick else []
    게시글번호 = onclick_ids[-1] if onclick_ids else ""

    if href and href.startswith("http"):
        상세링크 = href
    elif 게시글번호:
        상세링크 = f"{NIA_VIEW_URL}?cbIdx={NIA_CBIDX}&bcIdx={게시글번호}&parentSeq={게시글번호}"
    else:
        상세링크 = NIA_BASE_URL

    번호 = 게시글번호 or 조회수  # 게시글 고유번호 추정값

    return Announcement(
        기관명=org,
        번호=번호,
        남은기간="신규" if 신규 else "",
        제목=제목,
        비고=부서,
        작성자=작성자,
        공고일자=공고일자,
        상세링크=상세링크,
    )


def parse_nia_list(html: str, org: str = "NIA") -> list[Announcement]:
    soup = BeautifulSoup(html, "html.parser")
    # 목록 컨테이너의 정확한 class명을 모르므로 전체 <li> 중 "조회수"가 포함된
    # 항목만 공고로 간주 (메뉴/배너 등 다른 <li>는 자동으로 걸러짐)
    results = []
    for li in soup.find_all("li"):
        item = parse_nia_row(li, org)
        if item is not None:
            results.append(item)
    return results


def fetch_nia_detail(상세링크: str) -> str:
    resp = requests.get(상세링크, headers=HEADERS, timeout=15)
    resp.raise_for_status()
    resp.encoding = resp.apparent_encoding or "utf-8"
    return resp.text


def _extract_nia_period_from_detail(html: str) -> str:
    """상세페이지 본문에서 공고기간/접수기간/마감 관련 문구를 찾아 그 주변
    텍스트를 반환한다. 나라장터 연계형 공고는 이 정보 자체가 없고, 직접
    작성형 공고문은 날짜에 연도가 빠진 경우가 많아(예: '11.05(화)~11.19(화)')
    KISA보다도 인식률이 낮을 수 있음 - 실제 값과 비교 확인 권장. 못 찾으면
    빈 문자열."""
    soup = BeautifulSoup(html, "html.parser")
    body_text = _compact_loose_dates(soup.get_text("\n", strip=True))

    for keyword in ("접수기간", "공고기간", "제출기간", "공개기간", "마감일시", "입찰마감", "마감"):
        idx = body_text.find(keyword)
        if idx == -1:
            continue
        window = _keyword_window(body_text, idx)
        if _DATE_IN_TEXT_RE.search(window):
            return window
    return ""


def crawl_nia(start_page: int, end_page: int, org: str = "NIA", delay_range=(0.5, 1.0)) -> list[Announcement]:
    all_items: list[Announcement] = []
    for page in range(start_page, end_page + 1):
        print(f"[NIA {page}/{end_page}] 수집 중...", file=sys.stderr)
        try:
            html = fetch_nia_page(page)
        except requests.RequestException as e:
            print(f"  -> 요청 실패: {e}", file=sys.stderr)
            continue

        items = parse_nia_list(html, org)
        if not items:
            print("  -> 공고 없음 (마지막 페이지이거나 구조 변경 가능성)", file=sys.stderr)

        # 목록에 접수기간 정보가 없으므로, 상세페이지 본문에서 찾아 보완 시도
        for idx, item in enumerate(items):
            if item.접수시작일 or item.접수종료일 or not item.상세링크:
                continue
            try:
                detail_html = fetch_nia_detail(item.상세링크)
                period_text = _extract_nia_period_from_detail(detail_html)
                if period_text:
                    시작일, 종료일 = _extract_dates(period_text)
                    items[idx] = replace(item, 접수시작일=시작일, 접수종료일=종료일)
            except requests.RequestException as e:
                print(f"  -> 상세페이지 요청 실패({item.제목}): {e}", file=sys.stderr)
            time.sleep(random.uniform(*delay_range))

        all_items.extend(items)

        time.sleep(random.uniform(*delay_range))

    return all_items


# ---------------------------------------------------------------------------
# KEIT (한국산업기술기획평가원) S-Rome 과제공고
# ---------------------------------------------------------------------------
# 목록/페이지네이션은 pageIndex 파라미터의 실제 URL이라 requests로 수집 가능합니다.
# 실제 DOM 확인 결과(로컬 진단):
#   - 공고 하나 = <div class="table_box"> ... </div> (형제가 여러 개, 반복 컨테이너)
#   - 제목 링크 = <a onclick="f_detail('I22474', '2026'); return false;">
#                   <span class="title">제목</span></a>
#   - 접수기간/등록일은 <div class="info"> 안 <p><span class="label">라벨</span>값</p>
# 상세페이지 실제 URL 패턴은 확인 못했고, f_detail의 두 인자(공고ID, 연도)로
# 추정 URL을 만들어둔 상태입니다(미검증). 클릭해서 실제 페이지가 맞는지 확인 후
# 알려주시면 정확한 패턴으로 고치겠습니다.

_KEIT_STATUS_WORDS = ("종료", "접수마감", "접수중", "접수예정")  # 우선순위 순서
_KEIT_DETAIL_RE = re.compile(r"f_detail\(\s*'([^']+)'\s*,\s*'([^']+)'\s*\)")
_KEIT_PERIOD_RE = re.compile(
    r"접수기간\s*([\d]{4}-\d{2}-\d{2}\s+[\d:]+\s*~\s*[\d]{4}-\d{2}-\d{2}\s+[\d:]+)"
)
_KEIT_REGDATE_RE = re.compile(r"등록일\s*([\d]{4}-\d{2}-\d{2})")


def fetch_keit_page(page_index: int) -> str:
    params = {"prgmId": KEIT_PRGM_ID, "pageIndex": page_index}
    resp = requests.get(KEIT_BASE_URL, params=params, headers=HEADERS, timeout=15)
    resp.raise_for_status()
    resp.encoding = resp.apparent_encoding or "utf-8"
    return resp.text


def parse_keit_row(box, org: str = "KEIT") -> Announcement | None:
    """<div class="table_box"> 한 항목을 Announcement로 변환."""
    a_tag = box.find("a", onclick=lambda v: v and "f_detail" in v)
    if a_tag is None:
        return None

    title_span = a_tag.find("span", class_="title")
    제목 = title_span.get_text(strip=True) if title_span else a_tag.get_text(strip=True)
    if not 제목:
        return None

    onclick = a_tag.get("onclick", "")
    m = _KEIT_DETAIL_RE.search(onclick)
    ancm_id, ancm_year = (m.group(1), m.group(2)) if m else ("", "")

    if ancm_id:
        # 추정 URL: f_detail의 두 인자를 그대로 붙여봄 (미검증 - 실제 페이지 확인 필요)
        상세링크 = (
            KEIT_BASE_URL.replace("retrieveTaskAnncmListView.do", "retrieveTaskAnncmView.do")
            + f"?prgmId={KEIT_PRGM_ID}&ancmId={ancm_id}&ancmYy={ancm_year}"
        )
    else:
        상세링크 = onclick or KEIT_BASE_URL

    text = box.get_text(separator=" ", strip=True)

    접수기간_m = _KEIT_PERIOD_RE.search(text)
    접수기간 = 접수기간_m.group(1) if 접수기간_m else ""

    등록일_m = _KEIT_REGDATE_RE.search(text)
    등록일 = 등록일_m.group(1) if 등록일_m else ""

    상태배지 = "IRIS 공고" if "IRIS 공고" in text else ""

    마감상태 = ""
    for w in _KEIT_STATUS_WORDS:
        if w in text:
            마감상태 = w
            break

    접수시작일, 접수종료일 = _extract_dates(접수기간)

    return Announcement(
        기관명=org,
        번호=ancm_id,
        남은기간=마감상태,
        제목=제목,
        비고=상태배지,  # "IRIS 공고" 표시 (IRIS 연동 공고 여부)
        접수시작일=접수시작일,
        접수종료일=접수종료일,
        작성자="",
        공고일자=등록일,
        상세링크=상세링크,
    )


def parse_keit_list(html: str, org: str = "KEIT") -> list[Announcement]:
    soup = BeautifulSoup(html, "html.parser")
    results = []
    for box in soup.find_all("div", class_="table_box"):
        item = parse_keit_row(box, org)
        if item is not None:
            results.append(item)
    return results


def crawl_keit(start_page: int, end_page: int, org: str = "KEIT", delay_range=(0.5, 1.0)) -> list[Announcement]:
    all_items: list[Announcement] = []
    for page in range(start_page, end_page + 1):
        print(f"[KEIT {page}/{end_page}] 수집 중...", file=sys.stderr)
        try:
            html = fetch_keit_page(page)
        except requests.RequestException as e:
            print(f"  -> 요청 실패: {e}", file=sys.stderr)
            continue

        items = parse_keit_list(html, org)
        if not items:
            print("  -> 공고 없음 (마지막 페이지이거나 구조 변경 가능성)", file=sys.stderr)
        all_items.extend(items)

        time.sleep(random.uniform(*delay_range))

    return all_items


# ---------------------------------------------------------------------------
# KIAT (한국산업기술진흥원) 입찰공고
# ---------------------------------------------------------------------------
# 목록은 최초 페이지 로딩 시 비어있고, POST boardContentsListAjax.do 요청으로
# 채워지는 AJAX 방식입니다 (로컬 진단으로 확인됨). requests.post로 동일한
# 파라미터를 보내면 목록 HTML 조각을 그대로 받을 수 있습니다.
#
# 실제 확인된 행 구조:
#   <table class="list fixed listTypeA"> 안 각 <tr>에
#     td.td_number(번호), td.td_title > a[href="javascript:contentsView('CONTENTS_ID')"](제목),
#     td.td_reg_date(등록일), td.td_app_term(접수기간, "YYYY-MM-DD~YYYY-MM-DD"),
#     td.td_app_state > span.app_state[data-start][data-end] (상태는 JS로 계산되어 서버 응답엔 비어있음
#     -> data-start/data-end를 이용해 파이썬에서 직접 계산)
#
# 상세링크는 contentsView('CONTENTS_ID')의 CONTENTS_ID로 boardContentsView.do를
# 추정 조합한 것으로, 정확한 URL 패턴은 로컬에서 한 번 클릭 확인이 필요합니다.


def fetch_kiat_page(page_no: int) -> str:
    data = {
        "miv_pageNo": str(page_no),
        "miv_pageSize": "",
        "total_cnt": "",
        "LISTOP": "",
        "mode": "W",
        "contents_id": "",
        "board_id": KIAT_BOARD_ID,
        "cate_id": "",
        "field_id": "",
        "intropage_boardUseYn": "",
        "MenuId": KIAT_MENU_ID,
        "state_filter": "W",
        "contents_year": "",
        "start_date": "",
        "end_date": "",
        "searchkey": "T",
        "searchtxt": "",
    }
    resp = requests.post(KIAT_LIST_AJAX_URL, data=data, headers=HEADERS, timeout=15)
    resp.raise_for_status()
    resp.encoding = resp.apparent_encoding or "utf-8"
    return resp.text


def _kiat_status_from_dates(start: str, end: str) -> str:
    """data-start/data-end(YYYY-MM-DD)를 오늘 날짜와 비교해 상태 문자열 계산."""
    try:
        today = _dt.date.today()
        start_d = _dt.date.fromisoformat(start) if start else None
        end_d = _dt.date.fromisoformat(end) if end else None
    except ValueError:
        return ""

    if start_d and today < start_d:
        return "진행전"
    if end_d and today > end_d:
        return "접수마감"
    if start_d and end_d and start_d <= today <= end_d:
        return "접수중"
    return ""


def parse_kiat_row(tr, org: str = "KIAT") -> Announcement | None:
    td_title = tr.find("td", class_="td_title")
    if td_title is None:
        return None

    a_tag = td_title.find("a")
    if a_tag is None:
        return None

    제목 = a_tag.get_text(strip=True)
    href = a_tag.get("href", "")
    m = re.search(r"contentsView\('([^']+)'\)", href)
    contents_id = m.group(1) if m else ""

    if contents_id:
        # 추정 URL (미검증 - 실제 페이지 확인 필요)
        상세링크 = (
            f"{KIAT_VIEW_URL}?board_id={KIAT_BOARD_ID}"
            f"&contents_id={contents_id}&MenuId={KIAT_MENU_ID}"
        )
    else:
        상세링크 = href or KIAT_LIST_AJAX_URL

    td_number = tr.find("td", class_="td_number")
    번호 = td_number.get_text(strip=True) if td_number else ""

    td_reg_date = tr.find("td", class_="td_reg_date")
    공고일자 = td_reg_date.get_text(strip=True) if td_reg_date else ""

    td_app_term = tr.find("td", class_="td_app_term")
    신청기간_원문 = td_app_term.get_text(strip=True) if td_app_term else ""
    접수시작일, 접수종료일 = _extract_dates(신청기간_원문)

    남은기간 = ""
    state_span = tr.find("span", class_="app_state")
    if state_span is not None:
        남은기간 = _kiat_status_from_dates(
            state_span.get("data-start", ""), state_span.get("data-end", "")
        )

    return Announcement(
        기관명=org,
        번호=번호,
        남은기간=남은기간,
        제목=제목,
        비고="",
        접수시작일=접수시작일,
        접수종료일=접수종료일,
        작성자="",
        공고일자=공고일자,
        상세링크=상세링크,
    )


def parse_kiat_list(html: str, org: str = "KIAT") -> list[Announcement]:
    soup = BeautifulSoup(html, "html.parser")
    table = soup.find("table", class_="list")
    if table is None:
        return []

    results = []
    for tr in table.find_all("tr"):
        item = parse_kiat_row(tr, org)
        if item is not None:
            results.append(item)
    return results


def fetch_kiat_detail(상세링크: str) -> str:
    resp = requests.get(상세링크, headers=HEADERS, timeout=15)
    resp.raise_for_status()
    resp.encoding = resp.apparent_encoding or "utf-8"
    return resp.text


def _extract_kiat_period_from_detail(html: str) -> str:
    """상세페이지의 '접수기간 / 상태' 표 행(예: '2026-07-24~2026-08-05 [진행중]')에서
    접수기간 텍스트를 추출한다. 못 찾으면 빈 문자열."""
    soup = BeautifulSoup(html, "html.parser")
    for cell in soup.find_all(["th", "td"]):
        label = cell.get_text(strip=True)
        if "접수기간" in label:
            value_cell = cell.find_next_sibling("td")
            if value_cell is not None:
                return value_cell.get_text(strip=True)
    return ""


def crawl_kiat(start_page: int, end_page: int, org: str = "KIAT", delay_range=(0.5, 1.0)) -> list[Announcement]:
    all_items: list[Announcement] = []
    for page in range(start_page, end_page + 1):
        print(f"[KIAT {page}/{end_page}] 수집 중...", file=sys.stderr)
        try:
            html = fetch_kiat_page(page)
        except requests.RequestException as e:
            print(f"  -> 요청 실패: {e}", file=sys.stderr)
            continue

        items = parse_kiat_list(html, org)
        if not items:
            print("  -> 공고 없음 (마지막 페이지이거나 구조 변경 가능성)", file=sys.stderr)

        # 목록에서 접수기간을 못 얻은 경우, 상세페이지의 '접수기간 / 상태'에서 보완
        for idx, item in enumerate(items):
            if item.접수시작일 or item.접수종료일 or not item.상세링크:
                continue
            try:
                detail_html = fetch_kiat_detail(item.상세링크)
                period_text = _extract_kiat_period_from_detail(detail_html)
                if period_text:
                    시작일, 종료일 = _extract_dates(period_text)
                    items[idx] = replace(item, 접수시작일=시작일, 접수종료일=종료일)
            except requests.RequestException as e:
                print(f"  -> 상세페이지 요청 실패({item.제목}): {e}", file=sys.stderr)
            time.sleep(random.uniform(*delay_range))

        all_items.extend(items)

        time.sleep(random.uniform(*delay_range))

    return all_items


# ---------------------------------------------------------------------------
# KISA (한국인터넷진흥원) 입찰공고
# ---------------------------------------------------------------------------
# 일반 HTML <table>로 서버 렌더링되고, 상세링크도 실제 URL이 그대로 노출되어
# (예: /403/form?postSeq=10778&page=1) requests + BeautifulSoup만으로 충분합니다.
# 페이지네이션은 ?page= 파라미터로 추정했으나(상세링크에 &page=1이 붙는 것으로 유추),
# 실제 목록 URL에서도 동일하게 동작하는지는 로컬 확인 권장.

KISA_COLS = ("번호", "제목", "등록일", "조회수", "첨부파일")


def fetch_kisa_page(page: int) -> str:
    resp = requests.get(KISA_URL, params={"page": page}, headers=HEADERS, timeout=15)
    resp.raise_for_status()
    resp.encoding = resp.apparent_encoding or "utf-8"
    return resp.text


def parse_kisa_row(tr, org: str = "KISA") -> Announcement | None:
    tds = tr.find_all("td")
    if len(tds) < 3:
        return None  # 헤더 행 등

    번호 = tds[0].get_text(strip=True)

    a_tag = tds[1].find("a")
    if a_tag is None:
        return None
    제목 = a_tag.get_text(strip=True)
    href = a_tag.get("href", "")
    상세링크 = href if href.startswith("http") else f"https://www.kisa.or.kr{href}"

    공고일자 = tds[2].get_text(strip=True) if len(tds) > 2 else ""

    return Announcement(
        기관명=org,
        번호=번호,
        남은기간="",
        제목=제목,
        비고="",
        작성자="",
        공고일자=공고일자,
        상세링크=상세링크,
    )


def parse_kisa_list(html: str, org: str = "KISA") -> list[Announcement]:
    soup = BeautifulSoup(html, "html.parser")
    table = soup.find("table")
    if table is None:
        return []

    results = []
    for tr in table.find_all("tr"):
        item = parse_kisa_row(tr, org)
        if item is not None:
            results.append(item)
    return results


def fetch_kisa_detail(상세링크: str) -> str:
    resp = requests.get(상세링크, headers=HEADERS, timeout=15)
    resp.raise_for_status()
    resp.encoding = resp.apparent_encoding or "utf-8"
    return resp.text


def _extract_kisa_period_from_detail(html: str) -> str:
    """상세페이지 공고문 본문에서 접수/마감 관련 문구를 찾아 그 주변 텍스트를
    반환한다(표가 아니라 서술형 문장이라 KIAT보다 정확도가 낮을 수 있음 -
    실제 값과 비교 확인 권장). 못 찾으면 빈 문자열."""
    soup = BeautifulSoup(html, "html.parser")
    body_text = _compact_loose_dates(soup.get_text("\n", strip=True))

    for keyword in ("접수기간", "제출기간", "신청기간", "공개기간", "마감일시", "등록마감", "마감"):
        idx = body_text.find(keyword)
        if idx == -1:
            continue
        window = _keyword_window(body_text, idx)
        if _DATE_IN_TEXT_RE.search(window):
            return window
    return ""


def crawl_kisa(start_page: int, end_page: int, org: str = "KISA", delay_range=(0.5, 1.0)) -> list[Announcement]:
    all_items: list[Announcement] = []
    for page in range(start_page, end_page + 1):
        print(f"[KISA {page}/{end_page}] 수집 중...", file=sys.stderr)
        try:
            html = fetch_kisa_page(page)
        except requests.RequestException as e:
            print(f"  -> 요청 실패: {e}", file=sys.stderr)
            continue

        items = parse_kisa_list(html, org)
        if not items:
            print("  -> 공고 없음 (마지막 페이지이거나 구조 변경 가능성)", file=sys.stderr)

        # 목록에 접수기간이 없으므로, 상세페이지 본문에서 마감/접수기간 문구를 찾아 보완
        for idx, item in enumerate(items):
            if item.접수시작일 or item.접수종료일 or not item.상세링크:
                continue
            try:
                detail_html = fetch_kisa_detail(item.상세링크)
                period_text = _extract_kisa_period_from_detail(detail_html)
                if period_text:
                    시작일, 종료일 = _extract_dates(period_text)
                    items[idx] = replace(item, 접수시작일=시작일, 접수종료일=종료일)
            except requests.RequestException as e:
                print(f"  -> 상세페이지 요청 실패({item.제목}): {e}", file=sys.stderr)
            time.sleep(random.uniform(*delay_range))

        all_items.extend(items)

        time.sleep(random.uniform(*delay_range))

    return all_items


# ---------------------------------------------------------------------------
# SMTECH (중소기업기술개발사업 종합관리시스템) 사업공고
# ---------------------------------------------------------------------------
# 일반 HTML <table>로 서버 렌더링됩니다. 컬럼: No / 시스템구분(SMTECH·IRIS) /
# 비고 / 제목(링크) / 접수기간 / 공고일 / 상태(아이콘 alt 텍스트).
# 게시 건수가 많아 --end로 지정한 페이지마다 SMTECH_ROWS_PER_PAGE(기본 30)건씩
# 가져와 요청 횟수를 줄입니다.
# 주의: 시스템구분이 "IRIS"인 행은 상세링크가 javascript:goMove()로 되어 있어
# 실제 URL을 얻을 수 없습니다 (IRIS 자체 시스템에서 관리되는 공고). 이 경우
# 상세링크는 목록 URL로 대체되고, 작성자 컬럼에 "IRIS"라고 표시됩니다.


def fetch_smtech_page(page: int, session: requests.Session) -> str:
    params = {"cpage": page, "sort": "latest", "rows": SMTECH_ROWS_PER_PAGE}
    resp = session.get(SMTECH_LIST_URL, params=params, headers=HEADERS, timeout=15)
    resp.raise_for_status()
    resp.encoding = resp.apparent_encoding or "utf-8"
    return resp.text


def parse_smtech_row(tr, org: str = "SMTECH") -> Announcement | None:
    tds = tr.find_all("td")
    if len(tds) < 7:
        return None  # 헤더 행 등

    번호 = tds[0].get_text(strip=True)
    시스템구분 = tds[1].get_text(strip=True)
    비고 = tds[2].get_text(strip=True)

    a_tag = tds[3].find("a")
    제목 = a_tag.get_text(strip=True) if a_tag else tds[3].get_text(strip=True)
    if not 제목:
        return None

    href = a_tag.get("href", "") if a_tag else ""
    if href.startswith("http"):
        상세링크 = _strip_session_id(href)
    elif href.startswith("/"):
        상세링크 = _strip_session_id(f"https://www.smtech.go.kr{href}")
    else:
        상세링크 = SMTECH_LIST_URL  # javascript:goMove() 등 (IRIS 연동 공고, 실제 링크 없음)

    접수기간 = tds[4].get_text(strip=True)
    공고일 = tds[5].get_text(strip=True)
    접수시작일, 접수종료일 = _extract_dates(접수기간)

    상태_img = tds[6].find("img")
    남은기간 = 상태_img.get("alt", "") if 상태_img else tds[6].get_text(strip=True)

    return Announcement(
        기관명=org,
        번호=번호,
        남은기간=남은기간,
        제목=제목,
        비고=비고,
        접수시작일=접수시작일,
        접수종료일=접수종료일,
        작성자=시스템구분,  # "SMTECH" 또는 "IRIS" (연동 출처 구분)
        공고일자=공고일,
        상세링크=상세링크,
    )


def parse_smtech_list(html: str, org: str = "SMTECH") -> list[Announcement]:
    soup = BeautifulSoup(html, "html.parser")
    tables = soup.find_all("table")
    if not tables:
        return []

    # 페이지에 검색 필터용 작은 table도 함께 있어서, tr이 가장 많은(=실제 목록) table을 선택
    table = max(tables, key=lambda t: len(t.find_all("tr")))

    results = []
    for tr in table.find_all("tr"):
        item = parse_smtech_row(tr, org)
        if item is not None:
            results.append(item)
    return results


def crawl_smtech(start_page: int, end_page: int, org: str = "SMTECH", delay_range=(0.5, 1.0)) -> list[Announcement]:
    all_items: list[Announcement] = []
    session = requests.Session()
    for page in range(start_page, end_page + 1):
        print(f"[SMTECH {page}/{end_page}] 수집 중...", file=sys.stderr)
        try:
            html = fetch_smtech_page(page, session)
        except requests.RequestException as e:
            print(f"  -> 요청 실패: {e}", file=sys.stderr)
            continue

        items = parse_smtech_list(html, org)
        if not items:
            print("  -> 공고 없음 (마지막 페이지이거나 구조 변경 가능성)", file=sys.stderr)
        all_items.extend(items)

        time.sleep(random.uniform(*delay_range))

    return all_items


# ---------------------------------------------------------------------------
# IRIS (범부처통합연구지원시스템)
# ---------------------------------------------------------------------------
# 로컬 진단 결과, 목록 자체는 requests만으로 정상적으로 옵니다 (Playwright 불필요).
# 실제 확인된 구조:
#   <li>
#     <span class="inst_title">소관부처 &gt; 전문기관</span>
#     <div class="form-row">
#       <div class="group1">
#         <strong class="title"><a onclick="f_bsnsAncmBtinSituListForm_view('023097','ancmIng'); return false;">제목</a></strong>
#         <div class="etc_info">
#           <span><em>공고번호 :</em>...</span>
#           <span class="ancmDe"><em>공고일자 :</em>2026-07-15</span>
#           <span class="rcveSttSeNmLst"><em>공고상태 :</em>공고접수중</span>
#           <span class="pbofrTpSeNmLst"><em>공모유형 :</em>지정공모</span>
#         </div>
#       </div>
#       <div class="group2"><span class="d_day end">접수중</span></div>
#     </div>
#   </li>
# 상세링크는 onclick의 첫 번째 인자(공고 ID)로 추정 URL을 구성했으나 미검증입니다.
# 페이지네이션 파라미터(pageIndex로 추정)도 로컬 확인이 필요합니다.

IRIS_ANCM_RE = re.compile(r"f_bsnsAncmBtinSituListForm_view\('([^']+)'\s*,\s*'([^']+)'\)")


def fetch_iris_page(page_index: int) -> str:
    resp = requests.get(IRIS_URL, params={"pageIndex": page_index}, headers=HEADERS, timeout=15)
    resp.raise_for_status()
    resp.encoding = resp.apparent_encoding or "utf-8"
    return resp.text


def parse_iris_row(li) -> Announcement | None:
    a_tag = li.find("a", onclick=lambda v: v and "f_bsnsAncmBtinSituListForm_view" in v)
    if a_tag is None:
        return None

    제목 = a_tag.get_text(strip=True)
    if not 제목:
        return None

    onclick = a_tag.get("onclick", "")
    m = IRIS_ANCM_RE.search(onclick)
    ancm_id, ancm_stt = (m.group(1), m.group(2)) if m else ("", "")

    if ancm_id:
        # 추정 URL (미검증 - 실제 페이지 확인 필요)
        상세링크 = f"https://www.iris.go.kr/contents/retrieveBsnsAncmView.do?ancmId={ancm_id}"
    else:
        상세링크 = onclick or IRIS_URL

    inst_span = li.find("span", class_="inst_title")
    소속 = inst_span.get_text(strip=True) if inst_span else ""
    소관부처, _, 전문기관 = 소속.partition(">")
    소관부처 = 소관부처.strip()
    전문기관 = 전문기관.strip() or 소속

    ancmDe = li.find("span", class_="ancmDe")
    공고일자 = ancmDe.get_text(strip=True).replace("공고일자 :", "").replace("공고일자", "").strip() if ancmDe else ""

    pbofr = li.find("span", class_="pbofrTpSeNmLst")
    공모유형 = pbofr.get_text(strip=True).replace("공모유형 :", "").replace("공모유형", "").strip() if pbofr else ""

    dday = li.find("span", class_="d_day")
    남은기간 = dday.get_text(strip=True) if dday else ""

    return Announcement(
        기관명=전문기관 or 소관부처 or "IRIS",
        번호=ancm_id,
        남은기간=남은기간,
        제목=제목,
        비고=f"{소관부처} / {공모유형}".strip(" /") if 공모유형 else 소관부처,
        작성자="",
        공고일자=공고일자,
        상세링크=상세링크,
    )


def parse_iris_list(html: str) -> list[Announcement]:
    soup = BeautifulSoup(html, "html.parser")
    results = []
    for li in soup.find_all("li"):
        item = parse_iris_row(li)
        if item is not None:
            results.append(item)
    return results


def fetch_iris_detail(상세링크: str) -> str:
    resp = requests.get(상세링크, headers=HEADERS, timeout=15)
    resp.raise_for_status()
    resp.encoding = resp.apparent_encoding or "utf-8"
    return resp.text


def _extract_iris_period_from_detail(html: str) -> str:
    """상세페이지의 '접수기간' 항목(예: '2022-04-01 ~ 2022-04-20')에서
    날짜 범위 텍스트를 가져온다. 실제 상세페이지(retrieveBsnsAncmView.do)에서
    이 라벨과 값이 존재함을 확인함. 못 찾으면 빈 문자열."""
    soup = BeautifulSoup(html, "html.parser")
    label_tag = soup.find(lambda t: t.name not in ("script", "style") and t.get_text(strip=True) == "접수기간")
    if label_tag is None:
        return ""
    value_tag = label_tag.find_next_sibling()
    if value_tag is None:
        value_tag = label_tag.find_next(["dd", "td", "span", "div", "p"])
    return value_tag.get_text(strip=True) if value_tag is not None else ""


def crawl_iris(start_page: int, end_page: int, delay_range=(0.5, 1.0)) -> list[Announcement]:
    all_items: list[Announcement] = []
    for page in range(start_page, end_page + 1):
        print(f"[IRIS {page}/{end_page}] 수집 중...", file=sys.stderr)
        try:
            html = fetch_iris_page(page)
        except requests.RequestException as e:
            print(f"  -> 요청 실패: {e}", file=sys.stderr)
            continue

        items = parse_iris_list(html)
        if not items:
            print("  -> 공고 없음 (마지막 페이지이거나 구조 변경 가능성)", file=sys.stderr)

        # 목록에 접수기간이 없으므로, 상세페이지의 '접수기간' 항목에서 보완
        for idx, item in enumerate(items):
            if item.접수시작일 or item.접수종료일 or not item.상세링크:
                continue
            try:
                detail_html = fetch_iris_detail(item.상세링크)
                period_text = _extract_iris_period_from_detail(detail_html)
                if period_text:
                    시작일, 종료일 = _extract_dates(period_text)
                    items[idx] = replace(item, 접수시작일=시작일, 접수종료일=종료일)
            except requests.RequestException as e:
                print(f"  -> 상세페이지 요청 실패({item.제목}): {e}", file=sys.stderr)
            time.sleep(random.uniform(*delay_range))

        all_items.extend(items)

        time.sleep(random.uniform(*delay_range))

    return all_items


# ---------------------------------------------------------------------------
# 중소벤처24 (SMES24) 공고정보 Open API
# ---------------------------------------------------------------------------
# 화면 크롤링이 아니라 공식 API를 사용한다(문서 확인 완료: 요청/응답 스펙은
# https://portal.smes.go.kr/home/cs/opndata/UI_USR_L_210/supportBusinessInfoApi
# 참고). GET 방식, JSON 응답, 페이지네이션 파라미터가 없어 한 번의 호출로
# 전체 데이터를 받아온다. 인증키는 환경변수 SMES_API_KEY로 전달한다
# (GitHub Actions Secrets 또는 로컬 환경변수로 설정).

SMES_API_URL = "https://www.smes.go.kr/fnct/apiReqst/extPblancInfo"


def crawl_smes(delay_range=(0.5, 1.0), lookback_days: int = 90) -> list[Announcement]:
    token = os.environ.get("SMES_API_KEY")
    if not token:
        print("[SMES] SMES_API_KEY 환경변수가 없어 건너뜁니다", file=sys.stderr)
        return []

    # 페이지네이션이 없는 API라 기간을 지정하지 않으면 전체 데이터를 한 번에
    # 내려받으려다 응답이 느려질 수 있어(타임아웃 발생 확인됨), 최근
    # lookback_days일 치만 조회한다. 신규 공고 판별은 어차피 DB 대조로
    # 하므로 이 정도 범위면 충분하다.
    오늘 = _dt.date.today()
    strDt = (오늘 - _dt.timedelta(days=lookback_days)).strftime("%Y%m%d")
    endDt = 오늘.strftime("%Y%m%d")

    print(f"[SMES] 공고정보 API 조회 중... ({strDt}~{endDt})", file=sys.stderr)
    try:
        resp = requests.get(
            SMES_API_URL,
            params={"token": token, "html": "no", "strDt": strDt, "endDt": endDt},
            headers=HEADERS, timeout=30,
        )
        resp.raise_for_status()
        payload = resp.json()
    except requests.RequestException as e:
        print(f"  -> 요청 실패: {e}", file=sys.stderr)
        return []
    except ValueError as e:
        print(f"  -> 응답 JSON 파싱 실패: {e}", file=sys.stderr)
        return []

    data = payload.get("data")
    if not isinstance(data, list):
        print(
            f"  -> 데이터 없음 (resultCd={payload.get('resultCd')}, "
            f"resultMsg={payload.get('resultMsg')})", file=sys.stderr,
        )
        return []

    items = []
    for row in data:
        공고일자 = str(row.get("creatDt") or "").split(" ")[0]  # 'yyyy-MM-dd HH:mm:ss' -> 날짜만
        items.append(Announcement(
            기관명=row.get("sportInsttNm") or "중소벤처24",
            번호=str(row.get("pblancSeq") or ""),
            남은기간="",
            제목=row.get("pblancNm") or "",
            비고=row.get("detailBsnsNm") or "",
            접수시작일=row.get("pblancBgnDt") or "",
            접수종료일=row.get("pblancEndDt") or "",
            작성자="",
            공고일자=공고일자,
            상세링크=row.get("pblancDtlUrl") or "",
        ))

    print(f"  -> {len(items)}건 수집", file=sys.stderr)
    time.sleep(random.uniform(*delay_range))
    return items



def _load_dataframe(path: str):
    import pandas as pd

    if not os.path.exists(path):
        return None

    try:
        if path.lower().endswith(".xlsx"):
            return pd.read_excel(path, dtype=str)
        return pd.read_csv(path, dtype=str, encoding="utf-8-sig")
    except pd.errors.EmptyDataError:
        # 파일은 있지만 내용이 비어있는 경우(헤더까지 삭제된 경우) - 기존 데이터 없음으로 취급
        return None
    except pd.errors.ParserError as e:
        # 파일 내용이 정상적인 CSV 형식이 아닌 경우(수동 편집 중 깨짐 등) -
        # 죽지 않고 기존 데이터 없음으로 취급, 다음 저장 시 정상 형식으로 다시 만들어진다.
        print(f"[경고] {path} 파일이 정상적인 CSV 형식이 아니라 읽지 못했습니다 ({e}). "
              f"기존 데이터 없음으로 취급하고 새로 만듭니다.", file=sys.stderr)
        return None


def _save_dataframe(df, path: str) -> None:
    if path.lower().endswith(".xlsx"):
        df.to_excel(path, index=False)
    else:
        df.to_csv(path, index=False, encoding="utf-8-sig")


def find_new_items(items: list[Announcement], db_path: str) -> list[Announcement]:
    """db_path에 없는(=새로 올라온) 공고만 골라낸다."""
    existing_df = _load_dataframe(db_path)
    if existing_df is None or UNIQUE_KEY not in existing_df.columns:
        # DB가 없으면 이번에 수집한 전체가 신규
        return items

    existing_keys = set(existing_df[UNIQUE_KEY].dropna().astype(str))
    return [i for i in items if i.상세링크 not in existing_keys]


def update_db(items: list[Announcement], db_path: str) -> int:
    """새로 발견/재크롤링된 공고를 DB 파일에 반영. 반환값: DB 전체 건수.

    - 같은 상세링크가 기존 DB에도 있으면, 이번에 새로 긁어온 값(남은기간 등 최신
      정보)으로 덮어쓴다(keep="last"). 이미 목록에서 사라져 이번에 다시 크롤링
      되지 않은 옛 공고는 그대로 유지된다.
    - 마감Dday는 재크롤링 여부와 상관없이 DB 전체를 대상으로 매번 다시 계산한다
      (저장된 접수종료일 + 오늘 날짜만 있으면 계산 가능하므로 항상 최신 유지 가능).
    """
    import pandas as pd

    field_order = [f.name for f in fields(Announcement)]
    new_df = pd.DataFrame([asdict(i) for i in items])[field_order]

    existing_df = _load_dataframe(db_path)
    if existing_df is not None:
        merged = pd.concat([existing_df, new_df], ignore_index=True)
        # 새로 크롤링한 행이 new_df(뒤쪽)에 있으므로 keep="last"로 최신값 우선
        merged = merged.drop_duplicates(subset=[UNIQUE_KEY], keep="last")
    else:
        merged = new_df.drop_duplicates(subset=[UNIQUE_KEY], keep="last")

    if "접수종료일" in merged.columns:
        merged["마감Dday"] = merged["접수종료일"].fillna("").apply(_compute_dday)

    if "공고일자" in merged.columns:
        merged["공고일자"] = merged["공고일자"].fillna("").apply(_normalize_written_date)
        merged = merged.sort_values("공고일자", ascending=False, kind="stable")

    merged = merged[field_order]  # 컬럼 순서를 항상 최신 Announcement 정의 순서로 고정

    _save_dataframe(merged, db_path)
    return len(merged)


def upload_to_google_sheet(db_path, sheet_name="gongo"):
    """DB 파일을 읽어 Google Sheet에 덮어쓴다.
    인증 정보는 GitHub Actions 환경변수(GCP_SA_KEY) 또는 로컬 service_account.json
    파일에서 가져온다."""
    print("[구글시트] 업로드 시작", file=sys.stderr)
    try:
        SCOPES = [
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive",
        ]

        json_creds = os.environ.get("GCP_SA_KEY")
        sa_file_path = os.path.join(SCRIPT_DIR, "service_account.json")
        if json_creds:
            print("[구글시트] GCP_SA_KEY 환경변수로 인증 시도", file=sys.stderr)
            creds_dict = json.loads(json_creds)
            creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
        elif os.path.exists(sa_file_path):
            print(f"[구글시트] service_account.json 파일로 인증 시도 ({sa_file_path})", file=sys.stderr)
            creds = Credentials.from_service_account_file(
                sa_file_path, scopes=SCOPES
            )
        else:
            print(f"[구글시트] 인증 정보를 찾을 수 없습니다. GCP_SA_KEY 환경변수도 없고, {sa_file_path} 파일도 없습니다", file=sys.stderr)
            return

        client = gspread.authorize(creds)
        print("[구글시트] 인증 성공", file=sys.stderr)

        # 구글 시트 열기
        spreadsheet = client.open(sheet_name)
        worksheet = spreadsheet.sheet1
        print(f"[구글시트] '{sheet_name}' 시트 열기 성공", file=sys.stderr)

        # DB 파일 읽어와서 시트에 덮어쓰기
        if not os.path.exists(db_path):
            print(f"[구글시트] DB 파일을 찾을 수 없습니다: {db_path}", file=sys.stderr)
            return

        df = _load_dataframe(db_path)
        if df is None:
            print(f"[구글시트] DB 파일을 읽지 못했습니다(빈 파일이거나 형식 문제): {db_path}", file=sys.stderr)
            return

        df = df.fillna("")
        data = [df.columns.values.tolist()] + df.values.tolist()
        worksheet.clear()
        worksheet.update(data)
        print(f"[구글시트] 업로드 성공! ({len(df)}행)", file=sys.stderr)

    except Exception as e:
        print(f"[구글시트] 업로드 실패 - 원인: {type(e).__name__}: {e}", file=sys.stderr)
        

def _run_git(args_list: list[str]) -> bool:
    """git 명령을 SCRIPT_DIR 기준으로 실행. 성공하면 True, 실패해도 프로그램은 계속 진행."""
    import subprocess

    try:
        result = subprocess.run(
            ["git", "-C", SCRIPT_DIR] + args_list,
            capture_output=True, text=True, encoding="utf-8", errors="replace",
        )
        output = (result.stdout or "").strip() + (result.stderr or "").strip()
        if result.returncode == 0:
            print(f"[git] {' '.join(args_list)} 성공" + (f" - {output}" if output else ""), file=sys.stderr)
            return True
        else:
            print(f"[git] {' '.join(args_list)} 실패 - {output}", file=sys.stderr)
            return False
    except FileNotFoundError:
        print("[git] git이 설치되어 있지 않거나 PATH에 없어 건너뜁니다", file=sys.stderr)
        return False
    except Exception as e:
        print(f"[git] {' '.join(args_list)} 중 오류 - {type(e).__name__}: {e}", file=sys.stderr)
        return False


def git_pull_latest_db() -> None:
    """크롤링 시작 전, GitHub 저장소의 최신 DB 내용을 받아온다(로컬/온라인 공통)."""
    print("[git] 최신 DB 받아오는 중 (git pull)...", file=sys.stderr)
    _run_git(["pull", "--rebase"])


def git_push_db(db_path: str) -> None:
    """크롤링 및 구글시트 업로드 후, 갱신된 DB 파일을 GitHub 저장소에 반영한다(로컬/온라인 공통)."""
    print("[git] 결과를 저장소에 반영하는 중...", file=sys.stderr)
    _run_git(["add", db_path])
    committed = _run_git(["commit", "-m", "크롤링 결과 자동 반영"])
    if committed:
        _run_git(["push"])
    else:
        print("[git] 커밋할 변경 사항이 없어 push는 건너뜁니다", file=sys.stderr)


def main():
    parser = argparse.ArgumentParser(description="전담기관 사업공고 크롤러")
    parser.add_argument(
        "--site", type=str, choices=["nipa", "iris", "nia", "keit", "kiat", "kisa", "smtech", "smes", "all"], default="all",
        help="크롤링할 사이트 (기본 all = nipa+nia+keit+kiat+kisa+smtech+iris+smes 한 번에 실행)"
    )
    parser.add_argument("--start", type=int, default=1, help="시작 페이지 (기본 1)")
    parser.add_argument("--end", type=int, default=1, help="종료 페이지 (기본 1)")
    parser.add_argument(
        "--db", type=str, default=DEFAULT_DB_PATH,
        help="누적 DB 파일 경로 (.csv 또는 .xlsx, 기본: crawling.py와 같은 폴더의 crawling_db.csv). "
             "이 파일과 비교해서 신규 공고를 판별하고, 신규 공고를 여기에 누적 저장합니다."
    )
    parser.add_argument(
        "--new-out", type=str, default=None,
        help="이번 실행에서 새로 발견된 공고만 별도로 저장할 파일 경로 (선택, 예: new_today.csv)"
    )
    parser.add_argument(
        "--org", type=str, default="NIPA",
        help="[NIPA 전용] 공고를 게시한 기관명 (기본 NIPA). 다른 사이트는 자동으로 채워집니다."
    )
    args = parser.parse_args()

    git_pull_latest_db()

    if args.site == "nipa":
        items = crawl_nipa(args.start, args.end, args.org)
    elif args.site == "nia":
        items = crawl_nia(args.start, args.end)
    elif args.site == "keit":
        items = crawl_keit(args.start, args.end)
    elif args.site == "kiat":
        items = crawl_kiat(args.start, args.end)
    elif args.site == "kisa":
        items = crawl_kisa(args.start, args.end)
    elif args.site == "smtech":
        items = crawl_smtech(args.start, args.end)
    elif args.site == "iris":
        items = crawl_iris(args.start, args.end)
    elif args.site == "smes":
        items = crawl_smes()
    elif args.site == "all":
        items = []
        items += crawl_nipa(args.start, args.end, args.org)
        items += crawl_nia(args.start, args.end)
        items += crawl_keit(args.start, args.end)
        items += crawl_kiat(args.start, args.end)
        items += crawl_kisa(args.start, args.end)
        items += crawl_smtech(args.start, args.end)
        items += crawl_iris(args.start, args.end)
        items += crawl_smes()

    if not items:
        print("수집된 공고가 없습니다. 사이트 구조가 바뀌었을 수 있으니 파싱 로직을 확인하세요.", file=sys.stderr)
        sys.exit(1)

    items = [enrich_announcement(i) for i in items]

    new_items = find_new_items(items, args.db)

    if not new_items:
        print("신규 공고가 없습니다.", file=sys.stderr)
    else:
        print(f"신규 공고 {len(new_items)}건 발견:", file=sys.stderr)
        for it in new_items:
            print(f"  - [{it.번호}] {it.제목} ({it.접수시작일}~{it.접수종료일})", file=sys.stderr)

        if args.new_out:
            import pandas as pd
            field_order = [f.name for f in fields(Announcement)]
            pd.DataFrame([asdict(i) for i in new_items])[field_order].pipe(
                lambda df: _save_dataframe(df, args.new_out)
            )
            print(f"신규 공고 파일 저장 완료 -> {args.new_out}", file=sys.stderr)

    total = update_db(items, args.db)
    print(f"DB 업데이트 완료 (누적 {total}건) -> {args.db}", file=sys.stderr)

    # 기존 코드 맨 마지막 줄 아래에 추가
    upload_to_google_sheet(args.db, sheet_name="gongo")

    git_push_db(args.db)


if __name__ == "__main__":
    main()

