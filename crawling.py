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
    기관명: str = ""  # 가장 앞쪽 열
    번호: str = ""
    남은기간: str = ""
    제목: str = ""
    사업명: str = ""
    # 아래는 크롤링 시점에 제목/사업명/신청기간을 분석해 자동으로 채우는 파생 컬럼
    # (enrich_announcement 참고). 사업명 바로 뒤에 배치.
    AI관련여부: str = ""
    품질인증관련여부: str = ""
    관련키워드: str = ""
    공고유형: str = ""
    신청시작일: str = ""
    신청종료일: str = ""
    마감Dday: str = ""
    신청기간: str = ""
    작성자: str = ""
    작성일자: str = ""
    상세링크: str = ""


# ---------------------------------------------------------------------------
# AI품질역량센터 관점 자동 분류 (제목/사업명/신청기간 기반 규칙)
# ---------------------------------------------------------------------------
# 크롤링 직후 한 번만 계산해서 Announcement에 채워 넣는다. 신규 공고 판별/DB
# 누적 로직(find_new_items, update_db)은 상세링크 기준으로 동작하므로, 이미
# DB에 있는 공고는 재계산 없이 그대로 유지되고 새로 발견된 공고만 새로 분류된다.
# 규칙(키워드 사전 등)은 실제 데이터를 보며 계속 다듬어야 정확도가 올라간다.

AI_KEYWORDS = ("AI", "인공지능", "데이터", "지능형", "생성형", "머신러닝", "딥러닝", "빅데이터", "챗봇", "LLM")
QUALITY_KEYWORDS = ("품질", "인증", "검증", "시험", "신뢰성", "평가", "표준")

_DATE_IN_TEXT_RE = re.compile(r"(\d{4})[.\-](\d{1,2})[.\-](\d{1,2})")


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
    느슨하게 매칭)에서 시작일/종료일을 뽑는다. 둘 다 없으면 빈 문자열."""
    dates = _DATE_IN_TEXT_RE.findall(period_text or "")
    parsed = []
    for y, m, d in dates:
        try:
            parsed.append(_dt.date(int(y), int(m), int(d)))
        except ValueError:
            continue
    시작일 = parsed[0].isoformat() if len(parsed) >= 1 else ""
    종료일 = parsed[1].isoformat() if len(parsed) >= 2 else ""
    return 시작일, 종료일


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


def enrich_announcement(ann: Announcement) -> Announcement:
    """제목/사업명/신청기간을 분석해 AI품질역량센터 관점의 분류 컬럼을 채운 새 Announcement 반환."""
    text = f"{ann.제목} {ann.사업명}"

    ai_hits = _match_keywords(text, AI_KEYWORDS)
    quality_hits = _match_keywords(text, QUALITY_KEYWORDS)

    시작일, 종료일 = _extract_dates(ann.신청기간)

    return replace(
        ann,
        AI관련여부="Y" if ai_hits else "N",
        품질인증관련여부="Y" if quality_hits else "N",
        관련키워드=", ".join(ai_hits + quality_hits),
        공고유형=_classify_공고유형(ann.제목),
        신청시작일=시작일,
        신청종료일=종료일,
        마감Dday=_compute_dday(종료일),
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

    # title_cell 안의 전체 텍스트에서 제목/신청기간을 제외한 나머지를 사업명으로 추정
    full_text = title_cell.get_text(separator="\n", strip=True)
    lines = [ln.strip() for ln in full_text.split("\n") if ln.strip()]

    사업명 = ""
    신청기간 = ""
    for ln in lines:
        if ln == 제목:
            continue
        m = re.search(r"신청기간\s*:?\s*(.+)", ln)
        if m:
            신청기간 = m.group(1).strip()
        elif not 사업명:
            사업명 = ln

    작성자 = tds[3].get_text(strip=True)
    작성일자 = tds[4].get_text(strip=True)

    return Announcement(
        기관명=org,
        번호=번호,
        남은기간=남은기간,
        제목=제목,
        사업명=사업명,
        신청기간=신청기간,
        작성자=작성자,
        작성일자=작성일자,
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
# 가능합니다. 다만 각 공고의 상세보기는 <a href="#view" onclick="..."> 형태의
# 자바스크립트 기반이라, 이 환경에서는 실제 onclick 함수명/파라미터를 확인하지
# 못했습니다. 아래는 텍스트 패턴("조회수 N", "YYYY.MM.DD" 등) 기준으로 필드를
# 추출하고, 상세링크는 onclick 속성 원문을 그대로 담아둡니다.
# 로컬에서 실행해 onclick 값이 예: fn_egov_bbsView('78336','123456') 같은
# 형태로 나오면, 알려주시면 실제 상세 URL을 만들어내는 로직으로 바꿔드리겠습니다.

_NIA_DATE_RE = re.compile(r"^\d{4}\.\d{2}\.\d{2}$")
_NIA_VIEWCOUNT_RE = re.compile(r"조회수\s*(\d+)")
_NIA_ID_IN_ONCLICK_RE = re.compile(r"'(\d+)'")


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
    작성일자 = next((ln for ln in lines if _NIA_DATE_RE.match(ln)), "")

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

    if href and href.startswith("http"):
        상세링크 = href
    elif onclick:
        상세링크 = onclick  # 임시: 실제 URL 패턴은 로컬 확인 후 교체 필요
    else:
        상세링크 = NIA_BASE_URL

    onclick_ids = _NIA_ID_IN_ONCLICK_RE.findall(onclick) if onclick else []
    번호 = onclick_ids[-1] if onclick_ids else 조회수  # 게시글 고유번호 추정값

    return Announcement(
        기관명=org,
        번호=번호,
        남은기간="신규" if 신규 else "",
        제목=제목,
        사업명=부서,
        신청기간="",
        작성자=작성자,
        작성일자=작성일자,
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

    return Announcement(
        기관명=org,
        번호=ancm_id,
        남은기간=마감상태,
        제목=제목,
        사업명=상태배지,  # "IRIS 공고" 표시 (IRIS 연동 공고 여부)
        신청기간=접수기간,
        작성자="",
        작성일자=등록일,
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
    작성일자 = td_reg_date.get_text(strip=True) if td_reg_date else ""

    td_app_term = tr.find("td", class_="td_app_term")
    신청기간 = td_app_term.get_text(strip=True) if td_app_term else ""

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
        사업명="",
        신청기간=신청기간,
        작성자="",
        작성일자=작성일자,
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

    작성일자 = tds[2].get_text(strip=True) if len(tds) > 2 else ""

    return Announcement(
        기관명=org,
        번호=번호,
        남은기간="",
        제목=제목,
        사업명="",
        신청기간="",
        작성자="",
        작성일자=작성일자,
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
        all_items.extend(items)

        time.sleep(random.uniform(*delay_range))

    return all_items


# ---------------------------------------------------------------------------
# SMTECH (중소기업기술개발사업 종합관리시스템) 사업공고
# ---------------------------------------------------------------------------
# 일반 HTML <table>로 서버 렌더링됩니다. 컬럼: No / 시스템구분(SMTECH·IRIS) /
# 사업명 / 제목(링크) / 접수기간 / 공고일 / 상태(아이콘 alt 텍스트).
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
    사업명 = tds[2].get_text(strip=True)

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

    상태_img = tds[6].find("img")
    남은기간 = 상태_img.get("alt", "") if 상태_img else tds[6].get_text(strip=True)

    return Announcement(
        기관명=org,
        번호=번호,
        남은기간=남은기간,
        제목=제목,
        사업명=사업명,
        신청기간=접수기간,
        작성자=시스템구분,  # "SMTECH" 또는 "IRIS" (연동 출처 구분)
        작성일자=공고일,
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
        사업명=소관부처,
        신청기간=공모유형,
        작성자="",
        작성일자=공고일자,
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
        all_items.extend(items)

        time.sleep(random.uniform(*delay_range))

    return all_items


def _load_dataframe(path: str):
    import pandas as pd

    if not os.path.exists(path):
        return None

    if path.lower().endswith(".xlsx"):
        return pd.read_excel(path, dtype=str)
    return pd.read_csv(path, dtype=str, encoding="utf-8-sig")


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
      (저장된 신청종료일 + 오늘 날짜만 있으면 계산 가능하므로 항상 최신 유지 가능).
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

    if "신청종료일" in merged.columns:
        merged["마감Dday"] = merged["신청종료일"].fillna("").apply(_compute_dday)

    _save_dataframe(merged, db_path)
    return len(merged)


def upload_to_google_sheet(db_path, sheet_name="Gongo"):
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
        if json_creds:
            print("[구글시트] GCP_SA_KEY 환경변수로 인증 시도", file=sys.stderr)
            creds_dict = json.loads(json_creds)
            creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
        elif os.path.exists("service_account.json"):
            print("[구글시트] service_account.json 파일로 인증 시도", file=sys.stderr)
            creds = Credentials.from_service_account_file(
                "service_account.json", scopes=SCOPES
            )
        else:
            print("[구글시트] 인증 정보(GCP_SA_KEY / service_account.json)를 찾을 수 없습니다", file=sys.stderr)
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
        

def main():
    parser = argparse.ArgumentParser(description="전담기관 사업공고 크롤러")
    parser.add_argument(
        "--site", type=str, choices=["nipa", "iris", "nia", "keit", "kiat", "kisa", "smtech", "all"], default="all",
        help="크롤링할 사이트 (기본 all = nipa+nia+keit+kiat+kisa+smtech+iris 한 번에 실행)"
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
    elif args.site == "all":
        items = []
        items += crawl_nipa(args.start, args.end, args.org)
        items += crawl_nia(args.start, args.end)
        items += crawl_keit(args.start, args.end)
        items += crawl_kiat(args.start, args.end)
        items += crawl_kisa(args.start, args.end)
        items += crawl_smtech(args.start, args.end)
        items += crawl_iris(args.start, args.end)

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
            print(f"  - [{it.번호}] {it.제목} ({it.신청기간})", file=sys.stderr)

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
    upload_to_google_sheet(args.db, sheet_name="Gongo")


if __name__ == "__main__":
    main()
