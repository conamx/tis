"""
티스토리 자동 발행 매크로

흐름: 키워드 입력 → 네이버 블로그 검색 → 참고글 선택 → AI 원고 생성 (Claude Code 또는 Gemini)
      → 참고글과 유사도 검사(높으면 재작성) → 카드 이미지 생성/업로드
      → 티스토리 에디터에 글 주입 → 사용자가 직접 발행

API 키 등 설정은 config.json 에 저장됩니다 (앱의 [⚙ 설정] 버튼에서 수정).

원고 작성 AI
  - claude_cli (기본): PC에 설치·로그인된 Claude Code 를 `claude -p` 로 호출.
                       Claude 구독 사용량으로 처리되므로 API 비용 없음 (하루 몇 개 수준 용도)
  - gemini           : Gemini API 키 사용
"""
import warnings
warnings.filterwarnings("ignore")

import base64
import datetime
import html
import json
import os
import queue
import random
import re
import shutil
import subprocess
import sys
import tempfile
import textwrap
import threading
import time
import webbrowser
from urllib.parse import urljoin

import numpy as np
import requests
from bs4 import BeautifulSoup
from PIL import Image, ImageDraw, ImageEnhance, ImageFont

import tkinter as tk
from tkinter import messagebox, ttk


# ==========================================
# [경로]
# ==========================================
def get_resource_path(filename):
    """exe 에 묶인 리소스(템플릿, 폰트, 아이콘) 경로"""
    if hasattr(sys, "_MEIPASS"):
        return os.path.join(sys._MEIPASS, filename)
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), filename)


def get_app_dir():
    """config.json / output / logs 가 저장될 폴더 (exe 옆 또는 스크립트 옆)"""
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


APP_DIR = get_app_dir()
CONFIG_PATH = os.path.join(APP_DIR, "config.json")


# ==========================================
# [설정] — config.json 으로 분리 (코드에 키를 넣지 않음)
# ==========================================
DEFAULT_CONFIG = {
    "accounts": {
        "dndhkdusdl": "https://dndhkdusdl.tistory.com",
        "infopicktolife": "https://infopicktolife.tistory.com",
    },
    "last_account": "dndhkdusdl",
    "last_keywords": "",
    "naver_client_id": "",
    "naver_client_secret": "",
    "gemini_api_key": "",
    "imgbb_api_key": "",
    "gemini_models": ["gemini-2.5-flash", "gemini-2.5-flash-lite"],
    "ai_provider": "claude_cli",      # "claude_cli" 또는 "gemini"
    "claude_model": "",               # 비우면 Claude Code 기본 모델, 예: sonnet / opus
    "claude_cli_path": "",            # 비우면 자동 탐색
    "fallback_to_gemini": True,       # Claude Code 실패 시 Gemini 키가 있으면 Gemini 로 재시도
    "chrome_version_main": None,      # None = 설치된 크롬 버전 자동 감지
    "tistory_write_path": "/manage/post",
    "search_count": 5,                # 네이버 검색 결과 개수
    "manual_select": True,            # 참고글 직접 선택 여부
    "auto_pick_count": 3,             # 자동 선택 시 사용할 상위 N개
    "max_source_chars": 20000,
    "min_body_chars": 2500,
    "max_similarity": 0.12,           # 참고글과 겹치는 비율 상한 (0~1)
    "rewrite_retries": 2,             # 유사도 초과 시 재작성 횟수
    "wait_timeout_min": 30,           # 로그인/발행 대기 시간(분), 0 = 무제한
}

# 환경변수가 있으면 config.json 보다 우선
ENV_OVERRIDES = {
    "naver_client_id": "NAVER_CLIENT_ID",
    "naver_client_secret": "NAVER_CLIENT_SECRET",
    "gemini_api_key": "GEMINI_API_KEY",
    "imgbb_api_key": "IMGBB_API_KEY",
}

KEY_LABELS = {
    "naver_client_id": "네이버 Client ID",
    "naver_client_secret": "네이버 Client Secret",
    "gemini_api_key": "Gemini API 키",
    "imgbb_api_key": "imgbb API 키",
}


def load_config():
    cfg = json.loads(json.dumps(DEFAULT_CONFIG))
    if os.path.exists(CONFIG_PATH):
        try:
            with open(CONFIG_PATH, encoding="utf-8") as f:
                cfg.update(json.load(f))
        except Exception as e:
            print(f"config.json 읽기 실패 (기본값 사용): {e}")
    for key, env in ENV_OVERRIDES.items():
        if os.environ.get(env):
            cfg[key] = os.environ[env]
    return cfg


def save_config(cfg):
    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)


def missing_keys(cfg):
    required = ["naver_client_id", "naver_client_secret", "imgbb_api_key"]
    if cfg.get("ai_provider") == "gemini":
        required.append("gemini_api_key")
    return [KEY_LABELS[k] for k in required if not str(cfg.get(k, "")).strip()]


CONFIG = load_config()

# ==========================================
# [다크모드 색상]
# ==========================================
BG = "#1e1e1e"
BG2 = "#2a2a2a"
BG3 = "#333333"
ACCENT = "#4f9eff"
TEXT = "#e8e8e8"
TEXT_DIM = "#888888"
SUCCESS = "#4caf50"
WARNING = "#ff9800"
ERROR = "#f44336"
BORDER = "#3a3a3a"
FONT = "Malgun Gothic"

# ==========================================
# [5가지 SEO 글쓰기 스타일]
# ==========================================
STYLES = [
    {
        "name": "경험담형",
        "desc": """
- 실제로 써본 사람처럼 1인칭으로. '저는', '제가' 가끔, 일기 느낌은 NO
- 잘 안 알려진 팁이나 실수 경험을 중간에 한두 군데 자연스럽게 삽입
- 문장 끝: '~거든요', '~더라고요', '~수 있어요' 위주, '~입니다'는 30% 이하
- 도입부에서 독자가 공감할 상황 묘사로 시작
""",
    },
    {
        "name": "칼럼형",
        "desc": """
- 잡지 칼럼처럼 서술형 위주. 리스트 남발 금지, 문장으로 풀어가기
- 도입부에서 독자 상황을 콕 집어주고, 문단이 자연스럽게 이어지게
- 문어체와 구어체 6:4 혼합. '~네요', '~죠', '~잖아요' 자연스럽게
- 중간에 반전 포인트나 의외의 사실 한 군데 삽입
""",
    },
    {
        "name": "심화정보형",
        "desc": """
- 기본은 아는 독자 대상으로, 한 단계 더 깊은 내용 위주
- '사실 이건 대부분 잘못 알고 있는데...' 같은 흐름으로 시작
- 문어체 기반이지만 '솔직히', '사실', '근데' 구어 표현 자연스럽게 삽입
- 주의사항이나 예외 케이스를 반드시 한 섹션 포함
""",
    },
    {
        "name": "직관형",
        "desc": """
- 짧고 직관적인 문장 위주. 한 문단 = 핵심 하나
- 수치나 근거 있으면 구체적으로 (예: "3일 만에", "20% 절감"), 없으면 비워두기
- 결론을 각 섹션 첫 문장에 먼저, 이후 이유 설명
- '~입니다/합니다'와 '~해요/예요' 반반 혼합
""",
    },
    {
        "name": "Q&A형",
        "desc": """
- 독자가 자주 묻는 질문 형식으로 구성 (h2 소제목을 의문문으로)
- 각 질문에 핵심 답변을 먼저 한 문장으로, 이후 상세 설명
- 마지막 섹션은 '많이 하는 실수' 또는 '주의할 점'으로 마무리
- 구어체 위주, '~거든요', '~편이에요', '~더라고요' 자연스럽게
""",
    },
]

TITLE_FORMATS = [
    "의문형. 예: '왜 ~가 잘 안 될까?', '~하면 어떻게 될까'",
    "숫자 포함 (단 'N가지 방법' 말고). 예: '~할 때 꼭 확인할 3가지', '직접 해본 ~, N번의 결과'",
    "짧고 직관적 10자 내외. 예: '~의 진짜 문제', '~, 이렇게 쓰세요'",
    "대상 명시. 예: '~초보라면', '~가 고민인 분', '~를 처음 시작하는 분께'",
    "검색 의도 직접 반영. 예: '~하는 법 (직접 해봄)', '~ 완전 정리'",
]

BANNED_PHRASES = [
    "살펴보겠습니다", "알아보도록 하겠습니다", "중요합니다", "바로 시작해볼게요",
    "다양한", "효과적인", "체계적으로", "혁신적인", "탁월한", "놀라운", "최적화된",
    "이상으로 마치겠습니다",
]

_HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                  "(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
}


def strip_tags(text):
    return re.sub(r"<[^>]+>", "", text or "")


# ==========================================
# [네이버 검색]
# ==========================================
def search_naver_blog(kw, count):
    res = requests.get(
        "https://openapi.naver.com/v1/search/blog.json",
        headers={"X-Naver-Client-Id": CONFIG["naver_client_id"],
                 "X-Naver-Client-Secret": CONFIG["naver_client_secret"]},
        params={"query": kw, "display": count, "sort": "sim"},
        timeout=10,
    )
    if res.status_code != 200:
        raise RuntimeError(f"네이버 검색 API 오류 {res.status_code}: {res.text[:120]}")
    return res.json().get("items", [])


# ==========================================
# [소스 수집]
# ==========================================
def _fetch(url):
    res = requests.get(url, headers=_HEADERS, timeout=10)
    res.raise_for_status()
    # 헤더에 charset 이 없으면 requests 가 ISO-8859-1 로 읽어 한글이 깨짐
    if not res.encoding or res.encoding.lower() == "iso-8859-1":
        res.encoding = res.apparent_encoding
    return res


def _clean_text(node):
    for s in node(["script", "style", "nav", "footer", "aside", "header", "form", "noscript"]):
        s.decompose()
    text = node.get_text(separator="\n", strip=True)
    return re.sub(r"\n{3,}", "\n\n", text)


def get_any_web_content(url):
    """URL 에서 본문 텍스트 추출. 실패 시 빈 문자열."""
    try:
        soup = BeautifulSoup(_fetch(url).text, "html.parser")

        if "blog.naver.com" in url:
            iframe = soup.find("iframe", id="mainFrame")
            if iframe and iframe.get("src"):
                soup = BeautifulSoup(_fetch(urljoin("https://blog.naver.com", iframe["src"])).text,
                                     "html.parser")
            # 스마트에디터 ONE / 구 에디터 / 모바일
            for sel in [".se-main-container", "#postViewArea", ".post_ct", "#viewTypeSelector"]:
                content = soup.select_one(sel)
                if content:
                    return _clean_text(content)
            return ""

        # 후보 중 가장 긴 본문을 사용 (첫 매치가 작은 위젯인 경우 방지)
        best = ""
        for sel in ["article", "main", "#content", ".content", ".post_content",
                    ".entry-content", ".tt_article_useless_p_margin", "#article-view-content-div"]:
            for target in soup.select(sel):
                text = _clean_text(target)
                if len(text) > len(best):
                    best = text
        if best:
            return best
        return _clean_text(soup.body) if soup.body else ""
    except Exception as e:
        print(f"스크래핑 실패 {url}: {e}")
        return ""


# ==========================================
# [참고글 유사도 검사] — 원고가 참고글을 베끼지 않았는지 확인
# ==========================================
_NGRAM = 8  # 공백 제거 후 8글자 연속 일치를 '겹침'으로 판단


def _normalize(text):
    return re.sub(r"[^0-9A-Za-z가-힣]", "", strip_tags(text)).lower()


def _shingles(text, n=_NGRAM):
    t = _normalize(text)
    return {t[i:i + n] for i in range(len(t) - n + 1)}


def similarity_report(article_html, source_text):
    """
    returns (ratio, copied_sentences)
      ratio: 원고의 n-gram 중 참고글에도 있는 비율 (0~1)
      copied_sentences: 절반 이상이 참고글과 겹치는 원고 문장 목록
    """
    src = _shingles(source_text)
    art = _shingles(article_html)
    if not art or not src:
        return 0.0, []
    ratio = len(art & src) / len(art)

    plain = strip_tags(re.sub(r"<(br|/p|/li|/h\d|/td|/th)[^>]*>", "\n", article_html))
    sentences = [s.strip() for s in re.split(r"(?<=[.!?。])\s+|\n", plain) if len(s.strip()) >= 15]
    copied = []
    for s in sentences:
        sh = _shingles(s)
        if sh and len(sh & src) / len(sh) >= 0.5:
            copied.append(s)
    return ratio, copied


def find_banned(text):
    plain = strip_tags(text)
    return [p for p in BANNED_PHRASES if p in plain]


# ==========================================
# [AI 원고 작성] — Claude Code CLI / Gemini API
# ==========================================
class AIError(Exception):
    pass


def find_claude_cli():
    """설치된 Claude Code 실행 파일 경로 (없으면 None)"""
    custom = (CONFIG.get("claude_cli_path") or "").strip()
    if custom:
        return custom if os.path.exists(custom) else None
    found = shutil.which("claude")
    if found:
        return found
    home = os.path.expanduser("~")
    for p in [os.path.join(home, ".local", "bin", "claude.exe"),
              os.path.join(home, ".local", "bin", "claude"),
              os.path.join(os.environ.get("APPDATA", ""), "npm", "claude.cmd"),
              os.path.join(home, ".claude", "local", "claude")]:
        if os.path.exists(p):
            return p
    return None


def call_claude_cli(prompt, log=print, stop_flag=None, timeout=900):
    """
    `claude -p` 로 원고 생성. 프롬프트는 stdin 으로 전달(명령줄 길이 제한 회피).
    --tools "" : 파일 읽기/명령 실행 없이 글만 쓰게 함
    """
    exe = find_claude_cli()
    if not exe:
        raise AIError("Claude Code 를 찾지 못했습니다. 설치 후 터미널에서 `claude` 를 한 번 실행해 "
                      "로그인하거나, ⚙ 설정에 실행 파일 경로를 넣어주세요.")
    cmd = [exe, "-p", "--output-format", "json", "--tools", "", "--no-session-persistence"]
    model = (CONFIG.get("claude_model") or "").strip()
    if model:
        cmd += ["--model", model]

    env = os.environ.copy()
    # API 키가 환경변수에 있으면 구독 대신 API 로 과금되므로 제거
    env.pop("ANTHROPIC_API_KEY", None)
    env.pop("ANTHROPIC_AUTH_TOKEN", None)
    flags = getattr(subprocess, "CREATE_NO_WINDOW", 0)  # 윈도우에서 검은 창 안 뜨게

    proc = subprocess.Popen(cmd, stdin=subprocess.PIPE, stdout=subprocess.PIPE,
                            stderr=subprocess.PIPE, text=True, encoding="utf-8",
                            errors="replace", cwd=tempfile.gettempdir(), env=env,
                            creationflags=flags)
    started = time.time()
    pending_input = prompt
    while True:
        try:
            out, err = proc.communicate(input=pending_input, timeout=2)
            break
        except subprocess.TimeoutExpired:
            pending_input = None  # stdin 은 첫 호출에서 이미 전달됨
            if stop_flag and stop_flag.is_set():
                proc.kill()
                proc.communicate()
                raise AIError("사용자 중지")
            if time.time() - started > timeout:
                proc.kill()
                proc.communicate()
                raise AIError(f"Claude Code 응답 시간 초과 ({timeout // 60}분)")

    try:
        data = json.loads(out)
    except (json.JSONDecodeError, TypeError):
        msg = (err or out or "").strip()[:300]
        raise AIError(f"Claude Code 실행 실패 (코드 {proc.returncode}): {msg or '출력 없음'}")

    result = data.get("result") or ""
    if data.get("is_error") or data.get("subtype") != "success" or not result.strip():
        hint = ""
        low = result.lower()
        if "login" in low or "auth" in low:
            hint = " → 터미널에서 `claude` 실행 후 /login 으로 로그인하세요"
        elif "limit" in low:
            hint = " → 구독 사용량 한도에 걸렸습니다. 나중에 다시 시도하세요"
        raise AIError(f"Claude Code 오류: {result[:200] or data.get('subtype')}{hint}")
    return result


def call_ai(prompt, log=print, stop_flag=None):
    """설정된 AI 로 원고 생성. Claude Code 실패 시 Gemini 키가 있으면 Gemini 로 재시도"""
    if CONFIG.get("ai_provider", "claude_cli") == "gemini":
        return call_gemini(prompt, log, stop_flag)
    try:
        return call_claude_cli(prompt, log, stop_flag)
    except AIError as e:
        if str(e) == "사용자 중지":
            raise
        if CONFIG.get("fallback_to_gemini", True) and CONFIG.get("gemini_api_key", "").strip():
            log(f"  ⚠️ {e}", "warning")
            log("  ↪ Gemini 로 재시도합니다", "warning")
            return call_gemini(prompt, log, stop_flag)
        raise


def ai_label():
    if CONFIG.get("ai_provider", "claude_cli") == "gemini":
        return "Gemini"
    model = CONFIG.get("claude_model")
    return f"Claude Code ({model})" if model else "Claude Code"


def call_gemini(prompt, log=print, stop_flag=None):
    models = CONFIG.get("gemini_models") or DEFAULT_CONFIG["gemini_models"]
    last_err = "알 수 없는 오류"
    for model in models:
        url = f"https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent"
        for attempt in range(3):
            if stop_flag and stop_flag.is_set():
                raise AIError("사용자 중지")
            try:
                res = requests.post(
                    url,
                    headers={"x-goog-api-key": CONFIG["gemini_api_key"]},
                    json={"contents": [{"parts": [{"text": prompt}]}],
                          "generationConfig": {"temperature": 1.0}},
                    timeout=180,
                )
            except requests.RequestException as e:
                last_err = f"[{model}] 네트워크 오류: {e}"
                log(f"  {last_err}", "warning")
                time.sleep(5)
                continue

            if res.status_code == 200:
                data = res.json()
                cands = data.get("candidates") or []
                parts = (cands[0].get("content") or {}).get("parts") if cands else None
                if parts:
                    return "".join(p.get("text", "") for p in parts)
                reason = (data.get("promptFeedback") or {}).get("blockReason") \
                    or (cands[0].get("finishReason") if cands else "빈 응답")
                last_err = f"[{model}] 응답 없음 ({reason})"
                log(f"  {last_err}", "warning")
                break  # 같은 모델 재시도해도 결과 같음 → 다음 모델

            last_err = f"[{model}] HTTP {res.status_code}: {res.text[:150]}"
            if res.status_code in (429, 500, 502, 503, 504):
                wait = 10 * (attempt + 1)
                log(f"  {last_err} → {wait}초 후 재시도", "warning")
                time.sleep(wait)
                continue
            log(f"  {last_err}", "warning")
            break  # 400/403/404 등은 재시도 의미 없음
    raise AIError(last_err)


def clean_model_output(text):
    text = text.replace("```html", "").replace("```", "").strip()
    text = re.sub(r"</?(!DOCTYPE|html|head|body)[^>]*>", "", text, flags=re.IGNORECASE)
    text = re.sub(r"^###\s+(.+)$", r"<h3>\1</h3>", text, flags=re.MULTILINE)
    text = re.sub(r"^##\s+(.+)$", r"<h2>\1</h2>", text, flags=re.MULTILINE)
    text = re.sub(r"\*\*(.+?)\*\*", r"<strong>\1</strong>", text)
    return text.strip()


def build_prompt(kw, source_text, style):
    title_format = random.choice(TITLE_FORMATS)
    return f"""블로그 글을 작성해주세요. 아래 조건을 반드시 따르세요.

[주제 키워드] {kw}

[제목]
- 첫 줄에 제목만 텍스트로 (HTML 태그, 따옴표, 접두사 없이)
- 형식: {title_format}
- 키워드 '{kw}'를 자연스럽게 포함

[글쓰기 스타일: {style['name']}]
{style['desc']}

[SEO 필수 조건 — 반드시 지킬 것]
1. 분량: {CONFIG['min_body_chars']:,}자 이상. 부족하면 각 섹션을 더 상세하게 작성
2. 구조: <h2> 소제목 최소 3개 이상, 필요시 <h3> 추가
3. 키워드 '{kw}' → 제목·첫 문단·소제목 1개 이상·본문에 자연스럽게 총 4~6회 포함
4. 도입부 첫 문단: 독자가 이 글을 읽어야 하는 이유 명확히 제시
5. 각 h2 섹션 최소 200자 이상
6. 마지막 섹션: 핵심 요약 또는 실천 방법으로 마무리
7. 본문 중간에 <table> 비교표 또는 정리표를 1개 이상 반드시 포함

[참고자료 활용 규칙 — 가장 중요]
- 참고자료에서는 '사실(수치, 절차, 명칭)'만 가져오고 문장은 100% 새로 쓸 것
- 참고자료의 문장을 옮기거나, 어순·조사만 바꾸거나, 단어만 동의어로 바꾸는 것 금지
- 참고자료와 같은 표현이 8글자 이상 연속으로 겹치지 않게 할 것
- 소제목 구성과 설명 순서도 참고자료와 다르게 새로 짤 것
- 참고자료에 없는 관점(독자 상황별 조언, 흔한 오해, 판단 기준 등)을 최소 2개 추가

[이미지]
- [IMG1], [IMG2], [IMG3]를 본문 앞·중간·뒤에 고르게 단독 줄 배치

[HTML 규칙]
- <h2>, <h3>, <p>, <ul>, <li>, <strong>, <table> 사용
- 마크다운(##, **, __ 등) 절대 금지
- 인사말, 닉네임, 블로그명 절대 금지

[마지막 줄 — 반드시 이 형식]
이미지문구: 문구1 | 문구2 | 문구3
- 각 문구는 본문의 서로 다른 소주제를 반영, 15자 이내

[절대 금지 표현]
- "살펴보겠습니다", "알아보도록 하겠습니다", "중요합니다", "바로 시작해볼게요"
- "다양한", "효과적인", "체계적으로", "혁신적인", "탁월한", "놀라운", "최적화된"
- "지금까지 ~에 대해 알아봤습니다", "이상으로 마치겠습니다"
- 모든 h2 섹션을 같은 분량/구조로 반복

[표 작성 규칙]
- <table>, <thead>, <tbody>, <tr>, <th>, <td> 태그 사용
- 표 위에 <h3>으로 표 제목 달기
- 2열 이상, 3행 이상으로 내용이 있을 때만 사용

[참고자료 — 사실 정보만 참고, 문장은 완전히 새로 쓸 것]
{source_text[:CONFIG['max_source_chars']]}
"""


def build_rewrite_prompt(kw, raw, copied, banned):
    copied_list = "\n".join(f"- {s[:200]}" for s in copied[:40]) or "- (전체적으로 표현이 많이 겹침)"
    banned_list = ", ".join(banned) if banned else "없음"
    return f"""아래 블로그 원고는 참고자료와 표현이 너무 많이 겹칩니다. 다시 써주세요.

[재작성 규칙]
- 정보(사실, 수치)와 HTML 구조, 분량은 유지
- 아래 '겹치는 문장'은 의미만 남기고 어휘·문장 구조·어순을 완전히 바꿀 것
- 겹치는 문장 외의 부분도 문장을 더 자연스럽고 독창적으로 다듬을 것
- 8글자 이상 같은 표현이 연속되지 않게 할 것
- 다음 금지 표현 제거: {banned_list}
- 키워드 '{kw}'는 4~6회 유지
- 첫 줄 제목(태그 없이)도 새로 지을 것, [IMG1]~[IMG3] 자리표시와 마지막 줄 '이미지문구: 문구1 | 문구2 | 문구3' 형식 유지
- 마크다운 금지, 설명 없이 결과물만 출력

[겹치는 문장]
{copied_list}

[원고]
{raw}
"""


def parse_article(raw, kw):
    """모델 출력 → (제목, 본문 HTML, 이미지 문구 3개)"""
    lines = [l.strip() for l in raw.split("\n") if l.strip()]
    if not lines:
        return "", "", [kw, kw, kw]
    title = strip_tags(lines[0])
    title = re.sub(r"^(#+\s*|제목\s*[:：]\s*)", "", title).strip().strip("\"'“”‘’")
    body = "\n".join(lines[1:]).strip()

    img_texts = []
    if "이미지문구" in body:
        # 마지막 줄의 첫 '이미지문구:' 기준으로 분리 ('이미지문구: 이미지문구: ...' 처럼 중복돼도 안전)
        line_start = body.rfind("\n", 0, body.rfind("이미지문구")) + 1
        idx = body.find("이미지문구", line_start)
        tail = body[idx:]
        body = body[:idx]
        tail = re.sub(r"^(이미지문구\s*[:：]\s*)+", "", strip_tags(tail)).strip()
        img_texts = [t.strip() for t in tail.split("|") if t.strip()]
        body = re.sub(r"(<p>\s*)$", "", body.strip()).strip()

    # 제목이 본문 맨 앞에 <h1> 으로 한 번 더 나오는 경우 제거
    body = re.sub(r"^\s*<h1[^>]*>.*?</h1>", "", body, flags=re.DOTALL | re.IGNORECASE).strip()

    while len(img_texts) < 3:
        img_texts.append(kw)
    return title, body, img_texts[:3]


def ensure_img_placeholders(body):
    """모델이 [IMGn] 을 빠뜨렸을 때 h2 경계에 고르게 끼워넣기"""
    missing = [n for n in (1, 2, 3) if f"[IMG{n}]" not in body]
    if not missing:
        return body
    h2_pos = [m.start() for m in re.finditer(r"<h2", body)]
    slots = []
    if h2_pos:
        slots = [h2_pos[0], h2_pos[len(h2_pos) // 2], len(body)]
    else:
        slots = [0, len(body) // 2, len(body)]
    # 뒤에서부터 넣어야 앞 위치가 안 밀림
    for n in sorted(missing, reverse=True):
        pos = slots[n - 1]
        body = body[:pos] + f"\n[IMG{n}]\n" + body[pos:]
    return body


# ==========================================
# [카드 이미지 생성 + imgbb 업로드]
# ==========================================
FONT_FILES = ["NanumSquareNeo-cBd.ttf", "NanumSquareNeo-dEb.ttf", "NanumSquareNeo-eHv.ttf"]
TEMPLATE_FILES = ["template.png", "template_2.png"]
# 템플릿별 글자 세로 중심 (1280x720 기준) — 흰 영역 가운데에 오도록. 목록에 없으면 360
TEMPLATE_TEXT_CENTER = {"template.png": 300, "template_2.png": 320}
FALLBACK_FONTS = ["C:/Windows/Fonts/malgunbd.ttf", "C:/Windows/Fonts/malgun.ttf",
                  "/System/Library/Fonts/AppleSDGothicNeo.ttc"]


def _load_font(size):
    fonts = [get_resource_path(f) for f in FONT_FILES if os.path.exists(get_resource_path(f))]
    candidates = ([random.choice(fonts)] if fonts else []) + FALLBACK_FONTS
    for path in candidates:
        try:
            return ImageFont.truetype(path, size), True
        except Exception:
            continue
    return ImageFont.load_default(), False  # 한글이 □ 로 나올 수 있음


def create_card(text, save_path):
    """카드 이미지를 만들어 save_path 에 저장. returns 경고 메시지 목록"""
    warns = []
    templates = [f for f in TEMPLATE_FILES if os.path.exists(get_resource_path(f))]
    center_y = 360
    if templates:
        name = random.choice(templates)
        center_y = TEMPLATE_TEXT_CENTER.get(name, 360)
        img = Image.open(get_resource_path(name)).convert("RGB").resize((1280, 720), Image.LANCZOS)
    else:
        img = Image.new("RGB", (1280, 720), color="white")
        warns.append("템플릿 이미지 없음 → 흰 배경 사용")

    # 살짝 확대 후 크롭 (축소하면 가장자리에 검은 띠가 생기므로 확대만)
    scale = random.uniform(1.0, 1.05)
    w, h = int(1280 * scale), int(720 * scale)
    img = img.resize((w, h), Image.LANCZOS)
    left, top = random.randint(0, w - 1280), random.randint(0, h - 720)
    img = img.crop((left, top, left + 1280, top + 720))

    img = ImageEnhance.Brightness(img).enhance(random.uniform(0.95, 1.05))
    img = ImageEnhance.Contrast(img).enhance(random.uniform(0.95, 1.05))

    arr = np.array(img, dtype=np.int16)
    noise = np.random.randint(-3, 4, arr.shape, dtype=np.int16)
    img = Image.fromarray(np.clip(arr + noise, 0, 255).astype(np.uint8))

    draw = ImageDraw.Draw(img)
    font_size = random.randint(72, 88)
    font, ok = _load_font(font_size)
    if not ok:
        warns.append("한글 폰트 없음 → 글자가 깨질 수 있음")

    c = random.randint(0, 15)
    lines = textwrap.wrap(strip_tags(text).strip(), width=12)[:4]
    line_h = font_size + 20
    y = center_y - line_h * len(lines) / 2 + random.randint(-15, 15)
    for line in lines:
        bbox = draw.textbbox((0, 0), line, font=font)
        x = (1280 - (bbox[2] - bbox[0])) / 2 + random.randint(-5, 5)
        draw.text((x, y), line, fill=(c, c, c), font=font)
        y += line_h

    img.save(save_path, quality=random.randint(88, 96))
    return warns


def upload_imgbb(path, name):
    with open(path, "rb") as f:
        img_b64 = base64.b64encode(f.read()).decode("utf-8")
    last = ""
    for _ in range(2):
        try:
            r = requests.post("https://api.imgbb.com/1/upload",
                              data={"key": CONFIG["imgbb_api_key"], "image": img_b64, "name": name},
                              timeout=30)
            if r.status_code == 200 and r.json().get("success"):
                return r.json()["data"]["url"], ""
            last = f"HTTP {r.status_code}: {r.text[:120]}"
        except Exception as e:
            last = str(e)
        time.sleep(2)
    return "", last


# ==========================================
# [셀레늄]
# ==========================================
def setup_driver():
    import undetected_chromedriver as uc
    options = uc.ChromeOptions()
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--window-size=1400,1000")
    options.add_argument("--lang=ko-KR")
    version = CONFIG.get("chrome_version_main")
    return uc.Chrome(options=options, version_main=int(version) if version else None)


def driver_alive(driver):
    if driver is None:
        return False
    try:
        _ = driver.current_url
        return True
    except Exception:
        return False


TITLE_SELECTORS = [".textarea_tit", "#post-title-inp", "textarea[placeholder*='제목']"]


def _dismiss_alerts(driver):
    """'작성 중인 글이 있습니다' 같은 알림창 닫기"""
    for _ in range(3):
        try:
            driver.switch_to.alert.dismiss()
            time.sleep(0.5)
        except Exception:
            return


def _bmp_only(text):
    # chromedriver send_keys 는 이모지 등 BMP 밖 문자를 입력하지 못함
    return "".join(ch for ch in text if ord(ch) <= 0xFFFF)


def inject_post(driver, base_url, title, body):
    """티스토리 글쓰기 화면에 제목/본문 입력. returns (성공여부, 메시지)"""
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support import expected_conditions as EC
    from selenium.webdriver.support.ui import WebDriverWait

    try:  # 이전 글 화면의 '페이지를 나가시겠습니까?' 확인창
        driver.switch_to.alert.accept()
    except Exception:
        pass
    driver.get(base_url.rstrip("/") + CONFIG.get("tistory_write_path", "/manage/post"))
    time.sleep(1)
    _dismiss_alerts(driver)  # '작성 중인 글 불러오기' 는 취소

    title_el = None
    for sel in TITLE_SELECTORS:
        try:
            title_el = WebDriverWait(driver, 8 if sel == TITLE_SELECTORS[0] else 2).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, sel)))
            break
        except Exception:
            _dismiss_alerts(driver)
    if title_el is None:
        return False, "제목 입력칸을 찾지 못했습니다 (로그인 상태/페이지 확인)"

    title_el.click()
    title_el.clear()
    title_el.send_keys(_bmp_only(title))
    time.sleep(0.5)

    # 1순위: TinyMCE API 로 넣기 (에디터 내부 상태와 동기화됨)
    try:
        WebDriverWait(driver, 10).until(lambda d: d.execute_script(
            "return !!(window.tinymce && tinymce.activeEditor && tinymce.activeEditor.initialized)"))
        ok = driver.execute_script("""
            var ed = tinymce.activeEditor;
            ed.setContent(arguments[0]);
            ed.fire('change'); ed.fire('input'); ed.save();
            return ed.getContent().length > 0;
        """, body)
        if ok:
            return True, ""
    except Exception:
        pass

    # 2순위: iframe 직접 주입 (기존 방식)
    try:
        driver.switch_to.frame("editor-tistory_ifr")
        editor_body = driver.find_element(By.ID, "tinymce")
        driver.execute_script("arguments[0].innerHTML = arguments[1]", editor_body, body)
        return True, ""
    except Exception as e:
        return False, f"본문 주입 실패: {e.__class__.__name__}"
    finally:
        driver.switch_to.default_content()


# ==========================================
# [결과 저장]
# ==========================================
def safe_filename(name, limit=40):
    return re.sub(r'[\\/:*?"<>|\s]+', "_", name).strip("_")[:limit] or "untitled"


def save_article_html(folder, idx, kw, title, body):
    path = os.path.join(folder, f"{idx:02d}_{safe_filename(kw)}.html")
    with open(path, "w", encoding="utf-8") as f:
        f.write(f"<!doctype html><meta charset='utf-8'><title>{html.escape(title)}</title>\n"
                f"<h1>{html.escape(title)}</h1>\n{body}\n")
    return path


def open_folder(path):
    try:
        if sys.platform.startswith("win"):
            os.startfile(path)
        elif sys.platform == "darwin":
            subprocess.Popen(["open", path])
        else:
            subprocess.Popen(["xdg-open", path])
    except Exception:
        pass


# ==========================================
# [GUI 앱]
# ==========================================
class StopRequested(Exception):
    pass


class MacroApp:
    def __init__(self, root):
        self.root = root
        self.root.title("티스토리 자동 발행 매크로")
        self.root.geometry("980x740")
        self.root.minsize(860, 640)
        self.root.configure(bg=BG)
        try:
            from PIL import ImageTk
            self._icon_img = ImageTk.PhotoImage(Image.open(get_resource_path("icon.png")))
            self.root.iconphoto(True, self._icon_img)
        except Exception:
            pass

        self.running = False
        self.driver = None
        self.stop_flag = threading.Event()
        self.login_event = threading.Event()
        self.next_event = threading.Event()
        self._ui_queue = queue.Queue()
        self._log_file = None
        self._out_dir = None

        self._setup_style()
        self._build_ui()
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        self.root.bind("<Control-Return>", lambda e: self._run())
        self.root.after(50, self._poll_ui_queue)

        if missing_keys(CONFIG):
            self.root.after(300, lambda: self._open_settings(first_run=True))

    # ── 스레드 안전 UI 호출 ──
    def ui(self, fn, *args, **kwargs):
        """백그라운드 스레드에서 UI 를 건드릴 때는 항상 이걸로"""
        self._ui_queue.put((fn, args, kwargs))

    def _poll_ui_queue(self):
        try:
            while True:
                fn, args, kwargs = self._ui_queue.get_nowait()
                try:
                    fn(*args, **kwargs)
                except Exception as e:
                    print(f"UI 오류: {e}")
        except queue.Empty:
            pass
        self.root.after(50, self._poll_ui_queue)

    def log(self, msg, tag=""):
        if threading.current_thread() is not threading.main_thread():
            self.ui(self.log, msg, tag)
            return
        ts = datetime.datetime.now().strftime("%H:%M:%S")
        self.log_box.config(state="normal")
        self.log_box.insert("end", f"[{ts}] ", "dim")
        self.log_box.insert("end", msg + "\n", tag)
        self.log_box.see("end")
        self.log_box.config(state="disabled")
        if self._log_file:
            try:
                with open(self._log_file, "a", encoding="utf-8") as f:
                    f.write(f"[{ts}] {msg}\n")
            except Exception:
                pass

    def set_progress(self, text, value=None, maximum=None):
        if threading.current_thread() is not threading.main_thread():
            self.ui(self.set_progress, text, value, maximum)
            return
        self.progress_var.set(text)
        if maximum is not None:
            self.progress_bar.configure(maximum=maximum)
        if value is not None:
            self.progress_bar.configure(value=value)

    # ── UI 구성 ──
    def _setup_style(self):
        style = ttk.Style(self.root)
        try:
            style.theme_use("clam")
        except Exception:
            pass
        style.configure("Dark.Horizontal.TProgressbar", troughcolor=BG2, background=ACCENT,
                        bordercolor=BORDER, lightcolor=ACCENT, darkcolor=ACCENT)

    def _btn(self, parent, text, cmd, bg=ACCENT, fg="white", size=12, bold=True, pady=12):
        return tk.Button(parent, text=text, font=(FONT, size, "bold" if bold else "normal"),
                         bg=bg, fg=fg, activebackground=bg, activeforeground=fg,
                         relief="flat", bd=0, pady=pady, cursor="hand2", command=cmd)

    def _build_ui(self):
        title_bar = tk.Frame(self.root, bg=BG, pady=14)
        title_bar.pack(fill="x", padx=24)
        tk.Label(title_bar, text="티스토리 자동 발행", font=(FONT, 18, "bold"),
                 bg=BG, fg=TEXT).pack(side="left")
        self._btn(title_bar, "📁 결과 폴더", self._open_output, bg=BG3, fg=TEXT,
                  size=9, bold=False, pady=4).pack(side="right", padx=(6, 0), ipadx=8)
        self._btn(title_bar, "⚙ 설정", self._open_settings, bg=BG3, fg=TEXT,
                  size=9, bold=False, pady=4).pack(side="right", ipadx=8)
        tk.Frame(self.root, bg=BORDER, height=1).pack(fill="x", padx=24)

        body = tk.Frame(self.root, bg=BG)
        body.pack(fill="both", expand=True, padx=24, pady=14)
        left = tk.Frame(body, bg=BG, width=390)
        left.pack(side="left", fill="y", padx=(0, 12))
        left.pack_propagate(False)
        right = tk.Frame(body, bg=BG)
        right.pack(side="left", fill="both", expand=True)

        self._build_left(left)
        self._build_right(right)

    def _build_left(self, parent):
        # 하단 고정 영역 (진행상황 + 버튼) — 화면 전환과 무관하게 항상 표시
        bottom = tk.Frame(parent, bg=BG)
        bottom.pack(side="bottom", fill="x")

        self.progress_var = tk.StringVar(value="대기 중")
        tk.Label(bottom, textvariable=self.progress_var, font=(FONT, 9), bg=BG, fg=ACCENT,
                 wraplength=370, justify="left").pack(anchor="w", pady=(8, 4))
        self.progress_bar = ttk.Progressbar(bottom, style="Dark.Horizontal.TProgressbar",
                                            mode="determinate")
        self.progress_bar.pack(fill="x", pady=(0, 10))

        # 실행 버튼 / 발행 버튼이 번갈아 들어가는 자리
        btn_slot = tk.Frame(bottom, bg=BG)
        btn_slot.pack(fill="x")
        self.publish_btn = self._btn(btn_slot, "✅  로그인 완료", self._on_publish_btn, bg=SUCCESS)
        self.run_btn = self._btn(btn_slot, "▶   실행 시작  (Ctrl+Enter)", self._run)
        self.run_btn.pack(fill="x")
        self.stop_btn = self._btn(bottom, "■  중지", self._stop, bg=BG3, fg=TEXT,
                                  size=9, bold=False, pady=6)

        # ── 키워드 입력 화면 ──
        self._kw_screen = tk.Frame(parent, bg=BG)
        self._kw_screen.pack(fill="both", expand=True)

        tk.Label(self._kw_screen, text="티스토리 계정", font=(FONT, 11, "bold"),
                 bg=BG, fg=TEXT).pack(anchor="w", pady=(0, 6))
        self._account_var = tk.StringVar()
        self._acc_frame = tk.Frame(self._kw_screen, bg=BG2, padx=10, pady=6)
        self._acc_frame.pack(fill="x", pady=(0, 10))
        self._refresh_accounts()

        opt = tk.Frame(self._kw_screen, bg=BG)
        opt.pack(fill="x", pady=(0, 10))
        self._manual_var = tk.BooleanVar(value=bool(CONFIG.get("manual_select", True)))
        tk.Checkbutton(opt, text="참고글 직접 선택", variable=self._manual_var,
                       font=(FONT, 9), bg=BG, fg=TEXT, selectcolor=BG3,
                       activebackground=BG, activeforeground=TEXT).pack(side="left")
        tk.Label(opt, text="(끄면 상위", font=(FONT, 8), bg=BG, fg=TEXT_DIM).pack(side="left")
        self._auto_n = tk.Spinbox(opt, from_=1, to=10, width=3, font=(FONT, 9),
                                  bg=BG2, fg=TEXT, buttonbackground=BG3, relief="flat")
        self._auto_n.delete(0, "end")
        self._auto_n.insert(0, str(CONFIG.get("auto_pick_count", 3)))
        self._auto_n.pack(side="left", padx=2)
        tk.Label(opt, text="개 자동 사용)", font=(FONT, 8), bg=BG, fg=TEXT_DIM).pack(side="left")

        hdr = tk.Frame(self._kw_screen, bg=BG)
        hdr.pack(fill="x")
        tk.Label(hdr, text="키워드 목록", font=(FONT, 11, "bold"), bg=BG, fg=TEXT).pack(side="left")
        self._kw_count = tk.Label(hdr, text="0개", font=(FONT, 9), bg=BG, fg=TEXT_DIM)
        self._kw_count.pack(side="right")
        tk.Label(self._kw_screen, text="한 줄에 하나씩 입력 (중복은 자동 제거)",
                 font=(FONT, 8), bg=BG, fg=TEXT_DIM).pack(anchor="w", pady=(0, 6))

        kw_frame = tk.Frame(self._kw_screen, bg=BORDER, padx=1, pady=1)
        kw_frame.pack(fill="both", expand=True)
        self.kw_text = tk.Text(kw_frame, font=(FONT, 10), bg=BG2, fg=TEXT, insertbackground=TEXT,
                               relief="flat", bd=6, wrap="none", undo=True)
        kw_scroll = tk.Scrollbar(kw_frame, command=self.kw_text.yview, bg=BG3)
        self.kw_text.config(yscrollcommand=kw_scroll.set)
        kw_scroll.pack(side="right", fill="y")
        self.kw_text.pack(fill="both", expand=True)
        self.kw_text.insert("1.0", CONFIG.get("last_keywords", ""))
        self.kw_text.bind("<<Modified>>", self._on_kw_modified)
        self._update_kw_count()

        # ── 검색결과 선택 화면 (기본 숨김) ──
        self._sel_screen = tk.Frame(parent, bg=BG)
        sel_hdr = tk.Frame(self._sel_screen, bg=BG)
        sel_hdr.pack(fill="x", pady=(0, 4))
        self._sel_title = tk.Label(sel_hdr, text="", font=(FONT, 11, "bold"), bg=BG, fg=TEXT,
                                   wraplength=300, justify="left")
        self._sel_title.pack(side="left")
        self._sel_step = tk.Label(sel_hdr, text="", font=(FONT, 9), bg=BG, fg=TEXT_DIM)
        self._sel_step.pack(side="right")
        tk.Label(self._sel_screen, text="참고할 글을 선택하세요 (복수 선택, '열기'로 미리보기)",
                 font=(FONT, 8), bg=BG, fg=TEXT_DIM).pack(anchor="w", pady=(0, 6))

        sel_list_wrap = tk.Frame(self._sel_screen, bg=BORDER, padx=1, pady=1)
        sel_list_wrap.pack(fill="both", expand=True)
        self._sel_list = tk.Frame(sel_list_wrap, bg=BG2)
        self._sel_list.pack(fill="both", expand=True, padx=6, pady=6)

        tk.Label(self._sel_screen, text="직접 URL 추가 (여러 개는 공백으로 구분)",
                 font=(FONT, 8), bg=BG, fg=TEXT_DIM).pack(anchor="w", pady=(8, 2))
        self._extra_url = tk.Entry(self._sel_screen, font=(FONT, 9), bg=BG2, fg=TEXT,
                                   insertbackground=TEXT, relief="flat")
        self._extra_url.pack(fill="x", ipady=4)

        self._sel_next_btn = self._btn(self._sel_screen, "다음 →", self._sel_next)
        self._sel_next_btn.pack(fill="x", pady=(10, 0))
        row = tk.Frame(self._sel_screen, bg=BG)
        row.pack(fill="x", pady=(6, 0))
        self._btn(row, "이 키워드 건너뛰기", self._sel_skip, bg=BG3, fg=TEXT_DIM,
                  size=8, bold=False, pady=6).pack(side="left", fill="x", expand=True, padx=(0, 3))
        self._btn(row, "전체 취소", self._stop, bg=BG3, fg=ERROR,
                  size=8, bold=False, pady=6).pack(side="left", fill="x", expand=True, padx=(3, 0))
        self.root.bind("<Return>", self._on_enter)

    def _build_right(self, parent):
        tk.Label(parent, text="■  실행 로그", font=(FONT, 9, "bold"),
                 bg=BG, fg=TEXT_DIM).pack(anchor="w", pady=(4, 4))
        log_frame = tk.Frame(parent, bg=BORDER, padx=1, pady=1)
        log_frame.pack(fill="both", expand=True)
        self.log_box = tk.Text(log_frame, font=("Consolas", 9), bg=BG2, fg=TEXT,
                               insertbackground=TEXT, relief="flat", bd=6,
                               wrap="word", state="disabled")
        scroll = tk.Scrollbar(log_frame, command=self.log_box.yview, bg=BG3)
        self.log_box.config(yscrollcommand=scroll.set)
        scroll.pack(side="right", fill="y")
        self.log_box.pack(fill="both", expand=True)
        for tag, color in [("success", SUCCESS), ("warning", WARNING), ("error", ERROR),
                           ("info", ACCENT), ("dim", TEXT_DIM)]:
            self.log_box.tag_config(tag, foreground=color)
        self._btn(parent, "로그 지우기", self._clear_log, bg=BG3, fg=TEXT_DIM,
                  size=8, bold=False, pady=4).pack(anchor="e", pady=(4, 0), ipadx=8)

    def _refresh_accounts(self):
        for w in self._acc_frame.winfo_children():
            w.destroy()
        accounts = CONFIG.get("accounts") or {}
        last = CONFIG.get("last_account")
        self._account_var.set(last if last in accounts else next(iter(accounts), ""))
        if not accounts:
            tk.Label(self._acc_frame, text="계정 없음 — ⚙ 설정에서 추가하세요",
                     font=(FONT, 9), bg=BG2, fg=ERROR).pack(anchor="w")
        for key, url in accounts.items():
            tk.Radiobutton(self._acc_frame, text=f"{key}   ", variable=self._account_var, value=key,
                           font=(FONT, 10), bg=BG2, fg=TEXT, selectcolor=BG3,
                           activebackground=BG2, activeforeground=TEXT).pack(anchor="w")

    # ── 작은 헬퍼 ──
    def _clear_log(self):
        self.log_box.config(state="normal")
        self.log_box.delete("1.0", "end")
        self.log_box.config(state="disabled")

    def _get_keywords(self):
        seen, result = set(), []
        for k in self.kw_text.get("1.0", "end").splitlines():
            k = k.strip()
            if k and k not in seen:
                seen.add(k)
                result.append(k)
        return result

    def _update_kw_count(self):
        self._kw_count.configure(text=f"{len(self._get_keywords())}개")

    def _on_kw_modified(self, _):
        self._update_kw_count()
        self.kw_text.edit_modified(False)

    def _on_enter(self, event):
        if self._sel_screen.winfo_ismapped() and event.widget is not self.kw_text:
            self._sel_next()

    def _open_output(self):
        path = self._out_dir or os.path.join(APP_DIR, "output")
        os.makedirs(path, exist_ok=True)
        open_folder(path)

    def _set_running_ui(self, running):
        if running:
            self.run_btn.config(state="disabled", text="⏳  실행 중...", bg=BG3)
            self.stop_btn.pack(fill="x", pady=(6, 0))
        else:
            self.run_btn.config(state="normal", text="▶   실행 시작  (Ctrl+Enter)", bg=ACCENT)
            self.stop_btn.pack_forget()
            self.publish_btn.pack_forget()
            self.run_btn.pack(fill="x")
            self._show_kw_screen()

    def _show_publish_btn(self, text):
        if self.stop_flag.is_set():
            return
        self.publish_btn.configure(text=text, state="normal", bg=SUCCESS)
        self.run_btn.pack_forget()
        self.publish_btn.pack(fill="x")

    def _show_kw_screen(self):
        self._sel_screen.pack_forget()
        self._kw_screen.pack(fill="both", expand=True)

    def _wait(self, event):
        """event 가 설정될 때까지 대기. 중지/시간초과 시 StopRequested"""
        timeout_min = CONFIG.get("wait_timeout_min", 30) or 0
        deadline = time.time() + timeout_min * 60 if timeout_min else None
        while not event.is_set():
            if self.stop_flag.is_set():
                raise StopRequested()
            if deadline and time.time() > deadline:
                raise StopRequested(f"{timeout_min}분 동안 응답이 없어 중단합니다")
            event.wait(0.3)
        if self.stop_flag.is_set():  # 중지 버튼도 event 를 깨우므로 한 번 더 확인
            raise StopRequested()
        event.clear()

    def _check_stop(self):
        if self.stop_flag.is_set():
            raise StopRequested()

    # ── 버튼 콜백 ──
    def _on_publish_btn(self):
        if not self.login_event.is_set() and not getattr(self, "_logged_in", False):
            self._logged_in = True
            self.login_event.set()
            self.publish_btn.configure(text="⏳  원고 생성 기다리는 중...", state="disabled", bg=BG3)
            self.log("🔓 로그인 완료 확인", "success")
        else:
            self.next_event.set()
            self.publish_btn.configure(text="⏳  주입 중...", state="disabled", bg=BG3)

    def _stop(self):
        if not self.running:
            return
        self.stop_flag.set()
        self.login_event.set()
        self.next_event.set()
        self.log("■ 중지 요청 — 현재 작업이 끝나는 대로 멈춥니다", "warning")
        if self._sel_screen.winfo_ismapped():
            self._finish_run()

    def _on_close(self):
        if self.running and not messagebox.askokcancel("종료", "작업 중입니다. 종료할까요?"):
            return
        self._remember_inputs()
        self.stop_flag.set()
        if self.driver is not None:
            try:
                self.driver.quit()
            except Exception:
                pass
        self.root.destroy()

    def _remember_inputs(self):
        CONFIG["last_keywords"] = self.kw_text.get("1.0", "end").strip()
        CONFIG["last_account"] = self._account_var.get()
        CONFIG["manual_select"] = bool(self._manual_var.get())
        try:
            CONFIG["auto_pick_count"] = max(1, int(self._auto_n.get()))
        except ValueError:
            pass
        try:
            save_config(CONFIG)
        except Exception as e:
            self.log(f"⚠️ 설정 저장 실패: {e}", "warning")

    # ── 실행 흐름 ──
    def _run(self):
        if self.running:
            return
        miss = missing_keys(CONFIG)
        if miss:
            self.log(f"❌ 설정 필요: {', '.join(miss)}", "error")
            self._open_settings()
            return
        keywords = self._get_keywords()
        if not keywords:
            self.log("❌ 키워드를 입력해주세요", "error")
            return
        account = self._account_var.get()
        if account not in (CONFIG.get("accounts") or {}):
            self.log("❌ 티스토리 계정을 선택해주세요", "error")
            return
        self._remember_inputs()

        self.running = True
        self.stop_flag.clear()
        self.login_event.clear()
        self.next_event.clear()
        self._logged_in = False
        self._keywords = keywords
        self._account = account
        self._base_url = CONFIG["accounts"][account]
        self._sel_results = []
        self._search_results = {}
        self._warned = set()
        # 백그라운드 스레드에서 tk 변수를 읽지 않도록 미리 복사
        self._manual = bool(self._manual_var.get())
        self._auto_count = int(CONFIG.get("auto_pick_count", 3))

        stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        self._out_dir = os.path.join(APP_DIR, "output", f"{stamp}_{safe_filename(account, 20)}")
        os.makedirs(self._out_dir, exist_ok=True)
        os.makedirs(os.path.join(APP_DIR, "logs"), exist_ok=True)
        self._log_file = os.path.join(APP_DIR, "logs", f"{stamp}.log")

        self._set_running_ui(True)
        self.log(f"📌 계정: {account} ({self._base_url}) / 키워드 {len(keywords)}개", "info")
        self.set_progress("검색 중...", 0, len(keywords))
        threading.Thread(target=self._search_all, daemon=True).start()

    def _search_all(self):
        """백그라운드: 모든 키워드를 미리 검색해 선택 화면을 바로바로 넘길 수 있게"""
        total = len(self._keywords)
        for idx, kw in enumerate(self._keywords):
            if self.stop_flag.is_set():
                return
            self.set_progress(f"검색 중 [{idx + 1}/{total}] {kw}", idx)
            try:
                items = search_naver_blog(kw, int(CONFIG.get("search_count", 5)))
                self.log(f"🔍 {kw}: 검색결과 {len(items)}개", "dim")
            except Exception as e:
                items = []
                self.log(f"⚠️ [{kw}] 검색 실패: {e}", "warning")
            self._search_results[idx] = items
            if idx == 0 and self._manual:
                self.ui(self._show_sel_screen_ui, 0)

        if not self._manual:
            n = self._auto_count
            self._sel_results = [[it["link"] for it in self._search_results.get(i, [])[:n]]
                                 for i in range(total)]
            self.log(f"✅ 상위 {n}개 글 자동 선택", "success")
            self.ui(self._start_generation)

    def _show_sel_screen_ui(self, idx):
        if self.stop_flag.is_set():
            return
        kw = self._keywords[idx]
        total = len(self._keywords)
        items = self._search_results.get(idx)
        if items is None:  # 아직 검색 안 끝남
            self.root.after(200, lambda: self._show_sel_screen_ui(idx))
            return

        self._kw_screen.pack_forget()
        self._sel_screen.pack(fill="both", expand=True)
        self._sel_title.configure(text=f"🔍  {kw}")
        self._sel_step.configure(text=f"[{idx + 1}/{total}]")
        self.set_progress(f"참고글 선택 [{idx + 1}/{total}] {kw}", idx)
        self._extra_url.delete(0, "end")

        for w in self._sel_list.winfo_children():
            w.destroy()
        self._cur_check_vars = []
        if not items:
            tk.Label(self._sel_list, text="검색 결과 없음 — 아래에 URL을 직접 넣거나 건너뛰세요",
                     font=(FONT, 9), bg=BG2, fg=ERROR, wraplength=320).pack(anchor="w", pady=4)
        for i, item in enumerate(items):
            title = html.unescape(strip_tags(item["title"]))
            desc = html.unescape(strip_tags(item.get("description", "")))[:70]
            date = item.get("postdate", "")
            date = f"{date[:4]}.{date[4:6]}.{date[6:]}" if len(date) == 8 else ""
            var = tk.BooleanVar(value=(i == 0))
            row = tk.Frame(self._sel_list, bg=BG2)
            row.pack(fill="x", pady=3)
            top = tk.Frame(row, bg=BG2)
            top.pack(fill="x")
            tk.Checkbutton(top, text=title[:40] + ("…" if len(title) > 40 else ""), variable=var,
                           font=(FONT, 9, "bold"), bg=BG2, fg=TEXT, selectcolor=BG3,
                           activebackground=BG2, activeforeground=TEXT,
                           anchor="w", wraplength=280, justify="left").pack(side="left", anchor="w")
            link = item["link"]
            tk.Button(top, text="열기", font=(FONT, 7), bg=BG3, fg=TEXT_DIM, relief="flat", bd=0,
                      cursor="hand2", command=lambda u=link: webbrowser.open(u)).pack(side="right")
            meta = " · ".join(x for x in [item.get("bloggername", ""), date] if x)
            if desc or meta:
                tk.Label(row, text=(meta + "\n" if meta else "") + desc, font=(FONT, 8),
                         bg=BG2, fg=TEXT_DIM, wraplength=320, justify="left",
                         anchor="w").pack(anchor="w", padx=(22, 0))
            self._cur_check_vars.append((var, item))

        if idx + 1 == total:
            self._sel_next_btn.configure(text="✅  선택 완료 → 원고 생성 시작  (Enter)")
        else:
            self._sel_next_btn.configure(text=f"다음 →  ({idx + 2}/{total}번째)  (Enter)")

    def _sel_next(self):
        if not self._sel_screen.winfo_ismapped():
            return
        urls = [item["link"] for var, item in self._cur_check_vars if var.get()]
        urls += [u for u in self._extra_url.get().split() if u.startswith("http")]
        self._sel_results.append(urls)
        self._advance_sel()

    def _sel_skip(self):
        self._sel_results.append([])
        self._advance_sel()

    def _advance_sel(self):
        idx = len(self._sel_results)
        if idx < len(self._keywords):
            self._show_sel_screen_ui(idx)
        else:
            self._show_kw_screen()
            self.log("✅ 참고 글 선택 완료! 원고 생성 시작합니다.", "success")
            self._start_generation()

    def _start_generation(self):
        if self.stop_flag.is_set():
            return
        self._show_kw_screen()
        threading.Thread(target=self._generate_and_publish, daemon=True).start()

    def _finish_run(self):
        self.running = False
        self._set_running_ui(False)

    # ── 원고 생성 ──
    def _generate_article(self, idx, kw, urls):
        """키워드 하나 처리. returns dict 또는 None"""
        style = random.choice(STYLES)
        self.log(f"📝 스타일: {style['name']}", "info")

        source_text = ""
        for url in urls:
            self._check_stop()
            c = get_any_web_content(url)
            mark = "✓" if len(c) > 200 else "✗ (본문 추출 실패)"
            self.log(f"  스크래핑 {mark} {len(c):,}자  {url}", "dim")
            if len(c) > 200:
                source_text += f"\n--- 참고 ---\n{c}\n"
            time.sleep(random.uniform(0.5, 1.0))

        if not source_text:
            self.log(f"⚠️ [{kw}] 참고할 본문이 없어 건너뜀", "warning")
            return None

        self.log(f"🤖 {ai_label()} 원고 생성 중... (1~3분 걸릴 수 있어요)", "info")
        try:
            raw = clean_model_output(call_ai(build_prompt(kw, source_text, style),
                                                 self.log, self.stop_flag))
        except AIError as e:
            self.log(f"❌ [{kw}] 생성 실패: {e}", "error")
            return None

        # 참고글과 겹치는 표현이 많으면 재작성
        max_sim = float(CONFIG.get("max_similarity", 0.12))
        retries = int(CONFIG.get("rewrite_retries", 2))
        best_raw, best_ratio = raw, None
        for attempt in range(retries + 1):
            _, body, _ = parse_article(raw, kw)
            ratio, copied = similarity_report(body, source_text)
            banned = find_banned(body)
            if best_ratio is None or ratio < best_ratio:
                best_raw, best_ratio = raw, ratio
            msg = f"  🔎 참고글 겹침 {ratio:.1%} (기준 {max_sim:.0%}) · 겹친 문장 {len(copied)}개"
            if banned:
                msg += f" · 금지표현 {', '.join(banned)}"
            self.log(msg, "success" if ratio <= max_sim and not copied else "warning")
            if (ratio <= max_sim and not copied) or attempt == retries:
                break
            self._check_stop()
            self.log(f"  ✍️ 원고 재작성 중... ({attempt + 1}/{retries})", "info")
            try:
                raw = clean_model_output(call_ai(build_rewrite_prompt(kw, raw, copied, banned),
                                                     self.log, self.stop_flag))
            except AIError as e:
                self.log(f"  ⚠️ 재작성 실패, 이전 원고 사용: {e}", "warning")
                break

        title, body, img_texts = parse_article(best_raw, kw)
        if not title or len(strip_tags(body)) < 300:
            self.log(f"❌ [{kw}] 원고 형식이 이상해서 건너뜀", "error")
            return None
        if best_ratio is not None and best_ratio > max_sim:
            self.log(f"  ⚠️ 겹침 {best_ratio:.1%} — 기준보다 높으니 발행 전 직접 확인하세요", "warning")

        body_len = len(strip_tags(body))
        min_len = int(CONFIG.get("min_body_chars", 2500))
        self.log(f"✅ 제목: {title} ({body_len:,}자)", "success" if body_len >= min_len else "warning")
        if body_len < min_len:
            self.log(f"  ⚠️ 목표 분량({min_len:,}자)보다 짧습니다", "warning")

        # 이미지
        self.log("🖼️ 이미지 생성/업로드 중...", "info")
        body = ensure_img_placeholders(body)
        img_dir = os.path.join(self._out_dir, "images")
        os.makedirs(img_dir, exist_ok=True)
        for i, text in enumerate(img_texts):
            self._check_stop()
            placeholder = f"[IMG{i + 1}]"
            clean_text = strip_tags(text).strip() or kw
            path = os.path.join(img_dir, f"{idx:02d}_{i + 1}.jpg")
            for w in create_card(clean_text, path):
                if w not in self._warned:  # 같은 경고는 한 번만
                    self._warned.add(w)
                    self.log(f"  ⚠️ {w}", "warning")
            url, err = upload_imgbb(path, f"card_{idx}_{i + 1}")
            if url:
                img_html = (f'<div style="text-align:center; margin:30px 0;">'
                            f'<img src="{html.escape(url, quote=True)}" '
                            f'alt="{html.escape(clean_text, quote=True)}" '
                            f'style="max-width:100%; height:auto;"></div>')
                body = body.replace(placeholder, img_html)
                self.log(f"  ✅ IMG{i + 1}: {clean_text}", "success")
            else:
                body = body.replace(placeholder, "")
                self.log(f"  ⚠️ IMG{i + 1} 업로드 실패: {err}", "warning")

        path = save_article_html(self._out_dir, idx + 1, kw, title, body)
        self.log(f"  💾 저장: {os.path.relpath(path, APP_DIR)}", "dim")
        return {"kw": kw, "title": title, "body": body}

    # ── 원고 생성 + 주입 ──
    def _generate_and_publish(self):
        keywords = self._keywords
        total = len(keywords)
        try:
            # 브라우저를 먼저 띄워서, 원고 생성되는 동안 로그인해둘 수 있게
            self._open_browser()

            articles = []
            for idx, kw in enumerate(keywords):
                self._check_stop()
                self.set_progress(f"원고 생성 [{idx + 1}/{total}] {kw}", idx, total)
                self.log(f"\n{'=' * 36}\n[{idx + 1}/{total}] 키워드: {kw}", "info")
                urls = self._sel_results[idx] if idx < len(self._sel_results) else []
                try:
                    art = self._generate_article(idx, kw, urls)
                except StopRequested:
                    raise
                except Exception as e:
                    self.log(f"❌ [{kw}] 처리 중 오류: {e}", "error")
                    art = None
                if art:
                    articles.append(art)

            self.set_progress("원고 생성 완료", total, total)
            if not articles:
                self.log("❌ 생성된 원고가 없습니다", "error")
                return
            self.log(f"\n✅ 원고 {len(articles)}/{total}개 완료 (저장 폴더: 📁 결과 폴더)", "success")

            if self.driver is None:
                self.log("⚠️ 브라우저를 열 수 없어 주입을 건너뜁니다. 저장된 HTML 파일을 사용하세요.",
                         "warning")
                return

            if not self._logged_in:
                self.log("🔐 브라우저에서 로그인 후 [로그인 완료] 버튼을 눌러주세요.", "info")
                self.ui(self._show_publish_btn, "✅  로그인 완료 → 첫 번째 글 주입")
                self._wait(self.login_event)

            self._inject_all(articles)

        except StopRequested as e:
            self.log(f"■ {e}" if str(e) else "■ 중지되었습니다", "warning")
        except Exception as e:
            self.log(f"❌ 실행 오류: {e}", "error")
        finally:
            self.ui(self._finish_run)

    def _open_browser(self):
        if driver_alive(self.driver):
            self.log("🌐 열려 있는 브라우저 재사용 (이미 로그인돼 있으면 바로 [로그인 완료])", "info")
        else:
            self.log("🌐 크롬 실행 중...", "info")
            try:
                self.driver = setup_driver()
            except Exception as e:
                self.driver = None
                self.log(f"⚠️ 크롬 실행 실패: {e}", "warning")
                self.log("   크롬 버전 문제라면 ⚙ 설정에서 '크롬 메인 버전'을 맞춰주세요.", "warning")
                return
        try:
            self.driver.get(self._base_url.rstrip("/") + CONFIG.get("tistory_write_path", "/manage/post"))
        except Exception:
            pass
        self.log("🔐 원고를 만드는 동안 브라우저에서 티스토리에 로그인해두세요. "
                 "로그인 후 [로그인 완료] 버튼을 누르면 됩니다.", "info")
        self.ui(self._show_publish_btn, "✅  로그인 완료")

    def _inject_all(self, articles):
        n = len(articles)
        done = 0
        for idx, art in enumerate(articles):
            self._check_stop()
            self.set_progress(f"주입 [{idx + 1}/{n}] {art['title']}", idx, n)
            self.log(f"\n📤 [{idx + 1}/{n}] 주입 중: {art['title']}", "info")
            try:
                ok, msg = inject_post(self.driver, self._base_url, art["title"], art["body"])
            except Exception as e:
                ok, msg = False, f"{e.__class__.__name__}: {str(e)[:120]}"

            if not ok:
                self.log(f"⚠️ {msg} — 저장된 HTML 파일에서 복사해 쓰세요", "warning")
                if idx == n - 1:
                    break
                self.ui(self._show_publish_btn, f"⏭  다음 글 주입 (남은 {n - idx - 1}개)")
                self._wait(self.next_event)
                continue

            done += 1
            if idx == n - 1:
                self.log(f"✅ 마지막 글 [{art['title']}] 주입 완료! 브라우저에서 발행 버튼을 눌러주세요.",
                         "success")
                break
            self.log(f"✅ 주입 완료! 브라우저에서 발행한 뒤 [발행 완료] 버튼을 누르면 "
                     f"[{articles[idx + 1]['title']}] 이(가) 주입됩니다.", "success")
            self.ui(self._show_publish_btn, f"✅  발행 완료 → 다음 글 주입 (남은 {n - idx - 1}개)")
            self._wait(self.next_event)

        self.set_progress(f"✅ {done}/{n}개 주입 완료", n, n)
        self.log(f"\n🎉 {done}/{n}개 주입 완료! (브라우저는 열어둡니다)", "success")

    # ── 설정 창 ──
    def _open_settings(self, first_run=False):
        if getattr(self, "_settings_win", None) and self._settings_win.winfo_exists():
            self._settings_win.lift()
            return
        win = tk.Toplevel(self.root, bg=BG, padx=20, pady=16)
        self._settings_win = win
        win.title("설정")
        win.geometry("540x820")
        win.transient(self.root)

        if first_run:
            tk.Label(win, text="처음 실행입니다. 키를 입력하고 저장해주세요.",
                     font=(FONT, 9, "bold"), bg=BG, fg=WARNING).pack(anchor="w", pady=(0, 8))

        entries = {}

        def field(label, key, secret=False, value=None):
            tk.Label(win, text=label, font=(FONT, 9), bg=BG, fg=TEXT).pack(anchor="w", pady=(6, 2))
            e = tk.Entry(win, font=(FONT, 10), bg=BG2, fg=TEXT, insertbackground=TEXT,
                         relief="flat", show="•" if secret else "")
            e.insert(0, str(CONFIG.get(key, "") if value is None else value))
            e.pack(fill="x", ipady=4)
            entries[key] = e

        # ── 원고 작성 AI ──
        tk.Label(win, text="원고 작성 AI", font=(FONT, 9, "bold"), bg=BG, fg=TEXT).pack(anchor="w")
        provider_var = tk.StringVar(value=CONFIG.get("ai_provider", "claude_cli"))
        prow = tk.Frame(win, bg=BG)
        prow.pack(fill="x")
        for val, text in [("claude_cli", "Claude Code (구독, API 비용 없음)"), ("gemini", "Gemini API")]:
            tk.Radiobutton(prow, text=text, variable=provider_var, value=val, font=(FONT, 9),
                           bg=BG, fg=TEXT, selectcolor=BG3, activebackground=BG,
                           activeforeground=TEXT).pack(side="left", padx=(0, 10))
        field("Claude 모델 (비우면 기본값, 예: sonnet / opus)", "claude_model")
        field("Claude Code 실행 파일 경로 (비우면 자동 찾기)", "claude_cli_path")
        test_row = tk.Frame(win, bg=BG)
        test_row.pack(fill="x", pady=(6, 0))
        test_label = tk.Label(test_row, text="", font=(FONT, 8), bg=BG, fg=TEXT_DIM,
                              wraplength=330, justify="left")

        def test_claude():
            # 저장 전 입력값으로 테스트
            CONFIG["claude_model"] = entries["claude_model"].get().strip()
            CONFIG["claude_cli_path"] = entries["claude_cli_path"].get().strip()
            test_label.configure(text="테스트 중... (10초 정도)", fg=TEXT_DIM)

            def work():
                try:
                    r = call_claude_cli("'연결 성공'이라고만 답해주세요.", timeout=120)
                    self.ui(lambda: test_label.winfo_exists() and
                            test_label.configure(text=f"✅ 연결됨: {r.strip()[:30]}", fg=SUCCESS))
                except Exception as e:
                    msg = str(e)
                    self.ui(lambda: test_label.winfo_exists() and
                            test_label.configure(text=f"❌ {msg}", fg=ERROR))
            threading.Thread(target=work, daemon=True).start()

        self._btn(test_row, "Claude 연결 테스트", test_claude, bg=BG3, fg=TEXT, size=8,
                  bold=False, pady=4).pack(side="left", ipadx=8)
        test_label.pack(side="left", padx=8)

        field("Gemini API 키 (Gemini 사용 시 또는 Claude 실패 시 백업)", "gemini_api_key", secret=True)
        field("Gemini 모델 (쉼표로 구분, 앞에서부터 시도)", "gemini_models",
              value=", ".join(CONFIG.get("gemini_models", [])))

        tk.Frame(win, bg=BORDER, height=1).pack(fill="x", pady=(12, 4))
        field("네이버 Client ID", "naver_client_id")
        field("네이버 Client Secret", "naver_client_secret", secret=True)
        field("imgbb API 키", "imgbb_api_key", secret=True)
        field("크롬 메인 버전 (비우면 자동 감지, 오류 날 때만 예: 145)", "chrome_version_main",
              value=CONFIG.get("chrome_version_main") or "")
        field("참고글 겹침 허용치 (0~1, 낮을수록 원고를 더 많이 바꿈)", "max_similarity")

        tk.Label(win, text="티스토리 계정 (한 줄에 '이름 URL')", font=(FONT, 9),
                 bg=BG, fg=TEXT).pack(anchor="w", pady=(6, 2))
        acc_text = tk.Text(win, font=(FONT, 10), bg=BG2, fg=TEXT, insertbackground=TEXT,
                           relief="flat", height=4)
        acc_text.insert("1.0", "\n".join(f"{k} {v}" for k, v in (CONFIG.get("accounts") or {}).items()))
        acc_text.pack(fill="x")

        def show_secrets():
            for k in ("naver_client_secret", "gemini_api_key", "imgbb_api_key"):
                entries[k].configure(show="" if entries[k].cget("show") else "•")

        def save():
            for k in KEY_LABELS:
                CONFIG[k] = entries[k].get().strip()
            CONFIG["ai_provider"] = provider_var.get()
            CONFIG["claude_model"] = entries["claude_model"].get().strip()
            CONFIG["claude_cli_path"] = entries["claude_cli_path"].get().strip()
            CONFIG["gemini_models"] = [m.strip() for m in entries["gemini_models"].get().split(",")
                                       if m.strip()] or DEFAULT_CONFIG["gemini_models"]
            ver = entries["chrome_version_main"].get().strip()
            CONFIG["chrome_version_main"] = int(ver) if ver.isdigit() else None
            try:
                CONFIG["max_similarity"] = min(1.0, max(0.0, float(entries["max_similarity"].get())))
            except ValueError:
                pass
            accounts = {}
            for line in acc_text.get("1.0", "end").splitlines():
                parts = line.split()
                if len(parts) >= 2 and parts[1].startswith("http"):
                    accounts[parts[0]] = parts[1].rstrip("/")
                elif len(parts) == 1 and parts[0]:
                    # 이름만 쓰면 기본 주소로
                    accounts[parts[0]] = f"https://{parts[0]}.tistory.com"
            CONFIG["accounts"] = accounts
            try:
                save_config(CONFIG)
            except Exception as e:
                messagebox.showerror("저장 실패", str(e), parent=win)
                return
            self._refresh_accounts()
            miss = missing_keys(CONFIG)
            self.log("⚙ 설정 저장 완료" + (f" (아직 비어있음: {', '.join(miss)})" if miss else ""),
                     "warning" if miss else "success")
            win.destroy()

        row = tk.Frame(win, bg=BG)
        row.pack(fill="x", side="bottom", pady=(12, 0))
        self._btn(row, "저장", save, pady=8).pack(side="right", ipadx=20)
        self._btn(row, "키 보기/숨기기", show_secrets, bg=BG3, fg=TEXT, size=9,
                  bold=False, pady=8).pack(side="left", ipadx=8)
        tk.Label(win, text=f"저장 위치: {CONFIG_PATH}", font=(FONT, 7), bg=BG, fg=TEXT_DIM,
                 wraplength=470, justify="left").pack(anchor="w", side="bottom", pady=(8, 0))


def main():
    root = tk.Tk()
    try:
        root.iconbitmap(get_resource_path("icon.ico"))
    except Exception:
        pass
    MacroApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
