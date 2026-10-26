# parser.py
import re
from pathlib import Path


# ──────────────────────────────────────────────────────────────
# 한글 서수 매핑
# ──────────────────────────────────────────────────────────────
_KO_ORDINALS = {
    '첫 번째': '1', '두 번째': '2', '세 번째': '3',
    '네 번째': '4', '다섯 번째': '5', '여섯 번째': '6',
    '일곱 번째': '7', '여덟 번째': '8', '아홉 번째': '9',
    '열 번째': '10',
    '첫번째': '1', '두번째': '2', '세번째': '3',
    '네번째': '4', '다섯번째': '5', '여섯번째': '6',
}


# ──────────────────────────────────────────────────────────────
# 글 분리 패턴
# ──────────────────────────────────────────────────────────────
_ARTICLE_SPLIT_PATTERNS = [
    # # 글 N
    r'\n+# 글 (\d+)\n',
    r'\n+## 글 (\d+)\n',
    r'\n+# 글(\d+)\n',
    r'\n---\n# 글 (\d+)\n',
    # # 버전 N / ## 버전 N (괄호 키워드 허용)
    r'(?:\n|^)## 버전 (\d+)[\s\S]*?\n',
    r'(?:\n|^)# 버전 (\d+)[\s\S]*?\n',
    # 버전 N / 버전N (단독 줄, 공백 있음/없음 모두, --- 구분자 앞뒤 모두 허용)
    r'(?m)(?:^---\n)?^버전\s*(\d+)\s*$\n',
    # **버전 N** / **버전N** (볼드 형식, 공백 있음/없음 모두, --- 구분자 앞뒤 모두 허용)
    r'(?m)(?:^-{3}\n){0,2}^\*\*버전\s*(\d+)\*\*\n',
    # ✅ VERSION N
    r'(?:\n|^)✅\s*VERSION\s*(\d+)\s*\n',
    # **N번 글** / **N번** 형식 (문서 시작 포함)
    r'(?:^|\n)\*\*(\d+)번 글\*\*\n',
    r'(?:^|\n)\*\*(\d+)번\*\*\n',
    # 글 N (# 없이, 공백 있음/없음 모두)
    r'\n글 (\d+)\n',
    r'(?m)(?:^|\n)글(\d+)\s*$\n',  # 글1, 글2 (공백 없이)
    # 📄 칼럼 N / 📝 칼럼 N / 칼럼 N
    r'\n📄\s*칼럼\s*(\d+)\s*\n',
    r'\n📝\s*칼럼\s*(\d+)\s*\n',
    r'\n칼럼\s*(\d+)\s*\n',
    # 📄 글 N
    r'\n📄\s*글\s*(\d+)\s*\n',
    # [Article N] / [글 N] / [칼럼 N]
    r'\n\[(?:Article|글|칼럼)\s*(\d+)\]\s*\n',
]


def _ko_ordinal_pattern():
    keys = '|'.join(re.escape(k) for k in _KO_ORDINALS)
    return rf'\n\*\*\[({keys})\s*글?\]?\*\*\n'


# ──────────────────────────────────────────────────────────────
# 메타 패턴
# ──────────────────────────────────────────────────────────────
_FOCUS_PATTERNS = [
    r'\*\*Focus Keyword:\*\*\s*(.+)',
    r'\*\*Focus Keyword\*\*:\s*(.+)',
    r'Focus Keyword:\s*(.+)',
    r'\*\*포커스 키워드:\*\*\s*(.+)',
    r'포커스 키워드:\s*(.+)',
]
_SUB_PATTERNS = [
    r'\*\*보조 키워드:\*\*\s*(.+)',
    r'\*\*보조 키워드\*\*:\s*(.+)',
    r'보조 키워드:\s*(.+)',
]
_TITLE_PATTERNS = [
    r'## 제목\n(.+)',
    r'\*\*제목:\*\*\s*(.+)',
    r'제목:\s*(.+)',
    r'---\n+# (.+)\n',
    # # 제목 (단독 # 헤딩)
    r'(?m)^# (.+)\n',
    # **볼드 제목** 단독 줄 (숫자. 로 시작하는 목차 항목 제외)
    r'(?m)^\*\*([^*\[\n\d][^*\n]{4,})\*\*\s*$',
    # 글 N\n\n제목 형태
    r'^글\s*\d+\n+([^\n#\*\-Q].+?)\n',
    # content 첫 비어있지 않은 줄 폴백
    r'^\s*\n([^\n#\*\-\d].+?)\n',
]
_SUMMARY_PATTERNS = [
    # 라벨형 포맷: "요약문\n(내용)" 뒤에 번호(1. 2. 3.)로 시작하는 줄이 나오면 그 직전에서 끊는다.
    # (AI가 요약문 뒤에 소제목 줄을 실수로 중복 출력해도 요약문에 딸려 들어가지 않도록 방지)
    r'요약문\n(.+?)(?=\n\d+[.\s]|\Z)',
    r'## 요약문\n(.+?)(?=\n##|\Z)',
    r'## 요약\n(.+?)(?=\n##|\Z)',
    r'\*\*요약문\*\*\n(.+?)(?=\n\*\*목차\*\*|\n\*\*[①②③④⑤⑥]|\n---|\Z)',
    r'\*\*요약\*\*\n(.+?)(?=\n\*\*|\n---|\Z)',
    r'\*\*요약:\*\*\s*(.+?)(?=\n\*\*|\Z)',
    # 제목 블록 이후 나오는 요약문 (📄 칼럼 형식: 제목\n...\n요약문\n내용)
    r'제목\n.+\n(?:.*\n)*?요약문\n(.+?)(?=\n[가-힣A-Za-z""\']|\Z)',
    # 번호 목차 다음 요약문
    r'목차\n(?:[ \t]*\d+\..+\n){2,}\n?요약문\n(.+?)(?=\n[가-힣""\']|\Z)',
]
_TOC_PATTERNS = [
    r'##[^\n]*목차[^\n]*\n((?:.+\n)+?)(?=\n##)',
    r'\*\*목차\*\*\n((?:.+\n)+?)(?=\n\*\*[①②③④⑤⑥]|\n\*\*요약|\n##|\n---)',
    r'목차\n\n?((?:\*\*.+\*\*\n\n?)+?)(?=\n##|\nQ|\n---|\Z)',
    r'목차\n((?:[ \t]*\d+\..+\n)+?)(?=\n---|\n\*\*)',
    # 번호 목차 (요약문 전까지)
    r'목차\n((?:[ \t]*\d+\..+\n)+?)(?=\n?요약문|\n?요약)',
    # 번호 목차 (Q1. 또는 빈줄 전까지) — Q&A 형식 원고
    r'목차\n((?:[ \t]*\d+\..+\n)+?)(?=\n?Q\d+\.|\n?\*\*Q\d+|\Z)',
    # 번호 목차 (일반 — 뒤에 무엇이 오든 잡기)
    r'목차\n((?:[ \t]*\d+\..+\n){2,})',
    # 번호/볼드 없이 줄글로만 나열된 목차 (빈 줄 뒤 짧은 문장들, 다음 빈 줄까지)
    r'목차\n\n?((?:[^\n#*\-\d][^\n]{1,100}\n){2,})(?=\n)',
]


# ──────────────────────────────────────────────────────────────
# 유틸
# ──────────────────────────────────────────────────────────────
def _first_match(patterns, text, flags=0):
    for p in patterns:
        m = re.search(p, text, flags)
        if m:
            return m.group(1).strip()
    return ''


def strip_markdown(text):
    """마크다운 기호 제거 + ①②③ → 1.2.3."""
    circle = '①②③④⑤⑥⑦⑧⑨⑩'
    text = re.sub(r'[①②③④⑤⑥⑦⑧⑨⑩]',
                  lambda m: str(circle.index(m.group()) + 1) + '.', text)
    text = re.sub(r'\*\*(.+?)\*\*', r'\1', text)
    text = re.sub(r'\*(.+?)\*',     r'\1', text)
    text = re.sub(r'__(.+?)__',     r'\1', text)
    text = re.sub(r'_(.+?)_',       r'\1', text)
    text = re.sub(r'^#{1,6}\s+', '', text, flags=re.MULTILINE)
    text = re.sub(r'^-{3,}$',    '', text, flags=re.MULTILINE)
    text = re.sub(r'^\*{3,}$',   '', text, flags=re.MULTILINE)
    text = re.sub(r'`(.+?)`',    r'\1', text)
    text = re.sub(r'\[(.+?)\]\(.+?\)', r'\1', text)
    text = re.sub(r'^[•·]\s*',   '', text, flags=re.MULTILINE)
    text = re.sub(r'^&nbsp;\s*', '', text, flags=re.MULTILINE)
    return text.strip()


def _clean_but_keep_bold(text):
    """strip_markdown과 동일하게 구조적 마크다운 기호(헤딩/구분선/코드/링크/글머리표)를
    제거하되, **볼드**·*이탤릭* 강조 표시는 그대로 보존한다.
    (이후 to_onenote_richtext에서 <b>/<i> 태그로 변환하기 위한 중간 단계)"""
    circle = '①②③④⑤⑥⑦⑧⑨⑩'
    text = re.sub(r'[①②③④⑤⑥⑦⑧⑨⑩]',
                  lambda m: str(circle.index(m.group()) + 1) + '.', text)
    text = re.sub(r'^#{1,6}\s+', '', text, flags=re.MULTILINE)
    text = re.sub(r'^-{3,}$',    '', text, flags=re.MULTILINE)
    text = re.sub(r'^\*{3,}$',   '', text, flags=re.MULTILINE)
    text = re.sub(r'`(.+?)`',    r'\1', text)
    text = re.sub(r'\[(.+?)\]\(.+?\)', r'\1', text)
    # 글머리표(•/·)는 삭제하지 않고 "• "로 정규화한다.
    # (wp_upload.py의 "종합 결론" 리스트 항목 인식이 이 불릿에 의존하므로 보존 필요)
    text = re.sub(r'^[•·]\s*', '• ', text, flags=re.MULTILINE)
    text = re.sub(r'^&nbsp;\s*', '', text, flags=re.MULTILINE)
    return text.strip()


def to_onenote_richtext(text):
    """OneNote one:T 요소(CDATA)에 그대로 넣을 수 있는 인라인 서식 HTML로 변환.
    **볼드**/__볼드__ → <b>, *이탤릭*/_이탤릭_ → <i> 로 유지하고,
    나머지 마크다운 기호는 strip_markdown과 동일하게 제거한다."""
    text = _clean_but_keep_bold(text)
    # 태그 삽입 전에 HTML 특수문자를 이스케이프 (우리가 넣을 <b>/<i>는 이후 삽입되므로 안전)
    text = text.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
    # re.DOTALL: 한 단락 안에서 줄바꿈을 사이에 두고 걸쳐 있는 볼드/이탤릭도 인식
    text = re.sub(r'\*\*(.+?)\*\*', r'<b>\1</b>', text, flags=re.DOTALL)
    text = re.sub(r'__(.+?)__',     r'<b>\1</b>', text, flags=re.DOTALL)
    text = re.sub(r'\*(.+?)\*',     r'<i>\1</i>', text, flags=re.DOTALL)
    text = re.sub(r'_(.+?)_',       r'<i>\1</i>', text, flags=re.DOTALL)
    # 짝이 맞지 않아 변환되지 못하고 남은 마크다운 기호(예: 원문 오타로 **가 홀수 개)는
    # 화면에 별표가 그대로 노출되지 않도록 제거 (서식은 못 살리더라도 텍스트는 깨끗하게)
    text = text.replace('**', '').replace('__', '')
    return text


def _normalize_invisible_chars(text):
    """노션/워드 등에서 '빈 줄'을 유지하려고 넣는 보이지 않는 특수문자를 제거한다.

    점자 빈칸(⠀ U+2800)이 대표적 — 화면·채팅창에는 빈 줄처럼 보이지만
    Python 정규식의 \\s 는 이 문자를 공백으로 인식하지 않는다. 그 결과 실제로는
    빈 줄이 하나도 없는 텍스트로 취급되어, 빈 줄(\\n\\n)에 의존하는 제목/목차/
    섹션 분리 로직이 전부 실패하고 제목 자리에 이 문자 자체가 잡히는 문제가 있었다.
    글 분리(_split_articles)보다 먼저, 파싱 진입점에서 한 번만 실행한다."""
    if not text:
        return text
    # 제로폭 문자·BOM류: 완전히 제거
    for ch in ('\ufeff', '\u200b', '\u200c', '\u200d', '\u2800'):
        text = text.replace(ch, '')
    # 줄바꿈처럼 쓰이는 유니코드 구분자는 실제 줄바꿈으로 통일
    text = text.replace('\u2028', '\n').replace('\u2029', '\n')
    # 줄바꿈 없는 공백류(NBSP 등)는 일반 공백으로 치환
    for ch in ('\u00a0', '\u2007', '\u202f'):
        text = text.replace(ch, ' ')
    return text


def _normalize_paragraph_breaks(content):
    """문단 사이에 빈 줄이 전혀(또는 거의) 없이, 한 줄바꿈(\\n)만으로 문단을
    구분하는 원고를 정규화한다. (💡 핵심 포인트 / 📋 체크리스트 형식 원고에서
    자주 나타남 — 아래 예시처럼 모든 문단이 개행 하나로만 붙어 있는 경우)

        제목 줄
        키워드 반복 줄
        본문 문단1
        💡 핵심 포인트
        본문 문단2

    제목 추출 자체는 줄 단위 폴백이라 영향이 없지만, OneNote로 내보낼 때
    문단을 나누는 로직(onenote_com.py)과 섹션 추출 로직은 전부 빈 줄(\\n\\n)에
    의존하기 때문에, 빈 줄이 없으면 글 전체가 한 덩어리로 뭉쳐버린다.
    이미 빈 줄 기반으로 작성된 원고(목차·Q&A 등 줄 단위 구조가 의미를 갖는
    경우)는 건드리면 오히려 망가지므로, 빈 줄 비율이 충분히 낮을 때만 적용한다."""
    lines = content.split('\n')
    non_empty = [l for l in lines if l.strip()]
    if len(non_empty) < 4:
        return content
    blank_breaks = len(re.findall(r'\n[ \t]*\n', content))
    # 이미 빈 줄 구분이 어느 정도 쓰이고 있으면 원문을 그대로 둔다
    # (목차/Q&A처럼 줄 단위 나열이 의도적인 원고를 오염시키지 않기 위함)
    if blank_breaks >= len(non_empty) * 0.3:
        return content
    out = []
    for i, ln in enumerate(lines):
        out.append(ln)
        if i == len(lines) - 1:
            continue
        if ln.strip() == '' or lines[i + 1].strip() == '':
            continue  # 이미 빈 줄이면 중복으로 추가하지 않음
        out.append('')
    return '\n'.join(out)


def _preprocess(content):
    # 글머리표(•/·)는 삭제하지 않고 "• "로 정규화한다.
    # (wp_upload.py의 "종합 결론" 리스트 항목 인식이 이 불릿에 의존하므로 보존 필요)
    content = re.sub(r'^[•·]\s*', '• ', content, flags=re.MULTILINE)
    content = re.sub(r'^-{3,}\s*$', '---', content, flags=re.MULTILINE)
    content = re.sub(r'^&nbsp;\s*$', '', content, flags=re.MULTILINE)
    content = _normalize_paragraph_breaks(content)
    return content


def _trim_body(body):
    m = re.search(r'\n(?!\*\*)\s*(요약문|목차)\s*\n', body)
    if m:
        body = body[:m.start()]
    body = re.sub(r'\n{3,}', '\n\n', body)
    return body.strip()


def _has_article_content(text):
    return bool(re.search(
        r'Focus Keyword|포커스 키워드|보조 키워드|##\s|\*\*제목|Q1\.|목차|\[사건 개요\]|\[해당 사건',
        text))


def _build_toc_number_map(content):
    """목차에서 '번호 → 소제목 텍스트' 역매핑을 반환. 소제목에 번호 붙일 때 사용."""
    toc_map = {}
    m = re.search(r'목차\n((?:\d+\..+\n)+)', content)
    if m:
        for line in m.group(1).strip().split('\n'):
            lm = re.match(r'(\d+)\.\s*(.+)', line.strip())
            if lm:
                toc_map[lm.group(2).strip()] = lm.group(1)
    return toc_map


# ──────────────────────────────────────────────────────────────
# 섹션 추출
# ──────────────────────────────────────────────────────────────
_FAQ_HEADING_RE = re.compile(r'\n#{0,3}\s*자주\s*묻는\s*질문[^\n]*\n+')


_BARE_Q1_RE = re.compile(r'\n(?:\*\*)?Q1\.\s')


def _split_faq_block(content):
    """'자주 묻는 질문' 헤딩을 기준으로 본문과 FAQ 블록을 분리한다.
    (본문 소제목 섹션(1./2./3. 등)과 글 말미의 FAQ Q&A가 함께 있는 형식을
    지원하기 위함 — 이전에는 Q&A가 발견되면 본문 섹션을 통째로 버렸음)"""
    m = _FAQ_HEADING_RE.search(content)
    if m:
        return content[:m.start()], content[m.end():]
    # 명시적 "자주 묻는 질문" 헤딩이 없어도 "Q1. ..."로 시작하는 FAQ 블록이
    # 본문 뒤에 바로 붙는 경우가 있으므로, 첫 Q1. 위치를 기준으로 분리를 시도한다.
    # (본문 소제목 섹션이 앞에 있는 경우만 분리 — Q1.이 글 맨 앞이면 전체가 Q&A 형식)
    m2 = _BARE_Q1_RE.search(content)
    if m2 and m2.start() > 30:
        return content[:m2.start()], content[m2.start() + 1:]
    return content, ''


def _extract_mixed_numeric_qa_sections(content):
    """숫자 단독 줄(1,2,3...)로 구분된 블록 중 일부만 Q&A 형식이고
    나머지는 도입부/전환용 평문 문단인 혼합 원고를 처리한다.
    (예: 1=도입, 2=Q&A, 3=평문, 4=Q&A, 5=평문, 6=Q&A, 7=마무리 평문)
    기존 방식(등장한 첫 Q1. 이후를 통째로 FAQ로 간주)은 평문 블록이
    통째로 유실되거나 마지막 Q&A 답변에 뒤 문단이 흡수되는 문제가 있어,
    블록 단위로 개별 검사해 유실/오염을 방지한다."""
    numeric_blocks = _extract_numeric_delimiter_sections(content)
    if not numeric_blocks:
        return []
    result = []
    for label, body in numeric_blocks:
        block_qa = _extract_qa_sections(body)
        if block_qa:
            # 블록을 구분하던 숫자(1,2,3...)가 사라지지 않도록 첫 Q&A 본문 앞에 되살려 둔다.
            first_t, first_b = block_qa[0]
            if label:
                block_qa[0] = (first_t, label + '\n\n' + first_b)
            result.extend(block_qa)
        elif body.strip():
            merged = (label + '\n\n' + body) if label else body
            result.append(('', _trim_body(merged)))
    return result


def _extract_sections(content, title=None):
    # 숫자 단독 줄 블록 수가 Q&A 쌍 수보다 많으면, 도입부/전환 평문이 섞인
    # 혼합 원고일 가능성이 높음 → 블록 단위 처리 결과가 더 많은 섹션을
    # 살려내면(=기존 방식이 뭔가를 잃어버렸다는 뜻) 그쪽을 우선 사용한다.
    if len(re.findall(r'(?m)^\d+\s*$', content)) >= 3:
        _mixed = _extract_mixed_numeric_qa_sections(content)
        _naive_qa = _extract_qa_sections(content)
        if len(_mixed) > len(_naive_qa):
            return _mixed

    # '자주 묻는 질문' 헤딩이 있으면 본문/FAQ를 분리해서 각각 처리 후 합친다.
    main_content, faq_content = _split_faq_block(content)
    if faq_content:
        # FAQ 헤딩 앞에 남는 "---" 구분선 잔여물 제거 (연속으로 "---"가 여러 번 나오는 경우까지 모두 제거)
        main_content = re.sub(r'(?:\n-{3,}\s*)+\Z', '', main_content).rstrip()

    # 섹션 추출 전 중복 목차 블록(📋 목차)만 제거한 사본 사용
    # 일반 '목차\n1. ...' 는 제거하지 않음 — TOC 추출에 필요
    _c = re.sub(
        r'(?m)^(?:#{1,3}\s*)?📋\s*목차\s*\n+((?:\d+\..+\n?)+)\n?',
        '', main_content
    )
    _c = re.sub(r'(?s)(?:##\s*)?요약문\n(.+?)(?=\n\n|\Z)', '', _c)
    # 일반 '목차\n1. ...\n2. ...\n' 번호 목차 블록도 섹션 탐지용 사본(_c)에서는 제거해야
    # 목차의 "N. 소제목" 줄이 본문 섹션으로 오탐되는 것을 막을 수 있다.
    _c = re.sub(r'(?m)^목차\s*\n((?:[ \t]*\d+\..*\n)+)', '', _c)

    # FAQ 블록이 있으면 그 안에서만 Q&A 추출
    qa = _extract_qa_sections(faq_content) if faq_content else []

    # FAQ 헤딩 자체가 없는 경우 → 글 전체가 Q&A 형식일 수 있음 (기존 동작 유지)
    if not faq_content:
        qa = _extract_qa_sections(main_content)
        if qa:
            return qa

    # 2. [대괄호 제목] 형식
    bracket = _extract_bracket_sections(main_content, title)
    if bracket:
        return bracket + qa

    # 3. --- 블록 분리 (연속 구분선 ---\n--- 정규화 후 분리)
    _content_for_split = re.sub(r'\n---\n---\n', '\n---\n', main_content)
    blocks = re.split(r'\n---\n', _content_for_split)
    sections = []
    for block in blocks:
        block = block.strip()
        if not block:
            continue
        sec = _block_to_section(block)
        if sec:
            sections.append(sec)

    # 3-1. --- 없이 ## N. 형식이 연속으로 오는 경우
    if not sections or len(sections) == 1:
        _pat31 = r'(?m)^## (\d+\.[^\n]+)\n+([\s\S]+?)(?=\n## \d+\.|\Z)'
        _num31 = re.findall(_pat31, main_content)
        if len(_num31) > len(sections):
            sections = [(t.strip(), b.strip()) for t, b in _num31]

    # 3-2. "숫자. 소제목\n\n본문" 형식 — _c(목차 제거본) 사용으로 오탐 방지
    if not sections or len(sections) == 1:
        _pat32 = r'(?m)^(\d+\.\s+[^\n]{5,}[?？]?)\n\n([\s\S]+?)(?=\n\d+\.\s+[^\n]{5,}|\Z)'
        _num32 = re.findall(_pat32, _c)
        if len(_num32) >= 2 and len(_num32) > len(sections):
            sections = [(t.strip(), b.strip()) for t, b in _num32]

    # 3-3. 번호 없이 소제목 단독 줄 (물음표/마침표 끝) + 빈줄 + 본문 — _c 사용
    if not sections or len(sections) == 1:
        _pat33 = (
            r'\n\n([^#\n\*\-①②③④⑤⑥⑦⑧⑨⑩>][^\n]{4,80}[?？])\n\n'
            r'([\s\S]+?)(?=\n\n[^#\n\*\-①②③④⑤⑥⑦⑧⑨⑩>][^\n]{4,80}[?？]\n\n|\Z)'
        )
        _num33 = re.findall(_pat33, _c)
        _skip = re.compile(r'^(목차|요약문|요약|Focus|보조 키워드|제목|📄|📝)')
        _num33 = [(t, b) for t, b in _num33 if not _skip.match(t.strip())]
        if len(_num33) >= 2 and len(_num33) > len(sections):
            _toc_map = _build_toc_number_map(main_content)
            _numbered = []
            for idx, (t, b) in enumerate(_num33, start=1):
                t_clean = t.strip()
                num = _toc_map.get(t_clean, str(idx))
                _numbered.append((f'{num}. {t_clean}', b.strip()))
            sections = _numbered

    # 3-4. "숫자. 소제목\n본문" 형식 (제목과 본문 사이에 빈 줄이 없는 경우) — _c 사용
    # (예: "1. 강간 사건에서 CCTV가...\n강간 사건에서 CCTV 영상이...")
    if not sections or len(sections) == 1:
        _pat34 = r'(?m)^(\d+)\.\s*([^\n]{5,})\n([\s\S]+?)(?=\n\d+\.\s*[^\n]{5,}\n|\Z)'
        _num34 = re.findall(_pat34, _c)
        if len(_num34) >= 2 and len(_num34) > len(sections):
            sections = [(f'{n}. {t.strip()}', b.strip()) for n, t, b in _num34]

    # 4. 숫자 단독 줄이 단락 구분자인 형식 — _c 사용
    if not sections:
        nd = _extract_numeric_delimiter_sections(_c)
        if nd:
            sections = nd

    # 4-1. 단독 숫자 패턴 (기존 fallback)
    if not sections:
        sections = _extract_numbered_sections(main_content)

    # 5. **볼드 제목** + 본문 (✅ VERSION 형식)
    if not sections:
        sections = _extract_bold_sections(main_content)

    return [(t, _trim_body(b)) for t, b in sections] + qa


def _extract_qa_sections(content):
    """Q&A 형식 - 여러 변형 지원"""
    # "A." 뒤에 줄바꿈을 넣으면 OneNote 등으로 넘어갈 때 한 문단 텍스트 안에 개행 문자가 그대로
    # 박혀버려서, 다른 곳에 복사해 붙여넣으면 줄바꿈이 무시되고 "Q2. ... A. ..."처럼
    # 한 줄로 뭉쳐 보이는 문제가 있었음. 공백으로 바꿔 원문처럼 한 줄(문단)로 자연스럽게 이어지게 함.
    sep = " "

    # **Q1. 질문?**\n\n**A. 관련 문의 답변**\n\n답변 형식 (Q와 A 라벨이 모두 볼드로 감싸지고,
    # "A." 뒤에 "관련 문의 답변"이라는 라벨이 붙은 뒤 실제 답변이 다음 문단에 오는 경우)
    p0a = r'\*\*Q(\d+)\.\s*(.+?)\*\*\s*\n+\*\*A\d*\.\s*관련\s*문의\s*답변\*\*\s*\n+([\s\S]+?)(?=\n\*{0,2}Q\d+\.|\Z)'
    found0a = re.findall(p0a, content, re.DOTALL)
    if found0a:
        result = []
        for n, t, b in found0a:
            title = strip_markdown("Q" + n + ". " + t).strip()
            body  = "A." + sep + b.strip()
            result.append((title, body))
        return result

    # **Q1. 질문?**\nA. 답변 형식 (질문이 볼드로 감싸진 경우)
    p0 = r'\*\*Q(\d+)\.\s*(.+?)\*\*\s*\n+A\.\s*([\s\S]+?)(?=\n\*{0,2}Q\d+\.|\Z)'
    found0 = re.findall(p0, content, re.DOTALL)
    if found0:
        result = []
        for n, t, b in found0:
            title = strip_markdown("Q" + n + ". " + t).strip()
            body  = "A." + sep + b.strip()
            result.append((title, body))
        return result

    # ### Q1. ... ### A. 형식
    p1 = r'### Q(\d+)\.\s*(.+?)\n+(?:[\s\S]*?### A\d*\..*?\n+)([\s\S]+?)(?=\n### Q\d+\.|\Z)'
    found = re.findall(p1, content, re.DOTALL)
    if found:
        result = []
        for n, t, b in found:
            title = strip_markdown("Q" + n + ". " + t).strip()
            body  = "A." + sep + b.strip()
            result.append((title, body))
        return result

    # Q1. 제목\n\n질문본문\n\n### A. 형식
    p3 = r'Q(\d+)\.\s*(.+?)\n+[\s\S]*?### A\d*\.\s*관련 문의 답변\s*\n([\s\S]+?)(?=\nQ\d+\.|\Z)'
    found3 = re.findall(p3, content, re.DOTALL)
    if found3:
        result = []
        for n, t, b in found3:
            title = strip_markdown("Q" + n + ". " + t).strip()
            body  = "A." + sep + b.strip()
            result.append((title, body))
        return result

    # Q1. ... A. / A1. 형식 (### 없이)
    p2 = r'Q(\d+)\.\s*(.+?)\n+A\d*\.\s*관련 문의 답변\s*\n([\s\S]+?)(?=\nQ\d+\.|\Z)'
    found2 = re.findall(p2, content, re.DOTALL)
    if found2:
        result = []
        for n, t, b in found2:
            title = strip_markdown("Q" + n + ". " + t).strip()
            body  = "A." + sep + b.strip()
            result.append((title, body))
        return result

    # 일반 Q&A: Q1. 질문\n\nA. 답변 또는 A1. 답변 (번호 붙은 답변 라벨도 허용)
    p4 = r'Q(\d+)\.\s*(.+?)\n+A\d*\.\s*([\s\S]+?)(?=\nQ\d+\.|\Z)'
    found4 = re.findall(p4, content, re.DOTALL)
    if found4:
        result = []
        for n, t, b in found4:
            title = strip_markdown("Q" + n + ". " + t).strip()
            body  = "A." + sep + b.strip()
            result.append((title, body))
        return result

    return []


def _extract_bracket_sections(content, title=None):
    """[대괄호 제목] 형식 — 앞에 단독 숫자 줄이 있으면 본문 앞에 숫자 보존"""

    def _prepend_intro(results, first_match_start):
        """첫 대괄호 마커 이전 텍스트(도입부)가 있으면 별도 섹션으로 보존.
        (제목 줄과 중복되면 그 줄은 제거)"""
        intro = content[:first_match_start].strip()
        if not intro:
            return results
        if title:
            t_stripped = title.strip()
            first_line, sep, rest = intro.partition('\n')
            if first_line.strip() == t_stripped:
                intro = rest.strip()
            elif intro.startswith(t_stripped):
                intro = intro[len(t_stripped):].strip()
        if len(intro) >= 20:
            results.insert(0, ('', intro))
        return results

    # 숫자 단독줄 + **[제목]** 형식
    # 제목은 **[원문]** 볼드+괄호 그대로 유지, 숫자는 본문 맨 앞에
    p_numbered = r'(?:---\s*\n+)?(\d+)\s*\n+(\*{0,2}\[[^\]]+\]\*{0,2})\n+([\s\S]+?)(?=\n---|\n\d+\s*\n\n|\Z)'
    found_num = list(re.finditer(p_numbered, content, re.DOTALL))
    if len(found_num) >= 2:
        results = [(m.group(2).strip(), m.group(1) + '\n\n' + m.group(3).strip())
                   for m in found_num]
        return _prepend_intro(results, found_num[0].start())

    # 일반 [대괄호 제목] 형식
    pattern = r'\[([^\]]+)\]\n+([\s\S]+?)(?=\n\[[^\]]+\]|\Z)'
    found = list(re.finditer(pattern, content, re.DOTALL))
    if len(found) >= 2:
        results = [(strip_markdown(m.group(1)).strip(), m.group(2).strip()) for m in found]
        return _prepend_intro(results, found[0].start())
    return []


def _extract_bold_sections(content):
    """**볼드 제목**\n\n본문 형식 (✅ VERSION 등에서 사용)"""
    pattern = r'\*\*([^*\n]+)\*\*\n\n([\s\S]+?)(?=\n\*\*[^*\n]+\*\*\n\n|\Z)'
    found = re.findall(pattern, content, re.DOTALL)
    if len(found) >= 2:
        return [(strip_markdown(t).strip(), b.strip()) for t, b in found]
    return []


def _block_to_section(block):
    patterns = [
        # 번호 포함 — 캡처 그룹에 번호까지 포함
        (r'\*\*([①②③④⑤⑥⑦⑧⑨⑩]\s*.+?)\*\*\n+([\s\S]+)', True),
        (r'\*\*(\d+\.\s*.+?)\*\*\n+([\s\S]+)',              True),
        (r'### (.+?)\n+([\s\S]+)',                           False),
        (r'## (?!목차|요약|제목)(.+?)\n+([\s\S]+)',          False),
        # 숫자\n소제목(짧은 한 줄)\n\n본문 — 소제목은 100자 이내로 제한해
        # 번호로만 구분된 긴 단락 전체가 "소제목"으로 잘못 캡처되는 것을 방지
        (r'^\d+\n([^\n]{1,100})\n\n([\s\S]+)',               False),
    ]
    for p, has_num in patterns:
        m = re.match(p, block, re.DOTALL)
        if m:
            title = strip_markdown(m.group(1)).strip()
            return (title, m.group(2).strip())
    return None


def _extract_numbered_sections(content):
    p1 = r'\n\n(\d+)\n\n(.+?)\n\n([\s\S]+?)(?=\n\n\d+\n\n|\Z)'
    found = re.findall(p1, content, re.DOTALL)
    if found:
        return [(strip_markdown(title).strip(), body.strip())
                for _, title, body in found]
    p2 = r'\n(\d+)\n(.+?)\n\n([\s\S]+?)(?=\n\d+\n|\Z)'
    found2 = re.findall(p2, content, re.DOTALL)
    if found2:
        return [(strip_markdown(title).strip(), body.strip())
                for _, title, body in found2]
    return []


def _extract_numeric_delimiter_sections(content):
    """
    숫자 단독 줄이 단락 구분자인 형식.
    예:
        도입 본문 (번호 없음)

        1

        첫 번째 단락 본문

        2

        두 번째 단락 본문

    주의: 목차(1. 소제목), ## 헤딩 형식과 충돌하지 않도록
    - 콘텐츠에 ## 헤딩이 2개 이상 있으면 이 방식 사용 안 함
    - 콘텐츠에 '목차' 섹션이 있고 그 안에 숫자 줄이 있으면 사용 안 함
    """
    # ## 헤딩이 2개 이상이면 헤딩 기반 섹션 방식과 충돌 → 사용 안 함
    if len(re.findall(r'(?m)^##\s+', content)) >= 2:
        return []

    # 숫자를 캡처해서 분리 (실제 번호를 그대로 라벨로 사용하기 위함 - 재순번 매기지 않음)
    parts = re.split(r'(?m)^(\d+)\s*$', content)
    if len(parts) < 3:
        return []

    intro = parts[0].strip()
    intro = re.sub(
        r'^(?:제목|Focus Keyword|포커스 키워드|보조 키워드)[^\n]*\n?',
        '', intro, flags=re.MULTILINE
    ).strip()

    numbered = []
    pending_nums = []
    for i in range(1, len(parts) - 1, 2):
        num  = parts[i].strip()
        body = parts[i + 1].strip()
        pending_nums.append(num)
        # 본문이 없으면(연속된 숫자 단독 줄, 예: "1\n2") 번호만 쌓아두고
        # 다음 숫자와 합쳐서 라벨로 사용한다 (번호가 사라지지 않도록 보존)
        if not body or len(body) < 10:
            continue
        label = '\n'.join(pending_nums)
        numbered.append((label, body))
        pending_nums = []

    # 최소 3개 이상일 때만 이 방식 사용 (오탐 방지)
    if len(numbered) < 3:
        return []

    sections = []
    if intro and len(intro) >= 10:
        sections.append(('', intro))
    sections.extend(numbered)
    return sections


# ──────────────────────────────────────────────────────────────
# 글 분리
# ──────────────────────────────────────────────────────────────
def _split_articles(text):
    for p in _ARTICLE_SPLIT_PATTERNS:
        splits = re.split(p, text, flags=re.MULTILINE)
        if len(splits) >= 3:
            return splits, 'numeric'

    # 한글 서수
    ko_pat = _ko_ordinal_pattern()
    splits = re.split(ko_pat, text)
    if len(splits) >= 3:
        result = [splits[0]]
        for i in range(1, len(splits), 2):
            num = _KO_ORDINALS.get(splits[i].strip(), splits[i])
            result.append(num)
            if i + 1 < len(splits):
                result.append(splits[i + 1])
        return result, 'korean'

    return None, None


# ──────────────────────────────────────────────────────────────
# 휴리스틱 폴백: 알 수 없는 형식에서 글 단위 분리 시도
# ──────────────────────────────────────────────────────────────
_HEURISTIC_SPLIT_PATTERNS = [
    # == 제목 == 형식
    r'\n={2,}\s*(.+?)\s*={2,}\n',
    # *** 제목 *** 형식
    r'\n\*{3}\s*(.+?)\s*\*{3}\n',
    # 【제목】형식
    r'\n【([^】]+)】\n',
    # ─────── 수평선 3번 이상 반복
    r'\n[─━=]{5,}\n',
    # 빈 줄 4개 이상 연속 (단락 구분자로 사용하는 경우)
    r'\n{4,}',
    # < 제목 > 형식
    r'\n<\s*([^>]{2,40})\s*>\n',
    # [제목] 단독 줄 (대괄호, 짧은 제목)
    r'\n\[([^\]]{2,30})\]\n(?!\()',  # 링크 형식 제외
    # ◆◇■□▶ 등 특수문자로 시작하는 구분자
    r'\n[◆◇■□▶▷●○★☆]\s*(\d+)[^\n]*\n',
    # 숫자. 제목 (줄 시작, 앞에 빈 줄)
    r'\n\n(\d+)\.\s+([^\n]{5,})\n',
]


def _heuristic_split(text):
    """
    기존 패턴으로 분리 안 될 때 동작하는 폴백.
    글처럼 보이는 덩어리를 최대한 뽑아낸다.
    반환: [(번호문자열, 내용문자열), ...] 또는 []
    """
    # 1순위: 위 패턴 중 하나라도 2회 이상 등장하면 그걸로 분리
    for pat in _HEURISTIC_SPLIT_PATTERNS:
        parts = re.split(pat, text, flags=re.MULTILINE)
        if len(parts) < 3:
            continue
        # 캡처 그룹이 있으면 홀수 인덱스가 캡처값 → 짝수 인덱스가 내용
        # 캡처 그룹 없으면 모든 조각이 내용
        has_capture = bool(re.search(r'\((?!\?)', pat))
        if has_capture:
            # parts: [pre, cap1, body1, cap2, body2, ...]
            # 제목+본문을 합쳐서 글 단위로 구성
            chunks = []
            for i in range(1, len(parts) - 1, 2):
                title_part = parts[i].strip() if parts[i] else ''
                body_part  = parts[i + 1].strip() if i + 1 < len(parts) else ''
                combined   = (title_part + '\n\n' + body_part).strip()
                if len(combined) >= 100:
                    chunks.append(combined)
        else:
            chunks = [p.strip() for p in parts if p and len(p.strip()) >= 100]

        if len(chunks) >= 2:
            return [(str(idx + 1), chunk) for idx, chunk in enumerate(chunks)]

    # 2순위: 빈 줄 3개 이상으로 나뉜 큰 덩어리들
    parts = re.split(r'\n{3,}', text)
    big_parts = [p.strip() for p in parts if len(p.strip()) >= 100]
    if len(big_parts) >= 2:
        return [(str(i + 1), chunk) for i, chunk in enumerate(big_parts)]

    # 3순위: ## 헤딩이 여러 개 있으면 헤딩 단위로 분리
    heading_split = re.split(r'\n(## .+)\n', text)
    if len(heading_split) >= 5:  # 헤딩 2개 이상
        chunks = []
        for i in range(1, len(heading_split), 2):
            title = heading_split[i].strip()
            body  = heading_split[i + 1].strip() if i + 1 < len(heading_split) else ''
            if body and len(body) >= 100:
                chunks.append((str(len(chunks) + 1), title + '\n\n' + body))
        if len(chunks) >= 2:
            return chunks

    return []


# ──────────────────────────────────────────────────────────────
# ArticleParser
# ──────────────────────────────────────────────────────────────
class ArticleParser:

    @staticmethod
    def parse(text):
        if not text or not text.strip():
            return []
        text = _normalize_invisible_chars(text)
        splits, fmt = _split_articles(text)

        # 기존 패턴으로 분리 실패 → 휴리스틱 폴백
        if splits is None:
            # Focus Keyword/포커스 키워드 라벨이 1개 이하면 글이 1개뿐이라는 뜻이므로
            # 휴리스틱 분리(특히 "숫자. 소제목" 패턴)가 본문 내 소제목 1./2./3.을
            # 별개의 글로 잘못 쪼개는 것을 막는다.
            _focus_label_count = len(re.findall(r'Focus Keyword\s*:|포커스\s*키워드\s*:', text))
            if _focus_label_count <= 1:
                try:
                    art = ArticleParser._parse_single('1', text)
                    if ArticleParser._is_valid(art):
                        print("[Parser] 단일 글로 처리 (Focus Keyword 라벨 1개)")
                        return [art]
                except Exception as e:
                    print(f"[Parser] 단일 글 처리 오류: {e}")
                return []

            fallback = _heuristic_split(text)
            if fallback:
                articles = []
                for num, content in fallback:
                    try:
                        art = ArticleParser._parse_single(num, content)
                        if ArticleParser._is_valid(art):
                            articles.append(art)
                    except Exception as e:
                        print(f"[Parser][폴백] 글 {num} 오류: {e}")
                if articles:
                    print(f"[Parser] 휴리스틱 폴백으로 {len(articles)}개 글 추출")
                    return articles
            # 폴백도 실패 → 전체를 글 1개로 취급
            try:
                art = ArticleParser._parse_single('1', text)
                if ArticleParser._is_valid(art):
                    print("[Parser] 분리 실패, 전체를 글 1개로 처리")
                    return [art]
            except Exception as e:
                print(f"[Parser] 단일 글 처리 오류: {e}")
            return []
        articles = []

        # splits[0]: 마커 이전 미리보기 블록인 경우 글로 추가하지 않음
        # (📄 칼럼 N 앞의 목차+요약문만 있는 블록)
        first_content = splits[0].strip() if splits else ''
        if first_content and _has_article_content(first_content):
            # 미리보기 블록 판별:
            # - 요약문은 있고
            # - 실제 소제목+본문 구조가 없는 경우만 스킵
            _has_real_section = (
                # 물음표 끝 소제목
                bool(re.search(r'\n\n[^#\n\*\-①-⑩>][^\n]{4,80}[?？]\n\n', first_content)) or
                # **N. 소제목** 형태
                bool(re.search(r'\*\*\d+\.\s*.+?\*\*\n', first_content)) or
                # ## N. 소제목 형태
                bool(re.search(r'(?m)^##\s+\d+\.', first_content))
            )
            _is_preview = (
                bool(re.search(r'요약문|요약', first_content)) and
                not _has_real_section
            )
            if not _is_preview:
                try:
                    art = ArticleParser._parse_single('1', first_content)
                    if ArticleParser._is_valid(art):
                        articles.append(art)
                except Exception as e:
                    print(f"[Parser] 첫 번째 글 오류: {e}")

        for i in range(1, len(splits), 2):
            try:
                num     = splits[i].strip()
                content = splits[i + 1] if i + 1 < len(splits) else ''
                art     = ArticleParser._parse_single(num, content)
                if ArticleParser._is_valid(art):
                    articles.append(art)
            except Exception as e:
                print(f"[Parser] 글 {splits[i]} 오류: {e}")
        return articles

    @staticmethod
    def parse_from_file(file_path):
        path = Path(file_path)
        if not path.exists():
            raise FileNotFoundError(f"파일 없음: {file_path}")
        for enc in ('utf-8-sig', 'utf-8', 'cp949', 'euc-kr'):
            try:
                return ArticleParser.parse(path.read_text(encoding=enc))
            except UnicodeDecodeError:
                continue
        raise UnicodeDecodeError('인코딩 감지 실패')

    @staticmethod
    def get_preview(article, max_summary=300):
        lines = [
            f"📝 글 {article['number']}",
            '',
            f"🎯 Focus: {article['focus_keyword'] or '(없음)'}",
            f"🔑 보조: {article['sub_keywords'] or '(없음)'}",
            f"📌 제목: {article['title']}",
            '',
        ]
        if article['toc']:
            lines.append('📑 목차')
            for item in article['toc']:
                lines.append(f'  {item}')
            lines.append('')
        if article['summary']:
            lines.append('💡 요약')
            s = article['summary']
            lines.append(s[:max_summary] + ('…' if len(s) > max_summary else ''))
            lines.append('')
        if article['sections']:
            lines.append(f"📖 섹션 {len(article['sections'])}개")
            for t, _ in article['sections'][:3]:
                lines.append(f"  • {t}")
            if len(article['sections']) > 3:
                lines.append(f"  … 외 {len(article['sections'])-3}개")
        lines += ['', f"📊 {len(article['raw_content']):,}자"]
        return '\n'.join(lines)

    @staticmethod
    def detect_format(text):
        _, fmt = _split_articles(_normalize_invisible_chars(text))
        return fmt or '알 수 없음'

    @staticmethod
    def _parse_single(num, content):
        content = _preprocess(content)
        # 글 분리 시 다음 글 마커 앞에 남는 "---" / "===" 구분선 잔여물 제거
        content = re.sub(r'(?:\n-{3,}\s*)+\Z', '', content).rstrip()
        content = re.sub(r'(?:\n={3,}\s*)+\Z', '', content).rstrip()

        # 📄 칼럼 N / 글 N 마커가 있으면 그 이전 미리보기 블록 제거
        # (요약문+목차가 마커 앞에 붙어있는 경우)
        _marker_m = re.search(
            r'^(?:📄|📝)\s*(?:칼럼|글)\s*\d+\s*$',
            content, re.MULTILINE
        )
        if _marker_m:
            after = content[_marker_m.end():].strip()
            if after:
                content = after

        focus = _first_match(_FOCUS_PATTERNS, content)
        sub   = _first_match(_SUB_PATTERNS, content)

        _raw_title = _first_match(_TITLE_PATTERNS, content)
        if not _raw_title or re.match(r'^글 \d+$', _raw_title.strip()):
            _lines = [l.strip() for l in content.split('\n') if l.strip()]
            for _l in _lines[:5]:
                if not re.match(
                    r'^(\*\*|#|Q\d|A\.|목차|---|✅|&nbsp|Focus|포커스|보조 키워드|제목\s*$|📄|📝'
                    r'|\d+[.\s)]|[가-힣]\.\s)',  # 목차/본문 소제목(숫자·가나다 순번)은 제목 후보에서 제외
                    _l
                ):
                    _raw_title = _l
                    break
        title = _raw_title or f"글 {num}"

        # "포커스 키워드: X" / "제목: Y" 라벨 줄만 제거 (제목 값이 같은 줄이든 다음 줄이든 모두 처리).
        # 뒤따르는 키워드 태그 줄은 실제 본문 콘텐츠이므로 그대로 둔다.
        _preamble_m = re.match(
            r'\s*(?P<focus>포커스\s*키워드\s*:[^\n]*\n+)?제목\s*:\s*[^\n]*\n+',
            content
        )
        _had_focus_label = bool(_preamble_m and _preamble_m.group('focus'))
        if _preamble_m:
            content = content[_preamble_m.end():]

        # 글N 포맷(포커스 키워드 라벨이 있던 경우)은 태그줄 + "**제목 반복**" 줄이 모두 중복이라 함께 제거.
        # 버전N 포맷(포커스 키워드 라벨 없음)은 태그줄과 볼드 제목 줄을 그대로 본문에 남긴다.
        if _had_focus_label and title:
            content = re.sub(
                r'\A(?:[^\n]*\n\n)?\*\*' + re.escape(title) + r'\*\*\s*\n+',
                '', content, count=1
            )

        # 글 끝에 붙는 "키워드 삽입 횟수: ...", "총 글자수 (공백 제외): ...자" 같은
        # 작성 메모/통계 문구 제거 (본문이 아니라 검수용 메모)
        content = re.sub(
            r'\n+(?:키워드\s*삽입\s*횟수|총\s*글자수)[^\n]*'
            r'(?:\n(?:키워드\s*삽입\s*횟수|총\s*글자수)[^\n]*)*\s*\Z',
            '', content
        )

        toc_raw = _first_match(_TOC_PATTERNS, content)
        toc = []
        if toc_raw:
            for ln in toc_raw.split('\n'):
                ln = _clean_but_keep_bold(ln).strip()
                if ln:
                    # 원문에 번호가 없는 목차 줄(점/불릿으로만 표시되던 항목)에
                    # "1. 2. 3." 형식의 순번을 붙여준다. 이미 번호가 있으면 그대로 둔다.
                    if not re.match(r'^\d+[.\)]\s', ln):
                        ln = f'{len(toc) + 1}. {ln}'
                    toc.append(ln)

        # 요약문 추출 (중복 방지: 첫 번째 등장만 사용)
        summary_raw = ''
        for p in _SUMMARY_PATTERNS:
            m = re.search(p, content, re.DOTALL)
            if m:
                summary_raw = m.group(1).strip()
                break
        summary = _clean_but_keep_bold(summary_raw)

        sections = _extract_sections(content, title)

        return {
            'number':        num,
            'focus_keyword': focus,
            'sub_keywords':  sub,
            'title':         title,
            'toc':           toc,
            'summary':       summary,
            'sections':      sections,
            'raw_content':   content,
            'char_count':    len(content),
        }

    @staticmethod
    def _is_valid(article):
        return bool(article.get('title') or article.get('focus_keyword'))
