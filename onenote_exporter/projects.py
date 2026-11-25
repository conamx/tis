# projects.py
"""프로젝트 = projects/ 아래 폴더 하나.

  projects/니케_로톡성공사례/
    지침.md       ← 메인 지침 (claude.ai 프로젝트 지침을 그대로 붙여넣기)
    *.md / *.txt  ← 세부지침·참고자료 (프로젝트 소스 파일). 폴더에 넣기만 하면 같이 읽음
    설정.json     ← 편수, 접두어, 하단 프리셋, 엑셀 사용 여부, 자동 검사 규칙

지침을 바꾸려면 파일만 고치면 된다. 프로그램은 매번 생성할 때마다 새로 읽는다.
"""
import json
import os
import re

from app_config import PROJECTS_DIR

MAIN_GUIDE = '지침.md'
SETTINGS   = '설정.json'
TEXT_EXT   = ('.md', '.txt')

DEFAULT_SETTINGS = {
    '편수': 10,
    '접두어': '',
    '상단키워드': '',
    '하단프리셋': [],
    '엑셀사용': True,
    '입력안내': '사건자료(또는 주제·키워드)를 입력하세요.',
    '모델': '',        # 비우면 ⚙ 설정의 모델. claude.ai 프로젝트에서 쓰던 모델과 맞추기 (예: opus / sonnet)
    'effort': 'medium',  # 생각 깊이 low/medium/high. medium이 품질 유지 + 사용량 절약 (실측 기준)
    '웹검색': False,   # claude.ai에서 웹 검색을 켜고 쓰던 프로젝트면 true (사용량 늘어남)
    '검사': {
        '원고당_정확히1회': [],
        '금지문구': [],
        '사건번호_노출금지': True,
        '제목_최대글자': 0,     # 0이면 검사 안 함. 마지막 결과 괄호(와 그 앞 공백)는 빼고 셈
        '본문_최소글자': 0,     # 0이면 검사 안 함. 제목·소제목·표·목록·해시태그·기호·공백 제외
    },
}

NEW_GUIDE_TEMPLATE = """# (프로젝트 이름) 원고 생성 지침

여기에 claude.ai 프로젝트의 "지침" 내용을 그대로 붙여넣으세요.

## 출력 형식 (OneNote 내보내기용 — 지우지 마세요)
- 원고는 `# 1`, `# 2` ... 처럼 번호만 있는 줄로 구분한다.
- 각 번호 아래 첫 줄은 제목, 그 다음 줄부터 본문.
- 소제목은 `## 소제목`, 목록은 `- 항목`, 표는 `| 칸 | 칸 |` 형식.
- 원고 외의 인사·설명·검수 결과는 출력하지 않는다.
"""


class Project:
    def __init__(self, folder):
        self.folder = folder
        self.name = os.path.basename(folder)

    @property
    def settings(self):
        s = json.loads(json.dumps(DEFAULT_SETTINGS))
        p = os.path.join(self.folder, SETTINGS)
        if os.path.exists(p):
            with open(p, encoding='utf-8') as f:
                user = json.load(f)
            checks = user.pop('검사', None)
            s.update(user)
            if isinstance(checks, dict):
                s['검사'].update(checks)
        return s

    def guide_files(self):
        """지침.md 먼저, 나머지 텍스트 파일은 이름순."""
        files = [f for f in os.listdir(self.folder)
                 if f.lower().endswith(TEXT_EXT) and not f.startswith(('~', '.'))]
        files.sort(key=lambda f: (f != MAIN_GUIDE, f))
        return [os.path.join(self.folder, f) for f in files]

    def build_system_prompt(self, single=None):
        files = self.guide_files()
        if not files:
            raise RuntimeError(f"'{self.name}' 폴더에 지침 파일(.md/.txt)이 없습니다.")
        parts = []
        for path in files:
            text = _read_text(path).strip()
            if not text:
                continue
            fname = os.path.basename(path)
            if fname == MAIN_GUIDE:
                parts.append(text)
            else:
                parts.append(f'<프로젝트_소스 파일명="{fname}">\n{text}\n</프로젝트_소스>')
        n = self.settings.get('편수', 10)
        mn = int((self.settings.get('검사') or {}).get('본문_최소글자') or 0)
        if mn:
            # Claude는 글자 수를 정확히 못 세고 공백까지 세는 경향 → 문장·문단 수로 바꿔서 지시 (실측: 문장 1개 ≈ 공백 제외 42자)
            sents = -(-int(mn * 1.12) // 40)          # 1500자 → 42문장
            paras = -(-sents // 3.5)
            length_rule = (
                f'- 분량(중요): 원고마다 제목·소제목·표·목록·해시태그를 뺀 서술문이 공백 제외 {mn:,}자 이상이어야 한다. '
                f'공백을 포함하면 약 {int(mn * 1.4):,}자 이상이다. 글자 수를 세지 말고 다음 기준으로 쓴다: '
                f'서술 문단 {int(paras)}개 이상, 문단마다 3~4문장, 원고당 최소 {sents}문장 '
                '(도입·각 소제목 본문·Q&A 답변·마무리 문단의 문장을 모두 합쳐서). '
                '부족하면 사건 경위, 해당 법적 기준의 의미, 처분의 의미, 독자가 확인할 점을 더 풀어 쓴다. '
                '근거 없는 사실로 채우지는 않는다.\n')
        else:
            length_rule = ''
        out_rule = (f'- 이번 응답에는 `# {single}` 원고 1편만 출력한다. 번호 줄은 `# {single}` 한 줄, '
                    '바로 다음 줄이 제목이다. 앞뒤 인사·설명·검수표는 출력하지 않는다.') if single else (
                    f'- 응답에는 `# 1`부터 `# {n}`까지의 원고만 출력한다. 각 번호 줄은 `# 숫자` 한 줄만 쓰고, '
                    '바로 다음 줄이 제목이다. 앞뒤 인사·설명·검수표는 출력하지 않는다.')
        parts.append(
            '---\n'
            '[매크로 실행 규칙]\n'
            '- 위 지침에서 언급하는 프로젝트 소스 파일은 위 <프로젝트_소스> 블록으로 제공된다. '
            '제공되지 않은 파일은 없는 것으로 보고 지침 본문 기준으로 작성한다.\n'
            + length_rule
            + '- 사건자료에 나오는 사람 이름(피의자·피고인·피해자·관련인 실명)은 원고에 절대 쓰지 않는다. '
            '필요하면 \'의뢰인\', \'상대방\' 등으로 바꾼다.\n'
            + out_rule
        )
        return '\n\n'.join(parts)


def _read_text(path):
    for enc in ('utf-8-sig', 'utf-8', 'cp949'):
        try:
            with open(path, encoding=enc) as f:
                return f.read()
        except UnicodeDecodeError:
            continue
    return ''


def list_projects():
    if not os.path.isdir(PROJECTS_DIR):
        return []
    out = []
    for d in sorted(os.listdir(PROJECTS_DIR)):
        full = os.path.join(PROJECTS_DIR, d)
        if os.path.isdir(full) and not d.startswith(('.', '_')):
            out.append(Project(full))
    return out


def create_project(name):
    name = re.sub(r'[\\/:*?"<>|]', '_', name.strip())
    if not name:
        raise ValueError('프로젝트 이름을 입력하세요.')
    folder = os.path.join(PROJECTS_DIR, name)
    if os.path.exists(folder):
        raise ValueError('같은 이름의 프로젝트가 이미 있습니다.')
    os.makedirs(folder)
    with open(os.path.join(folder, MAIN_GUIDE), 'w', encoding='utf-8') as f:
        f.write(NEW_GUIDE_TEMPLATE)
    with open(os.path.join(folder, SETTINGS), 'w', encoding='utf-8') as f:
        json.dump(DEFAULT_SETTINGS, f, ensure_ascii=False, indent=2)
    return Project(folder)


def check_output(text, project, case_id=''):
    """설정.json의 '검사' 규칙으로 원고를 점검. 반환: [(원고번호 또는 None, 문제), ...]
    (원고번호가 있는 문제는 그 1편만 다시 쓰면 고쳐진다)"""
    s = project.settings
    rules = s.get('검사', {})
    want = int(s.get('편수', 10) or 0)
    issues = []
    arts = split_articles(text)
    if want and len(arts) != want:
        issues.append((None, f'원고 수 {len(arts)}편 (목표 {want}편)'))
    for num, body in arts:
        for phrase in rules.get('원고당_정확히1회', []):
            c = body.count(phrase)
            if c != 1:
                issues.append((num, f'"{phrase}" {c}회 (정확히 1회여야 함)'))
        for phrase in rules.get('금지문구', []):
            if phrase and phrase in body:
                issues.append((num, f'금지 문구 "{phrase}" 포함'))
        if rules.get('사건번호_노출금지') and case_id and case_id in body:
            issues.append((num, '사건번호가 원고에 노출됨'))
        mx = int(rules.get('제목_최대글자') or 0)
        if mx:
            tl = title_length(body.strip().split('\n')[0])
            if tl > mx:
                issues.append((num, f'제목 {tl}자 (결과 괄호 빼고 {mx}자 이하여야 함)'))
        mn = int(rules.get('본문_최소글자') or 0)
        if mn:
            bc = body_chars(body)
            if bc < mn:
                # 기준보다 15% 넘게 짧을 때만 다시 쓰기 대상, 조금 모자라면 경고만 (사용량 절약)
                tag = '' if bc < mn * 0.85 else '[경고] '
                issues.append((num, f'{tag}본문 {bc:,}자 (최소 {mn:,}자)'))
    return issues


def title_length(title):
    """제목 글자 수 — 맨 끝 결과 괄호(안의 중첩 괄호 포함)와 그 앞 공백 1개는 뺀다."""
    t = title.strip()
    if t.endswith(')'):
        depth = 0
        for i in range(len(t) - 1, -1, -1):
            if t[i] == ')':
                depth += 1
            elif t[i] == '(':
                depth -= 1
                if depth == 0:
                    t = t[:i]
                    if t.endswith(' '):
                        t = t[:-1]
                    break
    return len(t)


def body_chars(body):
    """실제 서술문 글자 수 — 제목·소제목·표·목록·해시태그 줄, 마크다운 기호, 공백 제외"""
    lines = body.strip().split('\n')[1:]          # 첫 줄 = 제목
    keep = []
    for ln in lines:
        t = ln.strip()
        if not t or t.startswith(('#', '|', '- ', '* ', '• ')):
            continue
        keep.append(t)
    txt = re.sub(r'[*_`>]', '', ''.join(keep))
    return len(re.sub(r'\s+', '', txt))


def split_articles(text):
    """'# 1' ~ '# N' 원고 원문 → [(번호, 본문), ...]"""
    parts = re.split(r'(?m)^#[ \t]+(\d+)[ \t]*$', text)
    return [(parts[i], parts[i + 1]) for i in range(1, len(parts) - 1, 2)]


def join_articles(arts):
    return '\n\n'.join(f'# {n}\n{b.strip()}' for n, b in arts) + '\n'


def save_settings(project, updates):
    """화면에서 바꾼 값을 프로젝트 설정.json에 자동 저장 (다른 키는 그대로 둠)."""
    p = os.path.join(project.folder, SETTINGS)
    data = {}
    if os.path.exists(p):
        with open(p, encoding='utf-8') as f:
            data = json.load(f)
    changed = False
    for k, v in updates.items():
        if data.get(k) != v:
            data[k] = v
            changed = True
    if changed:
        with open(p, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)


def file_title(path, limit=30):
    """파일의 첫 줄(제목)을 이름으로 보여주기 위해 꺼낸다"""
    for ln in _read_text(path).splitlines():
        t = re.sub(r'^[#\s*]+', '', ln).strip().strip('*').strip()
        if t:
            return t if len(t) <= limit else t[:limit] + '…'
    return '(빈 파일)'


def rename_project(project, new_name):
    new_name = re.sub(r'[\\/:*?"<>|]', '_', (new_name or '').strip())
    if not new_name:
        raise ValueError('이름을 입력하세요.')
    dst = os.path.join(PROJECTS_DIR, new_name)
    if os.path.exists(dst):
        raise ValueError('같은 이름의 지침이 이미 있습니다.')
    os.rename(project.folder, dst)
    return Project(dst)


def delete_project(project):
    import shutil
    shutil.rmtree(project.folder)
