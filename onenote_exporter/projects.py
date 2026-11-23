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
    'effort': '',      # 비우면 기본. claude.ai '확장 사고'를 켜고 쓰던 프로젝트면 high
    '웹검색': False,   # claude.ai에서 웹 검색을 켜고 쓰던 프로젝트면 true (사용량 늘어남)
    '검사': {
        '원고당_정확히1회': [],
        '금지문구': [],
        '사건번호_노출금지': True,
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

    def build_system_prompt(self):
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
        parts.append(
            '---\n'
            '[매크로 실행 규칙]\n'
            '- 위 지침에서 언급하는 프로젝트 소스 파일은 위 <프로젝트_소스> 블록으로 제공된다. '
            '제공되지 않은 파일은 없는 것으로 보고 지침 본문 기준으로 작성한다.\n'
            '- 사건자료에 나오는 사람 이름(피의자·피고인·피해자·관련인 실명)은 원고에 절대 쓰지 않는다. '
            '필요하면 \'의뢰인\', \'상대방\' 등으로 바꾼다.\n'
            f'- 응답에는 `# 1`부터 `# {n}`까지의 원고만 출력한다. 각 번호 줄은 `# 숫자` 한 줄만 쓰고, '
            '바로 다음 줄이 제목이다. 앞뒤 인사·설명·검수표는 출력하지 않는다.'
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
    """설정.json의 '검사' 규칙으로 결과 원고를 점검해 경고 목록을 돌려준다 (자동 재생성은 하지 않음)."""
    s = project.settings
    rules = s.get('검사', {})
    want = int(s.get('편수', 10) or 0)
    warnings = []
    parts = re.split(r'(?m)^#[ \t]+(\d+)[ \t]*$', text)
    arts = [(parts[i], parts[i + 1]) for i in range(1, len(parts) - 1, 2)]
    if want and len(arts) != want:
        warnings.append(f'원고 수 {len(arts)}편 (목표 {want}편)')
    for num, body in arts:
        for phrase in rules.get('원고당_정확히1회', []):
            c = body.count(phrase)
            if c != 1:
                warnings.append(f'# {num}: "{phrase}" {c}회 (정확히 1회여야 함)')
        for phrase in rules.get('금지문구', []):
            if phrase and phrase in body:
                warnings.append(f'# {num}: 금지 문구 "{phrase}" 포함')
    if rules.get('사건번호_노출금지') and case_id and case_id in text:
        warnings.append(f'사건번호 {case_id} 가 원고에 노출됨')
    return warnings
