# generator.py
"""PC에 설치·로그인된 Claude Code(`claude -p`)로 원고 생성. (티스토리 매크로와 같은 방식)

계정 2개: 계정 1은 Claude Code 기본 로그인, 계정 2는 claude_account2 폴더에 따로 로그인.
"""
import json
import os
import shutil
import subprocess
import tempfile
import re
import time

from app_config import CONFIG, ACCOUNT2_DIR


class AIError(Exception):
    pass


def find_claude_cli():
    """설치된 Claude Code 실행 파일 경로 (없으면 None)"""
    custom = (CONFIG.get('claude_cli_path') or '').strip()
    if custom:
        return custom if os.path.exists(custom) else None
    found = shutil.which('claude')
    if found:
        return found
    home = os.path.expanduser('~')
    for p in [os.path.join(home, '.local', 'bin', 'claude.exe'),
              os.path.join(home, '.local', 'bin', 'claude'),
              os.path.join(os.environ.get('APPDATA', ''), 'npm', 'claude.cmd'),
              os.path.join(home, '.claude', 'local', 'claude')]:
        if os.path.exists(p):
            return p
    return None


def call_claude_cli(prompt, stop_flag=None, timeout=None, system_file=None,
                    model=None, effort=None, web_search=False, account=None):
    """`claude -p` 호출. 프롬프트는 stdin으로 전달(명령줄 길이 제한 회피).
    --tools "" : 파일 읽기/명령 실행 없이 글만 쓰게 함.
    system_file: Claude Code 기본(코딩용) 시스템 프롬프트를 이 파일 내용(지침)으로 교체
                 → claude.ai 프로젝트에 지침을 넣은 것과 같은 구조, 사용량도 줄어듦."""
    exe = find_claude_cli()
    if not exe:
        raise AIError('Claude Code를 찾지 못했습니다. 설치 후 터미널에서 `claude`를 한 번 실행해 '
                      '로그인하거나, ⚙ 설정에 실행 파일 경로를 넣어주세요.')
    if timeout is None:
        timeout = int(CONFIG.get('claude_timeout_min') or 30) * 60
    cmd = [exe, '-p', '--output-format', 'json', '--no-session-persistence']
    if web_search:
        cmd += ['--tools', 'WebSearch', '--allowedTools', 'WebSearch']
    else:
        cmd += ['--tools', '']
    model = (model or CONFIG.get('claude_model') or '').strip()
    if model:
        cmd += ['--model', model]
    if effort:
        cmd += ['--effort', effort]
    if system_file:
        cmd += ['--system-prompt-file', system_file]

    env = os.environ.copy()
    # API 키가 환경변수에 있으면 구독 대신 API로 과금되므로 제거
    env.pop('ANTHROPIC_API_KEY', None)
    env.pop('ANTHROPIC_AUTH_TOKEN', None)
    # 계정 2: 로그인 정보를 별도 폴더에 두고 그 폴더로 실행
    if int(account or CONFIG.get('account') or 1) == 2:
        os.makedirs(ACCOUNT2_DIR, exist_ok=True)
        env['CLAUDE_CONFIG_DIR'] = ACCOUNT2_DIR
    else:
        env.pop('CLAUDE_CONFIG_DIR', None)
    # 원고 10편은 길어서 기본 출력 한도에 잘릴 수 있음 → 한도를 넉넉히
    env.setdefault('CLAUDE_CODE_MAX_OUTPUT_TOKENS', '64000')
    flags = getattr(subprocess, 'CREATE_NO_WINDOW', 0)  # 윈도우에서 검은 창 안 뜨게

    proc = subprocess.Popen(cmd, stdin=subprocess.PIPE, stdout=subprocess.PIPE,
                            stderr=subprocess.PIPE, text=True, encoding='utf-8',
                            errors='replace', cwd=tempfile.gettempdir(), env=env,
                            creationflags=flags)
    started = time.time()
    pending_input = prompt
    while True:
        try:
            out, err = proc.communicate(input=pending_input, timeout=2)
            break
        except subprocess.TimeoutExpired:
            pending_input = None  # stdin은 첫 호출에서 이미 전달됨
            if stop_flag and stop_flag.is_set():
                proc.kill(); proc.communicate()
                raise AIError('사용자 중지')
            if time.time() - started > timeout:
                proc.kill(); proc.communicate()
                raise AIError(f'Claude Code 응답 시간 초과 ({timeout // 60}분)')

    try:
        data = json.loads(out)
    except (json.JSONDecodeError, TypeError):
        msg = (err or out or '').strip()[:300]
        raise AIError(f'Claude Code 실행 실패 (코드 {proc.returncode}): {msg or "출력 없음"}')

    result = data.get('result') or ''
    if data.get('is_error') or data.get('subtype') != 'success' or not result.strip():
        hint = ''
        low = result.lower()
        if 'login' in low or 'auth' in low:
            hint = ' → 터미널에서 `claude` 실행 후 /login 으로 로그인하세요'
        elif 'limit' in low:
            hint = ' → 구독 사용량 한도에 걸렸습니다. 나중에 다시 시도하세요'
        raise AIError(f'Claude Code 오류: {result[:200] or data.get("subtype")}{hint}')
    return result


def generate(project, case_text, stop_flag=None):
    """프로젝트 지침 + 사건자료 → 원고 원문(`# 1` ~ `# N`).

    claude.ai 프로젝트와 같은 구조로 보낸다:
      시스템 프롬프트 = 지침 + 프로젝트 소스 파일 (Claude Code 기본 프롬프트는 쓰지 않음)
      사용자 메시지   = 사건자료만 (프로젝트 채팅창에 붙여넣던 그대로)"""
    system = project.build_system_prompt()
    s = project.settings
    opts = dict(model=(s.get('모델') or '').strip() or None,
                effort=(s.get('effort') or '').strip() or None,
                web_search=bool(s.get('웹검색')))
    fd, sys_path = tempfile.mkstemp(prefix='onenote_guide_', suffix='.md')
    try:
        with os.fdopen(fd, 'w', encoding='utf-8') as f:
            f.write(system)
        try:
            text = call_claude_cli(case_text.strip(), stop_flag=stop_flag,
                                   system_file=sys_path, **opts)
        except AIError as e:
            # 아주 오래된 Claude Code라 옵션을 모르면 예전 방식(지침을 메시지에 같이)으로 재시도
            if 'unknown option' not in str(e).lower():
                raise
            prompt = ('<지침>\n' + system + '\n</지침>\n\n'
                      '아래 사건자료로 위 지침에 따라 원고를 작성하라.\n\n'
                      '<사건자료>\n' + case_text.strip() + '\n</사건자료>')
            text = call_claude_cli(prompt, stop_flag=stop_flag, **opts)
    finally:
        try:
            os.remove(sys_path)
        except OSError:
            pass
    # 혹시 원고 앞에 설명이 붙으면 첫 '# 1' 앞부분은 버린다
    m = re.search(r'(?m)^#[ \t]+1[ \t]*$', text)
    return text[m.start():] if m else text


def open_login(account):
    """해당 계정으로 Claude Code 로그인 창을 띄운다."""
    import sys
    if not sys.platform.startswith('win'):
        raise AIError('Windows에서만 지원합니다.')
    exe = find_claude_cli()
    if not exe:
        raise AIError('Claude Code가 설치되어 있지 않습니다. 먼저 [Claude Code 설치하기]를 누르세요.')
    pre = ''
    if int(account) == 2:
        os.makedirs(ACCOUNT2_DIR, exist_ok=True)
        pre = "$env:CLAUDE_CONFIG_DIR = '" + ACCOUNT2_DIR.replace("'", "''") + "'; "
    script = (pre +
              f"Write-Host '계정 {account} 로그인: 브라우저가 열리면 계정 {account}로 로그인하세요.' -ForegroundColor Green; "
              "Write-Host '이미 로그인된 상태면 /login 을 입력해 계정을 바꿀 수 있습니다. 끝나면 창을 닫으세요.' -ForegroundColor Green; "
              "& '" + exe.replace("'", "''") + "'")
    subprocess.Popen(['powershell', '-NoExit', '-ExecutionPolicy', 'Bypass', '-Command', script],
                     creationflags=getattr(subprocess, 'CREATE_NEW_CONSOLE', 0))


def open_installer():
    """Claude Code 설치 + 로그인 창을 띄운다 (Windows PowerShell).
    설치가 끝나면 같은 창에서 claude 가 실행되어 브라우저 로그인 화면이 열린다."""
    import sys
    if not sys.platform.startswith('win'):
        raise AIError('Windows에서만 자동 설치를 지원합니다.')
    script = (
        "Write-Host 'Claude Code 설치 중... (1~2분)' -ForegroundColor Cyan; "
        "irm https://claude.ai/install.ps1 | iex; "
        "$exe = Join-Path $env:USERPROFILE '.local\\bin\\claude.exe'; "
        "if (-not (Test-Path $exe)) { $exe = 'claude' }; "
        "Write-Host ''; "
        "Write-Host '설치 완료! 이어서 로그인합니다. 브라우저가 열리면 Claude 계정으로 로그인하세요.' -ForegroundColor Green; "
        "Write-Host '로그인 후 이 창은 닫아도 됩니다. 매크로에서 [Claude 연결 테스트]를 눌러 확인하세요.' -ForegroundColor Green; "
        "& $exe"
    )
    subprocess.Popen(['powershell', '-NoExit', '-ExecutionPolicy', 'Bypass', '-Command', script],
                     creationflags=getattr(subprocess, 'CREATE_NEW_CONSOLE', 0))


def label():
    model = (CONFIG.get('claude_model') or '').strip()
    return f'Claude Code ({model})' if model else 'Claude Code'
