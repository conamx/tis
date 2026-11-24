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
    def __init__(self, msg, partial=''):
        super().__init__(msg)
        self.partial = partial   # 실패 전까지 써진 글 (버리지 않고 저장용)


class LimitError(AIError):
    """사용량 한도. resets_at = 한도가 풀리는 시각(유닉스 초), 모르면 None"""
    def __init__(self, msg, resets_at=None, partial=''):
        super().__init__(msg, partial)
        self.resets_at = resets_at


# 계정별 마지막 사용량 정보 (Claude Code가 알려주는 값)
#   {'five_hour': (사용률 0~1, 초기화 시각), 'seven_day': (...), 'status': 'allowed'/'rejected', 'resetsAt': ...}
USAGE = {}


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


def _account_env(account):
    env = os.environ.copy()
    # API 키가 환경변수에 있으면 구독 대신 API로 과금되므로 제거
    env.pop('ANTHROPIC_API_KEY', None)
    env.pop('ANTHROPIC_AUTH_TOKEN', None)
    # 계정 2: 로그인 정보를 별도 폴더에 두고 그 폴더로 실행
    if str(account or 1) == '2':
        os.makedirs(ACCOUNT2_DIR, exist_ok=True)
        env['CLAUDE_CONFIG_DIR'] = ACCOUNT2_DIR
    else:
        env.pop('CLAUDE_CONFIG_DIR', None)
    return env


def check_login(account):
    """사용량을 쓰지 않고 로그인 상태만 확인. (로그인됨 여부, 표시용 문구)"""
    exe = find_claude_cli()
    if not exe:
        return False, 'Claude Code 없음'
    try:
        r = subprocess.run([exe, 'auth', 'status', '--json'], capture_output=True, text=True,
                           encoding='utf-8', errors='replace', timeout=30,
                           env=_account_env(account), cwd=tempfile.gettempdir(),
                           creationflags=getattr(subprocess, 'CREATE_NO_WINDOW', 0))
        data = json.loads(r.stdout or '{}')
    except Exception:
        return False, '확인 실패'
    if not data.get('loggedIn'):
        return False, '로그인 필요'
    acc = data.get('account') if isinstance(data.get('account'), dict) else {}
    who = data.get('email') or acc.get('email')
    return True, (who or '로그인됨')


def call_claude_cli(prompt, stop_flag=None, timeout=None, system_file=None,
                    model=None, effort=None, web_search=False, account=None, progress=None):
    """`claude -p` 호출. 프롬프트는 stdin으로 전달(명령줄 길이 제한 회피).
    --tools "" : 파일 읽기/명령 실행 없이 글만 쓰게 함.
    system_file: Claude Code 기본(코딩용) 시스템 프롬프트를 이 파일 내용(지침)으로 교체
    progress(stage, text): 작성 중인 글을 실시간으로 받는 함수 (stage: 'think' / 'write')"""
    import queue
    import threading
    exe = find_claude_cli()
    if not exe:
        raise AIError('Claude Code를 찾지 못했습니다. 설정 → [Claude Code 설치하기]를 누르세요.')
    if timeout is None:
        timeout = int(CONFIG.get('claude_timeout_min') or 30) * 60
    cmd = [exe, '-p', '--output-format', 'stream-json', '--verbose',
           '--include-partial-messages', '--no-session-persistence',
           '--strict-mcp-config']   # PC에 깔린 MCP 서버 도구 설명이 따라붙지 않게
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

    env = _account_env(account)
    # 원고 10편은 길어서 기본 출력 한도에 잘릴 수 있음 → 한도를 넉넉히
    env.setdefault('CLAUDE_CODE_MAX_OUTPUT_TOKENS', '64000')
    flags = getattr(subprocess, 'CREATE_NO_WINDOW', 0)  # 윈도우에서 검은 창 안 뜨게

    proc = subprocess.Popen(cmd, stdin=subprocess.PIPE, stdout=subprocess.PIPE,
                            stderr=subprocess.PIPE, text=True, encoding='utf-8',
                            errors='replace', cwd=tempfile.gettempdir(), env=env,
                            creationflags=flags)
    lines, err_buf = queue.Queue(), []

    def _read_out():
        for line in proc.stdout:
            lines.put(line)
        lines.put(None)
    threading.Thread(target=_read_out, daemon=True).start()
    threading.Thread(target=lambda: err_buf.append(proc.stderr.read()), daemon=True).start()
    try:
        proc.stdin.write(prompt)
        proc.stdin.close()
    except OSError:
        pass

    started = last_cb = time.time()
    text, result, stage, raw_tail, limit = [], None, 'think', [], {}
    while True:
        try:
            line = lines.get(timeout=1)
        except queue.Empty:
            line = ''
        if stop_flag and stop_flag.is_set():
            proc.kill()
            raise AIError('사용자 중지', ''.join(text))
        if time.time() - started > timeout:
            proc.kill()
            raise AIError(f'Claude Code 응답 시간 초과 ({timeout // 60}분)', ''.join(text))
        if line is None:
            break
        if line:
            raw_tail = (raw_tail + [line])[-5:]
            try:
                ev = json.loads(line)
            except json.JSONDecodeError:
                continue
            t = ev.get('type')
            if t == 'stream_event':
                e = ev.get('event', {})
                if e.get('type') == 'content_block_delta':
                    d = e.get('delta', {})
                    if d.get('type') == 'text_delta':
                        text.append(d.get('text', ''))
                        stage = 'write'
            elif t == 'result':
                result = ev
            elif t == 'rate_limit_event':
                limit = ev.get('rate_limit_info') or {}
                win = limit.get('unifiedWindows') or {}
                USAGE[str(account or 1)] = {
                    'status': limit.get('status'), 'resetsAt': limit.get('resetsAt'),
                    **{k: (v.get('utilization'), v.get('resetsAt')) for k, v in win.items()
                       if isinstance(v, dict)}}
        if progress and time.time() - last_cb >= 1:
            last_cb = time.time()
            try:
                progress(stage, ''.join(text))
            except Exception:
                pass
    proc.wait(timeout=10)

    rejected = limit.get('status') == 'rejected'
    if result is None:
        msg = ((err_buf[0] if err_buf else '') or ''.join(raw_tail)).strip()[:300]
        if rejected or 'limit' in msg.lower():
            raise LimitError(f'사용량 한도에 걸렸습니다', limit.get('resetsAt'), ''.join(text))
        raise AIError(f'Claude Code 실행 실패 (코드 {proc.returncode}): {msg or "출력 없음"}',
                      ''.join(text))
    out = result.get('result') or ''.join(text)
    if result.get('is_error') or result.get('subtype') != 'success' or not out.strip():
        hint = ''
        low = out.lower()
        if rejected or ('limit' in low and ('usage' in low or 'reset' in low or 'hit' in low)):
            raise LimitError('사용량 한도에 걸렸습니다', limit.get('resetsAt'), ''.join(text))
        if 'login' in low or 'auth' in low:
            hint = ' → 로그인이 필요합니다'
        elif 'limit' in low:
            hint = ' → 사용량 한도에 걸렸습니다. 나중에 다시 시도하세요'
        raise AIError(f'Claude Code 오류: {out[:200] or result.get("subtype")}{hint}',
                      ''.join(text))
    return out


def generate(project, case_text, stop_flag=None, account=None, progress=None, single=None,
             extra=''):
    """프로젝트 지침 + 사건자료 → 원고 원문(`# 1` ~ `# N`).

    claude.ai 프로젝트와 같은 구조로 보낸다:
      시스템 프롬프트 = 지침 + 프로젝트 소스 파일 (Claude Code 기본 프롬프트는 쓰지 않음)
      사용자 메시지   = 사건자료만 (프로젝트 채팅창에 붙여넣던 그대로)"""
    system = project.build_system_prompt(single=single)
    s = project.settings
    opts = dict(model=(s.get('모델') or '').strip() or None,
                effort=(s.get('effort') or '').strip() or None,
                web_search=bool(s.get('웹검색')), account=account, progress=progress)
    fd, sys_path = tempfile.mkstemp(prefix='onenote_guide_', suffix='.md')
    try:
        with os.fdopen(fd, 'w', encoding='utf-8') as f:
            f.write(system)
        try:
            msg = case_text.strip() + (('\n\n' + extra.strip()) if extra else '')
            text = call_claude_cli(msg, stop_flag=stop_flag, system_file=sys_path, **opts)
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


def rewrite_one(project, case_text, num, old_article, other_titles, problems=None,
                stop_flag=None, account=None, progress=None):
    """원고 1편만 다시 쓰기. 같은 사건의 다른 원고 제목을 알려줘서 겹치지 않게 한다.
    10편을 통째로 다시 쓰는 것보다 사용량이 약 1/10."""
    old_title = (old_article or '').strip().split('\n')[0]
    extra = [f'[다시 쓰기 요청] 위 사건으로 이미 원고 {project.settings.get("편수", 10)}편을 썼다. '
             f'그중 # {num} 원고 1편만 새로 써라. 출력은 `# {num}` 줄로 시작하는 원고 1편뿐이다.',
             f'- 기존 # {num} 제목: {old_title}',
             '- 제목 접두어 종류(전문변호사/변호사)와 원고 유형은 기존 # {0}과 같게 유지하고, '
             '내용·구성·표현은 새로 쓴다.'.format(num)]
    if problems:
        extra.append('- 기존 원고의 문제(반드시 고칠 것): ' + ' / '.join(problems))
    if other_titles:
        extra.append('- 다른 원고 제목(겹치지 않게): ' + ' | '.join(other_titles))
    text = generate(project, case_text, stop_flag=stop_flag, account=account,
                    progress=progress, single=num, extra='\n'.join(extra))
    m = re.search(r'(?m)^#[ \t]+\d+[ \t]*$', text)
    body = text[m.end():] if m else text
    return f'# {num}\n' + body.strip() + '\n'
