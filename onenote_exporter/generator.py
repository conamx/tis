# generator.py
"""PC에 설치·로그인된 Claude Code(`claude -p`)로 원고 생성.

Claude 구독(Pro/Max) 사용량으로 처리되므로 API 키·별도 결제가 필요 없다.
(티스토리 매크로와 같은 방식)
"""
import json
import os
import shutil
import subprocess
import tempfile
import re
import time

from app_config import CONFIG


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


def call_claude_cli(prompt, stop_flag=None, timeout=None):
    """`claude -p` 호출. 프롬프트는 stdin으로 전달(명령줄 길이 제한 회피).
    --tools "" : 파일 읽기/명령 실행 없이 글만 쓰게 함."""
    exe = find_claude_cli()
    if not exe:
        raise AIError('Claude Code를 찾지 못했습니다. 설치 후 터미널에서 `claude`를 한 번 실행해 '
                      '로그인하거나, ⚙ 설정에 실행 파일 경로를 넣어주세요.')
    if timeout is None:
        timeout = int(CONFIG.get('claude_timeout_min') or 30) * 60
    cmd = [exe, '-p', '--output-format', 'json', '--tools', '', '--no-session-persistence']
    model = (CONFIG.get('claude_model') or '').strip()
    if model:
        cmd += ['--model', model]

    env = os.environ.copy()
    # API 키가 환경변수에 있으면 구독 대신 API로 과금되므로 제거
    env.pop('ANTHROPIC_API_KEY', None)
    env.pop('ANTHROPIC_AUTH_TOKEN', None)
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
    """프로젝트 지침 + 사건자료 → 원고 원문(`# 1` ~ `# N`)."""
    system = project.build_system_prompt()
    prompt = (
        '<지침>\n' + system + '\n</지침>\n\n'
        '아래 사건자료로 위 지침에 따라 원고를 작성하라.\n\n'
        '<사건자료>\n' + case_text.strip() + '\n</사건자료>'
    )
    text = call_claude_cli(prompt, stop_flag=stop_flag)
    # 혹시 원고 앞에 설명이 붙으면 첫 '# 1' 앞부분은 버린다
    m = re.search(r'(?m)^#[ \t]+1[ \t]*$', text)
    return text[m.start():] if m else text


def label():
    model = (CONFIG.get('claude_model') or '').strip()
    return f'Claude Code ({model})' if model else 'Claude Code'
