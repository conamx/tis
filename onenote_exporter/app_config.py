# app_config.py
"""설정 파일(config.json)과 폴더 위치 관리.

exe로 빌드하면 config.json / projects / output / 사용기록.csv 는 exe 옆에 생긴다.
"""
import json
import os
import shutil
import sys

# 사용자 파일이 놓이는 폴더 (exe 옆 또는 이 소스 폴더)
if getattr(sys, 'frozen', False):
    APP_DIR = os.path.dirname(os.path.abspath(sys.executable))
else:
    APP_DIR = os.path.dirname(os.path.abspath(__file__))

# exe 안에 묶인 기본 리소스 위치
BUNDLE_DIR = getattr(sys, '_MEIPASS', APP_DIR)

CONFIG_PATH   = os.path.join(APP_DIR, 'config.json')
PROJECTS_DIR  = os.path.join(APP_DIR, 'projects')
OUTPUT_DIR    = os.path.join(APP_DIR, 'output')
USAGE_LOG     = os.path.join(APP_DIR, '사용기록.csv')
UPLOAD_XLSX   = os.path.join(APP_DIR, '업로드기록.xlsx')   # 로톡 링크 기록 (사건번호 | 링크 | 제목)
LOG_DIR       = os.path.join(APP_DIR, 'logs')     # 날짜별 작업 기록 (밤새 돌린 결과 확인용)
# 두 번째 Claude 계정의 로그인 정보가 저장되는 폴더 (첫 번째 계정은 Claude Code 기본 위치)
ACCOUNT2_DIR  = os.path.join(APP_DIR, 'claude_account2')
# 토큰으로 연결한 계정은 계정마다 따로 쓰는 작업 폴더 (로그인 정보끼리 절대 안 섞이게)
def account_dir(n):
    return os.path.join(APP_DIR, f'claude_account{n}')

DEFAULTS = {
    'claude_model': '',         # 비우면 Claude Code 기본값. 예: sonnet / opus
    'claude_cli_path': '',      # 비우면 자동 찾기
    'claude_timeout_min': 30,   # 사례 1건(원고 10편) 최대 대기 시간
    'excel_path': '',
    'excel_used_check': True,   # 엑셀 '사용체크' 열에 표시가 있으면 쓴 사례로 본다
    'auto_export': True,        # 원고 생성 후 바로 OneNote로 내보내기
    'account': 1,               # 원고 작성에 쓸 Claude 계정 (1 또는 2)
    'last_project': '',
    'lawtalk_id': '',           # 링크 기록 엑셀의 '아이디' 칸
    'link_hide_prefix': '사용',  # 링크 기록 창에서 이 글자로 시작하는 제목은 숨김 (쉼표로 여러 개)
    'title_strip_result': False,  # 링크 기록 제목에서 끝의 결과 괄호 빼기
    'upload_xlsx': '',
    'link_strip_dash': True,     # 링크 넣을 때 OneNote 제목 앞 'ㅡ' 지우기          # 링크 기록 엑셀 (비우면 exe 옆 업로드기록.xlsx)
    'token_1': '',              # 계정 1 장기 토큰 (claude setup-token) — config.json은 남에게 주지 말 것
    'token_2': '',
}

CONFIG = dict(DEFAULTS)


def load_config():
    CONFIG.clear()
    CONFIG.update(DEFAULTS)
    if os.path.exists(CONFIG_PATH):
        try:
            with open(CONFIG_PATH, encoding='utf-8') as f:
                CONFIG.update(json.load(f))
        except Exception as e:
            print(f'[설정] config.json 읽기 실패: {e}')
    return CONFIG


def save_config():
    with open(CONFIG_PATH, 'w', encoding='utf-8') as f:
        json.dump(CONFIG, f, ensure_ascii=False, indent=2)


def ensure_projects_dir():
    """projects 폴더가 없으면 exe에 묶인 기본 프로젝트를 복사해 둔다."""
    if os.path.isdir(PROJECTS_DIR):
        return
    src = os.path.join(BUNDLE_DIR, 'projects')
    if os.path.isdir(src) and os.path.abspath(src) != os.path.abspath(PROJECTS_DIR):
        shutil.copytree(src, PROJECTS_DIR)
    else:
        os.makedirs(PROJECTS_DIR, exist_ok=True)


def open_folder(path):
    os.makedirs(path, exist_ok=True)
    if sys.platform.startswith('win'):
        os.startfile(path)  # noqa
    else:
        import subprocess
        subprocess.Popen(['xdg-open', path])


def open_file(path):
    """파일을 기본 프로그램(엑셀 등)으로 연다. 없으면 False"""
    if not os.path.exists(path):
        return False
    if sys.platform.startswith('win'):
        os.startfile(path)  # noqa
    else:
        import subprocess
        subprocess.Popen(['xdg-open', path])
    return True
