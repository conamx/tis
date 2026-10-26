# main.py
"""
OneNote 글 내보내기 — 진입점
사용법:
  python main.py                     # GUI 실행
  python main.py --input FILE.txt    # 파일을 자동으로 열어 GUI 실행
  python main.py --help              # 도움말

빌드:
  pyinstaller --onefile --noconsole --icon=icon.ico main.py
"""

import argparse
import os
import sys

# exe(PyInstaller) 환경에서 경로 기준점
BASE_DIR = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))


def _parse_args():
    parser = argparse.ArgumentParser(
        prog='onenote_exporter',
        description='Claude 응답 글들을 OneNote로 내보내는 GUI 프로그램',
    )
    parser.add_argument(
        '--input', '-i',
        metavar='FILE',
        default='',
        help='시작 시 자동으로 불러올 텍스트 파일 경로 (txt 또는 md)',
    )
    parser.add_argument(
        '--version', '-v',
        action='version',
        version='OneNote Exporter 2.0',
    )
    return parser.parse_args()


def _set_icon(root):
    """아이콘 설정 — icon.ico 우선, 없으면 icon.png 시도."""
    ico = os.path.join(BASE_DIR, 'icon.ico')
    png = os.path.join(BASE_DIR, 'icon.png')

    if os.path.exists(ico):
        try:
            root.iconbitmap(ico)
            return
        except Exception:
            pass

    if os.path.exists(png):
        try:
            from PIL import Image, ImageTk
            img = ImageTk.PhotoImage(Image.open(png))
            root.iconphoto(True, img)
            root._icon_ref = img  # GC 방지
            return
        except Exception:
            pass

    # 둘 다 없으면 기본 아이콘 유지


def main():
    args = _parse_args()

    try:
        import tkinter as tk
    except Exception as e:
        print(f'[오류] GUI를 실행할 수 없습니다: {e}', file=sys.stderr)
        print('Python tkinter가 설치되어 있는지 확인하세요.', file=sys.stderr)
        sys.exit(1)

    root = tk.Tk()
    _set_icon(root)

    from ui import run
    run(root=root, input_file=args.input)


if __name__ == '__main__':
    main()
