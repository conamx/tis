# lawtalk.py
"""로톡용 복사: 원고를 대제목(h2)·소제목(h3)이 들어간 HTML로 만들어 클립보드에 올린다.
OneNote를 거치지 않으므로 로톡에 바로 붙여넣으면 제목 구조가 그대로 들어간다.
(로톡 대제목·소제목 변환기가 하던 일을 매크로가 대신함)"""
import html as _html
import re

from onenote_com import HEADING_STYLE, HEADING_FONT_PT, SIZE_BODY, _inline_runs


def _inline(text, force_bold=False):
    out = []
    for t, b, i, _ in _inline_runs(text, force_bold=force_bold):
        s = _html.escape(t, quote=False)
        if i:
            s = f'<i>{s}</i>'
        if b:
            s = f'<b>{s}</b>'
        out.append(s)
    return ''.join(out)


def _plain(text):
    return re.sub(r'\*\*|\*', '', text)


def article_html(art, keyword='', footers=None):
    """→ (HTML 조각, 일반 텍스트). 제목(페이지 제목)은 포함하지 않음 — 로톡 제목 칸에 따로 넣음"""
    h, p = [], []
    if keyword:
        h.append(f'<p><b>{_html.escape(keyword, quote=False)}</b></p>')
        p.append(keyword)
    blocks = art.get('blocks')
    if blocks is None:   # 붙여넣기로 가져온 예전 형식 글: 문단만
        for para in re.split(r'\n\s*\n', art.get('raw_content', '')):
            if para.strip():
                h.append(f'<p>{_inline(para.strip())}</p>'); p.append(_plain(para.strip()))
    else:
        bullets = []

        def flush():
            if bullets:
                h.append('<ul>' + ''.join(f'<li>{_inline(x)}</li>' for x in bullets) + '</ul>')
                bullets.clear()

        for b in blocks:
            t = b.get('type')
            if t != 'bullet':
                flush()
            if t == 'blank':
                continue
            if t == 'heading':
                tag = HEADING_STYLE.get(b.get('level', 2), 'h3')
                h.append(f'<{tag}>{_html.escape(_plain(b["text"]), quote=False)}</{tag}>')
                p.append(_plain(b['text']))
            elif t == 'bullet':
                bullets.append(b['text']); p.append('• ' + _plain(b['text']))
            elif t == 'table':
                rows = []
                for ri, row in enumerate(b['rows']):
                    cells = [re.sub(r'\s*<br\s*/?>\s*', ' ', c, flags=re.I).strip() for c in row]
                    line = ' | '.join(cells)
                    bold = b.get('header') and ri == 0
                    rows.append(f'<p style="text-align:center">{_inline(line, force_bold=bold)}</p>')
                    p.append(_plain(line))
                h.extend(rows)
            else:
                h.append(f'<p>{_inline(b.get("text", ""))}</p>'); p.append(_plain(b.get('text', '')))
        flush()
    for preset in footers or []:
        for text, bold in preset:
            s = _html.escape(text, quote=False)
            h.append(f'<p><b>{s}</b></p>' if bold else f'<p>{s}</p>')
            p.append(text)
    return '\n'.join(h), '\n\n'.join(p)


def cf_html(fragment):
    """윈도우 클립보드 'HTML Format' 형식 (앞에 바이트 위치 머리글이 붙음)"""
    head = ('Version:0.9\r\nStartHTML:{0:010d}\r\nEndHTML:{1:010d}\r\n'
            'StartFragment:{2:010d}\r\nEndFragment:{3:010d}\r\n')
    pre = '<html><head><meta charset="utf-8"></head><body>\r\n<!--StartFragment-->'
    post = '<!--EndFragment-->\r\n</body></html>'
    dummy = head.format(0, 0, 0, 0)
    start_html = len(dummy.encode('utf-8'))
    start_frag = start_html + len(pre.encode('utf-8'))
    end_frag = start_frag + len(fragment.encode('utf-8'))
    end_html = end_frag + len(post.encode('utf-8'))
    return head.format(start_html, end_html, start_frag, end_frag) + pre + fragment + post


def copy_to_clipboard(fragment, text):
    """윈도우 클립보드에 HTML + 일반 텍스트를 함께 올린다. 성공하면 True"""
    try:
        import ctypes
        from ctypes import wintypes
    except Exception:
        return False
    if not hasattr(ctypes, 'windll'):
        return False
    u32, k32 = ctypes.windll.user32, ctypes.windll.kernel32
    k32.GlobalAlloc.restype = wintypes.HGLOBAL
    k32.GlobalAlloc.argtypes = [wintypes.UINT, ctypes.c_size_t]
    k32.GlobalLock.restype = ctypes.c_void_p
    k32.GlobalLock.argtypes = [wintypes.HGLOBAL]
    k32.GlobalUnlock.argtypes = [wintypes.HGLOBAL]
    u32.SetClipboardData.argtypes = [wintypes.UINT, wintypes.HANDLE]
    u32.SetClipboardData.restype = wintypes.HANDLE
    GMEM_MOVEABLE, CF_UNICODETEXT = 0x0002, 13

    def put(fmt, data):
        h = k32.GlobalAlloc(GMEM_MOVEABLE, len(data))
        ptr = k32.GlobalLock(h)
        ctypes.memmove(ptr, data, len(data))
        k32.GlobalUnlock(h)
        u32.SetClipboardData(fmt, h)

    html_fmt = u32.RegisterClipboardFormatW('HTML Format')
    if not u32.OpenClipboard(None):
        return False
    try:
        u32.EmptyClipboard()
        put(html_fmt, cf_html(fragment).encode('utf-8') + b'\0')
        put(CF_UNICODETEXT, text.encode('utf-16-le') + b'\0\0')
        return True
    finally:
        u32.CloseClipboard()
