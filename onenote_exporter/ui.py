# ui.py
import tkinter as tk
from tkinter import scrolledtext, messagebox, filedialog, ttk
import os
import re
import threading
import time
from datetime import datetime
from parser import ArticleParser
import onenote_com
from onenote_com import OneNoteManager, SectionInfo
import app_config
from app_config import CONFIG
import projects as proj
import generator
import excel_source

FOOTER_PRESETS = [
    ('오토트러스트', [
        ('(오토트러스트 사진 1)', True),
        ('(오토트러스트 사진 2)', True),
        ('(오토트러스트 사진 3)', True),
        ('https://m.site.naver.com/1hhgJ', False),
        ('(오토트러스트 사진 4)', True),
        ('tel:042-623-5800', False),
    ]),
    ('대덕블로그', [
        ('지도', True),
        ('대덕자동차직업전문학교', False),
        ('http://www.dcaredu.com/', False),
    ]),
    ('파이브성범죄', [
        ('전화 예약', True),
        ('tel:010-9276-9874', False),
        ('네이버 예약', True),
        ('https://naver.me/5tf4735M', False),
        ('카카오톡', True),
        ('http://pf.kakao.com/_kGWxhn', False),
    ]),
    ('니케임대차', [
        ('니케 임대차소송센터', True),
        ('https://lknike.com/rent/', False),
        ('전화연결 바로하기', True),
        ('tel:1544-5943', False),
    ]),
    ('니케부동산', [
        ('전화연결 바로가기', True),
        ('tel:02-6953-6936', False),
    ]),
    ('니케성범죄', [
        ('법률사무소니케 메인홈페이지 (성범죄, 학폭)', True),
        ('https://m.site.naver.com/1OUY4', False),
        ('전화연결 바로가기', True),
        ('tel: 1644-5468', False),
    ]),
    ('성지성범죄', [
        ('성지파트너스 홈페이지', True),
        ('https://szpcrime.com/', False),
        ('형사전화연결 바로가기 ', True),
        ('tel:1551-7501', False),
        ('성공사례 확인하러 가기', True),
        ('https://blog.naver.com/sjpcrime', False),
    ]),
    ('오엔블로그', [
        ('전화연결 바로가기', True),
        ('tel:1660-2241', False),
        ('카카오톡 바로가기', True),
        ('https://pf.kakao.com/_vxjeRG', False),
        ('법률사무소오엔 홈페이지', True),
        ('https://onlaw.co.kr/onlaw/counsel/index.html', False),
    ]),
    ('진잠골프', [
        ('지도', True),
        ('진잠실내골프연습장', False),
        ('(12번 사진 삽입)', True),
        ('tel:0425439991', False),
    ]),
]

C = {
    'bg':      '#f4f5f7',
    'white':   '#ffffff',
    'header':  '#1e293b',
    'accent':  '#2563eb',
    'green':   '#16a34a',
    'red':     '#dc2626',
    'orange':  '#ea580c',
    'gray':    '#64748b',
    'light':   '#e2e8f0',
    'text':    '#0f172a',
    'subtext': '#475569',
    'log_bg':  '#1e293b',
    'log_fg':  '#94a3b8',
    'log_ok':  '#4ade80',
    'log_err': '#f87171',
    'log_inf': '#7dd3fc',
}
F = '맑은 고딕'


# 접두어·키워드에 쓸 수 있는 자리표시 → 글마다 자동으로 채워짐
#   {사건번호} {범죄명} {결과1} {결과2} {시트} {날짜}(예: 11/24)
_CASE_NO_RE = re.compile(r'\d{2,4}\s*[가-힣]{1,4}\s*\d{1,7}(?:호)?')


def case_meta(case=None, text=''):
    """엑셀 사례(또는 직접 입력한 사건자료)에서 자리표시 값을 만든다."""
    if case is not None:
        return {'사건번호': case.case_id.split('\n')[0].strip(), '범죄명': case.crime.split('\n')[0],
                '결과1': case.result1.replace('\n', ' '), '결과2': case.result2.replace('\n', ' '),
                '시트': case.sheet}
    m = _CASE_NO_RE.search(text or '')
    return {'사건번호': m.group().replace(' ', '') if m else ''}


def fill_template(tpl, meta=None):
    """'사건번호 : {사건번호}' → '사건번호 : 2021형제55'. 값이 없는 자리표시가 있으면 빈 문자열."""
    if not tpl or '{' not in tpl:
        return tpl or ''
    now = datetime.now()
    vals = dict(meta or {})
    vals.setdefault('날짜', f'{now.month}/{now.day}')
    missing = False

    def sub(m):
        nonlocal missing
        v = (vals.get(m.group(1)) or '').strip()
        if not v:
            missing = True
        return v
    out = re.sub(r'\{([^{}]+)\}', sub, tpl)
    return '' if missing else out


def _write_log_file(msg, level='inf'):
    """logs/날짜.log 에 한 줄씩 남김 (밤새 돌린 결과를 아침에 확인)"""
    try:
        os.makedirs(app_config.LOG_DIR, exist_ok=True)
        now = datetime.now()
        mark = {'err': '[오류]', 'warn': '[경고]', 'ok': '[완료]'}.get(level, '')
        with open(os.path.join(app_config.LOG_DIR, now.strftime('%Y-%m-%d') + '.log'), 'a', encoding='utf-8') as f:
            f.write(f'{now:%H:%M:%S} {mark} {msg}\n')
    except Exception:
        pass


LIMIT_GRACE = 60   # 한도 초기화 시각 + 1분 뒤에 다시 시작


def _keep_awake(on):
    """Windows: 작업·예약 중 PC가 절전 모드로 들어가지 않게 한다."""
    try:
        import ctypes
        ES_CONTINUOUS, ES_SYSTEM_REQUIRED = 0x80000000, 0x00000001
        ctypes.windll.kernel32.SetThreadExecutionState(
            ES_CONTINUOUS | (ES_SYSTEM_REQUIRED if on else 0))
    except Exception:
        pass


def _btn(p, text, cmd, bg, fg='white', fs=10, px=16, py=8, **kw):
    b = tk.Button(p, text=text, command=cmd, bg=bg, fg=fg,
                  font=(F, fs, 'bold'), padx=px, pady=py,
                  relief=tk.FLAT, bd=0, cursor='hand2',
                  activebackground=bg, activeforeground=fg, **kw)
    return b


class OneNoteExporterGUI:
    def __init__(self, root):
        self.root = root
        self.root.title('OneNote 글 내보내기')
        self.root.geometry('860x760')
        self.root.minsize(640, 500)
        self.root.configure(bg=C['bg'])
        self.root.resizable(True, True)

        self.articles     = []
        self.check_vars   = []  
        self.footer_var   = tk.BooleanVar(value=False) 
        self.footer_preset_vars = []  
        self._loading_job  = None  
        self.onenote      = OneNoteManager()
        self.is_connected = False
        self._section_map = {}

        app_config.load_config()
        app_config.ensure_projects_dir()
        self.ai_win       = None
        self._acc_login   = {}
        self._pending_section = None
        self._loaded_project  = None
        self.schedule     = None    # 예약 실행 정보
        self.projects     = []
        self.cases        = {}
        self.case_rows    = {}
        self.gen_stop     = threading.Event()
        self.generating   = False

        self._build()
        self._shortcuts()
        self.root.after(600, self._auto_connect)

    # ── 빌드 ──────────────────────────────────────────────────

    def _build(self):
        # 헤더
        hdr = tk.Frame(self.root, bg=C['header'])
        hdr.pack(fill=tk.X)

        tk.Label(hdr, text='📝', font=('Segoe UI Emoji', 13),
                 bg=C['header'], fg='white').pack(side=tk.LEFT, padx=(16,4), pady=0)
        tk.Label(hdr, text='OneNote 글 내보내기',
                 font=(F, 14, 'bold'), bg=C['header'], fg='white'
                 ).pack(side=tk.LEFT, padx=20, pady=12)

        self.conn_dot = tk.Label(hdr, text='● 연결 대기',
                                  font=(F, 9), bg=C['header'], fg='#f59e0b')
        self.conn_dot.pack(side=tk.RIGHT, padx=16)
        tk.Button(hdr, text='설정', command=self._open_settings, font=(F, 9),
                  bg=C['header'], fg='#cbd5e1', activebackground=C['header'],
                  activeforeground='white', relief=tk.FLAT, bd=0, cursor='hand2'
                  ).pack(side=tk.RIGHT, padx=(0, 4))


        cfg = tk.Frame(self.root, bg=C['white'],
                       highlightbackground=C['light'], highlightthickness=1)
        cfg.pack(fill=tk.X, padx=12, pady=(8, 0))

        tk.Label(cfg, text='섹션', font=(F, 9, 'bold'),
                 bg=C['white'], fg=C['subtext']).grid(row=0, column=0, padx=(12,4), pady=6, sticky='w')
        self.section_var = tk.StringVar(value='(자동)')
        self.section_cb  = ttk.Combobox(cfg, textvariable=self.section_var,
                                         state='readonly', width=22, font=(F, 9))
        self.section_cb.grid(row=0, column=1, padx=(0,3), pady=6)
        self.section_cb.bind('<<ComboboxSelected>>', self._on_section)

        tk.Frame(cfg, bg=C['light'], width=1).grid(row=0, column=2, sticky='ns', pady=4)

        tk.Checkbutton(cfg, text='하단 구분선', variable=self.footer_var,
                      font=(F, 9), bg=C['white'], fg=C['subtext'],
                      activebackground=C['white'], selectcolor=C['white']
                      ).grid(row=0, column=3, padx=(10,6), pady=6)
        tk.Label(cfg, text='접두어', font=(F, 9, 'bold'),
                 bg=C['white'], fg=C['subtext']).grid(row=0, column=4, padx=(4,4), pady=6)
        self.prefix_var = tk.StringVar()
        tk.Entry(cfg, textvariable=self.prefix_var, font=(F, 9),
                 width=18, relief=tk.FLAT, bg=C['bg'], insertbackground=C['text']
                 ).grid(row=0, column=5, padx=(0,3), pady=6, ipady=4)
        tk.Label(cfg, text='예) ○ 성지 [블로그]', font=(F, 8),
                 bg=C['white'], fg='#94a3b8').grid(row=0, column=6, padx=(0,6))

        cfg.columnconfigure(7, weight=1)
        self.export_btn = _btn(cfg, 'OneNote로 내보내기', self._export,
             bg=C['red'], fs=10, px=16, py=6)
        self.export_btn.grid(row=0, column=8, padx=(4,12), pady=6)

        tk.Frame(cfg, bg=C['light'], height=1).grid(
            row=1, column=0, columnspan=9, sticky='ew', padx=8, pady=(0,4))
        tk.Label(cfg, text='키워드', font=(F, 9, 'bold'),
                 bg=C['white'], fg=C['subtext']).grid(row=2, column=0, padx=(12,4), pady=(0,8), sticky='w')
        self.keyword_var = tk.StringVar()
        tk.Entry(cfg, textvariable=self.keyword_var, font=(F, 9),
                 width=22, relief=tk.FLAT, bg=C['bg'], insertbackground=C['text']
                 ).grid(row=2, column=1, padx=(0,3), pady=(0,8), ipady=4, sticky='w')
        tk.Label(cfg, text='본문 맨 위에 굵게 삽입 · {사건번호} {날짜} {범죄명} 쓰면 글마다 자동 입력', font=(F, 8),
                 bg=C['white'], fg='#94a3b8').grid(row=2, column=2, columnspan=6, padx=(6,6), pady=(0,8), sticky='w')


        inp = tk.Frame(self.root, bg=C['white'],
                       highlightbackground=C['light'], highlightthickness=1)
        inp.pack(fill=tk.X, padx=12, pady=(8, 0))

        inp_top = tk.Frame(inp, bg=C['bg'])
        inp_top.pack(fill=tk.X, padx=1, pady=1)
        tk.Label(inp_top, text='Claude 응답 붙여넣기',
                 font=(F, 9, 'bold'), bg=C['bg'], fg=C['subtext']
                 ).pack(side=tk.LEFT, padx=12, pady=6)
        tk.Label(inp_top, text='Ctrl+V 붙여넣기  →  F5 파싱',
                 font=(F, 8), bg=C['bg'], fg='#94a3b8'
                 ).pack(side=tk.RIGHT, padx=12)

        self.input_txt = scrolledtext.ScrolledText(
            inp, font=(F, 9), wrap=tk.WORD, bg=C['white'], fg=C['subtext'],
            relief=tk.FLAT, padx=12, pady=8, height=4,
            highlightthickness=0)
        self.input_txt.pack(fill=tk.X, padx=1, pady=(0,1))
        self._ph = "여기에 Claude 응답을 붙여넣으세요...\n\n예시:\n# 글 201\n**Focus Keyword:** 강제추행"
        self.input_txt.insert('1.0', self._ph)
        self.input_txt.bind('<FocusIn>',  self._ph_in)
        self.input_txt.bind('<FocusOut>', self._ph_out)

        inp_bot = tk.Frame(inp, bg=C['bg'])
        inp_bot.pack(fill=tk.X, padx=1, pady=1)
        self.status_lbl = tk.Label(inp_bot, text='응답을 붙여넣고 파싱하세요',
                                    font=(F, 9), bg=C['bg'], fg=C['subtext'])
        self.status_lbl.pack(side=tk.LEFT, padx=12, pady=6)
        _btn(inp_bot, '🔍  파싱  (F5)', self._parse,
             bg=C['green'], fs=10, px=24, py=7
             ).pack(side=tk.RIGHT, padx=8, pady=6)
        _btn(inp_bot, '✨  AI 원고 생성', self._open_ai_window,
             bg=C['accent'], fs=10, px=18, py=7
             ).pack(side=tk.RIGHT, padx=0, pady=6)

        lst = tk.Frame(self.root, bg=C['white'],
                       highlightbackground=C['light'], highlightthickness=1)
        lst.pack(fill=tk.BOTH, expand=True, padx=12, pady=(8,0))

        lst_hdr = tk.Frame(lst, bg=C['bg'])
        lst_hdr.pack(fill=tk.X, padx=1)
        tk.Label(lst_hdr, text='글 목록', font=(F, 10, 'bold'),
                 bg=C['bg'], fg=C['text']).pack(side=tk.LEFT, padx=12, pady=5)
        self.count_lbl = tk.Label(lst_hdr, text='0개',
                                   font=(F, 9), bg=C['bg'], fg=C['subtext'])
        self.count_lbl.pack(side=tk.LEFT, padx=4, pady=4)

        lw = tk.Frame(lst, bg=C['white'])
        lw.pack(fill=tk.BOTH, expand=True, padx=1, pady=0)
        sb = tk.Scrollbar(lw)
        sb.pack(side=tk.RIGHT, fill=tk.Y)
        self.check_canvas = tk.Canvas(lw, bg=C['white'],
                                       yscrollcommand=sb.set,
                                       highlightthickness=0)
        self.check_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb.config(command=self.check_canvas.yview)
        self.check_frame = tk.Frame(self.check_canvas, bg=C['white'])
        self.check_canvas_window = self.check_canvas.create_window(
            (0, 0), window=self.check_frame, anchor='nw')
        self.check_frame.bind('<Configure>', self._on_check_frame_configure)
        self.check_canvas.bind('<Configure>', self._on_canvas_configure)

        self.check_canvas.bind('<MouseWheel>',
            lambda e: self.check_canvas.yview_scroll(-1*(e.delta//120), 'units'))

        lst_bot = tk.Frame(lst, bg=C['bg'])
        lst_bot.pack(fill=tk.X, padx=1, pady=1)
        for t, cmd in [('✓ 전체', self._sel_all),
                        ('✗ 해제',  self._desel),
                        ('🗑 삭제',  self._delete),
                        ('↻ 다시 쓰기', self._rewrite_selected),
                        ('📋 로톡용 복사', self._copy_lawtalk),
                        ('제목 복사', self._copy_title),
                        ('🔗 링크 기록', self._link_dialog)]:
            tk.Button(lst_bot, text=t, command=cmd,
                      font=(F, 9), bg=C['light'], fg=C['text'],
                      relief=tk.FLAT, padx=12, pady=5, cursor='hand2'
                      ).pack(side=tk.LEFT, padx=(8,0), pady=6)

 
        footer_frame = tk.Frame(self.root, bg=C['white'],
                                highlightbackground=C['light'], highlightthickness=1)
        footer_frame.pack(fill=tk.X, padx=12, pady=(4, 0))

        self.footer_preset_vars = []
        boxes = []
        for name, _ in FOOTER_PRESETS:
            var = tk.BooleanVar(value=False)
            self.footer_preset_vars.append(var)
            boxes.append(tk.Checkbutton(footer_frame, text=name, variable=var,
                                        font=(F, 9), bg=C['white'], fg=C['text'],
                                        activebackground=C['white'], selectcolor=C['white']))

        def _reflow(e=None):
            # 창 너비에 들어가는 만큼 한 줄에 놓고, 넘치면 다음 줄로
            width = max(footer_frame.winfo_width(), 200) - 24
            x = row = col = 0
            for b in boxes:
                w = b.winfo_reqwidth() + 4
                if x and x + w > width:
                    row += 1; col = 0; x = 0
                b.grid(row=row, column=col, sticky='w', padx=(12 if col == 0 else 0, 4), pady=3)
                col += 1; x += w
        footer_frame.bind('<Configure>', _reflow)
        footer_frame.after(50, _reflow)



        pb_area = tk.Frame(self.root, bg=C['bg'])
        pb_area.pack(fill=tk.X, padx=12, pady=(4, 2))
        self.prog_status = tk.Label(pb_area, text='', font=(F, 8),
                                     bg=C['bg'], fg=C['gray'])
        self.prog_status.pack(side=tk.RIGHT)
        pb_wrap = tk.Frame(pb_area, bg=C['light'], height=6)
        pb_wrap.pack(fill=tk.X, side=tk.LEFT, expand=True, pady=6)
        pb_wrap.pack_propagate(False)
        self.pb_fill = tk.Frame(pb_wrap, bg=C['accent'], height=6)
        self.pb_fill.place(x=0, y=0, relwidth=0.0, height=6)

  
        self.log_txt = scrolledtext.ScrolledText(
            self.root, font=('Consolas', 8), bg=C['log_bg'], fg=C['log_fg'],
            relief=tk.FLAT, padx=8, pady=4, state=tk.DISABLED, height=1)

        bot = tk.Frame(self.root, bg=C['bg'], height=28)
        bot.pack(fill=tk.X, padx=12, pady=(0, 6))
        bot.pack_propagate(False)
        self.prog_lbl = tk.Label(bot, text='', font=(F, 9),
                                  bg=C['bg'], fg=C['green'])
        self.prog_lbl.pack(side=tk.LEFT, padx=4)

    def _rebuild_checks(self):
        """체크박스 목록 재생성"""
        for w in self.check_frame.winfo_children():
            w.destroy()
        self.check_vars = []
        for a in self.articles:
            var = tk.BooleanVar(value=True)
            self.check_vars.append(var)
            t = ('🔗 ' if a.get('lawtalk_url') else '') + a['title'][:60] + ('…' if len(a['title']) > 60 else '')
            tk.Checkbutton(
                self.check_frame, text=t,
                variable=var,
                font=(F, 9), bg=C['white'], fg=C['text'],
                activebackground=C['white'],
                selectcolor=C['white'],
                anchor='w', padx=8, pady=2,
            ).pack(fill=tk.X, padx=4, pady=0)
        self.check_frame.update_idletasks()
        self.check_canvas.config(scrollregion=self.check_canvas.bbox('all'))


    def _start_loading(self):
        dots = ['', '.', '..', '...']
        idx = [0]
        def tick():
            self.export_btn.config(text='내보내는 중' + dots[idx[0] % 4])
            idx[0] += 1
            self._loading_job = self.root.after(400, tick)
        tick()

    def _stop_loading(self):
        if self._loading_job:
            self.root.after_cancel(self._loading_job)
            self._loading_job = None
        self.export_btn.config(text='OneNote로 내보내기', state=tk.NORMAL)

    def _on_check_frame_configure(self, _):
        self.check_canvas.config(scrollregion=self.check_canvas.bbox('all'))

    def _on_canvas_configure(self, e):
        self.check_canvas.itemconfig(self.check_canvas_window, width=e.width)


    def _shortcuts(self):
        self.root.bind('<F5>',        lambda e: self._parse())
        self.root.bind('<Control-a>', lambda e: self._sel_all())
        self.root.bind('<Control-A>', lambda e: self._sel_all())


    def _ph_in(self, _):
        if self.input_txt.get('1.0', tk.END).strip() == self._ph.strip():
            self.input_txt.delete('1.0', tk.END)
            self.input_txt.config(fg=C['text'])

    def _ph_out(self, _):
        if not self.input_txt.get('1.0', tk.END).strip():
            self.input_txt.insert('1.0', self._ph)
            self.input_txt.config(fg=C['subtext'])


    def _log(self, msg, level='inf'):
        from datetime import datetime
        ts = datetime.now().strftime('%H:%M:%S')
        self.log_txt.config(state=tk.NORMAL)
        self.log_txt.insert(tk.END, f"[{ts}] {msg}\n", level)
        self.log_txt.see(tk.END)
        self.log_txt.config(state=tk.DISABLED)
        _write_log_file(msg, level)


    def _auto_connect(self):
        self.conn_dot.config(text='● 연결 중...', fg='#f59e0b')
        self.export_btn.config(state=tk.DISABLED)
        import threading
        def _auto():
            ok = self.onenote.connect()
            self.root.after(0, lambda: self._connect_done(ok))
        threading.Thread(target=_auto, daemon=True).start()



    def _connect_done(self, ok):
        self._set_conn(ok)
        if ok:
            self._refresh_sections()
            self.export_btn.config(state=tk.NORMAL)
        else:
            messagebox.showerror('오류', 'OneNote 연결 실패.\nOneNote가 실행 중인지 확인 후 앱을 재시작하세요.')

    def _set_conn(self, ok):
        self.is_connected = ok
        if ok:
            self.conn_dot.config(text='● 연결됨', fg='#4ade80')
            self._log('연결 성공', 'ok')
        else:
            self.conn_dot.config(text='● 연결 실패', fg='#f87171')
            self._log('연결 실패', 'err')


    def _refresh_sections(self):
        if not self.is_connected:
            return
        try:
            secs = self.onenote.get_all_sections()
            self._section_map = {s.display_name(): s for s in secs}
            self.section_cb['values'] = ['(자동)'] + list(self._section_map.keys())
            self.section_cb.set('(자동)')
            self._sync_ai_sections()
            self._log(f"섹션 {len(secs)}개 로드", 'ok')
        except Exception as e:
            self._log(f"섹션 오류: {e}", 'err')

    def _on_section(self, _):
        val = self.section_var.get()
        if val == '(자동)':
            self.onenote.section_id = ''
        elif val in self._section_map:
            self.onenote.set_section(self._section_map[val].id)
            self._log(f"섹션: {val}", 'ok')

    def _open_file(self):
        path = filedialog.askopenfilename(
            title='파일 열기',
            filetypes=[('텍스트', '*.txt *.md'), ('전체', '*.*')])
        if path:
            self._open_file_path(path)

    def _open_file_path(self, path):
        """--input 옵션 / 파일 열기 공통"""
        try:
            arts = ArticleParser.parse_from_file(path)
            if not arts:
                messagebox.showwarning('경고', '파싱 실패'); return
            self._load(arts)
            self._log(f"파일: {path} ({len(arts)}개)", 'ok')
        except Exception as e:
            messagebox.showerror('오류', f'파일 오류:\n{e}')


    def _parse(self):
        content = self.input_txt.get('1.0', tk.END).strip()
        if not content or content == self._ph.strip():
            messagebox.showwarning('경고', '응답을 먼저 붙여넣으세요.')
            return
        # 최소 내용 있는지만 확인
        if len(content) < 10:
            messagebox.showwarning('경고', '내용이 너무 짧습니다. Claude 응답을 붙여넣으세요.')
            return
        self._set_status('파싱 중…', C['orange'])
        arts = ArticleParser.parse(content)
        if not arts:
            messagebox.showerror('오류', '파싱 실패. 형식을 확인하세요.')
            self._set_status('파싱 실패', C['red'])
            return
        self._load(arts)

    def _load(self, arts):
        self.articles = arts
        self._rebuild_checks()
        self.count_lbl.config(text=f"{len(arts)}개")
        self._set_status(f"✅ {len(arts)}개 로드 완료", C['green'])
        self._log(f"파싱 완료: {len(arts)}개", 'ok')

    def _sel_all(self):
        for v in self.check_vars:
            v.set(True)

    def _desel(self):
        for v in self.check_vars:
            v.set(False)

    def _delete(self):
        sel = [i for i, v in enumerate(self.check_vars) if v.get()]
        if not sel:
            messagebox.showwarning('경고', '삭제할 글을 체크하세요.')
            return
        if not messagebox.askyesno('확인', f"{len(sel)}개 삭제?"):
            return
        self.articles = [a for i, a in enumerate(self.articles) if i not in sel]
        self._rebuild_checks()
        self.count_lbl.config(text=f"{len(self.articles)}개")


    def _export(self):
        if not self.is_connected:
            messagebox.showwarning('경고', 'OneNote 연결 필요')
            return
        sel = [i for i, v in enumerate(self.check_vars) if v.get()]
        if not sel:
            messagebox.showwarning('경고', '내보낼 글을 체크하세요.')
            return
        selected = [self.articles[i] for i in sel]
        prefix   = self.prefix_var.get().strip()
        keyword  = self.keyword_var.get().strip()
        msg = f"💾 {len(selected)}개 글을 내보낼까요?\n섹션: {self.section_var.get()}"
        if prefix:
            msg += f"\n접두어: {prefix}"
        if keyword:
            msg += f"\n키워드(상단삽입): {keyword}"
        if not messagebox.askyesno('확인', msg):
            return
        self._log(f"내보내기 시작: {len(selected)}개", 'inf')
        self.export_btn.config(state=tk.DISABLED)


        selected_footers = [
            lines
            for var, (_, lines) in zip(self.footer_preset_vars, FOOTER_PRESETS)
            if var.get()
        ]

        threading.Thread(
            target=self._export_thread,
            args=(selected, prefix, selected_footers, keyword),
            daemon=True).start()


    def _export_articles(self, arts, prefix, keyword, footer_presets, add_footer, progress_cb=None):
        """글마다 자리표시({사건번호} 등)를 채워서 OneNote에 넣는다."""
        result = {'success': 0, 'fail': 0, 'failed_nums': []}
        for i, art in enumerate(arts, 1):
            if progress_cb:
                progress_cb(i, len(arts), art)
            meta = art.get('meta') or {}
            ok = self.onenote.create_page(art, title_prefix=fill_template(prefix, meta),
                                          footer_presets=footer_presets,
                                          top_keyword=fill_template(keyword, meta))
            if ok:
                art['page_id'] = getattr(self.onenote, 'last_page_id', '')
                result['success'] += 1
            else:
                result['fail'] += 1; result['failed_nums'].append(art.get('number', '?'))
        if add_footer:
            try:
                self.onenote.create_page({'number': '0', 'title': 'ㅡ' * 20, 'focus_keyword': '',
                                          'sub_keywords': '', 'toc': [], 'summary': '',
                                          'sections': [], 'raw_content': ''}, title_prefix='')
            except Exception:
                pass
        return result

    def _export_thread(self, selected, prefix, footer_presets, keyword=''):
        def cb(i, total, art):
            pct = i / total
            self.root.after(0, lambda p=pct: self.pb_fill.place(
                x=0, y=0, relwidth=p, height=5))
            self.root.after(0, lambda: self.prog_status.config(
                text=f"{i}/{total}"))
            self.root.after(0, lambda: self._log(
                f"({i}/{total}) 글 {art['number']}: {art['title'][:30]}"))
        add_footer = self.footer_var.get()
        result = self._export_articles(selected, prefix, keyword, footer_presets,
                                       add_footer, progress_cb=cb)
        self.root.after(0, lambda: self._export_done(result, len(selected)))

    def _export_done(self, result, total):
        self.export_btn.config(state=tk.NORMAL)
        self.prog_lbl.config(text='')
        self.pb_fill.place(x=0, y=0, relwidth=0.0, height=5)
        self.prog_status.config(text='')
        s, f = result['success'], result['fail']
        if f == 0:
            self._set_status(f"✅ {s}개 내보내기 완료", C['green'])
            self._log(f"완료: {s}개", 'ok')
        else:
            failed = ', '.join(result['failed_nums'])
            self._set_status(f"⚠️ {s}/{total} 성공  /  {f}개 실패: {failed}", C['orange'])
            self._log(f"{s}개 성공 / {f}개 실패 ({failed})", 'err')

    def _set_status(self, text, color=None):
        self.status_lbl.config(text=text, fg=color or C['subtext'])

    # ══════════════════════════════════════════════════════════
    # AI 원고 생성 창
    # ══════════════════════════════════════════════════════════

    def _card(self, parent, title=None, side=tk.TOP):
        f = tk.Frame(parent, bg=C['white'], highlightbackground=C['light'], highlightthickness=1)
        f.pack(fill=tk.X, padx=12, pady=(8, 0), side=side)
        if title:
            top = tk.Frame(f, bg=C['bg'])
            top.pack(fill=tk.X, padx=1, pady=1)
            tk.Label(top, text=title, font=(F, 9, 'bold'), bg=C['bg'],
                     fg=C['subtext']).pack(side=tk.LEFT, padx=12, pady=6)
            f._top = top
        return f

    def _small_btn(self, parent, text, cmd):
        return tk.Button(parent, text=text, command=cmd, font=(F, 9), bg=C['light'],
                         fg=C['text'], relief=tk.FLAT, padx=12, pady=4, cursor='hand2')

    def _dropdown(self, parent, width=14, command=None, var=None):
        """ttk 콤보박스 대신 쓰는 깔끔한 드롭다운 (설정창과 같은 납작한 스타일)"""
        var = var or tk.StringVar()
        mb = tk.Menubutton(parent, text='', font=(F, 9), bg=C['bg'], fg=C['text'],
                           activebackground=C['light'], activeforeground=C['text'],
                           relief=tk.FLAT, bd=0, padx=10, pady=5, width=width, anchor='w',
                           cursor='hand2', highlightthickness=0, direction='below')
        menu = tk.Menu(mb, tearoff=0, font=(F, 9), bg=C['white'], fg=C['text'], bd=0,
                       activebackground=C['accent'], activeforeground='white')
        mb['menu'] = menu

        def _show(*_):
            mb.config(text=f'{var.get()}  ▾')
        var.trace_add('write', _show)

        def set_values(values):
            menu.delete(0, tk.END)
            for v in values:
                menu.add_command(label=v, command=lambda v=v: (var.set(v), command and command()))
        mb.var, mb.set_values = var, set_values
        _show()
        return mb

    def _segment(self, parent, var, options, command=None):
        """버튼처럼 생긴 라디오 (선택된 칸만 연한 파랑)"""
        f = tk.Frame(parent, bg=C['light'], padx=1, pady=1)
        for val, text in options:
            tk.Radiobutton(f, text=text, variable=var, value=val, command=command,
                           indicatoron=0, font=(F, 9), bg=C['white'], fg=C['subtext'],
                           selectcolor='#dbeafe', activebackground=C['bg'], relief=tk.FLAT,
                           bd=0, offrelief=tk.FLAT, padx=14, pady=4, cursor='hand2'
                           ).pack(side=tk.LEFT, padx=(0, 1))
        return f

    def _open_ai_window(self):
        if self.ai_win and self.ai_win.winfo_exists():
            self.ai_win.deiconify(); self.ai_win.lift(); return
        win = tk.Toplevel(self.root)
        self.ai_win = win
        win.title('AI 원고 생성')
        win.geometry('860x880')
        win.minsize(640, 540)
        win.configure(bg=C['bg'])
        win.protocol('WM_DELETE_WINDOW', lambda: (self._save_project_ui(), win.withdraw()))

        # ── 프로젝트 / 입력 / 계정 ──
        card = self._card(win, '프로젝트')
        for t, cmd in [('새로고침', self._refresh_projects), ('새 프로젝트', self._new_project),
                       ('지침 폴더', self._open_project_folder), ('✎ 지침 관리', self._guide_manager)]:
            tk.Button(card._top, text=t, command=cmd, font=(F, 8), bg=C['bg'], fg=C['subtext'],
                      activebackground=C['light'], relief=tk.FLAT, bd=0, padx=8,
                      cursor='hand2').pack(side=tk.RIGHT, padx=(0, 6))
        body = tk.Frame(card, bg=C['white'])
        body.pack(fill=tk.X, padx=12, pady=(6, 10))

        def label(r, text):
            tk.Label(body, text=text, font=(F, 9), bg=C['white'], fg=C['subtext'], width=7,
                     anchor='w').grid(row=r, column=0, sticky='w', pady=3)

        label(0, '프로젝트')
        self.project_dd = self._dropdown(body, width=26, command=self._on_project)
        self.project_var = self.project_dd.var
        self.project_dd.grid(row=0, column=1, sticky='w', pady=3)
        self.project_info = tk.Label(body, text='', font=(F, 8), bg=C['white'], fg='#94a3b8')
        self.project_info.grid(row=0, column=2, sticky='w', padx=10)

        label(1, '입력')
        self.source_var = tk.StringVar(value='excel')
        self._segment(body, self.source_var, [('excel', '엑셀 사례'), ('direct', '직접 입력')],
                      self._on_source).grid(row=1, column=1, sticky='w', pady=3)
        qf = tk.Frame(body, bg=C['white'])
        qf.grid(row=1, column=2, sticky='w', padx=10)
        tk.Label(qf, text='사례당', font=(F, 9), bg=C['white'], fg=C['subtext']).pack(side=tk.LEFT)
        self.count_var = tk.StringVar(value='10')
        tk.Entry(qf, textvariable=self.count_var, width=3, font=(F, 9), relief=tk.FLAT,
                 bg=C['bg'], justify='center').pack(side=tk.LEFT, padx=4, ipady=3)
        tk.Label(qf, text='편', font=(F, 9), bg=C['white'], fg=C['subtext']).pack(side=tk.LEFT)

        label(2, '계정')
        acc = CONFIG.get('account') or 1
        self.account_var = tk.StringVar(value=str(acc))
        self._segment(body, self.account_var, [('1', '계정 1'), ('2', '계정 2'),
                                               ('both', '1 + 2 동시')]
                      ).grid(row=2, column=1, sticky='w', pady=3)
        tk.Label(body, text='동시: 두 계정이 사례를 나눠서 함께 작성', font=(F, 8),
                 bg=C['white'], fg='#94a3b8').grid(row=2, column=2, sticky='w', padx=10)
        st = tk.Frame(body, bg=C['white'])
        st.grid(row=3, column=1, columnspan=2, sticky='w', pady=(2, 0))
        self.acc_status = {}
        for n in (1, 2):
            tk.Label(st, text=f'계정 {n}', font=(F, 8, 'bold'), bg=C['white'],
                     fg=C['subtext']).pack(side=tk.LEFT)
            lb = tk.Label(st, text='● 확인 중', font=(F, 8), bg=C['white'], fg='#94a3b8')
            lb.pack(side=tk.LEFT, padx=(4, 4))
            self.acc_status[n] = lb
            tk.Button(st, text='연결', command=lambda n=n: self._token_dialog(n), font=(F, 8),
                      bg=C['bg'], fg=C['subtext'], activebackground=C['light'], relief=tk.FLAT,
                      bd=0, padx=6, cursor='hand2').pack(side=tk.LEFT, padx=(0, 14))
        tk.Button(st, text='↻ 다시 확인', command=self._check_accounts, font=(F, 8),
                  bg=C['white'], fg=C['accent'], activebackground=C['white'], relief=tk.FLAT,
                  bd=0, cursor='hand2').pack(side=tk.LEFT)

        # ── OneNote (메인 화면과 같은 값 공유) ──
        oc = self._card(win, 'OneNote')
        self.auto_export_var = tk.BooleanVar(value=CONFIG.get('auto_export', True))
        tk.Checkbutton(oc._top, text='생성 후 바로 내보내기', variable=self.auto_export_var,
                       font=(F, 8), bg=C['bg'], fg=C['subtext'], activebackground=C['bg'],
                       selectcolor=C['white']).pack(side=tk.RIGHT, padx=8)
        ob = tk.Frame(oc, bg=C['white'])
        ob.pack(fill=tk.X, padx=12, pady=(6, 10))
        tk.Label(ob, text='섹션', font=(F, 9), bg=C['white'], fg=C['subtext'], width=7,
                 anchor='w').grid(row=0, column=0, sticky='w', pady=3)
        self.section_dd = self._dropdown(ob, width=30, var=self.section_var,
                                         command=lambda: self._on_section(None))
        self.section_dd.grid(row=0, column=1, columnspan=3, sticky='w', pady=3)
        self._sync_ai_sections()
        tk.Checkbutton(ob, text='하단 구분선', variable=self.footer_var, font=(F, 9),
                       bg=C['white'], fg=C['subtext'], activebackground=C['white'],
                       selectcolor=C['white']).grid(row=0, column=4, sticky='w', padx=10)
        tk.Label(ob, text='접두어', font=(F, 9), bg=C['white'], fg=C['subtext'], width=7,
                 anchor='w').grid(row=1, column=0, sticky='w', pady=3)
        tk.Entry(ob, textvariable=self.prefix_var, font=(F, 9), width=20, relief=tk.FLAT,
                 bg=C['bg'], insertbackground=C['text']).grid(row=1, column=1, sticky='w', ipady=4)
        tk.Label(ob, text='키워드', font=(F, 9), bg=C['white'], fg=C['subtext']
                 ).grid(row=1, column=2, sticky='w', padx=(14, 6))
        tk.Entry(ob, textvariable=self.keyword_var, font=(F, 9), width=20, relief=tk.FLAT,
                 bg=C['bg'], insertbackground=C['text']).grid(row=1, column=3, sticky='w', ipady=4)
        tk.Label(ob, text='{사건번호} {날짜} {범죄명} {결과2} → 글마다 자동 입력', font=(F, 8),
                 bg=C['white'], fg='#94a3b8').grid(row=1, column=4, sticky='w', padx=10)
        tk.Label(ob, text='하단 문구', font=(F, 9), bg=C['white'], fg=C['subtext'], width=7,
                 anchor='nw').grid(row=2, column=0, sticky='nw', pady=(6, 0))
        pf = tk.Frame(ob, bg=C['white'])
        pf.grid(row=2, column=1, columnspan=4, sticky='w', pady=(4, 0))
        for i, ((name, _), var) in enumerate(zip(FOOTER_PRESETS, self.footer_preset_vars)):
            tk.Checkbutton(pf, text=name, variable=var, font=(F, 9), bg=C['white'],
                           fg=C['text'], activebackground=C['white'], selectcolor=C['white']
                           ).grid(row=i // 4, column=i % 4, sticky='w', padx=(0, 10))

        # ── 실행 ──
        run = tk.Frame(win, bg=C['bg'])
        run.pack(fill=tk.X, padx=12, pady=10, side=tk.BOTTOM)
        tk.Button(run, text='결과 폴더', command=lambda: app_config.open_folder(app_config.OUTPUT_DIR),
                  font=(F, 9), bg=C['light'], fg=C['text'], relief=tk.FLAT, padx=12, pady=6,
                  cursor='hand2').pack(side=tk.LEFT)
        tk.Button(run, text='기록(로그)', command=lambda: app_config.open_folder(app_config.LOG_DIR),
                  font=(F, 9), bg=C['light'], fg=C['text'], relief=tk.FLAT, padx=12, pady=6,
                  cursor='hand2').pack(side=tk.LEFT, padx=(6, 0))
        self.stop_btn = tk.Button(run, text='중지', command=self._stop_generate, font=(F, 10),
                                  bg=C['light'], fg=C['subtext'], relief=tk.FLAT, padx=16,
                                  pady=6, cursor='hand2', state=tk.DISABLED)
        self.stop_btn.pack(side=tk.RIGHT, padx=(6, 0))
        self.gen_btn = _btn(run, '원고 생성', self._start_generate, bg=C['accent'], fs=10,
                            px=28, py=7)
        self.gen_btn.pack(side=tk.RIGHT)
        tk.Button(run, text='⏰ 예약', command=self._open_schedule, font=(F, 9), bg=C['light'],
                  fg=C['text'], relief=tk.FLAT, padx=12, pady=6, cursor='hand2'
                  ).pack(side=tk.RIGHT, padx=(0, 6))
        self.sched_lbl = tk.Label(run, text='', font=(F, 9), bg=C['bg'], fg=C['accent'])
        self.sched_lbl.pack(side=tk.LEFT, padx=10)

        # ── 진행 상황 ──
        lc = self._card(win, '진행 상황', side=tk.BOTTOM)
        self.wait_limit_var = tk.BooleanVar(value=True)
        self.autofix_var = tk.BooleanVar(value=True)
        for text, var in [('한도에 걸리면 기다렸다 이어 쓰기', self.wait_limit_var),
                          ('검사에 걸린 원고는 그 편만 자동으로 다시 쓰기', self.autofix_var)]:
            tk.Checkbutton(lc._top, text=text, variable=var, font=(F, 8), bg=C['bg'],
                           fg=C['subtext'], activebackground=C['bg'], selectcolor=C['white']
                           ).pack(side=tk.RIGHT, padx=(0, 8))
        self.prog_area = tk.Frame(lc, bg=C['white'])
        self.prog_area.pack(fill=tk.X, padx=12)
        self.prog_rows = {}
        self.ai_log = tk.Text(lc, font=(F, 9), bg=C['white'], fg=C['subtext'], relief=tk.FLAT,
                              padx=12, pady=6, height=3, state=tk.DISABLED, wrap=tk.WORD,
                              highlightthickness=0)
        self.ai_log.pack(fill=tk.X, padx=1, pady=(0, 1))
        for tag, col in [('ok', C['green']), ('err', C['red']), ('inf', C['subtext']),
                         ('warn', C['orange'])]:
            self.ai_log.tag_configure(tag, foreground=col)

        # ── 엑셀 사례 ──
        self.src_area = tk.Frame(win, bg=C['bg'])
        self.src_area.pack(fill=tk.BOTH, expand=True)

        ef = tk.Frame(self.src_area, bg=C['white'], highlightbackground=C['light'],
                      highlightthickness=1)
        self.excel_frame = ef
        top = tk.Frame(ef, bg=C['bg'])
        top.pack(fill=tk.X, padx=1, pady=1)
        tk.Label(top, text='사례 선택', font=(F, 9, 'bold'), bg=C['bg'],
                 fg=C['subtext']).pack(side=tk.LEFT, padx=12, pady=6)
        self.excel_lbl = tk.Label(top, text='', font=(F, 8), bg=C['bg'], fg='#94a3b8')
        self.excel_lbl.pack(side=tk.LEFT, padx=4)
        tk.Button(top, text='엑셀 파일 선택', command=self._pick_excel, font=(F, 8),
                  bg=C['bg'], fg=C['subtext'], activebackground=C['light'], relief=tk.FLAT,
                  bd=0, padx=8, cursor='hand2').pack(side=tk.RIGHT, padx=6)

        flt = tk.Frame(ef, bg=C['white'])
        flt.pack(fill=tk.X, padx=12, pady=(8, 6))
        tk.Label(flt, text='시트', font=(F, 9), bg=C['white'], fg=C['subtext']).pack(side=tk.LEFT)
        self.sheet_dd = self._dropdown(flt, width=12, command=self._on_sheet)
        self.sheet_var = self.sheet_dd.var
        self.sheet_dd.pack(side=tk.LEFT, padx=(6, 14))
        tk.Label(flt, text='분류', font=(F, 9), bg=C['white'], fg=C['subtext']).pack(side=tk.LEFT)
        self.cat_dd = self._dropdown(flt, width=14, command=self._fill_cases)
        self.cat_var = self.cat_dd.var
        self.cat_var.set('(전체)')
        self.cat_dd.pack(side=tk.LEFT, padx=(6, 14))
        self.unused_var = tk.BooleanVar(value=True)
        tk.Checkbutton(flt, text='안 쓴 사례만', variable=self.unused_var,
                       command=self._fill_cases, font=(F, 9), bg=C['white'],
                       fg=C['subtext'], activebackground=C['white'], selectcolor=C['white']
                       ).pack(side=tk.LEFT)
        self.case_count = tk.Label(flt, text='', font=(F, 8), bg=C['white'], fg='#94a3b8')
        self.case_count.pack(side=tk.RIGHT)

        tk.Frame(ef, bg=C['light'], height=1).pack(fill=tk.X)
        lw = tk.Frame(ef, bg=C['white'])
        lw.pack(fill=tk.BOTH, expand=True)
        sb = tk.Scrollbar(lw)
        sb.pack(side=tk.RIGHT, fill=tk.Y)
        self.case_canvas = tk.Canvas(lw, bg=C['white'], yscrollcommand=sb.set,
                                     highlightthickness=0)
        self.case_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb.config(command=self.case_canvas.yview)
        self.case_frame = tk.Frame(self.case_canvas, bg=C['white'])
        cw = self.case_canvas.create_window((0, 0), window=self.case_frame, anchor='nw')
        self.case_frame.bind('<Configure>', lambda e: self.case_canvas.config(
            scrollregion=self.case_canvas.bbox('all')))
        self.case_canvas.bind('<Configure>', lambda e: self.case_canvas.itemconfig(cw, width=e.width))
        wheel = lambda e: self.case_canvas.yview_scroll(-1 * (e.delta // 120), 'units')
        self.case_canvas.bind('<Enter>', lambda e: self.case_canvas.bind_all('<MouseWheel>', wheel))
        self.case_canvas.bind('<Leave>', lambda e: self.case_canvas.unbind_all('<MouseWheel>'))
        self.case_checks = []

        bot = tk.Frame(ef, bg=C['bg'])
        bot.pack(fill=tk.X, padx=1, pady=1)
        tk.Label(bot, text='위에서', font=(F, 9), bg=C['bg'], fg=C['subtext']
                 ).pack(side=tk.LEFT, padx=(12, 4), pady=6)
        self.pick_n = tk.Entry(bot, width=3, font=(F, 9), relief=tk.FLAT, bg=C['white'],
                               justify='center')
        self.pick_n.insert(0, '1')
        self.pick_n.pack(side=tk.LEFT, ipady=3)
        for t, cmd in [('개 선택', self._auto_pick), ('✗ 해제', self._clear_pick)]:
            tk.Button(bot, text=t, command=cmd, font=(F, 9), bg=C['light'], fg=C['text'],
                      relief=tk.FLAT, padx=12, pady=4, cursor='hand2'
                      ).pack(side=tk.LEFT, padx=(6, 0))
        self.pick_lbl = tk.Label(bot, text='', font=(F, 9), bg=C['bg'], fg=C['accent'])
        self.pick_lbl.pack(side=tk.RIGHT, padx=12)

        # ── 직접 입력 ──
        df = tk.Frame(self.src_area, bg=C['white'], highlightbackground=C['light'],
                      highlightthickness=1)
        self.direct_frame = df
        dtop = tk.Frame(df, bg=C['bg'])
        dtop.pack(fill=tk.X, padx=1, pady=1)
        tk.Label(dtop, text='사건자료 · 주제', font=(F, 9, 'bold'), bg=C['bg'],
                 fg=C['subtext']).pack(side=tk.LEFT, padx=12, pady=6)
        self.direct_hint = tk.Label(dtop, text='', font=(F, 8), bg=C['bg'], fg='#94a3b8')
        self.direct_hint.pack(side=tk.RIGHT, padx=12)
        self.direct_txt = scrolledtext.ScrolledText(df, font=(F, 9), wrap=tk.WORD, height=8,
                                                    relief=tk.FLAT, bg=C['white'], padx=12, pady=8,
                                                    highlightthickness=0)
        self.direct_txt.pack(fill=tk.BOTH, expand=True, padx=1, pady=(0, 1))

        self._refresh_projects()
        self._check_accounts()
        if CONFIG.get('excel_path') and os.path.exists(CONFIG['excel_path']):
            win.after(200, lambda: self._load_excel(CONFIG['excel_path']))
        else:
            self.excel_lbl.config(text='엑셀 파일을 선택하세요')
        self._on_source()

    def _sync_ai_sections(self):
        if getattr(self, 'section_dd', None) and self.section_dd.winfo_exists():
            self.section_dd.set_values(['(자동)'] + list(self._section_map.keys()))
            self.section_var.set(self.section_var.get())   # 표시 갱신
        if getattr(self, '_pending_section', None):
            self._apply_section(self._pending_section)

    def _check_accounts(self):
        for n, lb in self.acc_status.items():
            lb.config(text='● 확인 중', fg='#94a3b8')

        def work():
            for n in (1, 2):
                ok, msg = generator.check_login(n)
                self._acc_login[n] = (ok, msg)
                self._ui(lambda n=n: self._show_acc(n))
        threading.Thread(target=work, daemon=True).start()

    def _show_acc(self, n):
        if not (self.ai_win and self.ai_win.winfo_exists()):
            return
        ok, msg = self._acc_login.get(n, (False, '확인 중'))
        u = generator.USAGE.get(str(n)) or {}
        parts = [msg]
        for k, lab in (('five_hour', '5시간'), ('seven_day', '7일')):
            if u.get(k) and u[k][0] is not None:
                parts.append(f'{lab} {int(u[k][0] * 100)}%')
        if u.get('status') == 'rejected':
            parts.append('한도 도달')
        self.acc_status[n].config(text='● ' + ' · '.join(parts),
                                  fg=(C['green'] if ok and u.get('status') != 'rejected' else
                                      C['orange'] if ok else C['red']))

    def _token_dialog(self, n):
        """계정 연결: 장기 토큰(1년)을 발급받아 붙여넣기. 계정끼리 서로 로그아웃시키지 않음."""
        if not generator.find_claude_cli():
            self._offer_install(); return
        w = tk.Toplevel(self.ai_win)
        w.title(f'계정 {n} 연결')
        w.configure(bg=C['bg'])
        w.transient(self.ai_win)
        w.resizable(False, False)
        card = self._card(w, f'계정 {n} 연결 (한 번만 하면 약 1년 유지)')
        b = tk.Frame(card, bg=C['white'])
        b.pack(fill=tk.X, padx=12, pady=(8, 12))
        step = lambda r, t: tk.Label(b, text=t, font=(F, 9), bg=C['white'], fg=C['text'],
                                     justify='left', anchor='w').grid(row=r, column=0, columnspan=3,
                                                                      sticky='w', pady=(6, 2))
        step(0, '① 아래 버튼 → 검은 창이 뜨고 브라우저가 열립니다')
        def open_setup():
            try:
                generator.open_token_setup(n)
            except Exception as e:
                messagebox.showerror('연결', str(e), parent=w)
        self._small_btn(b, '토큰 발급 창 열기', open_setup).grid(row=1, column=0, sticky='w')
        tk.Label(b, text=f'브라우저에서 계정 {n}로 로그인 → 승인\n'
                         '(다른 계정으로 로그인돼 있으면 검은 창의 주소를 시크릿 창에 붙여넣기)',
                 font=(F, 8), bg=C['white'], fg='#94a3b8', justify='left'
                 ).grid(row=2, column=0, columnspan=3, sticky='w', pady=(4, 0))
        step(3, '② 검은 창에 나온 sk-ant-oat… 토큰을 복사해서 붙여넣기')
        ent = tk.Entry(b, font=(F, 9), width=40, relief=tk.FLAT, bg=C['bg'], show='•')
        ent.grid(row=4, column=0, sticky='w', ipady=4)
        if generator.get_token(n):
            ent.insert(0, generator.get_token(n))
        def paste():
            try:
                ent.delete(0, tk.END); ent.insert(0, w.clipboard_get())
            except Exception:
                pass
        self._small_btn(b, '붙여넣기', paste).grid(row=4, column=1, padx=(6, 0))
        status = tk.Label(b, text='', font=(F, 9), bg=C['white'], fg=C['subtext'],
                          wraplength=380, justify='left')
        status.grid(row=6, column=0, columnspan=3, sticky='w', pady=(8, 0))

        def save():
            tok = generator.clean_token(ent.get())
            if not tok:
                status.config(text='❌ sk-ant-oat 로 시작하는 토큰이 아닙니다. 검은 창의 토큰 전체를 복사하세요.',
                              fg=C['red']); return
            CONFIG[f'token_{n}'] = tok
            app_config.save_config()
            status.config(text='확인 중… (10초 정도)', fg=C['subtext'])
            def work():
                ok, msg = generator.verify_account(n)
                self._acc_login[n] = (ok, '토큰 연결됨' if ok else '토큰 오류')
                self._ui(lambda: (status.winfo_exists() and status.config(
                    text=('✅ ' if ok else '❌ ') + msg, fg=C['green'] if ok else C['red']),
                    self._show_acc(n)))
            threading.Thread(target=work, daemon=True).start()

        def remove():
            CONFIG[f'token_{n}'] = ''
            app_config.save_config()
            ent.delete(0, tk.END)
            status.config(text='연결을 해제했습니다.', fg=C['subtext'])
            self._check_accounts()

        bf = tk.Frame(w, bg=C['bg'])
        bf.pack(fill=tk.X, padx=12, pady=10)
        _btn(bf, '저장하고 확인', save, bg=C['accent'], fs=10, px=20, py=6).pack(side=tk.RIGHT)
        tk.Button(bf, text='연결 해제', command=remove, font=(F, 9), bg=C['light'],
                  relief=tk.FLAT, padx=12, pady=5, cursor='hand2').pack(side=tk.RIGHT, padx=6)

    def _login_account(self, n):
        if not generator.find_claude_cli():
            self._offer_install(); return
        try:
            generator.open_login(n)
        except Exception as e:
            messagebox.showerror('로그인', str(e), parent=self.ai_win); return
        self.acc_status[n].config(text='● 로그인 창에서 진행 후 [다시 확인]', fg=C['orange'])

    # ── 진행 표시 ──

    def _prog_set(self, key, text, frac=None):
        if not (self.ai_win and self.ai_win.winfo_exists()):
            return
        row = self.prog_rows.get(key)
        if row is None:
            f = tk.Frame(self.prog_area, bg=C['white'])
            f.pack(fill=tk.X, pady=(6, 0))
            lb = tk.Label(f, text='', font=(F, 9), bg=C['white'], fg=C['text'], anchor='w')
            lb.pack(fill=tk.X)
            track = tk.Frame(f, bg=C['light'], height=4)
            track.pack(fill=tk.X, pady=(3, 0))
            track.pack_propagate(False)
            bar = tk.Frame(track, bg=C['accent'], height=4)
            bar.place(x=0, y=0, relwidth=0.0, height=4)
            row = self.prog_rows[key] = (f, lb, bar)
        _, lb, bar = row
        lb.config(text=text)
        if frac is not None:
            bar.place(x=0, y=0, relwidth=max(0.0, min(1.0, frac)), height=4)

    def _prog_clear(self, key=None):
        for k in ([key] if key is not None else list(self.prog_rows)):
            row = self.prog_rows.pop(k, None)
            if row and row[0].winfo_exists():
                row[0].destroy()

    def _ai_log_msg(self, msg, level='inf'):
        self._log(msg, level)
        if not (self.ai_win and self.ai_win.winfo_exists()):
            return
        ts = datetime.now().strftime('%H:%M')
        self.ai_log.config(state=tk.NORMAL)
        self.ai_log.insert(tk.END, f'{ts}  {msg}\n', level)
        self.ai_log.see(tk.END)
        self.ai_log.config(state=tk.DISABLED)

    # ── 프로젝트 ──

    def _current_project(self):
        name = self.project_var.get()
        return next((p for p in self.projects if p.name == name), None)

    def _refresh_projects(self):
        self.projects = proj.list_projects()
        names = [p.name for p in self.projects]
        self.project_dd.set_values(names)
        last = CONFIG.get('last_project')
        if self.project_var.get() in names:
            pass
        elif last in names:
            self.project_var.set(last)
        elif names:
            self.project_var.set(names[0])
        else:
            self.project_var.set('')
        self._on_project()

    def _ui_settings(self):
        """화면 값 → 프로젝트 설정.json에 저장할 값"""
        try:
            n = max(1, min(30, int(self.count_var.get())))
        except ValueError:
            n = 10
        return {
            '편수': n,
            '접두어': self.prefix_var.get(), '상단키워드': self.keyword_var.get(),
            '하단프리셋': [name for var, (name, _) in zip(self.footer_preset_vars, FOOTER_PRESETS)
                       if var.get()],
            '하단구분선': bool(self.footer_var.get()),
            '원노트섹션': self.section_var.get() if self.section_var.get() != '(자동)' else '',
            '계정': self.account_var.get(),
            '자동내보내기': bool(self.auto_export_var.get()),
            '자동다시쓰기': bool(self.autofix_var.get()),
            '한도대기': bool(self.wait_limit_var.get()),
        }

    def _save_project_ui(self):
        """지금 화면 설정을 현재 프로젝트 기본값으로 자동 저장 (따로 저장 버튼 없음)"""
        name = getattr(self, '_loaded_project', None)
        p = next((x for x in self.projects if x.name == name), None)
        if not p or not (self.ai_win and self.ai_win.winfo_exists()):
            return
        try:
            proj.save_settings(p, self._ui_settings())
        except Exception as e:
            self._log(f'설정 저장 실패: {e}', 'err')

    def _apply_section(self, name):
        if name and name in self._section_map:
            self.section_var.set(name)
            self._on_section(None)
            self._pending_section = None
        else:
            self._pending_section = name   # 섹션 목록이 아직 안 왔으면 받은 뒤 적용

    def _on_project(self):
        if getattr(self, '_loaded_project', None) and self._loaded_project != self.project_var.get():
            self._save_project_ui()
        p = self._current_project()
        if not p:
            self.project_info.config(text='새 프로젝트를 만드세요')
            return
        try:
            s = p.settings
        except Exception as e:
            self.project_info.config(text=f'설정.json 오류: {e}')
            return
        self.project_info.config(text=f"{s.get('편수')}편 · 지침 파일 {len(p.guide_files())}개")
        # 프로젝트 기본값을 메인 화면 접두어/키워드/하단 프리셋에 반영
        self.prefix_var.set(s.get('접두어', '') or '')
        self.keyword_var.set(s.get('상단키워드', '') or '')
        wanted = set(s.get('하단프리셋') or [])
        for var, (name, _) in zip(self.footer_preset_vars, FOOTER_PRESETS):
            var.set(name in wanted)
        self.direct_hint.config(text=s.get('입력안내', ''))
        self.count_var.set(str(s.get('편수', 10)))
        self.footer_var.set(bool(s.get('하단구분선', False)))
        if s.get('계정'):
            self.account_var.set(str(s.get('계정')))
        self.auto_export_var.set(bool(s.get('자동내보내기', True)))
        self.autofix_var.set(bool(s.get('자동다시쓰기', True)))
        self.wait_limit_var.set(bool(s.get('한도대기', True)))
        self._apply_section(s.get('원노트섹션', ''))
        self._loaded_project = p.name
        if not s.get('엑셀사용', True):
            self.source_var.set('direct')
            self._on_source()
        if CONFIG.get('last_project') != p.name:
            CONFIG['last_project'] = p.name
            app_config.save_config()
        if self.cases:
            self._fill_cases()

    def _open_project_folder(self):
        p = self._current_project()
        app_config.open_folder(p.folder if p else app_config.PROJECTS_DIR)

    def _new_project(self):
        from tkinter import simpledialog
        name = simpledialog.askstring('새 프로젝트', '프로젝트 이름 (폴더 이름이 됩니다)\n예) 오엔_블로그칼럼',
                                      parent=self.ai_win)
        if not name:
            return
        try:
            p = proj.create_project(name)
        except Exception as e:
            messagebox.showerror('오류', str(e), parent=self.ai_win); return
        self._refresh_projects()
        self.project_var.set(p.name)
        self._on_project()
        self._guide_manager()
        messagebox.showinfo('새 프로젝트',
                            "지침 관리 창이 열렸습니다.\n\n"
                            "1) claude.ai 프로젝트 지침을 복사(Ctrl+C) → [복사해 둔 내용으로 전체 바꾸기]\n"
                            "2) 프로젝트 소스 파일은 [가져오기] 또는 [+ 새 파일]\n"
                            "3) 필요하면 [⚙ 자동 검사 규칙]\n\n"
                            "편수·하단 문구·섹션 등은 AI 원고 생성 창에서 고르면 자동 저장됩니다.", parent=self.ai_win)

    # ── 엑셀 ──

    def _on_source(self):
        self.excel_frame.pack_forget(); self.direct_frame.pack_forget()
        f = self.excel_frame if self.source_var.get() == 'excel' else self.direct_frame
        f.pack(fill=tk.BOTH, expand=True, padx=12, pady=(8, 0))

    def _pick_excel(self):
        path = filedialog.askopenfilename(title='사례 엑셀 선택', parent=self.ai_win,
                                          filetypes=[('엑셀', '*.xlsx *.xlsm'), ('전체', '*.*')])
        if path:
            self._load_excel(path)

    def _load_excel(self, path):
        try:
            self.cases = excel_source.load_cases(path)
        except Exception as e:
            messagebox.showerror('엑셀 오류', f'엑셀을 읽지 못했습니다.\n(엑셀에서 열려 있으면 닫고 다시 시도)\n\n{e}',
                                 parent=self.ai_win)
            return
        if not self.cases:
            messagebox.showwarning('엑셀', "'사건번호'와 '요약' 머리글이 있는 시트를 찾지 못했습니다.",
                                   parent=self.ai_win)
            return
        CONFIG['excel_path'] = path
        app_config.save_config()
        total = sum(len(v) for v in self.cases.values())
        self.excel_lbl.config(text=f'{os.path.basename(path)} · 사례 {total}건')
        self.sheet_dd.set_values(list(self.cases.keys()))
        if self.sheet_var.get() not in self.cases:
            self.sheet_var.set(next(iter(self.cases)))
        self._on_sheet()

    def _on_sheet(self):
        cases = self.cases.get(self.sheet_var.get(), [])
        cats = sorted({c.category for c in cases if c.category})
        self.cat_dd.set_values(['(전체)'] + cats)
        if self.cat_var.get() not in cats:
            self.cat_var.set('(전체)')
        self._fill_cases()

    def _is_used(self, case, used):
        if case.key in used:
            return True
        return bool(CONFIG.get('excel_used_check', True) and case.used_mark)

    def _fill_cases(self):
        if not (self.ai_win and self.ai_win.winfo_exists()):
            return
        for w in self.case_frame.winfo_children():
            w.destroy()
        self.case_checks = []
        p = self._current_project()
        used = excel_source.load_usage(p.name) if p else set()
        cat = self.cat_var.get()
        n_used = 0
        for c in self.cases.get(self.sheet_var.get(), []):
            if cat != '(전체)' and c.category != cat:
                continue
            is_used = self._is_used(c, used)
            n_used += is_used
            if is_used and self.unused_var.get():
                continue
            var = tk.BooleanVar(value=False)
            head = ' · '.join(x.replace('\n', ', ') for x in (c.crime, c.result1, c.result2) if x)
            text = f'{head}    {c.short()[:48]}'
            tk.Checkbutton(self.case_frame, text=text, variable=var, command=self._update_pick,
                           font=(F, 9), bg=C['white'], fg='#94a3b8' if is_used else C['text'],
                           activebackground=C['bg'], selectcolor=C['white'], anchor='w',
                           padx=8, pady=1).pack(fill=tk.X, padx=4)
            self.case_checks.append((var, c))
        self.case_canvas.yview_moveto(0)
        self.case_count.config(text=f'{len(self.case_checks)}건 · 사용함 {n_used}건')
        self._update_pick()

    def _update_pick(self):
        n = sum(1 for v, _ in self.case_checks if v.get())
        self.pick_lbl.config(text=f'{n}건 선택' if n else '')

    def _auto_pick(self):
        try:
            n = int(self.pick_n.get())
        except ValueError:
            n = 1
        for i, (v, _) in enumerate(self.case_checks):
            v.set(i < n)
        self._update_pick()

    def _clear_pick(self):
        for v, _ in self.case_checks:
            v.set(False)
        self._update_pick()

    # ── 생성 ──

    def _collect_jobs(self, auto_n=None):
        """선택한 사례(또는 직접 입력) → [(제목, 사건자료, Case 또는 None)]"""
        if self.source_var.get() == 'excel':
            if auto_n:
                self._fill_cases()
                for i, (v, _) in enumerate(self.case_checks):
                    v.set(i < auto_n)
                self._update_pick()
            sel = [c for v, c in self.case_checks if v.get()]
            return [(c.crime.split('\n')[0], c.to_prompt(), c) for c in sel]
        text = self.direct_txt.get('1.0', tk.END).strip()
        return [('직접입력', text, None)] if len(text) >= 5 else []

    def _start_generate(self, unattended=False, auto_n=None):
        """unattended=True: 예약 실행 (확인 창 없이 바로 시작)"""
        if self.generating:
            return False
        w = self.ai_win
        warn = (lambda t, m: self._ai_log_msg(f'{t}: {m}', 'err')) if unattended else \
            (lambda t, m: messagebox.showwarning(t, m, parent=w))
        p = self._current_project()
        if not p:
            warn('경고', '프로젝트를 먼저 고르세요.'); return False
        self._save_project_ui()
        try:
            p.build_system_prompt()
        except Exception as e:
            warn('지침 오류', str(e)); return False
        if not generator.find_claude_cli():
            if unattended:
                warn('오류', 'Claude Code가 없습니다')
            else:
                self._offer_install()
            return False

        jobs = self._collect_jobs(auto_n)
        if not jobs:
            warn('경고', '사례를 체크하세요. (또는 [위에서 N개 선택])' if self.source_var.get() == 'excel'
                 else '사건자료(또는 주제)를 입력하세요.')
            return False

        acc = self.account_var.get()
        accounts = [1, 2] if acc == 'both' else [int(acc)]
        bad = [n for n in accounts if not generator.check_login(n)[0]]
        if bad:
            warn('로그인 필요', ', '.join(f'계정 {n}' for n in bad) +
                 ' 로그인이 필요합니다. [로그인] 버튼으로 로그인한 뒤 다시 시도하세요.')
            self._check_accounts()
            return False

        auto_export = self.auto_export_var.get()
        if auto_export and not self.is_connected:
            if unattended:
                self.is_connected = self.onenote.connect()   # 밤새 꺼져 있었을 수 있어 다시 연결
                auto_export = self.is_connected
                if not auto_export:
                    self._ai_log_msg('OneNote 연결 실패 → 원고만 파일로 저장합니다', 'warn')
            elif not messagebox.askyesno('확인', 'OneNote가 연결되지 않았습니다.\n원고만 만들고 파일로 저장할까요?',
                                         parent=w):
                return False
            else:
                auto_export = False

        n = int(p.settings.get('편수', 10) or 10)
        if not unattended:
            acc_txt = '계정 1 + 2 동시' if len(accounts) == 2 else f'계정 {accounts[0]}'
            msg = (f'{len(jobs)}건 × {n}편 = 원고 {len(jobs) * n}편을 생성합니다.\n'
                   f'{acc_txt} · 건당 수 분 걸립니다.')
            if auto_export:
                msg += f'\n\n완료되는 대로 OneNote "{self.section_var.get()}"에 내보냅니다.'
            if not messagebox.askyesno('원고 생성', msg, parent=w):
                return False

        opts = {
            'prefix': self.prefix_var.get().strip(),
            'keyword': self.keyword_var.get().strip(),
            'add_footer': self.footer_var.get(),
            'footers': [lines for var, (_, lines) in zip(self.footer_preset_vars, FOOTER_PRESETS)
                        if var.get()],
            'autofix': self.autofix_var.get(),
            'wait_limit': self.wait_limit_var.get(),
        }
        self.generating = True
        self.gen_stop.clear()
        _keep_awake(True)
        self.gen_btn.config(state=tk.DISABLED, text='생성 중…')
        self.stop_btn.config(state=tk.NORMAL)
        self.export_btn.config(state=tk.DISABLED)
        threading.Thread(target=self._generate_run,
                         args=(p, jobs, accounts, auto_export, opts), daemon=True).start()
        return True

    def _stop_generate(self):
        self.gen_stop.set()
        self._ai_log_msg('중지 요청 — 진행 중인 건이 끝나면 멈춥니다', 'warn')

    def _ui(self, fn):
        self.root.after(0, fn)

    def _wait_limit(self, key, tag, resets_at, log):
        """한도가 풀릴 때까지 기다림. 풀리면 True, 중지/너무 길면 False"""
        now = time.time()
        target = (resets_at + LIMIT_GRACE) if resets_at else now + 30 * 60   # 시각을 모르면 30분 뒤 다시 시도
        if target - now > 12 * 3600:
            log(f'{tag} — 한도가 {datetime.fromtimestamp(target):%m/%d %H:%M}에 풀려서 이 계정은 멈춥니다', 'err')
            return False
        at = datetime.fromtimestamp(target).strftime('%H:%M')
        log(f'{tag} — 사용량 한도 도달 · {at}에 이어서 씁니다', 'warn')
        while time.time() < target:
            if self.gen_stop.is_set():
                return False
            left = int(target - time.time())
            self._ui(lambda left=left: self._prog_set(
                key, f'{tag}  ·  한도 대기 중 · {at}에 다시 시작 (남은 {left // 3600}시간 {left % 3600 // 60}분)', 0.0))
            time.sleep(5)
        return True

    def _generate_run(self, p, jobs, accounts, auto_export, opts):
        """계정 수만큼 작업자를 띄워 사례를 나눠 처리. OneNote 내보내기는 한 번에 하나씩."""
        import queue
        out_dir = os.path.join(app_config.OUTPUT_DIR,
                               datetime.now().strftime('%Y%m%d') + '_' + p.name)
        os.makedirs(out_dir, exist_ok=True)
        q = queue.Queue()
        for i, job in enumerate(jobs, 1):
            q.put((i, job))
        total = len(jobs)
        stats = {'ok': 0, 'fail': 0, 'done': 0}
        lock = threading.Lock()          # 통계·사용기록
        export_lock = threading.Lock()   # OneNote는 동시에 쓰지 않음
        multi = len(accounts) > 1
        per = int(p.settings.get('편수', 10) or 10)
        log = lambda m, l='inf': self._ui(lambda: self._ai_log_msg(m, l))

        def finish(ok):
            with lock:
                stats['ok' if ok else 'fail'] += 1
                stats['done'] += 1
                d = stats['done']
            self._ui(lambda: self.prog_status.config(text=f'생성 {d}/{total}'))

        def progress_fn(key, tag, t0, total_n):
            def on_progress(stage, txt):
                sec = int(time.time() - t0)
                el = f'{sec // 60}분 {sec % 60:02d}초'
                if stage == 'think' or not txt:
                    self._ui(lambda: self._prog_set(key, f'{tag}  ·  구상 중  ·  {el}', 0.02))
                    return
                k = len(re.findall(r'(?m)^#[ \t]+\d+[ \t]*$', txt))
                k = max(1, min(k, total_n))
                head = f'원고 {k}/{total_n}편 작성 중' if total_n > 1 else '작성 중'
                msg = f'{tag}  ·  {head}  ·  {len(txt):,}자  ·  {el}'
                self._ui(lambda: self._prog_set(key, msg, (k - 0.5) / total_n))
            return on_progress

        def worker(account):
            who = f'계정{account} ' if multi else ''
            key = account
            while not self.gen_stop.is_set():
                try:
                    idx, job = q.get_nowait()
                except queue.Empty:
                    return
                title, case_text, case = job[:3]
                saved = job[3] if len(job) > 3 else ''   # 한도로 끊겼을 때 이미 받은 원고
                tag = f'{who}[{idx}/{total}] {title}'
                log(f'{tag} — ' + ('이어서 시작 (받아 둔 원고 있음)' if saved else '시작'))
                t0 = time.time()
                self._ui(lambda tag=tag: self._prog_set(key, f'{tag}  ·  시작 중…', 0.0))
                try:
                    if saved:
                        text = saved
                    else:
                        text = generator.generate(p, case_text, stop_flag=self.gen_stop, account=account,
                                                  progress=progress_fn(key, tag, t0, per))
                except generator.LimitError as e:
                    self._ui(lambda: self._show_acc(account))
                    # 끊기기 전까지 다 쓴 원고는 버리지 않음 (마지막 1편은 쓰다 만 것이라 제외)
                    done_arts = proj.split_articles(e.partial or '')[:-1]
                    keep = proj.join_articles(done_arts) if done_arts else ''
                    if keep:
                        log(f'{tag} — 한도로 끊김 · 다 쓴 {len(done_arts)}편은 보관, 나머지만 이어서 씁니다', 'warn')
                    q.put((idx, (title, case_text, case, keep)))   # 다시 대기열로 (다른 계정이 먼저 가져갈 수도 있음)
                    if opts['wait_limit'] and self._wait_limit(key, f'계정{account}', e.resets_at, log):
                        continue
                    log(f'계정{account} — 사용량 한도로 멈춤', 'err')
                    self._ui(lambda: self._prog_clear(key))
                    return
                except Exception as e:
                    log(f'{tag} 실패: {e}', 'err')
                    partial = getattr(e, 'partial', '')
                    if partial and len(partial) > 200:
                        safe = (case.key if case else '직접입력').replace('/', '_').replace('\\', '_')
                        pp = os.path.join(out_dir, datetime.now().strftime('%H%M%S') + f'_{safe}_미완성.md')
                        with open(pp, 'w', encoding='utf-8') as f:
                            f.write(partial)
                        log(f'  써진 부분까지 저장: {os.path.basename(pp)}', 'warn')
                    self._ui(lambda: self._prog_clear(key))
                    finish(False)
                    if '중지' in str(e) or '로그인' in str(e):
                        return
                    continue

                # 원고가 덜 왔으면(예: 7편) 빠진 번호만 이어서 쓰기
                for attempt in range(2):
                    arts_raw = proj.split_articles(text)
                    have = {n for n, _ in arts_raw}
                    missing = [str(i) for i in range(1, per + 1) if str(i) not in have]
                    if not missing or self.gen_stop.is_set():
                        break
                    info = generator.last_info()
                    why = ('사용량 한도로 끊겼던 사례' if (saved and attempt == 0) else
                           f'종료: {info.get("stop", "?")}, 출력 {info.get("out_tokens") or "?"}토큰')
                    log(f'{tag} — {len(arts_raw)}편만 받음 ({why}) → {", ".join(missing)}번만 이어서 씁니다', 'warn')
                    try:
                        more = generator.write_missing(
                            p, case_text, missing, [b.strip().split('\n')[0] for _, b in arts_raw],
                            stop_flag=self.gen_stop, account=account,
                            progress=progress_fn(key, f'{tag} 이어 쓰기', time.time(), len(missing)))
                    except generator.LimitError as e:
                        if opts['wait_limit'] and self._wait_limit(key, f'계정{account}', e.resets_at, log):
                            continue
                        break
                    except Exception as e:
                        log(f'{tag} — 이어 쓰기 실패: {e}', 'err')
                        break
                    add = [(n, b) for n, b in proj.split_articles(more) if n in missing]
                    merged = dict(arts_raw); merged.update(add)
                    text = proj.join_articles(sorted(merged.items(), key=lambda x: int(x[0])))

                case_id = case.case_id if case else ''
                issues = proj.check_output(text, p, case_id)
                # 검사에 걸린 원고는 그 편만 다시 쓰기 (1회)
                if opts['autofix']:
                    bad = {}
                    for num, msg in issues:
                        if num and not msg.startswith('[경고]'):
                            bad.setdefault(num, []).append(msg)
                    if bad:
                        arts_raw = proj.split_articles(text)
                        titles = {n: b.strip().split('\n')[0] for n, b in arts_raw}
                        for num, probs in bad.items():
                            if self.gen_stop.is_set():
                                break
                            log(f'{tag} — # {num} 다시 쓰기 ({probs[0]})', 'warn')
                            try:
                                new = generator.rewrite_one(
                                    p, case_text, num, dict(arts_raw).get(num, ''),
                                    [t for n, t in titles.items() if n != num], probs,
                                    stop_flag=self.gen_stop, account=account,
                                    progress=progress_fn(key, f'{tag} # {num} 다시 쓰기', time.time(), 1))
                                nb = proj.split_articles(new)
                                if nb:
                                    arts_raw = [(n, nb[0][1] if n == num else b) for n, b in arts_raw]
                            except Exception as e:
                                log(f'{tag} — # {num} 다시 쓰기 실패: {e}', 'err')
                        text = proj.join_articles(arts_raw)
                        issues = proj.check_output(text, p, case_id)
                for num, msg in issues:
                    log(f'  ⚠ {"# " + num + " " if num else ""}{msg}', 'warn')

                stamp = datetime.now().strftime('%H%M%S')
                safe = (case.key if case else '직접입력').replace('/', '_').replace('\\', '_')
                fpath = os.path.join(out_dir, f'{stamp}_{safe}.md')
                with open(fpath, 'w', encoding='utf-8') as f:
                    f.write(text)

                arts = ArticleParser.parse(text)
                if not arts:
                    log(f'{tag}: 원고 형식을 읽지 못했습니다 → {os.path.basename(fpath)}', 'err')
                    self._ui(lambda: self._prog_clear(key))
                    finish(False)
                    continue
                meta = case_meta(case, case_text)
                group = f'{stamp}_{safe}'
                for a in arts:
                    a['meta'] = meta
                    a['gen'] = {'project': p.name, 'case_text': case_text, 'account': account,
                                'group': group, 'file': fpath}
                self._ui(lambda a=arts: self._add_articles(a))

                exported = True
                if auto_export:
                    self._ui(lambda tag=tag, n=len(arts): self._prog_set(
                        key, f'{tag}  ·  OneNote에 {n}편 넣는 중…', 1.0))
                    with export_lock:
                        res = self._export_articles(arts, opts['prefix'], opts['keyword'],
                                                    opts['footers'], opts['add_footer'])
                    exported = res['fail'] == 0
                    if not exported:
                        log(f"{tag} — OneNote {res['fail']}편 실패 (글 목록에서 다시 내보내기)", 'err')
                if case and exported:
                    with lock:
                        excel_source.record_usage(p.name, case)
                sec = int(time.time() - t0)
                log(f'{tag} 완료 ({len(arts)}편 · {sec // 60}분 {sec % 60:02d}초)', 'ok')
                self._ui(lambda: (self._prog_clear(key), self._show_acc(account)))
                finish(True)

        threads = [threading.Thread(target=worker, args=(a,), daemon=True) for a in accounts]
        for t in threads:
            t.start()
        for t in threads:
            t.join()
        self._ui(lambda: self._generate_done(stats['ok'], stats['fail']))

    def _add_articles(self, arts):
        self.articles.extend(arts)
        self._rebuild_checks()
        self.count_lbl.config(text=f'{len(self.articles)}개')

    def _generate_done(self, ok, fail):
        self.generating = False
        if not self.schedule:
            _keep_awake(False)
        self._prog_clear()
        if self.ai_win and self.ai_win.winfo_exists():
            self.gen_btn.config(state=tk.NORMAL, text='원고 생성')
            self.stop_btn.config(state=tk.DISABLED)
        if self.is_connected:
            self.export_btn.config(state=tk.NORMAL)
        self.prog_status.config(text='')
        self._ai_log_msg(f'끝: 성공 {ok}건 / 실패 {fail}건', 'ok' if not fail else 'warn')
        self._set_status(f'AI 생성 완료: {ok}건', C['green'] if not fail else C['orange'])
        self._fill_cases()
        try:
            self.root.bell()
        except Exception:
            pass

    # ── 로톡용 복사 (대제목·소제목 그대로) ──

    def _first_checked(self):
        sel = [i for i, v in enumerate(self.check_vars) if v.get()]
        if not sel:
            messagebox.showwarning('로톡용 복사', '복사할 글을 하나 체크하세요.'); return None
        if len(sel) > 1:
            self._set_status(f'여러 개가 체크돼 있어 맨 위 글을 복사합니다 ({len(sel)}개 중 1번째)', C['orange'])
        return self.articles[sel[0]]

    def _copy_lawtalk(self):
        import lawtalk
        art = self._first_checked()
        if not art:
            return
        kw = fill_template(self.keyword_var.get().strip(), art.get('meta'))
        footers = [lines for var, (_, lines) in zip(self.footer_preset_vars, FOOTER_PRESETS) if var.get()]
        frag, text = lawtalk.article_html(art, keyword=kw, footers=footers)
        if lawtalk.copy_to_clipboard(frag, text):
            n2, n3 = frag.count('<h2>'), frag.count('<h3>')
            self._set_status(f'로톡용 복사 완료 (대제목 {n2}개 · 소제목 {n3}개) → 로톡 본문에 Ctrl+V', C['green'])
        else:
            self.root.clipboard_clear(); self.root.clipboard_append(text)
            self._set_status('서식 복사가 안 돼서 글자만 복사했습니다 (윈도우에서만 서식 복사 가능)', C['orange'])

    def _copy_title(self):
        art = self._first_checked()
        if not art:
            return
        title = fill_template(self.prefix_var.get().strip(), art.get('meta'))
        title = (title + ' ' if title else '') + art.get('title', '')
        self.root.clipboard_clear(); self.root.clipboard_append(title)
        self._set_status(f'제목 복사: {title[:40]}', C['green'])

    # ── 로톡 링크 기록 (엑셀 + OneNote 페이지) ──

    def _full_title(self, art):
        pre = fill_template(self.prefix_var.get().strip(), art.get('meta'))
        return (pre + ' ' if pre else '') + art.get('title', '')

    def _link_dialog(self):
        """OneNote 섹션의 페이지를 직접 읽어서 링크 기록 — 매크로를 껐다 켜도, 예전 페이지도 언제든 가능"""
        if not self.is_connected:
            messagebox.showwarning('링크 기록', 'OneNote가 연결되어 있어야 합니다.'); return
        w = tk.Toplevel(self.root)
        w.title('로톡 링크 기록')
        w.geometry('960x640')
        w.configure(bg=C['bg'])
        w.transient(self.root)
        top = self._card(w, '로톡에 올린 글 주소를 붙여넣으세요 — 기록 엑셀과 OneNote 페이지에 함께 기록')
        bar = tk.Frame(top, bg=C['white']); bar.pack(fill=tk.X, padx=12, pady=8)
        tk.Label(bar, text='섹션', font=(F, 9), bg=C['white'], fg=C['subtext']).pack(side=tk.LEFT)
        sec_var = tk.StringVar(value=self.section_var.get() if self.section_var.get() in self._section_map
                               else (next(iter(self._section_map), '')))
        sec_dd = self._dropdown(bar, width=30, var=sec_var, command=lambda: load())
        sec_dd.set_values(list(self._section_map.keys()))
        sec_var.set(sec_var.get())
        sec_dd.pack(side=tk.LEFT, padx=(6, 12))
        tk.Label(bar, text='찾기', font=(F, 9), bg=C['white'], fg=C['subtext']).pack(side=tk.LEFT)
        q_var = tk.StringVar()
        qe = tk.Entry(bar, textvariable=q_var, font=(F, 9), width=18, relief=tk.FLAT, bg=C['bg'])
        qe.pack(side=tk.LEFT, padx=6, ipady=3)
        self._small_btn(bar, '↻ 다시 불러오기', lambda: load()).pack(side=tk.RIGHT)

        # 기록 엑셀 형식 설정 (자동 저장)
        cfgr = tk.Frame(top, bg=C['white']); cfgr.pack(fill=tk.X, padx=12, pady=(0, 8))
        tk.Label(cfgr, text='아이디', font=(F, 9), bg=C['white'], fg=C['subtext']).pack(side=tk.LEFT)
        id_var = tk.StringVar(value=CONFIG.get('lawtalk_id', ''))
        tk.Entry(cfgr, textvariable=id_var, font=(F, 9), width=12, relief=tk.FLAT, bg=C['bg']
                 ).pack(side=tk.LEFT, padx=(6, 12), ipady=3)
        tk.Label(cfgr, text='숨길 제목', font=(F, 9), bg=C['white'], fg=C['subtext']).pack(side=tk.LEFT)
        hide_var = tk.StringVar(value=CONFIG.get('link_hide_prefix', '사용'))
        tk.Entry(cfgr, textvariable=hide_var, font=(F, 9), width=14, relief=tk.FLAT, bg=C['bg']
                 ).pack(side=tk.LEFT, padx=(6, 4), ipady=3)
        tk.Label(cfgr, text='(이 글자로 시작하면 숨김)', font=(F, 8), bg=C['white'],
                 fg='#94a3b8').pack(side=tk.LEFT)
        hide_var.trace_add('write', lambda *a: show())
        strip_var = tk.BooleanVar(value=bool(CONFIG.get('title_strip_result', False)))
        tk.Checkbutton(cfgr, text='끝의 결과 괄호 빼기', variable=strip_var, font=(F, 9), bg=C['white'],
                       activebackground=C['white'], selectcolor=C['white']).pack(side=tk.LEFT, padx=10)
        dash_var = tk.BooleanVar(value=bool(CONFIG.get('link_strip_dash', True)))
        tk.Checkbutton(cfgr, text='OneNote 제목 앞 ㅡ 지우기', variable=dash_var, font=(F, 9), bg=C['white'],
                       activebackground=C['white'], selectcolor=C['white']).pack(side=tk.LEFT)
        cfgr = tk.Frame(top, bg=C['white']); cfgr.pack(fill=tk.X, padx=12, pady=(0, 8))
        xl_var = tk.StringVar(value=CONFIG.get('upload_xlsx') or app_config.UPLOAD_XLSX)
        def pick_xl():
            pth = filedialog.asksaveasfilename(parent=w, title='링크를 이어서 기록할 엑셀 (기존 파일 선택 또는 새 이름)',
                                               defaultextension='.xlsx',
                                               initialfile=os.path.basename(xl_var.get()),
                                               filetypes=[('엑셀', '*.xlsx')], confirmoverwrite=False)
            if pth:
                xl_var.set(pth); xl_lbl.config(text=os.path.basename(pth))
        tk.Label(cfgr, text='기록 엑셀', font=(F, 9), bg=C['white'], fg=C['subtext']).pack(side=tk.LEFT)
        xl_lbl = tk.Label(cfgr, text=os.path.basename(xl_var.get()), font=(F, 9), bg=C['white'], fg=C['text'])
        xl_lbl.pack(side=tk.LEFT, padx=6)
        self._small_btn(cfgr, '바꾸기', pick_xl).pack(side=tk.LEFT)
        tk.Label(cfgr, text='기존 엑셀을 고르면 맨 아래에 이어서 추가 (날짜 | 아이디 | 사건번호 | 제목 | url)',
                 font=(F, 8), bg=C['white'], fg='#94a3b8').pack(side=tk.LEFT, padx=10)

        lc = tk.Frame(w, bg=C['white'], highlightbackground=C['light'], highlightthickness=1)
        lc.pack(fill=tk.BOTH, expand=True, padx=12, pady=(8, 0))
        hdr = tk.Frame(lc, bg=C['bg']); hdr.pack(fill=tk.X, padx=1, pady=1)
        for t, wd in (('사건번호', 14), ('OneNote 페이지 제목', 50), ('로톡 링크', 32)):
            tk.Label(hdr, text=t, font=(F, 9, 'bold'), bg=C['bg'], fg=C['subtext'], width=wd,
                     anchor='w').pack(side=tk.LEFT, padx=(8, 0), pady=5)
        wrap = tk.Frame(lc, bg=C['white']); wrap.pack(fill=tk.BOTH, expand=True)
        sb = tk.Scrollbar(wrap); sb.pack(side=tk.RIGHT, fill=tk.Y)
        cv = tk.Canvas(wrap, bg=C['white'], highlightthickness=0, yscrollcommand=sb.set)
        cv.pack(side=tk.LEFT, fill=tk.BOTH, expand=True); sb.config(command=cv.yview)
        inner = tk.Frame(cv, bg=C['white'])
        cv.create_window((0, 0), window=inner, anchor='nw')
        inner.bind('<Configure>', lambda e: cv.config(scrollregion=cv.bbox('all')))
        wheel = lambda e: cv.yview_scroll(-1 * (e.delta // 120), 'units')
        cv.bind('<Enter>', lambda e: cv.bind_all('<MouseWheel>', wheel))
        cv.bind('<Leave>', lambda e: cv.unbind_all('<MouseWheel>'))
        status = tk.Label(w, text='', font=(F, 9), bg=C['bg'], fg=C['subtext'])
        st = {'pages': [], 'typed': {}}   # typed: 페이지ID → 입력한 링크 (필터 바뀌어도 유지)
        rows = []

        def remember():
            for pg, e in rows:
                if e.winfo_exists() and e.cget('state') != 'disabled':
                    st['typed'][pg['id']] = e.get().strip()

        def show():
            remember()
            for c in inner.winfo_children():
                c.destroy()
            rows.clear()
            q = q_var.get().strip()
            n = 0
            for pg in st['pages']:
                if not pg['title'] or set(pg['title']) <= set('ㅡ-─ '):
                    continue   # 구분선 페이지
                hides = [x.strip() for x in hide_var.get().split(',') if x.strip()]
                if any(pg['title'].lstrip().startswith(h) for h in hides):
                    continue   # 예) '사용'으로 시작 = 이미 올린 글 (ㅡ 지워진 것)
                if q and q not in pg['title'] and q not in pg['case_no']:
                    continue
                r = tk.Frame(inner, bg=C['white']); r.pack(fill=tk.X, pady=2)
                tk.Label(r, text=pg['case_no'] or '-', font=(F, 9), bg=C['white'], width=14,
                         anchor='w').pack(side=tk.LEFT, padx=(8, 0))
                t = pg['title']
                tk.Label(r, text=t[:48] + ('…' if len(t) > 48 else ''), font=(F, 9), bg=C['white'],
                         width=50, anchor='w').pack(side=tk.LEFT, padx=(8, 0))
                e = tk.Entry(r, font=(F, 9), width=32, relief=tk.FLAT, bg=C['bg'])
                e.pack(side=tk.LEFT, ipady=3, padx=(8, 0))
                if pg['link']:
                    e.insert(0, pg['link']); e.config(state='disabled', disabledbackground=C['white'])
                else:
                    e.insert(0, st['typed'].get(pg['id'], ''))
                    def paste(e=e):
                        try:
                            e.delete(0, tk.END); e.insert(0, w.clipboard_get().strip())
                        except Exception:
                            pass
                    tk.Button(r, text='붙여넣기', command=paste, font=(F, 8), bg=C['light'],
                              relief=tk.FLAT, padx=6, cursor='hand2').pack(side=tk.LEFT, padx=4)
                rows.append((pg, e))
                n += 1
            cv.yview_moveto(0)
            status.config(text=f'표시 {n}개 / 전체 {len(st["pages"])}개 (숨길 제목: {hide_var.get() or "없음"})',
                          fg=C['subtext'])

        q_var.trace_add('write', lambda *a: show())

        def load():
            name = sec_var.get()
            sec = self._section_map.get(name)
            if not sec:
                status.config(text='섹션을 고르세요.'); return
            remember()
            status.config(text='OneNote에서 페이지 읽는 중… (페이지가 많으면 수십 초)', fg=C['orange'])
            def work():
                try:
                    pages = self.onenote.list_pages(sec.id)
                except Exception as ex:
                    msg = str(ex)
                    self._ui(lambda: status.winfo_exists() and status.config(text=f'불러오기 실패: {msg[:80]}', fg=C['red']))
                    return
                def done():
                    if w.winfo_exists():
                        st['pages'] = pages; show()
                self._ui(done)
            threading.Thread(target=work, daemon=True).start()

        def save():
            remember()
            items = []
            for pg in st['pages']:
                url = st['typed'].get(pg['id'], '')
                if url and not pg['link']:
                    if not re.match(r'^https?://', url):
                        messagebox.showwarning('링크 기록', f'주소 형식이 아닙니다:\n{url}', parent=w); return
                    items.append((pg, url))
            if not items:
                status.config(text='새로 넣은 링크가 없습니다.', fg=C['subtext']); return
            words = []
            CONFIG.update({'lawtalk_id': id_var.get().strip(), 'link_hide_prefix': hide_var.get().strip(),
                           'title_strip_result': bool(strip_var.get()), 'link_strip_dash': bool(dash_var.get()),
                           'upload_xlsx': '' if xl_var.get() == app_config.UPLOAD_XLSX else xl_var.get()})
            app_config.save_config()
            try:
                n = excel_source.record_upload(xl_var.get(), [
                    {'아이디': id_var.get().strip(), '사건번호': pg['case_no'], 'url': u,
                     '제목': excel_source.clean_title(pg['title'], words, strip_var.get())}
                    for pg, u in items])
            except PermissionError:
                messagebox.showerror('링크 기록', f'{os.path.basename(xl_var.get())} 가 엑셀에서 열려 있어 저장하지 못했습니다.\n닫고 다시 저장하세요.', parent=w)
                return
            except ValueError as ex:
                messagebox.showerror('링크 기록', str(ex), parent=w); return
            except Exception as ex:
                messagebox.showerror('링크 기록', f'엑셀 저장 실패: {ex}', parent=w); return
            status.config(text=f'엑셀에 {n}건 기록. OneNote 페이지에 넣는 중…', fg=C['orange'])
            def work():
                ok = 0
                for pg, u in items:
                    if self.onenote.add_link(pg['id'], u, strip_dash=dash_var.get()):
                        pg['link'] = u; ok += 1
                        if dash_var.get():
                            pg['title'] = onenote_com.strip_title_dash(pg['title'])
                        st['typed'].pop(pg['id'], None)
                msg = f'엑셀 {n}건 · OneNote {ok}건 기록 완료'
                if ok < len(items):
                    msg += f' ({len(items) - ok}건은 OneNote에 못 넣음 — 페이지가 지워졌는지 확인)'
                self._ui(lambda: (w.winfo_exists() and (show(), status.config(text=msg, fg=C['green'])),
                                  self._log(msg, 'ok')))
            threading.Thread(target=work, daemon=True).start()

        bf = tk.Frame(w, bg=C['bg']); bf.pack(fill=tk.X, padx=12, pady=10, side=tk.BOTTOM, before=lc)
        _btn(bf, '저장', save, bg=C['accent'], fs=10, px=24, py=6).pack(side=tk.RIGHT)
        tk.Button(bf, text='기록 엑셀 열기',
                  command=lambda: app_config.open_file(xl_var.get()) or
                  status.config(text='아직 기록이 없습니다.', fg=C['subtext']),
                  font=(F, 9), bg=C['light'], relief=tk.FLAT, padx=12, pady=5, cursor='hand2').pack(side=tk.RIGHT, padx=6)
        status.pack(side=tk.BOTTOM, anchor='w', padx=14, before=lc)
        load()

    # ── 1편 다시 쓰기 (메인 화면 글 목록) ──

    def _rewrite_selected(self):
        if self.generating:
            messagebox.showwarning('경고', '원고 생성이 끝난 뒤에 다시 쓰기를 하세요.'); return
        sel = [i for i, v in enumerate(self.check_vars) if v.get()]
        targets = [i for i in sel if self.articles[i].get('gen')]
        if not sel:
            messagebox.showwarning('경고', '다시 쓸 글을 체크하세요.'); return
        if not targets:
            messagebox.showwarning('경고', 'AI로 생성한 글만 다시 쓸 수 있습니다.\n(붙여넣기로 가져온 글은 사건자료 정보가 없음)')
            return
        replace = any(self.articles[i].get('page_id') for i in targets) and self.is_connected
        msg = f'{len(targets)}편을 한 편씩 다시 씁니다. (편당 1~2분)'
        if replace:
            msg += '\n\nOneNote에 이미 넣은 글은 새 글로 바꿉니다.\n(옛 페이지는 OneNote 휴지통으로 이동)'
        if not messagebox.askyesno('다시 쓰기', msg):
            return
        opts = (self.prefix_var.get().strip(), self.keyword_var.get().strip(),
                [lines for var, (_, lines) in zip(self.footer_preset_vars, FOOTER_PRESETS) if var.get()])
        self.generating = True
        self.export_btn.config(state=tk.DISABLED)
        threading.Thread(target=self._rewrite_thread, args=(targets, opts), daemon=True).start()

    def _rewrite_thread(self, targets, opts):
        prefix, keyword, footers = opts
        done = 0
        for k, i in enumerate(targets, 1):
            art = self.articles[i]
            g = art['gen']
            p = next((x for x in proj.list_projects() if x.name == g['project']), None)
            if not p:
                self._ui(lambda: self._log('프로젝트를 찾을 수 없습니다', 'err')); continue
            num = str(art.get('number', '1'))
            siblings = [a['title'] for a in self.articles
                        if a.get('gen', {}).get('group') == g['group'] and a is not art]
            self._ui(lambda k=k: self._set_status(f'다시 쓰는 중 {k}/{len(targets)}…', C['orange']))
            try:
                old = art.get('title', '') + '\n' + art.get('raw_content', '')
                new_text = generator.rewrite_one(p, g['case_text'], num, old,
                                                 siblings, account=g['account'])
            except Exception as e:
                self._ui(lambda e=e: self._log(f'다시 쓰기 실패: {e}', 'err')); continue
            new = ArticleParser.parse(new_text)
            if not new:
                self._ui(lambda: self._log('다시 쓴 원고 형식 오류', 'err')); continue
            na = new[0]
            na['number'] = num
            na['meta'], na['gen'] = art.get('meta'), g
            if art.get('page_id') and self.is_connected:
                res = self._export_articles([na], prefix, keyword, footers, False)
                if res['success']:
                    self.onenote.delete_page(art['page_id'])
            self.articles[i] = na
            done += 1
            # 원고 파일에도 반영
            try:
                with open(g['file'], encoding='utf-8') as f:
                    raw = proj.split_articles(f.read())
                raw = [(n, new_text.split('\n', 1)[1] if n == num else b) for n, b in raw]
                with open(g['file'], 'w', encoding='utf-8') as f:
                    f.write(proj.join_articles(raw))
            except Exception:
                pass
        self._ui(lambda: self._rewrite_done(done, len(targets)))

    def _rewrite_done(self, done, total):
        self.generating = False
        if self.is_connected:
            self.export_btn.config(state=tk.NORMAL)
        self._rebuild_checks()
        self._set_status(f'다시 쓰기 완료: {done}/{total}편', C['green'] if done == total else C['orange'])
        self._log(f'다시 쓰기 완료: {done}/{total}편', 'ok')

    # ── 예약 실행 ──

    def _open_schedule(self):
        w = tk.Toplevel(self.ai_win)
        w.title('예약 실행')
        w.configure(bg=C['bg'])
        w.transient(self.ai_win)
        w.resizable(False, False)
        card = self._card(w, '정해진 시간에 자동으로 원고 생성')
        b = tk.Frame(card, bg=C['white'])
        b.pack(fill=tk.X, padx=12, pady=(8, 12))
        tk.Label(b, text='시작 시간', font=(F, 9), bg=C['white'], fg=C['subtext'], width=9,
                 anchor='w').grid(row=0, column=0, sticky='w', pady=4)
        tf = tk.Frame(b, bg=C['white'])
        tf.grid(row=0, column=1, sticky='w')
        hh = tk.Entry(tf, width=3, font=(F, 10), relief=tk.FLAT, bg=C['bg'], justify='center')
        mm = tk.Entry(tf, width=3, font=(F, 10), relief=tk.FLAT, bg=C['bg'], justify='center')
        hh.insert(0, '02'); mm.insert(0, '00')
        hh.pack(side=tk.LEFT, ipady=3)
        tk.Label(tf, text=':', bg=C['white']).pack(side=tk.LEFT, padx=2)
        mm.pack(side=tk.LEFT, ipady=3)
        tk.Label(tf, text='  (24시간, 지난 시간이면 내일)', font=(F, 8), bg=C['white'],
                 fg='#94a3b8').pack(side=tk.LEFT)

        tk.Label(b, text='무엇을', font=(F, 9), bg=C['white'], fg=C['subtext'], width=9,
                 anchor='w').grid(row=1, column=0, sticky='nw', pady=4)
        mode = tk.StringVar(value='auto' if self.source_var.get() == 'excel' else 'now')
        mf = tk.Frame(b, bg=C['white'])
        mf.grid(row=1, column=1, sticky='w')
        af = tk.Frame(mf, bg=C['white'])
        af.pack(anchor='w')
        tk.Radiobutton(af, text='그 시간에 안 쓴 사례 위에서', variable=mode, value='auto',
                       font=(F, 9), bg=C['white'], activebackground=C['white']).pack(side=tk.LEFT)
        cnt = tk.Entry(af, width=4, font=(F, 9), relief=tk.FLAT, bg=C['bg'], justify='center')
        cnt.insert(0, '10')
        cnt.pack(side=tk.LEFT, ipady=2)
        tk.Label(af, text='건', font=(F, 9), bg=C['white']).pack(side=tk.LEFT, padx=(2, 0))
        tk.Radiobutton(mf, text='지금 체크해 둔 사례 / 입력한 내용', variable=mode, value='now',
                       font=(F, 9), bg=C['white'], activebackground=C['white']).pack(anchor='w')
        tk.Label(b, text='· 지금 화면의 프로젝트·계정·OneNote 설정 그대로 실행합니다\n'
                         '· 매크로를 켜 둔 채로 두세요. 예약~작업 끝까지 PC가 절전 모드로 가지 않게 합니다\n'
                         '· "한도에 걸리면 기다렸다 이어 쓰기"를 켜 두면 밤새 이어서 씁니다',
                 font=(F, 8), bg=C['white'], fg='#94a3b8', justify='left'
                 ).grid(row=2, column=0, columnspan=2, sticky='w', pady=(8, 0))

        def ok():
            try:
                h, m = int(hh.get()), int(mm.get())
                assert 0 <= h < 24 and 0 <= m < 60
                n = int(cnt.get()) if mode.get() == 'auto' else None
            except Exception:
                messagebox.showwarning('예약', '시간(0~23 : 0~59)과 건수를 숫자로 넣으세요.', parent=w); return
            if mode.get() == 'now' and not self._collect_jobs():
                messagebox.showwarning('예약', '체크한 사례(또는 입력 내용)가 없습니다.', parent=w); return
            now = datetime.now()
            at = now.replace(hour=h, minute=m, second=0, microsecond=0)
            if at <= now:
                from datetime import timedelta
                at += timedelta(days=1)
            self._save_project_ui()
            self.schedule = {'at': at, 'auto_n': n}
            _keep_awake(True)
            self._tick_schedule()
            self._ai_log_msg(f'예약: {at:%m/%d %H:%M} · ' + (f'안 쓴 사례 {n}건' if n else '체크한 사례'), 'inf')
            w.destroy()

        def cancel():
            self._cancel_schedule(); w.destroy()

        bf = tk.Frame(w, bg=C['bg'])
        bf.pack(fill=tk.X, padx=12, pady=10)
        _btn(bf, '예약하기', ok, bg=C['accent'], fs=10, px=22, py=6).pack(side=tk.RIGHT)
        if self.schedule:
            tk.Button(bf, text='예약 취소', command=cancel, font=(F, 9), bg=C['light'],
                      relief=tk.FLAT, padx=12, pady=5).pack(side=tk.RIGHT, padx=6)

    def _cancel_schedule(self):
        self.schedule = None
        if hasattr(self, 'sched_lbl') and self.sched_lbl.winfo_exists():
            self.sched_lbl.config(text='')
        if not self.generating:
            _keep_awake(False)
        self._ai_log_msg('예약 취소', 'warn')

    def _tick_schedule(self):
        sc = self.schedule
        if not sc:
            return
        left = (sc['at'] - datetime.now()).total_seconds()
        if left > 0:
            h, m = int(left // 3600), int(left % 3600 // 60)
            if self.sched_lbl.winfo_exists():
                self.sched_lbl.config(text=f'⏰ {sc["at"]:%H:%M} 예약 (남은 {h}시간 {m}분)  ')
            self.root.after(20000 if left > 60 else 1000, self._tick_schedule)
            return
        self.schedule = None
        self.sched_lbl.config(text='')
        if self.ai_win:
            self.ai_win.deiconify()
        self._ai_log_msg('예약 시간 — 원고 생성 시작', 'ok')
        if not self._start_generate(unattended=True, auto_n=sc['auto_n']):
            self._ai_log_msg('예약 실행을 시작하지 못했습니다 (위 메시지 참고)', 'err')
            _keep_awake(False)

    # ── 지침 관리 (파일 추가·삭제·수정, 검사 규칙) ──

    def _guide_manager(self):
        """지침 관리: 모든 지침을 이름으로 보여주고 추가·이름 변경·삭제·편집, 지침별 검사 규칙"""
        w = tk.Toplevel(self.ai_win)
        w.title('지침 관리')
        w.geometry('1000x700')
        w.minsize(800, 540)
        w.configure(bg=C['bg'])
        w.transient(self.ai_win)
        st = {'rows': [], 'cur': None, 'dirty': False}   # cur = (kind, project, filename)

        # ── 왼쪽: 지침 목록 ──
        left = tk.Frame(w, bg=C['white'], highlightbackground=C['light'], highlightthickness=1, width=300)
        left.pack(side=tk.LEFT, fill=tk.Y, padx=(12, 0), pady=12)
        left.pack_propagate(False)
        lt = tk.Frame(left, bg=C['bg']); lt.pack(fill=tk.X, padx=1, pady=1)
        tk.Label(lt, text='지침 목록', font=(F, 9, 'bold'), bg=C['bg'],
                 fg=C['subtext']).pack(side=tk.LEFT, padx=12, pady=6)
        lb = tk.Listbox(left, font=(F, 9), relief=tk.FLAT, bd=0, highlightthickness=0,
                        selectbackground='#dbeafe', selectforeground=C['text'], fg=C['text'],
                        activestyle='none', exportselection=False)
        lb.pack(fill=tk.BOTH, expand=True, padx=6, pady=6)
        lbtn = tk.Frame(left, bg=C['white']); lbtn.pack(fill=tk.X, padx=6, pady=(0, 4))
        lbtn2 = tk.Frame(left, bg=C['white']); lbtn2.pack(fill=tk.X, padx=6, pady=(0, 8))

        # ── 오른쪽 ──
        right = tk.Frame(w, bg=C['white'], highlightbackground=C['light'], highlightthickness=1)
        right.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=12, pady=12)
        rt = tk.Frame(right, bg=C['bg']); rt.pack(fill=tk.X, padx=1, pady=1)
        title = tk.Label(rt, text='', font=(F, 9, 'bold'), bg=C['bg'], fg=C['subtext'])
        title.pack(side=tk.LEFT, padx=12, pady=6)
        info = tk.Label(rt, text='', font=(F, 8), bg=C['bg'], fg='#94a3b8')
        info.pack(side=tk.RIGHT, padx=12)
        rbtn = tk.Frame(right, bg=C['bg']); rbtn.pack(fill=tk.X, padx=1, pady=1, side=tk.BOTTOM)
        body = tk.Frame(right, bg=C['white']); body.pack(fill=tk.BOTH, expand=True)
        editor = scrolledtext.ScrolledText(body, font=(F, 10), wrap=tk.WORD, relief=tk.FLAT,
                                           padx=14, pady=10, undo=True, highlightthickness=0)
        panel = tk.Frame(body, bg=C['white'])
        rv = {}

        def projects_now():
            return proj.list_projects()

        def refresh(select=None):
            """select = (kind, project_name, filename)"""
            lb.delete(0, tk.END)
            st['rows'] = []
            for p in projects_now():
                st['rows'].append(('project', p.name, None))
                lb.insert(tk.END, f'■  {p.name}')
                lb.itemconfig(tk.END, fg=C['text'])
                for f in p.guide_files():
                    fn = os.path.basename(f)
                    kind = '메인 지침' if fn == proj.MAIN_GUIDE else '참고'
                    st['rows'].append(('file', p.name, fn))
                    lb.insert(tk.END, f'      {kind} · {proj.file_title(f, 22)}')
                    lb.itemconfig(tk.END, fg=C['subtext'])
                r = p.settings.get('검사', {})
                n = (len(r.get('원고당_정확히1회', [])) + len(r.get('금지문구', [])) +
                     bool(r.get('제목_최대글자')) + bool(r.get('본문_최소글자')))
                st['rows'].append(('rules', p.name, None))
                lb.insert(tk.END, f'      ⚙ 자동 검사 규칙 ({n}개)' if n else '      ⚙ 자동 검사 규칙 (없음)')
                lb.itemconfig(tk.END, fg='#94a3b8')
            if not st['rows']:
                open_item(None); return
            idx = 0
            if select:
                for i, row in enumerate(st['rows']):
                    if row[:2] == select[:2] and (row[2] == select[2] or select[0] != 'file'):
                        idx = i; break
            lb.selection_clear(0, tk.END); lb.selection_set(idx); lb.see(idx)
            open_item(st['rows'][idx])

        def get_p(name):
            return next((x for x in projects_now() if x.name == name), None)

        def count():
            info.config(text=f"{len(editor.get('1.0', 'end-1c')):,}자" + ('  · 수정됨' if st['dirty'] else ''))

        def save_cur():
            cur = st['cur']
            if not cur:
                return
            kind, pname, fn = cur
            p = get_p(pname)
            if kind == 'file' and st['dirty'] and p:
                with open(os.path.join(p.folder, fn), 'w', encoding='utf-8') as f:
                    f.write(editor.get('1.0', 'end-1c'))
                st['dirty'] = False
                count()
            elif kind == 'rules' and p:
                save_rules(p)

        def clear_right():
            editor.pack_forget(); panel.pack_forget()
            for c in panel.winfo_children():
                c.destroy()
            for c in rbtn.winfo_children():
                c.destroy()

        def open_item(row):
            if st['cur'] and st['cur'] != row:
                save_cur()
            st['cur'] = row
            clear_right()
            if not row:
                title.config(text='지침이 없습니다'); info.config(text='')
                return
            kind, pname, fn = row
            p = get_p(pname)
            if kind == 'project':
                show_project(p)
            elif kind == 'rules':
                show_rules(p)
            else:
                show_file(p, fn)

        # 지침(프로젝트) 개요
        def show_project(p):
            title.config(text=p.name); info.config(text='')
            panel.pack(fill=tk.BOTH, expand=True, padx=16, pady=12)
            tk.Label(panel, text='지침 이름', font=(F, 9, 'bold'), bg=C['white']).pack(anchor='w')
            row = tk.Frame(panel, bg=C['white']); row.pack(fill=tk.X, pady=(4, 12))
            ent = tk.Entry(row, font=(F, 10), width=30, relief=tk.FLAT, bg=C['bg'])
            ent.insert(0, p.name); ent.pack(side=tk.LEFT, ipady=4)
            def rename():
                if self.generating:
                    messagebox.showwarning('지침', '원고 생성 중에는 바꿀 수 없습니다.', parent=w); return
                if ent.get().strip() == p.name:
                    return
                try:
                    np_ = proj.rename_project(p, ent.get())
                except Exception as e:
                    messagebox.showerror('지침', str(e), parent=w); return
                if CONFIG.get('last_project') == p.name:
                    CONFIG['last_project'] = np_.name; app_config.save_config()
                if self.project_var.get() == p.name:
                    self._loaded_project = None
                    self.project_var.set(np_.name)
                self._refresh_projects()
                st['cur'] = None
                refresh(('project', np_.name, None))
            self._small_btn(row, '이름 바꾸기', rename).pack(side=tk.LEFT, padx=6)
            files = [os.path.basename(f) for f in p.guide_files()]
            tk.Label(panel, text=f'파일 {len(files)}개 — 모두 함께 Claude에게 전달됩니다', font=(F, 9),
                     bg=C['white'], fg=C['subtext']).pack(anchor='w')
            for f in p.guide_files():
                fn = os.path.basename(f)
                tk.Label(panel, text=f"  · {'메인 지침' if fn == proj.MAIN_GUIDE else '참고'}  "
                                     f"{proj.file_title(f, 50)}", font=(F, 9), bg=C['white'],
                         fg=C['text']).pack(anchor='w', pady=1)
            tk.Label(panel, text='\n· 메인 지침을 새 버전으로 바꾸려면 왼쪽에서 "메인 지침"을 고르고 [복사해 둔 내용으로 전체 바꾸기]\n'
                                 '· 참고 파일(세부지침 등)은 아래 [+ 파일] 또는 [가져오기]\n'
                                 '· 이 지침으로 원고를 쓰려면 AI 원고 생성 창의 프로젝트에서 이 이름을 고르세요',
                     font=(F, 8), bg=C['white'], fg='#94a3b8', justify='left').pack(anchor='w')

        # 파일 편집
        def show_file(p, fn):
            kind = '메인 지침' if fn == proj.MAIN_GUIDE else '참고 파일'
            title.config(text=f'{p.name}  ›  {kind}  ·  {proj.file_title(os.path.join(p.folder, fn), 40)}')
            editor.delete('1.0', tk.END)
            editor.insert('1.0', proj._read_text(os.path.join(p.folder, fn)))
            editor.edit_reset(); editor.edit_modified(False)
            st['dirty'] = False
            editor.pack(fill=tk.BOTH, expand=True)
            count()
            _btn(rbtn, '저장', lambda: (save_cur(), refresh(st['cur'])), bg=C['accent'], fs=10,
                 px=22, py=6).pack(side=tk.RIGHT, padx=8, pady=6)
            self._small_btn(rbtn, '복사해 둔 내용으로 전체 바꾸기', replace_clip).pack(side=tk.RIGHT, pady=6)
            tk.Label(rbtn, text='다른 항목으로 넘어가거나 창을 닫으면 자동 저장', font=(F, 8),
                     bg=C['bg'], fg='#94a3b8').pack(side=tk.LEFT, padx=12)

        def on_mod(_=None):
            if editor.edit_modified():
                st['dirty'] = True
                editor.edit_modified(False)
                count()
        editor.bind('<<Modified>>', on_mod)
        editor.bind('<Control-s>', lambda e: (save_cur(), 'break')[1])

        def replace_clip():
            try:
                txt = w.clipboard_get()
            except Exception:
                txt = ''
            if not txt.strip():
                messagebox.showwarning('지침', '먼저 새 지침 내용을 복사(Ctrl+C)하세요.', parent=w); return
            if not messagebox.askyesno('지침', f'복사해 둔 내용({len(txt):,}자)으로 전체를 바꿀까요?', parent=w):
                return
            editor.delete('1.0', tk.END); editor.insert('1.0', txt)
            st['dirty'] = True
            cur = st['cur']
            save_cur()
            refresh(cur)
            offer_rules(get_p(cur[1]))

        # 검사 규칙
        def show_rules(p):
            title.config(text=f'{p.name}  ›  자동 검사 규칙'); info.config(text='')
            panel.pack(fill=tk.BOTH, expand=True, padx=16, pady=10)
            tk.Label(panel, text='원고를 다 쓴 뒤 매크로가 이 지침의 규칙을 기계적으로 점검합니다. '
                                 '걸린 원고는 그 편만 자동으로 다시 씁니다.\n'
                                 '지침마다 따로 저장되고, 비워 두면 검사하지 않습니다.',
                     font=(F, 8), bg=C['white'], fg=C['subtext'], justify='left').pack(anchor='w')
            ar = tk.Frame(panel, bg=C['white']); ar.pack(fill=tk.X, pady=(6, 4))
            self._small_btn(ar, '✨ 지침에서 자동으로 만들기', lambda: auto_rules(p)).pack(side=tk.LEFT)
            rv['status'] = tk.Label(ar, text='', font=(F, 8), bg=C['white'], fg=C['subtext'])
            rv['status'].pack(side=tk.LEFT, padx=8)
            build_rule_form(p.settings.get('검사', {}))
            _btn(rbtn, '저장', lambda: (save_rules(p), refresh(('rules', p.name, None))),
                 bg=C['accent'], fs=10, px=22, py=6).pack(side=tk.RIGHT, padx=8, pady=6)

        def build_rule_form(r):
            if rv.get('form') and rv['form'].winfo_exists():
                rv['form'].destroy()
            f = tk.Frame(panel, bg=C['white']); f.pack(fill=tk.BOTH, expand=True)
            rv['form'] = f
            def lab(t):
                tk.Label(f, text=t, font=(F, 9, 'bold'), bg=C['white'], fg=C['text']).pack(anchor='w', pady=(8, 0))
            lab('원고마다 정확히 1번 나와야 하는 문구 (한 줄에 하나)')
            rv['once'] = tk.Text(f, font=(F, 9), height=3, relief=tk.FLAT, bg=C['bg'], padx=8, pady=4)
            rv['once'].insert('1.0', '\n'.join(r.get('원고당_정확히1회', [])))
            rv['once'].pack(fill=tk.X, pady=(4, 0))
            lab('원고에 나오면 안 되는 문구 (한 줄에 하나)')
            rv['ban'] = tk.Text(f, font=(F, 9), height=6, relief=tk.FLAT, bg=C['bg'], padx=8, pady=4)
            rv['ban'].insert('1.0', '\n'.join(r.get('금지문구', [])))
            rv['ban'].pack(fill=tk.X, pady=(4, 0))
            row = tk.Frame(f, bg=C['white']); row.pack(fill=tk.X, pady=(12, 0))
            for key, text, unit, k2 in [('title', '제목 최대', '자 (끝 결과 괄호 제외)', '제목_최대글자'),
                                        ('body', '    본문 최소', '자 (0 = 검사 안 함)', '본문_최소글자')]:
                tk.Label(row, text=text, font=(F, 9, 'bold'), bg=C['white']).pack(side=tk.LEFT)
                e = tk.Entry(row, width=6, font=(F, 9), relief=tk.FLAT, bg=C['bg'], justify='center')
                e.insert(0, str(r.get(k2, 0) or 0)); e.pack(side=tk.LEFT, padx=6, ipady=3)
                tk.Label(row, text=unit, font=(F, 8), bg=C['white'], fg='#94a3b8').pack(side=tk.LEFT)
                rv[key] = e
            rv['case'] = tk.BooleanVar(value=bool(r.get('사건번호_노출금지', True)))
            tk.Checkbutton(f, text='엑셀 사건번호가 원고에 나오면 걸리게', variable=rv['case'], font=(F, 9),
                           bg=C['white'], activebackground=C['white'], selectcolor=C['white']
                           ).pack(anchor='w', pady=(10, 0))

        def save_rules(p):
            if not (rv.get('once') and rv['once'].winfo_exists()):
                return
            lines = lambda t: [x.strip() for x in t.get('1.0', tk.END).split('\n') if x.strip()]
            num = lambda e: max(0, int(e.get())) if e.get().strip().isdigit() else 0
            rules = dict(p.settings.get('검사', {}))
            rules.update({'원고당_정확히1회': lines(rv['once']), '금지문구': lines(rv['ban']),
                          '제목_최대글자': num(rv['title']), '본문_최소글자': num(rv['body']),
                          '사건번호_노출금지': bool(rv['case'].get())})
            proj.save_settings(p, {'검사': rules})

        def auto_rules(p, then_open=False):
            acc = self.account_var.get()
            acc = 1 if acc == 'both' else int(acc)
            if rv.get('status') and rv['status'].winfo_exists():
                rv['status'].config(text='지침 읽는 중… (1분 정도)', fg=C['subtext'])
            self._ai_log_msg(f'{p.name}: 지침에서 검사 규칙 만드는 중…', 'inf')
            def work():
                try:
                    r = generator.extract_rules(p, account=acc)
                except Exception as e:
                    msg = str(e)
                    self._ui(lambda: (rv.get('status') and rv['status'].winfo_exists() and
                                      rv['status'].config(text=f'❌ {msg[:60]}', fg=C['red']),
                                      self._ai_log_msg(f'검사 규칙 만들기 실패: {msg}', 'err')))
                    return
                def done():
                    rules = dict(p.settings.get('검사', {})); rules.update(r)
                    proj.save_settings(p, {'검사': rules})
                    self._ai_log_msg(f'{p.name}: 검사 규칙을 지침에 맞게 새로 만들었습니다', 'ok')
                    if w.winfo_exists():
                        st['cur'] = None
                        refresh(('rules', p.name, None))
                        if rv.get('status') and rv['status'].winfo_exists():
                            rv['status'].config(text='✅ 지침에 맞게 채웠습니다. 확인 후 필요하면 고치세요', fg=C['green'])
                self._ui(done)
            threading.Thread(target=work, daemon=True).start()

        def offer_rules(p):
            if p and messagebox.askyesno('검사 규칙', '지침이 바뀌었습니다.\n자동 검사 규칙도 새 지침에 맞게 자동으로 만들까요?\n(1분 정도, 사용량 조금)',
                                         parent=w):
                auto_rules(p)

        def on_select(_=None):
            sel = lb.curselection()
            if sel and st['rows'][sel[0]] != st['cur']:
                open_item(st['rows'][sel[0]])
        lb.bind('<<ListboxSelect>>', on_select)

        # 버튼
        def cur_project():
            return get_p(st['cur'][1]) if st['cur'] else None

        def new_guide():
            from tkinter import simpledialog
            name = simpledialog.askstring('새 지침', '지침 이름\n예) 오엔_블로그칼럼', parent=w)
            if not name:
                return
            try:
                p = proj.create_project(name)
            except Exception as e:
                messagebox.showerror('새 지침', str(e), parent=w); return
            self._refresh_projects()
            st['cur'] = None
            refresh(('file', p.name, proj.MAIN_GUIDE))
            messagebox.showinfo('새 지침', '새 지침을 만들었습니다.\n지침 내용을 복사(Ctrl+C)한 뒤 [복사해 둔 내용으로 전체 바꾸기]를 누르세요.',
                                parent=w)

        def new_file():
            p = cur_project()
            if not p:
                return
            from tkinter import simpledialog
            name = simpledialog.askstring('새 파일', '파일 이름 (예: 세부지침)', parent=w)
            if not name:
                return
            name = re.sub(r'[\\/:*?"<>|]', '_', name.strip())
            if not name.lower().endswith(('.md', '.txt')):
                name += '.md'
            path = os.path.join(p.folder, name)
            if os.path.exists(path):
                messagebox.showwarning('새 파일', '같은 이름의 파일이 있습니다.', parent=w); return
            save_cur()
            with open(path, 'w', encoding='utf-8') as f:
                f.write(os.path.splitext(name)[0] + '\n')
            st['cur'] = None
            refresh(('file', p.name, name))

        def import_files():
            p = cur_project()
            if not p:
                return
            paths = filedialog.askopenfilenames(parent=w, title='지침·참고 파일 가져오기',
                                                filetypes=[('텍스트', '*.md *.txt'), ('전체', '*.*')])
            last = None
            for src in paths:
                name = os.path.basename(src)
                if not name.lower().endswith(('.md', '.txt')):
                    name = os.path.splitext(name)[0] + '.md'
                dst = os.path.join(p.folder, name)
                if os.path.exists(dst) and not messagebox.askyesno('가져오기', f'{name} 이(가) 이미 있습니다. 덮어쓸까요?',
                                                                    parent=w):
                    continue
                with open(dst, 'w', encoding='utf-8') as f:
                    f.write(proj._read_text(src))
                last = name
            if last:
                save_cur(); st['cur'] = None
                refresh(('file', p.name, last))
                offer_rules(p)

        def delete_item():
            cur = st['cur']
            if not cur:
                return
            kind, pname, fn = cur
            p = get_p(pname)
            if kind == 'project':
                if self.generating:
                    messagebox.showwarning('삭제', '원고 생성 중에는 지울 수 없습니다.', parent=w); return
                if not messagebox.askyesno('지침 삭제', f'지침 "{pname}" 전체(파일·설정)를 삭제할까요?\n되돌릴 수 없습니다.',
                                           parent=w):
                    return
                proj.delete_project(p)
                st['cur'] = None
                self._refresh_projects()
                refresh()
            elif kind == 'file':
                if fn == proj.MAIN_GUIDE:
                    messagebox.showinfo('삭제', '메인 지침은 지울 수 없습니다.\n지침 전체를 지우려면 지침 이름(■)을 고르고 삭제하세요.',
                                        parent=w); return
                if not messagebox.askyesno('삭제', f'참고 파일 "{proj.file_title(os.path.join(p.folder, fn), 40)}"을(를) 삭제할까요?',
                                           parent=w):
                    return
                os.remove(os.path.join(p.folder, fn))
                st['cur'] = None; st['dirty'] = False
                refresh(('project', pname, None))

        for t, cmd in [('+ 새 지침', new_guide), ('+ 파일', new_file), ('가져오기', import_files)]:
            self._small_btn(lbtn, t, cmd).pack(side=tk.LEFT, padx=(0, 4))
        self._small_btn(lbtn2, '삭제', delete_item).pack(side=tk.LEFT)
        tk.Label(lbtn2, text='고른 항목(지침 또는 파일)을 삭제', font=(F, 8), bg=C['white'],
                 fg='#94a3b8').pack(side=tk.LEFT, padx=6)

        def close():
            save_cur(); w.destroy()
            self._refresh_projects()
        w.protocol('WM_DELETE_WINDOW', close)
        cp = self._current_project()
        refresh(('project', cp.name, None) if cp else None)

    # ── Claude Code 설치 / 로그인 ──

    def _install_claude(self):
        try:
            generator.open_installer()
        except Exception as e:
            messagebox.showerror('설치', f'설치 창을 열지 못했습니다: {e}\n\n'
                                 'PowerShell에서 직접 실행하세요:\nirm https://claude.ai/install.ps1 | iex')

    def _offer_install(self):
        if messagebox.askyesno('Claude Code 없음',
                               '이 PC에서 Claude Code를 찾지 못했습니다.\n\n'
                               '지금 설치할까요? (1~2분)\n'
                               '설치가 끝나면 브라우저에서 Claude 계정으로 로그인하면 됩니다.'):
            self._install_claude()

    # ── 설정 ──

    def _open_settings(self):
        win = tk.Toplevel(self.root)
        win.title('설정')
        win.configure(bg=C['bg'])
        win.transient(self.root)
        win.resizable(False, False)
        entries = {}

        card = self._card(win, 'Claude Code')
        body = tk.Frame(card, bg=C['white'])
        body.pack(fill=tk.X, padx=12, pady=(4, 10))

        def field(label, key, width=40):
            row = tk.Frame(body, bg=C['white'])
            row.pack(fill=tk.X, pady=3)
            tk.Label(row, text=label, font=(F, 9), bg=C['white'], fg=C['subtext'],
                     width=12, anchor='w').pack(side=tk.LEFT)
            e = tk.Entry(row, font=(F, 9), width=width, relief=tk.FLAT, bg=C['bg'],
                         insertbackground=C['text'])
            e.insert(0, str(CONFIG.get(key, '')))
            e.pack(side=tk.LEFT, ipady=4, fill=tk.X, expand=True)
            entries[key] = e

        field('모델', 'claude_model')
        tk.Label(body, text='비우면 기본값 · 예) sonnet / opus', font=(F, 8), bg=C['white'],
                 fg='#94a3b8').pack(anchor='w', padx=(88, 0))
        field('실행 파일 경로', 'claude_cli_path')
        tk.Label(body, text='비우면 자동으로 찾음', font=(F, 8), bg=C['white'],
                 fg='#94a3b8').pack(anchor='w', padx=(88, 0))
        field('최대 대기(분)', 'claude_timeout_min', width=6)

        irow = tk.Frame(body, bg=C['white'])
        irow.pack(fill=tk.X, pady=(8, 0))
        tk.Label(irow, text='Claude Code가 없는 PC', font=(F, 8), bg=C['white'],
                 fg='#94a3b8', width=16, anchor='w').pack(side=tk.LEFT)
        self._small_btn(irow, 'Claude Code 설치하기', self._install_claude).pack(side=tk.LEFT)
        tk.Label(body, text='계정 로그인·확인은 AI 원고 생성 창에서 합니다', font=(F, 8),
                 bg=C['white'], fg='#94a3b8').pack(anchor='w', pady=(6, 0))

        def apply_entries():
            CONFIG['claude_model'] = entries['claude_model'].get().strip()
            CONFIG['claude_cli_path'] = entries['claude_cli_path'].get().strip()

        ex = self._card(win, '엑셀')
        used_var = tk.BooleanVar(value=CONFIG.get('excel_used_check', True))
        tk.Checkbutton(ex, text="'사용체크' 열에 표시된 사례도 사용한 것으로 보기",
                       variable=used_var, font=(F, 9), bg=C['white'], fg=C['subtext'],
                       activebackground=C['white'], selectcolor=C['white']
                       ).pack(anchor='w', padx=12, pady=8)

        def save():
            apply_entries()
            try:
                CONFIG['claude_timeout_min'] = max(5, int(entries['claude_timeout_min'].get()))
            except ValueError:
                pass
            CONFIG['excel_used_check'] = used_var.get()
            app_config.save_config()
            self._fill_cases()
            win.destroy()

        _btn(win, '저장', save, bg=C['accent'], fs=10, px=28, py=7).pack(pady=12)


def run(root=None, input_file=''):
    if root is None:
        root = tk.Tk()
    app = OneNoteExporterGUI(root)
    if input_file:
        root.after(200, lambda: app._open_file_path(input_file))
    root.mainloop()


if __name__ == '__main__':
    run()
