# ui.py
import tkinter as tk
from tkinter import scrolledtext, messagebox, filedialog, ttk
import os
import threading
from datetime import datetime
from parser import ArticleParser
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
        tk.Label(cfg, text='입력 시 본문 맨 위에 굵게 삽입됩니다', font=(F, 8),
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
                        ('🗑 삭제',  self._delete)]:
            tk.Button(lst_bot, text=t, command=cmd,
                      font=(F, 9), bg=C['light'], fg=C['text'],
                      relief=tk.FLAT, padx=12, pady=5, cursor='hand2'
                      ).pack(side=tk.LEFT, padx=(8,0), pady=6)

 
        footer_frame = tk.Frame(self.root, bg=C['white'],
                                highlightbackground=C['light'], highlightthickness=1)
        footer_frame.pack(fill=tk.X, padx=12, pady=(4, 0))

        tk.Label(footer_frame, text='',
                 font=(F, 9, 'bold'), bg=C['white'], fg=C['subtext']
                 ).pack(side=tk.LEFT, padx=(12, 8), pady=6)

        self.footer_preset_vars = []
        for name, _ in FOOTER_PRESETS:
            var = tk.BooleanVar(value=False)
            self.footer_preset_vars.append(var)
            tk.Checkbutton(footer_frame, text=name, variable=var,
                           font=(F, 9), bg=C['white'], fg=C['text'],
                           activebackground=C['white'], selectcolor=C['white']
                           ).pack(side=tk.LEFT, padx=(0, 4), pady=6)



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
            t = a['title'][:60] + ('…' if len(a['title']) > 60 else '')
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
        result = self.onenote.batch_create_pages(
            selected, title_prefix=prefix, progress_cb=cb,
            add_footer=add_footer, footer_presets=footer_presets,
            top_keyword=keyword)
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

    def _card(self, parent, title=None):
        f = tk.Frame(parent, bg=C['white'], highlightbackground=C['light'], highlightthickness=1)
        f.pack(fill=tk.X, padx=12, pady=(8, 0))
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

    def _open_ai_window(self):
        if self.ai_win and self.ai_win.winfo_exists():
            self.ai_win.deiconify(); self.ai_win.lift(); return
        win = tk.Toplevel(self.root)
        self.ai_win = win
        win.title('AI 원고 생성')
        win.geometry('820x700')
        win.minsize(640, 520)
        win.configure(bg=C['bg'])
        win.protocol('WM_DELETE_WINDOW', win.withdraw)

        # 프로젝트
        pf = self._card(win)
        tk.Label(pf, text='프로젝트', font=(F, 9, 'bold'), bg=C['white'],
                 fg=C['subtext']).grid(row=0, column=0, padx=(12, 4), pady=8, sticky='w')
        self.project_var = tk.StringVar()
        self.project_cb = ttk.Combobox(pf, textvariable=self.project_var, state='readonly',
                                       width=24, font=(F, 9))
        self.project_cb.grid(row=0, column=1, padx=(0, 6), pady=8)
        self.project_cb.bind('<<ComboboxSelected>>', lambda e: self._on_project())
        self._small_btn(pf, '지침 폴더', self._open_project_folder).grid(row=0, column=2, padx=2)
        self._small_btn(pf, '새 프로젝트', self._new_project).grid(row=0, column=3, padx=2)
        self._small_btn(pf, '새로고침', self._refresh_projects).grid(row=0, column=4, padx=2)
        self.project_info = tk.Label(pf, text='', font=(F, 8), bg=C['white'], fg='#94a3b8')
        self.project_info.grid(row=0, column=5, padx=8, sticky='w')

        tk.Frame(pf, bg=C['light'], height=1).grid(row=1, column=0, columnspan=7,
                                                    sticky='ew', padx=8)
        tk.Label(pf, text='입력', font=(F, 9, 'bold'), bg=C['white'],
                 fg=C['subtext']).grid(row=2, column=0, padx=(12, 4), pady=8, sticky='w')
        self.source_var = tk.StringVar(value='excel')
        sf = tk.Frame(pf, bg=C['white'])
        sf.grid(row=2, column=1, columnspan=3, sticky='w')
        for val, text in [('excel', '엑셀 사례'), ('direct', '직접 입력')]:
            tk.Radiobutton(sf, text=text, variable=self.source_var, value=val,
                           command=self._on_source, font=(F, 9), bg=C['white'],
                           activebackground=C['white'], selectcolor=C['white']
                           ).pack(side=tk.LEFT, padx=(0, 10))
        tk.Label(pf, text='계정', font=(F, 9, 'bold'), bg=C['white'],
                 fg=C['subtext']).grid(row=2, column=4, padx=(4, 4), sticky='e')
        self.account_var = tk.IntVar(value=int(CONFIG.get('account') or 1))
        af = tk.Frame(pf, bg=C['white'])
        af.grid(row=2, column=5, sticky='w')
        for n in (1, 2):
            tk.Radiobutton(af, text=f'계정 {n}', variable=self.account_var, value=n,
                           font=(F, 9), bg=C['white'], activebackground=C['white'],
                           selectcolor=C['white']).pack(side=tk.LEFT, padx=(0, 6))
        pf.columnconfigure(6, weight=1)

        # 엑셀 / 직접 입력 영역
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
        tk.Button(top, text='엑셀 파일 선택', command=self._pick_excel, font=(F, 9),
                  bg=C['light'], fg=C['text'], relief=tk.FLAT, padx=10, pady=2,
                  cursor='hand2').pack(side=tk.RIGHT, padx=8)

        r2 = tk.Frame(ef, bg=C['white'])
        r2.pack(fill=tk.X, padx=12, pady=(6, 4))
        tk.Label(r2, text='시트', font=(F, 9), bg=C['white'], fg=C['subtext']).pack(side=tk.LEFT)
        self.sheet_var = tk.StringVar()
        self.sheet_cb = ttk.Combobox(r2, textvariable=self.sheet_var, state='readonly',
                                     width=13, font=(F, 9))
        self.sheet_cb.pack(side=tk.LEFT, padx=(4, 10))
        self.sheet_cb.bind('<<ComboboxSelected>>', lambda e: self._on_sheet())
        tk.Label(r2, text='분류', font=(F, 9), bg=C['white'], fg=C['subtext']).pack(side=tk.LEFT)
        self.cat_var = tk.StringVar(value='(전체)')
        self.cat_cb = ttk.Combobox(r2, textvariable=self.cat_var, state='readonly',
                                   width=15, font=(F, 9))
        self.cat_cb.pack(side=tk.LEFT, padx=(4, 10))
        self.cat_cb.bind('<<ComboboxSelected>>', lambda e: self._fill_cases())
        self.unused_var = tk.BooleanVar(value=True)
        tk.Checkbutton(r2, text='안 쓴 사례만', variable=self.unused_var,
                       command=self._fill_cases, font=(F, 9), bg=C['white'],
                       fg=C['subtext'], activebackground=C['white'], selectcolor=C['white']
                       ).pack(side=tk.LEFT)
        self.case_count = tk.Label(r2, text='', font=(F, 8), bg=C['white'], fg='#94a3b8')
        self.case_count.pack(side=tk.RIGHT)

        tw = tk.Frame(ef, bg=C['white'])
        tw.pack(fill=tk.BOTH, expand=True, padx=12, pady=(0, 4))
        cols = ('crime', 'r1', 'r2', 'summary')
        self.case_tree = ttk.Treeview(tw, columns=cols, show='headings', height=8,
                                      selectmode='extended')
        for c, t, w in [('crime', '범죄명', 120), ('r1', '결과', 100),
                        ('r2', '결과2', 140), ('summary', '요약', 360)]:
            self.case_tree.heading(c, text=t)
            self.case_tree.column(c, width=w, anchor='w')
        vsb = ttk.Scrollbar(tw, orient='vertical', command=self.case_tree.yview)
        self.case_tree.configure(yscrollcommand=vsb.set)
        self.case_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        self.case_tree.tag_configure('used', foreground='#94a3b8')

        bot = tk.Frame(ef, bg=C['bg'])
        bot.pack(fill=tk.X, padx=1, pady=1)
        tk.Label(bot, text='위에서', font=(F, 9), bg=C['bg'], fg=C['subtext']
                 ).pack(side=tk.LEFT, padx=(12, 4), pady=6)
        self.pick_n = tk.Spinbox(bot, from_=1, to=50, width=3, font=(F, 9))
        self.pick_n.delete(0, tk.END); self.pick_n.insert(0, '1')
        self.pick_n.pack(side=tk.LEFT)
        tk.Button(bot, text='개 선택', command=self._auto_pick, font=(F, 9), bg=C['light'],
                  fg=C['text'], relief=tk.FLAT, padx=10, pady=3, cursor='hand2'
                  ).pack(side=tk.LEFT, padx=6)

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

        # 실행
        run = tk.Frame(win, bg=C['bg'])
        run.pack(fill=tk.X, padx=12, pady=(8, 0))
        self.auto_export_var = tk.BooleanVar(value=CONFIG.get('auto_export', True))
        tk.Checkbutton(run, text='생성 후 바로 OneNote로 내보내기', variable=self.auto_export_var,
                       font=(F, 9), bg=C['bg'], fg=C['subtext'], activebackground=C['bg'],
                       selectcolor=C['white']).pack(side=tk.LEFT)
        self.stop_btn = _btn(run, '중지', self._stop_generate, bg=C['gray'], fs=10, px=16, py=7)
        self.stop_btn.pack(side=tk.RIGHT, padx=(6, 0))
        self.stop_btn.config(state=tk.DISABLED)
        self.gen_btn = _btn(run, '원고 생성', self._start_generate, bg=C['accent'], fs=10,
                            px=24, py=7)
        self.gen_btn.pack(side=tk.RIGHT)
        tk.Button(run, text='결과 폴더', command=lambda: app_config.open_folder(app_config.OUTPUT_DIR),
                  font=(F, 9), bg=C['light'], fg=C['text'], relief=tk.FLAT, padx=12, pady=5,
                  cursor='hand2').pack(side=tk.RIGHT, padx=6)

        self.ai_log = scrolledtext.ScrolledText(
            win, font=('Consolas', 8), bg=C['log_bg'], fg=C['log_fg'],
            relief=tk.FLAT, padx=8, pady=4, state=tk.DISABLED, height=6)
        self.ai_log.pack(fill=tk.X, padx=12, pady=(8, 10))
        for tag, col in [('ok', C['log_ok']), ('err', C['log_err']),
                         ('inf', C['log_inf']), ('warn', '#fbbf24')]:
            self.ai_log.tag_configure(tag, foreground=col)

        self._refresh_projects()
        if CONFIG.get('excel_path') and os.path.exists(CONFIG['excel_path']):
            win.after(200, lambda: self._load_excel(CONFIG['excel_path']))
        else:
            self.excel_lbl.config(text='엑셀 파일을 선택하세요')
        self._on_source()

    def _ai_log_msg(self, msg, level='inf'):
        self._log(msg, level)
        if not (self.ai_win and self.ai_win.winfo_exists()):
            return
        ts = datetime.now().strftime('%H:%M:%S')
        self.ai_log.config(state=tk.NORMAL)
        self.ai_log.insert(tk.END, f'[{ts}] {msg}\n', level)
        self.ai_log.see(tk.END)
        self.ai_log.config(state=tk.DISABLED)

    # ── 프로젝트 ──

    def _current_project(self):
        name = self.project_var.get()
        return next((p for p in self.projects if p.name == name), None)

    def _refresh_projects(self):
        self.projects = proj.list_projects()
        names = [p.name for p in self.projects]
        self.project_cb['values'] = names
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

    def _on_project(self):
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
        app_config.open_folder(p.folder)
        messagebox.showinfo('새 프로젝트',
                            "폴더가 열렸습니다.\n\n"
                            "1) 지침.md 에 claude.ai 프로젝트 지침을 붙여넣기\n"
                            "2) 프로젝트 소스 파일(.md/.txt)은 같은 폴더에 넣기\n"
                            "3) 설정.json 에서 편수·하단 프리셋 등 수정\n\n"
                            "다 하면 [새로고침]을 누르세요.", parent=self.ai_win)

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
        self.sheet_cb['values'] = list(self.cases.keys())
        if self.sheet_var.get() not in self.cases:
            self.sheet_var.set(next(iter(self.cases)))
        self._on_sheet()

    def _on_sheet(self):
        cases = self.cases.get(self.sheet_var.get(), [])
        cats = sorted({c.category for c in cases if c.category})
        self.cat_cb['values'] = ['(전체)'] + cats
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
        self.case_tree.delete(*self.case_tree.get_children())
        self.case_rows = {}
        p = self._current_project()
        used = excel_source.load_usage(p.name) if p else set()
        cat = self.cat_var.get()
        shown = n_used = 0
        for c in self.cases.get(self.sheet_var.get(), []):
            if cat != '(전체)' and c.category != cat:
                continue
            is_used = self._is_used(c, used)
            n_used += is_used
            if is_used and self.unused_var.get():
                continue
            iid = self.case_tree.insert('', tk.END, values=(
                c.crime.replace('\n', ', '), c.result1.replace('\n', ', '),
                c.result2.replace('\n', ', '), c.short()),
                tags=('used',) if is_used else ())
            self.case_rows[iid] = c
            shown += 1
        self.case_count.config(text=f'{shown}건 · 사용함 {n_used}건')

    def _auto_pick(self):
        try:
            n = int(self.pick_n.get())
        except ValueError:
            n = 1
        items = self.case_tree.get_children()[:n]
        self.case_tree.selection_set(items)
        if items:
            self.case_tree.see(items[0])

    # ── 생성 ──

    def _start_generate(self):
        if self.generating:
            return
        w = self.ai_win
        p = self._current_project()
        if not p:
            messagebox.showwarning('경고', '프로젝트를 먼저 고르세요.', parent=w); return
        try:
            p.build_system_prompt()
        except Exception as e:
            messagebox.showerror('지침 오류', str(e), parent=w); return
        if not generator.find_claude_cli():
            self._offer_install()
            return

        jobs = []
        if self.source_var.get() == 'excel':
            sel = self.case_tree.selection()
            if not sel:
                messagebox.showwarning('경고', '목록에서 사례를 선택하세요.\n(Ctrl/Shift로 여러 개, 또는 [N개 선택])',
                                       parent=w)
                return
            for iid in sel:
                c = self.case_rows[iid]
                jobs.append((c.crime.split('\n')[0], c.to_prompt(), c))
        else:
            text = self.direct_txt.get('1.0', tk.END).strip()
            if len(text) < 5:
                messagebox.showwarning('경고', '사건자료(또는 주제)를 입력하세요.', parent=w); return
            jobs.append(('직접입력', text, None))

        auto_export = self.auto_export_var.get()
        CONFIG['auto_export'] = auto_export
        CONFIG['account'] = self.account_var.get()
        app_config.save_config()
        if auto_export and not self.is_connected:
            if not messagebox.askyesno('확인', 'OneNote가 연결되지 않았습니다.\n원고만 만들고 파일로 저장할까요?',
                                       parent=w):
                return
            auto_export = False

        n = int(p.settings.get('편수', 10) or 10)
        msg = (f'{len(jobs)}건 × {n}편 = 원고 {len(jobs) * n}편을 생성합니다.\n'
               f'계정 {CONFIG["account"]} 사용 · 건당 수 분 걸립니다.')
        if auto_export:
            msg += f'\n\n완료되는 대로 OneNote "{self.section_var.get()}"에 내보냅니다.'
        if not messagebox.askyesno('원고 생성', msg, parent=w):
            return

        opts = {
            'prefix': self.prefix_var.get().strip(),
            'keyword': self.keyword_var.get().strip(),
            'add_footer': self.footer_var.get(),
            'footers': [lines for var, (_, lines) in zip(self.footer_preset_vars, FOOTER_PRESETS)
                        if var.get()],
        }
        self.generating = True
        self.gen_stop.clear()
        self.gen_btn.config(state=tk.DISABLED, text='생성 중…')
        self.stop_btn.config(state=tk.NORMAL)
        self.export_btn.config(state=tk.DISABLED)
        threading.Thread(target=self._generate_thread,
                         args=(p, jobs, auto_export, opts), daemon=True).start()

    def _stop_generate(self):
        self.gen_stop.set()
        self._ai_log_msg('중지 요청 — 진행 중인 건이 끝나면 멈춥니다', 'warn')

    def _ui(self, fn):
        self.root.after(0, fn)

    def _generate_thread(self, p, jobs, auto_export, opts):
        out_dir = os.path.join(app_config.OUTPUT_DIR,
                               datetime.now().strftime('%Y%m%d') + '_' + p.name)
        os.makedirs(out_dir, exist_ok=True)
        ok = fail = 0
        log = lambda m, l='inf': self._ui(lambda: self._ai_log_msg(m, l))
        for idx, (title, case_text, case) in enumerate(jobs, 1):
            if self.gen_stop.is_set():
                break
            tag = f'[{idx}/{len(jobs)}] {title}'
            log(f'{tag} — 원고 작성 중…')
            self._ui(lambda i=idx: self.prog_status.config(text=f'생성 {i}/{len(jobs)}'))
            try:
                text = generator.generate(p, case_text, stop_flag=self.gen_stop)
            except Exception as e:
                fail += 1
                log(f'{tag} 실패: {e}', 'err')
                if '중지' in str(e) or '한도' in str(e):
                    break
                continue

            stamp = datetime.now().strftime('%H%M%S')
            safe = (case.key if case else '직접입력').replace('/', '_').replace('\\', '_')
            fpath = os.path.join(out_dir, f'{stamp}_{safe}.md')
            with open(fpath, 'w', encoding='utf-8') as f:
                f.write(text)
            for w in proj.check_output(text, p, case.case_id if case else ''):
                log(f'  ⚠ {w}', 'warn')

            arts = ArticleParser.parse(text)
            if not arts:
                fail += 1
                log(f'{tag}: 원고 형식을 읽지 못했습니다 → {os.path.basename(fpath)}', 'err')
                continue
            self._ui(lambda a=arts: self._add_articles(a))

            exported = True
            if auto_export:
                log(f'{tag} — OneNote로 {len(arts)}편 내보내는 중…')
                res = self.onenote.batch_create_pages(
                    arts, title_prefix=opts['prefix'], add_footer=opts['add_footer'],
                    footer_presets=opts['footers'], top_keyword=opts['keyword'])
                exported = res['fail'] == 0
                log(f"{tag} — OneNote {res['success']}편 성공 / {res['fail']}편 실패",
                    'ok' if exported else 'err')
            if case and exported:
                excel_source.record_usage(p.name, case)
            ok += 1
            log(f'{tag} 완료', 'ok')

        self._ui(lambda: self._generate_done(ok, fail))

    def _add_articles(self, arts):
        self.articles.extend(arts)
        self._rebuild_checks()
        self.count_lbl.config(text=f'{len(self.articles)}개')

    def _generate_done(self, ok, fail):
        self.generating = False
        if self.ai_win and self.ai_win.winfo_exists():
            self.gen_btn.config(state=tk.NORMAL, text='원고 생성')
            self.stop_btn.config(state=tk.DISABLED)
        if self.is_connected:
            self.export_btn.config(state=tk.NORMAL)
        self.prog_status.config(text='')
        self._ai_log_msg(f'종료: 성공 {ok}건 / 실패 {fail}건', 'ok' if not fail else 'warn')
        self._set_status(f'AI 생성 완료: {ok}건', C['green'] if not fail else C['orange'])
        self._fill_cases()

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

        acc = self._card(win, '계정')
        ab = tk.Frame(acc, bg=C['white'])
        ab.pack(fill=tk.X, padx=12, pady=(4, 10))
        status = {}

        def apply_entries():
            CONFIG['claude_model'] = entries['claude_model'].get().strip()
            CONFIG['claude_cli_path'] = entries['claude_cli_path'].get().strip()

        def test(n):
            apply_entries()
            status[n].config(text='확인 중…', fg=C['subtext'])

            def work():
                try:
                    r = generator.call_claude_cli("'연결 성공'이라고만 답해주세요.", timeout=120,
                                                  account=n)
                    self._ui(lambda: status[n].winfo_exists() and
                             status[n].config(text=f'✅ {r.strip()[:20]}', fg=C['green']))
                except Exception as e:
                    m = str(e)
                    self._ui(lambda: status[n].winfo_exists() and
                             status[n].config(text=f'❌ {m}', fg=C['red']))
            threading.Thread(target=work, daemon=True).start()

        def login(n):
            apply_entries()
            try:
                generator.open_login(n)
            except Exception as e:
                messagebox.showerror('로그인', str(e), parent=win)

        for n in (1, 2):
            row = tk.Frame(ab, bg=C['white'])
            row.pack(fill=tk.X, pady=3)
            tk.Label(row, text=f'계정 {n}', font=(F, 9, 'bold'), bg=C['white'], fg=C['text'],
                     width=8, anchor='w').pack(side=tk.LEFT)
            self._small_btn(row, '로그인', lambda n=n: login(n)).pack(side=tk.LEFT, padx=(0, 4))
            self._small_btn(row, '연결 테스트', lambda n=n: test(n)).pack(side=tk.LEFT)
            status[n] = tk.Label(row, text='', font=(F, 8), bg=C['white'], fg=C['subtext'],
                                 wraplength=260, justify='left')
            status[n].pack(side=tk.LEFT, padx=8)
        irow = tk.Frame(ab, bg=C['white'])
        irow.pack(fill=tk.X, pady=(6, 0))
        tk.Label(irow, text='Claude Code가 없는 PC', font=(F, 8), bg=C['white'],
                 fg='#94a3b8').pack(side=tk.LEFT)
        self._small_btn(irow, 'Claude Code 설치하기', self._install_claude).pack(side=tk.LEFT, padx=8)

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
