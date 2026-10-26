# ui.py
import tkinter as tk
from tkinter import scrolledtext, messagebox, filedialog, ttk
import re
import threading
from parser import ArticleParser
from onenote_com import OneNoteManager, SectionInfo

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
        self.root.geometry('860x830')
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

        tk.Label(cfg, text='영상\n링크', font=(F, 9, 'bold'),
                 bg=C['white'], fg=C['subtext']).grid(row=3, column=0, padx=(12,4), pady=(0,8), sticky='nw')
        self.links_txt = tk.Text(cfg, font=(F, 9), width=1, height=3, wrap=tk.NONE,
                                 relief=tk.FLAT, bg=C['bg'], insertbackground=C['text'])
        self.links_txt.grid(row=3, column=1, columnspan=5, padx=(0,3), pady=(0,8), sticky='ew')
        tk.Label(cfg, text='한 줄에 하나씩\n본문 맨 앞에\n따로 삽입', font=(F, 8),
                 bg=C['white'], fg='#94a3b8', justify=tk.LEFT
                 ).grid(row=3, column=6, columnspan=3, padx=(6,6), pady=(0,8), sticky='nw')


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
        if not path:
            return
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
        links, bad = self._get_links()
        if bad:
            messagebox.showwarning('경고', '링크가 아닌 줄이 있습니다 (http로 시작해야 함):\n' + '\n'.join(bad))
            return
        msg = f"💾 {len(selected)}개 글을 내보낼까요?\n섹션: {self.section_var.get()}"
        if prefix:
            msg += f"\n접두어: {prefix}"
        if links:
            msg += f"\n영상 링크(맨 앞 삽입): {len(links)}개"
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
            args=(selected, prefix, selected_footers, keyword, links),
            daemon=True).start()

    def _get_links(self):
        """영상 링크 칸에서 링크 목록을 읽는다. 한 줄에 여러 개를 붙여넣어도 공백 기준으로 나눈다.
        반환: (링크 목록, 링크가 아닌 항목 목록)"""
        links, bad = [], []
        for tok in self.links_txt.get('1.0', tk.END).split():
            if re.match(r'https?://', tok, re.I):
                links.append(tok)
            else:
                bad.append(tok)
        return links, bad


    def _export_thread(self, selected, prefix, footer_presets, keyword='', links=None):
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
            top_keyword=keyword, top_links=links)
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


def run(root=None, input_file=''):
    if root is None:
        root = tk.Tk()
    app = OneNoteExporterGUI(root)
    if input_file:
        root.after(200, lambda: app._open_file_path(input_file))
    root.mainloop()


if __name__ == '__main__':
    run()
