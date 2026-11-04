"""
판매중 상품 점검 (무신사 · 29CM) — 옵션 재고 + 검색 노출 — 실행 화면 (Windows exe 진입점)
"""
import os, sys, json, queue, threading, traceback, webbrowser
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinter.scrolledtext import ScrolledText

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import runner  # noqa: E402
cfg = runner.cfg

APP_TITLE = '판매채널 상품 점검 (무신사 · 29CM)'
PLATFORMS = ('무신사', '29CM')
FONT = ('맑은 고딕', 10)
FONT_B = ('맑은 고딕', 10, 'bold')
LAST_FILE = os.path.join(cfg.ROOT, 'last_paths.json')


def open_path(p):
    try:
        if sys.platform.startswith('win'):
            os.startfile(p)  # noqa
        elif sys.platform == 'darwin':
            os.system(f'open "{p}"')
        else:
            os.system(f'xdg-open "{p}"')
    except Exception as e:
        messagebox.showerror(APP_TITLE, f'열 수 없습니다: {e}')


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry('880x780')
        self.minsize(740, 600)
        self.option_add('*Font', FONT)
        self.q = queue.Queue()
        self.busy = False
        self.stop_ev = threading.Event()
        self.in_var = tk.StringVar()
        self.cm_var = tk.StringVar()
        self.plat = tk.StringVar(value='무신사')
        self.sm_var = tk.StringVar()
        self.out_var = tk.StringVar(value=os.path.join(cfg.ROOT, '결과'))
        self._load_last()
        self._build()
        self.after(100, self._poll)

    # ---------- 화면 ----------
    def _build(self):
        pad = dict(padx=10, pady=4)
        top = ttk.Frame(self); top.pack(fill='x', **pad)
        ttk.Label(top, text=APP_TITLE, font=('맑은 고딕', 14, 'bold')).pack(side='left')
        ttk.Button(top, text='설정(브랜드명)', command=self.open_settings).pack(side='right')
        ttk.Button(top, text='사용법', command=self.show_help).pack(side='right', padx=6)

        pf = ttk.LabelFrame(self, text=' 판매채널 '); pf.pack(fill='x', **pad)
        for i, name in enumerate(PLATFORMS):
            ttk.Radiobutton(pf, text=name, value=name, variable=self.plat, command=self._on_platform).pack(side='left', padx=(10 if i == 0 else 4, 8), pady=6)
        ttk.Label(pf, foreground='#777', text='채널을 고르면 아래 1번 입력 파일과 결과 엑셀이 그 채널용으로 바뀝니다.').pack(side='left', padx=8)

        self.box = box = ttk.LabelFrame(self, text=' 1. 무신사 재고 관리(v2) 복사본 '); box.pack(fill='x', **pad)
        self.in_entry = ttk.Entry(box, textvariable=self.in_var)
        self.in_entry.grid(row=0, column=0, sticky='ew', padx=8, pady=(8, 0))
        ttk.Button(box, text='찾아보기', command=self.pick_in).grid(row=0, column=1, padx=8, pady=(8, 0))
        self.in_help = ttk.Label(box, foreground='#555', justify='left', wraplength=760)
        self.in_help.grid(row=1, column=0, columnspan=2, sticky='w', padx=8, pady=(2, 8))
        box.columnconfigure(0, weight=1)

        sm = ttk.LabelFrame(self, text=' 2. 셀메이트 재고목록 (stk_stockList_*.csv) '); sm.pack(fill='x', **pad)
        ttk.Entry(sm, textvariable=self.sm_var).grid(row=0, column=0, sticky='ew', padx=8, pady=(8, 0))
        ttk.Button(sm, text='찾아보기', command=self.pick_sm).grid(row=0, column=1, padx=8, pady=(8, 0))
        ttk.Label(sm, foreground='#555', justify='left', wraplength=760, text=(
            '셀메이트 → 재고목록 → 엑셀(CSV) 다운로드한 파일 그대로. 셀메이트에 재고가 있는데 무신사 옵션이 품절된 것을 찾는 데 씁니다.'
            )).grid(row=1, column=0, columnspan=2, sticky='w', padx=8, pady=(2, 8))
        sm.columnconfigure(0, weight=1)

        ob = ttk.LabelFrame(self, text=' 3. 결과 저장 폴더 '); ob.pack(fill='x', **pad)
        ttk.Entry(ob, textvariable=self.out_var).pack(side='left', fill='x', expand=True, padx=8, pady=6)
        ttk.Button(ob, text='변경', command=self.pick_out).pack(side='left', padx=4)
        ttk.Button(ob, text='폴더 열기', command=self.open_out).pack(side='left', padx=8)

        sb = ttk.LabelFrame(self, text=' 4. 점검 '); sb.pack(fill='x', **pad)
        self.btn1 = ttk.Button(sb, text='▶ 점검 시작', command=self.start_all)
        self.btn1.grid(row=0, column=0, padx=8, pady=6, sticky='ew', ipady=6)
        self.run_help = ttk.Label(sb, foreground='#444', justify='left', wraplength=520)
        self.run_help.grid(row=0, column=1, columnspan=2, sticky='w')
        self.btn_stop = ttk.Button(sb, text='■ 중지', command=self.stop_search, state='disabled')
        self.btn_stop.grid(row=1, column=0, padx=8, pady=(0, 6), sticky='ew')
        self.btn2 = ttk.Button(sb, text='검색만 이어서 하기', command=self.resume_search)
        self.btn2.grid(row=1, column=1, padx=(0, 8), pady=(0, 6), sticky='w')
        ttk.Label(sb, foreground='#777', text='중지했거나 창을 닫았을 때, 했던 곳 다음부터 검색합니다.').grid(row=1, column=2, sticky='w')
        self.pb = ttk.Progressbar(sb, mode='determinate', length=160)
        self.pb.grid(row=2, column=0, columnspan=3, sticky='ew', padx=8, pady=(0, 6))
        mf = ttk.Frame(sb); mf.grid(row=3, column=0, columnspan=3, sticky='w', padx=8, pady=(0, 6))
        ttk.Label(mf, foreground='#777', text='예비(자동 검색이 안 될 때): ').pack(side='left')
        self.btn3 = ttk.Button(mf, text='콘솔 스크립트 복사 + 무신사 열기', command=self.copy_search); self.btn3.pack(side='left', padx=4)
        self.btn4 = ttk.Button(mf, text='콘솔 결과 붙여넣기', command=self.paste_search); self.btn4.pack(side='left', padx=4)
        sb.columnconfigure(2, weight=1)

        lb = ttk.LabelFrame(self, text=' 진행 상황 '); lb.pack(fill='both', expand=True, **pad)
        self.log_box = ScrolledText(lb, height=10, font=('Consolas', 9), state='disabled')
        self.log_box.pack(fill='both', expand=True, padx=6, pady=6)
        self._on_platform()
        self.log('판매채널을 고르고, 1번 파일과 셀메이트 재고목록을 고른 뒤 [▶ 점검 시작]을 누르세요.')

    def _on_platform(self):
        cm = self.plat.get() == '29CM'
        self.box.configure(text=' 1. 29CM 옵션재고 일괄변경 파일 (여러 개 한꺼번에 선택) ' if cm else ' 1. 무신사 재고 관리(v2) 복사본 ')
        self.in_entry.configure(textvariable=self.cm_var if cm else self.in_var)
        self.in_help.configure(text=(
            '29Connect → 옵션/재고 관리 → 판매상태 "판매중" → 페이지마다 엑셀 다운로드한 '
            '"29CM_옵션재고_일괄변경(파트너)_페이지…" 파일들. [찾아보기]에서 Ctrl·Shift로 여러 개를 한꺼번에 고르세요. '
            '페이지 순서는 상관없고, 겹친 옵션은 한 번만 셉니다.') if cm else (
            '파트너센터 → 상품/재고 → 재고 관리(v2) → 상품상태 "판매중", 출력수 5000 → 표를 전부 드래그해 복사(Ctrl+C) '
            '→ 엑셀 새 파일에 붙여넣기(Ctrl+V) → 저장.  머리글 없이 붙여넣어도 되고, 필요한 열만 자동으로 골라 씁니다. '
            '(5000건이 넘으면 2페이지도 이어서 붙여넣기)'))
        p = self.plat.get()
        self.run_help.configure(text=(
            f'① 셀메이트 재고 ↔ {p} 옵션 대조 → {p}_옵션재고점검 엑셀\n'
            f'② 크롬을 띄워 {p} 검색창에 판매중 상품을 하나씩 검색 → {p}_검색노출점검 엑셀 + 미노출 화면 캡처\n'
            '   (상품 1,000개 기준 40~60분. 검색하는 크롬 창은 건드리지 마세요. 최소화는 괜찮습니다)'))
        if not self.busy:
            for b in (self.btn3, self.btn4):
                b.configure(state='disabled' if cm else 'normal')

    # ---------- 동작 ----------
    def _cm_files(self):
        return [p.strip() for p in self.cm_var.get().split(';') if p.strip()]

    def pick_in(self):
        if self.plat.get() == '29CM':
            ps = filedialog.askopenfilenames(title='29CM 옵션재고 일괄변경 파일 선택 (여러 개)', filetypes=[('Excel', '*.xlsx'), ('모든 파일', '*.*')],
                                             initialdir=os.path.dirname((self._cm_files() or [os.path.join(cfg.ROOT, 'x')])[0]))
            if ps:
                self.cm_var.set('; '.join(os.path.normpath(p) for p in ps))
                self.log(f'29CM 파일 {len(ps)}개를 골랐습니다.')
            return
        p = filedialog.askopenfilename(title='재고 관리(v2) 복사본 선택', filetypes=[('Excel', '*.xlsx'), ('모든 파일', '*.*')],
                                       initialdir=os.path.dirname(self.in_var.get() or cfg.ROOT))
        if p:
            self.in_var.set(os.path.normpath(p))

    def pick_sm(self):
        p = filedialog.askopenfilename(title='셀메이트 재고목록 선택', filetypes=[('CSV', '*.csv'), ('모든 파일', '*.*')],
                                       initialdir=os.path.dirname(self.sm_var.get() or self.in_var.get() or cfg.ROOT))
        if p:
            self.sm_var.set(os.path.normpath(p))

    def pick_out(self):
        p = filedialog.askdirectory(title='결과 저장 폴더', initialdir=self.out_var.get() or cfg.ROOT)
        if p:
            self.out_var.set(os.path.normpath(p))

    def open_out(self):
        d = self.out_var.get()
        os.makedirs(d, exist_ok=True)
        open_path(d)

    def log(self, msg):
        self.log_box.configure(state='normal')
        self.log_box.insert('end', msg + '\n')
        self.log_box.see('end')
        self.log_box.configure(state='disabled')

    def _poll(self):
        try:
            while True:
                kind, payload = self.q.get_nowait()
                if kind == 'log':
                    self.log(payload)
                elif kind == 'prog':
                    i, n = payload
                    self.pb.configure(mode='determinate', maximum=max(n, 1), value=i)
                elif kind == 'done':
                    self._finish(True, payload)
                elif kind == 'error':
                    self._finish(False, payload)
        except queue.Empty:
            pass
        self.after(100, self._poll)

    def _set_busy(self, busy):
        self.busy = busy
        cm = self.plat.get() == '29CM'
        for b in (self.btn1, self.btn2, self.btn3, self.btn4):
            b.configure(state='disabled' if busy or (cm and b in (self.btn3, self.btn4)) else 'normal')
        self.btn_stop.configure(state='normal' if busy else 'disabled')
        if not busy:
            self.pb.stop(); self.pb.configure(mode='determinate', value=0)
        else:
            self.pb.configure(mode='indeterminate'); self.pb.start(12)

    def _work(self, fn, on_done):
        self._on_done = on_done
        def target():
            try:
                res = fn(lambda m: self.q.put(('log', m)))
                self.q.put(('done', res))
            except Exception as e:
                self.q.put(('log', traceback.format_exc()))
                self.q.put(('error', str(e)))
        self._set_busy(True)
        threading.Thread(target=target, daemon=True).start()

    def _finish(self, ok, payload):
        self._set_busy(False)
        if ok:
            self._on_done(payload)
        else:
            messagebox.showerror(APP_TITLE, f'오류가 났습니다.\n\n{payload}')

    def _out(self):
        out = self.out_var.get().strip()
        if not out:
            messagebox.showwarning(APP_TITLE, '결과 저장 폴더를 정해 주세요.'); return None
        os.makedirs(out, exist_ok=True)
        return out

    def _search_fn(self, out):
        self.stop_ev.clear()
        def prog(i, n):
            if i == 1:
                self.q.put(('log', ''))
            self.q.put(('prog', (i, n)))
        if self.plat.get() == '29CM':
            return lambda log: runner.auto_search_29cm(out, log=log, stop=self.stop_ev, progress=prog)
        return lambda log: runner.auto_search(out, log=log, stop=self.stop_ev, progress=prog)

    def _done_search(self, report):
        if not report:
            return
        p = self.plat.get(); cap = '29CM_미노출_캡처' if p == '29CM' else '미노출_캡처'
        msg = f'점검을 마쳤습니다.\n\n· {p}_검색노출점검 → "검색_미노출" 시트\n· {p}_옵션재고점검 → "상품별_한눈에보기" 시트\n· {cap} 폴더: 검색 안 된 상품의 검색 화면\n\n결과 폴더를 열까요?'
        if self.stop_ev.is_set():
            msg = '검색을 중지했습니다. 지금까지 검색한 결과로 엑셀을 만들었습니다 (나머지는 "미확인").\n[검색만 이어서 하기]로 계속할 수 있습니다.\n\n결과 폴더를 열까요?'
        if messagebox.askyesno(APP_TITLE, msg):
            open_path(os.path.dirname(report))

    def start_all(self):
        if self.busy:
            return
        out = self._out()
        if not out:
            return
        smf = self.sm_var.get().strip()
        self._save_last()
        self.log('')
        search = self._search_fn(out)
        if self.plat.get() == '29CM':
            files = self._cm_files()
            def both(log):
                runner.prepare_29cm(files, smf, out, log=log)
                return search(log)
        else:
            src = self.in_var.get().strip()
            def both(log):
                runner.prepare(src, smf, out, log=log)
                return search(log)
        self._work(both, self._done_search)

    def resume_search(self):
        if self.busy:
            return
        out = self._out()
        if not out:
            return
        self.log('')
        self._work(self._search_fn(out), self._done_search)

    def stop_search(self):
        self.stop_ev.set()
        self.log('중지 요청 — 지금 검색 중인 상품까지 하고 멈춥니다 ...')

    def copy_search(self):
        js = os.path.join(self.out_var.get(), 'musinsa_brand_search_ready.js')
        if not os.path.exists(js):
            messagebox.showinfo(APP_TITLE, '먼저 [▶ 점검 시작]을 한 번 실행해 주세요 (옵션 재고 점검까지 끝나면 중지해도 됩니다).'); return
        with open(js, encoding='utf-8') as f:
            text = f.read()
        self.clipboard_clear(); self.clipboard_append(text); self.update()
        webbrowser.open('https://www.musinsa.com')
        self.log('검색 스크립트를 복사했습니다. 무신사 페이지에서 F12 → Console → Ctrl+V → Enter')

    def paste_search(self):
        if self.busy:
            return
        try:
            text = self.clipboard_get()
        except tk.TclError:
            text = ''
        if '|' not in text:
            messagebox.showwarning(APP_TITLE, '클립보드에 검색 결과가 없습니다.\n무신사 콘솔에 "완료"가 나온 뒤 눌러 주세요.'); return
        out = self._out()
        if not out:
            return
        def done(report):
            if messagebox.askyesno(APP_TITLE, '검색 결과를 두 엑셀에 반영했습니다. 검색 노출 점검 엑셀을 열어볼까요?'):
                open_path(report)
        self._work(lambda log: runner.apply_search(text, out, log=log), done)

    def open_settings(self):
        if not os.path.exists(cfg.SETTINGS_FILE):
            cfg.save_settings(keys=['BRAND_KR', 'EXCLUDE_CODE_PATTERN', 'SEARCH_DROP_WORDS'])
        messagebox.showinfo(APP_TITLE, '설정 파일(settings.json)을 엽니다.\n\n'
                            'BRAND_KR: 무신사·29CM에 표시되는 브랜드명 (검색어 앞에 붙고, 검색 결과의 브랜드를 이 이름으로 확인)\n'
                            'EXCLUDE_CODE_PATTERN: 점검에서 뺄 스타일넘버 규칙 (예: ^(ZD|VV) = ZD·VV로 시작하는 상품 제외)\n'
                            'SEARCH_DROP_WORDS: 검색해서 안 나오면 빼고 다시 검색할 단어 (기본: UNISEX, 유니섹스)\n\n'
                            '저장한 뒤 [▶ 점검 시작]부터 다시 실행하면 반영됩니다.')
        open_path(cfg.SETTINGS_FILE)

    def show_help(self):
        messagebox.showinfo(APP_TITLE, (
            '무신사 판매중 상품 중\n'
            ' · 셀메이트에 재고가 있는데 무신사 옵션이 품절(판매 제한)인 것\n'
            ' · 무신사에서 상품명으로 검색했을 때 안 나오는 것\n'
            '을 찾습니다.\n\n'
            '1) 재고 관리(v2) 표를 복사해 엑셀로 저장한 파일과 셀메이트 재고목록(CSV)을 고릅니다.\n'
            '2) [▶ 점검 시작]\n'
            '   · 옵션 재고 점검 엑셀이 먼저 만들어지고\n'
            '   · 크롬이 열려 무신사 검색창에 "브랜드명 + 상품명(컬러·_3COL 표기 제외)"을 하나씩 입력해 검색합니다.\n'
            '   · 검색 결과 화면에 그 상품이 우리 브랜드로 안 보이면 화면을 캡처해 "미노출_캡처" 폴더에 저장합니다.\n'
            '   · 검색하는 크롬 창은 건드리지 마세요 (최소화는 괜찮습니다). 중간에 멈추면 [검색만 이어서 하기].\n'
            '3) 결과 폴더\n'
            '   · 무신사_검색노출점검 → "검색_미노출" 시트 ("캡처 보기"로 화면 확인)\n'
            '   · 무신사_옵션재고점검 → "상품별_한눈에보기" 시트\n\n'
            '자동 검색이 안 되면 아래 "예비" 버튼으로 크롬 콘솔(F12)에서 검색할 수 있습니다 (무신사만).\n\n'
            '[29CM]\n'
            '판매채널에서 29CM를 고르고, 1번에 "29CM_옵션재고_일괄변경(파트너)_페이지…" 파일을 전부 고릅니다(여러 개 선택).\n'
            '29CM 파일에는 셀메이트 상품코드가 없어서 상품명(컬러·_3Colors 표기 제외, 한글↔영문 포함)으로 셀메이트 상품을 찾습니다.\n'
            '결과: 29CM_옵션재고점검(상품별_한눈에보기·옵션별_대조·29CM에_옵션없음·상품매칭_확인필요), 29CM_검색노출점검, 29CM_미노출_캡처'))

    def _load_last(self):
        try:
            with open(LAST_FILE, encoding='utf-8') as f:
                d = json.load(f)
            if d.get('in'): self.in_var.set(d['in'])
            if d.get('cm29'): self.cm_var.set(d['cm29'])
            if d.get('platform') in PLATFORMS: self.plat.set(d['platform'])
            if d.get('sellmate'): self.sm_var.set(d['sellmate'])
            if d.get('out'): self.out_var.set(d['out'])
        except Exception:
            pass

    def _save_last(self):
        try:
            with open(LAST_FILE, 'w', encoding='utf-8') as f:
                json.dump({'in': self.in_var.get(), 'cm29': self.cm_var.get(), 'platform': self.plat.get(),
                           'sellmate': self.sm_var.get(), 'out': self.out_var.get()}, f, ensure_ascii=False, indent=2)
        except Exception:
            pass


def selftest_29cm(fixture_dir, out_dir, log, auto_base=None):
    """29CM: 페이지 파일 2개(1줄 겹침) + 셀메이트 → 옵션 판정, (가짜 29CM 페이지에서) 자동 검색"""
    import pandas as pd
    out = os.path.join(out_dir, '29cm')
    files = [os.path.join(fixture_dir, f) for f in ('cm29_page1.xlsx', 'cm29_page2.xlsx')]
    P = runner.prepare_29cm(files, os.path.join(fixture_dir, 'sellmate_stock.csv'), out, log=log)
    O = pd.read_excel(P['option_report'], '옵션별_대조', header=1, dtype=str)
    j = {(a, b): c for a, b, c in zip(O['29CM 상품명'], O['29CM 옵션명'], O['판정'])}
    log(f'29CM 옵션 판정: {j}')
    ok = len(O) == 7 and j.get(('Test Crop Tee_2Colors', '[색상]BLACK [사이즈]M')) == '셀메이트 재고 있는데 29CM 품절' \
        and j.get(('테스트 와이드 팬츠 BLUE', '[사이즈]S')) == '셀메이트 재고 있는데 29CM 판매중지'
    if ok and auto_base:
        report = runner.auto_search_29cm(out, log=log, headless=True, base=auto_base)
        R = pd.read_excel(report, '전체_검색결과', header=1, dtype=str)
        st = dict(zip(R['29CM 상품번호'], R['검색여부']))
        log(f'29CM 자동 검색 결과: {st}')
        ok = st.get('3000001') == '검색됨' and st.get('3000002') == '검색 안 됨' and st.get('3000003', '').startswith('검색됨') \
            and 'UNISEX' in st.get('3000004', '') \
            and os.path.exists(os.path.join(out, '29CM_미노출_캡처'))
    log('29CM ' + ('OK' if ok else 'FAIL'))
    return ok


def selftest(fixture_dir, out_dir, auto_base=None, auto_base_29cm=None):
    """화면 없이 실행 (빌드 확인용): SalesChannelChecker.exe --selftest <입력폴더> <결과폴더> [가짜 무신사 주소] [가짜 29CM 주소]"""
    os.makedirs(out_dir, exist_ok=True)
    logf = open(os.path.join(out_dir, 'selftest.log'), 'w', encoding='utf-8')
    log = lambda m: (logf.write(m + '\n'), logf.flush())
    try:
        ok = True
        for name in ['stock_copy_noheader.xlsx', 'stock_copy_header_bottom.xlsx', 'stock_copy_header_top14.xlsx']:
            out = os.path.join(out_dir, name.replace('.xlsx', ''))
            P = runner.prepare(os.path.join(fixture_dir, name), os.path.join(fixture_dir, 'sellmate_stock.csv'), out, log=log)
            with open(os.path.join(fixture_dir, 'search_result.txt'), encoding='utf-8') as f:
                report = runner.apply_search(f.read(), out, log=log)
            ok = ok and os.path.exists(report) and os.path.exists(P['script']) and os.path.exists(P['option_report'])
            if ok and auto_base:   # 크롬 자동 검색 (가짜 무신사 페이지)
                report = runner.auto_search(out, log=log, headless=True, base=auto_base)
                import pandas as _pd
                R = _pd.read_excel(report, '전체_검색결과', header=1, dtype=str)
                st = dict(zip(R.iloc[:, 1], R.iloc[:, 0]))
                log(f'자동 검색 결과: {st}')
                ok = st.get('9000001', '').startswith('검색됨') and st.get('9000002') == '검색 안 됨' \
                    and 'UNISEX' in st.get('9000003', '') \
                    and os.path.exists(os.path.join(out, '미노출_캡처'))
            if ok:   # 셀메이트 재고 5인데 무신사 품절인 9000001 BLACK^M 이 잡혀야 함
                import pandas as pd
                hv = pd.read_excel(P['option_report'], '상품별_한눈에보기', header=1, dtype=str)
                hit = hv[hv.iloc[:, 1] == '9000001']
                ok = len(hit) == 1 and hit.iloc[0, 0].startswith('①')
                log(f'옵션 점검 확인: 9000001 → {hit.iloc[0, 0] if len(hit) else "없음"}')
        ok = selftest_29cm(fixture_dir, out_dir, log, auto_base_29cm) and ok
        log('SELFTEST ' + ('OK' if ok else 'FAIL'))
        return 0 if ok else 1
    except Exception:
        log(traceback.format_exc()); log('SELFTEST FAIL')
        return 1
    finally:
        logf.close()


if __name__ == '__main__':
    if len(sys.argv) >= 4 and sys.argv[1] == '--selftest':
        sys.exit(selftest(sys.argv[2], sys.argv[3], sys.argv[4] if len(sys.argv) > 4 else None,
                          sys.argv[5] if len(sys.argv) > 5 else None))
    App().mainloop()
