"""
무신사 검색 노출 점검 — 실행 화면 (Windows exe 진입점)
"""
import os, sys, json, queue, threading, traceback, webbrowser
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinter.scrolledtext import ScrolledText

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import runner  # noqa: E402
cfg = runner.cfg

APP_TITLE = '무신사 검색 노출 점검'
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
        self.geometry('860x700')
        self.minsize(740, 600)
        self.option_add('*Font', FONT)
        self.q = queue.Queue()
        self.busy = False
        self.in_var = tk.StringVar()
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

        box = ttk.LabelFrame(self, text=' 1. 무신사 재고 관리(v2) 복사본 '); box.pack(fill='x', **pad)
        ttk.Entry(box, textvariable=self.in_var).grid(row=0, column=0, sticky='ew', padx=8, pady=(8, 0))
        ttk.Button(box, text='찾아보기', command=self.pick_in).grid(row=0, column=1, padx=8, pady=(8, 0))
        ttk.Label(box, foreground='#555', justify='left', wraplength=760, text=(
            '파트너센터 → 상품/재고 → 재고 관리(v2) → 상품상태 "판매중", 출력수 5000 → 표를 전부 드래그해 복사(Ctrl+C) '
            '→ 엑셀 새 파일에 붙여넣기(Ctrl+V) → 저장.  머리글 없이 붙여넣어도 되고, 필요한 열만 자동으로 골라 씁니다. '
            '(5000건이 넘으면 2페이지도 이어서 붙여넣기)')).grid(row=1, column=0, columnspan=2, sticky='w', padx=8, pady=(2, 8))
        box.columnconfigure(0, weight=1)

        ob = ttk.LabelFrame(self, text=' 2. 결과 저장 폴더 '); ob.pack(fill='x', **pad)
        ttk.Entry(ob, textvariable=self.out_var).pack(side='left', fill='x', expand=True, padx=8, pady=6)
        ttk.Button(ob, text='변경', command=self.pick_out).pack(side='left', padx=4)
        ttk.Button(ob, text='폴더 열기', command=self.open_out).pack(side='left', padx=8)

        sb = ttk.LabelFrame(self, text=' 3. 점검 '); sb.pack(fill='x', **pad)
        self.btn1 = ttk.Button(sb, text='① 상품 목록 만들기', command=self.prepare)
        self.btn1.grid(row=0, column=0, padx=8, pady=6, sticky='ew', ipady=3)
        ttk.Label(sb, foreground='#444', text='판매중 상품을 묶고 검색어(상품명에서 컬러·COL 표기 제외)를 만듭니다.').grid(row=0, column=1, sticky='w')
        self.btn2 = ttk.Button(sb, text='② 검색 스크립트 복사 + 무신사 열기', command=self.copy_search)
        self.btn2.grid(row=1, column=0, padx=8, pady=6, sticky='ew', ipady=3)
        ttk.Label(sb, foreground='#444', justify='left', wraplength=470, text=(
            '열린 무신사 페이지(크롬)에서 F12 → Console 탭 → Ctrl+V → Enter\n'
            '(처음이면 "allow pasting" 입력 후 Enter, 다시 Ctrl+V). 끝나면 "완료"가 뜨고 결과가 자동 복사됩니다.')).grid(row=1, column=1, sticky='w')
        self.btn3 = ttk.Button(sb, text='③ 검색 결과 붙여넣기 → 엑셀 완성', command=self.paste_search)
        self.btn3.grid(row=2, column=0, padx=8, pady=6, sticky='ew', ipady=3)
        ttk.Label(sb, foreground='#444', text='콘솔에 "완료"가 나온 뒤 누르세요. 결과 엑셀의 "검색_미노출" 시트를 보면 됩니다.').grid(row=2, column=1, sticky='w')
        self.pb = ttk.Progressbar(sb, mode='indeterminate', length=160)
        self.pb.grid(row=3, column=0, columnspan=2, sticky='ew', padx=8, pady=(0, 6))
        sb.columnconfigure(1, weight=1)

        lb = ttk.LabelFrame(self, text=' 진행 상황 '); lb.pack(fill='both', expand=True, **pad)
        self.log_box = ScrolledText(lb, height=10, font=('Consolas', 9), state='disabled')
        self.log_box.pack(fill='both', expand=True, padx=6, pady=6)
        self.log('재고 관리(v2) 복사본을 고르고 ① → ② → ③ 순서로 누르세요.')

    # ---------- 동작 ----------
    def pick_in(self):
        p = filedialog.askopenfilename(title='재고 관리(v2) 복사본 선택', filetypes=[('Excel', '*.xlsx'), ('모든 파일', '*.*')],
                                       initialdir=os.path.dirname(self.in_var.get() or cfg.ROOT))
        if p:
            self.in_var.set(os.path.normpath(p))

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
                elif kind == 'done':
                    self._finish(True, payload)
                elif kind == 'error':
                    self._finish(False, payload)
        except queue.Empty:
            pass
        self.after(100, self._poll)

    def _set_busy(self, busy):
        self.busy = busy
        for b in (self.btn1, self.btn2, self.btn3):
            b.configure(state='disabled' if busy else 'normal')
        (self.pb.start(12) if busy else self.pb.stop())

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

    def prepare(self):
        if self.busy:
            return
        out = self._out()
        if not out:
            return
        src = self.in_var.get().strip()
        self._save_last()
        self.log('')
        self._work(lambda log: runner.prepare(src, out, log=log),
                   lambda P: messagebox.showinfo(APP_TITLE, '상품 목록을 만들었습니다.\n이제 ② 버튼을 눌러 무신사 검색을 실행하세요.'))

    def copy_search(self):
        js = os.path.join(self.out_var.get(), 'musinsa_brand_search_ready.js')
        if not os.path.exists(js):
            messagebox.showinfo(APP_TITLE, '먼저 ① 상품 목록 만들기를 실행해 주세요.'); return
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
            if messagebox.askyesno(APP_TITLE, '결과 엑셀을 만들었습니다. 열어볼까요?'):
                open_path(report)
        self._work(lambda log: runner.apply_search(text, out, log=log), done)

    def open_settings(self):
        if not os.path.exists(cfg.SETTINGS_FILE):
            cfg.save_settings(keys=['BRAND_KR', 'EXCLUDE_CODE_PATTERN'])
        messagebox.showinfo(APP_TITLE, '설정 파일(settings.json)을 엽니다.\n\n'
                            'BRAND_KR: 무신사에 표시되는 브랜드명 (검색어 앞에 붙고, 검색 결과의 브랜드를 이 이름으로 확인)\n'
                            'EXCLUDE_CODE_PATTERN: 점검에서 뺄 스타일넘버 규칙 (예: ^(ZD|VV) = ZD·VV로 시작하는 상품 제외)\n\n'
                            '저장한 뒤 ① 상품 목록 만들기부터 다시 실행하면 반영됩니다.')
        open_path(cfg.SETTINGS_FILE)

    def show_help(self):
        messagebox.showinfo(APP_TITLE, (
            '판매중인데 무신사에서 상품명으로 검색했을 때 안 나오는 상품을 찾습니다.\n\n'
            '1) 재고 관리(v2) 표를 복사해 엑셀로 저장한 파일을 고릅니다.\n'
            '2) ① 상품 목록 만들기\n'
            '3) ② 검색 스크립트 복사 + 무신사 열기 → 크롬 F12 → Console → Ctrl+V → Enter\n'
            '   (상품 1,000개 기준 약 10분, 그동안 그 탭을 닫지 마세요)\n'
            '4) 콘솔에 "완료"가 나오면 ③ 검색 결과 붙여넣기\n'
            '5) 결과 엑셀 "검색_미노출" 시트 확인\n\n'
            '검색은 "브랜드명 + 상품명(컬러·_3COL 표기 제외)"으로 하고,\n'
            '검색 결과 화면에 그 상품이 우리 브랜드로 나오는지 확인합니다.'))

    def _load_last(self):
        try:
            with open(LAST_FILE, encoding='utf-8') as f:
                d = json.load(f)
            if d.get('in'): self.in_var.set(d['in'])
            if d.get('out'): self.out_var.set(d['out'])
        except Exception:
            pass

    def _save_last(self):
        try:
            with open(LAST_FILE, 'w', encoding='utf-8') as f:
                json.dump({'in': self.in_var.get(), 'out': self.out_var.get()}, f, ensure_ascii=False, indent=2)
        except Exception:
            pass


def selftest(fixture_dir, out_dir):
    """화면 없이 실행 (빌드 확인용): MusinsaChecker.exe --selftest <입력폴더> <결과폴더>"""
    os.makedirs(out_dir, exist_ok=True)
    logf = open(os.path.join(out_dir, 'selftest.log'), 'w', encoding='utf-8')
    log = lambda m: (logf.write(m + '\n'), logf.flush())
    try:
        ok = True
        for name in ['stock_copy_noheader.xlsx', 'stock_copy_header_bottom.xlsx', 'stock_copy_header_top14.xlsx']:
            out = os.path.join(out_dir, name.replace('.xlsx', ''))
            P = runner.prepare(os.path.join(fixture_dir, name), out, log=log)
            with open(os.path.join(fixture_dir, 'search_result.txt'), encoding='utf-8') as f:
                report = runner.apply_search(f.read(), out, log=log)
            ok = ok and os.path.exists(report) and os.path.exists(P['script'])
        log('SELFTEST ' + ('OK' if ok else 'FAIL'))
        return 0 if ok else 1
    except Exception:
        log(traceback.format_exc()); log('SELFTEST FAIL')
        return 1
    finally:
        logf.close()


if __name__ == '__main__':
    if len(sys.argv) >= 4 and sys.argv[1] == '--selftest':
        sys.exit(selftest(sys.argv[2], sys.argv[3]))
    App().mainloop()
