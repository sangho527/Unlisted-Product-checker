"""
무신사 미등록·옵션 품절 점검 — 실행 화면 (Windows exe 진입점)
"""
import os, sys, json, queue, threading, traceback, webbrowser
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinter.scrolledtext import ScrolledText

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import runner  # noqa: E402
cfg = runner.cfg

APP_TITLE = '무신사 미등록·옵션 품절 점검'
FONT = ('맑은 고딕', 10)
FONT_B = ('맑은 고딕', 10, 'bold')
LAST_FILE = os.path.join(cfg.ROOT, 'last_paths.json')

INPUTS = [
    ('sellmate',        '셀메이트 재고목록 (.csv)',         '셀메이트 → 재고 목록 → 엑셀 다운로드 (stk_stockList_*.csv)', [('CSV', '*.csv'), ('모든 파일', '*.*')]),
    ('musinsa_all',     '무신사 상품정보조회 – 전체 (.xlsx)', '파트너센터 → 상품정보조회 → 상품상태 "전체" → 엑셀 다운로드',   [('Excel', '*.xlsx'), ('모든 파일', '*.*')]),
    ('musinsa_selling', '무신사 상품정보조회 – 판매중 (.xlsx)', '같은 화면에서 상품상태 "판매중" → 엑셀 다운로드',            [('Excel', '*.xlsx'), ('모든 파일', '*.*')]),
    ('musinsa_option',  '무신사 재고관리(v2) 옵션 목록 (.xlsx)', '재고 관리(v2) → 판매중·출력수 5000 → 표 전체 복사 → 엑셀에 붙여넣어 저장', [('Excel', '*.xlsx'), ('모든 파일', '*.*')]),
]


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
        self.geometry('880x760')
        self.minsize(760, 640)
        self.option_add('*Font', FONT)
        self.q = queue.Queue()
        self.busy = False
        self.vars = {k: tk.StringVar() for k, *_ in INPUTS}
        self.out_var = tk.StringVar(value=os.path.join(cfg.ROOT, '결과'))
        self.listing_only = tk.BooleanVar(value=False)
        self._load_last()
        self._build()
        self.after(100, self._poll)

    # ---------- 화면 ----------
    def _build(self):
        pad = dict(padx=10, pady=4)
        top = ttk.Frame(self); top.pack(fill='x', **pad)
        ttk.Label(top, text=APP_TITLE, font=('맑은 고딕', 14, 'bold')).pack(side='left')
        ttk.Button(top, text='설정(브랜드·수동 연결)', command=self.open_settings).pack(side='right')
        ttk.Button(top, text='사용법', command=self.show_help).pack(side='right', padx=6)

        box = ttk.LabelFrame(self, text=' 1. 입력 파일 '); box.pack(fill='x', **pad)
        for i, (k, label, hint, types) in enumerate(INPUTS):
            ttk.Label(box, text=label, font=FONT_B).grid(row=i * 2, column=0, sticky='w', padx=8, pady=(6, 0))
            e = ttk.Entry(box, textvariable=self.vars[k]); e.grid(row=i * 2, column=1, sticky='ew', padx=4, pady=(6, 0))
            ttk.Button(box, text='찾아보기', command=lambda k=k, t=types: self.pick(k, t)).grid(row=i * 2, column=2, padx=8, pady=(6, 0))
            ttk.Label(box, text=hint, foreground='#666').grid(row=i * 2 + 1, column=1, sticky='w', padx=4)
        box.columnconfigure(1, weight=1)

        ob = ttk.LabelFrame(self, text=' 2. 결과 저장 폴더 '); ob.pack(fill='x', **pad)
        ttk.Entry(ob, textvariable=self.out_var).pack(side='left', fill='x', expand=True, padx=8, pady=6)
        ttk.Button(ob, text='변경', command=self.pick_out).pack(side='left', padx=4)
        ttk.Button(ob, text='폴더 열기', command=self.open_out).pack(side='left', padx=8)

        rb = ttk.Frame(self); rb.pack(fill='x', **pad)
        self.run_btn = ttk.Button(rb, text='▶  점검 시작', command=self.start)
        self.run_btn.pack(side='left', ipadx=20, ipady=6)
        ttk.Checkbutton(rb, text='등록 현황만 (셀메이트 + 상품정보조회 전체 파일만 필요)', variable=self.listing_only).pack(side='left', padx=12)
        self.pb = ttk.Progressbar(rb, mode='indeterminate', length=160); self.pb.pack(side='right')

        sb = ttk.LabelFrame(self, text=' 3. (선택) 무신사 검색 노출 확인 — 점검을 먼저 실행한 뒤 '); sb.pack(fill='x', **pad)
        ttk.Button(sb, text='① 검색 스크립트 복사 + 무신사 열기', command=self.copy_search).grid(row=0, column=0, padx=8, pady=6, sticky='w')
        ttk.Label(sb, foreground='#444', justify='left', wraplength=500, text=(
            '무신사 페이지에서 F12 → Console 탭 → Ctrl+V 붙여넣기 → Enter (처음이면 "allow pasting" 입력 후 다시 붙여넣기)\n'
            '2~3분 뒤 "완료" 메시지가 나오면 결과가 자동으로 복사됩니다.')).grid(row=0, column=1, sticky='w')
        ttk.Button(sb, text='② 복사된 검색 결과 붙여넣기 → 엑셀 반영', command=self.paste_search).grid(row=1, column=0, padx=8, pady=6, sticky='w')

        lb = ttk.LabelFrame(self, text=' 진행 상황 '); lb.pack(fill='both', expand=True, **pad)
        self.log_box = ScrolledText(lb, height=12, font=('Consolas', 9), state='disabled')
        self.log_box.pack(fill='both', expand=True, padx=6, pady=6)
        self.log('입력 파일 4개를 고르고 [점검 시작]을 누르세요. 결과 엑셀은 결과 저장 폴더에 만들어집니다.')

    # ---------- 동작 ----------
    def pick(self, key, types):
        p = filedialog.askopenfilename(title='파일 선택', filetypes=types,
                                       initialdir=os.path.dirname(self.vars[key].get() or cfg.ROOT))
        if p:
            self.vars[key].set(os.path.normpath(p))

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
        self.run_btn.configure(state='disabled' if busy else 'normal')
        (self.pb.start(12) if busy else self.pb.stop())

    def _work(self, fn):
        def target():
            try:
                res = fn(lambda m: self.q.put(('log', m)))
                self.q.put(('done', res))
            except Exception as e:
                self.q.put(('log', traceback.format_exc()))
                self.q.put(('error', str(e)))
        self._set_busy(True)
        threading.Thread(target=target, daemon=True).start()

    def start(self):
        if self.busy:
            return
        files = {k: v.get().strip() for k, v in self.vars.items()}
        out = self.out_var.get().strip()
        if not out:
            messagebox.showwarning(APP_TITLE, '결과 저장 폴더를 정해 주세요.'); return
        os.makedirs(out, exist_ok=True)
        self._save_last()
        self.log('')
        self.log('=== 점검 시작 ===')
        listing_only = self.listing_only.get()
        self._work(lambda log: runner.run_check(files, out, listing_only=listing_only, log=log))

    def _finish(self, ok, payload):
        self._set_busy(False)
        if ok:
            if messagebox.askyesno(APP_TITLE, '점검이 끝났습니다. 결과 폴더를 열까요?'):
                self.open_out()
        else:
            messagebox.showerror(APP_TITLE, f'오류가 났습니다.\n\n{payload}\n\n진행 상황 창의 내용을 확인해 주세요.')

    def copy_search(self):
        js = os.path.join(self.out_var.get(), 'musinsa_brand_search_ready.js')
        if not os.path.exists(js):
            messagebox.showinfo(APP_TITLE, '먼저 [점검 시작]으로 점검을 실행해 주세요. (검색할 상품 목록이 만들어집니다)'); return
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
            messagebox.showwarning(APP_TITLE, '클립보드에 검색 결과가 없습니다.\n무신사 콘솔에서 "완료" 메시지가 나온 뒤 눌러 주세요.'); return
        out = self.out_var.get().strip()
        self._work(lambda log: runner.rebuild_with_search(text, out, log=log))

    def open_settings(self):
        if not os.path.exists(cfg.SETTINGS_FILE):
            cfg.save_settings()
        messagebox.showinfo(APP_TITLE, '설정 파일(settings.json)을 엽니다.\n\n'
                            'BRAND_KR: 브랜드명 / TARGET_CODE_PATTERN: 점검할 상품코드 규칙\n'
                            'EXCLUDE_CODE_PATTERN: 제외할 다른 브랜드 코드 / MANUAL_LINKS: 수동 연결\n\n'
                            '저장한 뒤 다시 [점검 시작]을 누르면 반영됩니다.')
        open_path(cfg.SETTINGS_FILE)

    def show_help(self):
        messagebox.showinfo(APP_TITLE, (
            '1) 입력 파일 4개를 고릅니다.\n'
            '2) [점검 시작] → 결과 폴더에 엑셀 2개가 만들어집니다.\n'
            '   · 무신사_옵션재고점검_날짜.xlsx  ← 메인 (상품별_한눈에보기부터 보세요)\n'
            '   · 무신사_어반드레스_등록점검_날짜.xlsx\n'
            '3) (선택) 검색 노출 확인: ① 버튼 → 무신사 콘솔에 붙여넣기 → 끝나면 ② 버튼\n\n'
            '각 시트 맨 위에 "왜 만들었나 / 읽는 법"이 적혀 있습니다.\n'
            '재고·원가가 들어 있는 파일이니 외부에 공유하지 마세요.'))

    def _load_last(self):
        try:
            with open(LAST_FILE, encoding='utf-8') as f:
                d = json.load(f)
            for k, v in self.vars.items():
                if d.get(k): v.set(d[k])
            if d.get('out'): self.out_var.set(d['out'])
        except Exception:
            pass

    def _save_last(self):
        try:
            d = {k: v.get() for k, v in self.vars.items()}; d['out'] = self.out_var.get()
            with open(LAST_FILE, 'w', encoding='utf-8') as f:
                json.dump(d, f, ensure_ascii=False, indent=2)
        except Exception:
            pass


if __name__ == '__main__':
    App().mainloop()
