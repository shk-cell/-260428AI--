import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
import threading
import time

try:
    from selenium import webdriver
    from selenium.webdriver.common.by import By
    from selenium.webdriver.chrome.options import Options
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    HAS_SELENIUM = True
except ImportError:
    HAS_SELENIUM = False

class TagMacro:
    def __init__(self, root):
        self.root = root
        self.root.title("HTML 태그 탐색 매크로")
        self.root.geometry("580x750")
        self.root.resizable(False, False)

        self.driver  = None
        self.running = False
        self.actions = []  # [(by_label, selector, action, input_text, delay), ...]

        self._build_ui()

        if not HAS_SELENIUM:
            messagebox.showwarning("라이브러리 없음",
                "selenium이 설치되지 않았습니다.\n\n"
                "터미널에서 실행하세요:\n pip install selenium")

    def _build_ui(self):
        # ── 브라우저 설정 ──
        frm0 = tk.LabelFrame(self.root, text="브라우저 설정", padx=10, pady=8)
        frm0.pack(fill="x", padx=12, pady=(12,4))

        tk.Label(frm0, text="URL:").grid(row=0, column=0, sticky="e", pady=4)
        self.url_var = tk.StringVar(value="http://localhost:3000")
        tk.Entry(frm0, textvariable=self.url_var, width=36).grid(row=0, column=1, columnspan=2, padx=6)

        tk.Label(frm0, text="브라우저:").grid(row=1, column=0, sticky="e", pady=4)
        self.browser_var = tk.StringVar(value="Chrome")
        ttk.Combobox(frm0, textvariable=self.browser_var,
                     values=["Chrome", "Edge"], width=12, state="readonly").grid(row=1, column=1, padx=6, sticky="w")

        self.connect_btn = tk.Button(frm0, text="브라우저 열기", command=self._open_browser,
                                     bg="#2563eb", fg="white", width=14)
        self.connect_btn.grid(row=1, column=2, padx=6)

        # ── 동작 추가 ──
        frm = tk.LabelFrame(self.root, text="동작 추가", padx=10, pady=8)
        frm.pack(fill="x", padx=12, pady=4)

        tk.Label(frm, text="탐색 방식:").grid(row=0, column=0, sticky="e", pady=4)
        self.by_var = tk.StringVar(value="ID")
        by_box = ttk.Combobox(frm, textvariable=self.by_var,
                               values=["ID", "CSS 선택자", "클래스명", "태그명", "XPath", "텍스트 포함"],
                               width=14, state="readonly")
        by_box.grid(row=0, column=1, padx=6, sticky="w")
        by_box.bind("<<ComboboxSelected>>", self._update_hint)

        tk.Label(frm, text="값:").grid(row=1, column=0, sticky="e", pady=4)
        self.selector_var = tk.StringVar()
        tk.Entry(frm, textvariable=self.selector_var, width=32).grid(row=1, column=1, columnspan=2, padx=6)

        self.hint_var = tk.StringVar(value="예: doneBtn  /  quizDoneBtn")
        tk.Label(frm, textvariable=self.hint_var, fg="#888", font=("", 9)).grid(
            row=2, column=1, columnspan=2, sticky="w", padx=6)

        tk.Label(frm, text="동작:").grid(row=3, column=0, sticky="e", pady=4)
        self.action_var = tk.StringVar(value="클릭")
        ttk.Combobox(frm, textvariable=self.action_var,
                     values=["클릭", "더블클릭", "텍스트 입력", "값 읽기"],
                     width=14, state="readonly").grid(row=3, column=1, padx=6, sticky="w")

        tk.Label(frm, text="입력 텍스트:").grid(row=4, column=0, sticky="e", pady=4)
        self.input_text_var = tk.StringVar()
        tk.Entry(frm, textvariable=self.input_text_var, width=22).grid(row=4, column=1, padx=6, sticky="w")

        tk.Label(frm, text="실행 전 대기(초):").grid(row=4, column=2, sticky="e")
        self.delay_var = tk.StringVar(value="1.0")
        tk.Entry(frm, textvariable=self.delay_var, width=6).grid(row=4, column=3, padx=6)

        tk.Button(frm, text="동작 추가 ＋", command=self._add_action,
                  bg="#2563eb", fg="white", width=12).grid(row=5, column=0, columnspan=4, pady=(8,2))

        # ── 동작 목록 ──
        frm2 = tk.LabelFrame(self.root, text="동작 목록 (순서대로 실행)", padx=10, pady=8)
        frm2.pack(fill="x", padx=12, pady=4)

        cols = ("순서", "방식", "값", "동작", "대기(초)")
        self.tree = ttk.Treeview(frm2, columns=cols, show="headings", height=5)
        widths = [45, 90, 160, 80, 65]
        for c, w in zip(cols, widths):
            self.tree.heading(c, text=c)
            self.tree.column(c, width=w, anchor="center")
        self.tree.pack(fill="x")

        tk.Button(frm2, text="선택 삭제", command=self._del_action).pack(anchor="e", pady=(4,0))

        # ── 반복 설정 ──
        frm3 = tk.LabelFrame(self.root, text="반복 설정", padx=10, pady=8)
        frm3.pack(fill="x", padx=12, pady=4)

        tk.Label(frm3, text="반복 횟수 (0=무한):").grid(row=0, column=0, sticky="e")
        self.repeat_var = tk.StringVar(value="1")
        tk.Entry(frm3, textvariable=self.repeat_var, width=8).grid(row=0, column=1, padx=6)

        tk.Label(frm3, text="반복 간격(초):").grid(row=0, column=2, sticky="e")
        self.interval_var = tk.StringVar(value="3.0")
        tk.Entry(frm3, textvariable=self.interval_var, width=8).grid(row=0, column=3, padx=6)

        tk.Label(frm3, text="요소 대기(초):").grid(row=1, column=0, sticky="e", pady=4)
        self.wait_var = tk.StringVar(value="5")
        tk.Entry(frm3, textvariable=self.wait_var, width=8).grid(row=1, column=1, padx=6)
        tk.Label(frm3, text="(요소를 못 찾을 때 기다리는 시간)", fg="#888", font=("",9)).grid(row=1, column=2, columnspan=2, sticky="w")

        # ── 실행 버튼 ──
        ctrl = tk.Frame(self.root)
        ctrl.pack(pady=8)

        self.start_btn = tk.Button(ctrl, text="▶  실행", command=self._start,
                                   bg="#16a34a", fg="white", width=12, font=("", 11, "bold"))
        self.start_btn.pack(side="left", padx=6)

        self.stop_btn = tk.Button(ctrl, text="■  정지", command=self._stop,
                                  bg="#dc2626", fg="white", width=12, font=("", 11, "bold"), state="disabled")
        self.stop_btn.pack(side="left", padx=6)

        # ── 현재 동작 상황 ──
        frm5 = tk.LabelFrame(self.root, text="현재 동작 상황", padx=8, pady=6)
        frm5.pack(fill="x", padx=12, pady=(0,4))

        self.status_var = tk.StringVar(value="대기 중")
        tk.Label(frm5, textvariable=self.status_var,
                 bg="#1e293b", fg="#7dd3fc",
                 font=("Consolas", 10, "bold"),
                 anchor="w", padx=10, pady=6,
                 relief="flat").pack(fill="x")

        # ── 로그 ──
        frm4 = tk.LabelFrame(self.root, text="로그", padx=6, pady=4)
        frm4.pack(fill="both", expand=True, padx=12, pady=(0,12))
        self.log = scrolledtext.ScrolledText(frm4, height=4, state="disabled",
                                              bg="#f8f8f8", font=("Consolas", 9))
        self.log.pack(fill="both", expand=True)

    def _update_hint(self, _=None):
        hints = {
            "ID":          "예: doneBtn  /  quizDoneBtn",
            "CSS 선택자":  "예: button#doneBtn  /  .btn-primary",
            "클래스명":    "예: btn-primary",
            "태그명":      "예: button  /  video",
            "XPath":       "예: //button[contains(text(),'완료')]",
            "텍스트 포함": "예: 완료하기",
        }
        self.hint_var.set(hints.get(self.by_var.get(), ""))

    def _add_action(self):
        sel = self.selector_var.get().strip()
        if not sel:
            messagebox.showwarning("경고", "값을 입력하세요.")
            return
        try:
            delay = float(self.delay_var.get())
        except ValueError:
            messagebox.showerror("오류", "대기 시간을 숫자로 입력하세요.")
            return

        by     = self.by_var.get()
        action = self.action_var.get()
        itext  = self.input_text_var.get()

        self.actions.append((by, sel, action, itext, delay))
        self.tree.insert("", "end", values=(len(self.actions), by, sel, action, delay))
        self.selector_var.set("")

    def _del_action(self):
        sel = self.tree.selection()
        if not sel:
            return
        for item in sel:
            idx = self.tree.index(item)
            self.tree.delete(item)
            self.actions.pop(idx)
        for i, item in enumerate(self.tree.get_children()):
            vals = list(self.tree.item(item, "values"))
            vals[0] = i + 1
            self.tree.item(item, values=vals)

    def _open_browser(self):
        if not HAS_SELENIUM:
            messagebox.showerror("오류", "selenium을 먼저 설치하세요.\npip install selenium")
            return

        def _open():
            try:
                self._log("브라우저 열기...")
                if self.browser_var.get() == "Edge":
                    from selenium.webdriver.edge.options import Options as EdgeOptions
                    opts = EdgeOptions()
                    self.driver = webdriver.Edge(options=opts)
                else:
                    opts = Options()
                    self.driver = webdriver.Chrome(options=opts)

                self.driver.get(self.url_var.get())
                self._log(f"브라우저 열림: {self.url_var.get()}")
                self.status_var.set("브라우저 연결됨")
                self.root.after(0, lambda: self.connect_btn.config(
                    text="브라우저 닫기", command=self._close_browser))
            except Exception as e:
                self._log(f"오류: {e}")
                messagebox.showerror("브라우저 오류", str(e))

        threading.Thread(target=_open, daemon=True).start()

    def _close_browser(self):
        if self.driver:
            self.driver.quit()
            self.driver = None
        self.connect_btn.config(text="브라우저 열기", command=self._open_browser)
        self.status_var.set("브라우저 닫힘")
        self._log("브라우저 닫힘")

    def _get_by(self, by_label):
        mapping = {
            "ID":          By.ID,
            "CSS 선택자":  By.CSS_SELECTOR,
            "클래스명":    By.CLASS_NAME,
            "태그명":      By.TAG_NAME,
            "XPath":       By.XPATH,
            "텍스트 포함": By.XPATH,
        }
        return mapping[by_label]

    def _get_selector(self, by_label, sel):
        if by_label == "텍스트 포함":
            return f"//*[contains(text(),'{sel}')]"
        return sel

    def _execute_one(self, by_label, sel, action, itext, delay):
        by       = self._get_by(by_label)
        selector = self._get_selector(by_label, sel)
        wait_sec = float(self.wait_var.get())

        self._set_status(f"'{sel}' 탐색합니다...")
        self._log(f"'{sel}' 탐색 시작")

        # 탐색
        try:
            elem = WebDriverWait(self.driver, wait_sec).until(
                EC.element_to_be_clickable((by, selector))
            )
        except Exception:
            self._set_status(f"❌ '{sel}' 요소를 찾지 못했습니다. (없거나 비활성 상태)")
            self._log(f"'{sel}' 요소를 찾지 못했습니다.")
            time.sleep(2)
            return False

        # 찾은 후 대기 및 실행
        self._set_status(f"✅ '{sel}' 요소를 찾았습니다. {delay}초 뒤 {action}합니다.")
        self._log(f"'{sel}' 요소 발견 → {delay}초 대기 후 {action}")
        time.sleep(delay)

        try:
            if action == "클릭":
                elem.click()
                self._set_status(f"'{sel}' 클릭 완료!")
                self._log(f"'{sel}' 클릭 완료")
            elif action == "더블클릭":
                from selenium.webdriver.common.action_chains import ActionChains
                ActionChains(self.driver).double_click(elem).perform()
                self._set_status(f"'{sel}' 더블클릭 완료!")
                self._log(f"'{sel}' 더블클릭 완료")
            elif action == "텍스트 입력":
                elem.clear()
                elem.send_keys(itext)
                self._set_status(f"'{sel}' 텍스트 입력 완료!")
                self._log(f"'{sel}' 텍스트 입력 완료: '{itext}'")
            elif action == "값 읽기":
                txt = elem.text or elem.get_attribute("value") or ""
                self._set_status(f"'{sel}' 값 읽기: {txt}")
                self._log(f"'{sel}' 값 읽기: '{txt}'")
        except Exception as e:
            self._set_status(f"⚠️ '{sel}' {action} 실패: {e}")
            self._log(f"'{sel}' {action} 실패: {e}")
            return False

        return True

    def _start(self):
        if not self.driver:
            messagebox.showwarning("경고", "브라우저를 먼저 여세요.")
            return
        if not self.actions:
            messagebox.showwarning("경고", "동작을 먼저 추가하세요.")
            return
        self.running = True
        self.start_btn.config(state="disabled")
        self.stop_btn.config(state="normal")
        threading.Thread(target=self._run_loop, daemon=True).start()

    def _stop(self):
        self.running = False
        self.start_btn.config(state="normal")
        self.stop_btn.config(state="disabled")
        self.status_var.set("정지됨")

    def _run_loop(self):
        repeat   = int(self.repeat_var.get())
        interval = float(self.interval_var.get())
        loop     = 0

        while self.running:
            # ① 대기
            if interval > 0:
                for remaining in range(int(interval), 0, -1):
                    if not self.running:
                        return
                    self.status_var.set(f"대기 중... {remaining}초 후 탐색 시작")
                    time.sleep(1)

            if not self.running:
                return

            # ② 순차 탐색 실행
            loop += 1
            self._log(f"── {loop}회차 탐색 시작 ──")
            for i, (by, sel, action, itext, delay) in enumerate(self.actions):
                if not self.running:
                    return
                self._execute_one(by, sel, action, itext, delay)

            self.status_var.set(f"[{loop}회차] 완료 → 다음 대기 중...")

            if repeat != 0 and loop >= repeat:
                break

        self.root.after(0, self._stop)
        self._log("전체 완료!")

    def _set_status(self, msg):
        self.root.after(0, lambda: self.status_var.set(msg))

    def _log(self, msg):
        def _write():
            self.log.config(state="normal")
            self.log.insert("end", f"[{time.strftime('%H:%M:%S')}] {msg}\n")
            self.log.see("end")
            self.log.config(state="disabled")
        self.root.after(0, _write)

    def on_close(self):
        if self.driver:
            self.driver.quit()
        self.root.destroy()


if __name__ == "__main__":
    root = tk.Tk()
    app = TagMacro(root)
    root.protocol("WM_DELETE_WINDOW", app.on_close)
    root.mainloop()
