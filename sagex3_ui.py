import os
import shutil
import queue
import threading
import traceback
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from datetime import datetime

try:
    from PIL import Image, ImageTk
    _PIL_OK = True
except Exception:
    _PIL_OK = False

import sagex3

# Assets are bundled read-only resources; use the PyInstaller resource dir when
# frozen so the logo is found inside the packaged executable.
BASE_DIR = getattr(sagex3, "RESOURCE_DIR", os.path.dirname(os.path.abspath(__file__)))
LOGO_WHITE = os.path.join(BASE_DIR, "assets", "tiauto_logo_white.png")
LOGO_FULL = os.path.join(BASE_DIR, "assets", "tiauto_logo.png")
LOGO_DARK = os.path.join(BASE_DIR, "assets", "tiauto_logo_dark.png")
LOGO_RED = os.path.join(BASE_DIR, "assets", "tiauto_logo_red.png")

# -------------------------------------------------------------- palette
# Futuristic 25th century dark neon theme
BG = "#050510"            # deep space background
PANEL = "#0a0a1a"         # card background
PANEL_HI = "#101025"      # elevated panel
BORDER = "#1f1f3f"        # neon border
GREEN = "#00ff9d"         # neon cyan-green
GREEN_DK = "#00cc7d"      # darker neon
GREEN_SOFT = "#22d3ee"    # cyan accent
FG = "#e6e6ff"            # primary text
MUTED = "#7c7c9d"         # secondary text
AMBER = "#f59e0b"
RED = "#ff3366"
CONSOLE_BG = "#060614"
BLUE = "#00e5ff"          # neon blue/cyan
BLUE_DK = "#00bcd4"
BLUE_SOFT = "#7dd3fc"

LOG_DIR = sagex3.LOG_DIR
STEPS_LOG = sagex3.STEPS_LOG
DEFAULT_FILE = sagex3.file_path


class Spinner(tk.Canvas):
    """Smooth professional blue arc spinner."""

    def __init__(self, parent, size=22, bg=PANEL, color=BLUE, width=3):
        super().__init__(parent, width=size, height=size, bg=bg,
                         highlightthickness=0, bd=0)
        self._size = size
        self._color = color
        self._width = width
        self._angle = 0
        self._running = False
        self._job = None
        self._arc = self.create_arc(2, 2, size - 2, size - 2, start=0, extent=90,
                                    style="arc", outline=color, width=width)

    def start(self):
        if self._running:
            return
        self._running = True
        self._spin()

    def stop(self):
        self._running = False
        if self._job:
            self.after_cancel(self._job)
            self._job = None
        self.itemconfigure(self._arc, state="hidden")

    def _spin(self):
        if not self._running:
            return
        self._angle = (self._angle - 12) % 360
        self.itemconfigure(self._arc, start=self._angle, state="normal")
        self._job = self.after(40, self._spin)


class ProgressPulse(tk.Canvas):
    """Indeterminate blue progress bar for long-running work."""

    def __init__(self, parent, height=4, bg=PANEL):
        super().__init__(parent, height=height, bg=bg, highlightthickness=0, bd=0)
        self._h = height
        self._pos = 0.0
        self._running = False
        self._job = None
        self._bar = self.create_rectangle(0, 0, 0, height, fill=BLUE, outline="")
        self.bind("<Configure>", self._draw)

    def start(self):
        if self._running:
            return
        self._running = True
        self.pack(fill="x") if not self.winfo_ismapped() else None
        self._animate()

    def stop(self):
        self._running = False
        if self._job:
            self.after_cancel(self._job)
            self._job = None
        self.coords(self._bar, 0, 0, 0, self._h)

    def _animate(self):
        if not self._running:
            return
        w = max(self.winfo_width(), 1)
        seg = w * 0.28
        x = self._pos * (w + seg) - seg
        self.coords(self._bar, x, 0, x + seg, self._h)
        self._pos = (self._pos + 0.02) % 1.0
        self._job = self.after(30, self._animate)

    def _draw(self, _evt=None):
        pass


class APP(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("AP Payment Process")
        self.geometry("1120x760")
        self.minsize(980, 640)
        self.configure(bg=BG)
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self._set_window_icon()

        self.log_queue = queue.Queue()
        self.stop_requested = threading.Event()
        self.worker = None
        self.current_file = DEFAULT_FILE
        self.headless_var = tk.BooleanVar(value=True)

        self._build_style()
        self._build_header()
        self._build_body()
        self._build_footer()

        self.reload_file(self.current_file, silent=True)
        self.after(80, self._drain_log)

    def _set_window_icon(self):
        """Sets the taskbar/titlebar icon from the TiAuto logo."""
        if not _PIL_OK:
            return
        path = LOGO_FULL if os.path.exists(LOGO_FULL) else LOGO_WHITE
        if not os.path.exists(path):
            return
        try:
            src = Image.open(path).convert("RGBA")
            # square canvas so the icon is not distorted
            side = max(src.size)
            sq = Image.new("RGBA", (side, side), (0, 0, 0, 0))
            sq.paste(src, ((side - src.width) // 2, (side - src.height) // 2), src)
            sq = sq.resize((64, 64), Image.LANCZOS)
            self._icon_img = ImageTk.PhotoImage(sq)
            self.iconphoto(True, self._icon_img)
        except Exception:
            pass

    # --------------------------------------------------------------- styling
    def _build_style(self):
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure(".", background=BG, foreground=FG)
        style.configure("TFrame", background=BG)
        style.configure("TLabel", background=BG, foreground=FG)
        style.configure("Card.TFrame", background=PANEL)
        style.configure("Treeview", background=PANEL, fieldbackground=PANEL,
                        foreground=FG, rowheight=28, borderwidth=1)
        style.configure("Treeview.Heading", background=PANEL_HI, foreground=BLUE,
                        font=("Consolas", 9), relief="flat", padding=6)
        style.map("Treeview", background=[("selected", "#003344")],
                  foreground=[("selected", GREEN)])
        style.map("Treeview.Heading", background=[("active", "#151535")])

    # --------------------------------------------------------------- header
    def _build_header(self):
        bar = tk.Frame(self, bg=PANEL, height=78)
        bar.pack(fill="x", side="top")
        bar.pack_propagate(False)

        # neon top accent bar
        accent = tk.Frame(self, bg=BLUE, height=2)
        accent.pack(fill="x", side="top")
        accent.pack_propagate(False)

        # logo
        self._logo_img = None
        if _PIL_OK and os.path.exists(LOGO_WHITE):
            try:
                src = Image.open(LOGO_WHITE).convert("RGBA")
                hh = 48
                hw = max(1, int(src.width * hh / src.height))
                src = src.resize((hw, hh), Image.LANCZOS)
                self._logo_img = ImageTk.PhotoImage(src)
                logo_lbl = tk.Label(bar, image=self._logo_img, bg=PANEL, bd=0,
                                    highlightthickness=0)
                logo_lbl.pack(side="left", padx=(22, 16), pady=15)
            except Exception:
                self._logo_img = None
        if self._logo_img is None:
            logo = tk.Canvas(bar, width=48, height=48, bg=PANEL, highlightthickness=0)
            logo.pack(side="left", padx=(22, 14), pady=15)
            logo.create_oval(4, 4, 44, 44, outline=BLUE, width=2)
            logo.create_text(24, 24, text="AP", fill=GREEN,
                             font=("Consolas", 14, "bold"))

        title = tk.Frame(bar, bg=PANEL)
        title.pack(side="left", pady=16)
        tk.Label(title, text="AP PAYMENT PROCESS // NEXUS-25", bg=PANEL, fg=GREEN,
                 font=("Consolas", 16, "bold")).pack(anchor="w")
        tk.Label(title, text="SAGE X3 • AUTONOMOUS REMITTANCE INTERFACE",
                 bg=PANEL, fg=BLUE_SOFT, font=("Consolas", 8)).pack(anchor="w")

        right = tk.Frame(bar, bg=PANEL)
        right.pack(side="right", padx=22)
        self.spinner = Spinner(right, size=18, bg=PANEL, color=GREEN)
        self.spinner.pack(side="left", padx=(0, 10))
        self.state_dot = tk.Canvas(right, width=10, height=10, bg=PANEL,
                                   highlightthickness=0)
        self.state_dot.pack(side="left", padx=(0, 8))
        self._dot = self.state_dot.create_oval(1, 1, 9, 9, fill=MUTED, outline=BLUE)
        self.state_label = tk.Label(right, text="STANDBY", bg=PANEL, fg=MUTED,
                                    font=("Consolas", 9))
        self.state_label.pack(side="left")

    # --------------------------------------------------------------- body
    def _build_body(self):
        body = tk.Frame(self, bg=BG)
        body.pack(fill="both", expand=True, padx=16, pady=14)

        top = tk.Frame(body, bg=BG)
        top.pack(fill="x")

        self._build_file_card(top)
        self._build_config_card(top)
        self._build_controls(body)

        lower = tk.Frame(body, bg=BG)
        lower.pack(fill="both", expand=True, pady=(4, 0))
        self._build_table_card(lower)
        self._build_log_card(lower)

    def _card(self, parent, **kw):
        return tk.Frame(parent, bg=PANEL, highlightthickness=1,
                        highlightbackground=BORDER, **kw)

    def _build_file_card(self, parent):
        card = self._card(parent)
        card.pack(side="left", fill="both", expand=True)

        tk.Label(card, text="[ DATA SOURCE ]", bg=PANEL, fg=GREEN,
                 font=("Consolas", 8)).pack(anchor="w", padx=18, pady=(14, 4))

        row = tk.Frame(card, bg=PANEL)
        row.pack(fill="x", padx=18, pady=(0, 8))
        self.file_label = tk.Label(row, text=os.path.basename(self.current_file),
                                   bg=PANEL, fg=FG, font=("Consolas", 10),
                                   anchor="w")
        self.file_label.pack(side="left", fill="x", expand=True)

        self.upload_btn = tk.Button(
            card, text="> LOAD FILE", command=self.upload_file,
            bg=PANEL_HI, fg=GREEN, activebackground="#003322",
            activeforeground=GREEN, relief="flat", bd=0, cursor="hand2",
            font=("Consolas", 9), padx=16, pady=6, highlightbackground=BORDER,
            highlightthickness=1)
        self.upload_btn.pack(anchor="w", padx=18, pady=(0, 16))

        self.file_meta = tk.Label(card, text="", bg=PANEL, fg=MUTED,
                                  font=("Consolas", 7), anchor="w", justify="left")
        self.file_meta.pack(anchor="w", padx=18, pady=(0, 14))

    def _build_config_card(self, parent):
        card = self._card(parent)
        card.pack(side="left", fill="both", expand=True, padx=(14, 0))

        tk.Label(card, text="[ RUNTIME PARAMS ]", bg=PANEL, fg=GREEN,
                 font=("Consolas", 8)).grid(row=0, column=0, columnspan=3,
                                            sticky="w", padx=18, pady=(14, 8))

        self.cfg_vars = {}
        fields = [("User", "username"), ("Company", "company_code"),
                  ("Bank", "bank_code"), ("Payment Date", "payment_date")]
        for i, (label, key) in enumerate(fields):
            col, r = i % 2, 1 + i // 2
            cell = tk.Frame(card, bg=PANEL)
            cell.grid(row=r, column=col, sticky="w", padx=18, pady=(0, 12))
            tk.Label(cell, text=f"::{label.upper()}::", bg=PANEL, fg=MUTED,
                     font=("Consolas", 7)).pack(anchor="w")
            var = tk.StringVar(value="—")
            self.cfg_vars[key] = var
            tk.Label(cell, textvariable=var, bg=PANEL, fg=FG,
                     font=("Consolas", 10, "bold")).pack(anchor="w")

        toggle = tk.Frame(card, bg=PANEL)
        toggle.grid(row=3, column=0, columnspan=2, sticky="w", padx=18, pady=(2, 16))
        self.headless_chk = tk.Checkbutton(
            toggle, text="[ ] HEADLESS MODE", variable=self.headless_var,
            bg=PANEL, fg=BLUE_SOFT, activebackground=PANEL, activeforeground=GREEN,
            selectcolor=PANEL, font=("Consolas", 8), bd=0, cursor="hand2",
            highlightthickness=0)
        self.headless_chk.pack(anchor="w")

    def _build_controls(self, parent):
        bar = tk.Frame(parent, bg=BG)
        bar.pack(fill="x", pady=(16, 8))

        self.run_btn = tk.Button(
            bar, text="> INITIATE SEQUENCE", command=self.start_run,
            bg=PANEL_HI, fg=GREEN, activebackground="#003322",
            activeforeground=GREEN, relief="flat", bd=0, cursor="hand2",
            font=("Consolas", 10, "bold"), padx=26, pady=10,
            highlightbackground=GREEN, highlightthickness=1)
        self.run_btn.pack(side="left")

        self.stop_btn = tk.Button(
            bar, text="> ABORT", command=self.stop_run, bg=PANEL_HI, fg=RED,
            activebackground="#330011", activeforeground=RED, relief="flat", bd=0,
            cursor="hand2", font=("Consolas", 9, "bold"), padx=18, pady=10,
            state="disabled", highlightbackground=RED, highlightthickness=1)
        self.stop_btn.pack(side="left", padx=(10, 0))

        self.reload_btn = tk.Button(
            bar, text="> REFRESH", command=lambda: self.reload_file(self.current_file),
            bg=PANEL_HI, fg=BLUE_SOFT, activebackground="#001122", activeforeground=BLUE_SOFT,
            relief="flat", bd=0, cursor="hand2", font=("Consolas", 9),
            padx=16, pady=10, highlightbackground=BORDER, highlightthickness=1)
        self.reload_btn.pack(side="left", padx=(10, 0))

        self.open_btn = tk.Button(
            bar, text="> VIEW DATA", command=self.open_excel, bg=PANEL_HI, fg=BLUE_SOFT,
            activebackground="#001122", activeforeground=BLUE_SOFT, relief="flat", bd=0,
            cursor="hand2", font=("Consolas", 9), padx=16, pady=10,
            highlightbackground=BORDER, highlightthickness=1)
        self.open_btn.pack(side="left", padx=(10, 0))

        self.progress = ProgressPulse(bar, height=2, bg=BG)
        self.progress.pack(side="bottom", fill="x", pady=(8, 0))

    def _build_table_card(self, parent):
        card = self._card(parent)
        card.pack(side="left", fill="both", expand=True)

        head = tk.Frame(card, bg=PANEL)
        head.pack(fill="x", padx=18, pady=(14, 8))
        tk.Label(head, text="[ TARGET QUEUE ]", bg=PANEL, fg=GREEN,
                 font=("Consolas", 8)).pack(side="left")
        self.rows_summary = tk.Label(head, text="", bg=PANEL, fg=MUTED,
                                     font=("Consolas", 8))
        self.rows_summary.pack(side="right")

        wrap = tk.Frame(card, bg=PANEL)
        wrap.pack(fill="both", expand=True, padx=18, pady=(0, 16))

        cols = ("row", "supplier", "vendor", "pay", "status")
        self.tree = ttk.Treeview(wrap, columns=cols, show="headings", height=7)
        cfg = {"row": ("ROW", 50), "supplier": ("SUPPLIER", 90),
               "vendor": ("VENDOR", 180), "pay": ("PAY NUMBER", 210),
               "status": ("STATUS", 100)}
        for c in cols:
            self.tree.heading(c, text=cfg[c][0])
            self.tree.column(c, width=cfg[c][1], anchor="w",
                             stretch=(c in ("vendor", "pay")))
        self.tree.pack(side="left", fill="both", expand=True)
        sb = ttk.Scrollbar(wrap, orient="vertical", command=self.tree.yview)
        sb.pack(side="right", fill="y")
        self.tree.configure(yscrollcommand=sb.set)

        self.tree.tag_configure("done", foreground=GREEN)
        self.tree.tag_configure("pending", foreground=AMBER)
        self.tree.tag_configure("active", background="#003344", foreground=GREEN)
        self.tree.tag_configure("failed", foreground=RED)

    def _build_log_card(self, parent):
        card = self._card(parent)
        card.pack(side="left", fill="both", expand=True, padx=(14, 0))

        head = tk.Frame(card, bg=PANEL)
        head.pack(fill="x", padx=18, pady=(14, 8))
        tk.Label(head, text="[ SYSTEM LOG ]", bg=PANEL, fg=GREEN,
                 font=("Consolas", 8)).pack(side="left")
        tk.Button(head, text="[CLEAR]", command=self.clear_log, bg=PANEL, fg=MUTED,
                  activebackground=PANEL, activeforeground=GREEN, bd=0,
                  relief="flat", font=("Consolas", 7), cursor="hand2").pack(side="right")

        self.log_text = tk.Text(card, bg=CONSOLE_BG, fg="#c9d1d9", relief="flat",
                                font=("Consolas", 8), wrap="word", padx=12, pady=10,
                                insertbackground=GREEN, height=10, borderwidth=1,
                                highlightbackground=BORDER, highlightthickness=1)
        self.log_text.pack(fill="both", expand=True, padx=18, pady=(0, 16))
        self.log_text.configure(state="disabled")
        self.log_text.tag_configure("step", foreground=GREEN)
        self.log_text.tag_configure("ok", foreground="#00ff9d")
        self.log_text.tag_configure("warn", foreground=AMBER)
        self.log_text.tag_configure("err", foreground=RED)
        self.log_text.tag_configure("muted", foreground=MUTED)

    def _build_footer(self):
        bar = tk.Frame(self, bg=PANEL, height=26)
        bar.pack(fill="x", side="bottom")
        bar.pack_propagate(False)
        accent = tk.Frame(self, bg=BLUE, height=1)
        accent.pack(fill="x", side="bottom")
        accent.pack_propagate(False)
        self.status = tk.Label(bar, text="  SYSTEM IDLE :: AWAITING INPUT", bg=PANEL, fg=MUTED,
                               font=("Consolas", 7), anchor="w")
        self.status.pack(fill="x", padx=12)

    # -------------------------------------------------------------- logging
    def _tag_for(self, text):
        low = text.lower()
        if text.startswith("[STEP]"):
            return "step"
        if "[failed]" in low or "fail" in low or "error" in low:
            return "err"
        if "[warn" in low or "[retry]" in low or "stop" in low:
            return "warn"
        if "extracted" in low or "successfully" in low or "paynumber" in low:
            return "ok"
        if "[skip]" in low or "[log]" in low or "[info]" in low:
            return "muted"
        return None

    def append_log(self, text, tag=None):
        self.log_text.configure(state="normal")
        self.log_text.insert("end", text + "\n", tag or self._tag_for(text) or "")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def clear_log(self):
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")

    def _queue_log(self, text):
        self.log_queue.put(text)

    def _drain_log(self):
        try:
            while True:
                self.append_log(self.log_queue.get_nowait())
        except queue.Empty:
            pass
        self.after(80, self._drain_log)

    # ------------------------------------------------------------ file input
    def upload_file(self):
        path = filedialog.askopenfilename(
            title="Select the input parameter file",
            filetypes=[("Excel workbooks", "*.xlsx *.xlsm"), ("All files", "*.*")])
        if path:
            self.reload_file(path)

    def reload_file(self, path, silent=False):
        if not path or not os.path.exists(path):
            if not silent:
                messagebox.showerror("File not found", f"Could not find:\n{path}")
            return
        try:
            cfg = sagex3.read_parameters(path)
            rows = sagex3.read_accounts(path)
        except Exception as e:
            if not silent:
                messagebox.showerror("Read error",
                                     f"Could not read this workbook:\n{e}\n\n"
                                     "Make sure it has 'SAGEX3_Credentials' and "
                                     "'Account_Details' sheets.")
            self.append_log(f"[WARN] Failed to read {os.path.basename(path)}: {e}")
            return

        self.current_file = path
        self.file_label.config(text=os.path.basename(path))
        size_kb = os.path.getsize(path) / 1024
        self.file_meta.config(text=f"{size_kb:,.1f} KB  •  {os.path.dirname(path)}")

        for key, var in self.cfg_vars.items():
            var.set(cfg.get(key, "—"))

        for item in self.tree.get_children():
            self.tree.delete(item)

        pending = 0
        for r in rows:
            if r["pay_number"]:
                status, tag = "Done", "done"
            else:
                status, tag = "Pending", "pending"
                pending += 1
            self.tree.insert("", "end", iid=str(r["row"]), tags=(tag,),
                             values=(r["row"], r["supplier"], r["vendor_name"],
                                     r["pay_number"] or "—", status))

        self.rows_summary.config(text=f"{len(rows)} row(s)  •  {pending} pending")
        self.status.config(text=f"  Loaded {os.path.basename(path)}  •  "
                                f"{len(rows)} rows, {pending} pending")
        if not silent:
            self.append_log(f"[INFO] Loaded '{os.path.basename(path)}' "
                            f"({len(rows)} rows, {pending} pending).")

    def open_excel(self):
        try:
            os.startfile(self.current_file)
        except Exception as e:
            messagebox.showerror("Open failed", str(e))

    # -------------------------------------------------------------- run flow
    def _set_state(self, text, color, spinning=False):
        self.state_label.config(text=text.upper(), fg=color)
        self.state_dot.itemconfigure(self._dot, fill=color, outline=color)
        if spinning:
            self.spinner.start()
        else:
            self.spinner.stop()

    def start_run(self):
        if self.worker and self.worker.is_alive():
            return
        if not self.current_file or not os.path.exists(self.current_file):
            messagebox.showerror("No input file", "Please upload an input file first.")
            return

        self.stop_requested.clear()
        sagex3.LOG_HOOK = self._queue_log
        sagex3.STOP_FLAG = self.stop_requested.is_set

        self.run_btn.config(state="disabled")
        self.stop_btn.config(state="normal")
        self.reload_btn.config(state="disabled")
        self.upload_btn.config(state="disabled")
        self._set_state("EXECUTING", GREEN, spinning=True)
        self.progress.start()

        mode = "HEADLESS" if self.headless_var.get() else "VISIBLE"
        self.append_log(f"\n>> INITIATED [{datetime.now():%Y-%m-%d %H:%M:%S}] [{mode}]")
        self.status.config(text=f"  EXECUTING :: BROWSER MODE::{mode}")

        self.worker = threading.Thread(target=self._run_worker, daemon=True)
        self.worker.start()
        self.after(200, self._tick_elapsed)
        self._start_time = datetime.now()

    def _tick_elapsed(self):
        if self.worker and self.worker.is_alive():
            secs = int((datetime.now() - self._start_time).total_seconds())
            self.status.config(text=f"  EXECUTING :: RUNTIME {secs // 60:02d}:{secs % 60:02d}")
            self.after(1000, self._tick_elapsed)

    def _run_worker(self):
        result = None
        try:
            result = sagex3.run(excel_path=self.current_file,
                                headless=self.headless_var.get())
        except Exception as e:
            self._queue_log(f"[FAILED] {type(e).__name__}: {e}")
            self._queue_log(traceback.format_exc())
        finally:
            self.after(0, lambda: self._on_run_finished(result))

    def _on_run_finished(self, result):
        sagex3.LOG_HOOK = None
        sagex3.STOP_FLAG = None
        self.run_btn.config(state="normal")
        self.stop_btn.config(state="disabled")
        self.reload_btn.config(state="normal")
        self.upload_btn.config(state="normal")
        self.progress.stop()

        if result is None:
            self._set_state("FAULT", RED)
        elif result.get("stopped"):
            self._set_state("ABORTED", AMBER)
        elif result.get("failed"):
            self._set_state("COMPLETE::ERR", AMBER)
        else:
            self._set_state("COMPLETE", GREEN)

        self.reload_file(self.current_file, silent=True)
        self.status.config(text=f"  SEQUENCE END :: {datetime.now():%H:%M:%S}")
        if result:
            self.append_log(f">> SUMMARY :: PROCESSED={result.get('processed', 0)} | FAILED={result.get('failed', 0)} | SKIPPED={result.get('skipped', 0)}")

    def stop_run(self):
        if self.worker and self.worker.is_alive():
            self.stop_requested.set()
            try:
                import sagex3 as sx
                if getattr(sx, "CURRENT_DRIVER", None) is not None:
                    try:
                        sx.CURRENT_DRIVER.quit()
                    except Exception:
                        pass
            except Exception:
                pass
            self._set_state("ABORTING", AMBER, spinning=True)
            self.append_log(">> ABORT SIGNAL TRANSMITTED")

    def _on_close(self):
        if self.worker and self.worker.is_alive():
            if not messagebox.askyesno("TERMINATE", "SEQUENCE IN PROGRESS - ABORT AND EXIT?"):
                return
            self.stop_requested.set()
            try:
                import sagex3 as sx
                if getattr(sx, "CURRENT_DRIVER", None) is not None:
                    try:
                        sx.CURRENT_DRIVER.quit()
                    except Exception:
                        pass
            except Exception:
                pass
        self.destroy()


if __name__ == "__main__":
    app = APP()
    app.mainloop()
