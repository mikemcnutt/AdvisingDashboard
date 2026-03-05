import json
import platform
import sys
import html
from dataclasses import dataclass
from pathlib import Path
from typing import Optional

import tkinter as tk
from tkinter import ttk, messagebox, filedialog

try:
    import win32com.client  # pywin32 (Windows only)
except Exception:
    win32com = None


# -----------------------------
# Theme
# -----------------------------
ROYAL_BLUE = "#1e3a8a"
ROYAL_BLUE_DARK = "#1e40af"
ROYAL_BLUE_LIGHT = "#3b82f6"
ROYAL_BLUE_CARD = "#e0e7ff"
BORDER_BLUE = "#93c5fd"

TEXT_LIGHT = "#ffffff"
TEXT_DARK = "#0f172a"
TEXT_MUTED = "#334155"


# -----------------------------
# Data
# -----------------------------
@dataclass
class StudentInfo:
    first_name: str
    last_name: str
    student_id: str
    kctcs_email: str
    personal_email: str
    notes: str
    json_path: str

    @property
    def display_name(self) -> str:
        name = f"{self.first_name} {self.last_name}".strip()
        return name if name else (self.kctcs_email or "(Unnamed Student)")


# -----------------------------
# Helpers
# -----------------------------
def safe_str(v) -> str:
    return "" if v is None else str(v)

def app_base_dir() -> Path:
    return Path(sys.executable).parent if getattr(sys, "frozen", False) else Path(__file__).parent

def settings_path() -> Path:
    return app_base_dir() / "settings.json"

def load_settings() -> dict:
    p = settings_path()
    if not p.exists():
        return {}
    try:
        return json.loads(p.read_text(encoding="utf-8"))
    except Exception:
        return {}

def save_settings(d: dict):
    p = settings_path()
    try:
        p.write_text(json.dumps(d, indent=2), encoding="utf-8")
    except Exception:
        pass

def load_json(path: Path) -> dict:
    with path.open("r", encoding="utf-8") as f:
        return json.load(f)

def iter_json_files(root: Path):
    for p in root.rglob("*.json"):
        yield p

def extract_student_info(obj: dict, json_path: str) -> StudentInfo:
    student = obj.get("student") if isinstance(obj.get("student"), dict) else {}
    data = obj.get("data") if isinstance(obj.get("data"), dict) else {}

    first_name = safe_str(student.get("firstName")).strip()
    last_name = safe_str(student.get("lastName")).strip()
    student_id = safe_str(student.get("studentId")).strip()
    kctcs_email = safe_str(student.get("kctcsEmail")).strip()
    personal_email = safe_str(student.get("personalEmail")).strip()
    notes = safe_str(data.get("notes")).strip()

    return StudentInfo(
        first_name=first_name,
        last_name=last_name,
        student_id=student_id,
        kctcs_email=kctcs_email,
        personal_email=personal_email,
        notes=notes,
        json_path=json_path
    )

def find_semester_plan(obj: dict, season: str, year: str) -> Optional[dict]:
    data = obj.get("data")
    if not isinstance(data, dict):
        return None
    plans = data.get("semesterPlans")
    if not isinstance(plans, list):
        return None

    for p in plans:
        if not isinstance(p, dict):
            continue
        if safe_str(p.get("season")).strip() == season and safe_str(p.get("year")).strip() == year:
            return p
    return None

def classify(obj: dict, season: str, year: str) -> str:
    """
    Returns: "needs" | "partial" | "done"
    """
    plan = find_semester_plan(obj, season, year)
    if plan is None:
        return "needs"
    if bool(plan.get("notComplete")):
        return "partial"
    return "done"


# -----------------------------
# Tooltip
# -----------------------------
class Tooltip:
    def __init__(self, parent: tk.Widget):
        self.parent = parent
        self.tip = None

    def show(self, x: int, y: int, text: str):
        self.hide()
        if not text:
            return
        self.tip = tk.Toplevel(self.parent)
        self.tip.wm_overrideredirect(True)
        self.tip.wm_geometry(f"+{x}+{y}")

        lbl = tk.Label(
            self.tip,
            text=text,
            justify="left",
            background="#0b1220",
            foreground="#e5e7eb",
            relief="solid",
            borderwidth=1,
            wraplength=420,
            padx=10,
            pady=8,
            font=("Segoe UI", 9),
        )
        lbl.pack()

    def hide(self):
        if self.tip is not None:
            try:
                self.tip.destroy()
            except Exception:
                pass
        self.tip = None


# -----------------------------
# Scrollable Frame
# -----------------------------
class ScrollableFrame(ttk.Frame):
    def __init__(self, parent, *args, **kwargs):
        super().__init__(parent, *args, **kwargs)

        self.canvas = tk.Canvas(self, highlightthickness=0, bd=0)
        self.vsb = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=self.vsb.set)

        self.inner = ttk.Frame(self.canvas)
        self.inner_id = self.canvas.create_window((0, 0), window=self.inner, anchor="nw")

        self.canvas.pack(side="left", fill="both", expand=True)
        self.vsb.pack(side="right", fill="y")

        self.inner.bind("<Configure>", self._on_inner_configure)
        self.canvas.bind("<Configure>", self._on_canvas_configure)

        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel, add=True)

    def _on_inner_configure(self, _):
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def _on_canvas_configure(self, event):
        self.canvas.itemconfigure(self.inner_id, width=event.width)

    def _on_mousewheel(self, event):
        if self.winfo_containing(event.x_root, event.y_root) in (self.canvas, self.inner):
            delta = -1 * int(event.delta / 120)
            self.canvas.yview_scroll(delta, "units")

    def clear(self):
        for child in self.inner.winfo_children():
            child.destroy()


# -----------------------------
# Outlook email (HTML)
# -----------------------------
def ensure_outlook_ready():
    if platform.system().lower() != "windows":
        raise RuntimeError("Outlook emailing is only supported on Windows.")
    if win32com is None:
        raise RuntimeError("pywin32 is not installed. Install on Windows with: pip install pywin32")

def _nl2br(text: str) -> str:
    return html.escape(text or "").replace("\r\n", "\n").replace("\r", "\n").replace("\n", "<br>")

def build_email_subject(base_subject: str, term_label: str) -> str:
    s = (base_subject or "").strip()
    if not s:
        s = "Advising Appointment Needed"
    if term_label.lower() in s.lower():
        return s
    return f"{s} — {term_label}"

def build_email_html(first_name: str, message_text: str, scheduling_link: str) -> str:
    first = (first_name or "").strip() or "there"
    msg_html = _nl2br(message_text)

    link = (scheduling_link or "").strip()
    button_block = ""
    if link:
        safe_link = html.escape(link, quote=True)
        button_block = f"""
          <div style="margin-top:18px;">
            <a href="{safe_link}"
               style="display:inline-block;background:#3b82f6;color:#ffffff;text-decoration:none;
                      padding:10px 14px;border-radius:999px;font-weight:700;font-size:14px;">
              Schedule Appointment
            </a>
          </div>
        """

    return f"""
<!doctype html>
<html>
  <body style="margin:0;padding:0;background:#f1f5f9;font-family:Segoe UI, Arial, sans-serif;">
    <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%">
      <tr>
        <td align="center" style="padding:18px;">
          <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="640"
                 style="max-width:640px;background:#ffffff;border:1px solid #dbeafe;border-radius:14px;overflow:hidden;">
            <tr>
              <td style="background:linear-gradient(90deg,#1e3a8a,#3b82f6);padding:18px 20px;">
                <div style="color:#ffffff;font-size:16px;font-weight:700;letter-spacing:.2px;">
                  Advising Appointment
                </div>
              </td>
            </tr>
            <tr>
              <td style="padding:20px;">
                <div style="color:#0f172a;font-size:15px;font-weight:700;margin-bottom:12px;">
                  Hello {html.escape(first)},
                </div>

                <div style="color:#334155;font-size:14px;line-height:1.55;">
                  {msg_html}
                </div>

                {button_block}

                <div style="margin-top:18px;color:#0f172a;font-size:14px;">
                  Thanks,<br>
                  <span style="color:#334155;">(Your Advisor)</span>
                </div>
              </td>
            </tr>
            <tr>
              <td style="padding:14px 20px;background:#f8fafc;border-top:1px solid #e2e8f0;">
                <div style="color:#64748b;font-size:12px;line-height:1.4;">
                  This email was generated from the advising dashboard.
                </div>
              </td>
            </tr>
          </table>
        </td>
      </tr>
    </table>
  </body>
</html>
""".strip()

def outlook_create_email_html(kctcs_email: str, personal_email: str, subject: str, html_body: str, draft: bool = True):
    ensure_outlook_ready()

    to_list = [e.strip() for e in [kctcs_email, personal_email] if e and str(e).strip()]
    if not to_list:
        raise RuntimeError("Student has no email addresses in JSON (KCTCS or personal).")

    outlook = win32com.client.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)
    mail.To = "; ".join(to_list)
    mail.Subject = subject
    mail.HTMLBody = html_body

    if draft:
        mail.Save()
    else:
        mail.Send()


# -----------------------------
# App
# -----------------------------
class AdvisingDashboardApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Advising Dashboard")
        self.geometry("1220x790")
        self.minsize(1100, 700)

        self.tooltip = Tooltip(self)

        self.season_var = tk.StringVar(value="Fall")
        self.year_var = tk.StringVar(value="2026")
        self.folder_var = tk.StringVar(value=str(self.default_advising_folder()))

        self.count_needs = tk.StringVar(value="Needs Advised: 0")
        self.count_partial = tk.StringVar(value="Advised Not Complete: 0")
        self.count_done = tk.StringVar(value="Advised: 0")

        self.needs_students: list[StudentInfo] = []
        self.partial_students: list[StudentInfo] = []
        self.done_students: list[StudentInfo] = []

        self.needs_checks: dict[str, tk.BooleanVar] = {}

        s = load_settings()
        self.subject_var = tk.StringVar(value=s.get("subject", "Advising Appointment Needed"))
        self.scheduling_link_var = tk.StringVar(value=s.get("schedulingLink", ""))

        self._apply_theme()
        self._build_ui()

    def default_advising_folder(self) -> Path:
        base = app_base_dir()
        return base / "Advising"

    def term_label(self) -> str:
        return f"{self.season_var.get()} {self.year_var.get()}"

    def _apply_theme(self):
        self.configure(bg=ROYAL_BLUE)
        style = ttk.Style()
        style.theme_use("clam")

        style.configure("Main.TFrame", background=ROYAL_BLUE)

        style.configure("Top.TLabelframe", background=ROYAL_BLUE, foreground=TEXT_LIGHT, bordercolor=BORDER_BLUE)
        style.configure("Top.TLabelframe.Label", background=ROYAL_BLUE, foreground=TEXT_LIGHT, font=("Segoe UI", 11, "bold"))

        style.configure("Card.TLabelframe", background=ROYAL_BLUE_CARD, foreground=TEXT_DARK, bordercolor=BORDER_BLUE)
        style.configure("Card.TLabelframe.Label", background=ROYAL_BLUE_CARD, foreground=ROYAL_BLUE_DARK, font=("Segoe UI", 11, "bold"))

        style.configure("Blue.TButton", background=ROYAL_BLUE_LIGHT, foreground="white",
                        font=("Segoe UI", 10, "bold"), padding=(12, 7), borderwidth=0, focusthickness=0)
        style.map("Blue.TButton", background=[("active", ROYAL_BLUE_DARK)])

        style.configure("Pill.TButton", background=ROYAL_BLUE_LIGHT, foreground="white",
                        font=("Segoe UI", 9, "bold"), padding=(12, 6), borderwidth=0, focusthickness=0)
        style.map("Pill.TButton", background=[("active", ROYAL_BLUE_DARK)])

        style.configure("Ghost.TButton", background=ROYAL_BLUE_CARD, foreground=ROYAL_BLUE_DARK,
                        font=("Segoe UI", 9, "bold"), padding=(10, 6), borderwidth=1)
        style.map("Ghost.TButton", background=[("active", "#c7d2fe")])

        style.configure("Summary.TLabel", background=ROYAL_BLUE, foreground=TEXT_LIGHT, font=("Segoe UI", 10, "bold"))

    def _save_settings(self):
        save_settings({
            "subject": self.subject_var.get(),
            "schedulingLink": self.scheduling_link_var.get()
        })

    def _build_ui(self):
        top = ttk.Labelframe(self, text="Controls", padding=12, style="Top.TLabelframe")
        top.pack(fill="x", padx=12, pady=(12, 10))

        ttk.Label(top, text="Semester:", foreground=TEXT_LIGHT, background=ROYAL_BLUE, font=("Segoe UI", 10, "bold")).pack(side="left")
        ttk.Combobox(top, textvariable=self.season_var, state="readonly",
                     values=["Spring", "Summer", "Fall"], width=10).pack(side="left", padx=(8, 16))

        ttk.Label(top, text="Year:", foreground=TEXT_LIGHT, background=ROYAL_BLUE, font=("Segoe UI", 10, "bold")).pack(side="left")
        ttk.Combobox(top, textvariable=self.year_var, state="readonly",
                     values=[str(y) for y in range(2026, 2041)], width=8).pack(side="left", padx=(8, 16))

        ttk.Label(top, text="Advising folder:", foreground=TEXT_LIGHT, background=ROYAL_BLUE, font=("Segoe UI", 10, "bold")).pack(side="left")
        ttk.Entry(top, textvariable=self.folder_var, width=54).pack(side="left", padx=(8, 8))

        ttk.Button(top, text="Browse…", style="Ghost.TButton", command=self.browse_folder).pack(side="left", padx=(0, 10))
        ttk.Button(top, text="Scan", style="Blue.TButton", command=self.scan).pack(side="left")

        self.status_label = ttk.Label(top, text="Ready", style="Summary.TLabel")
        self.status_label.pack(side="right")

        summary = ttk.Frame(self, padding=(12, 0), style="Main.TFrame")
        summary.pack(fill="x")
        ttk.Label(summary, textvariable=self.count_needs, style="Summary.TLabel").pack(side="left", padx=(0, 14))
        ttk.Label(summary, textvariable=self.count_partial, style="Summary.TLabel").pack(side="left", padx=(0, 14))
        ttk.Label(summary, textvariable=self.count_done, style="Summary.TLabel").pack(side="left")

        email_box = ttk.Labelframe(self, text="Email settings", padding=10, style="Card.TLabelframe")
        email_box.pack(fill="x", padx=12, pady=(10, 10))

        row1 = ttk.Frame(email_box)
        row1.pack(fill="x", pady=(0, 8))

        ttk.Label(row1, text="Subject:", background=ROYAL_BLUE_CARD, foreground=TEXT_DARK,
                  font=("Segoe UI", 10, "bold")).pack(side="left")
        subj_entry = ttk.Entry(row1, textvariable=self.subject_var)
        subj_entry.pack(side="left", fill="x", expand=True, padx=(8, 12))

        ttk.Label(row1, text="Scheduling link:", background=ROYAL_BLUE_CARD, foreground=TEXT_DARK,
                  font=("Segoe UI", 10, "bold")).pack(side="left")
        link_entry = ttk.Entry(row1, textvariable=self.scheduling_link_var, width=40)
        link_entry.pack(side="left", padx=(8, 0))

        subj_entry.bind("<FocusOut>", lambda _e: self._save_settings())
        link_entry.bind("<FocusOut>", lambda _e: self._save_settings())

        ttk.Label(email_box, text="Message:", background=ROYAL_BLUE_CARD, foreground=TEXT_DARK,
                  font=("Segoe UI", 10, "bold")).pack(anchor="w")

        self.email_body = tk.Text(email_box, height=4, wrap="word", bd=1, relief="solid", highlightthickness=0)
        self.email_body.pack(fill="x", expand=True)
        self.email_body.insert("1.0", "Please reply to schedule an advising appointment for the selected semester.")
        self.email_body.bind("<FocusOut>", lambda _e: self._save_settings())

        main = ttk.Frame(self, padding=(12, 0, 12, 12), style="Main.TFrame")
        main.pack(fill="both", expand=True)

        main.columnconfigure(0, weight=1)
        main.columnconfigure(1, weight=1)
        main.columnconfigure(2, weight=1)
        main.rowconfigure(0, weight=1)

        self.frame_needs = ttk.Labelframe(main, text="Needs advised (0)", padding=10, style="Card.TLabelframe")
        self.frame_partial = ttk.Labelframe(main, text="Advised (not complete) (0)", padding=10, style="Card.TLabelframe")
        self.frame_done = ttk.Labelframe(main, text="Advised (0)", padding=10, style="Card.TLabelframe")

        self.frame_needs.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        self.frame_partial.grid(row=0, column=1, sticky="nsew", padx=8)
        self.frame_done.grid(row=0, column=2, sticky="nsew", padx=(8, 0))

        needs_controls = ttk.Frame(self.frame_needs)
        needs_controls.pack(fill="x", pady=(0, 8))

        ttk.Button(needs_controls, text="Select all", style="Ghost.TButton", command=self.needs_select_all).pack(side="left")
        ttk.Button(needs_controls, text="Select none", style="Ghost.TButton", command=self.needs_select_none).pack(side="left", padx=(8, 0))

        ttk.Button(needs_controls, text="Draft selected", style="Blue.TButton",
                   command=lambda: self.email_selected_needs(draft=True)).pack(side="right")
        ttk.Button(needs_controls, text="Send selected", style="Ghost.TButton",
                   command=lambda: self.email_selected_needs(draft=False)).pack(side="right", padx=(0, 8))

        self.needs_list = ScrollableFrame(self.frame_needs)
        self.needs_list.pack(fill="both", expand=True)

        self.partial_list = ScrollableFrame(self.frame_partial)
        self.partial_list.pack(fill="both", expand=True)

        self.done_list = ScrollableFrame(self.frame_done)
        self.done_list.pack(fill="both", expand=True)

    def browse_folder(self):
        chosen = filedialog.askdirectory(title="Select Advising folder")
        if chosen:
            self.folder_var.set(chosen)

    def set_status(self, text: str):
        self.status_label.config(text=text)
        self.update_idletasks()

    def needs_select_all(self):
        for var in self.needs_checks.values():
            var.set(True)

    def needs_select_none(self):
        for var in self.needs_checks.values():
            var.set(False)

    def _current_subject(self) -> str:
        return build_email_subject(self.subject_var.get(), self.term_label())

    def _current_message_text(self) -> str:
        return self.email_body.get("1.0", "end").strip()

    def _current_link(self) -> str:
        return self.scheduling_link_var.get().strip()

    def email_selected_needs(self, draft: bool):
        selected = []
        for s in self.needs_students:
            var = self.needs_checks.get(s.json_path)
            if var and var.get():
                selected.append(s)

        if not selected:
            messagebox.showinfo("No selection", "Select at least one student to email.")
            return

        try:
            ensure_outlook_ready()
        except Exception as e:
            messagebox.showerror("Email unavailable", str(e))
            return

        if not draft:
            confirm = messagebox.askyesno("Confirm send", f"Send {len(selected)} email(s) now?")
            if not confirm:
                return

        subject = self._current_subject()
        message_text = self._current_message_text()
        link = self._current_link()

        ok = 0
        err = 0

        for s in selected:
            try:
                html_body = build_email_html(s.first_name, message_text, link)
                outlook_create_email_html(s.kctcs_email, s.personal_email, subject, html_body, draft=draft)
                ok += 1
            except Exception:
                err += 1

        mode = "Drafted" if draft else "Sent"
        messagebox.showinfo("Email complete", f"{mode}: {ok}\nErrors: {err}")

    def email_one_partial(self, s: StudentInfo):
        try:
            ensure_outlook_ready()
        except Exception as e:
            messagebox.showerror("Email unavailable", str(e))
            return

        subject = self._current_subject()
        message_text = self._current_message_text()
        link = self._current_link()

        try:
            html_body = build_email_html(s.first_name, message_text, link)
            outlook_create_email_html(s.kctcs_email, s.personal_email, subject, html_body, draft=True)
            messagebox.showinfo("Draft created", f"Draft email created for {s.display_name}.")
        except Exception as e:
            messagebox.showerror("Email failed", str(e))

    def _render_needs(self):
        self.needs_list.clear()
        self.needs_checks.clear()

        for s in self.needs_students:
            holder = tk.Frame(self.needs_list.inner, bg=ROYAL_BLUE_CARD, highlightbackground=BORDER_BLUE, highlightthickness=1)
            holder.pack(fill="x", pady=4)

            row = ttk.Frame(holder)
            row.pack(fill="x", padx=8, pady=6)

            var = tk.BooleanVar(value=True)
            self.needs_checks[s.json_path] = var

            tk.Checkbutton(row, variable=var, bg=ROYAL_BLUE_CARD, activebackground=ROYAL_BLUE_CARD,
                          highlightthickness=0).pack(side="left", padx=(0, 10))

            ttk.Label(row, text=s.display_name, background=ROYAL_BLUE_CARD, foreground=TEXT_DARK,
                      font=("Segoe UI", 10, "bold")).pack(side="left", padx=(0, 10))
            ttk.Label(row, text=s.student_id, background=ROYAL_BLUE_CARD, foreground=TEXT_MUTED,
                      font=("Segoe UI", 9)).pack(side="left")

    def _render_partial(self):
        self.partial_list.clear()

        for s in self.partial_students:
            holder = tk.Frame(self.partial_list.inner, bg=ROYAL_BLUE_CARD, highlightbackground=BORDER_BLUE, highlightthickness=1)
            holder.pack(fill="x", pady=4)

            row = ttk.Frame(holder)
            row.pack(fill="x", padx=8, pady=8)

            left = ttk.Frame(row)
            left.pack(side="left", fill="x", expand=True)

            ttk.Label(left, text=s.display_name, background=ROYAL_BLUE_CARD, foreground=TEXT_DARK,
                      font=("Segoe UI", 10, "bold")).pack(anchor="w")
            ttk.Label(left, text=s.student_id, background=ROYAL_BLUE_CARD, foreground=TEXT_MUTED,
                      font=("Segoe UI", 9)).pack(anchor="w")

            right = ttk.Frame(row)
            right.pack(side="right")

            if s.notes.strip():
                notes_lbl = tk.Label(right, text="Notes", bg="#dbeafe", fg=ROYAL_BLUE_DARK,
                                     padx=10, pady=4, font=("Segoe UI", 9, "bold"), cursor="question_arrow")
                notes_lbl.pack(side="left", padx=(0, 8))

                notes_lbl.bind("<Enter>", lambda _e, n=s.notes: self.tooltip.show(self.winfo_pointerx()+12, self.winfo_pointery()+12, n))
                notes_lbl.bind("<Leave>", lambda _e: self.tooltip.hide())

            ttk.Button(right, text="Email", style="Pill.TButton",
                       command=lambda stu=s: self.email_one_partial(stu)).pack(side="left")

    def _render_done(self):
        self.done_list.clear()

        for s in self.done_students:
            holder = tk.Frame(self.done_list.inner, bg=ROYAL_BLUE_CARD, highlightbackground=BORDER_BLUE, highlightthickness=1)
            holder.pack(fill="x", pady=4)

            row = ttk.Frame(holder)
            row.pack(fill="x", padx=8, pady=8)

            left = ttk.Frame(row)
            left.pack(side="left", fill="x", expand=True)

            ttk.Label(left, text=s.display_name, background=ROYAL_BLUE_CARD, foreground=TEXT_DARK,
                      font=("Segoe UI", 10, "bold")).pack(anchor="w")
            ttk.Label(left, text=s.student_id, background=ROYAL_BLUE_CARD, foreground=TEXT_MUTED,
                      font=("Segoe UI", 9)).pack(anchor="w")

    def scan(self):
        folder = Path(self.folder_var.get()).expanduser()
        if not folder.exists() or not folder.is_dir():
            messagebox.showerror("Folder not found", f"Advising folder does not exist:\n{folder}")
            return

        season = self.season_var.get()
        year = self.year_var.get()
        term = self.term_label()
        self.set_status(f"Scanning for {term}…")

        needs: list[StudentInfo] = []
        partial: list[StudentInfo] = []
        done: list[StudentInfo] = []

        files = list(iter_json_files(folder))
        bad_files = 0

        for p in files:
            try:
                obj = load_json(p)
                bucket = classify(obj, season, year)
                info = extract_student_info(obj, str(p))

                if bucket == "needs":
                    needs.append(info)
                elif bucket == "partial":
                    partial.append(info)
                else:
                    done.append(info)
            except Exception:
                bad_files += 1
                continue

        needs.sort(key=lambda s: s.display_name.lower())
        partial.sort(key=lambda s: s.display_name.lower())
        done.sort(key=lambda s: s.display_name.lower())

        self.needs_students = needs
        self.partial_students = partial
        self.done_students = done

        self._render_needs()
        self._render_partial()
        self._render_done()

        n_needs = len(needs)
        n_partial = len(partial)
        n_done = len(done)

        self.count_needs.set(f"Needs Advised: {n_needs}")
        self.count_partial.set(f"Advised Not Complete: {n_partial}")
        self.count_done.set(f"Advised: {n_done}")

        self.frame_needs.config(text=f"Needs advised ({n_needs})")
        self.frame_partial.config(text=f"Advised (not complete) ({n_partial})")
        self.frame_done.config(text=f"Advised ({n_done})")

        msg = f"{len(files)} file(s) scanned • {n_needs} need • {n_partial} partial • {n_done} advised"
        if bad_files:
            msg += f" • {bad_files} unreadable"
        self.set_status(msg)


def main():
    app = AdvisingDashboardApp()
    app.mainloop()

if __name__ == "__main__":
    main()
