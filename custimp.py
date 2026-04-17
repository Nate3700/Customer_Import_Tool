import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import os
import sys
import json
import datetime
import threading
import urllib.request
import shutil
import subprocess
import tempfile

# ── Version ───────────────────────────────────────────────────────────────────
APP_VERSION = "1.0.5"

# ── UPDATE THESE TWO LINES with your GitHub username and repo name ────────────
GITHUB_USER = "Nate3700"
GITHUB_REPO = "Customer_Import_Tool"
# ─────────────────────────────────────────────────────────────────────────────

VERSION_URL  = f"https://raw.githubusercontent.com/{GITHUB_USER}/{GITHUB_REPO}/main/version.json"
DOWNLOAD_URL = f"https://github.com/{GITHUB_USER}/{GITHUB_REPO}/releases/latest/download/CustImpTool.exe"

# ── Auto-updater ──────────────────────────────────────────────────────────────
def check_for_updates(app):
    """Run in a background thread. Silently skips if no internet."""
    try:
        with urllib.request.urlopen(VERSION_URL, timeout=5) as resp:
            data = json.loads(resp.read().decode())
        latest = data.get("version", "0.0.0")
        if _version_tuple(latest) > _version_tuple(APP_VERSION):
            # Switch back to main thread to show dialog
            app.after(0, lambda: _prompt_update(app, latest))
    except Exception:
        pass  # No internet or file not found — silently continue

def _version_tuple(v):
    return tuple(int(x) for x in v.strip().split("."))

def _prompt_update(app, latest_version):
    answer = messagebox.askyesno(
        "Update Available",
        f"A new version of custimp is available.\n\n"
        f"  Current version:  {APP_VERSION}\n"
        f"  New version:      {latest_version}\n\n"
        "Would you like to download and install the update now?\n"
        "(The app will restart automatically.)"
    )
    if answer:
        _download_and_install(app, latest_version)

def _download_and_install(app, latest_version):
    """Download new .exe to temp, write a batch script to replace and relaunch."""
    progress = tk.Toplevel(app)
    progress.title("Downloading Update...")
    progress.geometry("360x110")
    progress.resizable(False, False)
    progress.configure(bg='#F7F8FA')
    progress.grab_set()

    tk.Label(progress, text=f"Downloading version {latest_version}...",
             font=('Segoe UI', 11), bg='#F7F8FA', fg='#1C2331').pack(pady=(20, 8))
    bar = ttk.Progressbar(progress, mode='indeterminate', length=300)
    bar.pack()
    bar.start(12)

    def do_download():
        try:
            # Download new exe to a temp file
            tmp_fd, tmp_path = tempfile.mkstemp(suffix='.exe')
            os.close(tmp_fd)
            urllib.request.urlretrieve(DOWNLOAD_URL, tmp_path)

            current_exe = os.path.abspath(sys.executable
                                          if getattr(sys, 'frozen', False)
                                          else __file__)

            # Write a small batch script that:
            # 1. Waits for this process to exit
            # 2. Replaces the old .exe with the new one
            # 3. Relaunches the app
            # 4. Deletes itself
            batch = tempfile.NamedTemporaryFile(
                suffix='.bat', delete=False, mode='w')
            batch.write(
                f'@echo off\n'
                f'ping 127.0.0.1 -n 3 > nul\n'          # wait ~2 sec
                f'move /Y "{tmp_path}" "{current_exe}"\n'
                f'start "" "{current_exe}"\n'
                f'del "%~f0"\n'
            )
            batch.close()

            progress.after(0, progress.destroy)
            subprocess.Popen(['cmd', '/c', batch.name],
                             creationflags=subprocess.CREATE_NO_WINDOW
                             if sys.platform == 'win32' else 0)
            app.after(200, app.destroy)

        except Exception as e:
            progress.after(0, progress.destroy)
            app.after(0, lambda: messagebox.showerror(
                "Update Failed",
                f"Could not download the update:\n\n{e}\n\n"
                "Please try again later or download manually from GitHub."))

    threading.Thread(target=do_download, daemon=True).start()

# ── Config file ───────────────────────────────────────────────────────────────
def get_config_path():
    base = os.path.dirname(os.path.abspath(
        sys.executable if getattr(sys, 'frozen', False) else __file__))
    return os.path.join(base, 'custimp_config.json')

def load_config():
    path = get_config_path()
    if os.path.exists(path):
        try:
            with open(path, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception:
            pass
    return {}

def save_config(config):
    try:
        with open(get_config_path(), 'w', encoding='utf-8') as f:
            json.dump(config, f, indent=2)
    except Exception as e:
        messagebox.showwarning('Config Warning',
                               f'Could not save settings:\n{e}')

# ── UPC-A Check Digit ─────────────────────────────────────────────────────────
def upca_check_digit(base11):
    d = str(base11).zfill(11)
    total = sum(int(d[i]) * (3 if i % 2 == 0 else 1) for i in range(11))
    return (10 - (total % 10)) % 10

def to_upca(id11):
    base = str(id11).zfill(11)
    return base + str(upca_check_digit(base))

def normalize_id(raw):
    cleaned = ''.join(c for c in raw.strip() if c.isdigit())
    if len(cleaned) == 12:
        return cleaned[:11]
    return cleaned

# ── CSV Builder ───────────────────────────────────────────────────────────────
def build_csv(rows):
    TOTAL_COLS = 37
    lines = []
    for r in rows:
        cols = [''] * TOTAL_COLS
        try:
            bal = float(r['balance']) if r['balance'] else 0.0
        except ValueError:
            bal = 0.0
        bal_str = f"{-abs(bal):.2f}"
        cols[0]  = r['id']
        cols[2]  = f'"{r["first_name"]}"'
        cols[3]  = f'"{r["last_name"]}"'
        cols[32] = f'"{r["company"]}"'
        cols[33] = '1'
        cols[34] = '"Net 30"'
        cols[35] = '0'
        cols[36] = bal_str
        lines.append(','.join(cols))
    return '\r\n'.join(lines)

# ── Archive helper ────────────────────────────────────────────────────────────
def archive_existing(folder):
    target = os.path.join(folder, 'custimp.txt')
    if not os.path.exists(target):
        return True, None
    try:
        mtime        = os.path.getmtime(target)
        dt           = datetime.datetime.fromtimestamp(mtime)
        stamp        = dt.strftime('%Y-%m-%d_%H-%M')
        archive_name = f'custimp_{stamp}.txt'
        archive_path = os.path.join(folder, archive_name)
        if os.path.exists(archive_path):
            stamp        = dt.strftime('%Y-%m-%d_%H-%M-%S')
            archive_name = f'custimp_{stamp}.txt'
            archive_path = os.path.join(folder, archive_name)
        os.rename(target, archive_path)
        return True, archive_path
    except Exception as e:
        return False, str(e)

# ── Main App ──────────────────────────────────────────────────────────────────
class CustImpApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(f"Customer Import Tool — custimp.txt  v{APP_VERSION}")
        self.geometry("900x700")
        self.minsize(920, 740)
        self.resizable(True, True)

        self.rows      = []
        self.form_data = {}
        self.config    = load_config()

        # ── Design tokens ─────────────────────────────────────────────
        self.C_BG       = '#F7F8FA'   # page background
        self.C_CARD     = '#FFFFFF'   # card / panel background
        self.C_BORDER   = '#E2E6EA'   # borders
        self.C_TEXT     = '#1C2331'   # primary text
        self.C_MUTED    = '#6B7685'   # secondary / label text
        self.C_ACCENT   = '#2563EB'   # blue primary
        self.C_ACCENT_H = '#1D4ED8'   # blue hover
        self.C_SUCCESS  = '#16A34A'   # green done
        self.C_DANGER   = '#DC2626'   # red danger
        self.C_DANGER_H = '#B91C1C'
        self.C_COUNT_BG = '#EFF6FF'   # card count background
        self.C_COUNT_BD = '#BFDBFE'   # card count border
        self.C_ROW_ODD  = '#F9FAFB'
        self.C_ROW_EVEN = '#FFFFFF'
        self.C_ROW_SEL  = '#DBEAFE'

        self.configure(bg=self.C_BG)

        self.font_label  = ('Segoe UI', 12, 'bold')
        self.font_input  = ('Segoe UI', 13)
        self.font_mono   = ('Consolas', 11)
        self.font_btn    = ('Segoe UI', 12, 'bold')
        self.font_title  = ('Segoe UI', 26, 'bold')
        self.font_sub    = ('Segoe UI', 11)
        self.font_table  = ('Segoe UI', 11)
        self.font_head   = ('Segoe UI', 10, 'bold')

        self.step1_frame = None
        self.step2_frame = None
        self._build_header()
        self._build_step1()
        self._build_step2()
        self._show_step(1)

        # Check for updates in background after window is ready
        self.after(1000, lambda: threading.Thread(
            target=check_for_updates, args=(self,), daemon=True).start())

    # ── Header ────────────────────────────────────────────────────────────────
    def _build_header(self):
        hdr = tk.Frame(self, bg=self.C_BG)
        hdr.pack(fill='x', padx=48, pady=(28, 8))

        # Top row: title left, settings + version right
        top = tk.Frame(hdr, bg=self.C_BG)
        top.pack(fill='x')

        title_frame = tk.Frame(top, bg=self.C_BG)
        title_frame.pack(side='left')
        tk.Label(title_frame, text='custimp', font=self.font_title,
                 fg=self.C_TEXT, bg=self.C_BG).pack(side='left')
        tk.Label(title_frame, text='.txt', font=self.font_title,
                 fg=self.C_ACCENT, bg=self.C_BG).pack(side='left')

        right_frame = tk.Frame(top, bg=self.C_BG)
        right_frame.pack(side='right', anchor='s', pady=(0, 4))
        tk.Label(right_frame, text=f'v{APP_VERSION}', font=('Segoe UI', 9),
                 fg=self.C_MUTED, bg=self.C_BG).pack(side='left', padx=(0, 12))
        tk.Button(right_frame, text='⚙  Settings', font=('Segoe UI', 9, 'bold'),
                  bg=self.C_BG, fg=self.C_MUTED, relief='flat', cursor='hand2',
                  activebackground=self.C_BORDER, activeforeground=self.C_TEXT,
                  padx=10, pady=4,
                  command=self._open_settings).pack(side='left')

        # Divider
        tk.Frame(hdr, bg=self.C_BORDER, height=1).pack(fill='x', pady=(10, 12))

        # Step indicator
        self.step_frame = tk.Frame(hdr, bg=self.C_BG)
        self.step_frame.pack(anchor='w')

        self.s1_circle = tk.Label(self.step_frame, text='1', width=2,
                                   font=('Segoe UI', 10, 'bold'), relief='flat',
                                   padx=2, pady=2)
        self.s1_circle.pack(side='left', padx=(0, 6))
        self.s1_label = tk.Label(self.step_frame, text='Enter Details',
                                  font=('Segoe UI', 10, 'bold'), bg=self.C_BG)
        self.s1_label.pack(side='left')

        tk.Label(self.step_frame, text='  ──────  ', fg=self.C_BORDER,
                 bg=self.C_BG, font=('Segoe UI', 10)).pack(side='left')

        self.s2_circle = tk.Label(self.step_frame, text='2', width=2,
                                   font=('Segoe UI', 10, 'bold'), relief='flat',
                                   padx=2, pady=2)
        self.s2_circle.pack(side='left', padx=(0, 6))
        self.s2_label = tk.Label(self.step_frame, text='Fill Names & Export',
                                  font=('Segoe UI', 10, 'bold'), bg=self.C_BG)
        self.s2_label.pack(side='left')

        self._update_step_indicators(1)

    def _open_settings(self):
        win = tk.Toplevel(self)
        win.title('Settings')
        win.geometry('480x180')
        win.resizable(False, False)
        win.configure(bg=self.C_BG)
        win.grab_set()

        tk.Label(win, text='Save Folder', font=self.font_label,
                 fg=self.C_TEXT, bg=self.C_BG).pack(anchor='w', padx=24, pady=(20, 4))

        folder_inner = tk.Frame(win, bg=self.C_BG)
        folder_inner.pack(fill='x', padx=24)

        saved   = self.config.get('save_folder', '')
        display = saved if saved else 'No folder set — will prompt on first save'
        self.folder_var = tk.StringVar(value=display)

        tk.Label(folder_inner, textvariable=self.folder_var,
                 font=('Segoe UI', 9), fg=self.C_MUTED, bg='#F9FAFB',
                 anchor='w', relief='flat', padx=10, pady=6,
                 highlightthickness=1,
                 highlightbackground=self.C_BORDER).pack(side='left', fill='x', expand=True)

        tk.Button(folder_inner, text='Change Folder', font=('Segoe UI', 9, 'bold'),
                  bg=self.C_ACCENT, fg='white', relief='flat', cursor='hand2',
                  activebackground=self.C_ACCENT_H, activeforeground='white',
                  padx=12, pady=6,
                  command=lambda: self._change_folder(win)).pack(side='left', padx=(8, 0))

        tk.Label(win, text='The selected folder is where custimp.txt will be saved each time.',
                 font=('Segoe UI', 9), fg=self.C_MUTED, bg=self.C_BG,
                 anchor='w').pack(anchor='w', padx=24, pady=(8, 0))

        tk.Button(win, text='Close', font=('Segoe UI', 10, 'bold'),
                  bg=self.C_BORDER, fg=self.C_TEXT, relief='flat', cursor='hand2',
                  activebackground=self.C_MUTED, activeforeground='white',
                  padx=16, pady=7,
                  command=win.destroy).pack(pady=(20, 0))

    def _update_step_indicators(self, step):
        if step == 1:
            self.s1_circle.config(bg=self.C_ACCENT, fg='white', text='1')
            self.s1_label.config(fg=self.C_TEXT)
            self.s2_circle.config(bg=self.C_BORDER, fg=self.C_MUTED, text='2')
            self.s2_label.config(fg=self.C_MUTED)
        else:
            self.s1_circle.config(bg=self.C_SUCCESS, fg='white', text='✓')
            self.s1_label.config(fg=self.C_TEXT)
            self.s2_circle.config(bg=self.C_ACCENT, fg='white', text='2')
            self.s2_label.config(fg=self.C_TEXT)

    # ── Step 1 ────────────────────────────────────────────────────────────────
    def _build_step1(self):
        outer = tk.Frame(self, bg=self.C_BG)
        self.step1_frame = outer

        # Scrollable centred column
        content = tk.Frame(outer, bg=self.C_BG)
        content.pack(padx=48, pady=20, fill='x')

        def make_box(parent, label_text, var, placeholder='', optional_text=''):
            """Each field in its own card-like box."""
            box = tk.Frame(parent, bg=self.C_CARD,
                           highlightbackground=self.C_BORDER, highlightthickness=1)
            box.pack(fill='x', pady=(0, 16), ipady=4)

            # Label row
            lbl_row = tk.Frame(box, bg=self.C_CARD)
            lbl_row.pack(fill='x', padx=14, pady=(14, 4))
            tk.Label(lbl_row, text=label_text, font=self.font_label,
                     fg=self.C_TEXT, bg=self.C_CARD).pack(side='left')
            if optional_text:
                tk.Label(lbl_row, text=optional_text, font=self.font_sub,
                         fg=self.C_MUTED, bg=self.C_CARD).pack(side='left', padx=(6, 0))

            # Entry
            entry = tk.Entry(box, textvariable=var, font=self.font_input,
                             bg=self.C_CARD, fg=self.C_TEXT, relief='flat',
                             highlightthickness=0, insertbackground=self.C_ACCENT)
            entry.pack(fill='x', padx=14, pady=(0, 4), ipady=10)

            # Bottom border acting as focus indicator
            border = tk.Frame(box, bg=self.C_BORDER, height=2)
            border.pack(fill='x', padx=14)

            if placeholder:
                entry.insert(0, placeholder)
                entry.config(fg=self.C_MUTED)
                def on_focus_in(e, en=entry, bd=border, ph=placeholder):
                    if en.get() == ph:
                        en.delete(0, 'end')
                        en.config(fg=self.C_TEXT)
                    bd.config(bg=self.C_ACCENT)
                def on_focus_out(e, en=entry, bd=border, ph=placeholder):
                    bd.config(bg=self.C_BORDER)
                    if not en.get():
                        en.insert(0, ph)
                        en.config(fg=self.C_MUTED)
                entry.bind('<FocusIn>',  on_focus_in)
                entry.bind('<FocusOut>', on_focus_out)
            else:
                def on_focus_in_plain(e, bd=border):
                    bd.config(bg=self.C_ACCENT)
                def on_focus_out_plain(e, bd=border):
                    bd.config(bg=self.C_BORDER)
                entry.bind('<FocusIn>',  on_focus_in_plain)
                entry.bind('<FocusOut>', on_focus_out_plain)

            err_lbl = tk.Label(box, text='', font=('Segoe UI', 9, 'bold'),
                               fg=self.C_DANGER, bg=self.C_CARD, anchor='w')
            err_lbl.pack(fill='x', padx=14, pady=(2, 10))
            return entry, err_lbl

        self.v_first_id = tk.StringVar()
        self.v_last_id  = tk.StringVar()
        self.v_company  = tk.StringVar()
        self.v_balance  = tk.StringVar()

        # ── Row 1: First ID and Last ID side by side ──────────────────────
        id_row = tk.Frame(content, bg=self.C_BG)
        id_row.pack(fill='x', pady=(0, 0))
        id_row.columnconfigure(0, weight=1, uniform='id')
        id_row.columnconfigure(1, weight=1, uniform='id')

        left  = tk.Frame(id_row, bg=self.C_BG)
        right = tk.Frame(id_row, bg=self.C_BG)
        left.grid(row=0, column=0, sticky='ew', padx=(0, 6))
        right.grid(row=0, column=1, sticky='ew', padx=(6, 0))

        self.e_first_id, self.err_first_id = make_box(
            left,  'First ID Card Number', self.v_first_id, 'Scan or type 11-digit ID')
        self.e_last_id, self.err_last_id = make_box(
            right, 'Last ID Card Number',  self.v_last_id,  'Scan or type 11-digit ID')

        # ── Row 2: Company Name and Balance side by side ──────────────────
        opt_row = tk.Frame(content, bg=self.C_BG)
        opt_row.pack(fill='x')
        opt_row.columnconfigure(0, weight=2, uniform='opt')
        opt_row.columnconfigure(1, weight=1, uniform='opt')

        co_frame  = tk.Frame(opt_row, bg=self.C_BG)
        bal_frame = tk.Frame(opt_row, bg=self.C_BG)
        co_frame.grid(row=0, column=0, sticky='ew', padx=(0, 6))
        bal_frame.grid(row=0, column=1, sticky='ew', padx=(6, 0))

        self.e_company, self.err_company = make_box(
            co_frame,  'Company Name', self.v_company, 'e.g. Acme Corp', '(optional)')
        self.e_balance, self.err_balance = make_box(
            bal_frame, 'Balance', self.v_balance, 'e.g. 100.00', '(optional)')

        # ── Row 4: IDs to be Generated (left) + Generate button (right) ──
        bottom_row = tk.Frame(content, bg=self.C_BG)
        bottom_row.pack(fill='x', pady=(8, 0))
        bottom_row.columnconfigure(0, weight=1, uniform='bot')
        bottom_row.columnconfigure(1, weight=1, uniform='bot')

        count_frame = tk.Frame(bottom_row, bg=self.C_BG)
        count_frame.grid(row=0, column=0, sticky='ew', padx=(0, 6))

        count_box = tk.Frame(count_frame, bg=self.C_COUNT_BG,
                             highlightbackground=self.C_COUNT_BD, highlightthickness=1)
        count_box.pack(fill='both', expand=True)
        tk.Label(count_box, text='IDs to be Generated',
                 font=('Segoe UI', 12, 'bold'), fg=self.C_MUTED,
                 bg=self.C_COUNT_BG).pack(side='left', padx=16, pady=16)
        self.card_count_var = tk.StringVar(value='—')
        self.card_count_lbl = tk.Label(count_box, textvariable=self.card_count_var,
                                        font=('Segoe UI', 24, 'bold'),
                                        fg=self.C_ACCENT, bg=self.C_COUNT_BG)
        self.card_count_lbl.pack(side='right', padx=16, pady=16)

        btn_frame = tk.Frame(bottom_row, bg=self.C_BG)
        btn_frame.grid(row=0, column=1, sticky='ew', padx=(6, 0))
        tk.Button(btn_frame, text='Generate Spreadsheet  →',
                  font=('Segoe UI', 11, 'bold'),
                  bg=self.C_ACCENT, fg='white', relief='flat', cursor='hand2',
                  activebackground=self.C_ACCENT_H, activeforeground='white',
                  padx=20, pady=16,
                  command=self._handle_generate).pack(fill='both', expand=True)

        # Initialise folder_var so settings dialog works before it's opened
        saved = self.config.get('save_folder', '')
        self.folder_var = tk.StringVar(
            value=saved if saved else 'No folder set — will prompt on first save')

        self.v_first_id.trace_add('write', lambda *a: self._live_validate())
        self.v_last_id.trace_add('write',  lambda *a: self._live_validate())
        for e in [self.e_first_id, self.e_last_id, self.e_company, self.e_balance]:
            e.bind('<Return>', lambda ev: self._handle_generate())

    def _change_folder(self, parent=None):
        folder = filedialog.askdirectory(title='Select Save Folder')
        if folder:
            self.config['save_folder'] = folder
            save_config(self.config)
            self.folder_var.set(folder)

    def _live_validate(self):
        first = normalize_id(
            self.e_first_id.get() if self.e_first_id.get() != 'Scan or type 11-digit ID' else '')
        last  = normalize_id(
            self.e_last_id.get()  if self.e_last_id.get()  != 'Scan or type 11-digit ID' else '')

        if first:
            if not first.isdigit() or len(first) != 11:
                self.err_first_id.config(
                    text='Must be exactly 11 digits (or scan 12-digit barcode)')
                self.e_first_id.config(highlightbackground=self.C_DANGER)
            else:
                self.err_first_id.config(text='')
                self.e_first_id.config(highlightbackground=self.C_BORDER)
        else:
            self.err_first_id.config(text='')
            self.e_first_id.config(highlightbackground=self.C_BORDER)

        if last:
            if not last.isdigit() or len(last) != 11:
                self.err_last_id.config(
                    text='Must be exactly 11 digits (or scan 12-digit barcode)')
                self.e_last_id.config(highlightbackground=self.C_DANGER)
            elif first and len(first) == 11 and int(last) <= int(first):
                self.err_last_id.config(text='Last ID must be greater than First ID')
                self.e_last_id.config(highlightbackground=self.C_DANGER)
            elif first and len(first) == 11 and int(last) - int(first) > 999:
                self.err_last_id.config(text='Max 1000 rows at once')
                self.e_last_id.config(highlightbackground=self.C_DANGER)
            else:
                self.err_last_id.config(text='')
                self.e_last_id.config(highlightbackground=self.C_BORDER)
        else:
            self.err_last_id.config(text='')
            self.e_last_id.config(highlightbackground=self.C_BORDER)

        if (len(first) == 11 and len(last) == 11 and
                first.isdigit() and last.isdigit() and int(last) > int(first)):
            self.card_count_var.set(f'{int(last) - int(first) + 1:,}')
            self.card_count_lbl.config(fg=self.C_ACCENT)
        else:
            self.card_count_var.set('—')
            self.card_count_lbl.config(
                fg=self.C_DANGER if (first or last) else self.C_ACCENT)

    def _validate_full(self):
        first_raw = self.e_first_id.get()
        last_raw  = self.e_last_id.get()
        first = normalize_id(
            first_raw if first_raw != 'Scan or type 11-digit ID' else '')
        last  = normalize_id(
            last_raw  if last_raw  != 'Scan or type 11-digit ID' else '')
        company = self.e_company.get().strip()
        balance = self.e_balance.get().strip()
        if company == 'e.g. Acme Corp': company = ''
        if balance == 'e.g. 100.00':   balance = ''

        ok = True
        if not first or not first.isdigit() or len(first) != 11:
            self.err_first_id.config(
                text='Must be exactly 11 digits (or scan 12-digit barcode)')
            ok = False
        else:
            self.err_first_id.config(text='')
        if not last or not last.isdigit() or len(last) != 11:
            self.err_last_id.config(
                text='Must be exactly 11 digits (or scan 12-digit barcode)')
            ok = False
        elif int(last) <= int(first):
            self.err_last_id.config(text='Last ID must be greater than First ID')
            ok = False
        elif int(last) - int(first) > 999:
            self.err_last_id.config(text='Max 1000 rows at once')
            ok = False
        else:
            self.err_last_id.config(text='')
        if balance:
            try:
                float(balance)
                self.err_balance.config(text='')
            except ValueError:
                self.err_balance.config(text='Enter a valid number')
                ok = False
        if ok:
            self.form_data = {
                'first_id': first, 'last_id': last,
                'company':  company, 'balance': balance or '0',
            }
        return ok

    def _handle_generate(self):
        if not self._validate_full():
            return
        self.rows = []
        for id_num in range(int(self.form_data['first_id']),
                            int(self.form_data['last_id']) + 1):
            self.rows.append({
                'id': to_upca(id_num), 'company': self.form_data['company'],
                'first_name': '', 'last_name': '',
                'balance': self.form_data['balance'],
            })
        self._render_table()
        self._update_info_bar()
        self._show_step(2)

    # ── Step 2 ────────────────────────────────────────────────────────────────
    def _build_step2(self):
        outer = tk.Frame(self, bg=self.C_BG)
        self.step2_frame = outer

        info_bar = tk.Frame(outer, bg='white', bd=1, relief='solid',
                            highlightbackground=self.C_BORDER, highlightthickness=1)
        info_bar.pack(fill='x', padx=40, pady=(10, 6), ipady=8)
        self.info_text_var = tk.StringVar(value='')
        tk.Label(info_bar, textvariable=self.info_text_var, font=('Arial', 11),
                 fg='#374151', bg='white').pack(side='left', padx=14)
        btn_row = tk.Frame(info_bar, bg='white')
        btn_row.pack(side='right', padx=10)
        tk.Button(btn_row, text='← Back', font=self.font_btn,
                  bg='white', fg='#374151', relief='solid', cursor='hand2',
                  padx=12, pady=6, bd=1,
                  command=self._go_back).pack(side='left', padx=(0, 6))
        tk.Button(btn_row, text='✕ Start Over', font=self.font_btn,
                  bg='#c62828', fg='white', relief='flat', cursor='hand2',
                  activebackground='#b71c1c', activeforeground='white',
                  padx=12, pady=6,
                  command=self._start_over).pack(side='left')

        tk.Label(outer,
                 text='Click a First Name or Last Name cell to edit.  Paste (Ctrl+V) to fill down from the selected row.',
                 font=('Segoe UI', 10), fg=self.C_MUTED, bg=self.C_BG,
                 anchor='w').pack(fill='x', padx=48, pady=(0, 6))

        self.warn_frame = tk.Frame(outer, bg='#FEF9C3',
                                    highlightbackground='#FDE047', highlightthickness=1)
        self.warn_label = tk.Label(self.warn_frame, text='',
                                    font=('Segoe UI', 10, 'bold'),
                                    fg='#713F12', bg='#FEF9C3',
                                    anchor='w', wraplength=860)
        self.warn_label.pack(fill='x', padx=12, pady=7)

        table_frame = tk.Frame(outer, bg=self.C_BG,
                               highlightbackground=self.C_BORDER, highlightthickness=1)
        table_frame.pack(fill='both', expand=True, padx=48, pady=(0, 10))

        cols = ('upca', 'company', 'first_name', 'last_name', 'balance')
        self.tree = ttk.Treeview(table_frame, columns=cols, show='headings',
                                  selectmode='browse')
        style = ttk.Style()
        style.theme_use('clam')
        style.configure('Treeview', font=('Segoe UI', 11), rowheight=30,
                         background=self.C_CARD, fieldbackground=self.C_CARD,
                         foreground=self.C_TEXT, borderwidth=0)
        style.configure('Treeview.Heading', font=('Segoe UI', 10, 'bold'),
                         background='#F1F5F9', foreground=self.C_MUTED,
                         relief='flat', borderwidth=0)
        style.map('Treeview',
                  background=[('selected', self.C_ROW_SEL)],
                  foreground=[('selected', self.C_TEXT)])

        for col, heading, width in [
            ('upca',       'UPC-A (12 digit)', 160),
            ('company',    'Company Name',     160),
            ('first_name', 'First Name ✏',    150),
            ('last_name',  'Last Name ✏',     150),
            ('balance',    'Balance',          100),
        ]:
            self.tree.heading(col, text=heading, anchor='center')
            self.tree.column(col, width=width, minwidth=80, anchor='center')

        vsb = ttk.Scrollbar(table_frame, orient='vertical', command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(side='left', fill='both', expand=True)
        vsb.pack(side='right', fill='y')
        self.tree.bind('<ButtonRelease-1>', self._on_double_click)
        self.tree.bind('<Return>',    self._on_double_click)
        self.tree.bind('<Control-v>', self._on_tree_paste)
        self.tree.tag_configure('odd',  background=self.C_ROW_ODD)
        self.tree.tag_configure('even', background=self.C_ROW_EVEN)

        tk.Button(outer, text='⬇ Save custimp.txt', font=('Segoe UI', 11, 'bold'),
                  bg=self.C_ACCENT, fg='white', relief='flat', cursor='hand2',
                  activebackground=self.C_ACCENT_H, activeforeground='white',
                  padx=20, pady=10,
                  command=self._handle_download).pack(padx=48, pady=(4, 16), anchor='e')

    def _render_table(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        for i, r in enumerate(self.rows):
            try:
                bal = float(r['balance']) if r['balance'] else 0.0
            except ValueError:
                bal = 0.0
            tag = 'even' if i % 2 == 0 else 'odd'
            self.tree.insert('', 'end', iid=str(i), tags=(tag,),
                             values=(r['id'], r['company'],
                                     r['first_name'], r['last_name'],
                                     f"{abs(bal):.2f}"))

    def _update_info_bar(self):
        filled = sum(1 for r in self.rows if r['first_name'] or r['last_name'])
        count  = len(self.rows)
        fid    = self.form_data.get('first_id', '')
        lid    = self.form_data.get('last_id', '')
        msg    = f"{count} IDs Generated  –  ({fid}  to  {lid})"
        if filled:
            msg += f"  ·  {filled} name{'s' if filled != 1 else ''} filled"
        self.info_text_var.set(msg)

    def _on_double_click(self, event):
        region = self.tree.identify_region(event.x, event.y)
        if region != 'cell':
            return
        col_id  = self.tree.identify_column(event.x)
        col_idx = int(col_id.replace('#', '')) - 1
        col_key = self.tree['columns'][col_idx]
        if col_key not in ('first_name', 'last_name'):
            return
        row_id = self.tree.identify_row(event.y)
        if not row_id:
            return
        self._open_cell_editor(row_id, col_id, col_idx)

    def _open_cell_editor(self, row_id, col_id, col_idx):
        x, y, w, h = self.tree.bbox(row_id, col_id)
        val     = self.tree.item(row_id, 'values')[col_idx]
        col_key = self.tree['columns'][col_idx]
        ri      = int(row_id)

        entry = tk.Entry(self.tree, font=('Segoe UI', 11), relief='flat',
                         bg='#EFF6FF', fg=self.C_TEXT,
                         selectbackground=self.C_ACCENT, selectforeground='white',
                         insertbackground=self.C_ACCENT,
                         highlightthickness=1, highlightcolor=self.C_ACCENT,
                         highlightbackground=self.C_ACCENT)
        entry.place(x=x, y=y, width=w, height=h)
        entry.insert(0, val)
        entry.select_range(0, 'end')
        entry.focus()

        def save(ev=None):
            new_val = entry.get()
            self.rows[ri][col_key] = new_val
            vals = list(self.tree.item(row_id, 'values'))
            vals[col_idx] = new_val
            self.tree.item(row_id, values=vals)
            entry.destroy()
            self._update_info_bar()

        def on_paste(ev):
            try:
                clipboard = self.clipboard_get()
            except tk.TclError:
                return
            names     = [line.split('\t') for line in clipboard.splitlines()]
            edit_keys = ['first_name', 'last_name']
            start_col = edit_keys.index(col_key)
            available = len(self.rows) - ri
            if len(names) > available:
                self.warn_label.config(
                    text=(f"⚠  Too many names pasted: {len(names)} pasted but only "
                          f"{available} row{'s' if available != 1 else ''} "
                          f"remain{'s' if available == 1 else ''}. Extra entries ignored."))
                self.warn_frame.pack(fill='x', padx=48, pady=(0, 6), ipady=2,
                                     before=self.tree.master)
            else:
                self.warn_frame.pack_forget()
            for dr, row_names in enumerate(names):
                target_row = ri + dr
                if target_row >= len(self.rows):
                    break
                for dc, name_val in enumerate(row_names):
                    tc = start_col + dc
                    if tc < len(edit_keys):
                        self.rows[target_row][edit_keys[tc]] = name_val.strip()
            self._render_table()
            self._update_info_bar()
            entry.destroy()
            return 'break'

        entry.bind('<Return>',    save)
        entry.bind('<Escape>',    lambda e: entry.destroy())
        entry.bind('<FocusOut>',  save)
        entry.bind('<Control-v>', on_paste)

    def _on_tree_paste(self, event):
        sel = self.tree.selection()
        if not sel:
            return
        row_id    = sel[0]
        ri        = int(row_id)
        edit_keys = ['first_name', 'last_name']
        try:
            clipboard = self.clipboard_get()
        except tk.TclError:
            return
        names     = [line.split('\t') for line in clipboard.splitlines()]
        available = len(self.rows) - ri
        if len(names) > available:
            self.warn_label.config(
                text=(f"⚠  Too many names pasted: {len(names)} pasted but only "
                      f"{available} row{'s' if available != 1 else ''} "
                      f"remain{'s' if available == 1 else ''}. Extra entries ignored."))
            self.warn_frame.pack(fill='x', padx=48, pady=(0, 6), ipady=2)
        else:
            self.warn_frame.pack_forget()
        for dr, row_names in enumerate(names):
            target_row = ri + dr
            if target_row >= len(self.rows):
                break
            for dc, name_val in enumerate(row_names):
                if dc < len(edit_keys):
                    self.rows[target_row][edit_keys[dc]] = name_val.strip()
        self._render_table()
        self._update_info_bar()

    def _handle_download(self):
        folder = self.config.get('save_folder', '').strip()
        if folder and not os.path.isdir(folder):
            messagebox.showwarning(
                'Folder Not Found',
                f'The saved folder could not be found:\n\n{folder}\n\n'
                'Please choose a new save location.')
            folder = ''
        if not folder:
            folder = filedialog.askdirectory(title='Select Save Folder for custimp.txt')
            if not folder:
                return
            self.config['save_folder'] = folder
            save_config(self.config)
            self.folder_var.set(folder)

        ok, result = archive_existing(folder)
        if not ok:
            messagebox.showerror(
                'Archive Error',
                f'Could not rename the existing custimp.txt:\n\n{result}\n\n'
                'Please make sure the file is not open in another program.')
            return

        dest = os.path.join(folder, 'custimp.txt')
        try:
            with open(dest, 'w', newline='', encoding='utf-8') as f:
                f.write(build_csv(self.rows))
        except Exception as e:
            messagebox.showerror('Save Error', f'Could not save file:\n{e}')
            return

        if result:
            messagebox.showinfo(
                'Saved',
                f'custimp.txt saved to:\n{folder}\n\n'
                f'Previous file archived as:\n{os.path.basename(result)}')
        else:
            messagebox.showinfo('Saved', f'custimp.txt saved to:\n{folder}')

    def _go_back(self):
        self._show_step(1)

    def _start_over(self):
        self.rows      = []
        self.form_data = {}
        for entry, placeholder in [
            (self.e_first_id, 'Scan or type 11-digit ID'),
            (self.e_last_id,  'Scan or type 11-digit ID'),
            (self.e_company,  'e.g. Acme Corp'),
            (self.e_balance,  'e.g. 100.00'),
        ]:
            entry.delete(0, 'end')
            entry.insert(0, placeholder)
            entry.config(fg='#9ca3af')
        for err in [self.err_first_id, self.err_last_id,
                    self.err_company,  self.err_balance]:
            err.config(text='')
        self.card_count_var.set('—')
        self.card_count_lbl.config(fg=self.C_ACCENT)
        self.warn_frame.pack_forget()
        self._show_step(1)

    def _show_step(self, n):
        if n == 1:
            if self.step2_frame:
                self.step2_frame.pack_forget()
            self.step1_frame.pack(fill='both', expand=True)
        else:
            self.step1_frame.pack_forget()
            self.step2_frame.pack(fill='both', expand=True)
        self._update_step_indicators(n)


if __name__ == '__main__':
    app = CustImpApp()
    app.mainloop()
