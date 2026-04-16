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
APP_VERSION = "1.0.0"

# ── UPDATE THESE TWO LINES with your GitHub username and repo name ────────────
GITHUB_USER = "Nate3700"
GITHUB_REPO = "Customer_Import_Tool"
# ─────────────────────────────────────────────────────────────────────────────

VERSION_URL  = f"https://raw.githubusercontent.com/{GITHUB_USER}/{GITHUB_REPO}/main/version.json"
DOWNLOAD_URL = f"https://github.com/{GITHUB_USER}/{GITHUB_REPO}/releases/latest/download/custimp.exe"

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
    progress.configure(bg='#f0f2f5')
    progress.grab_set()

    tk.Label(progress, text=f"Downloading version {latest_version}...",
             font=('Arial', 11), bg='#f0f2f5', fg='#1a1a2e').pack(pady=(20, 8))
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
        self.minsize(700, 560)
        self.configure(bg='#f0f2f5')
        self.resizable(True, True)

        self.rows      = []
        self.form_data = {}
        self.config    = load_config()

        self.font_label = ('Arial', 10, 'bold')
        self.font_input = ('Arial', 12)
        self.font_mono  = ('Courier New', 11)
        self.font_btn   = ('Arial', 10, 'bold')

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
        hdr = tk.Frame(self, bg='#f0f2f5')
        hdr.pack(fill='x', padx=40, pady=(30, 10))

        tk.Label(hdr, text='CUSTOMER IMPORT TOOL', font=('Arial', 9, 'bold'),
                 fg='#1565c0', bg='#f0f2f5').pack()

        title_frame = tk.Frame(hdr, bg='#f0f2f5')
        title_frame.pack()
        tk.Label(title_frame, text='custimp', font=('Arial', 28, 'bold'),
                 fg='#1a1a2e', bg='#f0f2f5').pack(side='left')
        tk.Label(title_frame, text='.txt', font=('Arial', 28, 'bold'),
                 fg='#1565c0', bg='#f0f2f5').pack(side='left')
        tk.Label(title_frame, text=f'  v{APP_VERSION}', font=('Arial', 11),
                 fg='#9ca3af', bg='#f0f2f5').pack(side='left', pady=(10, 0))

        self.step_frame = tk.Frame(hdr, bg='#f0f2f5')
        self.step_frame.pack(pady=(12, 0))

        self.s1_circle = tk.Label(self.step_frame, text='1', width=2,
                                   font=('Arial', 10, 'bold'), relief='flat')
        self.s1_circle.pack(side='left', padx=(0, 4))
        self.s1_label = tk.Label(self.step_frame, text='Enter Details',
                                  font=('Arial', 10, 'bold'), bg='#f0f2f5')
        self.s1_label.pack(side='left')

        tk.Label(self.step_frame, text='────', fg='#cbd5e1',
                 bg='#f0f2f5', font=('Arial', 10)).pack(side='left', padx=6)

        self.s2_circle = tk.Label(self.step_frame, text='2', width=2,
                                   font=('Arial', 10, 'bold'), relief='flat')
        self.s2_circle.pack(side='left', padx=(0, 4))
        self.s2_label = tk.Label(self.step_frame, text='Fill Names & Export',
                                  font=('Arial', 10, 'bold'), bg='#f0f2f5')
        self.s2_label.pack(side='left')

        self._update_step_indicators(1)

    def _update_step_indicators(self, step):
        if step == 1:
            self.s1_circle.config(bg='#1565c0', fg='white', text='1')
            self.s1_label.config(fg='#1a1a2e')
            self.s2_circle.config(bg='#cbd5e1', fg='#64748b', text='2')
            self.s2_label.config(fg='#94a3b8')
        else:
            self.s1_circle.config(bg='#2e7d32', fg='white', text='✓')
            self.s1_label.config(fg='#1a1a2e')
            self.s2_circle.config(bg='#1565c0', fg='white', text='2')
            self.s2_label.config(fg='#1a1a2e')

    # ── Step 1 ────────────────────────────────────────────────────────────────
    def _build_step1(self):
        outer = tk.Frame(self, bg='#f0f2f5')
        self.step1_frame = outer

        card = tk.Frame(outer, bg='white', bd=1, relief='solid',
                        highlightbackground='#d1d9e0', highlightthickness=1)
        card.pack(padx=40, pady=10, ipadx=30, ipady=20)

        def add_field(parent, label_text, var, placeholder='', optional_text=''):
            row = tk.Frame(parent, bg='white')
            row.pack(fill='x', pady=(0, 14))
            lbl_frame = tk.Frame(row, bg='white')
            lbl_frame.pack(fill='x')
            tk.Label(lbl_frame, text=label_text, font=self.font_label,
                     fg='#374151', bg='white').pack(side='left')
            if optional_text:
                tk.Label(lbl_frame, text=optional_text, font=('Arial', 9),
                         fg='#6b7280', bg='white').pack(side='left', padx=(4, 0))
            entry = tk.Entry(row, textvariable=var, font=self.font_input,
                             bg='#f9fafb', fg='#1a1a2e', relief='flat',
                             highlightthickness=1, highlightbackground='#d1d9e0',
                             highlightcolor='#1565c0', width=36)
            entry.pack(fill='x', ipady=6)
            if placeholder:
                entry.insert(0, placeholder)
                entry.config(fg='#9ca3af')
                def on_focus_in(e, en=entry, ph=placeholder):
                    if en.get() == ph:
                        en.delete(0, 'end')
                        en.config(fg='#1a1a2e')
                def on_focus_out(e, en=entry, ph=placeholder):
                    if not en.get():
                        en.insert(0, ph)
                        en.config(fg='#9ca3af')
                entry.bind('<FocusIn>',  on_focus_in)
                entry.bind('<FocusOut>', on_focus_out)
            err_lbl = tk.Label(row, text='', font=('Arial', 9, 'bold'),
                               fg='#c62828', bg='white')
            err_lbl.pack(fill='x')
            return entry, err_lbl

        self.v_first_id = tk.StringVar()
        self.v_last_id  = tk.StringVar()
        self.v_company  = tk.StringVar()
        self.v_balance  = tk.StringVar()

        self.e_first_id, self.err_first_id = add_field(
            card, 'First ID Card Number', self.v_first_id, 'Scan or type 11-digit ID')
        self.e_last_id, self.err_last_id = add_field(
            card, 'Last ID Card Number', self.v_last_id, 'Scan or type 11-digit ID')
        self.e_company, self.err_company = add_field(
            card, 'Company Name', self.v_company, 'e.g. Acme Corp', '(optional)')
        self.e_balance, self.err_balance = add_field(
            card, 'Balance', self.v_balance, 'e.g. 100.00', '(optional, defaults to 0)')

        # Card count
        count_frame = tk.Frame(card, bg='#f0f6ff', bd=1, relief='solid',
                               highlightbackground='#90b8f8', highlightthickness=1)
        count_frame.pack(fill='x', pady=(6, 14), ipady=8, ipadx=10)
        tk.Label(count_frame, text='Cards to be Generated', font=('Arial', 10, 'bold'),
                 fg='#374151', bg='#f0f6ff').pack(side='left', padx=10)
        self.card_count_var = tk.StringVar(value='—')
        self.card_count_lbl = tk.Label(count_frame, textvariable=self.card_count_var,
                                        font=('Arial', 16, 'bold'), fg='#1565c0', bg='#f0f6ff')
        self.card_count_lbl.pack(side='right', padx=10)

        # Save folder row
        folder_row = tk.Frame(card, bg='white')
        folder_row.pack(fill='x', pady=(0, 14))
        tk.Label(folder_row, text='Save Folder', font=self.font_label,
                 fg='#374151', bg='white').pack(anchor='w')
        folder_inner = tk.Frame(folder_row, bg='white')
        folder_inner.pack(fill='x')
        saved   = self.config.get('save_folder', '')
        display = saved if saved else 'No folder set — will prompt on first save'
        self.folder_var = tk.StringVar(value=display)
        tk.Label(folder_inner, textvariable=self.folder_var,
                 font=('Arial', 9), fg='#4b5563', bg='#f9fafb',
                 anchor='w', relief='flat', padx=8, pady=4,
                 highlightthickness=1,
                 highlightbackground='#d1d9e0').pack(side='left', fill='x', expand=True)
        tk.Button(folder_inner, text='Change', font=('Arial', 9, 'bold'),
                  bg='#e8edf5', fg='#374151', relief='flat', cursor='hand2',
                  padx=10, pady=4,
                  command=self._change_folder).pack(side='left', padx=(6, 0))

        tk.Button(card, text='Generate Spreadsheet →', font=self.font_btn,
                  bg='#1565c0', fg='white', relief='flat', cursor='hand2',
                  activebackground='#0d47a1', activeforeground='white',
                  padx=20, pady=10,
                  command=self._handle_generate).pack(fill='x')

        self.v_first_id.trace_add('write', lambda *a: self._live_validate())
        self.v_last_id.trace_add('write',  lambda *a: self._live_validate())
        for e in [self.e_first_id, self.e_last_id, self.e_company, self.e_balance]:
            e.bind('<Return>', lambda ev: self._handle_generate())

    def _change_folder(self):
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
                self.e_first_id.config(highlightbackground='#c62828')
            else:
                self.err_first_id.config(text='')
                self.e_first_id.config(highlightbackground='#d1d9e0')
        else:
            self.err_first_id.config(text='')
            self.e_first_id.config(highlightbackground='#d1d9e0')

        if last:
            if not last.isdigit() or len(last) != 11:
                self.err_last_id.config(
                    text='Must be exactly 11 digits (or scan 12-digit barcode)')
                self.e_last_id.config(highlightbackground='#c62828')
            elif first and len(first) == 11 and int(last) <= int(first):
                self.err_last_id.config(text='Last ID must be greater than First ID')
                self.e_last_id.config(highlightbackground='#c62828')
            elif first and len(first) == 11 and int(last) - int(first) > 999:
                self.err_last_id.config(text='Max 1000 rows at once')
                self.e_last_id.config(highlightbackground='#c62828')
            else:
                self.err_last_id.config(text='')
                self.e_last_id.config(highlightbackground='#d1d9e0')
        else:
            self.err_last_id.config(text='')
            self.e_last_id.config(highlightbackground='#d1d9e0')

        if (len(first) == 11 and len(last) == 11 and
                first.isdigit() and last.isdigit() and int(last) > int(first)):
            self.card_count_var.set(f'{int(last) - int(first) + 1:,}')
            self.card_count_lbl.config(fg='#1565c0')
        else:
            self.card_count_var.set('—')
            self.card_count_lbl.config(
                fg='#c62828' if (first or last) else '#1565c0')

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
        outer = tk.Frame(self, bg='#f0f2f5')
        self.step2_frame = outer

        info_bar = tk.Frame(outer, bg='white', bd=1, relief='solid',
                            highlightbackground='#d1d9e0', highlightthickness=1)
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
                  command=self._start_over).pack(side='left', padx=(0, 6))
        tk.Button(btn_row, text='⬇ Save custimp.txt', font=self.font_btn,
                  bg='#1565c0', fg='white', relief='flat', cursor='hand2',
                  activebackground='#0d47a1', activeforeground='white',
                  padx=12, pady=6,
                  command=self._handle_download).pack(side='left')

        tip = tk.Frame(outer, bg='#fffbea', bd=1, relief='solid',
                       highlightbackground='#f9c600', highlightthickness=1)
        tip.pack(fill='x', padx=40, pady=(0, 6), ipady=6)
        tk.Label(tip,
                 text='💡  Tip: Click a First Name or Last Name cell and paste (Ctrl+V) to fill down automatically.',
                 font=('Arial', 10), fg='#374151', bg='#fffbea',
                 anchor='w').pack(fill='x', padx=10)

        self.warn_frame = tk.Frame(outer, bg='#fff8e1', bd=1, relief='solid',
                                    highlightbackground='#f9c600', highlightthickness=1)
        self.warn_label = tk.Label(self.warn_frame, text='',
                                    font=('Arial', 10, 'bold'),
                                    fg='#7b3f00', bg='#fff8e1',
                                    anchor='w', wraplength=800)
        self.warn_label.pack(fill='x', padx=10, pady=6)

        table_frame = tk.Frame(outer, bg='#f0f2f5')
        table_frame.pack(fill='both', expand=True, padx=40, pady=(0, 10))

        cols = ('upca', 'company', 'first_name', 'last_name', 'balance')
        self.tree = ttk.Treeview(table_frame, columns=cols, show='headings',
                                  selectmode='browse')
        style = ttk.Style()
        style.theme_use('clam')
        style.configure('Treeview', font=('Arial', 11), rowheight=28,
                         background='white', fieldbackground='white',
                         foreground='#1a1a2e')
        style.configure('Treeview.Heading', font=('Arial', 10, 'bold'),
                         background='#e8edf5', foreground='#4b5563')
        style.map('Treeview', background=[('selected', '#dbeafe')])

        for col, heading, width in [
            ('upca',       'UPC-A (12 digit)', 160),
            ('company',    'Company Name',     160),
            ('first_name', 'First Name ✏',    150),
            ('last_name',  'Last Name ✏',     150),
            ('balance',    'Balance',          100),
        ]:
            self.tree.heading(col, text=heading)
            self.tree.column(col, width=width, minwidth=80, anchor='w')

        vsb = ttk.Scrollbar(table_frame, orient='vertical', command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(side='left', fill='both', expand=True)
        vsb.pack(side='right', fill='y')
        self.tree.bind('<Double-1>',  self._on_double_click)
        self.tree.bind('<Return>',    self._on_double_click)
        self.tree.bind('<Control-v>', self._on_tree_paste)
        self.tree.tag_configure('odd',  background='#f8fafc')
        self.tree.tag_configure('even', background='white')

        tk.Button(outer, text='⬇ Save custimp.txt', font=self.font_btn,
                  bg='#1565c0', fg='white', relief='flat', cursor='hand2',
                  activebackground='#0d47a1', activeforeground='white',
                  padx=20, pady=10,
                  command=self._handle_download).pack(padx=40, pady=(0, 16), anchor='e')

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
                                     f"{-abs(bal):.2f}"))

    def _update_info_bar(self):
        filled = sum(1 for r in self.rows if r['first_name'] or r['last_name'])
        count  = len(self.rows)
        fid    = self.form_data.get('first_id', '')
        lid    = self.form_data.get('last_id', '')
        msg    = f"{count} rows  ·  IDs {fid} – {lid}"
        msg   += (f"  ·  {filled} name{'s' if filled != 1 else ''} filled"
                  if filled else "  ·  Paste names into First Name / Last Name cells")
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

        entry = tk.Entry(self.tree, font=('Arial', 11), relief='flat',
                         bg='#e8f0fe', fg='#1565c0',
                         highlightthickness=1, highlightcolor='#1565c0',
                         highlightbackground='#1565c0')
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
                self.warn_frame.pack(fill='x', padx=40, pady=(0, 6), ipady=2,
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
            self.warn_frame.pack(fill='x', padx=40, pady=(0, 6), ipady=2)
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
        self.card_count_lbl.config(fg='#1565c0')
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
