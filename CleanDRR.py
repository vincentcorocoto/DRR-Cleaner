import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import threading
try:
    from tkinterdnd2 import TkinterDnD, DND_FILES
    _DND_AVAILABLE = True
except ImportError:
    _DND_AVAILABLE = False
import os
import io
import sqlite3
import hashlib
import pandas as pd
from openpyxl.styles import PatternFill, Font
import openpyxl

# ══════════════════════════════════════════════════════════════════════════════
# THEME
# ══════════════════════════════════════════════════════════════════════════════
BG        = "#0F1117"
CARD      = "#1A1D27"
BORDER    = "#2A2D3E"
ACCENT    = "#4F8EF7"
ACCENT2   = "#7C3AED"
SUCCESS   = "#22C55E"
DANGER    = "#EF4444"
WARNING   = "#F59E0B"
TEXT      = "#F1F5F9"
SUBTEXT   = "#94A3B8"
HEADER_BG = "#1E2235"
DARK      = "#0A0C14"
GOLD      = "#F59E0B"

# ══════════════════════════════════════════════════════════════════════════════
# DRR CLEANER CONFIG
# ══════════════════════════════════════════════════════════════════════════════
STATUS_COL_NAME    = "Status"
STATUS_COL_INDEX   = 9
DATE_COL_NAME      = "Date"
TIME_COL_NAME      = "Time"
COL_E_INDEX        = 4
DIALED_COL_INDEX   = 5
PTP_AMOUNT_COL     = "PTP Amount"
CLAIM_PAID_COL     = "Claim Paid Amount"
REMARK_COL         = "Remark"
REMOVE_STATUSES    = {"BP", "REACTIVE", "SMS FAILED", "NEW", "ABORTED"}
COLS_TO_DROP_START = 27
COLS_TO_DROP_END   = 50

# ══════════════════════════════════════════════════════════════════════════════
# DATABASE — SQLite user store
# ══════════════════════════════════════════════════════════════════════════════
DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "users.db")

def _hash(password: str) -> str:
    return hashlib.sha256(password.encode()).hexdigest()

def db_init():
    """Create the users table and seed default accounts if empty."""
    con = sqlite3.connect(DB_PATH)
    cur = con.cursor()
    cur.execute("""
        CREATE TABLE IF NOT EXISTS users (
            username TEXT PRIMARY KEY,
            password_hash TEXT NOT NULL,
            role TEXT NOT NULL DEFAULT 'user',
            created_at TEXT
        )
    """)
    con.commit()
    # Seed defaults only when the table is empty
    if cur.execute("SELECT COUNT(*) FROM users").fetchone()[0] == 0:
        seeds = [
            ("admin",      "admin123",    "admin"),
            ("supervisor", "super2024",   "supervisor"),
            ("vincent",    "orico2024",   "user"),
            ("user",       "password",    "user"),
        ]
        for u, p, r in seeds:
            cur.execute("INSERT INTO users VALUES (?,?,?,datetime('now','localtime'))",
                        (u, _hash(p), r))
        con.commit()
    con.close()

def db_check_login(username: str, password: str):
    """Return role string on success, None on failure."""
    con = sqlite3.connect(DB_PATH)
    row = con.execute(
        "SELECT role FROM users WHERE username=? AND password_hash=?",
        (username, _hash(password))
    ).fetchone()
    con.close()
    return row[0] if row else None

def db_list_users():
    """Return list of (username, role, created_at)."""
    con = sqlite3.connect(DB_PATH)
    rows = con.execute(
        "SELECT username, role, created_at FROM users ORDER BY username"
    ).fetchall()
    con.close()
    return rows

def db_add_user(username: str, password: str, role: str = "user") -> str:
    """Add a new user. Returns '' on success or an error message."""
    if not username or not password:
        return "Username and password are required."
    try:
        con = sqlite3.connect(DB_PATH)
        con.execute("INSERT INTO users (username, password_hash, role, created_at) VALUES (?,?,?,datetime('now','localtime'))",
                    (username, _hash(password), role))
        con.commit()
        con.close()
        return ""
    except sqlite3.IntegrityError:
        return f"Username '{username}' already exists."

def db_delete_user(username: str) -> str:
    """Delete a user. Prevents deleting the last admin."""
    con = sqlite3.connect(DB_PATH)
    admin_count = con.execute(
        "SELECT COUNT(*) FROM users WHERE role='admin'"
    ).fetchone()[0]
    role = con.execute(
        "SELECT role FROM users WHERE username=?", (username,)
    ).fetchone()
    if role and role[0] == "admin" and admin_count <= 1:
        con.close()
        return "Cannot delete the only admin account."
    con.execute("DELETE FROM users WHERE username=?", (username,))
    con.commit()
    con.close()
    return ""

def db_update_user(username: str, new_password: str = "", new_role: str = "") -> str:
    """Update password and/or role for a user."""
    con = sqlite3.connect(DB_PATH)
    if new_password:
        con.execute("UPDATE users SET password_hash=? WHERE username=?",
                    (_hash(new_password), username))
    if new_role:
        con.execute("UPDATE users SET role=? WHERE username=?",
                    (new_role, username))
    con.commit()
    con.close()
    return ""

# Initialise DB on import
db_init()

# ══════════════════════════════════════════════════════════════════════════════
# HELPERS
# ══════════════════════════════════════════════════════════════════════════════
def show_toast(parent, message, color=SUCCESS, duration=3000):
    toast = tk.Toplevel(parent)
    toast.overrideredirect(True)
    toast.attributes("-topmost", True)
    toast.attributes("-alpha", 0.95)
    frame = tk.Frame(toast, bg=color, padx=20, pady=12)
    frame.pack()
    tk.Label(frame, text="✔  " + message, font=("Segoe UI", 10, "bold"),
             bg=color, fg="white").pack()
    parent.update_idletasks()
    pw, ph = parent.winfo_width(), parent.winfo_height()
    px, py = parent.winfo_x(), parent.winfo_y()
    toast.update_idletasks()
    tw, th = toast.winfo_reqwidth(), toast.winfo_reqheight()
    toast.geometry(f"+{px + pw - tw - 20}+{py + ph - th - 40}")

    def fade(alpha=0.95):
        alpha -= 0.05
        if alpha <= 0:
            toast.destroy()
        else:
            toast.attributes("-alpha", alpha)
            toast.after(40, lambda: fade(alpha))

    toast.after(duration, fade)


# ══════════════════════════════════════════════════════════════════════════════
# DRR PROCESSING LOGIC
# ══════════════════════════════════════════════════════════════════════════════
def process_file(filepath):
    df_headers = pd.read_excel(filepath, nrows=0)
    str_cols = {}
    if COL_E_INDEX < len(df_headers.columns):
        str_cols[df_headers.columns[COL_E_INDEX]] = str
    col_e_name      = df_headers.columns[COL_E_INDEX]      if COL_E_INDEX      < len(df_headers.columns) else None
    dialed_col_name = df_headers.columns[DIALED_COL_INDEX] if DIALED_COL_INDEX < len(df_headers.columns) else None
    if dialed_col_name:
        str_cols[dialed_col_name] = str
    for col in df_headers.columns:
        if "dialed" in str(col).lower():
            str_cols[col] = str
            dialed_col_name = col
            break
    df = pd.read_excel(filepath, dtype=str_cols)

    def vectorized_strip_decimal(series):
        s = series.fillna("").astype(str).str.strip()
        has_dot   = s.str.contains(".", regex=False)
        int_part  = s.str.split(".").str[0]
        is_numeric = int_part.str.lstrip("-").str.isdigit()
        return s.where(~(has_dot & is_numeric), int_part).replace("nan", "")

    def vectorized_clean_account(series):
        s        = series.fillna("").astype(str).str.strip()
        ends_dot0 = s.str.endswith(".0")
        base      = s.str[:-2]
        is_safe   = ends_dot0 & base.str.lstrip("-").str.isdigit() & ~base.str.startswith("0")
        return s.where(~is_safe, base).replace("nan", "")

    if col_e_name and col_e_name in df.columns:
        df[col_e_name] = vectorized_clean_account(df[col_e_name])
    if dialed_col_name and dialed_col_name in df.columns:
        df[dialed_col_name] = vectorized_strip_decimal(df[dialed_col_name])
    else:
        for col in df.columns:
            if "dialed" in str(col).lower():
                df[col] = vectorized_strip_decimal(df[col])
                dialed_col_name = col
                break

    STATUS_COL = STATUS_COL_NAME if STATUS_COL_NAME in df.columns else (
        df.columns[STATUS_COL_INDEX] if STATUS_COL_INDEX < len(df.columns) else None
    )
    if not STATUS_COL:
        raise ValueError("Status column not found.")

    status_normalized = df[STATUS_COL].astype(str).str.strip().str.upper()
    is_blank          = df[STATUS_COL].isna() | (df[STATUS_COL].astype(str).str.strip() == "")
    has_cease         = status_normalized.str.contains("CEASE", na=False)
    is_removable      = status_normalized.isin(REMOVE_STATUSES) | is_blank | has_cease

    if PTP_AMOUNT_COL in df.columns:
        has_ptp       = status_normalized.str.contains(r"\bPTP\b", regex=True, na=False)
        ptp_numeric   = pd.to_numeric(df[PTP_AMOUNT_COL].astype(str).str.replace(",", "", regex=False), errors="coerce").fillna(0)
        ptp_has_value = has_ptp & (ptp_numeric > 0)
        ptp_no_value  = has_ptp & (ptp_numeric <= 0)
        is_removable  = (is_removable | ptp_no_value) & ~ptp_has_value

    if CLAIM_PAID_COL in df.columns:
        has_kept       = status_normalized.str.contains(r"\bKEPT\b", regex=True, na=False)
        claim_numeric  = pd.to_numeric(df[CLAIM_PAID_COL].astype(str).str.replace(",", "", regex=False), errors="coerce").fillna(0)
        kept_has_value = has_kept & (claim_numeric > 0)
        kept_no_value  = has_kept & (claim_numeric <= 0)
        is_removable   = (is_removable | kept_no_value) & ~kept_has_value

    cleaned_df = df[~is_removable].copy()
    removed_df = df[is_removable].copy()

    s_norm  = removed_df[STATUS_COL].fillna("").astype(str).str.strip()
    s_upper = s_norm.str.upper()
    reason  = pd.Series("Status: " + s_norm, index=removed_df.index)
    reason  = reason.where(~s_upper.isin(REMOVE_STATUSES), "Status: " + s_norm)
    reason  = reason.where(~s_upper.str.contains("CEASE", na=False), "Status contains CEASE: " + s_norm)
    reason  = reason.where(~s_upper.str.contains("PTP",   na=False), "PTP with no PTP Amount")
    reason  = reason.where(~s_upper.str.contains("KEPT",  na=False), "KEPT with no Claim Paid Amount")
    reason  = reason.where(s_upper != "",                             reason)
    reason  = reason.where(~(s_upper.isin(["", "NAN"])),             "Blank Status")
    removed_df.insert(0, "Removed Reason", reason)

    srp_mask   = pd.Series([False] * len(cleaned_df), index=cleaned_df.index)
    remarks_df = pd.DataFrame(columns=["Row #", STATUS_COL,
                                        REMARK_COL + " (Before)",
                                        REMARK_COL + " (After)"])  # always empty DataFrame with correct cols

    if REMARK_COL in cleaned_df.columns:
        remark_norm    = cleaned_df[REMARK_COL].astype(str).str.strip().str.upper()
        status_clean   = cleaned_df[STATUS_COL].astype(str).str.strip().str.upper()
        has_action_ptp = remark_norm.str.contains(r"ACTION\s*:\s*PTP", regex=True, na=False)
        status_not_ptp_kept = (
            ~status_clean.str.contains("PTP",  na=False) &
            ~status_clean.str.contains("KEPT", na=False)
        )
        srp_mask = has_action_ptp & status_not_ptp_kept
        if srp_mask.any():
            # Build remarks log strictly from cleaned_df rows matching srp_mask only
            changed_rows = cleaned_df.loc[srp_mask, [STATUS_COL, REMARK_COL]].copy()
            remarks_df = pd.DataFrame({
                "Row #":                  range(1, srp_mask.sum() + 1),
                STATUS_COL:               changed_rows[STATUS_COL].values,
                REMARK_COL + " (Before)": changed_rows[REMARK_COL].astype(str).values,
                REMARK_COL + " (After)":  changed_rows[REMARK_COL].astype(str).str.replace(
                    r"(?i)Action\s*:\s*PTP", "Action: SRP", regex=True).values,
            }).reset_index(drop=True)
        cleaned_df.loc[srp_mask, REMARK_COL] = cleaned_df.loc[srp_mask, REMARK_COL].astype(str).str.replace(
            r"(?i)Action\s*:\s*PTP", "Action: SRP", regex=True)

    cols_to_drop = [df.columns[i] for i in range(COLS_TO_DROP_START, min(COLS_TO_DROP_END + 1, len(df.columns)))]
    cleaned_df.drop(columns=cols_to_drop, inplace=True, errors="ignore")
    removed_df.drop(columns=cols_to_drop, inplace=True, errors="ignore")

    if DATE_COL_NAME in cleaned_df.columns and TIME_COL_NAME in cleaned_df.columns:
        # Parse the Date column (Column B) for the date portion
        raw_date = pd.to_datetime(cleaned_df[DATE_COL_NAME], errors="coerce")
        # Parse the Time column (Column C) — extract time portion only
        raw_time = pd.to_datetime(cleaned_df[TIME_COL_NAME], errors="coerce")
        date_str = raw_date.dt.strftime("%m/%d/%Y")
        # Time only — strip any date portion from Time column
        time_str = raw_time.dt.strftime("%I:%M:%S %p").str.lstrip("0")
        # Combine: "06/16/2026  4:51:16 PM"
        cleaned_df[TIME_COL_NAME] = date_str + "  " + time_str
    if DATE_COL_NAME in cleaned_df.columns:
        cleaned_df[DATE_COL_NAME] = pd.to_datetime(cleaned_df[DATE_COL_NAME], errors="coerce").dt.strftime("%m/%d/%Y")

    stats = {
        "total":       len(df),
        "retained":    len(cleaned_df),
        "removed":     len(removed_df),
        "srp_changed": int(srp_mask.sum()),
    }

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        for sheet_name, frame in [("Cleaned", cleaned_df), ("Removed", removed_df)]:
            frame.to_excel(writer, index=False, sheet_name=sheet_name)
            ws = writer.sheets[sheet_name]
            normal_fill  = PatternFill("solid", fgColor="1E2235")
            orange_fill  = PatternFill("solid", fgColor="C05621")
            white_bold   = Font(bold=True, color="FFFFFF")
            for cell in ws[1]:
                cell.font = white_bold
                cell.fill = orange_fill if cell.value == "Removed Reason" else normal_fill
            for col in ws.columns:
                header_len = len(str(col[0].value)) if col[0].value else 10
                ws.column_dimensions[col[0].column_letter].width = min(header_len + 6, 40)
        if not remarks_df.empty:
            remarks_df.to_excel(writer, index=False, sheet_name="Remarks Changes")
            ws = writer.sheets["Remarks Changes"]
            for cell in ws[1]:
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = PatternFill("solid", fgColor="7C3AED")
            for col in ws.columns:
                max_len = max((len(str(c.value)) for c in col if c.value), default=10)
                ws.column_dimensions[col[0].column_letter].width = min(max_len + 4, 60)

    buf.seek(0)
    return cleaned_df, removed_df, remarks_df, stats, buf.read(), col_e_name, dialed_col_name


# ══════════════════════════════════════════════════════════════════════════════
# SCREEN 1 — LOGIN
# ══════════════════════════════════════════════════════════════════════════════
class LoginScreen(tk.Frame):
    def __init__(self, master, on_login):
        super().__init__(master, bg=BG)
        self.on_login = on_login
        self._build()

    def _build(self):
        # Animated gradient banner
        banner = tk.Frame(self, bg=HEADER_BG, height=8)
        banner.pack(fill="x")

        # Center card container
        outer = tk.Frame(self, bg=BG)
        outer.pack(expand=True)

        card = tk.Frame(outer, bg=CARD, padx=48, pady=44,
                        highlightthickness=1, highlightbackground=BORDER)
        card.pack(padx=20, pady=20)

        # Logo area
        logo_frame = tk.Frame(card, bg=CARD)
        logo_frame.pack(pady=(0, 24))

        logo_circle = tk.Canvas(logo_frame, width=72, height=72, bg=CARD,
                                highlightthickness=0)
        logo_circle.pack()
        logo_circle.create_oval(4, 4, 68, 68, fill=ACCENT, outline="")
        logo_circle.create_text(36, 36, text="⬡", font=("Segoe UI", 26, "bold"), fill="white")

        tk.Label(card, text="Welcome Back", font=("Segoe UI", 20, "bold"),
                 bg=CARD, fg=TEXT).pack()
        tk.Label(card, text="Sign in to your account to continue",
                 font=("Segoe UI", 9), bg=CARD, fg=SUBTEXT).pack(pady=(4, 28))

        # Username field
        self._field(card, "👤  Username", "username")
        tk.Frame(card, bg=BG, height=10).pack()
        # Password field
        self._field(card, "🔒  Password", "password", show="●")

        # Error label
        self.err_var = tk.StringVar()
        self.err_lbl = tk.Label(card, textvariable=self.err_var,
                                font=("Segoe UI", 9), bg=CARD, fg=DANGER)
        self.err_lbl.pack(pady=(8, 0))

        # Login button
        btn_frame = tk.Frame(card, bg=CARD)
        btn_frame.pack(fill="x", pady=(16, 0))

        self.login_btn = tk.Button(
            btn_frame, text="Sign In  →",
            font=("Segoe UI", 11, "bold"),
            bg=ACCENT, fg="white", relief="flat",
            pady=11, cursor="hand2",
            activebackground="#3a7be8", activeforeground="white",
            command=self._attempt_login
        )
        self.login_btn.pack(fill="x")

        # Bind Enter key
        self.master.bind("<Return>", lambda e: self._attempt_login())



        # Bottom bar
        footer = tk.Frame(self, bg=DARK)
        footer.pack(fill="x", side="bottom")
        tk.Label(footer, text="Created by  Vincent Corocoto  ·  09567796275",
                 font=("Segoe UI", 8, "bold"), bg=DARK, fg=ACCENT).pack(side="left", padx=16, pady=8)
        tk.Label(footer, text='"Kapag ang palay naging bigas, May bumayo."',
                 font=("Segoe UI", 8, "italic"), bg=DARK, fg="#4A5568").pack(side="right", padx=16, pady=8)

    def _field(self, parent, label_text, attr, show=None):
        lbl = tk.Label(parent, text=label_text, font=("Segoe UI", 9, "bold"),
                       bg=CARD, fg=SUBTEXT, anchor="w")
        lbl.pack(fill="x", pady=(0, 4))

        wrapper = tk.Frame(parent, bg=BORDER, padx=1, pady=1)
        wrapper.pack(fill="x")
        inner = tk.Frame(wrapper, bg=CARD)
        inner.pack(fill="x")

        var = tk.StringVar()
        entry = tk.Entry(inner, textvariable=var,
                         font=("Segoe UI", 11), bg=CARD, fg=TEXT,
                         insertbackground=TEXT, relief="flat", bd=8,
                         show=show or "")
        entry.pack(fill="x")
        setattr(self, f"{attr}_var", var)
        setattr(self, f"{attr}_entry", entry)

        # Focus glow effect
        def on_focus_in(e):
            wrapper.config(bg=ACCENT)
        def on_focus_out(e):
            wrapper.config(bg=BORDER)
        entry.bind("<FocusIn>", on_focus_in)
        entry.bind("<FocusOut>", on_focus_out)

    def _attempt_login(self):
        user = self.username_var.get().strip()
        pw   = self.password_var.get().strip()
        if not user or not pw:
            self.err_var.set("⚠  Please enter username and password.")
            return
        role = db_check_login(user, pw)
        if role is not None:
            self.err_var.set("")
            self.on_login(user, role)
        else:
            self.err_var.set("✗  Invalid username or password.")
            self.password_var.set("")
            self.password_entry.focus_set()


# ══════════════════════════════════════════════════════════════════════════════
# SCREEN 2 — DASHBOARD  (dropdown selector)
# ══════════════════════════════════════════════════════════════════════════════
class DashboardScreen(tk.Frame):
    def __init__(self, master, username, role, on_select, on_logout, on_manage_accounts):
        super().__init__(master, bg=BG)
        self.username  = username
        self.role      = role
        self.on_select = on_select
        self.on_logout = on_logout
        self.on_manage_accounts = on_manage_accounts
        self._build()

    def _build(self):
        # Header
        header = tk.Frame(self, bg=HEADER_BG, height=60)
        header.pack(fill="x")
        header.pack_propagate(False)

        tk.Label(header, text="⬡  Operations Portal",
                 font=("Segoe UI", 14, "bold"), bg=HEADER_BG, fg=TEXT
                 ).pack(side="left", padx=20, pady=14)

        # User badge + logout
        right_hdr = tk.Frame(header, bg=HEADER_BG)
        right_hdr.pack(side="right", padx=16)

        role_color = GOLD if self.role == "admin" else ACCENT
        role_icon  = "👑" if self.role == "admin" else "👤"
        tk.Label(right_hdr, text=f"{role_icon}  {self.username}",
                 font=("Segoe UI", 9, "bold"), bg=HEADER_BG, fg=role_color
                 ).pack(side="left", padx=(0, 12))

        if self.role == "admin":
            tk.Button(right_hdr, text="⚙ Manage Accounts",
                      font=("Segoe UI", 8, "bold"),
                      bg=ACCENT2, fg="white", relief="flat",
                      padx=10, pady=4, cursor="hand2",
                      command=self.on_manage_accounts
                      ).pack(side="left", padx=(0, 8))

        tk.Button(right_hdr, text="Log Out",
                  font=("Segoe UI", 8, "bold"),
                  bg=DANGER, fg="white", relief="flat",
                  padx=12, pady=4, cursor="hand2",
                  command=self.on_logout
                  ).pack(side="left")

        # Thin accent stripe
        tk.Frame(self, bg=ACCENT, height=3).pack(fill="x")

        # ── Center content ────────────────────────────────────────────────────
        outer = tk.Frame(self, bg=BG)
        outer.pack(expand=True)

        card = tk.Frame(outer, bg=CARD, padx=56, pady=52,
                        highlightthickness=1, highlightbackground=BORDER)
        card.pack(padx=20, pady=20)

        tk.Label(card, text="Select Your Platform",
                 font=("Segoe UI", 18, "bold"), bg=CARD, fg=TEXT).pack()
        tk.Label(card, text="Choose the client account you want to work with",
                 font=("Segoe UI", 9), bg=CARD, fg=SUBTEXT).pack(pady=(4, 32))

        # Dropdown label
        tk.Label(card, text="CLIENT ACCOUNT", font=("Segoe UI", 8, "bold"),
                 bg=CARD, fg=SUBTEXT).pack(anchor="w")

        # Custom styled combobox container
        combo_wrapper = tk.Frame(card, bg=BORDER, padx=1, pady=1)
        combo_wrapper.pack(fill="x", pady=(4, 24))

        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Dark.TCombobox",
                        fieldbackground=CARD,
                        background=CARD,
                        foreground=TEXT,
                        arrowcolor=ACCENT,
                        bordercolor=BORDER,
                        relief="flat",
                        padding=10,
                        font=("Segoe UI", 11))
        style.map("Dark.TCombobox",
                  fieldbackground=[("readonly", CARD)],
                  selectbackground=[("readonly", CARD)],
                  selectforeground=[("readonly", TEXT)])

        self.selected = tk.StringVar(value="— Select an account —")
        combo = ttk.Combobox(
            combo_wrapper,
            textvariable=self.selected,
            values=["Orico Auto Loan", "Bank of Makati"],
            state="readonly",
            style="Dark.TCombobox",
            font=("Segoe UI", 11),
        )
        combo.pack(fill="x")

        # Error label
        self.err_var = tk.StringVar()
        tk.Label(card, textvariable=self.err_var,
                 font=("Segoe UI", 9), bg=CARD, fg=DANGER).pack(pady=(0, 8))

        # Proceed button
        tk.Button(card, text="Proceed  →",
                  font=("Segoe UI", 11, "bold"),
                  bg=ACCENT, fg="white", relief="flat",
                  pady=11, cursor="hand2",
                  activebackground="#3a7be8",
                  command=self._proceed
                  ).pack(fill="x")

        # Quick-info chips
        chips_frame = tk.Frame(card, bg=CARD)
        chips_frame.pack(pady=(28, 0))
        for label, color in [("Orico Auto Loan", ACCENT), ("Bank of Makati", ACCENT2)]:
            chip = tk.Frame(chips_frame, bg=color, padx=14, pady=5)
            chip.pack(side="left", padx=6)
            tk.Label(chip, text=label, font=("Segoe UI", 9, "bold"),
                     bg=color, fg="white").pack()

        # Footer
        footer = tk.Frame(self, bg=DARK)
        footer.pack(fill="x", side="bottom")
        tk.Label(footer, text="Created by  Vincent Corocoto  ·  09567796275",
                 font=("Segoe UI", 8, "bold"), bg=DARK, fg=ACCENT).pack(side="left", padx=16, pady=8)
        tk.Label(footer, text='"Kapag ang palay naging bigas, May bumayo."',
                 font=("Segoe UI", 8, "italic"), bg=DARK, fg="#4A5568").pack(side="right", padx=16, pady=8)

    def _proceed(self):
        val = self.selected.get()
        if val.startswith("—"):
            self.err_var.set("⚠  Please select an account first.")
            return
        self.err_var.set("")
        self.on_select(val)


# ══════════════════════════════════════════════════════════════════════════════
# SCREEN 3 — ORICO TOOLS  (clickable tool cards)
# ══════════════════════════════════════════════════════════════════════════════
class OricoToolsScreen(tk.Frame):
    def __init__(self, master, username, role="user", on_tool=None, on_back=None):
        super().__init__(master, bg=BG)
        self.username = username
        self.role     = role
        self.on_tool  = on_tool
        self.on_back  = on_back
        self._build()

    def _build(self):
        # Header
        header = tk.Frame(self, bg=HEADER_BG, height=60)
        header.pack(fill="x")
        header.pack_propagate(False)

        left_hdr = tk.Frame(header, bg=HEADER_BG)
        left_hdr.pack(side="left", padx=12, pady=8)
        tk.Button(left_hdr, text="←  Back",
                  font=("Segoe UI", 9, "bold"),
                  bg=CARD, fg=TEXT, relief="flat",
                  padx=12, pady=6, cursor="hand2",
                  command=self.on_back).pack(side="left")
        tk.Label(left_hdr, text="  ⬡  Orico Auto Loan",
                 font=("Segoe UI", 13, "bold"), bg=HEADER_BG, fg=TEXT).pack(side="left", padx=8)

        right_hdr = tk.Frame(header, bg=HEADER_BG)
        right_hdr.pack(side="right", padx=16)
        role_colors = {"admin": DANGER, "supervisor": WARNING, "user": SUCCESS}
        role_color  = role_colors.get(self.role, ACCENT)
        tk.Label(right_hdr, text=f"👤  {self.username}",
                 font=("Segoe UI", 9, "bold"), bg=HEADER_BG, fg=ACCENT).pack()
        tk.Label(right_hdr, text=f"● {self.role.upper()}",
                 font=("Segoe UI", 7, "bold"), bg=HEADER_BG, fg=role_color).pack()

        tk.Frame(self, bg=ACCENT, height=3).pack(fill="x")

        # Subtitle
        sub = tk.Frame(self, bg=BG)
        sub.pack(pady=(32, 8))
        tk.Label(sub, text="Tools & Utilities",
                 font=("Segoe UI", 18, "bold"), bg=BG, fg=TEXT).pack()
        tk.Label(sub, text="Select a tool to get started",
                 font=("Segoe UI", 9), bg=BG, fg=SUBTEXT).pack(pady=4)

        # Tool cards grid
        grid = tk.Frame(self, bg=BG)
        grid.pack(expand=True)

        tools = [
            {
                "name":    "DRR Cleaner",
                "icon":    "🧹",
                "desc":    "Clean & process Daily\nRemittance Reports",
                "color":   ACCENT,
                "tag":     "drr_cleaner",
                "badge":   "Excel",
            },
            {
                "name":    "Coming Soon",
                "icon":    "🔧",
                "desc":    "More tools will be\nadded here",
                "color":   "#3D4153",
                "tag":     None,
                "badge":   "Soon",
            },
            {
                "name":    "Coming Soon",
                "icon":    "📊",
                "desc":    "Analytics &\nReporting",
                "color":   "#3D4153",
                "tag":     None,
                "badge":   "Soon",
            },
        ]

        for i, tool in enumerate(tools):
            col_frame = tk.Frame(grid, bg=BG)
            col_frame.grid(row=0, column=i, padx=14)
            self._tool_card(col_frame, tool)

        # Footer
        footer = tk.Frame(self, bg=DARK)
        footer.pack(fill="x", side="bottom")
        tk.Label(footer, text="Created by  Vincent Corocoto  ·  09567796275",
                 font=("Segoe UI", 8, "bold"), bg=DARK, fg=ACCENT).pack(side="left", padx=16, pady=8)
        tk.Label(footer, text='"Kapag ang palay naging bigas, May bumayo."',
                 font=("Segoe UI", 8, "italic"), bg=DARK, fg="#4A5568").pack(side="right", padx=16, pady=8)

    def _tool_card(self, parent, tool):
        is_active = tool["tag"] is not None
        border_color = tool["color"] if is_active else BORDER

        card = tk.Frame(parent, bg=CARD, width=200, height=220,
                        highlightthickness=2,
                        highlightbackground=border_color,
                        cursor="hand2" if is_active else "arrow")
        card.pack()
        card.pack_propagate(False)

        inner = tk.Frame(card, bg=CARD, padx=20, pady=20)
        inner.pack(fill="both", expand=True)

        # Badge
        badge_color = tool["color"] if is_active else "#3D4153"
        badge = tk.Frame(inner, bg=badge_color, padx=8, pady=2)
        badge.pack(anchor="e")
        tk.Label(badge, text=tool["badge"], font=("Segoe UI", 7, "bold"),
                 bg=badge_color, fg="white").pack()

        # Icon
        tk.Label(inner, text=tool["icon"], font=("Segoe UI", 34),
                 bg=CARD, fg=tool["color"] if is_active else SUBTEXT).pack(pady=(8, 8))

        # Name
        tk.Label(inner, text=tool["name"], font=("Segoe UI", 12, "bold"),
                 bg=CARD, fg=TEXT if is_active else SUBTEXT).pack()

        # Description
        tk.Label(inner, text=tool["desc"], font=("Segoe UI", 8),
                 bg=CARD, fg=SUBTEXT, justify="center").pack(pady=(4, 0))

        # Hover effects for active cards
        if is_active:
            tag = tool["tag"]

            def on_enter(e, c=card, col=tool["color"]):
                c.config(highlightbackground=col, bg="#22263A")
                for w in c.winfo_children():
                    _set_bg_recursive(w, "#22263A")

            def on_leave(e, c=card, col=tool["color"]):
                c.config(highlightbackground=col, bg=CARD)
                for w in c.winfo_children():
                    _set_bg_recursive(w, CARD)

            def on_click(e, t=tag):
                self.on_tool(t)

            for widget in [card, inner] + _all_children(inner):
                widget.bind("<Enter>", on_enter)
                widget.bind("<Leave>", on_leave)
                widget.bind("<Button-1>", on_click)


def _all_children(widget):
    result = []
    for child in widget.winfo_children():
        result.append(child)
        result.extend(_all_children(child))
    return result


def _set_bg_recursive(widget, color):
    try:
        widget.config(bg=color)
    except Exception:
        pass
    for child in widget.winfo_children():
        _set_bg_recursive(child, color)


# ══════════════════════════════════════════════════════════════════════════════
# SCREEN 4 — DRR CLEANER APP
# ══════════════════════════════════════════════════════════════════════════════
class DRRCleanerScreen(tk.Frame):
    PAGE_SIZE = 2000

    def __init__(self, master, username, role="user", on_back=None):
        super().__init__(master, bg=BG)
        self.username           = username
        self.role               = role
        self.on_back            = on_back
        self.file_path          = None
        self.output_bytes       = None
        self.cleaned_df         = None
        self.removed_df         = None
        self.removed_reason_df  = None
        self.remarks_df         = None
        self._current_df        = None
        self._current_page      = 0
        self._build_ui()

    def _build_ui(self):
        # ── Header ────────────────────────────────────────────────────────────
        header = tk.Frame(self, bg=HEADER_BG, height=60)
        header.pack(fill="x")
        header.pack_propagate(False)

        left_hdr = tk.Frame(header, bg=HEADER_BG)
        left_hdr.pack(side="left", padx=12, pady=8)
        tk.Button(left_hdr, text="←  Back",
                  font=("Segoe UI", 9, "bold"),
                  bg=CARD, fg=TEXT, relief="flat",
                  padx=12, pady=6, cursor="hand2",
                  command=self.on_back).pack(side="left")
        tk.Label(left_hdr, text="  🧹  DRR Cleaner  —  Orico Auto Loan",
                 font=("Segoe UI", 13, "bold"), bg=HEADER_BG, fg=TEXT).pack(side="left", padx=8)
        tk.Label(left_hdr, text="Drop · Preview · Verify · Download",
                 font=("Segoe UI", 8), bg=HEADER_BG, fg=SUBTEXT).pack(side="left", padx=4)

        right_hdr = tk.Frame(header, bg=HEADER_BG)
        right_hdr.pack(side="right", padx=16)
        role_colors = {"admin": DANGER, "supervisor": WARNING, "user": SUCCESS}
        role_color  = role_colors.get(self.role, ACCENT)
        tk.Label(right_hdr, text=f"👤  {self.username}",
                 font=("Segoe UI", 9, "bold"), bg=HEADER_BG, fg=ACCENT).pack()
        tk.Label(right_hdr, text=f"● {self.role.upper()}",
                 font=("Segoe UI", 7, "bold"), bg=HEADER_BG, fg=role_color).pack()

        tk.Frame(self, bg=ACCENT, height=3).pack(fill="x")

        # ── Body ──────────────────────────────────────────────────────────────
        body = tk.Frame(self, bg=BG)
        body.pack(fill="both", expand=True, padx=20, pady=16)

        # ── LEFT PANEL ────────────────────────────────────────────────────────
        left = tk.Frame(body, bg=BG, width=300)
        left.pack(side="left", fill="y", padx=(0, 14))
        left.pack_propagate(False)

        # ── Drop Zone ─────────────────────────────────────────────────────────
        self._drop_frame = tk.Frame(left, bg=CARD, highlightthickness=2,
                                    highlightbackground=BORDER)
        self._drop_frame.pack(fill="x", pady=(0, 10))

        # ── State A: empty drop prompt ────────────────────────────────────────
        self._drop_empty = tk.Frame(self._drop_frame, bg=CARD)
        self._drop_empty.pack(fill="both", padx=20, pady=20)

        self._drop_icon  = tk.Label(self._drop_empty, text="📂", font=("Segoe UI", 26),
                                    bg=CARD, fg=ACCENT)
        self._drop_icon.pack()
        self._drop_title = tk.Label(self._drop_empty, text="Drop Excel file here",
                                    font=("Segoe UI", 11, "bold"), bg=CARD, fg=TEXT)
        self._drop_title.pack(pady=(6, 2))
        tk.Label(self._drop_empty, text=".xlsx  ·  .xlsm  ·  .xls",
                 font=("Segoe UI", 8), bg=CARD, fg=SUBTEXT).pack()
        tk.Label(self._drop_empty, text="or", font=("Segoe UI", 8),
                 bg=CARD, fg=SUBTEXT).pack(pady=(8, 4))
        tk.Button(self._drop_empty, text="Browse File", font=("Segoe UI", 9, "bold"),
                  bg=ACCENT, fg="white", relief="flat", padx=16, pady=6,
                  cursor="hand2", command=self._browse).pack()

        # ── State B: file loaded display (hidden until a file is picked) ──────
        self._drop_loaded = tk.Frame(self._drop_frame, bg=CARD)
        # NOT packed yet — shown in _load_file

        # Top: green checkmark banner
        loaded_top = tk.Frame(self._drop_loaded, bg="#1A3A2A")
        loaded_top.pack(fill="x")
        tk.Label(loaded_top, text="✅  File Loaded", font=("Segoe UI", 9, "bold"),
                 bg="#1A3A2A", fg=SUCCESS).pack(side="left", padx=12, pady=6)
        tk.Button(loaded_top, text="✕  Clear", font=("Segoe UI", 8),
                  bg="#1A3A2A", fg=SUBTEXT, relief="flat",
                  cursor="hand2", command=self._clear_file
                  ).pack(side="right", padx=8)

        # File icon + details
        loaded_body = tk.Frame(self._drop_loaded, bg=CARD)
        loaded_body.pack(fill="x", padx=14, pady=12)

        tk.Label(loaded_body, text="📄", font=("Segoe UI", 22),
                 bg=CARD, fg=ACCENT).pack(side="left", padx=(0, 10))

        loaded_text = tk.Frame(loaded_body, bg=CARD)
        loaded_text.pack(side="left", fill="x", expand=True)
        self._fi_name = tk.Label(loaded_text, text="", font=("Segoe UI", 9, "bold"),
                                  bg=CARD, fg=TEXT, anchor="w", wraplength=180, justify="left")
        self._fi_name.pack(fill="x")
        self._fi_meta = tk.Label(loaded_text, text="", font=("Segoe UI", 7),
                                  bg=CARD, fg=SUBTEXT, anchor="w")
        self._fi_meta.pack(fill="x", pady=(2, 0))

        # Change file button
        tk.Button(self._drop_loaded, text="🔄  Change File",
                  font=("Segoe UI", 8), bg=CARD, fg=SUBTEXT,
                  relief="flat", cursor="hand2",
                  command=self._browse).pack(pady=(0, 10))

        # Register DnD on the frame AND every child widget so the whole zone is a target
        self._dnd_enabled = False
        self._dnd_widgets = []
        try:
            for widget in [self._drop_frame, self._drop_empty, self._drop_loaded,
                           self._drop_icon, self._drop_title, loaded_top, loaded_body]:
                widget.drop_target_register('DND_Files')
                widget.dnd_bind('<<Drop>>',      self._on_drop)
                widget.dnd_bind('<<DragEnter>>', self._on_drag_enter)
                widget.dnd_bind('<<DragLeave>>', self._on_drag_leave)
                self._dnd_widgets.append(widget)
            self._dnd_enabled = True
        except Exception:
            pass

        # Legacy label kept for compatibility (hidden)
        self.file_label = tk.Label(left, text="", font=("Segoe UI", 8),
                                   bg=BG, fg=SUBTEXT, wraplength=270, justify="left")


        tk.Label(left, text="SUMMARY", font=("Segoe UI", 8, "bold"),
                 bg=BG, fg=SUBTEXT).pack(anchor="w", pady=(0, 6))

        self.stat_vars = {}
        for key, label, color, icon in [
            ("total",    "Total Rows",    TEXT,    "📋"),
            ("retained", "Rows Retained", SUCCESS, "✅"),
            ("removed",  "Rows Removed",  DANGER,  "🗑"),
            ("srp",      "Remarks Fixed", WARNING, "✏️"),
        ]:
            c = tk.Frame(left, bg=CARD, highlightthickness=1, highlightbackground=BORDER)
            c.pack(fill="x", pady=3)
            inn = tk.Frame(c, bg=CARD)
            inn.pack(fill="x", padx=12, pady=8)
            tk.Label(inn, text=icon, font=("Segoe UI", 11), bg=CARD, fg=color).pack(side="left")
            tk.Label(inn, text=label, font=("Segoe UI", 9), bg=CARD, fg=SUBTEXT).pack(side="left", padx=8)
            var = tk.StringVar(value="—")
            self.stat_vars[key] = var
            tk.Label(inn, textvariable=var, font=("Segoe UI", 11, "bold"),
                     bg=CARD, fg=color).pack(side="right")

        # Supervisor: can view + download but NOT upload/clean
        # User: can upload + clean but NOT download
        can_upload   = self.role in ("admin", "user")
        can_download = self.role in ("admin", "supervisor")

        self.process_btn = tk.Button(left, text="▶  Run Cleaner",
                                     font=("Segoe UI", 10, "bold"), bg=ACCENT, fg="white",
                                     relief="flat", pady=10, cursor="hand2",
                                     state="disabled" if can_upload else "disabled",
                                     command=self._run_process)
        self.process_btn.pack(fill="x", pady=(14, 6))
        if not can_upload:
            self.process_btn.config(state="disabled", bg=CARD, fg=SUBTEXT,
                                    text="▶  Run Cleaner  (No Access)")

        self.download_btn = tk.Button(left, text="⬇  Download Output",
                                      font=("Segoe UI", 9, "bold"),
                                      bg=SUCCESS if can_download else CARD,
                                      fg="white" if can_download else SUBTEXT,
                                      relief="flat", pady=10, cursor="hand2",
                                      state="disabled", command=self._download)
        self.download_btn.pack(fill="x")
        if not can_download:
            self.download_btn.config(text="⬇  Download  (No Access)")

        style = ttk.Style()
        style.configure("TProgressbar", troughcolor=CARD, background=ACCENT, thickness=4)
        self.progress = ttk.Progressbar(left, mode="indeterminate")
        self.progress.pack(fill="x", pady=(10, 0))

        # ── RIGHT PANEL ───────────────────────────────────────────────────────
        right = tk.Frame(body, bg=BG)
        right.pack(side="left", fill="both", expand=True)

        tab_bar = tk.Frame(right, bg=BG)
        tab_bar.pack(fill="x", pady=(0, 8))

        self.active_tab = tk.StringVar(value="cleaned")
        self.tab_btns   = {}
        for tab_id, tab_label, tab_color in [
            ("cleaned",        "✅ Cleaned Rows",    ACCENT),
            ("removed",        "🗑 Removed Rows",    CARD),
            ("removed_reason", "📋 Removed Status",  CARD),
            ("remarks",        "✏️ Remarks Changes", CARD),
        ]:
            btn = tk.Button(tab_bar, text=tab_label,
                            font=("Segoe UI", 9, "bold"),
                            bg=tab_color, fg="white", relief="flat",
                            padx=14, pady=7, cursor="hand2",
                            command=lambda t=tab_id: self._switch_tab(t))
            btn.pack(side="left", padx=(0, 6))
            self.tab_btns[tab_id] = btn

        # ── Search bar ────────────────────────────────────────────────────────
        search_frame = tk.Frame(right, bg=CARD, highlightthickness=1, highlightbackground=BORDER)
        search_frame.pack(fill="x", pady=(0, 4))
        tk.Label(search_frame, text="🔍", bg=CARD, fg=SUBTEXT, font=("Segoe UI", 10)).pack(side="left", padx=8)
        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", self._on_search)
        tk.Entry(search_frame, textvariable=self.search_var, font=("Segoe UI", 9),
                 bg=CARD, fg=TEXT, insertbackground=TEXT, relief="flat", bd=0
                 ).pack(side="left", fill="x", expand=True, pady=8, padx=4)
        tk.Label(search_frame, text="Search all columns…", bg=CARD, fg=SUBTEXT,
                 font=("Segoe UI", 8)).pack(side="right", padx=8)

        # ── Column Dropdown Filter bar ─────────────────────────────────────────
        filter_outer = tk.Frame(right, bg=CARD, highlightthickness=1, highlightbackground=BORDER)
        filter_outer.pack(fill="x", pady=(0, 8))
        tk.Label(filter_outer, text="▼ Column Filter:", bg=CARD, fg=SUBTEXT,
                 font=("Segoe UI", 8, "bold")).pack(side="left", padx=8, pady=6)
        self._filter_col_var = tk.StringVar(value="— Select Column —")
        self._filter_val_var = tk.StringVar(value="— Select Value —")
        self._filter_col_cb  = ttk.Combobox(filter_outer, textvariable=self._filter_col_var,
                                             state="readonly", width=24,
                                             font=("Segoe UI", 9))
        self._filter_col_cb.pack(side="left", padx=(0, 6), pady=6)
        self._filter_val_cb  = ttk.Combobox(filter_outer, textvariable=self._filter_val_var,
                                             state="readonly", width=28,
                                             font=("Segoe UI", 9))
        self._filter_val_cb.pack(side="left", padx=(0, 6), pady=6)
        self._filter_col_cb.bind("<<ComboboxSelected>>", self._on_filter_col_change)
        self._filter_val_cb.bind("<<ComboboxSelected>>", self._on_filter_apply)
        tk.Button(filter_outer, text="✕ Clear Filter", font=("Segoe UI", 8, "bold"),
                  bg=DANGER, fg="white", relief="flat", padx=10, pady=4,
                  cursor="hand2", command=self._clear_filter).pack(side="left", padx=(0, 6))
        self._filter_active = False

        table_frame = tk.Frame(right, bg=BG)
        table_frame.pack(fill="both", expand=True)
        vsb = ttk.Scrollbar(table_frame, orient="vertical")
        hsb = ttk.Scrollbar(table_frame, orient="horizontal")
        vsb.pack(side="right", fill="y")
        hsb.pack(side="bottom", fill="x")

        style.configure("Custom.Treeview", background=CARD, foreground=TEXT,
                         fieldbackground=CARD, rowheight=28, font=("Segoe UI", 9), borderwidth=0)
        style.configure("Custom.Treeview.Heading", background=HEADER_BG, foreground=TEXT,
                         font=("Segoe UI", 9, "bold"), relief="flat")
        style.map("Custom.Treeview",
                  background=[("selected", ACCENT2)], foreground=[("selected", "white")])

        self.tree = ttk.Treeview(table_frame, style="Custom.Treeview",
                                  yscrollcommand=vsb.set, xscrollcommand=hsb.set,
                                  show="headings", selectmode="browse")
        self.tree.pack(fill="both", expand=True)
        vsb.config(command=self.tree.yview)
        hsb.config(command=self.tree.xview)
        self.tree.bind("<<TreeviewSelect>>", self._on_cell_click)
        self.tree.bind("<ButtonRelease-1>",  self._on_tree_click)
        self.tree.bind("<Double-1>",         self._on_double_click)
        self.tree.bind("<Return>",           self._on_double_click)

        self.row_count_label = tk.Label(right, text="", font=("Segoe UI", 8), bg=BG, fg=SUBTEXT)
        self.row_count_label.pack(anchor="e", pady=(4, 0))

        # ── Excel-style formula bar ───────────────────────────────────────────
        fbar = tk.Frame(right, bg="#1C1F2E",
                        highlightthickness=1, highlightbackground=BORDER)
        fbar.pack(fill="x", pady=(4, 2))

        # fx badge
        tk.Label(fbar, text=" fx ", font=("Segoe UI", 9, "bold"),
                 bg="#2E7D32", fg="white", padx=6, pady=4
                 ).pack(side="left")

        # Column name box
        self._fbar_col_var = tk.StringVar(value="")
        col_box = tk.Entry(fbar, textvariable=self._fbar_col_var,
                           font=("Segoe UI", 9, "bold"),
                           bg="#252840", fg=ACCENT,
                           relief="flat", bd=0,
                           width=22, justify="center",
                           insertbackground=ACCENT,
                           readonlybackground="#252840",
                           state="readonly")
        col_box.pack(side="left", padx=(1, 0), ipady=4)

        # Divider
        tk.Frame(fbar, bg=BORDER, width=1).pack(side="left", fill="y", pady=2)

        # Value text — full, selectable, scrollable
        self._fbar_val = tk.Text(fbar, font=("Segoe UI", 9),
                                 bg="#1C1F2E", fg=TEXT,
                                 relief="flat", bd=0,
                                 height=1, wrap="none",
                                 cursor="xterm",
                                 selectbackground=ACCENT,
                                 selectforeground="white",
                                 insertbackground=TEXT,
                                 state="disabled")
        fbar_xsb = ttk.Scrollbar(fbar, orient="horizontal",
                                  command=self._fbar_val.xview)
        self._fbar_val.configure(xscrollcommand=fbar_xsb.set)
        self._fbar_val.pack(side="left", fill="x", expand=True, padx=(6, 0), pady=2)
        # Copy button on right
        tk.Button(fbar, text="⎘", font=("Segoe UI", 10),
                  bg="#1C1F2E", fg=SUBTEXT, relief="flat",
                  cursor="hand2", padx=6,
                  command=self._fbar_copy
                  ).pack(side="right", padx=4)

        nav_frame = tk.Frame(right, bg=BG)
        nav_frame.pack(fill="x", pady=(4, 0))
        self.prev_btn = tk.Button(nav_frame, text="◀  Prev", font=("Segoe UI", 8, "bold"),
                                   bg=CARD, fg=TEXT, relief="flat", padx=12, pady=4,
                                   cursor="hand2", state="disabled", command=self._prev_page)
        self.prev_btn.pack(side="left", padx=(0, 6))
        self.page_label = tk.Label(nav_frame, text="", font=("Segoe UI", 8), bg=BG, fg=SUBTEXT)
        self.page_label.pack(side="left")
        self.next_btn = tk.Button(nav_frame, text="Next  ▶", font=("Segoe UI", 8, "bold"),
                                   bg=CARD, fg=TEXT, relief="flat", padx=12, pady=4,
                                   cursor="hand2", state="disabled", command=self._next_page)
        self.next_btn.pack(side="left", padx=6)

        # Status bar
        self.status_var = tk.StringVar(value="Ready — drop or browse an Excel file to begin.")
        tk.Label(self, textvariable=self.status_var, font=("Segoe UI", 8),
                 bg=HEADER_BG, fg=SUBTEXT, anchor="w", padx=16, pady=6).pack(fill="x", side="bottom")

        # Footer
        watermark = tk.Frame(self, bg=DARK)
        watermark.pack(fill="x", side="bottom")
        tk.Label(watermark, text="Created by  Vincent Corocoto  ·  09567796275",
                 font=("Segoe UI", 8, "bold"), bg=DARK, fg=ACCENT).pack(side="left", padx=16, pady=6)
        tk.Label(watermark, text='"Kapag ang palay naging bigas, May bumayo."',
                 font=("Segoe UI", 8, "italic"), bg=DARK, fg="#4A5568").pack(side="right", padx=16, pady=6)

    # ── FILE LOAD ─────────────────────────────────────────────────────────────
    def _on_drag_enter(self, event):
        self._drop_frame.config(highlightbackground=ACCENT, highlightthickness=2,
                                bg="#1A2540")
        self._drop_icon.config(bg="#1A2540", fg=ACCENT, text="⬇️")
        self._drop_title.config(bg="#1A2540", fg=ACCENT, text="Release to load file")
        for w in self._drop_frame.winfo_children():
            try: w.config(bg="#1A2540")
            except Exception: pass

    def _on_drag_leave(self, event):
        self._drop_frame.config(highlightbackground=BORDER, bg=CARD)
        self._drop_icon.config(bg=CARD, fg=ACCENT, text="📂")
        self._drop_title.config(bg=CARD, fg=TEXT, text="Drop Excel file here")
        for w in self._drop_frame.winfo_children():
            try: w.config(bg=CARD)
            except Exception: pass

    def _on_drop(self, event):
        self._on_drag_leave(None)
        raw = event.data.strip()
        # tkinterdnd2 may wrap paths with braces when they contain spaces
        paths = []
        if raw.startswith("{"):
            import re
            paths = re.findall(r'\{([^}]+)\}', raw)
        if not paths:
            paths = [raw]
        self._load_file(paths[0])

    def _browse(self):
        path = filedialog.askopenfilename(title="Select Excel file",
                                          filetypes=[("Excel files", "*.xlsx *.xlsm *.xls")])
        if path:
            self._load_file(path)

    def _clear_file(self):
        self.file_path = None
        self._drop_loaded.pack_forget()
        self._drop_empty.pack(fill="both", padx=20, pady=20)
        self._drop_frame.config(highlightbackground=BORDER)
        self.process_btn.config(state="disabled")
        self.download_btn.config(state="disabled")
        self.status_var.set("Ready — drop or browse an Excel file to begin.")
        self._clear_table()
        for k in self.stat_vars:
            self.stat_vars[k].set("—")

    def _load_file(self, path):
        if not os.path.exists(path):
            messagebox.showerror("File not found", f"Cannot find:\n{path}")
            return
        self.file_path = path
        fname = os.path.basename(path)
        fsize = os.path.getsize(path)
        fsize_str = (f"{fsize / 1024:.1f} KB" if fsize < 1_048_576
                     else f"{fsize / 1_048_576:.2f} MB")
        import time
        mtime     = os.path.getmtime(path)
        mtime_str = time.strftime("%b %d, %Y  %I:%M %p", time.localtime(mtime))

        # Swap drop zone: hide empty state, show loaded state
        self._drop_empty.pack_forget()
        self._fi_name.config(text=fname)
        self._fi_meta.config(text=f"{fsize_str}  ·  {mtime_str}")
        self._drop_loaded.pack(fill="both")
        self._drop_frame.config(highlightbackground=SUCCESS)

        self.file_label.config(text="")
        self.process_btn.config(state="normal" if self.role in ("admin", "user") else "disabled")
        self.download_btn.config(state="disabled")
        self.status_var.set(f"Loaded: {fname}  —  Click ▶ Run Cleaner to process.")
        self._clear_table()
        for k in self.stat_vars:
            self.stat_vars[k].set("—")

    # ── PROCESSING ────────────────────────────────────────────────────────────
    def _run_process(self):
        if self.role not in ("admin", "user"):
            messagebox.showwarning("Access Denied", "Your role does not have upload/clean access.")
            return
        if not self.file_path:
            return
        self.process_btn.config(state="disabled")
        self.download_btn.config(state="disabled")
        self.status_var.set("Processing…")
        self.progress.start(10)
        threading.Thread(target=self._process_thread, daemon=True).start()

    def _process_thread(self):
        try:
            result = process_file(self.file_path)
            self.after(0, lambda: self._on_done(*result))
        except Exception as e:
            self.after(0, lambda: self._on_error(str(e)))

    def _on_done(self, cleaned, removed, remarks, stats, out_bytes, col_e, dialed_col):
        self.progress.stop()
        self.cleaned_df        = cleaned
        self.removed_df        = removed
        if "Removed Reason" in removed.columns:
            status_col = [c for c in removed.columns if c != "Removed Reason"]
            mask = removed["Removed Reason"].str.startswith("Status: ")
            self.removed_reason_df = removed[mask][["Removed Reason"] + status_col].copy()
        else:
            self.removed_reason_df = removed.copy()
        self.remarks_df        = remarks
        self.output_bytes      = out_bytes

        self.stat_vars["total"].set(str(stats["total"]))
        self.stat_vars["retained"].set(str(stats["retained"]))
        self.stat_vars["removed"].set(str(stats["removed"]))
        self.stat_vars["srp"].set(str(stats["srp_changed"]))

        self.download_btn.config(state="normal")
        self.process_btn.config(state="normal" if self.role in ("admin", "user") else "disabled")
        self.status_var.set(
            f"Done  ·  {stats['retained']} retained  ·  {stats['removed']} removed  ·  "
            f"{stats['srp_changed']} remarks → SRP"
        )
        self._switch_tab("cleaned")
        show_toast(self.master, "Processing complete!", color=SUCCESS)

    def _on_error(self, msg):
        self.progress.stop()
        self.process_btn.config(state="normal" if self.role in ("admin", "user") else "disabled")
        self.status_var.set(f"Error: {msg}")
        messagebox.showerror("Processing Error", msg)

    # ── TABS & TABLE ──────────────────────────────────────────────────────────
    def _switch_tab(self, tab_id):
        self.active_tab.set(tab_id)
        colors = {k: CARD for k in self.tab_btns}
        colors[tab_id] = ACCENT2 if tab_id in ("remarks", "removed_reason") else ACCENT
        for tid, btn in self.tab_btns.items():
            btn.config(bg=colors[tid])
        df_map = {
            "cleaned":        self.cleaned_df,
            "removed":        self.removed_df,
            "removed_reason": self.removed_reason_df,
            "remarks":        self.remarks_df,
        }
        df = df_map.get(tab_id)
        empty_msgs = {
            "remarks":        "No remark changes found.",
            "removed_reason": "No removed rows found.",
            "removed":        "No removed rows.",
            "cleaned":        "No retained rows.",
        }
        # Refresh column filter dropdown for the new tab
        if df is not None and not (hasattr(df, "empty") and df.empty):
            self._filter_col_cb["values"] = ["— Select Column —"] + list(df.columns)
            self._filter_col_var.set("— Select Column —")
            self._filter_val_var.set("— Select Value —")
            self._filter_val_cb["values"] = []
            self._filter_active = False
            self._populate_table(df)
        else:
            self._clear_table()
            self.row_count_label.config(text=empty_msgs.get(tab_id, "No rows."))

    def _populate_table(self, df, filter_text=""):
        self._clear_table()
        if df is None or df.empty:
            self.row_count_label.config(text="No rows to display.")
            return
        cols = list(df.columns)
        self.tree["columns"] = cols
        for col in cols:
            sample    = df[col].astype(str)
            max_len   = max(len(str(col)), int(sample.str.len().quantile(0.90)) if len(df) > 0 else 10)
            col_width = max(100, min(max_len * 8, 280))
            self.tree.heading(col, text=col, anchor="w",
                              command=lambda c=col: self._sort_col(c))
            self.tree.column(col, width=col_width, minwidth=60, anchor="w", stretch=False)
        filt = filter_text.lower().strip()
        if filt:
            combined = df.fillna("").astype(str).agg(" ".join, axis=1).str.lower()
            filtered = df[combined.str.contains(filt, regex=False)].reset_index(drop=True)
        else:
            filtered = df.reset_index(drop=True)
        self._current_df   = filtered
        self._current_page = 0
        self._render_page()

    def _render_page(self):
        self.tree.delete(*self.tree.get_children())
        df    = self._current_df
        page  = self._current_page
        start = page * self.PAGE_SIZE
        end   = min(start + self.PAGE_SIZE, len(df))
        str_df = df.iloc[start:end].fillna("").astype(str)
        for i, vals in enumerate(str_df.values.tolist()):
            tag = "even" if i % 2 == 0 else "odd"
            self.tree.insert("", "end", values=vals, tags=(tag,))
        self.tree.tag_configure("even", background=CARD)
        self.tree.tag_configure("odd",  background="#14172A")
        total_pages = max(1, (len(df) + self.PAGE_SIZE - 1) // self.PAGE_SIZE)
        self.row_count_label.config(
            text=(f"Page {page + 1} of {total_pages}  ·  "
                  f"Showing rows {start + 1}–{end} of {len(df)}  ·  "
                  f"{len(df.columns)} columns  ·  ← → to navigate pages")
        )
        if hasattr(self, "prev_btn"):
            self.prev_btn.config(state="normal" if page > 0 else "disabled")
            self.next_btn.config(state="normal" if page < total_pages - 1 else "disabled")
            self.page_label.config(text=f"Page {page + 1} / {total_pages}")

    def _sort_col(self, col):
        tab = self.active_tab.get()
        df  = {
            "cleaned":        self.cleaned_df,
            "removed":        self.removed_df,
            "removed_reason": self.removed_reason_df,
            "remarks":        self.remarks_df,
        }.get(tab)
        if df is None:
            return
        asc = getattr(self, "_sort_asc", {})
        ascending = not asc.get(col, True)
        asc[col] = ascending
        self._sort_asc = asc
        try:
            df_sorted = df.sort_values(by=col, ascending=ascending,
                                        key=lambda x: x.astype(str).str.lower())
        except Exception:
            df_sorted = df
        if tab == "cleaned":
            self.cleaned_df = df_sorted
        elif tab == "removed":
            self.removed_df = df_sorted
        else:
            self.remarks_df = df_sorted
        self._populate_table(df_sorted, self.search_var.get())

    def _on_filter_col_change(self, _=None):
        """When a column is chosen, populate the value dropdown with unique values."""
        tab = self.active_tab.get()
        df  = {"cleaned": self.cleaned_df, "removed": self.removed_df,
               "removed_reason": self.removed_reason_df, "remarks": self.remarks_df}.get(tab)
        if df is None:
            return
        col = self._filter_col_var.get()
        if col not in df.columns:
            return
        unique_vals = ["(All)"] + sorted(df[col].fillna("").astype(str).unique().tolist())
        self._filter_val_cb["values"] = unique_vals
        self._filter_val_var.set("— Select Value —")

    def _on_filter_apply(self, _=None):
        """Apply the column+value filter to the current tab."""
        tab = self.active_tab.get()
        df  = {"cleaned": self.cleaned_df, "removed": self.removed_df,
               "removed_reason": self.removed_reason_df, "remarks": self.remarks_df}.get(tab)
        if df is None:
            return
        col = self._filter_col_var.get()
        val = self._filter_val_var.get()
        if col not in df.columns or val in ("— Select Value —", "(All)"):
            self._populate_table(df, self.search_var.get())
            self._filter_active = False
            return
        filtered = df[df[col].fillna("").astype(str) == val].copy()
        self._filter_active = True
        self._populate_table(filtered, self.search_var.get())

    def _clear_filter(self):
        self._filter_col_var.set("— Select Column —")
        self._filter_val_var.set("— Select Value —")
        self._filter_val_cb["values"] = []
        self._filter_active = False
        tab = self.active_tab.get()
        df  = {"cleaned": self.cleaned_df, "removed": self.removed_df,
               "removed_reason": self.removed_reason_df, "remarks": self.remarks_df}.get(tab)
        if df is not None:
            self._populate_table(df, self.search_var.get())

    def _clear_table(self):
        self.tree.delete(*self.tree.get_children())
        self.tree["columns"] = []
        self.row_count_label.config(text="")
        self._current_df   = None
        self._current_page = 0
        if hasattr(self, "prev_btn"):
            self.prev_btn.config(state="disabled")
            self.next_btn.config(state="disabled")
            self.page_label.config(text="")

    def _prev_page(self):
        if self._current_page > 0:
            self._current_page -= 1
            self._render_page()

    def _next_page(self):
        if self._current_df is not None:
            total = (len(self._current_df) + self.PAGE_SIZE - 1) // self.PAGE_SIZE
            if self._current_page < total - 1:
                self._current_page += 1
                self._render_page()

    def _on_search(self, *_):
        tab = self.active_tab.get()
        df  = {
            "cleaned":        self.cleaned_df,
            "removed":        self.removed_df,
            "removed_reason": self.removed_reason_df,
            "remarks":        self.remarks_df,
        }.get(tab)
        if df is not None:
            self._populate_table(df, self.search_var.get())

    # ── FORMULA BAR ───────────────────────────────────────────────────────────
    def _on_tree_click(self, event):
        """Detect which column was clicked and update the formula bar."""
        region = self.tree.identify_region(event.x, event.y)
        if region != "cell":
            return
        col_id = self.tree.identify_column(event.x)
        item   = self.tree.identify_row(event.y)
        if not item or not col_id:
            return
        cols = self.tree["columns"]
        try:
            col_idx = int(col_id.replace("#", "")) - 1
            col_name = cols[col_idx]
        except (ValueError, IndexError):
            return
        values = self.tree.item(item, "values")
        val = values[col_idx] if col_idx < len(values) else ""
        self._update_fbar(col_name, val)

    def _on_cell_click(self, event=None):
        """Selection change — update bar with first column of selected row."""
        sel = self.tree.selection()
        if not sel:
            return
        cols   = self.tree["columns"]
        values = self.tree.item(sel[0], "values")
        if not cols or not values:
            return
        # Only refresh bar if no specific column was detected by _on_tree_click
        # (this fires for keyboard nav / programmatic selection)
        if not hasattr(self, "_fbar_last_item") or self._fbar_last_item != sel[0]:
            self._fbar_last_item = sel[0]
            self._update_fbar(cols[0], values[0])

    def _update_fbar(self, col_name, value):
        """Push column name + value into the Excel-style formula bar."""
        self._fbar_col_var.set(f"  {col_name}  ")
        self._fbar_val.config(state="normal")
        self._fbar_val.delete("1.0", "end")
        self._fbar_val.insert("1.0", str(value) if value != "" else "")
        self._fbar_val.config(state="disabled")

    def _fbar_copy(self):
        val = self._fbar_val.get("1.0", "end-1c")
        self.clipboard_clear()
        self.clipboard_append(val)
        show_toast(self.master, "Value copied!", color=SUCCESS, duration=1500)

    # ── ROW DETAIL POPUP (double-click) ───────────────────────────────────────
    def _on_double_click(self, event=None):
        sel = self.tree.selection()
        if not sel:
            return
        item    = sel[0]
        values  = self.tree.item(item, "values")
        columns = self.tree["columns"]
        if not columns or not values:
            return
        clicked_col = None
        if event and event.x:
            col_id = self.tree.identify_column(event.x)
            try:
                col_idx  = int(col_id.replace("#", "")) - 1
                clicked_col = columns[col_idx]
            except (ValueError, IndexError):
                pass
        self._show_row_detail(columns, values, clicked_col)

    def _show_row_detail(self, columns, values, clicked_col=None):
        popup = tk.Toplevel(self)
        popup.title("Row Details")
        popup.configure(bg=BG)
        popup.resizable(True, True)
        popup.grab_set()

        # ── Header bar ───────────────────────────────────────────────────────
        hdr = tk.Frame(popup, bg=HEADER_BG, height=48)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)
        tk.Label(hdr, text="🔍  Row Details",
                 font=("Segoe UI", 12, "bold"), bg=HEADER_BG, fg=TEXT
                 ).pack(side="left", padx=16, pady=10)
        tk.Label(hdr, text=f"{len(columns)} fields",
                 font=("Segoe UI", 9), bg=HEADER_BG, fg=SUBTEXT
                 ).pack(side="left")
        tk.Button(hdr, text="✕  Close",
                  font=("Segoe UI", 8, "bold"),
                  bg=CARD, fg=TEXT, relief="flat", padx=10, pady=4,
                  cursor="hand2", command=popup.destroy
                  ).pack(side="right", padx=12, pady=8)
        tk.Frame(popup, bg=ACCENT, height=2).pack(fill="x")

        # ── Scrollable body ───────────────────────────────────────────────────
        container = tk.Frame(popup, bg=BG)
        container.pack(fill="both", expand=True)

        canvas = tk.Canvas(container, bg=BG, highlightthickness=0, bd=0)
        vsb2   = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=vsb2.set)
        vsb2.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)

        inner  = tk.Frame(canvas, bg=BG)
        win_id = canvas.create_window((0, 0), window=inner, anchor="nw")

        inner.bind("<Configure>",  lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind("<Configure>", lambda e: canvas.itemconfig(win_id, width=e.width))

        def _mw(e): canvas.yview_scroll(int(-1 * (e.delta / 120)), "units")
        canvas.bind_all("<MouseWheel>", _mw)
        popup.bind("<Destroy>", lambda e: canvas.unbind_all("<MouseWheel>"))

        # ── Field rows ────────────────────────────────────────────────────────
        col_data = list(zip(columns, values))
        for i, (col, val) in enumerate(col_data):
            is_hl  = (col == clicked_col)
            row_bg = ACCENT2 if is_hl else (CARD if i % 2 == 0 else "#14172A")
            lbl_fg = "white"  if is_hl else SUBTEXT
            val_fg = "white"  if is_hl else TEXT

            row = tk.Frame(inner, bg=row_bg)
            row.pack(fill="x", padx=12, pady=1)

            tk.Label(row, text=f"{i+1:>3}", font=("Consolas", 8),
                     bg=row_bg, fg=SUBTEXT, width=3, anchor="e"
                     ).pack(side="left", padx=(8, 4), pady=6)

            tk.Label(row, text=col, font=("Segoe UI", 9, "bold"),
                     bg=row_bg, fg=lbl_fg, width=28, anchor="w"
                     ).pack(side="left", padx=(0, 10), pady=6)

            tk.Frame(row, bg=BORDER, width=1).pack(side="left", fill="y", pady=4)

            val_str  = str(val) if val != "" else "—"
            vt = tk.Text(row, font=("Segoe UI", 9),
                         bg=row_bg, fg=val_fg,
                         relief="flat", bd=0, height=1, wrap="none",
                         cursor="xterm",
                         selectbackground=ACCENT, selectforeground="white")
            vt.insert("1.0", val_str)
            if len(val_str) > 80:
                vt.config(height=min((len(val_str) // 80) + 1, 4), wrap="word")
            vt.config(state="disabled")
            vt.pack(side="left", fill="x", expand=True, padx=(10, 8), pady=4)

        # ── Footer ────────────────────────────────────────────────────────────
        footer = tk.Frame(popup, bg=DARK)
        footer.pack(fill="x", side="bottom")

        def _copy_all():
            popup.clipboard_clear()
            popup.clipboard_append("\n".join(f"{c}\t{v}" for c, v in col_data))
            show_toast(popup, "All fields copied!", color=SUCCESS)

        tk.Button(footer, text="📋  Copy All Fields",
                  font=("Segoe UI", 9, "bold"),
                  bg=ACCENT, fg="white", relief="flat",
                  padx=16, pady=7, cursor="hand2",
                  command=_copy_all).pack(side="left", padx=12, pady=8)
        tk.Label(footer, text="Double-click any row to view full details.",
                 font=("Segoe UI", 8), bg=DARK, fg=SUBTEXT
                 ).pack(side="right", padx=16)

        popup.update_idletasks()
        popup.geometry("700x560")
        px = self.winfo_rootx() + (self.winfo_width()  - 700) // 2
        py = self.winfo_rooty() + (self.winfo_height() - 560) // 2
        popup.geometry(f"700x560+{max(px,0)}+{max(py,0)}")

    def _download(self):
        save_path = filedialog.asksaveasfilename(
            title="Save output as",
            defaultextension=".xlsx",
            initialfile="cleaned_output.xlsx",
            filetypes=[("Excel files", "*.xlsx")]
        )
        if not save_path:
            return
        try:
            with open(save_path, "wb") as f:
                f.write(self.output_bytes)
            self.status_var.set(f"Saved → {save_path}")
            show_toast(self.master, f"Saved: {os.path.basename(save_path)}", color=SUCCESS)
        except Exception as e:
            show_toast(self.master, f"Save failed: {e}", color=DANGER, duration=5000)


# ══════════════════════════════════════════════════════════════════════════════
# BANK OF MAKATI PLACEHOLDER
# ══════════════════════════════════════════════════════════════════════════════
class BankOfMakatiScreen(tk.Frame):
    def __init__(self, master, username, on_back):
        super().__init__(master, bg=BG)
        self.username = username
        self.on_back  = on_back
        self._build()

    def _build(self):
        header = tk.Frame(self, bg=HEADER_BG, height=60)
        header.pack(fill="x")
        header.pack_propagate(False)

        left_hdr = tk.Frame(header, bg=HEADER_BG)
        left_hdr.pack(side="left", padx=12, pady=8)
        tk.Button(left_hdr, text="←  Back",
                  font=("Segoe UI", 9, "bold"),
                  bg=CARD, fg=TEXT, relief="flat",
                  padx=12, pady=6, cursor="hand2",
                  command=self.on_back).pack(side="left")
        tk.Label(left_hdr, text="  🏦  Bank of Makati",
                 font=("Segoe UI", 13, "bold"), bg=HEADER_BG, fg=TEXT).pack(side="left", padx=8)

        tk.Frame(self, bg=ACCENT2, height=3).pack(fill="x")

        outer = tk.Frame(self, bg=BG)
        outer.pack(expand=True)

        card = tk.Frame(outer, bg=CARD, padx=60, pady=60,
                        highlightthickness=1, highlightbackground=BORDER)
        card.pack(padx=20, pady=20)

        tk.Label(card, text="🏦", font=("Segoe UI", 48), bg=CARD, fg=ACCENT2).pack()
        tk.Label(card, text="Bank of Makati", font=("Segoe UI", 20, "bold"),
                 bg=CARD, fg=TEXT).pack(pady=(12, 4))
        tk.Label(card, text="Tools for this account are coming soon.",
                 font=("Segoe UI", 10), bg=CARD, fg=SUBTEXT).pack()

        footer = tk.Frame(self, bg=DARK)
        footer.pack(fill="x", side="bottom")
        tk.Label(footer, text="Created by  Vincent Corocoto  ·  09567796275",
                 font=("Segoe UI", 8, "bold"), bg=DARK, fg=ACCENT).pack(side="left", padx=16, pady=8)
        tk.Label(footer, text='"Kapag ang palay naging bigas, May bumayo."',
                 font=("Segoe UI", 8, "italic"), bg=DARK, fg="#4A5568").pack(side="right", padx=16, pady=8)


# ══════════════════════════════════════════════════════════════════════════════
# ACCOUNT MANAGEMENT SCREEN  (admin only)
# ══════════════════════════════════════════════════════════════════════════════
class AccountManagementScreen(tk.Frame):
    def __init__(self, master, username, on_back):
        super().__init__(master, bg=BG)
        self.username = username
        self.on_back  = on_back
        self._build()

    # ── layout ────────────────────────────────────────────────────────────────
    def _build(self):
        # Header
        header = tk.Frame(self, bg=HEADER_BG, height=60)
        header.pack(fill="x")
        header.pack_propagate(False)

        left_hdr = tk.Frame(header, bg=HEADER_BG)
        left_hdr.pack(side="left", padx=12, pady=8)
        tk.Button(left_hdr, text="←  Back",
                  font=("Segoe UI", 9, "bold"),
                  bg=CARD, fg=TEXT, relief="flat",
                  padx=12, pady=6, cursor="hand2",
                  command=self.on_back).pack(side="left")
        tk.Label(left_hdr, text="  ⚙  Account Management",
                 font=("Segoe UI", 13, "bold"), bg=HEADER_BG, fg=TEXT).pack(side="left", padx=8)

        tk.Frame(self, bg=GOLD, height=3).pack(fill="x")

        # Two-column layout
        body = tk.Frame(self, bg=BG)
        body.pack(fill="both", expand=True, padx=24, pady=20)

        # LEFT — user table
        left = tk.Frame(body, bg=CARD, highlightthickness=1, highlightbackground=BORDER)
        left.pack(side="left", fill="both", expand=True, padx=(0, 12))

        tk.Label(left, text="👥  Registered Users",
                 font=("Segoe UI", 11, "bold"), bg=CARD, fg=TEXT
                 ).pack(anchor="w", padx=16, pady=(14, 8))

        # Treeview
        tree_frame = tk.Frame(left, bg=CARD)
        tree_frame.pack(fill="both", expand=True, padx=12, pady=(0, 12))

        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Mgmt.Treeview",
                        background=DARK, fieldbackground=DARK,
                        foreground=TEXT, rowheight=30,
                        bordercolor=BORDER, relief="flat")
        style.configure("Mgmt.Treeview.Heading",
                        background=HEADER_BG, foreground=ACCENT,
                        font=("Segoe UI", 9, "bold"), relief="flat")
        style.map("Mgmt.Treeview",
                  background=[("selected", ACCENT)],
                  foreground=[("selected", "white")])

        cols = ("username", "role", "created_at")
        self.tree = ttk.Treeview(tree_frame, columns=cols, show="headings",
                                 style="Mgmt.Treeview", selectmode="browse")
        for col, w, label in [("username", 140, "Username"),
                               ("role",     80,  "Role"),
                               ("created_at", 160, "Created At")]:
            self.tree.heading(col, text=label)
            self.tree.column(col, width=w, anchor="w")
        self.tree.pack(fill="both", expand=True)
        self.tree.bind("<<TreeviewSelect>>", self._on_select)

        # Action buttons below table
        btns = tk.Frame(left, bg=CARD)
        btns.pack(fill="x", padx=12, pady=(0, 14))
        tk.Button(btns, text="🗑  Delete Selected",
                  font=("Segoe UI", 9, "bold"),
                  bg=DANGER, fg="white", relief="flat",
                  padx=10, pady=6, cursor="hand2",
                  command=self._delete_selected).pack(side="left", padx=(0, 8))
        tk.Button(btns, text="↺  Refresh",
                  font=("Segoe UI", 9, "bold"),
                  bg=CARD, fg=SUBTEXT, relief="flat",
                  padx=10, pady=6, cursor="hand2",
                  command=self._refresh).pack(side="left")

        # RIGHT — add / edit panel
        right = tk.Frame(body, bg=CARD, highlightthickness=1,
                         highlightbackground=BORDER, width=300)
        right.pack(side="right", fill="y")
        right.pack_propagate(False)

        tk.Label(right, text="➕  Add / Edit Account",
                 font=("Segoe UI", 11, "bold"), bg=CARD, fg=TEXT
                 ).pack(anchor="w", padx=16, pady=(14, 4))
        tk.Frame(right, bg=BORDER, height=1).pack(fill="x", padx=12, pady=(0, 12))

        # Form fields
        self._form_username = self._rfield(right, "Username")
        self._form_password = self._rfield(right, "Password", show="●")
        self._form_confirm  = self._rfield(right, "Confirm Password", show="●")

        tk.Label(right, text="ROLE", font=("Segoe UI", 8, "bold"),
                 bg=CARD, fg=SUBTEXT).pack(anchor="w", padx=16, pady=(8, 2))
        self._role_var = tk.StringVar(value="user")
        role_frame = tk.Frame(right, bg=CARD)
        role_frame.pack(anchor="w", padx=16)
        for val, label in [("user", "User"), ("supervisor", "Supervisor"), ("admin", "Admin")]:
            tk.Radiobutton(role_frame, text=label, variable=self._role_var, value=val,
                           font=("Segoe UI", 9), bg=CARD, fg=TEXT,
                           selectcolor=CARD, activebackground=CARD,
                           activeforeground=ACCENT).pack(side="left", padx=(0, 12))

        # Status / error
        self._form_err = tk.StringVar()
        tk.Label(right, textvariable=self._form_err,
                 font=("Segoe UI", 8), bg=CARD, fg=DANGER,
                 wraplength=260, justify="left").pack(anchor="w", padx=16, pady=(8, 0))

        # Buttons
        tk.Button(right, text="✚  Add New User",
                  font=("Segoe UI", 10, "bold"),
                  bg=SUCCESS, fg="white", relief="flat",
                  pady=9, cursor="hand2",
                  command=self._add_user).pack(fill="x", padx=16, pady=(12, 6))

        tk.Button(right, text="✎  Update Selected",
                  font=("Segoe UI", 10, "bold"),
                  bg=ACCENT, fg="white", relief="flat",
                  pady=9, cursor="hand2",
                  command=self._update_user).pack(fill="x", padx=16, pady=(0, 6))

        tk.Button(right, text="✕  Clear Form",
                  font=("Segoe UI", 9),
                  bg=CARD, fg=SUBTEXT, relief="flat",
                  pady=7, cursor="hand2",
                  command=self._clear_form).pack(fill="x", padx=16)

        # Footer
        footer = tk.Frame(self, bg=DARK)
        footer.pack(fill="x", side="bottom")
        tk.Label(footer, text="Created by  Vincent Corocoto  ·  09567796275",
                 font=("Segoe UI", 8, "bold"), bg=DARK, fg=ACCENT).pack(side="left", padx=16, pady=8)
        tk.Label(footer, text='"Kapag ang palay naging bigas, May bumayo."',
                 font=("Segoe UI", 8, "italic"), bg=DARK, fg="#4A5568").pack(side="right", padx=16, pady=8)

        self._refresh()

    # ── helpers ───────────────────────────────────────────────────────────────
    def _rfield(self, parent, label_text, show=None):
        tk.Label(parent, text=label_text.upper(),
                 font=("Segoe UI", 8, "bold"), bg=CARD, fg=SUBTEXT
                 ).pack(anchor="w", padx=16, pady=(8, 2))
        var = tk.StringVar()
        wrapper = tk.Frame(parent, bg=BORDER, padx=1, pady=1)
        wrapper.pack(fill="x", padx=12)
        inner = tk.Frame(wrapper, bg=CARD)
        inner.pack(fill="x")
        entry = tk.Entry(inner, textvariable=var,
                         font=("Segoe UI", 10), bg=CARD, fg=TEXT,
                         insertbackground=TEXT, relief="flat", bd=6, show=show or "")
        entry.pack(fill="x")
        def fi(e): wrapper.config(bg=ACCENT)
        def fo(e): wrapper.config(bg=BORDER)
        entry.bind("<FocusIn>", fi)
        entry.bind("<FocusOut>", fo)
        return var

    def _refresh(self):
        self.tree.delete(*self.tree.get_children())
        for row in db_list_users():
            tag = "admin" if row[1] == "admin" else ""
            self.tree.insert("", "end", values=row, tags=(tag,))
        self.tree.tag_configure("admin", foreground=GOLD)

    def _on_select(self, _=None):
        sel = self.tree.selection()
        if not sel:
            return
        vals = self.tree.item(sel[0], "values")
        self._form_username.set(vals[0])
        self._form_password.set("")
        self._form_confirm.set("")
        self._role_var.set(vals[1])
        self._form_err.set("")

    def _clear_form(self):
        self._form_username.set("")
        self._form_password.set("")
        self._form_confirm.set("")
        self._role_var.set("user")
        self._form_err.set("")

    def _add_user(self):
        u  = self._form_username.get().strip()
        p  = self._form_password.get().strip()
        p2 = self._form_confirm.get().strip()
        r  = self._role_var.get()
        if not u or not p:
            self._form_err.set("⚠ Username and password are required.")
            return
        if p != p2:
            self._form_err.set("⚠ Passwords do not match.")
            return
        err = db_add_user(u, p, r)
        if err:
            self._form_err.set(f"✗ {err}")
        else:
            self._form_err.set("")
            self._clear_form()
            self._refresh()
            show_toast(self.master, f"User '{u}' added.", color=SUCCESS)

    def _update_user(self):
        sel = self.tree.selection()
        if not sel:
            self._form_err.set("⚠ Select a user from the table first.")
            return
        original_username = self.tree.item(sel[0], "values")[0]
        p  = self._form_password.get().strip()
        p2 = self._form_confirm.get().strip()
        r  = self._role_var.get()
        if p and p != p2:
            self._form_err.set("⚠ Passwords do not match.")
            return
        # Protect last admin
        if r != "admin":
            rows = db_list_users()
            admins = [row for row in rows if row[1] == "admin"]
            if len(admins) == 1 and admins[0][0] == original_username:
                self._form_err.set("✗ Cannot remove admin role from the only admin.")
                return
        db_update_user(original_username, new_password=p, new_role=r)
        self._form_err.set("")
        self._clear_form()
        self._refresh()
        show_toast(self.master, f"User '{original_username}' updated.", color=SUCCESS)

    def _delete_selected(self):
        sel = self.tree.selection()
        if not sel:
            return
        username = self.tree.item(sel[0], "values")[0]
        if username == self.username:
            messagebox.showwarning("Cannot Delete", "You cannot delete your own account.", parent=self)
            return
        if not messagebox.askyesno("Confirm Delete",
                                   f"Delete user '{username}'? This cannot be undone.",
                                   parent=self):
            return
        err = db_delete_user(username)
        if err:
            messagebox.showerror("Error", err, parent=self)
        else:
            self._refresh()
            show_toast(self.master, f"User '{username}' deleted.", color=WARNING)


# ══════════════════════════════════════════════════════════════════════════════
# ROOT  — navigation controller
# ══════════════════════════════════════════════════════════════════════════════
_AppBase = TkinterDnD.Tk if _DND_AVAILABLE else tk.Tk

class App(_AppBase):
    def __init__(self):
        super().__init__()
        self.title("Operations Portal")
        self.geometry("1180x760")
        self.minsize(900, 620)
        self.configure(bg=BG)
        self.resizable(True, True)
        self._username = None
        self._role     = None
        self._current  = None
        self._show_login()

    # ── screen switcher ───────────────────────────────────────────────────────
    def _show(self, screen):
        if self._current:
            self._current.destroy()
        self._current = screen
        self._current.pack(fill="both", expand=True)

    def _show_login(self):
        self._show(LoginScreen(self, on_login=self._on_login))

    def _on_login(self, username, role):
        self._username = username
        self._role     = role
        self._show_dashboard()

    def _show_dashboard(self):
        self._show(DashboardScreen(
            self,
            username=self._username,
            role=self._role,
            on_select=self._on_account_select,
            on_logout=self._show_login,
            on_manage_accounts=self._show_account_management,
        ))

    def _show_account_management(self):
        self._show(AccountManagementScreen(
            self,
            username=self._username,
            on_back=self._show_dashboard,
        ))

    def _on_account_select(self, account):
        if account == "Orico Auto Loan":
            self._show(OricoToolsScreen(
                self,
                username=self._username,
                role=self._role,
                on_tool=self._on_tool_select,
                on_back=self._show_dashboard,
            ))
        elif account == "Bank of Makati":
            self._show(BankOfMakatiScreen(
                self,
                username=self._username,
                on_back=self._show_dashboard,
            ))

    def _on_tool_select(self, tool_tag):
        if tool_tag == "drr_cleaner":
            self._show(DRRCleanerScreen(
                self,
                username=self._username,
                role=self._role,
                on_back=lambda: self._on_account_select("Orico Auto Loan"),
            ))


# ══════════════════════════════════════════════════════════════════════════════
if __name__ == "__main__":
    app = App()
    app.mainloop()