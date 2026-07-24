#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
EESISP Remittance Calculator  --  v1.0
American Power Electrical Supply Co.

Desktop version of the union remittance calculator.

  * Employees are added, renamed, removed or moved between classifications
    without touching any formula.
  * Premium owed per classification = ROUND((gross wages - deductions) * rate, 2)
  * JIB = (JIB rate) % of total gross wages, all classifications, before deductions.
  * The remittance total = JIB + 401K C&L.
  * Every percentage asks for confirmation before it changes.
  * Wages and deductions can be pasted straight out of Excel.

Data is stored as JSON in the user's application-data folder, so it survives
rebuilds of the .exe.  Standard library only -- no third-party packages.

Python 3.8+
"""

import csv
import html
import json
import math
import os
import re
import tempfile
import time
import uuid
import webbrowser
from decimal import Decimal, ROUND_HALF_UP

import tkinter as tk
import tkinter.font as tkfont
from tkinter import ttk, filedialog, messagebox

APP_NAME = "EESISP Remittance Calculator"
VERSION = "1.0"
COMPANY = "American Power Electrical Supply"

# --------------------------------------------------------------------------
# theme
# --------------------------------------------------------------------------
BG = "#12161c"
PANEL = "#1a212b"
PANEL2 = "#212a36"
FIELD = "#0e1319"
LINE = "#2c3746"
INK = "#e6edf5"
MUTED = "#8b9bad"
ACCENT = "#3ea6ff"
GOOD = "#5bd17a"
WARN = "#f5a524"
DANGER = "#ff6b6b"

SANS_STACK = ["Segoe UI", "Helvetica Neue", "DejaVu Sans", "Arial"]
MONO_STACK = ["Consolas", "SF Mono", "DejaVu Sans Mono", "Courier New"]
COND_STACK = ["Arial Narrow", "Liberation Sans Narrow", "DejaVu Sans Condensed",
              "Segoe UI", "DejaVu Sans"]

# --------------------------------------------------------------------------
# defaults -- carried over from EESISP_CALC__REV_12_18_25.xlsx
# --------------------------------------------------------------------------
DEFAULT_ROLES = [
    {"id": "SALES", "name": "Sales", "code": "CC", "pct": 1.09},
    {"id": "DRIVERS", "name": "Drivers", "code": "C6", "pct": 13.80},
    {"id": "WAREHOUSE", "name": "Warehouse", "code": "C5", "pct": 5.07},
    {"id": "ADMIN", "name": "Admin", "code": "C4", "pct": 0.72},
]
DEFAULT_JIB_PCT = 29.0

SEED_EMPLOYEES = [
    ("Anichino, Jonathan P", "DRIVERS"),
    ("Dillon, Patrick", "WAREHOUSE"),
    ("Hughie, Derrick L", "DRIVERS"),
    ("Kelly, Thomas Patrick", "WAREHOUSE"),
    ("Kok, Iaroslav", "DRIVERS"),
    ("Paz, Elvis", "DRIVERS"),
    ("Perez, Gonzalo", "WAREHOUSE"),
    ("Rivera, Ricky", "WAREHOUSE"),
]


# --------------------------------------------------------------------------
# numbers
# --------------------------------------------------------------------------
def parse_num(value):
    """Read a number the way a person types it: $1,842.50  (500)  1842.5"""
    if isinstance(value, bool):
        return 0.0
    if isinstance(value, (int, float)):
        return float(value) if math.isfinite(float(value)) else 0.0
    raw = str(value if value is not None else "").strip()
    if not raw:
        return 0.0
    negative = raw.startswith("(") and raw.endswith(")")
    cleaned = re.sub(r"[^0-9.\-]", "", raw)
    parts = cleaned.split(".")
    if len(parts) > 2:
        cleaned = parts[0] + "." + "".join(parts[1:])
    try:
        number = float(cleaned)
    except ValueError:
        return 0.0
    if not math.isfinite(number):
        return 0.0
    return -abs(number) if negative else number


def round2(value):
    """Round half away from zero, the way Excel's ROUND() does."""
    try:
        return float(Decimal(str(float(value))).quantize(Decimal("0.01"),
                                                         rounding=ROUND_HALF_UP))
    except Exception:
        return 0.0


def money(value):
    value = float(value)
    return ("-$" if value < 0 else "$") + "{:,.2f}".format(abs(value))


def fix2(value):
    return "{:.2f}".format(parse_num(value))


def pct_text(value):
    text = "{:.4f}".format(round(parse_num(value), 4)).rstrip("0").rstrip(".")
    return text if text else "0"


# --------------------------------------------------------------------------
# state
# --------------------------------------------------------------------------
def data_dir():
    if os.name == "nt":
        base = os.environ.get("APPDATA") or os.path.expanduser("~")
    else:
        base = (os.environ.get("XDG_DATA_HOME")
                or os.path.join(os.path.expanduser("~"), ".local", "share"))
    folder = os.path.join(base, "EESISP Calculator")
    try:
        os.makedirs(folder, exist_ok=True)
    except OSError:
        folder = os.path.abspath(".")
    return folder


DATA_FILE = os.path.join(data_dir(), "eesisp_data.json")


def new_id():
    return "e" + uuid.uuid4().hex[:10]


def fresh_state():
    return {
        "period": "",
        "roles": [dict(r) for r in DEFAULT_ROLES],
        "employees": [
            {"id": new_id(), "name": name, "role": role,
             "wages": 0.0, "deds": 0.0, "active": True}
            for name, role in SEED_EMPLOYEES
        ],
        "jibPct": DEFAULT_JIB_PCT,
        "k401": 0.0,
        "history": [],
    }


def normalise_state(raw):
    """Merge a loaded file over the defaults and repair anything missing."""
    state = fresh_state()
    if not isinstance(raw, dict):
        return state
    state["period"] = str(raw.get("period", "") or "")
    state["k401"] = parse_num(raw.get("k401", 0))
    state["jibPct"] = parse_num(raw.get("jibPct", DEFAULT_JIB_PCT))

    roles = raw.get("roles")
    if isinstance(roles, list) and roles:
        merged = []
        for default in DEFAULT_ROLES:
            found = next((r for r in roles
                          if isinstance(r, dict) and r.get("id") == default["id"]), None)
            merged.append({
                "id": default["id"],
                "name": default["name"],
                "code": str((found or {}).get("code", default["code"]) or default["code"]),
                "pct": parse_num((found or {}).get("pct", default["pct"])),
            })
        state["roles"] = merged

    employees = raw.get("employees")
    if isinstance(employees, list):
        valid_roles = {r["id"] for r in state["roles"]}
        cleaned = []
        for emp in employees:
            if not isinstance(emp, dict):
                continue
            role = emp.get("role")
            cleaned.append({
                "id": str(emp.get("id") or new_id()),
                "name": str(emp.get("name", "") or ""),
                "role": role if role in valid_roles else state["roles"][1]["id"],
                "wages": parse_num(emp.get("wages", 0)),
                "deds": parse_num(emp.get("deds", 0)),
                "active": bool(emp.get("active", True)),
            })
        state["employees"] = cleaned

    history = raw.get("history")
    if isinstance(history, list):
        state["history"] = [h for h in history if isinstance(h, dict)]
    return state


def load_state():
    try:
        with open(DATA_FILE, "r", encoding="utf-8") as handle:
            return normalise_state(json.load(handle)), None
    except FileNotFoundError:
        return fresh_state(), None
    except Exception as exc:
        return fresh_state(), str(exc)


def save_state(state):
    tmp = DATA_FILE + ".tmp"
    with open(tmp, "w", encoding="utf-8") as handle:
        json.dump(state, handle, indent=2)
    os.replace(tmp, DATA_FILE)


def role_by_id(state, role_id):
    for role in state["roles"]:
        if role["id"] == role_id:
            return role
    return state["roles"][0]


# --------------------------------------------------------------------------
# the calculation
# --------------------------------------------------------------------------
def compute(state):
    by_role = {}
    for role in state["roles"]:
        by_role[role["id"]] = {"role": role, "count": 0, "wages": 0.0,
                               "deds": 0.0, "tgw": 0.0, "owed": 0.0}
    gross_all = 0.0
    for emp in state["employees"]:
        if not emp.get("active", True):
            continue
        bucket = by_role.get(emp.get("role")) or by_role[state["roles"][0]["id"]]
        bucket["count"] += 1
        bucket["wages"] += parse_num(emp.get("wages"))
        bucket["deds"] += parse_num(emp.get("deds"))
        gross_all += parse_num(emp.get("wages"))

    premiums = 0.0
    for bucket in by_role.values():
        bucket["tgw"] = round2(bucket["wages"] - bucket["deds"])
        bucket["owed"] = round2(bucket["tgw"] * (parse_num(bucket["role"]["pct"]) / 100.0))
        premiums += bucket["owed"]

    gross_all = round2(gross_all)
    jib_pct = parse_num(state.get("jibPct", DEFAULT_JIB_PCT))
    jib = round2(gross_all * (jib_pct / 100.0))
    k401 = round2(parse_num(state.get("k401")))
    return {
        "byRole": by_role,
        "grossAll": gross_all,
        "premiums": round2(premiums),
        "jibPct": jib_pct,
        "jib": jib,
        "k401": k401,
        "jibTotal": round2(jib + k401),
    }


# --------------------------------------------------------------------------
# excel paste helpers
# --------------------------------------------------------------------------
def parse_tsv(text):
    """Split clipboard text into a grid, honouring Excel's quoted cells."""
    text = str(text or "").replace("\r\n", "\n").replace("\r", "\n")
    rows, row, cell, quoted = [], [], "", False
    index = 0
    while index < len(text):
        char = text[index]
        if quoted:
            if char == '"':
                if index + 1 < len(text) and text[index + 1] == '"':
                    cell += '"'
                    index += 1
                else:
                    quoted = False
            else:
                cell += char
        elif char == '"' and cell == "":
            quoted = True
        elif char == "\t":
            row.append(cell)
            cell = ""
        elif char == "\n":
            row.append(cell)
            rows.append(row)
            row, cell = [], ""
        else:
            cell += char
        index += 1
    row.append(cell)
    rows.append(row)
    return [r for r in rows if any(str(c).strip() for c in r)]


def parse_role(state, text):
    value = str(text or "").strip().lower()
    if not value:
        return None
    for role in state["roles"]:
        if value in (role["id"].lower(), role["name"].lower(), role["code"].lower()):
            return role["id"]
    if re.match(r"^(wh|ware)", value):
        return "WAREHOUSE"
    if re.match(r"^(dr|driv|cdl)", value):
        return "DRIVERS"
    if re.match(r"^(sal|sls|inside|outside)", value):
        return "SALES"
    if re.match(r"^(adm|off|acct|clerk)", value):
        return "ADMIN"
    return None


def norm_name(value):
    return re.sub(r"\s+", " ", re.sub(r"[.,]", " ", str(value or "").lower())).strip()


def token_key(value):
    return " ".join(sorted(t for t in norm_name(value).split(" ") if len(t) > 1))


def surname(value):
    raw = str(value or "").strip()
    if "," in raw:
        return norm_name(raw.split(",")[0])
    parts = [p for p in norm_name(raw).split(" ") if p]
    return parts[-1] if parts else ""


def build_matcher(state):
    exact, token, sur = {}, {}, {}
    for emp in state["employees"]:
        for table, key in ((exact, norm_name(emp["name"])),
                           (token, token_key(emp["name"])),
                           (sur, surname(emp["name"]))):
            if key:
                table.setdefault(key, []).append(emp["id"])

    def match(raw):
        if not norm_name(raw):
            return None
        for table, key in ((exact, norm_name(raw)),
                           (token, token_key(raw)),
                           (sur, surname(raw))):
            hits = table.get(key)
            if hits and len(hits) == 1:
                return hits[0]
        return None

    return match


def has_digit(value):
    return any(ch.isdigit() for ch in str(value))


def guess_columns(grid, header):
    """Work out which pasted column holds what."""
    data = grid[1:] if header else grid
    col_count = max(len(r) for r in grid)
    idx = {"name": -1, "wages": -1, "deds": -1, "role": -1}

    if header:
        for i, cell in enumerate(grid[0]):
            text = str(cell).lower()
            if idx["name"] < 0 and re.search(r"name|employee|worker", text):
                idx["name"] = i
            if idx["deds"] < 0 and re.search(r"ded|deduct|pre.?tax", text):
                idx["deds"] = i
            if idx["wages"] < 0 and re.search(r"wage|gross|earn|salary|amount", text):
                idx["wages"] = i
            if idx["role"] < 0 and re.search(r"role|class|dept|depart|division|job|type", text):
                idx["role"] = i

    numeric = []
    for col in range(col_count):
        values = [str(r[col]).strip() for r in data
                  if col < len(r) and str(r[col]).strip()]
        if not values:
            continue
        count = sum(1 for v in values
                    if has_digit(v) and not re.search(r"[a-z]{2}", v, re.I))
        if count >= len(values) * 0.7:
            numeric.append(col)
        elif idx["name"] < 0:
            idx["name"] = col

    free = [c for c in numeric if c not in (idx["wages"], idx["deds"])]
    if idx["wages"] < 0 and free:
        idx["wages"] = free.pop(0)
    if idx["deds"] < 0 and free:
        idx["deds"] = free.pop(0)
    if idx["deds"] == idx["wages"]:
        idx["deds"] = -1
    return idx


def looks_like_header(grid):
    if len(grid) < 2:
        return False
    first = [str(c).strip() for c in grid[0]]
    if any(has_digit(c) for c in first):
        return False
    return any(re.search(r"name|employee|wage|gross|deduct|role|class|earn|pay", c, re.I)
               for c in first)


# --------------------------------------------------------------------------
# small widget helpers
# --------------------------------------------------------------------------
def pick_font(family_stack, size, weight="normal", slant="roman"):
    available = set(tkfont.families())
    for family in family_stack:
        if family in available:
            return tkfont.Font(family=family, size=size, weight=weight, slant=slant)
    return tkfont.Font(size=size, weight=weight, slant=slant)


class Fonts(object):
    def __init__(self):
        self.sans = pick_font(SANS_STACK, 10)
        self.sans_small = pick_font(SANS_STACK, 9)
        self.sans_bold = pick_font(SANS_STACK, 10, "bold")
        self.mono = pick_font(MONO_STACK, 10)
        self.mono_small = pick_font(MONO_STACK, 8)
        self.mono_big = pick_font(MONO_STACK, 22, "bold")
        self.cond = pick_font(COND_STACK, 10, "bold")
        self.cond_small = pick_font(COND_STACK, 8, "bold")
        self.title = pick_font(COND_STACK, 20, "bold")


def flat_button(parent, text, command, kind="normal", width=None, fonts=None):
    colours = {
        "normal": (PANEL2, INK, LINE),
        "accent": (ACCENT, "#08121c", ACCENT),
        "danger": (DANGER, "#1a0606", DANGER),
        "ghost": (PANEL, MUTED, LINE),
    }[kind]
    button = tk.Button(
        parent, text=text, command=command,
        bg=colours[0], fg=colours[1],
        activebackground=colours[0], activeforeground=colours[1],
        highlightbackground=colours[2], highlightcolor=ACCENT, highlightthickness=1,
        relief="flat", bd=0, padx=10, pady=4, cursor="hand2",
        font=(fonts.sans if fonts else None),
    )
    if width:
        button.configure(width=width)
    return button


def card(parent, title, fonts, subtitle=None):
    """A titled panel. Returns (outer frame, body frame, header frame)."""
    outer = tk.Frame(parent, bg=PANEL, highlightbackground=LINE,
                     highlightthickness=1, bd=0)
    header = tk.Frame(outer, bg=PANEL, height=1)
    header.pack(fill="x", padx=13, pady=(11, 9))
    label = tk.Label(header, text=title.upper(), bg=PANEL, fg=INK,
                     font=fonts.cond, anchor="w")
    label.pack(side="left")
    if subtitle:
        tk.Label(header, text=subtitle, bg=PANEL, fg=MUTED,
                 font=fonts.sans_small).pack(side="left", padx=(10, 0))
    tk.Frame(outer, bg=LINE, height=1).pack(fill="x")
    body = tk.Frame(outer, bg=PANEL)
    body.pack(fill="both", expand=True)
    return outer, body, header


# --------------------------------------------------------------------------
# confirmation dialog
# --------------------------------------------------------------------------
class ConfirmDialog(tk.Toplevel):
    def __init__(self, parent, fonts, title, body, yes="Confirm", danger=False):
        tk.Toplevel.__init__(self, parent)
        self.result = False
        self.configure(bg=PANEL)
        self.title(title)
        self.resizable(False, False)
        self.transient(parent)

        wrap = tk.Frame(self, bg=PANEL, highlightbackground=LINE, highlightthickness=1)
        wrap.pack(fill="both", expand=True)

        tk.Label(wrap, text=title.upper(), bg=PANEL, fg=INK, font=fonts.cond,
                 anchor="w", justify="left").pack(fill="x", padx=18, pady=(16, 6))
        tk.Label(wrap, text=body, bg=PANEL, fg=MUTED, font=fonts.sans,
                 wraplength=380, justify="left", anchor="w").pack(fill="x", padx=18)

        tk.Frame(wrap, bg=LINE, height=1).pack(fill="x", pady=(16, 0))
        row = tk.Frame(wrap, bg=PANEL)
        row.pack(fill="x", padx=14, pady=11)
        ok_btn = flat_button(row, yes, self._yes,
                             "danger" if danger else "accent", fonts=fonts)
        ok_btn.pack(side="right")
        flat_button(row, "Cancel", self._no, "normal",
                    fonts=fonts).pack(side="right", padx=(0, 8))

        self.bind("<Escape>", lambda e: self._no())
        self.bind("<Return>", lambda e: self._yes())
        self.protocol("WM_DELETE_WINDOW", self._no)

        self.update_idletasks()
        self._centre(parent)
        self.grab_set()
        ok_btn.focus_set()

    def _centre(self, parent):
        width, height = self.winfo_width(), self.winfo_height()
        x = parent.winfo_rootx() + (parent.winfo_width() - width) // 2
        y = parent.winfo_rooty() + (parent.winfo_height() - height) // 3
        self.geometry("+{}+{}".format(max(x, 0), max(y, 0)))

    def _yes(self):
        self.result = True
        self.destroy()

    def _no(self):
        self.result = False
        self.destroy()


def ask_confirm(parent, fonts, title, body, yes="Confirm", danger=False):
    dialog = ConfirmDialog(parent, fonts, title, body, yes, danger)
    parent.wait_window(dialog)
    return dialog.result


# --------------------------------------------------------------------------
# import-from-Excel dialog
# --------------------------------------------------------------------------
class ImportDialog(tk.Toplevel):
    def __init__(self, parent, fonts, state, prefill=""):
        tk.Toplevel.__init__(self, parent)
        self.fonts = fonts
        self.state_ref = state
        self.result = None
        self.rows = []
        self.grid_data = []
        self.target_vars = []

        self.configure(bg=PANEL)
        self.title("Paste from Excel")
        self.geometry("780x680")
        self.minsize(660, 560)
        self.transient(parent)

        head = tk.Frame(self, bg=PANEL)
        head.pack(fill="x", padx=18, pady=(15, 0))
        tk.Label(head, text="PASTE FROM EXCEL", bg=PANEL, fg=INK,
                 font=fonts.cond, anchor="w").pack(fill="x")
        tk.Label(head, bg=PANEL, fg=MUTED, font=fonts.sans_small,
                 justify="left", anchor="w", wraplength=700,
                 text=("Select the rows in your spreadsheet and copy them, then use "
                       "Paste from clipboard (or press Ctrl+V in the box). Headers are "
                       "fine. Names get matched to the people already in the list.")
                 ).pack(fill="x", pady=(3, 0))

        box = tk.Frame(self, bg=PANEL)
        box.pack(fill="x", padx=18, pady=(12, 0))
        self.text = tk.Text(box, height=7, bg=FIELD, fg=INK, insertbackground=ACCENT,
                            font=fonts.mono_small, relief="flat",
                            highlightbackground=ACCENT, highlightthickness=1, wrap="none",
                            selectbackground=ACCENT, selectforeground="#08121c")
        self.text.pack(fill="x")
        self.text.bind("<<Paste>>", lambda e: self.after(30, self.read_paste))
        self.text.bind("<KeyRelease>", lambda e: self.after(120, self.read_paste))

        tools = tk.Frame(self, bg=PANEL)
        tools.pack(fill="x", padx=18, pady=(8, 0))
        flat_button(tools, "Paste from clipboard", self.paste_clipboard,
                    "accent", fonts=fonts).pack(side="left")
        flat_button(tools, "Clear", self.clear_text, "ghost",
                    fonts=fonts).pack(side="left", padx=(8, 0))

        # column mapping ---------------------------------------------------
        self.map_frame = tk.Frame(self, bg=PANEL)
        self.map_frame.pack(fill="x", padx=18, pady=(12, 0))
        self.col_vars = {}
        for label in ("Name", "Gross wages", "Deductions", "Role"):
            cell = tk.Frame(self.map_frame, bg=PANEL)
            cell.pack(side="left", padx=(0, 14))
            tk.Label(cell, text=label, bg=PANEL, fg=MUTED,
                     font=fonts.sans_small).pack(side="left", padx=(0, 5))
            var = tk.StringVar()
            combo = ttk.Combobox(cell, textvariable=var, state="readonly",
                                 width=17, style="Dark.TCombobox", font=fonts.sans_small)
            combo.pack(side="left")
            combo.bind("<<ComboboxSelected>>", lambda e: self.build_preview())
            self.col_vars[label] = (var, combo)

        self.header_var = tk.BooleanVar(value=False)
        tk.Checkbutton(self.map_frame, text="First row is a header",
                       variable=self.header_var, command=self.on_header_toggle,
                       bg=PANEL, fg=MUTED, selectcolor=FIELD, font=fonts.sans_small,
                       activebackground=PANEL, activeforeground=INK,
                       highlightthickness=0, bd=0).pack(side="left")

        # preview ----------------------------------------------------------
        prev_wrap = tk.Frame(self, bg=PANEL, highlightbackground=LINE,
                             highlightthickness=1)
        prev_wrap.pack(fill="both", expand=True, padx=18, pady=(12, 0))

        heads = tk.Frame(prev_wrap, bg=PANEL2)
        heads.pack(fill="x")
        for text, width, anchor in (("FROM YOUR SPREADSHEET", 26, "w"),
                                    ("GROSS WAGES", 13, "e"),
                                    ("DEDUCTIONS", 12, "e"),
                                    ("APPLIES TO", 24, "w")):
            tk.Label(heads, text=text, bg=PANEL2, fg=MUTED, font=fonts.cond_small,
                     width=width, anchor=anchor).pack(side="left", padx=6, pady=5)

        canvas_holder = tk.Frame(prev_wrap, bg=PANEL)
        canvas_holder.pack(fill="both", expand=True)
        self.canvas = tk.Canvas(canvas_holder, bg=PANEL, highlightthickness=0)
        bar = ttk.Scrollbar(canvas_holder, orient="vertical",
                            command=self.canvas.yview, style="Dark.Vertical.TScrollbar")
        self.canvas.configure(yscrollcommand=bar.set)
        bar.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)
        self.prev_inner = tk.Frame(self.canvas, bg=PANEL)
        self.prev_window = self.canvas.create_window((0, 0), window=self.prev_inner,
                                                     anchor="nw")
        self.prev_inner.bind("<Configure>", lambda e: self.canvas.configure(
            scrollregion=self.canvas.bbox("all")))
        self.canvas.bind("<Configure>", lambda e: self.canvas.itemconfigure(
            self.prev_window, width=e.width))

        self.zero_var = tk.BooleanVar(value=False)
        self.zero_check = tk.Checkbutton(
            self, variable=self.zero_var, bg=PANEL, fg=MUTED, selectcolor=FIELD,
            font=fonts.sans_small, activebackground=PANEL, activeforeground=INK,
            highlightthickness=0, bd=0, anchor="w", justify="left", wraplength=700,
            text=("Set wages and deductions to 0.00 for anyone not in this paste "
                  "- use it when the spreadsheet is the whole payroll for the period."))
        self.zero_check.pack(fill="x", padx=16, pady=(10, 0))

        tk.Frame(self, bg=LINE, height=1).pack(fill="x", pady=(12, 0))
        foot = tk.Frame(self, bg=PANEL)
        foot.pack(fill="x", padx=14, pady=11)
        self.summary = tk.Label(foot, text="", bg=PANEL, fg=MUTED, font=fonts.sans_small)
        self.summary.pack(side="left")
        self.apply_btn = flat_button(foot, "Apply", self.apply, "accent", fonts=fonts)
        self.apply_btn.pack(side="right")
        self.apply_btn.configure(state="disabled")
        flat_button(foot, "Cancel", self.cancel, "normal",
                    fonts=fonts).pack(side="right", padx=(0, 8))

        self.bind("<Escape>", lambda e: self.cancel())
        self.protocol("WM_DELETE_WINDOW", self.cancel)

        if prefill:
            self.text.insert("1.0", prefill)
            self.read_paste()
        self.update_idletasks()
        self.grab_set()
        self.text.focus_set()

    # ------------------------------------------------------------------
    def paste_clipboard(self):
        try:
            data = self.clipboard_get()
        except tk.TclError:
            messagebox.showinfo("Nothing to paste",
                                "The clipboard is empty, or it does not hold text.",
                                parent=self)
            return
        self.text.delete("1.0", "end")
        self.text.insert("1.0", data)
        self.read_paste()

    def clear_text(self):
        self.text.delete("1.0", "end")
        self.read_paste()

    def raw_text(self):
        return self.text.get("1.0", "end-1c")

    def on_header_toggle(self):
        self.fill_column_pickers()
        self.build_preview()

    def read_paste(self):
        grid = parse_tsv(self.raw_text())
        self.grid_data = grid
        if not grid:
            self.reset_preview()
            return
        self.header_var.set(looks_like_header(grid))
        self.fill_column_pickers()
        self.build_preview()

    def reset_preview(self):
        for child in self.prev_inner.winfo_children():
            child.destroy()
        self.rows, self.target_vars = [], []
        self.summary.configure(text="")
        self.apply_btn.configure(state="disabled")

    def column_labels(self):
        grid, header = self.grid_data, self.header_var.get()
        labels = ["- none -"]
        if not grid:
            return labels
        count = max(len(r) for r in grid)
        sample_row = grid[1] if (header and len(grid) > 1) else grid[0]
        for col in range(count):
            head = str(grid[0][col]).strip() if (header and col < len(grid[0])) else ""
            sample = str(sample_row[col]).strip() if col < len(sample_row) else ""
            tail = head or sample
            labels.append("Col {}{}".format(col + 1,
                                            " - " + tail[:16] if tail else ""))
        return labels

    def fill_column_pickers(self):
        if not self.grid_data:
            return
        labels = self.column_labels()
        guess = guess_columns(self.grid_data, self.header_var.get())
        keys = {"Name": "name", "Gross wages": "wages",
                "Deductions": "deds", "Role": "role"}
        for label, (var, combo) in self.col_vars.items():
            combo.configure(values=labels)
            index = guess[keys[label]]
            var.set(labels[index + 1] if 0 <= index < len(labels) - 1 else labels[0])

    def selected_col(self, label):
        var, _ = self.col_vars[label]
        try:
            return self.column_labels().index(var.get()) - 1
        except ValueError:
            return -1

    def build_preview(self):
        for child in self.prev_inner.winfo_children():
            child.destroy()
        self.rows, self.target_vars = [], []
        grid = self.grid_data
        if not grid:
            self.reset_preview()
            return

        data = grid[1:] if self.header_var.get() else grid
        c_name = self.selected_col("Name")
        c_wages = self.selected_col("Gross wages")
        c_deds = self.selected_col("Deductions")
        c_role = self.selected_col("Role")
        match = build_matcher(self.state_ref)

        def cell(row, index):
            return str(row[index]).strip() if 0 <= index < len(row) else ""

        for row in data:
            name = cell(row, c_name)
            wages = parse_num(cell(row, c_wages)) if c_wages >= 0 else None
            deds = parse_num(cell(row, c_deds)) if c_deds >= 0 else None
            role = parse_role(self.state_ref, cell(row, c_role)) if c_role >= 0 else None
            if not name and wages is None:
                continue
            target = (match(name) or "__new") if name else "__skip"
            self.rows.append({"name": name, "wages": wages, "deds": deds,
                              "role": role, "target": target})

        options = ["Add as new employee", "Skip this row"]
        id_for_option = {options[0]: "__new", options[1]: "__skip"}
        option_for_id = {"__new": options[0], "__skip": options[1]}
        for emp in self.state_ref["employees"]:
            label = emp["name"].strip() or "(unnamed)"
            while label in id_for_option:
                label += " "
            options.append(label)
            id_for_option[label] = emp["id"]
            option_for_id[emp["id"]] = label

        for index, entry in enumerate(self.rows):
            line = tk.Frame(self.prev_inner, bg=PANEL)
            line.pack(fill="x")
            tk.Label(line, text=entry["name"] or "-", bg=PANEL, fg=INK,
                     font=self.fonts.sans, width=26, anchor="w"
                     ).pack(side="left", padx=6, pady=3)
            tk.Label(line, text="-" if entry["wages"] is None else money(entry["wages"]),
                     bg=PANEL, fg=INK, font=self.fonts.mono, width=13, anchor="e"
                     ).pack(side="left", padx=6)
            tk.Label(line, text="-" if entry["deds"] is None else money(entry["deds"]),
                     bg=PANEL, fg=INK, font=self.fonts.mono, width=12, anchor="e"
                     ).pack(side="left", padx=6)
            var = tk.StringVar(value=option_for_id.get(entry["target"], options[0]))
            combo = ttk.Combobox(line, textvariable=var, values=options, state="readonly",
                                 width=24, style="Dark.TCombobox", font=self.fonts.sans_small)
            combo.pack(side="left", padx=6)

            def on_pick(_event, i=index, v=var):
                self.rows[i]["target"] = id_for_option.get(v.get(), "__skip")
                self.update_summary()

            combo.bind("<<ComboboxSelected>>", on_pick)
            self.target_vars.append(var)
            tk.Frame(self.prev_inner, bg=LINE, height=1).pack(fill="x")

        self.apply_btn.configure(state=("normal" if self.rows else "disabled"))
        self.update_summary()

    def update_summary(self):
        updated = sum(1 for r in self.rows if r["target"] not in ("__new", "__skip"))
        added = sum(1 for r in self.rows if r["target"] == "__new")
        skipped = sum(1 for r in self.rows if r["target"] == "__skip")
        bits = []
        if updated:
            bits.append("{} matched".format(updated))
        if added:
            bits.append("{} new".format(added))
        if skipped:
            bits.append("{} skipped".format(skipped))
        self.summary.configure(text="  ".join(bits))

    def apply(self):
        self.result = {"rows": list(self.rows), "zero_rest": bool(self.zero_var.get())}
        self.destroy()

    def cancel(self):
        self.result = None
        self.destroy()


# --------------------------------------------------------------------------
# main window
# --------------------------------------------------------------------------
class App(tk.Tk):
    def __init__(self):
        tk.Tk.__init__(self)
        self.title("{}  v{}".format(APP_NAME, VERSION))
        self.geometry("1200x820")
        self.minsize(980, 640)
        self.configure(bg=BG)

        self.fonts = Fonts()
        self._install_styles()

        self.state_data, load_error = load_state()
        self._building = False
        self._confirm_open = False
        self._pct_prev = {}
        self._save_job = None
        self._status_job = None
        self._emp_widgets = []

        self._build_header()
        self._build_body()

        self.sort_employees()
        self.render_employees()
        self.render_rates()
        self.render_history()
        self.render_slip()

        if load_error:
            self.set_status("Saved data could not be read - starting fresh", warn=True)
        else:
            self.set_status("Ready")

        self.protocol("WM_DELETE_WINDOW", self.on_close)

    # ------------------------------------------------------------------
    def _install_styles(self):
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure("Dark.TCombobox", fieldbackground=PANEL2, background=PANEL2,
                        foreground=INK, arrowcolor=MUTED, bordercolor=LINE,
                        lightcolor=PANEL2, darkcolor=PANEL2, relief="flat",
                        selectbackground=PANEL2, selectforeground=INK)
        style.map("Dark.TCombobox",
                  fieldbackground=[("readonly", PANEL2), ("disabled", PANEL)],
                  foreground=[("readonly", INK)],
                  selectbackground=[("readonly", PANEL2)],
                  selectforeground=[("readonly", INK)],
                  bordercolor=[("focus", ACCENT)])
        self.option_add("*TCombobox*Listbox.background", PANEL2)
        self.option_add("*TCombobox*Listbox.foreground", INK)
        self.option_add("*TCombobox*Listbox.selectBackground", ACCENT)
        self.option_add("*TCombobox*Listbox.selectForeground", "#08121c")
        style.configure("Dark.Vertical.TScrollbar", background=PANEL2, troughcolor=BG,
                        bordercolor=BG, arrowcolor=MUTED, relief="flat")
        style.map("Dark.Vertical.TScrollbar", background=[("active", LINE)])

    # ------------------------------------------------------------------
    def _build_header(self):
        head = tk.Frame(self, bg=BG)
        head.pack(fill="x", padx=18, pady=(14, 0))

        left = tk.Frame(head, bg=BG)
        left.pack(side="left", anchor="w")
        tk.Label(left, text=COMPANY.upper(), bg=BG, fg=ACCENT,
                 font=self.fonts.cond_small).pack(anchor="w")
        tk.Label(left, text="EESISP REMITTANCE", bg=BG, fg=INK,
                 font=self.fonts.title).pack(anchor="w")

        right = tk.Frame(head, bg=BG)
        right.pack(side="right", anchor="e")
        row = tk.Frame(right, bg=BG)
        row.pack(anchor="e")
        tk.Label(row, text="PAY PERIOD", bg=BG, fg=MUTED,
                 font=self.fonts.cond_small).pack(side="left", padx=(0, 8))
        self.period_var = tk.StringVar(value=self.state_data["period"])
        period = tk.Entry(row, textvariable=self.period_var, bg=PANEL, fg=INK,
                          insertbackground=ACCENT, relief="flat", width=26,
                          highlightbackground=LINE, highlightcolor=ACCENT,
                          highlightthickness=1, font=self.fonts.sans)
        period.pack(side="left", ipady=4)
        self.period_var.trace_add("write", lambda *a: self.on_period_change())
        self.status = tk.Label(right, text="", bg=BG, fg=GOOD,
                               font=self.fonts.cond_small, anchor="e")
        self.status.pack(anchor="e", pady=(5, 0))

        tk.Frame(self, bg=LINE, height=1).pack(fill="x", padx=18, pady=(12, 0))

    # ------------------------------------------------------------------
    def _build_body(self):
        body = tk.Frame(self, bg=BG)
        body.pack(fill="both", expand=True, padx=18, pady=16)

        # ---- scrollable left column ----
        left_holder = tk.Frame(body, bg=BG)
        left_holder.pack(side="left", fill="both", expand=True)
        self.left_canvas = tk.Canvas(left_holder, bg=BG, highlightthickness=0)
        left_bar = ttk.Scrollbar(left_holder, orient="vertical",
                                 command=self.left_canvas.yview,
                                 style="Dark.Vertical.TScrollbar")
        self.left_canvas.configure(yscrollcommand=left_bar.set)
        left_bar.pack(side="right", fill="y")
        self.left_canvas.pack(side="left", fill="both", expand=True)
        self.left_inner = tk.Frame(self.left_canvas, bg=BG)
        self.left_window = self.left_canvas.create_window((0, 0), window=self.left_inner,
                                                          anchor="nw")
        self.left_inner.bind("<Configure>", lambda e: self.left_canvas.configure(
            scrollregion=self.left_canvas.bbox("all")))
        self.left_canvas.bind("<Configure>", lambda e: self.left_canvas.itemconfigure(
            self.left_window, width=e.width))
        self.left_canvas.bind("<Enter>", lambda e: self._bind_wheel())
        self.left_canvas.bind("<Leave>", lambda e: self._unbind_wheel())

        self._build_employees(self.left_inner)
        self._build_rates(self.left_inner)
        self._build_history(self.left_inner)

        # ---- fixed slip on the right ----
        self.slip_holder = tk.Frame(body, bg=BG, width=350)
        self.slip_holder.pack(side="right", fill="y", padx=(18, 0))
        self.slip_holder.pack_propagate(False)
        self._build_slip(self.slip_holder)

    def _bind_wheel(self):
        self.bind_all("<MouseWheel>", self._on_wheel)
        self.bind_all("<Button-4>", self._on_wheel)
        self.bind_all("<Button-5>", self._on_wheel)

    def _unbind_wheel(self):
        self.unbind_all("<MouseWheel>")
        self.unbind_all("<Button-4>")
        self.unbind_all("<Button-5>")

    def _on_wheel(self, event):
        if getattr(event, "num", None) == 4:
            step = -1
        elif getattr(event, "num", None) == 5:
            step = 1
        else:
            step = -1 if event.delta > 0 else 1
        self.left_canvas.yview_scroll(step, "units")

    # ------------------------------------------------------------------
    def _build_employees(self, parent):
        outer, body, header = card(parent, "Employees", self.fonts)
        outer.pack(fill="x", pady=(0, 16))

        tools = tk.Frame(header, bg=PANEL)
        tools.pack(side="right")
        flat_button(tools, "+ Add employee", self.add_employee, "accent",
                    fonts=self.fonts).pack(side="right")
        flat_button(tools, "Paste from Excel", self.open_import, "normal",
                    fonts=self.fonts).pack(side="right", padx=(0, 8))

        note = tk.Frame(body, bg="#241f10", highlightbackground=WARN,
                        highlightthickness=1)
        note.pack(fill="x", padx=13, pady=(11, 4))
        tk.Label(note, bg="#241f10", fg="#f0dcb6", font=self.fonts.sans_small,
                 justify="left", anchor="w", wraplength=620,
                 text=("Check the roles below. In the spreadsheet the group formulas "
                       "pointed at fixed row numbers, so after the names were "
                       "alphabetised some people may have been counted in the wrong "
                       "classification. Set each person's role once here and it stays "
                       "correct.")).pack(side="left", padx=10, pady=7)
        flat_button(note, "Dismiss", lambda: note.destroy(), "ghost",
                    fonts=self.fonts).pack(side="right", padx=8)

        self.emp_grid = tk.Frame(body, bg=PANEL)
        self.emp_grid.pack(fill="x", padx=13, pady=(6, 12))
        for col, weight in ((0, 3), (1, 0), (2, 0), (3, 0), (4, 0), (5, 0), (6, 0)):
            self.emp_grid.columnconfigure(col, weight=weight)

    def _build_rates(self, parent):
        outer, body, _ = card(parent, "Classifications & premium rates", self.fonts,
                              "Changing a rate asks for confirmation first.")
        outer.pack(fill="x", pady=(0, 16))
        self.rate_grid = tk.Frame(body, bg=PANEL)
        self.rate_grid.pack(fill="x", padx=13, pady=10)
        self.rate_grid.columnconfigure(0, weight=1)

    def _build_history(self, parent):
        outer, body, header = card(parent, "Saved periods", self.fonts)
        outer.pack(fill="x", pady=(0, 4))
        self.hist_count = tk.Label(header, text="(0)", bg=PANEL, fg=MUTED,
                                   font=self.fonts.sans_small)
        self.hist_count.pack(side="left", padx=(6, 0))
        self.hist_grid = tk.Frame(body, bg=PANEL)
        self.hist_grid.pack(fill="x", padx=13, pady=10)
        self.hist_grid.columnconfigure(0, weight=1)

    # ------------------------------------------------------------------
    def _build_slip(self, parent):
        outer = tk.Frame(parent, bg=PANEL, highlightbackground=LINE,
                         highlightthickness=1)
        outer.pack(fill="both", expand=True)

        top = tk.Frame(outer, bg=PANEL)
        top.pack(fill="x", padx=15, pady=(14, 10))
        tk.Label(top, text="AMOUNT DUE", bg=PANEL, fg=ACCENT,
                 font=self.fonts.cond_small).pack(anchor="w")
        self.slip_period = tk.Label(top, text="No period entered", bg=PANEL, fg=MUTED,
                                    font=self.fonts.sans_small, anchor="w")
        self.slip_period.pack(anchor="w", pady=(2, 0))
        tk.Frame(outer, bg=LINE, height=1).pack(fill="x")

        self.class_frame = tk.Frame(outer, bg=PANEL)
        self.class_frame.pack(fill="x", padx=15, pady=(10, 4))

        tk.Frame(outer, bg=LINE, height=1).pack(fill="x", padx=15, pady=4)

        jib_row = tk.Frame(outer, bg=PANEL)
        jib_row.pack(fill="x", padx=15, pady=(4, 0))
        jib_left = tk.Frame(jib_row, bg=PANEL)
        jib_left.pack(side="left")
        tk.Label(jib_left, text="JIB", bg=PANEL, fg=INK,
                 font=self.fonts.cond).pack(side="left")
        self.jib_tag = tk.Label(jib_left, text="29%", bg=PANEL, fg=MUTED,
                                font=self.fonts.mono_small)
        self.jib_tag.pack(side="left", padx=(6, 0))
        jib_right = tk.Frame(jib_row, bg=PANEL)
        jib_right.pack(side="right")
        self.jib_amount = tk.Label(jib_right, text="$0.00", bg=PANEL, fg=INK,
                                   font=self.fonts.mono, anchor="e")
        self.jib_amount.pack(anchor="e")
        self.jib_basis = tk.Label(jib_right, text="", bg=PANEL, fg=MUTED,
                                  font=self.fonts.mono_small, anchor="e")
        self.jib_basis.pack(anchor="e")

        k_row = tk.Frame(outer, bg=PANEL)
        k_row.pack(fill="x", padx=15, pady=(8, 12))
        tk.Label(k_row, text="401K C&L", bg=PANEL, fg=INK,
                 font=self.fonts.cond).pack(side="left")
        self.k401_var = tk.StringVar(value=fix2(self.state_data["k401"]))
        k_entry = tk.Entry(k_row, textvariable=self.k401_var, bg=FIELD, fg=INK,
                           insertbackground=ACCENT, relief="flat", width=12,
                           justify="right", highlightbackground=WARN,
                           highlightcolor=ACCENT, highlightthickness=1,
                           font=self.fonts.mono)
        k_entry.pack(side="right", ipady=3)
        self.k401_var.trace_add("write", lambda *a: self.on_k401_change())
        k_entry.bind("<FocusOut>", lambda e: self.k401_var.set(fix2(self.k401_var.get())))

        tk.Frame(outer, bg=LINE, height=1).pack(fill="x")
        total_box = tk.Frame(outer, bg=PANEL2)
        total_box.pack(fill="x")
        tk.Label(total_box, text="JIB + 401K C&L", bg=PANEL2, fg=MUTED,
                 font=self.fonts.cond_small, anchor="w").pack(fill="x", padx=15,
                                                              pady=(12, 0))
        self.total_label = tk.Label(total_box, text="$0.00", bg=PANEL2, fg=GOOD,
                                    font=self.fonts.mono_big, anchor="w")
        self.total_label.pack(fill="x", padx=15, pady=(0, 13))

        tk.Frame(outer, bg=LINE, height=1).pack(fill="x")
        buttons = tk.Frame(outer, bg=PANEL)
        buttons.pack(fill="x", padx=12, pady=12)
        grid_buttons = [
            ("Save period", self.save_period, "accent"),
            ("Printable summary", self.print_summary, "normal"),
            ("Export CSV", self.export_csv, "normal"),
            ("Clear amounts", self.clear_amounts, "normal"),
            ("Backup", self.backup, "ghost"),
            ("Restore", self.restore, "ghost"),
        ]
        for index, (text, command, kind) in enumerate(grid_buttons):
            button = flat_button(buttons, text, command, kind, fonts=self.fonts)
            button.grid(row=index // 2, column=index % 2, sticky="ew", padx=3, pady=3)
        buttons.columnconfigure(0, weight=1)
        buttons.columnconfigure(1, weight=1)

    # ------------------------------------------------------------------
    # rendering
    # ------------------------------------------------------------------
    def render_employees(self):
        self._building = True
        for child in self.emp_grid.winfo_children():
            child.destroy()
        self._emp_widgets = []

        headers = [("EMPLOYEE", "w"), ("ROLE", "w"), ("GROSS WAGES", "e"),
                   ("DEDUCTIONS", "e"), ("TGW", "e"), ("", "w"), ("", "w")]
        for col, (text, anchor) in enumerate(headers):
            tk.Label(self.emp_grid, text=text, bg=PANEL, fg=MUTED,
                     font=self.fonts.cond_small, anchor=anchor
                     ).grid(row=0, column=col, sticky="ew", padx=4, pady=(0, 6))

        role_names = [r["name"] for r in self.state_data["roles"]]
        role_ids = [r["id"] for r in self.state_data["roles"]]

        if not self.state_data["employees"]:
            tk.Label(self.emp_grid, text="No employees yet. Use + Add employee.",
                     bg=PANEL, fg=MUTED, font=self.fonts.sans
                     ).grid(row=1, column=0, columnspan=7, pady=18)

        for index, emp in enumerate(self.state_data["employees"]):
            row = index + 1
            fg = INK if emp.get("active", True) else MUTED

            name_var = tk.StringVar(value=emp["name"])
            name_entry = tk.Entry(self.emp_grid, textvariable=name_var, bg=PANEL, fg=fg,
                                  insertbackground=ACCENT, relief="flat",
                                  highlightbackground=LINE, highlightcolor=ACCENT,
                                  highlightthickness=1, font=self.fonts.sans)
            name_entry.grid(row=row, column=0, sticky="ew", padx=4, pady=2, ipady=3)

            role_var = tk.StringVar(
                value=role_by_id(self.state_data, emp["role"])["name"])
            role_combo = ttk.Combobox(self.emp_grid, textvariable=role_var,
                                      values=role_names, state="readonly", width=12,
                                      style="Dark.TCombobox", font=self.fonts.sans)
            role_combo.grid(row=row, column=1, sticky="ew", padx=4, pady=2)

            wages_var = tk.StringVar(value=fix2(emp["wages"]))
            wages_entry = tk.Entry(self.emp_grid, textvariable=wages_var, bg=PANEL, fg=fg,
                                   insertbackground=ACCENT, relief="flat", width=12,
                                   justify="right", highlightbackground=LINE,
                                   highlightcolor=ACCENT, highlightthickness=1,
                                   font=self.fonts.mono)
            wages_entry.grid(row=row, column=2, sticky="ew", padx=4, pady=2, ipady=3)

            deds_var = tk.StringVar(value=fix2(emp["deds"]))
            deds_entry = tk.Entry(self.emp_grid, textvariable=deds_var, bg=PANEL, fg=fg,
                                  insertbackground=ACCENT, relief="flat", width=12,
                                  justify="right", highlightbackground=LINE,
                                  highlightcolor=ACCENT, highlightthickness=1,
                                  font=self.fonts.mono)
            deds_entry.grid(row=row, column=3, sticky="ew", padx=4, pady=2, ipady=3)

            tgw = tk.Label(self.emp_grid,
                           text=money(round2(parse_num(emp["wages"]) - parse_num(emp["deds"]))),
                           bg=PANEL, fg=MUTED, font=self.fonts.mono, anchor="e", width=12)
            tgw.grid(row=row, column=4, sticky="ew", padx=4)

            emp_id = emp["id"]
            toggle = tk.Button(
                self.emp_grid, text=("On" if emp.get("active", True) else "Off"),
                command=lambda i=emp_id: self.toggle_employee(i),
                bg=PANEL, fg=(GOOD if emp.get("active", True) else MUTED),
                activebackground=PANEL2, activeforeground=INK, relief="flat", bd=0,
                width=4, cursor="hand2", font=self.fonts.sans_small,
                highlightthickness=0)
            toggle.grid(row=row, column=5, padx=2)

            remove = tk.Button(self.emp_grid, text="X",
                               command=lambda i=emp_id: self.remove_employee(i),
                               bg=PANEL, fg=MUTED, activebackground=PANEL2,
                               activeforeground=DANGER, relief="flat", bd=0, width=3,
                               cursor="hand2", font=self.fonts.sans_small,
                               highlightthickness=0)
            remove.grid(row=row, column=6, padx=(2, 0))

            name_var.trace_add("write",
                               lambda *a, i=emp_id, v=name_var: self.on_field(i, "name", v))
            wages_var.trace_add("write",
                                lambda *a, i=emp_id, v=wages_var, lab=tgw:
                                self.on_field(i, "wages", v, lab))
            deds_var.trace_add("write",
                               lambda *a, i=emp_id, v=deds_var, lab=tgw:
                               self.on_field(i, "deds", v, lab))
            wages_entry.bind("<FocusOut>",
                             lambda e, v=wages_var: v.set(fix2(v.get())))
            deds_entry.bind("<FocusOut>",
                            lambda e, v=deds_var: v.set(fix2(v.get())))
            role_combo.bind("<<ComboboxSelected>>",
                            lambda e, i=emp_id, v=role_var, ids=role_ids,
                            names=role_names: self.on_role(i, ids[names.index(v.get())]))
            for widget in (name_entry, wages_entry, deds_entry):
                widget.bind("<Control-v>", lambda e, i=emp_id: self.on_grid_paste(e, i))
                widget.bind("<Control-V>", lambda e, i=emp_id: self.on_grid_paste(e, i))

            name_entry.bind("<FocusOut>",
                            lambda e, i=emp_id: self.on_name_commit(i))
            name_entry.bind("<Return>", lambda e: self.focus_set())

            self._emp_widgets.append({
                "id": emp_id, "tgw": tgw,
                "fields": {"name": name_entry, "role": role_combo,
                           "wages": wages_entry, "deds": deds_entry}})

        self._building = False

    def render_rates(self):
        self._building = True
        for child in self.rate_grid.winfo_children():
            child.destroy()

        def rate_row(row, label, sub, code_var, pct_var, key, kind):
            name_box = tk.Frame(self.rate_grid, bg=PANEL)
            name_box.grid(row=row, column=0, sticky="ew", pady=6)
            tk.Label(name_box, text=label.upper(), bg=PANEL,
                     fg=(WARN if kind == "jib" else INK),
                     font=self.fonts.cond, anchor="w").pack(anchor="w")
            tk.Label(name_box, text=sub, bg=PANEL, fg=MUTED,
                     font=self.fonts.sans_small, anchor="w").pack(anchor="w")

            if code_var is not None:
                code_entry = tk.Entry(self.rate_grid, textvariable=code_var, bg=FIELD,
                                      fg=MUTED, insertbackground=ACCENT, relief="flat",
                                      width=6, justify="center",
                                      highlightbackground=LINE, highlightcolor=ACCENT,
                                      highlightthickness=1, font=self.fonts.mono)
                code_entry.grid(row=row, column=1, padx=8, ipady=3)
                code_var.trace_add("write",
                                   lambda *a, k=key, v=code_var: self.on_code(k, v))
            else:
                tk.Label(self.rate_grid, text="", bg=PANEL, width=6
                         ).grid(row=row, column=1, padx=8)

            pct_box = tk.Frame(self.rate_grid, bg=PANEL)
            pct_box.grid(row=row, column=2, sticky="e")
            pct_entry = tk.Entry(pct_box, textvariable=pct_var, bg=FIELD, fg=INK,
                                 insertbackground=ACCENT, relief="flat", width=9,
                                 justify="right", highlightbackground=LINE,
                                 highlightcolor=ACCENT, highlightthickness=1,
                                 font=self.fonts.mono)
            pct_entry.pack(side="left", ipady=3)
            tk.Label(pct_box, text="%", bg=PANEL, fg=MUTED,
                     font=self.fonts.sans).pack(side="left", padx=(4, 0))
            pct_entry.bind("<FocusIn>",
                           lambda e, k=(kind, key), v=pct_var: self._pct_prev.__setitem__(k, v.get()))
            pct_entry.bind("<FocusOut>",
                           lambda e, k=kind, i=key, v=pct_var: self.on_pct_commit(k, i, v))
            pct_entry.bind("<Return>", lambda e: self.focus_set())

        for index, role in enumerate(self.state_data["roles"]):
            count = sum(1 for e in self.state_data["employees"]
                        if e.get("active", True) and e["role"] == role["id"])
            code_var = tk.StringVar(value=role["code"])
            pct_var = tk.StringVar(value=pct_text(role["pct"]))
            rate_row(index, role["name"],
                     "1 employee" if count == 1 else "{} employees".format(count),
                     code_var, pct_var, role["id"], "role")

        tk.Frame(self.rate_grid, bg=LINE, height=1).grid(
            row=len(self.state_data["roles"]), column=0, columnspan=3,
            sticky="ew", pady=(8, 4))
        jib_var = tk.StringVar(value=pct_text(self.state_data["jibPct"]))
        rate_row(len(self.state_data["roles"]) + 1, "JIB",
                 "Percent of total gross wages, all classifications",
                 None, jib_var, "__JIB", "jib")
        self._building = False

    def render_history(self):
        for child in self.hist_grid.winfo_children():
            child.destroy()
        history = self.state_data["history"]
        self.hist_count.configure(text="({})".format(len(history)))
        if not history:
            tk.Label(self.hist_grid,
                     text=("Nothing saved yet. Fill in a period and choose "
                           "Save period to keep a copy of the totals."),
                     bg=PANEL, fg=MUTED, font=self.fonts.sans_small, anchor="w"
                     ).grid(row=0, column=0, sticky="w", pady=8)
            return
        for index, entry in enumerate(history):
            box = tk.Frame(self.hist_grid, bg=PANEL)
            box.grid(row=index, column=0, sticky="ew", pady=3)
            tk.Label(box, text=entry.get("label") or "(untitled period)", bg=PANEL,
                     fg=INK, font=self.fonts.sans, anchor="w").pack(anchor="w")
            tk.Label(box, text=entry.get("saved", ""), bg=PANEL, fg=MUTED,
                     font=self.fonts.sans_small, anchor="w").pack(anchor="w")
            tk.Label(self.hist_grid, text=money(entry.get("jibTotal", 0)), bg=PANEL,
                     fg=GOOD, font=self.fonts.mono, anchor="e"
                     ).grid(row=index, column=1, sticky="e", padx=10)
            tk.Button(self.hist_grid, text="X",
                      command=lambda i=entry.get("id"): self.delete_history(i),
                      bg=PANEL, fg=MUTED, activebackground=PANEL2,
                      activeforeground=DANGER, relief="flat", bd=0, width=3,
                      cursor="hand2", font=self.fonts.sans_small, highlightthickness=0
                      ).grid(row=index, column=2)
            tk.Frame(self.hist_grid, bg=LINE, height=1).grid(
                row=index, column=0, columnspan=3, sticky="sew")

    def render_slip(self):
        result = compute(self.state_data)
        for child in self.class_frame.winfo_children():
            child.destroy()

        for role in self.state_data["roles"]:
            bucket = result["byRole"][role["id"]]
            line = tk.Frame(self.class_frame, bg=PANEL)
            line.pack(fill="x", pady=3)
            left = tk.Frame(line, bg=PANEL)
            left.pack(side="left")
            tk.Label(left, text=role["name"].upper(), bg=PANEL,
                     fg=(INK if bucket["owed"] else MUTED),
                     font=self.fonts.cond).pack(side="left")
            tk.Label(left, text=role["code"], bg=PANEL, fg=MUTED,
                     font=self.fonts.mono_small).pack(side="left", padx=(6, 0))
            right = tk.Frame(line, bg=PANEL)
            right.pack(side="right")
            tk.Label(right, text=money(bucket["owed"]), bg=PANEL,
                     fg=(INK if bucket["owed"] else MUTED),
                     font=self.fonts.mono, anchor="e").pack(anchor="e")
            tk.Label(right, text="{} x {}%".format(money(bucket["tgw"]),
                                                   pct_text(role["pct"])),
                     bg=PANEL, fg=MUTED, font=self.fonts.mono_small,
                     anchor="e").pack(anchor="e")

        self.jib_tag.configure(text=pct_text(result["jibPct"]) + "%")
        self.jib_amount.configure(text=money(result["jib"]),
                                  fg=(INK if result["jib"] else MUTED))
        self.jib_basis.configure(text="{}% x {} gross".format(
            pct_text(result["jibPct"]), money(result["grossAll"])))
        self.total_label.configure(text=money(result["jibTotal"]))
        self.slip_period.configure(
            text=self.state_data["period"].strip() or "No period entered")

    # ------------------------------------------------------------------
    # field handlers
    # ------------------------------------------------------------------
    def find_employee(self, emp_id):
        for emp in self.state_data["employees"]:
            if emp["id"] == emp_id:
                return emp
        return None

    def on_field(self, emp_id, field, var, tgw_label=None):
        if self._building:
            return
        emp = self.find_employee(emp_id)
        if emp is None:
            return
        emp[field] = var.get() if field == "name" else parse_num(var.get())
        if tgw_label is not None:
            tgw_label.configure(text=money(round2(parse_num(emp["wages"])
                                                  - parse_num(emp["deds"]))))
        self.render_slip()
        self.queue_save()

    def on_role(self, emp_id, role_id):
        if self._building:
            return
        emp = self.find_employee(emp_id)
        if emp is None:
            return
        emp["role"] = role_id
        self.render_rates()
        self.render_slip()
        self.queue_save()

    def on_code(self, role_id, var):
        if self._building:
            return
        role_by_id(self.state_data, role_id)["code"] = var.get().upper()
        self.render_slip()
        self.queue_save()

    def on_pct_commit(self, kind, key, var):
        """Percentages only change after the user confirms."""
        if self._building or self._confirm_open:
            return
        previous = self._pct_prev.get((kind, key))
        if previous is None:
            return
        old_value = parse_num(previous)
        new_value = parse_num(var.get())
        if abs(new_value - old_value) < 1e-9:
            var.set(pct_text(old_value))
            return

        label = "JIB" if kind == "jib" else role_by_id(self.state_data, key)["name"]
        affects = ("the JIB amount" if kind == "jib"
                   else "every {} premium".format(label))
        self._confirm_open = True
        try:
            confirmed = ask_confirm(
                self, self.fonts,
                "Change the {} rate?".format(label),
                "This moves it from {}% to {}% and recalculates {} from here on. "
                "Periods you have already saved keep their original figures."
                .format(pct_text(old_value), pct_text(new_value), affects),
                yes="Change rate")
        finally:
            self._confirm_open = False

        if not confirmed:
            var.set(pct_text(old_value))
            return

        if kind == "jib":
            self.state_data["jibPct"] = new_value
        else:
            role_by_id(self.state_data, key)["pct"] = new_value
        var.set(pct_text(new_value))
        self._pct_prev[(kind, key)] = var.get()
        self.render_slip()
        self.queue_save()
        self.set_status("{} rate is now {}%".format(label, pct_text(new_value)))

    def on_period_change(self):
        if self._building:
            return
        self.state_data["period"] = self.period_var.get()
        self.render_slip()
        self.queue_save()

    def on_k401_change(self):
        if self._building:
            return
        self.state_data["k401"] = parse_num(self.k401_var.get())
        self.render_slip()
        self.queue_save()

    # ------------------------------------------------------------------
    # employee actions
    # ------------------------------------------------------------------
    def add_employee(self):
        self.state_data["employees"].append({
            "id": new_id(), "name": "", "role": self.state_data["roles"][1]["id"],
            "wages": 0.0, "deds": 0.0, "active": True})
        self.sort_employees()
        self.render_employees()
        self.render_rates()
        self.render_slip()
        self.queue_save()
        self.refocus(self.state_data["employees"][-1]["id"], "name")

    def remove_employee(self, emp_id):
        emp = self.find_employee(emp_id)
        if emp is None:
            return
        who = emp["name"].strip() or "this employee"
        if not ask_confirm(self, self.fonts, "Remove {}?".format(who),
                           "{} comes out of the totals right away. To keep the record "
                           "but stop counting them, use the On/Off button instead."
                           .format(who), yes="Remove", danger=True):
            return
        self.state_data["employees"] = [e for e in self.state_data["employees"]
                                        if e["id"] != emp_id]
        self.render_employees()
        self.render_rates()
        self.render_slip()
        self.queue_save()

    def toggle_employee(self, emp_id):
        emp = self.find_employee(emp_id)
        if emp is None:
            return
        emp["active"] = not emp.get("active", True)
        self.render_employees()
        self.render_rates()
        self.render_slip()
        self.queue_save()

    def sort_employees(self):
        """The list stays alphabetical. Someone still being typed in has no name
        yet, so blanks sit at the bottom until there is something to sort on."""
        self.state_data["employees"].sort(
            key=lambda e: (not e["name"].strip(), e["name"].strip().lower()))

    def order_key(self):
        return [e["id"] for e in self.state_data["employees"]]

    def on_name_commit(self, emp_id):
        """Called when a name field is left. Re-orders only if the row moves."""
        if self._building:
            return
        self.after_idle(self._apply_order)

    def _apply_order(self):
        before = self.order_key()
        self.sort_employees()
        if self.order_key() == before:
            return
        # rebuilding the grid destroys the widget that has focus, so note where
        # the cursor is now and put it back on the same field afterwards
        try:
            focused = self.focus_get()
        except (KeyError, tk.TclError):
            focused = None
        target = None
        for record in self._emp_widgets:
            for field, widget in record["fields"].items():
                if widget is focused:
                    target = (record["id"], field)
                    break
            if target:
                break
        self.render_employees()
        self.render_rates()
        self.render_slip()
        self.queue_save()
        if target:
            self.refocus(*target)

    def refocus(self, emp_id, field):
        for record in self._emp_widgets:
            if record["id"] == emp_id:
                widget = record["fields"].get(field)
                if widget is not None:
                    widget.focus_set()
                return

    # ------------------------------------------------------------------
    # excel paste
    # ------------------------------------------------------------------
    def on_grid_paste(self, event, emp_id):
        """Ctrl+V on a row: fill down from that cell when a block was copied."""
        try:
            data = self.clipboard_get()
        except tk.TclError:
            return None
        grid = parse_tsv(data)
        if not grid or (len(grid) == 1 and len(grid[0]) == 1):
            return None  # single value - let Tk paste it normally

        widget = event.widget
        info = widget.grid_info()
        start_row = int(info.get("row", 1)) - 1
        start_col = {0: 0, 1: 1, 2: 2, 3: 3}.get(int(info.get("column", 0)), 0)
        columns = ["name", "role", "wages", "deds"]

        filled, overflow = 0, 0
        for row_offset, row in enumerate(grid):
            index = start_row + row_offset
            if index >= len(self.state_data["employees"]):
                overflow += 1
                continue
            emp = self.state_data["employees"][index]
            for col_offset, value in enumerate(row):
                position = start_col + col_offset
                if position >= len(columns):
                    continue
                field = columns[position]
                text = str(value).strip()
                if field == "name":
                    if text:
                        emp["name"] = text
                        filled += 1
                elif field == "role":
                    role = parse_role(self.state_data, text)
                    if role:
                        emp["role"] = role
                        filled += 1
                else:
                    emp[field] = parse_num(text)
                    filled += 1

        self.sort_employees()
        self.render_employees()
        self.render_rates()
        self.render_slip()
        self.queue_save()
        if overflow:
            self.set_status("Filled {} cells - {} row(s) had no employee to land on"
                            .format(filled, overflow), warn=True)
        else:
            self.set_status("Filled {} cells from your spreadsheet".format(filled))
        return "break"

    def open_import(self):
        prefill = ""
        try:
            candidate = self.clipboard_get()
            if "\t" in candidate:
                prefill = candidate
        except tk.TclError:
            pass

        dialog = ImportDialog(self, self.fonts, self.state_data, prefill)
        self.wait_window(dialog)
        if not dialog.result:
            return

        rows = dialog.result["rows"]
        updated, added, touched = 0, 0, set()
        for entry in rows:
            if entry["target"] == "__skip":
                continue
            if entry["target"] == "__new":
                emp = {"id": new_id(), "name": entry["name"],
                       "role": entry["role"] or self.state_data["roles"][1]["id"],
                       "wages": entry["wages"] or 0.0,
                       "deds": entry["deds"] or 0.0, "active": True}
                self.state_data["employees"].append(emp)
                touched.add(emp["id"])
                added += 1
            else:
                emp = self.find_employee(entry["target"])
                if emp is None:
                    continue
                if entry["wages"] is not None:
                    emp["wages"] = entry["wages"]
                if entry["deds"] is not None:
                    emp["deds"] = entry["deds"]
                if entry["role"]:
                    emp["role"] = entry["role"]
                touched.add(emp["id"])
                updated += 1

        zeroed = 0
        if dialog.result["zero_rest"]:
            for emp in self.state_data["employees"]:
                if emp["id"] not in touched and (parse_num(emp["wages"])
                                                 or parse_num(emp["deds"])):
                    emp["wages"] = 0.0
                    emp["deds"] = 0.0
                    zeroed += 1

        self.sort_employees()
        self.render_employees()
        self.render_rates()
        self.render_slip()
        self.queue_save()
        bits = []
        if updated:
            bits.append("updated {}".format(updated))
        if added:
            bits.append("added {}".format(added))
        if zeroed:
            bits.append("zeroed {}".format(zeroed))
        self.set_status("Imported - " + ", ".join(bits) if bits else "Nothing changed")

    # ------------------------------------------------------------------
    # period actions
    # ------------------------------------------------------------------
    def save_period(self):
        result = compute(self.state_data)
        self.state_data["history"].insert(0, {
            "id": new_id(),
            "label": self.state_data["period"].strip(),
            "saved": time.strftime("%b %d, %Y"),
            "premiums": result["premiums"],
            "grossAll": result["grossAll"],
            "jibPct": result["jibPct"],
            "jib": result["jib"],
            "k401": result["k401"],
            "jibTotal": result["jibTotal"],
            "lines": [{"name": r["name"], "code": r["code"],
                       "pct": parse_num(r["pct"]),
                       "tgw": result["byRole"][r["id"]]["tgw"],
                       "owed": result["byRole"][r["id"]]["owed"]}
                      for r in self.state_data["roles"]],
        })
        self.render_history()
        self.queue_save()
        self.set_status("Period saved")

    def delete_history(self, entry_id):
        self.state_data["history"] = [h for h in self.state_data["history"]
                                      if h.get("id") != entry_id]
        self.render_history()
        self.queue_save()

    def clear_amounts(self):
        if not ask_confirm(self, self.fonts, "Start a new pay period?",
                           "Wages, deductions and the 401K amount go back to zero. "
                           "Employees, roles and every rate stay exactly as they are.",
                           yes="Clear amounts"):
            return
        for emp in self.state_data["employees"]:
            emp["wages"] = 0.0
            emp["deds"] = 0.0
        self.state_data["k401"] = 0.0
        self.state_data["period"] = ""
        self._building = True
        self.k401_var.set("0.00")
        self.period_var.set("")
        self._building = False
        self.render_employees()
        self.render_slip()
        self.queue_save()
        self.set_status("Cleared - ready for the next period")

    # ------------------------------------------------------------------
    # export
    # ------------------------------------------------------------------
    def default_tag(self):
        tag = self.state_data["period"].strip() or time.strftime("%Y-%m-%d")
        return re.sub(r"[^A-Za-z0-9]+", "-", tag).strip("-") or "period"

    def export_csv(self):
        path = filedialog.asksaveasfilename(
            parent=self, defaultextension=".csv",
            initialfile="eesisp-{}.csv".format(self.default_tag()),
            filetypes=[("CSV file", "*.csv"), ("All files", "*.*")])
        if not path:
            return
        result = compute(self.state_data)
        try:
            with open(path, "w", newline="", encoding="utf-8") as handle:
                writer = csv.writer(handle)
                writer.writerow(["EESISP Remittance", self.state_data["period"].strip()])
                writer.writerow([])
                writer.writerow(["Employee", "Role", "Gross wages", "Deductions",
                                 "TGW", "Status"])
                for emp in self.state_data["employees"]:
                    writer.writerow([
                        emp["name"], role_by_id(self.state_data, emp["role"])["name"],
                        fix2(emp["wages"]), fix2(emp["deds"]),
                        fix2(parse_num(emp["wages"]) - parse_num(emp["deds"])),
                        "Active" if emp.get("active", True) else "Inactive"])
                writer.writerow([])
                writer.writerow(["Class", "Code", "Employees", "Gross wages",
                                 "Deductions", "TGW", "Rate %", "Owed"])
                for role in self.state_data["roles"]:
                    bucket = result["byRole"][role["id"]]
                    writer.writerow([role["name"], role["code"], bucket["count"],
                                     fix2(bucket["wages"]), fix2(bucket["deds"]),
                                     fix2(bucket["tgw"]), pct_text(role["pct"]),
                                     fix2(bucket["owed"])])
                writer.writerow([])
                writer.writerow(["Total gross wages", fix2(result["grossAll"])])
                writer.writerow(["JIB rate %", pct_text(result["jibPct"])])
                writer.writerow(["JIB", fix2(result["jib"])])
                writer.writerow(["401K C&L", fix2(result["k401"])])
                writer.writerow(["JIB + 401K C&L TOTAL", fix2(result["jibTotal"])])
        except OSError as exc:
            messagebox.showerror("Could not save", str(exc), parent=self)
            return
        self.set_status("CSV saved")

    def summary_html(self):
        result = compute(self.state_data)
        esc = html.escape
        rows = "".join(
            "<tr><td>{}</td><td>{}</td><td class=n>{}</td><td class=n>{}</td>"
            "<td class=n>{}</td></tr>".format(
                esc(e["name"]), esc(role_by_id(self.state_data, e["role"])["name"]),
                money(parse_num(e["wages"])), money(parse_num(e["deds"])),
                money(round2(parse_num(e["wages"]) - parse_num(e["deds"]))))
            for e in self.state_data["employees"] if e.get("active", True))
        classes = "".join(
            "<tr><td>{} <small>{}</small></td><td class=n>{}</td>"
            "<td class=n>{}%</td><td class=n>{}</td></tr>".format(
                esc(r["name"]), esc(r["code"]),
                money(result["byRole"][r["id"]]["tgw"]), pct_text(r["pct"]),
                money(result["byRole"][r["id"]]["owed"]))
            for r in self.state_data["roles"])
        return """<!doctype html><meta charset="utf-8">
<title>EESISP Remittance</title>
<style>
body{{font-family:Arial,Helvetica,sans-serif;color:#000;margin:32px;font-size:12pt}}
h1{{font-size:19pt;margin:0}} .sub{{color:#555;font-size:10pt;margin:2px 0 18px}}
table{{border-collapse:collapse;width:100%;margin-bottom:22px}}
th{{text-align:left;font-size:9pt;letter-spacing:.08em;text-transform:uppercase;
color:#555;border-bottom:1px solid #999;padding:5px 7px}}
td{{padding:5px 7px;border-bottom:1px solid #ddd}}
td.n,th.n{{text-align:right;font-family:Consolas,monospace}}
small{{color:#777}}
.total{{border:2px solid #000;padding:11px 14px;display:inline-block;margin-top:6px}}
.total .cap{{font-size:9pt;letter-spacing:.1em;text-transform:uppercase;color:#555}}
.total .val{{font-size:22pt;font-family:Consolas,monospace;font-weight:bold}}
@media print{{body{{margin:12mm}}}}
</style>
<h1>EESISP Remittance</h1>
<div class=sub>{company} &middot; {period}</div>
<table><tr><th>Employee</th><th>Role</th><th class=n>Gross wages</th>
<th class=n>Deductions</th><th class=n>TGW</th></tr>{rows}</table>
<table><tr><th>Classification</th><th class=n>TGW</th><th class=n>Rate</th>
<th class=n>Premium</th></tr>{classes}</table>
<table><tr><th>&nbsp;</th><th class=n>&nbsp;</th></tr>
<tr><td>Total gross wages</td><td class=n>{gross}</td></tr>
<tr><td>JIB ({jibpct}% of gross wages)</td><td class=n>{jib}</td></tr>
<tr><td>401K C&amp;L</td><td class=n>{k401}</td></tr></table>
<div class=total><div class=cap>JIB + 401K C&amp;L</div>
<div class=val>{total}</div></div>
""".format(company=esc(COMPANY),
           period=esc(self.state_data["period"].strip() or "No period entered"),
           rows=rows, classes=classes, gross=money(result["grossAll"]),
           jibpct=pct_text(result["jibPct"]), jib=money(result["jib"]),
           k401=money(result["k401"]), total=money(result["jibTotal"]))

    def print_summary(self):
        try:
            path = os.path.join(tempfile.gettempdir(),
                                "eesisp-{}.html".format(self.default_tag()))
            with open(path, "w", encoding="utf-8") as handle:
                handle.write(self.summary_html())
            webbrowser.open("file://" + os.path.abspath(path))
            self.set_status("Summary opened in your browser - print from there")
        except OSError as exc:
            messagebox.showerror("Could not open summary", str(exc), parent=self)

    def backup(self):
        path = filedialog.asksaveasfilename(
            parent=self, defaultextension=".json",
            initialfile="eesisp-backup-{}.json".format(time.strftime("%Y-%m-%d")),
            filetypes=[("JSON file", "*.json"), ("All files", "*.*")])
        if not path:
            return
        try:
            with open(path, "w", encoding="utf-8") as handle:
                json.dump(self.state_data, handle, indent=2)
        except OSError as exc:
            messagebox.showerror("Could not save", str(exc), parent=self)
            return
        self.set_status("Backup saved")

    def restore(self):
        path = filedialog.askopenfilename(
            parent=self, filetypes=[("JSON file", "*.json"), ("All files", "*.*")])
        if not path:
            return
        try:
            with open(path, "r", encoding="utf-8") as handle:
                raw = json.load(handle)
        except Exception as exc:
            messagebox.showerror("Could not read that file", str(exc), parent=self)
            return
        if not isinstance(raw, dict) or "employees" not in raw:
            messagebox.showerror("Not a backup",
                                 "That file is not an EESISP backup. Pick a file "
                                 "saved with Backup.", parent=self)
            return
        if not ask_confirm(self, self.fonts, "Restore this backup?",
                           "Everything on screen is replaced with the contents of "
                           "that file.", yes="Restore", danger=True):
            return
        self.state_data = normalise_state(raw)
        self.sort_employees()
        self._building = True
        self.period_var.set(self.state_data["period"])
        self.k401_var.set(fix2(self.state_data["k401"]))
        self._building = False
        self.render_employees()
        self.render_rates()
        self.render_history()
        self.render_slip()
        self.queue_save()
        self.set_status("Backup restored")

    # ------------------------------------------------------------------
    # saving / status
    # ------------------------------------------------------------------
    def queue_save(self):
        if self._save_job is not None:
            self.after_cancel(self._save_job)
        self._save_job = self.after(500, self.persist)

    def persist(self):
        self._save_job = None
        try:
            save_state(self.state_data)
            self.set_status("Saved " + time.strftime("%I:%M %p").lstrip("0"))
        except OSError as exc:
            self.set_status("Not saved - {}".format(exc), warn=True)

    def set_status(self, message, warn=False):
        self.status.configure(text=message.upper(), fg=(WARN if warn else GOOD))
        if self._status_job is not None:
            self.after_cancel(self._status_job)
        self._status_job = self.after(6000, lambda: self.status.configure(text=""))

    def on_close(self):
        if self._save_job is not None:
            self.after_cancel(self._save_job)
        try:
            save_state(self.state_data)
        except OSError:
            pass
        self.destroy()


def main():
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
