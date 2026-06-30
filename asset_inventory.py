"""
Server Asset Inventory  v2.0  — Three-Tier Access
====================================================
Run  :  python asset_inventory.py
Open :  http://localhost:5050

TIERS
  Public      — Dashboard, Search, Bulk Lookup, Month Compare (no login)
  User        — + Add Tags, Add Notes, Flag servers   (login required)
  Admin       — + Upload Excel, Manage Users, Reports  (admin role)

100% OFFLINE · Single file · No internet required
Packages: flask, pandas, numpy, matplotlib, werkzeug, openpyxl
"""

import io, json, warnings, logging, re, base64, os, secrets
from datetime import datetime, timedelta
from functools import wraps
from pathlib import Path

import pandas as pd
import numpy as np
from flask import (Flask, request, render_template_string,
                   jsonify, session, redirect, send_file, make_response)
from werkzeug.security import generate_password_hash, check_password_hash

warnings.filterwarnings("ignore")
logging.basicConfig(level=logging.INFO,
                    format="%(asctime)s  %(levelname)-7s  %(message)s",
                    datefmt="%H:%M:%S")
log = logging.getLogger(__name__)

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import numpy as np

# ── File paths ────────────────────────────────────────────────────────────────
_USERS_FILE = Path("users.json")
_KEY_FILE   = Path(".flask_secret")
_AUDIT_FILE = Path("audit.log")
_TAGS_FILE   = Path("tags.json")
_NOTES_FILE  = Path("notes.json")
_FLAGS_FILE  = Path("flags.json")
_CORR_FILE   = Path("corrections.json")
_DECOM_FILE  = Path("decommissions.json")
_NEWSRV_FILE = Path("new_servers.json")
_TSHIRT_FILE = Path("tshirt_changes.json")
_INVENTORY_FILE      = Path("current_inventory.xlsx")
_INVENTORY_META_FILE = Path("current_inventory_meta.json")
_INVENTORY_PREV_FILE = Path("prev_inventory.xlsx")
_INVENTORY_PREV_META = Path("prev_inventory_meta.json")
_ARCHIVE_DIR = Path("archive")

# ── Colour palette ────────────────────────────────────────────────────────────
BLUE   = "#4a8cff"; RED    = "#ff4f6a"; GREEN  = "#30d988"
YELLOW = "#ffc240"; PURPLE = "#a78bfa"; CYAN   = "#22d3ee"
ORANGE = "#fb923c"; MUTED  = "#7b8db0"; TEXT   = "#edf2ff"
BG     = "#07090f"; SURFACE= "#111827"; BORDER = "#1e2a40"
PALETTE = [BLUE, GREEN, YELLOW, PURPLE, CYAN, ORANGE, RED,
           "#f472b6", "#34d399", "#f87171", "#60a5fa", "#a3e635"]
PLATFORM_COLORS = {
    "Dynamo": BLUE, "Digital Journey": GREEN,
    "EMA": PURPLE, "EPMC": CYAN, "JEA": ORANGE,
}

plt.rcParams.update({
    "figure.facecolor": BG,    "axes.facecolor":  SURFACE,
    "axes.edgecolor":   BORDER,"axes.labelcolor": MUTED,
    "xtick.color":      MUTED, "ytick.color":     MUTED,
    "text.color":       TEXT,  "grid.color":      BORDER,
    "grid.linewidth": 0.6, "font.family": "DejaVu Sans",
    "font.size": 9, "axes.titlesize": 10,
    "axes.titlecolor": TEXT, "axes.titlepad": 8,
    "legend.facecolor": SURFACE, "legend.edgecolor": BORDER,
    "legend.fontsize": 8,
})

# ═══════════════════════════════════════════════════════════════════════════════
#  AUTH SYSTEM  (transplanted & adapted from app_v28_dcss.py)
# ═══════════════════════════════════════════════════════════════════════════════
def _load_users():
    if _USERS_FILE.exists():
        try:   return json.loads(_USERS_FILE.read_text())
        except: return {}
    return {}

def _save_users(u):
    _USERS_FILE.write_text(json.dumps(u, indent=2))

def _app_setup():
    """First-run bootstrap — generates secret key and default admin."""
    if _KEY_FILE.exists():
        secret = _KEY_FILE.read_bytes()
    else:
        secret = secrets.token_bytes(32)
        _KEY_FILE.write_bytes(secret)

    users = _load_users()
    if not users:
        pw = secrets.token_urlsafe(12)
        users["admin"] = {
            "password":   generate_password_hash(pw),
            "role":       "admin",
            "enabled":    True,
            "full_name":  "Administrator",
            "created_at": datetime.now().isoformat(),
        }
        _save_users(users)
        log.info("=" * 60)
        log.info("  First run — default admin account created")
        log.info("  Username : admin")
        log.info("  Password : %s", pw)
        log.info("  SAVE THIS — it will not be shown again")
        log.info("  Change via Admin Panel → Manage Users")
        log.info("=" * 60)
    return secret

def _audit(event, extra=""):
    try:
        ip   = request.remote_addr if request else "—"
        line = f"{datetime.now().isoformat()} | {event:<24} | ip={ip} | {extra}\n"
        with open(_AUDIT_FILE, "a") as f:
            f.write(line)
    except Exception:
        pass

def _current_user():
    if "username" not in session:
        return None
    if "last_active" in session:
        elapsed = (datetime.now() -
                   datetime.fromisoformat(session["last_active"])).total_seconds()
        if elapsed > 8 * 3600:          # 8-hour session timeout
            session.clear(); return None
    session["last_active"] = datetime.now().isoformat()
    users = _load_users()
    u = users.get(session["username"])
    if not u or not u.get("enabled", True):
        session.clear(); return None
    return {**u, "username": session["username"]}

def _login_required(fn):
    @wraps(fn)
    def wrapper(*args, **kwargs):
        if not _current_user():
            if request.is_json or request.method == "POST":
                return jsonify({"error": "Login required", "auth": False}), 401
            return redirect("/login?next=" + request.path)
        return fn(*args, **kwargs)
    return wrapper

def _admin_required(fn):
    @wraps(fn)
    def wrapper(*args, **kwargs):
        u = _current_user()
        if not u or u.get("role") != "admin":
            return jsonify({"error": "Admin access required"}), 403
        return fn(*args, **kwargs)
    return wrapper

# ═══════════════════════════════════════════════════════════════════════════════
#  CHART HELPERS
# ═══════════════════════════════════════════════════════════════════════════════
def fig_to_b64(fig, dpi=110):
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=dpi, bbox_inches="tight",
                facecolor=fig.get_facecolor())
    buf.seek(0)
    data = base64.b64encode(buf.read()).decode()
    plt.close(fig)
    return f"data:image/png;base64,{data}"

def make_donut(labels, values, colors=None, title="", size=(4.2, 3.6)):
    if not labels or not values or sum(values) == 0: return None
    cols  = colors or [PALETTE[i % len(PALETTE)] for i in range(len(labels))]
    total = sum(values)
    patches = [mpatches.Patch(color=c, label=f"{l} ({v})")
               for l, v, c in zip(labels, values, cols)]
    min_c  = max(1, round(total * 0.015))
    rv     = [max(v, min_c) if v > 0 else 0 for v in values]
    nz     = [(vr, c) for vr, vo, c in zip(rv, values, cols) if vo > 0]
    if not nz: return None
    nzr, nzc = zip(*nz)
    nzo    = [v for v in values if v > 0]
    expl   = [0.06 if v < total * 0.03 else 0 for v in nzo]
    fig, ax = plt.subplots(figsize=size)
    fig.patch.set_facecolor(BG); ax.set_facecolor(BG)
    ax.pie(nzr, colors=nzc, startangle=90, explode=expl,
           wedgeprops=dict(width=0.55, edgecolor=BG, linewidth=2))
    ax.text(0, 0, str(total), ha="center", va="center",
            fontsize=15, fontweight="bold", color=TEXT)
    ax.set_title(title, color=TEXT, pad=6)
    ax.legend(handles=patches, loc="lower center",
              bbox_to_anchor=(0.5, -0.22), ncol=2, framealpha=0, fontsize=7.5)
    fig.tight_layout()
    return fig_to_b64(fig)

def make_hbar(labels, values, colors=None, title="", size=(5.5, 0.45)):
    if not labels or not values: return None
    n = len(labels); h = max(2.5, n * size[1])
    fig, ax = plt.subplots(figsize=(size[0], h))
    col = colors if colors else [BLUE] * n
    if isinstance(col, str): col = [col] * n
    bars = ax.barh(list(range(n)), values, color=col,
                   height=0.6, edgecolor=BG, linewidth=0.4)
    ax.set_yticks(list(range(n))); ax.set_yticklabels(labels, fontsize=8.5)
    ax.invert_yaxis(); ax.set_title(title, color=TEXT)
    ax.grid(axis="x", alpha=0.4)
    ax.spines[["top", "right", "left"]].set_visible(False)
    mx = max(values) if values else 1
    for bar, val in zip(bars, values):
        ax.text(bar.get_width() + mx * 0.01,
                bar.get_y() + bar.get_height() / 2,
                str(val), va="center", fontsize=7.5, color=TEXT)
    fig.tight_layout()
    return fig_to_b64(fig)

def make_vbar(labels, values, colors=None, title="", size=(7, 3.2)):
    if not labels or not values: return None
    fig, ax = plt.subplots(figsize=size)
    cols = colors or [BLUE] * len(labels)
    ax.bar(list(range(len(labels))), values, color=cols,
           edgecolor=BG, linewidth=0.4, width=0.65)
    ax.set_xticks(list(range(len(labels))))
    ax.set_xticklabels(labels, rotation=30, ha="right", fontsize=8)
    ax.set_title(title, color=TEXT)
    ax.grid(axis="y", alpha=0.4)
    ax.spines[["top", "right"]].set_visible(False)
    fig.tight_layout()
    return fig_to_b64(fig)

def make_hbar_compare(labels, vals_a, vals_b, label_a="Current",
                      label_b="Previous", title="", size=(6, 0.5)):
    if not labels: return None
    n = len(labels); h = max(3, n * size[1])
    fig, ax = plt.subplots(figsize=(size[0], h))
    y = np.arange(n); bh = 0.35
    ax.barh(y + bh/2, vals_a, bh, color=BLUE,  label=label_a,
            edgecolor=BG, linewidth=0.3)
    ax.barh(y - bh/2, vals_b, bh, color=GREEN, label=label_b,
            edgecolor=BG, linewidth=0.3)
    ax.set_yticks(y); ax.set_yticklabels(labels, fontsize=8)
    ax.invert_yaxis(); ax.set_title(title, color=TEXT)
    ax.grid(axis="x", alpha=0.35)
    ax.spines[["top", "right", "left"]].set_visible(False)
    ax.legend(framealpha=0.2, fontsize=8)
    fig.tight_layout()
    return fig_to_b64(fig)

# ═══════════════════════════════════════════════════════════════════════════════
#  COLUMN DETECTION
# ═══════════════════════════════════════════════════════════════════════════════
COL_MAP = {
    "Server HostName":               ["server hostname","hostname","host name","servername",
                                      "server host name"],
    "Server Type(Physical/Virtual)": ["server type","server type(physical/virtual)",
                                      "server type (physical/virtual)",
                                      "physical/virtual","type"],
    "Platform":                      ["platform","plateform","infra","infrastructure"],
    "Server DC Location":            ["server dc location","dc location","datacenter",
                                      "data center","location"],
    "HPC or NON HPC or JPC":         ["hpc or non hpc or jpc","hpc/non hpc/jpc","hpc","group"],
    "Server Role":                   ["server role","role"],
    "Final OS":                      ["final os","os","operating system","os version"],
    "Commercial Category":           ["commercial category","category","contract type"],
    "Reference":                     ["reference","ref"],
    "Application Name":              ["application name","application","app name","app"],
}
STATUS_PAT = re.compile(
    r"status\s+as\s+on\s+(?:1st|1)\s+(\w+)\s+(\d{4})", re.IGNORECASE)

def detect_status_col(columns):
    for col in columns:
        if STATUS_PAT.search(col.strip()):
            return col
    return None

def _clean_header_key(col):
    """Normalise a raw header into a matchable key: lowercase, strip
    surrounding whitespace, trailing colons, and collapse internal
    whitespace. Handles real-world headers like 'Server Role:',
    'Reference : ', 'HPC or NON HPC or JPC  :' etc."""
    key = str(col).strip().lower()
    key = key.rstrip(":").strip()          # drop trailing colon(s)
    key = re.sub(r"\s+", " ", key)          # collapse multi-space
    return key

def normalise_columns(df):
    df.columns = [str(c).strip() for c in df.columns]
    rename, used = {}, set()
    for col in df.columns:
        key = _clean_header_key(col)
        for canon, aliases in COL_MAP.items():
            if canon in used: continue
            if key in aliases or key == canon.lower():
                rename[col] = canon; used.add(canon); break
    if rename: log.info("Column rename: %s", rename)
    return df.rename(columns=rename)

# ═══════════════════════════════════════════════════════════════════════════════
#  STORE + DATA HELPERS
# ═══════════════════════════════════════════════════════════════════════════════
STORE = {
    "df": None, "df_prev": None,
    "status_col": None, "prev_status": None,
    "filename": None, "prev_file": None,
    "uploaded_at": None, "total_rows": 0,
    "quality": {},
}

def load_excel(file_bytes, filename):
    ext = filename.rsplit(".", 1)[-1].lower()
    if ext in ("xlsx", "xls"):
        df = pd.read_excel(io.BytesIO(file_bytes), dtype=str)
    else:
        df = pd.read_csv(io.BytesIO(file_bytes), dtype=str)
    df = df.fillna("").astype(str)
    df = df[~df.apply(lambda r: r.str.strip().eq("").all(), axis=1)]
    return normalise_columns(df)

# ── Inventory persistence (survives server restart/crash) ───────────────────
def persist_inventory(df, filename, status_col, uploaded_at, total_rows, quality,
                       which="current"):
    """Save the uploaded inventory + its metadata to disk so a server restart
    or crash doesn't lose the data. Also archives a dated copy."""
    xlsx_path = _INVENTORY_FILE if which == "current" else _INVENTORY_PREV_FILE
    meta_path = _INVENTORY_META_FILE if which == "current" else _INVENTORY_PREV_META
    try:
        df.to_excel(xlsx_path, index=False)
        meta = {
            "filename": filename, "status_col": status_col,
            "uploaded_at": uploaded_at, "total_rows": total_rows,
            "quality": quality,
        }
        meta_path.write_text(json.dumps(meta, indent=2))
        if which == "current":
            archive_inventory(df, filename)
    except Exception:
        log.exception("Failed to persist inventory to disk")

def archive_inventory(df, filename):
    """Keep a dated snapshot of every uploaded inventory file for history."""
    try:
        _ARCHIVE_DIR.mkdir(exist_ok=True)
        stamp = datetime.now().strftime("%Y-%m-%d_%H%M%S")
        safe_name = re.sub(r"[^A-Za-z0-9_.-]+", "_", filename)[:60]
        archive_path = _ARCHIVE_DIR / f"{stamp}__{safe_name}"
        df.to_excel(archive_path, index=False)
    except Exception:
        log.exception("Failed to archive inventory")

def load_persisted_inventory():
    """On startup, restore the last uploaded inventory (and previous-month
    file, if any) from disk so a crash/restart doesn't require re-upload."""
    if _INVENTORY_FILE.exists() and _INVENTORY_META_FILE.exists():
        try:
            df = pd.read_excel(_INVENTORY_FILE, dtype=str).fillna("").astype(str)
            df = normalise_columns(df)
            meta = json.loads(_INVENTORY_META_FILE.read_text())
            STORE["df"]          = df
            STORE["status_col"]  = meta.get("status_col")
            STORE["filename"]    = meta.get("filename")
            STORE["uploaded_at"] = meta.get("uploaded_at")
            STORE["total_rows"]  = meta.get("total_rows", len(df))
            STORE["quality"]     = meta.get("quality", {})
            log.info("Restored inventory from disk: %s (%d rows)",
                     STORE["filename"], STORE["total_rows"])
        except Exception:
            log.exception("Failed to restore persisted inventory")

    if _INVENTORY_PREV_FILE.exists() and _INVENTORY_PREV_META.exists():
        try:
            df = pd.read_excel(_INVENTORY_PREV_FILE, dtype=str).fillna("").astype(str)
            df = normalise_columns(df)
            meta = json.loads(_INVENTORY_PREV_META.read_text())
            STORE["df_prev"]    = df
            STORE["prev_status"] = meta.get("status_col")
            STORE["prev_file"]   = meta.get("filename")
            log.info("Restored previous-month inventory from disk: %s (%d rows)",
                     STORE["prev_file"], len(df))
        except Exception:
            log.exception("Failed to restore persisted previous inventory")

def quality_check(df, sc):
    issues = {}
    for c in ["Application Name", "Commercial Category",
              "Server DC Location", "Final OS", "Platform"]:
        if c in df.columns:
            b = int((df[c].str.strip() == "").sum())
            if b: issues[c] = b
    if sc and sc in df.columns:
        b = int((df[sc].str.strip() == "").sum())
        if b: issues["Status (blank)"] = b
    return issues

def vc(series):
    s = series.replace("", np.nan).dropna()
    v = s.value_counts()
    return v.index.tolist(), v.values.tolist()

def safe_col(df, col):
    return df[col] if col in df.columns else pd.Series(dtype=str)

def analyse(df, sc):
    total = len(df)
    r = {"total": total}
    if sc and sc in df.columns:
        st = df[sc].str.strip()
        r["live"]     = int((st.str.lower() == "live").sum())
        r["not_live"] = int((st.str.lower() == "not live").sum())
        r["status_other"] = total - r["live"] - r["not_live"]
    else:
        r["live"] = r["not_live"] = r["status_other"] = 0
    if "Server Type(Physical/Virtual)" in df.columns:
        st2 = df["Server Type(Physical/Virtual)"].str.lower()
        r["physical"] = int(st2.str.contains("physical|bare", na=False).sum())
        r["virtual"]  = int(st2.str.contains("virtual|vm",    na=False).sum())
    else:
        r["physical"] = r["virtual"] = 0
    pl_l, pl_v = vc(safe_col(df, "Platform").str.strip())
    r["platform_labels"] = pl_l
    r["platform_values"] = pl_v
    r["platform_colors"] = [PLATFORM_COLORS.get(p, MUTED) for p in pl_l]
    loc_l, loc_v = vc(safe_col(df, "Server DC Location").str.strip())
    r["loc_labels"] = loc_l; r["loc_values"] = loc_v
    hpc_l, hpc_v = vc(safe_col(df, "HPC or NON HPC or JPC").str.strip())
    r["hpc_labels"] = hpc_l; r["hpc_values"] = hpc_v
    role_l, role_v = vc(safe_col(df, "Server Role").str.strip())
    r["role_labels"] = role_l; r["role_values"] = role_v
    os_l, os_v = vc(safe_col(df, "Final OS").str.strip())
    r["os_labels"] = os_l[:15]; r["os_values"] = os_v[:15]
    app_l, app_v = vc(safe_col(df, "Application Name").str.strip())
    r["app_labels"] = app_l[:15]; r["app_values"] = app_v[:15]
    return r

def compare_months(df_c, df_p, sc_c, sc_p):
    hn = "Server HostName"
    out = {"curr_total": len(df_c), "prev_total": len(df_p)}
    out["diff_total"] = out["curr_total"] - out["prev_total"]
    ch = set(df_c[hn].str.strip().str.lower()) if hn in df_c.columns else set()
    ph = set(df_p[hn].str.strip().str.lower()) if hn in df_p.columns else set()
    new_h  = ch - ph; gone_h = ph - ch
    out["new_count"] = len(new_h); out["removed_count"] = len(gone_h)
    disp = [c for c in [hn, "Platform", "Server Role",
                        "Final OS", "Server DC Location",
                        "HPC or NON HPC or JPC"] if c in df_c.columns]
    out["new_servers"]     = (df_c[df_c[hn].str.strip().str.lower().isin(new_h)][disp]
                              .fillna("").to_dict("records") if hn in df_c.columns else [])
    disp_p = [c for c in disp if c in df_p.columns]
    out["removed_servers"] = (df_p[df_p[hn].str.strip().str.lower().isin(gone_h)][disp_p]
                              .fillna("").to_dict("records") if hn in df_p.columns else [])
    out["to_live"] = out["to_not_live"] = []
    if sc_c and sc_p and hn in df_c.columns and hn in df_p.columns:
        common = ch & ph
        ci = df_c.set_index(df_c[hn].str.strip().str.lower())
        pi = df_p.set_index(df_p[hn].str.strip().str.lower())
        to_live = []; to_not = []
        for h in common:
            try:
                cs = str(ci.loc[h, sc_c] if not isinstance(ci.loc[h, sc_c], pd.Series)
                         else ci.loc[h, sc_c].iloc[0]).strip().lower()
                ps = str(pi.loc[h, sc_p] if not isinstance(pi.loc[h, sc_p], pd.Series)
                         else pi.loc[h, sc_p].iloc[0]).strip().lower()
                row = ci.loc[h]; row = row.iloc[0] if isinstance(row, pd.DataFrame) else row
                base = {hn: row.get(hn, h), "Platform": row.get("Platform", ""),
                        "Server Role": row.get("Server Role", ""),
                        "Prev Status": ps.title(), "Curr Status": cs.title()}
                if ps != "live" and cs == "live":   to_live.append(base)
                elif ps == "live" and cs != "live": to_not.append(base)
            except Exception: continue
        out["to_live"] = to_live; out["to_not_live"] = to_not
    if "Platform" in df_c.columns and "Platform" in df_p.columns:
        plats = sorted(set(df_c["Platform"].str.strip().unique()) |
                       set(df_p["Platform"].str.strip().unique()))
        plats = [p for p in plats if p]
        cp = df_c["Platform"].str.strip().value_counts()
        pp = df_p["Platform"].str.strip().value_counts()
        out["plat_labels"] = plats
        out["plat_curr"]   = [int(cp.get(p, 0)) for p in plats]
        out["plat_prev"]   = [int(pp.get(p, 0)) for p in plats]
        if sc_c and sc_p:
            clv = df_c[df_c[sc_c].str.strip().str.lower() == "live"]["Platform"].str.strip().value_counts()
            plv = df_p[df_p[sc_p].str.strip().str.lower() == "live"]["Platform"].str.strip().value_counts()
            out["live_curr"] = [int(clv.get(p, 0)) for p in plats]
            out["live_prev"] = [int(plv.get(p, 0)) for p in plats]
    else:
        out["plat_labels"] = out["plat_curr"] = out["plat_prev"] = []
        out["live_curr"]   = out["live_prev"] = []
    return out

# ═══════════════════════════════════════════════════════════════════════════════
#  TAG / NOTE / FLAG HELPERS
# ═══════════════════════════════════════════════════════════════════════════════
PRESET_TAGS = [
    "OS Incorrect", "OS Outdated", "Decommission Pending",
    "Details Missing", "Application Missing", "Platform Wrong",
    "DC Location Wrong", "Verify with Team", "Duplicate Entry",
]

def _load_json(path):
    if path.exists():
        try:   return json.loads(path.read_text())
        except: return {}
    return {}

def _save_json(path, data):
    path.write_text(json.dumps(data, indent=2))

def get_tags():         return _load_json(_TAGS_FILE)
def get_notes():        return _load_json(_NOTES_FILE)
def get_flags():        return _load_json(_FLAGS_FILE)
def get_corrections():  return _load_json(_CORR_FILE)

# Columns users are allowed to suggest corrections for
CORRECTABLE_COLS = [
    "Server Type(Physical/Virtual)",
    "Platform",
    "Server DC Location",
    "HPC or NON HPC or JPC",
    "Server Role",
    "Final OS",
    "Commercial Category",
    "Application Name",
    "Reference",
]

def add_correction(hostname, column, current_val, suggested_val, reason, username, full_name):
    corrs = get_corrections()
    entry = {
        "id":            secrets.token_hex(8),
        "hostname":      hostname.strip(),
        "column":        column.strip(),
        "current_val":   current_val.strip(),
        "suggested_val": suggested_val.strip(),
        "reason":        reason.strip(),
        "user":          username,
        "name":          full_name,
        "ts":            datetime.now().strftime("%d %b %Y %H:%M"),
        "status":        "Pending",   # Pending | Approved | Rejected
    }
    corrs[entry["id"]] = entry
    _save_json(_CORR_FILE, corrs)
    _audit("CORRECTION_ADD",
           f"user={username} host={hostname} col={column} val={suggested_val}")
    return entry

def mark_correction_decision(corr_id, decision, reason, username):
    """decision: 'Approved' or 'Rejected'"""
    corrs = get_corrections()
    if corr_id in corrs:
        corrs[corr_id]["status"]           = decision
        corrs[corr_id]["decision_reason"]  = (reason or "").strip()
        corrs[corr_id]["reviewed_by"]      = username
        corrs[corr_id]["reviewed_ts"]      = datetime.now().strftime("%d %b %Y %H:%M")
        _save_json(_CORR_FILE, corrs)
        _audit(f"CORRECTION_{decision.upper()}", f"admin={username} id={corr_id} reason={reason}")

def delete_correction(corr_id, username):
    corrs = get_corrections()
    if corr_id in corrs:
        del corrs[corr_id]
        _save_json(_CORR_FILE, corrs)
        _audit("CORRECTION_DELETE", f"admin={username} id={corr_id}")

# ── Decommission helpers ──────────────────────────────────────────────────────
DECOM_REASONS = ["Shutdown", "Decommissioned"]

def get_decommissions(): return _load_json(_DECOM_FILE)

def add_decommission(hostname, reason, eff_date, comment, username, full_name):
    decs = get_decommissions()
    entry = {
        "id":          secrets.token_hex(8),
        "hostname":    hostname.strip(),
        "reason":      reason.strip(),
        "eff_date":    eff_date.strip(),
        "comment":     comment.strip(),
        "user":        username,
        "name":        full_name,
        "ts":          datetime.now().strftime("%d %b %Y %H:%M"),
        "status":      "Pending",   # Pending | Approved | Rejected
    }
    decs[entry["id"]] = entry
    _save_json(_DECOM_FILE, decs)
    _audit("DECOM_ADD", f"user={username} host={hostname} reason={reason}")
    return entry

def mark_decom_decision(dec_id, decision, reason, username):
    decs = get_decommissions()
    if dec_id in decs:
        decs[dec_id]["status"]          = decision
        decs[dec_id]["decision_reason"] = (reason or "").strip()
        decs[dec_id]["reviewed_by"]     = username
        decs[dec_id]["reviewed_ts"]     = datetime.now().strftime("%d %b %Y %H:%M")
        _save_json(_DECOM_FILE, decs)
        _audit(f"DECOM_{decision.upper()}", f"admin={username} id={dec_id} reason={reason}")

def delete_decommission(dec_id, username):
    decs = get_decommissions()
    if dec_id in decs:
        del decs[dec_id]
        _save_json(_DECOM_FILE, decs)
        _audit("DECOM_DELETE", f"admin={username} id={dec_id}")

# ── New server suggestion helpers ─────────────────────────────────────────────
NEW_SERVER_FIELDS = [
    "Server HostName", "Server Type(Physical/Virtual)", "Platform",
    "Server DC Location", "HPC or NON HPC or JPC", "Server Role",
    "Final OS", "Commercial Category", "Reference", "Application Name",
]

def get_new_servers(): return _load_json(_NEWSRV_FILE)

def add_new_server(fields, comment, username, full_name):
    news = get_new_servers()
    entry = {
        "id":       secrets.token_hex(8),
        "fields":   {k: fields.get(k, "").strip() for k in NEW_SERVER_FIELDS},
        "comment":  comment.strip(),
        "user":     username,
        "name":     full_name,
        "ts":       datetime.now().strftime("%d %b %Y %H:%M"),
        "status":   "Pending",   # Pending | Approved | Rejected
    }
    news[entry["id"]] = entry
    _save_json(_NEWSRV_FILE, news)
    _audit("NEWSRV_ADD",
           f"user={username} host={fields.get('Server HostName','')}")
    return entry

def mark_newsrv_decision(ns_id, decision, reason, username):
    news = get_new_servers()
    if ns_id in news:
        news[ns_id]["status"]          = decision
        news[ns_id]["decision_reason"] = (reason or "").strip()
        news[ns_id]["reviewed_by"]     = username
        news[ns_id]["reviewed_ts"]     = datetime.now().strftime("%d %b %Y %H:%M")
        _save_json(_NEWSRV_FILE, news)
        _audit(f"NEWSRV_{decision.upper()}", f"admin={username} id={ns_id} reason={reason}")

def delete_new_server(ns_id, username):
    news = get_new_servers()
    if ns_id in news:
        del news[ns_id]
        _save_json(_NEWSRV_FILE, news)
        _audit("NEWSRV_DELETE", f"admin={username} id={ns_id}")

# ── T-shirt size (CPU/RAM) change helpers ─────────────────────────────────────
# Platforms where CPU/RAM directly affects contract pricing —
# only these accept T-shirt size change submissions.
TSHIRT_PLATFORMS = ["EMA", "EPMC", "JEA"]

def get_tshirt_changes(): return _load_json(_TSHIRT_FILE)

def add_tshirt_change(hostname, platform, current_cpu, current_ram,
                      new_cpu, new_ram, reason, username, full_name):
    changes = get_tshirt_changes()
    entry = {
        "id":          secrets.token_hex(8),
        "hostname":    hostname.strip(),
        "platform":    platform.strip(),
        "current_cpu": current_cpu.strip(),
        "current_ram": current_ram.strip(),
        "new_cpu":     new_cpu.strip(),
        "new_ram":     new_ram.strip(),
        "reason":      reason.strip(),
        "user":        username,
        "name":        full_name,
        "ts":          datetime.now().strftime("%d %b %Y %H:%M"),
        "status":      "Pending",   # Pending | Approved | Rejected
    }
    changes[entry["id"]] = entry
    _save_json(_TSHIRT_FILE, changes)
    _audit("TSHIRT_ADD",
           f"user={username} host={hostname} cpu={current_cpu}->{new_cpu} ram={current_ram}->{new_ram}")
    return entry

def mark_tshirt_decision(t_id, decision, reason, username):
    changes = get_tshirt_changes()
    if t_id in changes:
        changes[t_id]["status"]          = decision
        changes[t_id]["decision_reason"] = (reason or "").strip()
        changes[t_id]["reviewed_by"]     = username
        changes[t_id]["reviewed_ts"]     = datetime.now().strftime("%d %b %Y %H:%M")
        _save_json(_TSHIRT_FILE, changes)
        _audit(f"TSHIRT_{decision.upper()}", f"admin={username} id={t_id} reason={reason}")

def delete_tshirt_change(t_id, username):
    changes = get_tshirt_changes()
    if t_id in changes:
        del changes[t_id]
        _save_json(_TSHIRT_FILE, changes)
        _audit("TSHIRT_DELETE", f"admin={username} id={t_id}")

def add_tag(hostname, tag, username, full_name):
    """Tags are stored as rich objects so we know who added them and
    whether Admin has accepted them as official."""
    tags = get_tags()
    hn = hostname.strip().lower()
    if hn not in tags: tags[hn] = []
    if not any(t["tag"] == tag for t in tags[hn]):
        entry = {
            "id":     secrets.token_hex(6),
            "tag":    tag,
            "user":   username,
            "name":   full_name,
            "ts":     datetime.now().strftime("%d %b %Y %H:%M"),
            "status": "Unverified",   # Unverified | Accepted
        }
        tags[hn].append(entry)
        _save_json(_TAGS_FILE, tags)
        _audit("TAG_ADD", f"user={username} host={hostname} tag={tag}")
    return tags[hn]

def remove_tag_by_id(hostname, tag_id, username):
    tags = get_tags()
    hn = hostname.strip().lower()
    if hn in tags:
        removed = next((t for t in tags[hn] if t["id"] == tag_id), None)
        tags[hn] = [t for t in tags[hn] if t["id"] != tag_id]
        if not tags[hn]: del tags[hn]
        _save_json(_TAGS_FILE, tags)
        _audit("TAG_REMOVE", f"user={username} host={hostname} "
               f"tag={removed['tag'] if removed else tag_id}")
    return tags.get(hn, [])

def accept_tag(hostname, tag_id, username):
    """Admin marks a tag as Accepted / official."""
    tags = get_tags()
    hn = hostname.strip().lower()
    if hn in tags:
        for t in tags[hn]:
            if t["id"] == tag_id:
                t["status"]      = "Accepted"
                t["accepted_by"] = username
                t["accepted_ts"] = datetime.now().strftime("%d %b %Y %H:%M")
        _save_json(_TAGS_FILE, tags)
        _audit("TAG_ACCEPT", f"admin={username} host={hostname} id={tag_id}")
    return tags.get(hn, [])

def add_note(hostname, note_text, username, full_name):
    notes = get_notes()
    hn = hostname.strip().lower()
    if hn not in notes: notes[hn] = []
    entry = {
        "id":       secrets.token_hex(6),
        "note":     note_text.strip(),
        "user":     username,
        "name":     full_name,
        "ts":       datetime.now().strftime("%d %b %Y %H:%M"),
    }
    notes[hn].append(entry)
    _save_json(_NOTES_FILE, notes)
    _audit("NOTE_ADD", f"user={username} host={hostname}")
    return notes[hn]

def delete_note(hostname, note_id, username, is_admin):
    notes = get_notes()
    hn = hostname.strip().lower()
    if hn not in notes: return []
    before = notes[hn]
    if is_admin:
        notes[hn] = [n for n in before if n["id"] != note_id]
    else:
        notes[hn] = [n for n in before
                     if not (n["id"] == note_id and n["user"] == username)]
    _save_json(_NOTES_FILE, notes)
    _audit("NOTE_DELETE", f"user={username} host={hostname} id={note_id}")
    return notes[hn]

def flag_server(hostname, reason, username):
    flags = get_flags()
    hn = hostname.strip().lower()
    flags[hn] = {"reason": reason, "user": username,
                 "ts": datetime.now().strftime("%d %b %Y %H:%M")}
    _save_json(_FLAGS_FILE, flags)
    _audit("FLAG_ADD", f"user={username} host={hostname} reason={reason}")

def unflag_server(hostname, username):
    flags = get_flags()
    hn = hostname.strip().lower()
    if hn in flags: del flags[hn]
    _save_json(_FLAGS_FILE, flags)
    _audit("FLAG_REMOVE", f"user={username} host={hostname}")

def to_excel_bytes(df):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        df.to_excel(w, index=False)
    return buf.getvalue()

# ═══════════════════════════════════════════════════════════════════════════════
#  FLASK APP
# ═══════════════════════════════════════════════════════════════════════════════
secret = _app_setup()
app = Flask(__name__)
app.secret_key = secret
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024
app.config["SESSION_COOKIE_HTTPONLY"] = True
app.json.sort_keys = False   # preserve column order (Server HostName first) in JSON responses

# Restore last uploaded inventory from disk (survives crash/restart) — runs once on import
load_persisted_inventory()

# ── Auth routes ───────────────────────────────────────────────────────────────
@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        uname = request.form.get("username", "").strip()
        pw    = request.form.get("password", "")
        users = _load_users()
        u     = users.get(uname)
        if u and u.get("enabled", True) and check_password_hash(u["password"], pw):
            session.clear()
            session["username"]    = uname
            session["last_active"] = datetime.now().isoformat()
            _audit("LOGIN_OK", f"user={uname}")
            nxt = request.args.get("next", "/")
            return redirect(nxt)
        _audit("LOGIN_FAIL", f"user={uname}")
        return render_template_string(LOGIN_HTML, error="Invalid username or password")
    return render_template_string(LOGIN_HTML, error=None)

@app.route("/logout")
def logout():
    u = session.get("username", "?")
    _audit("LOGOUT", f"user={u}")
    session.clear()
    return redirect("/")

# ── Admin: upload ─────────────────────────────────────────────────────────────
@app.route("/upload", methods=["POST"])
@_login_required
@_admin_required
def upload():
    which = request.form.get("which", "current")
    f     = request.files.get("file")
    if not f: return jsonify({"error": "No file"}), 400
    try:
        raw  = f.read()
        df   = load_excel(raw, f.filename)
        sc   = detect_status_col(df.columns.tolist())
        if which == "prev":
            STORE["df_prev"]    = df
            STORE["prev_status"] = sc
            STORE["prev_file"]   = f.filename
            persist_inventory(df, f.filename, sc, None, len(df), {}, which="prev")
            _audit("UPLOAD_PREV", f"user={session.get('username')} file={f.filename} rows={len(df)}")
            return jsonify({"ok": True, "filename": f.filename,
                            "rows": len(df), "status_col": sc or "Not detected"})
        qc = quality_check(df, sc)
        uploaded_at = datetime.now().strftime("%d %b %Y %H:%M")
        STORE.update({"df": df, "status_col": sc, "filename": f.filename,
                      "uploaded_at": uploaded_at,
                      "total_rows": len(df), "quality": qc})
        persist_inventory(df, f.filename, sc, uploaded_at, len(df), qc, which="current")
        _audit("UPLOAD", f"user={session.get('username')} file={f.filename} rows={len(df)}")
        return jsonify({"ok": True, "filename": f.filename, "rows": len(df),
                        "status_col": sc or "Not detected", "quality": qc})
    except Exception as e:
        log.exception("Upload error")
        return jsonify({"error": str(e)}), 500

# ── Public: dashboard ─────────────────────────────────────────────────────────
@app.route("/dashboard")
def dashboard():
    df = STORE["df"]
    if df is None: return jsonify({"error": "No data loaded"})
    sc  = STORE["status_col"]
    res = analyse(df, sc)
    charts = {}
    if sc:
        charts["status_donut"] = make_donut(
            ["Live", "Not Live", "Other"],
            [res["live"], res["not_live"], res["status_other"]],
            colors=[GREEN, RED, MUTED], title="Live vs Not Live")
    charts["type_donut"] = make_donut(
        ["Physical", "Virtual"], [res["physical"], res["virtual"]],
        colors=[BLUE, PURPLE], title="Physical vs Virtual")
    if res["platform_labels"]:
        charts["platform_hbar"] = make_hbar(
            res["platform_labels"], res["platform_values"],
            colors=res["platform_colors"], title="Servers by Platform")
    if res["loc_labels"]:
        charts["loc_hbar"] = make_hbar(
            res["loc_labels"], res["loc_values"],
            colors=[CYAN]*len(res["loc_labels"]), title="Servers by DC Location")
    if res["hpc_labels"]:
        charts["hpc_hbar"] = make_hbar(
            res["hpc_labels"], res["hpc_values"],
            colors=[YELLOW]*len(res["hpc_labels"]), title="HPC / NON-HPC / JPC")
    if res["role_labels"]:
        charts["role_vbar"] = make_vbar(
            res["role_labels"], res["role_values"],
            colors=[PURPLE]*len(res["role_labels"]), title="Servers by Role")
    if res["os_labels"]:
        charts["os_vbar"] = make_vbar(
            res["os_labels"], res["os_values"],
            colors=[ORANGE]*len(res["os_labels"]), title="OS Distribution (Top 15)")
    if res["app_labels"]:
        charts["app_hbar"] = make_hbar(
            res["app_labels"], res["app_values"],
            colors=[BLUE]*len(res["app_labels"]), title="Top 15 Applications")
    return jsonify({
        "metrics": {k: res[k] for k in
                    ["total","live","not_live","physical","virtual"]},
        "charts": charts, "status_col": sc or "",
        "filename": STORE["filename"] or "",
        "uploaded_at": STORE["uploaded_at"] or "",
        "quality": STORE["quality"],
    })

# ── Public: filter options ────────────────────────────────────────────────────
@app.route("/filter_options")
def filter_options():
    df = STORE["df"]
    if df is None: return jsonify({})
    sc = STORE["status_col"]
    def opts(col):
        if col not in df.columns: return []
        return sorted(df[col].str.strip().replace("", np.nan).dropna().unique().tolist())
    st_opts = []
    if sc and sc in df.columns:
        st_opts = sorted(df[sc].str.strip().replace("", np.nan)
                         .dropna().unique().tolist())
    return jsonify({
        "platform": opts("Platform"), "status": st_opts,
        "role": opts("Server Role"), "hpc": opts("HPC or NON HPC or JPC"),
        "loc": opts("Server DC Location"), "os": opts("Final OS"),
        "stype": opts("Server Type(Physical/Virtual)"),
    })

# ── Public: search ────────────────────────────────────────────────────────────
@app.route("/search")
def search():
    df = STORE["df"]
    if df is None: return jsonify({"rows": [], "total": 0})
    sc = STORE["status_col"]
    q        = request.args.get("q", "").strip().lower()
    platform = request.args.get("platform", "")
    status   = request.args.get("status", "")
    role     = request.args.get("role", "")
    hpc      = request.args.get("hpc", "")
    loc      = request.args.get("loc", "")
    os_f     = request.args.get("os", "")
    stype    = request.args.get("stype", "")
    sort_by  = request.args.get("sort_by", "")
    sort_dir = request.args.get("sort_dir", "asc")
    page     = max(1, int(request.args.get("page", 1)))
    per_page = 50

    filt = df.copy()
    if q:
        mask = pd.Series(False, index=filt.index)
        for col in ["Server HostName", "Application Name", "Reference"]:
            if col in filt.columns:
                mask |= filt[col].str.lower().str.contains(q, na=False)
        filt = filt[mask]
    if platform and "Platform" in filt.columns:
        filt = filt[filt["Platform"].str.strip() == platform]
    if status and sc and sc in filt.columns:
        filt = filt[filt[sc].str.strip().str.lower() == status.lower()]
    if role and "Server Role" in filt.columns:
        filt = filt[filt["Server Role"].str.strip() == role]
    if hpc and "HPC or NON HPC or JPC" in filt.columns:
        filt = filt[filt["HPC or NON HPC or JPC"].str.strip() == hpc]
    if loc and "Server DC Location" in filt.columns:
        filt = filt[filt["Server DC Location"].str.strip() == loc]
    if os_f and "Final OS" in filt.columns:
        filt = filt[filt["Final OS"].str.strip() == os_f]
    if stype and "Server Type(Physical/Virtual)" in filt.columns:
        filt = filt[filt["Server Type(Physical/Virtual)"].str.lower()
                    .str.contains(stype.lower(), na=False)]

    # Sorting — applied on the full filtered set, before pagination
    if sort_by and sort_by in filt.columns:
        ascending = (sort_dir != "desc")
        filt = filt.sort_values(
            by=sort_by,
            key=lambda s: s.str.strip().str.lower(),
            ascending=ascending, kind="mergesort"  # stable sort
        )

    total   = len(filt)
    page_df = filt.iloc[(page-1)*per_page : page*per_page]
    disp    = [c for c in ["Server HostName","Server Type(Physical/Virtual)",
               "Platform", sc if sc else None, "Server DC Location",
               "HPC or NON HPC or JPC","Server Role","Final OS",
               "Application Name","Commercial Category","Reference"]
               if c and c in page_df.columns]

    # enrich with tags/notes/flags for display
    all_tags  = get_tags()
    all_notes = get_notes()
    all_flags = get_flags()
    rows = []
    for _, r in page_df[disp].fillna("").iterrows():
        d = r.to_dict()
        hn = d.get("Server HostName", "").strip().lower()
        d["_tags"]    = all_tags.get(hn, [])
        d["_notes"]   = len(all_notes.get(hn, []))
        d["_flagged"] = hn in all_flags
        rows.append(d)

    return jsonify({"rows": rows, "total": total, "page": page,
                    "pages": max(1, -(-total // per_page)),
                    "status_col": sc or "",
                    "sort_by": sort_by, "sort_dir": sort_dir})

# ── Public: grouped summary ───────────────────────────────────────────────────
GROUPABLE_COLS = {
    "platform":  "Platform",
    "loc":       "Server DC Location",
    "hpc":       "HPC or NON HPC or JPC",
    "role":      "Server Role",
    "os":        "Final OS",
    "stype":     "Server Type(Physical/Virtual)",
    "app":       "Application Name",
}

@app.route("/grouped")
def grouped():
    """Group servers by a chosen dimension, returning counts + live/not-live
    split per group. Drill-down into a single group's rows is a separate call."""
    df = STORE["df"]
    if df is None: return jsonify({"groups": [], "total": 0})
    sc = STORE["status_col"]
    by_key = request.args.get("by", "platform")
    by_col = GROUPABLE_COLS.get(by_key)
    if not by_col or by_col not in df.columns:
        return jsonify({"error": f"Cannot group by '{by_key}'"}), 400

    work = df.copy()
    work[by_col] = work[by_col].str.strip().replace("", "(blank)")

    groups = []
    for grp_val, grp_df in work.groupby(by_col, sort=False):
        entry = {"value": grp_val, "count": len(grp_df)}
        if sc and sc in grp_df.columns:
            st = grp_df[sc].str.strip().str.lower()
            entry["live"]     = int((st == "live").sum())
            entry["not_live"] = int((st == "not live").sum())
        else:
            entry["live"] = entry["not_live"] = 0
        groups.append(entry)

    groups.sort(key=lambda g: g["count"], reverse=True)
    return jsonify({"groups": groups, "total": len(work), "by": by_key,
                    "by_label": by_col, "status_col": sc or ""})

@app.route("/grouped/drill")
def grouped_drill():
    """Return rows belonging to one group value, paginated, for drill-down."""
    df = STORE["df"]
    if df is None: return jsonify({"rows": [], "total": 0})
    sc = STORE["status_col"]
    by_key = request.args.get("by", "platform")
    value  = request.args.get("value", "")
    page   = max(1, int(request.args.get("page", 1)))
    per_page = 50
    by_col = GROUPABLE_COLS.get(by_key)
    if not by_col or by_col not in df.columns:
        return jsonify({"error": f"Cannot group by '{by_key}'"}), 400

    filt = df.copy()
    filt[by_col] = filt[by_col].str.strip().replace("", "(blank)")
    filt = filt[filt[by_col] == value]

    total   = len(filt)
    page_df = filt.iloc[(page-1)*per_page : page*per_page]
    disp    = [c for c in ["Server HostName","Server Type(Physical/Virtual)",
               "Platform", sc if sc else None, "Server DC Location",
               "HPC or NON HPC or JPC","Server Role","Final OS",
               "Application Name","Commercial Category","Reference"]
               if c and c in page_df.columns]

    all_tags  = get_tags()
    all_flags = get_flags()
    rows = []
    for _, r in page_df[disp].fillna("").iterrows():
        d = r.to_dict()
        hn = d.get("Server HostName", "").strip().lower()
        d["_tags"]    = all_tags.get(hn, [])
        d["_flagged"] = hn in all_flags
        rows.append(d)

    return jsonify({"rows": rows, "total": total, "page": page,
                    "pages": max(1, -(-total // per_page)),
                    "status_col": sc or "", "value": value})

# ── Public: server detail ─────────────────────────────────────────────────────
@app.route("/detail")
def detail():
    df  = STORE["df"]
    sc  = STORE["status_col"]
    hn  = request.args.get("hostname", "").strip()
    if df is None or not hn or "Server HostName" not in df.columns:
        return jsonify({})
    row = df[df["Server HostName"].str.strip().str.lower() == hn.lower()]
    if row.empty: return jsonify({})
    r = row.iloc[0].to_dict()
    if sc: r["_status_col"] = sc; r["_status_val"] = r.get(sc, "")
    hn_key = hn.lower()
    r["_tags"]             = get_tags().get(hn_key, [])
    r["_notes"]            = get_notes().get(hn_key, [])
    r["_flag"]             = get_flags().get(hn_key)
    r["_correctable_cols"] = CORRECTABLE_COLS   # ← always sent to client
    return jsonify(r)

# ── Public: bulk lookup ───────────────────────────────────────────────────────
@app.route("/bulk", methods=["POST"])
def bulk():
    df = STORE["df"]
    sc = STORE["status_col"]
    if df is None: return jsonify({"found": [], "not_found": [], "total_found": 0})
    data  = request.get_json(force=True)
    names = [n.strip() for n in data.get("names", []) if n.strip()]
    if not names or "Server HostName" not in df.columns:
        return jsonify({"found": [], "not_found": names, "total_found": 0})
    lower = [n.lower() for n in names]
    df["_hn_l"] = df["Server HostName"].str.strip().str.lower()
    found_df = df[df["_hn_l"].isin(lower)]
    found_set = set(found_df["_hn_l"].tolist())
    not_found = [n for n in names if n.lower() not in found_set]
    disp = [c for c in ["Server HostName","Server Type(Physical/Virtual)","Platform",
            sc if sc else None,"Server DC Location","HPC or NON HPC or JPC",
            "Server Role","Final OS","Application Name","Commercial Category"]
            if c and c in found_df.columns]
    all_tags  = get_tags()
    all_flags = get_flags()
    rows = []
    for _, r in found_df[disp].fillna("").iterrows():
        d = r.to_dict()
        hn_key = d.get("Server HostName", "").strip().lower()
        d["_tags"]    = all_tags.get(hn_key, [])
        d["_flagged"] = hn_key in all_flags
        rows.append(d)
    df.drop(columns=["_hn_l"], inplace=True)
    return jsonify({"found": rows, "not_found": not_found,
                    "total_found": len(rows), "status_col": sc or ""})

# ── Public: compare ───────────────────────────────────────────────────────────
@app.route("/compare")
@_login_required
@_admin_required
def compare():
    dc = STORE["df"]; dp = STORE["df_prev"]
    if dc is None or dp is None:
        return jsonify({"error": "Upload both months first"})
    try:
        out = compare_months(dc, dp, STORE["status_col"], STORE["prev_status"])
        charts = {}
        if out.get("plat_labels"):
            charts["plat_compare"] = make_hbar_compare(
                out["plat_labels"], out["plat_curr"], out["plat_prev"],
                label_a=STORE["filename"] or "Current",
                label_b=STORE["prev_file"] or "Previous",
                title="Total Servers by Platform")
        if out.get("plat_labels") and out.get("live_curr"):
            charts["live_compare"] = make_hbar_compare(
                out["plat_labels"], out["live_curr"], out["live_prev"],
                label_a=STORE["filename"] or "Current",
                label_b=STORE["prev_file"] or "Previous",
                title="Live Servers by Platform")
        out["charts"] = charts
        out["curr_file"] = STORE["filename"] or "Current"
        out["prev_file"] = STORE["prev_file"] or "Previous"
        return jsonify(out)
    except Exception as e:
        log.exception("Compare error")
        return jsonify({"error": str(e)})

# ── User: tag routes ──────────────────────────────────────────────────────────
@app.route("/tag/add", methods=["POST"])
@_login_required
def tag_add():
    u    = _current_user()
    data = request.get_json(force=True)
    hn   = data.get("hostname", "").strip()
    tag  = data.get("tag", "").strip()
    if not hn or not tag: return jsonify({"error": "Missing fields"}), 400
    tags = add_tag(hn, tag, u["username"], u.get("full_name", u["username"]))
    return jsonify({"ok": True, "tags": tags})

@app.route("/tag/remove", methods=["POST"])
@_login_required
def tag_remove():
    u    = _current_user()
    data = request.get_json(force=True)
    hn      = data.get("hostname", "").strip()
    tag_id  = data.get("id", "").strip()
    tags = remove_tag_by_id(hn, tag_id, u["username"])
    return jsonify({"ok": True, "tags": tags})

@app.route("/admin/tag/accept", methods=["POST"])
@_login_required
@_admin_required
def admin_tag_accept():
    data   = request.get_json(force=True)
    hn     = data.get("hostname", "").strip()
    tag_id = data.get("id", "").strip()
    tags = accept_tag(hn, tag_id, session.get("username", ""))
    return jsonify({"ok": True, "tags": tags})

# ── User: note routes ─────────────────────────────────────────────────────────
@app.route("/note/add", methods=["POST"])
@_login_required
def note_add():
    u    = _current_user()
    data = request.get_json(force=True)
    hn   = data.get("hostname", "").strip()
    txt  = data.get("note", "").strip()
    if not hn or not txt: return jsonify({"error": "Missing fields"}), 400
    notes = add_note(hn, txt, u["username"], u.get("full_name", u["username"]))
    return jsonify({"ok": True, "notes": notes})

@app.route("/note/delete", methods=["POST"])
@_login_required
def note_delete():
    u     = _current_user()
    data  = request.get_json(force=True)
    hn    = data.get("hostname", "").strip()
    nid   = data.get("id", "").strip()
    is_admin = u.get("role") == "admin"
    notes = delete_note(hn, nid, u["username"], is_admin)
    return jsonify({"ok": True, "notes": notes})

# ── User: flag routes ─────────────────────────────────────────────────────────
@app.route("/flag/add", methods=["POST"])
@_login_required
def flag_add():
    u    = _current_user()
    data = request.get_json(force=True)
    hn   = data.get("hostname", "").strip()
    rsn  = data.get("reason", "").strip()
    if not hn: return jsonify({"error": "Missing hostname"}), 400
    flag_server(hn, rsn, u["username"])
    return jsonify({"ok": True})

@app.route("/flag/remove", methods=["POST"])
@_login_required
def flag_remove():
    u    = _current_user()
    data = request.get_json(force=True)
    hn   = data.get("hostname", "").strip()
    unflag_server(hn, u["username"])
    return jsonify({"ok": True})

# ── User: correction routes ──────────────────────────────────────────────────
@app.route("/correction/add", methods=["POST"])
@_login_required
def correction_add():
    u    = _current_user()
    data = request.get_json(force=True)
    hn   = data.get("hostname","").strip()
    col  = data.get("column","").strip()
    cur  = data.get("current_val","").strip()
    sug  = data.get("suggested_val","").strip()
    rsn  = data.get("reason","").strip()
    if not hn or not col or not sug:
        return jsonify({"error": "Hostname, column and suggested value required"}), 400
    if col not in CORRECTABLE_COLS:
        return jsonify({"error": "Column not correctable"}), 400
    entry = add_correction(hn, col, cur, sug, rsn,
                           u["username"], u.get("full_name", u["username"]))
    return jsonify({"ok": True, "id": entry["id"]})

@app.route("/admin/correction/review", methods=["POST"])
@_login_required
@_admin_required
def admin_correction_review():
    data     = request.get_json(force=True)
    cid      = data.get("id","").strip()
    decision = data.get("decision","Approved").strip()
    reason   = data.get("reason","")
    if decision not in ("Approved","Rejected"):
        return jsonify({"error":"Invalid decision"}), 400
    mark_correction_decision(cid, decision, reason, session.get("username",""))
    return jsonify({"ok": True})

@app.route("/admin/correction/delete", methods=["POST"])
@_login_required
@_admin_required
def admin_correction_delete():
    data = request.get_json(force=True)
    cid  = data.get("id","").strip()
    delete_correction(cid, session.get("username",""))
    return jsonify({"ok": True})

# ── User: decommission routes ────────────────────────────────────────────────
@app.route("/decommission/add", methods=["POST"])
@_login_required
def decommission_add():
    u    = _current_user()
    data = request.get_json(force=True)
    hn   = data.get("hostname","").strip()
    rsn  = data.get("reason","").strip()
    dt   = data.get("eff_date","").strip()
    cmt  = data.get("comment","").strip()
    if not hn or rsn not in DECOM_REASONS:
        return jsonify({"error": "Hostname and a valid reason are required"}), 400
    entry = add_decommission(hn, rsn, dt, cmt, u["username"],
                             u.get("full_name", u["username"]))
    return jsonify({"ok": True, "id": entry["id"]})

@app.route("/admin/decommission/review", methods=["POST"])
@_login_required
@_admin_required
def admin_decom_review():
    data     = request.get_json(force=True)
    decision = data.get("decision","Approved").strip()
    reason   = data.get("reason","")
    if decision not in ("Approved","Rejected"):
        return jsonify({"error":"Invalid decision"}), 400
    mark_decom_decision(data.get("id","").strip(), decision, reason, session.get("username",""))
    return jsonify({"ok": True})

@app.route("/admin/decommission/delete", methods=["POST"])
@_login_required
@_admin_required
def admin_decom_delete():
    data = request.get_json(force=True)
    delete_decommission(data.get("id","").strip(), session.get("username",""))
    return jsonify({"ok": True})

@app.route("/admin/decommissions_report")
@_login_required
@_admin_required
def admin_decom_report():
    decs = get_decommissions()
    if not decs: return "No decommission entries yet", 200
    rows = [{"ID": v["id"], "Hostname": v["hostname"], "Reason": v["reason"],
             "Effective Date": v["eff_date"], "Comment": v["comment"],
             "Submitted By": v["name"], "Username": v["user"],
             "Submitted At": v["ts"], "Status": v["status"],
             "Decision Reason": v.get("decision_reason",""),
             "Reviewed By": v.get("reviewed_by",""),
             "Reviewed At": v.get("reviewed_ts","")} for v in decs.values()]
    buf = io.BytesIO()
    pd.DataFrame(rows).to_excel(buf, index=False)
    buf.seek(0)
    return send_file(buf, as_attachment=True,
                     download_name=f"decommissions_report_{datetime.now().strftime('%Y%m%d')}.xlsx",
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ── User: new server suggestion routes ───────────────────────────────────────
@app.route("/newserver/add", methods=["POST"])
@_login_required
def newserver_add():
    u    = _current_user()
    data = request.get_json(force=True)
    fields  = data.get("fields", {})
    comment = data.get("comment", "")
    if not fields.get("Server HostName","").strip():
        return jsonify({"error": "Hostname is required"}), 400
    entry = add_new_server(fields, comment, u["username"],
                           u.get("full_name", u["username"]))
    return jsonify({"ok": True, "id": entry["id"]})

@app.route("/admin/newserver/review", methods=["POST"])
@_login_required
@_admin_required
def admin_newsrv_review():
    data     = request.get_json(force=True)
    decision = data.get("decision","Approved").strip()
    reason   = data.get("reason","")
    if decision not in ("Approved","Rejected"):
        return jsonify({"error":"Invalid decision"}), 400
    mark_newsrv_decision(data.get("id","").strip(), decision, reason, session.get("username",""))
    return jsonify({"ok": True})

@app.route("/admin/newserver/delete", methods=["POST"])
@_login_required
@_admin_required
def admin_newsrv_delete():
    data = request.get_json(force=True)
    delete_new_server(data.get("id","").strip(), session.get("username",""))
    return jsonify({"ok": True})

@app.route("/admin/newservers_report")
@_login_required
@_admin_required
def admin_newsrv_report():
    news = get_new_servers()
    if not news: return "No new server suggestions yet", 200
    rows = []
    for v in news.values():
        row = {"ID": v["id"], **v["fields"], "Comment": v["comment"],
               "Submitted By": v["name"], "Username": v["user"],
               "Submitted At": v["ts"], "Status": v["status"],
               "Decision Reason": v.get("decision_reason",""),
               "Reviewed By": v.get("reviewed_by",""),
               "Reviewed At": v.get("reviewed_ts","")}
        rows.append(row)
    buf = io.BytesIO()
    pd.DataFrame(rows).to_excel(buf, index=False)
    buf.seek(0)
    return send_file(buf, as_attachment=True,
                     download_name=f"new_servers_report_{datetime.now().strftime('%Y%m%d')}.xlsx",
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ── User: T-shirt size (CPU/RAM) change routes ────────────────────────────────
@app.route("/tshirt/add", methods=["POST"])
@_login_required
def tshirt_add():
    u    = _current_user()
    data = request.get_json(force=True)
    hn   = data.get("hostname","").strip()
    pl   = data.get("platform","").strip()
    cc   = data.get("current_cpu","").strip()
    cr   = data.get("current_ram","").strip()
    nc   = data.get("new_cpu","").strip()
    nr   = data.get("new_ram","").strip()
    rsn  = data.get("reason","").strip()
    if not hn:
        return jsonify({"error": "Hostname is required"}), 400
    if pl not in TSHIRT_PLATFORMS:
        return jsonify({"error": f"T-shirt size changes only apply to {', '.join(TSHIRT_PLATFORMS)} servers"}), 400
    if not nc and not nr:
        return jsonify({"error": "Provide a new CPU or new RAM value"}), 400
    entry = add_tshirt_change(hn, pl, cc, cr, nc, nr, rsn,
                              u["username"], u.get("full_name", u["username"]))
    return jsonify({"ok": True, "id": entry["id"]})

@app.route("/admin/tshirt/review", methods=["POST"])
@_login_required
@_admin_required
def admin_tshirt_review():
    data     = request.get_json(force=True)
    decision = data.get("decision","Approved").strip()
    reason   = data.get("reason","")
    if decision not in ("Approved","Rejected"):
        return jsonify({"error":"Invalid decision"}), 400
    mark_tshirt_decision(data.get("id","").strip(), decision, reason, session.get("username",""))
    return jsonify({"ok": True})

@app.route("/admin/tshirt/delete", methods=["POST"])
@_login_required
@_admin_required
def admin_tshirt_delete():
    data = request.get_json(force=True)
    delete_tshirt_change(data.get("id","").strip(), session.get("username",""))
    return jsonify({"ok": True})

@app.route("/admin/tshirt_report")
@_login_required
@_admin_required
def admin_tshirt_report():
    changes = get_tshirt_changes()
    if not changes: return "No T-shirt size changes yet", 200
    rows = [{
        "ID": v["id"], "Hostname": v["hostname"], "Platform": v["platform"],
        "Current CPU": v["current_cpu"], "Current RAM": v["current_ram"],
        "New CPU": v["new_cpu"], "New RAM": v["new_ram"],
        "Reason": v["reason"], "Submitted By": v["name"], "Username": v["user"],
        "Submitted At": v["ts"], "Status": v["status"],
        "Decision Reason": v.get("decision_reason",""),
        "Reviewed By": v.get("reviewed_by",""),
        "Reviewed At": v.get("reviewed_ts",""),
    } for v in changes.values()]
    buf = io.BytesIO()
    pd.DataFrame(rows).to_excel(buf, index=False)
    buf.seek(0)
    return send_file(buf, as_attachment=True,
                     download_name=f"tshirt_size_changes_{datetime.now().strftime('%Y%m%d')}.xlsx",
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

@app.route("/admin/monthly_changes_report")
@_login_required
@_admin_required
def admin_monthly_changes_report():
    """Consolidated, ready-to-apply record of every APPROVED change this
    month — corrections, decommissions, new servers, and T-shirt size
    (CPU/RAM) changes — across four sheets in one Excel file. This is what
    admin works from when updating next month's master inventory and
    when reconciling EMA/EPMC/JEA contract billing."""
    corrs   = get_corrections()
    decs    = get_decommissions()
    news    = get_new_servers()
    tshirts = get_tshirt_changes()

    corr_rows = [{
        "Hostname": v["hostname"], "Column": v["column"],
        "Current Value": v["current_val"], "New Value": v["suggested_val"],
        "Reason Given": v["reason"], "Submitted By": v["name"],
        "Submitted At": v["ts"], "Approved By": v.get("reviewed_by",""),
        "Approved At": v.get("reviewed_ts",""),
        "Admin Note": v.get("decision_reason",""),
    } for v in corrs.values() if v.get("status") == "Approved"]

    decom_rows = [{
        "Hostname": v["hostname"], "Reason": v["reason"],
        "Effective Date": v["eff_date"], "Comment": v["comment"],
        "Submitted By": v["name"], "Submitted At": v["ts"],
        "Approved By": v.get("reviewed_by",""), "Approved At": v.get("reviewed_ts",""),
        "Admin Note": v.get("decision_reason",""),
    } for v in decs.values() if v.get("status") == "Approved"]

    newsrv_rows = [{
        **v["fields"], "Comment": v["comment"],
        "Submitted By": v["name"], "Submitted At": v["ts"],
        "Approved By": v.get("reviewed_by",""), "Approved At": v.get("reviewed_ts",""),
        "Admin Note": v.get("decision_reason",""),
    } for v in news.values() if v.get("status") == "Approved"]

    tshirt_rows = [{
        "Hostname": v["hostname"], "Platform": v["platform"],
        "Current CPU": v["current_cpu"], "Current RAM": v["current_ram"],
        "New CPU": v["new_cpu"], "New RAM": v["new_ram"],
        "Reason Given": v["reason"], "Submitted By": v["name"],
        "Submitted At": v["ts"], "Approved By": v.get("reviewed_by",""),
        "Approved At": v.get("reviewed_ts",""),
        "Admin Note": v.get("decision_reason",""),
    } for v in tshirts.values() if v.get("status") == "Approved"]

    if not corr_rows and not decom_rows and not newsrv_rows and not tshirt_rows:
        return "No approved changes to report yet", 200

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        (pd.DataFrame(corr_rows) if corr_rows else pd.DataFrame(columns=["Hostname"])) \
            .to_excel(writer, sheet_name="Corrections", index=False)
        (pd.DataFrame(decom_rows) if decom_rows else pd.DataFrame(columns=["Hostname"])) \
            .to_excel(writer, sheet_name="Decommissions", index=False)
        (pd.DataFrame(newsrv_rows) if newsrv_rows else pd.DataFrame(columns=["Server HostName"])) \
            .to_excel(writer, sheet_name="New Servers", index=False)
        (pd.DataFrame(tshirt_rows) if tshirt_rows else pd.DataFrame(columns=["Hostname"])) \
            .to_excel(writer, sheet_name="T-Shirt Size Changes", index=False)
    buf.seek(0)
    fname = f"monthly_changes_to_apply_{datetime.now().strftime('%Y%m%d')}.xlsx"
    return send_file(buf, as_attachment=True, download_name=fname,
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ── Admin: user management ────────────────────────────────────────────────────
@app.route("/admin/users")
@_login_required
@_admin_required
def admin_users():
    users = _load_users()
    out   = []
    for uname, u in users.items():
        out.append({"username": uname, "full_name": u.get("full_name",""),
                    "role": u.get("role","user"),
                    "enabled": u.get("enabled", True),
                    "created_at": u.get("created_at","")})
    return jsonify({"users": out})

@app.route("/admin/user/create", methods=["POST"])
@_login_required
@_admin_required
def admin_user_create():
    data  = request.get_json(force=True)
    uname = data.get("username","").strip()
    pw    = data.get("password","").strip()
    role  = data.get("role","user")
    fname = data.get("full_name","").strip()
    if not uname or not pw:
        return jsonify({"error": "Username and password required"}), 400
    if len(pw) < 6:
        return jsonify({"error": "Password must be at least 6 characters"}), 400
    users = _load_users()
    if uname in users:
        return jsonify({"error": f"User '{uname}' already exists"}), 400
    users[uname] = {
        "password":   generate_password_hash(pw),
        "role":       role,
        "enabled":    True,
        "full_name":  fname or uname,
        "created_at": datetime.now().isoformat(),
    }
    _save_users(users)
    _audit("USER_CREATE", f"admin={session.get('username')} new_user={uname} role={role}")
    return jsonify({"ok": True})

@app.route("/admin/user/toggle", methods=["POST"])
@_login_required
@_admin_required
def admin_user_toggle():
    data  = request.get_json(force=True)
    uname = data.get("username","").strip()
    if uname == session.get("username"):
        return jsonify({"error": "Cannot disable your own account"}), 400
    users = _load_users()
    if uname not in users: return jsonify({"error": "User not found"}), 404
    users[uname]["enabled"] = not users[uname].get("enabled", True)
    _save_users(users)
    _audit("USER_TOGGLE", f"admin={session.get('username')} user={uname} enabled={users[uname]['enabled']}")
    return jsonify({"ok": True, "enabled": users[uname]["enabled"]})

@app.route("/admin/user/reset_pw", methods=["POST"])
@_login_required
@_admin_required
def admin_reset_pw():
    data  = request.get_json(force=True)
    uname = data.get("username","").strip()
    pw    = data.get("password","").strip()
    if len(pw) < 6: return jsonify({"error": "Min 6 chars"}), 400
    users = _load_users()
    if uname not in users: return jsonify({"error": "User not found"}), 404
    users[uname]["password"] = generate_password_hash(pw)
    _save_users(users)
    _audit("PW_RESET", f"admin={session.get('username')} target={uname}")
    return jsonify({"ok": True})

@app.route("/admin/user/change_role", methods=["POST"])
@_login_required
@_admin_required
def admin_change_role():
    data  = request.get_json(force=True)
    uname = data.get("username","").strip()
    role  = data.get("role","user")
    if uname == session.get("username"):
        return jsonify({"error": "Cannot change your own role"}), 400
    users = _load_users()
    if uname not in users: return jsonify({"error": "User not found"}), 404
    users[uname]["role"] = role
    _save_users(users)
    _audit("ROLE_CHANGE", f"admin={session.get('username')} user={uname} role={role}")
    return jsonify({"ok": True})

# ── Admin: reports ────────────────────────────────────────────────────────────
@app.route("/admin/tags_report")
@_login_required
@_admin_required
def admin_tags_report():
    df = STORE["df"]
    tags = get_tags()
    rows = []
    for hn_low, tag_list in tags.items():
        srv_row = {}
        if df is not None and "Server HostName" in df.columns:
            match = df[df["Server HostName"].str.strip().str.lower() == hn_low]
            if not match.empty:
                srv_row = match.iloc[0].to_dict()
        for t in tag_list:
            is_dict = isinstance(t, dict)
            rows.append({
                "Hostname":    srv_row.get("Server HostName", hn_low),
                "Tag":         t.get("tag", t) if is_dict else t,
                "Status":      t.get("status","Unverified") if is_dict else "Unverified",
                "Submitted By": t.get("name","") if is_dict else "",
                "Submitted At": t.get("ts","") if is_dict else "",
                "Accepted By": t.get("accepted_by","") if is_dict else "",
                "Platform":    srv_row.get("Platform",""),
                "Server Role": srv_row.get("Server Role",""),
                "Final OS":    srv_row.get("Final OS",""),
                "DC Location": srv_row.get("Server DC Location",""),
            })
    if not rows: return "No tags yet", 200
    buf   = io.BytesIO()
    pd.DataFrame(rows).to_excel(buf, index=False)
    buf.seek(0)
    return send_file(buf, as_attachment=True,
                     download_name=f"tags_report_{datetime.now().strftime('%Y%m%d')}.xlsx",
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.route("/admin/notes_report")
@_login_required
@_admin_required
def admin_notes_report():
    df = STORE["df"]
    notes = get_notes()
    rows = []
    for hn_low, note_list in notes.items():
        srv_row = {}
        if df is not None and "Server HostName" in df.columns:
            match = df[df["Server HostName"].str.strip().str.lower() == hn_low]
            if not match.empty: srv_row = match.iloc[0].to_dict()
        for n in note_list:
            rows.append({
                "Hostname":    srv_row.get("Server HostName", hn_low),
                "Note":        n["note"],
                "By":          n["name"],
                "Username":    n["user"],
                "Timestamp":   n["ts"],
                "Platform":    srv_row.get("Platform",""),
                "Server Role": srv_row.get("Server Role",""),
            })
    if not rows: return "No notes yet", 200
    buf = io.BytesIO()
    pd.DataFrame(rows).to_excel(buf, index=False)
    buf.seek(0)
    return send_file(buf, as_attachment=True,
                     download_name=f"notes_report_{datetime.now().strftime('%Y%m%d')}.xlsx",
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

@app.route("/admin/flags_report")
@_login_required
@_admin_required
def admin_flags_report():
    df    = STORE["df"]
    flags = get_flags()
    rows  = []
    for hn_low, info in flags.items():
        srv_row = {}
        if df is not None and "Server HostName" in df.columns:
            match = df[df["Server HostName"].str.strip().str.lower() == hn_low]
            if not match.empty: srv_row = match.iloc[0].to_dict()
        rows.append({
            "Hostname":  srv_row.get("Server HostName", hn_low),
            "Reason":    info.get("reason",""),
            "Flagged By": info.get("user",""),
            "Timestamp": info.get("ts",""),
            "Platform":  srv_row.get("Platform",""),
            "Final OS":  srv_row.get("Final OS",""),
        })
    if not rows: return "No flags yet", 200
    buf = io.BytesIO()
    pd.DataFrame(rows).to_excel(buf, index=False)
    buf.seek(0)
    return send_file(buf, as_attachment=True,
                     download_name=f"flags_report_{datetime.now().strftime('%Y%m%d')}.xlsx",
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

@app.route("/admin/corrections_report")
@_login_required
@_admin_required
def admin_corrections_report():
    corrs = get_corrections()
    if not corrs: return "No corrections yet", 200
    rows = [{"ID": v["id"], "Hostname": v["hostname"],
             "Column": v["column"], "Current Value": v["current_val"],
             "Suggested Value": v["suggested_val"], "Reason": v["reason"],
             "Submitted By": v["name"], "Username": v["user"],
             "Submitted At": v["ts"], "Status": v["status"],
             "Decision Reason": v.get("decision_reason",""),
             "Reviewed By": v.get("reviewed_by",""),
             "Reviewed At": v.get("reviewed_ts","")} for v in corrs.values()]
    buf = io.BytesIO()
    pd.DataFrame(rows).to_excel(buf, index=False)
    buf.seek(0)
    return send_file(buf, as_attachment=True,
                     download_name=f"corrections_report_{datetime.now().strftime('%Y%m%d')}.xlsx",
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

@app.route("/admin/data")
@_login_required
@_admin_required
def admin_data():
    """Return full data for admin panel — inline tables + counts."""
    df    = STORE["df"]
    tags  = get_tags()
    notes = get_notes()
    flags = get_flags()
    corrs = get_corrections()

    # ── Build inline tags table ──────────────────────────────────────────────
    tag_rows = []
    for hn_low, tag_list in tags.items():
        srv = {}
        if df is not None and "Server HostName" in df.columns:
            m = df[df["Server HostName"].str.strip().str.lower() == hn_low]
            if not m.empty: srv = m.iloc[0].to_dict()
        for t in tag_list:
            tag_rows.append({
                "id":          t.get("id",""),
                "hn_key":      hn_low,
                "hostname":    srv.get("Server HostName", hn_low),
                "tag":         t.get("tag", t) if isinstance(t, dict) else t,
                "user":        t.get("user","") if isinstance(t, dict) else "",
                "name":        t.get("name","") if isinstance(t, dict) else "",
                "ts":          t.get("ts","") if isinstance(t, dict) else "",
                "status":      t.get("status","Unverified") if isinstance(t, dict) else "Unverified",
                "accepted_by": t.get("accepted_by","") if isinstance(t, dict) else "",
                "platform":    srv.get("Platform","—"),
                "role":        srv.get("Server Role","—"),
                "os":          srv.get("Final OS","—"),
                "dc":          srv.get("Server DC Location","—"),
            })

    # ── Build inline notes table ─────────────────────────────────────────────
    note_rows = []
    for hn_low, note_list in notes.items():
        srv = {}
        if df is not None and "Server HostName" in df.columns:
            m = df[df["Server HostName"].str.strip().str.lower() == hn_low]
            if not m.empty: srv = m.iloc[0].to_dict()
        for n in note_list:
            note_rows.append({
                "id":       n["id"],
                "hostname": srv.get("Server HostName", hn_low),
                "note":     n["note"],
                "by":       n["name"],
                "user":     n["user"],
                "ts":       n["ts"],
                "platform": srv.get("Platform","—"),
                "role":     srv.get("Server Role","—"),
            })

    # ── Build inline flags list ──────────────────────────────────────────────
    flag_rows = []
    for hn_low, info in flags.items():
        srv = {}
        if df is not None and "Server HostName" in df.columns:
            m = df[df["Server HostName"].str.strip().str.lower() == hn_low]
            if not m.empty: srv = m.iloc[0].to_dict()
        flag_rows.append({
            "hostname": srv.get("Server HostName", hn_low),
            "reason":   info.get("reason",""),
            "user":     info.get("user",""),
            "ts":       info.get("ts",""),
            "platform": srv.get("Platform","—"),
            "os":       srv.get("Final OS","—"),
        })

    # ── Build corrections list ───────────────────────────────────────────────
    pending_corrs   = [v for v in corrs.values() if v.get("status") == "Pending"]
    approved_corrs  = [v for v in corrs.values() if v.get("status") == "Approved"]
    rejected_corrs  = [v for v in corrs.values() if v.get("status") == "Rejected"]

    # ── Decommissions + new servers (enrich with current inventory data) ────
    decs = get_decommissions()
    for v in decs.values():
        hn_low = v["hostname"].strip().lower()
        v["_in_inventory"] = bool(
            df is not None and "Server HostName" in df.columns and
            not df[df["Server HostName"].str.strip().str.lower() == hn_low].empty
        )
    pending_decs  = [v for v in decs.values() if v.get("status") == "Pending"]
    approved_decs = [v for v in decs.values() if v.get("status") == "Approved"]
    rejected_decs = [v for v in decs.values() if v.get("status") == "Rejected"]

    news = get_new_servers()
    pending_news  = [v for v in news.values() if v.get("status") == "Pending"]
    approved_news = [v for v in news.values() if v.get("status") == "Approved"]
    rejected_news = [v for v in news.values() if v.get("status") == "Rejected"]

    tshirts = get_tshirt_changes()
    pending_tshirts  = [v for v in tshirts.values() if v.get("status") == "Pending"]
    approved_tshirts = [v for v in tshirts.values() if v.get("status") == "Approved"]
    rejected_tshirts = [v for v in tshirts.values() if v.get("status") == "Rejected"]

    return jsonify({
        "tag_count":        sum(len(v) for v in tags.values()),
        "tag_unverified":   len([t for r in tag_rows for t in [r] if t["status"]=="Unverified"]),
        "tag_accepted":     len([t for r in tag_rows for t in [r] if t["status"]=="Accepted"]),
        "note_count":       sum(len(v) for v in notes.values()),
        "flag_count":       len(flags),
        "corr_pending":     len(pending_corrs),
        "corr_approved":    len(approved_corrs),
        "corr_rejected":    len(rejected_corrs),
        "decom_pending":    len(pending_decs),
        "decom_approved":   len(approved_decs),
        "decom_rejected":   len(rejected_decs),
        "newsrv_pending":   len(pending_news),
        "newsrv_approved":  len(approved_news),
        "newsrv_rejected":  len(rejected_news),
        "tshirt_pending":   len(pending_tshirts),
        "tshirt_approved":  len(approved_tshirts),
        "tshirt_rejected":  len(rejected_tshirts),
        "server_count":     STORE["total_rows"],
        "filename":         STORE["filename"] or "None",
        "uploaded_at":      STORE["uploaded_at"] or "—",
        "quality":          STORE["quality"],
        "tag_rows":         tag_rows,
        "note_rows":        note_rows,
        "flag_rows":        flag_rows,
        "corrections":      list(corrs.values()),
        "decommissions":    list(decs.values()),
        "new_servers":      list(news.values()),
        "tshirt_changes":   list(tshirts.values()),
        "preset_tags":      PRESET_TAGS,
        "correctable_cols": CORRECTABLE_COLS,
        "decom_reasons":    DECOM_REASONS,
        "newsrv_fields":    NEW_SERVER_FIELDS,
        "tshirt_platforms": TSHIRT_PLATFORMS,
    })

# ── Export ────────────────────────────────────────────────────────────────────
@app.route("/export")
def export():
    df = STORE["df"]
    if df is None: return "No data", 400
    sc = STORE["status_col"]
    filt = df.copy()
    q  = request.args.get("q","").strip().lower()
    pl = request.args.get("platform","")
    st = request.args.get("status","")
    ro = request.args.get("role","")
    if q:
        mask = pd.Series(False, index=filt.index)
        for col in ["Server HostName","Application Name"]:
            if col in filt.columns:
                mask |= filt[col].str.lower().str.contains(q, na=False)
        filt = filt[mask]
    if pl and "Platform" in filt.columns:
        filt = filt[filt["Platform"].str.strip() == pl]
    if st and sc and sc in filt.columns:
        filt = filt[filt[sc].str.strip().str.lower() == st.lower()]
    if ro and "Server Role" in filt.columns:
        filt = filt[filt["Server Role"].str.strip() == ro]
    buf = io.BytesIO()
    filt.drop(columns=[c for c in filt.columns if c.startswith("_")],
              errors="ignore").to_excel(buf, index=False)
    buf.seek(0)
    fname = f"inventory_export_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
    return send_file(buf, as_attachment=True, download_name=fname,
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

@app.route("/export_group")
def export_group():
    df = STORE["df"]
    if df is None: return "No data", 400
    by_key = request.args.get("by", "platform")
    value  = request.args.get("value", "")
    by_col = GROUPABLE_COLS.get(by_key)
    if not by_col or by_col not in df.columns:
        return "Invalid group", 400
    filt = df.copy()
    filt[by_col] = filt[by_col].str.strip().replace("", "(blank)")
    filt = filt[filt[by_col] == value]
    buf = io.BytesIO()
    filt.drop(columns=[c for c in filt.columns if c.startswith("_")],
              errors="ignore").to_excel(buf, index=False)
    buf.seek(0)
    safe_val = re.sub(r"[^A-Za-z0-9_-]+", "_", value)[:40]
    fname = f"group_{by_key}_{safe_val}_{datetime.now().strftime('%Y%m%d')}.xlsx"
    return send_file(buf, as_attachment=True, download_name=fname,
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

@app.route("/export_bulk", methods=["POST"])
def export_bulk():
    rows = request.get_json(force=True).get("rows", [])
    if not rows: return "No data", 400
    buf = io.BytesIO()
    pd.DataFrame(rows).drop(
        columns=[c for c in pd.DataFrame(rows).columns if c.startswith("_")],
        errors="ignore").to_excel(buf, index=False)
    buf.seek(0)
    return send_file(buf, as_attachment=True,
                     download_name=f"bulk_lookup_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

@app.route("/audit")
@_login_required
@_admin_required
def view_audit():
    if not _AUDIT_FILE.exists(): return "No audit log yet.", 200
    lines = _AUDIT_FILE.read_text().splitlines()
    body  = "\n".join(lines[-300:])
    return (f"<pre style='background:#07090f;color:#edf2ff;padding:20px;"
            f"font-family:monospace;font-size:.8rem;min-height:100vh'>"
            f"Asset Inventory Audit Log — last {min(len(lines),300)} entries\n"
            f"{'─'*60}\n{body}</pre>")

# ── Session info (for JS) ─────────────────────────────────────────────────────
@app.route("/me")
def me():
    u = _current_user()
    if not u: return jsonify({"logged_in": False})
    corrs   = get_corrections()
    decs    = get_decommissions()
    news    = get_new_servers()
    tshirts = get_tshirt_changes()
    return jsonify({"logged_in": True, "username": u["username"],
                    "full_name": u.get("full_name", u["username"]),
                    "role": u.get("role","user"),
                    "pending_corrections": sum(1 for v in corrs.values() if v.get("status")=="Pending"),
                    "pending_decommissions": sum(1 for v in decs.values() if v.get("status")=="Pending"),
                    "pending_newservers": sum(1 for v in news.values() if v.get("status")=="Pending"),
                    "pending_tshirts": sum(1 for v in tshirts.values() if v.get("status")=="Pending"),
                    "decom_reasons": DECOM_REASONS,
                    "newsrv_fields": NEW_SERVER_FIELDS,
                    "tshirt_platforms": TSHIRT_PLATFORMS,
                    "preset_tags": PRESET_TAGS})

@app.route("/my_activity")
@_login_required
def my_activity():
    """Activity summary for the logged-in user — accessible to any user,
    not just admin (unlike /admin/data)."""
    u = _current_user()
    uname = u["username"]
    df = STORE["df"]

    tags  = get_tags()
    notes = get_notes()
    corrs = get_corrections()

    my_tag_rows = []
    for hn_low, tag_list in tags.items():
        srv = {}
        if df is not None and "Server HostName" in df.columns:
            m = df[df["Server HostName"].str.strip().str.lower() == hn_low]
            if not m.empty: srv = m.iloc[0].to_dict()
        for t in tag_list:
            if isinstance(t, dict) and t.get("user") == uname:
                my_tag_rows.append({
                    "hostname": srv.get("Server HostName", hn_low),
                    "tag": t.get("tag",""), "status": t.get("status","Unverified"),
                    "ts": t.get("ts",""),
                })

    my_note_rows = []
    for hn_low, note_list in notes.items():
        srv = {}
        if df is not None and "Server HostName" in df.columns:
            m = df[df["Server HostName"].str.strip().str.lower() == hn_low]
            if not m.empty: srv = m.iloc[0].to_dict()
        for n in note_list:
            if n.get("user") == uname:
                my_note_rows.append({
                    "hostname": srv.get("Server HostName", hn_low),
                    "note": n.get("note",""), "ts": n.get("ts",""),
                    "platform": srv.get("Platform",""),
                })

    my_corr_rows = [c for c in corrs.values() if c.get("user") == uname]

    return jsonify({
        "my_tag_count":   len(my_tag_rows),
        "my_note_count":  len(my_note_rows),
        "my_corr_count":  len(my_corr_rows),
        "my_tags":        my_tag_rows,
        "my_notes":       my_note_rows,
        "my_corrections": my_corr_rows,
    })

@app.route("/")
def index():
    return render_template_string(MAIN_HTML)

# ═══════════════════════════════════════════════════════════════════════════════
#  LOGIN PAGE HTML
# ═══════════════════════════════════════════════════════════════════════════════
LOGIN_HTML = """<!DOCTYPE html>
<html lang="en"><head>
<meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Login — Server Asset Inventory</title>
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{background:#07090f;color:#edf2ff;font-family:'DejaVu Sans',system-ui,sans-serif;
  min-height:100vh;display:flex;align-items:center;justify-content:center}
.card{background:#111827;border:1px solid #1e2a40;border-radius:14px;
  padding:40px 36px;width:360px;box-shadow:0 8px 40px rgba(0,0,0,.5)}
.logo{font-size:2rem;text-align:center;margin-bottom:6px}
h1{text-align:center;font-size:1.15rem;color:#edf2ff;margin-bottom:4px}
.sub{text-align:center;font-size:.78rem;color:#7b8db0;margin-bottom:28px}
label{display:block;font-size:.76rem;color:#7b8db0;margin-bottom:5px;
  text-transform:uppercase;letter-spacing:.08em}
input{width:100%;background:#0d1321;border:1px solid #1e2a40;color:#edf2ff;
  padding:10px 13px;border-radius:8px;font-size:.9rem;margin-bottom:16px}
input:focus{outline:none;border-color:#4a8cff}
.btn{width:100%;background:linear-gradient(135deg,#4a8cff,#6366f1);color:#fff;
  border:none;padding:11px;border-radius:8px;font-size:.95rem;font-weight:700;
  cursor:pointer;transition:opacity .2s}
.btn:hover{opacity:.88}
.err{background:rgba(255,79,106,.12);border:1px solid rgba(255,79,106,.35);
  color:#ff4f6a;border-radius:7px;padding:9px 14px;font-size:.83rem;margin-bottom:14px}
.back{text-align:center;margin-top:18px;font-size:.8rem;color:#7b8db0}
.back a{color:#4a8cff}
</style></head><body>
<div class="card">
  <div class="logo">🖥️</div>
  <h1>Server Asset Inventory</h1>
  <div class="sub">Sign in to access user features</div>
  {% if error %}<div class="err">⚠ {{ error }}</div>{% endif %}
  <form method="POST">
    <label>Username</label>
    <input type="text" name="username" autocomplete="username" autofocus required>
    <label>Password</label>
    <input type="password" name="password" autocomplete="current-password" required>
    <button class="btn" type="submit">Sign In</button>
  </form>
  <div class="back"><a href="/">← Back to Inventory</a></div>
</div>
</body></html>"""

# ═══════════════════════════════════════════════════════════════════════════════
#  MAIN APP HTML
# ═══════════════════════════════════════════════════════════════════════════════
MAIN_HTML = r"""<!DOCTYPE html>
<html lang="en"><head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Server Asset Inventory</title>
<style>
/* ── CSS variables ── */
:root{
  --bg:#07090f;--surf:#111827;--surf2:#161f30;--border:#1e2a40;--border2:#243452;
  --text:#edf2ff;--muted:#7b8db0;--dim:#3d5070;
  --blue:#4a8cff;--green:#30d988;--yellow:#ffc240;--red:#ff4f6a;
  --purple:#a78bfa;--cyan:#22d3ee;--orange:#fb923c;
  --radius:10px;--shadow:0 4px 24px rgba(0,0,0,.45);
}
body.light{
  --bg:#f0f4fb;--surf:#fff;--surf2:#e8edf7;--border:#d0d8ea;--border2:#b8c4da;
  --text:#1a2340;--muted:#5a6a8a;--dim:#a0aec0;
  --blue:#2563eb;--green:#16a34a;--yellow:#d97706;--red:#dc2626;
  --purple:#7c3aed;--cyan:#0891b2;--orange:#ea580c;
}
*{box-sizing:border-box;margin:0;padding:0}
body{background:var(--bg);color:var(--text);font-family:'DejaVu Sans',system-ui,sans-serif;
  font-size:14px;min-height:100vh;transition:background .25s,color .25s}

/* ── Top bar ── */
.topbar{background:var(--surf);border-bottom:1px solid var(--border);padding:0 20px;
  display:flex;align-items:center;gap:12px;height:52px;position:sticky;top:0;z-index:100}
.topbar-title{font-size:1rem;font-weight:700;color:var(--blue);flex:1}
.topbar-meta{font-size:.72rem;color:var(--muted);max-width:320px;overflow:hidden;
  text-overflow:ellipsis;white-space:nowrap}
.user-pill{display:flex;align-items:center;gap:7px;background:var(--surf2);
  border:1px solid var(--border);border-radius:99px;padding:4px 12px 4px 8px;
  font-size:.78rem;font-weight:600;cursor:pointer;color:var(--text)}
.user-pill:hover{border-color:var(--blue)}
.user-avatar{width:22px;height:22px;border-radius:50%;background:var(--blue);
  color:#fff;display:flex;align-items:center;justify-content:center;font-size:.72rem;font-weight:700}
.theme-btn{background:var(--surf2);border:1px solid var(--border);color:var(--text);
  padding:5px 11px;border-radius:6px;cursor:pointer;font-size:.8rem;font-weight:600}
.theme-btn:hover{border-color:var(--blue)}

/* ── Sidebar ── */
.sidebar{width:215px;background:var(--surf);border-right:1px solid var(--border);
  position:fixed;top:52px;left:0;bottom:0;overflow-y:auto;padding:12px 0}
.nav-section{padding:12px 18px 5px;font-size:.65rem;font-weight:700;
  text-transform:uppercase;letter-spacing:.12em;color:var(--dim)}
.nav-item{display:flex;align-items:center;gap:9px;padding:8px 18px;cursor:pointer;
  color:var(--muted);font-size:.85rem;font-weight:500;border-left:3px solid transparent;
  transition:all .15s;user-select:none}
.nav-item:hover{background:var(--surf2);color:var(--text)}
.nav-item.active{color:var(--blue);border-left-color:var(--blue);background:var(--surf2)}
.nav-icon{font-size:.95rem;width:18px;text-align:center}
.nav-badge{margin-left:auto;background:var(--red);color:#fff;border-radius:99px;
  padding:1px 7px;font-size:.65rem;font-weight:700}

/* ── Main ── */
.main{margin-left:215px;padding:22px;margin-top:52px;min-height:calc(100vh - 52px);
  max-width:calc(100vw - 215px);overflow-x:auto}
.page{display:none;max-width:100%}.page.active{display:block}

/* ── Section header ── */
.shdr{font-size:.78rem;font-weight:700;text-transform:uppercase;letter-spacing:.1em;
  color:var(--blue);border-bottom:1px solid var(--border);padding-bottom:7px;margin-bottom:14px}

/* ── Metric cards ── */
.metrics-row{display:grid;grid-template-columns:repeat(auto-fit,minmax(130px,1fr));
  gap:12px;margin-bottom:20px}
.metric-card{background:var(--surf);border:1px solid var(--border);border-radius:var(--radius);
  padding:14px 16px;text-align:center}
.metric-val{font-size:1.9rem;font-weight:700;font-family:monospace;letter-spacing:-1px}
.metric-lbl{font-size:.68rem;font-weight:600;text-transform:uppercase;letter-spacing:.1em;
  color:var(--muted);margin-top:3px}
.c-blue{color:var(--blue)}.c-green{color:var(--green)}.c-red{color:var(--red)}
.c-purple{color:var(--purple)}.c-cyan{color:var(--cyan)}.c-yellow{color:var(--yellow)}
.c-orange{color:var(--orange)}

/* ── Charts ── */
.charts-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(300px,1fr));gap:16px}
.chart-card{background:var(--surf);border:1px solid var(--border);border-radius:var(--radius);
  padding:12px;overflow:hidden}
.chart-card img{width:100%;height:auto;border-radius:6px}

/* ── Grouped summary cards ── */
.group-cards-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(240px,1fr));gap:12px}
.group-card{background:var(--surf);border:1px solid var(--border);border-radius:var(--radius);
  padding:14px 16px;cursor:pointer;transition:all .15s}
.group-card:hover{border-color:var(--blue);transform:translateY(-1px)}
.group-card.active{border-color:var(--blue);background:var(--surf2)}
.group-card-name{font-size:.92rem;font-weight:700;color:var(--text);margin-bottom:8px;
  white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.group-card-count{font-size:1.7rem;font-weight:700;font-family:monospace;color:var(--blue)}
.group-card-bar{display:flex;height:6px;border-radius:3px;overflow:hidden;margin-top:8px;
  background:var(--border)}
.group-card-bar .seg-live{background:var(--green)}
.group-card-bar .seg-notlive{background:var(--red)}
.group-card-legend{display:flex;justify-content:space-between;font-size:.68rem;
  color:var(--muted);margin-top:5px}

/* ── Sortable table headers ── */
th.sortable{cursor:pointer;user-select:none;position:relative}
th.sortable:hover{color:var(--blue)}
th.sortable .sort-arrow{margin-left:4px;font-size:.7rem;opacity:.5}
th.sortable.sort-active .sort-arrow{opacity:1;color:var(--blue)}

/* ── Upload zone ── */
.upload-zone{border:2px dashed var(--border);border-radius:var(--radius);padding:30px;
  text-align:center;cursor:pointer;transition:all .2s;background:var(--surf)}
.upload-zone:hover,.upload-zone.drag{border-color:var(--blue);background:var(--surf2)}
.upload-zone input{display:none}

/* ── Filter bar ── */
.filter-bar{background:var(--surf);border:1px solid var(--border);border-radius:var(--radius);
  padding:12px 16px;margin-bottom:14px;display:flex;flex-wrap:wrap;gap:9px;align-items:flex-end}
.filter-bar label{font-size:.7rem;color:var(--muted);display:block;margin-bottom:3px}
.filter-bar input,.filter-bar select{background:var(--surf2);border:1px solid var(--border);
  color:var(--text);padding:7px 10px;border-radius:7px;font-size:.82rem;min-width:130px}
.filter-bar input:focus,.filter-bar select:focus{outline:none;border-color:var(--blue)}

/* ── Chips ── */
.chip-row{display:flex;flex-wrap:wrap;gap:7px;margin-bottom:12px}
.chip{padding:4px 13px;border-radius:99px;border:1px solid var(--border);
  background:var(--surf2);color:var(--muted);cursor:pointer;font-size:.78rem;
  font-weight:600;transition:all .15s}
.chip:hover{border-color:var(--blue);color:var(--blue)}
.chip.active{background:var(--blue);color:#fff;border-color:var(--blue)}

/* ── Buttons ── */
.btn{background:var(--blue);color:#fff;border:none;padding:7px 15px;border-radius:7px;
  cursor:pointer;font-size:.82rem;font-weight:600;transition:opacity .15s}
.btn:hover{opacity:.85}
.btn-sm{padding:3px 10px;font-size:.74rem}
.btn-xs{padding:2px 7px;font-size:.68rem;border-radius:5px}
.btn-outline{background:transparent;border:1px solid var(--border);color:var(--muted)}
.btn-outline:hover{border-color:var(--blue);color:var(--blue)}
.btn-green{background:var(--green);color:#000}
.btn-red{background:var(--red)}
.btn-yellow{background:var(--yellow);color:#000}
.btn-purple{background:var(--purple)}
.btn-orange{background:var(--orange)}

/* ── Table ── */
.tbl-wrap{overflow-x:auto;border-radius:var(--radius);border:1px solid var(--border)}
table{width:100%;border-collapse:collapse;font-size:.79rem}
th{background:var(--surf2);padding:8px 11px;text-align:left;font-size:.7rem;
  font-weight:700;text-transform:uppercase;letter-spacing:.06em;color:var(--muted);
  border-bottom:1px solid var(--border);white-space:nowrap}
td{padding:7px 11px;border-bottom:1px solid var(--border);color:var(--text);vertical-align:middle}
tr:last-child td{border-bottom:none}
tr:hover td{background:var(--surf2)}
.tbl-count{font-size:.76rem;color:var(--muted);margin-bottom:7px}
.pagination{display:flex;gap:5px;margin-top:10px;flex-wrap:wrap;align-items:center}
.page-btn{background:var(--surf2);border:1px solid var(--border);color:var(--muted);
  padding:3px 9px;border-radius:5px;cursor:pointer;font-size:.76rem}
.page-btn:hover,.page-btn.active{background:var(--blue);color:#fff;border-color:var(--blue)}

/* ── Badges ── */
.badge{display:inline-block;padding:2px 8px;border-radius:99px;font-size:.7rem;font-weight:700}
.badge-live{background:rgba(48,217,136,.15);color:var(--green);border:1px solid rgba(48,217,136,.3)}
.badge-notlive{background:rgba(255,79,106,.15);color:var(--red);border:1px solid rgba(255,79,106,.3)}
.badge-physical{background:rgba(74,140,255,.15);color:var(--blue);border:1px solid rgba(74,140,255,.3)}
.badge-virtual{background:rgba(167,139,250,.15);color:var(--purple);border:1px solid rgba(167,139,250,.3)}
.badge-other{background:var(--surf2);color:var(--muted);border:1px solid var(--border)}

/* ── Tag pills ── */
.tag-pill{display:inline-flex;align-items:center;gap:4px;background:rgba(167,139,250,.15);
  border:1px solid rgba(167,139,250,.3);color:var(--purple);border-radius:99px;
  padding:2px 8px;font-size:.68rem;font-weight:600;margin:2px}
.tag-pill .rm{cursor:pointer;color:var(--muted);font-size:.75rem}
.tag-pill .rm:hover{color:var(--red)}

/* ── Modal ── */
.modal-overlay{display:none;position:fixed;inset:0;background:rgba(0,0,0,.65);
  z-index:500;align-items:flex-start;justify-content:center;padding-top:60px;overflow-y:auto}
.modal-overlay.show{display:flex}
.modal{background:var(--surf);border:1px solid var(--border2);border-radius:14px;
  padding:26px;max-width:720px;width:94%;box-shadow:var(--shadow);margin-bottom:40px}
.modal-title{font-size:1.1rem;font-weight:700;color:var(--blue);margin-bottom:16px;
  display:flex;justify-content:space-between;align-items:center}
.modal-close{cursor:pointer;color:var(--muted);font-size:1.3rem}
.modal-close:hover{color:var(--red)}
.detail-grid{display:grid;grid-template-columns:1fr 1fr;gap:10px;margin-bottom:16px}
.detail-row{background:var(--surf2);border-radius:8px;padding:9px 13px}
.detail-key{font-size:.67rem;text-transform:uppercase;letter-spacing:.08em;color:var(--muted);margin-bottom:2px}
.detail-val{font-size:.86rem;font-weight:600;word-break:break-word}
.detail-full{grid-column:1/-1}
.notes-list{max-height:220px;overflow-y:auto}
.note-item{background:var(--surf2);border-radius:8px;padding:10px 13px;margin-bottom:8px;
  border-left:3px solid var(--blue)}
.note-meta{font-size:.68rem;color:var(--muted);margin-bottom:4px}
.note-text{font-size:.83rem}
.note-del{float:right;cursor:pointer;color:var(--muted);font-size:.72rem}
.note-del:hover{color:var(--red)}
.note-input-row{display:flex;gap:8px;margin-top:10px}
.note-input-row textarea{flex:1;background:var(--surf2);border:1px solid var(--border);
  color:var(--text);padding:8px 11px;border-radius:7px;font-size:.83rem;
  font-family:inherit;resize:none;height:60px}
.note-input-row textarea:focus{outline:none;border-color:var(--blue)}
.flag-box{background:rgba(255,79,106,.08);border:1px solid rgba(255,79,106,.25);
  border-radius:8px;padding:10px 14px;margin-bottom:12px}
.flag-box h4{color:var(--red);font-size:.8rem;margin-bottom:4px}

/* ── Alerts ── */
.alert{border-radius:8px;padding:11px 14px;font-size:.84rem;margin-bottom:12px}
.alert-info{background:rgba(74,140,255,.1);border:1px solid rgba(74,140,255,.3);color:var(--blue)}
.alert-success{background:rgba(48,217,136,.1);border:1px solid rgba(48,217,136,.3);color:var(--green)}
.alert-warn{background:rgba(255,194,64,.1);border:1px solid rgba(255,194,64,.3);color:var(--yellow)}
.alert-error{background:rgba(255,79,106,.1);border:1px solid rgba(255,79,106,.3);color:var(--red)}

/* ── Admin panel ── */
.admin-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(200px,1fr));gap:14px;margin-bottom:20px}
.admin-card{background:var(--surf);border:1px solid var(--border);border-radius:var(--radius);padding:18px;text-align:center}
.admin-card .num{font-size:2rem;font-weight:700;font-family:monospace}
.admin-card .lbl{font-size:.72rem;color:var(--muted);text-transform:uppercase;letter-spacing:.09em;margin-top:4px}
.admin-card .actions{margin-top:12px;display:flex;gap:6px;justify-content:center;flex-wrap:wrap}
.user-row{display:flex;align-items:center;gap:10px;padding:10px 14px;
  background:var(--surf2);border-radius:8px;margin-bottom:7px}
.user-info{flex:1}
.user-name{font-weight:700;font-size:.88rem}
.user-meta{font-size:.72rem;color:var(--muted)}
.role-badge{padding:2px 8px;border-radius:99px;font-size:.7rem;font-weight:700}
.role-admin{background:rgba(255,194,64,.18);color:var(--yellow);border:1px solid rgba(255,194,64,.35)}
.role-user{background:rgba(74,140,255,.15);color:var(--blue);border:1px solid rgba(74,140,255,.3)}
.disabled-row{opacity:.45}

/* ── Bulk lookup ── */
.bulk-area{width:100%;min-height:110px;background:var(--surf2);border:1px solid var(--border);
  color:var(--text);padding:11px;border-radius:8px;font-family:monospace;font-size:.84rem;resize:vertical}
.bulk-area:focus{outline:none;border-color:var(--blue)}
.not-found-box{background:rgba(255,79,106,.08);border:1px solid rgba(255,79,106,.25);
  border-radius:8px;padding:11px 14px;margin-top:10px}
.not-found-box h4{color:var(--red);font-size:.8rem;margin-bottom:6px}

/* ── Compare ── */
.compare-uploads{display:grid;grid-template-columns:1fr 1fr;gap:14px;margin-bottom:18px}
.compare-box{background:var(--surf);border:1px solid var(--border);border-radius:var(--radius);padding:16px}
.diff-cards{display:grid;grid-template-columns:repeat(auto-fit,minmax(130px,1fr));gap:11px;margin-bottom:18px}
.diff-card{background:var(--surf);border:1px solid var(--border);border-radius:var(--radius);
  padding:12px 14px;text-align:center}
.diff-val{font-size:1.6rem;font-weight:700;font-family:monospace}
.diff-lbl{font-size:.7rem;color:var(--muted);margin-top:3px;text-transform:uppercase;letter-spacing:.06em}

/* ── Quality panel ── */
.quality-panel{background:rgba(255,194,64,.07);border:1px solid rgba(255,194,64,.25);
  border-radius:8px;padding:11px 14px;margin-bottom:14px}
.quality-panel h4{color:var(--yellow);font-size:.8rem;margin-bottom:6px}

/* ── Login required notice ── */
.login-notice{background:var(--surf2);border:1px solid var(--border);border-radius:8px;
  padding:12px 16px;font-size:.83rem;color:var(--muted);display:flex;align-items:center;gap:10px}
.login-notice a{color:var(--blue);font-weight:600}

/* ── Spinner ── */
.spinner{display:inline-block;width:16px;height:16px;border:2px solid var(--border);
  border-top-color:var(--blue);border-radius:50%;animation:spin .7s linear infinite}
@keyframes spin{to{transform:rotate(360deg)}}
.loading-row{text-align:center;padding:28px;color:var(--muted)}

/* ── Input group ── */
.input-row{display:flex;gap:8px;margin-bottom:10px;flex-wrap:wrap}
.input-row input,.input-row select{background:var(--surf2);border:1px solid var(--border);
  color:var(--text);padding:8px 11px;border-radius:7px;font-size:.84rem;flex:1;min-width:130px}
.input-row input:focus,.input-row select:focus{outline:none;border-color:var(--blue)}

@media(max-width:768px){
  .sidebar{display:none}.main{margin-left:0;max-width:100vw;padding:16px;overflow-x:auto}
  .compare-uploads,.detail-grid{grid-template-columns:1fr}
}
</style>
</head>
<body>

<!-- ── Top bar ── -->
<div class="topbar">
  <div class="topbar-title">🖥️ Server Asset Inventory</div>
  <div class="topbar-meta" id="topbar-meta">No data loaded</div>
  <div id="user-area"></div>
  <button class="theme-btn" onclick="toggleTheme()">🌙/☀️</button>
</div>

<!-- ── Sidebar ── -->
<div class="sidebar">
  <div class="nav-section">Public</div>
  <div class="nav-item active" onclick="showPage('dashboard',this)">
    <span class="nav-icon">📊</span> Dashboard</div>
  <div class="nav-item" onclick="showPage('search',this)">
    <span class="nav-icon">🔍</span> Search & Filter</div>
  <div class="nav-item" onclick="showPage('grouped',this)">
    <span class="nav-icon">🗂️</span> Grouped Summary</div>
  <div class="nav-item" onclick="showPage('bulk',this)">
    <span class="nav-icon">📋</span> Bulk Lookup</div>

  <div class="nav-section" id="nav-user-section" style="display:none">Logged In</div>
  <div class="nav-item" id="nav-activity" style="display:none"
       onclick="showPage('activity',this)">
    <span class="nav-icon">🏷️</span> My Activity</div>

  <div class="nav-section" id="nav-report-section" style="display:none">Report Changes</div>
  <div class="nav-item" id="nav-decom" style="display:none"
       onclick="showPage('decom',this)">
    <span class="nav-icon">🗑️</span> Mark Decommission</div>
  <div class="nav-item" id="nav-newsrv" style="display:none"
       onclick="showPage('newsrv',this)">
    <span class="nav-icon">➕</span> Suggest New Server</div>
  <div class="nav-item" id="nav-tshirt" style="display:none"
       onclick="showPage('tshirt',this)">
    <span class="nav-icon">👕</span> Update CPU / RAM</div>

  <div class="nav-section" id="nav-admin-section" style="display:none">Admin</div>
  <div class="nav-item" id="nav-upload" style="display:none"
       onclick="showPage('upload',this)">
    <span class="nav-icon">📤</span> Upload Data</div>
  <div class="nav-item" id="nav-compare" style="display:none"
       onclick="showPage('compare',this)">
    <span class="nav-icon">📈</span> Month Compare</div>
  <div class="nav-item" id="nav-admin" style="display:none"
       onclick="showPage('admin',this)">
    <span class="nav-icon">⚙️</span> Admin Panel
    <span class="nav-badge" id="corr-badge" style="display:none">0</span>
  </div>
  <div class="nav-item" id="nav-users" style="display:none"
       onclick="showPage('users',this)">
    <span class="nav-icon">👥</span> Manage Users</div>
</div>

<!-- ── Main ── -->
<div class="main">

<!-- DASHBOARD -->
<div class="page active" id="page-dashboard">
  <div class="shdr">Dashboard</div>
  <div id="dash-loading" class="loading-row"><div class="spinner"></div> Loading…</div>
  <div id="dash-content" style="display:none">
    <div class="metrics-row" id="metrics-row"></div>
    <div class="charts-grid" id="charts-grid"></div>
  </div>
</div>

<!-- SEARCH -->
<div class="page" id="page-search">
  <div class="shdr">Search & Filter</div>
  <div class="chip-row">
    <span style="font-size:.75rem;color:var(--muted);align-self:center">Quick:</span>
    <span class="chip" onclick="qf('platform','Dynamo',this)">Dynamo</span>
    <span class="chip" onclick="qf('platform','Digital Journey',this)">Digital Journey</span>
    <span class="chip" onclick="qf('platform','EMA',this)">EMA</span>
    <span class="chip" onclick="qf('platform','EPMC',this)">EPMC</span>
    <span class="chip" onclick="qf('platform','JEA',this)">JEA</span>
    <span class="chip" onclick="qf('hpc','HPC',this)">HPC</span>
    <span class="chip" onclick="qf('hpc','JPC',this)">JPC</span>
    <span class="chip btn-outline" onclick="clearFilters()">✕ Clear</span>
  </div>
  <div class="filter-bar">
    <div><label>Search hostname / app</label>
      <input type="text" id="f-q" placeholder="Type to search…" oninput="debSearch()"></div>
    <div><label>Platform</label>
      <select id="f-platform" onchange="runSearch(1)"><option value="">All</option></select></div>
    <div><label title="Live = server is in BAU. Not Live = server is in project/testing phase, before operational acceptance into BAU.">Status ℹ️</label>
      <select id="f-status" onchange="runSearch(1)"><option value="">All</option></select></div>
    <div><label>Server Role</label>
      <select id="f-role" onchange="runSearch(1)"><option value="">All</option></select></div>
    <div><label>HPC Group</label>
      <select id="f-hpc" onchange="runSearch(1)"><option value="">All</option></select></div>
    <div><label>DC Location</label>
      <select id="f-loc" onchange="runSearch(1)"><option value="">All</option></select></div>
    <div><label>OS</label>
      <select id="f-os" onchange="runSearch(1)"><option value="">All</option></select></div>
    <div><label>Server Type</label>
      <select id="f-stype" onchange="runSearch(1)"><option value="">All</option></select></div>
    <div style="align-self:flex-end">
      <button class="btn btn-green btn-sm" onclick="exportSearch()">📥 Export</button>
    </div>
  </div>
  <div class="tbl-count" id="search-count"></div>
  <div class="tbl-wrap">
    <table><thead><tr id="search-thead"></tr></thead>
    <tbody id="search-tbody">
      <tr><td class="loading-row" colspan="12">Upload a file to search.</td></tr>
    </tbody></table>
  </div>
  <div class="pagination" id="search-pagination"></div>
</div>

<!-- GROUPED SUMMARY -->
<div class="page" id="page-grouped">
  <div class="shdr">🗂️ Grouped Summary</div>
  <p style="color:var(--muted);font-size:.84rem;margin-bottom:14px">
    See your inventory organized by a dimension — counts and Live/Not-Live split per group.
    Click any group to drill into its servers.
  </p>

  <div class="chip-row" id="grouped-by-chips">
    <span style="font-size:.78rem;color:var(--muted);align-self:center">Group by:</span>
    <span class="chip active" data-by="platform" onclick="setGroupBy('platform',this)">Platform</span>
    <span class="chip" data-by="loc" onclick="setGroupBy('loc',this)">DC Location</span>
    <span class="chip" data-by="hpc" onclick="setGroupBy('hpc',this)">HPC Group</span>
    <span class="chip" data-by="role" onclick="setGroupBy('role',this)">Server Role</span>
    <span class="chip" data-by="os" onclick="setGroupBy('os',this)">OS</span>
    <span class="chip" data-by="stype" onclick="setGroupBy('stype',this)">Server Type</span>
    <span class="chip" data-by="app" onclick="setGroupBy('app',this)">Application</span>
  </div>

  <div id="grouped-summary-loading" class="loading-row"><div class="spinner"></div></div>
  <div id="grouped-cards" style="display:none"></div>

  <!-- Drill-down panel -->
  <div id="grouped-drill-panel" style="display:none;margin-top:20px">
    <div class="shdr" id="grouped-drill-title" style="color:var(--green)"></div>
    <div class="tbl-count" id="grouped-drill-count"></div>
    <div style="margin-bottom:10px">
      <button class="btn btn-sm btn-green" id="grouped-drill-export">📥 Export this group</button>
      <button class="btn btn-sm btn-outline" onclick="closeGroupedDrill()">✕ Close</button>
    </div>
    <div class="tbl-wrap">
      <table><thead><tr id="grouped-drill-thead"></tr></thead>
      <tbody id="grouped-drill-tbody"></tbody></table>
    </div>
    <div class="pagination" id="grouped-drill-pagination"></div>
  </div>
</div>

<!-- BULK -->
<div class="page" id="page-bulk">
  <div class="shdr">Bulk Server Lookup</div>
  <div style="max-width:660px">
    <p style="color:var(--muted);font-size:.84rem;margin-bottom:10px">
      Paste hostnames — one per line. Useful for tickets, change requests, or incident response.
    </p>
    <textarea class="bulk-area" id="bulk-input"
      placeholder="srv-prod-001&#10;srv-prod-002&#10;appserver-12&#10;…"></textarea>
    <div style="display:flex;gap:8px;margin-top:9px">
      <button class="btn" onclick="runBulk()">🔍 Look Up</button>
      <button class="btn btn-outline" onclick="document.getElementById('bulk-input').value=''">Clear</button>
    </div>
    <div id="bulk-summary" style="margin-top:12px"></div>
    <div id="bulk-not-found"></div>
    <div id="bulk-results" style="margin-top:12px"></div>
  </div>
</div>

<!-- COMPARE -->
<div class="page" id="page-compare">
  <div class="shdr">Month-on-Month Comparison</div>
  <div class="compare-uploads">
    <div class="compare-box">
      <h4 style="font-size:.85rem;color:var(--muted);margin-bottom:8px">📅 Current Month</h4>
      <div style="font-size:.82rem">Loaded: <span id="curr-file-lbl" style="color:var(--blue)">None</span></div>
    </div>
    <div class="compare-box">
      <h4 style="font-size:.85rem;color:var(--muted);margin-bottom:8px">📅 Previous Month</h4>
      <div class="upload-zone" style="padding:16px"
           onclick="document.getElementById('prev-input').click()">
        <div style="font-size:.88rem">📂 Click to upload previous month</div>
        <div style="font-size:.72rem;color:var(--muted);margin-top:4px">.xlsx / .xls / .csv</div>
        <input type="file" id="prev-input" accept=".xlsx,.xls,.csv"
               onchange="uploadPrev(this.files[0])">
      </div>
      <div id="prev-status" style="margin-top:8px;font-size:.8rem;color:var(--muted)"></div>
    </div>
  </div>
  <button class="btn" onclick="runCompare()" style="margin-bottom:18px">📊 Run Comparison</button>
  <div id="compare-out"></div>
</div>

<!-- ACTIVITY (logged-in users) -->
<div class="page" id="page-activity">
  <div class="shdr">🏷️ My Activity — Tags, Notes & Flags</div>
  <div id="activity-out">
    <div class="alert alert-info">Loading activity data…</div>
  </div>
</div>

<!-- MARK DECOMMISSION -->
<div class="page" id="page-decom">
  <div class="shdr">🗑️ Mark Servers for Decommission / Shutdown</div>
  <div style="max-width:660px">
    <p style="color:var(--muted);font-size:.84rem;margin-bottom:14px">
      Report one or more servers shut down or decommissioned together
      (e.g. as part of one activity). Admin will review and apply this in
      next month's inventory.
    </p>
    <div style="font-size:.74rem;color:var(--muted);margin-bottom:4px">Add Servers</div>
    <input type="text" id="decom-hostname" placeholder="Start typing hostname, then pick from the list…"
      style="width:100%;background:var(--surf2);border:1px solid var(--border);color:var(--text);
      padding:9px 12px;border-radius:8px;font-size:.86rem;margin-bottom:4px"
      oninput="decomHostSearch()" autocomplete="off">
    <div id="decom-host-suggest" style="margin-bottom:8px"></div>

    <!-- Selected server chips -->
    <div id="decom-selected-chips" style="display:flex;flex-wrap:wrap;gap:6px;margin-bottom:12px"></div>

    <div style="display:flex;gap:10px;margin-bottom:10px">
      <div style="flex:1">
        <div style="font-size:.74rem;color:var(--muted);margin-bottom:4px">Reason (applies to all selected servers)</div>
        <select id="decom-reason" style="width:100%;background:var(--surf2);border:1px solid var(--border);
          color:var(--text);padding:9px 12px;border-radius:8px;font-size:.86rem">
          <option value="">— Select reason —</option>
        </select>
      </div>
      <div style="flex:1">
        <div style="font-size:.74rem;color:var(--muted);margin-bottom:4px">Effective Date (optional)</div>
        <input type="date" id="decom-date"
          style="width:100%;background:var(--surf2);border:1px solid var(--border);color:var(--text);
          padding:9px 12px;border-radius:8px;font-size:.86rem">
      </div>
    </div>
    <div style="font-size:.74rem;color:var(--muted);margin-bottom:4px">Comment (optional — e.g. activity / ticket reference)</div>
    <input type="text" id="decom-comment" placeholder="e.g. Confirmed by infra team — Change CHG0012345"
      style="width:100%;background:var(--surf2);border:1px solid var(--border);color:var(--text);
      padding:9px 12px;border-radius:8px;font-size:.86rem;margin-bottom:14px">

    <button class="btn btn-red" onclick="submitDecom()">🗑️ Submit <span id="decom-submit-count"></span> for Review</button>
    <div id="decom-msg" style="margin-top:12px"></div>

    <div class="shdr" style="margin-top:24px;color:var(--red)">My Submitted Decommissions</div>
    <div id="decom-my-list"></div>
  </div>
</div>

<!-- SUGGEST NEW SERVER -->
<div class="page" id="page-newsrv">
  <div class="shdr">➕ Suggest New Server (Not Yet in Inventory)</div>
  <div style="max-width:680px">
    <p style="color:var(--muted);font-size:.84rem;margin-bottom:14px">
      Report a server that exists but isn't in the current inventory yet.
      Admin will review and add it to next month's upload.
    </p>
    <div id="newsrv-form-grid" style="display:grid;grid-template-columns:repeat(auto-fit,minmax(260px,1fr));gap:10px;margin-bottom:12px"></div>
    <div style="font-size:.74rem;color:var(--muted);margin-bottom:4px">Comment (optional)</div>
    <input type="text" id="newsrv-comment" placeholder="e.g. New VM provisioned for Project X"
      style="width:100%;background:var(--surf2);border:1px solid var(--border);color:var(--text);
      padding:9px 12px;border-radius:8px;font-size:.86rem;margin-bottom:14px">

    <button class="btn btn-green" onclick="submitNewServer()">➕ Submit for Review</button>
    <div id="newsrv-msg" style="margin-top:12px"></div>

    <div class="shdr" style="margin-top:24px;color:var(--green)">My Submitted New Servers</div>
    <div id="newsrv-my-list"></div>
  </div>
</div>

<!-- UPDATE CPU / RAM (T-SHIRT SIZE) -->
<div class="page" id="page-tshirt">
  <div class="shdr">👕 Update CPU / RAM (T-Shirt Size)</div>
  <div style="max-width:660px">
    <p style="color:var(--muted);font-size:.84rem;margin-bottom:14px">
      Report a CPU or RAM change for an <b>EMA</b>, <b>EPMC</b>, or <b>JEA</b> server —
      these platforms' contract pricing depends on server T-shirt size.
      Admin will review and reconcile this with billing.
    </p>

    <div style="font-size:.74rem;color:var(--muted);margin-bottom:4px">Server Hostname</div>
    <input type="text" id="tshirt-hostname" placeholder="Start typing hostname…"
      style="width:100%;background:var(--surf2);border:1px solid var(--border);color:var(--text);
      padding:9px 12px;border-radius:8px;font-size:.86rem;margin-bottom:4px"
      oninput="tshirtHostSearch()" autocomplete="off">
    <div id="tshirt-host-suggest" style="margin-bottom:4px"></div>
    <div id="tshirt-host-selected" style="margin-bottom:10px"></div>

    <div style="display:flex;gap:10px;margin-bottom:10px">
      <div style="flex:1">
        <div style="font-size:.74rem;color:var(--muted);margin-bottom:4px">Current CPU (optional)</div>
        <input type="text" id="tshirt-current-cpu" placeholder="e.g. 4 vCPU"
          style="width:100%;background:var(--surf2);border:1px solid var(--border);color:var(--text);
          padding:9px 12px;border-radius:8px;font-size:.86rem">
      </div>
      <div style="flex:1">
        <div style="font-size:.74rem;color:var(--muted);margin-bottom:4px">Current RAM (optional)</div>
        <input type="text" id="tshirt-current-ram" placeholder="e.g. 16 GB"
          style="width:100%;background:var(--surf2);border:1px solid var(--border);color:var(--text);
          padding:9px 12px;border-radius:8px;font-size:.86rem">
      </div>
    </div>
    <div style="display:flex;gap:10px;margin-bottom:10px">
      <div style="flex:1">
        <div style="font-size:.74rem;color:var(--muted);margin-bottom:4px">New CPU</div>
        <input type="text" id="tshirt-new-cpu" placeholder="e.g. 8 vCPU"
          style="width:100%;background:var(--surf2);border:1px solid var(--border);color:var(--text);
          padding:9px 12px;border-radius:8px;font-size:.86rem">
      </div>
      <div style="flex:1">
        <div style="font-size:.74rem;color:var(--muted);margin-bottom:4px">New RAM</div>
        <input type="text" id="tshirt-new-ram" placeholder="e.g. 32 GB"
          style="width:100%;background:var(--surf2);border:1px solid var(--border);color:var(--text);
          padding:9px 12px;border-radius:8px;font-size:.86rem">
      </div>
    </div>
    <div style="font-size:.74rem;color:var(--muted);margin-bottom:4px">Reason (optional)</div>
    <input type="text" id="tshirt-reason" placeholder="e.g. Application load increased, approved by project lead"
      style="width:100%;background:var(--surf2);border:1px solid var(--border);color:var(--text);
      padding:9px 12px;border-radius:8px;font-size:.86rem;margin-bottom:14px">

    <button class="btn" style="background:var(--cyan);color:#000" onclick="submitTshirt()">👕 Submit for Review</button>
    <div id="tshirt-msg" style="margin-top:12px"></div>

    <div class="shdr" style="margin-top:24px;color:var(--cyan)">My Submitted CPU/RAM Changes</div>
    <div id="tshirt-my-list"></div>
  </div>
</div>

<!-- UPLOAD (admin) -->
<div class="page" id="page-upload">
  <div class="shdr">📤 Upload Inventory File</div>
  <div style="max-width:520px">
    <div class="upload-zone" id="drop-zone"
         onclick="document.getElementById('file-input').click()"
         ondragover="event.preventDefault();this.classList.add('drag')"
         ondragleave="this.classList.remove('drag')"
         ondrop="handleDrop(event)">
      <div style="font-size:2rem">📂</div>
      <div style="font-size:.95rem;font-weight:600;margin-top:8px">
        Click or drag your Excel / CSV file here</div>
      <div style="font-size:.78rem;color:var(--muted);margin-top:4px">
        .xlsx · .xls · .csv — max 50 MB</div>
      <input type="file" id="file-input" accept=".xlsx,.xls,.csv"
             onchange="uploadFile(this.files[0])">
    </div>
    <div id="upload-status" style="margin-top:12px"></div>
  </div>
</div>

<!-- ADMIN PANEL -->
<div class="page" id="page-admin">
  <div class="shdr">⚙️ Admin Panel</div>
  <div id="admin-loading" class="loading-row"><div class="spinner"></div></div>
  <div id="admin-content" style="display:none">

    <!-- Monthly Changes banner -->
    <div id="monthly-changes-banner" style="margin-bottom:16px"></div>

    <!-- Summary cards -->
    <div class="admin-grid" id="admin-stats"></div>

    <!-- ── CORRECTIONS ── -->
    <div style="margin-top:20px">
      <div class="shdr" style="color:var(--orange)">
        ✏️ Column Correction Suggestions
        <span id="corr-pending-badge" style="display:none;margin-left:8px;
          background:var(--orange);color:#000;border-radius:99px;padding:1px 9px;
          font-size:.72rem">0 pending</span>
        <a href="/admin/corrections_report" style="float:right">
          <button class="btn btn-xs btn-orange">📥 Export All</button></a>
      </div>
      <div style="display:flex;gap:6px;margin-bottom:12px">
        <button class="btn btn-sm" id="corr-tab-pending"
          onclick="showCorrTab('pending')" style="background:var(--orange)">
          ⏳ Pending <span id="corr-pending-count"></span></button>
        <button class="btn btn-sm btn-outline" id="corr-tab-approved"
          onclick="showCorrTab('approved')">
          ✅ Approved <span id="corr-approved-count"></span></button>
        <button class="btn btn-sm btn-outline" id="corr-tab-rejected"
          onclick="showCorrTab('rejected')">
          ❌ Rejected <span id="corr-rejected-count"></span></button>
      </div>
      <div id="corr-pending-table"></div>
      <div id="corr-approved-table" style="display:none"></div>
      <div id="corr-rejected-table" style="display:none"></div>
    </div>

    <!-- ── DECOMMISSIONS ── -->
    <div style="margin-top:20px">
      <div class="shdr" style="color:var(--red)">
        🗑️ Decommission / Shutdown Reports
        <span id="decom-pending-badge" style="display:none;margin-left:8px;
          background:var(--red);color:#fff;border-radius:99px;padding:1px 9px;
          font-size:.72rem">0 pending</span>
        <a href="/admin/decommissions_report" style="float:right">
          <button class="btn btn-xs btn-red">📥 Export All</button></a>
      </div>
      <div style="display:flex;gap:6px;margin-bottom:12px">
        <button class="btn btn-sm" id="decom-tab-pending"
          onclick="showDecomTab('pending')" style="background:var(--red)">
          ⏳ Pending <span id="decom-pending-count"></span></button>
        <button class="btn btn-sm btn-outline" id="decom-tab-approved"
          onclick="showDecomTab('approved')">
          ✅ Approved <span id="decom-approved-count"></span></button>
        <button class="btn btn-sm btn-outline" id="decom-tab-rejected"
          onclick="showDecomTab('rejected')">
          ❌ Rejected <span id="decom-rejected-count"></span></button>
      </div>
      <div id="decom-pending-table"></div>
      <div id="decom-approved-table" style="display:none"></div>
      <div id="decom-rejected-table" style="display:none"></div>
    </div>

    <!-- ── NEW SERVERS ── -->
    <div style="margin-top:20px">
      <div class="shdr" style="color:var(--green)">
        ➕ New Server Suggestions
        <span id="newsrv-pending-badge" style="display:none;margin-left:8px;
          background:var(--green);color:#000;border-radius:99px;padding:1px 9px;
          font-size:.72rem">0 pending</span>
        <a href="/admin/newservers_report" style="float:right">
          <button class="btn btn-xs btn-green">📥 Export All</button></a>
      </div>
      <div style="display:flex;gap:6px;margin-bottom:12px">
        <button class="btn btn-sm" id="newsrv-tab-pending"
          onclick="showNewsrvTab('pending')" style="background:var(--green);color:#000">
          ⏳ Pending <span id="newsrv-pending-count"></span></button>
        <button class="btn btn-sm btn-outline" id="newsrv-tab-approved"
          onclick="showNewsrvTab('approved')">
          ✅ Approved <span id="newsrv-approved-count"></span></button>
        <button class="btn btn-sm btn-outline" id="newsrv-tab-rejected"
          onclick="showNewsrvTab('rejected')">
          ❌ Rejected <span id="newsrv-rejected-count"></span></button>
      </div>
      <div id="newsrv-pending-table"></div>
      <div id="newsrv-approved-table" style="display:none"></div>
      <div id="newsrv-rejected-table" style="display:none"></div>
    </div>

    <!-- ── T-SHIRT SIZE CHANGES ── -->
    <div style="margin-top:20px">
      <div class="shdr" style="color:var(--cyan)">
        👕 CPU / RAM (T-Shirt Size) Changes — EMA / EPMC / JEA
        <span id="tshirt-pending-badge" style="display:none;margin-left:8px;
          background:var(--cyan);color:#000;border-radius:99px;padding:1px 9px;
          font-size:.72rem">0 pending</span>
        <a href="/admin/tshirt_report" style="float:right">
          <button class="btn btn-xs" style="background:var(--cyan);color:#000">📥 Export All</button></a>
      </div>
      <div style="display:flex;gap:6px;margin-bottom:12px">
        <button class="btn btn-sm" id="tshirt-tab-pending"
          onclick="showTshirtTab('pending')" style="background:var(--cyan);color:#000">
          ⏳ Pending <span id="tshirt-pending-count"></span></button>
        <button class="btn btn-sm btn-outline" id="tshirt-tab-approved"
          onclick="showTshirtTab('approved')">
          ✅ Approved <span id="tshirt-approved-count"></span></button>
        <button class="btn btn-sm btn-outline" id="tshirt-tab-rejected"
          onclick="showTshirtTab('rejected')">
          ❌ Rejected <span id="tshirt-rejected-count"></span></button>
      </div>
      <div id="tshirt-pending-table"></div>
      <div id="tshirt-approved-table" style="display:none"></div>
      <div id="tshirt-rejected-table" style="display:none"></div>
    </div>

    <!-- ── FLAGS ── -->
    <div style="margin-top:20px">
      <div class="shdr" style="color:var(--red)">
        🚩 Flagged Servers
        <a href="/admin/flags_report" style="float:right">
          <button class="btn btn-xs btn-red">📥 Export</button></a>
      </div>
      <div id="admin-flags"></div>
    </div>

    <!-- ── TAGS ── -->
    <div style="margin-top:20px">
      <div class="shdr" style="color:var(--purple)">
        🏷️ Server Tags
        <a href="/admin/tags_report" style="float:right">
          <button class="btn btn-xs btn-purple">📥 Export</button></a>
      </div>
      <div id="admin-tags-table"></div>
    </div>

    <!-- ── NOTES ── -->
    <div style="margin-top:20px">
      <div class="shdr" style="color:var(--blue)">
        📝 Server Notes
        <a href="/admin/notes_report" style="float:right">
          <button class="btn btn-xs">📥 Export</button></a>
      </div>
      <div id="admin-notes-table"></div>
    </div>

  </div>
</div>

<!-- MANAGE USERS -->
<div class="page" id="page-users">
  <div class="shdr">👥 Manage Users</div>
  <div style="max-width:640px">
    <div class="shdr" style="color:var(--green)">➕ Create New User</div>
    <div class="input-row">
      <input type="text" id="nu-uname" placeholder="Username">
      <input type="text" id="nu-fname" placeholder="Full Name">
    </div>
    <div class="input-row">
      <input type="password" id="nu-pw" placeholder="Password (min 6 chars)">
      <select id="nu-role">
        <option value="user">User (can tag/note/flag)</option>
        <option value="admin">Admin (full access)</option>
      </select>
    </div>
    <button class="btn btn-green btn-sm" onclick="createUser()" style="margin-bottom:20px">
      ➕ Create User
    </button>
    <div id="user-create-msg" style="margin-bottom:12px"></div>

    <div class="shdr">Existing Users</div>
    <div id="users-list"></div>
  </div>
</div>

</div><!-- /main -->

<!-- ── Detail Modal ── -->
<div class="modal-overlay" id="modal-overlay" onclick="closeModalOuter(event)">
  <div class="modal">
    <div class="modal-title">
      <span id="modal-hostname" style="font-family:monospace">Server Detail</span>
      <span class="modal-close" onclick="closeModal()">✕</span>
    </div>
    <div class="detail-grid" id="modal-body"></div>
    <!-- Tags section -->
    <div style="margin-top:14px">
      <div style="font-size:.75rem;font-weight:700;text-transform:uppercase;
           letter-spacing:.1em;color:var(--purple);margin-bottom:8px">🏷️ Tags</div>
      <div id="modal-tags"></div>
      <div id="modal-tag-editor" style="display:none;margin-top:8px">
        <div style="font-size:.75rem;color:var(--muted);margin-bottom:5px">Preset tags:</div>
        <div id="modal-preset-tags" style="display:flex;flex-wrap:wrap;gap:5px;margin-bottom:8px"></div>
        <div style="display:flex;gap:6px">
          <input type="text" id="custom-tag-input" placeholder="Or type a custom tag…"
            style="flex:1;background:var(--surf2);border:1px solid var(--border);
            color:var(--text);padding:6px 10px;border-radius:6px;font-size:.82rem">
          <button class="btn btn-purple btn-sm" onclick="addCustomTag()">Add</button>
        </div>
      </div>
    </div>
    <!-- Flag section -->
    <div style="margin-top:14px" id="modal-flag-section"></div>
    <!-- Notes section -->
    <div style="margin-top:14px">
      <div style="font-size:.75rem;font-weight:700;text-transform:uppercase;
           letter-spacing:.1em;color:var(--blue);margin-bottom:8px">📝 Notes</div>
      <div class="notes-list" id="modal-notes"></div>
      <div id="modal-note-editor" style="display:none">
        <div class="note-input-row">
          <textarea id="note-text-input" placeholder="Add a note…"></textarea>
          <button class="btn btn-sm" onclick="submitNote()" style="align-self:flex-end">Post</button>
        </div>
      </div>
    </div>
    <!-- Correction suggestion section -->
    <div style="margin-top:16px" id="modal-correction-section"></div>
  </div>
</div>

<!-- ── Add User Password Reset Modal ── -->
<div class="modal-overlay" id="pw-modal-overlay" onclick="closePwModal(event)">
  <div class="modal" style="max-width:380px">
    <div class="modal-title">
      🔑 Reset Password — <span id="pw-modal-uname"></span>
      <span class="modal-close" onclick="document.getElementById('pw-modal-overlay').classList.remove('show')">✕</span>
    </div>
    <div class="input-row">
      <input type="password" id="pw-new" placeholder="New password (min 6 chars)">
    </div>
    <button class="btn btn-sm" onclick="submitPwReset()">Update Password</button>
    <div id="pw-reset-msg" style="margin-top:8px"></div>
  </div>
</div>

<!-- ── Approve / Reject Decision Modal (shared by Corrections / Decommissions / New Servers) ── -->
<div class="modal-overlay" id="decision-modal-overlay" onclick="closeDecisionModal(event)">
  <div class="modal" style="max-width:440px">
    <div class="modal-title">
      <span id="decision-modal-title">Review Item</span>
      <span class="modal-close" onclick="document.getElementById('decision-modal-overlay').classList.remove('show')">✕</span>
    </div>
    <div style="display:flex;gap:8px;margin-bottom:14px">
      <button class="btn" id="decision-btn-approve" onclick="setDecisionChoice('Approved')"
        style="flex:1;background:var(--green);color:#000">✅ Approve</button>
      <button class="btn" id="decision-btn-reject" onclick="setDecisionChoice('Rejected')"
        style="flex:1;background:var(--red)">❌ Reject</button>
    </div>
    <div style="font-size:.74rem;color:var(--muted);margin-bottom:4px">
      Note (optional) — visible in the monthly changes report</div>
    <input type="text" id="decision-reason-input" placeholder="e.g. Confirmed with infra team, applying next upload"
      style="width:100%;background:var(--surf2);border:1px solid var(--border);color:var(--text);
      padding:8px 11px;border-radius:8px;font-size:.84rem;margin-bottom:12px">
    <button class="btn btn-sm" id="decision-submit-btn" onclick="submitDecisionModal()" disabled
      style="opacity:.5">Select Approve or Reject first</button>
    <div id="decision-modal-msg" style="margin-top:8px"></div>
  </div>
</div>

<script>
// ════════════════════════════════════════════════════════════════
//  STATE
// ════════════════════════════════════════════════════════════════
let ME = {logged_in: false, username:'', full_name:'', role:''};
let PRESET_TAGS = [];
// Fallback — always populated from /detail response, but defined here as safety net
const CORRECTABLE_COLS = [
  "Server Type(Physical/Virtual)", "Platform", "Server DC Location",
  "HPC or NON HPC or JPC", "Server Role", "Final OS",
  "Commercial Category", "Application Name", "Reference"
];
let _modalHostname = '';
let _bulkRows = [];
let _debTimer;
let _filterLoaded = false;
let _currPage = 1;
let _scol = '';

// ════════════════════════════════════════════════════════════════
//  INIT
// ════════════════════════════════════════════════════════════════
async function init(){
  if(localStorage.getItem('theme')==='light') document.body.classList.add('light');
  const r = await fetch('/me').then(x=>x.json());
  ME = r;
  if(r.decom_reasons)     _decomReasons      = r.decom_reasons;
  if(r.newsrv_fields)     _newsrvFields      = r.newsrv_fields;
  if(r.preset_tags)       PRESET_TAGS        = r.preset_tags;
  if(r.tshirt_platforms)  _tshirtPlatforms   = r.tshirt_platforms;
  renderUserArea();
  renderSidebar();
  loadDashboard();
}

function renderUserArea(){
  const el = document.getElementById('user-area');
  if(ME.logged_in){
    const init = (ME.full_name||ME.username||'?')[0].toUpperCase();
    el.innerHTML = `
      <div class="user-pill" onclick="window.location='/logout'">
        <div class="user-avatar">${init}</div>
        <span>${ME.full_name||ME.username}</span>
        <span style="color:var(--muted);font-size:.7rem">[${ME.role}]</span>
        <span style="color:var(--muted);font-size:.75rem">↩ Logout</span>
      </div>`;
  } else {
    el.innerHTML = `<a href="/login"><button class="btn btn-sm">🔐 Login</button></a>`;
  }
}

function renderSidebar(){
  const show = (id, vis) => {
    const el = document.getElementById(id);
    if(el) el.style.display = vis ? '' : 'none';
  };
  if(ME.logged_in){
    show('nav-user-section', true);
    show('nav-activity', true);
    show('nav-report-section', true);
    show('nav-decom', true);
    show('nav-newsrv', true);
    show('nav-tshirt', true);
  }
  if(ME.role === 'admin'){
    show('nav-admin-section', true);
    show('nav-upload', true);
    show('nav-compare', true);
    show('nav-admin', true);
    show('nav-users', true);
    const totalPending = (ME.pending_corrections||0) + (ME.pending_decommissions||0) + (ME.pending_newservers||0);
    if(totalPending > 0){
      const badge = document.getElementById('corr-badge');
      if(badge){ badge.style.display=''; badge.textContent=totalPending; }
    }
  }
}

// ════════════════════════════════════════════════════════════════
//  THEME + ROUTING
// ════════════════════════════════════════════════════════════════
function toggleTheme(){
  document.body.classList.toggle('light');
  localStorage.setItem('theme', document.body.classList.contains('light')?'light':'dark');
}

function showPage(name, el){
  const adminOnlyPages = ['compare','upload','admin','users'];
  if(adminOnlyPages.includes(name) && ME.role !== 'admin'){
    name = 'dashboard';
    el = document.querySelector('.nav-item[onclick*="dashboard"]');
  }
  document.querySelectorAll('.page').forEach(p=>p.classList.remove('active'));
  document.querySelectorAll('.nav-item').forEach(n=>n.classList.remove('active'));
  document.getElementById('page-'+name).classList.add('active');
  if(el) el.classList.add('active');
  if(name==='dashboard')  loadDashboard();
  if(name==='search')    { loadFilterOptions(); runSearch(1); }
  if(name==='grouped')   loadGroupedSummary();
  if(name==='compare')   { document.getElementById('curr-file-lbl').textContent = window._currFile||'None'; }
  if(name==='admin')     loadAdmin();
  if(name==='users')     loadUsers();
  if(name==='activity')  loadActivity();
  if(name==='decom')     loadDecomPage();
  if(name==='newsrv')    loadNewsrvPage();
  if(name==='tshirt')    loadTshirtPage();
}

// ════════════════════════════════════════════════════════════════
//  DASHBOARD
// ════════════════════════════════════════════════════════════════
function loadDashboard(){
  document.getElementById('dash-loading').style.display='';
  document.getElementById('dash-content').style.display='none';
  fetch('/dashboard').then(r=>r.json()).then(d=>{
    if(d.error){
      document.getElementById('dash-loading').innerHTML=
        `<div class="alert alert-info">${esc(d.error)}</div>`; return;
    }
    document.getElementById('dash-loading').style.display='none';
    document.getElementById('dash-content').style.display='';
    const m=d.metrics;
    document.getElementById('topbar-meta').textContent =
      d.filename ? `${d.filename} · ${m.total.toLocaleString()} servers · ${d.uploaded_at}` : 'No data loaded';
    document.getElementById('metrics-row').innerHTML = `
      <div class="metric-card"><div class="metric-val c-blue">${m.total.toLocaleString()}</div>
        <div class="metric-lbl">Total Servers</div></div>
      <div class="metric-card"><div class="metric-val c-green">${m.live.toLocaleString()}</div>
        <div class="metric-lbl">Live (BAU)</div></div>
      <div class="metric-card"><div class="metric-val c-red">${m.not_live.toLocaleString()}</div>
        <div class="metric-lbl">Not Live (Project Phase)</div></div>
      <div class="metric-card"><div class="metric-val c-cyan">${m.physical.toLocaleString()}</div>
        <div class="metric-lbl">Physical</div></div>
      <div class="metric-card"><div class="metric-val c-purple">${m.virtual.toLocaleString()}</div>
        <div class="metric-lbl">Virtual</div></div>`;
    const cg = document.getElementById('charts-grid');
    cg.innerHTML='';
    ['status_donut','type_donut','platform_hbar','loc_hbar',
     'hpc_hbar','role_vbar','os_vbar','app_hbar'].forEach(k=>{
      if(d.charts[k]) cg.innerHTML+=
        `<div class="chart-card"><img src="${d.charts[k]}" alt="${k}"></div>`;
    });
    if(d.quality && Object.keys(d.quality).length){
      const items = Object.entries(d.quality)
        .map(([k,v])=>`<div style="font-size:.8rem;color:var(--muted)">⚠️ <b>${esc(k)}</b>: <span style="color:var(--yellow)">${v}</span> blank rows</div>`)
        .join('');
      cg.innerHTML += `<div class="chart-card" style="grid-column:1/-1">
        <div class="quality-panel"><h4>⚠️ Data Quality Warnings</h4>${items}</div></div>`;
    }
  });
}

// ════════════════════════════════════════════════════════════════
//  SEARCH
// ════════════════════════════════════════════════════════════════
function loadFilterOptions(){
  if(_filterLoaded) return;
  fetch('/filter_options').then(r=>r.json()).then(d=>{
    fillSel('f-platform', d.platform||[]);
    fillSel('f-status',   d.status  ||[]);
    fillSel('f-role',     d.role    ||[]);
    fillSel('f-hpc',      d.hpc     ||[]);
    fillSel('f-loc',      d.loc     ||[]);
    fillSel('f-os',       d.os      ||[]);
    fillSel('f-stype',    d.stype   ||[]);
    _filterLoaded = true;
  });
}
function fillSel(id, opts){
  const el=document.getElementById(id);
  el.innerHTML='<option value="">All</option>'+
    opts.map(o=>`<option value="${esc(o)}">${esc(o)}</option>`).join('');
}

function debSearch(){ clearTimeout(_debTimer); _debTimer=setTimeout(()=>runSearch(1),350); }

let _sortBy = '';
let _sortDir = 'asc';

function runSearch(page){
  _currPage=page||1;
  const params=new URLSearchParams({
    q:        document.getElementById('f-q').value,
    platform: document.getElementById('f-platform').value,
    status:   document.getElementById('f-status').value,
    role:     document.getElementById('f-role').value,
    hpc:      document.getElementById('f-hpc').value,
    loc:      document.getElementById('f-loc').value,
    os:       document.getElementById('f-os').value,
    stype:    document.getElementById('f-stype').value,
    sort_by:  _sortBy,
    sort_dir: _sortDir,
    page:     _currPage,
  });
  document.getElementById('search-tbody').innerHTML=
    '<tr><td colspan="14" class="loading-row"><div class="spinner"></div></td></tr>';
  fetch('/search?'+params).then(r=>r.json()).then(d=>{
    if(d.status_col) _scol=d.status_col;
    document.getElementById('search-count').textContent=
      `Showing ${d.rows.length} of ${d.total.toLocaleString()} servers (page ${d.page}/${d.pages})`;
    renderSearchTable(d.rows, d.status_col||_scol);
    renderPagination(d.page, d.pages);
  });
}

function toggleSort(col){
  if(_sortBy === col){
    _sortDir = (_sortDir === 'asc') ? 'desc' : 'asc';
  } else {
    _sortBy = col;
    _sortDir = 'asc';
  }
  runSearch(1);
}

function renderSearchTable(rows, scol){
  if(!rows.length){
    document.getElementById('search-thead').innerHTML='';
    document.getElementById('search-tbody').innerHTML=
      '<tr><td colspan="14" class="loading-row">No results.</td></tr>'; return;
  }
  // columns to show (exclude internal _ keys)
  const cols=Object.keys(rows[0]).filter(c=>!c.startsWith('_'));
  document.getElementById('search-thead').innerHTML=
    cols.map(c=>{
      const isActive = (_sortBy === c);
      const arrow = isActive ? (_sortDir==='asc' ? '▲' : '▼') : '↕';
      return `<th class="sortable ${isActive?'sort-active':''}" onclick="toggleSort('${esc(c)}')"
        title="Click to sort">${esc(c)} <span class="sort-arrow">${arrow}</span></th>`;
    }).join('')+'<th>Tags / Notes</th>';
  document.getElementById('search-tbody').innerHTML=rows.map(r=>`
    <tr style="cursor:pointer" onclick="showDetail('${esc(r['Server HostName']||'')}')">
      ${cols.map(c=>{
        const v=r[c]||'';
        if(c==='Server HostName')
          return `<td style="font-weight:700;color:var(--blue);font-family:monospace">${esc(v)}</td>`;
        if(c===scol) return `<td>${statusBadge(v)}</td>`;
        if(c==='Server Type(Physical/Virtual)') return `<td>${typeBadge(v)}</td>`;
        return `<td>${esc(v)}</td>`;
      }).join('')}
      <td>${renderMiniTags(r._tags, r._notes, r._flagged)}</td>
    </tr>`).join('');
}

function renderMiniTags(tags, noteCount, flagged){
  let h='';
  if(flagged) h+=`<span class="badge" style="background:rgba(255,79,106,.15);color:var(--red);border:1px solid rgba(255,79,106,.3);margin:1px">🚩</span>`;
  (tags||[]).forEach(t=>{
    const label = t.tag || t;  // supports both rich objects and legacy strings
    const isAccepted = t.status === 'Accepted';
    const dot = t.status ? (isAccepted ? '✓ ' : '? ') : '';
    h+=`<span class="tag-pill" style="font-size:.65rem;padding:1px 6px;
      ${isAccepted?'':'opacity:.75;border-style:dashed'}" title="${isAccepted?'Accepted':'Unverified — pending admin review'}">${dot}${esc(label)}</span>`;
  });
  if(noteCount) h+=`<span class="badge badge-other" style="margin:1px">📝${noteCount}</span>`;
  return h||'<span style="color:var(--dim);font-size:.72rem">—</span>';
}

function renderPagination(page, pages){
  const pg=document.getElementById('search-pagination');
  if(pages<=1){pg.innerHTML='';return;}
  let h='';
  if(page>1) h+=`<button class="page-btn" onclick="runSearch(${page-1})">‹</button>`;
  const s=Math.max(1,page-2),e=Math.min(pages,page+2);
  if(s>1) h+=`<button class="page-btn" onclick="runSearch(1)">1</button>${s>2?'…':''}`;
  for(let i=s;i<=e;i++) h+=`<button class="page-btn ${i===page?'active':''}" onclick="runSearch(${i})">${i}</button>`;
  if(e<pages) h+=`${e<pages-1?'…':''}<button class="page-btn" onclick="runSearch(${pages})">${pages}</button>`;
  if(page<pages) h+=`<button class="page-btn" onclick="runSearch(${page+1})">›</button>`;
  pg.innerHTML=h;
}

function qf(field, val, el){
  document.querySelectorAll('.chip-row .chip').forEach(c=>c.classList.remove('active'));
  const map={platform:'f-platform',hpc:'f-hpc',role:'f-role',status:'f-status'};
  const sel=document.getElementById(map[field]);
  if(!sel) return;
  if(sel.value===val){ sel.value=''; } else { sel.value=val; el.classList.add('active'); }
  runSearch(1);
}
function clearFilters(){
  document.getElementById('f-q').value='';
  ['f-platform','f-status','f-role','f-hpc','f-loc','f-os','f-stype']
    .forEach(id=>document.getElementById(id).value='');
  document.querySelectorAll('.chip-row .chip').forEach(c=>c.classList.remove('active'));
  _sortBy=''; _sortDir='asc';
  runSearch(1);
}
function exportSearch(){
  const p=new URLSearchParams({
    q:       document.getElementById('f-q').value,
    platform:document.getElementById('f-platform').value,
    status:  document.getElementById('f-status').value,
    role:    document.getElementById('f-role').value,
  });
  window.location='/export?'+p;
}

// ════════════════════════════════════════════════════════════════
//  GROUPED SUMMARY
// ════════════════════════════════════════════════════════════════
let _groupBy = 'platform';
let _groupDrillValue = '';
let _groupDrillPage = 1;

function loadGroupedSummary(){
  document.getElementById('grouped-summary-loading').style.display='';
  document.getElementById('grouped-cards').style.display='none';
  document.getElementById('grouped-drill-panel').style.display='none';
  fetch('/grouped?by='+encodeURIComponent(_groupBy)).then(r=>r.json()).then(d=>{
    document.getElementById('grouped-summary-loading').style.display='none';
    if(d.error){
      document.getElementById('grouped-cards').innerHTML =
        `<div class="alert alert-info">${esc(d.error)}</div>`;
      document.getElementById('grouped-cards').style.display='';
      return;
    }
    if(!d.groups || !d.groups.length){
      document.getElementById('grouped-cards').innerHTML =
        '<div class="alert alert-info">No data loaded yet — upload an inventory file first.</div>';
      document.getElementById('grouped-cards').style.display='';
      return;
    }
    renderGroupCards(d.groups, d.total, !!d.status_col);
    document.getElementById('grouped-cards').style.display='';
  });
}

function setGroupBy(by, el){
  _groupBy = by;
  document.querySelectorAll('#grouped-by-chips .chip').forEach(c=>c.classList.remove('active'));
  el.classList.add('active');
  closeGroupedDrill();
  loadGroupedSummary();
}

function renderGroupCards(groups, total, hasStatus){
  const el = document.getElementById('grouped-cards');
  const header = `<div class="tbl-count" style="margin-bottom:12px">
    ${groups.length} groups · ${total.toLocaleString()} total servers</div>`;
  const cards = groups.map(g => {
    const livePct    = g.count ? (g.live    / g.count * 100) : 0;
    const notLivePct = g.count ? (g.not_live / g.count * 100) : 0;
    const barHtml = hasStatus ? `
      <div class="group-card-bar">
        ${g.live    ? `<div class="seg-live"    style="width:${livePct}%"></div>`    : ''}
        ${g.not_live? `<div class="seg-notlive" style="width:${notLivePct}%"></div>` : ''}
      </div>
      <div class="group-card-legend">
        <span style="color:var(--green)">● ${g.live} Live</span>
        <span style="color:var(--red)">● ${g.not_live} Not Live</span>
      </div>` : '';
    return `<div class="group-card" onclick="drillGroup('${esc(g.value).replace(/'/g,"\\'")}')">
      <div class="group-card-name" title="${esc(g.value)}">${esc(g.value)}</div>
      <div class="group-card-count">${g.count.toLocaleString()}</div>
      ${barHtml}
    </div>`;
  }).join('');
  el.innerHTML = header + `<div class="group-cards-grid">${cards}</div>`;
}

function drillGroup(value){
  _groupDrillValue = value;
  _groupDrillPage = 1;
  document.getElementById('grouped-drill-panel').style.display = '';
  document.getElementById('grouped-drill-title').textContent = `📂 ${value}`;
  document.getElementById('grouped-drill-export').onclick = () => {
    window.location = `/export_group?by=${encodeURIComponent(_groupBy)}&value=${encodeURIComponent(value)}`;
  };
  runGroupDrill(1);
  document.getElementById('grouped-drill-panel').scrollIntoView({behavior:'smooth', block:'nearest'});
}

function closeGroupedDrill(){
  document.getElementById('grouped-drill-panel').style.display = 'none';
  _groupDrillValue = '';
}

function runGroupDrill(page){
  _groupDrillPage = page || 1;
  const params = new URLSearchParams({
    by: _groupBy, value: _groupDrillValue, page: _groupDrillPage
  });
  document.getElementById('grouped-drill-tbody').innerHTML =
    '<tr><td colspan="12" class="loading-row"><div class="spinner"></div></td></tr>';
  fetch('/grouped/drill?'+params).then(r=>r.json()).then(d=>{
    document.getElementById('grouped-drill-count').textContent =
      `Showing ${d.rows.length} of ${d.total.toLocaleString()} servers (page ${d.page}/${d.pages})`;
    renderGroupDrillTable(d.rows, d.status_col);
    renderGroupDrillPagination(d.page, d.pages);
  });
}

function renderGroupDrillTable(rows, scol){
  if(!rows.length){
    document.getElementById('grouped-drill-thead').innerHTML = '';
    document.getElementById('grouped-drill-tbody').innerHTML =
      '<tr><td colspan="12" class="loading-row">No servers found.</td></tr>';
    return;
  }
  const cols = Object.keys(rows[0]).filter(c=>!c.startsWith('_'));
  document.getElementById('grouped-drill-thead').innerHTML =
    cols.map(c=>`<th>${esc(c)}</th>`).join('') + '<th>Tags</th>';
  document.getElementById('grouped-drill-tbody').innerHTML = rows.map(r=>`
    <tr style="cursor:pointer" onclick="showDetail('${esc(r['Server HostName']||'')}')">
      ${cols.map(c=>{
        const v=r[c]||'';
        if(c==='Server HostName')
          return `<td style="font-weight:700;color:var(--blue);font-family:monospace">${esc(v)}</td>`;
        if(c===scol) return `<td>${statusBadge(v)}</td>`;
        if(c==='Server Type(Physical/Virtual)') return `<td>${typeBadge(v)}</td>`;
        return `<td>${esc(v)}</td>`;
      }).join('')}
      <td>${renderMiniTags(r._tags, 0, r._flagged)}</td>
    </tr>`).join('');
}

function renderGroupDrillPagination(page, pages){
  const pg = document.getElementById('grouped-drill-pagination');
  if(pages<=1){ pg.innerHTML=''; return; }
  let h='';
  if(page>1) h+=`<button class="page-btn" onclick="runGroupDrill(${page-1})">‹</button>`;
  const s=Math.max(1,page-2), e=Math.min(pages,page+2);
  if(s>1) h+=`<button class="page-btn" onclick="runGroupDrill(1)">1</button>${s>2?'…':''}`;
  for(let i=s;i<=e;i++) h+=`<button class="page-btn ${i===page?'active':''}" onclick="runGroupDrill(${i})">${i}</button>`;
  if(e<pages) h+=`${e<pages-1?'…':''}<button class="page-btn" onclick="runGroupDrill(${pages})">${pages}</button>`;
  if(page<pages) h+=`<button class="page-btn" onclick="runGroupDrill(${page+1})">›</button>`;
  pg.innerHTML = h;
}

// ════════════════════════════════════════════════════════════════
//  SERVER DETAIL MODAL
// ════════════════════════════════════════════════════════════════
function showDetail(hostname){
  if(!hostname) return;
  _modalHostname = hostname;
  fetch(`/detail?hostname=${encodeURIComponent(hostname)}`).then(r=>r.json()).then(d=>{
    if(!d||!Object.keys(d).length) return;
    document.getElementById('modal-hostname').textContent = d['Server HostName']||hostname;
    // Detail grid
    const scol=d['_status_col']||'', skip=new Set(['_status_col','_status_val','_tags','_notes','_flag','_correctable_cols']);
    const full=['Application Name','Reference','Commercial Category'];
    let html='';
    Object.entries(d).forEach(([k,v])=>{
      if(skip.has(k)||!v||String(v).trim()==='') return;
      let vHtml=esc(v);
      if(k===scol) vHtml=statusBadge(v);
      if(k==='Server Type(Physical/Virtual)') vHtml=typeBadge(v);
      html+=`<div class="detail-row ${full.includes(k)?'detail-full':''}">
        <div class="detail-key">${esc(k)}</div>
        <div class="detail-val">${vHtml}</div></div>`;
    });
    document.getElementById('modal-body').innerHTML=html;
    renderModalTags(d._tags||[], d);
    renderModalNotes(d._notes||[], d);
    renderModalFlag(d._flag, d);
    renderModalCorrection(d);
    document.getElementById('modal-overlay').classList.add('show');
  });
}

function renderModalTags(tags, d){
  const el=document.getElementById('modal-tags');
  const editor=document.getElementById('modal-tag-editor');
  const presetEl=document.getElementById('modal-preset-tags');
  // Show existing tags — each as a rich object {id, tag, user, name, ts, status}
  if(tags.length){
    el.innerHTML=tags.map(t=>{
      const isAccepted = t.status === 'Accepted';
      const canRemove  = ME.logged_in && (ME.role==='admin' || ME.username===t.user);
      const canAccept  = ME.role==='admin' && !isAccepted;
      const rm = canRemove ? `<span class="rm" onclick="removeTagById('${t.id}')">✕</span>` : '';
      const accept = canAccept ? `<span class="rm" style="color:var(--green)"
        title="Accept as official" onclick="acceptTagById('${t.id}')">✓</span>` : '';
      const statusDot = isAccepted
        ? `<span style="color:var(--green);font-size:.65rem" title="Accepted by ${esc(t.accepted_by||'')}">✓</span>`
        : `<span style="color:var(--yellow);font-size:.65rem" title="Unverified — submitted by ${esc(t.name||t.user||'')}">?</span>`;
      return `<span class="tag-pill" style="${isAccepted?'':'border-style:dashed;opacity:.85'}">
        ${statusDot} ${esc(t.tag)}${accept}${rm}</span>`;
    }).join('');
  } else {
    el.innerHTML=`<span style="color:var(--dim);font-size:.8rem">No tags yet</span>`;
  }
  // Show editor only if logged in
  if(ME.logged_in){
    editor.style.display='';
    const usedSet=new Set(tags.map(t=>t.tag));
    presetEl.innerHTML=PRESET_TAGS.map(t=>{
      const used=usedSet.has(t);
      return `<span class="chip ${used?'active':''}" style="font-size:.72rem"
        onclick="togglePresetTag('${esc(t)}','${used}')">${esc(t)}</span>`;
    }).join('');
  } else {
    editor.style.display='none';
    if(!tags.length) el.innerHTML=loginNotice('tag this server');
  }
}

function renderModalNotes(notes, d){
  const el=document.getElementById('modal-notes');
  const editor=document.getElementById('modal-note-editor');
  if(notes.length){
    el.innerHTML=notes.map(n=>{
      const canDel=ME.logged_in&&(ME.role==='admin'||ME.username===n.user);
      const del=canDel?`<span class="note-del" onclick="deleteNote('${n.id}')">🗑</span>`:'';
      return `<div class="note-item">${del}
        <div class="note-meta">✍️ ${esc(n.name)} · ${esc(n.ts)}</div>
        <div class="note-text">${esc(n.note)}</div></div>`;
    }).join('');
  } else {
    el.innerHTML=`<div style="color:var(--dim);font-size:.8rem;padding:8px 0">No notes yet.</div>`;
  }
  if(ME.logged_in){
    editor.style.display='';
  } else {
    editor.style.display='none';
    if(!notes.length) el.innerHTML=loginNotice('add notes');
  }
}

function renderModalFlag(flag, d){
  const el=document.getElementById('modal-flag-section');
  if(flag){
    const canUnflag=ME.logged_in;
    el.innerHTML=`<div class="flag-box">
      <h4>🚩 Flagged for Review</h4>
      <div style="font-size:.8rem;color:var(--muted)">
        Reason: <b style="color:var(--text)">${esc(flag.reason||'No reason given')}</b><br>
        By: ${esc(flag.user)} · ${esc(flag.ts)}</div>
      ${canUnflag?`<button class="btn btn-sm btn-outline" style="margin-top:8px"
        onclick="unflagServer()">✓ Remove Flag</button>`:''}
    </div>`;
  } else if(ME.logged_in){
    el.innerHTML=`<button class="btn btn-sm btn-outline" style="border-color:var(--orange);color:var(--orange)"
      onclick="flagServer()">🚩 Flag for Review</button>`;
  } else {
    el.innerHTML='';
  }
}

function renderModalCorrection(serverData){
  const el=document.getElementById('modal-correction-section');
  if(!ME.logged_in){
    el.innerHTML=`<div class="login-notice" style="margin-top:4px">🔐 <a href="/login">Log in</a> to suggest a correction</div>`;
    return;
  }
  // Use cols from server response — always available, no dependency on admin panel
  const corrCols = (serverData._correctable_cols || CORRECTABLE_COLS || [])
    .filter(c => serverData[c] !== undefined && !c.startsWith('_'));
  if(!corrCols.length){
    el.innerHTML=`<div style="color:var(--dim);font-size:.78rem;margin-top:8px">No correctable columns found for this server.</div>`;
    return;
  }
  const colOpts = corrCols.map(c =>
    `<option value="${esc(c)}">${esc(c)}</option>`).join('');
  el.innerHTML=`
    <div style="border-top:1px solid var(--border);padding-top:14px">
      <div style="font-size:.75rem;font-weight:700;text-transform:uppercase;
           letter-spacing:.1em;color:var(--orange);margin-bottom:10px">✏️ Suggest a Correction</div>
      <div id="corr-form-area">
        <div style="display:flex;gap:8px;flex-wrap:wrap;margin-bottom:8px">
          <div style="flex:1;min-width:160px">
            <div style="font-size:.7rem;color:var(--muted);margin-bottom:3px">Column with wrong value</div>
            <select id="corr-col-sel" onchange="onCorrColChange()"
              style="width:100%;background:var(--surf2);border:1px solid var(--border);
              color:var(--text);padding:7px 10px;border-radius:7px;font-size:.82rem">
              <option value="">— Select column —</option>${colOpts}
            </select>
          </div>
          <div style="flex:1;min-width:160px" id="corr-current-box">
            <div style="font-size:.7rem;color:var(--muted);margin-bottom:3px">Current value</div>
            <div id="corr-current-val"
              style="background:rgba(255,79,106,.08);border:1px solid rgba(255,79,106,.2);
              color:var(--red);padding:7px 10px;border-radius:7px;font-size:.82rem;
              font-weight:600;min-height:34px;color:var(--dim)">← select a column first</div>
          </div>
        </div>
        <div id="corr-bottom" style="display:none">
          <div style="font-size:.7rem;color:var(--muted);margin-bottom:3px">Correct value should be</div>
          <input type="text" id="corr-suggested"
            style="width:100%;background:var(--surf2);border:1px solid var(--border);
            color:var(--text);padding:7px 10px;border-radius:7px;font-size:.82rem;margin-bottom:8px"
            placeholder="Type the correct value…">
          <div style="font-size:.7rem;color:var(--muted);margin-bottom:3px">Reason / evidence (optional)</div>
          <input type="text" id="corr-reason"
            style="width:100%;background:var(--surf2);border:1px solid var(--border);
            color:var(--text);padding:7px 10px;border-radius:7px;font-size:.82rem;margin-bottom:10px"
            placeholder="e.g. Confirmed with server team on 28 Jun">
          <button class="btn btn-sm btn-orange" onclick="submitCorrection()">✏️ Submit Correction</button>
        </div>
      </div>
      <div id="corr-success" style="display:none" class="alert alert-success">
        ✅ Correction submitted! Admin will review it.</div>
    </div>`;
  // Store server data for column value lookup
  window._corrServerData = serverData;
}

function onCorrColChange(){
  const col    = document.getElementById('corr-col-sel').value;
  const bottom = document.getElementById('corr-bottom');
  const curVal = document.getElementById('corr-current-val');
  if(!col){
    curVal.textContent = '← select a column first';
    curVal.style.color = 'var(--dim)';
    bottom.style.display = 'none';
    return;
  }
  const val = (window._corrServerData||{})[col];
  const display = (val && String(val).trim()) ? String(val).trim() : '(blank / missing)';
  curVal.textContent = display;
  curVal.style.color = (val && String(val).trim()) ? 'var(--red)' : 'var(--muted)';
  bottom.style.display = '';
  // Pre-fill suggested with current so user just edits it
  const sugInput = document.getElementById('corr-suggested');
  if(sugInput && val && String(val).trim()) sugInput.value = String(val).trim();
}

async function submitCorrection(){
  const col = document.getElementById('corr-col-sel').value;
  const sug = document.getElementById('corr-suggested').value.trim();
  const rsn = document.getElementById('corr-reason').value.trim();
  if(!col||!sug){ alert('Please select a column and provide a suggested value.'); return; }
  const cur = (window._corrServerData||{})[col]||'';
  const r = await fetch('/correction/add',{method:'POST',
    headers:{'Content-Type':'application/json'},
    body:JSON.stringify({hostname:_modalHostname,column:col,
      current_val:cur,suggested_val:sug,reason:rsn})}).then(x=>x.json());
  if(r.error){ alert(r.error); return; }
  document.getElementById('corr-form-area').style.display='none';
  document.getElementById('corr-success').style.display='';
}

function loginNotice(action){
  return `<div class="login-notice">🔐 <a href="/login">Log in</a> to ${action}</div>`;
}

function closeModal(){ document.getElementById('modal-overlay').classList.remove('show'); }
function closeModalOuter(e){
  if(e.target===document.getElementById('modal-overlay')) closeModal();
}

// ── Tag actions ──
async function togglePresetTag(tag, isUsed){
  if(isUsed==='true'){
    // find the tag's id from current modal state to remove it
    const r = await fetch(`/detail?hostname=${encodeURIComponent(_modalHostname)}`).then(x=>x.json());
    const existing = (r._tags||[]).find(t=>t.tag===tag);
    if(existing) await removeTagById(existing.id);
  } else {
    await addTagToServer(tag);
  }
}
async function addCustomTag(){
  const inp=document.getElementById('custom-tag-input');
  const tag=inp.value.trim();
  if(!tag) return;
  await addTagToServer(tag);
  inp.value='';
}
async function addTagToServer(tag){
  const r=await fetch('/tag/add',{method:'POST',
    headers:{'Content-Type':'application/json'},
    body:JSON.stringify({hostname:_modalHostname,tag})}).then(x=>x.json());
  if(r.error){ alert(r.error); return; }
  renderModalTags(r.tags, {});
  refreshRowTags(_modalHostname, r.tags);
}
async function removeTagById(tagId){
  const r=await fetch('/tag/remove',{method:'POST',
    headers:{'Content-Type':'application/json'},
    body:JSON.stringify({hostname:_modalHostname,id:tagId})}).then(x=>x.json());
  if(r.error){ alert(r.error); return; }
  renderModalTags(r.tags, {});
  refreshRowTags(_modalHostname, r.tags);
}
async function acceptTagById(tagId){
  const r=await fetch('/admin/tag/accept',{method:'POST',
    headers:{'Content-Type':'application/json'},
    body:JSON.stringify({hostname:_modalHostname,id:tagId})}).then(x=>x.json());
  if(r.error){ alert(r.error); return; }
  renderModalTags(r.tags, {});
  refreshRowTags(_modalHostname, r.tags);
}

// ── Note actions ──
async function submitNote(){
  const txt=document.getElementById('note-text-input').value.trim();
  if(!txt) return;
  const r=await fetch('/note/add',{method:'POST',
    headers:{'Content-Type':'application/json'},
    body:JSON.stringify({hostname:_modalHostname,note:txt})}).then(x=>x.json());
  if(r.error){ alert(r.error); return; }
  document.getElementById('note-text-input').value='';
  renderModalNotes(r.notes, {});
}
async function deleteNote(id){
  if(!confirm('Delete this note?')) return;
  const r=await fetch('/note/delete',{method:'POST',
    headers:{'Content-Type':'application/json'},
    body:JSON.stringify({hostname:_modalHostname,id})}).then(x=>x.json());
  if(r.ok) renderModalNotes(r.notes, {});
}

// ── Flag actions ──
async function flagServer(){
  const reason=prompt('Reason for flagging (optional):') || '';
  const r=await fetch('/flag/add',{method:'POST',
    headers:{'Content-Type':'application/json'},
    body:JSON.stringify({hostname:_modalHostname,reason})}).then(x=>x.json());
  if(r.ok) showDetail(_modalHostname);
}
async function unflagServer(){
  const r=await fetch('/flag/remove',{method:'POST',
    headers:{'Content-Type':'application/json'},
    body:JSON.stringify({hostname:_modalHostname})}).then(x=>x.json());
  if(r.ok) showDetail(_modalHostname);
}

function refreshRowTags(hostname, tags){
  // Update tags cell in search table without full reload
  document.querySelectorAll('#search-tbody tr').forEach(tr=>{
    const hnCell=tr.querySelector('td');
    if(hnCell&&hnCell.textContent.trim().toLowerCase()===hostname.toLowerCase()){
      const lastCell=tr.querySelector('td:last-child');
      if(lastCell) lastCell.innerHTML=renderMiniTags(tags,0,false);
    }
  });
}

// ════════════════════════════════════════════════════════════════
//  BULK LOOKUP
// ════════════════════════════════════════════════════════════════
function runBulk(){
  const raw=document.getElementById('bulk-input').value;
  const names=raw.split('\n').map(s=>s.trim()).filter(Boolean);
  if(!names.length) return;
  document.getElementById('bulk-summary').innerHTML='<div class="spinner"></div> Looking up…';
  document.getElementById('bulk-results').innerHTML='';
  document.getElementById('bulk-not-found').innerHTML='';
  fetch('/bulk',{method:'POST',headers:{'Content-Type':'application/json'},
    body:JSON.stringify({names})}).then(r=>r.json()).then(d=>{
    _bulkRows=d.found||[];
    const sc=d.status_col||'';
    const fc=d.total_found, nc=(d.not_found||[]).length;
    document.getElementById('bulk-summary').innerHTML=`
      <div class="alert ${fc?'alert-success':'alert-warn'}">
        ✅ Found <b>${fc}</b> of <b>${names.length}</b> servers
        ${nc?`· ❌ <b>${nc}</b> not found`:''}
        ${fc?`<button class="btn btn-sm btn-green" style="margin-left:10px"
          onclick="exportBulk()">📥 Export</button>`:''}
      </div>`;
    if(d.not_found&&d.not_found.length){
      document.getElementById('bulk-not-found').innerHTML=
        `<div class="not-found-box"><h4>❌ Not found (${d.not_found.length})</h4>
        ${d.not_found.map(n=>`<div style="font-family:monospace;font-size:.78rem;color:var(--muted);padding:2px 0">· ${esc(n)}</div>`).join('')}</div>`;
    }
    if(_bulkRows.length){
      const cols=Object.keys(_bulkRows[0]).filter(c=>!c.startsWith('_'));
      document.getElementById('bulk-results').innerHTML=`
        <div class="tbl-wrap"><table>
          <thead><tr>${cols.map(c=>`<th>${esc(c)}</th>`).join('')}<th>Tags</th></tr></thead>
          <tbody>${_bulkRows.map(r=>`<tr style="cursor:pointer"
            onclick="showDetail('${esc(r['Server HostName']||'')}')">
            ${cols.map(c=>{
              const v=r[c]||'';
              if(c==='Server HostName') return `<td style="font-weight:700;color:var(--blue);font-family:monospace">${esc(v)}</td>`;
              if(c===sc) return `<td>${statusBadge(v)}</td>`;
              return `<td>${esc(v)}</td>`;
            }).join('')}
            <td>${renderMiniTags(r._tags,0,r._flagged)}</td>
          </tr>`).join('')}
          </tbody></table></div>`;
    }
  });
}
function exportBulk(){
  fetch('/export_bulk',{method:'POST',headers:{'Content-Type':'application/json'},
    body:JSON.stringify({rows:_bulkRows})}).then(r=>r.blob()).then(b=>{
    const a=document.createElement('a');a.href=URL.createObjectURL(b);
    a.download='bulk_lookup.xlsx';a.click();
  });
}

// ════════════════════════════════════════════════════════════════
//  MONTH COMPARE
// ════════════════════════════════════════════════════════════════
function uploadPrev(file){
  if(!file) return;
  const fd=new FormData();fd.append('file',file);fd.append('which','prev');
  document.getElementById('prev-status').innerHTML='<div class="spinner"></div>';
  fetch('/upload',{method:'POST',body:fd}).then(r=>r.json()).then(d=>{
    document.getElementById('prev-status').innerHTML=
      d.error?`<span style="color:var(--red)">${esc(d.error)}</span>`
      :`<span style="color:var(--green)">✅ ${esc(d.filename)} · ${d.rows.toLocaleString()} rows · Status: ${esc(d.status_col)}</span>`;
  });
}

function runCompare(){
  document.getElementById('compare-out').innerHTML=
    '<div class="loading-row"><div class="spinner"></div> Comparing…</div>';
  fetch('/compare').then(r=>r.json()).then(d=>{
    if(d.error){
      document.getElementById('compare-out').innerHTML=
        `<div class="alert alert-warn">⚠️ ${esc(d.error)}</div>`; return;
    }
    const dc=n=>n>0?'c-green':n<0?'c-red':'c-blue';
    let h=`<div class="diff-cards">
      <div class="diff-card"><div class="diff-val c-blue">${d.curr_total.toLocaleString()}</div>
        <div class="diff-lbl">Current Total</div></div>
      <div class="diff-card"><div class="diff-val c-blue">${d.prev_total.toLocaleString()}</div>
        <div class="diff-lbl">Previous Total</div></div>
      <div class="diff-card"><div class="diff-val ${dc(d.diff_total)}">${d.diff_total>=0?'+':''}${d.diff_total}</div>
        <div class="diff-lbl">Net Change</div></div>
      <div class="diff-card"><div class="diff-val c-green">${d.new_count}</div>
        <div class="diff-lbl">New Servers</div></div>
      <div class="diff-card"><div class="diff-val c-red">${d.removed_count}</div>
        <div class="diff-lbl">Removed</div></div>
      <div class="diff-card"><div class="diff-val c-cyan">${(d.to_live||[]).length}</div>
        <div class="diff-lbl">Went Live</div></div>
      <div class="diff-card"><div class="diff-val c-orange">${(d.to_not_live||[]).length}</div>
        <div class="diff-lbl">Moved to Project Phase</div></div>
    </div>`;
    if(d.charts){
      h+='<div class="charts-grid">';
      ['plat_compare','live_compare'].forEach(k=>{
        if(d.charts[k]) h+=`<div class="chart-card" style="grid-column:1/-1">
          <img src="${d.charts[k]}" alt="${k}"></div>`;
      });
      h+='</div>';
    }
    if(d.to_live&&d.to_live.length)      h+=cmpTable('✅ Went Live this month (now in BAU)',d.to_live,'var(--green)');
    if(d.to_not_live&&d.to_not_live.length) h+=cmpTable('⚠️ Moved back to Project Phase (Live→Not Live)',d.to_not_live,'var(--orange)');
    if(d.new_servers&&d.new_servers.length)  h+=cmpTable('🆕 New Servers',d.new_servers,'var(--blue)');
    if(d.removed_servers&&d.removed_servers.length) h+=cmpTable('🗑️ Removed Servers',d.removed_servers,'var(--red)');
    document.getElementById('compare-out').innerHTML=h;
  });
}

function cmpTable(title, rows, color){
  if(!rows||!rows.length) return '';
  const cols=Object.keys(rows[0]);
  return `<div style="margin-bottom:18px">
    <div class="shdr" style="color:${color}">${title} (${rows.length})</div>
    <div class="tbl-wrap"><table>
      <thead><tr>${cols.map(c=>`<th>${esc(c)}</th>`).join('')}</tr></thead>
      <tbody>${rows.map(r=>`<tr>${cols.map(c=>{
        const v=r[c]||'';
        if(c==='Server HostName') return `<td style="font-weight:700;color:var(--blue);font-family:monospace">${esc(v)}</td>`;
        if(c==='Curr Status') return `<td>${statusBadge(v)}</td>`;
        return `<td>${esc(v)}</td>`;
      }).join('')}</tr>`).join('')}
      </tbody></table></div></div>`;
}

// ════════════════════════════════════════════════════════════════
//  UPLOAD (admin)
// ════════════════════════════════════════════════════════════════
function handleDrop(e){
  e.preventDefault();
  document.getElementById('drop-zone').classList.remove('drag');
  if(e.dataTransfer.files[0]) uploadFile(e.dataTransfer.files[0]);
}
function uploadFile(file){
  if(!file) return;
  const fd=new FormData();fd.append('file',file);fd.append('which','current');
  const el=document.getElementById('upload-status');
  el.innerHTML='<div class="alert alert-info"><div class="spinner"></div> Uploading…</div>';
  fetch('/upload',{method:'POST',body:fd}).then(r=>r.json()).then(d=>{
    if(d.error){ el.innerHTML=`<div class="alert alert-error">⚠️ ${esc(d.error)}</div>`; return; }
    window._currFile=d.filename;
    document.getElementById('topbar-meta').textContent=
      `${d.filename} · ${d.rows.toLocaleString()} servers · Status: ${d.status_col}`;
    let qHtml='';
    if(d.quality&&Object.keys(d.quality).length){
      const items=Object.entries(d.quality).map(([k,v])=>
        `<div style="font-size:.8rem;color:var(--muted)">⚠️ <b>${esc(k)}</b>: <span style="color:var(--yellow)">${v}</span> blank rows</div>`).join('');
      qHtml=`<div class="quality-panel" style="margin-top:10px"><h4>⚠️ Quality Warnings</h4>${items}</div>`;
    }
    el.innerHTML=`<div class="alert alert-success">
      ✅ Loaded <b>${d.rows.toLocaleString()}</b> servers from <b>${esc(d.filename)}</b><br>
      Status column: <b style="color:var(--green)">${esc(d.status_col)}</b>
    </div>${qHtml}`;
    _filterLoaded=false; // reset so filter options reload
  });
}

// ════════════════════════════════════════════════════════════════
//  ADMIN PANEL
// ════════════════════════════════════════════════════════════════
let _adminData = null;

function loadAdmin(){
  document.getElementById('admin-loading').style.display='';
  document.getElementById('admin-content').style.display='none';
  fetch('/admin/data').then(r=>r.json()).then(d=>{
    _adminData = d;
    document.getElementById('admin-loading').style.display='none';
    document.getElementById('admin-content').style.display='';

    // Store correctable cols for modal
    window._correctable_cols = d.correctable_cols||[];
    PRESET_TAGS = d.preset_tags||[];

    // ── Summary cards ──
    document.getElementById('admin-stats').innerHTML=`
      <div class="admin-card">
        <div class="num c-blue">${d.server_count.toLocaleString()}</div>
        <div class="lbl">Servers Loaded</div>
        <div style="font-size:.7rem;color:var(--muted);margin-top:4px">${esc(d.filename)}<br>${esc(d.uploaded_at)}</div>
      </div>
      <div class="admin-card">
        <div class="num c-purple">${d.tag_count}</div><div class="lbl">Total Tags</div></div>
      <div class="admin-card">
        <div class="num c-blue">${d.note_count}</div><div class="lbl">Total Notes</div></div>
      <div class="admin-card">
        <div class="num c-red">${d.flag_count}</div><div class="lbl">Flagged Servers</div></div>
      <div class="admin-card">
        <div class="num c-orange">${d.corr_pending}</div>
        <div class="lbl">Pending Corrections</div></div>
      <div class="admin-card">
        <div class="num c-green">${d.corr_approved}</div>
        <div class="lbl">Approved Corrections</div></div>
      <div class="admin-card">
        <div class="num c-red">${d.decom_pending}</div>
        <div class="lbl">Pending Decommissions</div></div>
      <div class="admin-card">
        <div class="num c-green">${d.decom_approved}</div>
        <div class="lbl">Approved Decommissions</div></div>
      <div class="admin-card">
        <div class="num c-green">${d.newsrv_pending}</div>
        <div class="lbl">Pending New Servers</div></div>
      <div class="admin-card">
        <div class="num c-green">${d.newsrv_approved}</div>
        <div class="lbl">Approved New Servers</div></div>
      <div class="admin-card">
        <div class="num c-cyan">${d.tshirt_pending}</div>
        <div class="lbl">Pending CPU/RAM Changes</div></div>
      <div class="admin-card">
        <div class="num c-green">${d.tshirt_approved}</div>
        <div class="lbl">Approved CPU/RAM Changes</div></div>
      ${d.quality&&Object.keys(d.quality).length?
        `<div class="admin-card" style="text-align:left;grid-column:1/-1">
          <div class="quality-panel"><h4>⚠️ Data Quality</h4>
          ${Object.entries(d.quality).map(([k,v])=>
            `<div style="font-size:.8rem;color:var(--muted)">⚠️ <b>${esc(k)}</b>: <span style="color:var(--yellow)">${v}</span> blank rows</div>`
          ).join('')}</div></div>`:''}`;

    // ── Monthly Changes (Approved items, ready to apply) ──
    const totalApproved = (d.corr_approved||0) + (d.decom_approved||0) + (d.newsrv_approved||0) + (d.tshirt_approved||0);
    const mcEl = document.getElementById('monthly-changes-banner');
    if(mcEl){
      mcEl.innerHTML = totalApproved
        ? `<div class="alert alert-success" style="display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:10px">
            <span>📦 <b>${totalApproved}</b> approved change${totalApproved===1?'':'s'} ready to apply in next month's master Excel.</span>
            <a href="/admin/monthly_changes_report"><button class="btn btn-sm btn-green">📥 Download Monthly Changes Report</button></a>
          </div>`
        : `<div class="alert alert-info">No approved changes yet this cycle. Review pending items below.</div>`;
    }

    // ── Corrections ──
    const pending   = (d.corrections||[]).filter(c=>c.status==='Pending');
    const approved  = (d.corrections||[]).filter(c=>c.status==='Approved');
    const rejected  = (d.corrections||[]).filter(c=>c.status==='Rejected');
    const pendBadge = document.getElementById('corr-pending-badge');
    if(pendBadge){ pendBadge.style.display=pending.length?'':'none'; pendBadge.textContent=`${pending.length} pending`; }
    document.getElementById('corr-pending-count').textContent=`(${pending.length})`;
    document.getElementById('corr-approved-count').textContent=`(${approved.length})`;
    document.getElementById('corr-rejected-count').textContent=`(${rejected.length})`;

    renderCorrTable('corr-pending-table', pending, 'pending');
    renderCorrTable('corr-approved-table', approved, 'approved');
    renderCorrTable('corr-rejected-table', rejected, 'rejected');

    // ── Decommissions ──
    const decPending  = (d.decommissions||[]).filter(c=>c.status==='Pending');
    const decApproved = (d.decommissions||[]).filter(c=>c.status==='Approved');
    const decRejected = (d.decommissions||[]).filter(c=>c.status==='Rejected');
    const decBadge = document.getElementById('decom-pending-badge');
    if(decBadge){ decBadge.style.display=decPending.length?'':'none'; decBadge.textContent=`${decPending.length} pending`; }
    document.getElementById('decom-pending-count').textContent=`(${decPending.length})`;
    document.getElementById('decom-approved-count').textContent=`(${decApproved.length})`;
    document.getElementById('decom-rejected-count').textContent=`(${decRejected.length})`;
    renderDecomTable('decom-pending-table', decPending, 'pending');
    renderDecomTable('decom-approved-table', decApproved, 'approved');
    renderDecomTable('decom-rejected-table', decRejected, 'rejected');

    // ── New Servers ──
    const nsPending  = (d.new_servers||[]).filter(c=>c.status==='Pending');
    const nsApproved = (d.new_servers||[]).filter(c=>c.status==='Approved');
    const nsRejected = (d.new_servers||[]).filter(c=>c.status==='Rejected');
    const nsBadge = document.getElementById('newsrv-pending-badge');
    if(nsBadge){ nsBadge.style.display=nsPending.length?'':'none'; nsBadge.textContent=`${nsPending.length} pending`; }
    document.getElementById('newsrv-pending-count').textContent=`(${nsPending.length})`;
    document.getElementById('newsrv-approved-count').textContent=`(${nsApproved.length})`;
    document.getElementById('newsrv-rejected-count').textContent=`(${nsRejected.length})`;
    renderNewsrvTable('newsrv-pending-table', nsPending, 'pending');
    renderNewsrvTable('newsrv-approved-table', nsApproved, 'approved');
    renderNewsrvTable('newsrv-rejected-table', nsRejected, 'rejected');

    // ── T-Shirt Size Changes ──
    const tsPending  = (d.tshirt_changes||[]).filter(c=>c.status==='Pending');
    const tsApproved = (d.tshirt_changes||[]).filter(c=>c.status==='Approved');
    const tsRejected = (d.tshirt_changes||[]).filter(c=>c.status==='Rejected');
    const tsBadge = document.getElementById('tshirt-pending-badge');
    if(tsBadge){ tsBadge.style.display=tsPending.length?'':'none'; tsBadge.textContent=`${tsPending.length} pending`; }
    document.getElementById('tshirt-pending-count').textContent=`(${tsPending.length})`;
    document.getElementById('tshirt-approved-count').textContent=`(${tsApproved.length})`;
    document.getElementById('tshirt-rejected-count').textContent=`(${tsRejected.length})`;
    renderTshirtTable('tshirt-pending-table', tsPending, 'pending');
    renderTshirtTable('tshirt-approved-table', tsApproved, 'approved');
    renderTshirtTable('tshirt-rejected-table', tsRejected, 'rejected');

    // ── Flags ──
    const flags = d.flag_rows||[];
    const flagEl = document.getElementById('admin-flags');
    if(!flags.length){
      flagEl.innerHTML='<div class="alert alert-success">✅ No flagged servers.</div>';
    } else {
      flagEl.innerHTML=`<div class="tbl-wrap"><table>
        <thead><tr><th>Hostname</th><th>Reason</th><th>Flagged By</th><th>Time</th>
          <th>Platform</th><th>OS</th><th>Action</th></tr></thead>
        <tbody>${flags.map(f=>`<tr>
          <td style="font-weight:700;color:var(--blue);font-family:monospace;cursor:pointer"
            onclick="showDetail('${esc(f.hostname)}')">${esc(f.hostname)}</td>
          <td style="font-size:.8rem">${esc(f.reason||'—')}</td>
          <td style="font-size:.78rem;color:var(--muted)">${esc(f.user)}</td>
          <td style="font-size:.75rem;color:var(--muted)">${esc(f.ts)}</td>
          <td style="font-size:.78rem">${esc(f.platform)}</td>
          <td style="font-size:.78rem">${esc(f.os)}</td>
          <td><button class="btn btn-xs btn-outline"
            onclick="adminUnflag('${esc(f.hostname)}')">✓ Clear</button></td>
        </tr>`).join('')}</tbody></table></div>`;
    }

    // ── Tags inline table ──
    const tagRows = d.tag_rows||[];
    const tagEl   = document.getElementById('admin-tags-table');
    if(!tagRows.length){
      tagEl.innerHTML='<div class="alert alert-info">No tags added yet.</div>';
    } else {
      const unverifiedCount = tagRows.filter(r=>r.status==='Unverified').length;
      tagEl.innerHTML=`<div style="font-size:.78rem;color:var(--muted);margin-bottom:6px">
        ${tagRows.length} tag entries across ${new Set(tagRows.map(r=>r.hostname)).size} servers
        ${unverifiedCount?` · <span style="color:var(--yellow)">${unverifiedCount} unverified</span>`:''}</div>
        <div class="tbl-wrap"><table>
        <thead><tr><th>Hostname</th><th>Tag</th><th>Status</th><th>Submitted By</th>
          <th>Platform</th><th>Server Role</th><th>OS</th><th>Action</th></tr></thead>
        <tbody>${tagRows.map(r=>`<tr>
          <td style="font-weight:700;color:var(--blue);font-family:monospace;cursor:pointer"
            onclick="showDetail('${esc(r.hostname)}')">${esc(r.tag.length?r.hostname:r.hostname)}</td>
          <td><span class="tag-pill" style="${r.status==='Accepted'?'':'border-style:dashed;opacity:.85'}">${esc(r.tag)}</span></td>
          <td>${r.status==='Accepted'
            ?`<span class="badge badge-live">✓ Accepted</span>`
            :`<span class="badge" style="background:rgba(255,194,64,.15);color:var(--yellow);border-color:rgba(255,194,64,.3)">? Unverified</span>`}</td>
          <td style="font-size:.78rem">${esc(r.name||r.user||'—')}<br><span style="font-size:.7rem;color:var(--muted)">${esc(r.ts||'')}</span></td>
          <td style="font-size:.8rem">${esc(r.platform)}</td>
          <td style="font-size:.8rem">${esc(r.role)}</td>
          <td style="font-size:.78rem">${esc(r.os)}</td>
          <td style="display:flex;gap:5px">
            ${r.status!=='Accepted'?`<button class="btn btn-xs btn-green" onclick="adminAcceptTag('${esc(r.hn_key)}','${r.id}')">✓ Accept</button>`:''}
            <button class="btn btn-xs btn-red" onclick="adminRemoveTag('${esc(r.hn_key)}','${r.id}')">🗑 Remove</button>
          </td>
        </tr>`).join('')}</tbody></table></div>`;
    }

    // ── Notes inline table ──
    const noteRows = d.note_rows||[];
    const noteEl   = document.getElementById('admin-notes-table');
    if(!noteRows.length){
      noteEl.innerHTML='<div class="alert alert-info">No notes added yet.</div>';
    } else {
      noteEl.innerHTML=`<div style="font-size:.78rem;color:var(--muted);margin-bottom:6px">
        ${noteRows.length} notes across ${new Set(noteRows.map(r=>r.hostname)).size} servers</div>
        <div class="tbl-wrap"><table>
        <thead><tr><th>Hostname</th><th>Note</th><th>Posted By</th>
          <th>Time</th><th>Platform</th><th>Role</th><th>Action</th></tr></thead>
        <tbody>${noteRows.map(n=>`<tr>
          <td style="font-weight:700;color:var(--blue);font-family:monospace;cursor:pointer"
            onclick="showDetail('${esc(n.hostname)}')">${esc(n.hostname)}</td>
          <td style="font-size:.8rem;max-width:260px;white-space:normal;line-height:1.4">${esc(n.note)}</td>
          <td style="font-size:.78rem">${esc(n.by)}</td>
          <td style="font-size:.75rem;color:var(--muted)">${esc(n.ts)}</td>
          <td style="font-size:.78rem">${esc(n.platform)}</td>
          <td style="font-size:.78rem">${esc(n.role)}</td>
          <td><button class="btn btn-xs btn-outline"
            onclick="adminDeleteNote('${esc(n.hostname)}','${esc(n.id)}',this)">🗑</button></td>
        </tr>`).join('')}</tbody></table></div>`;
    }
  });
}

function renderCorrTable(containerId, rows, mode){
  // mode: 'pending' | 'approved' | 'rejected'
  const el=document.getElementById(containerId);
  if(!rows.length){
    const msg = {pending:'No pending corrections — all clear! ✅',
                 approved:'No approved corrections yet.',
                 rejected:'No rejected corrections.'}[mode];
    el.innerHTML=`<div class="alert ${mode==='pending'?'alert-info':'alert-success'}">${msg}</div>`;
    return;
  }
  const isPending = mode === 'pending';
  el.innerHTML=`<div class="tbl-wrap"><table>
    <thead><tr>
      <th>Hostname</th><th>Column</th><th>Current Value</th>
      <th>Suggested Value</th><th>Reason</th>
      <th>Submitted By</th><th>Date</th>
      ${isPending?'<th>Actions</th>':'<th>Decision By</th><th>Decision Note</th>'}
    </tr></thead>
    <tbody>${rows.map(r=>`<tr>
      <td style="font-weight:700;color:var(--blue);font-family:monospace;cursor:pointer"
        onclick="showDetail('${esc(r.hostname)}')">${esc(r.hostname)}</td>
      <td><span class="badge badge-other" style="font-size:.7rem">${esc(r.column)}</span></td>
      <td style="font-size:.78rem;color:var(--red)">${esc(r.current_val||'(blank)')}</td>
      <td style="font-size:.78rem;color:var(--green);font-weight:600">${esc(r.suggested_val)}</td>
      <td style="font-size:.76rem;color:var(--muted);max-width:180px;white-space:normal">${esc(r.reason||'—')}</td>
      <td style="font-size:.76rem">${esc(r.name)}</td>
      <td style="font-size:.74rem;color:var(--muted)">${esc(r.ts)}</td>
      ${isPending
        ?`<td style="display:flex;gap:5px">
          <button class="btn btn-xs btn-green" onclick="openDecisionModal('correction','${r.id}','${esc(r.hostname)}')">⚖️ Review</button>
          <button class="btn btn-xs btn-red"   onclick="deleteCorr('${r.id}',this)">🗑 Delete</button>
        </td>`
        :`<td style="font-size:.74rem">${esc(r.reviewed_by||'—')} · ${esc(r.reviewed_ts||'')}</td>
          <td style="font-size:.74rem;color:var(--muted)">${esc(r.decision_reason||'—')}</td>`}
    </tr>`).join('')}</tbody></table></div>`;
}

function showCorrTab(tab){
  ['pending','approved','rejected'].forEach(t=>{
    document.getElementById(`corr-${t}-table`).style.display = (t===tab)?'':'none';
    const btn = document.getElementById(`corr-tab-${t}`);
    btn.className = (t===tab) ? 'btn btn-sm' : 'btn btn-sm btn-outline';
    btn.style.background = (t===tab) ? {pending:'var(--orange)',approved:'var(--green)',rejected:'var(--red)'}[t] : '';
    btn.style.color = (t===tab) ? (t==='rejected'?'#fff':'#000') : '';
  });
}

async function deleteCorr(id, btn){
  if(!confirm('Delete this correction suggestion?')) return;
  btn.disabled=true;
  const r=await fetch('/admin/correction/delete',{method:'POST',
    headers:{'Content-Type':'application/json'},
    body:JSON.stringify({id})}).then(x=>x.json());
  if(r.ok) loadAdmin();
}

// ── Decommissions admin table ──
function renderDecomTable(containerId, rows, mode){
  const el=document.getElementById(containerId);
  if(!rows.length){
    const msg = {pending:'No pending decommission reports — all clear! ✅',
                 approved:'No approved decommissions yet.',
                 rejected:'No rejected decommissions.'}[mode];
    el.innerHTML=`<div class="alert ${mode==='pending'?'alert-info':'alert-success'}">${msg}</div>`;
    return;
  }
  const isPending = mode === 'pending';
  el.innerHTML=`<div class="tbl-wrap"><table>
    <thead><tr>
      <th>Hostname</th><th>Still in Inventory?</th><th>Reason</th><th>Eff. Date</th>
      <th>Comment</th><th>Submitted By</th><th>Date</th>
      ${isPending?'<th>Actions</th>':'<th>Decision By</th><th>Decision Note</th>'}
    </tr></thead>
    <tbody>${rows.map(r=>`<tr>
      <td style="font-weight:700;color:var(--blue);font-family:monospace;cursor:pointer"
        onclick="showDetail('${esc(r.hostname)}')">${esc(r.hostname)}</td>
      <td>${r._in_inventory
        ?'<span class="badge badge-live">Yes</span>'
        :'<span class="badge badge-other">Not Found</span>'}</td>
      <td><span class="badge" style="background:rgba(255,79,106,.15);color:var(--red);
        border:1px solid rgba(255,79,106,.3)">${esc(r.reason)}</span></td>
      <td style="font-size:.78rem">${esc(r.eff_date||'—')}</td>
      <td style="font-size:.76rem;color:var(--muted);max-width:160px;white-space:normal">${esc(r.comment||'—')}</td>
      <td style="font-size:.76rem">${esc(r.name)}</td>
      <td style="font-size:.74rem;color:var(--muted)">${esc(r.ts)}</td>
      ${isPending
        ?`<td style="display:flex;gap:5px">
          <button class="btn btn-xs btn-green" onclick="openDecisionModal('decommission','${r.id}','${esc(r.hostname)}')">⚖️ Review</button>
          <button class="btn btn-xs btn-red"   onclick="deleteDecom('${r.id}',this)">🗑 Delete</button>
        </td>`
        :`<td style="font-size:.74rem">${esc(r.reviewed_by||'—')} · ${esc(r.reviewed_ts||'')}</td>
          <td style="font-size:.74rem;color:var(--muted)">${esc(r.decision_reason||'—')}</td>`}
    </tr>`).join('')}</tbody></table></div>`;
}
function showDecomTab(tab){
  ['pending','approved','rejected'].forEach(t=>{
    document.getElementById(`decom-${t}-table`).style.display = (t===tab)?'':'none';
    const btn = document.getElementById(`decom-tab-${t}`);
    btn.className = (t===tab) ? 'btn btn-sm' : 'btn btn-sm btn-outline';
    btn.style.background = (t===tab) ? {pending:'var(--red)',approved:'var(--green)',rejected:'var(--red)'}[t] : '';
    btn.style.color = (t===tab) ? '#fff' : '';
    if(t===tab && t==='approved') btn.style.color = '#000';
  });
}
async function deleteDecom(id, btn){
  if(!confirm('Delete this decommission report?')) return;
  btn.disabled=true;
  const r=await fetch('/admin/decommission/delete',{method:'POST',
    headers:{'Content-Type':'application/json'},body:JSON.stringify({id})}).then(x=>x.json());
  if(r.ok) loadAdmin();
}

// ── New servers admin table ──
function renderNewsrvTable(containerId, rows, mode){
  const el=document.getElementById(containerId);
  if(!rows.length){
    const msg = {pending:'No pending new-server suggestions — all clear! ✅',
                 approved:'No approved new servers yet.',
                 rejected:'No rejected new servers.'}[mode];
    el.innerHTML=`<div class="alert ${mode==='pending'?'alert-info':'alert-success'}">${msg}</div>`;
    return;
  }
  const isPending = mode === 'pending';
  const fieldCols = ["Server HostName","Platform","Server Type(Physical/Virtual)",
    "Final OS","Server Role","Server DC Location"];
  el.innerHTML=`<div class="tbl-wrap"><table>
    <thead><tr>
      ${fieldCols.map(c=>`<th>${esc(c)}</th>`).join('')}
      <th>Comment</th><th>Submitted By</th><th>Date</th>
      ${isPending?'<th>Actions</th>':'<th>Decision By</th><th>Decision Note</th>'}
    </tr></thead>
    <tbody>${rows.map(r=>{const f=r.fields||{};return `<tr>
      ${fieldCols.map((c,i)=>i===0
        ?`<td style="font-weight:700;color:var(--blue);font-family:monospace;cursor:pointer"
            onclick="showDetail('${esc(f[c]||'')}')">${esc(f[c]||'—')}</td>`
        :`<td style="font-size:.78rem">${esc(f[c]||'—')}</td>`).join('')}
      <td style="font-size:.76rem;color:var(--muted);max-width:160px;white-space:normal">${esc(r.comment||'—')}</td>
      <td style="font-size:.76rem">${esc(r.name)}</td>
      <td style="font-size:.74rem;color:var(--muted)">${esc(r.ts)}</td>
      ${isPending
        ?`<td style="display:flex;gap:5px">
          <button class="btn btn-xs btn-green" onclick="openDecisionModal('newserver','${r.id}','${esc(f['Server HostName']||'')}')">⚖️ Review</button>
          <button class="btn btn-xs btn-red"   onclick="deleteNewsrv('${r.id}',this)">🗑 Delete</button>
        </td>`
        :`<td style="font-size:.74rem">${esc(r.reviewed_by||'—')} · ${esc(r.reviewed_ts||'')}</td>
          <td style="font-size:.74rem;color:var(--muted)">${esc(r.decision_reason||'—')}</td>`}
    </tr>`;}).join('')}</tbody></table></div>`;
}
function showNewsrvTab(tab){
  ['pending','approved','rejected'].forEach(t=>{
    document.getElementById(`newsrv-${t}-table`).style.display = (t===tab)?'':'none';
    const btn = document.getElementById(`newsrv-tab-${t}`);
    btn.className = (t===tab) ? 'btn btn-sm' : 'btn btn-sm btn-outline';
    btn.style.background = (t===tab) ? {pending:'var(--green)',approved:'var(--green)',rejected:'var(--red)'}[t] : '';
    btn.style.color = (t===tab) ? (t==='rejected'?'#fff':'#000') : '';
  });
}
async function deleteNewsrv(id, btn){
  if(!confirm('Delete this new-server suggestion?')) return;
  btn.disabled=true;
  const r=await fetch('/admin/newserver/delete',{method:'POST',
    headers:{'Content-Type':'application/json'},body:JSON.stringify({id})}).then(x=>x.json());
  if(r.ok) loadAdmin();
}

// ── T-Shirt size (CPU/RAM) admin table ──
function renderTshirtTable(containerId, rows, mode){
  const el=document.getElementById(containerId);
  if(!rows.length){
    const msg = {pending:'No pending CPU/RAM change requests — all clear! ✅',
                 approved:'No approved CPU/RAM changes yet.',
                 rejected:'No rejected CPU/RAM changes.'}[mode];
    el.innerHTML=`<div class="alert ${mode==='pending'?'alert-info':'alert-success'}">${msg}</div>`;
    return;
  }
  const isPending = mode === 'pending';
  el.innerHTML=`<div class="tbl-wrap"><table>
    <thead><tr>
      <th>Hostname</th><th>Platform</th><th>Current CPU/RAM</th><th>New CPU/RAM</th>
      <th>Reason</th><th>Submitted By</th><th>Date</th>
      ${isPending?'<th>Actions</th>':'<th>Decision By</th><th>Decision Note</th>'}
    </tr></thead>
    <tbody>${rows.map(r=>`<tr>
      <td style="font-weight:700;color:var(--blue);font-family:monospace;cursor:pointer"
        onclick="showDetail('${esc(r.hostname)}')">${esc(r.hostname)}</td>
      <td><span class="badge badge-other">${esc(r.platform)}</span></td>
      <td style="font-size:.78rem;color:var(--red)">${esc(r.current_cpu||'—')} / ${esc(r.current_ram||'—')}</td>
      <td style="font-size:.78rem;color:var(--green);font-weight:600">${esc(r.new_cpu||'—')} / ${esc(r.new_ram||'—')}</td>
      <td style="font-size:.76rem;color:var(--muted);max-width:160px;white-space:normal">${esc(r.reason||'—')}</td>
      <td style="font-size:.76rem">${esc(r.name)}</td>
      <td style="font-size:.74rem;color:var(--muted)">${esc(r.ts)}</td>
      ${isPending
        ?`<td style="display:flex;gap:5px">
          <button class="btn btn-xs btn-green" onclick="openDecisionModal('tshirt','${r.id}','${esc(r.hostname)}')">⚖️ Review</button>
          <button class="btn btn-xs btn-red"   onclick="deleteTshirt('${r.id}',this)">🗑 Delete</button>
        </td>`
        :`<td style="font-size:.74rem">${esc(r.reviewed_by||'—')} · ${esc(r.reviewed_ts||'')}</td>
          <td style="font-size:.74rem;color:var(--muted)">${esc(r.decision_reason||'—')}</td>`}
    </tr>`).join('')}</tbody></table></div>`;
}
function showTshirtTab(tab){
  ['pending','approved','rejected'].forEach(t=>{
    document.getElementById(`tshirt-${t}-table`).style.display = (t===tab)?'':'none';
    const btn = document.getElementById(`tshirt-tab-${t}`);
    btn.className = (t===tab) ? 'btn btn-sm' : 'btn btn-sm btn-outline';
    btn.style.background = (t===tab) ? {pending:'var(--cyan)',approved:'var(--green)',rejected:'var(--red)'}[t] : '';
    btn.style.color = (t===tab) ? (t==='rejected'?'#fff':'#000') : '';
  });
}
async function deleteTshirt(id, btn){
  if(!confirm('Delete this CPU/RAM change request?')) return;
  btn.disabled=true;
  const r=await fetch('/admin/tshirt/delete',{method:'POST',
    headers:{'Content-Type':'application/json'},body:JSON.stringify({id})}).then(x=>x.json());
  if(r.ok) loadAdmin();
}

// ── Shared Approve/Reject decision modal ──
let _decisionTarget = {type:'', id:'', hostname:''};
let _decisionChoice = '';

function openDecisionModal(type, id, hostname){
  _decisionTarget = {type, id, hostname};
  _decisionChoice = '';
  document.getElementById('decision-modal-title').textContent = `Review: ${hostname}`;
  document.getElementById('decision-reason-input').value = '';
  document.getElementById('decision-modal-msg').innerHTML = '';
  document.getElementById('decision-btn-approve').style.outline = 'none';
  document.getElementById('decision-btn-reject').style.outline = 'none';
  const submitBtn = document.getElementById('decision-submit-btn');
  submitBtn.disabled = true; submitBtn.style.opacity = '.5';
  submitBtn.textContent = 'Select Approve or Reject first';
  document.getElementById('decision-modal-overlay').classList.add('show');
}
function closeDecisionModal(e){
  if(e.target===document.getElementById('decision-modal-overlay'))
    document.getElementById('decision-modal-overlay').classList.remove('show');
}
function setDecisionChoice(choice){
  _decisionChoice = choice;
  document.getElementById('decision-btn-approve').style.outline =
    choice==='Approved' ? '2px solid #fff' : 'none';
  document.getElementById('decision-btn-reject').style.outline =
    choice==='Rejected' ? '2px solid #fff' : 'none';
  const submitBtn = document.getElementById('decision-submit-btn');
  submitBtn.disabled = false; submitBtn.style.opacity = '1';
  submitBtn.textContent = choice==='Approved' ? '✅ Confirm Approve' : '❌ Confirm Reject';
  submitBtn.style.background = choice==='Approved' ? 'var(--green)' : 'var(--red)';
  submitBtn.style.color = choice==='Approved' ? '#000' : '#fff';
}
async function submitDecisionModal(){
  if(!_decisionChoice) return;
  const reason = document.getElementById('decision-reason-input').value.trim();
  const endpoints = {
    correction: '/admin/correction/review',
    decommission: '/admin/decommission/review',
    newserver: '/admin/newserver/review',
    tshirt: '/admin/tshirt/review',
  };
  const r = await fetch(endpoints[_decisionTarget.type], {method:'POST',
    headers:{'Content-Type':'application/json'},
    body:JSON.stringify({id:_decisionTarget.id, decision:_decisionChoice, reason})}).then(x=>x.json());
  if(r.error){
    document.getElementById('decision-modal-msg').innerHTML =
      `<div class="alert alert-error">${esc(r.error)}</div>`;
    return;
  }
  document.getElementById('decision-modal-overlay').classList.remove('show');
  loadAdmin();
}

async function adminDeleteNote(hostname, noteId, btn){
  if(!confirm('Delete this note?')) return;
  btn.disabled=true;
  const r=await fetch('/note/delete',{method:'POST',
    headers:{'Content-Type':'application/json'},
    body:JSON.stringify({hostname,id:noteId})}).then(x=>x.json());
  if(r.ok) loadAdmin();
}

async function adminUnflag(hostname){
  await fetch('/flag/remove',{method:'POST',headers:{'Content-Type':'application/json'},
    body:JSON.stringify({hostname})});
  loadAdmin();
}

async function adminAcceptTag(hostnameKey, tagId){
  await fetch('/admin/tag/accept',{method:'POST',headers:{'Content-Type':'application/json'},
    body:JSON.stringify({hostname:hostnameKey,id:tagId})});
  loadAdmin();
}
async function adminRemoveTag(hostnameKey, tagId){
  if(!confirm('Remove this tag?')) return;
  await fetch('/tag/remove',{method:'POST',headers:{'Content-Type':'application/json'},
    body:JSON.stringify({hostname:hostnameKey,id:tagId})});
  loadAdmin();
}

// ════════════════════════════════════════════════════════════════
//  MANAGE USERS
// ════════════════════════════════════════════════════════════════
function loadUsers(){
  fetch('/admin/users').then(r=>r.json()).then(d=>{
    const ul=document.getElementById('users-list');
    ul.innerHTML=(d.users||[]).map(u=>`
      <div class="user-row ${u.enabled?'':'disabled-row'}" id="ur-${esc(u.username)}">
        <div class="user-avatar" style="width:34px;height:34px;font-size:.88rem">
          ${(u.full_name||u.username||'?')[0].toUpperCase()}</div>
        <div class="user-info">
          <div class="user-name">${esc(u.full_name||u.username)}
            <span class="role-badge ${u.role==='admin'?'role-admin':'role-user'}">${u.role}</span>
            ${!u.enabled?'<span style="color:var(--red);font-size:.7rem"> ● Disabled</span>':''}
          </div>
          <div class="user-meta">@${esc(u.username)} · Created ${esc(u.created_at?u.created_at.split('T')[0]:'')}</div>
        </div>
        <div style="display:flex;gap:6px;flex-wrap:wrap">
          <button class="btn btn-xs btn-outline" onclick="openPwModal('${esc(u.username)}')">🔑 Reset PW</button>
          <button class="btn btn-xs ${u.enabled?'btn-red':'btn-green'}"
            onclick="toggleUser('${esc(u.username)}')">${u.enabled?'Disable':'Enable'}</button>
          <button class="btn btn-xs ${u.role==='admin'?'btn-yellow':'btn-outline'}"
            onclick="changeRole('${esc(u.username)}','${u.role==='admin'?'user':'admin'}')">
            ${u.role==='admin'?'→ User':'→ Admin'}</button>
        </div>
      </div>`).join('');
  });
}

async function createUser(){
  const uname=document.getElementById('nu-uname').value.trim();
  const fname=document.getElementById('nu-fname').value.trim();
  const pw=document.getElementById('nu-pw').value.trim();
  const role=document.getElementById('nu-role').value;
  const msg=document.getElementById('user-create-msg');
  if(!uname||!pw){msg.innerHTML='<div class="alert alert-warn">Username and password required</div>';return;}
  const r=await fetch('/admin/user/create',{method:'POST',
    headers:{'Content-Type':'application/json'},
    body:JSON.stringify({username:uname,full_name:fname,password:pw,role})}).then(x=>x.json());
  if(r.error){ msg.innerHTML=`<div class="alert alert-error">${esc(r.error)}</div>`; return; }
  msg.innerHTML='<div class="alert alert-success">✅ User created!</div>';
  document.getElementById('nu-uname').value='';
  document.getElementById('nu-fname').value='';
  document.getElementById('nu-pw').value='';
  loadUsers();
}

async function toggleUser(uname){
  const r=await fetch('/admin/user/toggle',{method:'POST',
    headers:{'Content-Type':'application/json'},
    body:JSON.stringify({username:uname})}).then(x=>x.json());
  if(r.error) alert(r.error); else loadUsers();
}
async function changeRole(uname,newRole){
  const r=await fetch('/admin/user/change_role',{method:'POST',
    headers:{'Content-Type':'application/json'},
    body:JSON.stringify({username:uname,role:newRole})}).then(x=>x.json());
  if(r.error) alert(r.error); else loadUsers();
}

function openPwModal(uname){
  document.getElementById('pw-modal-uname').textContent=uname;
  document.getElementById('pw-new').value='';
  document.getElementById('pw-reset-msg').innerHTML='';
  document.getElementById('pw-modal-overlay').classList.add('show');
  window._pwResetTarget=uname;
}
function closePwModal(e){
  if(e.target===document.getElementById('pw-modal-overlay'))
    document.getElementById('pw-modal-overlay').classList.remove('show');
}
async function submitPwReset(){
  const pw=document.getElementById('pw-new').value.trim();
  const msg=document.getElementById('pw-reset-msg');
  const r=await fetch('/admin/user/reset_pw',{method:'POST',
    headers:{'Content-Type':'application/json'},
    body:JSON.stringify({username:window._pwResetTarget,password:pw})}).then(x=>x.json());
  if(r.error){ msg.innerHTML=`<div class="alert alert-error">${esc(r.error)}</div>`; return; }
  msg.innerHTML='<div class="alert alert-success">✅ Password updated!</div>';
  setTimeout(()=>document.getElementById('pw-modal-overlay').classList.remove('show'),1500);
}

// ════════════════════════════════════════════════════════════════
//  MARK DECOMMISSION PAGE
// ════════════════════════════════════════════════════════════════
let _decomReasons = ['Shutdown','Decommissioned'];

function loadDecomPage(){
  fillSel('decom-reason', _decomReasons);
  document.getElementById('decom-hostname').value = '';
  document.getElementById('decom-host-suggest').innerHTML = '';
  document.getElementById('decom-comment').value = '';
  document.getElementById('decom-date').value = '';
  document.getElementById('decom-msg').innerHTML = '';
  _decomSelected = [];
  renderDecomChips();
  loadMyDecoms();
}

let _decomDebTimer;
let _decomSelected = [];   // array of {hostname, platform}

function decomHostSearch(){
  clearTimeout(_decomDebTimer);
  const q = document.getElementById('decom-hostname').value.trim();
  if(q.length < 2){ document.getElementById('decom-host-suggest').innerHTML=''; return; }
  _decomDebTimer = setTimeout(()=>{
    fetch('/search?q='+encodeURIComponent(q)+'&page=1').then(r=>r.json()).then(d=>{
      const already = new Set(_decomSelected.map(s=>s.hostname.toLowerCase()));
      const rows = (d.rows||[]).filter(r=>!already.has((r['Server HostName']||'').toLowerCase())).slice(0,6);
      if(!rows.length){
        document.getElementById('decom-host-suggest').innerHTML =
          `<div style="font-size:.78rem;color:var(--muted);padding:6px 0">No matching servers found.</div>`;
        return;
      }
      document.getElementById('decom-host-suggest').innerHTML =
        `<div style="border:1px solid var(--border);border-radius:8px;overflow:hidden">
        ${rows.map(r=>`<div style="padding:7px 12px;cursor:pointer;font-size:.83rem;
          font-family:monospace;border-bottom:1px solid var(--border)"
          onmouseover="this.style.background='var(--surf2)'" onmouseout="this.style.background=''"
          onclick="pickDecomHost('${esc(r['Server HostName'])}','${esc(r['Platform']||'')}')">${esc(r['Server HostName'])}
          <span style="color:var(--muted);font-family:inherit;font-size:.78rem"> — ${esc(r['Platform']||'')}</span>
        </div>`).join('')}</div>`;
    });
  }, 300);
}

function pickDecomHost(hostname, platform){
  if(!_decomSelected.some(s=>s.hostname.toLowerCase()===hostname.toLowerCase())){
    _decomSelected.push({hostname, platform});
  }
  document.getElementById('decom-hostname').value = '';
  document.getElementById('decom-host-suggest').innerHTML = '';
  renderDecomChips();
}

function removeDecomChip(hostname){
  _decomSelected = _decomSelected.filter(s=>s.hostname!==hostname);
  renderDecomChips();
}

function renderDecomChips(){
  const el = document.getElementById('decom-selected-chips');
  const countEl = document.getElementById('decom-submit-count');
  if(!_decomSelected.length){
    el.innerHTML = `<div style="color:var(--dim);font-size:.8rem">No servers added yet — search above and click a result to add it.</div>`;
    countEl.textContent = '';
    return;
  }
  el.innerHTML = _decomSelected.map(s=>`
    <span class="tag-pill" style="background:rgba(255,79,106,.12);border-color:rgba(255,79,106,.3);color:var(--red)">
      ${esc(s.hostname)}${s.platform?` <span style="color:var(--muted)">(${esc(s.platform)})</span>`:''}
      <span class="rm" onclick="removeDecomChip('${esc(s.hostname)}')">✕</span>
    </span>`).join('');
  countEl.textContent = `(${_decomSelected.length})`;
}

async function submitDecom(){
  const rsn = document.getElementById('decom-reason').value;
  const dt  = document.getElementById('decom-date').value;
  const cmt = document.getElementById('decom-comment').value.trim();
  const msg = document.getElementById('decom-msg');
  if(!_decomSelected.length){
    msg.innerHTML = '<div class="alert alert-warn">Add at least one server first.</div>';
    return;
  }
  if(!rsn){
    msg.innerHTML = '<div class="alert alert-warn">Please select a reason.</div>';
    return;
  }
  msg.innerHTML = '<div class="spinner"></div> Submitting…';
  let okCount = 0, errCount = 0;
  for(const s of _decomSelected){
    const r = await fetch('/decommission/add',{method:'POST',
      headers:{'Content-Type':'application/json'},
      body:JSON.stringify({hostname:s.hostname,reason:rsn,eff_date:dt,comment:cmt})}).then(x=>x.json());
    if(r.ok) okCount++; else errCount++;
  }
  msg.innerHTML = errCount
    ? `<div class="alert alert-warn">✅ ${okCount} submitted. ⚠️ ${errCount} failed.</div>`
    : `<div class="alert alert-success">✅ ${okCount} server(s) submitted! Admin will review and apply them in next month's inventory.</div>`;
  _decomSelected = [];
  renderDecomChips();
  document.getElementById('decom-comment').value = '';
  document.getElementById('decom-date').value = '';
  loadMyDecoms();
}

function loadMyDecoms(){
  fetch('/admin/data').then(r=>{
    if(r.status===403){ document.getElementById('decom-my-list').innerHTML=''; return null; }
    return r.json();
  }).then(d=>{
    if(!d) return;
    const mine = (d.decommissions||[]).filter(x=>x.user===ME.username);
    renderMyDecomList(mine);
  }).catch(()=>{});
}
function renderMyDecomList(mine){
  const el = document.getElementById('decom-my-list');
  if(!mine.length){ el.innerHTML = '<div style="color:var(--dim);font-size:.82rem">No submissions yet.</div>'; return; }
  el.innerHTML = `<div class="tbl-wrap"><table>
    <thead><tr><th>Hostname</th><th>Reason</th><th>Eff. Date</th><th>Comment</th><th>Submitted</th><th>Status</th><th>Admin Note</th></tr></thead>
    <tbody>${mine.map(m=>`<tr>
      <td style="font-weight:700;color:var(--blue);font-family:monospace">${esc(m.hostname)}</td>
      <td><span class="badge" style="background:rgba(255,79,106,.15);color:var(--red);border:1px solid rgba(255,79,106,.3)">${esc(m.reason)}</span></td>
      <td style="font-size:.78rem">${esc(m.eff_date||'—')}</td>
      <td style="font-size:.78rem;color:var(--muted)">${esc(m.comment||'—')}</td>
      <td style="font-size:.74rem;color:var(--muted)">${esc(m.ts)}</td>
      <td>${reviewStatusBadge(m.status)}</td>
      <td style="font-size:.76rem;color:var(--muted)">${esc(m.decision_reason||'—')}</td>
    </tr>`).join('')}</tbody></table></div>`;
}

// ════════════════════════════════════════════════════════════════
//  SUGGEST NEW SERVER PAGE
// ════════════════════════════════════════════════════════════════
let _newsrvFields = [];
let _filterOptCache = null;

function loadNewsrvPage(){
  document.getElementById('newsrv-comment').value = '';
  document.getElementById('newsrv-msg').innerHTML = '';
  buildNewsrvForm();
  loadMyNewServers();
}

function buildNewsrvForm(){
  const grid = document.getElementById('newsrv-form-grid');
  if(!_newsrvFields.length){
    _newsrvFields = ["Server HostName","Server Type(Physical/Virtual)","Platform",
      "Server DC Location","HPC or NON HPC or JPC","Server Role","Final OS",
      "Commercial Category","Reference","Application Name"];
  }
  const selectFields = {
    "Platform": "platform", "Server DC Location": "loc",
    "HPC or NON HPC or JPC": "hpc", "Server Role": "role",
    "Final OS": "os", "Server Type(Physical/Virtual)": "stype"
  };
  fetch('/filter_options').then(r=>r.json()).then(opts=>{
    _filterOptCache = opts;
    grid.innerHTML = _newsrvFields.map(f=>{
      const fid = 'ns-' + f.replace(/[^a-zA-Z0-9]/g,'_');
      if(selectFields[f] && opts[selectFields[f]] && opts[selectFields[f]].length){
        const optList = opts[selectFields[f]].map(o=>`<option value="${esc(o)}">${esc(o)}</option>`).join('');
        // Dropdown of existing values PLUS a free-text box for a new/custom value.
        // Whichever the user fills wins — text box takes priority if both are set.
        return `<div>
          <div style="font-size:.74rem;color:var(--muted);margin-bottom:4px">${esc(f)}</div>
          <select id="${fid}_sel" onchange="onNsSelChange('${fid}')"
            style="width:100%;background:var(--surf2);border:1px solid var(--border);
            color:var(--text);padding:8px 11px;border-radius:8px;font-size:.84rem;margin-bottom:5px">
            <option value="">— Pick existing value —</option>${optList}
          </select>
          <input type="text" id="${fid}" placeholder="…or type a new value here"
            style="width:100%;background:var(--surf2);border:1px solid var(--border);
            color:var(--text);padding:7px 11px;border-radius:8px;font-size:.8rem">
        </div>`;
      }
      return `<div>
        <div style="font-size:.74rem;color:var(--muted);margin-bottom:4px">${esc(f)}</div>
        <input type="text" id="${fid}" placeholder="${esc(f)}"
          style="width:100%;background:var(--surf2);border:1px solid var(--border);
          color:var(--text);padding:8px 11px;border-radius:8px;font-size:.84rem">
      </div>`;
    }).join('');
  });
}

function onNsSelChange(fid){
  // When user picks from dropdown, copy value into the text box so submit logic stays simple
  const sel = document.getElementById(fid+'_sel');
  const txt = document.getElementById(fid);
  if(sel && txt && sel.value) txt.value = sel.value;
}

async function submitNewServer(){
  const fields = {};
  _newsrvFields.forEach(f=>{
    const fid = 'ns-' + f.replace(/[^a-zA-Z0-9]/g,'_');
    const el = document.getElementById(fid);
    if(el) fields[f] = el.value.trim();
  });
  const comment = document.getElementById('newsrv-comment').value.trim();
  const msg = document.getElementById('newsrv-msg');
  if(!fields["Server HostName"]){
    msg.innerHTML = '<div class="alert alert-warn">Server Hostname is required.</div>';
    return;
  }
  const r = await fetch('/newserver/add',{method:'POST',
    headers:{'Content-Type':'application/json'},
    body:JSON.stringify({fields,comment})}).then(x=>x.json());
  if(r.error){ msg.innerHTML=`<div class="alert alert-error">${esc(r.error)}</div>`; return; }
  msg.innerHTML = '<div class="alert alert-success">✅ Submitted! Admin will review and add it to next month\'s inventory.</div>';
  buildNewsrvForm();
  document.getElementById('newsrv-comment').value = '';
  loadMyNewServers();
}

function loadMyNewServers(){
  fetch('/admin/data').then(r=>{
    if(r.status===403){ document.getElementById('newsrv-my-list').innerHTML=''; return null; }
    return r.json();
  }).then(d=>{
    if(!d) return;
    const mine = (d.new_servers||[]).filter(x=>x.user===ME.username);
    renderMyNewsrvList(mine);
  }).catch(()=>{});
}
function renderMyNewsrvList(mine){
  const el = document.getElementById('newsrv-my-list');
  if(!mine.length){ el.innerHTML = '<div style="color:var(--dim);font-size:.82rem">No submissions yet.</div>'; return; }
  el.innerHTML = `<div class="tbl-wrap"><table>
    <thead><tr><th>Hostname</th><th>Platform</th><th>OS</th><th>Role</th><th>Comment</th><th>Submitted</th><th>Status</th><th>Admin Note</th></tr></thead>
    <tbody>${mine.map(m=>{const f=m.fields||{};return `<tr>
      <td style="font-weight:700;color:var(--blue);font-family:monospace">${esc(f["Server HostName"]||'')}</td>
      <td style="font-size:.78rem">${esc(f["Platform"]||'—')}</td>
      <td style="font-size:.78rem">${esc(f["Final OS"]||'—')}</td>
      <td style="font-size:.78rem">${esc(f["Server Role"]||'—')}</td>
      <td style="font-size:.78rem;color:var(--muted)">${esc(m.comment||'—')}</td>
      <td style="font-size:.74rem;color:var(--muted)">${esc(m.ts)}</td>
      <td>${reviewStatusBadge(m.status)}</td>
      <td style="font-size:.76rem;color:var(--muted)">${esc(m.decision_reason||'—')}</td>
    </tr>`;}).join('')}</tbody></table></div>`;
}

// ════════════════════════════════════════════════════════════════
//  UPDATE CPU / RAM (T-SHIRT SIZE)
// ════════════════════════════════════════════════════════════════
let _tshirtPlatforms = ['EMA','EPMC','JEA'];
let _tshirtSelected = null;   // {hostname, platform}

function loadTshirtPage(){
  document.getElementById('tshirt-hostname').value = '';
  document.getElementById('tshirt-host-suggest').innerHTML = '';
  document.getElementById('tshirt-host-selected').innerHTML = '';
  document.getElementById('tshirt-current-cpu').value = '';
  document.getElementById('tshirt-current-ram').value = '';
  document.getElementById('tshirt-new-cpu').value = '';
  document.getElementById('tshirt-new-ram').value = '';
  document.getElementById('tshirt-reason').value = '';
  document.getElementById('tshirt-msg').innerHTML = '';
  _tshirtSelected = null;
  loadMyTshirtChanges();
}

let _tshirtDebTimer;
function tshirtHostSearch(){
  clearTimeout(_tshirtDebTimer);
  const q = document.getElementById('tshirt-hostname').value.trim();
  if(q.length < 2){ document.getElementById('tshirt-host-suggest').innerHTML=''; return; }
  _tshirtDebTimer = setTimeout(()=>{
    fetch('/search?q='+encodeURIComponent(q)+'&page=1').then(r=>r.json()).then(d=>{
      // Only show servers on EMA / EPMC / JEA — the platforms whose pricing depends on T-shirt size
      const rows = (d.rows||[]).filter(r=>_tshirtPlatforms.includes(r['Platform'])).slice(0,6);
      if(!rows.length){
        document.getElementById('tshirt-host-suggest').innerHTML =
          `<div style="font-size:.78rem;color:var(--muted);padding:6px 0">
            No matching ${_tshirtPlatforms.join('/')} servers found. T-shirt size changes only apply to these platforms.</div>`;
        return;
      }
      document.getElementById('tshirt-host-suggest').innerHTML =
        `<div style="border:1px solid var(--border);border-radius:8px;overflow:hidden">
        ${rows.map(r=>`<div style="padding:7px 12px;cursor:pointer;font-size:.83rem;
          font-family:monospace;border-bottom:1px solid var(--border)"
          onmouseover="this.style.background='var(--surf2)'" onmouseout="this.style.background=''"
          onclick="pickTshirtHost('${esc(r['Server HostName'])}','${esc(r['Platform']||'')}')">${esc(r['Server HostName'])}
          <span style="color:var(--muted);font-family:inherit;font-size:.78rem"> — ${esc(r['Platform']||'')}</span>
        </div>`).join('')}</div>`;
    });
  }, 300);
}
function pickTshirtHost(hostname, platform){
  _tshirtSelected = {hostname, platform};
  document.getElementById('tshirt-hostname').value = hostname;
  document.getElementById('tshirt-host-suggest').innerHTML = '';
  document.getElementById('tshirt-host-selected').innerHTML =
    `<div style="font-size:.78rem;color:var(--green)">✓ Selected: ${esc(hostname)}
      <span class="badge badge-other" style="margin-left:6px">${esc(platform)}</span></div>`;
}

async function submitTshirt(){
  const msg = document.getElementById('tshirt-msg');
  if(!_tshirtSelected){
    msg.innerHTML = '<div class="alert alert-warn">Please search and select a server first.</div>';
    return;
  }
  const cc  = document.getElementById('tshirt-current-cpu').value.trim();
  const cr  = document.getElementById('tshirt-current-ram').value.trim();
  const nc  = document.getElementById('tshirt-new-cpu').value.trim();
  const nr  = document.getElementById('tshirt-new-ram').value.trim();
  const rsn = document.getElementById('tshirt-reason').value.trim();
  if(!nc && !nr){
    msg.innerHTML = '<div class="alert alert-warn">Provide a new CPU or new RAM value.</div>';
    return;
  }
  const r = await fetch('/tshirt/add',{method:'POST',
    headers:{'Content-Type':'application/json'},
    body:JSON.stringify({
      hostname:_tshirtSelected.hostname, platform:_tshirtSelected.platform,
      current_cpu:cc, current_ram:cr, new_cpu:nc, new_ram:nr, reason:rsn
    })}).then(x=>x.json());
  if(r.error){ msg.innerHTML=`<div class="alert alert-error">${esc(r.error)}</div>`; return; }
  msg.innerHTML = '<div class="alert alert-success">✅ Submitted! Admin will review and reconcile with billing.</div>';
  loadTshirtPage();
}

function loadMyTshirtChanges(){
  fetch('/admin/data').then(r=>{
    if(r.status===403){ document.getElementById('tshirt-my-list').innerHTML=''; return null; }
    return r.json();
  }).then(d=>{
    if(!d) return;
    const mine = (d.tshirt_changes||[]).filter(x=>x.user===ME.username);
    renderMyTshirtList(mine);
  }).catch(()=>{});
}
function renderMyTshirtList(mine){
  const el = document.getElementById('tshirt-my-list');
  if(!mine.length){ el.innerHTML = '<div style="color:var(--dim);font-size:.82rem">No submissions yet.</div>'; return; }
  el.innerHTML = `<div class="tbl-wrap"><table>
    <thead><tr><th>Hostname</th><th>Platform</th><th>Current CPU/RAM</th><th>New CPU/RAM</th>
      <th>Reason</th><th>Submitted</th><th>Status</th><th>Admin Note</th></tr></thead>
    <tbody>${mine.map(m=>`<tr>
      <td style="font-weight:700;color:var(--blue);font-family:monospace">${esc(m.hostname)}</td>
      <td><span class="badge badge-other">${esc(m.platform)}</span></td>
      <td style="font-size:.78rem;color:var(--red)">${esc(m.current_cpu||'—')} / ${esc(m.current_ram||'—')}</td>
      <td style="font-size:.78rem;color:var(--green);font-weight:600">${esc(m.new_cpu||'—')} / ${esc(m.new_ram||'—')}</td>
      <td style="font-size:.76rem;color:var(--muted)">${esc(m.reason||'—')}</td>
      <td style="font-size:.74rem;color:var(--muted)">${esc(m.ts)}</td>
      <td>${reviewStatusBadge(m.status)}</td>
      <td style="font-size:.76rem;color:var(--muted)">${esc(m.decision_reason||'—')}</td>
    </tr>`).join('')}</tbody></table></div>`;
}

// ════════════════════════════════════════════════════════════════
//  ACTIVITY PAGE (logged-in users)
// ════════════════════════════════════════════════════════════════
function loadActivity(){
  const el=document.getElementById('activity-out');
  if(!ME.logged_in){
    el.innerHTML='<div class="alert alert-info">Please <a href="/login">log in</a> to view your activity.</div>';
    return;
  }
  fetch('/my_activity').then(r=>r.json()).then(d=>{
    const myTags  = d.my_tags||[];
    const myNotes = d.my_notes||[];
    const myCorrs = d.my_corrections||[];
    let html=`
      <div class="metrics-row" style="max-width:600px">
        <div class="metric-card"><div class="metric-val c-purple">${d.my_tag_count}</div>
          <div class="metric-lbl">My Tags</div></div>
        <div class="metric-card"><div class="metric-val c-blue">${d.my_note_count}</div>
          <div class="metric-lbl">My Notes</div></div>
        <div class="metric-card"><div class="metric-val c-orange">${d.my_corr_count}</div>
          <div class="metric-lbl">My Corrections</div></div>
      </div>`;

    // My tags
    if(myTags.length){
      html+=`<div class="shdr" style="color:var(--purple);margin-top:16px">🏷️ My Tags</div>
        <div class="tbl-wrap"><table>
        <thead><tr><th>Hostname</th><th>Tag</th><th>Status</th><th>Submitted</th></tr></thead>
        <tbody>${myTags.map(t=>`<tr>
          <td style="font-weight:700;color:var(--blue);font-family:monospace;cursor:pointer"
            onclick="showDetail('${esc(t.hostname)}')">${esc(t.hostname)}</td>
          <td><span class="tag-pill" style="${t.status==='Accepted'?'':'border-style:dashed;opacity:.85'}">${esc(t.tag)}</span></td>
          <td>${t.status==='Accepted'
            ?'<span class="badge badge-live">✓ Accepted</span>'
            :'<span class="badge" style="background:rgba(255,194,64,.15);color:var(--yellow);border-color:rgba(255,194,64,.3)">? Unverified</span>'}</td>
          <td style="font-size:.74rem;color:var(--muted)">${esc(t.ts)}</td>
        </tr>`).join('')}</tbody></table></div>`;
    }

    // My corrections
    if(myCorrs.length){
      html+=`<div class="shdr" style="color:var(--orange);margin-top:16px">✏️ My Correction Suggestions</div>
        <div class="tbl-wrap"><table>
        <thead><tr><th>Hostname</th><th>Column</th><th>Current</th><th>Suggested</th>
          <th>Reason</th><th>Submitted</th><th>Status</th><th>Admin Note</th></tr></thead>
        <tbody>${myCorrs.map(c=>`<tr>
          <td style="font-weight:700;color:var(--blue);font-family:monospace">${esc(c.hostname)}</td>
          <td><span class="badge badge-other" style="font-size:.68rem">${esc(c.column)}</span></td>
          <td style="font-size:.78rem;color:var(--red)">${esc(c.current_val||'(blank)')}</td>
          <td style="font-size:.78rem;color:var(--green);font-weight:600">${esc(c.suggested_val)}</td>
          <td style="font-size:.75rem;color:var(--muted)">${esc(c.reason||'—')}</td>
          <td style="font-size:.74rem;color:var(--muted)">${esc(c.ts)}</td>
          <td>${reviewStatusBadge(c.status)}</td>
          <td style="font-size:.74rem;color:var(--muted)">${esc(c.decision_reason||'—')}</td>
        </tr>`).join('')}</tbody></table></div>`;
    }

    // My notes
    if(myNotes.length){
      html+=`<div class="shdr" style="color:var(--blue);margin-top:16px">📝 My Notes</div>
        <div class="tbl-wrap"><table>
        <thead><tr><th>Hostname</th><th>Note</th><th>Platform</th><th>Time</th></tr></thead>
        <tbody>${myNotes.map(n=>`<tr>
          <td style="font-weight:700;color:var(--blue);font-family:monospace;cursor:pointer"
            onclick="showDetail('${esc(n.hostname)}')">${esc(n.hostname)}</td>
          <td style="font-size:.8rem;white-space:normal">${esc(n.note)}</td>
          <td style="font-size:.78rem">${esc(n.platform)}</td>
          <td style="font-size:.74rem;color:var(--muted)">${esc(n.ts)}</td>
        </tr>`).join('')}</tbody></table></div>`;
    }

    if(!myTags.length && !myCorrs.length && !myNotes.length){
      html+=`<div class="alert alert-info" style="max-width:540px;margin-top:8px">
        You haven't submitted anything yet.</div>`;
    }

    html+=`<div class="alert alert-info" style="max-width:540px;margin-top:16px">
      💡 Open any server from <b>Search & Filter</b> to add tags, notes, corrections or flags.
      ${ME.role==='admin'?'<br>Go to <b>Admin Panel</b> for full reports.':''}
    </div>`;
    el.innerHTML=html;
  });
}

// ════════════════════════════════════════════════════════════════
//  UTILITIES
// ════════════════════════════════════════════════════════════════
function statusBadge(v){
  const l=(v||'').toLowerCase();
  if(l==='live')     return `<span class="badge badge-live">Live</span>`;
  if(l==='not live') return `<span class="badge badge-notlive">Not Live</span>`;
  return `<span class="badge badge-other">${esc(v)}</span>`;
}
function typeBadge(v){
  const l=(v||'').toLowerCase();
  if(l.includes('physical')) return `<span class="badge badge-physical">Physical</span>`;
  if(l.includes('virtual'))  return `<span class="badge badge-virtual">Virtual</span>`;
  return `<span class="badge badge-other">${esc(v)||'—'}</span>`;
}
function reviewStatusBadge(status){
  if(status==='Approved') return `<span class="badge badge-live">✅ Approved</span>`;
  if(status==='Rejected') return `<span class="badge badge-notlive">❌ Rejected</span>`;
  return `<span class="badge" style="background:rgba(255,194,64,.15);color:var(--yellow);
    border-color:rgba(255,194,64,.3)">⏳ Pending</span>`;
}
function esc(s){
  return String(s||'').replace(/&/g,'&amp;').replace(/</g,'&lt;')
    .replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}

// ════════════════════════════════════════════════════════════════
//  BOOT
// ════════════════════════════════════════════════════════════════
init();
</script>
</body></html>"""

# ═══════════════════════════════════════════════════════════════════════════════
#  ENTRY POINT
# ═══════════════════════════════════════════════════════════════════════════════
if __name__ == "__main__":
    import socket
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80)); lan_ip = s.getsockname()[0]; s.close()
    except Exception: lan_ip = "YOUR_SERVER_IP"

    port = 5050
    users = _load_users()
    print("\n" + "=" * 64)
    print("  Server Asset Inventory  v2.0  — Three-Tier Access")
    print("=" * 64)
    print(f"  Local   :  http://localhost:{port}")
    print(f"  Network :  http://{lan_ip}:{port}   ← share with team")
    print("=" * 64)
    print("  🌐 Public  : Dashboard, Search, Bulk Lookup, Compare")
    print("  🔒 Users   : + Tags, Notes, Flags  (login required)")
    print("  🔐 Admin   : + Upload, Reports, User Mgmt  (admin role)")
    print("=" * 64)
    print(f"  👥 Users          : {len(users)} account(s) in users.json")
    print(f"  📋 Audit log      : {_AUDIT_FILE.resolve()}")
    print(f"  🏷️  Tags store     : {_TAGS_FILE.resolve()}")
    print(f"  📝 Notes store    : {_NOTES_FILE.resolve()}")
    print(f"  🚩 Flags store    : {_FLAGS_FILE.resolve()}")
    if STORE["df"] is not None:
        print(f"  💾 Inventory restored from disk: {STORE['filename']} ({STORE['total_rows']} rows)")
    else:
        print("  💾 No persisted inventory found — admin needs to upload")
    print(f"  🗄️  Archive folder : {_ARCHIVE_DIR.resolve()}")
    print("  🔌 100% OFFLINE — No internet required")
    print("  Press Ctrl+C to stop")
    print("=" * 64 + "\n")
    app.run(debug=False, port=port, host="0.0.0.0", threaded=True)
