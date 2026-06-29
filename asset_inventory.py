"""
Server Asset Inventory  v1.0
==============================
Run  :  python asset_inventory.py
Open :  http://localhost:5050

100% OFFLINE — no internet required.
All charts generated server-side with matplotlib (no CDN).
No external fonts — uses system fonts only.
Single file, no external assets needed.
"""

import io, json, warnings, logging, re, base64
from datetime import datetime
import pandas as pd
import numpy as np
from flask import Flask, request, render_template_string, jsonify, send_file

warnings.filterwarnings("ignore")
logging.basicConfig(level=logging.INFO,
                    format="%(asctime)s  %(levelname)-7s  %(message)s",
                    datefmt="%H:%M:%S")
log = logging.getLogger(__name__)

# ── matplotlib offline setup ──────────────────────────────────────────────────
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches

# ── Colour palette ────────────────────────────────────────────────────────────
BG      = "#07090f"
SURFACE = "#111827"
BORDER  = "#1e2a40"
BLUE    = "#4a8cff"
RED     = "#ff4f6a"
GREEN   = "#30d988"
YELLOW  = "#ffc240"
PURPLE  = "#a78bfa"
CYAN    = "#22d3ee"
ORANGE  = "#fb923c"
MUTED   = "#7b8db0"
TEXT    = "#edf2ff"
PALETTE = [BLUE, GREEN, YELLOW, PURPLE, CYAN, ORANGE, RED,
           "#f472b6", "#34d399", "#f87171", "#60a5fa", "#a3e635"]

PLATFORM_COLORS = {
    "Dynamo":          BLUE,
    "Digital Journey": GREEN,
    "EMA":             PURPLE,
    "EPMC":            CYAN,
    "JEA":             ORANGE,
}

plt.rcParams.update({
    "figure.facecolor": BG,
    "axes.facecolor":   SURFACE,
    "axes.edgecolor":   BORDER,
    "axes.labelcolor":  MUTED,
    "xtick.color":      MUTED,
    "ytick.color":      MUTED,
    "text.color":       TEXT,
    "grid.color":       BORDER,
    "grid.linewidth":   0.6,
    "font.family":      "DejaVu Sans",
    "font.size":        9,
    "axes.titlesize":   10,
    "axes.titlecolor":  TEXT,
    "axes.titlepad":    8,
    "legend.facecolor": SURFACE,
    "legend.edgecolor": BORDER,
    "legend.fontsize":  8,
})

# ── Chart helpers ─────────────────────────────────────────────────────────────
def fig_to_b64(fig, dpi=110):
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=dpi, bbox_inches="tight",
                facecolor=fig.get_facecolor())
    buf.seek(0)
    data = base64.b64encode(buf.read()).decode()
    plt.close(fig)
    return f"data:image/png;base64,{data}"

def make_donut(labels, values, colors=None, title="", size=(4.2, 3.6)):
    if not labels or not values or sum(values) == 0:
        return None
    cols  = colors if colors else [PALETTE[i % len(PALETTE)] for i in range(len(labels))]
    total = sum(values)
    legend_patches = [mpatches.Patch(color=c, label=f"{l} ({v})")
                      for l, v, c in zip(labels, values, cols)]
    min_count  = max(1, round(total * 0.015))
    render_vals = [max(v, min_count) if v > 0 else 0 for v in values]
    nz = [(vr, c) for vr, vo, c in zip(render_vals, values, cols) if vo > 0]
    if not nz:
        return None
    nz_render, nz_cols = zip(*nz)
    nz_orig  = [v for v in values if v > 0]
    explode  = [0.06 if v < total * 0.03 else 0 for v in nz_orig]
    fig, ax  = plt.subplots(figsize=size)
    fig.patch.set_facecolor(BG)
    ax.set_facecolor(BG)
    ax.pie(nz_render, colors=nz_cols, startangle=90, explode=explode,
           wedgeprops=dict(width=0.55, edgecolor=BG, linewidth=2))
    ax.text(0, 0, str(total), ha="center", va="center",
            fontsize=15, fontweight="bold", color=TEXT)
    ax.set_title(title, color=TEXT, pad=6)
    ax.legend(handles=legend_patches, loc="lower center",
              bbox_to_anchor=(0.5, -0.22), ncol=2, framealpha=0, fontsize=7.5)
    fig.tight_layout()
    return fig_to_b64(fig)

def make_hbar(labels, values, colors=None, title="", xlabel="", size=(5.5, 0.45)):
    if not labels or not values:
        return None
    n   = len(labels)
    h   = max(2.5, n * size[1])
    fig, ax = plt.subplots(figsize=(size[0], h))
    y   = list(range(n))
    col = colors if colors else [BLUE] * n
    if isinstance(col, str):
        col = [col] * n
    bars = ax.barh(y, values, color=col, height=0.6, edgecolor=BG, linewidth=0.4)
    ax.set_yticks(y)
    ax.set_yticklabels(labels, fontsize=8.5)
    ax.invert_yaxis()
    ax.set_xlabel(xlabel, color=MUTED)
    ax.set_title(title, color=TEXT)
    ax.grid(axis="x", alpha=0.4)
    ax.spines[["top", "right", "left"]].set_visible(False)
    mx = max(values) if values else 1
    for bar, val in zip(bars, values):
        ax.text(bar.get_width() + mx * 0.01,
                bar.get_y() + bar.get_height() / 2,
                str(val), va="center", fontsize=7.5, color=TEXT)
    fig.tight_layout()
    return fig_to_b64(fig)

def make_vbar(labels, values, colors=None, title="", ylabel="", size=(7, 3.2)):
    if not labels or not values:
        return None
    fig, ax = plt.subplots(figsize=size)
    x    = list(range(len(labels)))
    cols = colors if colors else [BLUE] * len(labels)
    ax.bar(x, values, color=cols, edgecolor=BG, linewidth=0.4, width=0.65)
    ax.set_xticks(x)
    ax.set_xticklabels(labels, rotation=30, ha="right", fontsize=8)
    ax.set_ylabel(ylabel, color=MUTED)
    ax.set_title(title, color=TEXT)
    ax.grid(axis="y", alpha=0.4)
    ax.spines[["top", "right"]].set_visible(False)
    fig.tight_layout()
    return fig_to_b64(fig)

def make_hbar_compare(labels, vals_a, vals_b, label_a="Month A", label_b="Month B",
                      title="", size=(6, 0.5)):
    """Side-by-side horizontal bars for month comparison."""
    if not labels:
        return None
    n   = len(labels)
    h   = max(3, n * size[1])
    fig, ax = plt.subplots(figsize=(size[0], h))
    y   = np.arange(n)
    bh  = 0.35
    ax.barh(y + bh/2, vals_a, bh, color=BLUE,  label=label_a, edgecolor=BG, linewidth=0.3)
    ax.barh(y - bh/2, vals_b, bh, color=GREEN, label=label_b, edgecolor=BG, linewidth=0.3)
    ax.set_yticks(y)
    ax.set_yticklabels(labels, fontsize=8)
    ax.invert_yaxis()
    ax.set_title(title, color=TEXT)
    ax.grid(axis="x", alpha=0.35)
    ax.spines[["top", "right", "left"]].set_visible(False)
    ax.legend(framealpha=0.2, fontsize=8)
    fig.tight_layout()
    return fig_to_b64(fig)

# ── Column detection ──────────────────────────────────────────────────────────
COL_MAP = {
    "Server HostName":            ["server hostname", "hostname", "host name",
                                   "server name", "servername"],
    "Server Type(Physical/Virtual)": ["server type", "server type(physical/virtual)",
                                      "type", "physical/virtual", "server type (physical/virtual)"],
    "Platform":                   ["platform", "infra", "infrastructure"],
    "Server DC Location":         ["server dc location", "dc location", "datacenter",
                                   "data center", "location", "server dc location :"],
    "HPC or NON HPC or JPC":      ["hpc or non hpc or jpc", "hpc/non hpc/jpc",
                                   "hpc", "group", "hpc or non hpc or jpc:"],
    "Server Role":                ["server role", "role", "server role :"],
    "Final OS":                   ["final os", "os", "operating system", "os version"],
    "Commercial Category":        ["commercial category", "category", "contract type",
                                   "commercial category :"],
    "Reference":                  ["reference", "ref", "reference :"],
    "Application Name":           ["application name", "application", "app name",
                                   "app", "application name :"],
}

STATUS_COL_PATTERN = re.compile(
    r"status\s+as\s+on\s+(?:1st|1)\s+(\w+)\s+(\d{4})", re.IGNORECASE
)

def detect_status_col(columns):
    """Find the monthly status column e.g. 'Status as on 1st June 2026'."""
    best = None
    for col in columns:
        m = STATUS_COL_PATTERN.search(col.strip())
        if m:
            best = col
    return best

def normalise_columns(df):
    df.columns = [str(c).strip() for c in df.columns]
    rename = {}
    used   = set()
    for col in df.columns:
        key = col.lower().strip()
        for canon, aliases in COL_MAP.items():
            if canon in used:
                continue
            if key in aliases or key == canon.lower():
                rename[col] = canon
                used.add(canon)
                break
    if rename:
        log.info("Column rename map: %s", rename)
    return df.rename(columns=rename)

# ── Flask app + in-memory store ───────────────────────────────────────────────
app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024   # 50 MB

STORE = {
    "df":          None,   # current month dataframe
    "df_prev":     None,   # previous month dataframe
    "status_col":  None,
    "prev_status": None,
    "filename":    None,
    "prev_file":   None,
    "uploaded_at": None,
    "total_rows":  0,
    "quality":     {},
}

# ── Data loading ──────────────────────────────────────────────────────────────
def load_excel(file_bytes, filename):
    ext = filename.rsplit(".", 1)[-1].lower()
    if ext in ("xlsx", "xls"):
        df = pd.read_excel(io.BytesIO(file_bytes), dtype=str)
    elif ext == "csv":
        df = pd.read_csv(io.BytesIO(file_bytes), dtype=str)
    else:
        raise ValueError("Unsupported file type")
    df = df.fillna("").astype(str)
    df = df[~df.apply(lambda r: r.str.strip().eq("").all(), axis=1)]
    df = normalise_columns(df)
    return df

def quality_check(df, status_col):
    issues = {}
    key_cols = ["Application Name", "Commercial Category",
                "Server DC Location", "Final OS", "Platform"]
    for c in key_cols:
        if c in df.columns:
            blank = int((df[c].str.strip() == "").sum())
            if blank:
                issues[c] = blank
    if status_col and status_col in df.columns:
        blank_st = int((df[status_col].str.strip() == "").sum())
        if blank_st:
            issues["Status (blank)"] = blank_st
    return issues

# ── Analysis helpers ──────────────────────────────────────────────────────────
def vc(series):
    """Value counts → (labels, values) lists, dropping blanks."""
    s = series.replace("", np.nan).dropna()
    v = s.value_counts()
    return v.index.tolist(), v.values.tolist()

def safe_col(df, col):
    return df[col] if col in df.columns else pd.Series(dtype=str)

def analyse(df, status_col):
    total = len(df)
    results = {"total": total}

    # Status summary
    if status_col and status_col in df.columns:
        st = df[status_col].str.strip()
        results["live"]     = int((st.str.lower() == "live").sum())
        results["not_live"] = int((st.str.lower() == "not live").sum())
        results["status_other"] = total - results["live"] - results["not_live"]
    else:
        results["live"] = results["not_live"] = results["status_other"] = 0

    # Server type
    if "Server Type(Physical/Virtual)" in df.columns:
        st2 = df["Server Type(Physical/Virtual)"].str.strip().str.lower()
        results["physical"] = int(st2.str.contains("physical|bare", na=False).sum())
        results["virtual"]  = int(st2.str.contains("virtual|vm",    na=False).sum())
    else:
        results["physical"] = results["virtual"] = 0

    # Platform
    if "Platform" in df.columns:
        pl_l, pl_v = vc(df["Platform"].str.strip())
        results["platform_labels"] = pl_l
        results["platform_values"] = pl_v
        results["platform_colors"] = [PLATFORM_COLORS.get(p, MUTED) for p in pl_l]
    else:
        results["platform_labels"] = results["platform_values"] = results["platform_colors"] = []

    # DC Location
    loc_l, loc_v = vc(safe_col(df, "Server DC Location").str.strip())
    results["loc_labels"] = loc_l
    results["loc_values"] = loc_v

    # HPC group
    hpc_l, hpc_v = vc(safe_col(df, "HPC or NON HPC or JPC").str.strip())
    results["hpc_labels"] = hpc_l
    results["hpc_values"] = hpc_v

    # Server Role
    role_l, role_v = vc(safe_col(df, "Server Role").str.strip())
    results["role_labels"] = role_l
    results["role_values"] = role_v

    # OS
    os_l, os_v = vc(safe_col(df, "Final OS").str.strip())
    results["os_labels"] = os_l[:15]
    results["os_values"] = os_v[:15]

    # Application
    app_l, app_v = vc(safe_col(df, "Application Name").str.strip())
    results["app_labels"] = app_l[:15]
    results["app_values"] = app_v[:15]

    return results

def compare_months(df_curr, df_prev, sc_curr, sc_prev):
    """Return comparison dict between two monthly uploads."""
    out = {}
    hn  = "Server HostName"

    # --- totals ---
    out["curr_total"] = len(df_curr)
    out["prev_total"] = len(df_prev)
    out["diff_total"] = out["curr_total"] - out["prev_total"]

    # --- new / removed servers ---
    curr_hosts = set(df_curr[hn].str.strip().str.lower()) if hn in df_curr.columns else set()
    prev_hosts = set(df_prev[hn].str.strip().str.lower()) if hn in df_prev.columns else set()
    new_hosts  = curr_hosts - prev_hosts
    gone_hosts = prev_hosts - curr_hosts
    out["new_count"]     = len(new_hosts)
    out["removed_count"] = len(gone_hosts)

    if hn in df_curr.columns:
        out["new_servers"] = df_curr[df_curr[hn].str.strip().str.lower().isin(new_hosts)][
            [c for c in [hn, "Platform", "Server Role", "Final OS", "Server DC Location",
                         "HPC or NON HPC or JPC"] if c in df_curr.columns]
        ].fillna("").to_dict("records")
    else:
        out["new_servers"] = []

    if hn in df_prev.columns:
        out["removed_servers"] = df_prev[df_prev[hn].str.strip().str.lower().isin(gone_hosts)][
            [c for c in [hn, "Platform", "Server Role", "Final OS", "Server DC Location",
                         "HPC or NON HPC or JPC"] if c in df_prev.columns]
        ].fillna("").to_dict("records")
    else:
        out["removed_servers"] = []

    # --- status changes (servers in both months) ---
    if sc_curr and sc_prev and hn in df_curr.columns and hn in df_prev.columns:
        common = curr_hosts & prev_hosts
        curr_st = df_curr.set_index(df_curr[hn].str.strip().str.lower())
        prev_st = df_prev.set_index(df_prev[hn].str.strip().str.lower())

        to_live     = []
        to_not_live = []
        for h in common:
            try:
                cs = str(curr_st.loc[h, sc_curr]).strip().lower() if h in curr_st.index else ""
                ps = str(prev_st.loc[h, sc_prev]).strip().lower() if h in prev_st.index else ""
                if isinstance(cs, pd.Series): cs = cs.iloc[0]
                if isinstance(ps, pd.Series): ps = ps.iloc[0]
                if ps != "live" and cs == "live":
                    row = curr_st.loc[h]
                    if isinstance(row, pd.DataFrame): row = row.iloc[0]
                    to_live.append({hn: row.get(hn, h),
                                    "Platform": row.get("Platform", ""),
                                    "Server Role": row.get("Server Role", ""),
                                    "Prev Status": ps.title(), "Curr Status": cs.title()})
                elif ps == "live" and cs != "live":
                    row = curr_st.loc[h]
                    if isinstance(row, pd.DataFrame): row = row.iloc[0]
                    to_not_live.append({hn: row.get(hn, h),
                                        "Platform": row.get("Platform", ""),
                                        "Server Role": row.get("Server Role", ""),
                                        "Prev Status": ps.title(), "Curr Status": cs.title()})
            except Exception:
                continue

        out["to_live"]     = to_live
        out["to_not_live"] = to_not_live
    else:
        out["to_live"] = out["to_not_live"] = []

    # --- platform comparison chart ---
    if "Platform" in df_curr.columns and "Platform" in df_prev.columns:
        all_platforms = sorted(
            set(df_curr["Platform"].str.strip().unique()) |
            set(df_prev["Platform"].str.strip().unique())
        )
        all_platforms = [p for p in all_platforms if p]
        curr_plat = df_curr["Platform"].str.strip().value_counts()
        prev_plat = df_prev["Platform"].str.strip().value_counts()
        out["plat_labels"] = all_platforms
        out["plat_curr"]   = [int(curr_plat.get(p, 0)) for p in all_platforms]
        out["plat_prev"]   = [int(prev_plat.get(p, 0)) for p in all_platforms]
    else:
        out["plat_labels"] = out["plat_curr"] = out["plat_prev"] = []

    # --- live count comparison per platform ---
    if sc_curr and sc_prev and "Platform" in df_curr.columns and "Platform" in df_prev.columns:
        all_platforms = out.get("plat_labels", [])
        curr_live = df_curr[df_curr[sc_curr].str.strip().str.lower() == "live"]["Platform"].str.strip().value_counts()
        prev_live = df_prev[df_prev[sc_prev].str.strip().str.lower() == "live"]["Platform"].str.strip().value_counts()
        out["live_curr"] = [int(curr_live.get(p, 0)) for p in all_platforms]
        out["live_prev"] = [int(prev_live.get(p, 0)) for p in all_platforms]
    else:
        out["live_curr"] = out["live_prev"] = []

    return out

# ── Flask routes ──────────────────────────────────────────────────────────────
@app.route("/")
def index():
    return render_template_string(HTML_TEMPLATE)

@app.route("/upload", methods=["POST"])
def upload():
    which = request.form.get("which", "current")   # "current" or "prev"
    f     = request.files.get("file")
    if not f:
        return jsonify({"error": "No file provided"}), 400
    try:
        raw  = f.read()
        df   = load_excel(raw, f.filename)
        scol = detect_status_col(df.columns.tolist())
        if which == "prev":
            STORE["df_prev"]    = df
            STORE["prev_status"] = scol
            STORE["prev_file"]   = f.filename
            return jsonify({
                "ok":         True,
                "filename":   f.filename,
                "rows":       len(df),
                "status_col": scol or "Not detected",
            })
        else:
            qc = quality_check(df, scol)
            STORE.update({
                "df":          df,
                "status_col":  scol,
                "filename":    f.filename,
                "uploaded_at": datetime.now().strftime("%d %b %Y %H:%M"),
                "total_rows":  len(df),
                "quality":     qc,
            })
            return jsonify({
                "ok":         True,
                "filename":   f.filename,
                "rows":       len(df),
                "status_col": scol or "Not detected",
                "quality":    qc,
                "columns":    df.columns.tolist(),
            })
    except Exception as e:
        log.exception("Upload error")
        return jsonify({"error": str(e)}), 500

@app.route("/dashboard")
def dashboard():
    df = STORE["df"]
    if df is None:
        return jsonify({"error": "No data loaded"})
    sc  = STORE["status_col"]
    res = analyse(df, sc)

    charts = {}
    # Donut: Live vs Not Live
    if sc:
        charts["status_donut"] = make_donut(
            ["Live", "Not Live", "Other"],
            [res["live"], res["not_live"], res["status_other"]],
            colors=[GREEN, RED, MUTED],
            title="Live vs Not Live"
        )
    # Donut: Physical vs Virtual
    charts["type_donut"] = make_donut(
        ["Physical", "Virtual"],
        [res["physical"], res["virtual"]],
        colors=[BLUE, PURPLE],
        title="Physical vs Virtual"
    )
    # Hbar: Platform
    if res["platform_labels"]:
        charts["platform_hbar"] = make_hbar(
            res["platform_labels"], res["platform_values"],
            colors=res["platform_colors"], title="Servers by Platform"
        )
    # Hbar: DC Location
    if res["loc_labels"]:
        charts["loc_hbar"] = make_hbar(
            res["loc_labels"], res["loc_values"],
            colors=[CYAN]*len(res["loc_labels"]), title="Servers by DC Location"
        )
    # Hbar: HPC Group
    if res["hpc_labels"]:
        charts["hpc_hbar"] = make_hbar(
            res["hpc_labels"], res["hpc_values"],
            colors=[YELLOW]*len(res["hpc_labels"]), title="HPC / NON-HPC / JPC"
        )
    # Vbar: OS
    if res["os_labels"]:
        charts["os_vbar"] = make_vbar(
            res["os_labels"], res["os_values"],
            colors=[ORANGE]*len(res["os_labels"]), title="Servers by OS (Top 15)"
        )
    # Vbar: Server Role
    if res["role_labels"]:
        charts["role_vbar"] = make_vbar(
            res["role_labels"], res["role_values"],
            colors=[PURPLE]*len(res["role_labels"]), title="Servers by Role"
        )
    # Hbar: Top Applications
    if res["app_labels"]:
        charts["app_hbar"] = make_hbar(
            res["app_labels"], res["app_values"],
            colors=[BLUE]*len(res["app_labels"]), title="Top 15 Applications"
        )

    return jsonify({
        "metrics": {
            "total":    res["total"],
            "live":     res["live"],
            "not_live": res["not_live"],
            "physical": res["physical"],
            "virtual":  res["virtual"],
        },
        "charts":      charts,
        "status_col":  sc or "",
        "filename":    STORE["filename"] or "",
        "uploaded_at": STORE["uploaded_at"] or "",
        "quality":     STORE["quality"],
    })

@app.route("/search")
def search():
    df = STORE["df"]
    if df is None:
        return jsonify({"rows": [], "total": 0})
    sc = STORE["status_col"]

    q        = request.args.get("q", "").strip().lower()
    platform = request.args.get("platform", "")
    status   = request.args.get("status", "")
    role     = request.args.get("role", "")
    hpc      = request.args.get("hpc", "")
    loc      = request.args.get("loc", "")
    os_f     = request.args.get("os", "")
    stype    = request.args.get("stype", "")
    page     = int(request.args.get("page", 1))
    per_page = 50

    filt = df.copy()

    # text search across hostname + application
    if q:
        mask = pd.Series(False, index=filt.index)
        for col in ["Server HostName", "Application Name", "Reference"]:
            if col in filt.columns:
                mask |= filt[col].str.lower().str.contains(q, na=False)
        filt = filt[mask]

    # dropdown filters
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
        filt = filt[filt["Server Type(Physical/Virtual)"].str.strip().str.lower()
                    .str.contains(stype.lower(), na=False)]

    total = len(filt)
    start = (page - 1) * per_page
    page_df = filt.iloc[start:start + per_page]

    display_cols = [c for c in [
        "Server HostName", "Server Type(Physical/Virtual)", "Platform",
        sc if sc else None, "Server DC Location", "HPC or NON HPC or JPC",
        "Server Role", "Final OS", "Application Name", "Commercial Category", "Reference"
    ] if c and c in page_df.columns]

    rows = page_df[display_cols].fillna("").to_dict("records")
    return jsonify({"rows": rows, "total": total, "page": page,
                    "pages": max(1, -(-total // per_page)),
                    "status_col": sc or ""})

@app.route("/detail")
def detail():
    df  = STORE["df"]
    sc  = STORE["status_col"]
    hn  = request.args.get("hostname", "").strip()
    if df is None or not hn or "Server HostName" not in df.columns:
        return jsonify({})
    row = df[df["Server HostName"].str.strip().str.lower() == hn.lower()]
    if row.empty:
        return jsonify({})
    r = row.iloc[0].to_dict()
    if sc:
        r["_status_col"] = sc
        r["_status_val"] = r.get(sc, "")
    return jsonify(r)

@app.route("/bulk", methods=["POST"])
def bulk():
    df = STORE["df"]
    sc = STORE["status_col"]
    if df is None:
        return jsonify({"found": [], "not_found": [], "total_found": 0})
    data  = request.get_json(force=True)
    names = [n.strip() for n in data.get("names", []) if n.strip()]
    if not names or "Server HostName" not in df.columns:
        return jsonify({"found": [], "not_found": names, "total_found": 0})

    lower_names = [n.lower() for n in names]
    df["_hn_lower"] = df["Server HostName"].str.strip().str.lower()
    found_df   = df[df["_hn_lower"].isin(lower_names)]
    found_set  = set(found_df["_hn_lower"].tolist())
    not_found  = [n for n in names if n.lower() not in found_set]

    display_cols = [c for c in [
        "Server HostName", "Server Type(Physical/Virtual)", "Platform",
        sc if sc else None, "Server DC Location", "HPC or NON HPC or JPC",
        "Server Role", "Final OS", "Application Name", "Commercial Category"
    ] if c and c in found_df.columns]

    df.drop(columns=["_hn_lower"], inplace=True)
    rows = found_df[display_cols].fillna("").to_dict("records")
    return jsonify({"found": rows, "not_found": not_found,
                    "total_found": len(rows), "status_col": sc or ""})

@app.route("/compare")
def compare():
    df_c = STORE["df"]
    df_p = STORE["df_prev"]
    if df_c is None or df_p is None:
        return jsonify({"error": "Upload both months first"})
    sc_c = STORE["status_col"]
    sc_p = STORE["prev_status"]
    try:
        out    = compare_months(df_c, df_p, sc_c, sc_p)
        charts = {}
        if out.get("plat_labels"):
            charts["plat_compare"] = make_hbar_compare(
                out["plat_labels"], out["plat_curr"], out["plat_prev"],
                label_a=STORE["filename"] or "Current",
                label_b=STORE["prev_file"] or "Previous",
                title="Total Servers by Platform — Month Comparison"
            )
        if out.get("plat_labels") and out.get("live_curr"):
            charts["live_compare"] = make_hbar_compare(
                out["plat_labels"], out["live_curr"], out["live_prev"],
                label_a=STORE["filename"] or "Current",
                label_b=STORE["prev_file"] or "Previous",
                title="Live Servers by Platform — Month Comparison"
            )
        out["charts"] = charts
        out["curr_file"] = STORE["filename"] or "Current"
        out["prev_file"] = STORE["prev_file"] or "Previous"
        return jsonify(out)
    except Exception as e:
        log.exception("Compare error")
        return jsonify({"error": str(e)})

@app.route("/filter_options")
def filter_options():
    df = STORE["df"]
    if df is None:
        return jsonify({})
    sc = STORE["status_col"]
    def opts(col):
        if col not in df.columns: return []
        return sorted(df[col].str.strip().replace("", np.nan).dropna().unique().tolist())
    status_opts = []
    if sc and sc in df.columns:
        status_opts = sorted(df[sc].str.strip().replace("", np.nan).dropna().unique().tolist())
    return jsonify({
        "platform": opts("Platform"),
        "status":   status_opts,
        "role":     opts("Server Role"),
        "hpc":      opts("HPC or NON HPC or JPC"),
        "loc":      opts("Server DC Location"),
        "os":       opts("Final OS"),
        "stype":    opts("Server Type(Physical/Virtual)"),
    })

@app.route("/export")
def export():
    df = STORE["df"]
    if df is None:
        return "No data", 400
    sc       = STORE["status_col"]
    platform = request.args.get("platform", "")
    status   = request.args.get("status", "")
    role     = request.args.get("role", "")
    q        = request.args.get("q", "").strip().lower()

    filt = df.copy()
    if q:
        mask = pd.Series(False, index=filt.index)
        for col in ["Server HostName", "Application Name"]:
            if col in filt.columns:
                mask |= filt[col].str.lower().str.contains(q, na=False)
        filt = filt[mask]
    if platform and "Platform" in filt.columns:
        filt = filt[filt["Platform"].str.strip() == platform]
    if status and sc and sc in filt.columns:
        filt = filt[filt[sc].str.strip().str.lower() == status.lower()]
    if role and "Server Role" in filt.columns:
        filt = filt[filt["Server Role"].str.strip() == role]

    buf  = io.BytesIO()
    filt.drop(columns=[c for c in filt.columns if c.startswith("_")], errors="ignore") \
        .to_excel(buf, index=False)
    buf.seek(0)
    fname = f"inventory_export_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
    return send_file(buf, as_attachment=True, download_name=fname,
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

@app.route("/export_bulk", methods=["POST"])
def export_bulk():
    data  = request.get_json(force=True)
    rows  = data.get("rows", [])
    if not rows:
        return "No data", 400
    df_out = pd.DataFrame(rows)
    buf    = io.BytesIO()
    df_out.to_excel(buf, index=False)
    buf.seek(0)
    fname = f"bulk_lookup_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
    return send_file(buf, as_attachment=True, download_name=fname,
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ── HTML template ─────────────────────────────────────────────────────────────
HTML_TEMPLATE = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Server Asset Inventory</title>
<style>
/* ── CSS variables — dark / light ── */
:root{
  --bg:#07090f; --surf:#111827; --surf2:#161f30; --border:#1e2a40; --border2:#243452;
  --text:#edf2ff; --muted:#7b8db0; --dim:#3d5070;
  --blue:#4a8cff; --green:#30d988; --yellow:#ffc240; --red:#ff4f6a;
  --purple:#a78bfa; --cyan:#22d3ee; --orange:#fb923c;
  --radius:10px; --shadow:0 4px 24px rgba(0,0,0,.45);
}
body.light{
  --bg:#f0f4fb; --surf:#ffffff; --surf2:#e8edf7; --border:#d0d8ea; --border2:#b8c4da;
  --text:#1a2340; --muted:#5a6a8a; --dim:#a0aec0;
  --blue:#2563eb; --green:#16a34a; --yellow:#d97706; --red:#dc2626;
  --purple:#7c3aed; --cyan:#0891b2; --orange:#ea580c;
}
*{box-sizing:border-box;margin:0;padding:0}
body{background:var(--bg);color:var(--text);font-family:'DejaVu Sans',system-ui,sans-serif;
  font-size:14px;min-height:100vh;transition:background .25s,color .25s}
a{color:var(--blue);text-decoration:none}

/* ── Layout ── */
.topbar{background:var(--surf);border-bottom:1px solid var(--border);
  padding:0 24px;display:flex;align-items:center;gap:16px;height:54px;
  position:sticky;top:0;z-index:100}
.topbar-title{font-size:1.05rem;font-weight:700;color:var(--blue);flex:1}
.topbar-meta{font-size:.75rem;color:var(--muted)}
.theme-btn{background:var(--surf2);border:1px solid var(--border);color:var(--text);
  padding:5px 12px;border-radius:6px;cursor:pointer;font-size:.8rem;font-weight:600}
.theme-btn:hover{border-color:var(--blue)}

.sidebar{width:220px;background:var(--surf);border-right:1px solid var(--border);
  position:fixed;top:54px;left:0;bottom:0;overflow-y:auto;padding:16px 0}
.nav-item{display:flex;align-items:center;gap:10px;padding:9px 20px;
  cursor:pointer;color:var(--muted);font-size:.88rem;font-weight:500;
  border-left:3px solid transparent;transition:all .15s}
.nav-item:hover{background:var(--surf2);color:var(--text)}
.nav-item.active{color:var(--blue);border-left-color:var(--blue);background:var(--surf2)}
.nav-icon{font-size:1rem;width:20px;text-align:center}
.nav-section{padding:14px 20px 6px;font-size:.68rem;font-weight:700;
  text-transform:uppercase;letter-spacing:.1em;color:var(--dim)}

.main{margin-left:220px;padding:24px;margin-top:54px;min-height:calc(100vh - 54px)}
.page{display:none}.page.active{display:block}

/* ── Upload zone ── */
.upload-zone{border:2px dashed var(--border);border-radius:var(--radius);
  padding:36px;text-align:center;cursor:pointer;transition:all .2s;
  background:var(--surf)}
.upload-zone:hover,.upload-zone.drag{border-color:var(--blue);background:var(--surf2)}
.upload-zone input{display:none}
.upload-label{font-size:.9rem;color:var(--muted);margin-top:8px}
.upload-icon{font-size:2.5rem}

/* ── Cards / metrics ── */
.metrics-row{display:grid;grid-template-columns:repeat(auto-fit,minmax(140px,1fr));gap:14px;margin-bottom:22px}
.metric-card{background:var(--surf);border:1px solid var(--border);border-radius:var(--radius);
  padding:16px 18px;text-align:center}
.metric-val{font-size:2rem;font-weight:700;font-family:monospace;letter-spacing:-1px}
.metric-lbl{font-size:.72rem;font-weight:600;text-transform:uppercase;letter-spacing:.1em;
  color:var(--muted);margin-top:4px}
.c-blue{color:var(--blue)}.c-green{color:var(--green)}.c-red{color:var(--red)}
.c-purple{color:var(--purple)}.c-cyan{color:var(--cyan)}.c-yellow{color:var(--yellow)}
.c-orange{color:var(--orange)}

/* ── Charts grid ── */
.charts-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(320px,1fr));gap:18px}
.chart-card{background:var(--surf);border:1px solid var(--border);border-radius:var(--radius);
  padding:14px;overflow:hidden}
.chart-card img{width:100%;height:auto;border-radius:6px}

/* ── Section header ── */
.section-hdr{font-size:.8rem;font-weight:700;text-transform:uppercase;letter-spacing:.1em;
  color:var(--blue);border-bottom:1px solid var(--border);padding-bottom:8px;margin-bottom:14px}

/* ── Filters ── */
.filter-bar{background:var(--surf);border:1px solid var(--border);border-radius:var(--radius);
  padding:14px 18px;margin-bottom:16px;display:flex;flex-wrap:wrap;gap:10px;align-items:flex-end}
.filter-bar input,.filter-bar select{background:var(--surf2);border:1px solid var(--border);
  color:var(--text);padding:7px 10px;border-radius:7px;font-size:.83rem;min-width:140px}
.filter-bar input:focus,.filter-bar select:focus{outline:none;border-color:var(--blue)}
.filter-bar label{font-size:.73rem;color:var(--muted);display:block;margin-bottom:3px}
.btn{background:var(--blue);color:#fff;border:none;padding:7px 16px;border-radius:7px;
  cursor:pointer;font-size:.83rem;font-weight:600;transition:opacity .15s}
.btn:hover{opacity:.85}
.btn-sm{padding:4px 11px;font-size:.76rem}
.btn-outline{background:transparent;border:1px solid var(--border);color:var(--muted)}
.btn-outline:hover{border-color:var(--blue);color:var(--blue)}
.btn-green{background:var(--green);color:#000}
.btn-red{background:var(--red)}
.btn-yellow{background:var(--yellow);color:#000}

/* ── Quick filter chips ── */
.chip-row{display:flex;flex-wrap:wrap;gap:8px;margin-bottom:14px}
.chip{padding:5px 14px;border-radius:99px;border:1px solid var(--border);
  background:var(--surf2);color:var(--muted);cursor:pointer;font-size:.8rem;
  font-weight:600;transition:all .15s}
.chip:hover{border-color:var(--blue);color:var(--blue)}
.chip.active{background:var(--blue);color:#fff;border-color:var(--blue)}
.chip-dynamo.active{background:#4a8cff}
.chip-dj.active{background:#30d988;color:#000}
.chip-ema.active{background:#a78bfa}
.chip-epmc.active{background:#22d3ee;color:#000}
.chip-jea.active{background:#fb923c}
.chip-hpc.active{background:#ffc240;color:#000}
.chip-jpc.active{background:#f472b6}

/* ── Table ── */
.tbl-wrap{overflow-x:auto;border-radius:var(--radius);border:1px solid var(--border)}
table{width:100%;border-collapse:collapse;font-size:.8rem}
th{background:var(--surf2);padding:9px 12px;text-align:left;font-size:.72rem;
  font-weight:700;text-transform:uppercase;letter-spacing:.06em;color:var(--muted);
  border-bottom:1px solid var(--border);white-space:nowrap}
td{padding:8px 12px;border-bottom:1px solid var(--border);color:var(--text)}
tr:last-child td{border-bottom:none}
tr:hover td{background:var(--surf2)}
.tbl-count{font-size:.78rem;color:var(--muted);margin-bottom:8px}
.pagination{display:flex;gap:6px;margin-top:12px;align-items:center;flex-wrap:wrap}
.page-btn{background:var(--surf2);border:1px solid var(--border);color:var(--muted);
  padding:4px 10px;border-radius:6px;cursor:pointer;font-size:.78rem}
.page-btn:hover,.page-btn.active{background:var(--blue);color:#fff;border-color:var(--blue)}

/* ── Status badge ── */
.badge{display:inline-block;padding:2px 9px;border-radius:99px;font-size:.72rem;font-weight:700}
.badge-live{background:rgba(48,217,136,.15);color:var(--green);border:1px solid rgba(48,217,136,.3)}
.badge-notlive{background:rgba(255,79,106,.15);color:var(--red);border:1px solid rgba(255,79,106,.3)}
.badge-physical{background:rgba(74,140,255,.15);color:var(--blue);border:1px solid rgba(74,140,255,.3)}
.badge-virtual{background:rgba(167,139,250,.15);color:var(--purple);border:1px solid rgba(167,139,250,.3)}
.badge-other{background:var(--surf2);color:var(--muted);border:1px solid var(--border)}

/* ── Detail modal ── */
.modal-overlay{display:none;position:fixed;inset:0;background:rgba(0,0,0,.65);
  z-index:500;align-items:center;justify-content:center}
.modal-overlay.show{display:flex}
.modal{background:var(--surf);border:1px solid var(--border2);border-radius:14px;
  padding:28px;max-width:680px;width:94%;max-height:85vh;overflow-y:auto;
  box-shadow:var(--shadow)}
.modal-title{font-size:1.15rem;font-weight:700;color:var(--blue);margin-bottom:18px;
  display:flex;justify-content:space-between;align-items:center}
.modal-close{cursor:pointer;color:var(--muted);font-size:1.4rem;line-height:1}
.modal-close:hover{color:var(--red)}
.detail-grid{display:grid;grid-template-columns:1fr 1fr;gap:12px}
.detail-row{background:var(--surf2);border-radius:8px;padding:10px 14px}
.detail-key{font-size:.7rem;text-transform:uppercase;letter-spacing:.08em;color:var(--muted);margin-bottom:3px}
.detail-val{font-size:.88rem;font-weight:600;color:var(--text);word-break:break-word}
.detail-full{grid-column:1/-1}

/* ── Bulk lookup ── */
.bulk-area{width:100%;min-height:120px;background:var(--surf2);border:1px solid var(--border);
  color:var(--text);padding:12px;border-radius:8px;font-family:monospace;font-size:.85rem;resize:vertical}
.bulk-area:focus{outline:none;border-color:var(--blue)}
.not-found-list{background:rgba(255,79,106,.08);border:1px solid rgba(255,79,106,.25);
  border-radius:8px;padding:12px 16px;margin-top:12px}
.not-found-list h4{color:var(--red);font-size:.82rem;margin-bottom:8px}
.not-found-list .item{font-family:monospace;font-size:.8rem;color:var(--muted);padding:2px 0}

/* ── Compare ── */
.compare-uploads{display:grid;grid-template-columns:1fr 1fr;gap:16px;margin-bottom:20px}
.compare-box{background:var(--surf);border:1px solid var(--border);border-radius:var(--radius);padding:18px}
.compare-box h4{font-size:.85rem;font-weight:700;color:var(--muted);margin-bottom:12px}
.diff-cards{display:grid;grid-template-columns:repeat(auto-fit,minmax(140px,1fr));gap:12px;margin-bottom:20px}
.diff-card{background:var(--surf);border:1px solid var(--border);border-radius:var(--radius);
  padding:14px 16px;text-align:center}
.diff-val{font-size:1.7rem;font-weight:700;font-family:monospace}
.diff-lbl{font-size:.72rem;color:var(--muted);margin-top:4px;text-transform:uppercase;letter-spacing:.07em}

/* ── Quality panel ── */
.quality-panel{background:rgba(255,194,64,.07);border:1px solid rgba(255,194,64,.25);
  border-radius:8px;padding:12px 16px;margin-bottom:16px}
.quality-panel h4{color:var(--yellow);font-size:.82rem;margin-bottom:8px}
.quality-item{font-size:.8rem;color:var(--muted);padding:2px 0}
.quality-item span{color:var(--yellow);font-weight:700;font-family:monospace}

/* ── Alert / info ── */
.alert{border-radius:8px;padding:12px 16px;font-size:.85rem;margin-bottom:14px}
.alert-info{background:rgba(74,140,255,.1);border:1px solid rgba(74,140,255,.3);color:var(--blue)}
.alert-success{background:rgba(48,217,136,.1);border:1px solid rgba(48,217,136,.3);color:var(--green)}
.alert-warn{background:rgba(255,194,64,.1);border:1px solid rgba(255,194,64,.3);color:var(--yellow)}

/* ── Loading ── */
.spinner{display:inline-block;width:18px;height:18px;border:2px solid var(--border);
  border-top-color:var(--blue);border-radius:50%;animation:spin .7s linear infinite}
@keyframes spin{to{transform:rotate(360deg)}}
.loading-row{text-align:center;padding:32px;color:var(--muted)}

/* ── Responsive ── */
@media(max-width:768px){
  .sidebar{display:none}
  .main{margin-left:0}
  .compare-uploads{grid-template-columns:1fr}
  .detail-grid{grid-template-columns:1fr}
}
</style>
</head>
<body>

<!-- Top bar -->
<div class="topbar">
  <div class="topbar-title">🖥️ Server Asset Inventory</div>
  <div class="topbar-meta" id="topbar-meta">No data loaded</div>
  <button class="theme-btn" onclick="toggleTheme()">🌙 / ☀️ Theme</button>
</div>

<!-- Sidebar -->
<div class="sidebar">
  <div class="nav-section">Main</div>
  <div class="nav-item active" onclick="showPage('upload')">
    <span class="nav-icon">📤</span> Upload Data
  </div>
  <div class="nav-item" onclick="showPage('dashboard')">
    <span class="nav-icon">📊</span> Dashboard
  </div>
  <div class="nav-section">Inventory</div>
  <div class="nav-item" onclick="showPage('search')">
    <span class="nav-icon">🔍</span> Search & Filter
  </div>
  <div class="nav-item" onclick="showPage('bulk')">
    <span class="nav-icon">📋</span> Bulk Lookup
  </div>
  <div class="nav-section">Insights</div>
  <div class="nav-item" onclick="showPage('compare')">
    <span class="nav-icon">📈</span> Month Compare
  </div>
  <div class="nav-item" onclick="showPage('quality')">
    <span class="nav-icon">🚨</span> Data Quality
  </div>
</div>

<!-- Main content -->
<div class="main">

<!-- ═══════════════════════════════════════════════════════ UPLOAD PAGE -->
<div class="page active" id="page-upload">
  <div class="section-hdr">Upload Inventory File</div>
  <div style="max-width:560px">
    <div class="upload-zone" id="drop-zone" onclick="document.getElementById('file-input').click()"
         ondragover="event.preventDefault();this.classList.add('drag')"
         ondragleave="this.classList.remove('drag')"
         ondrop="handleDrop(event)">
      <div class="upload-icon">📂</div>
      <div style="font-size:1rem;font-weight:600;margin-top:8px">
        Click or drag your Excel / CSV file here
      </div>
      <div class="upload-label">Supports .xlsx · .xls · .csv — max 50 MB</div>
      <input type="file" id="file-input" accept=".xlsx,.xls,.csv"
             onchange="uploadFile(this.files[0],'current')">
    </div>
    <div id="upload-status" style="margin-top:14px"></div>
  </div>
</div>

<!-- ═══════════════════════════════════════════════════════ DASHBOARD PAGE -->
<div class="page" id="page-dashboard">
  <div class="section-hdr">Dashboard</div>
  <div id="dash-loading" class="loading-row"><div class="spinner"></div> Loading…</div>
  <div id="dash-content" style="display:none">
    <div class="metrics-row" id="metrics-row"></div>
    <div class="charts-grid" id="charts-grid"></div>
  </div>
</div>

<!-- ═══════════════════════════════════════════════════════ SEARCH PAGE -->
<div class="page" id="page-search">
  <div class="section-hdr">Search & Filter</div>

  <!-- Team quick-view chips -->
  <div class="chip-row" id="team-chips">
    <span style="font-size:.78rem;color:var(--muted);align-self:center">Quick View:</span>
    <span class="chip chip-dynamo" onclick="quickFilter('platform','Dynamo',this)">Dynamo</span>
    <span class="chip chip-dj"    onclick="quickFilter('platform','Digital Journey',this)">Digital Journey</span>
    <span class="chip chip-ema"   onclick="quickFilter('platform','EMA',this)">EMA</span>
    <span class="chip chip-epmc"  onclick="quickFilter('platform','EPMC',this)">EPMC</span>
    <span class="chip chip-jea"   onclick="quickFilter('platform','JEA',this)">JEA</span>
    <span class="chip chip-hpc"   onclick="quickFilter('hpc','HPC',this)">HPC</span>
    <span class="chip chip-jpc"   onclick="quickFilter('hpc','JPC',this)">JPC</span>
    <span class="chip btn-outline" onclick="clearAllFilters()">✕ Clear All</span>
  </div>

  <div class="filter-bar">
    <div>
      <label>Search hostname / app / ref</label>
      <input type="text" id="f-search" placeholder="Type to search…" oninput="debSearch()">
    </div>
    <div>
      <label>Platform</label>
      <select id="f-platform" onchange="runSearch(1)"><option value="">All</option></select>
    </div>
    <div>
      <label>Status</label>
      <select id="f-status" onchange="runSearch(1)"><option value="">All</option></select>
    </div>
    <div>
      <label>Server Role</label>
      <select id="f-role" onchange="runSearch(1)"><option value="">All</option></select>
    </div>
    <div>
      <label>HPC / Group</label>
      <select id="f-hpc" onchange="runSearch(1)"><option value="">All</option></select>
    </div>
    <div>
      <label>DC Location</label>
      <select id="f-loc" onchange="runSearch(1)"><option value="">All</option></select>
    </div>
    <div>
      <label>OS</label>
      <select id="f-os" onchange="runSearch(1)"><option value="">All</option></select>
    </div>
    <div>
      <label>Server Type</label>
      <select id="f-stype" onchange="runSearch(1)"><option value="">All</option></select>
    </div>
    <div style="align-self:flex-end">
      <button class="btn btn-green btn-sm" onclick="exportCurrent()">📥 Export</button>
    </div>
  </div>

  <div class="tbl-count" id="search-count"></div>
  <div class="tbl-wrap">
    <table>
      <thead><tr id="search-thead"></tr></thead>
      <tbody id="search-tbody"><tr><td class="loading-row" colspan="12">Upload a file to search.</td></tr></tbody>
    </table>
  </div>
  <div class="pagination" id="search-pagination"></div>
</div>

<!-- ═══════════════════════════════════════════════════════ BULK PAGE -->
<div class="page" id="page-bulk">
  <div class="section-hdr">Bulk Server Lookup</div>
  <div style="max-width:680px">
    <p style="color:var(--muted);font-size:.85rem;margin-bottom:10px">
      Paste server hostnames below — one per line. Useful for cross-checking server lists
      from emails, tickets, or change requests.
    </p>
    <textarea class="bulk-area" id="bulk-input"
      placeholder="srv-prod-001&#10;srv-prod-002&#10;APPSERVER-12&#10;…"></textarea>
    <div style="display:flex;gap:10px;margin-top:10px">
      <button class="btn" onclick="runBulk()">🔍 Look Up Servers</button>
      <button class="btn btn-outline" onclick="document.getElementById('bulk-input').value=''">Clear</button>
    </div>
    <div id="bulk-summary" style="margin-top:14px"></div>
    <div id="bulk-not-found"></div>
    <div id="bulk-results" style="margin-top:14px"></div>
  </div>
</div>

<!-- ═══════════════════════════════════════════════════════ COMPARE PAGE -->
<div class="page" id="page-compare">
  <div class="section-hdr">Month-on-Month Comparison</div>
  <div class="compare-uploads">
    <div class="compare-box">
      <h4>📅 Current Month File</h4>
      <div style="font-size:.8rem;color:var(--muted);margin-bottom:8px">
        Already uploaded: <span id="curr-file-label" style="color:var(--blue)">None</span>
      </div>
    </div>
    <div class="compare-box">
      <h4>📅 Previous Month File</h4>
      <div class="upload-zone" style="padding:18px" onclick="document.getElementById('prev-input').click()">
        <div>📂 Click to upload previous month</div>
        <div style="font-size:.75rem;color:var(--muted);margin-top:4px">.xlsx / .xls / .csv</div>
        <input type="file" id="prev-input" accept=".xlsx,.xls,.csv" style="display:none"
               onchange="uploadPrev(this.files[0])">
      </div>
      <div id="prev-status" style="margin-top:8px;font-size:.8rem;color:var(--muted)"></div>
    </div>
  </div>
  <button class="btn" onclick="runCompare()" style="margin-bottom:20px">📊 Compare Months</button>
  <div id="compare-out"></div>
</div>

<!-- ═══════════════════════════════════════════════════════ QUALITY PAGE -->
<div class="page" id="page-quality">
  <div class="section-hdr">🚨 Data Quality</div>
  <div id="quality-out">
    <div class="alert alert-info">Upload a file to see data quality analysis.</div>
  </div>
</div>

</div><!-- /main -->

<!-- Detail modal -->
<div class="modal-overlay" id="modal-overlay" onclick="closeModal(event)">
  <div class="modal">
    <div class="modal-title">
      <span id="modal-hostname">Server Detail</span>
      <span class="modal-close" onclick="document.getElementById('modal-overlay').classList.remove('show')">✕</span>
    </div>
    <div class="detail-grid" id="modal-body"></div>
  </div>
</div>

<script>
// ── Theme ────────────────────────────────────────────────────────────────────
function toggleTheme(){
  document.body.classList.toggle('light');
  localStorage.setItem('theme', document.body.classList.contains('light')?'light':'dark');
}
if(localStorage.getItem('theme')==='light') document.body.classList.add('light');

// ── Page routing ─────────────────────────────────────────────────────────────
function showPage(name){
  document.querySelectorAll('.page').forEach(p=>p.classList.remove('active'));
  document.querySelectorAll('.nav-item').forEach(n=>n.classList.remove('active'));
  document.getElementById('page-'+name).classList.add('active');
  event.currentTarget.classList.add('active');
  if(name==='dashboard') loadDashboard();
  if(name==='search')    { loadFilterOptions(); runSearch(1); }
  if(name==='quality')   loadQuality();
  if(name==='compare')   {
    const cf = document.getElementById('curr-file-label');
    if(cf) cf.textContent = window._currFile || 'None';
  }
}

// ── Upload ───────────────────────────────────────────────────────────────────
function handleDrop(e){
  e.preventDefault();
  document.getElementById('drop-zone').classList.remove('drag');
  const f = e.dataTransfer.files[0];
  if(f) uploadFile(f,'current');
}

function uploadFile(file, which){
  if(!file) return;
  const fd = new FormData();
  fd.append('file', file);
  fd.append('which', which);
  const el = document.getElementById('upload-status');
  el.innerHTML = '<div class="alert alert-info"><div class="spinner"></div> Uploading…</div>';
  fetch('/upload',{method:'POST',body:fd})
    .then(r=>r.json())
    .then(d=>{
      if(d.error){
        el.innerHTML = `<div class="alert alert-warn">⚠️ ${d.error}</div>`;
        return;
      }
      window._currFile = d.filename;
      document.getElementById('topbar-meta').textContent =
        `${d.filename} · ${d.rows.toLocaleString()} servers · Status col: ${d.status_col}`;
      let qHtml = '';
      if(d.quality && Object.keys(d.quality).length){
        const items = Object.entries(d.quality).map(([k,v])=>
          `<div class="quality-item">⚠️ <b>${k}</b>: <span>${v}</span> blank rows</div>`
        ).join('');
        qHtml = `<div class="quality-panel" style="margin-top:10px">
          <h4>⚠️ Data Quality Warnings</h4>${items}</div>`;
      }
      el.innerHTML = `<div class="alert alert-success">
        ✅ Loaded <b>${d.rows.toLocaleString()}</b> servers from <b>${d.filename}</b><br>
        Status column detected: <b style="color:var(--green)">${d.status_col}</b>
      </div>${qHtml}`;
    })
    .catch(e=>{ el.innerHTML=`<div class="alert alert-warn">⚠️ ${e}</div>`; });
}

function uploadPrev(file){
  if(!file) return;
  const fd = new FormData();
  fd.append('file', file);
  fd.append('which', 'prev');
  document.getElementById('prev-status').innerHTML = '<div class="spinner"></div> Uploading…';
  fetch('/upload',{method:'POST',body:fd})
    .then(r=>r.json())
    .then(d=>{
      if(d.error){
        document.getElementById('prev-status').textContent = '⚠️ '+d.error;
        return;
      }
      document.getElementById('prev-status').innerHTML =
        `<span style="color:var(--green)">✅ ${d.filename} — ${d.rows.toLocaleString()} rows · Status: ${d.status_col}</span>`;
    });
}

// ── Dashboard ────────────────────────────────────────────────────────────────
function loadDashboard(){
  document.getElementById('dash-loading').style.display='';
  document.getElementById('dash-content').style.display='none';
  fetch('/dashboard').then(r=>r.json()).then(d=>{
    if(d.error){ document.getElementById('dash-loading').innerHTML=
      `<div class="alert alert-info">${d.error}</div>`; return; }
    document.getElementById('dash-loading').style.display='none';
    document.getElementById('dash-content').style.display='';

    const m = d.metrics;
    document.getElementById('metrics-row').innerHTML = `
      <div class="metric-card"><div class="metric-val c-blue">${m.total.toLocaleString()}</div>
        <div class="metric-lbl">Total Servers</div></div>
      <div class="metric-card"><div class="metric-val c-green">${m.live.toLocaleString()}</div>
        <div class="metric-lbl">Live (BAU)</div></div>
      <div class="metric-card"><div class="metric-val c-red">${m.not_live.toLocaleString()}</div>
        <div class="metric-lbl">Not Live</div></div>
      <div class="metric-card"><div class="metric-val c-cyan">${m.physical.toLocaleString()}</div>
        <div class="metric-lbl">Physical</div></div>
      <div class="metric-card"><div class="metric-val c-purple">${m.virtual.toLocaleString()}</div>
        <div class="metric-lbl">Virtual</div></div>`;

    const cg = document.getElementById('charts-grid');
    cg.innerHTML = '';
    const order = ['status_donut','type_donut','platform_hbar','loc_hbar',
                   'hpc_hbar','role_vbar','os_vbar','app_hbar'];
    order.forEach(k=>{
      if(d.charts[k]) cg.innerHTML +=
        `<div class="chart-card"><img src="${d.charts[k]}" alt="${k}"></div>`;
    });
  });
}

// ── Filter options ───────────────────────────────────────────────────────────
let _filterLoaded = false;
function loadFilterOptions(){
  if(_filterLoaded) return;
  fetch('/filter_options').then(r=>r.json()).then(d=>{
    fillSelect('f-platform', d.platform||[]);
    fillSelect('f-status',   d.status  ||[]);
    fillSelect('f-role',     d.role    ||[]);
    fillSelect('f-hpc',      d.hpc     ||[]);
    fillSelect('f-loc',      d.loc     ||[]);
    fillSelect('f-os',       d.os      ||[]);
    fillSelect('f-stype',    d.stype   ||[]);
    _filterLoaded = true;
  });
}
function fillSelect(id, opts){
  const el = document.getElementById(id);
  el.innerHTML = '<option value="">All</option>' +
    opts.map(o=>`<option value="${esc(o)}">${esc(o)}</option>`).join('');
}

// ── Search ───────────────────────────────────────────────────────────────────
let _debTimer;
function debSearch(){ clearTimeout(_debTimer); _debTimer=setTimeout(()=>runSearch(1),350); }

let _currPage = 1;
let _statusColCache = '';

function runSearch(page){
  _currPage = page||1;
  const params = new URLSearchParams({
    q:        document.getElementById('f-search').value,
    platform: document.getElementById('f-platform').value,
    status:   document.getElementById('f-status').value,
    role:     document.getElementById('f-role').value,
    hpc:      document.getElementById('f-hpc').value,
    loc:      document.getElementById('f-loc').value,
    os:       document.getElementById('f-os').value,
    stype:    document.getElementById('f-stype').value,
    page:     _currPage,
  });
  document.getElementById('search-tbody').innerHTML =
    '<tr><td colspan="12" class="loading-row"><div class="spinner"></div></td></tr>';
  fetch('/search?'+params).then(r=>r.json()).then(d=>{
    if(d.status_col) _statusColCache = d.status_col;
    document.getElementById('search-count').textContent =
      `Showing ${Math.min(d.total,50*(_currPage-1)+d.rows.length).toLocaleString()} of ${d.total.toLocaleString()} servers`;
    renderTable(d.rows, d.status_col||_statusColCache);
    renderPagination(d.page, d.pages);
  });
}

function renderTable(rows, scol){
  if(!rows.length){
    document.getElementById('search-thead').innerHTML = '';
    document.getElementById('search-tbody').innerHTML =
      '<tr><td colspan="12" class="loading-row">No results found.</td></tr>';
    return;
  }
  const cols = Object.keys(rows[0]);
  document.getElementById('search-thead').innerHTML =
    cols.map(c=>`<th>${c.replace(/_/g,' ')}</th>`).join('');
  document.getElementById('search-tbody').innerHTML = rows.map(r=>`<tr style="cursor:pointer"
    onclick="showDetail('${esc(r['Server HostName']||'')}')">${
    cols.map(c=>{
      const v = r[c]||'';
      if(c==='Server HostName') return `<td style="font-weight:700;color:var(--blue);font-family:monospace">${esc(v)}</td>`;
      if(c===scol) return `<td>${statusBadge(v)}</td>`;
      if(c==='Server Type(Physical/Virtual)') return `<td>${typeBadge(v)}</td>`;
      return `<td>${esc(v)}</td>`;
    }).join('')
  }</tr>`).join('');
}

function statusBadge(v){
  const l = (v||'').toLowerCase();
  if(l==='live')     return `<span class="badge badge-live">Live</span>`;
  if(l==='not live') return `<span class="badge badge-notlive">Not Live</span>`;
  return `<span class="badge badge-other">${esc(v)}</span>`;
}
function typeBadge(v){
  const l = (v||'').toLowerCase();
  if(l.includes('physical')) return `<span class="badge badge-physical">Physical</span>`;
  if(l.includes('virtual'))  return `<span class="badge badge-virtual">Virtual</span>`;
  return `<span class="badge badge-other">${esc(v)||'—'}</span>`;
}

function renderPagination(page, pages){
  const pg = document.getElementById('search-pagination');
  if(pages<=1){pg.innerHTML='';return;}
  let html = '';
  if(page>1) html+=`<button class="page-btn" onclick="runSearch(${page-1})">‹ Prev</button>`;
  const start=Math.max(1,page-2), end=Math.min(pages,page+2);
  if(start>1) html+=`<button class="page-btn" onclick="runSearch(1)">1</button>${start>2?'<span style="color:var(--muted)">…</span>':''}`;
  for(let i=start;i<=end;i++)
    html+=`<button class="page-btn ${i===page?'active':''}" onclick="runSearch(${i})">${i}</button>`;
  if(end<pages) html+=`${end<pages-1?'<span style="color:var(--muted)">…</span>':''}<button class="page-btn" onclick="runSearch(${pages})">${pages}</button>`;
  if(page<pages) html+=`<button class="page-btn" onclick="runSearch(${page+1})">Next ›</button>`;
  pg.innerHTML = html;
}

// ── Quick filters ─────────────────────────────────────────────────────────────
function quickFilter(field, value, el){
  document.querySelectorAll('.chip-row .chip').forEach(c=>c.classList.remove('active'));
  const map={platform:'f-platform',hpc:'f-hpc',role:'f-role',status:'f-status'};
  const selId = map[field];
  if(!selId) return;
  const sel = document.getElementById(selId);
  if(sel.value===value){ sel.value=''; }
  else { sel.value=value; el.classList.add('active'); }
  runSearch(1);
}
function clearAllFilters(){
  ['f-search',''].forEach(()=>{});
  document.getElementById('f-search').value='';
  ['f-platform','f-status','f-role','f-hpc','f-loc','f-os','f-stype'].forEach(id=>{
    document.getElementById(id).value='';
  });
  document.querySelectorAll('.chip-row .chip').forEach(c=>c.classList.remove('active'));
  runSearch(1);
}

// ── Export ────────────────────────────────────────────────────────────────────
function exportCurrent(){
  const params = new URLSearchParams({
    q:        document.getElementById('f-search').value,
    platform: document.getElementById('f-platform').value,
    status:   document.getElementById('f-status').value,
    role:     document.getElementById('f-role').value,
  });
  window.location = '/export?'+params;
}

// ── Detail modal ──────────────────────────────────────────────────────────────
function showDetail(hostname){
  if(!hostname) return;
  fetch(`/detail?hostname=${encodeURIComponent(hostname)}`).then(r=>r.json()).then(d=>{
    if(!d||!Object.keys(d).length) return;
    document.getElementById('modal-hostname').textContent = d['Server HostName']||hostname;
    const scol = d['_status_col']||'';
    const sval = d['_status_val']||'';
    const skip = new Set(['_status_col','_status_val']);
    let html = '';
    const full = ['Application Name','Reference','Commercial Category'];
    Object.entries(d).forEach(([k,v])=>{
      if(skip.has(k)||!v) return;
      const isFull = full.includes(k);
      let vHtml = esc(v);
      if(k===scol) vHtml = statusBadge(v);
      if(k==='Server Type(Physical/Virtual)') vHtml = typeBadge(v);
      html += `<div class="detail-row ${isFull?'detail-full':''}">
        <div class="detail-key">${esc(k)}</div>
        <div class="detail-val">${vHtml}</div></div>`;
    });
    document.getElementById('modal-body').innerHTML = html;
    document.getElementById('modal-overlay').classList.add('show');
  });
}
function closeModal(e){
  if(e.target===document.getElementById('modal-overlay'))
    document.getElementById('modal-overlay').classList.remove('show');
}

// ── Bulk lookup ───────────────────────────────────────────────────────────────
let _bulkRows = [];
function runBulk(){
  const raw = document.getElementById('bulk-input').value;
  const names = raw.split('\n').map(s=>s.trim()).filter(Boolean);
  if(!names.length) return;
  document.getElementById('bulk-summary').innerHTML =
    '<div class="spinner"></div> Looking up…';
  document.getElementById('bulk-results').innerHTML = '';
  document.getElementById('bulk-not-found').innerHTML = '';
  fetch('/bulk',{method:'POST',headers:{'Content-Type':'application/json'},
    body:JSON.stringify({names})})
    .then(r=>r.json())
    .then(d=>{
      _bulkRows = d.found||[];
      const scol = d.status_col||'';
      const foundC  = d.total_found;
      const notFoundC = (d.not_found||[]).length;
      document.getElementById('bulk-summary').innerHTML = `
        <div class="alert ${foundC?'alert-success':'alert-warn'}">
          ✅ Found <b>${foundC}</b> of <b>${names.length}</b> servers
          ${notFoundC?`&nbsp;·&nbsp; ❌ <b>${notFoundC}</b> not found`:''}
          ${foundC?`&nbsp;&nbsp;<button class="btn btn-sm btn-green" onclick="exportBulk()">📥 Export</button>`:''}
        </div>`;
      if(d.not_found&&d.not_found.length){
        document.getElementById('bulk-not-found').innerHTML =
          `<div class="not-found-list"><h4>❌ Not found (${d.not_found.length}) — may be typo or not in inventory</h4>
          ${d.not_found.map(n=>`<div class="item">· ${esc(n)}</div>`).join('')}</div>`;
      }
      if(_bulkRows.length){
        const cols = Object.keys(_bulkRows[0]);
        document.getElementById('bulk-results').innerHTML = `
          <div class="tbl-wrap" style="margin-top:12px"><table>
            <thead><tr>${cols.map(c=>`<th>${esc(c)}</th>`).join('')}</tr></thead>
            <tbody>${_bulkRows.map(r=>`<tr style="cursor:pointer"
              onclick="showDetail('${esc(r['Server HostName']||'')}')">${
              cols.map(c=>{
                const v=r[c]||'';
                if(c==='Server HostName') return `<td style="font-weight:700;color:var(--blue);font-family:monospace">${esc(v)}</td>`;
                if(c===scol) return `<td>${statusBadge(v)}</td>`;
                if(c==='Server Type(Physical/Virtual)') return `<td>${typeBadge(v)}</td>`;
                return `<td>${esc(v)}</td>`;
              }).join('')}</tr>`).join('')}
            </tbody></table></div>`;
      }
    });
}
function exportBulk(){
  fetch('/export_bulk',{method:'POST',headers:{'Content-Type':'application/json'},
    body:JSON.stringify({rows:_bulkRows})})
    .then(r=>r.blob())
    .then(b=>{
      const a=document.createElement('a');
      a.href=URL.createObjectURL(b);
      a.download=`bulk_lookup_${Date.now()}.xlsx`;
      a.click();
    });
}

// ── Compare ───────────────────────────────────────────────────────────────────
function runCompare(){
  document.getElementById('compare-out').innerHTML =
    '<div class="loading-row"><div class="spinner"></div> Comparing…</div>';
  fetch('/compare').then(r=>r.json()).then(d=>{
    if(d.error){
      document.getElementById('compare-out').innerHTML =
        `<div class="alert alert-warn">⚠️ ${d.error}</div>`;
      return;
    }
    const diffColor = n => n>0?'c-green':n<0?'c-red':'c-blue';
    let html = `
      <div class="diff-cards">
        <div class="diff-card">
          <div class="diff-val c-blue">${d.curr_total.toLocaleString()}</div>
          <div class="diff-lbl">Current Total</div></div>
        <div class="diff-card">
          <div class="diff-val c-blue">${d.prev_total.toLocaleString()}</div>
          <div class="diff-lbl">Previous Total</div></div>
        <div class="diff-card">
          <div class="diff-val ${diffColor(d.diff_total)}">${d.diff_total>=0?'+':''}${d.diff_total}</div>
          <div class="diff-lbl">Net Change</div></div>
        <div class="diff-card">
          <div class="diff-val c-green">${d.new_count}</div>
          <div class="diff-lbl">New Servers</div></div>
        <div class="diff-card">
          <div class="diff-val c-red">${d.removed_count}</div>
          <div class="diff-lbl">Removed</div></div>
        <div class="diff-card">
          <div class="diff-val c-cyan">${(d.to_live||[]).length}</div>
          <div class="diff-lbl">Went Live</div></div>
        <div class="diff-card">
          <div class="diff-val c-orange">${(d.to_not_live||[]).length}</div>
          <div class="diff-lbl">Left BAU</div></div>
      </div>`;

    // Charts
    if(d.charts){
      html += '<div class="charts-grid">';
      ['plat_compare','live_compare'].forEach(k=>{
        if(d.charts[k]) html += `<div class="chart-card" style="grid-column:1/-1">
          <img src="${d.charts[k]}" alt="${k}"></div>`;
      });
      html += '</div>';
    }

    // Went Live table
    if(d.to_live&&d.to_live.length){
      html += compareTable('✅ Servers that went Live this month',d.to_live,'var(--green)');
    }
    // Left BAU
    if(d.to_not_live&&d.to_not_live.length){
      html += compareTable('⚠️ Servers that left BAU (Live → Not Live)',d.to_not_live,'var(--orange)');
    }
    // New servers
    if(d.new_servers&&d.new_servers.length){
      html += compareTable('🆕 New Servers (not in previous month)',d.new_servers,'var(--blue)');
    }
    // Removed
    if(d.removed_servers&&d.removed_servers.length){
      html += compareTable('🗑️ Removed Servers (were in previous, gone now)',d.removed_servers,'var(--red)');
    }

    document.getElementById('compare-out').innerHTML = html;
  });
}

function compareTable(title, rows, color){
  if(!rows||!rows.length) return '';
  const cols = Object.keys(rows[0]);
  return `<div style="margin-bottom:20px">
    <div class="section-hdr" style="color:${color}">${title} (${rows.length})</div>
    <div class="tbl-wrap"><table>
      <thead><tr>${cols.map(c=>`<th>${esc(c)}</th>`).join('')}</tr></thead>
      <tbody>${rows.map(r=>`<tr>${cols.map(c=>{
        const v=r[c]||'';
        if(c==='Server HostName') return `<td style="font-weight:700;color:var(--blue);font-family:monospace">${esc(v)}</td>`;
        if(c==='Curr Status') return `<td>${statusBadge(v)}</td>`;
        if(c==='Prev Status') return `<td><span class="badge badge-other">${esc(v)}</span></td>`;
        return `<td>${esc(v)}</td>`;
      }).join('')}</tr>`).join('')}
      </tbody></table></div></div>`;
}

// ── Data quality ──────────────────────────────────────────────────────────────
function loadQuality(){
  fetch('/dashboard').then(r=>r.json()).then(d=>{
    const q = d.quality||{};
    const el = document.getElementById('quality-out');
    if(!Object.keys(q).length){
      el.innerHTML = `<div class="alert alert-success">✅ No data quality issues found! All key fields are complete.</div>`;
      return;
    }
    const items = Object.entries(q).map(([k,v])=>`
      <div class="quality-item">⚠️ <b>${esc(k)}</b> — <span>${v}</span> rows with blank values</div>
    `).join('');
    el.innerHTML = `
      <div class="alert alert-warn" style="margin-bottom:16px">
        Found issues in <b>${Object.keys(q).length}</b> column(s). Review below and fix in the source Excel.
      </div>
      <div class="quality-panel"><h4>⚠️ Issues Detected</h4>${items}</div>
      <p style="color:var(--muted);font-size:.82rem;margin-top:12px">
        💡 Fix these in your Excel file and re-upload for a cleaner inventory.
      </p>`;
  });
}

// ── Utility ───────────────────────────────────────────────────────────────────
function esc(s){
  return String(s||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;')
                      .replace(/"/g,'&quot;');
}
</script>
</body>
</html>"""

# ── Entry point ───────────────────────────────────────────────────────────────
if __name__ == "__main__":
    import socket
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        lan_ip = s.getsockname()[0]
        s.close()
    except Exception:
        lan_ip = "YOUR_SERVER_IP"

    port = 5050
    print("\n" + "=" * 64)
    print("  Server Asset Inventory  v1.0")
    print("=" * 64)
    print(f"  Local   :  http://localhost:{port}")
    print(f"  Network :  http://{lan_ip}:{port}   ← share with team")
    print("=" * 64)
    print("  100% OFFLINE — No internet required")
    print("  Charts via matplotlib · No CDN · Single .py file")
    print("  Formats : .xlsx  .xls  .csv")
    print("  Press Ctrl+C to stop")
    print("=" * 64 + "\n")
    app.run(debug=False, port=port, host="0.0.0.0", threaded=True)
