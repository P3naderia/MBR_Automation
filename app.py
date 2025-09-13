# -*- coding: utf-8 -*-
# python -m streamlit run test.py
# -*- coding: utf-8 -*-
# -*- coding: utf-8 -*-
# -*- coding: utf-8 -*-
# python -m streamlit run test.py
import os, re, io, tempfile, shutil
from datetime import datetime
from typing import Dict, List, Optional, Tuple

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
import streamlit as st

# python-pptx
from pptx import Presentation
from pptx.util import Pt
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_COLOR_TYPE, MSO_THEME_COLOR  # (남겨두되 사용 안 함: 호환성)

# =========================
# Global style / constants
# =========================
FONT_PATH = "C:/Windows/Fonts/malgun.ttf"  # Windows: 맑은 고딕 (없으면 기본 폰트 사용)
MONTH_LABELS = ['1월','2월','3월','4월','5월','6월','7월','8월','9월','10월','11월','12월']
PALETTE = {
    "primary": "#2F80ED", "green": "#27AE60", "orange": "#F2994A",
    "purple": "#9B51E0", "red": "#EB5757", "gray": "#BDBDBD", "dark": "#4F4F4F",
    "sp_fill": "#A9CEF8", "sb_fill": "#FAD4AD", "sd_fill": "#DCC4F6", "ba_fill": "#BFD3F2",
}

TOPCAT_METRICS: Dict[str, object] = {}   # 그래프11에서 채움 (Top1/2 카테고리용)
TOPASIN_METRICS: Dict[str, object] = {}  # 새로 추가 (Top1/2 ASIN 마커용)
GRAPH_ROOT: Optional[str] = None         # 세션 임시 작업폴더

# =========================
# Streamlit 기본 설정
# =========================
st.set_page_config(page_title="MBR PPT 자동 생성기 (3파일 업로드)", layout="wide")
st.title("📊 MBR PowerPoint 자동 생성기")
st.markdown("""
1) **CID 레벨 CSV/XLSX**, 2) **ASIN 레벨 CSV/XLSX**, 3) **PPT 템플릿(.pptx)** 을 업로드하세요.  
버튼을 누르면 **그래프(1~12, 16~17, 18~24) 생성 → 템플릿의 '그래프n' 자리 자동 삽입 → 텍스트/표 마커 치환 + YTD 테이블 채움 → 완성 PPT 다운로드**까지 자동 처리합니다.
""")

# =========================
# 파일 업로드
# =========================
cid_up = st.file_uploader("📁 CID 레벨 데이터 (CSV 또는 XLSX)", type=["csv","xlsx"], key="cid")
asin_up = st.file_uploader("📁 ASIN 레벨 데이터 (CSV 또는 XLSX)", type=["csv","xlsx"], key="asin")
ppt_up = st.file_uploader("📄 PowerPoint 템플릿 (.pptx)", type=["pptx"], key="ppt")

# =========================
# 공통 유틸
# =========================
def _set_korean_font_if_possible():
    try:
        if os.path.exists(FONT_PATH):
            font_prop = fm.FontProperties(fname=FONT_PATH)
            plt.rcParams['font.family'] = font_prop.get_name()
    except:
        pass

def ensure_graphs_folder() -> str:
    base = GRAPH_ROOT if GRAPH_ROOT else os.getcwd()
    graphs_folder = os.path.join(base, "graphs")
    os.makedirs(graphs_folder, exist_ok=True)
    return graphs_folder

def _save_fig(fig, name):
    graphs = ensure_graphs_folder()
    path = os.path.join(graphs, f"{name}.png")
    fig.savefig(path, dpi=300, bbox_inches="tight")
    plt.close(fig)
    return path

def parse_date_any(x):
    if pd.isna(x): return None
    try: return pd.to_datetime(str(x)).to_pydatetime().replace(day=1)
    except: return None

def parse_number_any(x, pct_to_100=False):
    if pd.isna(x): return None
    s = str(x).replace(',', '').replace('$', '').strip()
    m = re.search(r'[-+]?\d*\.?\d+', s)
    if not m: return None
    val = float(m.group(0))
    if '%' in s: return val
    return val*100 if pct_to_100 else val

def finalize_year_month(df, year_col='year', month_col='month'):
    df[year_col]  = pd.to_numeric(df[year_col], errors='coerce')
    df[month_col] = pd.to_numeric(df[month_col], errors='coerce')
    df = df.dropna(subset=[year_col, month_col]).copy()
    df[year_col]  = df[year_col].astype(int)
    df[month_col] = df[month_col].astype(int)
    df = df.sort_values([year_col, month_col]).reset_index(drop=True)
    df['date_str'] = df[year_col].astype(str) + '-' + df[month_col].astype(str).str.zfill(2)
    return df

def monthly_agg(dates, values, agg='mean'):
    rows=[]
    for d,v in zip(dates, values):
        dt = parse_date_any(d)
        val = parse_number_any(v) if v is not None else None
        if dt is None or val is None: continue
        rows.append({'year': dt.year, 'month': dt.month, 'value': val})
    if not rows: return pd.DataFrame(columns=['year','month','value','date_str'])
    t = pd.DataFrame(rows)
    gp = getattr(t.groupby(['year','month'])['value'], agg)()
    out = gp.reset_index()
    return finalize_year_month(out, 'year', 'month')

def _bi_theme(ax, ygrid=True):
    ax.set_facecolor("white")
    if ax.figure is not None: ax.figure.set_facecolor("white")
    for side in ["top","right"]: ax.spines[side].set_visible(False)
    for side in ["left","bottom"]: ax.spines[side].set_color("#BDBDBD")
    ax.tick_params(colors=PALETTE["dark"], labelsize=10)
    if ygrid: ax.yaxis.grid(True, color="#E6E6E6", linestyle="-", linewidth=1)
    ax.xaxis.grid(False)

def _yfmt_decimal(dec=1, suffix=""):
    return plt.FuncFormatter(lambda x, pos: f"{x:.{dec}f}{suffix}")

def _yfmt_k(dec=1):
    return plt.FuncFormatter(lambda x, pos: f"{x/1000:.{dec}f}K")

def _label_last(ax, xs, ys, text, dy=6):
    if len(xs)==0: return
    ax.annotate(text, (xs[-1], ys[-1]), textcoords="offset points", xytext=(0,dy),
                ha="center", va="bottom", fontsize=9,
                bbox=dict(boxstyle="round,pad=0.25", fc="white", ec="#DDDDDD", alpha=0.9))

# =========================
# (색상/폰트 보존 정책)
# =========================
# 텍스트 내 퍼센트에 ▲▼만 추가하고, 폰트/색상은 '절대' 건드리지 않음.
def style_percent_text_and_color(text, run):
    m = re.search(r'(-?\d+(?:\.\d+)?)\s*%', text or "")
    if not m:
        return text
    try:
        val = float(m.group(1))
    except ValueError:
        return text
    if val > 0:
        if "▲" not in text and "▼" not in text: return text + " ▲"
    elif val < 0:
        if "▲" not in text and "▼" not in text: return text + " ▼"
    return text

# =========================
# CID용 계산/마커 포맷
# =========================
def load_cid(file_path):
    ext = os.path.splitext(file_path)[1].lower()
    df = pd.read_csv(file_path) if ext==".csv" else pd.read_excel(file_path)
    if "Month" not in df.columns:
        raise KeyError("CID 파일에 Month 컬럼이 필요합니다.")
    df["Month"] = pd.to_datetime(df["Month"], errors="coerce")
    df["year"]  = df["Month"].dt.year
    df["month"] = df["Month"].dt.month
    return df

def safe_pct(a, b):
    if b == 0 or pd.isna(b): return np.nan
    return (a - b) / b * 100

def calc_metrics(sub):
    gms   = sub["GMS"].sum()
    gv    = sub["GV"].sum()
    units = sub["Units"].sum()
    ba    = sub["BA"].sum() if "BA" in sub.columns else 0
    asp   = gms / units if units > 0 else np.nan
    cr    = units / gv * 100 if gv > 0 else np.nan
    gcpb  = gms / ba if (ba is not None and ba > 0) else np.nan
    return {"gms": gms, "gv": gv, "units": units, "ba": ba, "asp": asp, "cr": cr, "gcpb": gcpb}

def fmt_gms(v):
    if pd.isna(v): return "N/A"
    if v >= 1_000_000: return f"{v/1_000_000:.1f}M"
    if v >= 1_000:     return f"{v/1_000:.0f}K"
    return f"{v:.0f}"

def fmt_k(v):
    if pd.isna(v) or v <= 0: return "N/A"
    return f"{v/1_000:.0f}K"

def fmt_pct(v, dec=1):
    if pd.isna(v): return "N/A"
    return f"{float(v):.{dec}f}%"

def month_back(year, month, n):
    y, m = year, month
    for _ in range(n):
        if m == 1: y -= 1; m = 12
        else: m -= 1
    return y, m

def metric_key_from_marker(marker):
    m = marker.lower()
    if   "gms"   in m: return "gms"
    elif "gv"    in m: return "gv"
    elif "units" in m: return "units"
    elif "ba"    in m: return "ba"
    elif "asp"   in m: return "asp"
    elif "cr"    in m: return "cr"
    elif "gcpb"  in m: return "gcpb"
    elif "ipi"   in m: return "ipi"
    elif "excess" in m: return "excess"
    elif "avgwoc" in m or "avg.woc" in m or "avg_woc" in m: return "avgwoc"
    elif "fbagms" in m or "fba gms" in m: return "fbagms"
    elif "awas" in m: return "awas"
    elif "awagv" in m: return "awagv"
    return None

# ----- Top1/Top2 카테고리 텍스트 (그래프11 연동 그대로 유지)
def _get_topcat_text_marker(marker):
    s = marker.lower().replace(" ", "")
    if "top1category" in s:      return TOPCAT_METRICS.get("top1_category", "N/A")
    if "top1gmsportion" in s:    return fmt_pct(TOPCAT_METRICS.get("top1_portion", np.nan))
    if "top1gmsgrowth" in s:     return fmt_pct(TOPCAT_METRICS.get("top1_growth", np.nan))
    if "top2category" in s:      return TOPCAT_METRICS.get("top2_category", "N/A")
    if "top2gmsportion" in s:    return fmt_pct(TOPCAT_METRICS.get("top2_portion", np.nan))
    if "top2gmsgrowth" in s:     return fmt_pct(TOPCAT_METRICS.get("top2_growth", np.nan))
    return None

# ----- Top1/Top2 ASIN 텍스트 마커(신규)
def _get_topasin_text_marker(marker):
    s = marker.strip()
    if s == "Top1 ASIN":              return TOPASIN_METRICS.get("top1_asin", "N/A")
    if s == "Top1 ASIN portion":      return f"{TOPASIN_METRICS.get('top1_portion','N/A')}"
    if s == "Top1 ASIN Growth":       return f"{TOPASIN_METRICS.get('top1_growth','N/A')}"
    if s == "Top2 ASIN":              return TOPASIN_METRICS.get("top2_asin", "N/A")
    if s == "Top2 ASIN portion":      return f"{TOPASIN_METRICS.get('top2_portion','N/A')}"
    if s == "Top2 ASIN Growth":       return f"{TOPASIN_METRICS.get('top2_growth','N/A')}"
    return None

def extract_value(df, marker):
    # Top1/Top2 카테고리 텍스트 마커 우선
    val_top = _get_topcat_text_marker(marker)
    if val_top is not None:
        return val_top
    # Top1/Top2 ASIN 텍스트 마커
    val_asin = _get_topasin_text_marker(marker)
    if val_asin is not None:
        return val_asin

    marker = marker.lower().strip()
    year_match = re.search(r"\d{4}", marker)
    year = int(year_match.group()) if year_match else datetime.now().year
    if df[df["year"] == year].empty:
        return "N/A"

    latest_month = int(df[df["year"] == year]["month"].max())
    anchor_year  = int(df["year"].max())
    anchor_month = int(df[df["year"] == anchor_year]["month"].max())

    mm_offset = None
    m = re.search(r"mm-(\d+)", marker)
    if m: mm_offset = int(m.group(1))

    def pick_month_subset(y, m_):
        return df[(df["year"] == y) & (df["month"] == m_)]

    try:
        if "y-1" in marker:
            prev_year = year - 1
            if not any(k in marker for k in ["gms","gv","units","ba","asp","cr","gcpb","ipi","excess","avgwoc","fbagms","awas","awagv"]):
                return f"{prev_year}-{str(latest_month).zfill(2)}"
            sub = pick_month_subset(prev_year, latest_month)
            if sub.empty: return "N/A"
            key = metric_key_from_marker(marker)
            if key in ("ipi","excess","avgwoc","fbagms","awas","awagv"): return "N/A"
            return _format_by_marker(key, calc_metrics(sub))

        if "m-1" in marker:
            prev_month = 12 if latest_month == 1 else latest_month - 1
            prev_year  = year - 1 if latest_month == 1 else year
            if not any(k in marker for k in ["gms","gv","units","ba","asp","cr","gcpb","ipi","excess","avgwoc","fbagms","awas","awagv"]):
                return f"{prev_year}-{str(prev_month).zfill(2)}"
            sub = pick_month_subset(prev_year, prev_month)
            if sub.empty: return "N/A"
            key = metric_key_from_marker(marker)
            if key in ("ipi","excess","avgwoc","fbagms","awas","awagv"): return "N/A"
            return _format_by_marker(key, calc_metrics(sub))

        if mm_offset is not None:
            ty, tm = month_back(year, latest_month, mm_offset)
            if not any(k in marker for k in ["gms","gv","units","ba","asp","cr","gcpb","ipi","excess","avgwoc","fbagms","awas","awagv"]):
                return f"{ty}-{str(tm).zfill(2)}"
            sub = pick_month_subset(ty, tm)
            if sub.empty: return "N/A"
            key = metric_key_from_marker(marker)
            if key in ("ipi","excess","avgwoc","fbagms","awas","awagv"): return "N/A"
            return _format_by_marker(key, calc_metrics(sub))

        if "mm" in marker and "mom" not in marker and "yoy" not in marker:
            if not any(k in marker for k in ["gms","gv","units","ba","asp","cr","gcpb","ipi","excess","avgwoc","fbagms","awas","awagv"]):
                return f"{year}-{str(latest_month).zfill(2)}"
            sub = pick_month_subset(year, latest_month)
            if sub.empty: return "N/A"
            key = metric_key_from_marker(marker)
            if key in ("ipi","excess","avgwoc","fbagms","awas","awagv"): return "N/A"
            return _format_by_marker(key, calc_metrics(sub))

        if "ytd" in marker and "yoy" not in marker:
            cur = calc_metrics(df[(df["year"] == year) & (df["month"] <= anchor_month)])
            key = metric_key_from_marker(marker)
            if key in ("ipi","excess","avgwoc","fbagms","awas","awagv"): return "N/A"
            return _format_by_marker(key, cur)

        if "ytd" in marker and "yoy" in marker:
            key = metric_key_from_marker(marker)
            if key in ("ipi","excess","avgwoc","fbagms","awas","awagv"): return "N/A"
            cur  = calc_metrics(df[(df["year"] == year)   & (df["month"] <= anchor_month)])
            prev = calc_metrics(df[(df["year"] == year-1) & (df["month"] <= anchor_month)])
            return fmt_pct(safe_pct(cur[key], prev[key]) if prev[key] > 0 else np.nan)

        if "mm" in marker and "yoy" in marker:
            key = metric_key_from_marker(marker)
            if key in ("ipi","excess","avgwoc","fbagms","awas","awagv"): return "N/A"
            cur  = calc_metrics(df[(df["year"] == year)   & (df["month"] == latest_month)])
            prev = calc_metrics(df[(df["year"] == year-1) & (df["month"] == latest_month)])
            return fmt_pct(safe_pct(cur[key], prev[key]) if prev[key] > 0 else np.nan)

        if "mm" in marker and "mom" in marker:
            key = metric_key_from_marker(marker)
            if key in ("ipi","excess","avgwoc","fbagms","awas","awagv"): return "N/A"
            cur = calc_metrics(df[(df["year"] == year) & (df["month"] == latest_month)])
            prev_month = 12 if latest_month == 1 else latest_month-1
            prev_year  = year - 1 if latest_month == 1 else year
            prev = calc_metrics(df[(df["year"] == prev_year) & (df["month"] == prev_month)])
            return fmt_pct(safe_pct(cur[key], prev[key]) if prev[key] > 0 else np.nan)

    except Exception as e:
        print(f"⚠️ 마커 처리 실패: {marker} → {e}")
        return "N/A"

    return "N/A"

def _format_by_marker(key, metrics):
    if key == "gms":   return fmt_gms(metrics["gms"])
    if key == "gv":    return fmt_k(metrics["gv"])
    if key == "units": return fmt_k(metrics["units"])
    if key == "ba":    return f"{metrics['ba']:.0f}" if (not pd.isna(metrics["ba"]) and metrics["ba"]>0) else "N/A"
    if key == "asp":   return f"{metrics['asp']:.1f}" if not pd.isna(metrics["asp"]) else "N/A"
    if key == "cr":    return fmt_pct(metrics["cr"])
    if key == "gcpb":  return f"{metrics['gcpb']:.1f}" if not pd.isna(metrics["gcpb"]) else "N/A"
    return "N/A"

# =========================
# PPT 텍스트/테이블 치환 (폰트/색상 보존)
# =========================
def iter_all_shapes(shapes):
    for sh in shapes:
        if sh.shape_type == MSO_SHAPE_TYPE.GROUP:
            for sub in iter_all_shapes(sh.shapes):
                yield sub
        else:
            yield sh

def replace_text_preserve_style(shape, df):
    """문단/런을 지우지 않고, 기존 첫 run의 스타일을 그대로 유지한 채 텍스트만 교체."""
    if not getattr(shape, "has_text_frame", False):
        return
    for paragraph in shape.text_frame.paragraphs:
        # 문단 전체 텍스트 합치기
        full_text = "".join(run.text for run in paragraph.runs)

        # 마커 치환
        def _replace_markers(s):
            out, last = [], 0
            for m in re.finditer(r"\{([^}]+)\}", s):
                out.append(s[last:m.start()])
                marker = m.group(1)
                val = extract_value(df, marker)
                out.append(val if val is not None else "N/A")
                last = m.end()
            out.append(s[last:])
            return "".join(out)

        new_text = _replace_markers(full_text)

        # run 스타일 보존: 첫 run에만 텍스트 대입, 나머지는 빈 문자열 처리
        runs = paragraph.runs
        if not runs:
            r = paragraph.add_run()
            r.text = new_text
            r.text = style_percent_text_and_color(r.text, r)
            continue
        r0 = runs[0]
        r0.text = new_text
        r0.text = style_percent_text_and_color(r0.text, r0)
        for r in runs[1:]:
            r.text = ""

def process_ppt_markers(prs, df):
    for slide in prs.slides:
        for shape in iter_all_shapes(slide.shapes):
            # 도형 텍스트 치환
            replace_text_preserve_style(shape, df)
            # 표 셀 치환
            if getattr(shape, "has_table", False) and shape.has_table:
                for row in shape.table.rows:
                    for cell in row.cells:
                        tf = cell.text_frame
                        for paragraph in tf.paragraphs:
                            for run in paragraph.runs:
                                t = run.text
                                matches = re.findall(r"\{([^}]+)\}", t)
                                for marker in matches:
                                    val = extract_value(df, marker)
                                    t = t.replace("{"+marker+"}", val if val else "N/A")
                                run.text = style_percent_text_and_color(t, run)

# =========================
# YTD 테이블 자동 채움 (폰트/색상 보존)
# =========================
def find_title_shape(slide, regex):
    pat = re.compile(regex)
    for sh in iter_all_shapes(slide.shapes):
        if getattr(sh, "has_text_frame", False) and sh.has_text_frame:
            txt = "\n".join(p.text for p in sh.text_frame.paragraphs)
            if pat.search(txt): 
                return sh
    return None

def center(shape):
    return (shape.left + shape.width/2, shape.top + shape.height/2)

def nearest_table(slide, anchor_shape):
    ax, ay = center(anchor_shape)
    best, best_d2 = None, None
    for sh in iter_all_shapes(slide.shapes):
        if hasattr(sh, "has_table") and sh.has_table:
            cx, cy = center(sh)
            d2 = (cx-ax)*(cx-ax) + (cy-ay)*(cy-ay)
            if best is None or d2 < best_d2:
                best, best_d2 = sh, d2
    return best.table if best else None

def write_cell(cell, text):
    """셀의 첫 run 스타일을 유지하고 텍스트만 교체."""
    tf = cell.text_frame
    if tf.paragraphs and tf.paragraphs[0].runs:
        p = tf.paragraphs[0]
        r0 = p.runs[0]
        r0.text = text
        r0.text = style_percent_text_and_color(r0.text, r0)
        for r in p.runs[1:]:
            r.text = ""
    else:
        p = tf.paragraphs[0]
        r = p.add_run()
        r.text = text
        r.text = style_percent_text_and_color(r.text, r)

def fill_ytd_table(table, df):
    anchor_year  = int(df["year"].max())
    anchor_month = int(df[df["year"] == anchor_year]["month"].max())
    years = [anchor_year, anchor_year-1, anchor_year-2]

    def ytd(y):
        sub = df[(df["year"] == y) & (df["month"] <= anchor_month)]
        if sub.empty:
            return {"gms": np.nan, "gv": np.nan, "units": np.nan, "cr": np.nan}
        gms   = sub["GMS"].sum()
        gv    = sub["GV"].sum()
        units = sub["Units"].sum()
        cr    = (units / gv * 100) if gv > 0 else np.nan
        return {"gms": gms, "gv": gv, "units": units, "cr": cr}

    yvals = {y: ytd(y) for y in years}

    def yoy(cur, prev, key):
        a = yvals[cur][key]; b = yvals[prev][key]
        return safe_pct(a, b) if (not pd.isna(b) and b > 0) else np.nan

    for ridx, y in enumerate(years, start=1):
        if ridx >= len(table.rows): break
        try:
            write_cell(table.cell(ridx, 0), f"{y} YTD (≤ {str(anchor_month).zfill(2)})")
        except Exception:
            pass
        v = yvals[y]
        write_cell(table.cell(ridx, 1), fmt_gms(v["gms"]))
        write_cell(table.cell(ridx, 2), fmt_k(v["gv"]))
        write_cell(table.cell(ridx, 3), fmt_k(v["units"]))
        write_cell(table.cell(ridx, 4), fmt_pct(v["cr"]))
        if y-1 in yvals:
            write_cell(table.cell(ridx, 5), fmt_pct(yoy(y, y-1, "gms")))
            write_cell(table.cell(ridx, 6), fmt_pct(yoy(y, y-1, "gv")))
            write_cell(table.cell(ridx, 7), fmt_pct(yoy(y, y-1, "units")))
            write_cell(table.cell(ridx, 8), fmt_pct(yoy(y, y-1, "cr")))
        else:
            for c in (5,6,7,8):
                write_cell(table.cell(ridx, c), "N/A")

def fill_ytd_table_on_slides(prs, df):
    count = 0
    for slide in prs.slides:
        title = find_title_shape(slide, r"연간\s*누적\s*매출\s*트렌드")
        if not title: continue
        table = nearest_table(slide, title)
        if not table: continue
        try:
            fill_ytd_table(table, df); count += 1
        except Exception as e:
            print(f"YTD 테이블 채우기 실패: {e}")
    return count

# =========================
# 그래프 생성 (원본 로직 유지)
# =========================
def create_line_by_year(dates, values, title, y_label, graph_name,
                        unit="none", decimal=1, percentage=False, annotate_year=None):
    df = monthly_agg(dates, values)
    if df.empty:
        return None
    df["value_conv"] = df["value"]
    _set_korean_font_if_possible()
    fig = plt.figure(figsize=(10, 6)); ax = plt.gca(); _bi_theme(ax)
    colors_cycle = [PALETTE["primary"], PALETTE["green"], PALETTE["orange"], PALETTE["purple"], PALETTE["red"]]
    for idx, y in enumerate(sorted(df["year"].unique())):
        sub = df[df["year"] == y].sort_values("month")
        xs = sub["month"].to_numpy(); ys = sub["value_conv"].to_numpy()
        ax.plot(xs, ys, linewidth=2.5, marker="o", markersize=5, color=colors_cycle[idx%len(colors_cycle)], label=f"{y}년")
        if annotate_year and y == annotate_year:
            txt = f"{ys[-1]:.{decimal}f}{'%' if percentage else ''}"
            _label_last(ax, xs, ys, txt)
    ax.set_title(title, fontsize=16, color=PALETTE["dark"])
    ax.set_xlabel("월", fontsize=11, color=PALETTE["dark"])
    ax.set_ylabel(y_label, fontsize=11, color=PALETTE["dark"])
    ax.set_xticks(range(1,13)); ax.set_xticklabels(MONTH_LABELS)
    if unit=="K": ax.yaxis.set_major_formatter(_yfmt_k(decimal))
    elif percentage: ax.yaxis.set_major_formatter(_yfmt_decimal(decimal, "%"))
    else: ax.yaxis.set_major_formatter(_yfmt_decimal(decimal))
    ax.legend(loc="upper left", fontsize=10, frameon=False)
    fig.tight_layout()
    return _save_fig(fig, graph_name)

def create_combo_ba_awagv_awas(dates, ba_values, awagv_values, awas_values, graph_name, title):
    rows=[]; n=len(dates)
    for i in range(n):
        dt = parse_date_any(dates[i])
        if dt is None: continue
        ba   = parse_number_any(ba_values[i]) if i<len(ba_values) else None
        awag = parse_number_any(avag:=awagv_values[i]) if i<len(awagv_values) else None  # noqa
        awas = parse_number_any(awas_values[i]) if i<len(awas_values) else None
        if ba is None or ba <= 0: continue
        disc = (awag/ba*100) if awag is not None else np.nan
        conv = (awas/ba*100) if awas is not None else np.nan
        rows.append({"year":dt.year,"month":dt.month,"ba":ba,"disc":disc,"conv":conv})
    if not rows: return None
    df = pd.DataFrame(rows).groupby(["year","month"]).mean().reset_index()
    df = finalize_year_month(df, "year", "month")
    _set_korean_font_if_possible()
    fig = plt.figure(figsize=(10, 6)); ax1 = plt.gca(); _bi_theme(ax1)
    x = np.arange(len(df))
    bars = ax1.bar(x, df["ba"], alpha=0.9, color=PALETTE["ba_fill"], label="BA")
    ax1.set_ylabel("취급 상품 개수 (개)", color=PALETTE["dark"])
    ax2 = ax1.twinx(); _bi_theme(ax2); ax2.set_yticks([])
    ax2.plot(x, df["disc"], "o-", linewidth=2.5, color=PALETTE["primary"])
    ax2.plot(x, df["conv"], "s-", linewidth=2.5, color=PALETTE["orange"])
    if pd.notna(df["conv"].iloc[-1]):
        _label_last(ax2, x, df["conv"].to_numpy(), f"{df['conv'].iloc[-1]:.1f}%")
    ax1.set_title(title, fontsize=16, color=PALETTE["dark"])
    ax1.set_xlabel("연월", color=PALETTE["dark"])
    ax1.set_xticks(x); ax1.set_xticklabels(df["date_str"], rotation=45, ha="right")
    ax1.legend([bars], ["BA"], loc="upper left", fontsize=9, frameon=False)
    fig.tight_layout()
    return _save_fig(fig, graph_name)

def create_ipi_combo_graph(dates, ipi_values, excess_pct_values, graph_name, title):
    rows = []
    for i in range(len(dates)):
        dt = parse_date_any(dates[i])
        ipi = parse_number_any(ipi_values[i]) if i<len(ipi_values) else None
        exc = parse_number_any(excess_pct_values[i], pct_to_100=True) if i<len(excess_pct_values) else None
        if dt is None or ipi is None: continue
        rows.append({"year":dt.year,"month":dt.month,"ipi":ipi,"excess":exc or np.nan})
    if not rows: return None
    df = pd.DataFrame(rows).groupby(["year","month"]).mean().reset_index()
    df = finalize_year_month(df, "year", "month")
    _set_korean_font_if_possible()
    fig = plt.figure(figsize=(10, 6)); ax1 = plt.gca(); _bi_theme(ax1)
    x = np.arange(len(df))
    bars = ax1.bar(x, df["ipi"], alpha=0.9, color=PALETTE["ba_fill"], label="IPI Score")
    ax1.set_ylabel("IPI Score", color=PALETTE["dark"])
    ax2 = ax1.twinx(); _bi_theme(ax2); ax2.set_yticks([])
    y2 = df["excess"].to_numpy()
    ax2.plot(x, y2, "o-", linewidth=2.5, color=PALETTE["orange"], label="Excess PCT")
    _label_last(ax2, x, y2, f"{y2[-1]:.1f}%")
    ax2.set_ylabel("Excess PCT (%)", color=PALETTE["dark"])
    ax1.set_title(title, fontsize=16, color=PALETTE["dark"])
    ax1.set_xlabel("연월", color=PALETTE["dark"])
    ax1.set_xticks(x); ax1.set_xticklabels(df["date_str"], rotation=45, ha="right")
    ax1.legend([bars], ["IPI Score"], loc="upper left", fontsize=10, frameon=False)
    fig.tight_layout()
    return _save_fig(fig, graph_name)

def create_merchandising_graph(dates, total_sales, bd_ops, ld_ops, dotd_ops, mario_ops, coupon_ops, graph_name, title):
    rows = []; n = len(dates)
    for i in range(n):
        dt = parse_date_any(dates[i])
        ts = parse_number_any(total_sales[i]) if i<len(total_sales) else None
        if dt is None or ts is None: continue
        bd = parse_number_any(bd_ops[i])    if i<len(bd_ops)    else 0
        ld = parse_number_any(ld_ops[i])    if i<len(ld_ops)    else 0
        do = parse_number_any(dotd_ops[i])  if i<len(dotd_ops)  else 0
        ma = parse_number_any(mario_ops[i]) if i<len(mario_ops) else 0
        cp = parse_number_any(coupon_ops[i])if i<len(coupon_ops)else 0
        total = (bd or 0)+(ld or 0)+(do or 0)+(ma or 0)+(cp or 0)
        pct = total/ts*100 if ts>0 else 0
        rows.append({"year":dt.year,"month":dt.month,"bd":bd or 0,"ld":ld or 0,"do":do or 0,"ma":ma or 0,"cp":cp or 0,"pct":pct})
    if not rows: return None
    df = pd.DataFrame(rows).groupby(["year","month"]).mean().reset_index()
    df = finalize_year_month(df, "year", "month")
    _set_korean_font_if_possible()
    fig = plt.figure(figsize=(12, 6)); ax1 = plt.gca(); _bi_theme(ax1)
    x = np.arange(len(df))
    bd, ld, do, ma, cp = df["bd"], df["ld"], df["do"], df["ma"], df["cp"]
    b1 = ax1.bar(x, bd, label="Best Deal", color=PALETTE["sp_fill"])
    b2 = ax1.bar(x, ld, bottom=bd, label="Lightning Deal", color=PALETTE["sb_fill"])
    b3 = ax1.bar(x, do, bottom=bd+ld, label="Deal of The Day", color=PALETTE["sd_fill"])
    b4 = ax1.bar(x, ma, bottom=bd+ld+do, label="Prime Exclusive Discount", color="#C7EBD0")
    b5 = ax1.bar(x, cp, bottom=bd+ld+do+ma, label="Coupon", color="#F8E3A2")
    ax1.set_ylabel("Merchandising OPS", color=PALETTE["dark"])
    ax2 = ax1.twinx(); _bi_theme(ax2); ax2.set_yticks([])
    y2 = df["pct"].to_numpy()
    ax2.plot(x, y2, "o-", linewidth=2.5, color=PALETTE["primary"], label="Merchandising OPS%")
    _label_last(ax2, x, y2, f"{y2[-1]:.1f}%")
    ax2.set_ylabel("Merchandising OPS%", color=PALETTE["dark"])
    ax1.set_title(title, fontsize=16, color=PALETTE["dark"])
    ax1.set_xlabel("연월", color=PALETTE["dark"])
    ax1.set_xticks(x); ax1.set_xticklabels(df["date_str"], rotation=45, ha="right")
    ax1.legend([b1,b2,b3,b4,b5], ["Best Deal","Lightning Deal","Deal of The Day","Prime Exclusive Discount","Coupon"],
               loc="upper left", fontsize=9, frameon=False)
    fig.tight_layout()
    return _save_fig(fig, graph_name)

def create_ads_tacos_graph(dates, total_sales, sp_spend, sb_spend, sd_spend, graph_name, title):
    rows=[]
    for i in range(len(dates)):
        dt = parse_date_any(dates[i])
        ts = parse_number_any(total_sales[i]) if i < len(total_sales) else None
        sp = parse_number_any(sp_spend[i])   if i < len(sp_spend)   else 0
        sb = parse_number_any(sb_spend[i])   if i < len(sb_spend)   else 0
        sd = parse_number_any(sd_spend[i])   if i < len(sd_spend)   else 0
        if dt is None or ts is None: continue
        total_ads = (sp or 0)+(sb or 0)+(sd or 0)
        tacos = total_ads/ts*100 if ts>0 else 0
        rows.append({"year":dt.year,"month":dt.month,"sp":sp or 0,"sb":sb or 0,"sd":sd or 0,"tacos":tacos})
    if not rows: return None
    df = pd.DataFrame(rows).groupby(["year","month"]).mean().reset_index()
    df = finalize_year_month(df, "year", "month")
    _set_korean_font_if_possible()
    fig = plt.figure(figsize=(12, 6)); ax1 = plt.gca(); _bi_theme(ax1)
    x = np.arange(len(df))
    b1 = ax1.bar(x, df["sp"], label="SP Spend", color=PALETTE["sp_fill"])
    b2 = ax1.bar(x, df["sb"], bottom=df["sp"], label="SB Spend", color=PALETTE["sb_fill"])
    b3 = ax1.bar(x, df["sd"], bottom=df["sp"]+df["sb"], label="SD Spend", color=PALETTE["sd_fill"])
    ax1.set_ylabel("Ads Spend", color=PALETTE["dark"])
    ax2 = ax1.twinx(); _bi_theme(ax2); ax2.set_yticks([])
    line_y = df["tacos"].to_numpy()
    ax2.plot(x, line_y, "o-", color=PALETTE["primary"], linewidth=2.5)
    _label_last(ax2, x, line_y, f"{line_y[-1]:.1f}%")
    ax1.set_title(title, fontsize=16, color=PALETTE["dark"])
    ax1.set_xlabel("연월", color=PALETTE["dark"]); ax1.set_xticks(x); ax1.set_xticklabels(df["date_str"], rotation=45, ha="right")
    ax1.legend([b1,b2,b3], ["SP Spend","SB Spend","SD Spend"], loc="upper left", fontsize=9, frameon=False)
    fig.tight_layout()
    return _save_fig(fig, graph_name)

def create_ads_impr_clicks_graph(dates, sp_imp, sp_clk, sb_imp, sb_clk, sd_imp, sd_clk, graph_name, title):
    rows=[]
    n=len(dates)
    for i in range(n):
        dt = parse_date_any(dates[i])
        if dt is None: continue
        sp_i = parse_number_any(sp_imp[i]) if i < len(sp_imp) else 0
        sp_c = parse_number_any(sp_clk[i]) if i < len(sp_clk) else 0
        sb_i = parse_number_any(sb_imp[i]) if i < len(sb_imp) else 0
        sb_c = parse_number_any(sb_clk[i]) if i < len(sb_clk) else 0
        sd_i = parse_number_any(sd_imp[i]) if i < len(sd_imp) else 0
        sd_c = parse_number_any(sd_clk[i]) if i < len(sd_clk) else 0
        rows.append({"year":dt.year,"month":dt.month,"sp_i":sp_i,"sb_i":sb_i,"sd_i":sd_i,
                     "clicks": (sp_c or 0)+(sb_c or 0)+(sd_c or 0)})
    if not rows: return None
    df = pd.DataFrame(rows).groupby(["year","month"]).mean().reset_index()
    df = finalize_year_month(df, "year", "month")
    _set_korean_font_if_possible()
    fig = plt.figure(figsize=(12, 6)); ax1 = plt.gca(); _bi_theme(ax1)
    x = np.arange(len(df))
    b1 = ax1.bar(x, df["sp_i"], label="SP Impression", color=PALETTE["sp_fill"])
    b2 = ax1.bar(x, df["sb_i"], bottom=df["sp_i"], label="SB Impression", color=PALETTE["sb_fill"])
    b3 = ax1.bar(x, df["sd_i"], bottom=df["sp_i"]+df["sb_i"], label="SD Impression", color=PALETTE["sd_fill"])
    ax1.set_ylabel("Impressions", color=PALETTE["dark"])
    ax2 = ax1.twinx(); _bi_theme(ax2); ax2.set_yticks([])
    clicks = df["clicks"].to_numpy()
    ax2.plot(x, clicks, "o-", color=PALETTE["primary"], linewidth=2.5, label="AD Clicks")
    _label_last(ax2, x, clicks, f"{clicks[-1]:.1f}")
    ax1.set_title(title, fontsize=16, color=PALETTE["dark"])
    ax1.set_xlabel("연월", color=PALETTE["dark"]); ax1.set_xticks(x); ax1.set_xticklabels(df["date_str"], rotation=45, ha="right")
    ax1.legend([b1,b2,b3], ["SP Impression","SB Impression","SD Impression"], loc="upper left", fontsize=9, frameon=False)
    fig.tight_layout()
    return _save_fig(fig, graph_name)

def create_three_line_pct(dates, a, b, c, labels, ylabel, graph_name, title):
    rows=[]
    n=len(dates)
    for i in range(n):
        dt = parse_date_any(dates[i])
        if dt is None: continue
        va = parse_number_any(a[i], pct_to_100=True) if i<len(a) else None
        vb = parse_number_any(b[i], pct_to_100=True) if i<len(b) else None
        vc = parse_number_any(c[i], pct_to_100=True) if i<len(c) else None
        rows.append({"year":dt.year,"month":dt.month, "A":va or 0, "B":vb or 0, "C":vc or 0})
    if not rows: return None
    df = pd.DataFrame(rows).groupby(["year","month"]).mean().reset_index()
    df = finalize_year_month(df, "year", "month")
    _set_korean_font_if_possible()
    fig = plt.figure(figsize=(12, 6)); ax = plt.gca(); _bi_theme(ax)
    x = np.arange(len(df))
    ax.plot(x, df["A"], "o-", linewidth=2.5, color=PALETTE["primary"], label=labels[0])
    ax.plot(x, df["B"], "s-", linewidth=2.5, color=PALETTE["orange"],  label=labels[1])
    ax.plot(x, df["C"], "^-", linewidth=2.5, color=PALETTE["purple"],  label=labels[2])
    _label_last(ax, x, df["A"].to_numpy(), f"{df['A'].iloc[-1]:.2f}%")
    ax.set_title(title, fontsize=16, color=PALETTE["dark"])
    ax.set_ylabel(ylabel, color=PALETTE["dark"]); ax.set_xlabel("연월", color=PALETTE["dark"])
    ax.set_xticks(x); ax.set_xticklabels(df["date_str"], rotation=45, ha="right")
    ax.yaxis.set_major_formatter(_yfmt_decimal(2, "%"))
    ax.legend(loc="upper left", fontsize=10, frameon=False)
    fig.tight_layout()
    return _save_fig(fig, graph_name)

def create_three_line_pct_nolabel(dates, a, b, c, labels, ylabel, graph_name, title):
    rows=[]; n=len(dates)
    for i in range(n):
        dt = parse_date_any(dates[i])
        if dt is None: continue
        va = parse_number_any(a[i], pct_to_100=True) if i<len(a) else None
        vb = parse_number_any(b[i], pct_to_100=True) if i<len(b) else None
        vc = parse_number_any(c[i], pct_to_100=True) if i<len(c) else None
        rows.append({'year':dt.year,'month':dt.month, 'A':va or 0, 'B':vb or 0, 'C':vc or 0})
    if not rows: return None
    df = pd.DataFrame(rows).groupby(['year','month']).mean().reset_index()
    df = finalize_year_month(df, 'year', 'month')
    _set_korean_font_if_possible()
    fig = plt.figure(figsize=(12,6)); ax = plt.gca(); _bi_theme(ax)
    x = np.arange(len(df))
    ax.plot(x, df['A'], 'o-', linewidth=2, label=labels[0])
    ax.plot(x, df['B'], 's-', linewidth=2, label=labels[1])
    ax.plot(x, df['C'], '^-', linewidth=2, label=labels[2])
    ax.set_title(title, fontsize=16)
    ax.set_ylabel(ylabel); ax.set_xlabel('연월')
    ax.set_xticks(x); ax.set_xticklabels(df['date_str'], rotation=45, ha='right')
    ax.legend(loc='upper left', fontsize=10, frameon=False)
    fig.tight_layout()
    return _save_fig(fig, graph_name)

def create_ads_sales_combo(dates, total_sales, sp_sales, sb_sales, sd_sales, graph_name, title):
    rows = []
    n = len(dates)
    for i in range(n):
        dt = parse_date_any(dates[i])
        if dt is None: 
            continue
        ts = parse_number_any(total_sales[i]) if i < len(total_sales) else None
        sp = parse_number_any(sp_sales[i])    if i < len(sp_sales)    else 0
        sb = parse_number_any(sb_sales[i])    if i < len(sb_sales)    else 0
        sd = parse_number_any(sd_sales[i])    if i < len(sd_sales)    else 0
        if ts is None:
            continue
        rows.append({
            "year": dt.year, "month": dt.month,
            "ts": ts or 0, "sp": sp or 0, "sb": sb or 0, "sd": sd or 0
        })

    if not rows:
        return None

    df = pd.DataFrame(rows).groupby(["year","month"]).sum(numeric_only=True).reset_index()
    df = finalize_year_month(df, "year", "month")
    df["ad_sales"] = df["sp"] + df["sb"] + df["sd"]
    df["ad_sales_pct"] = np.where(df["ts"] > 0, df["ad_sales"] / df["ts"] * 100, np.nan)

    _set_korean_font_if_possible()
    fig = plt.figure(figsize=(12, 6)); ax1 = plt.gca(); _bi_theme(ax1)

    x = np.arange(len(df))
    b1 = ax1.bar(x, df["sp"], label="SP Sales", color=PALETTE["sp_fill"])
    b2 = ax1.bar(x, df["sb"], bottom=df["sp"], label="SB Sales", color=PALETTE["sb_fill"])
    b3 = ax1.bar(x, df["sd"], bottom=df["sp"]+df["sb"], label="SD Sales", color=PALETTE["sd_fill"])

    ax1.set_ylabel("Ads Sales", color=PALETTE["dark"])

    ax2 = ax1.twinx(); _bi_theme(ax2); ax2.set_yticks([])
    y2 = df["ad_sales_pct"].to_numpy()
    ax2.plot(x, y2, "o-", linewidth=2.5, color=PALETTE["primary"], label="Ad sales%")
    if len(y2) > 0 and pd.notna(y2[-1]):
        _label_last(ax2, x, y2, f"{y2[-1]:.1f}%")

    ax1.set_title(title, fontsize=16, color=PALETTE["dark"])
    ax1.set_xlabel("연월", color=PALETTE["dark"])
    ax1.set_xticks(x); ax1.set_xticklabels(df["date_str"], rotation=45, ha="right")

    ax1.legend([b1,b2,b3], ["SP Sales","SB Sales","SD Sales"], loc="upper left", fontsize=9, frameon=False)
    fig.tight_layout()
    return _save_fig(fig, graph_name)

# =========================
# 그래프 11/12 (ASIN+CID)
# =========================
def _get_col(df, name):
    key = name.lower().replace(" ", "").replace("/", "")
    for c in df.columns:
        cc = str(c).lower().replace(" ", "").replace("/", "")
        if cc == key: return c
    raise KeyError(f"'{name}' 컬럼을 찾을 수 없습니다.")

def _fmt_big(v):
    if pd.isna(v): return "0"
    v = float(v)
    if v >= 1_000_000: return f"{v/1_000_000:.1f}M"
    if v >= 1_000:     return f"{v/1_000:.0f}K"
    return f"{v:.0f}"

def _heat_color(v, vmin=-100, vmax=100):
    import matplotlib.cm as cm, matplotlib.colors as colors
    if pd.isna(v): return "white"
    x = max(vmin, min(vmax, float(v)))
    norm = colors.Normalize(vmin=vmin, vmax=vmax)
    r,g,b,_ = cm.RdYlGn(norm(x))
    return (r, g, b)

def create_graph11_itkbn_dashboard(df_asin, df_monthly, sel_month=None, graph_name="Graph 11"):
    global TOPCAT_METRICS
    TOPCAT_METRICS = {}

    mcol = "Month" if "Month" in df_monthly.columns else None
    if mcol:
        df_monthly = df_monthly.copy()
        df_monthly["Month"] = pd.to_datetime(df_monthly["Month"], errors="coerce")
        anchor_dt = df_monthly["Month"].max()
        if sel_month:
            try: anchor_dt = pd.to_datetime(sel_month)
            except: pass
        month_str = anchor_dt.strftime("%Y-%m")
    else:
        month_str = "N/A"
        anchor_dt = pd.Timestamp.today().normalize()

    if "Month" not in df_asin.columns:
        raise KeyError("ASIN CSV에 Month 컬럼이 필요합니다. (YYYY-MM-DD 또는 YYYY-MM)")
    df_a = df_asin.copy()
    df_a["Month"] = pd.to_datetime(df_a["Month"], errors="coerce")
    df_a = df_a.dropna(subset=["Month"])

    col_itkbn = _get_col(df_a, "ITK/BN")
    col_gms   = _get_col(df_a, "GMS")
    col_gv    = _get_col(df_a, "GV")
    col_units = _get_col(df_a, "Units")

    cur = df_a[df_a["Month"].dt.to_period("M") == anchor_dt.to_period("M")]
    if cur.empty: cur = df_a
    donut_series = cur.groupby(col_itkbn, dropna=False)[col_gms].sum().sort_values(ascending=False)
    total_gms = float(donut_series.sum())

    top_list = donut_series.reset_index()
    top_list.columns = ["itkbn", "gms"]
    if not top_list.empty:
        top_list["portion"] = np.where(total_gms>0, top_list["gms"]/total_gms*100, np.nan)
        prev_dt = (anchor_dt - pd.offsets.MonthBegin(1))
        if len(top_list) >= 1:
            t1 = str(top_list.loc[0,"itkbn"])
            cur1 = cur[cur[col_itkbn]==t1][col_gms].sum()
            prev1 = df_a[(df_a[col_itkbn]==t1) & (df_a["Month"].dt.to_period("M")==prev_dt.to_period("M"))][col_gms].sum()
            g1 = safe_pct(cur1, prev1)
            TOPCAT_METRICS["top1_category"] = t1
            TOPCAT_METRICS["top1_portion"]  = float(top_list.loc[0,"portion"])
            TOPCAT_METRICS["top1_growth"]   = g1
        if len(top_list) >= 2:
            t2 = str(top_list.loc[1,"itkbn"])
            cur2 = cur[cur[col_itkbn]==t2][col_gms].sum()
            prev2 = df_a[(df_a[col_itkbn]==t2) & (df_a["Month"].dt.to_period("M")==prev_dt.to_period("M"))][col_gms].sum()
            g2 = safe_pct(cur2, prev2)
            TOPCAT_METRICS["top2_category"] = t2
            TOPCAT_METRICS["top2_portion"]  = float(top_list.loc[1,"portion"])
            TOPCAT_METRICS["top2_growth"]   = g2

    wc_series = cur.groupby(col_itkbn, dropna=False)[col_gms].sum().sort_values(ascending=False)
    wc_series = wc_series[wc_series>0].head(25)
    wc_labels = [str(k) if str(k)!="nan" else "Unknown" for k in wc_series.index]
    wc_vals   = wc_series.values.astype(float)

    TOP_N = 8
    donut_plot = donut_series.copy()
    if len(donut_plot) > TOP_N:
        others = donut_plot.iloc[TOP_N:].sum()
        donut_plot = donut_plot.iloc[:TOP_N]
        donut_plot.loc["Others"] = others
    donut_labels = [f"{k} ({_fmt_big(v)})" for k, v in zip(donut_plot.index, donut_plot.values)]
    donut_sizes  = donut_plot.values

    need_cols = {}
    for name in ["GMS","GV","Units"]:
        need_cols[name] = _get_col(df_monthly, name)
    ba_col = None
    for guess in ["BA","Buyable ASIN","BuyableASIN"]:
        try:
            ba_col = _get_col(df_monthly, guess); break
        except: pass

    mdf = df_monthly.copy()
    mdf["Month"] = pd.to_datetime(mdf["Month"], errors="coerce")
    mdf = mdf.dropna(subset=["Month"]).sort_values("Month")
    mdf["year"]  = mdf["Month"].dt.year
    mdf["month"] = mdf["Month"].dt.month

    agg = mdf.groupby(["year","month"]).agg({
        need_cols["GMS"]:"sum", need_cols["GV"]:"sum", need_cols["Units"]:"sum",
        ba_col if ba_col else need_cols["Units"]: "sum"
    }).reset_index()
    agg = agg.rename(columns={
        need_cols["GMS"]:"GMS", need_cols["GV"]:"GV", need_cols["Units"]:"Units",
        (ba_col if ba_col else need_cols["Units"]):"BA"
    })
    agg["ASP"] = np.where(agg["Units"]>0, agg["GMS"]/agg["Units"], np.nan)
    agg["CR"]  = np.where(agg["GV"]>0,    agg["Units"]/agg["GV"]*100, np.nan)
    agg["Month"] = pd.to_datetime(agg["year"].astype(str)+"-"+agg["month"].astype(str))
    agg = agg.sort_values("Month")

    for col in ["GMS","GV","Units","BA","ASP","CR"]:
        prev = agg[col].shift(1)
        agg[f"{col}_MoM"] = np.where(prev>0, (agg[col]-prev)/prev*100, np.nan)

    prev_y = agg[["year","month","GMS","GV","Units","BA","ASP","CR"]].copy()
    prev_y["year"] = prev_y["year"]+1
    merged = pd.merge(agg, prev_y, on=["year","month"], how="left", suffixes=("","_PY"))
    for col in ["GMS","GV","Units","BA","ASP","CR"]:
        agg[f"{col}_YoY"] = np.where(merged[f"{col}_PY"]>0, (agg[col]-merged[f"{col}_PY"])/merged[f"{col}_PY"]*100, np.nan)

    table_df = agg.sort_values("Month", ascending=False).copy()
    table_df["MonthStr"] = table_df["Month"].dt.strftime("%Y-%m")
    disp_cols = [
        "MonthStr","GMS","GV","Units","ASP","CR",
        "GMS_MoM","GV_MoM","Units_MoM","BA_MoM","ASP_MoM","CR_MoM",
        "GMS_YoY","GV_YoY","Units_YoY","BA_YoY","ASP_YoY","CR_YoY"
    ]
    table_df = table_df[disp_cols]

    _set_korean_font_if_possible()
    fig = plt.figure(figsize=(19.2, 10.8))
    fig.patch.set_facecolor("white")

    ax_wc = fig.add_axes([0.05, 0.60, 0.43, 0.32]); ax_wc.axis("off")
    ax_wc.set_title(f"ITK/BN Distribution by GMS  •  Month: {month_str}", fontsize=12, loc="left")
    if len(wc_labels) > 0:
        sizes = wc_vals / wc_vals.max()
        fs = 10 + (sizes**0.5) * 38
        cx, cy = 0.45, 0.45
        ax_wc.text(cx, cy, wc_labels[0], ha="center", va="center", fontsize=fs[0], color="#8AAAD6", alpha=0.85)
        rng = np.random.default_rng(42)
        xs = rng.uniform(0.05, 0.95, size=len(wc_labels)-1)
        ys = rng.uniform(0.05, 0.95, size=len(wc_labels)-1)
        for i, (tx, ty) in enumerate(zip(xs, ys), start=1):
            ax_wc.text(tx, ty, wc_labels[i], ha="center", va="center",
                       fontsize=fs[i], color="#AEC4E6", alpha=0.8, transform=ax_wc.transAxes)

    ax_d = fig.add_axes([0.56, 0.54, 0.38, 0.38])
    wedges, _ = ax_d.pie(donut_sizes, startangle=90, wedgeprops=dict(width=0.42, edgecolor="white"))
    ax_d.set_title("Sales% of ITK/BN", fontsize=12)
    ax_d.text(0, 0, _fmt_big(total_gms), ha="center", va="center", fontsize=18, fontweight="bold")
    ax_legend = fig.add_axes([0.80, 0.54, 0.18, 0.38]); ax_legend.axis("off")
    ax_legend.legend(wedges, donut_labels, loc="upper left", frameon=False, fontsize=9)

    ax_t = fig.add_axes([0.05, 0.06, 0.90, 0.44]); ax_t.axis("off")
    col_headers = ["Month","GMS","GV","Units","ASP","CR",
                   "GMS MoM","GV MoM","Units MoM","BA MoM","ASP MoM","CR MoM",
                   "GMS YoY","GV YoY","Units YoY","BA YoY","ASP YoY","CR YoY"]
    show = table_df.copy()
    show["GMS"]  = show["GMS"].map(lambda x: f"{x:,.0f}")
    show["GV"]   = show["GV"].map(lambda x: f"{x:,.0f}")
    show["Units"]= show["Units"].map(lambda x: f"{x:,.0f}")
    show["ASP"]  = show["ASP"].map(lambda x: f"{x:.2f}" if pd.notna(x) else "N/A")
    show["CR"]   = show["CR"].map(lambda x: fmt_pct(x))
    for c in ["GMS_MoM","GV_MoM","Units_MoM","BA_MoM","ASP_MoM","CR_MoM",
              "GMS_YoY","GV_YoY","Units_YoY","BA_YoY","ASP_YoY","CR_YoY"]:
        show[c] = show[c].map(lambda x: fmt_pct(x))

    cell_text = show.values.tolist()
    the_table = ax_t.table(cellText=cell_text, colLabels=col_headers, loc='center', cellLoc='center', colLoc='center')
    the_table.auto_set_font_size(False); the_table.set_fontsize(9); the_table.scale(1, 1.3)
    mom_idx = [col_headers.index(x) for x in ["GMS MoM","GV MoM","Units MoM","BA MoM","ASP MoM","CR MoM"]]
    yoy_idx = [col_headers.index(x) for x in ["GMS YoY","GV YoY","Units YoY","BA YoY","ASP YoY","CR YoY"]]
    for r in range(len(show)):
        for c_i, raw_col in zip(mom_idx, ["GMS_MoM","GV_MoM","Units_MoM","BA_MoM","ASP_MoM","CR_MoM"]):
            v = table_df.iloc[r][raw_col]; the_table[(r+1, c_i)].set_facecolor(_heat_color(v))
        for c_i, raw_col in zip(yoy_idx, ["GMS_YoY","GV_YoY","Units_YoY","BA_YoY","ASP_YoY","CR_YoY"]):
            v = table_df.iloc[r][raw_col]; the_table[(r+1, c_i)].set_facecolor(_heat_color(v))
    for c in range(len(col_headers)):
        the_table[(0, c)].set_facecolor("#F1F3F7"); the_table[(0, c)].set_text_props(weight='bold')

    fig.tight_layout()
    return _save_fig(fig, graph_name)

def create_graph12_top1_category_trends(df_asin, df_cid, graph_name="Graph 12"):
    def _get_col_safe(df, name):
        key = name.lower().replace(" ", "").replace("/", "")
        for c in df.columns:
            cc = str(c).lower().replace(" ", "").replace("/", "")
            if cc == key: return c
        raise KeyError(f"'{name}' 컬럼을 찾을 수 없습니다.")

    anchor = pd.to_datetime(df_cid["Month"], errors="coerce").max()
    if "Month" not in df_asin.columns:
        return None

    df_a = df_asin.copy()
    df_a["Month"] = pd.to_datetime(df_a["Month"], errors="coerce")
    df_a = df_a.dropna(subset=["Month"])

    col_itkbn = _get_col_safe(df_a, "ITK/BN")
    col_asin  = _get_col_safe(df_a, "ASIN")
    col_gms   = _get_col_safe(df_a, "GMS")
    col_gv    = _get_col_safe(df_a, "GV")
    col_units = _get_col_safe(df_a, "Units")

    cur = df_a[df_a["Month"].dt.to_period("M")==anchor.to_period("M")]
    if cur.empty: cur = df_a
    g_by = cur.groupby(col_itkbn, dropna=False)[col_gms].sum().sort_values(ascending=False)
    if g_by.empty:
        return None
    top1 = g_by.index[0]

    sub = df_a[df_a[col_itkbn]==top1].copy()
    sub["year"]  = sub["Month"].dt.year
    sub["month"] = sub["Month"].dt.month
    sub["BA_flag"] = (sub[col_units].fillna(0) > 0) | (sub[col_gms].fillna(0) > 0)
    agg = sub.groupby(["year","month"]).agg(
        GMS   =(col_gms,"sum"),
        GV    =(col_gv,"sum"),
        Units =(col_units,"sum"),
        BA    =(col_asin, lambda s: s[sub.loc[s.index,"BA_flag"]].nunique())
    ).reset_index()
    if agg.empty: return None
    agg["ASP"] = np.where(agg["Units"]>0, agg["GMS"]/agg["Units"], np.nan)
    agg["CR"]  = np.where(agg["GV"]>0,    agg["Units"]/agg["GV"]*100, np.nan)
    agg = finalize_year_month(agg, "year", "month")

    _set_korean_font_if_possible()
    fig, axes = plt.subplots(2,3, figsize=(19.2,10.8))
    fig.patch.set_facecolor("white")
    banner = fig.add_axes([0, 0.93, 1, 0.07])
    banner.set_facecolor("#0F6CBD"); banner.set_xticks([]); banner.set_yticks([])
    banner.text(0.5, 0.5, f"Top 1 카테고리 주요 지표 트렌드  •  {top1}",
                ha="center", va="center", color="white", fontsize=18, fontweight="bold")

    def draw_metric(ax, df, value_col, title, yfmt="raw", dec=1):
        _bi_theme(ax)
        years = sorted(df["year"].unique())
        hi_year = max(years)
        colors = {y: ("#F2B233" if y==hi_year else "#BDBDBD") for y in years}
        for y in years:
            suby = df[df["year"]==y].sort_values("month")
            x = suby["month"].to_numpy(); yv = suby[value_col].to_numpy()
            ax.plot(x, yv, marker="o", linewidth=2.2, color=colors[y], label=str(y))
            if y==hi_year and len(yv)>0 and pd.notna(yv[-1]):
                if yfmt=="k": txt = f"{yv[-1]/1000:.1f}K"
                elif yfmt=="pct": txt = f"{yv[-1]:.{dec}f}%"
                elif yfmt=="int": txt = f"{int(round(yv[-1]))}"
                else: txt = f"{yv[-1]:.{dec}f}"
                _label_last(ax, x, yv, txt, dy=8)
        ax.set_title(title, fontsize=14, color=PALETTE["dark"])
        ax.set_xticks(range(1,13)); ax.set_xlim(1,12)
        if yfmt=="k": ax.yaxis.set_major_formatter(_yfmt_k(dec))
        elif yfmt=="pct": ax.yaxis.set_major_formatter(_yfmt_decimal(dec, "%"))
        elif yfmt=="int": ax.yaxis.set_major_formatter(_yfmt_decimal(0, ""))        
        ax.legend(loc="upper left", fontsize=9, frameon=False)

    draw_metric(axes[0,0], agg, "GMS",   "매출",      yfmt="k",   dec=1)
    draw_metric(axes[0,1], agg, "GV",    "고객 유입", yfmt="k",   dec=1)
    draw_metric(axes[0,2], agg, "Units", "판매 수량", yfmt="k",   dec=1)
    draw_metric(axes[1,0], agg, "BA",    "판매 상품 개수", yfmt="int", dec=0)
    draw_metric(axes[1,1], agg, "ASP",   "판매 객단가", yfmt="raw", dec=2)
    draw_metric(axes[1,2], agg, "CR",    "구매 전환율", yfmt="pct", dec=1)

    fig.tight_layout(rect=[0,0,1,0.93])
    return _save_fig(fig, graph_name)

# =========================
# Top1/Top2 ASIN 그래프(그래프 16/17)
# =========================
def compute_top_asin_metrics(df_asin: pd.DataFrame, df_cid: pd.DataFrame) -> Dict[str, object]:
    """Top1/Top2 ASIN, portion(%), MoM growth(%) 계산하여 글로벌에 저장"""
    global TOPASIN_METRICS
    TOPASIN_METRICS = {}

    if df_asin is None or df_cid is None or df_asin.empty or df_cid.empty:
        return TOPASIN_METRICS

    anchor = pd.to_datetime(df_cid["Month"], errors="coerce").max()
    da = df_asin.copy()
    if "Month" not in da.columns:
        return TOPASIN_METRICS
    da["Month"] = pd.to_datetime(da["Month"], errors="coerce")
    da = da.dropna(subset=["Month"])
    if da.empty:
        return TOPASIN_METRICS

    col_asin = _get_col(da, "ASIN")
    col_gms  = _get_col(da, "GMS")

    cur = da[da["Month"].dt.to_period("M") == anchor.to_period("M")]
    if cur.empty:
        cur = da

    by_asin = cur.groupby(col_asin, dropna=False)[col_gms].sum().sort_values(ascending=False)
    total = float(by_asin.sum()) if by_asin.size else 0.0
    top_two = by_asin.head(2)

    prev_month = anchor - pd.offsets.MonthBegin(1)

    def _fmt1(x):
        if pd.isna(x): return "N/A"
        return f"{float(x):.1f}"

    for rank, (asin, gms_now) in enumerate(top_two.items(), start=1):
        gms_prev = da[(da[col_asin]==asin) & (da["Month"].dt.to_period("M")==prev_month.to_period("M"))][col_gms].sum()
        portion = (gms_now/total*100.0) if total>0 else np.nan
        growth  = safe_pct(gms_now, gms_prev) if gms_prev>0 else (np.nan if gms_prev==0 else np.nan)
        TOPASIN_METRICS[f"top{rank}_asin"]    = str(asin)
        TOPASIN_METRICS[f"top{rank}_portion"] = _fmt1(portion)
        TOPASIN_METRICS[f"top{rank}_growth"]  = _fmt1(growth)

    # 키 보정
    for k in ["top1_asin","top1_portion","top1_growth","top2_asin","top2_portion","top2_growth"]:
        TOPASIN_METRICS.setdefault(k, "N/A")
    return TOPASIN_METRICS

def create_top_asin_trend_graph(df_asin: pd.DataFrame, asin_value: str, graph_name: str, title_prefix: str):
    """단일 ASIN의 2x3 지표 트렌드 (GMS/GV/Units/BA/ASP/CR)"""
    if not asin_value or asin_value=="N/A":
        return None
    if "Month" not in df_asin.columns:
        return None

    col_asin  = _get_col(df_asin, "ASIN")
    col_gms   = _get_col(df_asin, "GMS")
    col_gv    = _get_col(df_asin, "GV")
    col_units = _get_col(df_asin, "Units")

    sub = df_asin[df_asin[col_asin]==asin_value].copy()
    if sub.empty:
        return None

    sub["Month"] = pd.to_datetime(sub["Month"], errors="coerce")
    sub = sub.dropna(subset=["Month"])
    sub["year"]  = sub["Month"].dt.year
    sub["month"] = sub["Month"].dt.month
    sub["BA_flag"] = (sub[col_units].fillna(0) > 0) | (sub[col_gms].fillna(0) > 0)

    agg = sub.groupby(["year","month"]).agg(
        GMS   =(col_gms,"sum"),
        GV    =(col_gv,"sum"),
        Units =(col_units,"sum"),
        BA    =(col_asin, lambda s: s[sub.loc[s.index,"BA_flag"]].nunique())
    ).reset_index()
    if agg.empty: return None
    agg["ASP"] = np.where(agg["Units"]>0, agg["GMS"]/agg["Units"], np.nan)
    agg["CR"]  = np.where(agg["GV"]>0,    agg["Units"]/agg["GV"]*100, np.nan)
    agg = finalize_year_month(agg, "year", "month")

    _set_korean_font_if_possible()
    fig, axes = plt.subplots(2,3, figsize=(19.2,10.8))
    fig.patch.set_facecolor("white")
    banner = fig.add_axes([0, 0.93, 1, 0.07])
    banner.set_facecolor("#0F6CBD"); banner.set_xticks([]); banner.set_yticks([])
    banner.text(0.5, 0.5, f"{title_prefix}  •  {asin_value}",
                ha="center", va="center", color="white", fontsize=18, fontweight="bold")

    def draw_metric(ax, df, value_col, title, yfmt="raw", dec=1):
        _bi_theme(ax)
        years = sorted(df["year"].unique())
        hi_year = max(years)
        colors = {y: ("#F2B233" if y==hi_year else "#BDBDBD") for y in years}
        for y in years:
            suby = df[df["year"]==y].sort_values("month")
            x = suby["month"].to_numpy(); yv = suby[value_col].to_numpy()
            ax.plot(x, yv, marker="o", linewidth=2.2, color=colors[y], label=str(y))
            if y==hi_year and len(yv)>0 and pd.notna(yv[-1]):
                if yfmt=="k": txt = f"{yv[-1]/1000:.1f}K"
                elif yfmt=="pct": txt = f"{yv[-1]:.{dec}f}%"
                elif yfmt=="int": txt = f"{int(round(yv[-1]))}"
                else: txt = f"{yv[-1]:.{dec}f}"
                _label_last(ax, x, yv, txt, dy=8)
        ax.set_title(title, fontsize=14, color=PALETTE["dark"])
        ax.set_xticks(range(1,13)); ax.set_xlim(1,12)
        if yfmt=="k": ax.yaxis.set_major_formatter(_yfmt_k(dec))
        elif yfmt=="pct": ax.yaxis.set_major_formatter(_yfmt_decimal(dec, "%"))
        elif yfmt=="int": ax.yaxis.set_major_formatter(_yfmt_decimal(0, ""))        
        ax.legend(loc="upper left", fontsize=9, frameon=False)

    draw_metric(axes[0,0], agg, "GMS",   "매출",      yfmt="k",   dec=1)
    draw_metric(axes[0,1], agg, "GV",    "고객 유입", yfmt="k",   dec=1)
    draw_metric(axes[0,2], agg, "Units", "판매 수량", yfmt="k",   dec=1)
    draw_metric(axes[1,0], agg, "BA",    "판매 상품 개수", yfmt="int", dec=0)
    draw_metric(axes[1,1], agg, "ASP",   "판매 객단가", yfmt="raw", dec=2)
    draw_metric(axes[1,2], agg, "CR",    "구매 전환율", yfmt="pct", dec=1)

    fig.tight_layout(rect=[0,0,1,0.93])
    return _save_fig(fig, graph_name)

# =========================
# PPT 그림 삽입 유틸
# =========================
def _marker_key(text: str) -> str:
    s = (text or "").strip().lower()
    s = s.replace(" ", "")
    s = s.replace("graph", "그래프")
    m = re.search(r"그래프0*(\d+)", s, flags=re.I)
    return f"그래프{int(m.group(1))}" if m else ""

def _iter_all_shapes(shapes):
    for sh in shapes:
        if sh.shape_type == MSO_SHAPE_TYPE.GROUP:
            for sub in _iter_all_shapes(sh.shapes):
                yield sub
        else:
            yield sh

def _delete_shape(shape):
    try:
        el = shape._element
        el.getparent().remove(el)
    except Exception:
        pass

def insert_graphs_by_markers(ppt_path: str, marker_to_image: dict, save_as: str | None = None):
    prs = Presentation(ppt_path)
    placed = 0
    for slide in prs.slides:
        for sh in list(_iter_all_shapes(slide.shapes)):
            if getattr(sh, "has_text_frame", False) and sh.has_text_frame:
                text = "\n".join(p.text for p in sh.text_frame.paragraphs)
                key = _marker_key(text)
                img = marker_to_image.get(key)
                if img and os.path.exists(img):
                    left, top, width, height = sh.left, sh.top, sh.width, sh.height
                    _delete_shape(sh)
                    slide.shapes.add_picture(img, left, top, width=width, height=height)
                    placed += 1
    out = save_as or os.path.join(os.path.dirname(ppt_path), f"Updated_{os.path.basename(ppt_path)}")
    prs.save(out)
    return out, placed

# =========================
# 원본 process_graphs (업로드 파일 경로 기준) + 그래프16/17 & TopASIN 마커 계산
# =========================
def process_graphs(cid_path, asin_path):
    ext = os.path.splitext(cid_path)[1].lower()
    df = pd.read_csv(cid_path) if ext == ".csv" else pd.read_excel(cid_path)
    col = lambda idx: (df.iloc[:, idx].tolist() if len(df.columns) > idx else [])

    dates = col(2)  # C열
    d = col(3); e = col(4); f = col(5); g = col(6); h = col(7); i = col(8); j = col(9)
    marker_to_path = {}
    marker_to_path["그래프1"]  = create_line_by_year(dates, d, "매출", "매출", "Graph 1", unit="K", decimal=1)
    marker_to_path["그래프2"]  = create_line_by_year(dates, e, "고객 유입", "GV", "Graph 2", unit="K", decimal=1)
    marker_to_path["그래프3"]  = create_line_by_year(dates, f, "판매 수량", "Units", "Graph 3", unit="K", decimal=1)
    marker_to_path["그래프4"]  = create_line_by_year(dates, g, "판매상품개수", "Buyable ASIN", "Graph 4", unit="none", decimal=0)
    marker_to_path["그래프5"]  = create_line_by_year(dates, h, "판매 객단가", "ASP", "Graph 5", unit="none", decimal=1)
    i_pct = [parse_number_any(x)*100 if parse_number_any(x) is not None else None for x in i]
    marker_to_path["그래프6"]  = create_line_by_year(dates, i_pct, "구매전환율", "Conversion %", "Graph 6", percentage=True)
    marker_to_path["그래프7"]  = create_line_by_year(dates, j, "SKU당 매출생산성", "GMS per BA", "Graph 7", unit="K", decimal=1)

    k = col(10); l = col(11)
    if g and (k or l):
        marker_to_path["그래프8"] = create_combo_ba_awagv_awas(dates, g, k, l, "Graph 8", "월별 BA / Discoverability / Conversion")
    cg = col(84); ch = col(85)
    if cg and ch:
        marker_to_path["그래프9"] = create_ipi_combo_graph(dates, cg, ch, "Graph 9", "IPI Score / Excess PCT")
    ak = col(36); al = col(37); am = col(38); an = col(39); ao = col(40)
    if d and (ak or al or am or an or ao):
        marker_to_path["그래프10"] = create_merchandising_graph(dates, d, ak, al, am, an, ao, "Graph 10", "Merchandising")

    df_cid = load_cid(cid_path)
    ext2 = os.path.splitext(asin_path)[1].lower()
    df_asin = pd.read_csv(asin_path) if ext2==".csv" else pd.read_excel(asin_path)

    # 그래프 11/12 (기존)
    marker_to_path["그래프11"] = create_graph11_itkbn_dashboard(df_asin, df_cid, None, "Graph 11")
    marker_to_path["그래프12"] = create_graph12_top1_category_trends(df_asin, df_cid, "Graph 12")

    # ===== 신규: Top1/Top2 ASIN 마커 계산 + 그래프 16/17 생성 =====
    compute_top_asin_metrics(df_asin, df_cid)
    top1 = TOPASIN_METRICS.get("top1_asin", "N/A")
    top2 = TOPASIN_METRICS.get("top2_asin", "N/A")
    # 그래프16/17
    g16 = create_top_asin_trend_graph(df_asin, top1, "Graph 16", "Top 1 ASIN 지표 트렌드")
    g17 = create_top_asin_trend_graph(df_asin, top2, "Graph 17", "Top 2 ASIN 지표 트렌드")
    if g16: marker_to_path["그래프16"] = g16
    if g17: marker_to_path["그래프17"] = g17
    # ============================================================

    az = col(51); bk = col(63); bv = col(74)
    if d and (az or bk or bv):
        marker_to_path["그래프18"] = create_ads_tacos_graph(dates, d, az, bk, bv, "Graph 18", "Ads Spend / TACOS")
    bc = col(54); bd = col(55); bn = col(65); bo = col(66); by = col(75); bz = col(76)
    if bc or bn or by:
        marker_to_path["그래프19"] = create_ads_impr_clicks_graph(dates, bc, bd, bn, bo, by, bz, "Graph 19", "Ads Impression / Clicks")
    be = col(56); bp = col(67); ca = col(77)
    if be or bp or ca:
        marker_to_path["그래프20"] = create_three_line_pct(dates, be, bp, ca, ['SP CTR','SB CTR','SD CTR'], 'CTR (%)', "Graph 20", "Ads CTR")
    bf = col(57); bq = col(68); cb = col(79)
    if bf or bq or cb:
        marker_to_path["그래프21"] = create_three_line_pct(dates, bf, bq, cb, ['SP CVR','SB CVR','SD CVR'], 'CVR (%)', "Graph 21", "Ads CVR")
    bg = col(58); br = col(69); cc = col(80)
    if bg or br or cc:
        marker_to_path["그래프22"] = create_three_line_pct_nolabel(dates, bg, br, cc, ['SP CPC','SB CPC','SD CPC'], 'CPC (%)', "Graph 22", "Ads CPC")
    bh = col(59); bs = col(70); cd = col(81)
    if bh or bs or cd:
        marker_to_path["그래프23"] = create_three_line_pct_nolabel(dates, bh, bs, cd, ['SP ACOS','SB ACOS','SD ACOS'], 'ACOS (%)', "Graph 23", "Ads ACOS")
    ba = col(52); bl = col(63); bw = col(74)
    if d and (ba or bl or bw):
        marker_to_path["그래프24"] = create_ads_sales_combo(dates, d, ba, bl, bw, "Graph 24", "Ads Sales")

    marker_to_path = {k:v for k,v in marker_to_path.items() if v}
    return marker_to_path, df_cid, df_asin

# =========================
# Streamlit 실행 버튼
# =========================
if st.button("🚀 PPT 생성하기", type="primary", disabled=not (cid_up and asin_up and ppt_up)):
    with st.spinner("임시 작업 폴더 준비 중..."):
        GRAPH_ROOT = tempfile.mkdtemp(prefix="mbr_")
        graphs_dir = ensure_graphs_folder()
        # 업로드 파일을 temp에 저장
        cid_path  = os.path.join(GRAPH_ROOT, cid_up.name)
        asin_path = os.path.join(GRAPH_ROOT, asin_up.name)
        ppt_path  = os.path.join(GRAPH_ROOT, ppt_up.name)
        with open(cid_path, "wb") as f: f.write(cid_up.getbuffer())
        with open(asin_path, "wb") as f: f.write(asin_up.getbuffer())
        with open(ppt_path, "wb") as f: f.write(ppt_up.getbuffer())

    try:
        with st.spinner("📈 그래프 생성 중 (1~12, 16~17, 18~24)..."):
            marker_to_path, df_cid, df_asin = process_graphs(cid_path, asin_path)
        st.success(f"그래프 생성 완료: {len(marker_to_path)}개")

        with st.spinner("🖼 PPT 템플릿에 그래프 자동 삽입 중..."):
            updated_ppt_path, placed = insert_graphs_by_markers(ppt_path, marker_to_path)

        with st.spinner("🔤 텍스트/표 마커 치환 + YTD 테이블 채우는 중..."):
            prs = Presentation(updated_ppt_path)
            process_ppt_markers(prs, df_cid)              # 폰트/색상 보존 치환
            ytd_cnt = fill_ytd_table_on_slides(prs, df_cid)
            final_path = os.path.join(GRAPH_ROOT, f"Filled_{os.path.basename(ppt_path)}")
            prs.save(final_path)

        st.success(f"완료! 그래프 삽입 {placed}개, YTD 테이블 {ytd_cnt}개 채움.")
        with open(final_path, "rb") as f:
            st.download_button(
                "📥 완성 PPT 다운로드",
                data=f.read(),
                file_name=os.path.basename(final_path),
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
            )

        with st.expander("🔎 디버그(삽입된 그래프 목록)"):
            mp = pd.DataFrame(sorted(marker_to_path.items()), columns=["marker","image_path"])
            st.dataframe(mp, use_container_width=True)

    except Exception as e:
        st.error(f"에러 발생: {e}")
    finally:
        # 필요 시 작업폴더 정리하려면 아래 주석 해제
        # try: shutil.rmtree(GRAPH_ROOT); GRAPH_ROOT=None
        # except: pass
        pass

st.caption("Tip) 템플릿에는 '그래프1' 같은 정확한 텍스트 자리표시자를 넣으세요. 표/텍스트의 {marker}도 자동 치환됩니다.  예: Top1 ASIN {Top1 ASIN} 매출 비중 {Top1 ASIN portion}%, +{Top1 ASIN Growth}% MoM")
