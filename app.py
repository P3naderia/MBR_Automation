import streamlit as st
import os
import re
import io
import tempfile
from datetime import datetime
from typing import Optional, Dict, List
import zipfile

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm

# PowerPoint (python-pptx)
try:
    from pptx import Presentation
    from pptx.util import Cm, Pt
    from pptx.enum.shapes import MSO_SHAPE_TYPE
    PPTX_AVAILABLE = True
except ImportError:
    PPTX_AVAILABLE = False

# =========================
# Global style / Utilities
# =========================
MONTH_LABELS = ['1월','2월','3월','4월','5월','6월','7월','8월','9월','10월','11월','12월']

@st.cache_resource
def _set_korean_font_if_possible():
    """한국어 폰트 설정"""
    try:
        plt.rcParams['font.family'] = ['DejaVu Sans', 'sans-serif']
        plt.rcParams['axes.unicode_minus'] = False
    except:
        pass

# =========================
# Parsing utils
# =========================
_DATE_PATTERNS = [
    ("%Y-%m-%d", r"^\d{4}-\d{1,2}-\d{1,2}$"),
    ("%m/%d/%Y", r"^\d{1,2}/\d{1,2}/\d{4}$"),
    ("%Y/%m/%d", r"^\d{4}/\d{1,2}/\d{1,2}$"),
    ("%Y.%m.%d", r"^\d{4}\.\d{1,2}\.\d{1,2}$"),
]

def parse_date_any(x, default_year=2023):
    if pd.isna(x):
        return None
    s = str(x).strip()
    s = re.sub(r'\s+\d{1,2}:\d{1,2}:\d{1,2}.*$', '', s)
    for fmt, pat in _DATE_PATTERNS:
        if re.match(pat, s):
            try:
                return datetime.strptime(s, fmt)
            except:
                pass
    m1 = re.match(r'^(\d{1,2})/(\d{4})$', s)
    m2 = re.match(r'^(\d{4})[-/\.](\d{1,2})$', s)
    if m1:
        month, year = int(m1.group(1)), int(m1.group(2))
        if 1 <= month <= 12: return datetime(year, month, 1)
    if m2:
        year, month = int(m2.group(1)), int(m2.group(2))
        if 1 <= month <= 12: return datetime(year, month, 1)
    # 영문 월
    month_names = ['jan','feb','mar','apr','may','jun','jul','aug','sep','oct','nov','dec']
    for idx, m in enumerate(month_names, start=1):
        if m in s.lower():
            y_m = re.search(r'(20\d{2}|\d{4})', s)
            year = int(y_m.group(0)) if y_m else default_year
            return datetime(year, idx, 1)
    # 숫자 하나만 있으면 월로 간주
    m = re.search(r'\d+', s)
    if m:
        num = int(m.group(0))
        if 1 <= num <= 12:
            y_m = re.search(r'(20\d{2}|\d{4})', s)
            year = int(y_m.group(0)) if y_m else default_year
            return datetime(year, num, 1)
    return None

def parse_number_any(x, pct_to_100=False):
    if pd.isna(x): 
        return None
    s = str(x).replace(',', '').replace('$', '').strip()
    m = re.search(r'[-+]?\d*\.?\d+', s)
    if not m:
        return None
    val = float(m.group(0))
    if '%' in s:
        return val  # '12.3%' -> 12.3
    return val*100 if pct_to_100 else val

def finalize_year_month(df, year_col='year', month_col='month'):
    if df.empty:
        return df.assign(date_str=pd.Series(dtype=str))
    df[year_col]  = pd.to_numeric(df[year_col], errors='coerce')
    df[month_col] = pd.to_numeric(df[month_col], errors='coerce')
    df = df.dropna(subset=[year_col, month_col]).copy()
    df[year_col]  = df[year_col].astype(int)
    df[month_col] = df[month_col].astype(int)
    df = df.sort_values([year_col, month_col]).reset_index(drop=True)
    df['date_str'] = df[year_col].astype(str) + '-' + df[month_col].astype(str).str.zfill(2)
    return df

def monthly_agg(dates, values, agg='mean', pct_to_100=False):
    rows = []
    n = len(dates)
    for i in range(n):
        dt = parse_date_any(dates[i])
        v  = parse_number_any(values[i], pct_to_100=pct_to_100) if i < len(values) else None
        if dt is None or v is None:
            continue
        rows.append({'year': dt.year, 'month': dt.month, 'value': v})
    if not rows:
        return pd.DataFrame(columns=['year','month','value','date_str'])
    df = pd.DataFrame(rows)
    gp = getattr(df.groupby(['year','month'])['value'], agg)()
    out = gp.reset_index()
    out = finalize_year_month(out, 'year', 'month')
    return out

def _build_month_index(dates):
    pairs = []
    for d in dates:
        dt = parse_date_any(d)
        if dt:
            pairs.append((dt.year, dt.month))
    if not pairs:
        return pd.DataFrame(columns=["date_str","year","month"])
    base = pd.DataFrame(pairs, columns=["year","month"]).drop_duplicates()
    base = finalize_year_month(base, "year", "month")
    return base[["date_str","year","month"]]

def _monthly_df(dates, values, pct_to_100=False, agg='mean'):
    df = monthly_agg(dates, values, agg=agg, pct_to_100=pct_to_100)
    return df[['date_str','value']] if not df.empty else pd.DataFrame(columns=['date_str','value'])

# =========================
# Plot helpers (BI 스타일)
# =========================
PALETTE = {
    "primary": "#2F80ED",
    "green":   "#27AE60",
    "orange":  "#F2994A",
    "purple":  "#9B51E0",
    "red":     "#EB5757",
    "gray":    "#BDBDBD",
    "dark":    "#4F4F4F",
    "sp_fill": "#A9CEF8",
    "sb_fill": "#FAD4AD",
    "sd_fill": "#DCC4F6",
    "ba_fill": "#BFD3F2",
}

def _bi_theme(ax, ygrid=True):
    ax.set_facecolor("white")
    if ax.figure is not None:
        ax.figure.set_facecolor("white")
    for side in ["top", "right"]:
        ax.spines[side].set_visible(False)
    for side in ["left", "bottom"]:
        ax.spines[side].set_color("#BDBDBD")
    ax.tick_params(colors=PALETTE["dark"], labelsize=10)
    if ygrid:
        ax.yaxis.grid(True, color="#E6E6E6", linestyle="-", linewidth=1)
    ax.xaxis.grid(False)

def _yfmt_decimal(dec=1, suffix=""):
    return plt.FuncFormatter(lambda x, pos: f"{x:.{dec}f}{suffix}")

def _yfmt_k(dec=1):
    return plt.FuncFormatter(lambda x, pos: f"{x:.{dec}f}K")

def _label_last(ax, xs, ys, text, dy=6):
    if len(xs) == 0:
        return
    x_last = xs[-1]
    y_last = ys[-1]
    ax.annotate(
        text, (x_last, y_last),
        textcoords="offset points", xytext=(0, dy),
        ha="center", va="bottom",
        fontsize=9,
        bbox=dict(boxstyle="round,pad=0.25", fc="white", ec="#DDDDDD", alpha=0.9),
    )

def _save_fig_to_bytes(fig, name):
    buf = io.BytesIO()
    fig.savefig(buf, format='png', dpi=300, bbox_inches="tight")
    buf.seek(0)
    plt.close(fig)
    return buf, f"{name}.png"

# =========================
# Graph makers
# =========================
def create_line_by_year(dates, values, title, y_label, graph_name,
                        unit="none", decimal=1, percentage=False, annotate_year=None):
    df = monthly_agg(dates, values)
    if df.empty:
        return None, None

    df["value_conv"] = df["value"] / 1000.0 if unit == "K" else df["value"]

    _set_korean_font_if_possible()
    fig = plt.figure(figsize=(10, 6))
    ax = plt.gca()
    _bi_theme(ax)

    colors_cycle = [PALETTE["primary"], PALETTE["green"], PALETTE["orange"], PALETTE["purple"], PALETTE["red"]]

    for idx, y in enumerate(sorted(df["year"].unique())):
        sub = df[df["year"] == y].sort_values("month")
        xs = sub["month"].to_numpy()
        ys = sub["value_conv"].to_numpy()
        ax.plot(xs, ys, linewidth=2.5, marker="o", markersize=5,
                color=colors_cycle[idx % len(colors_cycle)], label=f"{y}년")
        if annotate_year and y == annotate_year:
            if percentage:
                txt = f"{ys[-1]:.{decimal}f}%"
            elif unit == "K":
                txt = f"{ys[-1]:.{decimal}f}K"
            else:
                txt = f"{ys[-1]:.{decimal}f}"
            _label_last(ax, xs, ys, txt)

    ax.set_title(title, fontsize=16, color=PALETTE["dark"])
    ax.set_xlabel("월", fontsize=11, color=PALETTE["dark"])
    ax.set_ylabel(y_label, fontsize=11, color=PALETTE["dark"])
    ax.set_xticks(range(1, 13))
    ax.set_xticklabels(MONTH_LABELS, rotation=0)

    if unit == "K":
        ax.yaxis.set_major_formatter(_yfmt_k(decimal))
    elif percentage:
        ax.yaxis.set_major_formatter(_yfmt_decimal(decimal, "%"))
    else:
        ax.yaxis.set_major_formatter(_yfmt_decimal(decimal))

    ax.legend(loc="upper left", fontsize=10, frameon=False)
    fig.tight_layout()
    return _save_fig_to_bytes(fig, graph_name)

def create_three_line_pct_nolabel(dates, a, b, c, labels, ylabel, graph_name, title):
    rows=[]
    n=len(dates)
    for i in range(n):
        dt = parse_date_any(dates[i])
        if dt is None: 
            continue
        va = parse_number_any(a[i], pct_to_100=True) if i<len(a) else None
        vb = parse_number_any(b[i], pct_to_100=True) if i<len(b) else None
        vc = parse_number_any(c[i], pct_to_100=True) if i<len(c) else None
        rows.append({'year':dt.year,'month':dt.month, 'A':va or 0, 'B':vb or 0, 'C':vc or 0})
    if not rows:
        return None, None

    df = pd.DataFrame(rows).groupby(['year','month']).mean().reset_index()
    df = finalize_year_month(df, 'year', 'month')

    _set_korean_font_if_possible()
    fig = plt.figure(figsize=(12,6)); ax = plt.gca()
    _bi_theme(ax)
    x = np.arange(len(df))
    ax.plot(x, df['A'], 'o-', linewidth=2.5, label=labels[0], color=PALETTE["primary"])
    ax.plot(x, df['B'], 's-', linewidth=2.5, label=labels[1], color=PALETTE["orange"])
    ax.plot(x, df['C'], '^-', linewidth=2.5, label=labels[2], color=PALETTE["purple"])

    ax.set_title(title, fontsize=16, color=PALETTE["dark"])
    ax.set_ylabel(ylabel, color=PALETTE["dark"])
    ax.set_xlabel('연월', color=PALETTE["dark"])
    ax.set_xticks(x); ax.set_xticklabels(df['date_str'], rotation=45, ha='right')
    ax.yaxis.set_major_formatter(_yfmt_decimal(2, "%"))
    ax.legend(loc='upper left', fontsize=10, frameon=False)
    fig.tight_layout()
    return _save_fig_to_bytes(fig, graph_name)

def create_ads_sales_combo(dates, total_sales, sp_sales, sb_sales, sd_sales, graph_name, title):
    rows = []
    n = len(dates)
    for i in range(n):
        dt = parse_date_any(dates[i])
        if dt is None:
            continue
        ts  = parse_number_any(total_sales[i]) if i < len(total_sales) else None
        sp  = parse_number_any(sp_sales[i])    if i < len(sp_sales)    else 0
        sb  = parse_number_any(sb_sales[i])    if i < len(sb_sales)    else 0
        sd  = parse_number_any(sd_sales[i])    if i < len(sd_sales)    else 0
        if ts is None:
            continue
        rows.append({'year': dt.year, 'month': dt.month, 'ts': ts or 0, 'sp': sp or 0, 'sb': sb or 0, 'sd': sd or 0})

    if not rows:
        return None, None

    df = pd.DataFrame(rows).groupby(['year','month']).sum(numeric_only=True).reset_index()
    df = finalize_year_month(df, 'year', 'month')
    df['ad_sales'] = df['sp'] + df['sb'] + df['sd']
    df['ad_sales_pct'] = np.where(df['ts']>0, df['ad_sales']/df['ts']*100, 0)

    _set_korean_font_if_possible()
    fig = plt.figure(figsize=(12,6))
    ax1 = plt.gca()
    _bi_theme(ax1)
    x = np.arange(len(df))

    b1 = ax1.bar(x, df['sp'], label='SP Sales', color=PALETTE["sp_fill"])
    b2 = ax1.bar(x, df['sb'], bottom=df['sp'], label='SB Sales', color=PALETTE["sb_fill"])
    b3 = ax1.bar(x, df['sd'], bottom=df['sp']+df['sb'], label='SD Sales', color=PALETTE["sd_fill"])
    ax1.set_ylabel('Ads Sales', color=PALETTE["dark"])

    ax2 = ax1.twinx()
    _bi_theme(ax2)
    ax2.set_yticks([])
    line, = ax2.plot(x, df['ad_sales_pct'], 'o-', linewidth=2.5, label='Ad sales%', color=PALETTE["primary"])
    for i, v in enumerate(df['ad_sales_pct']):
        ax2.annotate(f"{v:.1f}%", (x[i], v), textcoords='offset points', xytext=(0,10),
                     ha='center', fontsize=9, bbox=dict(boxstyle='round,pad=0.3', fc='white', alpha=0.9))
    ax2.set_ylabel('Ad sales (%)', color=PALETTE["dark"])

    ax1.set_title(title, fontsize=16, color=PALETTE["dark"])
    ax1.set_xlabel('연월', color=PALETTE["dark"])
    ax1.set_xticks(x); ax1.set_xticklabels(df['date_str'], rotation=45, ha='right')
    ax1.legend([b1,b2,b3,line], ['SP Sales','SB Sales','SD Sales','Ad sales%'],
               loc='upper left', fontsize=9, frameon=False)

    fig.tight_layout()
    return _save_fig_to_bytes(fig, graph_name)

# =========================
# PowerPoint helpers
# =========================
def _marker_key(text: str) -> str:
    """'그래프1', '그래프 01', 'Graph 1' 등 → '그래프1'"""
    s = (text or "").strip().lower()
    s = s.replace(" ", "")
    s = s.replace("graph", "그래프")
    m = re.search(r"그래프0*(\d+)", s, flags=re.I)
    return f"그래프{int(m.group(1))}" if m else ""

def _comment_key(text: str) -> str:
    """'코멘트1', 'Comment 1' 등 → '코멘트1'"""
    s = (text or "").strip().lower()
    s = s.replace(" ", "")
    s = s.replace("comment", "코멘트")
    m = re.search(r"코멘트0*(\d+)", s, flags=re.I)
    return f"코멘트{int(m.group(1))}" if m else ""

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

def insert_graphs_by_markers(ppt_data, marker_to_image: dict):
    """PowerPoint에 그래프 삽입"""
    if not PPTX_AVAILABLE:
        return None, 0
    
    prs = Presentation(io.BytesIO(ppt_data))
    placed = 0
    
    for slide in prs.slides:
        for sh in list(_iter_all_shapes(slide.shapes)):
            if getattr(sh, "has_text_frame", False) and sh.has_text_frame:
                text = "\n".join(p.text for p in sh.text_frame.paragraphs)
                key = _marker_key(text)
                img_buf = marker_to_image.get(key)
                if img_buf:
                    left, top, width, height = sh.left, sh.top, sh.width, sh.height
                    _delete_shape(sh)
                    slide.shapes.add_picture(img_buf, left, top, width=width, height=height)
                    placed += 1
    
    output_buf = io.BytesIO()
    prs.save(output_buf)
    output_buf.seek(0)
    return output_buf, placed

def insert_comments_by_markers(ppt_data, comment_map: Dict[str, str]):
    """PowerPoint에 코멘트 삽입"""
    if not PPTX_AVAILABLE:
        return None, 0
        
    prs = Presentation(io.BytesIO(ppt_data))
    filled = 0
    
    for slide in prs.slides:
        for sh in _iter_all_shapes(slide.shapes):
            if getattr(sh, "has_text_frame", False) and sh.has_text_frame:
                raw = "\n".join(p.text for p in sh.text_frame.paragraphs)
                key = _comment_key(raw)
                if key and key in comment_map:
                    tf = sh.text_frame
                    tf.clear()
                    p = tf.paragraphs[0]
                    run = p.add_run()
                    run.text = comment_map[key]
                    try:
                        run.font.size = Pt(12)
                        run.font.bold = False
                        run.font.name = "Malgun Gothic"
                    except:
                        pass
                    filled += 1
    
    output_buf = io.BytesIO()
    prs.save(output_buf)
    output_buf.seek(0)
    return output_buf, filled

# =========================
# AI 컨설팅 코멘트 생성
# =========================
def _fmt(val, unit="%"):
    if val is None or pd.isna(val):
        return "-"
    if unit == "%":
        return f"{val:.1f}%"
    if unit == "K":
        return f"{val/1000:.1f}K"
    return f"{val:.1f}"

def _mom_yoy(df: pd.DataFrame):
    if df.empty or len(df) < 2:
        return None, None
    s = df["value"].astype(float).values
    mom = (s[-1]-s[-2])/abs(s[-2]) * 100 if abs(s[-2])>0 else np.nan
    if len(s) > 12 and abs(s[-13])>0:
        yoy = (s[-1]-s[-13])/abs(s[-13]) * 100
    else:
        yoy = np.nan
    return mom, yoy

def generate_consulting_comments(ctx: Dict[str, List]) -> Dict[str, str]:
    """각 그래프별 핵심 진단 & 액션 제안."""
    comments = {}
    base = _build_month_index(ctx["dates"])

    # 1: 매출
    df1 = _monthly_df(ctx["dates"], ctx["d"], False, "sum")
    mom, yoy = _mom_yoy(df1)
    comments["코멘트1"] = (
        f"매출 최근 {df1['date_str'].iloc[-1] if not df1.empty else '-'} 기준: "
        f"MoM {_fmt(mom)} / YoY {_fmt(yoy)}. "
        "MoM↓면: 프로모션·광고 집중 주간 재배치, 탑키워드 입찰 상향, MD/딜 캘린더 보강."
    )

    # 2: 고객유입(GV)
    df2 = _monthly_df(ctx["dates"], ctx["e"], False, "sum")
    mom, yoy = _mom_yoy(df2)
    comments["코멘트2"] = f"GV 트래픽 MoM {_fmt(mom)}, YoY {_fmt(yoy)}. " \
                         "상위 진입키워드 집중, 브랜드서치 방어, 크리에이티브 테스트 강화."

    # 3: Units
    df3 = _monthly_df(ctx["dates"], ctx["f"], False, "sum")
    mom, yoy = _mom_yoy(df3)
    comments["코멘트3"] = f"판매수량 MoM {_fmt(mom)}, YoY {_fmt(yoy)}. " \
                         "전환경로(장바구니→결제) 이탈구간 점검, 가격/적립/번들 최적화."

    # 4: BA
    df4 = _monthly_df(ctx["dates"], ctx["g"], False, "mean")
    comments["코멘트4"] = "취급상품(BA) 커버리지 확보가 핵심. 핵심 카테고리 신규/재고 안정화, 저회전 ASIN 정리."

    # 5: ASP
    df5 = _monthly_df(ctx["dates"], ctx["h"], False, "mean")
    comments["코멘트5"] = "ASP 변동이 매출/전환에 영향. 가격/쿠폰 전략 A/B, 업셀 번들 구성으로 객단가 방어."

    # 6: CVR
    df6 = _monthly_df(ctx["dates"], ctx["i"], True, "mean")
    mom, yoy = _mom_yoy(df6)
    comments["코멘트6"] = f"구매전환율 MoM {_fmt(mom)}, YoY {_fmt(yoy)}. " \
                         "상세페이지 상단 300px 개선, 리뷰/QA 확보, 배송비/리드타임 노출."

    # 7: GMS/BA
    df7 = _monthly_df(ctx["dates"], ctx["j"], False, "mean")
    comments["코멘트7"] = "SKU당 매출생산성은 롱테일 관리가 핵심. 상위 20% SKU 집중/재고 알람, 하위 라인업 정리."

    # 추가 코멘트들 (8-24)
    comments["코멘트8"] = "발견/전환 비중 최종값 분석. 검색가시성→전환 퍼널 병목을 같이 봐야 함."
    comments["코멘트9"] = "IPI Score와 Excess PCT 관리. 과잉재고 라인 정리 & 리드타임 단축."
    comments["코멘트10"] = "Merch OPS% 최적화. 행사 집중 주차/키워드 정렬, 비효율 딜 축소·전환형 딜 확대."

    # 광고 관련 코멘트 (18-24)
    comments["코멘트18"] = "TACOS 관리 필수. 과다시: 검색어 구조 정리·낮은 ROAS 그룹 축소, 브랜드방어는 유지."
    comments["코멘트19"] = "노출/클릭 총량 관리. CTR↓면: 썸네일·타이틀 A/B, 상위쿼리 맞춤 카피, 비표적 제외키워드 확대."
    comments["코멘트20"] = "CTR 채널별 최적화. SP/SB/SD 각각의 특성에 맞는 크리에이티브 전략 필요."
    comments["코멘트21"] = "CVR 채널별 분석. 랜딩페이지와 상품상세페이지 연계 최적화."
    comments["코멘트22"] = "CPC 관리로 광고효율 극대화. 입찰·매칭타입 정리, 비효율 키워드 정돈."
    comments["코멘트23"] = "ACOS 목표 달성. 목표 ROAS 기반 입찰 자동화·검색어 정리로 하향 안정화."
    comments["코멘트24"] = "광고매출 비중 관리. 과다 의존 시: 오가닉 비중 확대(카탈로그·SEO·MD/딜 밸런싱)."

    return comments

# =========================
# Streamlit App
# =========================
def main():
    st.set_page_config(
        page_title="Amazon Analytics Dashboard",
        page_icon="📊",
        layout="wide"
    )
    
    st.title("📊 Amazon Analytics Dashboard")
    st.markdown("Amazon 비즈니스 데이터 분석 및 자동 리포트 생성 도구")
    
    # 사이드바
    st.sidebar.header("설정")
    
    # 파일 업로드
    uploaded_file = st.file_uploader(
        "데이터 파일을 업로드하세요",
        type=['csv', 'xlsx', 'xls'],
        help="Amazon 비즈니스 데이터가 포함된 CSV 또는 Excel 파일"
    )
    
    # PowerPoint 템플릿 업로드 (선택사항)
    ppt_file = None
    if PPTX_AVAILABLE:
        ppt_file = st.file_uploader(
            "PowerPoint 템플릿 업로드 (선택사항)",
            type=['pptx'],
            help="그래프와 코멘트를 삽입할 PowerPoint 템플릿"
        )
    
    if uploaded_file is not None:
        try:
            # 파일 읽기
            if uploaded_file.name.endswith('.csv'):
                df = pd.read_csv(uploaded_file)
            else:
                df = pd.read_excel(uploaded_file)
            
            st.success(f"파일이 성공적으로 업로드되었습니다: {uploaded_file.name}")
            st.info(f"데이터 크기: {df.shape[0]}행 × {df.shape[1]}열")
            
            # 데이터 미리보기
            with st.expander("데이터 미리보기"):
                st.dataframe(df.head(10))
                
                # 컬럼 정보 표시
                st.subheader("컬럼 정보")
                col_info = []
                key_columns = {
                    2: "날짜 (C열)",
                    3: "총 매출 (D열)", 
                    4: "고객 유입 GV (E열)",
                    5: "판매 수량 (F열)",
                    6: "취급 상품 BA (G열)",
                    7: "평균 판매가 ASP (H열)",
                    8: "전환율 CVR (I열)",
                    9: "GMS/BA (J열)"
                }
                
                for idx, desc in key_columns.items():
                    if idx < len(df.columns):
                        col_info.append({
                            "컬럼": df.columns[idx], 
                            "설명": desc,
                            "샘플": str(df.iloc[0, idx]) if not df.empty else "N/A"
                        })
                
                st.dataframe(pd.DataFrame(col_info))
            
            # 그래프 생성 옵션
            st.header("📈 분석 옵션")
            
            col1, col2 = st.columns(2)
            
            with col1:
                generate_basic = st.checkbox("기본 지표 그래프 생성", value=True,
                    help="매출, GV, Units, BA, ASP, CVR, GMS/BA")
                generate_ads = st.checkbox("광고 관련 그래프 생성", value=True,
                    help="TACOS, 노출/클릭, CTR, CVR, CPC, ACOS")
            
            with col2:
                generate_inventory = st.checkbox("재고/머천다이징 그래프 생성", value=True,
                    help="IPI Score, Excess PCT, Merchandising")
                generate_comments = st.checkbox("AI 컨설팅 코멘트 생성", value=True,
                    help="각 그래프별 분석 및 개선 제안")
            
            if st.button("🚀 분석 시작", type="primary"):
                ppt_data = ppt_file.read() if ppt_file else None
                process_data(df, generate_basic, generate_ads, generate_inventory, generate_comments, ppt_data)
                
        except Exception as e:
            st.error(f"파일 처리 중 오류가 발생했습니다: {str(e)}")
    
    # 사용법 안내
    with st.expander("📋 사용법 안내"):
        st.markdown("""
        ### 사용 방법
        1. **데이터 준비**: Quicksight MBR Dashboard -> raw data (CID, ASIN) 다운로드 
        2. **파일 업로드**: 위의 파일 업로더를 사용하여 데이터 파일 업로드
        3. **템플릿 업로드** : PowerPoint 템플릿이 있다면 함께 업로드
        4. **옵션 선택**: 생성할 그래프 및 분석 옵션 선택
        5. **분석 실행**: '분석 시작' 버튼 클릭
        6. **결과 다운로드**: 생성된 그래프와 보고서를 다운로드
        
        ### 지원되는 지표
        - **기본 지표**: 매출, 고객유입(GV), 판매수량, 객단가, 전환율
        - **광고 지표**: TACOS, 광고비 지출, 노출/클릭, CTR, CVR, CPC, ACOS  
        - **재고 관리**: IPI Score, 과잉재고 비율, 머천다이징 효과
        - **AI 분석**: 자동 진단 및 개선 제안
        
        ### 데이터 형식 요구사항
        - C열: 날짜 (YYYY-MM-DD, MM/DD/YYYY 등)
        - D열: 총 매출
        - E열: 고객 유입 (GV)
        - F열: 판매 수량
        - G열: 취급 상품 수 (BA)
        - H열: 평균 판매가 (ASP)
        - I열: 전환율 (CVR)
        - J열: GMS/BA
        """)

def process_data(df, generate_basic, generate_ads, generate_inventory, generate_comments, ppt_data=None):
    """데이터 처리 및 그래프 생성"""
    
    progress_bar = st.progress(0)
    status_text = st.empty()
    
    try:
        # 데이터 추출 함수
        col = lambda idx: (df.iloc[:, idx].tolist() if len(df.columns) > idx else [])
        
        # 기본 열 정의
        dates = col(2)   # C열
        d = col(3)       # D열 - 총 매출
        e = col(4)       # E열 - GV
        f = col(5)       # F열 - Units
        g = col(6)       # G열 - BA
        h = col(7)       # H열 - ASP
        i = col(8)       # I열 - CVR
        j = col(9)       # J열 - GMS/BA

        # 추가 열들 (광고 및 재고 관련)
        k = col(10); l = col(11)  # AWAGV, AWAS
        cg = col(84); ch = col(85)  # IPI, Excess%
        ak = col(36); al = col(37); am = col(38); an = col(39); ao = col(40)  # Merch
        az = col(51); bk = col(63); bv = col(74)  # Spend
        bc = col(54); bd = col(55); bn = col(65); bo = col(66); by = col(75); bz = col(76)  # Impr/Clicks
        be = col(56); bp = col(67); ca = col(77)  # CTR
        bf = col(57); bq = col(68); cb = col(79)  # CVR
        bg = col(58); br = col(69); cc = col(80)  # CPC
        bh = col(59); bs = col(70); cd = col(81)  # ACOS
        ba = col(52); bl = col(64); bw = col(75)  # Ad Sales
        
        graphs_data = {}
        progress = 0.1
        
        # 기본 그래프 생성
        if generate_basic:
            status_text.text("기본 지표 그래프 생성 중...")
            progress_bar.progress(progress)
            
            # 1-7: 기본 그래프들
            basic_graphs = [
                (dates, d, "매출", "매출 (K)", "Graph_1", {"unit": "K", "decimal": 1}),
                (dates, e, "고객유입", "GV (K)", "Graph_2", {"unit": "K", "decimal": 1}),
                (dates, f, "판매 수량", "Units (K)", "Graph_3", {"unit": "K", "decimal": 1}),
                (dates, g, "판매상품개수", "Buyable ASIN", "Graph_4", {"unit": "none", "decimal": 0}),
                (dates, h, "판매 객단가", "ASP", "Graph_5", {"unit": "none", "decimal": 1}),
            ]
            
            # CVR 그래프 (퍼센트 변환)
            if i:
                i_pct = [parse_number_any(x)*100 if parse_number_any(x) is not None else None for x in i]
                buf, filename = create_line_by_year(dates, i_pct, "구매전환율", "Conversion %", "Graph_6", percentage=True, decimal=1)
                if buf: graphs_data[filename] = buf
            
            # GMS/BA 그래프
            if j:
                buf, filename = create_line_by_year(dates, j, "SKU당 매출생산성", "GMS Contribution per BA (K)", "Graph_7", unit="K", decimal=1)
                if buf: graphs_data[filename] = buf
            
            # 기본 그래프들 생성
            for date_data, value_data, title, ylabel, name, kwargs in basic_graphs:
                if value_data:
                    buf, filename = create_line_by_year(date_data, value_data, title, ylabel, name, **kwargs)
                    if buf: graphs_data[filename] = buf
                
                progress += 0.05
                progress_bar.progress(min(progress, 0.9))

        # 재고/머천다이징 그래프
        if generate_inventory:
            status_text.text("재고/머천다이징 그래프 생성 중...")
            
            # 8: BA + Discoverability/Conversion
            if g and (k or l):
                buf, filename = create_combo_ba_awagv_awas(dates, g, k, l, "Graph_8", "월별 취급 상품 개수 및 판매발생제품 비중")
                if buf: graphs_data[filename] = buf

            # 9: IPI / Excess
            if cg and ch:
                buf, filename = create_ipi_combo_graph(dates, cg, ch, "Graph_9", "IPI Score / Excess PCT")
                if buf: graphs_data[filename] = buf

            # 10: Merchandising
            if d and (ak or al or am or an or ao):
                buf, filename = create_merchandising_graph(dates, d, ak, al, am, an, ao, "Graph_10", "Merchandising")
                if buf: graphs_data[filename] = buf
                
            progress += 0.1
            progress_bar.progress(min(progress, 0.9))

        # 광고 관련 그래프
        if generate_ads:
            status_text.text("광고 관련 그래프 생성 중...")
            
            # 18: Ads Spend/TACOS
            if d and (az or bk or bv):
                buf, filename = create_ads_tacos_graph(dates, d, az, bk, bv, "Graph_18", "Ads Spend / TACOS")
                if buf: graphs_data[filename] = buf

            # 19: Impressions/Clicks
            if bc or bn or by:
                buf, filename = create_ads_impr_clicks_graph(dates, bc, bd, bn, bo, by, bz, "Graph_19", "Ads Impression / Clicks")
                if buf: graphs_data[filename] = buf

            # 20: CTR
            if be or bp or ca:
                buf, filename = create_three_line_pct(dates, be, bp, ca, ['SP CTR','SB CTR','SD CTR'], 'CTR (%)', "Graph_20", "Ads CTR")
                if buf: graphs_data[filename] = buf

            # 21: CVR
            if bf or bq or cb:
                buf, filename = create_three_line_pct(dates, bf, bq, cb, ['SP CVR','SB CVR','SD CVR'], 'CVR (%)', "Graph_21", "Ads CVR")
                if buf: graphs_data[filename] = buf

            # 22: CPC
            if bg or br or cc:
                buf, filename = create_three_line_pct_nolabel(dates, bg, br, cc, ['SP CPC','SB CPC','SD CPC'], 'CPC (%)', "Graph_22", "Ads CPC")
                if buf: graphs_data[filename] = buf

            # 23: ACOS
            if bh or bs or cd:
                buf, filename = create_three_line_pct_nolabel(dates, bh, bs, cd, ['SP ACOS','SB ACOS','SD ACOS'], 'ACOS (%)', "Graph_23", "Ads ACOS")
                if buf: graphs_data[filename] = buf

            # 24: Ads Sales
            if d and (ba or bl or bw):
                buf, filename = create_ads_sales_combo(dates, d, ba, bl, bw, "Graph_24", "Ads Sales")
                if buf: graphs_data[filename] = buf
                
            progress += 0.2
            progress_bar.progress(min(progress, 0.9))

        # AI 코멘트 생성
        if generate_comments:
            status_text.text("AI 컨설팅 코멘트 생성 중...")
            ctx = {
                "dates": dates, "d": d, "e": e, "f": f, "g": g, "h": h, "i": i, "j": j,
                "k": k, "l": l, "cg": cg, "ch": ch,
                "ak": ak, "al": al, "am": am, "an": an, "ao": ao,
                "az": az, "bk": bk, "bv": bv,
                "bc": bc, "bd": bd, "bn": bn, "bo": bo, "by": by, "bz": bz,
                "be": be, "bp": bp, "ca": ca,
                "bf": bf, "bq": bq, "cb": cb,
                "bg": bg, "br": br, "cc": cc,
                "bh": bh, "bs": bs, "cd": cd,
                "ba": ba, "bl": bl, "bw": bw
            }
            comments = generate_consulting_comments(ctx)
            
            # 코멘트를 텍스트 파일로 저장
            comment_text = "# Amazon Analytics 컨설팅 리포트\n\n"
            comment_text += f"생성일시: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n"
            
            for key, value in comments.items():
                graph_no = key.replace("코멘트", "")
                comment_text += f"## 그래프 {graph_no}\n{value}\n\n"
            
            comment_buf = io.BytesIO(comment_text.encode('utf-8'))
            graphs_data["consulting_comments.txt"] = comment_buf
            
            progress += 0.1

        # PowerPoint 처리
        ppt_result = None
        if ppt_data and PPTX_AVAILABLE and graphs_data:
            status_text.text("PowerPoint 템플릿에 그래프 삽입 중...")
            
            # 그래프 마커 매핑
            marker_to_image = {}
            for filename, buf in graphs_data.items():
                if filename.endswith('.png'):
                    graph_name = filename.replace('.png', '').replace('Graph_', '그래프')
                    marker_to_image[graph_name] = buf

            # 그래프 삽입
            ppt_with_graphs, placed_graphs = insert_graphs_by_markers(ppt_data, marker_to_image)
            
            if ppt_with_graphs and generate_comments and comments:
                # 코멘트도 삽입
                ppt_final, placed_comments = insert_comments_by_markers(ppt_with_graphs.getvalue(), comments)
                if ppt_final:
                    graphs_data[f"updated_presentation_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pptx"] = ppt_final
                    ppt_result = f"그래프 {placed_graphs}개, 코멘트 {placed_comments}개 삽입 완료"
            elif ppt_with_graphs:
                graphs_data[f"updated_presentation_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pptx"] = ppt_with_graphs
                ppt_result = f"그래프 {placed_graphs}개 삽입 완료"

        progress_bar.progress(1.0)
        status_text.text("분석 완료!")
        
        # 결과 표시
        st.success(f"총 {len(graphs_data)}개의 파일이 생성되었습니다.")
        
        if ppt_result:
            st.info(f"PowerPoint 처리: {ppt_result}")
        
        # 생성된 그래프 미리보기
        if graphs_data:
            st.header("생성된 그래프 미리보기")
            
            # 탭으로 그래프들 분류
            png_files = [(k, v) for k, v in graphs_data.items() if k.endswith('.png')]
            
            if len(png_files) > 0:
                # 기본 그래프들만 미리보기로 표시 (처음 6개)
                cols = st.columns(2)
                for idx, (filename, buf) in enumerate(png_files[:6]):
                    with cols[idx % 2]:
                        st.subheader(filename.replace('.png', '').replace('_', ' '))
                        if hasattr(buf, 'getvalue'):
                            st.image(buf.getvalue(), use_column_width=True)
                        else:
                            st.image(buf, use_column_width=True)
                
                if len(png_files) > 6:
                    with st.expander(f"추가 그래프 보기 ({len(png_files) - 6}개)"):
                        for idx, (filename, buf) in enumerate(png_files[6:]):
                            st.subheader(filename.replace('.png', '').replace('_', ' '))
                            if hasattr(buf, 'getvalue'):
                                st.image(buf.getvalue(), use_column_width=True)
                            else:
                                st.image(buf, use_column_width=True)

        # 코멘트 미리보기
        if generate_comments and "consulting_comments.txt" in graphs_data:
            with st.expander("AI 컨설팅 코멘트 미리보기"):
                comment_content = graphs_data["consulting_comments.txt"].getvalue().decode('utf-8')
                st.markdown(comment_content)

        # 다운로드 버튼들
        if graphs_data:
            st.header("다운로드")
            
            col1, col2 = st.columns(2)
            
            with col1:
                # 개별 파일 다운로드
                st.subheader("개별 파일 다운로드")
                for filename, buf in graphs_data.items():
                    if hasattr(buf, 'getvalue'):
                        data = buf.getvalue()
                    else:
                        data = buf
                    
                    if filename.endswith('.png'):
                        mime_type = "image/png"
                    elif filename.endswith('.txt'):
                        mime_type = "text/plain"
                    elif filename.endswith('.pptx'):
                        mime_type = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
                    else:
                        mime_type = "application/octet-stream"
                    
                    st.download_button(
                        label=f"📄 {filename}",
                        data=data,
                        file_name=filename,
                        mime=mime_type
                    )
            
            with col2:
                # 전체 ZIP 다운로드
                st.subheader("전체 다운로드")
                zip_buf = io.BytesIO()
                with zipfile.ZipFile(zip_buf, 'w', zipfile.ZIP_DEFLATED) as zip_file:
                    for filename, buf in graphs_data.items():
                        if hasattr(buf, 'getvalue'):
                            data = buf.getvalue()
                        else:
                            data = buf
                        zip_file.writestr(filename, data)
                
                zip_buf.seek(0)
                
                st.download_button(
                    label="📥 모든 파일 다운로드 (ZIP)",
                    data=zip_buf.getvalue(),
                    file_name=f"amazon_analytics_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip",
                    mime="application/zip",
                    type="primary"
                )
                
                st.markdown("*ZIP 파일에는 생성된 모든 그래프, 코멘트, PowerPoint 파일이 포함됩니다.*")

    except Exception as e:
        st.error(f"데이터 처리 중 오류가 발생했습니다: {str(e)}")
        with st.expander("오류 상세 정보"):
            st.exception(e)

if __name__ == "__main__":
    main()

def create_ads_tacos_graph(dates, total_sales, sp_spend, sb_spend, sd_spend, graph_name, title):
    rows = []
    for i in range(len(dates)):
        dt = parse_date_any(dates[i])
        ts = parse_number_any(total_sales[i]) if i < len(total_sales) else None
        sp = parse_number_any(sp_spend[i])   if i < len(sp_spend)   else 0
        sb = parse_number_any(sb_spend[i])   if i < len(sb_spend)   else 0
        sd = parse_number_any(sd_spend[i])   if i < len(sd_spend)   else 0
        if dt is None or ts is None:
            continue
        total_ads = (sp or 0) + (sb or 0) + (sd or 0)
        tacos = total_ads / ts * 100 if ts > 0 else 0
        rows.append({"year": dt.year, "month": dt.month, "sp": sp or 0, "sb": sb or 0, "sd": sd or 0, "tacos": tacos})
    if not rows:
        return None, None
    df = pd.DataFrame(rows).groupby(["year","month"]).mean().reset_index()
    df = finalize_year_month(df, "year", "month")

    _set_korean_font_if_possible()
    fig = plt.figure(figsize=(12, 6))
    ax1 = plt.gca()
    _bi_theme(ax1)
    x = np.arange(len(df))

    sp_k = df["sp"]/1000; sb_k = df["sb"]/1000; sd_k = df["sd"]/1000
    b1 = ax1.bar(x, sp_k, label="SP Spend", color=PALETTE["sp_fill"])
    b2 = ax1.bar(x, sb_k, bottom=sp_k, label="SB Spend", color=PALETTE["sb_fill"])
    b3 = ax1.bar(x, sd_k, bottom=sp_k+sb_k, label="SD Spend", color=PALETTE["sd_fill"])
    ax1.set_ylabel("Ads Spend (K)", color=PALETTE["dark"])
    ax1.yaxis.set_major_formatter(_yfmt_k(1))

    ax2 = ax1.twinx()
    _bi_theme(ax2)
    ax2.set_yticks([])
    line_x = x
    line_y = df["tacos"].to_numpy()
    ax2.plot(line_x, line_y, "o-", color=PALETTE["primary"], linewidth=2.5, label="TACOS")
    _label_last(ax2, line_x, line_y, f"{line_y[-1]:.1f}%")
    ax2.set_ylabel("TACOS (%)", color=PALETTE["dark"])

    ax1.set_title(title, fontsize=16, color=PALETTE["dark"])
    ax1.set_xlabel("연월", color=PALETTE["dark"])
    ax1.set_xticks(x); ax1.set_xticklabels(df["date_str"], rotation=45, ha="right")
    ax1.legend([b1, b2, b3], ["SP Spend","SB Spend","SD Spend"], loc="upper left", fontsize=9, frameon=False)
    fig.tight_layout()
    return _save_fig_to_bytes(fig, graph_name)

def create_ads_impr_clicks_graph(dates, sp_imp, sp_clk, sb_imp, sb_clk, sd_imp, sd_clk, graph_name, title):
    rows = []
    n = len(dates)
    for i in range(n):
        dt = parse_date_any(dates[i])
        if dt is None: 
            continue
        sp_i = parse_number_any(sp_imp[i]) if i < len(sp_imp) else 0
        sp_c = parse_number_any(sp_clk[i]) if i < len(sp_clk) else 0
        sb_i = parse_number_any(sb_imp[i]) if i < len(sb_imp) else 0
        sb_c = parse_number_any(sb_clk[i]) if i < len(sb_clk) else 0
        sd_i = parse_number_any(sd_imp[i]) if i < len(sd_imp) else 0
        sd_c = parse_number_any(sd_clk[i]) if i < len(sd_clk) else 0
        rows.append({"year":dt.year,"month":dt.month,"sp_i":sp_i,"sb_i":sb_i,"sd_i":sd_i,
                     "clicks": (sp_c or 0)+(sb_c or 0)+(sd_c or 0)})
    if not rows: 
        return None, None
    df = pd.DataFrame(rows).groupby(["year","month"]).mean().reset_index()
    df = finalize_year_month(df, "year", "month")

    _set_korean_font_if_possible()
    fig = plt.figure(figsize=(12, 6))
    ax1 = plt.gca()
    _bi_theme(ax1)
    x = np.arange(len(df))

    spk = df["sp_i"]/1000; sbk = df["sb_i"]/1000; sdk = df["sd_i"]/1000
    b1 = ax1.bar(x, spk, label="SP Impression", color=PALETTE["sp_fill"])
    b2 = ax1.bar(x, sbk, bottom=spk, label="SB Impression", color=PALETTE["sb_fill"])
    b3 = ax1.bar(x, sdk, bottom=spk+sbk, label="SD Impression", color=PALETTE["sd_fill"])
    ax1.set_ylabel("Impressions (K)", color=PALETTE["dark"])
    ax1.yaxis.set_major_formatter(_yfmt_k(1))

    ax2 = ax1.twinx()
    _bi_theme(ax2)
    ax2.set_yticks([])
    clicks_k = (df["clicks"]/1000).to_numpy()
    ax2.plot(x, clicks_k, "o-", color=PALETTE["primary"], linewidth=2.5, label="AD Clicks")
    _label_last(ax2, x, clicks_k, f"{clicks_k[-1]:.1f}K")

    ax1.set_title(title, fontsize=16, color=PALETTE["dark"])
    ax1.set_xlabel("연월", color=PALETTE["dark"])
    ax1.set_xticks(x); ax1.set_xticklabels(df["date_str"], rotation=45, ha="right")
    ax1.legend([b1,b2,b3], ["SP Impression","SB Impression","SD Impression"], loc="upper left", fontsize=9, frameon=False)
    fig.tight_layout()
    return _save_fig_to_bytes(fig, graph_name)

def create_three_line_pct(dates, a, b, c, labels, ylabel, graph_name, title):
    rows = []
    n = len(dates)
    for i in range(n):
        dt = parse_date_any(dates[i])
        if dt is None: 
            continue
        va = parse_number_any(a[i], pct_to_100=True) if i<len(a) else None
        vb = parse_number_any(b[i], pct_to_100=True) if i<len(b) else None
        vc = parse_number_any(c[i], pct_to_100=True) if i<len(c) else None
        rows.append({"year":dt.year,"month":dt.month, "A":va or 0, "B":vb or 0, "C":vc or 0})
    if not rows:
        return None, None
    df = pd.DataFrame(rows).groupby(["year","month"]).mean().reset_index()
    df = finalize_year_month(df, "year", "month")

    _set_korean_font_if_possible()
    fig = plt.figure(figsize=(12, 6))
    ax = plt.gca()
    _bi_theme(ax)

    x = np.arange(len(df))
    ax.plot(x, df["A"], "o-", linewidth=2.5, color=PALETTE["primary"], label=labels[0])
    ax.plot(x, df["B"], "s-", linewidth=2.5, color=PALETTE["orange"],  label=labels[1])
    ax.plot(x, df["C"], "^-", linewidth=2.5, color=PALETTE["purple"],  label=labels[2])

    _label_last(ax, x, df["A"].to_numpy(), f"{df['A'].iloc[-1]:.2f}%")

    ax.set_title(title, fontsize=16, color=PALETTE["dark"])
    ax.set_ylabel(ylabel, color=PALETTE["dark"])
    ax.set_xlabel("연월", color=PALETTE["dark"])
    ax.set_xticks(x); ax.set_xticklabels(df["date_str"], rotation=45, ha="right")
    ax.yaxis.set_major_formatter(_yfmt_decimal(2, "%"))
    ax.legend(loc="upper left", fontsize=10, frameon=False)
    fig.tight_layout()
    return _save_fig_to_bytes(fig, graph_name)

def create_ipi_combo_graph(dates, ipi_values, excess_pct_values, graph_name, title):
    rows = []
    for i in range(len(dates)):
        dt = parse_date_any(dates[i])
        ipi = parse_number_any(ipi_values[i]) if i<len(ipi_values) else None
        exc = parse_number_any(excess_pct_values[i], pct_to_100=True) if i<len(excess_pct_values) else None
        if dt is None or ipi is None: 
            continue
        rows.append({"year":dt.year,"month":dt.month,"ipi":ipi,"excess":exc or np.nan})
    if not rows:
        return None, None
    df = pd.DataFrame(rows).groupby(["year","month"]).mean().reset_index()
    df = finalize_year_month(df, "year", "month")

    _set_korean_font_if_possible()
    fig = plt.figure(figsize=(10, 6))
    ax1 = plt.gca()
    _bi_theme(ax1)
    x = np.arange(len(df))

    bars = ax1.bar(x, df["ipi"], alpha=0.9, color=PALETTE["ba_fill"], label="IPI Score")
    ax1.set_ylabel("IPI Score", color=PALETTE["dark"])

    ax2 = ax1.twinx()
    _bi_theme(ax2)
    ax2.set_yticks([])
    y2 = df["excess"].to_numpy()
    ax2.plot(x, y2, "o-", linewidth=2.5, color=PALETTE["orange"], label="Excess PCT")
    _label_last(ax2, x, y2, f"{y2[-1]:.1f}%")
    ax2.set_ylabel("Excess PCT (%)", color=PALETTE["dark"])

    ax1.set_title(title, fontsize=16, color=PALETTE["dark"])
    ax1.set_xlabel("연월", color=PALETTE["dark"])
    ax1.set_xticks(x); ax1.set_xticklabels(df["date_str"], rotation=45, ha="right")
    ax1.legend([bars], ["IPI Score"], loc="upper left", fontsize=10, frameon=False)
    fig.tight_layout()
    return _save_fig_to_bytes(fig, graph_name)

def create_combo_ba_awagv_awas(dates, ba_values, awagv_values, awas_values, graph_name, title):
    rows = []
    n = len(dates)
    for i in range(n):
        dt = parse_date_any(dates[i])
        if dt is None: 
            continue
        ba   = parse_number_any(ba_values[i]) if i<len(ba_values) else None
        awag = parse_number_any(awagv_values[i]) if i<len(awagv_values) else None
        awas = parse_number_any(awas_values[i]) if i<len(awas_values) else None
        if ba is None or ba <= 0: 
            continue
        disc = (awag/ba*100) if awag is not None else np.nan
        conv = (awas/ba*100) if awas is not None else np.nan
        rows.append({"year":dt.year,"month":dt.month,"ba":ba,"disc":disc,"conv":conv})
    if not rows:
        return None, None
    df = pd.DataFrame(rows).groupby(["year","month"]).mean().reset_index()
    df = finalize_year_month(df, "year", "month")

    _set_korean_font_if_possible()
    fig = plt.figure(figsize=(10, 6))
    ax1 = plt.gca()
    _bi_theme(ax1)
    x = np.arange(len(df))

    bars = ax1.bar(x, df["ba"], alpha=0.9, color=PALETTE["ba_fill"], label="BA")
    ax1.set_ylabel("취급 상품 개수 (개)", color=PALETTE["dark"])

    ax2 = ax1.twinx()
    _bi_theme(ax2)
    ax2.set_yticks([])
    l1 = ax2.plot(x, df["disc"], "o-", linewidth=2.5, color=PALETTE["primary"], label="ASIN Discoverability")
    l2 = ax2.plot(x, df["conv"], "s-", linewidth=2.5, color=PALETTE["orange"],  label="ASIN Conversion")
    if pd.notna(df["conv"].iloc[-1]):
        _label_last(ax2, x, df["conv"].to_numpy(), f"{df['conv'].iloc[-1]:.1f}%")
    ax2.set_ylabel("비율 (%)", color=PALETTE["dark"])

    ax1.set_title(title, fontsize=16, color=PALETTE["dark"])
    ax1.set_xlabel("연월", color=PALETTE["dark"])
    ax1.set_xticks(x); ax1.set_xticklabels(df["date_str"], rotation=45, ha="right")
    ax1.legend([bars], ["BA"], loc="upper left", fontsize=9, frameon=False)
    fig.tight_layout()
    return _save_fig_to_bytes(fig, graph_name)

def create_merchandising_graph(dates, total_sales, bd_ops, ld_ops, dotd_ops, mario_ops, coupon_ops, graph_name, title):
    rows=[]
    n=len(dates)
    for i in range(n):
        dt = parse_date_any(dates[i])
        ts = parse_number_any(total_sales[i]) if i<len(total_sales) else None
        if dt is None or ts is None: 
            continue
        bd = parse_number_any(bd_ops[i])    if i<len(bd_ops)    else 0
        ld = parse_number_any(ld_ops[i])    if i<len(ld_ops)    else 0
        do = parse_number_any(dotd_ops[i])  if i<len(dotd_ops)  else 0
        ma = parse_number_any(mario_ops[i]) if i<len(mario_ops) else 0
        cp = parse_number_any(coupon_ops[i])if i<len(coupon_ops)else 0
        total = (bd or 0)+(ld or 0)+(do or 0)+(ma or 0)+(cp or 0)
        pct = total/ts*100 if ts>0 else 0
        rows.append({"year":dt.year,"month":dt.month,"bd":bd or 0,"ld":ld or 0,"do":do or 0,"ma":ma or 0,"cp":cp or 0,"pct":pct})
    if not rows:
        return None, None
    df = pd.DataFrame(rows).groupby(["year","month"]).mean().reset_index()
    df = finalize_year_month(df, "year", "month")

    _set_korean_font_if_possible()
    fig = plt.figure(figsize=(12, 6))
    ax1 = plt.gca()
    _bi_theme(ax1)
    x = np.arange(len(df))

    bd, ld, do, ma, cp = df["bd"]/1000, df["ld"]/1000, df["do"]/1000, df["ma"]/1000, df["cp"]/1000
    b1 = ax1.bar(x, bd, label="Best Deal", color=PALETTE["sp_fill"])
    b2 = ax1.bar(x, ld, bottom=bd, label="Lightning Deal", color=PALETTE["sb_fill"])
    b3 = ax1.bar(x, do, bottom=bd+ld, label="Deal of The Day", color=PALETTE["sd_fill"])
    b4 = ax1.bar(x, ma, bottom=bd+ld+do, label="Prime Exclusive Discount", color="#C7EBD0")
    b5 = ax1.bar(x, cp, bottom=bd+ld+do+ma, label="Coupon", color="#F8E3A2")
    ax1.set_ylabel("Merchandising OPS (K)", color=PALETTE["dark"])
    ax1.yaxis.set_major_formatter(_yfmt_k(1))

    ax2 = ax1.twinx()
    _bi_theme(ax2)
    ax2.set_yticks([])
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
    return _save_fig_to_bytes(fig, graph_name)
