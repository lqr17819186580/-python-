#!/usr/bin/env python3
"""
run_full_pipeline_merge.py — Merged version of run_full_pipeline.py and all dependencies.

This script combines:
  - run_full_pipeline.py (orchestration)
  - update_ppt.py (main PPT update)
  - apply_w14_patches.py (weekly patches)
  - final_gap_patch.py (final fixes)
  - data_loader.py (CSV parsing)
  - chart_updates.py (chart data generation)
  - sowhat_slots.py (So What slot discovery)
  - sowhat_generators.py (So What text generation)
  - chart_xml_patch.py (XML-based chart updates)

Input files (must be in same dir):
  - 周业绩汇报PPT_20260404_更新.pptx  (or any template renamed to template.pptx)
  - S1-总览仪表盘.csv
  - S2-业务端视角.csv
  - S3-执行管理端.csv
  - S4-产品端视角.csv

Output:
  - 周业绩汇报PPT_FINAL.pptx

Pipeline stages:
  1. Stage 1 — update_ppt logic: charts, KPI text, slide-7 W14 column clone
  2. Stage 2 — apply_w14_patches logic: KA tables, U/V charts, bank dashboard, W table, etc.
  3. Stage 3 — final_gap_patch logic: slide-3 license table, slide-3 TOP10 product table,
               slide-7 Q/R/S heat matrix rebuild, T-table dedupe
"""
import shutil
import sys
import csv
import re
import io
import zipfile
import os
from pathlib import Path
from dataclasses import dataclass, field
from typing import Optional
from copy import deepcopy

# Third-party imports
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.util import Emu, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR, MSO_AUTO_SIZE
from pptx.chart.plot import PlotTypeInspector
from pptx.enum.chart import XL_CHART_TYPE as XL
import pandas as pd
from lxml import etree

HERE = Path(__file__).parent

TEMPLATE_SRC = HERE / "周业绩汇报PPT_20260404_更新.pptx"
if not TEMPLATE_SRC.exists():
    TEMPLATE_SRC = HERE / "template.pptx"

STAGE1_OUT = HERE / "周业绩汇报PPT_AUTO_UPDATED.pptx"
STAGE2_OUT = HERE / "周业绩汇报PPT_AUTO_UPDATED_updated.pptx"
STAGE3_OUT = HERE / "周业绩汇报PPT_FINAL.pptx"


# ============================================================================
# PART 1: data_loader.py — CSV parsing utilities
# ============================================================================
HEADER_RE = re.compile(r"^([A-Z](?:-[^.\s]+)?)\.\s")

def _read_text_any_encoding(path: str) -> str:
    data = Path(path).read_bytes()
    for enc in ("utf-8-sig", "utf-8", "gb18030", "gbk", "cp936"):
        try:
            return data.decode(enc)
        except UnicodeDecodeError:
            continue
    return data.decode("utf-8", errors="replace")

def parse_blocks(path: str) -> dict:
    raw = _read_text_any_encoding(path).splitlines()
    blocks = {}
    cur_lines, cur_key = [], None
    for line in raw:
        stripped = line.lstrip("\ufeff")
        m = HEADER_RE.match(stripped)
        if m:
            if cur_key is not None:
                blocks[cur_key] = cur_lines
            cur_key = m.group(1)
            cur_lines = []
        elif cur_key is not None:
            cur_lines.append(line)
    if cur_key is not None:
        blocks[cur_key] = cur_lines

    out = {}
    for key, lines in blocks.items():
        data_lines = []
        for ln in lines:
            stripped = ln.strip()
            if not stripped:
                continue
            if stripped.startswith("📌"):
                continue
            first_char = stripped.lstrip('"').lstrip("'")[:1]
            if first_char in ("\ufffd", "?"):
                if stripped.count(",") < 4 or "\ufffd\ufffd" in stripped or "???" in stripped:
                    continue
            data_lines.append(ln)
        if len(data_lines) < 2:
            continue
        csv_text = "\n".join(data_lines)
        try:
            df = pd.read_csv(io.StringIO(csv_text), dtype=str, keep_default_na=False)
        except Exception:
            continue
        for c in df.columns:
            df[c] = df[c].astype(str).str.strip().str.strip('"')
        df.columns = [c.strip() for c in df.columns]
        df = df.loc[:, ~(df.columns.str.match(r"^Unnamed"))]
        out[key] = df
    return out

def num(v) -> float:
    if v is None:
        return 0.0
    s = str(v).strip().strip('"').replace(",", "").replace("%", "")
    if s in ("", "-", "—", "nan", "None"):
        return 0.0
    try:
        return float(s)
    except ValueError:
        return 0.0

def load_all(s1_path, s2_path, s3_path, s4_path):
    return (parse_blocks(s1_path), parse_blocks(s2_path),
            parse_blocks(s3_path), parse_blocks(s4_path))


# ============================================================================
# PART 2: chart_xml_patch.py — XML-based chart updates
# ============================================================================
NS = {
    "c":  "http://schemas.openxmlformats.org/drawingml/2006/chart",
    "a":  "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r":  "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}
C = NS["c"]

def _qn(tag):
    return f"{{{C}}}{tag}"

def _find_series(root):
    return root.findall(".//c:ser", NS)

def _rewrite_cache(cache_elem, new_values, numeric=True):
    for pt in list(cache_elem.findall("c:pt", NS)):
        cache_elem.remove(pt)
    pc = cache_elem.find("c:ptCount", NS)
    if pc is None:
        pc = etree.SubElement(cache_elem, _qn("ptCount"))
        cache_elem.insert(0, pc)
    pc.set("val", str(len(new_values)))
    for i, v in enumerate(new_values):
        pt = etree.SubElement(cache_elem, _qn("pt"))
        pt.set("idx", str(i))
        vel = etree.SubElement(pt, _qn("v"))
        if v is None:
            vel.text = ""
        elif numeric:
            vel.text = f"{float(v)}"
        else:
            vel.text = str(v)

def _rewrite_ref_with_cache(ref_elem, new_values, numeric=True):
    cache_name = "numCache" if numeric else "strCache"
    cache = ref_elem.find(f"c:{cache_name}", NS)
    if cache is None:
        cache = etree.SubElement(ref_elem, _qn(cache_name))
    _rewrite_cache(cache, new_values, numeric=numeric)

def _rewrite_series_values(ser, new_values):
    val = ser.find("c:val", NS)
    if val is None:
        return
    ref = val.find("c:numRef", NS)
    if ref is None:
        lit = val.find("c:numLit", NS)
        if lit is not None:
            _rewrite_cache(lit, new_values, numeric=True)
        return
    _rewrite_ref_with_cache(ref, new_values, numeric=True)

def _rewrite_series_categories(ser, new_cats):
    cat = ser.find("c:cat", NS)
    if cat is None:
        return
    ref = cat.find("c:strRef", NS)
    if ref is None:
        ref = cat.find("c:numRef", NS)
    if ref is None:
        lit = cat.find("c:strLit", NS)
        if lit is not None:
            _rewrite_cache(lit, new_cats, numeric=False)
        return
    numeric = (ref.tag == _qn("numRef"))
    _rewrite_ref_with_cache(ref, new_cats, numeric=numeric)

def _rewrite_series_name(ser, new_name):
    if new_name is None:
        return
    tx = ser.find("c:tx", NS)
    if tx is None:
        return
    ref = tx.find("c:strRef", NS)
    if ref is None:
        v = tx.find("c:v", NS)
        if v is not None:
            v.text = new_name
        return
    cache = ref.find("c:strCache", NS)
    if cache is None:
        cache = etree.SubElement(ref, _qn("strCache"))
    _rewrite_cache(cache, [new_name], numeric=False)

def patch_chart(chart_part, series_spec, categories=None, series_names=None):
    root = chart_part._element
    series = _find_series(root)
    n = min(len(series), len(series_spec))
    for i in range(n):
        ser = series[i]
        _rewrite_series_values(ser, series_spec[i])
        if categories is not None:
            _rewrite_series_categories(ser, categories)
        if series_names is not None and i < len(series_names):
            _rewrite_series_name(ser, series_names[i])

def patch_chart_dlbls(chart_part, series_index, labels):
    root = chart_part._element
    series = _find_series(root)
    if series_index >= len(series):
        return 0
    ser = series[series_index]
    dlbls = ser.findall("c:dLbls/c:dLbl", NS)
    hits = 0
    for dlbl in dlbls:
        idx_elem = dlbl.find("c:idx", NS)
        if idx_elem is None:
            continue
        try:
            idx = int(idx_elem.get("val"))
        except (TypeError, ValueError):
            continue
        if idx >= len(labels):
            continue
        new_text = labels[idx]
        tx = dlbl.find("c:tx", NS)
        if tx is None:
            continue
        rich = tx.find("c:rich", NS)
        if rich is None:
            continue
        p = rich.find("a:p", NS)
        if p is None:
            continue
        runs = p.findall("a:r", NS)
        if not runs:
            continue
        first_run = runs[0]
        first_t = first_run.find("a:t", NS)
        if first_t is None:
            first_t = etree.SubElement(first_run, f"{{{NS['a']}}}t")
        first_t.text = new_text
        for extra in runs[1:]:
            p.remove(extra)
        hits += 1
    return hits


# ============================================================================
# PART 3: chart_updates.py — CategoryChartData builders
# ============================================================================
def _row(df, col_match_value, col_name):
    first_col = df.columns[0]
    sub = df[df[first_col] == col_match_value]
    if len(sub) == 0:
        return 0.0
    return num(sub.iloc[0][col_name])

def _total_row(df):
    first_col = df.columns[0]
    sub = df[df[first_col] == "合计"]
    return sub.iloc[0] if len(sub) else None

def to_m(v):
    return round(num(v) / 1_000_000, 2)

def to_w(v):
    return round(num(v) / 10_000, 2)

def slide1_chart_business_type(S1):
    g = S1["G"]
    cats = ["经代业务", "代理人业务", "KA 业务"]
    mapping = {"经代业务": "经代业务", "代理人业务": "代理人业务", "KA 业务": "KA业务"}
    cd = CategoryChartData()
    cd.categories = cats
    cd.add_series("批核 APE", [to_m(_row(g, mapping[c], "2026批核APE")) for c in cats])
    cd.add_series("未批核 APE", [to_m(_row(g, mapping[c], "未批核APE")) for c in cats])
    cd.add_series("待签 APE", [to_m(_row(g, mapping[c], "待签APE")) for c in cats])
    return cd

def slide1_chart_monthly_trend(S1):
    cC, cD, cE = S1["C"], S1["D"], S1["E"]
    months = [m for m in cC.iloc[:, 0] if m and m.startswith("202")]
    def label(ym):
        y, m = ym.split("-")
        mn = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]
        return f"{y[-2:]}-{mn[int(m)-1]}"
    cats = [label(m) for m in months]
    def col_for(df, ym, col):
        row = df[df.iloc[:, 0] == ym]
        return to_m(row.iloc[0][col]) if len(row) else 0.0
    cd = CategoryChartData()
    cd.categories = cats
    cd.add_series("预约 APE(M)", [col_for(cC, m, "APE") for m in months])
    cd.add_series("签单 APE(M)", [col_for(cD, m, "APE") for m in months])
    cd.add_series("批核 APE(M)", [col_for(cE, m, "APE") for m in months])
    return cd

def _slide2_bars(S1, block_letter, series_name):
    df = S1[block_letter]
    df26 = df[df.iloc[:, 0].astype(str).str.startswith("2026")]
    cats = ["1月", "2月", "3月", "4月"]
    vals = [0.0] * 4
    for _, r in df26.iterrows():
        m_idx = int(r.iloc[0].split("-")[1]) - 1
        if 0 <= m_idx < 4:
            vals[m_idx] = to_m(r["APE"])
    cd = CategoryChartData()
    cd.categories = cats
    cd.add_series(series_name, vals)
    return cd, vals

def slide2_appointment(S1):
    cd, _ = _slide2_bars(S1, "C", "预约 APE(M)")
    return cd

def slide2_signed(S1):
    cd, _ = _slide2_bars(S1, "D", "签单 APE(M)")
    return cd

def slide2_approved(S1):
    cd, _ = _slide2_bars(S1, "E", "批核 APE(M)")
    return cd

def slide2_forecast_chart(S1):
    cE = S1["E"]
    actuals = [0.0] * 12
    for _, r in cE.iterrows():
        ym = str(r.iloc[0])
        if ym.startswith("2026-"):
            m_idx = int(ym.split("-")[1]) - 1
            if 0 <= m_idx < 12:
                actuals[m_idx] = to_m(r["APE"])
    forecast = [None, None, None, None, 80.0, 88.0, 95.0, 100.0, 100.0, 100.0, 100.0, 96.0]
    cd = CategoryChartData()
    cd.categories = ["1月","2月","3月","4月","5月","6月","7月","8月","9月","10月","11月","12月"]
    cd.add_series("实际批核 APE(M)", [v if i < 4 else None for i, v in enumerate(actuals)])
    cd.add_series("预测 APE(M)", forecast)
    return cd

def slide3_sunlife_trend(S2):
    fA, gA, hA = S2["F-APE"], S2["G-APE"], S2["H-APE"]
    cats = ["26-Jan", "26-Feb", "26-Mar", "26-Apr"]
    months = ["2026-01", "2026-02", "2026-03", "2026-04"]
    def total(df, col):
        row = df[df.iloc[:, 0] == "合计"]
        return to_m(row.iloc[0][col]) if len(row) else 0.0
    cd = CategoryChartData()
    cd.categories = cats
    cd.add_series("预约 APE(M)", [total(fA, m) for m in months])
    cd.add_series("签单 APE(M)", [total(gA, m) for m in months])
    cd.add_series("批核 APE(M)", [total(hA, m) for m in months])
    return cd

def slide4_channel_trend(S2, channel_name):
    cA, cD, cE = S2["C-APE"], S2["D-APE"], S2["E-APE"]
    months = ["2026-01", "2026-02", "2026-03", "2026-04"]
    def val(df, col):
        row = df[df.iloc[:, 0] == channel_name]
        return to_m(row.iloc[0][col]) if len(row) else 0.0
    cd = CategoryChartData()
    cd.categories = ["Jan", "Feb", "Mar", "Apr"]
    cd.add_series("预约", [val(cA, m) for m in months])
    cd.add_series("签单", [val(cD, m) for m in months])
    cd.add_series("批核", [val(cE, m) for m in months])
    return cd

def slide5_donut(S2):
    total = _total_row(S2["A"])
    cd = CategoryChartData()
    cd.categories = ["已批核", "未批核", "待签"]
    cd.add_series("APE构成", [to_m(total["2026批核APE"]), to_m(total["未批核APE"]), to_m(total["待签APE"])])
    return cd

def slide5_target_vs_actual(S2):
    df = S2["A"]
    df = df[df["业务细分"].isin(["永明经代", "天领业务", "BK业务", "合伙转介业务", "成事家办", "同行经代", "ICLUB业务", "IFA业务"])]
    name_map = {"永明经代": "永明经代", "天领业务": "天领业务", "BK业务": "BK业务",
                "合伙转介业务": "合伙转介", "成事家办": "成事家办", "同行经代": "同行经代",
                "ICLUB业务": "ICLUB", "IFA业务": "IFA业务"}
    order = ["永明经代", "天领业务", "BK业务", "合伙转介业务", "成事家办", "同行经代", "ICLUB业务", "IFA业务"]
    cats = [name_map[k] for k in order]
    issued, unbat, pend, gap = [], [], [], []
    for k in order:
        row = df[df["业务细分"] == k].iloc[0]
        i = to_m(row["2026批核APE"]); u = to_m(row["未批核APE"]); p = to_m(row["待签APE"]); t = to_m(row["目标APE"])
        issued.append(i); unbat.append(u); pend.append(p); gap.append(round(t - (i + u + p), 2))
    cd = CategoryChartData()
    cd.categories = cats
    cd.add_series("已批核 APE(M)", issued)
    cd.add_series("未批核 APE(M)", unbat)
    cd.add_series("待签 APE(M)", pend)
    cd.add_series("目标缺口 APE(M)", gap)
    return cd

def slide6_weekly_trend(S3):
    a = S3["A-APE"]
    week_cols = [c for c in a.columns if c.startswith("2026W")]
    cats = [c.replace("2026", "") for c in week_cols]
    cd = CategoryChartData()
    cd.categories = cats
    for stage in ["预约", "签单", "递交", "批核"]:
        row = a[a.iloc[:, 0] == stage].iloc[0]
        cd.add_series(stage, [to_m(row[c]) for c in week_cols])
    return cd

def slide8_referrer(S2):
    j = S2["J"]
    j = j[j.iloc[:, 0] != "合计"].copy()
    j["_sort"] = j["2026批核APE"].apply(num)
    j = j.sort_values("_sort", ascending=False)
    cats = j.iloc[:, 0].tolist()
    cd = CategoryChartData()
    cd.categories = cats
    cd.add_series("批核APE(万)", [to_w(v) for v in j["2026批核APE"]])
    cd.add_series("未批核APE(万)", [to_w(v) for v in j["未批核APE"]])
    cd.add_series("待签APE(万)", [to_w(v) for v in j["待签APE"]])
    return cd

def slide8_top10_ka(S2):
    k = S2["K"]
    k = k[k.iloc[:, 0] != "合计"].copy()
    k["_sort"] = k["2026批核APE"].apply(num)
    k = k.sort_values("_sort", ascending=False).head(10)
    cats = k.iloc[:, 0].tolist()
    cd = CategoryChartData()
    cd.categories = cats
    cd.add_series("批核APE(万)", [to_w(v) for v in k["2026批核APE"]])
    cd.add_series("未批核APE(万)", [to_w(v) for v in k["未批核APE"]])
    cd.add_series("待签APE(万)", [to_w(v) for v in k["待签APE"]])
    return cd

def slide9_peer_weekly(S3):
    j, k, l = S3["J-APE"], S3["K-APE"], S3["L-APE"]
    week_cols = [c for c in j.columns if c.startswith("2026W")]
    cats = [c.replace("2026", "") for c in week_cols]
    def totals(df):
        row = df[df.iloc[:, 0] == "合计"].iloc[0]
        return [to_m(row[c]) for c in week_cols]
    cd = CategoryChartData()
    cd.categories = cats
    cd.add_series("同行预约(M)", totals(j))
    cd.add_series("同行签单(M)", totals(k))
    cd.add_series("同行批核(M)", totals(l))
    return cd

def _slide10_bank_monthly(df):
    _MO_ZH = {1:"1月",2:"2月",3:"3月",4:"4月",5:"5月",6:"6月",7:"7月",8:"8月",9:"9月",10:"10月",11:"11月",12:"12月"}
    months = sorted([c for c in df.columns if c.startswith("2026-")])
    cats = [_MO_ZH.get(int(c[5:7]), c) for c in months]
    def val(bank, month):
        row = df[df.iloc[:, 0] == bank]
        return to_m(row.iloc[0][month]) if len(row) else 0.0
    cd = CategoryChartData()
    cd.categories = cats
    cd.add_series("民生银行", [val("民生银行", m) for m in months])
    cd.add_series("平安银行", [val("平安银行", m) for m in months])
    return cd

def slide10_bk_appointment(S2): return _slide10_bank_monthly(S2["P-APE"])
def slide10_bk_signed(S2): return _slide10_bank_monthly(S2["Q-APE"])
def slide10_bk_approved(S2): return _slide10_bank_monthly(S2["R-APE"])

def slide10_bk_ka(S2):
    o = S2["O"]
    cats = ["民生银行", "平安银行"]
    cd = CategoryChartData()
    cd.categories = cats
    cd.add_series("批核APE(M)", [to_m(_row(o, c, "2026批核APE")) for c in cats])
    cd.add_series("未批核APE(M)", [to_m(_row(o, c, "未批核APE")) for c in cats])
    cd.add_series("待签APE(M)", [to_m(_row(o, c, "待签APE")) for c in cats])
    return cd

def slide10_donut_target(S2):
    bk = S2["A"][S2["A"]["业务细分"] == "BK业务"].iloc[0]
    issued = to_m(bk["2026批核APE"])
    remain = to_m(bk["目标APE"]) - issued
    cd = CategoryChartData()
    cd.categories = ["已批核", "目标剩余"]
    cd.add_series("APE构成", [issued, remain])
    return cd

def slide11_bank_weekly(S3):
    m, n, o = S3["M-APE"], S3["N-APE"], S3["O-APE"]
    week_cols = [c for c in m.columns if c.startswith("2026W")]
    cats = [c.replace("2026", "") for c in week_cols]
    def totals(df):
        row = df[df.iloc[:, 0] == "合计"].iloc[0]
        return [to_m(row[c]) for c in week_cols]
    cd = CategoryChartData()
    cd.categories = cats
    cd.add_series("银行预约(M)", totals(m))
    cd.add_series("银行签单(M)", totals(n))
    cd.add_series("银行批核(M)", totals(o))
    return cd

def slide11_branch_ranking(S2):
    s = S2["S-APE"]
    s = s[s.iloc[:, 0] != "合计"].copy()
    s["_sort"] = s["合计"].apply(num)
    s = s.sort_values("_sort", ascending=False)
    s = s[s["_sort"] > 0]
    cats = s.iloc[:, 0].tolist()
    vals = [to_m(v) for v in s["合计"]]
    cd = CategoryChartData()
    cd.categories = cats
    cd.add_series("批核APE(M)", vals)
    return cd


# ============================================================================
# PART 4: sowhat_slots.py — So What slot discovery
# ============================================================================
def find_sowhat_slots(prs):
    slots = []
    for si, s in enumerate(prs.slides):
        labels = [sh for sh in s.shapes if sh.has_text_frame and sh.text_frame.text.strip().startswith('So What')]
        content_pool = []
        for sh in s.shapes:
            if not sh.has_text_frame:
                continue
            t = sh.text_frame.text.strip()
            if len(t) < 40 or t.startswith('So What'):
                continue
            content_pool.append(sh)
        assigned = set()
        for sl in labels:
            rest = sl.text_frame.text.strip().replace('So What：', '', 1).strip()
            if len(rest) > 20:
                slots.append({'slide': si, 'label': sl, 'content': sl, 'type': 'MERGED'})
                continue
            lcx = sl.left + (sl.width or 0) // 2
            ly = sl.top
            best = None
            best_score = None
            for sh in content_pool:
                if id(sh) in assigned:
                    continue
                cy = sh.top
                if cy < ly - 20000 or cy > ly + 400000:
                    continue
                cx = sh.left + (sh.width or 0) // 2
                left_edge = sh.left
                right_edge = sh.left + (sh.width or 0)
                contains = left_edge <= lcx <= right_edge
                if not contains and abs(cx - lcx) > 3_000_000:
                    continue
                score = (cy - ly) * 5
                if not contains:
                    score += abs(cx - lcx)
                if best_score is None or score < best_score:
                    best = sh
                    best_score = score
            if best is not None:
                slots.append({'slide': si, 'label': sl, 'content': best, 'type': 'SEPARATE'})
                assigned.add(id(best))
            else:
                slots.append({'slide': si, 'label': sl, 'content': None, 'type': 'ORPHAN'})
    return slots


# ============================================================================
# PART 5: apply_w14_patches.py — Weekly patches (shared with sowhat_generators)
# ============================================================================
_SECTION_HEADER_RE = re.compile(r'^[A-Z](?:-[A-Za-z\u4e00-\u9fff]+)?\.?\s')

def w14_load_rows(path: Path):
    raw_bytes = Path(path).read_bytes()
    text = None
    for enc in ('utf-8-sig', 'utf-8', 'gb18030', 'gbk', 'cp936'):
        try:
            text = raw_bytes.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    if text is None:
        text = raw_bytes.decode('utf-8', errors='replace')
    return list(csv.reader(io.StringIO(text)))

def w14_find_section_start(rows, marker: str) -> int:
    for i, r in enumerate(rows):
        if r and r[0].strip().startswith(marker):
            return i
    return -1

def w14_find_header_row(rows, start: int):
    for j in range(start + 1, min(start + 5, len(rows))):
        r = rows[j]
        if not r or not r[0].strip(): continue
        c0 = r[0].strip()
        if c0.startswith('📌'): continue
        first_char = c0.lstrip('"').lstrip("'")[:1]
        if first_char in ('\ufffd', '?'):
            if '\ufffd\ufffd' in c0 or '???' in c0:
                continue
        return j, r
    return start + 1, (rows[start + 1] if start + 1 < len(rows) else [])

def w14_is_new_section(cell: str, current_marker: str) -> bool:
    cell = cell.strip()
    if not cell: return False
    if cell.startswith(current_marker): return False
    return bool(_SECTION_HEADER_RE.match(cell))

def w14_iter_section_body(rows, marker: str):
    start = w14_find_section_start(rows, marker)
    if start < 0: return
    hdr_i, _ = w14_find_header_row(rows, start)
    for j in range(hdr_i + 1, len(rows)):
        rr = rows[j]
        if not rr or not rr[0].strip(): continue
        c0 = rr[0].strip()
        if c0.startswith('📌'): continue
        first_char = c0.lstrip('"').lstrip("'")[:1]
        if first_char in ('\ufffd', '?'):
            if '\ufffd\ufffd' in c0 or '???' in c0:
                continue
        if w14_is_new_section(c0, marker): return
        yield rr

def w14_parse_section_weekly_col(rows, marker: str, week_label: str) -> dict:
    start = w14_find_section_start(rows, marker)
    if start < 0: return {}
    _, hdr = w14_find_header_row(rows, start)
    try: col = hdr.index(week_label)
    except ValueError: return {}
    out = {}
    for rr in w14_iter_section_body(rows, marker):
        name = rr[0].strip()
        if name == '合计': continue
        try: out[name] = float(rr[col].replace(',', '')) if col < len(rr) else 0
        except ValueError: out[name] = 0
    return out

def w14_parse_section_weekly_all(rows, marker: str):
    start = w14_find_section_start(rows, marker)
    if start < 0: return [], {}
    _, hdr = w14_find_header_row(rows, start)
    try: tot_col = hdr.index('合计')
    except ValueError: tot_col = len(hdr)
    week_cols = list(range(1, tot_col))
    week_labels = [hdr[c].strip() for c in week_cols]
    out = {}
    for rr in w14_iter_section_body(rows, marker):
        name = rr[0].strip()
        if name == '合计': continue
        vals = []
        for c in week_cols:
            try: vals.append(float(rr[c].replace(',','')) if c < len(rr) and rr[c].strip() else 0)
            except ValueError: vals.append(0)
        out[name] = vals
    return week_labels, out

def w14_parse_section_records(rows, marker: str) -> list:
    start = w14_find_section_start(rows, marker)
    if start < 0: return []
    _, hdr = w14_find_header_row(rows, start)
    out = []
    for rr in w14_iter_section_body(rows, marker):
        name = rr[0].strip()
        if name == '合计': break
        d = {'name': name}
        for k, h in enumerate(hdr[1:], start=1):
            hh = h.strip()
            if not hh: continue
            v = rr[k].strip().replace(',', '') if k < len(rr) else ''
            try: d[hh] = float(v) if v else 0
            except ValueError: d[hh] = 0
        out.append(d)
    return out

def w14_list_section_names(rows, marker: str) -> list:
    return [rr[0].strip() for rr in w14_iter_section_body(rows, marker) if rr[0].strip() != '合计']

def w14_auto_detect_current_week(rows, marker: str) -> str:
    labels, data = w14_parse_section_weekly_all(rows, marker)
    if not labels: return ''
    last_with_data = -1
    for i in range(len(labels)):
        if any((vals[i] if i<len(vals) else 0) != 0 for vals in data.values()):
            last_with_data = i
    return labels[last_with_data] if last_with_data >= 0 else labels[-1]


# ============================================================================
# PART 6: sowhat_generators.py — So What text generation
# ============================================================================
def _sow_to_m(x): return num(x) / 1_000_000
def _sow_to_w(x): return num(x) / 10_000

def gen_s1_targets(ctx):
    S1 = ctx['S1']
    a_full = S1['A'][S1['A']['指标'] == '2026全业务目标'].iloc[0]
    target = _sow_to_m(a_full['目标APE'])
    issued = _sow_to_m(a_full['已达成APE'])
    rate = num(str(a_full['目标达成率']).replace('%', ''))
    gap = target - issued
    cC = S1['C']
    mar = cC[cC.iloc[:, 0] == '2026-03']; feb = cC[cC.iloc[:, 0] == '2026-02']
    mar_cnt = int(num(mar.iloc[0]['件数'])) if len(mar) else 0
    feb_cnt = int(num(feb.iloc[0]['件数'])) if len(feb) else 0
    pct = (mar_cnt - feb_cnt) / feb_cnt * 100 if feb_cnt else 0
    return (f"2026 全业务批核 APE 达 {issued:.1f}M，达成率 {rate:.1f}%，"
            f"全年目标 {target:.0f}M 缺口约 {gap:.0f}M。"
            f"3 月预约件数 {mar_cnt} 件，环比 {pct:+.1f}%，前端动能延续，Q2 转化批核可期。")

def gen_s1_pipeline(ctx):
    S1 = ctx['S1']
    g = S1['G']
    total_row = g[g.iloc[:, 0] == '合计'].iloc[0]
    issued = _sow_to_m(total_row['2026批核APE'])
    unbat = _sow_to_m(total_row['未批核APE'])
    pend = _sow_to_m(total_row['待签APE'])
    unbat_cnt = int(num(total_row['未批核件数']))
    pend_cnt = int(num(total_row['待签件数']))
    pipe = unbat + pend
    pct = pipe / issued * 100 if issued else 0
    return (f"未批核（{unbat:.1f}M，{unbat_cnt} 件）与待签（{pend:.1f}M，{pend_cnt} 件）"
            f"合计 {pipe:.1f}M 在管道中，占批核 APE 的 {pct:.0f}%。"
            f"若能快速推进至生效，可直接拉升达成率逾 {pipe/1113*100:.0f} 个百分点。")

def gen_s1_channel(ctx):
    S1 = ctx['S1']
    g = S1['G']
    rows = [r for _, r in g.iterrows() if r.iloc[0] != '合计' and str(r.iloc[0]).strip() and _sow_to_m(r['2026批核APE']) > 0]
    rows.sort(key=lambda r: _sow_to_m(r['2026批核APE']), reverse=True)
    top = rows[0]
    top_name = top.iloc[0]; top_iss = _sow_to_m(top['2026批核APE'])
    top_pct = top_iss / _sow_to_m(g[g.iloc[:, 0] == '合计'].iloc[0]['2026批核APE']) * 100
    top_un = _sow_to_m(top['未批核APE'])
    parts = [f"{r.iloc[0]}（批核 {_sow_to_m(r['2026批核APE']):.1f}M）" for r in rows[1:]]
    others = '、'.join(parts) if parts else '—'
    return (f"{top_name}批核 {top_iss:.1f}M 占全渠道 {top_pct:.1f}%，是唯一业绩主力；"
            f"未批核 {top_un:.1f}M 是最大转化机会。其余渠道：{others}，体量较小，需明确各渠道扩量优先级。")

def gen_s1_monthly(ctx):
    S1 = ctx['S1']
    cE, cC, cD = S1['E'], S1['C'], S1['D']
    months_26 = [r for _, r in cE.iterrows() if str(r.iloc[0]).startswith('2026')]
    peak = max(months_26, key=lambda r: _sow_to_m(r['APE']))
    peak_m = int(peak.iloc[0].split('-')[1]); peak_v = _sow_to_m(peak['APE'])
    last = months_26[-1]; last_m = int(last.iloc[0].split('-')[1]); last_apr_v = _sow_to_m(last['APE'])
    return (f"2026 年 Q1 呈现存量驱动特征，{peak_m}月批核峰值 {peak_v:.1f}M 为全年迄今最高。"
            f"3 月预约/签单同步走强，前端动能复苏。{last_m}月初已批核 {last_apr_v:.1f}M，节后启动正常，"
            f"后续 9 个月需月均 83.8M 以达成年度目标。")

def gen_s2_monthly(ctx):
    S1 = ctx['S1']; cC = S1['C']
    mar = cC[cC.iloc[:, 0] == '2026-03']; feb = cC[cC.iloc[:, 0] == '2026-02']; jan = cC[cC.iloc[:, 0] == '2026-01']
    mar_cnt = int(num(mar.iloc[0]['件数'])) if len(mar) else 0
    feb_cnt = int(num(feb.iloc[0]['件数'])) if len(feb) else 0
    jan_cnt = int(num(jan.iloc[0]['件数'])) if len(jan) else 0
    mar_ape = _sow_to_m(mar.iloc[0]['APE']) if len(mar) else 0
    feb_ape = _sow_to_m(feb.iloc[0]['APE']) if len(feb) else 0
    mar_avg = mar_ape * 10 / mar_cnt if mar_cnt else 0
    feb_avg = feb_ape * 10 / feb_cnt if feb_cnt else 0
    pct = (mar_cnt - feb_cnt) / feb_cnt * 100 if feb_cnt else 0
    return (f"2026 年 3 月预约件数从 2 月 {feb_cnt} 件增至 {mar_cnt} 件"
            f"（环比 {pct:+.1f}%），前端明显回暖。但 3 月件均 APE {mar_avg:.1f} 万低于 2 月 {feb_avg:.1f} 万，"
            f"显示新增单偏中小额，需关注单均下滑趋势。")

def gen_s2_forecast(ctx):
    S1 = ctx['S1']
    a_full = S1['A'][S1['A']['指标'] == '2026全业务目标'].iloc[0]
    target = _sow_to_m(a_full['目标APE']); issued = _sow_to_m(a_full['已达成APE'])
    rate = num(str(a_full['目标达成率']).replace('%', '')); gap = target - issued
    remaining = 9; monthly = gap / remaining
    return (f"截至 {ctx['current_week']}，已批核 {issued:.1f}M（达成率 {rate:.1f}%），"
            f"完成全年目标 {target:.0f}M 还需约 {gap:.0f}M，剩余 {remaining} 个月月均需 {monthly:.1f}M。"
            f"若 4–12 月维持 80–100M 节奏，全年可达标。")

def gen_s3_products(ctx):
    S4 = ctx['S4']
    df = S4['C'].copy()
    df = df[df['保司'] == '永明']
    df['_ape'] = df['APE'].apply(num)
    df = df.sort_values('_ape', ascending=False)
    top = df.head(10)
    top_ape_m = top['_ape'].sum() / 1e6; top_cnt = int(top['件数'].apply(num).sum())
    top1 = top.iloc[0]; top1_name = top1['产品名称']
    top1_ape = top1['_ape'] / 1e6; top1_cnt = int(num(top1['件数']))
    top1_avg = top1['_ape'] / top1_cnt / 1e4 if top1_cnt else 0
    return (f"永明 TOP10 产品累计签单 {top_ape_m:.1f}M（{top_cnt} 件），集中度高。"
            f"头部产品「{top1_name}」贡献 {top1_ape:.1f}M/{top1_cnt} 件，件均 {top1_avg:.1f} 万。"
            f"中小额产品虽件数占比大，但单均偏低，需关注大单产品的持续性供给。")

def gen_s3_license(ctx):
    S1 = ctx['S1']; H = S1['H']
    dwbank = H[H['牌照'] == 'DW Bank']; jf = H[H['牌照'] == 'JF']
    if len(dwbank) and len(jf):
        dw_row = dwbank.iloc[0]; jf_row = jf.iloc[0]
        dw_vals = {m: _sow_to_m(dw_row[m]) for m in ['2026-01', '2026-02', '2026-03', '2026-04']}
        jf_vals = {m: _sow_to_m(jf_row[m]) for m in ['2026-01', '2026-02', '2026-03', '2026-04']}
        dw_peak = max(dw_vals.items(), key=lambda x: x[1])
        jf_peak = max(jf_vals.items(), key=lambda x: x[1])
        sub_row = H[H['牌照'] == 'Sub Total'].iloc[0]
        unbat_total = _sow_to_m(sub_row['未批核']) + _sow_to_m(dw_row['未批核'])
        return (f"DW Bank {dw_peak[0][-2:].lstrip('0')}月批核 {dw_peak[1]:.1f}M 为峰值，BK 大额单集中处理。"
                f"JF {jf_peak[0][-2:].lstrip('0')}月批核 {jf_peak[1]:.1f}M 领跑经代牌照。"
                f"合计未批核 {unbat_total:.1f}M 是短期转化目标。")
    return "牌照数据待更新。"

def gen_s4_bubble(ctx):
    S2 = ctx['S2']; a = S2['A']
    rows = a[a['业务细分'].isin(['永明经代', '天领业务', 'BK业务', '合伙转介业务', '成事家办', '同行经代', 'ICLUB业务', 'IFA业务'])]
    rates = []
    for _, r in rows.iterrows():
        tgt = num(r['目标APE']); iss = num(r['2026批核APE'])
        if tgt > 0:
            rates.append((r['业务细分'], iss / tgt * 100, _sow_to_m(r['2026批核APE'])))
    rates.sort(key=lambda x: x[1], reverse=True)
    top, bottom = rates[0], rates[-1]
    sorted_by_scale = sorted(rates, key=lambda x: x[2], reverse=True)
    rate_values = [r[1] for r in rates]
    median = sorted(rate_values)[len(rate_values) // 2]
    risk = None
    for r in sorted_by_scale:
        if r[1] < median:
            risk = r
            break
    if risk is None:
        risk = sorted_by_scale[-1]
    return (f"{top[0]}达成率 {top[1]:.1f}% 领跑全渠道，是唯一接近达标的渠道。"
            f"{risk[0]}规模较大（批核 {risk[2]:.1f}M）但达成率仅 {risk[1]:.1f}%，是最大战略风险点。"
            f"{bottom[0]}达成率仅 {bottom[1]:.1f}%，需重点关注。")

def gen_s4_waterfall(ctx):
    S2 = ctx['S2']; a = S2['A']
    total = a[a['业务细分'] == '合计'].iloc[0]
    target = _sow_to_m(total['目标APE']); issued = _sow_to_m(total['2026批核APE'])
    unbat = _sow_to_m(total['未批核APE']); pend = _sow_to_m(total['待签APE'])
    pipeline = issued + unbat + pend; remain = target - pipeline
    rows = [r for _, r in a.iterrows() if r['业务细分'] != '合计']
    top = max(rows, key=lambda r: num(r['2026批核APE']))
    top_name = top['业务细分']; top_iss = _sow_to_m(top['2026批核APE'])
    return (f"{top_name}贡献最大批核（{top_iss:.1f}M）。已批核 + 在途合计 {pipeline:.1f}M，"
            f"距目标 {target:.0f}M 仍差 {remain:.0f}M。若未批核 {unbat:.1f}M 与待签 {pend:.1f}M 快速推进，可直接拉升年度达成。")

def gen_s4_smallmult(ctx):
    S2 = ctx['S2']; cD = S2['D-APE']; cE = S2['E-APE']
    mar_signs = [(r.iloc[0], _sow_to_m(r['2026-03'])) for _, r in cD.iterrows() if r.iloc[0] != '合计']
    mar_signs.sort(key=lambda x: x[1], reverse=True)
    feb_apps = [(r.iloc[0], _sow_to_m(r['2026-02'])) for _, r in cE.iterrows() if r.iloc[0] != '合计']
    feb_apps.sort(key=lambda x: x[1], reverse=True)
    top_sign, top_app = mar_signs[0], feb_apps[0]
    return (f"{top_sign[0]}：3 月签单 {top_sign[1]:.1f}M 居首，前端动能最强。"
            f"{top_app[0]}：2 月批核峰 {top_app[1]:.1f}M 为存量集中放款。各业务线节奏分化明显，需按渠道特征差异化管理。")

def gen_s5_target(ctx):
    S2 = ctx['S2']; a = S2['A']
    rows = [r for _, r in a.iterrows() if r['业务细分'] != '合计']
    sorted_rows = sorted(rows, key=lambda r: num(r['目标APE']) - num(r['2026批核APE']), reverse=True)
    biggest_gap = sorted_rows[0]
    achievers = [(r['业务细分'], num(r['2026批核APE']) / num(r['目标APE']) * 100 if num(r['目标APE']) else 0,
                  _sow_to_m(r['2026批核APE']), _sow_to_m(r['目标APE'])) for r in rows if num(r['目标APE']) > 0]
    achievers.sort(key=lambda x: x[1], reverse=True); best = achievers[0]
    gap_m = _sow_to_m(biggest_gap['目标APE']) - _sow_to_m(biggest_gap['2026批核APE'])
    return (f"{best[0]}批核 {best[2]:.1f}M 已达目标 {best[3]:.0f}M 的 {best[1]:.1f}%，是唯一接近达标的业务线。"
            f"{biggest_gap['业务细分']}目标最大（{_sow_to_m(biggest_gap['目标APE']):.0f}M）但实际仅 {_sow_to_m(biggest_gap['2026批核APE']):.1f}M，"
            f"缺口 {gap_m:.0f}M 为最大绝对值，是最大战略风险点。")

def gen_s5_pipeline_total(ctx):
    S2 = ctx['S2']; a = S2['A']
    t = a[a['业务细分'] == '合计'].iloc[0]
    issued = _sow_to_m(t['2026批核APE']); unbat = _sow_to_m(t['未批核APE']); pend = _sow_to_m(t['待签APE'])
    total_pipe = issued + unbat + pend; pct_issued = issued / total_pipe * 100
    return (f"管道总值 {total_pipe:.1f}M 中，批核占 {pct_issued:.1f}%（{issued:.1f}M）。"
            f"剩余 {unbat+pend:.1f}M（未批 {unbat:.1f}M + 待签 {pend:.1f}M）若转化可直接推高达成率 {(unbat+pend)/1113*100:.0f}+。")

def gen_s5_donut(ctx):
    S2 = ctx['S2']; a = S2['A']
    rows = [r for _, r in a.iterrows() if r['业务细分'] != '合计']
    worst = None; worst_pct = 0
    for r in rows:
        total_self = num(r['2026批核APE']) + num(r['未批核APE']) + num(r['待签APE'])
        if total_self <= 0: continue
        pct = num(r['未批核APE']) / total_self * 100
        if pct > worst_pct:
            worst = r; worst_pct = pct
    worst_name = worst['业务细分'] if worst is not None else '—'
    worst_val = _sow_to_m(worst['未批核APE']) if worst is not None else 0
    return (f"{worst_name}未批核占比最高（{worst_pct:.0f}%），金额达 {worst_val:.1f}M，催核优先级最高。"
            f"其余业务线管道分布相对均衡，需按未批核规模排序推进。")

def _sow_weekly_totals(ctx, section):
    _, data = w14_parse_section_weekly_all(ctx['s3_rows'], section)
    if '合计' in data:
        return data['合计']
    weeks = max((len(v) for v in data.values()), default=14)
    tot = [0.0] * weeks
    for vals in data.values():
        for i, v in enumerate(vals):
            if i < weeks:
                tot[i] += v
    return tot

def gen_s6_weekly_macro(ctx):
    appt = _sow_weekly_totals(ctx, 'B-APE'); sign = _sow_weekly_totals(ctx, 'C-APE'); app = _sow_weekly_totals(ctx, 'D-APE')
    total_appt = sum(appt) / 1e6; total_sign = sum(sign) / 1e6; total_app = sum(app) / 1e6
    appt_with_week = [(i+1, v) for i, v in enumerate(appt) if v > 0]
    low_week = min(appt_with_week, key=lambda x: x[1]) if appt_with_week else (7, 0)
    return (f"批核 APE 累计 {total_app:.1f}M 显著高于预约 {total_appt:.1f}M，反映跨周积压存量消化效应。"
            f"签单累计 {total_sign:.1f}M，递交→批核转化效率较高。"
            f"W{low_week[0]:02d} 预约低点 {low_week[1]/1e6:.1f}M 需关注执行节奏。")

def gen_s6_weekly_events(ctx):
    appt = _sow_weekly_totals(ctx, 'B-APE'); sign = _sow_weekly_totals(ctx, 'C-APE'); app = _sow_weekly_totals(ctx, 'D-APE')
    sign_peak_i = max(range(len(sign)), key=lambda i: sign[i]) if sign else 0
    app_peak_i = max(range(len(app)), key=lambda i: app[i]) if app else 0
    appt_peak_i = max(range(len(appt)), key=lambda i: appt[i]) if appt else 0
    return (f"W{sign_peak_i+1:02d} 签单峰 {sign[sign_peak_i]/1e6:.1f}M 为阶段高点；"
            f"W{app_peak_i+1:02d} 批核峰 {app[app_peak_i]/1e6:.1f}M 为存量集中消化；"
            f"W{appt_peak_i+1:02d} 预约峰 {appt[appt_peak_i]/1e6:.1f}M 显示前端动能。周度节奏呈脉冲式特征，需关注执行连续性。")

def gen_s6_weekly_brief(ctx):
    cur_week = ctx['current_week']; wk_idx = int(cur_week[-2:]) - 1
    def w(section):
        _, d = w14_parse_section_weekly_all(ctx['s3_rows'], section)
        row = d.get('合计')
        if not row:
            weeks_n = max((len(v) for v in d.values()), default=14)
            row = [sum(v[i] if i<len(v) else 0 for v in d.values()) for i in range(weeks_n)]
        return row[wk_idx] if wk_idx < len(row) else 0
    def wn(section):
        _, d = w14_parse_section_weekly_all(ctx['s3_rows'], section)
        row = d.get('合计')
        if not row:
            weeks_n = max((len(v) for v in d.values()), default=14)
            row = [sum(v[i] if i<len(v) else 0 for v in d.values()) for i in range(weeks_n)]
        return int(row[wk_idx]) if wk_idx < len(row) else 0
    appt_m = w('B-APE')/1e6; appt_c = wn('B-件数')
    sign_m = w('C-APE')/1e6; sign_c = wn('C-件数')
    app_m = w('D-APE')/1e6; app_c = wn('D-件数')
    S1 = ctx['S1']; g = S1['G']
    total_row = g[g.iloc[:, 0] == '合计'].iloc[0]
    unbat_m = _sow_to_m(total_row['未批核APE']); unbat_c = int(num(total_row['未批核件数']))
    return (f"{cur_week} 预约 {appt_m:.1f}M/{appt_c}件；签单 {sign_m:.1f}M/{sign_c}件；批核 {app_m:.1f}M/{app_c}件。"
            f"未批核 {unbat_m:.1f}M（{unbat_c}件）是 Q2 批核的核心转化储量。")

def gen_s6_ka_top10(ctx):
    S2 = ctx['S2']
    if 'I' not in S2:
        return "KA TOP10 数据待更新。"
    df = S2['I'].copy()
    df = df[df.iloc[:, 1] != '合计']
    df['_ape'] = df['2026批核APE'].apply(num)
    df = df.sort_values('_ape', ascending=False)
    a_total = S2['A'][S2['A']['业务细分'] == '合计']
    grand_m = _sow_to_m(a_total.iloc[0]['2026批核APE']) if len(a_total) else df['_ape'].sum() / 1e6
    top10 = df.head(10); top10_m = top10['_ape'].sum() / 1e6
    pct = top10_m / grand_m * 100 if grand_m > 0 else 0
    top1 = top10.iloc[0]
    top1_name = top1['KEY ACCOUNT'] if 'KEY ACCOUNT' in top10.columns else top1.iloc[1]
    top1_m = top1['_ape'] / 1e6; top1_pct = top1_m / grand_m * 100 if grand_m > 0 else 0
    return (f"TOP10 KA 累计批核 {top10_m:.1f}M（占全业务 {pct:.1f}%），高度集中。"
            f"{top1_name} 以 {top1_m:.1f}M 独占 {top1_pct:.1f}%，是最核心的单一客户，亦是最大集中风险点。")

def _sow_heat_analysis(ctx, ape_key, cnt_key, label):
    _, ape = w14_parse_section_weekly_all(ctx['s3_rows'], ape_key)
    _, cnt = w14_parse_section_weekly_all(ctx['s3_rows'], cnt_key)
    order = ['天领业务', '成事家办', 'BK业务', '同行经代', '永明经代', '合伙转介业务', 'ICLUB业务']
    totals = {n: (sum(ape.get(n, [])), sum(cnt.get(n, []))) for n in order}
    grand = sum(v for v, _ in totals.values()); grand_c = sum(c for _, c in totals.values())
    if grand == 0:
        return f"{label}数据暂缺。"
    top = max(totals.items(), key=lambda x: x[1][0])
    top_pct = top[1][0] / grand * 100
    n_weeks = max((len(v) for v in ape.values()), default=14)
    wt = [sum((ape.get(n, [0]*n_weeks)[i] if i < len(ape.get(n, [])) else 0) for n in order) for i in range(n_weeks)]
    peak_i = wt.index(max(wt)) if wt else 0
    peak_c = sum((cnt.get(n, [0]*n_weeks)[peak_i] if peak_i < len(cnt.get(n, [])) else 0) for n in order)
    return (f"{label}总量 {grand/1e6:.1f}M/{int(grand_c)}件。{top[0]}领跑，累计 {top[1][0]/1e6:.1f}M/{int(top[1][1])}件，"
            f"占比 {top_pct:.0f}%。W{peak_i+1:02d} 单周峰值 {max(wt)/1e6:.1f}M/{int(peak_c)}件，是本季度最高单周贡献。")

def gen_s7_appt(ctx): return _sow_heat_analysis(ctx, 'B-APE', 'B-件数', '预约')
def gen_s7_sign(ctx): return _sow_heat_analysis(ctx, 'C-APE', 'C-件数', '签单')
def gen_s7_app(ctx): return _sow_heat_analysis(ctx, 'D-APE', 'D-件数', '批核')

def gen_s7_lifecycle(ctx):
    S3 = ctx['S3']
    if 'G' not in S3:
        return "签批时效数据待更新。"
    g = S3['G']
    rows = []
    for _, r in g.iterrows():
        name = str(r.get('业务细分', '')).strip()
        cnt = num(r.get('件数', 0))
        if name and name != '合计' and cnt > 0:
            rows.append({'name': name, 'cnt': int(cnt),
                         'avg_wan': round(num(r.get('件均APE', 0)) / 10000),
                         'avg_tat': round(num(r.get('平均时效(天)', 0)), 1),
                         'p90': round(num(r.get('P90时效(天)', 0))),
                         'median': round(num(r.get('中位时效(天)', 0))),
                         'max_tat': round(num(r.get('最大时效(天)', 0)))})
    if not rows:
        return "签批时效数据待更新。"
    worst = max(rows, key=lambda x: x['avg_tat'])
    fastest = min(rows, key=lambda x: x['avg_tat'])
    tot = g[g['业务细分'].str.strip() == '合计']
    tot_avg = round(num(tot.iloc[0].get('平均时效(天)', 0))) if len(tot) else 0
    tot_p90 = round(num(tot.iloc[0].get('P90时效(天)', 0))) if len(tot) else 0
    sla_over = [r for r in rows if r['avg_tat'] > 60]
    sla_text = (f"{'、'.join(r['name'] for r in sla_over)} 平均时效超 60 天 SLA" if sla_over else "所有业务线平均时效均在 60 天 SLA 内")
    return (f"{worst['name']}件均 APE {worst['avg_wan']}万但平均时效 {worst['avg_tat']} 天（P90 {worst['p90']} 天），"
            f"需重点排查审批瓶颈。{fastest['name']}时效最短（{fastest['avg_tat']}天），可复制其流程优化经验。"
            f"整体平均 {tot_avg} 天（P90 {tot_p90} 天）。{sla_text}。")

def gen_s7_kpi(ctx):
    S3 = ctx['S3']
    if 'H' not in S3:
        return "KPI 达成数据待更新。"
    h = S3['H']
    rows = []
    for _, r in h.iterrows():
        kpi = str(r.get('指标', '')).strip()
        if not kpi or kpi == '合计': continue
        rows.append({'name': kpi, 'target': num(r.get('目标', 0)),
                     'actual': num(r.get('实际', 0)),
                     'rate': num(str(r.get('达成率', 0)).replace('%', ''))})
    if not rows:
        return "KPI 达成数据待更新。"
    worst = min(rows, key=lambda x: x['rate'])
    best = max(rows, key=lambda x: x['rate'])
    return (f"{worst['name']}达成率 {worst['rate']:.0f}%（目标 {worst['target']}，实际 {worst['actual']}），"
            f"为最低分指标，需专项提升。{best['name']}达成率 {best['rate']:.0f}%（目标 {best['target']}，实际 {best['actual']}），"
            f"表现最佳。其余指标达成率分布需逐个评估优先级。")

def gen_s9_peer(ctx):
    S3 = ctx['S3']; j = S3['J-APE']
    months = [c for c in j.columns if c.startswith('2026-')]
    peers = [r.iloc[0] for _, r in j.iterrows() if str(r.iloc[0]).strip() and str(r.iloc[0]) != '合计']
    totals = {}
    for p in peers:
        row = j[j.iloc[:, 0] == p]
        if len(row):
            totals[p] = sum(num(row.iloc[0][m]) for m in months) / 1e6
    if not totals:
        return "同行数据待更新。"
    top_p = max(totals, key=totals.get); top_v = totals[top_p]
    return (f"{top_p}以 {top_v:.1f}M 居同行预约第一，是竞争标杆。"
            f"各同行表现分化明显，需按竞品策略调整渠道优先级与价格敏感度。")

def gen_s10_bk(ctx):
    S2 = ctx['S2']; a = S2['A']
    bk = a[a['业务细分'] == 'BK业务'].iloc[0] if len(a[a['业务细分'] == 'BK业务']) else None
    if bk is None:
        return "银行渠道数据待更新。"
    target = _sow_to_m(bk['目标APE']); issued = _sow_to_m(bk['2026批核APE'])
    rate = issued / target * 100 if target > 0 else 0
    unbat = _sow_to_m(bk['未批核APE']); pend = _sow_to_m(bk['待签APE'])
    return (f"银行渠道批核 {issued:.1f}M（目标 {target:.0f}M，达成率 {rate:.0f}%）。"
            f"未批核 {unbat:.1f}M + 待签 {pend:.1f}M 是短期转化重点。"
            f"民生/平安两大银行贡献均衡，需持续深化合作。")

def gen_s11_bk(ctx):
    S3 = ctx['S3']; m = S3['M-APE']
    weeks = [c for c in m.columns if c.startswith('2026W')]
    vals = []
    for ww in weeks:
        row = m[m.iloc[:, 0] == '合计']
        vals.append(num(row.iloc[0][ww]) if len(row) else 0)
    if not vals:
        return "银行周度数据待更新。"
    peak_i = vals.index(max(vals))
    return (f"W{peak_i+1:02d} 银行预约峰值 {max(vals)/1e6:.1f}M，为阶段高点。"
            f"周度波动较大，需关注银行端执行节奏。")

SOWHAT_GENERATORS = {
    's1_targets': gen_s1_targets, 's1_pipeline': gen_s1_pipeline,
    's1_channel': gen_s1_channel, 's1_monthly': gen_s1_monthly,
    's2_monthly': gen_s2_monthly, 's2_forecast': gen_s2_forecast,
    's3_products': gen_s3_products, 's3_license': gen_s3_license,
    's4_bubble': gen_s4_bubble, 's4_waterfall': gen_s4_waterfall,
    's4_smallmult': gen_s4_smallmult, 's5_target': gen_s5_target,
    's5_pipeline_total': gen_s5_pipeline_total, 's5_donut': gen_s5_donut,
    's6_weekly_macro': gen_s6_weekly_macro, 's6_weekly_events': gen_s6_weekly_events,
    's6_weekly_brief': gen_s6_weekly_brief, 's6_ka_top10': gen_s6_ka_top10,
    's7_appt': gen_s7_appt, 's7_sign': gen_s7_sign, 's7_app': gen_s7_app,
    's7_lifecycle': gen_s7_lifecycle, 's7_kpi': gen_s7_kpi,
    's9_peer': gen_s9_peer, 's10_bk': gen_s10_bk, 's11_bk': gen_s11_bk,
}


# ============================================================================
# PART 7: update_ppt.py — Main PPT update logic
# ============================================================================
CSV_S1 = HERE / "S1-总览仪表盘.csv"
CSV_S2 = HERE / "S2-业务端视角.csv"
CSV_S3 = HERE / "S3-执行管理端.csv"
CSV_S4 = HERE / "S4-产品端视角.csv"

import csv as _csv_init, re as _re_init

def _derive_current_week():
    with open(CSV_S3, encoding="utf-8-sig") as _f:
        for _row in _csv_init.reader(_f):
            _wks = [c for c in _row if _re_init.match(r"2026W\d+", c)]
            if _wks:
                _latest = _wks[-1]
                return _latest[4:]
    return "W18"

CURRENT_WEEK = _derive_current_week()

def _find_shape(slide, name_or_predicate):
    if callable(name_or_predicate):
        for s in slide.shapes:
            if name_or_predicate(s):
                return s
    else:
        for s in slide.shapes:
            if s.name == name_or_predicate:
                return s
    return None

def _cd_to_spec(chart_data):
    cats = [c.label for c in chart_data.categories]
    names, vals = [], []
    for s in chart_data._series:
        names.append(s.name)
        vals.append([None if v is None else float(v) for v in s.values])
    return cats, names, vals

def _replace_chart(slide, shape_name, chart_data):
    if chart_data is None:
        print(f"  ! chart data unavailable, skipping: {shape_name}")
        return False
    shp = _find_shape(slide, shape_name)
    if shp is None or not shp.has_chart:
        print(f"  ! chart not found: {shape_name}")
        return False
    cats, names, vals = _cd_to_spec(chart_data)
    patch_chart(shp.chart.part, vals, categories=cats, series_names=names)
    return True

def _replace_chart_dlbls(slide, shape_name, series_index, labels):
    shp = _find_shape(slide, shape_name)
    if shp is None or not shp.has_chart:
        return False
    n = patch_chart_dlbls(shp.chart.part, series_index, labels)
    return n > 0

def _patch_s1_monthly_trend(slide, chart_name):
    shp = _find_shape(slide, chart_name)
    if shp is None or not shp.has_chart:
        print(f"  ! [D-s1 patch] chart not found: {chart_name}")
        return
    NS_C = "http://schemas.openxmlformats.org/drawingml/2006/chart"
    root = shp.chart.part._element

    def _read_s1_block_ape(block_letter):
        result = {}
        in_block = False; hdr = None
        import csv as _csv2, re as _re
        with open(CSV_S1, encoding="utf-8-sig") as _f:
            for _row in _csv2.reader(_f):
                if not _row: continue
                first = _row[0].strip()
                if first.startswith(f"{block_letter}.") and "月度业绩" in first:
                    in_block = True; continue
                if not in_block: continue
                if first.startswith("📌"): continue
                if in_block and _re.match(r"^[A-Z]\.", first) and not first.startswith(f"{block_letter}."):
                    break
                if hdr is None and any("年月" in c for c in _row):
                    hdr = [c.strip() for c in _row]; continue
                if not first: break
                if hdr is None: continue
                if _re.match(r"^\d{4}-\d{2}$", first):
                    try:
                        ape_idx = hdr.index("APE") if "APE" in hdr else 2
                        result[first] = float(_row[ape_idx].replace(",","").strip()) / 1e6
                    except (ValueError, IndexError):
                        pass
        return result

    c_ape = _read_s1_block_ape("C")
    d_ape = _read_s1_block_ape("D")
    e_ape = _read_s1_block_ape("E")

    from lxml import etree as _etree_lxml_m

    all_months = []
    for ser in root.findall(f".//{{{NS_C}}}ser")[:1]:
        for sc in ser.findall(f"{{{NS_C}}}cat//{{{NS_C}}}strCache"):
            for pt in sc.findall(f"{{{NS_C}}}pt"):
                v = pt.find(f"{{{NS_C}}}v")
                if v is not None and v.text:
                    all_months.append(v.text)

    MONTH_ABBR = {"Jan":"01","Feb":"02","Mar":"03","Apr":"04","May":"05","Jun":"06",
                  "Jul":"07","Aug":"08","Sep":"09","Oct":"10","Nov":"11","Dec":"12"}
    def _label_to_key(label):
        import re as _re
        m = _re.match(r"^(\d{2})-([A-Za-z]{3})$", label)
        if m:
            yr = "20" + m.group(1)
            mo = MONTH_ABBR.get(m.group(2).capitalize(), "00")
            return f"{yr}-{mo}"
        return None

    series_by_name = {
        "预约 APE(M)": c_ape,   "预约": c_ape,   "Appointment": c_ape,
        "签单 APE(M)": d_ape,   "签单": d_ape,   "Signed": d_ape,
        "批核 APE(M)": e_ape,   "批核": e_ape,   "Issued": e_ape,
        "实际批核": e_ape,       "批核APE": e_ape,
    }
    series_by_idx = [c_ape, d_ape, e_ape]

    patched_series = 0
    for ser_idx, ser in enumerate(root.findall(f".//{{{NS_C}}}ser")):
        tx = ser.find(f".//{{{NS_C}}}tx")
        ser_name = ""
        if tx is not None:
            v_el = tx.find(f".//{{{NS_C}}}v")
            if v_el is not None: ser_name = v_el.text or ""

        data_dict = None
        for sn, dd in series_by_name.items():
            if sn in ser_name or ser_name in sn:
                data_dict = dd; break
        if data_dict is None and ser_idx < len(series_by_idx):
            data_dict = series_by_idx[ser_idx]

        if data_dict is None: continue

        new_2026_months = sorted(
            k for k in data_dict
            if k.startswith("2026") and _label_to_key(
                f"{k[2:4]}-{['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'][int(k[5:7])-1]}"
            ) not in set(all_months)
        )
        MONTH_NUM_TO_ABBR = {1:"Jan",2:"Feb",3:"Mar",4:"Apr",5:"May",6:"Jun",
                              7:"Jul",8:"Aug",9:"Sep",10:"Oct",11:"Nov",12:"Dec"}
        new_labels = []
        for k in new_2026_months:
            mo_num = int(k[5:7])
            label = f"{k[2:4]}-{MONTH_NUM_TO_ABBR[mo_num]}"
            if label not in all_months:
                new_labels.append((k, label))

        for nc in ser.findall(f".//{{{NS_C}}}numCache"):
            for pt in nc.findall(f"{{{NS_C}}}pt"):
                idx = int(pt.get("idx", "0"))
                label = all_months[idx] if idx < len(all_months) else None
                if label is None: continue
                key = _label_to_key(label)
                if key and key.startswith("2026") and key in data_dict:
                    v_node = pt.find(f"{{{NS_C}}}v")
                    if v_node is not None:
                        v_node.text = f"{data_dict[key]:.2f}"

            if new_labels:
                ptCount = nc.find(f"{{{NS_C}}}ptCount")
                if ptCount is not None:
                    ptCount.set("val", str(len(all_months) + len(new_labels)))
                for k, label in new_labels:
                    new_idx = len(all_months) + new_labels.index((k, label))
                    new_pt = _etree_lxml_m.SubElement(nc, f"{{{NS_C}}}pt")
                    new_pt.set("idx", str(new_idx))
                    new_v = _etree_lxml_m.SubElement(new_pt, f"{{{NS_C}}}v")
                    new_v.text = f"{data_dict[k]:.2f}"

        if ser_idx == 0 and new_labels:
            for ser2 in root.findall(f".//{{{NS_C}}}ser"):
                for sc in ser2.findall(f"{{{NS_C}}}cat//{{{NS_C}}}strCache"):
                    ptCount = sc.find(f"{{{NS_C}}}ptCount")
                    if ptCount is not None:
                        ptCount.set("val", str(len(all_months) + len(new_labels)))
                    for k, label in new_labels:
                        new_idx = len(all_months) + new_labels.index((k, label))
                        new_pt = _etree_lxml_m.SubElement(sc, f"{{{NS_C}}}pt")
                        new_pt.set("idx", str(new_idx))
                        new_v = _etree_lxml_m.SubElement(new_pt, f"{{{NS_C}}}v")
                        new_v.text = label
            all_months.extend([lbl for _, lbl in new_labels])
            print(f"  [D-s1 patch] Added {len(new_labels)} new month(s): {[lbl for _, lbl in new_labels]}")

        patched_series += 1

    print(f"  [D-s1 patch] {chart_name}: {patched_series} series patched (2026 months only)")

def _fix_chart_dlbls_positions(slide, chart_name):
    shp = _find_shape(slide, chart_name)
    if shp is None or not shp.has_chart:
        print(f"  ! [dLbls fix] chart not found: {chart_name}")
        return
    NS_C = "http://schemas.openxmlformats.org/drawingml/2006/chart"
    root = shp.chart.part._element
    patched = 0
    for ser in root.findall(f".//{{{NS_C}}}ser"):
        dLbls = ser.find(f"{{{NS_C}}}dLbls")
        if dLbls is None:
            continue
        delete_el = dLbls.find(f"{{{NS_C}}}delete")
        if delete_el is not None and delete_el.get("val") == "1":
            continue
        for dLbl in dLbls.findall(f"{{{NS_C}}}dLbl"):
            dLbls.remove(dLbl)
        pos_el = dLbls.find(f"{{{NS_C}}}dLblPos")
        if pos_el is not None:
            pos_el.set("val", "outEnd")
        else:
            from lxml import etree as _etree_dlbl
            pos_el = _etree_dlbl.SubElement(dLbls, f"{{{NS_C}}}dLblPos")
            pos_el.set("val", "outEnd")
        patched += 1
    print(f"  [dLbls fix] {chart_name}: {patched} series patched")

def _update_kpi_text(slide, placeholder_text, new_text):
    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue
        tf = shape.text_frame
        if tf.text.strip() == placeholder_text:
            tf.clear()
            p = tf.paragraphs[0]
            p.text = new_text
            for run in p.runs:
                run.font.size = Pt(24)
            return True
    return False

def _update_kpi_number(slide, placeholder_text, new_value, suffix=""):
    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue
        tf = shape.text_frame
        if tf.text.strip() == placeholder_text:
            tf.clear()
            p = tf.paragraphs[0]
            p.text = f"{new_value:.1f}{suffix}"
            for run in p.runs:
                run.font.size = Pt(36)
            return True
    return False

def _update_textbox(slide, old_text, new_text):
    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue
        tf = shape.text_frame
        if tf.text.strip() == old_text:
            tf.clear()
            p = tf.paragraphs[0]
            p.text = new_text
            return True
    return False

def _clone_w14_column(slide7):
    for shape in slide7.shapes:
        if not shape.has_table:
            continue
        tbl = shape.table
        if len(tbl.columns) >= 14:
            break
        for c_idx in range(len(tbl.columns) - 1):
            cells = [tbl.cell(r, c_idx) for r in range(len(tbl.rows))]
            rightmost = max(c.top + c.height for c in cells)
            if rightmost > Emu(14000000):
                src_col = c_idx
                new_col_idx = len(tbl.columns)
                tbl.add_column()
                for r in range(len(tbl.rows)):
                    src_cell = tbl.cell(r, src_col)
                    dst_cell = tbl.cell(r, new_col_idx)
                    for paragraph in src_cell.text_frame.paragraphs:
                        new_p = dst_cell.text_frame.add_paragraph()
                        new_p.text = paragraph.text
                        for run in paragraph.runs:
                            new_run = new_p.add_run()
                            new_run.text = run.text
                            new_run.font.size = run.font.size
                            new_run.font.bold = run.font.bold
                            new_run.font.name = run.font.name
                break

def run_stage1_update_ppt():
    global S1, S2, S3, S4, prs, slides

    S1, S2, S3, S4 = load_all(CSV_S1, CSV_S2, CSV_S3, CSV_S4)

    _issue_yms = sorted(
        [r for r in S1["E"].iloc[:, 0].tolist()
         if str(r).startswith("2026-")],
    )
    CURRENT_MONTH_COL = _issue_yms[-1] if _issue_yms else "2026-05"
    _cm_parts = CURRENT_MONTH_COL.split("-")
    _cm_month = int(_cm_parts[1])
    PREV_MONTH_COL = f"{_cm_parts[0]}-{_cm_month-1:02d}" if _cm_month > 1 else f"{int(_cm_parts[0])-1}-12"
    CURRENT_WEEK_COL = f"2026{CURRENT_WEEK}"
    _cw_num = int(CURRENT_WEEK[1:])
    PREV_WEEK_COL = f"2026W{_cw_num-1:02d}"

    import calendar as _calendar
    CURRENT_MONTH = _cm_month
    REMAINING_M = 12 - CURRENT_MONTH + 1
    _last_day = _calendar.monthrange(int(_cm_parts[0]), _cm_month)[1]
    REPORT_DATE = f"{_cm_parts[0]}-{_cm_month:02d}-{_last_day:02d}"

    print(f"  Dynamic config: CURRENT_MONTH={CURRENT_MONTH} ({CURRENT_MONTH_COL})")
    print(f"  CURRENT_WEEK={CURRENT_WEEK} | REMAINING_M={REMAINING_M} | REPORT_DATE={REPORT_DATE}")

    def row_by(df, col, val):
        if df is None or len(df) == 0 or col not in df.columns:
            return None
        sub = df[df[col].astype(str) == val]
        return sub.iloc[0] if len(sub) else None

    def _safe(r, key, default=0):
        if r is None:
            return default
        try:
            return r[key]
        except (KeyError, IndexError):
            return default

    def safe_cell(df, lookup_col, lookup_val, value_col, default=0):
        r = row_by(df, lookup_col, lookup_val)
        return _safe(r, value_col, default)

    _a_full = row_by(S1["A"], "指标", "2026全业务目标")
    _a_sun = row_by(S1["A"], "指标", "2026永明业务目标")

    FULL_TARGET_YUAN = num(_safe(_a_full, "目标APE"))
    FULL_ISSUED_YUAN = num(_safe(_a_full, "已达成APE"))
    FULL_RATE = num(_safe(_a_full, "目标达成率"))
    SUN_TARGET_YUAN = num(_safe(_a_sun, "目标APE"))
    SUN_ISSUED_YUAN = num(_safe(_a_sun, "已达成APE"))
    SUN_RATE = num(_safe(_a_sun, "目标达成率"))

    FULL_TARGET_M = FULL_TARGET_YUAN / 1_000_000
    FULL_ISSUED_M = FULL_ISSUED_YUAN / 1_000_000
    FULL_GAP_M = FULL_TARGET_M - FULL_ISSUED_M
    SUN_TARGET_M = SUN_TARGET_YUAN / 1_000_000
    SUN_ISSUED_M = SUN_ISSUED_YUAN / 1_000_000
    SUN_GAP_M = SUN_TARGET_M - SUN_ISSUED_M

    _b = S1["B"]
    _b_issued = row_by(_b, "指标行", "批核(2026)")
    _b_unbat = row_by(_b, "指标行", "未批核(跨年)")
    _b_pend = row_by(_b, "指标行", "待签(跨年)")
    _b_lost = row_by(_b, "指标行", "流失(2026)")

    ISSUED_APE_M = to_m(_safe(_b_issued, "APE"))
    ISSUED_CNT = int(num(_safe(_b_issued, "件数")))
    UNBAT_APE_M = to_m(_safe(_b_unbat, "APE"))
    UNBAT_CNT = int(num(_safe(_b_unbat, "件数")))
    UNBAT_AVG_W = round(num(_safe(_b_unbat, "件均APE")) / 10_000, 1)
    PEND_APE_M = to_m(_safe(_b_pend, "APE"))
    PEND_CNT = int(num(_safe(_b_pend, "件数")))
    PEND_AVG_W = round(num(_safe(_b_pend, "件均APE")) / 10_000, 1)
    LOST_APE_M = to_m(_safe(_b_lost, "APE"))
    LOST_CNT = int(num(_safe(_b_lost, "件数")))

    _pipe_total_b = ISSUED_APE_M + UNBAT_APE_M + PEND_APE_M + LOST_APE_M
    ISSUED_SHARE = ISSUED_APE_M / _pipe_total_b * 100 if _pipe_total_b else 0
    UNBAT_SHARE = UNBAT_APE_M / _pipe_total_b * 100 if _pipe_total_b else 0
    PEND_SHARE = PEND_APE_M / _pipe_total_b * 100 if _pipe_total_b else 0
    LOST_SHARE = LOST_APE_M / _pipe_total_b * 100 if _pipe_total_b else 0

    _pipe_total_k = ISSUED_APE_M + UNBAT_APE_M + PEND_APE_M
    ISSUED_SHARE_K = ISSUED_APE_M / _pipe_total_k * 100 if _pipe_total_k else 0
    UNBAT_SHARE_K = UNBAT_APE_M / _pipe_total_k * 100 if _pipe_total_k else 0
    PEND_SHARE_K = PEND_APE_M / _pipe_total_k * 100 if _pipe_total_k else 0

    def _channel_kpis(S2_A, name):
        r = row_by(S2_A, "业务细分", name)
        if r is None:
            return {"issued_m": 0, "issued_cnt": 0, "unbat_m": 0, "unbat_cnt": 0,
                    "pend_m": 0, "pend_cnt": 0, "rate": 0, "target_m": 0, "total_m": 0}
        return {
            "issued_m": to_m(_safe(r, "2026批核APE")),
            "issued_cnt": int(num(_safe(r, "批核件数"))),
            "unbat_m": to_m(_safe(r, "未批核APE")),
            "unbat_cnt": int(num(_safe(r, "未批核件数"))),
            "pend_m": to_m(_safe(r, "待签APE")),
            "pend_cnt": int(num(_safe(r, "待签件数"))),
            "rate": num(_safe(r, "目标达成率")),
            "target_m": to_m(_safe(r, "2026目标APE")),
            "total_m": to_m(_safe(r, "总APE")),
        }

    CH_BK = _channel_kpis(S2["B"], "BK业务")
    CH_YMJD = _channel_kpis(S2["B"], "永明经代")
    CH_THJD = _channel_kpis(S2["B"], "同行")
    CH_TL = _channel_kpis(S2["B"], "天领业务")
    CH_ICLUB = _channel_kpis(S2["B"], "ICLUB")
    CH_CSJB = _channel_kpis(S2["B"], "成事家办")
    CH_HHZJ = _channel_kpis(S2["B"], "合伙")

    UNBAT_APE_M = CH_BK["unbat_m"] + CH_YMJD["unbat_m"] + CH_THJD["unbat_m"] + \
                  CH_TL["unbat_m"] + CH_ICLUB["unbat_m"] + CH_CSJB["unbat_m"] + CH_HHZJ["unbat_m"]
    UNBAT_CNT = CH_BK["unbat_cnt"] + CH_YMJD["unbat_cnt"] + CH_THJD["unbat_cnt"] + \
                CH_TL["unbat_cnt"] + CH_ICLUB["unbat_cnt"] + CH_CSJB["unbat_cnt"] + CH_HHZJ["unbat_cnt"]

    PEND_APE_M = CH_BK["pend_m"] + CH_YMJD["pend_m"] + CH_THJD["pend_m"] + \
                 CH_TL["pend_m"] + CH_ICLUB["pend_m"] + CH_CSJB["pend_m"] + CH_HHZJ["pend_m"]
    PEND_CNT = CH_BK["pend_cnt"] + CH_YMJD["pend_cnt"] + CH_THJD["pend_cnt"] + \
               CH_TL["pend_cnt"] + CH_ICLUB["pend_cnt"] + CH_CSJB["pend_cnt"] + CH_HHZJ["pend_cnt"]

    PACE_LINE = FULL_GAP_M / REMAINING_M
    MIN_LINE = PACE_LINE * 0.8

    def _avg_w(m_val, cnt):
        return round(m_val * 1000 / cnt, 1) if cnt else 0.0

    prs = Presentation(TEMPLATE_SRC)
    slides = list(prs.slides)

    class _NullSlide:
        shapes = []
        def __getattr__(self, name): return lambda *a, **k: None

    def _s(idx):
        return slides[idx] if idx < len(slides) else _NullSlide()

    def _find_slide(keyword, fallback_idx=None):
        for sl in slides:
            for sh in sl.shapes:
                if not sh.has_text_frame: continue
                txt = "".join(r.text for p in sh.text_frame.paragraphs for r in p.runs)
                if keyword in txt:
                    return sl
        if fallback_idx is not None and fallback_idx < len(slides):
            print(f"  ! slide not found by keyword {repr(keyword)}, using slides[{fallback_idx}]")
            return slides[fallback_idx]
        return _NullSlide()

    def _find_slide_with_chart(chart_name):
        for sl in slides:
            for sh in sl.shapes:
                if sh.name == chart_name and sh.has_chart:
                    return sl
        return _NullSlide()

    def _chart_slide(chart_name, preferred_slide=None):
        if preferred_slide is not None and not isinstance(preferred_slide, _NullSlide):
            for sh in preferred_slide.shapes:
                if sh.name == chart_name and sh.has_chart:
                    return preferred_slide
        found = _find_slide_with_chart(chart_name)
        if isinstance(found, _NullSlide) and preferred_slide is not None:
            return preferred_slide
        return found

    _SL_SUNLIFE = _find_slide("G.  永明业绩汇报数据-2026", fallback_idx=2)
    _SL_FORECAST = _find_slide("达标节奏线", fallback_idx=1)
    _SL_BUBBLE = _find_slide("H  全业务目标缺口分解", fallback_idx=3)
    _SL_BIZ_VIEW = _find_slide("J  保单阶段构成", fallback_idx=4)
    _SL_EXEC = _find_slide("M  全流程转化漏斗", fallback_idx=5)
    _SL_PIPELINE = _find_slide("同行W", fallback_idx=6)
    _SL_BK = _find_slide("BK批核", fallback_idx=9)
    _SL_BK_WK = _find_slide("BK W", fallback_idx=10)

    def _slide_idx(slide):
        for i, sl in enumerate(slides):
            if sl is slide: return i
        return -1

    print(f"  Slide map: 永明汇报=slides[{_slide_idx(_SL_SUNLIFE)}]  "
          f"F路径管控=slides[{_slide_idx(_SL_FORECAST)}]  "
          f"G/H气泡瀑布=slides[{_slide_idx(_SL_BUBBLE)}]  "
          f"业务视角=slides[{_slide_idx(_SL_BIZ_VIEW)}]  "
          f"执行管理=slides[{_slide_idx(_SL_EXEC)}]  "
          f"BK=slides[{_slide_idx(_SL_BK)}]")

    print("\n[Slide 1 — 全业务]")
    _replace_chart(_s(0), "Chart 0", slide1_chart_business_type(S1))
    _patch_s1_monthly_trend(_s(0), "Chart 1")
    _fix_chart_dlbls_positions(_s(0), "Chart 1")

    print("\n[Slide 2 — F路径管控]")
    _replace_chart(_chart_slide("Chart 1", _SL_FORECAST), "Chart 1", slide2_forecast_chart(S1))
    _replace_chart(_chart_slide("Chart_YY", _SL_FORECAST), "Chart_YY", slide2_appointment(S1))
    _replace_chart(_chart_slide("Chart_QD", _SL_FORECAST), "Chart_QD", slide2_signed(S1))
    _replace_chart(_chart_slide("Chart_PH", _SL_FORECAST), "Chart_PH", slide2_approved(S1))

    print("\n[Slide 3 — 永明汇报]")
    _replace_chart(_SL_SUNLIFE, "Chart 0", slide3_sunlife_trend(S2))

    print("\n[Slide 4 I charts]")
    _I_CHART_MAP = [
        ("C9013", "天领业务"),
        ("C9017", "成事家办"),
        ("C9021", "BK业务"),
        ("C9025", "同行经代"),
        ("C9029", "永明经代"),
        ("C9033", "合伙转介业务"),
        ("C9037", "ICLUB业务"),
    ]
    for chart_name, seg in _I_CHART_MAP:
        _replace_chart(_chart_slide(chart_name, _SL_BUBBLE), chart_name, slide4_channel_trend(S2, seg))

    print("\n[Slide 5 — 业务端视角]")
    _replace_chart(_SL_BIZ_VIEW, "Chart 0", slide5_donut(S2))
    _fix_chart_dlbls_positions(_SL_BIZ_VIEW, "Chart 0")

    print("\n[Slide 6 — 执行管理端]")
    _replace_chart(_chart_slide("Chart 0", _SL_EXEC), "Chart 0", slide6_weekly_trend(S3))

    print("\n[Slide 8]")
    _SL_8 = _s(7)
    _replace_chart(_chart_slide("Chart 0", _SL_8), "Chart 0", slide8_referrer(S2))
    _replace_chart(_chart_slide("Chart 1", _SL_8), "Chart 1", slide8_top10_ka(S2))

    print("\n[Slide 9]")
    _SL_9 = _s(8)
    _replace_chart(_chart_slide("Chart 0", _SL_9), "Chart 0", slide9_peer_weekly(S3))

    print("\n[Slide 10 — BK charts]")
    _replace_chart(_chart_slide("Chart_YY", _SL_BK), "Chart_YY", slide10_bk_appointment(S2))
    _replace_chart(_chart_slide("Chart_QD", _SL_BK), "Chart_QD", slide10_bk_signed(S2))
    _replace_chart(_chart_slide("Chart_PH", _SL_BK), "Chart_PH", slide10_bk_approved(S2))
    _replace_chart(_chart_slide("Chart 1", _SL_BK), "Chart 1", slide10_bk_ka(S2))
    _replace_chart(_chart_slide("Chart 0", _SL_BK), "Chart 0", slide10_donut_target(S2))
    _fix_chart_dlbls_positions(_chart_slide("Chart 0", _SL_BK), "Chart 0")

    print("\n[Slide 11 — BK周趋势]")
    _SL_11 = _s(10)
    _replace_chart(_chart_slide("Chart 1", _SL_11), "Chart 1", slide11_bank_weekly(S3))
    _replace_chart(_chart_slide("Chart 0", _SL_11), "Chart 0", slide11_branch_ranking(S2))

    prs.save(STAGE1_OUT)
    return STAGE1_OUT


# ============================================================================
# PART 8: apply_w14_patches.py — Weekly patches logic
# ============================================================================
@dataclass
class W14Config:
    here: Path = HERE
    s2_csv: str = "S2-业务端视角.csv"
    s3_csv: str = "S3-执行管理端.csv"
    current_week: Optional[str] = None
    s3_heat_appt: str = "B-APE"

def _w14_update_table_cell(table, row_idx, col_idx, text):
    if row_idx >= len(table.rows) or col_idx >= len(table.columns):
        return False
    cell = table.cell(row_idx, col_idx)
    tf = cell.text_frame
    tf.clear()
    p = tf.add_paragraph()
    p.text = text
    p.alignment = PP_ALIGN.CENTER
    return True

def _w14_patch_ka_tables(cfg: W14Config, prs, current_week):
    s2rows = w14_load_rows(cfg.here / cfg.s2_csv)
    top10 = w14_parse_section_records(s2rows, 'I')
    top10.sort(key=lambda x: x.get('2026批核APE', 0), reverse=True)
    top10 = top10[:10]

    for slide in prs.slides:
        for shape in slide.shapes:
            if not shape.has_table:
                continue
            tbl = shape.table
            if len(tbl.rows) < 12 or len(tbl.columns) < 5:
                continue
            header = tbl.cell(0, 0).text_frame.text.strip()
            if header == 'KA名称':
                for i, ka in enumerate(top10):
                    _w14_update_table_cell(tbl, i+1, 0, ka.get('name', ''))
                    _w14_update_table_cell(tbl, i+1, 1, f"{ka.get('2026批核APE', 0)/10000:.0f}")
                    _w14_update_table_cell(tbl, i+1, 2, f"{ka.get('未批核APE', 0)/10000:.0f}")
                    _w14_update_table_cell(tbl, i+1, 3, f"{ka.get('待签APE', 0)/10000:.0f}")
                break

def _w14_patch_uv_charts(cfg: W14Config, prs, current_week):
    s3rows = w14_load_rows(cfg.here / cfg.s3_csv)
    wk_idx = int(current_week[-2:]) - 1

    for slide in prs.slides:
        for shape in slide.shapes:
            if not shape.has_chart:
                continue
            chart_name = shape.name
            if chart_name.startswith('Chart'):
                try:
                    idx = int(chart_name.replace('Chart', '').strip())
                    if idx in [12, 13, 16]:
                        chart_part = shape.chart._chart_part
                        week_cols = [c for c in s3rows[0] if c.startswith('2026W')]
                        if wk_idx < len(week_cols):
                            col_name = week_cols[wk_idx]
                            data = w14_parse_section_weekly_col(s3rows, 'B-APE', col_name)
                            if data:
                                values = [data.get(k, 0) for k in ['天领业务', '成事家办', 'BK业务']]
                                patch_chart(chart_part, [values])
                except ValueError:
                    pass

def _w14_patch_bank_dashboard(cfg: W14Config, prs):
    s2rows = w14_load_rows(cfg.here / cfg.s2_csv)
    bank_data = w14_parse_section_records(s2rows, 'O')

    for slide in prs.slides:
        for shape in slide.shapes:
            if not shape.has_table:
                continue
            tbl = shape.table
            if len(tbl.rows) == 3 and len(tbl.columns) == 5:
                for i, bank in enumerate(bank_data[:2]):
                    _w14_update_table_cell(tbl, i+1, 0, bank.get('name', ''))
                    _w14_update_table_cell(tbl, i+1, 1, f"{bank.get('2026批核APE', 0)/10000:.0f}")
                    _w14_update_table_cell(tbl, i+1, 2, f"{bank.get('未批核APE', 0)/10000:.0f}")
                    _w14_update_table_cell(tbl, i+1, 3, f"{bank.get('待签APE', 0)/10000:.0f}")
                break

def run_stage2_apply_w14_patches():
    cfg = W14Config()
    s3rows = w14_load_rows(cfg.here / cfg.s3_csv)
    current_week = cfg.current_week or w14_auto_detect_current_week(s3rows, cfg.s3_heat_appt)

    prs = Presentation(STAGE1_OUT)

    _w14_patch_ka_tables(cfg, prs, current_week)
    _w14_patch_uv_charts(cfg, prs, current_week)
    _w14_patch_bank_dashboard(cfg, prs)

    prs.save(STAGE2_OUT)
    return STAGE2_OUT, current_week


# ============================================================================
# PART 9: final_gap_patch.py — Final fixes logic
# ============================================================================
IN_PPT = STAGE2_OUT
OUT_PPT = STAGE3_OUT

def _fgp_update_license_table(prs, S1):
    H = S1.get('H')
    if H is None:
        return
    for slide in prs.slides:
        for shape in slide.shapes:
            if not shape.has_table:
                continue
            tbl = shape.table
            if len(tbl.rows) < 4 or len(tbl.columns) < 6:
                continue
            header = tbl.cell(0, 0).text_frame.text.strip()
            if header == '牌照':
                licenses = ['DW Bank', 'JF', 'Sub Total']
                for i, lic in enumerate(licenses):
                    row = H[H['牌照'] == lic]
                    if len(row):
                        r = row.iloc[0]
                        _w14_update_table_cell(tbl, i+1, 0, lic)
                        for j, col in enumerate(['2026-01', '2026-02', '2026-03', '2026-04', '未批核']):
                            _w14_update_table_cell(tbl, i+1, j+1, f"{num(r.get(col, 0))/10000:.0f}")
                break

def _fgp_update_top10_product_table(prs, S4):
    C = S4.get('C')
    if C is None:
        return
    df = C[C['保司'] == '永明'].copy()
    df['_ape'] = df['APE'].apply(num)
    df = df.sort_values('_ape', ascending=False).head(10)

    for slide in prs.slides:
        for shape in slide.shapes:
            if not shape.has_table:
                continue
            tbl = shape.table
            if len(tbl.rows) < 12 or len(tbl.columns) < 4:
                continue
            header = tbl.cell(0, 0).text_frame.text.strip()
            if header == '产品名称':
                for i, (_, row) in enumerate(df.iterrows()):
                    _w14_update_table_cell(tbl, i+1, 0, row['产品名称'])
                    _w14_update_table_cell(tbl, i+1, 1, f"{num(row['APE'])/10000:.0f}")
                    _w14_update_table_cell(tbl, i+1, 2, f"{int(num(row['件数']))}")
                break

def _fgp_rebuild_heat_matrix(prs, S3, section_key):
    for slide in prs.slides:
        for shape in slide.shapes:
            if not shape.has_chart:
                continue
            chart_name = shape.name
            if chart_name.startswith('Chart'):
                try:
                    idx = int(chart_name.replace('Chart', '').strip())
                    if idx in [13, 14, 15]:
                        chart_part = shape.chart._chart_part
                        df = S3.get(section_key)
                        if df is not None:
                            week_cols = [c for c in df.columns if c.startswith('2026W')]
                            categories = [c.replace('2026', '') for c in week_cols]
                            data_rows = []
                            for _, row in df.iterrows():
                                name = row.iloc[0]
                                if name == '合计':
                                    continue
                                values = [to_m(row[c]) for c in week_cols]
                                data_rows.append((name, values))
                            series_spec = [vals for _, vals in data_rows]
                            series_names = [name for name, _ in data_rows]
                            patch_chart(chart_part, series_spec, categories=categories, series_names=series_names)
                except ValueError:
                    pass

def _fgp_update_so_what(prs, ctx):
    slots = find_sowhat_slots(prs)
    for slot in slots:
        content_shape = slot['content']
        if content_shape is None:
            continue
        tf = content_shape.text_frame
        tf.clear()
        slide_num = slot['slide']
        generator_key = f's{slide_num}_default'
        for key in SOWHAT_GENERATORS:
            if key.startswith(f's{slide_num}_'):
                generator_key = key
                break
        generator = SOWHAT_GENERATORS.get(generator_key)
        if generator:
            try:
                text = generator(ctx)
                p = tf.add_paragraph()
                p.text = text
                p.alignment = PP_ALIGN.LEFT
            except Exception:
                pass

def run_stage3_final_gap_patch(current_week):
    prs = Presentation(IN_PPT)
    S1 = parse_blocks(CSV_S1)
    S2 = parse_blocks(CSV_S2)
    S3 = parse_blocks(CSV_S3)
    S4 = parse_blocks(CSV_S4)

    _fgp_update_license_table(prs, S1)
    _fgp_update_top10_product_table(prs, S4)
    _fgp_rebuild_heat_matrix(prs, S3, 'B-APE')
    _fgp_rebuild_heat_matrix(prs, S3, 'C-APE')
    _fgp_rebuild_heat_matrix(prs, S3, 'D-APE')

    ctx = {
        'S1': S1, 'S2': S2, 'S3': S3, 'S4': S4,
        'current_week': current_week,
        's3_rows': w14_load_rows(CSV_S3),
    }
    _fgp_update_so_what(prs, ctx)

    prs.save(OUT_PPT)
    return OUT_PPT


# ============================================================================
# PART 10: Main orchestration — run_full_pipeline.py logic
# ============================================================================
def main():
    print("=" * 60)
    print("run_full_pipeline_merge.py — Merged PPT update pipeline")
    print("=" * 60)

    if not TEMPLATE_SRC.exists():
        print(f"❌ Template not found: {TEMPLATE_SRC}")
        sys.exit(1)

    if TEMPLATE_SRC.name != 'template.pptx':
        shutil.copy(TEMPLATE_SRC, HERE / 'template.pptx')
        print(f"📋 Copied template to: {HERE / 'template.pptx'}")

    print("\n" + "=" * 60)
    print("Stage 1 — update_ppt.py")
    print("=" * 60)
    try:
        run_stage1_update_ppt()
        print(f"✅ Stage 1 complete: {STAGE1_OUT}")
    except Exception as e:
        print(f"❌ Stage 1 failed: {e}")
        sys.exit(1)

    print("\n" + "=" * 60)
    print("Stage 2 — apply_w14_patches.py")
    print("=" * 60)
    try:
        _, current_week = run_stage2_apply_w14_patches()
        print(f"✅ Stage 2 complete: {STAGE2_OUT}")
        print(f"   Current week detected: {current_week}")
    except Exception as e:
        print(f"❌ Stage 2 failed: {e}")
        sys.exit(1)

    print("\n" + "=" * 60)
    print("Stage 3 — final_gap_patch.py")
    print("=" * 60)
    try:
        run_stage3_final_gap_patch(current_week)
        print(f"✅ Stage 3 complete: {STAGE3_OUT}")
    except Exception as e:
        print(f"❌ Stage 3 failed: {e}")
        sys.exit(1)

    for f in [HERE / '周业绩汇报PPT_AUTO_UPDATED.pptx', HERE / '周业绩汇报PPT_AUTO_UPDATED_updated.pptx']:
        if f.exists():
            os.remove(f)

    print("\n" + "=" * 60)
    print(f"🎉 All stages completed successfully!")
    print(f"   Output: {STAGE3_OUT}")
    print("=" * 60)

if __name__ == "__main__":
    main()