#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
万相台无界 · 四象限周报生成器（计划级 + 商品级）

用法：
    python weekly_report.py --plans 计划报表.csv --items 商品报表.csv --period "09-21至09-27"
    python weekly_report.py --plans 计划报表.csv --roi-thr 6.0 --sp-thr 0.85
    python weekly_report.py --plans 计划报表.csv --output 周报.html

四象限口径（与阿里妈妈万相台无界运营方法论一致）：
    横轴 ROI  分界 = 大盘 ROI
    纵轴 潜客占比 分界 = 大盘潜客占比
    花费 < min-cost 或 访问人数 < min-uv 判为「样本不足」，不参与判定

产物是单个 HTML 文件：图表为自绘 SVG、交互为原生 JS，零外部依赖，
离线可开、可直接发给同事，体积通常在 100KB 级别。
"""

import os
import re
import sys
import json
import math
import argparse
from datetime import datetime

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

import importlib.util

HERE = os.path.dirname(os.path.abspath(__file__))
_spec = importlib.util.spec_from_file_location("az", os.path.join(HERE, "analyze.py"))
az = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(az)

# 象限定义
QUADS = {
    1: {"n": "优质拉新", "c": "#22a06b", "cond": "高 ROI ＋ 高潜客占比",
        "act": "拉新转化兼得 → 加大投放", "short": "加大投放"},
    2: {"n": "防守收割", "c": "#f0883e", "cond": "高 ROI ＋ 低潜客占比",
        "act": "存量收割型 → 控制花费占比 ≤30%", "short": "控比≤30%"},
    3: {"n": "人群矫正", "c": "#8b5cf6", "cond": "低 ROI ＋ 高潜客占比",
        "act": "拉新强转化弱 → 矫正人群精准度", "short": "修定向"},
    4: {"n": "建议调整", "c": "#e5484d", "cond": "低 ROI ＋ 低潜客占比",
        "act": "效率偏低 → 优先调整优化，确认无效再停投", "short": "调整"},
    0: {"n": "样本不足", "c": "#9ca3af", "cond": "花费或访问人数过低",
        "act": "暂不判定，先补量或合并计划", "short": "观察"},
}

# 商品规模分层（ROI × 花费）
TIERS = {
    "爆款": {"c": "#22a06b", "act": "规模与效率兼得 → 保预算、拓同类"},
    "潜力": {"c": "#2f6fed", "act": "效率高但没量 → 加预算、拓词拓人群"},
    "问题": {"c": "#e5484d", "act": "花钱多效率低 → 优先降价/重做定向"},
    "长尾": {"c": "#9ca3af", "act": "量小效低 → 收缩或合并投放"},
}


# ---------------------------------------------------------------------------
# 数据装载
# ---------------------------------------------------------------------------
def load_rows(path):
    """读报表 → 标准字段的 dict 列表"""
    df = az.load_table(path)
    m = az.build_mapping(df)
    if "cost" not in m:
        raise RuntimeError("未识别到「花费」列，该文件列名为：%s" % list(df.columns))

    def col(std, numeric=True):
        if std not in m:
            return None
        s = df[m[std]]
        return s.map(az.to_num) if numeric else s.astype(str)

    n = len(df)

    def arr(std, numeric=True):
        s = col(std, numeric)
        if s is None:
            return [0.0] * n if numeric else [""] * n
        return s.tolist()

    # 名称列：计划名优先，商品报表则取主体名称
    name_col = col("plan", numeric=False)
    if name_col is None:
        name_col = col("item", numeric=False)

    cols = {
        "s": arr("scene", False), "id": arr("cid", False),
        "c": arr("cost"), "g": arr("gmv"), "r": arr("roi"),
        "im": arr("imp"), "ck": arr("clk"), "ctr": arr("ctr"),
        "cpc": arr("ppc"), "cvr": arr("cvr"), "o": arr("order"),
        "u": arr("uv"), "q": arr("qk"), "p": arr("sp"),
        "nr": arr("newRate"), "nc": arr("newC"), "b": arr("byr"),
        "ca": arr("cart"), "f": arr("fav"),
        "ni": arr("natImp"), "ng": arr("natGmv"),
    }

    rows = []
    for i in range(n):
        r = {"n": (name_col[i] if name_col is not None else "")}
        for k, v in cols.items():
            r[k] = v[i]
        r["n"] = str(r["n"]).strip()
        # 派生
        if not r["r"] and r["c"]:
            r["r"] = r["g"] / r["c"]
        if not r["p"] and r["u"]:
            r["p"] = r["q"] / r["u"]
        if not r["ctr"] and r["im"]:
            r["ctr"] = r["ck"] / r["im"]
        if not r["cpc"] and r["ck"]:
            r["cpc"] = r["c"] / r["ck"]
        if not r["cvr"] and r["ck"]:
            r["cvr"] = r["o"] / r["ck"]
        if not r.get("ni"):
            r["ni"] = 0.0
        if not r.get("ng"):
            r["ng"] = 0.0
        r["nat_ratio"] = (r["ni"] / r["im"]) if r["im"] else 0.0
        rows.append(r)
    return rows


def total_of(rows, key):
    return sum(float(r.get(key) or 0) for r in rows)




def classify(r, roi_thr, sp_thr, min_cost, min_uv):
    if r["c"] < min_cost or r["u"] < min_uv or not r["r"] or not r["p"]:
        return 0
    hr = r["r"] >= roi_thr
    hs = r["p"] >= sp_thr
    return 1 if (hr and hs) else 2 if (hr and not hs) else 3 if (not hr and hs) else 4


def quadrant_stats(rows):
    st = {k: {"n": 0, "c": 0.0, "g": 0.0, "u": 0.0, "q": 0.0, "o": 0, "b": 0}
          for k in (0, 1, 2, 3, 4)}
    for r in rows:
        v = st[r["qd"]]
        v["n"] += 1
        v["c"] += r["c"]
        v["g"] += r["g"]
        v["u"] += r["u"]
        v["q"] += r["q"]
        v["o"] += r["o"]
        v["b"] += r["b"]
    for k, v in st.items():
        v["roi"] = v["g"] / v["c"] if v["c"] else 0
        v["sp"] = v["q"] / v["u"] if v["u"] else 0
    return st


def tier_of(r, roi_thr, cost_thr):
    hr = r["r"] >= roi_thr
    hc = r["c"] >= cost_thr
    return "爆款" if (hr and hc) else "潜力" if (hr and not hc) else \
        "问题" if (not hr and hc) else "长尾"


# ---------------------------------------------------------------------------
# SVG 图表（自绘，零依赖）
# ---------------------------------------------------------------------------
def esc(s):
    return (str(s).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
            .replace('"', "&quot;"))


def fmt_money(v):
    v = float(v or 0)
    return "%.2f万" % (v / 10000) if abs(v) >= 10000 else "%.0f" % v


def svg_scatter(rows, roi_thr, sp_thr, W=1080, H=460):
    """四象限散点：x=ROI，y=潜客占比，气泡=花费"""
    if not rows:
        return ""
    xs = sorted(r["r"] for r in rows if r["r"] > 0)
    if not xs:
        return ""
    xmax = xs[int(len(xs) * 0.97)] * 1.15
    xmax = max(xmax, roi_thr * 1.4)
    # 纵轴自适应：按实际潜客占比的分布范围确定显示区间，避免散点全部挤在上部
    yv = [r["p"] for r in rows if r["p"]] + [sp_thr]
    lo, hi = min(yv), max(yv)
    ypad = max(0.06, (hi - lo) * 0.15)
    ymin = max(0.0, lo - ypad)
    ymax = min(1.0, hi + ypad)
    if ymax - ymin < 0.12:                       # 分布过度集中时保底留白
        mid = (ymax + ymin) / 2
        ymin, ymax = max(0.0, mid - 0.06), min(1.0, mid + 0.06)
    cmax = max(r["c"] for r in rows) or 1

    ml, mr, mt, mb = 66, 24, 24, 48
    pw, ph = W - ml - mr, H - mt - mb

    def X(v):
        return ml + (min(v, xmax) / xmax) * pw

    def Y(v):
        vv = min(max(v, ymin), ymax)
        return mt + ph - ((vv - ymin) / (ymax - ymin)) * ph

    p = ['<svg viewBox="0 0 %d %d" width="100%%" style="max-width:1080px" id="scatter">'
         % (W, H)]
    p.append('<rect width="%d" height="%d" fill="#fff"/>' % (W, H))
    # 象限底色
    p.append('<rect x="%.1f" y="%.1f" width="%.1f" height="%.1f" fill="#22a06b" opacity="0.05"/>'
             % (X(roi_thr), mt, ml + pw - X(roi_thr), Y(sp_thr) - mt))
    p.append('<rect x="%.1f" y="%.1f" width="%.1f" height="%.1f" fill="#e5484d" opacity="0.05"/>'
             % (ml, Y(sp_thr), X(roi_thr) - ml, mt + ph - Y(sp_thr)))
    # 网格（纵轴刻度随自适应区间变化）
    for i in range(5):
        y = mt + ph * i / 4
        val = ymax - (ymax - ymin) * i / 4
        p.append('<line x1="%d" y1="%.1f" x2="%.1f" y2="%.1f" stroke="#eef1f5"/>' % (ml, y, ml + pw, y))
        p.append('<text x="%d" y="%.1f" font-size="11" fill="#94a3b8" text-anchor="end">%.0f%%</text>'
                 % (ml - 8, y + 4, val * 100))
    for i in range(6):
        x = ml + pw * i / 5
        p.append('<text x="%.1f" y="%.1f" font-size="11" fill="#94a3b8" text-anchor="middle">%.1f</text>'
                 % (x, mt + ph + 18, xmax * i / 5))
    # 阈值线
    p.append('<line x1="%.1f" y1="%d" x2="%.1f" y2="%.1f" stroke="#334155" stroke-dasharray="5 4"/>'
             % (X(roi_thr), mt, X(roi_thr), mt + ph))
    p.append('<line x1="%d" y1="%.1f" x2="%.1f" y2="%.1f" stroke="#334155" stroke-dasharray="5 4"/>'
             % (ml, Y(sp_thr), ml + pw, Y(sp_thr)))
    p.append('<text x="%.1f" y="%d" font-size="11" fill="#475569">ROI %.2f</text>'
             % (X(roi_thr) + 5, mt + 12, roi_thr))
    p.append('<text x="%d" y="%.1f" font-size="11" fill="#475569">潜客占比 %.1f%%</text>'
             % (ml + 5, Y(sp_thr) - 6, sp_thr * 100))
    # 气泡（不用 <title>：浏览器原生提示有约 1 秒延迟且样式不可控，改由 JS 自绘 tooltip）
    for i, r in sorted(enumerate(rows), key=lambda kv: -kv[1]["c"]):
        rad = 3 + 17 * math.sqrt(r["c"] / cmax)
        color = QUADS[r["qd"]]["c"]
        p.append('<circle class="pt" data-q="%d" data-i="%d" cx="%.1f" cy="%.1f" r="%.1f" '
                 'fill="%s" opacity="0.55" stroke="%s" stroke-width="1"/>'
                 % (r["qd"], i, X(r["r"]), Y(r["p"]), rad, color, color))
    p.append('<text x="%d" y="%d" font-size="11.5" fill="#94a3b8">横轴 ROI ｜ 纵轴 潜客占比 ｜ 气泡大小 花费</text>'
             % (ml, H - 8))
    p.append("</svg>")
    return "".join(p)


def svg_group_bars(st, tc, tg, W=1080, H=250):
    """各象限 花费占比 vs 成交占比"""
    keys = [1, 2, 3, 4, 0]
    labels = [QUADS[k]["n"] for k in keys]
    ml, mt, mb = 60, 30, 40
    pw, ph = W - ml - 30, H - mt - mb
    slot = pw / len(keys)
    vmax = max(max(st[k]["c"] / tc if tc else 0, st[k]["g"] / tg if tg else 0)
               for k in keys) * 1.25 or 1
    p = ['<svg viewBox="0 0 %d %d" width="100%%" style="max-width:1080px">' % (W, H)]
    p.append('<rect width="%d" height="%d" fill="#fff"/>' % (W, H))
    for i in range(4):
        y = mt + ph * i / 3
        p.append('<line x1="%d" y1="%.1f" x2="%.1f" y2="%.1f" stroke="#eef1f5"/>' % (ml, y, ml + pw, y))
        p.append('<text x="%d" y="%.1f" font-size="11" fill="#94a3b8" text-anchor="end">%.0f%%</text>'
                 % (ml - 8, y + 4, (1 - i / 3) * vmax * 100))
    for i, k in enumerate(keys):
        cx = ml + slot * i + slot / 2
        bw = min(30.0, slot * 0.3)
        cs = st[k]["c"] / tc if tc else 0
        gs = st[k]["g"] / tg if tg else 0
        h1 = (cs / vmax) * ph
        h2 = (gs / vmax) * ph
        p.append('<rect class="bpt" data-tip="%s · 花费占比 %.1f%%" x="%.1f" y="%.1f" '
                 'width="%.1f" height="%.1f" fill="#ff5000" rx="2"/>'
                 % (QUADS[k]["n"], cs * 100, cx - bw - 2, mt + ph - h1, bw, h1))
        p.append('<rect class="bpt" data-tip="%s · 成交占比 %.1f%%" x="%.1f" y="%.1f" '
                 'width="%.1f" height="%.1f" fill="#2f6fed" rx="2"/>'
                 % (QUADS[k]["n"], gs * 100, cx + 2, mt + ph - h2, bw, h2))
        p.append('<text x="%.1f" y="%.1f" font-size="12" fill="#334155" text-anchor="middle">%s</text>'
                 % (cx, mt + ph + 18, labels[i]))
        p.append('<text x="%.1f" y="%.1f" font-size="10.5" fill="#94a3b8" text-anchor="middle">%d个</text>'
                 % (cx, mt + ph + 32, st[k]["n"]))
    p.append('<rect x="%d" y="10" width="10" height="10" rx="2" fill="#ff5000"/>' % ml)
    p.append('<text x="%d" y="19" font-size="12" fill="#475569">花费占比</text>' % (ml + 15))
    p.append('<rect x="%d" y="10" width="10" height="10" rx="2" fill="#2f6fed"/>' % (ml + 90))
    p.append('<text x="%d" y="19" font-size="12" fill="#475569">成交占比</text>' % (ml + 105))
    p.append("</svg>")
    return "".join(p)


def svg_scene_stack(rows, W=1080, H=280):
    """场景 × 象限 花费堆叠"""
    scenes = {}
    for r in rows:
        scenes.setdefault(r["s"] or "未分类", {})
        scenes[r["s"] or "未分类"][r["qd"]] = scenes[r["s"] or "未分类"].get(r["qd"], 0) + r["c"]
    items = sorted(scenes.items(), key=lambda kv: -sum(kv[1].values()))
    if not items:
        return ""
    ml, mt, mb = 150, 26, 30
    pw, ph = W - ml - 90, H - mt - mb
    vmax = max(sum(v.values()) for _, v in items) or 1
    rowh = ph / len(items)
    p = ['<svg viewBox="0 0 %d %d" width="100%%" style="max-width:1080px">' % (W, H)]
    p.append('<rect width="%d" height="%d" fill="#fff"/>' % (W, H))
    for i, (s, v) in enumerate(items):
        y = mt + i * rowh + rowh * 0.15
        h = rowh * 0.7
        x = ml
        tot = sum(v.values())
        for q in (1, 2, 3, 4, 0):
            c = v.get(q, 0)
            if not c:
                continue
            w = (c / vmax) * pw
            p.append('<rect class="bpt" data-tip="%s · %s｜花费 ¥%s（%.1f%%）" x="%.1f" y="%.1f" '
                     'width="%.1f" height="%.1f" fill="%s" opacity="0.85"/>'
                     % (esc(s), QUADS[q]["n"], fmt_money(c), c / tot * 100,
                        x, y, w, h, QUADS[q]["c"]))
            x += w
        name = s if len(s) <= 12 else s[:11] + "…"
        p.append('<text x="%d" y="%.1f" font-size="12" fill="#334155" text-anchor="end">%s</text>'
                 % (ml - 10, y + h * 0.7, esc(name)))
        p.append('<text x="%.1f" y="%.1f" font-size="11" fill="#64748b">%s</text>'
                 % (x + 8, y + h * 0.7, fmt_money(tot)))
    lx = ml
    for q in (1, 2, 3, 4, 0):
        p.append('<rect x="%.1f" y="8" width="9" height="9" rx="2" fill="%s"/>' % (lx, QUADS[q]["c"]))
        p.append('<text x="%.1f" y="16.5" font-size="11.5" fill="#475569">%s</text>' % (lx + 13, QUADS[q]["n"]))
        lx += len(QUADS[q]["n"]) * 12 + 42
    p.append("</svg>")
    return "".join(p)


# ---------------------------------------------------------------------------
# HTML
# ---------------------------------------------------------------------------
CSS = """
*{box-sizing:border-box}
body{margin:0;background:#f4f6f9;color:#1f2937;font:14px/1.65 -apple-system,BlinkMacSystemFont,"PingFang SC","Microsoft YaHei",Arial,sans-serif}
.wrap{max-width:1120px;margin:0 auto;padding:28px 18px 60px}
h1{font-size:24px;margin:0 0 4px}
h2{font-size:18px;margin:34px 0 12px;padding-left:10px;border-left:4px solid #ff5000}
.sub{color:#64748b;font-size:12.5px;margin-bottom:20px}
.card{background:#fff;border:1px solid #e6eaf0;border-radius:12px;padding:16px 18px;margin-bottom:14px;box-shadow:0 1px 2px rgba(16,24,40,.04)}
.kpis{display:grid;grid-template-columns:repeat(7,1fr);gap:10px}
.kpi{background:#fff;border:1px solid #e6eaf0;border-radius:11px;padding:12px 14px}
.kpi .k{font-size:11.5px;color:#64748b;white-space:nowrap}
.kpi .v{font-size:20px;font-weight:600;margin-top:3px;letter-spacing:-.3px}
.kpi .u{font-size:11px;color:#94a3b8}
.qcards{display:grid;grid-template-columns:repeat(5,1fr);gap:10px}
.qcard{background:#fff;border:1px solid #e6eaf0;border-top:3px solid #ccc;border-radius:11px;padding:13px 14px;cursor:pointer;transition:.15s}
.qcard:hover{transform:translateY(-2px);box-shadow:0 4px 12px rgba(16,24,40,.09)}
.qcard .t{font-weight:600;font-size:14px;display:flex;align-items:center;gap:6px}
.qcard .dot{width:9px;height:9px;border-radius:50%}
.qcard .c{font-size:11.5px;color:#64748b;margin:4px 0 8px}
.qcard .m{font-size:12px;color:#334155}
.qcard .m b{font-size:15px}
.qcard .a{margin-top:8px;font-size:11.5px;color:#0f766e;background:#f0fdfa;padding:6px 8px;border-radius:6px}
table{width:100%;border-collapse:collapse;font-size:12.5px}
th{background:#f7f9fc;color:#475569;font-weight:600;text-align:right;padding:8px 9px;border-bottom:1px solid #e6eaf0;white-space:nowrap;cursor:pointer;user-select:none;position:sticky;top:0}
th:hover{background:#eef2f8}
th:first-child,td:first-child{text-align:left}
td{padding:7px 9px;border-bottom:1px solid #f2f4f7;text-align:right;white-space:nowrap}
td.nm{text-align:left;max-width:300px;overflow:hidden;text-overflow:ellipsis}
tr:hover td{background:#fafbfd}
.bad{color:#dc2626}.good{color:#16a34a}
.bar{display:flex;gap:8px;flex-wrap:wrap;align-items:center;margin-bottom:10px}
.bar input,.bar select{border:1px solid #d8dee8;border-radius:7px;padding:6px 10px;font-size:12.5px;outline:none;background:#fff}
.bar input:focus,.bar select:focus{border-color:#ff5000}
.btn{border:1px solid #d8dee8;background:#fff;border-radius:7px;padding:6px 12px;font-size:12.5px;cursor:pointer}
.btn:hover{background:#f7f9fc}
.btn.pri{background:#ff5000;border-color:#ff5000;color:#fff}
.pg{display:flex;gap:6px;align-items:center;justify-content:flex-end;margin-top:10px;font-size:12.5px;color:#64748b}
.qtag{display:inline-block;padding:1px 7px;border-radius:20px;font-size:11px;color:#fff}
.note{font-size:12px;color:#94a3b8;margin-top:8px}
.sim{display:grid;grid-template-columns:repeat(3,1fr);gap:12px}
.sim div{background:#fafbfd;border:1px solid #eef1f5;border-radius:9px;padding:12px 14px}
.sim .h{font-size:12.5px;color:#64748b}
.sim .b{font-size:19px;font-weight:600;margin:4px 0}
.sim .d{font-size:11.5px;color:#64748b}
.two{display:grid;grid-template-columns:1fr 1fr;gap:14px}
.chips{display:flex;gap:8px;flex-wrap:wrap;margin-bottom:11px}
.chip{border:1px solid #d8dee8;background:#fff;border-radius:20px;padding:5px 13px;font-size:12.5px;cursor:pointer;display:flex;align-items:center;gap:6px;transition:.15s;user-select:none}
.chip:hover{border-color:#94a3b8}
.chip.on{background:#1f2937;border-color:#1f2937;color:#fff}
.chip .n{font-weight:600}
.chip .c{opacity:.7;font-size:11.5px}
.doti{width:8px;height:8px;border-radius:50%;flex:none}
.mini{display:inline-flex;height:9px;border-radius:3px;overflow:hidden;width:104px;vertical-align:middle;background:#eef1f5}
.mini i{display:block;height:100%}
#tip{position:fixed;z-index:99;pointer-events:none;opacity:0;transition:opacity .08s;
  background:rgba(17,24,39,.95);color:#fff;border-radius:9px;padding:10px 13px;font-size:12px;
  line-height:1.75;box-shadow:0 8px 24px rgba(0,0,0,.22);max-width:300px}
#tip .tt{font-weight:600;font-size:12.5px;margin-bottom:3px;word-break:break-all}
#tip .sc{font-size:11px;color:#9ca3af;margin-bottom:5px}
#tip .rw{display:flex;justify-content:space-between;gap:14px}
#tip .rw span:first-child{color:#9ca3af}
#tip .hint{margin-top:6px;padding-top:5px;border-top:1px solid rgba(255,255,255,.15);
  font-size:11px;color:#fbbf24}
#scatter .pt{cursor:pointer;transition:opacity .1s,stroke-width .1s}
#scatter .pt:hover{opacity:.95!important;stroke-width:2.5}
.legend{display:flex;gap:8px;flex-wrap:wrap;margin-top:10px}
.legend .lg{border:1px solid #e6eaf0;background:#fff;border-radius:20px;padding:4px 12px;
  font-size:12px;cursor:pointer;display:flex;align-items:center;gap:6px;user-select:none}
.legend .lg.off{opacity:.35;text-decoration:line-through}
.toprow{cursor:pointer}
.toprow:hover td{background:#fff7ed!important}
.hl{background:#fff7ed!important}
@media(max-width:900px){.kpis{grid-template-columns:repeat(3,1fr)}.qcards{grid-template-columns:repeat(2,1fr)}.sim{grid-template-columns:1fr}.two{grid-template-columns:1fr}}
"""

JS = """
var Q=__QUADS__, TIER=__TIERS__;
function fm(v){v=v||0;return '¥'+(v>=10000?(v/10000).toFixed(2)+'万':Math.round(v).toLocaleString());}
function pc(v,n){return v==null?'—':(v*100).toFixed(n===undefined?1:n)+'%';}
function fx(v,n){return (v||0).toFixed(n===undefined?2:n);}

function Table(opt){
  var data=opt.data, cols=opt.cols, tb=document.getElementById(opt.tid),
      pager=document.getElementById(opt.pid), count=document.getElementById(opt.cid);
  var fq='', fs='', kw='', fIdx=null, sk=opt.sortKey||'c', sd=-1, page=1, PS=opt.ps||25;
  var view=function(){
    var d=data.filter(function(r){
      if(fIdx && !fIdx.has(r.ix)) return false;
      if(fq && String(r.qd)!==fq) return false;
      if(fs && r.s!==fs) return false;
      if(kw && (String(r.n)+' '+String(r.id||'')).toLowerCase().indexOf(kw)===-1) return false;
      return true;
    });
    d.sort(function(a,b){var x=a[sk],y=b[sk];if(typeof x==='string')return sd*x.localeCompare(y);return sd*(x-y);});
    return d;
  };
  function render(){
    var d=view(), tp=Math.max(1,Math.ceil(d.length/PS));
    if(page>tp) page=tp;
    var h='<thead><tr>';
    cols.forEach(function(c){
      var ar=sk===c.k?(sd<0?' ↓':' ↑'):'';
      h+='<th data-k="'+c.k+'">'+c.t+ar+'</th>';
    });
    h+='</tr></thead><tbody>';
    d.slice((page-1)*PS,page*PS).forEach(function(r){
      h+='<tr>';
      cols.forEach(function(c){
        var v=r[c.k];
        if(c.k==='qd'){h+='<td><span class="qtag" style="background:'+Q[v].c+'">'+Q[v].n+'</span></td>';}
        else if(c.k==='tier'){h+='<td><span class="qtag" style="background:'+TIER[v].c+'">'+v+'</span></td>';}
        else if(c.m==='money'){h+='<td>'+fm(v)+'</td>';}
        else if(c.m==='pct'){h+='<td>'+pc(v)+'</td>';}
        else if(c.m==='x'){h+='<td>'+fx(v)+'</td>';}
        else if(c.m==='num'){h+='<td>'+Math.round(v||0).toLocaleString()+'</td>';}
        else{h+='<td class="nm" title="'+String(v).replace(/"/g,'')+'">'+v+'</td>';}
      });
      h+='</tr>';
    });
    h+='</tbody>';
    tb.innerHTML=h;
    count.textContent='共 '+d.length+' 条';
    pager.innerHTML='<button class="btn" data-p="-1">上一页</button><span style="padding:0 8px">'
      +page+' / '+tp+'</span><button class="btn" data-p="1">下一页</button>';
    [].forEach.call(tb.querySelectorAll('th'),function(th){
      th.onclick=function(){var k=th.getAttribute('data-k');if(sk===k)sd=-sd;else{sd=(k==='n'||k==='s')?1:-1;}sk=k;render();};
    });
    [].forEach.call(pager.querySelectorAll('button'),function(b){
      b.onclick=function(){page+=parseInt(b.getAttribute('data-p'));render();};
    });
  }
  function csv(){
    var d=view(), head=cols.map(function(c){return c.t;}).join(',');
    var lines=d.map(function(r){return cols.map(function(c){
      var v=(c.k==='qd')?Q[r.qd].n:r[c.k];
      return '"'+String(v).replace(/"/g,'')+'"';}).join(',');});
    var blob=new Blob(['\\ufeff'+head+'\\n'+lines.join('\\n')],{type:'text/csv;charset=utf-8'});
    var a=document.createElement('a');a.href=URL.createObjectURL(blob);
    a.download=opt.csvName||'导出.csv';a.click();
  }
  return {render:render,
          setQ:function(v){fq=v;page=1;render();},
          setS:function(v){fs=v;page=1;render();},
          setKw:function(v){kw=v.toLowerCase().trim();page=1;render();},
          setIdx:function(list){fIdx=(list&&list.length)?new Set(list):null;page=1;render();},
          clear:function(){fIdx=null;fq='';fs='';kw='';page=1;render();},
          csv:csv,
          scenes:function(){var s={};data.forEach(function(r){s[r.s||'未分类']=1;});return Object.keys(s);}};
}

document.addEventListener('DOMContentLoaded',function(){
  var PD=__PLANS__, IDATA=__ITEMS__;

  /* ---- 散点：自绘 tooltip（替代原生 title，无延迟、样式可控） ---- */
  var tip=document.createElement('div');tip.id='tip';document.body.appendChild(tip);
  var SC=document.getElementById('scatter');
  function tipHTML(d){
    var s='<div class="tt">'+(d.n||'')+'</div><div class="sc">'+(d.s||'')+' ｜ '+Q[d.qd].n+'</div>';
    s+='<div class="rw"><span>花费</span><span>¥'+Math.round(d.c).toLocaleString()+'</span></div>';
    s+='<div class="rw"><span>成交金额</span><span>¥'+Math.round(d.g).toLocaleString()+'</span></div>';
    s+='<div class="rw"><span>投产比 ROI</span><span>'+d.r.toFixed(2)+'</span></div>';
    s+='<div class="rw"><span>潜客占比</span><span>'+(d.p*100).toFixed(1)+'%</span></div>';
    s+='<div class="rw"><span>CTR / CVR</span><span>'+(d.ctr*100).toFixed(2)+'% / '+(d.cvr*100).toFixed(2)+'%</span></div>';
    s+='<div class="rw"><span>点击 / 成交笔数</span><span>'+Math.round(d.ck).toLocaleString()+' / '+Math.round(d.o)+'</span></div>';
    s+='<div class="hint">点击可只看「'+Q[d.qd].n+'」</div>';
    return s;
  }
  function isPt(t){return t&&t.classList&&t.classList.contains('pt');}
  function hasTip(t){return t&&t.getAttribute&&t.getAttribute('data-tip');}
  document.addEventListener('mouseover',function(e){
    var t=e.target;
    if(isPt(t)){var d=PD[+t.getAttribute('data-i')];if(d)tip.innerHTML=tipHTML(d),tip.style.opacity=1;return;}
    if(hasTip(t)){tip.innerHTML='<div class="tt">'+t.getAttribute('data-tip')+'</div>';tip.style.opacity=1;}
  });
  document.addEventListener('mousemove',function(e){
    if(tip.style.opacity==='0'||tip.style.opacity==='')return;
    var x=e.clientX+16,y=e.clientY+16;
    if(x+310>window.innerWidth)x=e.clientX-310;
    if(y+190>window.innerHeight)y=e.clientY-190;
    tip.style.left=x+'px';tip.style.top=y+'px';
  });
  document.addEventListener('mouseout',function(e){
    if(isPt(e.target)||hasTip(e.target))tip.style.opacity=0;
  });
  var hidQ={};
  [].forEach.call(document.querySelectorAll('#sclegend .lg'),function(el){
    el.onclick=function(){
      var q=el.getAttribute('data-q');
      hidQ[q]=!hidQ[q];
      el.className='lg'+(hidQ[q]?' off':'');
      [].forEach.call(document.querySelectorAll('#scatter .pt'),function(c){
        c.style.display=hidQ[c.getAttribute('data-q')]?'none':'';
      });
    };
  });

  var PT=Table({data:PD,cols:__PLANCOLS__,tid:'ptb',pid:'ppg',cid:'pcnt',sortKey:'c',csvName:__CSVNAME1__});
  PT.render();

  function syncQ(v){
    v=v||'';
    [].forEach.call(document.querySelectorAll('.chip'),function(c){
      if(c.getAttribute('data-q')===v){c.classList.add('on');}else{c.classList.remove('on');}
    });
    [].forEach.call(document.querySelectorAll('.qcard'),function(x){
      x.style.outline=(v && x.getAttribute('data-q')===v)?'2px solid #ff5000':'';
    });
  }
  function goPlan(){
    var t=document.getElementById('plan-table');
    if(t)t.scrollIntoView({behavior:'smooth',block:'start'});
  }
  [].forEach.call(document.querySelectorAll('.chip'),function(el){
    el.onclick=function(){
      var v=el.getAttribute('data-q');
      PT.setIdx(null); PT.setQ(el.classList.contains('on')?'':v);
      syncQ(el.classList.contains('on')?'':v); goPlan();
    };
  });
  [].forEach.call(document.querySelectorAll('.qcard'),function(el){
    el.onclick=function(){var v=el.getAttribute('data-q');PT.setIdx(null);PT.setQ(v);syncQ(v);goPlan();};
  });
  [].forEach.call(document.querySelectorAll('#scatter .pt'),function(c){
    c.onclick=function(){var v=c.getAttribute('data-q');PT.setIdx(null);PT.setQ(v);syncQ(v);goPlan();};
  });
  document.getElementById('pclear').onclick=function(){
    PT.clear();syncQ('');
    document.getElementById('pkw').value='';
    document.getElementById('pss').value='__all__';
  };
  var ss=document.getElementById('pss');
  PT.scenes().forEach(function(s){var o=document.createElement('option');o.value=s;o.textContent=s;ss.appendChild(o);});
  ss.onchange=function(){PT.setS(ss.value==='__all__'?'':ss.value);};
  document.getElementById('pkw').oninput=function(){PT.setKw(this.value);};
  document.getElementById('pcsv').onclick=function(){PT.csv();};
  [].forEach.call(document.querySelectorAll('#scatter .pt'),function(c){
    c.onclick=function(){PT.setQ(c.getAttribute('data-q'));
      document.getElementById('plan-table').scrollIntoView({behavior:'smooth',block:'start'});};
  });

  var IT=Table({data:IDATA,cols:__ITEMCOLS__,tid:'itb',pid:'ipg',cid:'icnt',sortKey:'c',csvName:__CSVNAME2__});
  IT.render();
  var is=document.getElementById('iss');
  IT.scenes().forEach(function(s){var o=document.createElement('option');o.value=s;o.textContent=s;is.appendChild(o);});
  is.onchange=function(){IT.setS(is.value==='__all__'?'':is.value);};
  document.getElementById('ikw').oninput=function(){IT.setKw(this.value);};
  document.getElementById('icsv').onclick=function(){IT.csv();};
  [].forEach.call(document.querySelectorAll('.tcard'),function(el){
    el.onclick=function(){var v=el.getAttribute('data-t');IT.setS(v);
      document.getElementById('item-table').scrollIntoView({behavior:'smooth',block:'start'});};
  });
});
"""


def kpi(v, label, unit=""):
    return '<div class="kpi"><div class="k">%s</div><div class="v">%s</div><div class="u">%s</div></div>' % (
        label, v, unit)


def render(ctx, out_path):
    plans = ctx["plans"]
    items = ctx["items"]
    st = ctx["quad_stats"]
    acct = ctx["acct"]
    roi_thr, sp_thr = ctx["roi_thr"], ctx["sp_thr"]

    h = []
    h.append('<!DOCTYPE html><html lang="zh-CN"><head><meta charset="utf-8">')
    h.append('<meta name="viewport" content="width=device-width,initial-scale=1">')
    h.append("<title>%s</title><style>%s</style></head><body><div class=\"wrap\">" % (ctx["title"], CSS))

    h.append("<h1>%s</h1>" % ctx["title"])
    h.append('<div class="sub">%s</div>' % ctx["sub"])

    # KPI
    h.append('<div class="kpis">')
    h.append(kpi("¥%s" % fmt_money(acct["c"]), "总花费", ""))
    h.append(kpi("¥%s" % fmt_money(acct["g"]), "总成交金额", ""))
    h.append(kpi("%.2f" % acct["roi"], "大盘投产比 ROI", "判定分界线"))
    h.append(kpi("%.1f%%" % (acct["sp"] * 100), "大盘潜客占比", "判定分界线"))
    h.append(kpi("%.1f%%" % (acct["nr"] * 100), "成交新客占比", ""))
    h.append(kpi("%d / %d" % (len(plans), len(items)), "计划数 / 商品数", ""))
    h.append(kpi("%s" % fmt_money(acct["u"]), "引导访问人数", ""))
    h.append(kpi("%s" % fmt_money(acct["b"]), "成交人数", ""))
    h.append(kpi("%.2f" % acct["cpc"], "平均点击花费", "CPC"))
    h.append(kpi("%s" % fmt_money(acct["o"]), "成交笔数", ""))
    h.append(kpi("%.2f%%" % (acct["ctr"] * 100), "点击率 CTR", ""))
    h.append(kpi("%.2f%%" % (acct["cvr"] * 100), "点击转化率 CVR", ""))
    h.append('<div class="kpi" style="background:#fff7ed;border-color:#fed7aa">'
             '<div class="k">全站推广 · 花费占比</div>'
             '<div class="v" style="color:#ea580c">%.1f%%</div>'
             '<div class="u">%d 个计划 · ¥%s</div></div>'
             % (ctx["qz_info"]["share"] * 100, ctx["qz_info"]["n"],
                fmt_money(ctx["qz_info"]["c"])))
    h.append('<div class="kpi" style="background:#fff7ed;border-color:#fed7aa">'
             '<div class="k">全站推广 · ROI</div>'
             '<div class="v" style="color:#ea580c">%.2f</div>'
             '<div class="u">成交 ¥%s ｜ 大盘 %.2f</div></div>'
             % (ctx["qz_info"]["roi"], fmt_money(ctx["qz_info"]["g"]), acct["roi"]))
    h.append("</div>")

    # 店铺花费 TOP N 商品（店铺整体一览：紧跟大盘，放在四象限分布之前）
    if ctx["top_items"]:
        h.append('<h2 id="top-items">店铺花费 TOP%d 商品 · 推广效果</h2>' % ctx["top_n"])
        h.append('<div class="card">')
        h.append('<div class="note" style="margin:0 0 10px">数据直接取自商品级报表，按花费降序；'
                 '「分层」按 ROI × 花费 划分：爆款=高ROI高花费、潜力=高ROI低花费、'
                 '问题=低ROI高花费、长尾=低ROI低花费。</div>')
        h.append('<div style="overflow:auto"><table><thead><tr>'
                 '<th>#</th><th>商品ID</th><th>商品名称</th><th>分层</th>'
                 '<th>花费</th><th>占店铺花费</th>'
                 '<th>成交金额</th><th>ROI</th><th>潜客占比</th><th>新客占比</th>'
                 '<th>展现</th><th>点击</th><th>CTR</th><th>CPC</th><th>CVR</th>'
                 '<th>成交笔数</th><th>加购</th><th>收藏</th>'
                 + ('<th>自然曝光</th><th>自然溢出</th>' if ctx["has_nat"] else '')
                 + '</tr></thead><tbody>')
        for t in ctx["top_items"]:
            nm = t["name"] if len(t["name"]) <= 26 else t["name"][:25] + "…"
            tc_ = TIERS.get(t["tier"], {"c": "#9ca3af"})["c"]
            rcls = "good" if t["tier"] in ("爆款", "潜力") else "bad"
            row = ('<tr>'
                   '<td>%d</td><td class="nm" style="color:#475569">%s</td>'
                   '<td class="nm" title="%s">%s</td>'
                   '<td><span class="qtag" style="background:%s">%s</span></td>'
                   '<td>¥%s</td><td>%.1f%%</td><td>¥%s</td>'
                   '<td class="%s"><b>%.2f</b></td><td>%.1f%%</td><td>%.1f%%</td>'
                   '<td>%s</td><td>%s</td><td>%.2f%%</td><td>%.2f</td><td>%.2f%%</td>'
                   '<td>%s</td><td>%s</td><td>%s</td>'
                   % (t["rank"], esc(t["id"] or "—"),
                      esc(t["name"]), esc(nm), tc_, t["tier"] or "—",
                      fmt_money(t["c"]), t["share"] * 100, fmt_money(t["g"]),
                      rcls, t["r"], t["p"] * 100, t["nr"] * 100,
                      fmt_money(t["im"]), fmt_money(t["ck"]),
                      t["ctr"] * 100, t["cpc"], t["cvr"] * 100,
                      fmt_money(t["o"]), fmt_money(t["ca"]), fmt_money(t["f"])))
            if ctx["has_nat"]:
                row += '<td>%s</td><td>%.1f%%</td>' % (
                    fmt_money(t["ni"]), t["nat_ratio"] * 100)
            row += '</tr>'
            h.append(row)
        h.append('</tbody></table></div>')
        h.append('<div class="note">花费占比 = 该商品花费 ÷ 商品级报表总花费；'
                 'ROI 标红表示该商品低于大盘分界线。</div>')
        if ctx["has_nat"]:
            h.append('<div class="note">自然曝光 = 报表「自然流量曝光量」；自然溢出 = 自然曝光 ÷ 付费展现量，'
                     '反映付费投放带来的自然流量外溢（品牌 / 搜索溢出）。该指标仅作参考，'
                     '<b>不参与四象限判定</b>。</div>')
        h.append("</div>")

    # 商品结构与分层（店铺整体一览）
    if items:
        h.append('<h2 id="item-table">商品结构与分层</h2>')
        h.append('<div class="qcards">')
        for t, d in ctx["tier_stats"].items():
            h.append('<div class="qcard tcard" data-t="%s" style="border-top-color:%s">'
                     '<div class="t"><span class="dot" style="background:%s"></span>%s</div>'
                     '<div class="c">%s</div>'
                     '<div class="m"><b>%d</b> 个商品 · 花费 <b>¥%s</b></div>'
                     '<div class="m">占预算 %.1f%% ｜ ROI %.2f</div>'
                     '<div class="a">%s</div></div>'
                     % (t, d["c_color"], d["c_color"], t, d["cond"], d["n"],
                        fmt_money(d["cost"]), d["cost_share"] * 100, d["roi"], d["act"]))
        h.append("</div>")
        h.append('<div class="card" style="margin-top:14px"><div class="bar">'
                 '<input id="ikw" placeholder="搜索商品名称 / 商品ID…" style="min-width:220px">'
                 '<select id="iss"><option value="__all__">全部分层</option></select>'
                 '<button class="btn" id="icsv">导出 CSV</button>'
                 '<span id="icnt" style="color:#64748b;font-size:12.5px"></span></div>'
                 '<div style="max-height:620px;overflow:auto"><table id="itb"></table></div>'
                 '<div class="pg" id="ipg"></div></div>')

    # 四象限卡片
    h.append("<h2>四象限分布（点击卡片可筛选下方明细）</h2>")
    h.append('<div class="qcards">')
    for k in (1, 2, 3, 4, 0):
        q = QUADS[k]
        v = st[k]
        cs = v["c"] / acct["c"] if acct["c"] else 0
        gs = v["g"] / acct["g"] if acct["g"] else 0
        h.append('<div class="qcard" data-q="%d" style="border-top-color:%s">'
                 '<div class="t"><span class="dot" style="background:%s"></span>%s</div>'
                 '<div class="c">%s</div>'
                 '<div class="m"><b>%d</b> 个 · 花费 <b>¥%s</b></div>'
                 '<div class="m">占预算 %.1f%% ｜ 占成交 %.1f%%</div>'
                 '<div class="m">ROI %.2f ｜ 潜客 %.0f%%</div>'
                 '<div class="a">%s</div></div>'
                 % (k, q["c"], q["c"], q["n"], q["cond"], v["n"], fmt_money(v["c"]),
                    cs * 100, gs * 100, v["roi"], v["sp"] * 100, q["act"]))
    h.append("</div>")
    be_txt = ""
    if ctx.get("be_applied"):
        be_txt = "（取 max(基准 %.2f, 保本 %.2f)）" % (ctx["base_thr"], ctx["be_roi"])
    h.append('<div class="note">判定口径：ROI 分界 %.2f%s｜潜客占比分界 %.1f%%（大盘）｜'
             '花费 &lt; %.0f 元或访问人数 &lt; %.0f 记为「样本不足」不参与判定。'
             '四象限仅依据 ROI×潜客占比，未混入自然流量曝光量。</div>'
             % (roi_thr, be_txt, sp_thr * 100, ctx["min_cost"], ctx["min_uv"]))

    # 散点
    h.append("<h2>计划四象限散点</h2>")
    lg = ['<div class="legend" id="sclegend">']
    for k in (1, 2, 3, 4, 0):
        lg.append('<div class="lg" data-q="%d"><span class="doti" style="background:%s"></span>'
                  '%s<span style="color:#94a3b8">%d</span></div>'
                  % (k, QUADS[k]["c"], QUADS[k]["n"], st[k]["n"]))
    lg.append('<span class="note" style="margin-left:4px;align-self:center">'
              '悬停看明细 · 点击筛选下方列表 · 点图例可隐藏某个象限</span>')
    lg.append("</div>")
    h.append('<div class="card">%s%s</div>' % (svg_scatter(plans, roi_thr, sp_thr), "".join(lg)))

    # 错配
    h.append("<h2>预算错配：花了多少、赚回多少</h2>")
    h.append('<div class="card">%s<div class="note">橙条明显高于蓝条 = 该象限在净消耗预算。'
             '理想状态是「优质拉新」的蓝条 ≥ 橙条。</div></div>'
             % svg_group_bars(st, acct["c"], acct["g"]))

    # 场景
    h.append("<h2>分场景结构（按象限拆分花费）</h2>")
    h.append('<div class="card">%s</div>' % svg_scene_stack(plans))

    # 优化测算
    h.append("<h2>预算调整测算</h2>")
    h.append('<div class="sim">%s</div>' % ctx["sim_html"])
    h.append('<div class="note">测算为静态推算：假设被停投计划的成交全部流失（未计入自然流量承接与'
             '预算转移带来的增量），实际结果通常好于此测算。仅用于判断调整方向，不作为承诺值。</div>')

    # 计划明细
    h.append('<h2 id="plan-table">计划明细与处置建议</h2>')
    h.append('<div class="card">')
    chips = ['<div class="chips" id="pchips">']
    chips.append('<div class="chip on" data-q=""><span class="n">全部</span>'
                 '<span class="c">%d 计划</span></div>' % len(plans))
    for k in (1, 2, 3, 4, 0):
        v = st[k]
        chips.append('<div class="chip" data-q="%d"><span class="doti" style="background:%s"></span>'
                     '<span class="n">%s</span><span class="c">%d 个 · ¥%s</span></div>'
                     % (k, QUADS[k]["c"], QUADS[k]["n"], v["n"], fmt_money(v["c"])))
    chips.append("</div>")
    h.append("".join(chips))
    h.append('<div class="bar">'
             '<input id="pkw" placeholder="搜索计划名称 / 计划ID…" style="min-width:240px">'
             '<select id="pss"><option value="__all__">全部场景</option></select>'
             '<button class="btn" id="pclear">清除筛选</button>'
             '<button class="btn" id="pcsv">导出 CSV</button>'
             '<span id="pcnt" style="color:#64748b;font-size:12.5px"></span></div>'
             '<div style="max-height:620px;overflow:auto"><table id="ptb"></table></div>'
             '<div class="pg" id="ppg"></div></div>')

    h.append('<div class="note" style="text-align:center;margin-top:36px">'
             '本文件为单文件离线报告，无外部依赖，可直接发送给同事；'
             '数据口径以万相台无界后台导出为准。</div>')
    h.append("</div>")

    js = (JS.replace("__QUADS__", json.dumps({k: {"n": v["n"], "c": v["c"]} for k, v in QUADS.items()},
                                             ensure_ascii=False))
            .replace("__TIERS__", json.dumps({k: {"c": v["c"]} for k, v in TIERS.items()},
                                             ensure_ascii=False))
            .replace("__PLANS__", json.dumps(ctx["plans_json"], ensure_ascii=False, separators=(",", ":")))
            .replace("__ITEMS__", json.dumps(ctx["items_json"], ensure_ascii=False, separators=(",", ":")))
            .replace("__PLANCOLS__", json.dumps(ctx["plan_cols"], ensure_ascii=False))
            .replace("__ITEMCOLS__", json.dumps(ctx["item_cols"], ensure_ascii=False))
            .replace("__CSVNAME1__", json.dumps(ctx["csv1"], ensure_ascii=False))
            .replace("__CSVNAME2__", json.dumps(ctx["csv2"], ensure_ascii=False)))
    h.append("<script>%s</script></body></html>" % js)

    html = "".join(h)
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(html)
    return out_path


# ---------------------------------------------------------------------------
def build(plans, items, args):
    tc = total_of(plans, "c")
    tg = total_of(plans, "g")
    tu = total_of(plans, "u")
    tq = total_of(plans, "q")
    tb = total_of(plans, "b")
    tck = total_of(plans, "ck")
    tim = total_of(plans, "im")

    roi_thr = args.roi_thr if args.roi_thr else (tg / tc if tc else 0)
    sp_thr = args.sp_thr if args.sp_thr is not None else (tq / tu if tu else 0)

    # 保本 ROI 上限（可选）：若给出店铺整体毛利率 margin(0-1)，则
    # 保本ROI = 1/margin，ROI 分界线取 max(保本ROI, 大盘ROI)。
    # 仅一个店铺级输入，不需要逐商品毛利率；不填则沿用大盘线（历史行为）。
    be_roi = 0.0
    be_applied = False
    base_thr = roi_thr
    margin = getattr(args, "margin", 0)
    if margin and 0 < margin < 1:
        be_roi = 1.0 / margin
        if be_roi > roi_thr:
            roi_thr = be_roi
        be_applied = True

    # 自然流量曝光量是否存在（部分导出版本没有该列）
    has_nat = any((r.get("ni") or 0) > 0 for r in plans) or \
        any((r.get("ni") or 0) > 0 for r in items)

    for r in plans:
        r["qd"] = classify(r, roi_thr, sp_thr, args.min_cost, args.min_uv)
    st = quadrant_stats(plans)

    # 货品全站推广占比（全站计划花费 / 全店总花费）
    qz = [r for r in plans if "全站" in (r["s"] or "")]
    qz_c = sum(r["c"] for r in qz)
    qz_g = sum(r["g"] for r in qz)
    qz_info = {
        "share": qz_c / tc if tc else 0,
        "c": qz_c, "g": qz_g,
        "roi": qz_g / qz_c if qz_c else 0,
        "n": len(qz),
    }


    acct = {
        "c": tc, "g": tg, "roi": tg / tc if tc else 0,
        "sp": tq / tu if tu else 0, "u": tu, "q": tq, "b": tb,
        "nr": (total_of(plans, "nc") / tb) if tb else 0,
        "o": total_of(plans, "o"),
        "cpc": tc / tck if tck else 0,
        "ctr": tck / tim if tim else 0,
        "cvr": total_of(plans, "o") / tck if tck else 0,
    }

    # 预算调整测算
    c4, g4 = st[4]["c"], st[4]["g"]
    c3, g3 = st[3]["c"], st[3]["g"]
    base = acct["roi"]

    def roi_after(dc, dg):
        return (tg - dg) / (tc - dc) if (tc - dc) > 0 else 0

    s1 = roi_after(c4, g4)
    s2 = roi_after(c4 + c3 * 0.5, g4 + g3 * 0.5)
    c1, g1 = st[1]["c"], st[1]["g"]
    move = (c4 + c3 * 0.5) * 0.7
    add_g = move * (g1 / c1 if c1 else 0) * 0.8
    s3 = (tg - g4 - g3 * 0.5 + add_g) / (tc - c4 - c3 * 0.5 + move) if (tc - c4 - c3 * 0.5 + move) > 0 else 0

    def sim(title, roi_new, gmv_new, cost_new, desc):
        d_roi = (roi_new / base - 1) * 100 if base else 0
        d_gmv = (gmv_new / tg - 1) * 100 if tg else 0
        cls = "good" if d_roi >= 0 else "bad"
        gcls = "good" if d_gmv >= 0 else "bad"
        return ('<div><div class="h">%s</div>'
                '<div class="b">%.2f <span class="%s" style="font-size:13px">(%+.1f%%)</span></div>'
                '<div class="d">成交 ¥%s <span class="%s">(%+.1f%%)</span> ｜ 花费 ¥%s</div>'
                '<div class="d" style="margin-top:4px">%s</div></div>'
                % (title, roi_new, cls, d_roi, fmt_money(gmv_new), gcls, d_gmv,
                   fmt_money(cost_new), desc))

    g1 = tg - g4
    g2 = tg - g4 - g3 * 0.5
    g3s = g2 + add_g
    c2 = tc - c4 - c3 * 0.5
    c3s = c2 + move

    sim_html = (
        sim("① 当前基准", base, tg, tc, "不做任何调整")
        + sim("② 停投「建议调整」", s1, g1, tc - c4,
              "停掉 %d 个计划、¥%s 花费" % (st[4]["n"], fmt_money(c4)))
        + sim("③ 停投 + 「人群矫正」减半", s2, g2, c2,
              "再把 %d 个「人群矫正」计划预算压一半" % st[3]["n"])
    )
    sim_html += (
        '<div style="grid-column:1/-1;background:#fffaf5;border:1px solid #ffe0c2;'
        'border-radius:9px;padding:12px 14px;font-size:12.5px;color:#7c3f00">'
        '<b>④ 综合情景（保规模）</b>：在 ③ 的基础上，把释放预算的 70%%（¥%s）转投「优质拉新」，'
        '按该象限当前 ROI 的 8 折估算增量。ROI 约 <b>%.2f</b>（%+.1f%%），'
        '成交 ¥%s（%+.1f%%）。<br>'
        '<b>怎么选</b>：③ 的 ROI 更高（%.2f）但成交规模收缩 %.1f%%；'
        '④ 用一点效率换回规模，适合还要冲量的阶段。'
        '若当前目标就是保投产，走 ③；若目标是在投产达标前提下最大化成交，走 ④。</div>'
        % (fmt_money(move), s3, (s3 / base - 1) * 100 if base else 0,
           fmt_money(g3s), (g3s / tg - 1) * 100 if tg else 0,
           s2, abs((g2 / tg - 1) * 100) if tg else 0))

    # 商品分层
    tier_stats = {}
    if items:
        ic = total_of(items, "c")
        cost_thr = args.cost_thr if args.cost_thr else (
            sorted(r["c"] for r in items)[int(len(items) * 0.7)] if items else 0)
        iroi_thr = args.roi_thr if args.roi_thr else (
            total_of(items, "g") / ic if ic else 0)
        for r in items:
            r["tier"] = tier_of(r, iroi_thr, cost_thr)
        for t in TIERS:
            sub = [r for r in items if r["tier"] == t]
            c = sum(x["c"] for x in sub)
            g = sum(x["g"] for x in sub)
            tier_stats[t] = {
                "n": len(sub), "cost": c, "roi": g / c if c else 0,
                "cost_share": c / ic if ic else 0,
                "c_color": TIERS[t]["c"], "act": TIERS[t]["act"],
                "cond": ("高ROI 高花费" if t == "爆款" else "高ROI 低花费" if t == "潜力"
                         else "低ROI 高花费" if t == "问题" else "低ROI 低花费"),
            }
        # 商品表的「场景」列复用分层（便于同一套筛选组件）
        for r in items:
            r["s"] = r["tier"]

    # 店铺花费 TOP N 商品（直接用商品级报表数据，不做计划关联）
    ic_all = total_of(items, "c")
    top_idx = sorted(range(len(items)), key=lambda i: -items[i]["c"])[:args.top_n]
    top_items = []
    for rank, i in enumerate(top_idx, 1):
        it = items[i]
        top_items.append({
            "rank": rank, "ix": i, "name": it["n"], "id": it.get("id", ""),
            "tier": it.get("tier", ""),
            "c": it["c"], "g": it["g"], "r": it["r"], "p": it["p"],
            "im": it["im"], "ck": it["ck"], "ctr": it["ctr"], "cpc": it["cpc"],
            "cvr": it["cvr"], "o": it["o"], "u": it["u"], "nr": it["nr"],
            "ca": it["ca"], "f": it["f"], "b": it["b"],
            "ni": it.get("ni", 0), "ng": it.get("ng", 0),
            "nat_ratio": it.get("nat_ratio", 0),
            "share": it["c"] / ic_all if ic_all else 0,
        })

    def slim(r, keys):
        return {k: (round(r[k], 4) if isinstance(r[k], float) else r[k]) for k in keys}

    pkeys = ["n", "s", "c", "g", "r", "im", "ck", "ctr", "cpc", "cvr", "o",
             "u", "q", "p", "nr", "b", "ca", "f", "ni", "nat_ratio", "ng", "qd", "id"]
    ikeys = ["n", "s", "tier", "id", "c", "g", "r", "im", "ck", "ctr", "cpc", "cvr", "o",
             "u", "q", "p", "nr", "b", "ca", "f", "ni", "nat_ratio", "ng"]

    plan_cols = [
        {"k": "qd", "t": "象限"}, {"k": "n", "t": "计划名称"}, {"k": "s", "t": "场景"},
        {"k": "c", "t": "花费", "m": "money"}, {"k": "g", "t": "成交金额", "m": "money"},
        {"k": "r", "t": "ROI", "m": "x"}, {"k": "p", "t": "潜客占比", "m": "pct"},
        {"k": "nr", "t": "新客占比", "m": "pct"}, {"k": "im", "t": "展现", "m": "num"},
        {"k": "ck", "t": "点击", "m": "num"}, {"k": "ctr", "t": "CTR", "m": "pct"},
        {"k": "cpc", "t": "CPC", "m": "x"}, {"k": "cvr", "t": "CVR", "m": "pct"},
        {"k": "o", "t": "成交笔数", "m": "num"}, {"k": "u", "t": "访问人数", "m": "num"},
        {"k": "q", "t": "潜客数", "m": "num"}, {"k": "ca", "t": "加购", "m": "num"},
        {"k": "f", "t": "收藏", "m": "num"},
    ]
    if has_nat:
        plan_cols += [
            {"k": "ni", "t": "自然曝光", "m": "num"},
            {"k": "nat_ratio", "t": "自然溢出", "m": "pct"},
        ]
    item_cols = [
        {"k": "tier", "t": "分层"}, {"k": "id", "t": "商品ID"}, {"k": "n", "t": "商品名称"},
        {"k": "c", "t": "花费", "m": "money"}, {"k": "g", "t": "成交金额", "m": "money"},
        {"k": "r", "t": "ROI", "m": "x"}, {"k": "p", "t": "潜客占比", "m": "pct"},
        {"k": "im", "t": "展现", "m": "num"}, {"k": "ck", "t": "点击", "m": "num"},
        {"k": "ctr", "t": "CTR", "m": "pct"}, {"k": "cpc", "t": "CPC", "m": "x"},
        {"k": "cvr", "t": "CVR", "m": "pct"}, {"k": "o", "t": "成交笔数", "m": "num"},
        {"k": "u", "t": "访问人数", "m": "num"}, {"k": "ca", "t": "加购", "m": "num"},
        {"k": "f", "t": "收藏", "m": "num"},
    ]
    if has_nat:
        item_cols += [
            {"k": "ni", "t": "自然曝光", "m": "num"},
            {"k": "nat_ratio", "t": "自然溢出", "m": "pct"},
        ]

    period = args.period or ""
    src = []
    if args.plans:
        src.append(os.path.basename(args.plans))
    if args.items:
        src.append(os.path.basename(args.items))

    pjson = []
    for i, r in enumerate(plans):
        d = slim(r, pkeys)
        d["ix"] = i
        pjson.append(d)

    return {
        "plans": plans, "items": items, "quad_stats": st, "acct": acct,
        "roi_thr": roi_thr, "sp_thr": sp_thr,
        "has_nat": has_nat, "be_roi": be_roi, "be_applied": be_applied,
        "base_thr": base_thr,
        "min_cost": args.min_cost, "min_uv": args.min_uv,
        "sim_html": sim_html, "tier_stats": tier_stats,
        "qz_info": qz_info, "top_items": top_items, "top_n": args.top_n,
        "sim_roi": (s1, s2, s3),
        "plans_json": pjson,
        "items_json": [slim(r, ikeys) for r in items],
        "plan_cols": plan_cols, "item_cols": item_cols,
        "title": "万相台无界 · 计划四象限周报" + (" · " + period if period else ""),
        "sub": "%s ｜ 数据源：%s ｜ 全量 %d 个计划%s ｜ 生成于 %s"
               % (period or "统计周期见导出文件", "、".join(src), len(plans),
                  "、%d 个商品" % len(items) if items else "",
                  datetime.now().strftime("%Y-%m-%d %H:%M")),
        "csv1": "计划四象限_%s.csv" % (period or "报表"),
        "csv2": "商品分层_%s.csv" % (period or "报表"),
    }


def safe_period(p):
    """周期文案要进文件名：清掉 Windows 非法字符"""
    return re.sub(r'[\\/:*?"<>|]', "-", (p or "").strip())


def main():
    ap = argparse.ArgumentParser(description="万相台无界四象限周报生成器")
    ap.add_argument("--plans", required=True, help="计划级报表文件")
    ap.add_argument("--items", default="", help="商品级报表文件（可选）")
    ap.add_argument("--period", default="", help="统计周期文案，如 09-21至09-27")
    ap.add_argument("--roi-thr", type=float, default=0, help="ROI 分界线，默认取大盘 ROI")
    ap.add_argument("--sp-thr", type=float, default=None, help="潜客占比分界线，默认取大盘值")
    ap.add_argument("--margin", type=float, default=0,
                    help="店铺整体毛利率(0-1)，用于计算保本ROI=1/毛利率；"
                         "ROI 分界线取 max(保本ROI, 大盘ROI)。不填则沿用大盘线。")
    ap.add_argument("--cost-thr", type=float, default=0, help="商品规模分界花费，默认取 70 分位")
    ap.add_argument("--min-cost", type=float, default=300.0, help="样本不足：花费下限，默认 300")
    ap.add_argument("--min-uv", type=float, default=200.0, help="样本不足：访问人数下限，默认 200")
    ap.add_argument("--top-n", type=int, default=10, help="店铺花费 TOP N 商品，默认 10")
    ap.add_argument("--summary-file", default="", help="把文本摘要另存为文件（如 摘要.md）")
    ap.add_argument("--output", default="", help="输出 HTML 路径")
    args = ap.parse_args()

    plans = load_rows(args.plans)
    items = load_rows(args.items) if args.items else []

    ctx = build(plans, items, args)
    out = args.output or os.path.join(
        os.path.dirname(os.path.abspath(args.plans)),
        "万相台四象限周报_%s.html" % (safe_period(args.period) or datetime.now().strftime("%m%d")))
    render(ctx, out)

    st = ctx["quad_stats"]
    a = ctx["acct"]
    L = []
    L.append("=" * 60)
    L.append("周报已生成：%s（%.0f KB）" % (out, os.path.getsize(out) / 1024))
    L.append("周期：%s ｜ 计划 %d 个 ｜ 商品 %d 个"
             % (args.period or "见导出文件", len(plans), len(items)))
    L.append("")
    L.append("【大盘】花费 %.0f ｜ 成交 %.0f ｜ ROI %.2f ｜ 潜客占比 %.1f%% ｜ 新客占比 %.1f%%"
             % (a["c"], a["g"], a["roi"], a["sp"] * 100, a["nr"] * 100))
    L.append("【判定线】ROI %.2f ｜ 潜客占比 %.1f%% ｜ 样本不足线 花费<%.0f 或 访问<%.0f"
             % (ctx["roi_thr"], ctx["sp_thr"] * 100, args.min_cost, args.min_uv))
    q = ctx["qz_info"]
    L.append("【全站推广】%d 个计划 ｜ 占花费 %.1f%% ｜ ROI %.2f（大盘 %.2f）"
             % (q["n"], q["share"] * 100, q["roi"], a["roi"]))
    L.append("")
    L.append("【四象限】")
    for k in (1, 2, 3, 4, 0):
        v = st[k]
        L.append("  %-6s %3d 个 · 花费 %8s（%4.1f%%）· 成交 %9s（%4.1f%%）· ROI %5.2f"
                 % (QUADS[k]["n"], v["n"], fmt_money(v["c"]),
                    v["c"] / a["c"] * 100 if a["c"] else 0,
                    fmt_money(v["g"]), v["g"] / a["g"] * 100 if a["g"] else 0, v["roi"]))
    L.append("")
    # 待处理计划：建议调整 + 人群矫正里花钱最多的
    bad = sorted([r for r in plans if r["qd"] in (4, 3)], key=lambda z: -z["c"])[:8]
    if bad:
        L.append("【优先处理计划】消耗最高的 8 个（象限=4建议调整 / 3人群矫正）")
        for r in bad:
            L.append("  [Q%d] %-32s 花费 %8s · ROI %5.2f · 潜客 %5.1f%% · %s"
                     % (r["qd"], r["n"][:32], fmt_money(r["c"]), r["r"],
                        r["p"] * 100, r["s"]))
        L.append("")
    if ctx["top_items"]:
        L.append("【商品花费 TOP5】")
        for t in ctx["top_items"][:5]:
            L.append("  %d. %-28s ID %s · 花费 %8s（%4.1f%%）· ROI %5.2f · %s"
                     % (t["rank"], t["name"][:28], t["id"] or "—", fmt_money(t["c"]),
                        t["share"] * 100, t["r"], t["tier"]))
        L.append("")
    L.append("【测算】基准 ROI %.2f ｜ 停投建议调整 → %.2f ｜ 再压降人群矫正一半 → %.2f"
             % (a["roi"], ctx["sim_roi"][0], ctx["sim_roi"][1]))
    L.append("=" * 60)

    text = "\n".join(L)
    print(text)
    if args.summary_file:
        with open(args.summary_file, "w", encoding="utf-8") as f:
            f.write(text + "\n")
        print("摘要已写入：%s" % args.summary_file)


if __name__ == "__main__":
    main()
