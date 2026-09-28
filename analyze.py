#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
万相台无界推广分析器
==================
输入：万相台无界（阿里妈妈）后台导出的报表文件（.xlsx / .xls / .csv）
输出：一份自包含的 HTML 分析报告（离线可用，无需联网）

用法：
    python analyze.py 报表.xlsx
    python analyze.py 报表.xlsx --target-roi 3.0
    python analyze.py 报表.xlsx --output 我的报告.html

设计要点：
    1. 自动探测表头行（万相台导出文件通常前几行是标题/统计口径说明）
    2. 智能字段映射（兼容不同版本导出的列名差异，支持别名模糊匹配）
    3. 缺失指标自动用基础字段推导（CTR / PPC / CPM / CVR / ROI / 客单价）
    4. 规则化诊断引擎，输出可执行的预算调整建议
"""

import os
import re
import sys
import math
import argparse
from datetime import datetime

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

try:
    import pandas as pd
except ImportError:
    print("=" * 54)
    print("缺少 pandas 库，请先双击同目录的 setup.bat 完成一次安装。")
    print("（或命令行运行：python -m pip install pandas openpyxl）")
    print("装好后请通过 single.bat / batch.bat 启动，不要直接双击 .py 文件。")
    print("=" * 54)
    try:
        input("按回车退出…")
    except Exception:
        pass
    sys.exit(1)

# ---------------------------------------------------------------------------
# 字段别名表：标准字段 -> 可能出现的列名（按优先级排列）
# ---------------------------------------------------------------------------
FIELD_ALIASES = {
    "date":     ["日期", "统计日期", "时间", "账期", "日期区间", "day", "date"],
    "scene":    ["场景名字", "推广场景", "场景", "推广类型", "产品类型", "营销场景", "计划类型",
                 "adtype", "campaigntype", "producttype"],
    "plan":     ["计划名字", "计划名称", "推广计划", "计划名", "计划", "计划id名称",
                 "campaign", "planname", "name"],
    "item":     ["主体名称", "商品名称", "商品标题", "商品", "宝贝名称", "宝贝",
                 "item", "product", "title"],
    "unit":     ["单元名称", "推广单元", "单元名", "单元", "adgroup"],
    "creative": ["创意名称", "创意标题", "创意", "creative", "adcreative"],
    "keyword":  ["关键词", "搜索词", "关键词名称", "词", "人群名称", "人群",
                 "keyword", "searchterm", "query", "audience"],
    "cost":     ["消耗", "花费", "总消耗", "消耗金额", "花费金额", "支出", "cost", "spend",
                 "charge", "c"],
    "imp":      ["展现量", "曝光量", "展示量", "展现次数", "展现", "曝光",
                 "impressions", "impr", "views", "pv"],
    "clk":      ["点击量", "点击数", "点击次数", "点击", "clicks", "click", "ck"],
    "ctr":      ["点击率", "ctr"],
    "ppc":      ["平均点击成本", "平均点击花费", "点击均价", "单次点击成本", "cpc", "ppc"],
    "cpm":      ["千次展现花费", "千次展现成本", "千次曝光成本", "cpm"],
    # 注意：不要把「直接成交金额」列进来，它会抢在「总成交金额」前面被精确命中
    "gmv":      ["总成交金额", "成交金额", "总成交额", "支付金额", "成交额",
                 "gmv", "revenue", "sales", "saleamount", "conversionvalue", "turnover",
                 "amt", "g"],
    "order":    ["总成交笔数", "成交笔数", "成交订单数", "订单数", "支付订单数", "成交量", "订单量",
                 "orders", "conversions", "ordercount", "purchases"],
    "cvr":      ["转化率", "点击转化率", "成交转化率", "cvr", "conversionrate"],
    "roi":      ["投入产出比", "投产比", "roi", "roas", "回报率"],
    "cart":     ["总购物车数", "购物车数", "加购量", "加购数", "加购", "加入购物车", "addtocart"],
    "fav":      ["总收藏数", "收藏量", "收藏数", "收藏", "favorite", "col"],
    # —— 万相台无界特有：人群资产与成交结构 ——
    "cid":      ["计划id", "主体id", "商品id", "计划编号", "cid", "campaignid", "itemid"],
    "uv":       ["引导访问人数", "访问人数", "访客数", "访客", "uv", "uniqueshopper", "覆盖人数"],
    "qk":       ["引导访问潜客数", "潜客数", "潜客", "qk", "potential"],
    "sp":       ["引导访问潜客占比", "潜客占比", "潜客比例", "潜客率", "sp"],
    "newRate":  ["成交新客占比", "新客率", "新客占比", "新客比例", "newrate"],
    "newC":     ["成交新客数", "新客数", "新客", "newc", "newcustomer"],
    "byr":      ["买家数", "成交人数", "购买人数", "byr", "buyer"],
    "dirOrd":   ["直接成交笔数", "直接订单数", "直接成交", "dirord"],
}

# 表头探测关键字：命中即认为该行是列名行
HEADER_HINTS = ["消耗", "花费", "展现", "曝光", "点击", "成交金额", "计划", "日期", "花费金额"]

SCENE_TIPS = {
    "全站推广": "看整体投产与成交规模，调优动作以「目标ROI出价」和预算节奏为主，避免频繁改价导致模型重新学习。",
    "关键词推广": "看词效与匹配方式，重点处理高花费低产出词，加价高转化词，定期否词。",
    "精准人群推广": "看人群溢价与人群转化，弱人群降溢价或剔除，高转化人群提溢价扩量。",
    "商品运营": "看单品投放效率，低效商品及时替换，爆款加大预算承接。",
    "活动场景": "看活动期蓄水与爆发节奏，活动前加预算蓄客，活动后及时收量。",
}


# ---------------------------------------------------------------------------
# 工具函数
# ---------------------------------------------------------------------------
def normalize(s):
    """列名归一化：去空格、括号内容、百分号、货币符号，全角转半角，转小写"""
    s = str(s).strip()
    s = re.sub(r"[（(].*?[)）]", "", s)
    s = s.replace(" ", "").replace("　", "")
    s = s.replace("%", "").replace("¥", "").replace("￥", "").replace("$", "")
    s = s.replace("：", "").replace(":", "")
    out = []
    for ch in s:
        code = ord(ch)
        if 0xFF01 <= code <= 0xFF5E:
            out.append(chr(code - 0xFEE0))
        else:
            out.append(ch)
    return "".join(out).lower()


def to_num(v):
    """把 '1,234' / '12.5%' / '¥88' / '-' / '' 转成 float"""
    if v is None:
        return 0.0
    if isinstance(v, (int, float)):
        return 0.0 if pd.isna(v) else float(v)
    s = str(v).strip()
    if s in ("", "-", "--", "nan", "NaN", "None", "无", "/"):
        return 0.0
    s = s.replace(",", "").replace("¥", "").replace("￥", "").replace("$", "")
    pct = False
    if s.endswith("%"):
        pct = True
        s = s[:-1]
    try:
        n = float(s)
    except ValueError:
        return 0.0
    return n / 100.0 if pct else n


def safe_div(a, b):
    """安全的除法（Series）"""
    b = pd.Series(b).astype(float)
    a = pd.Series(a).astype(float)
    return (a / b.replace(0, pd.NA)).fillna(0.0)


def fmt_money(v):
    v = float(v or 0)
    if abs(v) >= 10000:
        return "%.2f万" % (v / 10000)
    return "%.0f" % v


def fmt_num(v):
    return "{:,.0f}".format(float(v or 0))


def fmt_pct(v, digits=2):
    return "{:.{d}f}%".format(float(v or 0) * 100, d=digits)


def fmt_x(v, digits=2):
    return "{:.{d}f}".format(float(v or 0), d=digits)


def esc(s):
    return (str(s).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
            .replace('"', "&quot;"))


# ---------------------------------------------------------------------------
# 读取与字段映射
# ---------------------------------------------------------------------------
def read_csv_safely(path):
    for enc in ("utf-8-sig", "gbk", "gb18030", "utf-8"):
        try:
            return pd.read_csv(path, header=None, encoding=enc, dtype=str), enc
        except UnicodeDecodeError:
            continue
    raise RuntimeError("无法识别 CSV 编码，请另存为 UTF-8 后再试")


def detect_header_row(path, is_csv):
    """探测真正的列名行"""
    try:
        if is_csv:
            raw, _ = read_csv_safely(path)
            raw = raw.head(15)
        else:
            raw = pd.read_excel(path, header=None, nrows=15, dtype=str)
    except Exception:
        return 0
    for i in range(len(raw)):
        joined = "".join(normalize(v) for v in raw.iloc[i].tolist() if str(v) != "nan")
        if sum(1 for h in HEADER_HINTS if h in joined) >= 2:
            return i
    return 0


def load_table(path):
    ext = os.path.splitext(path)[1].lower()
    is_csv = ext in (".csv", ".txt")
    hrow = detect_header_row(path, is_csv)
    if is_csv:
        _, enc = read_csv_safely(path)
        df = pd.read_csv(path, header=hrow, encoding=enc)
    else:
        df = pd.read_excel(path, header=hrow)
    df = df.dropna(how="all")
    df.columns = [str(c).strip() for c in df.columns]
    # 去掉全是 Unnamed 的辅助列
    df = df.loc[:, [not str(c).startswith("Unnamed") for c in df.columns]]
    return df


def build_mapping(df):
    cols = list(df.columns)
    ncols = {c: normalize(c) for c in cols}
    mapping, used = {}, set()

    for std, aliases in FIELD_ALIASES.items():
        # 1) 精确匹配
        for c in cols:
            if c in used:
                continue
            if ncols[c] in aliases:
                mapping[std] = c
                used.add(c)
                break
        if std in mapping:
            continue
        # 2) 包含匹配
        for c in cols:
            if c in used:
                continue
            n = ncols[c]
            if not n:
                continue
            for a in aliases:
                # 短别名（c / g / uv / 宝贝…）只走精确匹配：
                # 否则 c 会吞掉 ctr、ord 会被当成 keyword(key-ord) 的子串、宝贝会吞掉收藏宝贝数
                if not a or len(a) < 3:
                    continue
                # 反向包含（列名是别名的子串）只在列名较长时启用
                if a in n or (len(n) >= 4 and n in a):
                    mapping[std] = c
                    used.add(c)
                    break
            if std in mapping:
                break
    return mapping


def prepare(df, mapping):
    """抽出标准列并补算派生指标"""
    out = pd.DataFrame(index=df.index)

    def pick(std, numeric=True):
        if std in mapping:
            s = df[mapping[std]]
            return s.map(to_num) if numeric else s.astype(str)
        return None

    for std in ("scene", "plan", "unit", "creative", "keyword"):
        s = pick(std, numeric=False)
        if s is not None:
            out[std] = s

    if "date" in mapping:
        out["date"] = pd.to_datetime(df[mapping["date"]], errors="coerce")

    for std in ("cost", "imp", "clk", "gmv", "order", "cart", "fav"):
        s = pick(std)
        out[std] = s if s is not None else 0.0

    # 已有的比率字段：缺失或全为 0 时，标记为需要用基础字段推导
    need_derive = set()
    for std in ("ctr", "ppc", "cpm", "cvr", "roi"):
        s = pick(std)
        if s is None or float(pd.Series(s).fillna(0).abs().sum()) == 0:
            need_derive.add(std)
            out[std] = 0.0
        else:
            out[std] = s

    # 派生
    if "ctr" in need_derive:
        out["ctr"] = safe_div(out["clk"], out["imp"])
    if "ppc" in need_derive:
        out["ppc"] = safe_div(out["cost"], out["clk"])
    if "cpm" in need_derive:
        out["cpm"] = safe_div(out["cost"], out["imp"]) * 1000
    if "cvr" in need_derive:
        out["cvr"] = safe_div(out["order"], out["clk"])
    if "roi" in need_derive:
        out["roi"] = safe_div(out["gmv"], out["cost"])
    out["aov"] = safe_div(out["gmv"], out["order"])

    out = out.dropna(subset=["cost"])
    out = out[out["cost"].fillna(0).astype(float) >= 0]
    return out


# ---------------------------------------------------------------------------
# 汇总分析
# ---------------------------------------------------------------------------
def agg_group(df, key):
    """按维度聚合，保留比率重算"""
    if key not in df.columns:
        return None
    g = df.groupby(key, dropna=True).agg(
        cost=("cost", "sum"), imp=("imp", "sum"), clk=("clk", "sum"),
        gmv=("gmv", "sum"), order=("order", "sum"),
        cart=("cart", "sum") if "cart" in df.columns else ("cost", "sum"),
        fav=("fav", "sum") if "fav" in df.columns else ("cost", "sum"),
    ).reset_index()
    g["ctr"] = safe_div(g["clk"], g["imp"])
    g["ppc"] = safe_div(g["cost"], g["clk"])
    g["cvr"] = safe_div(g["order"], g["clk"])
    g["roi"] = safe_div(g["gmv"], g["cost"])
    g["aov"] = safe_div(g["gmv"], g["order"])
    return g.sort_values("cost", ascending=False)


def analyze(df, target_roi):
    res = {"target_roi": target_roi, "nrows": len(df)}

    total_cost = float(df["cost"].sum())
    total_gmv = float(df["gmv"].sum())
    total_imp = float(df["imp"].sum())
    total_clk = float(df["clk"].sum())
    total_order = float(df["order"].sum())

    res["overview"] = {
        "cost": total_cost,
        "gmv": total_gmv,
        "imp": total_imp,
        "clk": total_clk,
        "order": total_order,
        "ctr": total_clk / total_imp if total_imp else 0,
        "ppc": total_cost / total_clk if total_clk else 0,
        "cpm": total_cost / total_imp * 1000 if total_imp else 0,
        "cvr": total_order / total_clk if total_clk else 0,
        "roi": total_gmv / total_cost if total_cost else 0,
        "aov": total_gmv / total_order if total_order else 0,
        "profit_gap": total_cost * target_roi - total_gmv,
    }

    # 日趋势
    if "date" in df.columns and df["date"].notna().any():
        d = df.dropna(subset=["date"]).copy()
        d["d"] = d["date"].dt.strftime("%m-%d")
        daily = d.groupby("d").agg(
            cost=("cost", "sum"), imp=("imp", "sum"), clk=("clk", "sum"),
            gmv=("gmv", "sum"), order=("order", "sum")).reset_index().sort_values("d")
        daily["ctr"] = safe_div(daily["clk"], daily["imp"])
        daily["roi"] = safe_div(daily["gmv"], daily["cost"])
        daily["cvr"] = safe_div(daily["order"], daily["clk"])
        res["daily"] = daily
    else:
        res["daily"] = None

    for key in ("scene", "plan", "unit", "creative", "keyword"):
        res[key] = agg_group(df, key)

    res["diagnostics"] = diagnose(res, df, target_roi)
    return res


def diagnose(res, df, target_roi):
    """规则化诊断引擎"""
    items = []
    ov = res["overview"]
    total_cost = ov["cost"] or 1.0

    def add(level, title, evidence, action):
        items.append({"level": level, "title": title, "evidence": evidence, "action": action})

    # 1. 整体投产
    if ov["roi"] > 0:
        if ov["roi"] < target_roi:
            add("high",
                "整体投产未达标",
                "整体 ROI %.2f，低于目标 %.2f，消耗 %s 元，成交 %s 元" %
                (ov["roi"], target_roi, fmt_money(ov["cost"]), fmt_money(ov["gmv"])),
                "距离目标还差 %s 元成交额。按当前效率，需把 ROI 从 %.2f 提到 %.2f（+%.0f%%）；"
                "若短期提不动效率，建议先砍掉低效计划把消耗压下来，保住整体投产。"
                % (fmt_money(max(0, ov["profit_gap"])), ov["roi"], target_roi,
                   (target_roi / ov["roi"] - 1) * 100 if ov["roi"] else 0))
        else:
            add("low", "整体投产达标",
                "整体 ROI %.2f，高于目标 %.2f" % (ov["roi"], target_roi),
                "可考虑在高效计划上加预算放大，同时监控 ROI 是否在放量后下滑。")

    # 2. 计划维度：止损 / 加投
    plans = res.get("plan")
    if plans is not None and len(plans) > 0:
        thr = max(50.0, total_cost * 0.01)
        bad = plans[(plans["cost"] >= thr) & (plans["roi"] < target_roi * 0.8)]
        good = plans[(plans["cost"] >= thr) & (plans["roi"] >= target_roi * 1.5)]
        zero = plans[(plans["cost"] >= thr) & (plans["gmv"] <= 0)]

        if len(zero):
            add("high", "有消耗无成交的计划",
                "%d 个计划花了钱却零成交，合计消耗 %s 元（占比 %.1f%%）：%s" %
                (len(zero), fmt_money(zero["cost"].sum()),
                 zero["cost"].sum() / total_cost * 100,
                 "、".join(zero["plan"].astype(str).head(5).tolist())),
                "这类计划优先暂停或大幅降价，属于确定性亏损。")

        if len(bad):
            bad_cost = float(bad["cost"].sum())
            add("high" if bad_cost / total_cost > 0.25 else "mid", "低效计划吃掉了大量预算",
                "%d 个计划 ROI 低于 %.2f（目标的 80%%），合计消耗 %s 元，占总消耗 %.1f%%，"
                "这些计划整体 ROI 仅 %.2f" %
                (len(bad), target_roi * 0.8, fmt_money(bad_cost),
                 bad_cost / total_cost * 100, bad["gmv"].sum() / bad_cost if bad_cost else 0),
                "逐个降预算或降价；若连续 3 天无改善直接暂停，把预算挪给高效计划。")

        if len(good):
            add("low", "高效计划可以加投",
                "%d 个计划 ROI 高于 %.2f：%s" %
                (len(good), target_roi * 1.5,
                 "、".join(good["plan"].astype(str).head(5).tolist())),
                "分 2-3 次小步加预算（每次 +20%~30%），加完观察 1-2 天 ROI 是否守住。")

        # 集中度
        top1 = plans.iloc[0]
        share = float(top1["cost"]) / total_cost
        if share > 0.5 and len(plans) > 1:
            add("mid", "预算过度集中",
                "计划「%s」占了 %.1f%% 的消耗" % (top1["plan"], share * 100),
                "单计划占比过高，一旦模型波动整体投产会剧烈震荡，建议拆分成 2-3 个计划分散风险。")
    else:
        # 没有计划列时，用关键词/单元兜底
        alt = res.get("keyword") if res.get("keyword") is not None else res.get("unit")
        if alt is not None and len(alt):
            thr = max(50.0, total_cost * 0.01)
            bad = alt[(alt["cost"] >= thr) & (alt["roi"] < target_roi * 0.8)]
            if len(bad):
                add("mid", "低效投放对象占比偏高",
                    "%d 个对象 ROI 低于 %.2f，合计消耗 %s 元（%.1f%%）" %
                    (len(bad), target_roi * 0.8, fmt_money(bad["cost"].sum()),
                     bad["cost"].sum() / total_cost * 100),
                    "优先降价或暂停，把预算集中到 ROI 达标的对象上。")

    # 3. 点击率异常
    if ov["imp"] > 0:
        avg_ctr = ov["ctr"]
        if avg_ctr > 0 and avg_ctr < 0.01:
            add("mid", "整体点击率偏低",
                "整体 CTR %.2f%%，低于常见水平" % (avg_ctr * 100),
                "优先换创意图/标题、检查人群与商品匹配度；点击率上不去，出价再高也是在买无效曝光。")

    # 4. 高点击低转化
    plans2 = plans if plans is not None else None
    if plans2 is not None and len(plans2) >= 3 and ov["ctr"] > 0 and ov["cvr"] > 0:
        cand = plans2[(plans2["clk"] >= 30) & (plans2["ctr"] > ov["ctr"] * 1.2)
                      & (plans2["cvr"] < ov["cvr"] * 0.7)]
        if len(cand):
            add("mid", "有人点、没人买",
                "%d 个计划点击率高于均值 20%% 以上，但转化率低于均值 30%% 以上：%s" %
                (len(cand), "、".join(cand["plan"].astype(str).head(4).tolist())),
                "流量不精准或落地承接有问题：检查详情页/价格/评价/库存状态，同时收紧人群与关键词匹配。")

    # 5. 点击成本
    if ov["ppc"] > 0 and plans2 is not None and len(plans2) >= 3:
        avg_ppc = ov["ppc"]
        exp = plans2[(plans2["cost"] >= max(50.0, total_cost * 0.01))
                     & (plans2["ppc"] > avg_ppc * 1.5) & (plans2["roi"] < target_roi)]
        if len(exp):
            add("mid", "点击成本偏高的计划",
                "%d 个计划 CPC 高于均值 50%% 且 ROI 未达标：%s" %
                (len(exp), "、".join(exp["plan"].astype(str).head(4).tolist())),
                "降低出价/人群溢价，或收窄投放时段与地域，先把点击成本压到均值附近再看转化。")

    # 6. 趋势
    daily = res.get("daily")
    if daily is not None and len(daily) >= 8:
        recent = daily.tail(7)
        prev = daily.iloc[-14:-7] if len(daily) >= 14 else daily.head(max(1, len(daily) - 7))
        r_cost, p_cost = float(recent["cost"].sum()), float(prev["cost"].sum())
        r_gmv, p_gmv = float(recent["gmv"].sum()), float(prev["gmv"].sum())
        r_roi = r_gmv / r_cost if r_cost else 0
        p_roi = p_gmv / p_cost if p_cost else 0
        if p_roi > 0:
            delta = (r_roi / p_roi - 1) * 100
            if delta <= -20:
                add("high", "投产在下滑",
                    "近 7 天 ROI %.2f，前一期 %.2f，下滑 %.0f%%" % (r_roi, p_roi, abs(delta)),
                    "排查三点：是否近期加了预算/改了出价导致模型重学、竞品是否降价、主图详情或评价是否变差。")
            elif delta >= 20:
                add("low", "投产在改善",
                    "近 7 天 ROI %.2f，前一期 %.2f，提升 %.0f%%" % (r_roi, p_roi, delta),
                    "趋势向好，可小步加预算放大，同时记录当前配置便于复盘。")
        if p_cost > 0 and r_cost > p_cost * 1.2 and r_gmv <= p_gmv * 1.05:
            add("mid", "花费涨了、成交没跟上",
                "近 7 天消耗 %.0f 元（前期 %.0f 元，+%.0f%%），成交仅 %.0f 元（前期 %.0f 元）"
                % (r_cost, p_cost, (r_cost / p_cost - 1) * 100, r_gmv, p_gmv),
                "加预算后边际效率在下降，把增量预算撤回到原水平，观察 ROI 是否恢复。")

    # 7. 场景层面提示
    scenes = res.get("scene")
    if scenes is not None and len(scenes):
        worst = scenes.sort_values("roi").iloc[0]
        if worst["roi"] < target_roi:
            tip = SCENE_TIPS.get(str(worst["scene"]), "")
            add("mid", "场景「%s」投产最低" % worst["scene"],
                "消耗 %s 元，ROI %.2f（整体 %.2f）" %
                (fmt_money(worst["cost"]), worst["roi"], ov["roi"]),
                (tip + " 当前该场景 ROI 低于目标，优先在这里动手。") if tip else
                "该场景 ROI 低于目标，优先在这里做减法。")

    return items


# ---------------------------------------------------------------------------
# SVG 图表（零依赖、离线可用）
# ---------------------------------------------------------------------------
C_ORANGE = "#ff5000"
C_BLUE = "#2f6fed"
C_GREEN = "#16a34a"
C_RED = "#dc2626"
C_GREY = "#94a3b8"


def svg_combo(daily, target_roi):
    """柱=消耗/成交（左轴），线=ROI（右轴）"""
    if daily is None or len(daily) < 2:
        return ""
    n = len(daily)
    W, H = 960, 340
    ml, mr, mt, mb = 70, 70, 30, 46
    pw, ph = W - ml - mr, H - mt - mb

    costs = daily["cost"].astype(float).tolist()
    gmv = daily["gmv"].astype(float).tolist()
    rois = daily["roi"].astype(float).tolist()
    labels = daily["d"].astype(str).tolist()

    vmax = max(max(costs), max(gmv)) * 1.15 or 1
    rmax = max(max(rois), target_roi) * 1.25 or 1

    def y_money(v):
        return mt + ph - (v / vmax) * ph

    def y_roi(v):
        return mt + ph - (v / rmax) * ph

    slot = pw / n
    bw = min(18.0, slot * 0.34)
    parts = []
    parts.append('<svg viewBox="0 0 %d %d" width="100%%" style="max-width:960px">' % (W, H))
    parts.append('<rect x="0" y="0" width="%d" height="%d" fill="#ffffff"/>' % (W, H))

    # 网格 + 左轴刻度
    for i in range(5):
        v = vmax * i / 4
        y = y_money(v)
        parts.append('<line x1="%.1f" y1="%.1f" x2="%.1f" y2="%.1f" stroke="#eef1f5" stroke-width="1"/>'
                     % (ml, y, ml + pw, y))
        parts.append('<text x="%.1f" y="%.1f" font-size="11" fill="#94a3b8" text-anchor="end">%s</text>'
                     % (ml - 8, y + 4, fmt_money(v)))
    # 右轴刻度
    for i in range(5):
        v = rmax * i / 4
        y = y_roi(v)
        parts.append('<text x="%.1f" y="%.1f" font-size="11" fill="#c0c8d2" text-anchor="start">%.1f</text>'
                     % (ml + pw + 8, y + 4, v))

    # 柱
    for i in range(n):
        cx = ml + slot * i + slot / 2
        x1 = cx - bw - 1.5
        x2 = cx + 1.5
        y1 = y_money(costs[i])
        y2 = y_money(gmv[i])
        base = mt + ph
        parts.append('<rect x="%.1f" y="%.1f" width="%.1f" height="%.1f" fill="%s" rx="2"><title>日期 %s 消耗 %.0f</title></rect>'
                     % (x1, y1, bw, base - y1, C_ORANGE, labels[i], costs[i]))
        parts.append('<rect x="%.1f" y="%.1f" width="%.1f" height="%.1f" fill="%s" rx="2"><title>日期 %s 成交 %.0f</title></rect>'
                     % (x2, y2, bw, base - y2, C_BLUE, labels[i], gmv[i]))

    # ROI 折线
    pts = []
    for i in range(n):
        cx = ml + slot * i + slot / 2
        pts.append("%.1f,%.1f" % (cx, y_roi(rois[i])))
    parts.append('<polyline points="%s" fill="none" stroke="%s" stroke-width="2.2" stroke-linejoin="round"/>'
                 % (" ".join(pts), C_GREEN))
    # 目标线
    ty = y_roi(target_roi)
    parts.append('<line x1="%.1f" y1="%.1f" x2="%.1f" y2="%.1f" stroke="%s" stroke-width="1.4" stroke-dasharray="5 4"/>'
                 % (ml, ty, ml + pw, ty, C_RED))
    parts.append('<text x="%.1f" y="%.1f" font-size="11" fill="%s">目标 ROI %.1f</text>'
                 % (ml + 4, ty - 5, C_RED, target_roi))

    # x 轴标签
    step = max(1, math.ceil(n / 14))
    for i in range(0, n, step):
        cx = ml + slot * i + slot / 2
        parts.append('<text x="%.1f" y="%.1f" font-size="11" fill="#94a3b8" text-anchor="middle">%s</text>'
                     % (cx, mt + ph + 18, labels[i]))

    # 图例
    lg = [("消耗", C_ORANGE), ("成交金额", C_BLUE), ("ROI(右轴)", C_GREEN)]
    x = ml
    for name, color in lg:
        parts.append('<rect x="%.1f" y="%.1f" width="10" height="10" rx="2" fill="%s"/>' % (x, 12, color))
        parts.append('<text x="%.1f" y="%.1f" font-size="12" fill="#475569">%s</text>' % (x + 15, 21, name))
        x += len(name) * 13 + 55

    parts.append("</svg>")
    return "".join(parts)


def svg_bars(items, label_key, value_key, note_key, target_roi, topn=10,
             value_label="消耗", note_fmt="roi"):
    """横向条形图，条内标注 ROI"""
    if items is None or len(items) == 0:
        return ""
    d = items.head(topn)
    n = len(d)
    W = 960
    rowh, mt, mb = 34, 16, 16
    H = mt + mb + rowh * n + 30
    ml = 210
    mr = 90
    pw = W - ml - mr
    vmax = float(d[value_key].max()) or 1

    parts = ['<svg viewBox="0 0 %d %d" width="100%%" style="max-width:960px">' % (W, H)]
    parts.append('<rect x="0" y="0" width="%d" height="%d" fill="#ffffff"/>' % (W, H))
    for i in range(n):
        row = d.iloc[i]
        y = mt + i * rowh + 6
        v = float(row[value_key])
        w = (v / vmax) * pw
        roi = float(row["roi"]) if "roi" in row else 0
        color = C_GREEN if roi >= target_roi else (C_ORANGE if roi >= target_roi * 0.8 else C_RED)
        name = str(row[label_key])
        if len(name) > 16:
            name = name[:15] + "…"
        parts.append('<text x="%d" y="%.1f" font-size="12.5" fill="#334155" text-anchor="end">%s</text>'
                     % (ml - 10, y + 15, esc(name)))
        parts.append('<rect x="%d" y="%.1f" width="%.1f" height="17" rx="3" fill="%s" opacity="0.9"/>'
                     % (ml, y + 3, max(w, 2), color))
        note = ("ROI %.2f" % roi) if note_fmt == "roi" else ("%.0f" % v)
        parts.append('<text x="%.1f" y="%.1f" font-size="11.5" fill="#64748b">%s · %s</text>'
                     % (ml + max(w, 2) + 8, y + 16, fmt_money(v), note))
    parts.append("</svg>")
    return "".join(parts)


def svg_bubble(items, label_key, target_roi, topn=35):
    """x=消耗 y=ROI 的气泡图，气泡大小=成交金额"""
    if items is None or len(items) == 0:
        return ""
    d = items.head(topn)
    W, H = 960, 400
    ml, mr, mt, mb = 70, 30, 30, 46
    pw, ph = W - ml - mr, H - mt - mb
    xmax = float(d["cost"].max()) or 1
    ymax = max(float(d["roi"].max()), target_roi) * 1.2 or 1
    gmax = float(d["gmv"].max()) or 1

    parts = ['<svg viewBox="0 0 %d %d" width="100%%" style="max-width:960px">' % (W, H)]
    parts.append('<rect x="0" y="0" width="%d" height="%d" fill="#ffffff"/>' % (W, H))

    for i in range(5):
        y = mt + ph - ph * i / 4
        parts.append('<line x1="%d" y1="%.1f" x2="%.1f" y2="%.1f" stroke="#eef1f5"/>' % (ml, y, ml + pw, y))
        parts.append('<text x="%d" y="%.1f" font-size="11" fill="#94a3b8" text-anchor="end">%.1f</text>'
                     % (ml - 8, y + 4, ymax * i / 4))
        parts.append('<text x="%.1f" y="%.1f" font-size="11" fill="#94a3b8" text-anchor="middle">%s</text>'
                     % (ml + pw * i / 4, mt + ph + 18, fmt_money(xmax * i / 4)))

    ty = mt + ph - (target_roi / ymax) * ph
    parts.append('<line x1="%d" y1="%.1f" x2="%.1f" y2="%.1f" stroke="%s" stroke-dasharray="5 4" stroke-width="1.4"/>'
                 % (ml, ty, ml + pw, ty, C_RED))
    parts.append('<text x="%.1f" y="%.1f" font-size="11" fill="%s">目标 %.1f</text>' % (ml + 5, ty - 6, C_RED, target_roi))
    qy = mt + ph - ((target_roi * 0.8) / ymax) * ph
    parts.append('<line x1="%d" y1="%.1f" x2="%.1f" y2="%.1f" stroke="#fecaca" stroke-dasharray="3 3"/>'
                 % (ml, qy, ml + pw, qy))

    for i in range(len(d)):
        row = d.iloc[i]
        cx = ml + (float(row["cost"]) / xmax) * pw
        cy = mt + ph - (min(float(row["roi"]), ymax) / ymax) * ph
        r = 5 + 16 * math.sqrt(float(row["gmv"]) / gmax) if gmax else 6
        roi = float(row["roi"])
        color = C_GREEN if roi >= target_roi else (C_ORANGE if roi >= target_roi * 0.8 else C_RED)
        name = str(row[label_key])
        if len(name) > 12:
            name = name[:11] + "…"
        parts.append('<circle cx="%.1f" cy="%.1f" r="%.1f" fill="%s" opacity="0.45" stroke="%s" stroke-width="1.2">'
                     '<title>%s\n消耗 %.0f\n成交 %.0f\nROI %.2f</title></circle>'
                     % (cx, cy, r, color, color, esc(str(row[label_key])),
                        float(row["cost"]), float(row["gmv"]), roi))
        if r > 12:
            parts.append('<text x="%.1f" y="%.1f" font-size="10.5" fill="#334155" text-anchor="middle">%s</text>'
                         % (cx, cy + 3.5, esc(name)))
    parts.append('<text x="%d" y="%d" font-size="11" fill="#94a3b8">横轴：消耗 · 纵轴：ROI · 气泡大小：成交金额</text>'
                 % (ml, H - 8))
    parts.append("</svg>")
    return "".join(parts)


# ---------------------------------------------------------------------------
# HTML 渲染
# ---------------------------------------------------------------------------
CSS = """
*{box-sizing:border-box}
body{margin:0;background:#f5f6f8;color:#1f2937;font-family:-apple-system,BlinkMacSystemFont,"PingFang SC","Microsoft YaHei","Helvetica Neue",Arial,sans-serif;line-height:1.6}
.wrap{max-width:1080px;margin:0 auto;padding:32px 20px 64px}
h1{font-size:26px;margin:0 0 6px}
h2{font-size:19px;margin:38px 0 14px;padding-left:11px;border-left:4px solid #ff5000}
.sub{color:#64748b;font-size:13px;margin-bottom:24px}
.card{background:#fff;border:1px solid #e8ebef;border-radius:12px;padding:18px 20px;margin-bottom:16px;box-shadow:0 1px 3px rgba(16,24,40,.04)}
.kpis{display:grid;grid-template-columns:repeat(5,1fr);gap:12px}
.kpi{background:#fff;border:1px solid #e8ebef;border-radius:12px;padding:14px 16px}
.kpi .k{font-size:12px;color:#64748b}
.kpi .v{font-size:22px;font-weight:600;margin-top:4px;letter-spacing:-.3px}
.kpi .u{font-size:12px;color:#94a3b8;margin-top:2px}
.grid2{display:grid;grid-template-columns:1fr 1fr;gap:16px}
table{width:100%;border-collapse:collapse;font-size:13px}
th{background:#f8fafc;color:#475569;font-weight:600;text-align:right;padding:9px 10px;border-bottom:1px solid #e8ebef;white-space:nowrap}
th:first-child,td:first-child{text-align:left}
td{padding:9px 10px;border-bottom:1px solid #f1f3f6;text-align:right;white-space:nowrap}
td.name{text-align:left;max-width:280px;overflow:hidden;text-overflow:ellipsis}
tr:hover td{background:#fafbfc}
.bad{color:#dc2626}.good{color:#16a34a}.warn{color:#ea580c}
.diag{border-left:4px solid #94a3b8;padding:12px 16px;margin-bottom:12px;background:#fff;border-radius:0 10px 10px 0;border-top:1px solid #eef1f5;border-right:1px solid #eef1f5;border-bottom:1px solid #eef1f5}
.diag.high{border-left-color:#dc2626}
.diag.mid{border-left-color:#f59e0b}
.diag.low{border-left-color:#16a34a}
.diag .t{font-weight:600;font-size:15px;margin-bottom:3px}
.diag .e{font-size:13px;color:#475569}
.diag .a{font-size:13px;color:#0f766e;margin-top:6px;background:#f0fdfa;padding:8px 10px;border-radius:6px}
.tag{display:inline-block;font-size:11px;padding:2px 8px;border-radius:20px;margin-right:6px;vertical-align:2px}
.tag.high{background:#fee2e2;color:#b91c1c}
.tag.mid{background:#fef3c7;color:#b45309}
.tag.low{background:#dcfce7;color:#15803d}
.note{font-size:12.5px;color:#94a3b8;margin-top:10px}
.empty{color:#94a3b8;font-size:13px;padding:12px 0}
@media(max-width:860px){.kpis{grid-template-columns:repeat(2,1fr)}.grid2{grid-template-columns:1fr}}
"""


def kpi(v, label, unit=""):
    return '<div class="kpi"><div class="k">%s</div><div class="v">%s</div><div class="u">%s</div></div>' % (
        label, v, unit)


def roi_cls(v, target):
    return "good" if v >= target else ("bad" if v < target * 0.8 else "warn")


def render_table(g, name_key, target, limit=15, extra=None):
    if g is None or len(g) == 0:
        return '<div class="empty">该维度数据缺失</div>'
    d = g.head(limit)
    rows = []
    cols = [("消耗", "cost", "money"), ("展现", "imp", "num"), ("点击", "clk", "num"),
            ("CTR", "ctr", "pct"), ("CPC", "ppc", "x"), ("成交笔数", "order", "num"),
            ("成交金额", "gmv", "money"), ("CVR", "cvr", "pct"), ("ROI", "roi", "x")]
    head = "".join("<th>%s</th>" % c[0] for c in cols)
    if extra:
        head = "<th>%s</th>" % extra + head
    for _, r in d.iterrows():
        tds = ['<td class="name">%s</td>' % esc(str(r[name_key]))]
        for _, key, kind in cols:
            v = float(r[key])
            if kind == "money":
                s = fmt_money(v)
            elif kind == "num":
                s = fmt_num(v)
            elif kind == "pct":
                s = fmt_pct(v)
            else:
                s = fmt_x(v)
            cls = ' class="%s"' % roi_cls(v, target) if key == "roi" else ""
            tds.append("<td%s>%s</td>" % (cls, s))
        rows.append("<tr>%s</tr>" % "".join(tds))
    return '<table><thead><tr>%s</tr></thead><tbody>%s</tbody></table>' % (head, "".join(rows))


def render(res, source_name, out_path):
    ov = res["overview"]
    target = res["target_roi"]
    now = datetime.now().strftime("%Y-%m-%d %H:%M")
    plans = res.get("plan")
    scenes = res.get("scene")
    kws = res.get("keyword")

    roi_color = roi_cls(ov["roi"], target)
    h = []
    h.append("<!DOCTYPE html><html lang=\"zh-CN\"><head><meta charset=\"utf-8\">")
    h.append("<meta name=\"viewport\" content=\"width=device-width,initial-scale=1\">")
    h.append("<title>万相台无界推广分析报告</title><style>%s</style></head><body><div class=\"wrap\">" % CSS)
    h.append("<h1>万相台无界 · 推广分析报告</h1>")
    h.append('<div class="sub">数据源：%s ｜ 目标 ROI：%.2f ｜ 生成时间：%s ｜ 样本行数：%d</div>'
             % (esc(source_name), target, now, res["nrows"]))

    # KPI
    h.append('<div class="kpis">')
    h.append(kpi(fmt_money(ov["cost"]), "总消耗", "元"))
    h.append(kpi(fmt_money(ov["gmv"]), "总成交金额", "元"))
    h.append('<div class="kpi"><div class="k">投产比 ROI</div><div class="v %s">%.2f</div><div class="u">目标 %.2f</div></div>'
             % (roi_color, ov["roi"], target))
    h.append(kpi(fmt_pct(ov["ctr"]), "点击率 CTR", ""))
    h.append(kpi(fmt_pct(ov["cvr"]), "转化率 CVR", ""))
    h.append(kpi("¥%.2f" % ov["ppc"], "平均点击成本", "CPC"))
    h.append(kpi(fmt_num(ov["clk"]), "总点击", "次"))
    h.append(kpi(fmt_num(ov["order"]), "成交笔数", "笔"))
    h.append(kpi("¥%.2f" % ov["aov"], "客单价", ""))
    h.append(kpi("¥%.2f" % ov["cpm"], "千次展现成本", "CPM"))
    h.append("</div>")

    # 诊断
    h.append("<h2>核心诊断</h2>")
    if res["diagnostics"]:
        for d in res["diagnostics"]:
            lv = {"high": "紧急", "mid": "关注", "low": "机会"}[d["level"]]
            h.append('<div class="diag %s"><div class="t"><span class="tag %s">%s</span>%s</div>'
                     '<div class="e">%s</div><div class="a">建议：%s</div></div>'
                     % (d["level"], d["level"], lv, esc(d["title"]),
                        esc(d["evidence"]), esc(d["action"])))
    else:
        h.append('<div class="card"><div class="empty">数据维度不足，无法生成诊断（至少需要消耗与成交金额字段）</div></div>')

    # 趋势
    h.append("<h2>消耗 / 成交 / ROI 趋势</h2>")
    h.append('<div class="card">%s</div>' % (svg_combo(res["daily"], target)
                                             or '<div class="empty">未识别到日期字段，无法绘制趋势</div>'))

    # 场景
    h.append("<h2>推广场景表现</h2>")
    h.append('<div class="card">%s</div>' % render_table(scenes, "scene", target, limit=12))
    if scenes is not None and len(scenes):
        tips = []
        for _, r in scenes.iterrows():
            tip = SCENE_TIPS.get(str(r["scene"]))
            if tip:
                tips.append("<b>%s</b>：%s" % (esc(str(r["scene"])), tip))
        if tips:
            h.append('<div class="card"><div class="note">%s</div></div>' % "<br>".join(tips))

    # 计划
    h.append("<h2>计划消耗 TOP（按消耗排序，条色=ROI 达标情况）</h2>")
    h.append('<div class="card">%s</div>' % (svg_bars(plans, "plan", "cost", "roi", target)
                                             or '<div class="empty">未识别到计划字段</div>'))
    h.append("<h2>计划明细</h2>")
    h.append('<div class="card">%s</div>' % render_table(plans, "plan", target, limit=20))

    # 低效计划
    if plans is not None and len(plans):
        thr = max(50.0, ov["cost"] * 0.01)
        bad = plans[(plans["cost"] >= thr) & (plans["roi"] < target * 0.8)]
        if len(bad):
            h.append("<h2>优先止损名单（ROI 低于目标 80%）</h2>")
            h.append('<div class="card">%s<div class="note">阈值：消耗 ≥ %.0f 元。建议先降价 20%%~30%% 观察 2 天，无改善则暂停。</div></div>'
                     % (render_table(bad, "plan", target, limit=20), thr))
        good = plans[(plans["cost"] >= thr) & (plans["roi"] >= target * 1.5)]
        if len(good):
            h.append("<h2>可加投名单（ROI 高于目标 150%）</h2>")
            h.append('<div class="card">%s<div class="note">每次加预算 20%%~30%%，加完观察 1-2 天，确认 ROI 守住再继续加。</div></div>'
                     % render_table(good, "plan", target, limit=20))

    # 气泡
    h.append("<h2>消耗 vs ROI 分布（气泡=成交金额）</h2>")
    h.append('<div class="card">%s<div class="note">左上=花得少赚得多（可加投）；右下=花得多赚得少（要止损）。</div></div>'
             % (svg_bubble(plans if plans is not None else kws,
                           "plan" if plans is not None else "keyword", target)
                or '<div class="empty">维度数据不足</div>'))

    # 关键词 / 人群
    if kws is not None and len(kws):
        h.append("<h2>关键词 / 人群效率 TOP20</h2>")
        h.append('<div class="card">%s</div>' % render_table(kws, "keyword", target, limit=20))
        zero = kws[(kws["cost"] >= max(20.0, ov["cost"] * 0.005)) & (kws["gmv"] <= 0)]
        if len(zero):
            h.append('<div class="card"><b>零成交高消耗词/人群 %d 个</b>，合计消耗 %s 元：'
                     '<span class="bad">%s</span><div class="note">建议直接降价或否词/剔除人群。</div></div>'
                     % (len(zero), fmt_money(zero["cost"].sum()),
                        esc("、".join(zero["keyword"].astype(str).head(10).tolist()))))

    # 单元 / 创意
    for key, title in (("unit", "单元"), ("creative", "创意")):
        g = res.get(key)
        if g is not None and len(g):
            h.append("<h2>%s维度 TOP15</h2>" % title)
            h.append('<div class="card">%s</div>' % render_table(g, key, target, limit=15))

    h.append('<div class="note" style="text-align:center;margin-top:40px">'
             '本报告由本地脚本生成，数据不上传；指标口径以万相台后台导出为准。</div>')
    h.append("</div></body></html>")

    html = "".join(h)
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(html)
    return out_path


# ---------------------------------------------------------------------------
def main():
    ap = argparse.ArgumentParser(description="万相台无界推广分析器")
    ap.add_argument("file", help="万相台无界导出的报表文件（xlsx/xls/csv）")
    ap.add_argument("--target-roi", type=float, default=2.0, help="目标投产比，默认 2.0")
    ap.add_argument("--output", default="", help="输出 HTML 路径，默认与输入同目录")
    args = ap.parse_args()

    path = args.file
    if not os.path.exists(path):
        print("文件不存在：%s" % path)
        sys.exit(1)

    df = load_table(path)
    mapping = build_mapping(df)
    if "cost" not in mapping:
        print("未能在文件中识别到「消耗/花费」列。识别到的列：")
        for c in df.columns:
            print("  -", c)
        sys.exit(2)

    data = prepare(df, mapping)
    res = analyze(data, args.target_roi)

    out = args.output or (os.path.splitext(path)[0] + "_分析报告.html")
    render(res, os.path.basename(path), out)

    ov = res["overview"]
    print("=" * 52)
    print("分析完成：%s" % out)
    print("识别字段：%s" % "、".join("%s→%s" % (k, v) for k, v in mapping.items()))
    print("总消耗 %.0f 元 ｜ 成交 %.0f 元 ｜ ROI %.2f ｜ CTR %s ｜ CVR %s"
          % (ov["cost"], ov["gmv"], ov["roi"], fmt_pct(ov["ctr"]), fmt_pct(ov["cvr"])))
    print("诊断条目 %d 条" % len(res["diagnostics"]))
    print("=" * 52)


if __name__ == "__main__":
    main()
