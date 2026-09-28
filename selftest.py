#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""鲁棒性自检：用不同列名风格 / 不同文件格式跑通分析器，确认不会崩。

运行：python selftest.py
"""

import os
import sys
import random
import importlib.util
from datetime import date, timedelta

import pandas as pd

sys.stdout.reconfigure(encoding="utf-8")
HERE = os.path.dirname(os.path.abspath(__file__))
spec = importlib.util.spec_from_file_location("az", os.path.join(HERE, "analyze.py"))
az = importlib.util.module_from_spec(spec)
spec.loader.exec_module(az)

random.seed(7)
DAYS = 20


def base_data():
    rows = []
    start = date.today() - timedelta(days=DAYS)
    for d in range(DAYS):
        day = start + timedelta(days=d)
        for p, base in [("爆款计划", 3.2), ("长尾计划", 1.1), ("测款计划", 0.4)]:
            cost = round(random.uniform(100, 700), 2)
            roi = base * random.uniform(0.8, 1.2)
            clk = int(cost / random.uniform(0.4, 1.2))
            ctr = random.uniform(0.008, 0.03)
            order = max(0, int(clk * random.uniform(0.01, 0.05)))
            rows.append([day.strftime("%Y-%m-%d"), "关键词推广", p,
                         int(clk / ctr), clk, round(ctr, 4), cost,
                         order, round(cost * roi, 2)])
    return rows


COLS_CN = ["统计日期", "推广类型", "推广计划", "展现", "点击", "点击率", "花费", "支付订单数", "支付金额"]
COLS_EN = ["Date", "Ad Type", "Campaign", "Impressions", "Clicks", "CTR", "Cost", "Orders", "Revenue"]


def run(name, path):
    df = az.load_table(path)
    mapping = az.build_mapping(df)
    assert "cost" in mapping, "%s 未识别到消耗列：%s" % (name, list(df.columns))
    data = az.prepare(df, mapping)
    res = az.analyze(data, 2.0)
    out = os.path.join(HERE, "_selftest_%s.html" % name)
    az.render(res, name, out)
    ov = res["overview"]
    print("  [OK] %-14s 字段 %2d 个 | 消耗 %.0f | 成交 %.0f | ROI %.2f | 诊断 %d 条"
          % (name, len(mapping), ov["cost"], ov["gmv"], ov["roi"], len(res["diagnostics"])))
    return out


def main():
    print("万相台无界分析器 · 鲁棒性自检")
    outs = []

    # 1. 中文列名（无 scene/unit/keyword，且比率列缺失，靠派生）
    df = pd.DataFrame(base_data(), columns=COLS_CN)
    p = os.path.join(HERE, "_t_cn.xlsx")
    df.to_excel(p, index=False)
    outs.append(run("中文列名", p))

    # 2. 英文列名
    df = pd.DataFrame(base_data(), columns=COLS_EN)
    p = os.path.join(HERE, "_t_en.csv")
    df.to_csv(p, index=False, encoding="utf-8-sig")
    outs.append(run("英文列名CSV", p))

    # 3. 带前置说明行的 xlsx（模拟万相台导出：前两行是标题与统计口径）
    import openpyxl
    p = os.path.join(HERE, "_t_prefix.xlsx")
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["万相台无界 · 推广报表"])
    ws.append(["统计周期：近 20 天"])
    ws.append(COLS_CN)
    for r in base_data():
        ws.append(r)
    wb.save(p)
    outs.append(run("带标题行", p))

    # 4. 带千分位/百分号/货币符号的脏数据
    df = pd.DataFrame(base_data(), columns=COLS_CN)
    df["花费"] = df["花费"].map(lambda v: "¥{:,.2f}".format(v))
    df["点击率"] = df["点击率"].map(lambda v: "{:.2%}".format(v))
    p = os.path.join(HERE, "_t_dirty.xlsx")
    df.to_excel(p, index=False)
    outs.append(run("脏数据格式", p))

    # 5. 单日快照（无趋势，daily 维度退化）
    df = pd.DataFrame(base_data(), columns=COLS_CN)
    df["统计日期"] = date.today().strftime("%Y-%m-%d")
    p = os.path.join(HERE, "_t_single.xlsx")
    df.to_excel(p, index=False)
    outs.append(run("单日快照", p))

    print("全部通过。样例输出：%s" % outs[0])


if __name__ == "__main__":
    main()
