#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""生成一份模拟的「万相台无界」导出报表，用于验证分析器 / 对照字段格式。

真实使用时不需要这个文件，直接跑 analyze.py 分析你自己的导出文件即可。
"""

import os
import random
from datetime import date, timedelta

import pandas as pd

random.seed(42)

SCENES = ["全站推广", "关键词推广", "精准人群推广", "商品运营"]
# (计划名, 场景, ROI 基准)
PLANS = [
    ("全站推广-主推爆款A", "全站推广", 1.2),
    ("全站推广-新品B测款", "全站推广", 0.45),
    ("关键词推广-核心词包", "关键词推广", 3.6),
    ("关键词推广-长尾拓量", "关键词推广", 0.85),
    ("精准人群-老客召回", "精准人群推广", 4.2),
    ("精准人群-相似店铺", "精准人群推广", 0.55),
    ("商品运营-秋冬新款", "商品运营", 2.1),
    ("商品运营-清仓款", "商品运营", 0.05),
]
KEYWORDS = ["连衣裙 秋冬", "加绒卫衣", "直筒牛仔裤", "打底衫 女", "羊毛开衫",
            "半身裙 长款", "棉服 短款", "阔腿裤 加厚", "针织马甲", "风衣 中长款"]
UNITS = ["单元-高意向人群", "单元-泛人群拓量", "单元-店铺访客", "单元-竞品客群"]

rows = []
start = date.today() - timedelta(days=29)
for d in range(30):
    day = start + timedelta(days=d)
    # 全站整体效率随时间下滑，模拟"投产下滑"的真实场景
    drift = max(0.35, 1.0 - d * 0.022)
    for name, scene, roi_base in PLANS:
        cost = round(random.uniform(120, 900) * (roi_base ** 0.3), 2)
        roi = max(0.0, roi_base * drift * random.uniform(0.82, 1.2))
        ctr = random.uniform(0.006, 0.032)
        clk = int(cost / random.uniform(0.35, 1.4))
        imp = int(clk / ctr) if ctr else 0
        gmv = round(cost * roi, 2)
        # 清仓款这类低效计划几乎不成交，用于触发"有消耗无成交"诊断
        cvr = random.uniform(0.0005, 0.004) if roi_base < 0.3 else random.uniform(0.01, 0.06)
        order = max(0, int(clk * cvr))
        if order == 0:
            gmv = 0.0
        rows.append({
            "日期": day.strftime("%Y-%m-%d"),
            "推广场景": scene,
            "计划名称": name,
            "单元名称": random.choice(UNITS),
            "关键词": random.choice(KEYWORDS),
            "展现量": imp,
            "点击量": clk,
            "点击率": round(ctr, 4),
            "消耗": cost,
            "平均点击成本": round(cost / clk, 2) if clk else 0,
            "千次展现成本": round(cost / imp * 1000, 2) if imp else 0,
            "成交笔数": order,
            "成交金额": gmv,
            "点击转化率": round(order / clk, 4) if clk else 0,
            "投入产出比": round(gmv / cost, 2) if cost else 0,
            "加购量": int(clk * random.uniform(0.03, 0.12)),
            "收藏量": int(clk * random.uniform(0.01, 0.06)),
        })

df = pd.DataFrame(rows)
out = os.path.join(os.path.dirname(os.path.abspath(__file__)), "sample_wanxiangtai.xlsx")
with pd.ExcelWriter(out, engine="openpyxl") as w:
    df.to_excel(w, index=False, sheet_name="万相台无界报表")
print("已生成示例报表：%s（%d 行）" % (out, len(df)))
