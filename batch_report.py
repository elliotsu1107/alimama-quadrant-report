#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
批量生成多个店铺的万相台无界周报 / 日报

用法：
    python batch_report.py --dir D:/报表 --period "2026-09-21至09-27" --outdir outputs

文件识别规则（不要求固定命名，尽量宽松）：
    文件名含「计划」→ 计划级报表
    文件名含「商品」→ 商品级报表
    店铺名 = 文件名去掉日期戳与「计划/商品」等关键词后的剩余部分；
            若没有店铺名，则按文件顺序命名为 店铺1、店铺2…
    同店铺的计划与商品报表按排序顺序一一配对。

产出：每个店铺一份 HTML 报告 + 一份文本摘要，最后打印总览。
"""

import os
import re
import sys
import argparse
from datetime import datetime

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

import importlib.util

HERE = os.path.dirname(os.path.abspath(__file__))
_spec = importlib.util.spec_from_file_location("wr", os.path.join(HERE, "weekly_report.py"))
wr = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(wr)

EXTS = (".csv", ".xlsx", ".xls")


def safe_period(p):
    """周期文案要进文件名：清掉 Windows 非法字符（/ \\ : * ? " < > |）"""
    p = re.sub(r'[\\/:*?"<>|]', "-", (p or "").strip())
    return p or datetime.now().strftime("%m%d")


def guess_shop(filename):
    """从文件名猜店铺名：去扩展名 → 去日期戳 → 去「计划/商品报表」等关键词"""
    base = os.path.splitext(os.path.basename(filename))[0]
    base = re.sub(r"[_\-]?\d{8}([_\-]\d{6})?", "", base)      # 20260928_145046
    base = re.sub(r"[_\-]?\d{4}[_\-]\d{1,2}[_\-]\d{1,2}.*$", "", base)  # 2026-09-28
    for kw in ("计划报表", "商品报表", "计划级", "商品级", "计划", "商品", "推广"):
        base = base.replace(kw, "")
    base = re.sub(r"[_\-\s]+", " ", base).strip()
    return base


def main():
    ap = argparse.ArgumentParser(description="批量生成万相台无界周报/日报")
    ap.add_argument("--dir", default="", help="报表文件所在目录")
    ap.add_argument("--paths", nargs="*", default=[],
                    help="直接给文件或目录（可多个，例如 --paths 计划.csv 商品.csv）")
    ap.add_argument("--period", default="", help="统计周期文案")
    ap.add_argument("--outdir", default="", help="输出目录，默认为 --dir 下的 reports/")
    ap.add_argument("--roi-thr", type=float, default=0, help="ROI 分界线，默认取大盘")
    ap.add_argument("--sp-thr", type=float, default=None, help="潜客占比分界线，默认取大盘")
    ap.add_argument("--min-cost", type=float, default=300.0)
    ap.add_argument("--min-uv", type=float, default=200.0)
    ap.add_argument("--top-n", type=int, default=10)
    ap.add_argument("--suffix", default="", help="报告文件名后缀，如 周报 / 日报")
    args = ap.parse_args()

    # 收集报表：--paths 支持文件/目录混用（拖多个文件或整个文件夹都行），--dir 兼容旧用法
    sources = list(args.paths) if args.paths else ([args.dir] if args.dir else [])
    if not sources:
        print("请提供报表来源：--dir 报表目录 或 --paths 文件1 文件2 ...")
        return 1

    files = []
    for s in sources:
        if os.path.isdir(s):
            files += [os.path.join(s, f) for f in sorted(os.listdir(s))
                      if f.lower().endswith(EXTS) and not f.startswith("~$")]
        elif os.path.isfile(s) and s.lower().endswith(EXTS):
            files.append(s)
        else:
            print("跳过（不是目录也不是报表文件）：%s" % s)
    if not files:
        print("没有找到任何报表文件（.csv / .xlsx / .xls）。")
        return 1

    base = sources[0] if os.path.isdir(sources[0]) else os.path.dirname(os.path.abspath(sources[0]))
    out_dir = args.outdir or os.path.join(base, "reports")
    os.makedirs(out_dir, exist_ok=True)
    plans, items = {}, {}
    for full in files:
        f = os.path.basename(full)
        shop = guess_shop(f)
        if "计划" in f:
            plans.setdefault(shop, []).append(full)
        elif "商品" in f:
            items.setdefault(shop, []).append(full)

    if not plans:
        print("没有找到文件名含「计划」的报表文件。找到的文件：")
        for f in files:
            print("  -", os.path.basename(f))
        return 1

    shops = sorted(set(plans) | set(items))
    # 没店铺名的按顺序编号
    renamed = {}
    for i, s in enumerate(shops, 1):
        renamed[s] = s if s else "店铺%d" % i

    print("=" * 60)
    print("批量生成 ｜ 共 %d 个店铺 ｜ 输出：%s" % (len(shops), out_dir))
    print("=" * 60)

    tag = args.suffix or ("日报" if args.period and "至" not in args.period else "周报")
    ok, fail = 0, []
    summaries = []

    for s in shops:
        name = renamed[s]
        pl = sorted(plans.get(s, []))
        it = sorted(items.get(s, []))
        for idx, pf in enumerate(pl):
            itf = it[idx] if idx < len(it) else ""
            shop_label = name if len(pl) == 1 else "%s-%d" % (name, idx + 1)
            tag_p = safe_period(args.period)
            out_html = os.path.join(out_dir, "%s_无界%s_%s.html" % (shop_label, tag, tag_p))
            out_txt = os.path.join(out_dir, "%s_摘要_%s.txt" % (shop_label, tag_p))
            sub = argparse.Namespace(
                roi_thr=args.roi_thr, sp_thr=args.sp_thr, cost_thr=0,
                min_cost=args.min_cost, min_uv=args.min_uv,
                period=args.period, top_n=args.top_n,
                plans=pf, items=itf,
            )
            try:
                p_rows = wr.load_rows(pf)
                i_rows = wr.load_rows(itf) if itf else []
                ctx = wr.build(p_rows, i_rows, sub)
                ctx["title"] = "%s · 万相台无界%s" % (shop_label, tag)
                wr.render(ctx, out_html)
                a = ctx["acct"]
                with open(out_txt, "w", encoding="utf-8") as f:
                    f.write("%s %s\n" % (shop_label, args.period))
                    f.write("花费 %.0f ｜ 成交 %.0f ｜ ROI %.2f ｜ 潜客占比 %.1f%%\n"
                            % (a["c"], a["g"], a["roi"], a["sp"] * 100))
                    f.write("全站推广占花费 %.1f%% ｜ ROI %.2f\n"
                            % (ctx["qz_info"]["share"] * 100, ctx["qz_info"]["roi"]))
                    for k in (1, 2, 3, 4, 0):
                        v = ctx["quad_stats"][k]
                        f.write("  %s %d个 花费%s(%.1f%%) ROI %.2f\n"
                                % (wr.QUADS[k]["n"], v["n"], wr.fmt_money(v["c"]),
                                   v["c"] / a["c"] * 100 if a["c"] else 0, v["roi"]))
                    f.write("测算 ROI：基准 %.2f → 关停建议删除 %.2f → 再压降人群矫正 %.2f\n"
                            % (a["roi"], ctx["sim_roi"][0], ctx["sim_roi"][1]))
                ok += 1
                summaries.append((shop_label, a, ctx))
                print("  [OK] %-16s 花费 %9s ｜ ROI %5.2f ｜ 计划 %3d ｜ %s"
                      % (shop_label, wr.fmt_money(a["c"]), a["roi"], len(p_rows),
                         os.path.basename(out_html)))
            except Exception as e:
                fail.append((shop_label, str(e)))
                print("  [失败] %-16s %s" % (shop_label, e))

    print("=" * 60)
    print("完成 %d 份，失败 %d 份 ｜ 输出目录：%s" % (ok, len(fail), out_dir))
    if summaries:
        print("\n各店铺一览（按花费降序）：")
        for label, a, ctx in sorted(summaries, key=lambda x: -x[1]["c"]):
            q4 = ctx["quad_stats"][4]
            print("  %-16s 花费 %9s ｜ ROI %5.2f ｜ 待删 %2d个(占花费%4.1f%%) ｜ 全站占比 %4.1f%%"
                  % (label, wr.fmt_money(a["c"]), a["roi"], q4["n"],
                     q4["c"] / a["c"] * 100 if a["c"] else 0,
                     ctx["qz_info"]["share"] * 100))
    if fail:
        print("\n失败明细：")
        for label, err in fail:
            print("  %s：%s" % (label, err))
    print("=" * 60)
    return 0


if __name__ == "__main__":
    sys.exit(main())
