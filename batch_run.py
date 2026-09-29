#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""批量入口（中文交互）：由 batch.bat 调用，负责所有中文提示与输入，
再把参数交给 batch_report.py 真正生成报告。

也可以直接命令行跑：
    python batch_run.py D:/报表目录
    python batch_run.py 计划报表.csv 商品报表.csv
"""

import os
import sys
import importlib.util

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

HERE = os.path.dirname(os.path.abspath(__file__))
EXTS = (".csv", ".xlsx", ".xls")

VERSION = "0.1.1"

_spec = importlib.util.spec_from_file_location("br", os.path.join(HERE, "batch_report.py"))
br = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(br)


def ask(prompt=""):
    """读一行输入；非交互环境（EOF）直接退出，避免卡死"""
    try:
        return input(prompt)
    except EOFError:
        print("\n（未检测到交互式输入，已退出）")
        raise SystemExit(0)


def wait_exit():
    try:
        input("\n按回车关闭窗口…")
    except EOFError:
        pass


def clean(p):
    return (p or "").strip().strip('"').strip("'")


def collect_paths(preset):
    """收集报表来源：拖进来的参数优先，否则逐条输入（可无限加）"""
    paths = [clean(p) for p in preset if clean(p)]
    if paths:
        print("\n已检测到 %d 个来源：" % len(paths))
        for i, p in enumerate(paths, 1):
            print("  [%d] %s" % (i, p))
        return paths

    print("\n请把「报表文件夹」或「计划报表 + 商品报表」拖到本窗口后回车，")
    print("也可以直接粘贴完整路径（带空格不用加引号）。")
    print("多个来源请一条一条输入，全部输完后直接回车结束。")
    print("提示：一个店铺通常需要「计划报表」+「商品报表」两份。")
    idx = 1
    while True:
        v = clean(ask("  来源%d > " % idx))
        if not v:
            break
        if not os.path.exists(v):
            print("    路径不存在，已忽略：%s" % v)
            continue
        if os.path.isfile(v) and not v.lower().endswith(EXTS):
            print("    不是报表文件（.csv/.xlsx/.xls），已忽略")
            continue
        paths.append(v)
        idx += 1
    return paths


def pick_period():
    print("\n统计周期怎么写进报告标题？例如：2026-09-21至09-27")
    print("（日报就写单日，如 2026-09-28；直接回车则留空）")
    return clean(ask("  周期 > "))


def pick_shop():
    print("\n店铺名称写进报告标题？（直接回车 = 按文件名自动识别，如「九阳旗舰店_计划_…」→ 九阳旗舰店）")
    print("多个店铺混在一起时请留空，脚本会逐个自动识别。")
    return clean(ask("  店铺名称 > "))


def pick_tag(period):
    print("\n报告类型：[1] 周报  [2] 日报（直接回车按周期自动判断）")
    s = clean(ask("  选择 > "))
    if s == "1":
        return "周报"
    if s == "2":
        return "日报"
    return "日报" if period and "至" not in period else "周报"


def main():
    preset = [a for a in sys.argv[1:] if not a.startswith("--")]

    print("=" * 58)
    print("  万相台无界 · 多店铺周报 / 日报 批量生成工具  v%s" % VERSION)
    print("=" * 58)
    print("  用法：把报表「文件夹」拖到 batch.bat 图标上；")
    print("        或同时选中「计划报表 + 商品报表」两个文件一起拖上去。")
    print("        也支持一次拖多个文件夹（一个文件夹 = 一个店铺）。")

    paths = collect_paths(preset)
    if not paths:
        print("\n没有提供任何报表来源，已取消。")
        wait_exit()
        return 1

    period = pick_period()
    shop = pick_shop()
    tag = pick_tag(period)

    # 输出目录：第一个来源是文件夹就用它，否则用文件所在目录
    first = paths[0]
    base = first if os.path.isdir(first) else os.path.dirname(os.path.abspath(first))
    out_dir = os.path.join(base, "reports")

    print("\n" + "-" * 58)
    print("来源 %d 个 ｜ 周期「%s」｜ 类型 %s%s"
          % (len(paths), period or "(未填)", tag,
             " ｜ 店铺「%s」" % shop if shop else ""))
    print("输出目录：%s" % out_dir)
    print("-" * 58)
    print("\n正在生成报告，请稍候…\n")

    old_argv = sys.argv
    sys.argv = ["batch_report.py", "--paths"] + paths + \
               ["--period", period, "--outdir", out_dir, "--suffix", tag]
    if shop:
        sys.argv += ["--shop", shop]
    try:
        rc = br.main()
    except SystemExit as e:
        rc = e.code or 0
    except Exception as e:
        print("\n生成失败：%s" % e)
        rc = 1
    finally:
        sys.argv = old_argv

    print("\n" + "=" * 58)
    if rc == 0:
        print("全部完成。报告与摘要都在：\n  %s" % out_dir)
        s = ""
        try:
            s = input("\n是否打开该目录？(直接回车打开，输入 n 跳过) ").strip().lower()
        except EOFError:
            s = "n"
        if s != "n":
            try:
                os.startfile(out_dir)
            except Exception:
                pass
    else:
        print("未全部成功，请查看上面的报错信息。")
    print("=" * 58)
    wait_exit()
    return rc


if __name__ == "__main__":
    sys.exit(main())
