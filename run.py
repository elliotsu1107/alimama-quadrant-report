#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""交互式入口：双击运行即可，按提示选文件、填目标 ROI，自动生成并打开报告。

也可以直接带参数跑，跳过选择步骤：
    python run.py 报表.xlsx
    python run.py 报表.xlsx --target-roi 3.0
    python run.py 报表.xlsx --no-open      # 生成但不自动打开浏览器
"""

import os
import sys
import glob
import webbrowser
import importlib.util

sys.stdout.reconfigure(encoding="utf-8")

HERE = os.path.dirname(os.path.abspath(__file__))
EXTS = (".xlsx", ".xls", ".csv")

_spec = importlib.util.spec_from_file_location("az", os.path.join(HERE, "analyze.py"))
az = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(az)


def ask(prompt=""):
    """读取一行输入；非交互环境（EOF）下直接退出，避免卡死"""
    try:
        return input(prompt)
    except EOFError:
        print("\n（未检测到交互式输入，已退出）")
        raise SystemExit(0)


def wait_exit():
    try:
        input("\n按回车退出…")
    except EOFError:
        pass


def candidates():
    """扫描当前目录与脚本目录里的报表文件"""
    dirs = []
    for d in (os.getcwd(), HERE):
        ad = os.path.abspath(d)
        if ad not in dirs:
            dirs.append(ad)
    found, seen = [], set()
    for d in dirs:
        for ext in EXTS:
            for f in glob.glob(os.path.join(d, "*" + ext)):
                af = os.path.abspath(f)
                if af not in seen and os.path.basename(f).lower() != "sample_wanxiangtai.xlsx":
                    seen.add(af)
                    found.append(f)
    return sorted(found, key=lambda p: os.path.getmtime(p), reverse=True)


def pick_file(arg=None):
    if arg:
        if os.path.exists(arg):
            return arg
        print("  找不到文件：%s" % arg)

    files = candidates()
    print("\n请选择要分析的报表：")
    if files:
        for i, f in enumerate(files, 1):
            size = os.path.getsize(f) / 1024
            print("  [%d] %s  (%.0f KB)" % (i, os.path.basename(f), size))
    print("  [0] 用内置示例数据试跑（sample_wanxiangtai.xlsx）")

    while True:
        s = ask("\n输入编号，或直接拖入/粘贴文件完整路径后回车：").strip().strip('"')
        if not s:
            continue
        if s == "0":
            return os.path.join(HERE, "sample_wanxiangtai.xlsx")
        if s.isdigit() and 1 <= int(s) <= len(files):
            return files[int(s) - 1]
        if os.path.exists(s):
            return s
        p = os.path.join(os.getcwd(), s)
        if os.path.exists(p):
            return p
        print("  路径无效，请重新输入。")


def pick_roi(arg=None):
    if arg:
        return float(arg)
    print("\n目标投产比 ROI 是多少？（不知道就按回车，默认 2.0）")
    print("  参考：保本 ROI = 1 ÷ 毛利率，毛利率 40% 则保本 ROI = 2.5")
    while True:
        s = ask("目标 ROI = ").strip()
        if not s:
            return 2.0
        try:
            v = float(s)
            if v > 0:
                return v
        except ValueError:
            pass
        print("  请输入数字，例如 2.5")


def main():
    args = [a for a in sys.argv[1:]]
    no_open = "--no-open" in args
    args = [a for a in args if a != "--no-open"]
    target = None
    if "--target-roi" in args:
        i = args.index("--target-roi")
        if i + 1 < len(args):
            target = float(args[i + 1])
            args = args[:i] + args[i + 2:]
    file_arg = args[0] if args else None

    print("=" * 56)
    print("  万相台无界推广分析器")
    print("=" * 56)

    path = pick_file(file_arg)
    target = pick_roi(target)

    print("\n正在分析：%s" % os.path.basename(path))
    try:
        df = az.load_table(path)
    except Exception as e:
        print("读取失败：%s" % e)
        wait_exit()
        return 1

    mapping = az.build_mapping(df)
    if "cost" not in mapping:
        print("\n没认出「消耗/花费」这一列。该文件里的列为：")
        for c in df.columns:
            print("   - %s" % c)
        print("\n请确认导出时保留了消耗列，或把列名改成「消耗」后重试。")
        wait_exit()
        return 2

    data = az.prepare(df, mapping)
    res = az.analyze(data, target)
    out = os.path.join(os.path.dirname(os.path.abspath(path)),
                       os.path.splitext(os.path.basename(path))[0] + "_分析报告.html")
    az.render(res, os.path.basename(path), out)

    ov = res["overview"]
    print("\n" + "-" * 56)
    print("识别到 %d 个字段，%d 行数据" % (len(mapping), len(data)))
    print("总消耗 %.0f 元 ｜ 成交 %.0f 元 ｜ ROI %.2f（目标 %.2f）"
          % (ov["cost"], ov["gmv"], ov["roi"], target))
    print("点击率 %.2f%% ｜ 转化率 %.2f%% ｜ 平均点击成本 %.2f 元"
          % (ov["ctr"] * 100, ov["cvr"] * 100, ov["ppc"]))
    print("诊断 %d 条：" % len(res["diagnostics"]))
    for d in res["diagnostics"]:
        flag = {"high": "[紧急]", "mid": "[关注]", "low": "[机会]"}[d["level"]]
        print("  %s %s" % (flag, d["title"]))
    print("-" * 56)
    print("\n报告已生成：\n  %s" % out)

    if not no_open:
        try:
            webbrowser.open("file:///" + out.replace("\\", "/"))
            print("已在浏览器中打开。")
        except Exception:
            print("请手动打开上面的路径。")

    wait_exit()
    return 0


if __name__ == "__main__":
    sys.exit(main())
