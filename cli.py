# -*- coding: utf-8 -*-
"""
调度数据网批复单业务IP填充工具 —— 命令行模式
不装图形界面库也能用，适合批量、脚本化、服务器环境。

用法示例：
    # 自动识别模板/掩码/起始IP，填好批复单（推荐）
    python3 cli.py --word 批复单.docx --auto -o 批复单_已填.docx

    # 只看预览不导出
    python3 cli.py --word 批复单.docx --auto
    # 用默认 /27 模板，实时起始 IP 10.20.1.160，非实时 10.20.2.160，打印结果
    python3 cli.py --realtime 10.20.1.160 --nonrealtime 10.20.2.160

    # 导出 Excel
    python3 cli.py --realtime 10.20.1.160 --nonrealtime 10.20.2.160 -o 业务地址表.xlsx

    # 用 /28 模板
    python3 cli.py --template 28 --prefix 28 --realtime 10.20.1.0 -o out.xlsx

    # 自动避让冲突（实时/非实时起始 IP 相同时）
    python3 cli.py --realtime 10.20.1.160 --nonrealtime 10.20.1.160 --avoid-clash

    # 批量站点（CSV：站点名称,实时起始IP,非实时起始IP）
    python3 cli.py --sites sites.csv -o 批量业务地址表.xlsx

    # 只算某几个业务
    python3 cli.py --realtime 10.20.1.160 --only 远动,保护

    # 把业务 IP 按「策略名称」填进 Word 批复单
    python3 cli.py --word 批复单.docx -o 批复单_已填.docx

    # 指定每张纵密表的起始 IP（不指定则按表中现有地址自动推断）
    python3 cli.py --word 批复单.docx \
        --word-base "省调第一接入网实时纵密参数=54.100.134.96,省调第一接入网非实时纵密参数=54.100.135.96" \
        -o 批复单_已填.docx

    # 只填批复单里的空白格，不动已有地址
    python3 cli.py --word 批复单.docx --word-only-empty -o 批复单_已填.docx
"""

from __future__ import annotations

import argparse
import csv
import os
import sys

import ip_core as C
from ip_core import Category, BusinessItem, HostIP, Issue
import ip_export as X
import ip_word as W


def print_table(results, limit=None):
    hdr = ["业务大类", "业务种类", "网别", "偏移", "★业务地址",
           "所属网段", "网关", "网段可用范围"]
    widths = [10, 12, 6, 5, 16, 20, 15, 26]

    def line(cells):
        out = []
        for c, w in zip(cells, widths):
            s = str(c)
            # 中文按两个字符宽度计算
            n = sum(2 if ord(ch) > 127 else 1 for ch in s)
            out.append(s + " " * max(0, w - n))
        return "  ".join(out)

    print(line(hdr))
    print("-" * (sum(widths) + 2 * len(widths)))
    rows = results if limit is None else results[:limit]
    for sn in rows:
        print(line([
            sn.category, sn.business,
            f"{sn.seq}网" if sn.seq else "单网",
            sn.offset,
            sn.ip,
            sn.subnet,
            sn.gateway or "-",
            f"{sn.first_host} - {sn.last_host}",
        ]))
    if limit and len(results) > limit:
        print(f"  ... 其余 {len(results) - limit} 条已省略")


def load_sites_csv(path: str):
    specs = []
    with open(path, "r", encoding="utf-8-sig") as f:
        for row in csv.reader(f):
            if not row or not row[0].strip():
                continue
            if row[0].strip().lower() in ("站点名称", "站点", "site"):
                continue
            name = row[0].strip()
            nets = [row[i].strip() if len(row) > i else "" for i in (1, 2)]
            specs.append((name, nets))
    return specs


def main(argv=None):
    p = argparse.ArgumentParser(
        description="IP 地址自动分配工具（命令行版）",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__)

    p.add_argument("--template", choices=["27", "28"], default="27",
                   help="预置模板：27=/27 标准（默认），28=/28 备用")
    p.add_argument("--prefix", type=int, default=None,
                   help="掩码位数（默认随模板，可覆盖，范围 24-30）")
    p.add_argument("--realtime", default=None, help="实时业务起始 IP，如 10.20.1.160（第四字节为基准）")
    p.add_argument("--nonrealtime", default=None, help="非实时业务起始 IP，如 10.20.2.160")
    p.add_argument("--gateway", choices=["first", "last", "none"], default="first",
                   help="网关策略：first=子网首个可用地址（默认），last=末个，none=不显示")
    p.add_argument("--avoid-clash", action="store_true",
                   help="自动避让冲突（把冲突大类的基准地址往后顺移）")
    p.add_argument("--only", default=None,
                   help="只计算指定业务，逗号分隔，如 远动,保护")
    p.add_argument("--sites", default=None, help="批量站点 CSV 路径（三列：站点名称,实时起始IP,非实时起始IP）")
    p.add_argument("-o", "--output", default=None,
                   help="导出文件路径（.xlsx 或 .csv）；不填则只打印到屏幕")
    p.add_argument("--word", default=None,
                   help="Word 批复单模板路径；按「策略名称」填入业务 IP")
    p.add_argument("--word-base", default=None,
                   help="各纵密表起始 IP，格式：\"表标题=起始IP,表标题2=起始IP2\"；"
                        "省略则按表中现有地址自动推断")
    p.add_argument("--word-only-empty", action="store_true",
                   help="只填批复单里的空白格，保留已有地址")
    p.add_argument("--auto", action="store_true",
                   help="自动识别批复单用的业务模板、掩码位数和各表起始 IP"
                        "（穷举后按与已有地址的吻合度打分）")
    p.add_argument("--quiet", action="store_true", help="不打印表格，只输出结果摘要")
    a = p.parse_args(argv)

    # ---- 载入模板 ----
    data = C.template_27() if a.template == "27" else C.template_28()
    cats, prefix, policy = C.dict_to_config(data)
    if a.prefix:
        if not 24 <= a.prefix <= 30:
            p.error("--prefix 必须在 24~30 之间")
        prefix = a.prefix
    policy = a.gateway

    # ---- 覆盖起始网段 ----
    if a.realtime:
        if C.parse_base(a.realtime) is None:
            p.error(f"--realtime 起始 IP 格式不正确：{a.realtime}")
        cats[0].base_network = a.realtime
    if a.nonrealtime:
        if C.parse_base(a.nonrealtime) is None:
            p.error(f"--nonrealtime 起始 IP 格式不正确：{a.nonrealtime}")
        if len(cats) > 1:
            cats[1].base_network = a.nonrealtime

    # ---- 只算指定业务 ----
    if a.only:
        keep = {s.strip() for s in a.only.split(",") if s.strip()}
        for cat in cats:
            for it in cat.items:
                it.enabled = it.name in keep

    # ---- Word 批复单填充模式 ----
    if a.word:
        if not os.path.exists(a.word):
            p.error(f"Word 批复单不存在：{a.word}")
        try:
            tables = W.scan_template(a.word)
        except Exception as e:
            p.error(f"无法解析 Word 批复单：{e}")
        if not tables:
            p.error("该文档中未找到含「策略名称 / 本端业务IP地址」的表格")

        # 自动识别：穷举模板 × 掩码，按与表中已有地址的吻合度打分
        auto_note = ""
        override_auto = {}
        if a.auto:
            tpls = {n: C.dict_to_config(f())[0] for n, f in C.TEMPLATES.items()}
            best = W.auto_detect(tables, tpls)
            if best is None:
                print("[提示] 批复单中没有已填地址可供比对，无法自动识别",
                      file=sys.stderr)
            else:
                cats = tpls[best["template"]]
                prefix = best["mask"]
                ov = {t.title: best["bases"][t.table_index]
                      for t in tables if t.table_index in best["bases"]}
                override_auto = ov
                auto_note = (f"自动识别：{best['template']} + /{best['mask']} 掩码，"
                             f"{best['hit']}/{best['total']} 行吻合")
                print(f"[自动识别] {auto_note}\n", file=sys.stderr)

        W.apply_mapping_by_category(tables, cats)

        # 解析用户指定的各表起始 IP
        override = {}
        if a.word_base:
            for part in a.word_base.split(","):
                if "=" not in part:
                    p.error(f"--word-base 格式不正确：{part}")
                k, v = part.split("=", 1)
                override[k.strip()] = v.strip()
        for k, v in (override_auto if a.auto else {}).items():
            override.setdefault(k, v)

        for t in tables:
            base = override.get(t.title) or t.suggested_base
            if not base:
                print(f"[跳过] {t.title}：未能确定起始 IP，请用 --word-base 指定",
                      file=sys.stderr)
                continue
            t.suggested_base = base
            src_cat = cats[1] if "非实时" in t.title and len(cats) > 1 else cats[0]
            cat = Category(name=src_cat.name, base_network=base,
                           items=[BusinessItem(i.name, list(i.offsets),
                                               i.remark, i.enabled)
                                  for i in src_cat.items])
            res, iss = C.allocate_category(cat, prefix, policy)
            pool = {}
            for h in res:
                pool.setdefault(h.business, []).append(h.ip)
            for r in t.rows:
                if a.word_only_empty and r.old_value:
                    r.new_value = ""
                    continue
                if r.business and pool.get(r.business):
                    r.new_value = W.format_ip_list(pool[r.business])
                else:
                    r.new_value = ""

        if not a.quiet:
            src_txt = ("命令行指定" if a.word_base
                       else "自动识别推断" if a.auto
                       else "按表中现有地址推断")
            print(f"\n掩码 /{prefix}    起始 IP 来源：{src_txt}\n")
            for t in tables:
                print(f"【{t.title}】起始 IP = {t.suggested_base}")
                for r in t.rows:
                    mark = "→" if r.new_value else " "
                    print(f"    {r.tunnel:<3} {r.policy:<18} "
                          f"{r.old_value or '(空)':<22} {mark} {r.new_value or '(跳过)'}")
                print()

        if a.output:
            out = a.output if a.output.lower().endswith(".docx") else a.output + ".docx"
            W.fill_docx(a.word, out, tables)
            print(f"已导出：{out}")
        else:
            print("提示：加 -o 输出文件.docx 才会生成填充后的批复单")

        n = len([r for t in tables for r in t.rows if r.new_value])
        print(f"\n共填充 {n} 处「本端业务IP地址」")
        return 0

    # ---- 批量站点模式 ----
    if a.sites:
        if not os.path.exists(a.sites):
            p.error(f"站点 CSV 不存在：{a.sites}")
        specs = load_sites_csv(a.sites)
        if not specs:
            p.error("站点 CSV 中没有有效数据")
        rows, issues = C.allocate_sites(specs, cats, prefix, policy)
        headers = ["站点"] + HostIP.header()

        if not a.quiet:
            print(f"\n站点数：{len(specs)}    业务地址数：{len(rows)}\n")
            for site, nets in specs:
                print(f"【{site}】实时起始IP={nets[0] or '(默认)'}  非实时起始IP={nets[1] or '(默认)'}")
                sub = [r for r in rows if r["站点"] == site]
                for r in sub:
                    print(f"    {r['业务大类']:<8} {r['业务名称']:<12} {r['网别']:<4} "
                          f"{r['业务地址']:<15} 所属网段 {r['所属网段']:<18} "
                          f"gw={r['网关']:<14} {r['网段可用范围']}")
                print()

        if a.output:
            if a.output.lower().endswith(".csv"):
                X.export_dict_rows_csv(a.output, rows, headers)
            else:
                X.export_dict_rows_excel(a.output, rows, headers)
            print(f"已导出：{a.output}")

        n_err = len([i for i in issues if i.level == "error"])
        for i in issues:
            print(i, file=sys.stderr)
        return 1 if n_err else 0

    # ---- 单站点模式 ----
    if a.avoid_clash:
        cats, notes = C.auto_resolve_conflicts(cats, prefix)
        for n in notes:
            print(f"[自动避让] {n}")

    results, issues = C.allocate_all(cats, prefix, policy)

    if not a.quiet:
        print(f"\n掩码 /{prefix}    网关策略：{C.GATEWAY_POLICY_LABEL.get(policy, policy)}")
        for cat in cats:
            base = C.parse_base(cat.base_network)
            print(f"  {cat.name}：起始 IP {cat.base_network}"
                  f"（基准第四字节 = {base[3] if base else '?'}）")
        print()
        print_table(results)

    if a.output:
        if a.output.lower().endswith(".csv"):
            X.export_csv(a.output, results)
        else:
            X.export_excel(a.output, results, issues, prefix, policy, cats)
        print(f"\n已导出：{a.output}")

    print(f"\n共 {len(results)} 个业务地址", end="")
    n_err = len([i for i in issues if i.level == "error"])
    n_warn = len([i for i in issues if i.level == "warning"])
    print(f"，{n_err} 处错误，{n_warn} 处警告")
    for i in issues:
        print("  " + str(i), file=sys.stderr)

    return 1 if n_err else 0


if __name__ == "__main__":
    sys.exit(main())
