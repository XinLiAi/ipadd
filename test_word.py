# -*- coding: utf-8 -*-
"""
Word 批复单填充模块单元测试
运行： python3 test_word.py
"""

import os
import tempfile

import ip_core as C
import ip_word as W

SRC = "/data/inputs/那曲色尼风电调度数据网接入地址批复单.docx"
BUILTIN = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                       "templates", "approval_form_template.docx")
TMP = tempfile.mkdtemp(prefix="ipword_")

# 批复单实测的起始 IP（网络位）
BASES = {
    "省调第一接入网实时纵密参数": "54.100.134.96",
    "省调第一接入网非实时纵密参数": "54.100.135.96",
    "省调第二接入网实时纵密参数": "54.124.6.48",
    "省调第二接入网非实时纵密参数": "54.124.7.48",
}


def _tpl() -> str:
    return SRC if os.path.exists(SRC) else BUILTIN


def test_scan():
    """能扫出 4 张策略表，且列索引、推断起始 IP 正确"""
    tabs = W.scan_template(_tpl())
    assert len(tabs) == 4, f"期望 4 张表，实际 {len(tabs)}"

    titles = [t.title for t in tabs]
    for k in BASES:
        assert any(k in t for t in titles), f"缺少表：{k}"

    for t in tabs:
        assert t.col_policy == 3, f"{t.title} 策略名列索引 {t.col_policy}"
        assert t.col_value == 4, f"{t.title} 本端IP列索引 {t.col_value}"
        assert t.header_row == 2
        assert t.suggested_base == BASES[t.title], \
            f"{t.title} 推断起始IP {t.suggested_base} != {BASES[t.title]}"
        assert len(t.rows) > 0
    print("OK  test_scan  (4 张表，起始IP全部自动推断正确)")


def test_match():
    """策略名称 → 业务名映射"""
    cases = [
        ("备调远动fes1", "远动"), ("地调远动fes2", "远动"), ("区调远动fes4", "远动"),
        ("区调PMU1", "PMU"), ("区调PMU2", "PMU"),
        ("区调稳控", "稳控"),
        ("区调保护1", "保护"), ("区调保护2", "保护"),
        ("区调内网管控（网安）", "网安"),
        ("区调电量", "电量"), ("地调电量", "电量"),
        ("区调保信", "保信"),
        ("区调光功率", "光功率"),
        ("区调水情", "水情"),
        ("区调同步时钟", "同步时钟"),
        ("区调安控1", "安控"),
        ("区调风功率", "风功率"),
        ("区调现货市场", "现货市场"),
        ("区调一次调频", "一次调频"),
        ("未知业务XYZ", None),
    ]
    # 业务名取两套模板的合集（一次调频/一次设备只在 /27 模板里）
    names = []
    for fn in (C.template_27, C.template_28):
        cs, _p, _g = C.dict_to_config(fn())
        for c in cs:
            for i in c.items:
                if i.name not in names:
                    names.append(i.name)
    for policy, expect in cases:
        got = W.match_business(policy, names)
        assert got == expect, f"「{policy}」→ {got}，期望 {expect}"
    print("OK  test_match  (19 条策略名称映射全部正确)")


def test_category_isolation():
    """★ 关键 bug：跨大类撞名（实时「网安」vs 非实时「网络安全监测」）必须隔离"""
    cats, _p, _g = C.dict_to_config(C.template_27())
    rt, nrt = cats[0], cats[1]
    assert [i.name for i in rt.items].count("网安") == 1
    assert "网络安全监测" in [i.name for i in nrt.items]
    # 策略「区调内网管控（网安）」在实时大类 → 网安；在非实时大类 → 网络安全监测
    assert W.match_in_category("区调内网管控（网安）", rt) == "网安"
    assert W.match_in_category("区调内网管控（网安）", nrt) == "网络安全监测"
    # 按标题选大类的归属判断
    assert W.category_for_title("省调第一接入网非实时纵密参数", cats).name == nrt.name
    assert W.category_for_title("省调第一接入网实时纵密参数", cats).name == rt.name
    print("OK  test_category_isolation  (实时网安 / 非实时网络安全监测 已隔离)")


def test_disabled_not_matched():
    """停用的业务不参与匹配，标记为 business_disabled"""
    cats, _p, _g = C.dict_to_config(C.template_28())
    tabs = W.scan_template(_tpl())
    W.apply_mapping_by_category(tabs, cats)
    dis = [r.business_disabled for t in tabs for r in t.rows if r.business_disabled]
    assert "同步时钟" in dis and "安控" in dis, dis
    assert not [r for t in tabs for r in t.rows
                if r.business in ("同步时钟", "安控")]
    print(f"OK  test_disabled_not_matched  (停用业务标记未启用：{sorted(set(dis))})")


def test_auto_detect():
    """★ 自动识别：应挑出 /28 模板 + 掩码 28，且 32/32 吻合"""
    tabs = W.scan_template(_tpl())
    tpls = {n: C.dict_to_config(f())[0] for n, f in C.TEMPLATES.items()}
    best = W.auto_detect(tabs, tpls)
    assert best is not None
    assert best["template"] == "28位掩码（备用模板）", best["template"]
    assert best["mask"] == 28, best["mask"]
    assert best["hit"] == best["total"] == 32, (best["hit"], best["total"])
    assert best["bases"][tabs[0].table_index] == "54.100.134.96"
    assert best["bases"][tabs[1].table_index] == "54.100.135.96"
    print("OK  test_auto_detect  (自动选出 /28 模板 + 掩码28，32/32 行吻合)")


def test_infer_base_by_mask():
    """不同掩码下推断出的网络位应正确"""
    vals = ["54.100.134.103-104", "54.100.134.98", "54.100.134.99-100"]
    assert W.infer_base(vals, 28) == "54.100.134.96"
    assert W.infer_base(vals, 27) == "54.100.134.96"
    assert W.infer_base([], 28) == ""
    print("OK  test_infer_base_by_mask  (/27、/28 均能正确推断网络位)")


def test_format():
    """地址格式化与批复单写法一致"""
    assert W.format_ip_list(["54.100.134.98"]) == "54.100.134.98"
    assert W.format_ip_list(["54.100.134.103", "54.100.134.104"]) == "54.100.134.103-104"
    assert W.format_ip_list(["54.100.134.99", "54.100.134.100"]) == "54.100.134.99-100"
    assert W.format_ip_list(["54.100.135.98", "54.100.135.118"]) == "54.100.135.98、54.100.135.118"
    assert W.format_ip_list([]) == ""
    assert W.format_ip_list(["54.100.134.99", "54.100.134.100", "54.100.134.101"]) == \
        "54.100.134.99-101"
    print("OK  test_format  (单网/连续范围/不连续/多段)")


def _plan(only_empty=False, cats=None, prefix=None):
    if cats is None:
        cats, prefix, policy = C.dict_to_config(C.template_28())
    else:
        policy = C.GATEWAY_FIRST
    tabs = W.scan_template(_tpl())
    W.apply_mapping_by_category(tabs, cats)
    for t in tabs:
        base = BASES[t.title]
        src = cats[1] if "非实时" in t.title else cats[0]
        cat = C.Category(name=src.name, base_network=base,
                         items=[C.BusinessItem(i.name, list(i.offsets),
                                               i.remark, i.enabled) for i in src.items])
        res, _ = C.allocate_category(cat, prefix, policy)
        pool = {}
        for h in res:
            pool.setdefault(h.business, []).append(h.ip)
        for r in t.rows:
            if only_empty and r.old_value:
                r.new_value = ""
            elif r.business and pool.get(r.business):
                r.new_value = W.format_ip_list(pool[r.business])
            else:
                r.new_value = ""
    return tabs


def test_plan_matches_original():
    """★ 核心：算出的地址必须与批复单原文逐行一致"""
    tabs = _plan()
    bad = []
    filled = 0
    for t in tabs:
        for r in t.rows:
            if not r.new_value:
                continue
            filled += 1
            if r.old_value and r.old_value != r.new_value:
                bad.append((t.title, r.policy, r.old_value, r.new_value))
    assert not bad, f"与原文不一致：\n" + "\n".join(str(x) for x in bad)
    assert filled == 44, f"填充行数 {filled} != 44"
    print("OK  test_plan_matches_original  (已填的 32 行与批复单原文 100% 一致)")


def test_plan_fills_empty():
    """原本空白的 12 行远动地址应被补填"""
    tabs = _plan()
    new_rows = [(t.title, r.policy, r.new_value)
                for t in tabs for r in t.rows if not r.old_value and r.new_value]
    assert len(new_rows) == 12, f"补填 {len(new_rows)} 行，期望 12"
    for _t, p, v in new_rows:
        assert "远动" in p, p
    t1 = new_rows[0][2]
    assert t1 == "54.100.134.103-104", t1
    print("OK  test_plan_fills_empty  (补填 12 行远动，如 地调/区调远动fes)")


def test_only_empty_mode():
    """「仅填空白」模式只填 12 行，不动已有地址"""
    tabs = _plan(only_empty=True)
    n = len([r for t in tabs for r in t.rows if r.new_value])
    assert n == 12, n
    for t in tabs:
        for r in t.rows:
            if r.old_value:
                assert not r.new_value
    print("OK  test_only_empty_mode  (只补 12 处空白，已有地址原样保留)")


def test_unmatched_policy():
    """未启用/未分配地址的业务（同步时钟/安控）应标记并跳过，不填入"""
    tabs = _plan()
    skipped = [(t.title, r.policy, r.business_disabled)
               for t in tabs for r in t.rows if not r.new_value]
    dis = {d for _t, _p, d in skipped if d}
    assert "同步时钟" in dis, dis
    assert "安控" in dis, dis
    # 这些行的 business 必须为空（不参与填充）
    assert not [r for t in tabs for r in t.rows
                if r.business in ("同步时钟", "安控")]
    print(f"OK  test_unmatched_policy  (跳过 {len(skipped)} 行，"
          f"标记为未启用：{sorted(dis)})")


def test_fill_docx():
    """实际生成 Word，回填结果正确且原模板不变"""
    src = _tpl()
    before = os.path.getsize(src)
    tabs = _plan()
    out = os.path.join(TMP, "批复单_已填.docx")
    W.fill_docx(src, out, tabs)

    assert os.path.exists(out) and os.path.getsize(out) > 10000
    assert os.path.getsize(src) == before, "原模板不应被修改"

    import docx
    d = docx.Document(out)
    # 第一接入网实时表 = 表3
    t3 = d.tables[3]
    got = {}
    for ri in range(3, len(t3.rows)):
        c = t3.rows[ri].cells
        if len(c) > 4 and c[3].text.strip():
            got[c[3].text.strip()] = c[4].text.strip()
    assert got["备调远动fes1"] == "54.100.134.103-104", got["备调远动fes1"]
    assert got["区调稳控"] == "54.100.134.99-100"
    assert got["区调内网管控（网安）"] == "54.100.134.98"
    assert got["区调远动fes1"] == "54.100.134.103-104", "补填行"

    # 非实时表 = 表4
    t4 = d.tables[4]
    got4 = {}
    for ri in range(3, len(t4.rows)):
        c = t4.rows[ri].cells
        if len(c) > 4 and c[3].text.strip():
            got4[c[3].text.strip()] = c[4].text.strip()
    assert got4["区调电量"] == "54.100.135.106"
    assert got4["区调内网管控（网安）"] == "54.100.135.98", got4["区调内网管控（网安）"]
    assert got4["区调同步时钟"] == "", f"未分配地址的业务应保持空白，实际 {got4['区调同步时钟']!r}"
    print("OK  test_fill_docx  (生成成功，4 张表取值正确，原模板未改动)")


def test_fill_preserves_format():
    """填充后字体格式应与原单元格一致"""
    import docx
    d0 = docx.Document(_tpl())
    tabs = _plan()
    out = os.path.join(TMP, "fmt.docx")
    W.fill_docx(_tpl(), out, tabs)
    d1 = docx.Document(out)

    def font_of(doc, ti, ri, ci):
        c = doc.tables[ti].rows[ri].cells[ci]
        for p in c.paragraphs:
            for r in p.runs:
                if r.text.strip():
                    return r.font.name, r.font.size
        return None

    # 找一个原本有值的单元格（表3 R3 = 备调远动fes1）
    f_before = font_of(d0, 3, 3, 4)
    f_after = font_of(d1, 3, 3, 4)
    assert f_before == f_after, f"格式变了：{f_before} -> {f_after}"
    print(f"OK  test_fill_preserves_format  (字体沿用原单元格 {f_after})")


def test_custom_base():
    """改起始 IP 后地址随之变化"""
    cats, prefix, policy = C.dict_to_config(C.template_28())
    tabs = W.scan_template(_tpl())
    W.apply_mapping_by_category(tabs, cats)
    t = tabs[0]
    cat = C.Category(name=cats[0].name, base_network="54.100.134.160",
                     items=[C.BusinessItem(i.name, list(i.offsets), i.remark, i.enabled)
                            for i in cats[0].items])
    res, _ = C.allocate_category(cat, prefix, policy)
    pool = {}
    for h in res:
        pool.setdefault(h.business, []).append(h.ip)
    for r in t.rows:
        if r.business and pool.get(r.business):
            r.new_value = W.format_ip_list(pool[r.business])
    assert t.rows[0].new_value == "54.100.134.167-168", t.rows[0].new_value
    print("OK  test_custom_base  (基准 160 + 偏移7/8 → 54.100.134.167-168)")


def test_nrt_detected_after_autodetect():
    """★ 用户反馈的 bug：一/二平面非实时网安此前漏填，自动识别后必须填入"""
    tabs = W.scan_template(_tpl())
    tpls = {n: C.dict_to_config(f())[0] for n, f in C.TEMPLATES.items()}
    best = W.auto_detect(tabs, tpls)
    cats = tpls[best["template"]]
    W.apply_mapping_by_category(tabs, cats)
    for t in tabs:
        if "非实时" not in t.title:
            continue
        cat = W.category_for_title(t.title, cats)
        base = best["bases"][t.table_index]
        c = C.Category(cat.name, base,
                       [C.BusinessItem(i.name, list(i.offsets), i.remark, i.enabled)
                        for i in cat.items])
        res, _ = C.allocate_category(c, best["mask"], C.GATEWAY_FIRST)
        pool = {}
        for h in res:
            pool.setdefault(h.business, []).append(h.ip)
        rows = {r.policy: r for r in t.rows}
        for name in ("区调电量", "区调保信", "区调光功率", "区调水情", "区调内网管控（网安）"):
            if name not in rows:
                continue
            r = rows[name]
            assert r.business, f"{t.title}/{name} 未匹配到业务"
            ips = pool.get(r.business)
            r.new_value = W.format_ip_list(ips) if ips else ""
            assert r.new_value, f"{t.title}/{name} 仍未填入！"
            if r.old_value:
                assert r.new_value == r.old_value, \
                    f"{t.title}/{name}: {r.old_value} -> {r.new_value}"
    print("OK  test_nrt_detected_after_autodetect  (一/二平面非实时业务全部识别并填入)")


if __name__ == "__main__":
    test_scan()
    test_match()
    test_category_isolation()
    test_disabled_not_matched()
    test_auto_detect()
    test_infer_base_by_mask()
    test_format()
    test_plan_matches_original()
    test_plan_fills_empty()
    test_only_empty_mode()
    test_unmatched_policy()
    test_fill_docx()
    test_fill_preserves_format()
    test_custom_base()
    test_nrt_detected_after_autodetect()
    print("\n全部 Word 测试通过 ✓")
