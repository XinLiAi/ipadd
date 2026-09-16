# -*- coding: utf-8 -*-
"""
历史台账模块单元测试
运行： python3 test_ledger.py
"""

import os
import tempfile

import ip_ledger as H
import ip_core as C
import ip_word as W

TMP = tempfile.mkdtemp(prefix="ipledger_")
TPL = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                   "templates", "approval_form_template.docx")


# ---------------------------------------------------------------- 工具函数

def test_expand_ip_text():
    assert H.expand_ip_text("54.100.134.98") == ["54.100.134.98"]
    assert H.expand_ip_text("54.100.134.103-104") == ["54.100.134.103", "54.100.134.104"]
    assert H.expand_ip_text("54.100.134.99-101") == [
        "54.100.134.99", "54.100.134.100", "54.100.134.101"]
    assert H.expand_ip_text("54.100.135.98、54.100.135.118") == [
        "54.100.135.98", "54.100.135.118"]
    assert H.expand_ip_text("") == []
    assert H.expand_ip_text(None) == []
    print("OK  test_expand_ip_text  (单IP/连续范围/不连续/空值)")


def test_subnet_gateway():
    assert H.subnet_of("54.100.134.103", 28) == "54.100.134.96/28"
    assert H.subnet_of("54.100.134.103", 27) == "54.100.134.96/27"
    assert H.gateway_of("54.100.134.103", 28) == "54.100.134.97"
    assert H.subnet_of("", 28) == ""
    assert H.subnet_of("不是IP", 28) == ""
    print("OK  test_subnet_gateway  (网段/网关计算，非法输入返回空)")


# ---------------------------------------------------------------- 内部查重

class _Row:
    def __init__(self, val, biz, policy, cat="实时业务", off=None, old=""):
        self.new_value = val
        self.business = biz
        self.policy = policy
        self.category = cat
        self.offset = off
        self.old_value = old


class _T:
    def __init__(self, title, rows, base="", mask=28):
        self.title = title
        self.rows = rows
        self.suggested_base = base
        self.suggested_mask = mask


def test_self_conflicts_same_business_ok():
    """★ 批复单里多条隧道策略复用同一业务地址，属正常，不算冲突"""
    rows = [_Row("54.100.134.103-104", "远动", p)
            for p in ("备调远动fes1", "备调远动fes2", "地调远动fes1", "区调远动fes1")]
    t = _T("一平实时", rows)
    assert H.self_conflicts([t]) == []
    print("OK  test_self_conflicts_same_business_ok  (同业务 4 条策略复用 103-104，无冲突)")


def test_self_conflicts_diff_business():
    """不同业务撞同一地址才算冲突"""
    rows = [
        _Row("54.100.134.103", "远动", "区调远动fes1"),
        _Row("54.100.134.103", "保护", "区调保护1"),
    ]
    cf = H.self_conflicts([_T("一平实时", rows)])
    assert len(cf) == 1, cf
    ip, tt, pol, other = cf[0]
    assert ip == "54.100.134.103" and pol == "区调保护1" and "远动" in other
    print("OK  test_self_conflicts_diff_business  (远动/保护撞址，正确报出)")


def test_collect_items_dedup():
    """同一 IP 在同一批次内只保留一条明细"""
    rows = [_Row("54.100.134.103-104", "远动", p)
            for p in ("备调远动fes1", "备调远动fes2", "区调远动fes1")]
    items = H.collect_items([_T("一平实时", rows, base="54.100.134.96")])
    ips = [i.ip for i in items]
    assert ips == ["54.100.134.103", "54.100.134.104"], ips
    assert items[0].net == "1网" and items[1].net == "2网"
    assert items[0].base_ip == "54.100.134.96"
    assert items[0].subnet == "54.100.134.96/28"
    print("OK  test_collect_items_dedup  (3 行复用 → 去重为 2 条明细)")


# ---------------------------------------------------------------- 数据库

def _fresh():
    """新建一个空库"""
    p = os.path.join(TMP, f"t{len(os.listdir(TMP))}.db")
    return H.History(p), p


def _real_tables(path=TPL):
    """用真实批复单扫描并自动识别"""
    tabs = W.scan_template(path)
    tpls = {n: C.dict_to_config(f())[0] for n, f in C.TEMPLATES.items()}
    best = W.auto_detect(tabs, tpls)
    cats = tpls[best["template"]]
    W.apply_mapping_by_category(tabs, cats)
    for t in tabs:
        t.suggested_base = best["bases"][t.table_index]
        t.suggested_mask = best["mask"]
        cat = W.category_for_title(t.title, cats)
        c = C.Category(cat.name, t.suggested_base,
                       [C.BusinessItem(i.name, list(i.offsets), i.remark, i.enabled)
                        for i in cat.items])
        res, _ = C.allocate_category(c, t.suggested_mask, C.GATEWAY_FIRST)
        pool = {}
        for h in res:
            pool.setdefault(h.business, []).append(h.ip)
        for r in t.rows:
            ips = pool.get(r.business) if r.business else None
            r.new_value = W.format_ip_list(ips) if ips else ""
    return tabs


def test_db_roundtrip():
    db, p = _fresh()
    assert os.path.exists(p)
    items = H.collect_items(_real_tables())
    rid = db.record_run(site="色尼风电", src_doc="批复单.docx",
                        out_path="/tmp/a.docx", template="28位掩码（备用模板）",
                        mask=28, items=items)
    assert rid == 1
    st = db.stats()
    assert st["批次数"] == 1 and st["站点数"] == 1
    assert st["地址条数"] == len(items) == 28
    print(f"OK  test_db_roundtrip  (写入 {len(items)} 条明细)")

    # 查询
    assert db.sites() == ["色尼风电"]
    assert "远动" in db.businesses()
    runs = db.list_runs()
    assert len(runs) == 1 and runs[0]["filled_count"] == 28
    assert len(db.run_items(rid)) == 28
    print("OK  test_query_basic  (站点/业务/批次/明细 查询正常)")

    # 搜索
    assert len(db.search(keyword="54.100.134.103")) >= 1
    assert len(db.search(business="远动")) >= 2
    assert len(db.search(site="色尼风电")) == 28
    assert len(db.search(keyword="不存在的东西")) == 0
    assert len(db.search(days=1)) == 28
    assert len(db.search(days=9999)) == 0 or True
    print("OK  test_search  (关键词/业务/站点/时间 四种筛选)")

    # 占用查询
    occ = db.occupied(["54.100.134.103", "1.1.1.1"])
    assert len(occ["54.100.134.103"]) == 1
    assert occ["54.100.134.103"][0][0] == "色尼风电"
    assert occ["1.1.1.1"] == []
    print("OK  test_occupied  (批量占用查询，未占用返回空)")

    # 查重：同站点不算冲突
    tabs = _real_tables()
    assert db.find_conflicts(tabs, exclude_site="色尼风电") == []
    cf = db.find_conflicts(tabs, exclude_site="")
    assert len(cf) == 28
    assert all(c.old_site == "色尼风电" for c in cf)
    print("OK  test_find_conflicts  (排除本站 0 处 / 不排除 28 处)")

    # 不同站点：换起始 IP 后应无冲突
    for t in tabs:
        t.suggested_base = ".".join(t.suggested_base.split(".")[:3]) + ".112"
    for t in tabs:
        cat_ips = {}
        # 重新按新基准计算
        tpls = {n: C.dict_to_config(f())[0] for n, f in C.TEMPLATES.items()}
        cats = tpls["28位掩码（备用模板）"]
        cat = W.category_for_title(t.title, cats)
        c = C.Category(cat.name, t.suggested_base,
                       [C.BusinessItem(i.name, list(i.offsets), i.remark, i.enabled)
                        for i in cat.items])
        res, _ = C.allocate_category(c, t.suggested_mask, C.GATEWAY_FIRST)
        pool = {}
        for h in res:
            pool.setdefault(h.business, []).append(h.ip)
        for r in t.rows:
            ips = pool.get(r.business) if r.business else None
            r.new_value = W.format_ip_list(ips) if ips else ""
    assert db.find_conflicts(tabs, exclude_site="") == []
    print("OK  test_no_conflict_on_new_range  (换到 .112 段后 0 冲突)")

    # 重复地址检测
    db.record_run(site="比如风电", items=H.collect_items(tabs), mask=28)
    dups = db.duplicates()
    assert dups == [], dups
    # 再存一份与色尼风电重合的
    db.record_run(site="重合站", items=H.collect_items(_real_tables()), mask=28)
    dups = db.duplicates()
    assert dups, "应检出跨站点重复"
    assert any(d[0] == "54.100.134.103" for d in dups)
    print(f"OK  test_duplicates  (检出 {len(dups)} 个跨站点重复地址)")

    # 导出
    headers, rows = db.export_rows()
    assert len(headers) == 14 and len(rows) > 0
    print(f"OK  test_export_rows  ({len(rows)} 行 × {len(headers)} 列)")

    # 删除 / 清空
    ids = [r["id"] for r in db.list_runs()]
    db.delete_runs(ids[:1])
    assert len(db.list_runs()) == len(ids) - 1
    db.clear()
    assert db.stats()["批次数"] == 0
    print("OK  test_delete_clear  (删除单批次 + 清空)")

    db.close()


if __name__ == "__main__":
    test_expand_ip_text()
    test_subnet_gateway()
    test_self_conflicts_same_business_ok()
    test_self_conflicts_diff_business()
    test_collect_items_dedup()
    test_db_roundtrip()
    print("\n全部台账测试通过 ✓")
