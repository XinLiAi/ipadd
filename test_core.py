# -*- coding: utf-8 -*-
"""
ip_core 单元测试（单业务 IP 模式）
运行： python3 test_core.py
"""

import ip_core as C


def test_parse_base():
    assert C.parse_base("10.20.1.160") == (10, 20, 1, 160)
    assert C.parse_base("10.20.1.160/27") == (10, 20, 1, 160)
    assert C.parse_base("10.20.1") == (10, 20, 1, 0)      # 只给前三位，基准 0
    assert C.parse_base("abc") is None
    assert C.parse_base("10.20.1") == (10, 20, 1, 0)
    assert C.parse_base("10.20.1.300") is None
    assert C.parse_base("") is None
    print("OK  test_parse_base")


def test_match_source_table():
    """逐行核对《业务地址计算工具.xlsx》工作表「27位计算」的所得值"""
    cats, prefix, policy = C.dict_to_config(C.template_27())
    res, issues = C.allocate_all(cats, prefix, policy)

    # (大类, 业务, 偏移, 第四字节所得值) —— 第四字节即原表 F/G 列
    expect = [
        ("实时业务", "远动", 7, 167), ("实时业务", "远动", 8, 168),
        ("实时业务", "保护", 9, 169), ("实时业务", "保护", 10, 170),
        ("实时业务", "PMU", 15, 175), ("实时业务", "PMU", 16, 176),
        ("实时业务", "网安", 22, 182),
        ("实时业务", "稳控", 24, 184), ("实时业务", "稳控", 25, 185),
        ("实时业务", "一次调频", 3, 163),
        ("实时业务", "宽频", 1, 161), ("实时业务", "宽频", 2, 162),
        ("非实时业务", "电量", 10, 170),
        ("非实时业务", "保信（保护）", 5, 165),
        ("非实时业务", "光功率", 20, 180),
        ("非实时业务", "风功率", 19, 179),
        ("非实时业务", "水情", 7, 167),
        ("非实时业务", "网络安全监测", 22, 182),
        ("非实时业务", "现货市场", 9, 169),
        ("非实时业务", "一次设备", 3, 163),
    ]
    assert len(res) == len(expect), f"结果数 {len(res)} != 期望 {len(expect)}"
    for h, (cat, biz, off, last) in zip(res, expect):
        assert h.category == cat, f"{h.category} != {cat}"
        assert h.business == biz, f"{h.business} != {biz}"
        assert h.offset == off, f"{biz} 偏移 {h.offset} != {off}"
        assert h.last_octet == last, f"{biz} 所得值 {h.last_octet} != 期望 {last}"
    assert not [i for i in issues if i.level == "error"]
    print("OK  test_match_source_table  (20 行所得值与原表 F/G 列一致)")


def test_ip_format():
    """核心：业务地址必须是 10.20.1.167 这种完整 IP，而不是网段"""
    h = C.calc_host((10, 20, 1, 160), 7, 27)
    assert h.ip == "10.20.1.167", h.ip
    assert h.last_octet == 167
    assert h.prefix == 27
    # 所属网段 / 网关 / 广播 是配套说明信息
    assert h.subnet == "10.20.1.160/27"
    assert h.netmask == "255.255.255.224"
    assert h.gateway == "10.20.1.161"
    assert h.first_host == "10.20.1.161"
    assert h.last_host == "10.20.1.190"
    assert h.broadcast == "10.20.1.191"
    assert h.hosts == 30
    assert h.in_base_subnet is True
    print("OK  test_ip_format  (10.20.1.160 + 7 → 业务地址 10.20.1.167)")


def test_28_mask():
    h = C.calc_host((10, 20, 1, 0), 7, 28)
    assert h.ip == "10.20.1.7"
    assert h.subnet == "10.20.1.0/28"
    assert h.hosts == 14
    assert h.last_host == "10.20.1.14"
    print("OK  test_28_mask  (/28：10.20.1.0 + 7 → 10.20.1.7)")


def test_gateway_policy():
    b = (10, 20, 1, 160)
    assert C.calc_host(b, 7, 27, C.GATEWAY_FIRST).gateway == "10.20.1.161"
    assert C.calc_host(b, 7, 27, C.GATEWAY_LAST).gateway == "10.20.1.190"
    assert C.calc_host(b, 7, 27, C.GATEWAY_NONE).gateway == ""
    print("OK  test_gateway_policy  (首地址/末地址/不显示)")


def test_carry():
    """第四字节超过 255 → 向第三字节进位"""
    h = C.calc_host((10, 20, 1, 250), 10, 27)
    assert h.ip == "10.20.2.4", h.ip
    assert h.in_base_subnet is False
    print("OK  test_carry  (10.20.1.250 + 10 → 业务地址 10.20.2.4，标记跨段)")


def test_carry_deep():
    """第三字节也溢出 → 向第二字节进位"""
    h = C.calc_host((10, 20, 255, 250), 10, 27)
    assert h.ip == "10.21.0.4", h.ip
    print("OK  test_carry_deep  (10.20.255.250 + 10 → 10.21.0.4)")


def test_overflow_error():
    try:
        C.calc_host((10, 255, 255, 250), 10, 27)
        raise AssertionError("应当抛出异常")
    except ValueError:
        pass
    try:
        C.calc_host((10, 20, 1, 160), -1, 27)
        raise AssertionError("应当抛出异常")
    except ValueError:
        pass
    print("OK  test_overflow_error  (越界/负偏移均报错)")


def test_cross_subnet_warning():
    """偏移超出基准子网容量 → 给警告但仍算出地址"""
    cat = C.Category(name="测试", base_network="10.20.1.0",
                     items=[C.BusinessItem("超范围", [40])])
    res, issues = C.allocate_category(cat, 28)     # /28 只有 14 个可用地址
    assert res[0].ip == "10.20.1.40"
    assert res[0].in_base_subnet is False
    warns = [i for i in issues if i.level == "warning"]
    assert any("超出基准子网" in i.message for i in warns)
    print("OK  test_cross_subnet_warning  (/28 下偏移 40 → 10.20.1.40 并提示跨段)")


def test_gateway_clash_warning():
    """业务地址撞网关 → 警告"""
    cat = C.Category(name="测试", base_network="10.20.1.160",
                     items=[C.BusinessItem("撞网关", [1])])
    res, issues = C.allocate_category(cat, 27)
    assert res[0].ip == "10.20.1.161"
    assert any("与网关地址相同" in i.message for i in issues)
    print("OK  test_gateway_clash_warning  (偏移 1 → 10.20.1.161 撞网关，已提示)")


def test_conflict_detect():
    """实时/非实时起始 IP 相同 → 应报出 5 处业务地址重复"""
    cats, prefix, policy = C.dict_to_config(C.template_27())
    cats[1].base_network = "10.20.1.160"          # 改成与实时一样
    res, issues = C.allocate_all(cats, prefix, policy)
    errs = [i for i in issues if i.level == "error"]
    assert len(errs) == 5, f"期望 5 处冲突，实际 {len(errs)}"
    assert all("都是" in i.message for i in errs)
    print("OK  test_conflict_detect  (检出 5 处业务地址重复)")


def test_no_conflict_when_separated():
    cats, prefix, policy = C.dict_to_config(C.template_27())
    res, issues = C.allocate_all(cats, prefix, policy)
    assert not [i for i in issues if i.level == "error"]
    print("OK  test_no_conflict_when_separated  (10.20.1.x 与 10.20.2.x 不冲突)")


def test_disable_item():
    cats, prefix, policy = C.dict_to_config(C.template_27())
    cats[0].items[0].enabled = False
    res, issues = C.allocate_all(cats, prefix, policy)
    assert all(h.business != "远动" for h in res)
    print("OK  test_disable_item")


def test_batch_sites():
    cats, prefix, policy = C.dict_to_config(C.template_27())
    specs = [
        ("甲站", ["10.20.1.160", "10.20.2.160"]),
        ("乙站", ["10.20.11.160", "10.20.12.160"]),
        ("丙站", ["", ""]),
    ]
    rows, issues = C.allocate_sites(specs, cats, prefix, policy)
    assert len(rows) == 60, len(rows)
    assert {r["站点"] for r in rows} == {"甲站", "乙站", "丙站"}
    jia = [r for r in rows if r["站点"] == "甲站" and r["业务名称"] == "远动"]
    assert jia[0]["业务地址"] == "10.20.1.167"
    yi = [r for r in rows if r["站点"] == "乙站" and r["业务名称"] == "远动"]
    assert yi[0]["业务地址"] == "10.20.11.167", yi[0]["业务地址"]
    print("OK  test_batch_sites  (3 站点 × 20 = 60 个地址)")


def test_template_28():
    cats, prefix, policy = C.dict_to_config(C.template_28())
    assert prefix == 28
    res, issues = C.allocate_all(cats, prefix, policy)
    assert all(h.prefix == 28 for h in res)
    assert all(h.hosts == 14 for h in res)
    print("OK  test_template_28  (19 条全部 /28)")


def test_config_roundtrip(tmp="test_out.json"):
    cats, prefix, policy = C.dict_to_config(C.template_27())
    C.save_config(tmp, cats, prefix, policy)
    data = C.load_config(tmp)
    cats2, prefix2, policy2 = C.dict_to_config(data)
    assert prefix2 == prefix and policy2 == policy
    assert cats2[0].base_network == "10.20.1.160"
    assert cats2[0].items[0].name == "远动"
    assert cats2[0].items[0].offsets == [7, 8]
    import os
    os.remove(tmp)
    print("OK  test_config_roundtrip")


def test_suggest_base():
    """为非实时业务推荐一个不冲突的基准第四字节"""
    cats, prefix, policy = C.dict_to_config(C.template_27())
    used = sorted(C.used_last_octets(cats[0], prefix))
    base = C.suggest_base(used, cats[1].items, lo=0, hi=255)
    assert base is not None
    # 同一第三字节下（制造冲突场景），验证推荐值确实不冲突
    cats[1].base_network = f"10.20.1.{base}"
    res, issues = C.allocate_all(cats, prefix, policy)
    errs = [i for i in issues if i.level == "error"]
    assert not errs, f"基准 {base} 仍冲突：{[str(e) for e in errs]}"
    print(f"OK  test_suggest_base  (推荐非实时基准 = {base}，冲突归零)")


def test_auto_resolve():
    """自动避让：非实时基准 160 → 往后顺移"""
    cats, prefix, policy = C.dict_to_config(C.template_27())
    cats[1].base_network = "10.20.1.160"          # 制造冲突
    new_cats, notes = C.auto_resolve_conflicts(cats, prefix)
    assert notes, "应当产生调整说明"
    assert new_cats[1].base_network != "10.20.1.160"
    res, issues = C.allocate_all(new_cats, prefix, policy)
    assert not [i for i in issues if i.level == "error"]
    print(f"OK  test_auto_resolve  (非实时起始 IP → {new_cats[1].base_network})")


if __name__ == "__main__":
    test_parse_base()
    test_match_source_table()
    test_ip_format()
    test_28_mask()
    test_gateway_policy()
    test_carry()
    test_carry_deep()
    test_overflow_error()
    test_cross_subnet_warning()
    test_gateway_clash_warning()
    test_conflict_detect()
    test_no_conflict_when_separated()
    test_disable_item()
    test_batch_sites()
    test_template_28()
    test_config_roundtrip()
    test_suggest_base()
    test_auto_resolve()
    print("\n全部测试通过 ✓")
