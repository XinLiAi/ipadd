# -*- coding: utf-8 -*-
"""
IP 地址自动分配 —— 核心计算模块（单业务 IP 模式）

规则来源：《业务地址计算工具.xlsx》
    所得值 = 基准地址第四字节(B列) + 偏移(C列/D列)
    业务IP = 起始IP前三位 . 所得值

即：给定起始 IP 如 10.20.1.160（第四字节 160 为基准），
    某业务偏移为 7 时 → 业务 IP = 10.20.1.167

    ⚠ 注意：167 落在最后一位，不是第三字节。
      这是在一个子网内挑出「单个业务地址」，不是划分网段。

掩码位数（/27、/28）的作用：
    仅用于标出该业务 IP 所属的网段信息（网络号、网关、广播、可用范围），
    并对「偏移超出子网容量」给出提示。

纯 Python 实现，不依赖任何 GUI 库，可单独测试 / 命令行调用。
"""

from __future__ import annotations

import ipaddress
import json
import os
from dataclasses import dataclass, field, asdict
from typing import List, Optional, Tuple, Dict, Any

# ---------------------------------------------------------------- 常量

DEFAULT_PREFIX = 27              # 默认掩码位数
DEFAULT_BASE = "10.20.1.160"     # 默认起始 IP（第四字节 160 为基准）

# 网关取值策略
GATEWAY_FIRST = "first"      # 子网首个可用地址（默认，电力行业常用）
GATEWAY_LAST = "last"        # 子网末个可用地址
GATEWAY_NONE = "none"        # 不显示网关

GATEWAY_POLICY_LABEL = {
    GATEWAY_FIRST: "子网首个可用地址",
    GATEWAY_LAST: "子网末个可用地址",
    GATEWAY_NONE: "不显示网关",
}


# ---------------------------------------------------------------- 数据结构

@dataclass
class BusinessItem:
    """一条业务定义（对应原表一行）"""
    name: str                                        # 业务名称，如 "远动"
    offsets: List[int] = field(default_factory=list)  # 偏移序号，如 [7, 8]
    remark: str = ""                                 # 备注
    enabled: bool = True                             # 是否参与本次分配

    @property
    def dual(self) -> bool:
        """是否双网（两个及以上偏移 = 双网冗余）"""
        return len([o for o in self.offsets if o is not None]) >= 2

    def to_dict(self) -> dict:
        return asdict(self)

    @staticmethod
    def from_dict(d: dict) -> "BusinessItem":
        return BusinessItem(
            name=d.get("name", ""),
            offsets=[int(o) for o in d.get("offsets", []) if o is not None],
            remark=d.get("remark", ""),
            enabled=bool(d.get("enabled", True)),
        )


@dataclass
class Category:
    """一个业务大类（如 实时业务 / 非实时业务）"""
    name: str
    base_network: str = DEFAULT_BASE                 # 该大类的起始 IP（B 列基准）
    items: List[BusinessItem] = field(default_factory=list)

    def to_dict(self) -> dict:
        return {
            "name": self.name,
            "base_network": self.base_network,
            "items": [i.to_dict() for i in self.items],
        }

    @staticmethod
    def from_dict(d: dict) -> "Category":
        return Category(
            name=d.get("name", ""),
            base_network=d.get("base_network", DEFAULT_BASE),
            items=[BusinessItem.from_dict(x) for x in d.get("items", [])],
        )


@dataclass
class HostIP:
    """一个业务 IP 的分配结果"""
    category: str = ""       # 所属大类
    business: str = ""       # 业务名称
    seq: int = 0             # 第几张网（1=A网/主用，2=B网/备用，0=单网）
    offset: int = 0          # 本次使用的偏移序号
    ip: str = ""             # ★ 业务 IP，如 10.20.1.167
    prefix: int = 27         # 掩码位数
    netmask: str = ""        # 子网掩码 255.255.255.224
    subnet: str = ""         # 所属网段 10.20.1.160/27
    gateway: str = ""        # 网关
    first_host: str = ""     # 所属网段首个可用地址
    last_host: str = ""      # 所属网段末个可用地址
    broadcast: str = ""      # 所属网段广播地址
    hosts: int = 0           # 所属网段可用主机数
    in_base_subnet: bool = True   # 是否仍落在「基准地址所在子网」内
    remark: str = ""

    @property
    def last_octet(self) -> int:
        """业务 IP 的第四字节（即原表「所得值」）"""
        try:
            return int(self.ip.split(".")[3])
        except Exception:
            return -1

    @property
    def cidr(self) -> str:
        return f"{self.ip}/{self.prefix}"

    def as_row(self) -> list:
        """导出表格用的一行"""
        return [
            self.category,
            self.business,
            f"{self.seq}网" if self.seq else "单网",
            self.offset,
            self.ip,                       # ★ 业务地址
            f"/{self.prefix}",
            self.netmask,
            self.subnet,                   # 所属网段
            self.gateway,
            f"{self.first_host} - {self.last_host}",
            self.broadcast,
            self.hosts,
            "是" if self.in_base_subnet else "否（已跨段）",
            self.remark,
        ]

    @staticmethod
    def header() -> list:
        return ["业务大类", "业务名称", "网别", "偏移", "业务地址",
                "掩码位", "子网掩码", "所属网段", "网关", "网段可用范围",
                "广播地址", "网段可用数", "在基准段内", "备注"]


# 兼容旧名（导出模块等处引用）
Subnet = HostIP


@dataclass
class Issue:
    """一条校验问题"""
    level: str    # error / warning
    message: str

    def __str__(self):
        tag = "错误" if self.level == "error" else "警告"
        return f"[{tag}] {self.message}"


# ---------------------------------------------------------------- 核心计算

def parse_base(text: str) -> Optional[Tuple[int, int, int, int]]:
    """
    解析起始 IP，返回四个字节 (A, B, C, D)，其中 D 即原表 B 列基准值。

    接受写法：
        "10.20.1.160"      → (10, 20, 1, 160)   完整起始 IP（推荐）
        "10.20.1.160/27"   → (10, 20, 1, 160)   带掩码也认
        "10.20.1"          → (10, 20, 1, 0)     只填前三位，基准按 0
    """
    if not text:
        return None
    t = str(text).strip().split("/")[0]
    parts = t.split(".")
    if len(parts) == 3:
        parts = parts + ["0"]              # 只给前三位时，基准默认 0
    if len(parts) != 4:
        return None
    try:
        vals = tuple(int(p) for p in parts)
    except ValueError:
        return None
    if any(v < 0 or v > 255 for v in vals):
        return None
    return vals  # type: ignore


def calc_host(base: Tuple[int, int, int, int], offset: int,
              prefix: int = DEFAULT_PREFIX,
              gateway_policy: str = GATEWAY_FIRST) -> HostIP:
    """
    按「基准第四字节 + 偏移」算出一个业务 IP。

    :param base:   起始 IP 四字节 (A, B, C, D)，如 (10, 20, 1, 160)
    :param offset: 偏移序号，如 7
    :param prefix: 掩码位数，用于标注所属网段
    :param gateway_policy: 网关策略
    :return: HostIP
    """
    if offset < 0:
        raise ValueError(f"偏移 {offset} 非法（不能为负）")
    if not 0 <= prefix <= 32:
        raise ValueError(f"掩码位数 {prefix} 非法")

    # ---- 1. 业务 IP = 基准第四字节 + 偏移，溢出则逐字节进位 ----
    a, b, c, d = base
    total = d + offset
    if total > 255:
        carry, d2 = divmod(total, 256)      # 向第三字节进位
        c2 = c + carry
        if c2 > 255:
            carry2, c2 = divmod(c2, 256)
            b2 = b + carry2
            if b2 > 255:
                raise ValueError(f"偏移 {offset} 使地址超出范围（第二字节溢出）")
        else:
            b2 = b
        ip_int_base = (a, b2, c2, d2)
    else:
        ip_int_base = (a, b, c, total)

    ip = ".".join(str(x) for x in ip_int_base)

    # ---- 2. 该 IP 所属的网段（按掩码计算）----
    net = ipaddress.IPv4Network(f"{ip}/{prefix}", strict=False)
    hosts = list(net.hosts())

    if gateway_policy == GATEWAY_FIRST:
        gw = str(hosts[0]) if hosts else ""
    elif gateway_policy == GATEWAY_LAST:
        gw = str(hosts[-1]) if hosts else ""
    else:
        gw = ""

    # ---- 3. 基准地址本身所在的子网（用于判断是否跨段）----
    base_ip = ".".join(str(x) for x in base)
    base_net = ipaddress.IPv4Network(f"{base_ip}/{prefix}", strict=False)

    return HostIP(
        category="", business="", seq=0, offset=offset,
        ip=ip, prefix=prefix,
        netmask=str(net.netmask),
        subnet=str(net),
        gateway=gw,
        first_host=str(hosts[0]) if hosts else "",
        last_host=str(hosts[-1]) if hosts else "",
        broadcast=str(net.broadcast_address) if prefix <= 30 else "",
        hosts=len(hosts),
        in_base_subnet=(net == base_net),
    )


def allocate_category(cat: Category,
                      prefix: int = DEFAULT_PREFIX,
                      gateway_policy: str = GATEWAY_FIRST) -> Tuple[List[HostIP], List[Issue]]:
    """为一个业务大类生成全部业务 IP，同时返回校验问题"""
    results: List[HostIP] = []
    issues: List[Issue] = []

    base = parse_base(cat.base_network)
    if base is None:
        issues.append(Issue("error", f"【{cat.name}】起始 IP 格式不正确：{cat.base_network}"))
        return results, issues

    for item in cat.items:
        if not item.enabled:
            continue
        if not item.offsets:
            issues.append(Issue("warning", f"【{cat.name}/{item.name}】未设置偏移序号，已跳过"))
            continue

        for seq, off in enumerate(item.offsets, start=1):
            try:
                h = calc_host(base, off, prefix, gateway_policy)
            except ValueError as e:
                issues.append(Issue("error", f"【{cat.name}/{item.name}】{e}"))
                continue

            h.category = cat.name
            h.business = item.name
            h.seq = seq if len(item.offsets) > 1 else 0
            h.remark = item.remark

            # 业务 IP 与网关撞车提示
            if h.gateway and h.ip == h.gateway:
                issues.append(Issue(
                    "warning",
                    f"【{cat.name}/{item.name}】业务地址 {h.ip} 与网关地址相同，请调整偏移"))

            # 跨出基准子网提示
            if not h.in_base_subnet:
                issues.append(Issue(
                    "warning",
                    f"【{cat.name}/{item.name}】偏移 {off} 超出基准子网"
                    f"（/{prefix} 每段 {h.hosts} 个可用地址），"
                    f"{h.ip} 已落入 {h.subnet}"))

            results.append(h)

    return results, issues


def allocate_all(categories: List[Category],
                 prefix: int = DEFAULT_PREFIX,
                 gateway_policy: str = GATEWAY_FIRST
                 ) -> Tuple[List[HostIP], List[Issue]]:
    """
    生成全部大类的业务 IP，并做**跨大类**的 IP 重复检测。

    说明：实时业务与非实时业务通常分属不同 VLAN，起始 IP 前三位不同、互不冲突；
    但若两者起始 IP 填成一样（如原表基准同为 160），偏移重叠就会撞 IP，必须报出来。
    """
    all_res: List[HostIP] = []
    all_issues: List[Issue] = []
    for cat in categories:
        res, iss = allocate_category(cat, prefix, gateway_policy)
        all_res.extend(res)
        all_issues.extend(iss)

    # 全局重复检测：同一业务 IP 不得分给两个业务
    seen: Dict[str, HostIP] = {}
    for h in all_res:
        if h.ip in seen:
            o = seen[h.ip]
            all_issues.append(Issue(
                "error",
                f"地址冲突：{o.category}/{o.business} 与 "
                f"{h.category}/{h.business} 的业务地址都是 {h.ip}"))
        else:
            seen[h.ip] = h

    # 去重（保留首次出现）
    dedup: List[Issue] = []
    for i in all_issues:
        if i.message not in [d.message for d in dedup]:
            dedup.append(i)

    return all_res, dedup


# ---------------------------------------------------------------- 智能避让

def used_last_octets(cat: Category, prefix: int = DEFAULT_PREFIX) -> set:
    """统计某个大类已占用的第四字节"""
    base = parse_base(cat.base_network)
    if base is None:
        return set()
    out = set()
    for item in cat.items:
        if not item.enabled:
            continue
        for off in item.offsets:
            try:
                out.add(calc_host(base, off, prefix).last_octet)
            except ValueError:
                pass
    return out


def suggest_base(used: List[int], items: List[BusinessItem],
                 lo: int = 0, hi: int = 255, step: int = 32) -> Optional[int]:
    """
    为一个大类推荐一个「不与已占用地址冲突」的基准第四字节。

    :param used:  已被其它大类占用的第四字节列表
    :param items: 本大类全部业务（含 offsets）
    :param lo/hi: 基准搜索区间
    :param step:  推荐值的对齐步长（32 对齐 /27 子网边界，便于扩容）
    """
    used_set = set(used)
    offsets = []
    for it in items:
        if it.enabled:
            offsets.extend(it.offsets)
    if not offsets:
        return None
    max_off = max(offsets)

    cands = list(range(lo, hi + 1, step)) + list(range(lo, hi + 1, 1))
    for base in cands:
        if base + max_off > 255:
            continue
        if all((base + o) not in used_set for o in offsets):
            return base
    return None


def auto_resolve_conflicts(categories: List[Category],
                           prefix: int = DEFAULT_PREFIX) -> Tuple[List[Category], List[str]]:
    """
    自动为「与前面大类冲突」的大类重新挑选基准地址，返回调整说明。

    典型场景：实时业务与非实时业务都填了 10.20.1.160，
    偏移重叠 → 撞 IP。本函数把后面大类的基准往后顺移到不冲突的位置。
    """
    import copy
    cats = copy.deepcopy(categories)
    notes: List[str] = []
    used: set = set()

    for idx, cat in enumerate(cats):
        own = used_last_octets(cat, prefix)
        if idx > 0 and (own & used):
            base_t = parse_base(cat.base_network)
            if base_t is None:
                continue
            a, b, c, old_d = base_t
            step = 32 if prefix == 27 else (16 if prefix == 28 else 8)
            new_d = suggest_base(sorted(used), cat.items, lo=old_d, hi=255, step=step)
            if new_d is None:
                new_d = suggest_base(sorted(used), cat.items, lo=0, hi=255, step=step)
            if new_d is not None:
                cat.base_network = f"{a}.{b}.{c}.{new_d}"
                notes.append(
                    f"【{cat.name}】基准地址 {old_d} → {new_d}"
                    f"（起始 IP 改为 {a}.{b}.{c}.{new_d}，避开 {len(own & used)} 处冲突）")
                own = used_last_octets(cat, prefix)
        used |= own

    return cats, notes


# ---------------------------------------------------------------- 批量（多站点）

def allocate_sites(site_specs: List[Tuple[str, List[str]]],
                   categories: List[Category],
                   prefix: int = DEFAULT_PREFIX,
                   gateway_policy: str = GATEWAY_FIRST
                   ) -> Tuple[List[dict], List[Issue]]:
    """
    批量为多个站点生成业务 IP。

    :param site_specs: [(站点名称, [各大类起始IP...]), ...]
                       起始 IP 可填 "10.20.1.160"；留空则用模板默认值
    """
    rows: List[dict] = []
    issues: List[Issue] = []

    for site, ips in site_specs:
        site_cats = []
        for idx, cat in enumerate(categories):
            ip = ips[idx] if idx < len(ips) else ""
            c = Category(
                name=cat.name,
                base_network=ip.strip() if ip and ip.strip() else cat.base_network,
                items=[BusinessItem(i.name, list(i.offsets), i.remark, i.enabled)
                       for i in cat.items],
            )
            site_cats.append(c)

        res, iss = allocate_all(site_cats, prefix, gateway_policy)
        for r in res:
            d = {"站点": site}
            d.update(zip(HostIP.header(), r.as_row()))
            rows.append(d)
        for i in iss:
            issues.append(Issue(i.level, f"【站点 {site}】{i.message}"))

    return rows, issues


# ---------------------------------------------------------------- 预置模板

def template_27() -> Dict[str, Any]:
    """预置模板：/27（照搬工作表「27位计算」）

    起始 IP 说明：
      基准第四字节 = 160（原表 B 列），前三位取 10.20.1 / 10.20.2。
      实时与非实时分属不同 VLAN，故第三字节分别为 1、2，避免撞地址。
    """
    return {
        "prefix": 27,
        "gateway_policy": GATEWAY_FIRST,
        "categories": [
            {
                "name": "实时业务",
                "base_network": "10.20.1.160",
                "items": [
                    {"name": "远动",     "offsets": [7, 8],   "remark": "远动双网",    "enabled": True},
                    {"name": "保护",     "offsets": [9, 10],  "remark": "保护双网",    "enabled": True},
                    {"name": "PMU",      "offsets": [15, 16], "remark": "同步相量双网", "enabled": True},
                    {"name": "网安",     "offsets": [22],     "remark": "单网",       "enabled": True},
                    {"name": "稳控",     "offsets": [24, 25], "remark": "稳控双网",    "enabled": True},
                    {"name": "一次调频", "offsets": [3],      "remark": "单网",       "enabled": True},
                    {"name": "宽频",     "offsets": [1, 2],   "remark": "宽频双网",    "enabled": True},
                ],
            },
            {
                "name": "非实时业务",
                "base_network": "10.20.2.160",
                "items": [
                    {"name": "电量",         "offsets": [10], "remark": "单网", "enabled": True},
                    {"name": "保信（保护）",  "offsets": [5],  "remark": "单网", "enabled": True},
                    {"name": "光功率",       "offsets": [20], "remark": "单网", "enabled": True},
                    {"name": "风功率",       "offsets": [19], "remark": "单网", "enabled": True},
                    {"name": "水情",         "offsets": [7],  "remark": "单网", "enabled": True},
                    {"name": "网络安全监测",  "offsets": [22], "remark": "单网", "enabled": True},
                    {"name": "现货市场",     "offsets": [9],  "remark": "单网", "enabled": True},
                    {"name": "一次设备",     "offsets": [3],  "remark": "单网", "enabled": True},
                ],
            },
        ],
    }


def template_28() -> Dict[str, Any]:
    """预置模板：/28（照搬工作表「28位」）"""
    return {
        "prefix": 28,
        "gateway_policy": GATEWAY_FIRST,
        "categories": [
            {
                "name": "实时业务",
                "base_network": "10.20.1.0",
                "items": [
                    {"name": "远动",    "offsets": [7, 8],  "remark": "远动双网", "enabled": True},
                    {"name": "保护",    "offsets": [9, 10], "remark": "保护双网", "enabled": True},
                    {"name": "PMU",     "offsets": [5, 6],  "remark": "同步相量双网", "enabled": True},
                    {"name": "网安",    "offsets": [2],     "remark": "单网", "enabled": True},
                    {"name": "稳控",    "offsets": [3, 4],  "remark": "稳控双网", "enabled": True},
                    {"name": "宽频装置", "offsets": [1, 11], "remark": "宽频双网", "enabled": True},
                    {"name": "安控",     "offsets": [],       "remark": "批复单有此策略，待分配偏移", "enabled": False},
                ],
            },
            {
                "name": "非实时业务",
                "base_network": "10.20.2.0",
                "items": [
                    {"name": "网安",     "offsets": [2],     "remark": "单网（原表 D 列 22 为 27 位掩码用）", "enabled": True},
                    {"name": "电量",     "offsets": [10],    "remark": "单网", "enabled": True},
                    {"name": "保信",     "offsets": [5],     "remark": "单网", "enabled": True},
                    {"name": "光功率",   "offsets": [6],     "remark": "单网", "enabled": True},
                    {"name": "水情",     "offsets": [7],     "remark": "单网", "enabled": True},
                    {"name": "风功率",   "offsets": [8],     "remark": "单网", "enabled": True},
                    {"name": "现货市场", "offsets": [9],     "remark": "单网", "enabled": True},
                    {"name": "同步时钟", "offsets": [],       "remark": "批复单有此策略，待分配偏移", "enabled": False},
                ],
            },
        ],
    }


TEMPLATES = {
    "27位掩码（标准模板）": template_27,
    "28位掩码（备用模板）": template_28,
}


# ---------------------------------------------------------------- 配置存取

def config_to_dict(categories: List[Category],
                   prefix: int = DEFAULT_PREFIX,
                   gateway_policy: str = GATEWAY_FIRST) -> dict:
    return {
        "prefix": prefix,
        "gateway_policy": gateway_policy,
        "categories": [c.to_dict() for c in categories],
    }


def dict_to_config(data: dict) -> Tuple[List[Category], int, str]:
    cats = [Category.from_dict(c) for c in data.get("categories", [])]
    prefix = int(data.get("prefix", DEFAULT_PREFIX))
    policy = data.get("gateway_policy", GATEWAY_FIRST)
    return cats, prefix, policy


def load_config(path: str) -> dict:
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def save_config(path: str,
                categories: List[Category],
                prefix: int = DEFAULT_PREFIX,
                gateway_policy: str = GATEWAY_FIRST) -> None:
    d = os.path.dirname(os.path.abspath(path))
    if d:
        os.makedirs(d, exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(config_to_dict(categories, prefix, gateway_policy),
                  f, ensure_ascii=False, indent=2)
