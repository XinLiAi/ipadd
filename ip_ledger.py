# -*- coding: utf-8 -*-
"""
历史台账模块（SQLite）

记录每次导出的分配结果，支持按站点 / IP / 业务 / 时间查询，
并在下次分配时自动比对，防止地址重复分配。

数据模型：
    runs   一次导出 = 一条批次
    items  该批次算出的每个业务地址 = 一条明细
"""

from __future__ import annotations

import ipaddress
import os
import sqlite3
import sys
from dataclasses import dataclass
from datetime import datetime, timedelta
from typing import List, Dict, Optional, Tuple, Any

DB_NAME = "history.db"

SCHEMA = """
CREATE TABLE IF NOT EXISTS runs (
    id           INTEGER PRIMARY KEY AUTOINCREMENT,
    exported_at  TEXT    NOT NULL,
    site         TEXT    NOT NULL,
    src_doc      TEXT    DEFAULT '',
    out_path     TEXT    DEFAULT '',
    template     TEXT    DEFAULT '',
    mask         INTEGER,
    batch        INTEGER DEFAULT 0,
    filled_count INTEGER DEFAULT 0
);

CREATE TABLE IF NOT EXISTS items (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    run_id      INTEGER NOT NULL REFERENCES runs(id) ON DELETE CASCADE,
    table_title TEXT    DEFAULT '',
    category    TEXT    DEFAULT '',
    policy      TEXT    DEFAULT '',
    business    TEXT    DEFAULT '',
    net         TEXT    DEFAULT '',
    offset      INTEGER,
    ip          TEXT    NOT NULL,
    prefix      INTEGER,
    subnet      TEXT    DEFAULT '',
    gateway     TEXT    DEFAULT '',
    base_ip     TEXT    DEFAULT '',
    old_value   TEXT    DEFAULT ''
);

CREATE INDEX IF NOT EXISTS idx_items_ip    ON items(ip);
CREATE INDEX IF NOT EXISTS idx_items_run   ON items(run_id);
CREATE INDEX IF NOT EXISTS idx_items_biz   ON items(business);
CREATE INDEX IF NOT EXISTS idx_runs_site   ON runs(site);
CREATE INDEX IF NOT EXISTS idx_runs_time   ON runs(exported_at);
"""

# 明细字段（界面表格列顺序）
HIST_FIELDS = ["exported_at", "site", "table_title", "category", "business", "net",
               "offset", "ip", "prefix", "subnet", "gateway", "base_ip", "policy",
               "old_value"]


# ---------------------------------------------------------------- 工具

def db_path(dir_hint: str = "") -> str:
    """数据库路径：程序目录（打包后为 exe 同级目录）"""
    if dir_hint:
        return os.path.join(dir_hint, DB_NAME)
    base = os.path.dirname(os.path.abspath(__file__))
    if getattr(sys, "frozen", False):
        base = os.path.dirname(sys.executable)
    return os.path.join(base, DB_NAME)


def expand_ip_text(text: str) -> List[str]:
    """
    把批复单里的地址写法展开成单个 IP 列表。

        "54.100.134.98"          → ["54.100.134.98"]
        "54.100.134.103-104"     → ["54.100.134.103", "54.100.134.104"]
        "54.100.135.98、.118"    → ["54.100.135.98", "54.100.135.118"]
    """
    out: List[str] = []
    if not text:
        return out
    for part in str(text).split("、"):
        part = part.strip()
        if not part:
            continue
        if "-" in part and part.count(".") == 3:
            head, tail = part.rsplit(".", 1)
            if "-" in tail:
                a, b = tail.split("-", 1)
                try:
                    lo, hi = int(a), int(b)
                except ValueError:
                    out.append(part)
                    continue
                if 0 <= lo <= hi <= 255:
                    out.extend(f"{head}.{n}" for n in range(lo, hi + 1))
                    continue
        out.append(part)
    return out


def subnet_of(ip: str, prefix: Optional[int]) -> str:
    if not ip or not prefix:
        return ""
    try:
        return str(ipaddress.IPv4Network(f"{ip}/{prefix}", strict=False))
    except Exception:
        return ""


def gateway_of(ip: str, prefix: Optional[int]) -> str:
    if not ip or not prefix:
        return ""
    try:
        net = ipaddress.IPv4Network(f"{ip}/{prefix}", strict=False)
        hosts = list(net.hosts())
        return str(hosts[0]) if hosts else ""
    except Exception:
        return ""


@dataclass
class Item:
    """一条待归档的分配明细"""
    table_title: str = ""
    category: str = ""
    policy: str = ""
    business: str = ""
    net: str = ""
    offset: Optional[int] = None
    ip: str = ""
    prefix: Optional[int] = None
    subnet: str = ""
    gateway: str = ""
    base_ip: str = ""
    old_value: str = ""


@dataclass
class Conflict:
    """一条冲突：某个地址本次要用，但台账里已被别的站点占用"""
    ip: str = ""
    table_title: str = ""
    policy: str = ""
    business: str = ""
    old_site: str = ""
    old_business: str = ""
    old_at: str = ""
    # 便于界面直接引用
    new_site: str = ""
    new_policy: str = ""
    old_policy: str = ""
    created_at: str = ""


def collect_items(tables) -> List[Item]:
    """
    从批复单表对象里收集本次将填入的全部地址明细。
    同一 IP 在同一批次内只保留一条（批复单里多条隧道策略常复用同一地址）。
    """
    out: List[Item] = []
    seen = set()
    for t in tables:
        base = getattr(t, "suggested_base", "") or ""
        for r in t.rows:
            val = getattr(r, "new_value", "") or ""
            if not val:
                continue
            ips = expand_ip_text(val)
            for k, ip in enumerate(ips):
                if ip in seen:
                    continue
                seen.add(ip)
                net_tag = ("%d网" % (k + 1)) if len(ips) > 1 else "单网"
                out.append(Item(
                    table_title=t.title,
                    category=getattr(r, "category", "") or "",
                    policy=getattr(r, "policy", "") or "",
                    business=getattr(r, "business", "") or "",
                    net=net_tag,
                    offset=getattr(r, "offset", None),
                    ip=ip,
                    prefix=getattr(t, "suggested_mask", None) or 28,
                    subnet=subnet_of(ip, getattr(t, "suggested_mask", None) or 28),
                    gateway=gateway_of(ip, getattr(t, "suggested_mask", None) or 28),
                    base_ip=base,
                    old_value=getattr(r, "old_value", "") or "",
                ))
    return out


def self_conflicts(tables) -> List[Tuple[str, str, str, str]]:
    """
    检测本次导出内部的地址重复 —— 即【两个不同业务】占用了同一 IP。

    注意：批复单里多条隧道策略复用同一业务地址属正常结构
    （如 备调远动fes1~4、地调远动fes1~2、区调远动fes1~4 共 10 行都写
    远动的 103-104）。因此判重以「业务」为维度：
    同一业务的不同策略行复用同一地址 → 不算冲突；
    不同业务撞到同一地址 → 才算冲突。

    :return: [(IP, 表名, 策略名, 与之冲突的其他业务), ...]
    """
    owner: Dict[str, Tuple[str, str, str]] = {}    # ip -> (业务, 表名, 策略名)
    out: List[Tuple[str, str, str, str]] = []
    for t in tables:
        for r in t.rows:
            val = getattr(r, "new_value", "") or ""
            if not val:
                continue
            biz = getattr(r, "business", "") or "(未知业务)"
            policy = getattr(r, "policy", "") or ""
            for ip in expand_ip_text(val):
                prev = owner.get(ip)
                if prev is None:
                    owner[ip] = (biz, t.title, policy)
                elif prev[0] != biz:
                    out.append((ip, t.title, policy, f"{prev[1]}/{prev[0]}"))
    return out


# ---------------------------------------------------------------- 台账

class History:
    """历史台账数据库"""

    def __init__(self, path: str = ""):
        self.path = path or db_path()
        d = os.path.dirname(os.path.abspath(self.path))
        if d:
            os.makedirs(d, exist_ok=True)
        self.conn = sqlite3.connect(self.path)
        self.conn.row_factory = sqlite3.Row
        self.conn.execute("PRAGMA foreign_keys = ON")
        self.conn.executescript(SCHEMA)
        self.conn.commit()

    def close(self):
        try:
            self.conn.close()
        except Exception:
            pass

    # ---------------------------------------------------- 写入
    def record_run(self, site: str, src_doc: str = "", out_path: str = "",
                   template: str = "", mask: Optional[int] = None,
                   batch: bool = False, items: List[Item] = None,
                   exported_at: str = "") -> int:
        """写入一次导出记录及其明细，返回 run id"""
        items = items or []
        now = exported_at or datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        cur = self.conn.execute(
            "INSERT INTO runs (exported_at, site, src_doc, out_path, template,"
            " mask, batch, filled_count) VALUES (?,?,?,?,?,?,?,?)",
            (now, site, src_doc, out_path, template, mask,
             1 if batch else 0, len(items)))
        rid = cur.lastrowid
        if items:
            rows = [(rid, i.table_title, i.category, i.policy, i.business, i.net,
                     i.offset, i.ip, i.prefix, i.subnet, i.gateway, i.base_ip,
                     i.old_value) for i in items]
            self.conn.executemany(
                "INSERT INTO items (run_id, table_title, category, policy,"
                " business, net, offset, ip, prefix, subnet, gateway, base_ip,"
                " old_value) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)", rows)
        self.conn.commit()
        return rid

    def delete_runs(self, ids: List[int]) -> None:
        if not ids:
            return
        q = ",".join("?" * len(ids))
        self.conn.execute(f"DELETE FROM runs WHERE id IN ({q})", list(ids))
        self.conn.commit()

    def clear(self) -> None:
        self.conn.execute("DELETE FROM items")
        self.conn.execute("DELETE FROM runs")
        self.conn.commit()

    # ---------------------------------------------------- 查询
    def sites(self) -> List[str]:
        return [r[0] for r in self.conn.execute(
            "SELECT DISTINCT site FROM runs ORDER BY site").fetchall()]

    def businesses(self) -> List[str]:
        return [r[0] for r in self.conn.execute(
            "SELECT DISTINCT business FROM items WHERE business <> ''"
            " ORDER BY business").fetchall()]

    def list_runs(self) -> List[sqlite3.Row]:
        return self.conn.execute(
            "SELECT r.*, (SELECT COUNT(*) FROM items i WHERE i.run_id = r.id)"
            " AS n_items FROM runs r ORDER BY r.exported_at DESC, r.id DESC"
        ).fetchall()

    def run_items(self, rid: int) -> List[sqlite3.Row]:
        return self.conn.execute(
            "SELECT i.*, r.site, r.exported_at FROM items i "
            "JOIN runs r ON r.id = i.run_id WHERE i.run_id = ? ORDER BY i.id",
            (rid,)).fetchall()

    def search(self, keyword: str = "", site: str = "", business: str = "",
               days: Optional[int] = None) -> List[sqlite3.Row]:
        """
        查询明细。keyword 同时匹配 IP、业务、策略、站点。
        site / business 传「（全部站点）」「（全部业务）」或空表示不限。
        """
        sql = ("SELECT i.*, r.site, r.exported_at FROM items i "
               "JOIN runs r ON r.id = i.run_id WHERE 1=1")
        args: List[Any] = []

        if keyword:
            sql += (" AND (i.ip LIKE ? OR i.business LIKE ? OR i.policy LIKE ?"
                    " OR r.site LIKE ? OR i.table_title LIKE ?)")
            args.extend([f"%{keyword}%"] * 5)
        if site and not site.startswith("（"):
            sql += " AND r.site = ?"
            args.append(site)
        if business and not business.startswith("（"):
            sql += " AND i.business = ?"
            args.append(business)
        if days:
            since = (datetime.now() - timedelta(days=days)).strftime("%Y-%m-%d %H:%M:%S")
            sql += " AND r.exported_at >= ?"
            args.append(since)

        sql += " ORDER BY r.exported_at DESC, r.id DESC, i.id"
        return self.conn.execute(sql, args).fetchall()

    def occupied(self, ips: List[str]) -> Dict[str, List[Tuple[str, str, str]]]:
        """
        批量查询一组 IP 的历史占用情况。
        :return: {ip: [(站点, 业务, 导出时间), ...]}
        """
        out: Dict[str, List[Tuple[str, str, str]]] = {ip: [] for ip in ips}
        if not ips:
            return out
        uniq = list(dict.fromkeys(ips))
        for i in range(0, len(uniq), 400):
            chunk = uniq[i:i + 400]
            q = ",".join("?" * len(chunk))
            rows = self.conn.execute(
                f"SELECT i.ip, r.site, i.business, r.exported_at FROM items i "
                f"JOIN runs r ON r.id = i.run_id WHERE i.ip IN ({q})",
                chunk).fetchall()
            for r in rows:
                out.setdefault(r["ip"], []).append(
                    (r["site"], r["business"], r["exported_at"]))
        return out

    def duplicates(self) -> List[Tuple[str, int, List[str]]]:
        """台账里被多个站点占用的地址：(IP, 站点数, [站点...])"""
        rows = self.conn.execute(
            "SELECT i.ip, COUNT(DISTINCT r.site) AS n, "
            "GROUP_CONCAT(DISTINCT r.site) AS sites "
            "FROM items i JOIN runs r ON r.id = i.run_id "
            "GROUP BY i.ip HAVING n > 1 ORDER BY n DESC, i.ip").fetchall()
        return [(r["ip"], r["n"], (r["sites"] or "").split(",")) for r in rows]

    def find_conflicts(self, tables, exclude_site: str = "",
                       new_site: str = "") -> List[Conflict]:
        """
        把待填地址与台账比对，找出被【其它站点】占用的地址。

        :param exclude_site: 该站点的历史记录不算冲突（允许重复导出）
        """
        items = collect_items(tables)
        if not items:
            return []
        ips = [i.ip for i in items]
        occ = self.occupied(ips)

        out: List[Conflict] = []
        seen = set()
        for it in items:
            for old_site, old_biz, when in occ.get(it.ip, []):
                if exclude_site and old_site == exclude_site:
                    continue
                if it.ip in seen:
                    continue
                seen.add(it.ip)
                out.append(Conflict(
                    ip=it.ip, table_title=it.table_title, policy=it.policy,
                    business=it.business, old_site=old_site,
                    old_business=old_biz, old_at=when,
                    new_site=new_site, new_policy=it.policy,
                    old_policy=it.policy, created_at=when))
        return out

    def export_rows(self) -> Tuple[List[str], List[list]]:
        """导出全部明细，返回 (表头, 行)"""
        headers = ["导出时间", "站点", "纵密表", "大类", "策略名称", "业务",
                   "网别", "偏移", "业务地址", "掩码", "所属网段", "网关",
                   "起始IP", "原值"]
        rows = self.conn.execute(
            "SELECT r.exported_at, r.site, i.table_title, i.category, i.policy,"
            " i.business, i.net, i.offset, i.ip, i.prefix, i.subnet, i.gateway,"
            " i.base_ip, i.old_value "
            "FROM items i JOIN runs r ON r.id = i.run_id "
            "ORDER BY r.exported_at DESC, r.id DESC, i.id").fetchall()
        keys = ("exported_at", "site", "table_title", "category", "policy",
                "business", "net", "offset", "ip", "prefix", "subnet",
                "gateway", "base_ip", "old_value")
        return headers, [[("" if r[k] is None else
                           (f"/{r[k]}" if k == "prefix" and r[k] else r[k]))
                          for k in keys] for r in rows]

    def stats(self) -> dict:
        n_run = self.conn.execute("SELECT COUNT(*) FROM runs").fetchone()[0]
        n_item = self.conn.execute("SELECT COUNT(*) FROM items").fetchone()[0]
        n_site = self.conn.execute("SELECT COUNT(DISTINCT site) FROM runs").fetchone()[0]
        n_ip = self.conn.execute("SELECT COUNT(DISTINCT ip) FROM items").fetchone()[0]
        last = self.conn.execute(
            "SELECT exported_at FROM runs ORDER BY exported_at DESC LIMIT 1").fetchone()
        return {"批次数": n_run, "地址条数": n_item, "站点数": n_site,
                "不同IP数": n_ip,
                "最近导出": last["exported_at"] if last else "—"}
