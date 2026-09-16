# -*- coding: utf-8 -*-
"""
调度数据网接入地址批复单 —— 业务 IP 自动填充工具

界面结构：
    左：业务大类页签（起始 IP + 业务偏移表，可增删改）
    右：批复单填充（各纵密表起始 IP + 实时预览 + 导出 Word）

运行：  python3 app_gui.py
打包：  双击 build.bat（生成 dist/IP地址批复单填充工具.exe）
"""

from __future__ import annotations

import os
import sys
import traceback
from datetime import datetime
from typing import Tuple

# Qt binding: PySide6 on Win10+, PySide2 on Win7 / Server 2008 R2.
# See qt_compat.py for why (Qt 6 needs Windows 10 1809 or later).
#
# qt_compat already logs the reason to error.log and pops up a native
# Windows message box. Here we just exit quietly, so the user does not
# also get PyInstaller's cryptic "Unhandled exception in script" dialog.
try:
    import qt_compat as QC
    from qt_compat import (
        Qt, QDate,
        QFont, QColor, QBrush,
        QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QGridLayout,
        QTabWidget, QTableWidget, QTableWidgetItem, QHeaderView, QAbstractItemView,
        QPushButton, QLabel, QLineEdit, QSpinBox, QComboBox, QGroupBox, QSplitter,
        QFileDialog, QMessageBox, QListWidget, QListWidgetItem, QCheckBox, QStatusBar,
        QInputDialog, QDateEdit,
    )
except ImportError:
    # Message already shown by qt_compat._report_failure(); exit cleanly.
    sys.exit(1)

import ip_core as C
from ip_core import Category, BusinessItem, HostIP, Issue
import ip_word as W
import ip_export as X
import ip_ledger as H

APP_NAME = "调度数据网批复单业务IP填充工具"
APP_VER = "1.1"

# ---------------------------------------------------------------- 样式

QSS = """
QWidget { font-family: "Microsoft YaHei", "微软雅黑", "PingFang SC", "Noto Sans CJK SC", sans-serif; font-size: 12px; }
QMainWindow, QDialog { background: #F5F6F8; }
QGroupBox {
    border: 1px solid #D6DAE0; border-radius: 6px; margin-top: 14px;
    background: #FFFFFF; font-weight: bold; padding-top: 4px;
}
QGroupBox::title { subcontrol-origin: margin; left: 12px; padding: 0 6px; color: #2C3E50; }
QTableWidget { border: 1px solid #D6DAE0; border-radius: 4px; background: #FFFFFF;
               gridline-color: #E5E8EC; selection-background-color: #CCE4FF; }
QHeaderView::section { background: #E8EDF3; border: 0; border-right: 1px solid #D6DAE0;
                       border-bottom: 1px solid #D6DAE0; padding: 5px; font-weight: bold; }
QPushButton { background: #FFFFFF; border: 1px solid #C4CBD4; border-radius: 4px;
              padding: 5px 12px; min-height: 20px; }
QPushButton:hover { background: #EDF3FB; border-color: #4A90D9; }
QPushButton:pressed { background: #DCE8F7; }
QPushButton#primary { background: #2F6FD0; color: #FFFFFF; border: 1px solid #2F6FD0; font-weight: bold; }
QPushButton#primary:hover { background: #3C7FE0; }
QPushButton#primary:pressed { background: #2559AB; }
QPushButton#accent { background: #1E8449; color: #FFFFFF; border: 1px solid #1E8449; font-weight: bold; }
QPushButton#accent:hover { background: #229954; }
QPushButton#danger { color: #C0392B; }
QLineEdit, QSpinBox, QComboBox { border: 1px solid #C4CBD4; border-radius: 4px;
                                  padding: 4px 6px; background: #FFFFFF; min-height: 18px; }
QLineEdit:focus, QSpinBox:focus, QComboBox:focus { border-color: #4A90D9; }
QTabWidget::pane { border: 1px solid #D6DAE0; border-top: 0; background: #FFFFFF; }
QTabBar::tab { background: #E8EDF3; border: 1px solid #D6DAE0; padding: 6px 16px; margin-right: 2px; }
QTabBar::tab:selected { background: #FFFFFF; border-bottom-color: #FFFFFF; color: #2F6FD0; font-weight: bold; }
QListWidget { border: 1px solid #D6DAE0; border-radius: 4px; background: #FFFFFF; }
QStatusBar { background: #E8EDF3; }
"""

BIZ_HEADERS = ["启用", "业务种类", "偏移1 (C列)", "偏移2 (D列)", "备注"]

HIST_HEADERS = ["导出时间", "站点", "纵密表", "大类", "业务", "网别", "偏移",
                "业务地址", "掩码", "所属网段", "网关", "起始IP", "策略名称", "原值"]
HIST_FIELDS = ["exported_at", "site", "table_title", "category", "business", "net",
               "offset", "ip", "prefix", "subnet", "gateway", "base_ip", "policy",
               "old_value"]
RUN_HEADERS = ["批次ID", "导出时间", "站点", "源批复单", "业务模板", "掩码",
               "地址条数", "输出文件"]

# 历史明细表显示的列（对应 History.search 返回的字段名）
RUN_HEADERS = ["批次ID", "导出时间", "站点", "源批复单", "业务模板", "掩码",
               "地址条数", "输出文件"]
PREV_HEADERS = ["纵密表", "隧道", "策略名称", "匹配业务", "原本端业务IP", "将填入", "历史占用"]

CLR_NEW = QColor("#1E8449")      # 补填
CLR_CHG = QColor("#B9770E")      # 覆盖且变化
CLR_SAME = QColor("#7F8C8D")     # 无变化
CLR_SKIP = QColor("#C0392B")     # 跳过


def _center_item(text: str, editable: bool = True, color: QColor = None) -> QTableWidgetItem:
    it = QTableWidgetItem(str(text))
    it.setTextAlignment(Qt.AlignCenter)
    if not editable:
        it.setFlags(it.flags() & ~Qt.ItemIsEditable)
    if color is not None:
        it.setForeground(QBrush(color))
    return it


# ---------------------------------------------------------------- 业务偏移表

class CategoryPanel(QWidget):
    """一个大类的配置面板：起始 IP + 业务偏移表"""

    def __init__(self, cat: Category, parent=None):
        super().__init__(parent)
        self.cat = cat
        self._build()

    def _build(self):
        lay = QVBoxLayout(self)
        lay.setContentsMargins(8, 10, 8, 8)
        lay.setSpacing(8)

        hl = QHBoxLayout()
        hl.addWidget(QLabel("起始 IP（B列基准）："))
        self.ed_base = QLineEdit(self.cat.base_network)
        self.ed_base.setFixedWidth(150)
        self.ed_base.setToolTip(
            "填完整起始 IP，其第四字节即基准值（原表 B 列）\n"
            "例：10.20.1.160 → 基准 = 160，偏移 7 → 业务地址 10.20.1.167\n"
            "也可只填前三位 10.20.1，基准按 0 计")
        hl.addWidget(self.ed_base)
        hl.addWidget(QLabel("  业务地址 = 前三段 . (基准 + 偏移)"))
        hl.addStretch()
        lay.addLayout(hl)

        self.tbl = QTableWidget(0, len(BIZ_HEADERS))
        self.tbl.setHorizontalHeaderLabels(BIZ_HEADERS)
        self.tbl.verticalHeader().setVisible(False)
        self.tbl.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.tbl.setAlternatingRowColors(True)
        hh = self.tbl.horizontalHeader()
        hh.setSectionResizeMode(QHeaderView.ResizeToContents)
        hh.setSectionResizeMode(1, QHeaderView.Stretch)
        hh.setSectionResizeMode(4, QHeaderView.Stretch)
        lay.addWidget(self.tbl, 1)

        hb = QHBoxLayout()
        self.btn_add = QPushButton("＋ 新增业务")
        self.btn_del = QPushButton("－ 删除选中")
        self.btn_del.setObjectName("danger")
        self.btn_add.clicked.connect(self.add_row)
        self.btn_del.clicked.connect(self.del_row)
        hb.addWidget(self.btn_add)
        hb.addWidget(self.btn_del)
        hb.addStretch()
        lay.addLayout(hb)

        self.reload(self.cat)

    def reload(self, cat: Category):
        self.cat = cat
        self.ed_base.setText(cat.base_network)
        self.tbl.setRowCount(0)
        for it in cat.items:
            self._append_row(it)

    def _append_row(self, it: BusinessItem):
        r = self.tbl.rowCount()
        self.tbl.insertRow(r)

        cb = QTableWidgetItem()
        cb.setCheckState(Qt.Checked if it.enabled else Qt.Unchecked)
        cb.setTextAlignment(Qt.AlignCenter)
        self.tbl.setItem(r, 0, cb)

        name = QTableWidgetItem(it.name)
        name.setTextAlignment(Qt.AlignCenter)
        self.tbl.setItem(r, 1, name)

        o1 = it.offsets[0] if len(it.offsets) > 0 else None
        o2 = it.offsets[1] if len(it.offsets) > 1 else None
        self.tbl.setItem(r, 2, _center_item("" if o1 is None else o1))
        self.tbl.setItem(r, 3, _center_item("" if o2 is None else o2))
        self.tbl.setItem(r, 4, QTableWidgetItem(it.remark))

    def add_row(self):
        self._append_row(BusinessItem(name="新业务", offsets=[], remark=""))
        self.tbl.scrollToBottom()
        self.tbl.editItem(self.tbl.item(self.tbl.rowCount() - 1, 1))

    def del_row(self):
        rows = sorted({i.row() for i in self.tbl.selectedIndexes()}, reverse=True)
        if not rows:
            QMessageBox.information(self, "提示", "请先在表格中选中要删除的行")
            return
        for r in rows:
            self.tbl.removeRow(r)

    def get_category(self) -> Category:
        cat = Category(name=self.cat.name,
                       base_network=self.ed_base.text().strip() or self.cat.base_network)
        for r in range(self.tbl.rowCount()):
            name_it = self.tbl.item(r, 1)
            name = name_it.text().strip() if name_it else ""
            if not name:
                continue
            enabled = False
            c_it = self.tbl.item(r, 0)
            if c_it is not None:
                enabled = (c_it.checkState() == Qt.Checked)

            offs = []
            for col in (2, 3):
                it = self.tbl.item(r, col)
                raw = it.text().strip() if it else ""
                if raw:
                    try:
                        offs.append(int(raw))
                    except ValueError:
                        pass
            rem_it = self.tbl.item(r, 4)
            remark = rem_it.text().strip() if rem_it else ""
            cat.items.append(BusinessItem(name=name, offsets=offs,
                                          remark=remark, enabled=enabled))
        return cat


# ---------------------------------------------------------------- 主窗口

class MainWindow(QMainWindow):

    def __init__(self, db_path: str = None):
        super().__init__()
        self.categories: list = []
        self.tables: list = []
        self.row_edits: dict = {}     # table_index -> QLineEdit（起始 IP）
        self.mask_spins: dict = {}    # table_index -> QSpinBox
        self.last_word_dir: str = ""
        self.db_path: str = H.db_path()
        self.history_panel = None
        self.history = None
        self.hist_rows: list = []
        self.n_hist_conflict: int = 0
        self._db_path = db_path       # None 时用默认位置（程序目录/history）

        self.setWindowTitle(f"{APP_NAME}  v{APP_VER}")
        self.resize(1300, 820)
        self.setStyleSheet(QSS)
        self._build()

        # 载入内置模板 + 默认业务配置
        self.cmb_tpl.setCurrentIndex(0)
        self._load_template(0)
        init = self._builtin_template()
        if init:
            self.ed_tpl.setText(init)
        try:
            self.history = H.History(self._db_path)
        except Exception as e:
            self.history = None
            print(f"[警告] 历史台账数据库初始化失败：{e}")
        if self.history is not None:
            self._hist_reload()
        self.statusBar().showMessage("就绪")

    # ---------------------------------------------------------- 界面
    def _build(self):
        central = QWidget()
        self.setCentralWidget(central)
        root = QVBoxLayout(central)
        root.setContentsMargins(10, 10, 10, 8)
        root.setSpacing(8)

        root.addLayout(self._build_top_bar())
        root.addWidget(self._build_tabs(), 1)

        self.setStatusBar(QStatusBar())

    def _build_top_bar(self) -> QHBoxLayout:
        hl = QHBoxLayout()
        hl.setSpacing(8)

        hl.addWidget(QLabel("批复单："))
        self.ed_tpl = QLineEdit()
        self.ed_tpl.setPlaceholderText("选择一份 .docx 批复单（默认用内置模板）")
        self.ed_tpl.textChanged.connect(self._on_tpl_changed)
        hl.addWidget(self.ed_tpl, 1)
        b = QPushButton("浏览…")
        b.clicked.connect(self._browse_tpl)
        hl.addWidget(b)

        self.btn_detect = QPushButton("自动识别")
        self.btn_detect.setObjectName("accent")
        self.btn_detect.setToolTip(
            "穷举「业务模板 × 掩码位数」，用批复单里已有的地址打分，\n"
            "自动挑出最吻合的一套模板、掩码和每张表的起始 IP")
        self.btn_detect.clicked.connect(self.do_detect)
        hl.addWidget(self.btn_detect)

        self.btn_refresh = QPushButton("刷新预览")
        self.btn_refresh.clicked.connect(self.do_refresh)
        hl.addWidget(self.btn_refresh)

        hl.addSpacing(8)
        hl.addWidget(QLabel("业务模板："))
        self.cmb_tpl = QComboBox()
        self.cmb_tpl.addItems(list(C.TEMPLATES.keys()))
        self.cmb_tpl.setFixedWidth(170)
        self.cmb_tpl.currentIndexChanged.connect(self._load_template)
        hl.addWidget(self.cmb_tpl)

        hl.addSpacing(8)
        hl.addWidget(QLabel("站点名称："))
        self.ed_site = QLineEdit()
        self.ed_site.setPlaceholderText("如：色尼风电（导出时记入历史库）")
        self.ed_site.setFixedWidth(180)
        self.ed_site.setToolTip("填写后，本次导出会以该站点的名义存入历史记录，\n"
                                "便于日后按站点查询、以及给新站分配时检测地址冲突。")
        hl.addWidget(self.ed_site)

        self.btn_save_cfg = QPushButton("保存配置")
        self.btn_load_cfg = QPushButton("加载配置")
        self.btn_save_cfg.clicked.connect(self.do_save_config)
        self.btn_load_cfg.clicked.connect(self.do_load_config)
        hl.addWidget(self.btn_save_cfg)
        hl.addWidget(self.btn_load_cfg)

        return hl

    def _build_tabs(self) -> QTabWidget:
        self.tabs = QTabWidget()

        # ---------- Tab1 批复单填充 ----------
        w1 = QWidget()
        l1 = QHBoxLayout(w1)
        l1.setContentsMargins(6, 8, 6, 6)
        l1.setSpacing(8)

        left = QGroupBox("业务偏移表")
        ll = QVBoxLayout(left)
        ll.setContentsMargins(8, 12, 8, 8)
        self.cat_tabs = QTabWidget()
        ll.addWidget(self.cat_tabs)
        l1.addWidget(left, 4)

        right = QWidget()
        rl = QVBoxLayout(right)
        rl.setContentsMargins(0, 8, 0, 0)
        rl.setSpacing(8)

        self.gb_bases = QGroupBox("每张纵密表的起始 IP")
        self.bases_lay = QGridLayout(self.gb_bases)
        self.bases_lay.setContentsMargins(10, 14, 10, 10)
        self.bases_lay.setSpacing(6)
        self.bases_lay.setColumnStretch(1, 1)
        self.gb_bases.setMaximumHeight(210)
        rl.addWidget(self.gb_bases)

        gb = QGroupBox("填充预览")
        gl = QVBoxLayout(gb)
        gl.setContentsMargins(8, 12, 8, 8)
        self.tbl_prev = QTableWidget(0, len(PREV_HEADERS))
        self.tbl_prev.setHorizontalHeaderLabels(PREV_HEADERS)
        self.tbl_prev.verticalHeader().setVisible(False)
        self.tbl_prev.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.tbl_prev.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.tbl_prev.setAlternatingRowColors(True)
        self.tbl_prev.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        self.tbl_prev.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
        gl.addWidget(self.tbl_prev)

        hb = QHBoxLayout()
        self.chk_only_empty = QCheckBox("仅填空白处（保留表中已有地址）")
        self.chk_only_empty.stateChanged.connect(lambda _: self.do_refresh())
        hb.addWidget(self.chk_only_empty)
        hb.addStretch()
        self.lab_stat = QLabel("")
        hb.addWidget(self.lab_stat)
        self.btn_export = QPushButton("导出 Word 批复单")
        self.btn_export.setObjectName("primary")
        self.btn_export.setFixedWidth(170)
        self.btn_export.clicked.connect(self.do_export)
        hb.addWidget(self.btn_export)
        gl.addLayout(hb)

        rl.addWidget(gb, 1)
        l1.addWidget(right, 6)

        self.tabs.addTab(w1, "批复单填充")

        # ---------- Tab2 批量站点 ----------
        w2 = QWidget()
        l2 = QVBoxLayout(w2)
        l2.setContentsMargins(6, 8, 6, 6)
        l2.setSpacing(8)

        gb2 = QGroupBox("站点列表（每行一个站点，起始 IP 留空则沿用「批复单填充」页的设置）")
        g2 = QVBoxLayout(gb2)
        g2.setContentsMargins(8, 12, 8, 8)
        self.tbl_sites = QTableWidget(0, 1)
        self.tbl_sites.verticalHeader().setVisible(False)
        self.tbl_sites.setAlternatingRowColors(True)
        self.tbl_sites.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        g2.addWidget(self.tbl_sites, 1)

        hb2 = QHBoxLayout()
        b_add = QPushButton("＋ 新增站点")
        b_del = QPushButton("－ 删除选中")
        b_del.setObjectName("danger")
        b_add.clicked.connect(self._site_add)
        b_del.clicked.connect(self._site_del)
        self.btn_batch = QPushButton("批量导出批复单")
        self.btn_batch.setObjectName("primary")
        self.btn_batch.setFixedWidth(170)
        self.btn_batch.clicked.connect(self.do_batch_export)
        hb2.addWidget(b_add)
        hb2.addWidget(b_del)
        hb2.addStretch()
        hb2.addWidget(self.btn_batch)
        g2.addLayout(hb2)
        l2.addWidget(gb2, 1)

        tip = QLabel("批量导出会为每个站点生成一份填好的批复单，文件名 = 站点名.docx，"
                     "统一保存到你选择的文件夹。")
        tip.setStyleSheet("color:#5D6D7E; padding:2px 4px;")
        l2.addWidget(tip)

        self.tabs.addTab(w2, "批量站点")


        # ---------- Tab4 使用说明 ----------
        w3 = QWidget()
        l3 = QVBoxLayout(w3)
        l3.setContentsMargins(6, 8, 6, 6)
        txt = QLabel(self._help_text())
        txt.setWordWrap(True)
        txt.setTextInteractionFlags(Qt.TextSelectableByMouse)
        txt.setAlignment(Qt.AlignTop)
        txt.setStyleSheet("background:#FFFFFF; padding:14px; border:1px solid #D6DAE0; border-radius:6px;")
        l3.addWidget(txt)

        self.tabs.addTab(self._build_history_tab(), "历史台账")
        self.tabs.addTab(w3, "使用说明")
        self.tabs.currentChanged.connect(self._on_tab_changed)

        return self.tabs



    # ---------------------------------------------------------- 历史记录
    # ---------------------------------------------------------- 历史记录
    def _build_history_tab(self) -> QWidget:
        w = QWidget()
        lay = QVBoxLayout(w)
        lay.setContentsMargins(6, 8, 6, 6)
        lay.setSpacing(8)

        # --- 查询条件 ---
        gb = QGroupBox("查询条件")
        hl = QHBoxLayout(gb)
        hl.setContentsMargins(10, 14, 10, 10)
        hl.setSpacing(8)

        hl.addWidget(QLabel("关键词："))
        self.ed_kw = QLineEdit()
        self.ed_kw.setPlaceholderText("站点 / 业务 / IP 地址 / 策略名称，支持模糊匹配")
        self.ed_kw.setFixedWidth(260)
        self.ed_kw.returnPressed.connect(self._hist_search)
        hl.addWidget(self.ed_kw)

        hl.addWidget(QLabel("站点："))
        self.cmb_hist_site = QComboBox()
        self.cmb_hist_site.setMinimumWidth(120)
        hl.addWidget(self.cmb_hist_site)

        hl.addWidget(QLabel("业务："))
        self.cmb_hist_biz = QComboBox()
        self.cmb_hist_biz.setMinimumWidth(100)
        hl.addWidget(self.cmb_hist_biz)

        hl.addWidget(QLabel("时间："))
        self.cmb_hist_days = QComboBox()
        self.cmb_hist_days.addItems(["全部", "最近 7 天", "最近 30 天",
                                     "最近 90 天", "最近一年"])
        self.cmb_hist_days.setFixedWidth(110)
        hl.addWidget(self.cmb_hist_days)

        b = QPushButton("查询")
        b.setObjectName("primary")
        b.setFixedWidth(70)
        b.clicked.connect(self._hist_search)
        hl.addWidget(b)
        b = QPushButton("全部")
        b.setFixedWidth(60)
        b.clicked.connect(self._hist_all)
        hl.addWidget(b)
        hl.addStretch()
        lay.addWidget(gb)

        # --- 批次表 ---
        gb2 = QGroupBox("导出批次（单击查看该批次的地址明细）")
        vl = QVBoxLayout(gb2)
        vl.setContentsMargins(10, 14, 10, 10)
        self.tbl_runs = QTableWidget(0, len(RUN_HEADERS))
        self.tbl_runs.setHorizontalHeaderLabels(RUN_HEADERS)
        self.tbl_runs.verticalHeader().setVisible(False)
        self.tbl_runs.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.tbl_runs.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.tbl_runs.setAlternatingRowColors(True)
        self.tbl_runs.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        self.tbl_runs.horizontalHeader().setSectionResizeMode(7, QHeaderView.Stretch)
        self.tbl_runs.setMaximumHeight(190)
        self.tbl_runs.itemSelectionChanged.connect(self._hist_show_run)
        vl.addWidget(self.tbl_runs)

        hb = QHBoxLayout()
        b_del = QPushButton("删除选中批次")
        b_del.setObjectName("danger")
        b_del.clicked.connect(self._hist_del_run)
        b_clear = QPushButton("清空全部历史")
        b_clear.setObjectName("danger")
        b_clear.clicked.connect(self._hist_clear)
        hb.addWidget(b_del)
        hb.addWidget(b_clear)
        hb.addStretch()
        vl.addLayout(hb)
        lay.addWidget(gb2)

        # --- 明细表 ---
        gb3 = QGroupBox("地址明细")
        vl3 = QVBoxLayout(gb3)
        vl3.setContentsMargins(10, 14, 10, 10)
        self.tbl_hist = QTableWidget(0, len(HIST_HEADERS))
        self.tbl_hist.setHorizontalHeaderLabels(HIST_HEADERS)
        self.tbl_hist.verticalHeader().setVisible(False)
        self.tbl_hist.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.tbl_hist.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.tbl_hist.setAlternatingRowColors(True)
        self.tbl_hist.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        self.tbl_hist.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
        vl3.addWidget(self.tbl_hist)

        hb3 = QHBoxLayout()
        self.lab_hist = QLabel("尚未查询")
        hb3.addWidget(self.lab_hist)
        hb3.addStretch()
        b_dup = QPushButton("地址查重")
        b_dup.setObjectName("accent")
        b_dup.setToolTip("查出历史上同一个地址被分配到多处的情况")
        b_dup.clicked.connect(self._hist_dup)
        b_xls = QPushButton("导出 Excel")
        b_xls.clicked.connect(self._hist_export)
        hb3.addWidget(b_dup)
        hb3.addWidget(b_xls)
        vl3.addLayout(hb3)
        lay.addWidget(gb3, 1)

        return w

    def _on_tab_changed(self, idx: int):
        if self.tabs.tabText(idx) == "历史台账" and self.history is not None:
            self._hist_refresh_runs()
            self._hist_all()

    def _hist_reload(self):
        """导出后刷新历史台账页签"""
        if self.history is None:
            return
        self._hist_refresh_runs()
        self._hist_all()

    def _hist_refresh_runs(self):
        if self.history is None:
            return
        # 下拉框
        for cmb, vals, first in ((self.cmb_hist_site, self.history.sites(), "（全部站点）"),
                                 (self.cmb_hist_biz, self.history.businesses(), "（全部业务）")):
            cur = cmb.currentText()
            cmb.blockSignals(True)
            cmb.clear()
            cmb.addItem(first)
            cmb.addItems(vals)
            if cur in vals or cur == first:
                cmb.setCurrentText(cur)
            cmb.blockSignals(False)

        rows = self.history.list_runs()
        t = self.tbl_runs
        t.blockSignals(True)
        t.setRowCount(0)
        for r in rows:
            i = t.rowCount()
            t.insertRow(i)
            vals = [r["id"], r["exported_at"], r["site"], r["src_doc"] or "",
                    r["template"] or "", f"/{r['mask']}" if r["mask"] else "",
                    r["filled_count"], r["out_path"] or ""]
            for c, v in enumerate(vals):
                it = QTableWidgetItem(str(v))
                it.setTextAlignment(Qt.AlignCenter)
                if c == 2:
                    it.setForeground(QBrush(QColor("#2F6FD0")))
                    f = it.font()
                    f.setBold(True)
                    it.setFont(f)
                t.setItem(i, c, it)
        t.blockSignals(False)

    def _hist_fill(self, rows):
        t = self.tbl_hist
        t.setRowCount(0)
        for r in rows:
            i = t.rowCount()
            t.insertRow(i)
            for c, f in enumerate(HIST_FIELDS):
                v = r[f]
                if f == "prefix" and v:
                    v = f"/{v}"
                if f == "offset" and v is None:
                    v = ""
                txt = "" if v is None else str(v)
                it = QTableWidgetItem(txt)
                it.setTextAlignment(Qt.AlignCenter)
                if f == "ip":
                    it.setForeground(QBrush(QColor("#1E8449")))
                    fo = it.font()
                    fo.setBold(True)
                    it.setFont(fo)
                t.setItem(i, c, it)
        self.lab_hist.setText(f"共 {len(rows)} 条")

    def _hist_search(self):
        if self.history is None:
            return
        kw = self.ed_kw.text().strip()
        site = self.cmb_hist_site.currentText()
        site = "" if site.startswith("（") else site
        biz = self.cmb_hist_biz.currentText()
        biz = "" if biz.startswith("（") else biz
        days_map = {"全部": 0, "最近 7 天": 7, "最近 30 天": 30,
                    "最近 90 天": 90, "最近一年": 365}
        days = days_map.get(self.cmb_hist_days.currentText(), 0)
        rows = self.history.search(keyword=kw, site=site, business=biz, days=days)
        self._hist_fill(rows)
        self.tbl_runs.clearSelection()
        self.statusBar().showMessage(
            f"查询到 {len(rows)} 条记录"
            + (f"（关键词：{kw}）" if kw else ""))

    def _hist_all(self):
        self.ed_kw.clear()
        self.cmb_hist_site.setCurrentIndex(0)
        self.cmb_hist_biz.setCurrentIndex(0)
        self.cmb_hist_days.setCurrentIndex(0)
        if self.history is None:
            return
        self._hist_fill(self.history.search())

    def _hist_show_run(self):
        items = self.tbl_runs.selectedItems()
        if not items or self.history is None:
            return
        rid = int(self.tbl_runs.item(items[0].row(), 0).text())
        self._hist_fill(self.history.run_items(rid))
        self.statusBar().showMessage(f"显示批次 {rid} 的地址明细")

    def _hist_del_run(self):
        rows = {i.row() for i in self.tbl_runs.selectedIndexes()}
        if not rows:
            QMessageBox.information(self, "提示", "请先在上方批次表中选中要删除的批次")
            return
        if self.history is None:
            return
        ids = [int(self.tbl_runs.item(r, 0).text()) for r in rows]
        if QMessageBox.question(
                self, "确认删除",
                f"将删除 {len(ids)} 个批次及其全部地址明细，此操作不可撤销。\n确定继续？"
        ) != QMessageBox.Yes:
            return
        self.history.delete_runs(ids)
        self._hist_refresh_runs()
        self._hist_all()
        self.statusBar().showMessage(f"已删除 {len(ids)} 个批次")

    def _hist_clear(self):
        if self.history is None:
            return
        if QMessageBox.question(
                self, "确认清空",
                "将清空全部历史记录，此操作不可撤销。\n确定继续？"
        ) != QMessageBox.Yes:
            return
        self.history.clear()
        self._hist_refresh_runs()
        self._hist_all()
        self.statusBar().showMessage("历史记录已清空")

    def _hist_export(self):
        if self.history is None or self.tbl_hist.rowCount() == 0:
            QMessageBox.information(self, "提示", "没有可导出的记录")
            return
        path, _ = QFileDialog.getSaveFileName(
            self, "导出历史记录", "地址分配历史记录.xlsx", "Excel 工作簿 (*.xlsx)")
        if not path:
            return
        rows = []
        for r in range(self.tbl_hist.rowCount()):
            d = {}
            for c, h in enumerate(HIST_HEADERS):
                it = self.tbl_hist.item(r, c)
                d[h] = it.text() if it else ""
            rows.append(d)
        try:
            X.export_dict_rows_excel(path, rows, HIST_HEADERS, sheet_title="地址分配历史")
            QMessageBox.information(self, "导出成功", f"已导出 {len(rows)} 条：\n{path}")
        except Exception as e:
            QMessageBox.critical(self, "导出失败", str(e))

    def _hist_dup(self):
        if self.history is None:
            return
        dups = self.history.duplicates()
        if not dups:
            QMessageBox.information(self, "地址查重",
                                    "未发现重复地址 —— 历史上每个地址都只分配过一次。")
            return
        errs = [d for d in dups if d["level"] == "error"]
        warns = [d for d in dups if d["level"] == "warning"]
        lines = []
        if errs:
            lines.append(f"【站点内冲突】{len(errs)} 处，同一地址在同一站点分给了多个业务，"
                         f"需要立即核对：")
            for d in errs[:10]:
                lines.append(f"  · {d['ip']}  业务：{'、'.join(d['businesses'])}"
                             f"  站点：{'、'.join(d['sites'])}")
            if len(errs) > 10:
                lines.append(f"  … 另有 {len(errs) - 10} 处")
        if warns:
            lines.append("")
            lines.append(f"【站点间重复】{len(warns)} 处，同一地址被多个站点使用。"
                         f"若各站点分属不同 VPN 则属正常，否则是规划撞车：")
            for d in warns[:10]:
                lines.append(f"  · {d['ip']}  站点：{'、'.join(d['sites'])}"
                             f"  业务：{'、'.join(d['businesses'])}")
            if len(warns) > 10:
                lines.append(f"  … 另有 {len(warns) - 10} 处")
        QMessageBox.information(self, "地址查重结果", "\n".join(lines))

    @staticmethod
    def _help_text() -> str:
        return """<h3>这个工具做什么</h3>
<p>打开一份《调度数据网接入地址批复单》，把「<b>本端业务IP地址</b>」列按「<b>策略名称</b>」自动算好、填进去，另存成新文件。<b>原模板不会被改动。</b></p>

<h3>计算规则</h3>
<p><b>业务地址 = 起始 IP 前三段 . (基准 + 偏移)</b><br>
基准 = 起始 IP 的<b>第四字节</b>，偏移 = 左侧业务表里的 C 列 / D 列。</p>
<p><b>例：</b>起始 IP = <code>54.100.134.96</code>，远动偏移 7、8，掩码 /28<br>
&nbsp;&nbsp;&nbsp;&nbsp;→ 远动 1 网 = <b>54.100.134.103</b>，远动 2 网 = <b>54.100.134.104</b><br>
&nbsp;&nbsp;&nbsp;&nbsp;→ 批复单里写成 <b>54.100.134.103-104</b></p>
<p style="color:#C0392B"><b>注意：这是给单个业务分配一个 IP 地址，不是划分网段。</b>
掩码位数只用来标出该地址所属的网段信息（网络号、网关、广播）。</p>

<h3>三步上手</h3>
<ol>
<li><b>选批复单</b>：默认已载入内置的批复单模板，也可点「浏览」选自己的 .docx。</li>
<li><b>点「自动识别」</b>（推荐）：程序穷举「业务模板 × 掩码位数」，拿批复单里已有的地址逐个打分，
自动挑出最吻合的一套，并把<b>每张表的起始 IP</b> 一起填好。识别结果会显示在底部状态栏，
如「自动识别：28位掩码（备用模板）+ /28，32/32 行吻合」。</li>
<li><b>看预览 → 导出</b>：右侧预览列出每行「策略名称 → 匹配业务 → 原值 → 将填入」，
确认无误后点「导出 Word 批复单」。</li>
</ol>

<h3>预览里的颜色</h3>
<ul>
<li><b><span style="color:#1E8449">绿色</span></b>：原本空白、本次补填</li>
<li><b><span style="color:#B9770E">橙色</span></b>：原值会被覆盖（说明配置与批复单不一致，请核对）</li>
<li><b><span style="color:#7F8C8D">灰色</span></b>：算出来与原值一致，无变化</li>
<li><b><span style="color:#C0392B">红色</span></b>：未匹配到业务或未分配地址，跳过不动</li>
</ul>

<h3>策略名称是怎么匹配业务的</h3>
<p>关键词包含匹配，且<b>只在该表所属大类（实时 / 非实时）内匹配</b>，避免跨大类撞名
（例如实时的「网安」与非实时的「网络安全监测」是两个不同业务）。</p>
<table border="1" cellpadding="4" cellspacing="0" style="border-collapse:collapse;font-size:12px">
<tr><th>策略名称</th><th>匹配到的业务</th></tr>
<tr><td>备调/地调/区调远动fes1~4</td><td>远动</td></tr>
<tr><td>区调PMU1 / PMU2</td><td>PMU</td></tr>
<tr><td>区调稳控</td><td>稳控</td></tr>
<tr><td>区调保护1 / 保护2</td><td>保护</td></tr>
<tr><td>区调内网管控（网安）</td><td>网安</td></tr>
<tr><td>区调电量、地调电量</td><td>电量</td></tr>
<tr><td>区调保信</td><td>保信</td></tr>
<tr><td>区调光功率 / 风功率</td><td>光功率 / 风功率</td></tr>
<tr><td>区调水情</td><td>水情</td></tr>
<tr><td>区调同步时钟、区调安控1/2</td><td>同步时钟、安控（默认停用，需自行分配偏移）</td></tr>
</table>
<p>要加新规则，编辑 <code>ip_word.py</code> 顶部的 <code>POLICY_RULES</code>。</p>

<h3>地址写法</h3>
<ul>
<li>单网 → <code>54.100.134.98</code></li>
<li>双网 → <code>54.100.134.103-104</code>（前三段相同，末位写范围）</li>
<li>不连续 → <code>54.100.135.98、54.100.135.118</code></li>
</ul>

<h3>其他功能</h3>
<ul>
<li><b>仅填空白处</b>：只补批复单里空着的格子，已有地址原样保留。</li>
<li><b>批量站点</b>：第二个页签，每行一个站点，一次性为每个站点导出一份填好的批复单。</li>
<li><b>配置存取</b>：业务偏移表调好后「保存配置」成 JSON，下次「加载配置」直接还原。</li>
</ul>

<h3>已用真实批复单验证</h3>
<p>用《那曲色尼风电调度数据网接入地址批复单》反推，其地址与「28位掩码（备用模板）」的偏移量
<b>完全吻合</b>：网安 +2、稳控 +3/+4、PMU +5/+6、远动 +7/+8、保护 +9/+10；
非实时电量 +10、保信 +5、光功率 +6、水情 +7。两个接入网（一平面 / 二平面）均验证通过，
已填的 32 行与原文 <b>100% 一致</b>，另自动补填了 12 行原本空着的远动地址。</p>"""

    # ---------------------------------------------------------- 模板 / 配置
    @staticmethod
    def _builtin_template() -> str:
        """程序目录下的内置批复单模板（打包后位于 _MEIPASS/templates）"""
        cands = []
        if getattr(sys, "frozen", False):
            cands.append(os.path.join(sys._MEIPASS, "templates"))
        cands.append(os.path.join(os.path.dirname(os.path.abspath(__file__)), "templates"))
        for d in cands:
            if os.path.isdir(d):
                for f in sorted(os.listdir(d)):
                    if f.lower().endswith(".docx") and not f.startswith("~$"):
                        return os.path.join(d, f)
        return ""

    def _load_template(self, idx: int):
        name = self.cmb_tpl.itemText(idx)
        fn = C.TEMPLATES.get(name)
        if not fn:
            return
        data = fn()
        cats, prefix, policy = C.dict_to_config(data)
        self.categories = cats
        self.def_prefix = prefix
        self.def_policy = policy
        self.cat_tabs.clear()
        self.panels = []
        for cat in cats:
            p = CategoryPanel(cat)
            self.panels.append(p)
            self.cat_tabs.addTab(p, cat.name)
        self.statusBar().showMessage(f"已载入业务模板：{name}")
        if self.tables:
            self.do_refresh()

    def _collect_categories(self) -> list:
        return [p.get_category() for p in self.panels]

    def do_save_config(self):
        path, _ = QFileDialog.getSaveFileName(
            self, "保存配置", "批复单填充配置.json", "JSON 配置 (*.json)")
        if not path:
            return
        try:
            C.save_config(path, self._collect_categories(),
                          self.def_prefix, self.def_policy)
            QMessageBox.information(self, "成功", f"配置已保存：\n{path}")
        except Exception as e:
            QMessageBox.critical(self, "失败", str(e))

    def do_load_config(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "加载配置", "", "JSON 配置 (*.json)")
        if not path:
            return
        try:
            data = C.load_config(path)
            cats, prefix, policy = C.dict_to_config(data)
            self.def_prefix, self.def_policy = prefix, policy
            self.categories = cats
            self.cat_tabs.clear()
            self.panels = []
            for cat in cats:
                p = CategoryPanel(cat)
                self.panels.append(p)
                self.cat_tabs.addTab(p, cat.name)
            QMessageBox.information(self, "成功", f"配置已载入：\n{path}")
            if self.tables:
                self.do_refresh()
        except Exception as e:
            QMessageBox.critical(self, "失败", str(e))

    # ---------------------------------------------------------- 批复单
    def _browse_tpl(self):
        start = self.last_word_dir or os.path.dirname(self._builtin_template())
        path, _ = QFileDialog.getOpenFileName(
            self, "选择批复单", start, "Word 文档 (*.docx)")
        if path:
            self.last_word_dir = os.path.dirname(path)
            self.ed_tpl.setText(path)

    def _on_tpl_changed(self, path: str):
        self.tables = []
        self._clear_bases()
        if not path or not os.path.exists(path):
            self.lab_stat.setText("")
            self._fill_preview()
            return
        try:
            self.tables = W.scan_template(path)
        except Exception as e:
            QMessageBox.critical(self, "读取失败", f"无法解析该文档：\n{e}")
            return
        if not self.tables:
            self.statusBar().showMessage(
                "未在该文档中找到含「策略名称 / 本端业务IP地址」的表格")
            self._fill_preview()
            return

        self._build_bases()
        self.do_refresh()
        self.statusBar().showMessage(
            f"已载入批复单：{os.path.basename(path)}，识别到 {len(self.tables)} 张纵密表，"
            f"建议点「自动识别」")

    def _clear_bases(self):
        while self.bases_lay.count():
            item = self.bases_lay.takeAt(0)
            w = item.widget()
            if w:
                w.deleteLater()
        self.row_edits.clear()
        self.mask_spins.clear()

    def _build_bases(self):
        self._clear_bases()
        r = 0
        self.bases_lay.addWidget(QLabel("纵密表"), r, 0)
        self.bases_lay.addWidget(QLabel("起始 IP（第四字节为基准）"), r, 1)
        self.bases_lay.addWidget(QLabel("掩码"), r, 2)
        r += 1
        for t in self.tables:
            self.bases_lay.addWidget(QLabel(t.title), r, 0)

            ed = QLineEdit(t.suggested_base)
            ed.setToolTip("填完整起始 IP，第四字节即基准值\n"
                          "例：54.100.134.96 → 远动偏移 7 → 54.100.134.103")
            ed.textChanged.connect(lambda _: self.do_refresh())
            self.row_edits[t.table_index] = ed
            self.bases_lay.addWidget(ed, r, 1)

            sp = QSpinBox()
            sp.setRange(24, 30)
            sp.setValue(self.def_prefix)
            sp.setSuffix(" 位")
            sp.setFixedWidth(80)
            sp.valueChanged.connect(lambda _: self.do_refresh())
            self.mask_spins[t.table_index] = sp
            self.bases_lay.addWidget(sp, r, 2)
            r += 1

        # 同步批量站点表头
        self._sync_sites_header()

    def _sync_sites_header(self):
        if not self.tables:
            return
        headers = ["站点名称"] + [t.title for t in self.tables]
        self.tbl_sites.setColumnCount(len(headers))
        self.tbl_sites.setHorizontalHeaderLabels(headers)
        for r in range(self.tbl_sites.rowCount()):
            for c in range(1, len(headers)):
                if not self.tbl_sites.item(r, c):
                    self.tbl_sites.setItem(r, c, _center_item(""))

    # ---------------------------------------------------------- 自动识别
    def do_detect(self):
        if not self.tables:
            QMessageBox.information(self, "提示", "请先选择一份批复单")
            return
        tpls = {}
        for name, fn in C.TEMPLATES.items():
            cats, _p, _g = C.dict_to_config(fn())
            tpls[name] = cats

        best = W.auto_detect(self.tables, tpls)
        if best is None:
            QMessageBox.information(
                self, "无法自动识别",
                "批复单里没有已填写的地址可供比对。\n"
                "请手动在上方为每张表填写起始 IP 和掩码。")
            return

        # 切到识别出的业务模板
        idx = self.cmb_tpl.findText(best["template"])
        if idx >= 0 and idx != self.cmb_tpl.currentIndex():
            self.cmb_tpl.setCurrentIndex(idx)
        elif idx >= 0:
            self._load_template(idx)

        # 每张表套用起始 IP + 掩码
        for t in self.tables:
            base = best["bases"].get(t.table_index)
            if base:
                self.row_edits[t.table_index].setText(base)
            self.mask_spins[t.table_index].setValue(best["mask"])
            t.suggested_mask = best["mask"]

        self.do_refresh()

        if best["hit"] == best["total"]:
            msg = (f"自动识别完成：{best['template']} + /{best['mask']} 掩码，"
                   f"{best['hit']}/{best['total']} 行与批复单原有地址完全吻合。")
            self.statusBar().showMessage("自动识别：" + msg)
            QMessageBox.information(self, "自动识别成功", msg)
        else:
            msg = (f"最佳匹配：{best['template']} + /{best['mask']} 掩码，"
                   f"仅 {best['hit']}/{best['total']} 行吻合。\n"
                   f"该批复单的偏移规则可能与内置模板不同，请核对预览里的橙色行。")
            self.statusBar().showMessage("自动识别：" + msg)
            QMessageBox.warning(self, "吻合度偏低", msg)

    # ---------------------------------------------------------- 预览
    def do_refresh(self):
        if not self.tables:
            return
        cats = self._collect_categories()
        only_empty = self.chk_only_empty.isChecked()

        W.apply_mapping_by_category(self.tables, cats)

        for t in self.tables:
            ed = self.row_edits.get(t.table_index)
            sp = self.mask_spins.get(t.table_index)
            if ed is None or sp is None:
                continue
            base_txt = ed.text().strip()
            mask = sp.value()
            t.suggested_base = base_txt
            t.suggested_mask = mask
            if not base_txt or C.parse_base(base_txt) is None:
                for row in t.rows:
                    row.new_value = ""
                continue

            cat = W.category_for_title(t.title, cats)
            c2 = C.Category(name=cat.name, base_network=base_txt,
                            items=[C.BusinessItem(i.name, list(i.offsets),
                                                  i.remark, i.enabled) for i in cat.items])
            res, _ = C.allocate_category(c2, mask, self.def_policy)
            pool: dict = {}
            offmap: dict = {}
            for h in res:
                pool.setdefault(h.business, []).append(h.ip)
                offmap.setdefault(h.business, []).append(h.offset)

            for row in t.rows:
                if only_empty and row.old_value:
                    row.new_value = ""
                    row.ips = []
                    continue
                ips = pool.get(row.business) if row.business else None
                row.new_value = W.format_ip_list(ips) if ips else ""
                row.ips = list(ips) if ips else []
                row.offsets = list(offmap.get(row.business, [])) if row.business else []

        self._check_history()
        self._fill_preview()

    def _check_history(self):
        """
        预览刷新后，实时把本次地址与历史台账比对，逐行给出「历史占用」提示。

        提示分三种：
          ⚠ 已被某站占用 —— 该地址在台账里属于别的站点，属于撞地址
            本站已记录    —— 本站此前导出过，属正常重复导出
            新占用        —— 台账里没有，本次首次分配
        """
        self.n_hist_conflict = 0
        self.hist_rows = []

        for tb in self.tables:
            for row in tb.rows:
                row.hist_note = ""

        if self.history is None or not self.tables:
            return

        site = (self.ed_site.text() or "").strip() if hasattr(self, "ed_site") else ""

        # 一次性批量查出本次所有 IP 的历史占用情况
        ips: list = []
        for tb in self.tables:
            for row in tb.rows:
                if row.new_value:
                    ips.extend(H.expand_ip_text(row.new_value))
        try:
            occ = self.history.occupied(ips)
        except Exception:
            return

        for tb in self.tables:
            for row in tb.rows:
                if not row.new_value:
                    row.hist_note = ""
                    continue
                mine, others = [], []
                for ip in H.expand_ip_text(row.new_value):
                    for old_site, old_biz, _when in occ.get(ip, []):
                        if site and old_site == site:
                            if old_site not in mine:
                                mine.append(old_site)
                        elif old_site not in others:
                            others.append(old_site)
                if others:
                    row.hist_note = "\u26a0 已被 " + "、".join(others) + " 占用"
                    self.n_hist_conflict += 1
                    self.hist_rows.append((tb.title, row.policy, row.new_value, others))
                elif mine:
                    row.hist_note = "本站已记录"
                else:
                    row.hist_note = "新占用"

    def _fill_preview(self):
        t = self.tbl_prev
        t.setRowCount(0)
        n_fill = n_new = n_chg = n_skip = 0
        for tb in self.tables:
            for row in tb.rows:
                r = t.rowCount()
                t.insertRow(r)
                if row.business:
                    biz_txt = row.business
                elif row.business_disabled:
                    biz_txt = f"{row.business_disabled}（未启用/无偏移）"
                else:
                    biz_txt = "—（策略名未匹配）"
                vals = [tb.title, row.tunnel, row.policy, biz_txt,
                        row.old_value or "（空）", row.new_value or "—",
                        row.hist_note]
                if row.new_value:
                    n_fill += 1
                    if not row.old_value:
                        color = CLR_NEW
                        n_new += 1
                    elif row.old_value != row.new_value:
                        color = CLR_CHG
                        n_chg += 1
                    else:
                        color = CLR_SAME
                else:
                    color = CLR_SKIP
                    n_skip += 1
                for c, v in enumerate(vals):
                    it = QTableWidgetItem(str(v))
                    it.setTextAlignment(Qt.AlignCenter)
                    it.setForeground(QBrush(color))
                    if c == 5 and row.new_value:
                        f = it.font()
                        f.setBold(True)
                        it.setFont(f)
                    if c == 6:
                        txt = str(v)
                        if txt.startswith("\u26a0"):
                            it.setForeground(QBrush(QColor("#C0392B")))
                            f = it.font(); f.setBold(True); it.setFont(f)
                        elif txt.startswith("\u65b0\u5360\u7528"):
                            it.setForeground(QBrush(QColor("#1E8449")))
                        else:
                            it.setForeground(QBrush(QColor("#95A5A6")))
                    t.setItem(r, c, it)

        msg = (f"将写入 {n_fill} 行（补填 {n_new}、覆盖变化 {n_chg}），跳过 {n_skip} 行")
        if self.n_hist_conflict:
            msg += f"    ⚠ {self.n_hist_conflict} 行与历史记录冲突"
            self.lab_stat.setStyleSheet("color:#C0392B; font-weight:bold;")
        else:
            self.lab_stat.setStyleSheet("")
        self.lab_stat.setText(msg)

    def do_export(self):
        src = self.ed_tpl.text().strip()
        if not src or not os.path.exists(src):
            QMessageBox.information(self, "提示", "请先选择一份批复单")
            return
        if not self.tables:
            QMessageBox.information(self, "提示", "未检测到可填充的表格")
            return
        self.do_refresh()

        # 站点名称（用于台账归档与查重排除自身）
        site, ok = self._ask_site()
        if not ok:
            return

        if not self._check_conflicts(site):
            self.statusBar().showMessage("已取消导出：发现地址冲突")
            return

        base, ext = os.path.splitext(os.path.basename(src))
        default = f"{base}_已填业务IP{ext or '.docx'}"
        path, _ = QFileDialog.getSaveFileName(
            self, "导出批复单", default, "Word 文档 (*.docx)")
        if not path:
            return
        try:
            W.fill_docx(src, path, self.tables)
        except Exception as e:
            QMessageBox.critical(self, "导出失败", str(e))
            return

        n = len([r for tb in self.tables for r in tb.rows if r.new_value])
        aid = self._archive(site, path)
        tip = f"\n已归档到历史台账（ID={aid}）" if aid else ""
        QMessageBox.information(
            self, "导出成功",
            f"已填充 {n} 处「本端业务IP地址」：\n{path}\n\n"
            f"原模板未被修改。{tip}")
        self.statusBar().showMessage(
            f"已导出：{path}" + (f"（归档 ID={aid}）" if aid else ""))

    # ---------------------------------------------------------- 批量
    def _site_add(self):
        r = self.tbl_sites.rowCount()
        self.tbl_sites.insertRow(r)
        self.tbl_sites.setItem(r, 0, _center_item(f"站点{r + 1}"))
        for c in range(1, self.tbl_sites.columnCount()):
            self.tbl_sites.setItem(r, c, _center_item(""))

    def _site_del(self):
        rows = sorted({i.row() for i in self.tbl_sites.selectedIndexes()}, reverse=True)
        if not rows:
            QMessageBox.information(self, "提示", "请先选中要删除的站点行")
            return
        for r in rows:
            self.tbl_sites.removeRow(r)

    def do_batch_export(self):
        src = self.ed_tpl.text().strip()
        if not src or not os.path.exists(src):
            QMessageBox.information(self, "提示", "请先选择一份批复单")
            return
        if not self.tables:
            QMessageBox.information(self, "提示", "未检测到可填充的表格")
            return
        if self.tbl_sites.rowCount() == 0:
            QMessageBox.information(self, "提示", "请先添加至少一个站点")
            return

        outdir = QFileDialog.getExistingDirectory(self, "选择保存文件夹")
        if not outdir:
            return

        cats = self._collect_categories()
        ok, fail = 0, []
        warns: list = []
        archived = 0
        for r in range(self.tbl_sites.rowCount()):
            g = lambda c: (self.tbl_sites.item(r, c).text().strip()
                           if self.tbl_sites.item(r, c) else "")
            site = g(0)
            if not site:
                continue
            try:
                # 按该站点的起始 IP 重算一遍再导出
                for i, t in enumerate(self.tables):
                    base = g(i + 1) or self.row_edits[t.table_index].text().strip()
                    mask = self.mask_spins[t.table_index].value()
                    cat = W.category_for_title(t.title, cats)
                    c2 = C.Category(name=cat.name, base_network=base,
                                    items=[C.BusinessItem(i2.name, list(i2.offsets),
                                                          i2.remark, i2.enabled)
                                           for i2 in cat.items])
                    res, _ = C.allocate_category(c2, mask, self.def_policy)
                    pool: dict = {}
                    for h in res:
                        pool.setdefault(h.business, []).append(h.ip)
                    for row in t.rows:
                        if self.chk_only_empty.isChecked() and row.old_value:
                            row.new_value = ""
                            continue
                        ips = pool.get(row.business) if row.business else None
                        row.new_value = W.format_ip_list(ips) if ips else ""
                # 查重（批量模式只警告不阻断，避免卡住整批）
                cf = (self.history.find_conflicts(
                    self.tables, exclude_site=site, new_site=site)
                    if self.history is not None else [])
                if cf:
                    warns.append(f"{site}：{len(cf)} 处地址与历史台账冲突"
                                 f"（如 {cf[0].ip} 已被 {cf[0].old_site} 占用）")
                out = os.path.join(outdir, f"{site}.docx")
                W.fill_docx(src, out, self.tables)
                aid = self._archive(site, out)
                if aid:
                    archived += 1
                ok += 1
            except Exception as e:
                fail.append(f"{site}：{e}")

        self._hist_reload()
        msg = f"已导出 {ok} 份批复单到：\n{outdir}\n已归档 {archived} 个站点到历史台账。"
        if warns:
            msg += "\n\n⚠ 冲突提醒：\n" + "\n".join(warns[:10])
            if len(warns) > 10:
                msg += f"\n…… 另有 {len(warns) - 10} 个站点"
        if fail:
            msg += "\n\n失败：" + "\n".join(fail)
        QMessageBox.information(self, "批量导出完成", msg)
        self.statusBar().showMessage(f"批量导出 {ok} 份、归档 {archived} 个到 {outdir}")


    def _ask_site(self) -> Tuple[str, bool]:
        """取站点名：优先用顶部「站点名称」输入框，为空时才弹窗询问"""
        if hasattr(self, "ed_site"):
            name = (self.ed_site.text() or "").strip()
            if name:
                return name, True
        from qt_compat import QInputDialog
        sites = self.history.sites() if self.history is not None else []
        name, ok = QInputDialog.getItem(
            self, "站点名称",
            "请输入站点名称（用于历史台账归档与查重）：",
            sites, 0, True)
        if not ok:
            return "", False
        name = (name or "").strip()
        if not name:
            QMessageBox.information(self, "提示", "站点名称不能为空")
            return "", False
        if hasattr(self, "ed_site"):
            self.ed_site.setText(name)
        return name, True


    # ---------------------------------------------------------- 台账 / 查重
    def _check_conflicts(self, site: str) -> bool:
        """
        导出前查重。返回 True 表示可以继续导出，False 表示用户取消。
        """
        if not self.tables:
            return True

        # 本次内部重复（多张表之间撞地址）
        inner = H.self_conflicts(self.tables)
        # 与历史台账重复
        outer = (self.history.find_conflicts(self.tables, exclude_site=site,
                                            new_site=site)
                  if self.history is not None else [])

        if not inner and not outer:
            return True

        lines = []
        if inner:
            lines.append(f"【本次导出内部重复 {len(inner)} 处】")
            for ip, tt, pol, first in inner[:10]:
                lines.append(f"  {ip}  {tt}/{pol}（首次出现在 {first}）")
            if len(inner) > 10:
                lines.append(f"  …… 另有 {len(inner) - 10} 处")
            lines.append("")
        if outer:
            lines.append(f"【与历史台账冲突 {len(outer)} 处】")
            for c in outer[:15]:
                lines.append(f"  {c.ip}  本次【{c.new_site}/{c.new_policy}】"
                             f" ← 已被【{c.old_site}/{c.old_policy}】占用"
                             f"（{c.created_at[:10]}）")
            if len(outer) > 15:
                lines.append(f"  …… 另有 {len(outer) - 15} 处")
            lines.append("")
        lines.append("继续导出仍会写入这些地址，并一并归档到台账。")

        box = QMessageBox(self)
        box.setIcon(QMessageBox.Warning)
        box.setWindowTitle("地址冲突提醒")
        box.setText(f"发现 {len(inner) + len(outer)} 处地址冲突，建议核对后再导出。")
        box.setDetailedText("\n".join(lines))
        box.setStandardButtons(QMessageBox.Yes | QMessageBox.Cancel)
        box.setDefaultButton(QMessageBox.Cancel)
        box.button(QMessageBox.Yes).setText("仍然导出")
        box.button(QMessageBox.Cancel).setText("返回修改")
        return box.exec() == QMessageBox.Yes

    def _archive(self, site: str, out_file: str) -> int:
        """把本次导出结果存入历史台账，返回记录 ID"""
        if self.history is None:
            return 0
        try:
            items = H.collect_items(self.tables)
            if not items:
                return 0
            mask = (self.mask_spins[self.tables[0].table_index].value()
                    if self.tables and self.tables[0].table_index in self.mask_spins
                    else self.def_prefix)
            rid = self.history.record_run(
                site=site,
                src_doc=os.path.basename(self.ed_tpl.text().strip()),
                out_path=out_file,
                template=self.cmb_tpl.currentText(),
                mask=mask, batch=False, items=items)
            self._hist_reload()
            return rid
        except Exception as e:
            QMessageBox.warning(self, "归档失败", f"导出已完成，但写入台账失败：\n{e}")
            return 0


# ---------------------------------------------------------------- 启动入口

def _app_dir() -> str:
    """程序所在目录：打包后为 exe 同级目录，否则为源码目录"""
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def write_error(title: str, detail: str) -> str:
    """把错误信息追加写入 error.log，返回日志路径"""
    path = os.path.join(_app_dir(), "error.log")
    try:
        with open(path, "a", encoding="utf-8") as f:
            f.write("\n" + "=" * 62 + "\n")
            f.write(f"时间  ：{datetime.now():%Y-%m-%d %H:%M:%S}\n")
            f.write(f"版本  ：v{APP_VER}\n")
            f.write(f"环境  ：{'打包EXE' if getattr(sys, 'frozen', False) else '源码运行'}\n")
            f.write(f"Python：{sys.version.split()[0]}\n")
            f.write(f"{title}\n\n{detail}\n")
    except Exception:
        pass
    return path


def _excepthook(etype, value, tb):
    """捕获所有未处理异常：写日志 + 尽力弹窗，避免「双击没反应」"""
    detail = "".join(traceback.format_exception(etype, value, tb))
    log = write_error("未处理异常", detail)
    try:
        if QApplication.instance() is not None:
            QMessageBox.critical(
                None, "程序出错",
                f"程序遇到错误：\n\n{etype.__name__}: {value}\n\n"
                f"详细信息已写入：\n{log}")
    except Exception:
        pass
    sys.__excepthook__(etype, value, tb)


def main() -> int:
    # High-DPI: Qt 6 only, silently skipped on Qt 5
    QC.setup_high_dpi()

    app = QApplication(sys.argv)
    app.setApplicationName(APP_NAME)
    app.setApplicationVersion(APP_VER)
    f = QFont()
    f.setPointSize(10)
    app.setFont(f)

    win = MainWindow()
    win.show()
    return QC.exec_app(app)


if __name__ == "__main__":
    sys.excepthook = _excepthook
    try:
        sys.exit(main())
    except SystemExit:
        raise
    except Exception:
        detail = traceback.format_exc()
        log = write_error("启动失败", detail)
        # 尽力弹窗；若连 Qt 都起不来，就只有日志文件可查
        try:
            app = QApplication.instance() or QApplication(sys.argv)
            QMessageBox.critical(
                None, "启动失败",
                f"程序启动失败：\n\n{traceback.format_exc(limit=3)}\n\n"
                f"详细信息已写入：\n{log}")
        except Exception:
            pass
        sys.exit(1)
