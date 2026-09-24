#!/usr/bin/env python3
"""
GURI Database Viewer and Manager - GUI Application
Comprehensive interface for viewing and managing GURI records with document type support.
"""

import sys
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QGridLayout,
    QLabel, QLineEdit, QPushButton, QTextEdit, QTreeWidget, QTreeWidgetItem,
    QTabWidget, QMessageBox, QFileDialog, QDialog, QComboBox, QListWidget,
    QListWidgetItem, QMenu, QMenuBar, QStatusBar, QFrame, QGroupBox,
    QScrollArea, QSplitter, QRadioButton, QButtonGroup, QScrollBar, QSystemTrayIcon,
    QCheckBox, QToolButton, QDateEdit, QFormLayout, QSpinBox, QHeaderView,
    QInputDialog,
)
from PySide6.QtCore import Qt, QTimer, Signal, QSize, QPoint, QObject, QDate, QRectF
from PySide6.QtGui import (
    QFont,
    QColor,
    QPainter,
    QPen,
    QPixmap,
    QImage,
    QClipboard,
    QIcon,
    QAction,
    QCursor,
    QBrush,
    QConicalGradient,
    QRadialGradient,
)
from datetime import datetime, timedelta
import json
import math
import os
import re
import threading
from pathlib import Path
from typing import Any, Dict, List, Optional, Set
import logging

from guri import GURIDatabase, connect_guri_database
from guri_decompositor import GURIDecompositor
from guri_component_library import GURIComponentLibrary
from guri_action_detector import (
    CATEGORY_ACTIONABLE,
    CATEGORY_DEPRECATED,
    CATEGORY_IMPORTANT,
    CATEGORY_LESS,
    CATEGORY_RELEVANT,
    TIMELINE_CATEGORIES,
    TIMELINE_CATEGORY_LABELS,
    categorize_detected,
    detect_batch,
    extract_deadlines,
    highlight_action_html,
    merge_deadline_hits,
    sender_matches_important_domain,
)
from guri_action_store import (
    apply_relevance_feedback,
    apply_statuses,
    compute_relevance_scores,
    deprecate_email,
    deprecate_sender,
    derate_content_types,
    derate_sender,
    endorse_email_score,
    endorse_sender_score,
    ignore_email,
    ignore_recipient,
    ignore_sender,
    is_mail_junked,
    load_action_status,
    load_deadline_account_keys,
    load_manual_deadlines,
    load_relevance_feedback,
    load_scrape_cache,
    load_scrape_settings,
    mail_feedback_key,
    mark_spam,
    normalize_sender_address,
    save_action_status,
    save_deadline_account_keys,
    save_manual_deadlines,
    add_manual_deadline,
    delete_manual_deadline,
    save_relevance_feedback,
    save_scrape_cache,
    save_scrape_settings,
    lookback_months_to_days,
)
from guri_learning import (
    CATEGORY_CORRECTION_REASONS,
    CATEGORY_CORRECTION_REASON_GROUPS,
    CATEGORY_RECAT_TARGETS,
    DEADLINE_TIMELINE_CATEGORIES,
    DEADLINE_TIMELINE_LABELS,
    DL_CAT_EXTRACTED,
    DL_CAT_MANUAL,
    POSITIVE_REASON_LEARNING,
    add_user_rule,
    apply_user_rules_to_item,
    apply_user_rules_to_items,
    candidate_deadline_emails,
    confirm_deadline_mail,
    cue_patterns_from_learning,
    domain_from_sender,
    ensure_identity_defaults,
    extract_greeting_name,
    format_user_rules_for_llm,
    is_deadline_rejected,
    learn_deadline_cue,
    link_manual_deadline,
    load_learning,
    mail_is_for_me,
    mark_asked_important_sender,
    normalize_rule_effects,
    normalize_rule_match,
    record_category_correction,
    remember_important_sender,
    forget_important_sender,
    remember_other_name,
    reject_deadline_mail,
    resolve_timeline_category,
    rule_has_enforceable_effects,
    save_learning,
    set_my_identity,
    should_ask_important_sender,
    should_include_deadline_mail,
    step_category_toward,
    summarize_rule_effects,
    update_user_rule,
)
from guri_outlook_scraper import (
    AesAccountScanConfig,
    discover_outlook_accounts,
    fetch_mail_body,
    load_aes_account_configs,
    open_mail_in_outlook,
    save_aes_account_configs,
    scrape_outlook_inboxes,
)
from guri_inbox_search import (
    load_inbox_search_settings,
    save_inbox_search_settings,
    search_inbox_conversations,
)
import aes_score_history
from broker_removal.gui_tab import DataBrokersPanel
from guri_file_scraper import (
    ScanLocation,
    load_file_scrape_cache,
    load_scan_locations,
    save_file_scrape_cache,
    save_scan_locations,
    scrape_file_locations,
    scraped_files_to_detection_items,
)
from guri_ollama import (
    build_email_analysis_prompt,
    build_rule_compile_prompt,
    chat as ollama_chat,
    format_gpu_report as ollama_format_gpu_report,
    list_models as ollama_list_models,
    load_config as ollama_load_config,
    parse_compiled_rule_json,
    ping as ollama_ping,
    probe_gpu as ollama_probe_gpu,
    save_config as ollama_save_config,
)
from version import (
    APP_FULL_NAME,
    APP_NAME,
    APP_ORG,
    APP_SILLY_QUOTE,
    COPYRIGHT_YEAR,
    RELEASE_NOTES,
    SUITE_NAME,
    VERSION,
    about_summary,
    suite_banner,
    version_string,
)


# ---------------------------------------------------------------------------
# Design system — one palette, one stylesheet, applied app-wide.
# ---------------------------------------------------------------------------
PALETTE = {
    "bg": "#f4f6f8",          # window background
    "surface": "#ffffff",     # cards, inputs, panes
    "border": "#e2e8ee",      # hairline borders
    "border_strong": "#d5dde3",
    "text": "#1a2e35",        # primary text
    "muted": "#5a7280",       # secondary text
    "accent": "#0f6b7c",      # brand teal
    "accent_soft": "#e3f1f4", # selection / hover wash
    "ok": "#1b7a3d",
    "err": "#b42318",
}


def add_dialog_button_row(
    parent_layout,
    *,
    affirmative=None,
    cancel=None,
    leading=None,
) -> QHBoxLayout:
    """Dialog footer: [leading…] … stretch … affirmative(s) … Cancel/Close.

    Affirmative actions sit to the left of Cancel; Cancel is always bottom-right.
    """
    row = QHBoxLayout()
    row.setContentsMargins(0, 8, 0, 0)
    row.setSpacing(8)
    for btn in leading or []:
        if btn is not None:
            row.addWidget(btn)
    row.addStretch(1)
    affirm = affirmative
    if affirm is None:
        affirm = []
    elif not isinstance(affirm, (list, tuple)):
        affirm = [affirm]
    for btn in affirm:
        if btn is not None:
            row.addWidget(btn)
    if cancel is not None:
        row.addWidget(cancel)
    parent_layout.addLayout(row)
    return row


GURI_STYLESHEET = f"""
QMainWindow, QDialog {{
    background: {PALETTE['bg']};
}}
QWidget {{
    color: {PALETTE['text']};
    font-family: 'Segoe UI';
    font-size: 9pt;
}}

/* ---- Tabs: quiet underline style ---- */
QTabWidget::pane {{
    border: 1px solid {PALETTE['border']};
    border-radius: 8px;
    background: {PALETTE['surface']};
    top: -1px;
}}
QTabBar::tab {{
    background: transparent;
    border: none;
    border-bottom: 2px solid transparent;
    padding: 8px 18px;
    margin-right: 2px;
    color: {PALETTE['muted']};
    font-weight: 600;
}}
QTabBar::tab:selected {{
    color: {PALETTE['accent']};
    border-bottom: 2px solid {PALETTE['accent']};
}}
QTabBar::tab:hover:!selected {{
    color: {PALETTE['text']};
}}

/* ---- Group boxes as cards ---- */
QGroupBox {{
    background: {PALETTE['surface']};
    border: 1px solid {PALETTE['border']};
    border-radius: 10px;
    margin-top: 16px;
    padding: 12px 10px 10px 10px;
    font-weight: 600;
}}
QGroupBox::title {{
    subcontrol-origin: margin;
    left: 12px;
    padding: 0 6px;
    color: {PALETTE['accent']};
}}

/* ---- Buttons ---- */
QPushButton {{
    background: {PALETTE['surface']};
    border: 1px solid {PALETTE['border_strong']};
    border-radius: 6px;
    padding: 6px 14px;
    color: {PALETTE['text']};
}}
QPushButton:hover {{
    border-color: {PALETTE['accent']};
    color: {PALETTE['accent']};
    background: #f5fafb;
}}
QPushButton:pressed {{
    background: {PALETTE['accent_soft']};
}}
QPushButton:disabled {{
    color: #9fb0ba;
    background: {PALETTE['bg']};
    border-color: {PALETTE['border']};
}}
QPushButton#scoreTick {{
    padding: 2px 0;
    font-weight: 700;
    color: #1b5e20;
}}
QPushButton#scoreTick:hover {{
    border-color: #2e7d32;
    color: #1b5e20;
    background: #e8f5e9;
}}
QPushButton#scoreCross {{
    padding: 2px 0;
    font-weight: 700;
    color: #b71c1c;
}}
QPushButton#scoreCross:hover {{
    border-color: #c62828;
    color: #b71c1c;
    background: #ffebee;
}}
QToolButton#paneChrome {{
    background: transparent;
    border: 1px solid transparent;
    border-radius: 4px;
    padding: 2px 8px;
    color: {PALETTE['muted']};
    font-size: 11px;
    font-weight: 600;
}}
QToolButton#paneChrome:hover {{
    border-color: {PALETTE['border_strong']};
    color: {PALETTE['accent']};
    background: #f5fafb;
}}

/* ---- Inputs ---- */
QLineEdit, QTextEdit, QPlainTextEdit, QComboBox {{
    background: {PALETTE['surface']};
    border: 1px solid {PALETTE['border_strong']};
    border-radius: 6px;
    padding: 5px 8px;
    selection-background-color: {PALETTE['accent']};
    selection-color: white;
}}
QLineEdit:focus, QTextEdit:focus, QPlainTextEdit:focus, QComboBox:focus {{
    border-color: {PALETTE['accent']};
}}
QComboBox::drop-down {{
    border: none;
    width: 24px;
}}
QComboBox QAbstractItemView {{
    background: {PALETTE['surface']};
    border: 1px solid {PALETTE['border']};
    selection-background-color: {PALETTE['accent_soft']};
    selection-color: {PALETTE['text']};
}}

/* ---- Trees / lists / tables ---- */
QTreeWidget, QListWidget {{
    background: {PALETTE['surface']};
    border: 1px solid {PALETTE['border']};
    border-radius: 8px;
    alternate-background-color: #f8fafb;
    outline: none;
}}
QTreeWidget::item {{
    padding: 3px 2px;
}}
QTreeWidget::item:selected, QListWidget::item:selected {{
    background: {PALETTE['accent_soft']};
    color: #0f3c46;
}}
QHeaderView::section {{
    background: {PALETTE['bg']};
    border: none;
    border-bottom: 1px solid {PALETTE['border']};
    border-right: 1px solid #eef2f5;
    padding: 6px 8px;
    font-weight: 600;
    color: {PALETTE['muted']};
}}

/* ---- Scrollbars: slim and quiet ---- */
QScrollBar:vertical {{
    background: transparent;
    width: 10px;
    margin: 2px;
}}
QScrollBar:horizontal {{
    background: transparent;
    height: 10px;
    margin: 2px;
}}
QScrollBar::handle:vertical, QScrollBar::handle:horizontal {{
    background: #c7d2da;
    border-radius: 5px;
    min-height: 30px;
    min-width: 30px;
}}
QScrollBar::handle:vertical:hover, QScrollBar::handle:horizontal:hover {{
    background: #9fb0ba;
}}
QScrollBar::add-line, QScrollBar::sub-line {{
    height: 0;
    width: 0;
}}
QScrollBar::add-page, QScrollBar::sub-page {{
    background: transparent;
}}

/* ---- Menus / status bar ---- */
QMenuBar {{
    background: {PALETTE['bg']};
    border-bottom: 1px solid {PALETTE['border']};
}}
QMenuBar::item {{
    padding: 6px 10px;
    background: transparent;
}}
QMenuBar::item:selected {{
    background: {PALETTE['accent_soft']};
    border-radius: 4px;
}}
QMenu {{
    background: {PALETTE['surface']};
    border: 1px solid {PALETTE['border']};
    border-radius: 6px;
    padding: 4px;
}}
QMenu::item {{
    padding: 6px 24px 6px 12px;
    border-radius: 4px;
}}
QMenu::item:selected {{
    background: {PALETTE['accent_soft']};
}}

/* System tray menu — native Windows / Outlook style */
QMenu#trayMenu {{
    background: #f2f2f2;
    color: #000000;
    border: 1px solid #d0d0d0;
    border-radius: 0px;
    padding: 2px 0px;
    font-family: 'Segoe UI';
    font-size: 9pt;
}}
QMenu#trayMenu::item {{
    background: transparent;
    color: #000000;
    padding: 4px 36px 4px 28px;
    border-radius: 0px;
    margin: 0px;
}}
QMenu#trayMenu::item:selected {{
    background: #cce8ff;
    color: #000000;
}}
QMenu#trayMenu::item:disabled {{
    color: #6d6d6d;
}}
QMenu#trayMenu::separator {{
    height: 1px;
    background: #d1d1d1;
    margin: 4px 0px;
}}
QMenu#trayMenu::indicator {{
    width: 13px;
    height: 13px;
    margin-left: 8px;
}}
QStatusBar {{
    background: {PALETTE['surface']};
    border-top: 1px solid {PALETTE['border_strong']};
    color: {PALETTE['text']};
    min-height: 26px;
    padding: 2px 8px;
}}
QStatusBar::item {{
    border: none;
}}
QStatusBar QLabel {{
    color: {PALETTE['muted']};
    padding: 0 8px;
}}

/* ---- Misc ---- */
QToolTip {{
    background: {PALETTE['text']};
    color: #eef4f6;
    border: none;
    padding: 6px 8px;
    border-radius: 4px;
}}
QSplitter::handle {{
    background: transparent;
}}
QFrame#connBar {{
    background: {PALETTE['surface']};
    border: 1px solid {PALETTE['border']};
    border-radius: 8px;
}}
QRadioButton, QCheckBox {{
    spacing: 6px;
}}
"""


class _ScrapeBridge(QObject):
    """Marshal Outlook scrape results back to the UI thread."""

    finished = Signal(object, str)  # items list or None, error message


class _OllamaBridge(QObject):
    """Marshal Ollama HTTP results back to the UI thread."""

    models_ready = Signal(object, str)  # list[str] or None, error
    chat_ready = Signal(str, str)  # response or "", error
    gpu_ready = Signal(object, str)  # report dict or None, error


class _AccountsBridge(QObject):
    """Marshal Outlook account discovery back to the UI thread."""

    ready = Signal(object, str)  # list[AesAccountScanConfig] or None, error


class _MigrateBridge(QObject):
    """Marshal MySQL/SQLite → PostgreSQL migration progress back to the UI."""

    progress = Signal(str)
    finished = Signal(object, str)  # report dict or None, error


class _FileScrapeBridge(QObject):
    """Marshal filesystem location scrape results back to the UI thread."""

    progress = Signal(str)
    finished = Signal(object, str)  # items list or None, error message


class _RuleCompileBridge(QObject):
    """Marshal rule-compile LLM results back to the UI thread."""

    finished = Signal(object, str)  # parsed dict or None, error/raw


class _InboxSearchBridge(QObject):
    """Marshal conversation-level inbox search back to the UI thread."""

    progress = Signal(str)
    finished = Signal(object, str)  # payload dict or None, error


GURI_LOGO_PATH = r"C:\GeoFooter\smart-ass-email.svg"
GURI_APP_ICON_PATH = r"C:\GeoFooter\abyitself.ico"


def _load_guri_logo_pixmap(height: int = 72, width: Optional[int] = None) -> Optional[QPixmap]:
    """Load the GURI brand SVG at a given height (native aspect)."""
    path = GURI_LOGO_PATH
    if not os.path.isfile(path):
        return None
    try:
        from PySide6.QtCore import QRectF
        from PySide6.QtSvg import QSvgRenderer

        renderer = QSvgRenderer(path)
        if not renderer.isValid():
            return None
        vb = renderer.viewBoxF()
        if vb.width() <= 0 or vb.height() <= 0:
            size = renderer.defaultSize()
            aspect = (size.width() / size.height()) if size.height() else 10.0
        else:
            aspect = vb.width() / vb.height()
        if width is not None and width > 0:
            height = max(1, int(round(width / aspect)))
        else:
            width = max(1, int(round(height * aspect)))
        pix = QPixmap(int(width), int(height))
        pix.fill(Qt.GlobalColor.transparent)
        painter = QPainter(pix)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing, True)
        renderer.render(painter, QRectF(0, 0, width, height))
        painter.end()
        return pix
    except Exception:
        return None


class GuriBrandHeader(QWidget):
    """Title banner from smart-ass-email.svg — aspect preserved, bars span full width."""

    # End-cap colours = first / last stops of the SVG rainbow (solid, no gradient)
    _BAR_LEFT = QColor("#ff1010")
    _BAR_RIGHT = QColor("#ff0101")
    # Absolute bar centre Y in SVG user units (layer y − translate.y), viewBox h ≈ 82.31
    # Paths: 102.22363, 132.58524, 162.94685 with translate(…, -89.587082)
    _BAR_VIEWBOX_Y = (12.63655, 42.99816, 73.35977)
    _BAR_VIEWBOX_H = 82.306489
    _BAR_STROKE = 4.05  # stroke-width in SVG

    def __init__(self, parent=None):
        super().__init__(parent)
        from PySide6.QtWidgets import QSizePolicy
        from PySide6.QtSvg import QSvgRenderer

        self.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        self.setFixedHeight(72)
        self._renderer: Optional[QSvgRenderer] = None
        self._aspect = self._BAR_VIEWBOX_H and (1615.1971 / self._BAR_VIEWBOX_H)
        self._vb_h = self._BAR_VIEWBOX_H
        self._bar_ys = self._BAR_VIEWBOX_Y
        if os.path.isfile(GURI_LOGO_PATH):
            try:
                renderer = QSvgRenderer(GURI_LOGO_PATH)
                if renderer.isValid():
                    self._renderer = renderer
                    vb = renderer.viewBoxF()
                    if vb.width() > 0 and vb.height() > 0:
                        self._aspect = vb.width() / vb.height()
                        self._vb_h = float(vb.height())
            except Exception:
                self._renderer = None

    def _bar_width(self, logo_h: float) -> float:
        return max(2.0, logo_h * (self._BAR_STROKE / self._vb_h))

    def _draw_solid_bar_extensions(
        self,
        painter: QPainter,
        *,
        logo_x: float,
        logo_w: float,
        logo_y: float,
        logo_h: float,
        widget_w: int,
    ) -> None:
        """Continue each bar to the window edges in solid end-cap colours."""
        from PySide6.QtCore import QLineF

        width = self._bar_width(logo_h)
        # Overlap into the logo so joins don’t show a hairline gap
        left_end = logo_x + 1.5
        right_start = logo_x + logo_w - 1.5
        scale = logo_h / self._vb_h

        for vb_y in self._bar_ys:
            # Map SVG user-unit Y through the same scale as the rendered logo
            y = logo_y + vb_y * scale
            if left_end > 1:
                pen = QPen(self._BAR_LEFT)
                pen.setWidthF(width)
                pen.setCapStyle(Qt.PenCapStyle.FlatCap)
                painter.setPen(pen)
                painter.drawLine(QLineF(0.0, y, left_end, y))
            if right_start < widget_w - 1:
                pen = QPen(self._BAR_RIGHT)
                pen.setWidthF(width)
                pen.setCapStyle(Qt.PenCapStyle.FlatCap)
                painter.setPen(pen)
                painter.drawLine(QLineF(right_start, y, float(widget_w), y))

    def paintEvent(self, event):
        from PySide6.QtCore import QRectF

        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing, True)
        w = self.width()
        h = self.height()
        painter.fillRect(0, 0, w, h, QColor(PALETTE["surface"]))

        if self._renderer is not None and self._renderer.isValid():
            # Fit by height — never stretch the wordmark
            dest_h = float(h)
            dest_w = dest_h * self._aspect
            if dest_w > w:
                dest_w = float(w)
                dest_h = dest_w / self._aspect
            x = (w - dest_w) / 2.0
            y = (h - dest_h) / 2.0

            # Solid end-caps first, then the SVG (gradient) on top in the centre
            self._draw_solid_bar_extensions(
                painter,
                logo_x=x,
                logo_w=dest_w,
                logo_y=y,
                logo_h=dest_h,
                widget_w=w,
            )
            self._renderer.render(painter, QRectF(x, y, dest_w, dest_h))
        else:
            # Fallback: solid bars edge-to-edge + centred pixmap if available
            self._draw_solid_bar_extensions(
                painter,
                logo_x=w / 2,
                logo_w=0,
                logo_y=0,
                logo_h=float(h),
                widget_w=w,
            )
            pix = _load_guri_logo_pixmap(height=max(40, h - 8))
            if pix is not None and not pix.isNull():
                px = (w - pix.width()) // 2
                py = (h - pix.height()) // 2
                painter.drawPixmap(px, py, pix)
            else:
                painter.setPen(QColor(PALETTE["text"]))
                painter.setFont(QFont("Segoe UI", 16, QFont.Weight.DemiBold))
                painter.drawText(
                    self.rect(),
                    Qt.AlignmentFlag.AlignCenter,
                    "Smart Ass Email",
                )


class VipPoliceLightOverlay(QWidget):
    """Bottom-right red spinning police-light flash for new VIP mail (2 seconds)."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setFixedSize(88, 88)
        self.setAttribute(Qt.WidgetAttribute.WA_TransparentForMouseEvents, True)
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground, True)
        self.setWindowFlags(Qt.WindowType.Widget)
        self.hide()
        self._angle = 0.0
        self._spin = QTimer(self)
        self._spin.timeout.connect(self._tick)
        self._hide_timer = QTimer(self)
        self._hide_timer.setSingleShot(True)
        self._hide_timer.timeout.connect(self._stop)

    def flash(self, duration_ms: int = 2000) -> None:
        parent = self.parentWidget()
        if parent is None:
            return
        self._angle = 0.0
        self._reposition()
        self.show()
        self.raise_()
        self._spin.start(33)
        self._hide_timer.start(max(400, int(duration_ms)))
        self.update()

    def _stop(self) -> None:
        self._spin.stop()
        self.hide()

    def _tick(self) -> None:
        self._angle = (self._angle + 22.0) % 360.0
        self.update()

    def _reposition(self) -> None:
        parent = self.parentWidget()
        if parent is None:
            return
        margin = 18
        # Sit above the status bar region
        x = max(0, parent.width() - self.width() - margin)
        y = max(0, parent.height() - self.height() - margin - 28)
        self.move(x, y)

    def paintEvent(self, event) -> None:  # noqa: N802
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing, True)
        cx = self.width() / 2.0
        cy = self.height() / 2.0
        radius = min(cx, cy) - 4.0

        # Soft red glow base
        glow = QRadialGradient(cx, cy, radius)
        glow.setColorAt(0.0, QColor(255, 40, 40, 210))
        glow.setColorAt(0.45, QColor(220, 0, 0, 140))
        glow.setColorAt(1.0, QColor(120, 0, 0, 0))
        painter.setPen(Qt.PenStyle.NoPen)
        painter.setBrush(QBrush(glow))
        painter.drawEllipse(QRectF(cx - radius, cy - radius, radius * 2, radius * 2))

        # Spinning conical "beacon" lobes
        painter.save()
        painter.translate(cx, cy)
        painter.rotate(self._angle)
        cone = QConicalGradient(0, 0, 0)
        cone.setColorAt(0.0, QColor(255, 255, 255, 230))
        cone.setColorAt(0.12, QColor(255, 60, 60, 200))
        cone.setColorAt(0.28, QColor(180, 0, 0, 0))
        cone.setColorAt(0.5, QColor(255, 255, 255, 200))
        cone.setColorAt(0.62, QColor(255, 40, 40, 180))
        cone.setColorAt(0.78, QColor(160, 0, 0, 0))
        cone.setColorAt(1.0, QColor(255, 255, 255, 230))
        painter.setBrush(QBrush(cone))
        painter.drawEllipse(QRectF(-radius * 0.92, -radius * 0.92, radius * 1.84, radius * 1.84))
        painter.restore()

        # Chrome dome / light housing
        painter.setBrush(QBrush(QColor(40, 40, 45, 230)))
        painter.setPen(QPen(QColor(20, 20, 24), 1.5))
        painter.drawEllipse(QRectF(cx - 12, cy - 12, 24, 24))
        painter.setBrush(QBrush(QColor(255, 80, 80, 240)))
        painter.setPen(Qt.PenStyle.NoPen)
        painter.drawEllipse(QRectF(cx - 6, cy - 6, 12, 12))

        # Tiny "VIP" caption
        painter.setPen(QColor(255, 240, 240))
        painter.setFont(QFont("Segoe UI", 8, QFont.Weight.Bold))
        painter.drawText(
            self.rect().adjusted(0, 0, 0, -4),
            Qt.AlignmentFlag.AlignHCenter | Qt.AlignmentFlag.AlignBottom,
            "VIP",
        )


class TimelineHoverPopup(QFrame):
    """Dropdown under a timeline circle — sender list; click loads email preview.

    Implemented as an in-window child (not a top-level Tool window) to avoid
    Windows focus deadlocks when switching between circles.
    """

    email_chosen = Signal(object)  # detection dict
    wrong_category_requested = Signal(object)  # timeline entry dict
    dismiss_requested = Signal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("timelineHoverPopup")
        self.setFrameShape(QFrame.Shape.StyledPanel)
        self.setFrameShadow(QFrame.Shadow.Raised)
        self.setStyleSheet(
            f"""
            QFrame#timelineHoverPopup {{
                background: #ffffff;
                border: 1px solid {PALETTE['border_strong']};
            }}
            QLabel#popupHeader {{
                color: {PALETTE['muted']};
                font-size: 8pt;
                padding: 4px 8px 2px 8px;
            }}
            QListWidget {{
                border: none;
                background: #ffffff;
                outline: none;
                font-family: 'Segoe UI';
                font-size: 9pt;
            }}
            QListWidget::item {{
                padding: 5px 10px;
                border-bottom: 1px solid #eef2f5;
            }}
            QListWidget::item:hover {{
                background: {PALETTE['accent_soft']};
            }}
            QListWidget::item:selected {{
                background: {PALETTE['accent_soft']};
                color: {PALETTE['text']};
            }}
            """
        )
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)
        self.header = QLabel("")
        self.header.setObjectName("popupHeader")
        layout.addWidget(self.header)
        self.list = QListWidget()
        self.list.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.list.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        self.list.itemClicked.connect(self._on_item_clicked)
        self.list.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.list.customContextMenuRequested.connect(self._on_list_context_menu)
        layout.addWidget(self.list)
        self.setFixedWidth(320)
        self.hide()
        self._pinned = False
        self._hide_timer = QTimer(self)
        self._hide_timer.setSingleShot(True)
        self._hide_timer.setInterval(220)
        self._hide_timer.timeout.connect(self._request_dismiss)

    def populate(self, title: str, entries: List[Dict[str, Any]]) -> None:
        self.header.setText(title)
        self.list.blockSignals(True)
        self.list.clear()
        for entry in entries[:40]:
            sender = str(entry.get("sender") or "(unknown sender)")
            subject = str(entry.get("subject") or "")
            when = str(entry.get("datetime_str") or "")
            # One line: sender — subject
            line = sender
            if subject:
                short = subject if len(subject) <= 48 else subject[:47] + "…"
                line = f"{sender}  —  {short}"
            item = QListWidgetItem(line)
            tip = f"{when}\n{sender}\n{subject}\n\nRight-click: Wrong categorisation".strip()
            item.setToolTip(tip)
            item.setData(Qt.ItemDataRole.UserRole, entry)
            self.list.addItem(item)
        self.list.blockSignals(False)
        rows = min(max(len(entries), 1), 8)
        self.list.setFixedHeight(rows * 28 + 4)
        self.adjustSize()

    def _entry_from_item(self, item: QListWidgetItem) -> Dict[str, Any]:
        entry = item.data(Qt.ItemDataRole.UserRole) or {}
        return entry if isinstance(entry, dict) else {}

    def _on_item_clicked(self, item: QListWidgetItem) -> None:
        entry = self._entry_from_item(item)
        det = entry.get("det") if isinstance(entry.get("det"), dict) else None
        if not det:
            det = {
                "mail": {
                    "sender": entry.get("sender"),
                    "subject": entry.get("subject"),
                    "received": entry.get("datetime_str"),
                    "entry_id": "",
                    "store_id": "",
                    "body_preview": "",
                },
                "actions": [],
                "reasons": [],
            }
        self._pinned = False
        self.email_chosen.emit(det)
        self.hide()

    def _on_list_context_menu(self, position) -> None:
        item = self.list.itemAt(position)
        if not item:
            return
        self.list.setCurrentItem(item)
        entry = self._entry_from_item(item)
        if not entry:
            return
        self.cancel_hide()
        self.pin()
        menu = QMenu(self)
        menu.addAction(
            "Wrong categorisation…",
            lambda: self.wrong_category_requested.emit(entry),
        )
        menu.exec(self.list.mapToGlobal(position))

    def schedule_hide(self) -> None:
        if self._pinned:
            return
        self._hide_timer.start()

    def cancel_hide(self) -> None:
        self._hide_timer.stop()

    def pin(self) -> None:
        """Keep open until another circle is chosen or an item is picked."""
        self._pinned = True
        self.cancel_hide()

    def unpin(self) -> None:
        self._pinned = False

    def _request_dismiss(self) -> None:
        if self._pinned:
            return
        self.hide()
        self.dismiss_requested.emit()

    def enterEvent(self, event):
        self.cancel_hide()
        super().enterEvent(event)

    def leaveEvent(self, event):
        if not self._pinned:
            self.schedule_hide()
        super().leaveEvent(event)


class TimelineCanvas(QWidget):
    """Custom widget for drawing email timeline in category rows.

    Row = triage category. Circle colours encode sender/domain importance
    (VIP / important domain / elevated / normal).
    """

    # hour_bucket (0..23), category key, global anchor QPoint (below circle)
    cell_hovered = Signal(int, str, object)
    cell_left = Signal()
    cell_clicked = Signal(int, str, object)

    # Vertical padding above/below each hour-cell circle
    CELL_PAD_V = 3
    # Circles are half column width so the timeline stays compact
    CIRCLE_SCALE = 0.5

    # Importance tiers (highest wins when a cell aggregates several mails)
    IMPORTANCE_RANK = {"normal": 0, "elevated": 1, "domain": 2, "vip": 3}

    IMPORTANCE_STYLES = {
        "vip": {
            "fill": QColor("#D1FAE5"),
            "stroke": QColor("#047857"),
            "text": QColor("#064E3B"),
        },
        "domain": {
            "fill": QColor("#FEF3C7"),
            "stroke": QColor("#B45309"),
            "text": QColor("#78350F"),
        },
        "elevated": {
            "fill": QColor("#FFEDD5"),
            "stroke": QColor("#C2410C"),
            "text": QColor("#7C2D12"),
        },
    }

    ROW_STYLES = {
        CATEGORY_IMPORTANT: {
            "fill": QColor("#D1FAE5"),
            "stroke": QColor("#047857"),
            "text": QColor("#065F46"),
            "band": QColor("#F0FDF4"),
        },
        CATEGORY_ACTIONABLE: {
            "fill": QColor("#FFF3E0"),
            "stroke": QColor("#E65100"),
            "text": QColor("#BF360C"),
            "band": QColor("#FFF8F0"),
        },
        CATEGORY_RELEVANT: {
            "fill": QColor("#E3F2FD"),
            "stroke": QColor("#1565C0"),
            "text": QColor("#0D47A1"),
            "band": QColor("#F5FAFF"),
        },
        CATEGORY_LESS: {
            "fill": QColor("#ECEFF1"),
            "stroke": QColor("#607D8B"),
            "text": QColor("#37474F"),
            "band": QColor("#F7F9FA"),
        },
        CATEGORY_DEPRECATED: {
            "fill": QColor("#F3E5F5"),
            "stroke": QColor("#7B1FA2"),
            "text": QColor("#4A148C"),
            "band": QColor("#FAF5FB"),
        },
    }

    def __init__(self, parent=None):
        super().__init__(parent)
        # hour_bucket -> {category: count} or {category: {count, importance}}
        self.email_data: Dict[int, Dict[str, Any]] = {}
        self._mode = "rolling24"  # "today" | "rolling24"
        self._visible_hours = 24
        self._margin_left = 196
        self._margin_right = 16
        self._margin_top = 26
        self._margin_bottom = 8
        self._hour_width = 1.0
        self._row_h = 1.0
        self._hover_key: Optional[tuple] = None
        from PySide6.QtWidgets import QSizePolicy

        self.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        self.setStyleSheet("background-color: white;")
        self.setMouseTracking(True)
        self.setCursor(Qt.CursorShape.ArrowCursor)
        self.setToolTip("")  # custom popup replaces Qt tooltip
        self._apply_compact_height()

    @classmethod
    def _cell_payload(cls, raw: Any) -> Dict[str, Any]:
        """Normalize a cell to {count, importance}."""
        if isinstance(raw, dict):
            count_raw = raw.get("count", 0)
            try:
                count = int(count_raw) if not isinstance(count_raw, dict) else 0
            except (TypeError, ValueError):
                count = 0
            return {
                "count": count,
                "importance": str(raw.get("importance") or "normal"),
            }
        if isinstance(raw, bool):
            return {"count": int(raw), "importance": "normal"}
        if isinstance(raw, (int, float)):
            return {"count": int(raw), "importance": "normal"}
        if isinstance(raw, str):
            try:
                return {"count": int(raw.strip() or "0"), "importance": "normal"}
            except ValueError:
                return {"count": 0, "importance": "normal"}
        return {"count": 0, "importance": "normal"}

    def _style_for_cell(self, category: str, importance: str) -> Dict[str, QColor]:
        base = dict(self.ROW_STYLES.get(category) or self.ROW_STYLES[CATEGORY_LESS])
        imp = self.IMPORTANCE_STYLES.get(str(importance or "normal"))
        if imp:
            base["fill"] = imp["fill"]
            base["stroke"] = imp["stroke"]
            base["text"] = imp["text"]
        return base

    def set_timeline_mode(self, mode: str, visible_hours: int = 24) -> None:
        """today: fixed 1/24 column width, hours 00..now left-aligned.
        rolling24: full 24-hour rolling window."""
        self._mode = "today" if mode == "today" else "rolling24"
        self._visible_hours = max(1, min(24, int(visible_hours or 24)))
        self.update()

    def _column_count(self) -> int:
        return self._visible_hours if self._mode == "today" else 24

    def _layout_metrics(self):
        width = max(1, self.width())
        # Always size as 1/24 of full plot so Today doesn't stretch columns
        hour_width = (width - self._margin_left - self._margin_right) / 24.0
        # Row height hugs the circle (half column width) + padding
        circle_d = max(8.0, (hour_width - 2.0) * self.CIRCLE_SCALE)
        row_h = circle_d + (2 * self.CELL_PAD_V)
        self._hour_width = hour_width
        self._row_h = row_h
        return hour_width, row_h

    def _preferred_height_for_width(self, width: int) -> int:
        hour_width = (max(1, width) - self._margin_left - self._margin_right) / 24.0
        circle_d = max(8.0, (hour_width - 2.0) * self.CIRCLE_SCALE)
        row_h = circle_d + (2 * self.CELL_PAD_V)
        return int(
            self._margin_top
            + len(TIMELINE_CATEGORIES) * row_h
            + self._margin_bottom
        )

    def _apply_compact_height(self) -> None:
        if getattr(self, "_syncing_height", False):
            return
        h = self._preferred_height_for_width(max(self.width(), 400))
        if self.minimumHeight() == h and self.maximumHeight() == h and self.height() == h:
            return
        self._syncing_height = True
        try:
            self.setFixedHeight(h)
        finally:
            self._syncing_height = False

    def _deferred_height_sync(self) -> None:
        self._height_sync_pending = False
        self._apply_compact_height()

    def resizeEvent(self, event):
        super().resizeEvent(event)
        # Defer — calling setFixedHeight inside resizeEvent during maximize
        # can re-enter layout and freeze the UI thread.
        if getattr(self, "_height_sync_pending", False):
            return
        self._height_sync_pending = True
        QTimer.singleShot(0, self._deferred_height_sync)

    def sizeHint(self):
        w = max(self.width(), 600)
        return QSize(w, self._preferred_height_for_width(w))

    def _cell_anchor_global(self, hour_bucket: int, category: str) -> QPoint:
        """Global point just below the circle centre for popup placement."""
        hour_width, row_h = self._layout_metrics()
        row_idx = TIMELINE_CATEGORIES.index(category) if category in TIMELINE_CATEGORIES else 0
        cx = self._margin_left + (hour_bucket + 0.5) * hour_width
        cy = self._margin_top + (row_idx + 0.72) * row_h
        return self.mapToGlobal(QPoint(int(cx), int(cy)))

    def _hit_test(self, pos) -> Optional[tuple]:
        """Return (hour_bucket, category, count) under pos, or None."""
        hour_width, row_h = self._layout_metrics()
        x = pos.x()
        y = pos.y()
        n_cols = self._column_count()
        used_right = self._margin_left + n_cols * hour_width
        if x < self._margin_left or x > used_right:
            return None
        if y < self._margin_top or y > self.height() - self._margin_bottom:
            return None
        hour_bucket = int((x - self._margin_left) / hour_width)
        if hour_bucket < 0 or hour_bucket >= n_cols:
            return None
        row_idx = int((y - self._margin_top) / row_h)
        if row_idx < 0 or row_idx >= len(TIMELINE_CATEGORIES):
            return None
        category = TIMELINE_CATEGORIES[row_idx]
        cat_counts = self.email_data.get(hour_bucket) or {}
        if not isinstance(cat_counts, dict):
            count = int(cat_counts or 0) if category == CATEGORY_LESS else 0
        else:
            count = int(self._cell_payload(cat_counts.get(category)).get("count") or 0)
        if count <= 0:
            return None
        return hour_bucket, category, count

    def mouseMoveEvent(self, event):
        hit = self._hit_test(event.position().toPoint())
        if hit:
            hour_bucket, category, _count = hit
            key = (hour_bucket, category)
            self.setCursor(Qt.CursorShape.PointingHandCursor)
            if key != self._hover_key:
                self._hover_key = key
                anchor = self._cell_anchor_global(hour_bucket, category)
                self.cell_hovered.emit(hour_bucket, category, anchor)
        else:
            self.setCursor(Qt.CursorShape.ArrowCursor)
            if self._hover_key is not None:
                self._hover_key = None
                self.cell_left.emit()
        super().mouseMoveEvent(event)

    def leaveEvent(self, event):
        if self._hover_key is not None:
            self._hover_key = None
            self.cell_left.emit()
        self.setCursor(Qt.CursorShape.ArrowCursor)
        super().leaveEvent(event)

    def mousePressEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            hit = self._hit_test(event.position().toPoint())
            if hit:
                hour_bucket, category, _count = hit
                anchor = self._cell_anchor_global(hour_bucket, category)
                self.cell_clicked.emit(hour_bucket, category, anchor)
                event.accept()
                return
            # Empty area — dismiss any open list
            self.cell_clicked.emit(-1, "", QPoint())
            event.accept()
            return
        super().mousePressEvent(event)

    def paintEvent(self, event):
        """Draw the category-row timeline grid."""
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)

        margin_left = self._margin_left
        margin_top = self._margin_top
        margin_bottom = self._margin_bottom
        hour_width, row_h = self._layout_metrics()
        height = self.height()
        n_cols = self._column_count()
        used_w = hour_width * n_cols
        pad_v = self.CELL_PAD_V

        current_time = datetime.now()

        # Row bands + labels
        label_font = QFont("Segoe UI", 8)
        label_font.setBold(True)
        hour_font = QFont("Segoe UI", 8)
        hour_font.setBold(True)

        for row_idx, category in enumerate(TIMELINE_CATEGORIES):
            style = self.ROW_STYLES[category]
            y0 = margin_top + row_idx * row_h
            painter.fillRect(
                int(margin_left),
                int(y0),
                int(used_w),
                int(row_h),
                style["band"],
            )
            painter.setPen(style["text"])
            painter.setFont(label_font)
            painter.drawText(
                8,
                int(y0),
                margin_left - 16,
                int(row_h),
                Qt.AlignmentFlag.AlignVCenter | Qt.AlignmentFlag.AlignLeft,
                TIMELINE_CATEGORY_LABELS[category],
            )
            # Soft row separator
            painter.setPen(QPen(QColor("#d5dde3"), 1))
            painter.drawLine(
                int(margin_left),
                int(y0 + row_h),
                int(margin_left + used_w),
                int(y0 + row_h),
            )

        # Hour labels and vertical grid (Today: only elapsed hours, left-aligned)
        for hour in range(n_cols + 1):
            x = margin_left + (hour * hour_width)
            painter.setPen(QColor("#d0d7de"))
            painter.drawLine(int(x), margin_top, int(x), height - margin_bottom)
            if hour < n_cols:
                if self._mode == "today":
                    hour_label = f"{hour:02d}"
                else:
                    hour_time = current_time - timedelta(hours=23 - hour)
                    hour_label = hour_time.strftime("%H")
                painter.setPen(QColor("#1a2e35"))
                painter.setFont(hour_font)
                painter.drawText(
                    int(x),
                    4,
                    int(hour_width),
                    max(16, margin_top - 6),
                    Qt.AlignmentFlag.AlignCenter,
                    hour_label,
                )

        # Circles: half column width so the timeline stays compact
        max_diam = max(8.0, row_h - (2 * pad_v))
        max_diam = min(max_diam, max(8.0, (hour_width - 2.0) * self.CIRCLE_SCALE))
        pen_w = 1 if max_diam < 16 else 2

        for hour_bucket, cat_counts in self.email_data.items():
            if hour_bucket < 0 or hour_bucket >= n_cols:
                continue
            if not isinstance(cat_counts, dict):
                # Backward compat: single count → less relevant
                cat_counts = {CATEGORY_LESS: int(cat_counts or 0)}
            for row_idx, category in enumerate(TIMELINE_CATEGORIES):
                payload = self._cell_payload(cat_counts.get(category))
                count = int(payload.get("count") or 0)
                if count <= 0:
                    continue
                style = self._style_for_cell(
                    category, str(payload.get("importance") or "normal")
                )
                x = margin_left + (hour_bucket * hour_width) + hour_width / 2
                y = margin_top + row_idx * row_h + row_h / 2
                # Slightly grow with count, but keep ≥5px vertical padding
                grow = 0.85 + 0.15 * min(count, 6) / 6.0
                diameter = max(8.0, max_diam * grow)
                diameter = min(diameter, row_h - (2 * pad_v))
                radius = diameter / 2.0
                painter.setBrush(style["fill"])
                circle_pen = QPen(style["stroke"])
                circle_pen.setWidth(pen_w)
                painter.setPen(circle_pen)
                painter.drawEllipse(
                    int(x - radius),
                    int(y - radius),
                    int(diameter),
                    int(diameter),
                )
                font_px = max(7, min(11, int(diameter * 0.55)))
                cell_font = QFont("Segoe UI", font_px)
                cell_font.setBold(True)
                painter.setFont(cell_font)
                painter.setPen(style["text"])
                text_box = max(int(diameter), 12)
                painter.drawText(
                    int(x - text_box / 2),
                    int(y - text_box / 2),
                    text_box,
                    text_box,
                    Qt.AlignmentFlag.AlignCenter,
                    str(count),
                )

    def update_email_data(self, email_data):
        """Update email data and schedule a repaint (never force sync repaint).

        Expected shape: {hour_bucket: {category: count | {count, importance}}}
        """
        self.email_data = email_data or {}
        self.update()


class DeadlineTimelineCanvas(QWidget):
    """Day-bucket timeline for extracted + manual deadlines."""

    cell_hovered = Signal(int, str, object)
    cell_left = Signal()
    cell_clicked = Signal(int, str, object)

    CELL_PAD_V = 5

    ROW_STYLES = {
        DL_CAT_EXTRACTED: {
            "fill": QColor("#E8F5E9"),
            "stroke": QColor("#2E7D32"),
            "text": QColor("#1B5E20"),
            "band": QColor("#F3FAF4"),
        },
        DL_CAT_MANUAL: {
            "fill": QColor("#FFF8E1"),
            "stroke": QColor("#F9A825"),
            "text": QColor("#F57F17"),
            "band": QColor("#FFFCF3"),
        },
    }

    def __init__(self, parent=None):
        super().__init__(parent)
        self.day_data: Dict[int, Dict[str, int]] = {}
        self._n_days = 30
        self._start_date = datetime.now().date()
        self._margin_left = 168
        self._margin_right = 16
        self._margin_top = 26
        self._margin_bottom = 8
        self._day_width = 1.0
        self._row_h = 1.0
        self._hover_key: Optional[tuple] = None
        from PySide6.QtWidgets import QSizePolicy

        self.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        self.setStyleSheet("background-color: white;")
        self.setMouseTracking(True)
        self.setCursor(Qt.CursorShape.ArrowCursor)
        self._apply_compact_height()

    def set_range(self, start_date, n_days: int = 30) -> None:
        self._start_date = start_date
        self._n_days = max(7, min(60, int(n_days or 30)))
        self.update()

    def _layout_metrics(self):
        width = max(1, self.width())
        day_width = (width - self._margin_left - self._margin_right) / float(self._n_days)
        circle_d = max(8.0, min(day_width - 2.0, 28.0))
        row_h = circle_d + (2 * self.CELL_PAD_V)
        self._day_width = day_width
        self._row_h = row_h
        return day_width, row_h

    def _preferred_height_for_width(self, width: int) -> int:
        day_width = (max(1, width) - self._margin_left - self._margin_right) / float(
            max(1, self._n_days)
        )
        circle_d = max(8.0, min(day_width - 2.0, 28.0))
        row_h = circle_d + (2 * self.CELL_PAD_V)
        return int(
            self._margin_top
            + len(DEADLINE_TIMELINE_CATEGORIES) * row_h
            + self._margin_bottom
        )

    def _apply_compact_height(self) -> None:
        if getattr(self, "_syncing_height", False):
            return
        h = self._preferred_height_for_width(max(self.width(), 400))
        if self.minimumHeight() == h and self.maximumHeight() == h and self.height() == h:
            return
        self._syncing_height = True
        try:
            self.setFixedHeight(h)
        finally:
            self._syncing_height = False

    def _deferred_height_sync(self) -> None:
        self._height_sync_pending = False
        self._apply_compact_height()

    def resizeEvent(self, event):
        super().resizeEvent(event)
        if getattr(self, "_height_sync_pending", False):
            return
        self._height_sync_pending = True
        QTimer.singleShot(0, self._deferred_height_sync)

    def sizeHint(self):
        w = max(self.width(), 600)
        return QSize(w, self._preferred_height_for_width(w))

    def _cell_anchor_global(self, day_bucket: int, category: str) -> QPoint:
        day_width, row_h = self._layout_metrics()
        row_idx = (
            DEADLINE_TIMELINE_CATEGORIES.index(category)
            if category in DEADLINE_TIMELINE_CATEGORIES
            else 0
        )
        cx = self._margin_left + (day_bucket + 0.5) * day_width
        cy = self._margin_top + (row_idx + 0.72) * row_h
        return self.mapToGlobal(QPoint(int(cx), int(cy)))

    def _hit_test(self, pos) -> Optional[tuple]:
        day_width, row_h = self._layout_metrics()
        x = pos.x()
        y = pos.y()
        used_right = self._margin_left + self._n_days * day_width
        if x < self._margin_left or x > used_right:
            return None
        if y < self._margin_top or y > self.height() - self._margin_bottom:
            return None
        day_bucket = int((x - self._margin_left) / day_width)
        if day_bucket < 0 or day_bucket >= self._n_days:
            return None
        row_idx = int((y - self._margin_top) / row_h)
        if row_idx < 0 or row_idx >= len(DEADLINE_TIMELINE_CATEGORIES):
            return None
        category = DEADLINE_TIMELINE_CATEGORIES[row_idx]
        cat_counts = self.day_data.get(day_bucket) or {}
        count = int(cat_counts.get(category) or 0) if isinstance(cat_counts, dict) else 0
        if count <= 0:
            return None
        return day_bucket, category, count

    def mouseMoveEvent(self, event):
        hit = self._hit_test(event.position().toPoint())
        if hit:
            day_bucket, category, _count = hit
            key = (day_bucket, category)
            self.setCursor(Qt.CursorShape.PointingHandCursor)
            if key != self._hover_key:
                self._hover_key = key
                anchor = self._cell_anchor_global(day_bucket, category)
                self.cell_hovered.emit(day_bucket, category, anchor)
        else:
            self.setCursor(Qt.CursorShape.ArrowCursor)
            if self._hover_key is not None:
                self._hover_key = None
                self.cell_left.emit()
        super().mouseMoveEvent(event)

    def leaveEvent(self, event):
        if self._hover_key is not None:
            self._hover_key = None
            self.cell_left.emit()
        self.setCursor(Qt.CursorShape.ArrowCursor)
        super().leaveEvent(event)

    def mousePressEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            hit = self._hit_test(event.position().toPoint())
            if hit:
                day_bucket, category, _count = hit
                anchor = self._cell_anchor_global(day_bucket, category)
                self.cell_clicked.emit(day_bucket, category, anchor)
                event.accept()
                return
            self.cell_clicked.emit(-1, "", QPoint())
            event.accept()
            return
        super().mousePressEvent(event)

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)

        margin_left = self._margin_left
        margin_top = self._margin_top
        margin_bottom = self._margin_bottom
        day_width, row_h = self._layout_metrics()
        height = self.height()
        used_w = day_width * self._n_days
        pad_v = self.CELL_PAD_V
        today = datetime.now().date()

        label_font = QFont("Segoe UI", 8)
        label_font.setBold(True)
        day_font = QFont("Segoe UI", 7)
        day_font.setBold(True)

        for row_idx, category in enumerate(DEADLINE_TIMELINE_CATEGORIES):
            style = self.ROW_STYLES[category]
            y0 = margin_top + row_idx * row_h
            painter.fillRect(
                int(margin_left),
                int(y0),
                int(used_w),
                int(row_h),
                style["band"],
            )
            painter.setPen(style["text"])
            painter.setFont(label_font)
            painter.drawText(
                8,
                int(y0),
                margin_left - 16,
                int(row_h),
                Qt.AlignmentFlag.AlignVCenter | Qt.AlignmentFlag.AlignLeft,
                DEADLINE_TIMELINE_LABELS[category],
            )
            painter.setPen(QPen(QColor("#d5dde3"), 1))
            painter.drawLine(
                int(margin_left),
                int(y0 + row_h),
                int(margin_left + used_w),
                int(y0 + row_h),
            )

        for day in range(self._n_days + 1):
            x = margin_left + (day * day_width)
            painter.setPen(QColor("#d0d7de"))
            painter.drawLine(int(x), margin_top, int(x), height - margin_bottom)
            if day < self._n_days:
                d = self._start_date + timedelta(days=day)
                label = d.strftime("%d")
                if d == today:
                    painter.setPen(QColor("#c62828"))
                elif d.weekday() >= 5:
                    painter.setPen(QColor("#90a4ae"))
                else:
                    painter.setPen(QColor("#1a2e35"))
                painter.setFont(day_font)
                painter.drawText(
                    int(x),
                    4,
                    int(day_width),
                    max(16, margin_top - 6),
                    Qt.AlignmentFlag.AlignCenter,
                    label,
                )

        max_diam = max(8.0, row_h - (2 * pad_v))
        max_diam = min(max_diam, max(8.0, day_width - 2.0))
        pen_w = 1 if max_diam < 16 else 2

        for day_bucket, cat_counts in self.day_data.items():
            if day_bucket < 0 or day_bucket >= self._n_days:
                continue
            if not isinstance(cat_counts, dict):
                continue
            for row_idx, category in enumerate(DEADLINE_TIMELINE_CATEGORIES):
                count = int(cat_counts.get(category) or 0)
                if count <= 0:
                    continue
                style = self.ROW_STYLES[category]
                x = margin_left + (day_bucket * day_width) + day_width / 2
                y = margin_top + row_idx * row_h + row_h / 2
                grow = 0.85 + 0.15 * min(count, 6) / 6.0
                diameter = max(8.0, max_diam * grow)
                diameter = min(diameter, row_h - (2 * pad_v))
                radius = diameter / 2.0
                painter.setBrush(style["fill"])
                circle_pen = QPen(style["stroke"])
                circle_pen.setWidth(pen_w)
                painter.setPen(circle_pen)
                painter.drawEllipse(
                    int(x - radius),
                    int(y - radius),
                    int(diameter),
                    int(diameter),
                )
                font_px = max(7, min(11, int(diameter * 0.55)))
                cell_font = QFont("Segoe UI", font_px)
                cell_font.setBold(True)
                painter.setFont(cell_font)
                painter.setPen(style["text"])
                text_box = max(int(diameter), 12)
                painter.drawText(
                    int(x - text_box / 2),
                    int(y - text_box / 2),
                    text_box,
                    text_box,
                    Qt.AlignmentFlag.AlignCenter,
                    str(count),
                )

    def update_day_data(self, day_data) -> None:
        self.day_data = day_data or {}
        self.update()


# Document type codes and descriptions (using 0x hex notation)
DOCUMENT_TYPES = {
    # Category 1: Email (starts at 0x01)
    "01": "Email",
    
    # Category 2: General Documents (starts at 0x0a = 0x01+9)
    "0a": "Word Document",
    "0b": "Excel Spreadsheet",
    "0c": "PowerPoint Presentation",
    "0d": "PDF Document",
    "0e": "Text File",
    "0f": "Image File",
    "10": "Video File",
    "11": "Audio File",
    "12": "Archive/Zip",
    
    # Category 3: Aliniant Internal Documents (starts at 0x20 = decimal 32)
    "20": "Aliniant Policy Document",
    "21": "Aliniant Technical Document - .docx",
    "22": "Aliniant Technical Document - .pptx",
    "23": "Aliniant Compliance Document",
    "24": "Aliniant Pre-sales Document",
    "25": "Aliniant Contract Document",
    "26": "Legal Filing",
    "27": "Aliniant Finance - Invoice Out",
    "2c": "Aliniant Financial - Invoice In",
    
    # Category 4: Sensitive Documents (alphabetical order)
    "2a": "Protocol A",
    "2b": "Protocol B",
    "29": "Sensitive Correspondence",
    "28": "Sensitive Internal",
    
    # Other
    "99": "Other"
}

# Global type classification
GLOBAL_TYPES = {
    # Global Type 1: Email
    "01": "Global Type 1: Email",
    
    # Global Type 2: General Documents
    "0a": "Global Type 2: General Documents",
    "0b": "Global Type 2: General Documents",
    "0c": "Global Type 2: General Documents",
    "0d": "Global Type 2: General Documents",
    "0e": "Global Type 2: General Documents",
    "0f": "Global Type 2: General Documents",
    "10": "Global Type 2: General Documents",
    "11": "Global Type 2: General Documents",
    "12": "Global Type 2: General Documents",
    
    # Global Type 3: Aliniant Internal Documents (starts at 0x20)
    "20": "Global Type 3: Aliniant Internal Documents",
    "21": "Global Type 3: Aliniant Internal Documents",
    "22": "Global Type 3: Aliniant Internal Documents",
    "23": "Global Type 3: Aliniant Internal Documents",
    "24": "Global Type 3: Aliniant Internal Documents",
    "25": "Global Type 3: Aliniant Internal Documents",
    "26": "Global Type 3: Aliniant Internal Documents",
    "27": "Global Type 3: Aliniant Internal Documents",
    "2c": "Global Type 3: Aliniant Internal Documents",
    
    # Global Type 4: Sensitive Documents (alphabetical order)
    "2a": "Global Type 4: Sensitive Documents",
    "2b": "Global Type 4: Sensitive Documents",
    "29": "Global Type 4: Sensitive Documents",
    "28": "Global Type 4: Sensitive Documents",
    
    # Other
    "99": "Global Type 2: General Documents"
}

# Default values for non-email documents
NON_EMAIL_SENDER = "julian.garrett@aliniant.com"
NON_EMAIL_PLACEHOLDER = "FFFFF"
MY_ALINIANT_EMAIL = "julian.garrett@aliniant.com"


def _parse_record_datetime(value) -> Optional[datetime]:
    """Parse a GURI record datetime from string or datetime object."""
    if not value:
        return None
    if isinstance(value, datetime):
        return value
    value_str = str(value).strip()
    for fmt in ('%Y-%m-%d %H:%M:%S', '%Y-%m-%d %H:%M:%S.%f'):
        try:
            return datetime.strptime(value_str, fmt)
        except ValueError:
            continue
    return None


class _SortableTreeItem(QTreeWidgetItem):
    """Sorts on SORT_ROLE data when a column sets it, otherwise on text."""

    SORT_ROLE = Qt.ItemDataRole.UserRole + 1

    def __lt__(self, other):
        tree = self.treeWidget()
        col = tree.sortColumn() if tree is not None else 0
        mine = self.data(col, self.SORT_ROLE)
        theirs = other.data(col, self.SORT_ROLE)
        if mine is not None and theirs is not None:
            return mine < theirs
        return self.text(col).lower() < other.text(col).lower()


class GURIViewerGUI(QMainWindow):
    """Main GUI application for GURI database viewing and management."""
    
    def __init__(self):
        """Initialize the GUI application."""
        super().__init__()
        self.setWindowTitle(f"{APP_NAME} Database Viewer & Manager - v{version_string()}")
        self._apply_app_icon()
        
        # Setup logging
        logging.basicConfig(level=logging.INFO)
        self.logger = logging.getLogger(__name__)
        
        # Initialize database connection
        self.db = None
        self.db_type = "postgres"  # PostgreSQL required (MySQL/SQLite deprecated)
        self.current_records = []
        self.current_page = 0
        self.records_per_page = 50
        
        # Initialize recipients list for autocomplete
        self.recipients_list = []
        self._load_recipients()
        
        # Initialize important domains list
        self.important_domains = []
        self._load_important_domains()
        
        # Email reply-status tracking (guri -> status label)
        self.email_status = {}
        self._load_email_status()

        # Live Outlook scrape / action inbox
        self.action_status = load_action_status()  # entry_id -> status
        self.relevance_feedback = load_relevance_feedback()
        self.scrape_cache_items: List[Dict[str, Any]] = []
        self.scrape_days = 7
        self.scrape_unread_only = False
        self.scrape_lookback_months = 0
        self.scrape_lookback_locked = False
        self._scrape_busy = False
        self._scrape_bridge = _ScrapeBridge()
        self._scrape_bridge.finished.connect(self._on_scrape_finished)
        self._load_scrape_settings()
        self._load_scrape_cache_into_memory()
        self._welcome_refresh_busy = False
        self._welcome_refresh_scheduled = False
        self._pending_welcome_db_fallback = False
        self._expanding_to_screen = False

        self._ollama_bridge = _OllamaBridge()
        self._ollama_bridge.models_ready.connect(self._on_ollama_models_ready)
        self._ollama_bridge.chat_ready.connect(self._on_ollama_chat_ready)
        self._ollama_bridge.gpu_ready.connect(self._on_ollama_gpu_ready)
        self._ollama_busy = False
        self._ollama_gpu_busy = False
        self.ollama_cfg = ollama_load_config()

        self._accounts_bridge = _AccountsBridge()
        self._accounts_bridge.ready.connect(self._on_accounts_discovered)
        self._accounts_busy = False
        self._account_rows: List[Dict[str, Any]] = []
        self._inbox_search_bridge = _InboxSearchBridge()
        self._inbox_search_bridge.progress.connect(self._on_inbox_search_progress)
        self._inbox_search_bridge.finished.connect(self._on_inbox_search_finished)
        self._inbox_search_busy = False
        self._inbox_search_hits: List[Dict[str, Any]] = []
        try:
            self._inbox_search_settings = load_inbox_search_settings()
        except Exception:
            self._inbox_search_settings = {}
        self._migrate_bridge = _MigrateBridge()
        self._migrate_bridge.progress.connect(self._on_migrate_progress)
        self._migrate_bridge.finished.connect(self._on_migrate_finished)
        self._migrate_busy = False
        self._migrate_dialog = None
        self._file_scrape_bridge = _FileScrapeBridge()
        self._file_scrape_bridge.progress.connect(self._on_file_scrape_progress)
        self._file_scrape_bridge.finished.connect(self._on_file_scrape_finished)
        self._file_scrape_busy = False
        self.file_scrape_items: List[Dict[str, Any]] = []
        self._scan_locations: List[ScanLocation] = []
        self._rule_compile_bridge = _RuleCompileBridge()
        self._rule_compile_bridge.finished.connect(self._on_rule_compile_finished)
        self._rule_compile_busy = False
        self._rule_compile_pending: Optional[Dict[str, Any]] = None
        try:
            self.deadline_account_keys = load_deadline_account_keys()
        except Exception:
            self.deadline_account_keys = []
        try:
            self.manual_deadlines = load_manual_deadlines()
        except Exception:
            self.manual_deadlines = []
        try:
            self.learning = load_learning()
        except Exception:
            from guri_learning import normalize_learning

            self.learning = normalize_learning({})
        self._seed_identity_defaults()
        try:
            cache = load_file_scrape_cache()
            self.file_scrape_items = list(cache.get("items") or [])
        except Exception:
            self.file_scrape_items = []
        try:
            self._scan_locations = load_scan_locations()
        except Exception:
            self._scan_locations = []
        self._deadline_timeline_items: List[Dict[str, Any]] = []
        self._deadline_timeline_start = datetime.now().date()
        self._deadline_timeline_days = 30
        
        # Initialize auto-refresh timer (Welcome panels every 60s; Outlook scrape is slower)
        self.auto_refresh_timer = QTimer()
        self.auto_refresh_timer.timeout.connect(self._auto_refresh_welcome)
        self._auto_scrape_enabled = True
        self._auto_scrape_interval_ms = 5 * 60 * 1000  # 5 min — Outlook COM is expensive
        self._last_outlook_scrape_at: Optional[datetime] = None

        # Outlook scrape timer (separate from UI redraw so we don't hammer COM every minute)
        self.scrape_timer = QTimer()
        self.scrape_timer.setInterval(self._auto_scrape_interval_ms)
        self.scrape_timer.timeout.connect(self._auto_scrape_outlook)
        
        # Initialize timeline email data storage
        self.timeline_email_data = {}
        
        # Initialize decompositor
        self.decompositor = None
        
        # Initialize component library
        self.component_library = None
        
        # System tray icon
        self.tray_icon = None
        self.tray_available = QSystemTrayIcon.isSystemTrayAvailable()
        
        # Create GUI components (status bar first — tab refreshes report to it)
        self._create_menu()
        self._create_connection_frame()
        self._create_status_bar()
        self._create_notebook()

        # VIP arrival beacon (bottom-right spinning red light)
        self._vip_police_light = VipPoliceLightOverlay(self)
        self._seen_mail_keys: Set[str] = set()
        
        # Create system tray icon
        if self.tray_available:
            self._create_system_tray()
        
        # Auto-connect to database (PostgreSQL by default, falls back to SQLite)
        self._connect_database_default()
        # Seed "already seen" keys so startup scrape doesn't siren for old VIP mail
        self._seed_seen_mail_keys(self.scrape_cache_items)

    def resizeEvent(self, event) -> None:  # noqa: N802
        super().resizeEvent(event)
        light = getattr(self, "_vip_police_light", None)
        if light is not None:
            light._reposition()

    def _create_menu(self):
        """Create menu bar."""
        menubar = self.menuBar()
        
        # File menu
        file_menu = menubar.addMenu("File")
        file_menu.addAction("Connect to PostgreSQL...", self._show_postgres_connection_dialog)
        file_menu.addAction(
            "Connect to SQLite (deprecated)…",
            self._show_sqlite_connection_dialog,
        )
        file_menu.addSeparator()
        file_menu.addAction("Export Records...", self._export_records)
        file_menu.addSeparator()
        file_menu.addAction("Exit", self.close)
        
        # Tools menu
        tools_menu = menubar.addMenu("Tools")
        tools_menu.addAction("Refresh Records", self._load_records)
        tools_menu.addAction("Search", self._focus_search)
        tools_menu.addAction("Inbox Search", self._focus_inbox_search)
        tools_menu.addSeparator()
        tools_menu.addAction("Database Statistics", self._show_statistics)
        tools_menu.addAction("Refresh Database Browser", self._refresh_database_browser)
        tools_menu.addAction("Migrate to PostgreSQL…", self._show_migrate_dialog)
        tools_menu.addAction(
            "Repair GURI component tags…",
            self._repair_guri_components,
        )
        tools_menu.addSeparator()
        tools_menu.addAction("Scan locations…", self._focus_locations_tab)
        tools_menu.addAction("Scrape Files Now", self._start_file_scrape)
        
        # Help menu
        help_menu = menubar.addMenu("Help")
        help_menu.addAction("About", self._show_about)
        help_menu.addAction("Document Type Codes", self._show_doc_types)
    
    def _create_connection_frame(self):
        """Create the slim connection status bar."""
        conn_frame = QFrame()
        conn_frame.setObjectName("connBar")
        conn_layout = QHBoxLayout()
        conn_layout.setContentsMargins(12, 6, 8, 6)
        conn_layout.setSpacing(8)
        conn_frame.setLayout(conn_layout)

        self.conn_dot = QLabel("●")
        self.conn_dot.setStyleSheet(f"color: {PALETTE['err']}; font-size: 10pt;")
        conn_layout.addWidget(self.conn_dot)

        self.conn_label = QLabel("Not connected")
        self.conn_label.setStyleSheet(f"color: {PALETTE['muted']};")
        conn_layout.addWidget(self.conn_label)

        conn_layout.addStretch()

        pg_btn = QPushButton("PostgreSQL")
        pg_btn.clicked.connect(self._show_postgres_connection_dialog)
        conn_layout.addWidget(pg_btn)

        db_tab_btn = QPushButton("Database…")
        db_tab_btn.setToolTip("Open the Database browser tab")
        db_tab_btn.clicked.connect(self._focus_database_tab)
        conn_layout.addWidget(db_tab_btn)

        # Add to main widget
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        self.main_layout = QVBoxLayout()
        self.main_layout.setContentsMargins(10, 8, 10, 6)
        self.main_layout.setSpacing(8)
        central_widget.setLayout(self.main_layout)
        self.main_layout.addWidget(conn_frame)

    def _set_conn_status(self, text: str, ok: bool) -> None:
        """Update connection bar text + dot color (and status bar summary)."""
        self.conn_label.setText(text)
        color = PALETTE["ok"] if ok else PALETTE["err"]
        text_color = PALETTE["text"] if ok else PALETTE["muted"]
        self.conn_dot.setStyleSheet(f"color: {color}; font-size: 10pt;")
        self.conn_label.setStyleSheet(f"color: {text_color};")
        if hasattr(self, "status_db_label"):
            self.status_db_label.setText(text if ok else "DB: not connected")
    
    def _create_notebook(self):
        """Create main notebook with tabs."""
        self.notebook = QTabWidget()
        
        # Tab 0: About / version
        self._create_about_tab()

        # Tab 1: Welcome / Dashboard
        self._create_welcome_tab()
        
        # Tab 1: View Records
        self._create_view_tab()

        # Tab 2: Database browser (engine, schemas, tables, records)
        self._create_database_tab()
        
        # Tab 3: Create GURI
        self._create_create_tab()
        
        # Tab 4: Search
        self._create_search_tab()
        
        # Tab 5: Decompositor
        self._create_decompositor_tab()
        
        # Tab 6: Suppliers & Contracts
        self._create_suppliers_tab()
        
        # Tab 7: Library
        self._create_library_tab()

        # Tab 8: Email accounts (AES / GURI scrape rules)
        self._create_accounts_tab()

        # Tab 8b: Conversation / sentiment inbox search
        self._create_inbox_search_tab()

        # Tab 8c: AES per-sender threat scores (view / override / reset)
        self._create_aes_scores_tab()

        # Tab 8d: Data broker removal (Incogni-style tracking)
        self._create_data_brokers_tab()

        # Tab 9: Filesystem scan locations
        self._create_locations_tab()

        # Tab 10: Deadlines from scraped mail
        self._create_deadlines_tab()

        # Tab 10: Deadline timeline (extracted + manual by due date)
        self._create_deadline_timeline_tab()

        # Tab 11: 30-day deadline calendar
        self._create_calendar_tab()

        # Tab 12: Ollama
        self._create_ollama_tab()
        
        # Add notebook to main layout
        self.main_layout.addWidget(self.notebook)
    
    def _create_about_tab(self):
        """About / version tab (leftmost)."""
        frame = QWidget()
        self.notebook.addTab(frame, "About")
        root = QVBoxLayout(frame)
        root.setContentsMargins(20, 20, 20, 20)
        root.setSpacing(12)

        title = QLabel(f"{APP_FULL_NAME}  {APP_SILLY_QUOTE}")
        title.setStyleSheet(
            f"font-size: 18px; font-weight: 700; color: {PALETTE['accent']};"
        )
        title.setWordWrap(True)
        root.addWidget(title)

        ver = QLabel(f"{SUITE_NAME} suite  ·  Version {version_string()}")
        ver.setStyleSheet("font-size: 14px; font-weight: 600;")
        root.addWidget(ver)

        org = QLabel(f"© {COPYRIGHT_YEAR} {APP_ORG}  ·  {suite_banner()}")
        org.setStyleSheet(f"color: {PALETTE['muted']};")
        root.addWidget(org)

        blurb = QLabel(
            "GURI manages Globally Unique Record Identifiers for email and documents, "
            "scrapes Outlook for actions and deadlines, learns from your corrections, "
            "and can compile natural-language rules for local triage (with Ollama assist)."
        )
        blurb.setWordWrap(True)
        root.addWidget(blurb)

        notes_box = QGroupBox("Release notes")
        notes_layout = QVBoxLayout(notes_box)
        notes = QTextEdit()
        notes.setReadOnly(True)
        lines = []
        for entry in RELEASE_NOTES:
            if len(entry) >= 3:
                ver_s, stamp, note = entry[0], entry[1], entry[2]
                lines.append(f"v{ver_s}  ·  {stamp}\n{note}\n")
            else:
                ver_s, note = entry[0], entry[1]
                lines.append(f"v{ver_s}\n{note}\n")
        notes.setPlainText("\n".join(lines).strip())
        notes_layout.addWidget(notes)
        root.addWidget(notes_box, stretch=1)

        path_lbl = QLabel(
            "Runtime learning & scrape state: %LOCALAPPDATA%\\GeoFooter\\\n"
            f"version.py → {version_string()}"
        )
        path_lbl.setStyleSheet(f"color: {PALETTE['muted']}; font-size: 11px;")
        path_lbl.setWordWrap(True)
        root.addWidget(path_lbl)

    def _create_welcome_tab(self):
        """Create the welcome/dashboard tab."""
        welcome_frame = QWidget()
        self.notebook.addTab(welcome_frame, "Welcome")
        
        # Main content container
        content = QWidget()
        content_layout = QVBoxLayout()
        content.setLayout(content_layout)
        content_layout.setContentsMargins(12, 4, 12, 12)

        # Full-width rainbow brand header
        self.brand_header = GuriBrandHeader()
        content_layout.addWidget(self.brand_header)
        
        # Create horizontal split: LEFT side and RIGHT side
        main_split = QSplitter(Qt.Orientation.Horizontal)
        self.welcome_main_split = main_split
        
        # LEFT SIDE - Timeline + (Important Emails ⇄ Actions) splitter
        left_widget = QWidget()
        left_layout = QVBoxLayout()
        left_layout.setContentsMargins(0, 0, 0, 0)
        left_layout.setSpacing(8)
        left_widget.setLayout(left_layout)
        self.welcome_left_widget = left_widget

        from PySide6.QtWidgets import QSizePolicy
        
        # Email Timeline Section (compact — freed space goes to mail panes below)
        timeline_frame = QGroupBox("Email Timeline — Last 24 Hours")
        timeline_layout = QVBoxLayout()
        timeline_layout.setContentsMargins(8, 6, 8, 6)
        timeline_layout.setSpacing(4)
        timeline_frame.setLayout(timeline_layout)
        timeline_frame.setSizePolicy(
            QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Maximum
        )
        self.timeline_frame = timeline_frame
        self._timeline_mode = "rolling24"
        
        # Timeline info row: count · mode radios · refresh
        info_frame = QWidget()
        info_layout = QHBoxLayout()
        info_layout.setContentsMargins(0, 0, 0, 0)
        info_frame.setLayout(info_layout)
        
        self.timeline_count_label = QLabel("Emails: 0")
        font = QFont('Segoe UI', 9)
        font.setBold(True)
        self.timeline_count_label.setFont(font)
        self.timeline_count_label.setStyleSheet(f"color: {PALETTE['muted']};")
        info_layout.addWidget(self.timeline_count_label)

        # Gap between Deprecated counts and account filters
        accounts_gap = QWidget()
        accounts_gap.setFixedWidth(100)
        info_layout.addWidget(accounts_gap)

        accounts_heading = QLabel("Email Accounts")
        accounts_heading_font = QFont("Segoe UI", 9)
        accounts_heading_font.setBold(True)
        accounts_heading.setFont(accounts_heading_font)
        accounts_heading.setStyleSheet(f"color: {PALETTE['muted']};")
        info_layout.addWidget(accounts_heading)

        # Account filters — which mailboxes contribute to the timeline
        self._timeline_account_checks: Dict[str, QCheckBox] = {}
        account_cb_font = QFont("Segoe UI", 9)
        account_cb_font.setBold(True)
        for key, label, tip in (
            (
                "aliniant",
                "Aliniant",
                "julian.garrett@aliniant.com and administrator@aliniant.com",
            ),
            ("coleago", "Coleago", "Coleago mailbox(es)"),
            ("txo", "TXO", "TXO mailbox(es)"),
            ("others", "Others", "Hotmail / Gmail / Outlook and any other accounts"),
        ):
            cb = QCheckBox(label)
            cb.setChecked(True)
            cb.setToolTip(tip)
            cb.setFont(account_cb_font)
            cb.setStyleSheet(f"color: {PALETTE['muted']}; font-size: 11px; font-weight: 700;")
            cb.toggled.connect(self._on_timeline_account_filter_toggled)
            self._timeline_account_checks[key] = cb
            info_layout.addWidget(cb)

        info_layout.addStretch()

        self.timeline_radio_today = QRadioButton("Today")
        self.timeline_radio_24h = QRadioButton("Last 24 Hours")
        self.timeline_radio_24h.setChecked(True)
        self.timeline_mode_group = QButtonGroup(self)
        self.timeline_mode_group.addButton(self.timeline_radio_today)
        self.timeline_mode_group.addButton(self.timeline_radio_24h)
        self.timeline_radio_today.toggled.connect(self._on_timeline_mode_toggled)
        info_layout.addWidget(self.timeline_radio_today)
        info_layout.addWidget(self.timeline_radio_24h)
        
        refresh_timeline_btn = QPushButton("Refresh")
        refresh_timeline_btn.clicked.connect(self._refresh_email_timeline)
        info_layout.addWidget(refresh_timeline_btn)
        
        timeline_layout.addWidget(info_frame)

        legend = QLabel(
            "Circle colour = importance — "
            "<span style='color:#047857;font-weight:700;'>● VIP sender</span>  "
            "<span style='color:#B45309;font-weight:700;'>● Important domain</span>  "
            "<span style='color:#C2410C;font-weight:700;'>● Elevated</span>  "
            "<span style='color:#607D8B;'>● Normal (row colour)</span>"
            "  ·  Row = category"
        )
        legend.setTextFormat(Qt.TextFormat.RichText)
        legend.setWordWrap(True)
        legend.setStyleSheet(f"color: {PALETTE['muted']}; font-size: 11px;")
        timeline_layout.addWidget(legend)
        
        # Create canvas for hour grid
        self.timeline_canvas = TimelineCanvas()
        self.timeline_canvas.cell_hovered.connect(self._on_timeline_cell_hovered)
        self.timeline_canvas.cell_left.connect(self._on_timeline_cell_left)
        self.timeline_canvas.cell_clicked.connect(self._on_timeline_cell_clicked)
        self._timeline_emails: List[Dict[str, Any]] = []
        self._timeline_hover_key: Optional[tuple] = None
        self._timeline_popup_busy = False
        # Parent to the timeline group so the dropdown is an in-window overlay
        # (avoids Windows Tool-window focus freezes on the second click).
        self._timeline_popup = TimelineHoverPopup(timeline_frame)
        self._timeline_popup.email_chosen.connect(self._on_timeline_email_chosen)
        self._timeline_popup.wrong_category_requested.connect(
            self._on_timeline_wrong_category
        )
        timeline_layout.addWidget(self.timeline_canvas)
        
        left_layout.addWidget(timeline_frame, stretch=0)
        
        # Important Emails — 3 columns: domains | emails | stacked actions
        important_frame = QGroupBox("Important Emails")
        important_layout = QHBoxLayout()
        important_layout.setContentsMargins(8, 8, 8, 8)
        important_layout.setSpacing(10)
        important_frame.setLayout(important_layout)

        # Column A — Important domains
        domains_col = QWidget()
        domains_col_layout = QVBoxLayout()
        domains_col_layout.setContentsMargins(0, 0, 0, 0)
        domains_col_layout.setSpacing(6)
        domains_col.setLayout(domains_col_layout)
        domains_col.setMinimumWidth(180)
        domains_col.setMaximumWidth(260)

        domains_heading = QLabel("Important domains")
        domains_heading_font = QFont("Segoe UI", 9)
        domains_heading_font.setBold(True)
        domains_heading.setFont(domains_heading_font)
        domains_heading.setStyleSheet(f"color: {PALETTE['muted']};")
        domains_col_layout.addWidget(domains_heading)

        self.domain_entry = QLineEdit()
        self.domain_entry.setPlaceholderText("e.g. acme.com")
        self.domain_entry.returnPressed.connect(self._add_important_domain)
        domains_col_layout.addWidget(self.domain_entry)

        domain_btn_row = QHBoxLayout()
        domain_btn_row.setContentsMargins(0, 0, 0, 0)
        domain_btn_row.setSpacing(6)
        add_domain_btn = QPushButton("Add")
        add_domain_btn.clicked.connect(self._add_important_domain)
        domain_btn_row.addWidget(add_domain_btn)
        remove_domain_btn = QPushButton("Remove")
        remove_domain_btn.clicked.connect(self._remove_important_domain)
        domain_btn_row.addWidget(remove_domain_btn)
        domains_col_layout.addLayout(domain_btn_row)

        self.domains_listbox = QListWidget()
        self.domains_listbox.setSizePolicy(
            QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding
        )
        domains_col_layout.addWidget(self.domains_listbox, stretch=1)

        vip_heading = QLabel("VIP senders")
        vip_heading.setFont(domains_heading_font)
        vip_heading.setStyleSheet(f"color: {PALETTE['muted']};")
        vip_heading.setToolTip(
            "Individual addresses treated as VIPs (green timeline circles). "
            "Separate from Important domains."
        )
        domains_col_layout.addWidget(vip_heading)

        self.vip_entry = QLineEdit()
        self.vip_entry.setPlaceholderText("e.g. boss@acme.com")
        self.vip_entry.returnPressed.connect(self._add_vip_sender)
        domains_col_layout.addWidget(self.vip_entry)

        vip_btn_row = QHBoxLayout()
        vip_btn_row.setContentsMargins(0, 0, 0, 0)
        vip_btn_row.setSpacing(6)
        add_vip_btn = QPushButton("Add VIP")
        add_vip_btn.clicked.connect(self._add_vip_sender)
        vip_btn_row.addWidget(add_vip_btn)
        remove_vip_btn = QPushButton("Remove")
        remove_vip_btn.clicked.connect(self._remove_vip_sender)
        vip_btn_row.addWidget(remove_vip_btn)
        domains_col_layout.addLayout(vip_btn_row)

        self.vip_listbox = QListWidget()
        self.vip_listbox.setSizePolicy(
            QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding
        )
        domains_col_layout.addWidget(self.vip_listbox, stretch=1)
        important_layout.addWidget(domains_col, stretch=0)

        # Column B — Emails matching the domain filter + compact controls toolbar
        # (controls used to live in a 152px side column; inline row frees that
        # width for the table so more of each email is visible)
        emails_col = QWidget()
        emails_col_layout = QVBoxLayout()
        emails_col_layout.setContentsMargins(0, 0, 0, 0)
        emails_col_layout.setSpacing(6)
        emails_col.setLayout(emails_col_layout)

        emails_toolbar = QHBoxLayout()
        emails_toolbar.setContentsMargins(0, 0, 0, 0)
        emails_toolbar.setSpacing(6)

        emails_heading = QLabel("Matching emails")
        emails_heading_font = QFont("Segoe UI", 9)
        emails_heading_font.setBold(True)
        emails_heading.setFont(emails_heading_font)
        emails_heading.setStyleSheet(f"color: {PALETTE['muted']};")
        emails_toolbar.addWidget(emails_heading)
        emails_toolbar.addStretch(1)

        refresh_important_btn = QPushButton("Refresh")
        refresh_important_btn.clicked.connect(self._refresh_important_emails)
        emails_toolbar.addWidget(refresh_important_btn)

        scrape_btn = QPushButton("Scrape Outlook")
        scrape_btn.setToolTip(
            "Pull recent Inbox mail from AES-enabled accounts and detect actions."
        )
        scrape_btn.clicked.connect(lambda: self._start_outlook_scrape(manual=True))
        self.scrape_btn = scrape_btn
        emails_toolbar.addWidget(scrape_btn)

        self.unread_only_toggle = QPushButton("Unread only: OFF")
        self.unread_only_toggle.setCheckable(True)
        self.unread_only_toggle.setChecked(bool(self.scrape_unread_only))
        self.unread_only_toggle.setText(
            "Unread only: ON" if self.scrape_unread_only else "Unread only: OFF"
        )
        self.unread_only_toggle.toggled.connect(self._on_unread_only_toggled)
        emails_toolbar.addWidget(self.unread_only_toggle)

        go_back_lbl = QLabel("Go back")
        go_back_lbl.setStyleSheet(f"color: {PALETTE['muted']}; font-size: 11px;")
        emails_toolbar.addWidget(go_back_lbl)
        self.scrape_lookback_spin = QSpinBox()
        self.scrape_lookback_spin.setRange(1, 60)
        self.scrape_lookback_spin.setValue(
            max(1, int(getattr(self, "scrape_lookback_months", 0) or 1))
        )
        self.scrape_lookback_spin.setToolTip(
            "How many months of Inbox history Outlook scrape should cover."
        )
        self.scrape_lookback_spin.valueChanged.connect(self._on_lookback_months_edited)
        emails_toolbar.addWidget(self.scrape_lookback_spin)
        months_lbl = QLabel("Months")
        months_lbl.setStyleSheet(f"color: {PALETTE['muted']}; font-size: 11px;")
        emails_toolbar.addWidget(months_lbl)
        self.scrape_lookback_save_btn = QPushButton("Save")
        self.scrape_lookback_save_btn.setToolTip(
            "Lock in the lookback duration for Outlook scrapes."
        )
        self.scrape_lookback_save_btn.clicked.connect(self._save_scrape_lookback_months)
        emails_toolbar.addWidget(self.scrape_lookback_save_btn)
        self._update_lookback_save_button_style()

        emails_col_layout.addLayout(emails_toolbar)

        columns = ("Sender", "Recipients", "Subject", "DateTime", "Status")
        self.important_tree = QTreeWidget()
        self.important_tree.setAlternatingRowColors(True)
        self.important_tree.setRootIsDecorated(False)
        self.important_tree.setHeaderLabels(columns)
        self.important_tree.setColumnCount(len(columns))
        self.important_tree.setColumnWidth(0, 150)
        self.important_tree.setColumnWidth(1, 110)
        self.important_tree.setColumnWidth(3, 118)
        self.important_tree.setColumnWidth(4, 80)
        important_hdr = self.important_tree.header()
        important_hdr.setStretchLastSection(False)
        important_hdr.setSectionResizeMode(2, QHeaderView.ResizeMode.Stretch)
        self.important_tree.setSizePolicy(
            QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding
        )
        self.important_tree.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.important_tree.customContextMenuRequested.connect(self._show_email_context_menu)
        self.important_tree.itemClicked.connect(self._on_mail_item_selected)
        emails_col_layout.addWidget(self.important_tree, stretch=1)
        important_layout.addWidget(emails_col, stretch=1)

        # Actions needing attention (live scrape)
        actions_frame = QGroupBox("Actions to take")
        actions_layout = QVBoxLayout()
        actions_frame.setLayout(actions_layout)
        self.actions_hint = QLabel("Live Outlook scrape — reply / RSVP / approve / deadline signals.")
        self.actions_hint.setStyleSheet("color: #4a6570; font-size: 11px;")
        self.actions_hint.setWordWrap(True)
        actions_layout.addWidget(self.actions_hint)

        self.actions_tree = QTreeWidget()
        self.actions_tree.setAlternatingRowColors(True)
        self.actions_tree.setRootIsDecorated(False)
        self.actions_tree.setHeaderLabels(
            ("From", "Subject", "Action", "Signal", "When", "Status")
        )
        self.actions_tree.setColumnCount(6)
        self.actions_tree.setColumnWidth(0, 160)
        self.actions_tree.setColumnWidth(2, 96)
        self.actions_tree.setColumnWidth(3, 70)
        self.actions_tree.setColumnWidth(4, 118)
        self.actions_tree.setColumnWidth(5, 80)
        actions_hdr = self.actions_tree.header()
        actions_hdr.setStretchLastSection(False)
        actions_hdr.setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        self.actions_tree.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.actions_tree.customContextMenuRequested.connect(self._show_action_context_menu)
        self.actions_tree.itemClicked.connect(self._on_mail_item_selected)
        self.actions_tree.itemDoubleClicked.connect(self._open_action_item_in_outlook)
        actions_layout.addWidget(self.actions_tree)

        actions_btns = QHBoxLayout()
        refresh_actions_btn = QPushButton("Refresh Actions")
        refresh_actions_btn.clicked.connect(self._refresh_actions_panel)
        open_outlook_btn = QPushButton("Open in Outlook")
        open_outlook_btn.clicked.connect(self._open_selected_action_in_outlook)
        actions_btns.addWidget(refresh_actions_btn)
        actions_btns.addWidget(open_outlook_btn)
        actions_btns.addStretch(1)
        actions_layout.addLayout(actions_btns)

        # Side-by-side columns: Important Emails | Actions — both tables get the
        # full height of the lower pane, so far more rows are visible.
        mail_split = QSplitter(Qt.Orientation.Horizontal)
        mail_split.setChildrenCollapsible(False)
        mail_split.setHandleWidth(6)
        mail_split.addWidget(important_frame)
        mail_split.addWidget(actions_frame)
        mail_split.setStretchFactor(0, 11)
        mail_split.setStretchFactor(1, 9)
        mail_split.setSizes([880, 720])
        self.welcome_mail_split = mail_split
        left_layout.addWidget(mail_split, stretch=1)
        
        main_split.addWidget(left_widget)
        
        # RIGHT SIDE - Statistics
        right_widget = QWidget()
        right_layout = QVBoxLayout()
        right_widget.setLayout(right_layout)
        
        # Statistics Section - RIGHT SIDE
        stats_frame = QGroupBox("Email Statistics")
        stats_layout = QVBoxLayout()
        stats_frame.setLayout(stats_layout)
        
        # Last refresh time display
        refresh_info_frame = QWidget()
        refresh_info_layout = QHBoxLayout()
        refresh_info_frame.setLayout(refresh_info_layout)
        
        refresh_label = QLabel("Last refresh:")
        refresh_label.setStyleSheet(f"color: {PALETTE['muted']};")
        refresh_info_layout.addWidget(refresh_label)
        
        self.last_refresh_label = QLabel("Never")
        font = QFont('Segoe UI', 9)
        font.setItalic(True)
        self.last_refresh_label.setFont(font)
        self.last_refresh_label.setStyleSheet(f"color: {PALETTE['muted']};")
        refresh_info_layout.addWidget(self.last_refresh_label)
        
        refresh_info_layout.addStretch()
        
        auto_refresh_label = QLabel("Auto-refresh every 60 s")
        auto_refresh_label.setStyleSheet(f"color: {PALETTE['muted']};")
        refresh_info_layout.addWidget(auto_refresh_label)
        
        stats_layout.addWidget(refresh_info_frame)
        
        # Create stats display
        stats_container = QWidget()
        stats_container_layout = QHBoxLayout()
        stats_container.setLayout(stats_container_layout)
        
        # Today's stats (left column)
        today_frame = QWidget()
        today_layout = QVBoxLayout()
        today_frame.setLayout(today_layout)
        
        today_label = QLabel("Today")
        font = QFont('Segoe UI', 12)
        font.setBold(True)
        today_label.setFont(font)
        today_label.setStyleSheet(f"color: {PALETTE['accent']};")
        today_layout.addWidget(today_label)
        
        self.today_sent_label = QLabel("Sent: 0")
        self.today_sent_label.setFont(QFont('Segoe UI', 11))
        today_layout.addWidget(self.today_sent_label)
        
        self.today_received_label = QLabel("Received: 0")
        self.today_received_label.setFont(QFont('Segoe UI', 11))
        today_layout.addWidget(self.today_received_label)
        
        today_layout.addStretch()
        stats_container_layout.addWidget(today_frame)
        
        # Yesterday's stats (right column)
        yesterday_frame = QWidget()
        yesterday_layout = QVBoxLayout()
        yesterday_frame.setLayout(yesterday_layout)
        
        yesterday_label = QLabel("Yesterday")
        font = QFont('Segoe UI', 12)
        font.setBold(True)
        yesterday_label.setFont(font)
        yesterday_label.setStyleSheet(f"color: {PALETTE['accent']};")
        yesterday_layout.addWidget(yesterday_label)
        
        self.yesterday_sent_label = QLabel("Sent: 0")
        self.yesterday_sent_label.setFont(QFont('Segoe UI', 11))
        yesterday_layout.addWidget(self.yesterday_sent_label)
        
        self.yesterday_received_label = QLabel("Received: 0")
        self.yesterday_received_label.setFont(QFont('Segoe UI', 11))
        yesterday_layout.addWidget(self.yesterday_received_label)
        
        yesterday_layout.addStretch()
        stats_container_layout.addWidget(yesterday_frame)
        
        stats_layout.addWidget(stats_container)
        
        # Refresh button
        refresh_stats_btn = QPushButton("Refresh Statistics")
        refresh_stats_btn.clicked.connect(self._refresh_welcome_stats)
        stats_layout.addWidget(refresh_stats_btn)
        
        right_layout.addWidget(stats_frame)
        self.welcome_stats_frame = stats_frame

        # Email preview + action highlights (fills empty right-hand space)
        preview_frame = QGroupBox("Email preview")
        preview_layout = QVBoxLayout()
        preview_frame.setLayout(preview_layout)
        self.preview_frame = preview_frame
        self._preview_maximized = False
        self._preview_split_sizes: Optional[List[int]] = None
        self._preview_popout: Optional[QDialog] = None

        # Top-right chrome: Maximize / Minimize + Pop out
        chrome_row = QHBoxLayout()
        chrome_row.setContentsMargins(0, 0, 0, 2)
        chrome_row.addStretch(1)
        self.preview_max_btn = QToolButton()
        self.preview_max_btn.setObjectName("paneChrome")
        self.preview_max_btn.setText("Maximize")
        self.preview_max_btn.setToolTip("Expand email preview over the Welcome page")
        self.preview_max_btn.clicked.connect(self._toggle_preview_maximize)
        chrome_row.addWidget(self.preview_max_btn)
        self.preview_pop_btn = QToolButton()
        self.preview_pop_btn.setObjectName("paneChrome")
        self.preview_pop_btn.setText("Pop out")
        self.preview_pop_btn.setToolTip("Open email preview in its own window")
        self.preview_pop_btn.clicked.connect(self._popout_email_preview)
        chrome_row.addWidget(self.preview_pop_btn)
        preview_layout.addLayout(chrome_row)

        self.preview_meta = QLabel("Select an email from Important or Actions.")
        self.preview_meta.setWordWrap(True)
        self.preview_meta.setTextFormat(Qt.TextFormat.RichText)
        self.preview_meta.setStyleSheet("color: #4a6570; font-size: 11px;")
        preview_layout.addWidget(self.preview_meta)

        self.preview_reasons = QLabel("")
        self.preview_reasons.setWordWrap(True)
        self.preview_reasons.setStyleSheet("color: #0f6b7c; font-size: 11px;")
        preview_layout.addWidget(self.preview_reasons)

        self.email_preview = QTextEdit()
        self.email_preview.setReadOnly(True)
        self.email_preview.setPlaceholderText(
            "Message body appears here. Action phrases are highlighted."
        )
        preview_layout.addWidget(self.email_preview, stretch=1)

        preview_btns = QHBoxLayout()
        preview_btns.setSpacing(6)
        preview_open_btn = QPushButton("Open in Outlook")
        preview_open_btn.clicked.connect(self._open_preview_in_outlook)
        preview_btns.addWidget(preview_open_btn)

        ignore_email_btn = QPushButton("Ignore Email")
        ignore_email_btn.setToolTip(
            "Remove this message from prioritisation (Important, Actions, Timeline)"
        )
        ignore_email_btn.clicked.connect(self._preview_ignore_email)
        preview_btns.addWidget(ignore_email_btn)

        ignore_recipient_btn = QPushButton("Ignore Recipient")
        ignore_recipient_btn.setToolTip(
            "Remove this counterparty (From:) from prioritisation entirely"
        )
        ignore_recipient_btn.clicked.connect(self._preview_ignore_recipient)
        preview_btns.addWidget(ignore_recipient_btn)

        deprecate_email_btn = QPushButton("Deprecate Email")
        deprecate_email_btn.setToolTip(
            "Keep visible but score this email lower in prioritisation"
        )
        deprecate_email_btn.clicked.connect(self._preview_deprecate_email)
        preview_btns.addWidget(deprecate_email_btn)

        deprecate_sender_btn = QPushButton("Deprecate Sender")
        deprecate_sender_btn.setToolTip(
            "Keep visible but score this sender lower in prioritisation"
        )
        deprecate_sender_btn.clicked.connect(self._preview_deprecate_sender)
        preview_btns.addWidget(deprecate_sender_btn)

        self.preview_vip_btn = QPushButton("Mark VIP")
        self.preview_vip_btn.setToolTip(
            "Treat this sender as a VIP (green timeline importance). "
            "Click again to remove."
        )
        self.preview_vip_btn.clicked.connect(self._preview_toggle_vip_sender)
        preview_btns.addWidget(self.preview_vip_btn)
        preview_btns.addStretch(1)
        preview_layout.addLayout(preview_btns)

        score_row = QHBoxLayout()
        score_row.setSpacing(10)

        sender_score_wrap = QHBoxLayout()
        sender_score_wrap.setSpacing(4)
        sender_score_wrap.addWidget(QLabel("Sender Relevance"))
        self.preview_sender_score = QLabel("—")
        self.preview_sender_score.setMinimumWidth(28)
        self.preview_sender_score.setStyleSheet("font-weight: 700; color: #0f6b7c;")
        sender_score_wrap.addWidget(self.preview_sender_score)
        sender_tick = QPushButton("✓")
        sender_tick.setObjectName("scoreTick")
        sender_tick.setFixedWidth(28)
        sender_tick.setToolTip("Endorse Sender Ranking (gradually increase)")
        sender_tick.clicked.connect(self._preview_endorse_sender_relevance)
        sender_score_wrap.addWidget(sender_tick)
        sender_cross = QPushButton("✗")
        sender_cross.setObjectName("scoreCross")
        sender_cross.setFixedWidth(28)
        sender_cross.setToolTip("De-rate or mark spam")
        sender_cross.clicked.connect(self._preview_cross_relevance)
        sender_score_wrap.addWidget(sender_cross)
        score_row.addLayout(sender_score_wrap)

        email_score_wrap = QHBoxLayout()
        email_score_wrap.setSpacing(4)
        email_score_wrap.addWidget(QLabel("Email Relevance"))
        self.preview_email_score = QLabel("—")
        self.preview_email_score.setMinimumWidth(28)
        self.preview_email_score.setStyleSheet("font-weight: 700; color: #0f6b7c;")
        email_score_wrap.addWidget(self.preview_email_score)
        email_tick = QPushButton("✓")
        email_tick.setObjectName("scoreTick")
        email_tick.setFixedWidth(28)
        email_tick.setToolTip("Endorse Email Ranking (gradually increase)")
        email_tick.clicked.connect(self._preview_endorse_email_relevance)
        email_score_wrap.addWidget(email_tick)
        email_cross = QPushButton("✗")
        email_cross.setObjectName("scoreCross")
        email_cross.setFixedWidth(28)
        email_cross.setToolTip("De-rate or mark spam")
        email_cross.clicked.connect(self._preview_cross_relevance)
        email_score_wrap.addWidget(email_cross)
        score_row.addLayout(email_score_wrap)
        score_row.addStretch(1)

        describe_rule_btn = QPushButton("Describe a rule…")
        describe_rule_btn.setObjectName("describeRule")
        describe_rule_btn.setToolTip(
            "Write a natural-language rule for this email / sender / always. "
            "Saved rules are passed to Ollama so it can apply them."
        )
        describe_rule_btn.clicked.connect(self._show_describe_rule_dialog)
        score_row.addWidget(describe_rule_btn)
        preview_layout.addLayout(score_row)

        right_layout.addWidget(preview_frame, stretch=1)
        self._preview_det: Optional[Dict[str, Any]] = None
        
        main_split.addWidget(right_widget)
        main_split.setStretchFactor(0, 3)
        main_split.setStretchFactor(1, 2)
        
        content_layout.addWidget(main_split)
        
        # Set welcome_frame layout
        welcome_layout = QVBoxLayout()
        welcome_frame.setLayout(welcome_layout)
        welcome_layout.addWidget(content)
        
        # Load domain list UI (data refresh runs after database connects)
        self._refresh_domains_list()
        
        # Start auto-refresh (every 60 seconds)
        self._start_auto_refresh()
    
    def _create_view_tab(self):
        """Create the view records tab."""
        view_frame = QWidget()
        self.notebook.addTab(view_frame, "View Records")
        
        layout = QVBoxLayout()
        view_frame.setLayout(layout)
        
        # Toolbar
        toolbar = QWidget()
        toolbar_layout = QHBoxLayout()
        toolbar.setLayout(toolbar_layout)
        
        refresh_btn = QPushButton("Refresh")
        refresh_btn.clicked.connect(self._load_records)
        toolbar_layout.addWidget(refresh_btn)
        
        prev_btn = QPushButton("← Previous")
        prev_btn.clicked.connect(self._previous_page)
        toolbar_layout.addWidget(prev_btn)
        
        next_btn = QPushButton("Next →")
        next_btn.clicked.connect(self._next_page)
        toolbar_layout.addWidget(next_btn)
        
        self.page_label = QLabel("Page: 0")
        toolbar_layout.addWidget(self.page_label)
        
        self.record_count_label = QLabel("Total Records: 0")
        toolbar_layout.addWidget(self.record_count_label)
        
        toolbar_layout.addStretch()
        
        layout.addWidget(toolbar)
        
        # Records tree view
        columns = ("ID", "GURI", "Sender", "Recipients", "Subject", "DateTime", "Risk", "Doc Type", "Created")
        self.tree = QTreeWidget()
        self.tree.setAlternatingRowColors(True)
        self.tree.setRootIsDecorated(False)
        self.tree.setHeaderLabels(columns)
        self.tree.setColumnCount(len(columns))
        
        # Track sort state
        self.tree_sort_column = None
        self.tree_sort_reverse = False
        
        # Set column widths
        self.tree.setColumnWidth(0, 50)   # ID
        self.tree.setColumnWidth(1, 200)  # GURI
        self.tree.setColumnWidth(2, 180)  # Sender
        self.tree.setColumnWidth(3, 120)  # Recipients
        self.tree.setColumnWidth(4, 200)  # Subject
        self.tree.setColumnWidth(5, 130)  # DateTime
        self.tree.setColumnWidth(6, 100)  # Risk
        self.tree.setColumnWidth(7, 100)  # Doc Type
        self.tree.setColumnWidth(8, 150)  # Created
        
        # Enable sorting
        self.tree.setSortingEnabled(True)
        self.tree.header().sectionClicked.connect(self._on_tree_header_clicked)
        
        # Bind double-click to view details
        self.tree.itemDoubleClicked.connect(self._view_record_details)
        
        layout.addWidget(self.tree)

    def _create_database_tab(self):
        """Browser for the active database: engine, schemas, tables, sample records."""
        frame = QWidget()
        self.notebook.addTab(frame, "Database")
        root = QVBoxLayout(frame)
        root.setContentsMargins(12, 12, 12, 12)
        root.setSpacing(10)

        head = QLabel("Database")
        head.setStyleSheet(
            f"font-size: 14px; font-weight: 700; color: {PALETTE['accent']};"
        )
        root.addWidget(head)

        self.db_browser_hint = QLabel(
            "PostgreSQL is the supported engine. MySQL and SQLite are deprecated."
        )
        self.db_browser_hint.setWordWrap(True)
        self.db_browser_hint.setStyleSheet(f"color: {PALETTE['muted']};")
        root.addWidget(self.db_browser_hint)

        self.db_deprecated_banner = QLabel("")
        self.db_deprecated_banner.setWordWrap(True)
        self.db_deprecated_banner.setStyleSheet(
            "color: #8a4b08; background: #fff6e5; padding: 8px 10px; border-radius: 4px;"
        )
        self.db_deprecated_banner.setVisible(False)
        root.addWidget(self.db_deprecated_banner)

        toolbar = QHBoxLayout()
        refresh_btn = QPushButton("Refresh")
        refresh_btn.clicked.connect(self._refresh_database_browser)
        pg_btn = QPushButton("Connect PostgreSQL…")
        pg_btn.clicked.connect(self._show_postgres_connection_dialog)
        migrate_btn = QPushButton("Migrate…")
        migrate_btn.setToolTip(
            "Copy tables and data from MySQL or SQLite into PostgreSQL"
        )
        migrate_btn.clicked.connect(self._show_migrate_dialog)
        toolbar.addWidget(refresh_btn)
        toolbar.addWidget(pg_btn)
        toolbar.addWidget(migrate_btn)
        toolbar.addStretch(1)
        root.addLayout(toolbar)

        self.db_info_label = QLabel("Not connected")
        self.db_info_label.setWordWrap(True)
        self.db_info_label.setTextFormat(Qt.TextFormat.RichText)
        self.db_info_label.setStyleSheet(
            "background: #ffffff; border: 1px solid #e2e8ee; padding: 10px; border-radius: 4px;"
        )
        root.addWidget(self.db_info_label)

        split = QSplitter(Qt.Orientation.Horizontal)

        left = QWidget()
        left_layout = QVBoxLayout(left)
        left_layout.setContentsMargins(0, 0, 0, 0)
        left_layout.setSpacing(6)

        left_layout.addWidget(QLabel("Schemas"))
        self.db_schemas_list = QListWidget()
        self.db_schemas_list.setMaximumHeight(120)
        self.db_schemas_list.currentTextChanged.connect(self._on_db_schema_selected)
        left_layout.addWidget(self.db_schemas_list)

        left_layout.addWidget(QLabel("Tables"))
        self.db_tables_tree = QTreeWidget()
        self.db_tables_tree.setAlternatingRowColors(True)
        self.db_tables_tree.setRootIsDecorated(False)
        self.db_tables_tree.setUniformRowHeights(True)
        self.db_tables_tree.setHeaderLabels(("Table", "Approx. rows"))
        self.db_tables_tree.setColumnWidth(0, 180)
        self.db_tables_tree.itemClicked.connect(self._on_db_table_selected)
        left_layout.addWidget(self.db_tables_tree, stretch=1)
        split.addWidget(left)

        right = QWidget()
        right_layout = QVBoxLayout(right)
        right_layout.setContentsMargins(0, 0, 0, 0)
        right_layout.setSpacing(8)

        self.db_table_title = QLabel("Select a table")
        title_font = QFont("Segoe UI", 12)
        title_font.setBold(True)
        self.db_table_title.setFont(title_font)
        self.db_table_title.setStyleSheet(f"color: {PALETTE['accent']};")
        right_layout.addWidget(self.db_table_title)

        cols_group = QGroupBox("Columns")
        cols_layout = QVBoxLayout(cols_group)
        self.db_columns_tree = QTreeWidget()
        self.db_columns_tree.setAlternatingRowColors(True)
        self.db_columns_tree.setRootIsDecorated(False)
        self.db_columns_tree.setHeaderLabels(("Column", "Type", "Nullable", "Default"))
        self.db_columns_tree.setColumnWidth(0, 140)
        self.db_columns_tree.setColumnWidth(1, 120)
        self.db_columns_tree.setMaximumHeight(160)
        cols_layout.addWidget(self.db_columns_tree)
        right_layout.addWidget(cols_group)

        rows_group = QGroupBox("Sample records (first 50)")
        rows_layout = QVBoxLayout(rows_group)
        self.db_rows_tree = QTreeWidget()
        self.db_rows_tree.setAlternatingRowColors(True)
        self.db_rows_tree.setRootIsDecorated(False)
        self.db_rows_tree.setUniformRowHeights(True)
        rows_layout.addWidget(self.db_rows_tree)
        right_layout.addWidget(rows_group, stretch=1)

        split.addWidget(right)
        split.setStretchFactor(0, 2)
        split.setStretchFactor(1, 5)
        split.setSizes([280, 720])
        root.addWidget(split, stretch=1)

        self._db_browser_tables: List[Dict[str, Any]] = []
        QTimer.singleShot(600, self._refresh_database_browser)

    def _focus_database_tab(self) -> None:
        if not hasattr(self, "notebook"):
            return
        for i in range(self.notebook.count()):
            if self.notebook.tabText(i) == "Database":
                self.notebook.setCurrentIndex(i)
                self._refresh_database_browser()
                return

    def _refresh_database_browser(self) -> None:
        if not hasattr(self, "db_info_label"):
            return
        if not self.db:
            self.db_info_label.setText("<b>Not connected</b> — connect to PostgreSQL.")
            self.db_deprecated_banner.setVisible(False)
            if hasattr(self, "db_schemas_list"):
                self.db_schemas_list.clear()
            if hasattr(self, "db_tables_tree"):
                self.db_tables_tree.clear()
            return
        try:
            info = self.db.get_connection_info()
            guri_count = 0
            try:
                guri_count = int(self.db.get_guri_record_count())
            except Exception:
                guri_count = 0

            if info.get("deprecated"):
                self.db_deprecated_banner.setText(
                    f"⚠ Connected to {info.get('label')} — MySQL and SQLite are "
                    "deprecated. Configure PostgreSQL via guri_postgres_config.json "
                    "or Connect PostgreSQL…"
                )
                self.db_deprecated_banner.setVisible(True)
            else:
                self.db_deprecated_banner.setVisible(False)

            bits = [
                f"<b>Engine:</b> {info.get('label') or self.db_type}",
                f"<b>Database:</b> {info.get('database') or '—'}",
            ]
            if info.get("host"):
                bits.append(
                    f"<b>Host:</b> {info.get('host')}:{info.get('port') or ''}"
                )
            if info.get("user"):
                bits.append(f"<b>User:</b> {info.get('user')}")
            if info.get("path"):
                bits.append(f"<b>Path:</b> {info.get('path')}")
            bits.append(f"<b>guri_records:</b> {guri_count:,} row(s)")
            if info.get("server_version"):
                ver = str(info.get("server_version") or "").split(",")[0]
                bits.append(f"<b>Server:</b> {ver}")
            self.db_info_label.setText(" &nbsp;|&nbsp; ".join(bits))

            schemas = self.db.list_schemas()
            self.db_schemas_list.blockSignals(True)
            self.db_schemas_list.clear()
            self.db_schemas_list.addItem("(all schemas)")
            for schema in schemas:
                self.db_schemas_list.addItem(schema)
            self.db_schemas_list.setCurrentRow(0)
            self.db_schemas_list.blockSignals(False)

            self._load_db_tables_for_schema(None)
            self.db_browser_hint.setText(
                f"Browsing {info.get('label')} — select a schema and table to inspect columns and sample rows."
            )
            if hasattr(self, "status_bar"):
                self.status_bar.showMessage(
                    f"Database browser: {info.get('label')} / {info.get('database')}"
                )
        except Exception as exc:
            self.logger.error("Database browser refresh failed: %s", exc)
            self.db_info_label.setText(f"<b>Error:</b> {exc}")

    def _load_db_tables_for_schema(self, schema: Optional[str]) -> None:
        if not self.db or not hasattr(self, "db_tables_tree"):
            return
        self.db_tables_tree.clear()
        self.db_columns_tree.clear()
        self.db_rows_tree.clear()
        self.db_table_title.setText("Select a table")
        try:
            tables = self.db.list_tables(schema=schema)
        except Exception as exc:
            self.logger.error("list_tables failed: %s", exc)
            tables = []
        self._db_browser_tables = tables
        for tbl in tables:
            item = QTreeWidgetItem(self.db_tables_tree)
            name = str(tbl.get("full_name") or tbl.get("name") or "")
            item.setText(0, name)
            item.setText(1, f"{int(tbl.get('rows') or 0):,}")
            item.setData(0, Qt.ItemDataRole.UserRole, tbl)

    def _on_db_schema_selected(self, text: str) -> None:
        schema = None if not text or text.startswith("(all") else text
        self._load_db_tables_for_schema(schema)

    def _on_db_table_selected(self, item, _column=None) -> None:
        if item is None or not self.db:
            return
        tbl = item.data(0, Qt.ItemDataRole.UserRole) or {}
        if not isinstance(tbl, dict):
            return
        schema = str(tbl.get("schema") or "") or None
        name = str(tbl.get("name") or "")
        full = str(tbl.get("full_name") or name)
        self.db_table_title.setText(
            f"{full}  ·  ~{int(tbl.get('rows') or 0):,} rows"
        )

        self.db_columns_tree.clear()
        try:
            columns = self.db.describe_table(name, schema=schema)
        except Exception as exc:
            self.logger.error("describe_table failed: %s", exc)
            columns = []
        for col in columns:
            row = QTreeWidgetItem(self.db_columns_tree)
            row.setText(0, str(col.get("name") or ""))
            row.setText(1, str(col.get("type") or ""))
            row.setText(2, "YES" if col.get("nullable") else "NO")
            row.setText(3, str(col.get("default") or "")[:80])

        self.db_rows_tree.clear()
        try:
            preview = self.db.preview_table(name, schema=schema, limit=50)
        except Exception as exc:
            preview = {"columns": [], "rows": [], "error": str(exc)}
        if preview.get("error"):
            self.db_rows_tree.setHeaderLabels(("Error",))
            err = QTreeWidgetItem(self.db_rows_tree)
            err.setText(0, str(preview.get("error")))
            return
        colnames = list(preview.get("columns") or [])
        if not colnames:
            self.db_rows_tree.setHeaderLabels(("(no columns)",))
            return
        # Cap visible columns for very wide tables
        show_cols = colnames[:12]
        self.db_rows_tree.setHeaderLabels(tuple(show_cols))
        for col_i in range(len(show_cols)):
            self.db_rows_tree.setColumnWidth(col_i, 110)
        for data_row in preview.get("rows") or []:
            item_row = QTreeWidgetItem(self.db_rows_tree)
            for i, col_name in enumerate(show_cols):
                val = data_row.get(col_name)
                text = "" if val is None else str(val)
                if len(text) > 80:
                    text = text[:79] + "…"
                item_row.setText(i, text)

    def _show_migrate_dialog(self) -> None:
        """Dialog to copy MySQL/SQLite tables into PostgreSQL."""
        if getattr(self, "_migrate_busy", False):
            QMessageBox.information(
                self, "Migrate", "A migration is already running."
            )
            return

        dialog = QDialog(self)
        dialog.setWindowTitle("Migrate to PostgreSQL")
        dialog.setModal(True)
        dialog.resize(560, 520)
        layout = QVBoxLayout(dialog)

        intro = QLabel(
            "Copy GURI tables and content from a deprecated engine "
            "(MySQL or SQLite) into PostgreSQL. Existing PostgreSQL rows with "
            "the same unique keys are kept (ON CONFLICT DO NOTHING)."
        )
        intro.setWordWrap(True)
        intro.setStyleSheet(f"color: {PALETTE['muted']};")
        layout.addWidget(intro)

        source_box = QGroupBox("Source")
        source_layout = QVBoxLayout(source_box)
        source_group = QButtonGroup(dialog)
        src_mysql = QRadioButton("MySQL (guri_mysql_config.json)")
        src_sqlite = QRadioButton("SQLite file…")
        source_group.addButton(src_mysql)
        source_group.addButton(src_sqlite)
        src_mysql.setChecked(True)
        source_layout.addWidget(src_mysql)
        source_layout.addWidget(src_sqlite)

        sqlite_row = QHBoxLayout()
        sqlite_path_edit = QLineEdit(r"C:\GeoFooter\guri_records.db")
        sqlite_browse = QPushButton("Browse…")
        sqlite_row.addWidget(sqlite_path_edit, stretch=1)
        sqlite_row.addWidget(sqlite_browse)
        source_layout.addLayout(sqlite_row)
        layout.addWidget(source_box)

        # Prefill MySQL summary if config exists
        mysql_cfg_path = os.path.join(r"C:\GeoFooter", "guri_mysql_config.json")
        mysql_summary = QLabel("")
        mysql_summary.setWordWrap(True)
        mysql_summary.setStyleSheet(f"color: {PALETTE['muted']}; font-size: 11px;")
        mysql_defaults: Dict[str, Any] = {}
        if os.path.isfile(mysql_cfg_path):
            try:
                with open(mysql_cfg_path, "r", encoding="utf-8") as fh:
                    mysql_defaults = json.load(fh)
                mysql_summary.setText(
                    f"MySQL source: {mysql_defaults.get('host')}:{mysql_defaults.get('port', 3306)}/"
                    f"{mysql_defaults.get('database')} as {mysql_defaults.get('user')}"
                )
            except Exception as exc:
                mysql_summary.setText(f"Could not read MySQL config: {exc}")
                src_sqlite.setChecked(True)
        else:
            mysql_summary.setText(
                "No guri_mysql_config.json found — use SQLite or create that config file."
            )
            src_sqlite.setChecked(True)
        source_layout.addWidget(mysql_summary)

        dest_box = QGroupBox("Destination PostgreSQL")
        dest_form = QFormLayout(dest_box)
        pg_defaults = {
            "host": "localhost",
            "port": "5432",
            "user": "postgres",
            "password": "",
            "database": "guri_db",
        }
        pg_cfg_path = os.path.join(r"C:\GeoFooter", "guri_postgres_config.json")
        if os.path.isfile(pg_cfg_path):
            try:
                with open(pg_cfg_path, "r", encoding="utf-8") as fh:
                    saved = json.load(fh)
                for key in pg_defaults:
                    if key in saved and saved[key] is not None:
                        pg_defaults[key] = str(saved[key])
            except Exception:
                pass
        elif mysql_defaults:
            pg_defaults["host"] = str(mysql_defaults.get("host") or "localhost")
            pg_defaults["user"] = str(mysql_defaults.get("user") or "postgres")
            pg_defaults["password"] = str(mysql_defaults.get("password") or "")
            pg_defaults["database"] = str(mysql_defaults.get("database") or "guri_db")

        pg_host = QLineEdit(pg_defaults["host"])
        pg_port = QLineEdit(pg_defaults["port"])
        pg_user = QLineEdit(pg_defaults["user"])
        pg_password = QLineEdit(pg_defaults["password"])
        pg_password.setEchoMode(QLineEdit.EchoMode.Password)
        pg_database = QLineEdit(pg_defaults["database"])
        dest_form.addRow("Host", pg_host)
        dest_form.addRow("Port", pg_port)
        dest_form.addRow("User", pg_user)
        dest_form.addRow("Password", pg_password)
        dest_form.addRow("Database", pg_database)
        layout.addWidget(dest_box)

        log_box = QGroupBox("Progress")
        log_layout = QVBoxLayout(log_box)
        migrate_log = QTextEdit()
        migrate_log.setReadOnly(True)
        migrate_log.setMinimumHeight(140)
        migrate_log.setPlaceholderText("Migration output appears here…")
        log_layout.addWidget(migrate_log)
        layout.addWidget(log_box, stretch=1)

        run_btn = QPushButton("Start migration")
        run_btn.setDefault(True)
        cancel_btn = QPushButton("Cancel")
        cancel_btn.clicked.connect(dialog.reject)
        add_dialog_button_row(layout, affirmative=run_btn, cancel=cancel_btn)

        def _set_sqlite_enabled() -> None:
            use_sqlite = src_sqlite.isChecked()
            sqlite_path_edit.setEnabled(use_sqlite)
            sqlite_browse.setEnabled(use_sqlite)

        def _browse_sqlite() -> None:
            filename, _ = QFileDialog.getOpenFileName(
                dialog,
                "Select SQLite database",
                r"C:\GeoFooter",
                "SQLite Database (*.db);;All Files (*.*)",
            )
            if filename:
                sqlite_path_edit.setText(filename)

        src_mysql.toggled.connect(lambda _c: _set_sqlite_enabled())
        src_sqlite.toggled.connect(lambda _c: _set_sqlite_enabled())
        sqlite_browse.clicked.connect(_browse_sqlite)
        _set_sqlite_enabled()

        self._migrate_dialog = dialog
        self._migrate_log_widget = migrate_log
        self._migrate_run_btn = run_btn

        def on_run() -> None:
            if getattr(self, "_migrate_busy", False):
                return
            try:
                port = int(pg_port.text().strip() or "5432")
            except ValueError:
                QMessageBox.warning(dialog, "Migrate", "Port must be a number.")
                return
            pg_cfg = {
                "host": pg_host.text().strip() or "localhost",
                "port": port,
                "user": pg_user.text().strip(),
                "password": pg_password.text(),
                "database": pg_database.text().strip(),
            }
            for key in ("user", "database"):
                if not pg_cfg[key]:
                    QMessageBox.warning(
                        dialog, "Migrate", f"PostgreSQL {key} is required."
                    )
                    return

            mysql_cfg = None
            sqlite_path = None
            if src_sqlite.isChecked():
                path = sqlite_path_edit.text().strip()
                if not path or not os.path.isfile(path):
                    QMessageBox.warning(
                        dialog, "Migrate", "Choose a valid SQLite .db file."
                    )
                    return
                sqlite_path = path
            else:
                if not os.path.isfile(mysql_cfg_path):
                    QMessageBox.warning(
                        dialog,
                        "Migrate",
                        f"Missing MySQL config:\n{mysql_cfg_path}",
                    )
                    return
                try:
                    with open(mysql_cfg_path, "r", encoding="utf-8") as fh:
                        mysql_cfg = json.load(fh)
                except Exception as exc:
                    QMessageBox.critical(
                        dialog, "Migrate", f"Could not read MySQL config:\n{exc}"
                    )
                    return

            confirm = QMessageBox.question(
                dialog,
                "Confirm migration",
                "Copy all GURI tables from the selected source into PostgreSQL?\n\n"
                f"Destination: {pg_cfg['host']}:{pg_cfg['port']}/{pg_cfg['database']}",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.Yes,
            )
            if confirm != QMessageBox.StandardButton.Yes:
                return

            # Persist destination so GURI reconnects cleanly after migrate
            try:
                from migrate_mysql_to_postgres import save_postgres_config

                save_postgres_config(pg_cfg)
            except Exception as exc:
                self.logger.warning("Could not save postgres config: %s", exc)

            migrate_log.clear()
            migrate_log.append("Starting migration…")
            run_btn.setEnabled(False)
            self._migrate_busy = True
            self.status_bar.showMessage("Migrating database to PostgreSQL…")

            def work() -> None:
                try:
                    from pathlib import Path as _Path

                    from migrate_mysql_to_postgres import migrate as run_migrate

                    def _progress(msg: str) -> None:
                        self._migrate_bridge.progress.emit(str(msg))

                    report = run_migrate(
                        mysql_cfg=mysql_cfg,
                        sqlite_path=_Path(sqlite_path) if sqlite_path else None,
                        pg_cfg=pg_cfg,
                        log=_progress,
                    )
                    self._migrate_bridge.finished.emit(report, "")
                except Exception as exc:
                    self._migrate_bridge.finished.emit(None, str(exc))

            threading.Thread(target=work, daemon=True).start()

        run_btn.clicked.connect(on_run)
        dialog.exec()
        if getattr(self, "_migrate_dialog", None) is dialog:
            self._migrate_dialog = None
            self._migrate_log_widget = None
            self._migrate_run_btn = None

    def _on_migrate_progress(self, message: str) -> None:
        log = getattr(self, "_migrate_log_widget", None)
        if log is not None:
            log.append(message)
            log.ensureCursorVisible()

    def _on_migrate_finished(self, report_obj, error: str) -> None:
        self._migrate_busy = False
        run_btn = getattr(self, "_migrate_run_btn", None)
        if run_btn is not None:
            run_btn.setEnabled(True)
        log = getattr(self, "_migrate_log_widget", None)

        if error or not isinstance(report_obj, dict) or not report_obj.get("ok"):
            msg = error or "Migration failed."
            if log is not None:
                log.append(f"ERROR: {msg}")
            self.status_bar.showMessage(f"Migration failed: {msg}")
            QMessageBox.critical(self, "Migration failed", msg)
            return

        tables = report_obj.get("tables") or {}
        summary_lines = [
            f"{name}: {info.get('source_rows', 0)} row(s)"
            for name, info in tables.items()
        ]
        summary = "\n".join(summary_lines) or "(no table rows found)"
        if log is not None:
            log.append("Done.")
        self.status_bar.showMessage("Migration complete — reconnecting to PostgreSQL")

        # Reconnect to the migrated PostgreSQL database
        try:
            self.db, self.db_type = connect_guri_database(logger=self.logger)
            if self.db_type == "postgres" and self.db.pg_config:
                self._set_conn_status(
                    self._postgres_connection_label(self.db.pg_config), ok=True
                )
            self.decompositor = GURIDecompositor(database=self.db)
            self.component_library = GURIComponentLibrary(database=self.db)
            self._load_records()
            self._refresh_welcome_dashboard()
            self._refresh_database_browser()
        except Exception as exc:
            self.logger.error("Reconnect after migrate failed: %s", exc)

        QMessageBox.information(
            self,
            "Migration complete",
            "Tables copied to PostgreSQL.\n\n" + summary,
        )
        dlg = getattr(self, "_migrate_dialog", None)
        if dlg is not None and dlg.isVisible():
            dlg.accept()
    
    def _create_create_tab(self):
        """Create the create GURI tab."""
        create_frame = QWidget()
        self.notebook.addTab(create_frame, "Create GURI")
        
        main_layout = QVBoxLayout()
        create_frame.setLayout(main_layout)
        main_layout.setContentsMargins(20, 20, 20, 20)
        
        # Document Type Selection
        type_group = QGroupBox("Document Type")
        type_layout = QHBoxLayout()
        type_group.setLayout(type_layout)
        
        # Create radio button groups — one shared QButtonGroup so only one
        # document type can be selected across Standard / Aliniant / Sensitive.
        self.doc_type_var = "01"  # Default to email
        self.doc_type_button_group = QButtonGroup(create_frame)
        self.doc_type_button_group.setExclusive(True)
        
        # Separate types into three groups
        standard_types = []
        aliniant_types = []
        sensitive_types = []
        
        for code, name in DOCUMENT_TYPES.items():
            if code in ["2a", "2b", "29", "28"]:
                sensitive_types.append((code, name))
            elif name.startswith("Aliniant"):
                aliniant_types.append((code, name))
            else:
                standard_types.append((code, name))
        
        # Standard types (left column)
        standard_group = QGroupBox("📄 Standard")
        standard_layout = QGridLayout()
        standard_group.setLayout(standard_layout)
        
        row = 0
        col = 0
        for code, name in standard_types:
            display_code = f"0x{code}"
            rb = QRadioButton(f"{display_code} - {name}")
            rb.setProperty("doc_type", code)
            self.doc_type_button_group.addButton(rb)
            if code == "01":
                rb.setChecked(True)
                # For email, add strikethrough styling
                font = rb.font()
                font.setStrikeOut(True)
                rb.setFont(font)
            rb.toggled.connect(lambda checked, c=code: self._on_doc_type_radio_changed(c) if checked else None)
            standard_layout.addWidget(rb, row, col)
            col += 1
            if col > 1:
                col = 0
                row += 1
        
        type_layout.addWidget(standard_group)
        
        # Aliniant types (middle column)
        aliniant_group = QGroupBox("Aliniant Specific")
        aliniant_layout = QVBoxLayout()
        aliniant_group.setLayout(aliniant_layout)
        
        for code, name in aliniant_types:
            display_code = f"0x{code}"
            rb = QRadioButton(f"{display_code} - {name}")
            rb.setProperty("doc_type", code)
            self.doc_type_button_group.addButton(rb)
            rb.toggled.connect(lambda checked, c=code: self._on_doc_type_radio_changed(c) if checked else None)
            aliniant_layout.addWidget(rb)
        
        type_layout.addWidget(aliniant_group)
        
        # Sensitive types (right column)
        sensitive_group = QGroupBox("Sensitive / Classified")
        sensitive_layout = QVBoxLayout()
        sensitive_group.setLayout(sensitive_layout)
        
        for code, name in sensitive_types:
            display_code = f"0x{code}"
            rb = QRadioButton(f"{display_code} - {name}")
            rb.setProperty("doc_type", code)
            self.doc_type_button_group.addButton(rb)
            rb.toggled.connect(lambda checked, c=code: self._on_doc_type_radio_changed(c) if checked else None)
            sensitive_layout.addWidget(rb)
        
        type_layout.addWidget(sensitive_group)
        
        # Note frame
        note_group = QGroupBox("Note")
        note_layout = QVBoxLayout()
        note_group.setLayout(note_layout)
        
        note_text = QLabel(
            "Emails have been ruled through because\n"
            "GURIs are automatically added to both\n"
            "incoming and outgoing emails based on\n"
            "key data points:\n\n"
            "• Sender\n"
            "• Recipient\n"
            "• Subject\n"
            "• Time\n"
            "• Type"
        )
        note_text.setStyleSheet("color: #666666;")
        note_layout.addWidget(note_text)
        
        type_layout.addWidget(note_group)
        
        main_layout.addWidget(type_group)
        
        # Input fields
        fields_group = QGroupBox("GURI Parameters")
        fields_layout = QGridLayout()
        fields_group.setLayout(fields_layout)
        
        # Sender
        sender_label = QLabel("Sender:")
        fields_layout.addWidget(sender_label, 0, 0)
        
        sender_widget = QWidget()
        sender_hbox = QHBoxLayout()
        sender_widget.setLayout(sender_hbox)
        sender_hbox.setContentsMargins(0, 0, 0, 0)
        
        self.sender_entry = QLineEdit()
        self.sender_entry.setPlaceholderText("Enter sender email")
        sender_hbox.addWidget(self.sender_entry)
        
        # Quick-add buttons
        jg_btn = QPushButton("JG")
        jg_btn.setMaximumWidth(50)
        jg_btn.clicked.connect(lambda: self._set_sender("julian.garrett@aliniant.com"))
        sender_hbox.addWidget(jg_btn)
        
        admin_btn = QPushButton("Admin")
        admin_btn.setMaximumWidth(70)
        admin_btn.clicked.connect(lambda: self._set_sender("administrator@aliniant.com"))
        sender_hbox.addWidget(admin_btn)
        
        secure_btn = QPushButton("Secure")
        secure_btn.setMaximumWidth(70)
        secure_btn.clicked.connect(lambda: self._set_sender("aliniantsecurecontact@protonmail.com"))
        sender_hbox.addWidget(secure_btn)
        
        fields_layout.addWidget(sender_widget, 0, 1)
        
        # Recipients
        recipients_label = QLabel("Recipients:")
        fields_layout.addWidget(recipients_label, 1, 0)
        
        recipients_widget = QWidget()
        recipients_hbox = QHBoxLayout()
        recipients_widget.setLayout(recipients_hbox)
        recipients_hbox.setContentsMargins(0, 0, 0, 0)
        
        self.recipients_combobox = QComboBox()
        self.recipients_combobox.setEditable(True)
        self.recipients_combobox.addItems(sorted(self.recipients_list))
        self.recipients_combobox.currentTextChanged.connect(self._on_recipient_change)
        recipients_hbox.addWidget(self.recipients_combobox)
        
        add_recipient_btn = QPushButton("+ Add New")
        add_recipient_btn.clicked.connect(self._add_new_recipient)
        recipients_hbox.addWidget(add_recipient_btn)
        
        fields_layout.addWidget(recipients_widget, 1, 1)
        
        # Subject
        subject_label = QLabel("Subject/Name:")
        fields_layout.addWidget(subject_label, 2, 0)
        
        self.subject_combobox = QComboBox()
        self.subject_combobox.setEditable(True)
        fields_layout.addWidget(self.subject_combobox, 2, 1)
        
        # DateTime
        datetime_label = QLabel("Date/Time:")
        fields_layout.addWidget(datetime_label, 3, 0)
        
        datetime_widget = QWidget()
        datetime_hbox = QHBoxLayout()
        datetime_widget.setLayout(datetime_hbox)
        datetime_hbox.setContentsMargins(0, 0, 0, 0)
        
        self.datetime_entry = QLineEdit()
        self.datetime_entry.setText(datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
        datetime_hbox.addWidget(self.datetime_entry)
        
        now_btn = QPushButton("Now")
        now_btn.clicked.connect(self._set_current_datetime)
        datetime_hbox.addWidget(now_btn)
        
        fields_layout.addWidget(datetime_widget, 3, 1)
        
        # Risk / Document Location
        self.risk_label = QLabel("Document Location:")
        fields_layout.addWidget(self.risk_label, 4, 0)
        
        self.risk_entry = QLineEdit()
        self.risk_entry.setPlaceholderText("Enter document location or risk level")
        fields_layout.addWidget(self.risk_entry, 4, 1)
        
        fields_layout.setColumnStretch(1, 1)
        
        main_layout.addWidget(fields_group)
        
        # Buttons
        button_frame = QWidget()
        button_layout = QHBoxLayout()
        button_frame.setLayout(button_layout)
        
        generate_btn = QPushButton("Generate GURI")
        generate_btn.setStyleSheet("font-weight: bold;")
        generate_btn.clicked.connect(self._generate_guri)
        button_layout.addWidget(generate_btn)
        
        clear_btn = QPushButton("Clear Fields")
        clear_btn.clicked.connect(self._clear_create_fields)
        button_layout.addWidget(clear_btn)
        
        autofill_btn = QPushButton("Auto-fill Non-Email")
        autofill_btn.clicked.connect(self._autofill_non_email)
        button_layout.addWidget(autofill_btn)
        
        button_layout.addStretch()
        
        main_layout.addWidget(button_frame)
        
        # Result display
        result_group = QGroupBox("Generated GURI")
        result_layout = QVBoxLayout()
        result_group.setLayout(result_layout)
        
        self.guri_result_text = QTextEdit()
        self.guri_result_text.setReadOnly(True)
        self.guri_result_text.setMaximumHeight(100)
        result_layout.addWidget(self.guri_result_text)
        
        main_layout.addWidget(result_group)
        
        main_layout.addStretch()
        
        # Set initial state
        self._on_doc_type_change()
    
    def _create_search_tab(self):
        """Create the search tab."""
        search_frame = QWidget()
        self.notebook.addTab(search_frame, "Search")
        
        main_layout = QVBoxLayout()
        search_frame.setLayout(main_layout)
        main_layout.setContentsMargins(10, 10, 10, 10)
        
        # Search controls
        controls_group = QGroupBox("Search Criteria")
        controls_layout = QGridLayout()
        controls_group.setLayout(controls_layout)
        
        sender_label = QLabel("Sender:")
        controls_layout.addWidget(sender_label, 0, 0)
        
        self.search_sender_entry = QLineEdit()
        self.search_sender_entry.setPlaceholderText("Enter sender email to search")
        controls_layout.addWidget(self.search_sender_entry, 0, 1)
        
        search_btn = QPushButton("Search")
        search_btn.clicked.connect(self._perform_search)
        controls_layout.addWidget(search_btn, 0, 2)
        
        clear_btn = QPushButton("Clear")
        clear_btn.clicked.connect(self._clear_search)
        controls_layout.addWidget(clear_btn, 0, 3)
        
        controls_layout.setColumnStretch(1, 1)
        
        main_layout.addWidget(controls_group)
        
        # Results tree
        results_group = QGroupBox("Search Results")
        results_layout = QVBoxLayout()
        results_group.setLayout(results_layout)
        
        columns = ("ID", "GURI", "Sender", "Recipients", "Subject", "DateTime", "Risk", "Doc Type", "Created")
        self.search_tree = QTreeWidget()
        self.search_tree.setAlternatingRowColors(True)
        self.search_tree.setRootIsDecorated(False)
        self.search_tree.setHeaderLabels(columns)
        self.search_tree.setColumnCount(len(columns))
        
        # Set column widths
        self.search_tree.setColumnWidth(0, 50)   # ID
        self.search_tree.setColumnWidth(1, 200)  # GURI
        self.search_tree.setColumnWidth(2, 180)  # Sender
        self.search_tree.setColumnWidth(3, 120)  # Recipients
        self.search_tree.setColumnWidth(4, 200)  # Subject
        self.search_tree.setColumnWidth(5, 130)  # DateTime
        self.search_tree.setColumnWidth(6, 100)  # Risk
        self.search_tree.setColumnWidth(7, 100)  # Doc Type
        self.search_tree.setColumnWidth(8, 150)  # Created
        
        # Track sort state for search tree
        self.search_tree_sort_column = None
        self.search_tree_sort_reverse = False
        
        # Enable sorting
        self.search_tree.setSortingEnabled(True)
        self.search_tree.header().sectionClicked.connect(lambda col: self._on_search_tree_header_clicked(col))
        
        # Bind double-click to view details
        self.search_tree.itemDoubleClicked.connect(self._view_search_record_details)
        
        results_layout.addWidget(self.search_tree)
        
        main_layout.addWidget(results_group)
    
    def _create_decompositor_tab(self):
        """Create the GURI decompositor tab."""
        decomp_frame = QWidget()
        self.notebook.addTab(decomp_frame, "Decompositor")
        
        main_layout = QVBoxLayout()
        decomp_frame.setLayout(main_layout)
        main_layout.setContentsMargins(10, 10, 10, 10)
        
        # Input section
        input_group = QGroupBox("GURI Input")
        input_layout = QVBoxLayout()
        input_group.setLayout(input_layout)
        
        input_controls = QWidget()
        input_controls_layout = QHBoxLayout()
        input_controls.setLayout(input_controls_layout)
        
        guri_label = QLabel("GURI:")
        input_controls_layout.addWidget(guri_label)
        
        self.decomp_guri_entry = QLineEdit()
        self.decomp_guri_entry.setPlaceholderText("Enter GURI to decompose (e.g., a1b2cx3d4e5xf6g7h8i9xj0k1l2m3xn4oxp5)")
        self.decomp_guri_entry.returnPressed.connect(self._decompose_guri)
        input_controls_layout.addWidget(self.decomp_guri_entry)
        
        paste_btn = QPushButton("📋 Paste")
        paste_btn.clicked.connect(self._paste_guri)
        input_controls_layout.addWidget(paste_btn)
        
        decompose_btn = QPushButton("Decompose")
        decompose_btn.setStyleSheet("font-weight: bold;")
        decompose_btn.clicked.connect(self._decompose_guri)
        input_controls_layout.addWidget(decompose_btn)
        
        clear_btn = QPushButton("Clear")
        clear_btn.clicked.connect(self._clear_decompositor)
        input_controls_layout.addWidget(clear_btn)
        
        input_layout.addWidget(input_controls)
        main_layout.addWidget(input_group)
        
        # Summary section
        summary_group = QGroupBox("Decomposition Summary")
        summary_layout = QVBoxLayout()
        summary_group.setLayout(summary_layout)
        
        self.decomp_summary_text = QTextEdit()
        self.decomp_summary_text.setReadOnly(True)
        self.decomp_summary_text.setMaximumHeight(150)
        self.decomp_summary_text.setFont(QFont("Courier", 9))
        summary_layout.addWidget(self.decomp_summary_text)
        
        main_layout.addWidget(summary_group)
        
        # Components tree
        components_group = QGroupBox("Component Breakdown")
        components_layout = QVBoxLayout()
        components_group.setLayout(components_layout)
        
        columns = ("Component", "Real Value", "Hex Component", "Expected Len", "Actual Len", "Length ✓", "Hex ✓", "Status")
        self.decomp_tree = QTreeWidget()
        self.decomp_tree.setAlternatingRowColors(True)
        self.decomp_tree.setRootIsDecorated(False)
        self.decomp_tree.setHeaderLabels(columns)
        self.decomp_tree.setColumnCount(len(columns))
        
        # Set column widths
        self.decomp_tree.setColumnWidth(0, 180)  # Component
        self.decomp_tree.setColumnWidth(1, 300)   # Real Value
        self.decomp_tree.setColumnWidth(2, 120)   # Hex Component
        self.decomp_tree.setColumnWidth(3, 100)   # Expected Len
        self.decomp_tree.setColumnWidth(4, 100)   # Actual Len
        self.decomp_tree.setColumnWidth(5, 80)    # Length ✓
        self.decomp_tree.setColumnWidth(6, 80)    # Hex ✓
        self.decomp_tree.setColumnWidth(7, 120)   # Status
        
        # Enable sorting
        self.decomp_tree.setSortingEnabled(True)
        
        components_layout.addWidget(self.decomp_tree)
        main_layout.addWidget(components_group)
        
        # Database records section (for browsing and loading GURIs)
        db_group = QGroupBox("Database Records (Double-click to load GURI)")
        db_layout = QVBoxLayout()
        db_group.setLayout(db_layout)
        
        # Search controls
        search_controls = QWidget()
        search_controls_layout = QHBoxLayout()
        search_controls.setLayout(search_controls_layout)
        
        search_label = QLabel("Search:")
        search_controls_layout.addWidget(search_label)
        
        self.decomp_search_entry = QLineEdit()
        self.decomp_search_entry.setPlaceholderText("Enter sender email to search")
        self.decomp_search_entry.returnPressed.connect(self._decomp_search)
        search_controls_layout.addWidget(self.decomp_search_entry)
        
        search_btn = QPushButton("Search")
        search_btn.clicked.connect(self._decomp_search)
        search_controls_layout.addWidget(search_btn)
        
        clear_search_btn = QPushButton("Clear")
        clear_search_btn.clicked.connect(self._decomp_clear_search)
        search_controls_layout.addWidget(clear_search_btn)
        
        search_controls_layout.addStretch()
        
        prev_btn = QPushButton("⬅ Previous")
        prev_btn.clicked.connect(self._decomp_previous_page)
        search_controls_layout.addWidget(prev_btn)
        
        next_btn = QPushButton("Next ➡")
        next_btn.clicked.connect(self._decomp_next_page)
        search_controls_layout.addWidget(next_btn)
        
        self.decomp_page_label = QLabel("Page: 1")
        search_controls_layout.addWidget(self.decomp_page_label)
        
        self.decomp_record_count_label = QLabel("Total: 0")
        search_controls_layout.addWidget(self.decomp_record_count_label)
        
        db_layout.addWidget(search_controls)
        
        # Database tree
        db_columns = ("GURI", "Subject", "Doc Type", "C1 (Sender)", "C2 (Recip)", "C3 (Subject)", "C4 (Time)", "C5 (Risk/Loc)", "C6 (Type)")
        self.decomp_db_tree = QTreeWidget()
        self.decomp_db_tree.setAlternatingRowColors(True)
        self.decomp_db_tree.setRootIsDecorated(False)
        self.decomp_db_tree.setHeaderLabels(db_columns)
        self.decomp_db_tree.setColumnCount(len(db_columns))
        
        # Set column widths
        self.decomp_db_tree.setColumnWidth(0, 250)  # GURI
        self.decomp_db_tree.setColumnWidth(1, 200)  # Subject
        self.decomp_db_tree.setColumnWidth(2, 100)  # Doc Type
        self.decomp_db_tree.setColumnWidth(3, 80)   # C1
        self.decomp_db_tree.setColumnWidth(4, 80)   # C2
        self.decomp_db_tree.setColumnWidth(5, 80)   # C3
        self.decomp_db_tree.setColumnWidth(6, 80)   # C4
        self.decomp_db_tree.setColumnWidth(7, 80)   # C5
        self.decomp_db_tree.setColumnWidth(8, 80)   # C6
        
        # Bind click/double-click to load GURI
        self.decomp_db_tree.itemClicked.connect(self._decomp_preview_guri_from_db)
        self.decomp_db_tree.itemDoubleClicked.connect(self._decomp_load_guri_from_db)
        
        db_layout.addWidget(self.decomp_db_tree)
        main_layout.addWidget(db_group)
        
        # Initialize decompositor page counter
        self.decomp_current_page = 0
        
        # Load initial records
        self._decomp_load_records()
    
    def _create_status_bar(self):
        """Create the bottom status bar with a permanent info section."""
        self.status_bar = QStatusBar()
        self.status_bar.setSizeGripEnabled(True)

        # Permanent right-hand section: DB connection summary
        self.status_db_label = QLabel("DB: not connected")
        self.status_bar.addPermanentWidget(self.status_db_label)

        self.status_bar.showMessage("Ready")
        self.setStatusBar(self.status_bar)
    
    @staticmethod
    def _postgres_connection_label(cfg: Dict[str, Any]) -> str:
        """Format the status-bar label for a PostgreSQL connection."""
        return f"PostgreSQL — {cfg['host']}:{cfg.get('port', 5432)}/{cfg['database']}"

    def _connect_database_default(self):
        """Connect to PostgreSQL (required). Legacy engines are not auto-selected."""
        try:
            self.db, self.db_type = connect_guri_database(logger=self.logger)
            if self.db_type == "postgres":
                pg_cfg = self.db.pg_config
                if pg_cfg is None:
                    raise RuntimeError("PostgreSQL connection is missing configuration")
                self._set_conn_status(self._postgres_connection_label(pg_cfg), ok=True)
                self.status_bar.showMessage("Connected to PostgreSQL database")
                self.logger.info(
                    "Auto-connected to PostgreSQL: %s:%s/%s",
                    pg_cfg["host"],
                    pg_cfg.get("port", 5432),
                    pg_cfg["database"],
                )
            elif self.db_type == "mysql":
                mysql_cfg = self.db.mysql_config
                if mysql_cfg is None:
                    raise RuntimeError("MySQL connection is missing configuration")
                self._set_conn_status(
                    f"MySQL (deprecated) — {mysql_cfg['host']}:{mysql_cfg.get('port', 3306)}/{mysql_cfg['database']}",
                    ok=True,
                )
                self.status_bar.showMessage(
                    "Connected to deprecated MySQL — switch to PostgreSQL"
                )
            else:
                self._set_conn_status(
                    f"SQLite (deprecated) — {self.db.db_path}", ok=True
                )
                self.status_bar.showMessage(
                    "Connected to deprecated SQLite — switch to PostgreSQL"
                )

            self.decompositor = GURIDecompositor(database=self.db)
            self.component_library = GURIComponentLibrary(database=self.db)
            self._load_records()
            self._refresh_welcome_dashboard()
            self._refresh_database_browser()
        except Exception as e:
            self._set_conn_status("Not connected — PostgreSQL required", ok=False)
            QMessageBox.critical(
                self,
                "PostgreSQL required",
                "GURI now uses PostgreSQL only (MySQL and SQLite are deprecated).\n\n"
                f"{e}\n\n"
                "Create C:\\GeoFooter\\guri_postgres_config.json from the example file, "
                "then use Connect to PostgreSQL…",
            )
            self.logger.error(f"Database connection error: {e}")
    
    def _show_postgres_connection_dialog(self):
        """Show PostgreSQL connection dialog."""
        dialog = QDialog(self)
        dialog.setWindowTitle("Connect to PostgreSQL")
        dialog.setModal(True)
        dialog.resize(400, 300)
        
        layout = QVBoxLayout()
        dialog.setLayout(layout)

        defaults = {
            "host": "localhost",
            "port": "5432",
            "user": "postgres",
            "password": "",
            "database": "guri_db",
        }
        cfg_path = os.path.join(r"C:\GeoFooter", "guri_postgres_config.json")
        if os.path.exists(cfg_path):
            try:
                with open(cfg_path, "r", encoding="utf-8") as fh:
                    saved = json.load(fh)
                defaults["host"] = str(saved.get("host") or defaults["host"])
                defaults["port"] = str(saved.get("port") or defaults["port"])
                defaults["user"] = str(saved.get("user") or defaults["user"])
                defaults["password"] = str(saved.get("password") or "")
                defaults["database"] = str(saved.get("database") or defaults["database"])
            except Exception:
                pass
        
        # Fields
        fields_frame = QWidget()
        fields_layout = QGridLayout()
        fields_frame.setLayout(fields_layout)
        
        host_label = QLabel("Host:")
        host_entry = QLineEdit(defaults["host"])
        fields_layout.addWidget(host_label, 0, 0)
        fields_layout.addWidget(host_entry, 0, 1)
        
        port_label = QLabel("Port:")
        port_entry = QLineEdit(defaults["port"])
        fields_layout.addWidget(port_label, 1, 0)
        fields_layout.addWidget(port_entry, 1, 1)
        
        user_label = QLabel("User:")
        user_entry = QLineEdit(defaults["user"])
        fields_layout.addWidget(user_label, 2, 0)
        fields_layout.addWidget(user_entry, 2, 1)
        
        password_label = QLabel("Password:")
        password_entry = QLineEdit(defaults["password"])
        password_entry.setEchoMode(QLineEdit.EchoMode.Password)
        fields_layout.addWidget(password_label, 3, 0)
        fields_layout.addWidget(password_entry, 3, 1)
        
        database_label = QLabel("Database:")
        database_entry = QLineEdit(defaults["database"])
        fields_layout.addWidget(database_label, 4, 0)
        fields_layout.addWidget(database_entry, 4, 1)
        
        layout.addWidget(fields_frame)
        
        def connect():
            try:
                pg_config = {
                    "host": host_entry.text(),
                    "port": int(port_entry.text()),
                    "user": user_entry.text(),
                    "password": password_entry.text(),
                    "database": database_entry.text(),
                }
                
                self.db = GURIDatabase(db_type="postgres", pg_config=pg_config)
                self.db_type = "postgres"
                self._set_conn_status(
                    self._postgres_connection_label(pg_config),
                    ok=True,
                )
                
                # Initialize decompositor and component library with database
                self.decompositor = GURIDecompositor(database=self.db)
                self.component_library = GURIComponentLibrary(database=self.db)
                
                self._load_records()
                self._refresh_welcome_dashboard()
                self._refresh_database_browser()
                self.status_bar.showMessage("Connected to PostgreSQL database")
                dialog.accept()
                QMessageBox.information(
                    self, "Success", "Connected to PostgreSQL database successfully!"
                )
                
            except Exception as e:
                QMessageBox.critical(
                    self, "Connection Error", f"Failed to connect to PostgreSQL:\n{str(e)}"
                )
                self.logger.error(f"PostgreSQL connection error: {e}")
        
        # Buttons — affirmative left of Cancel (Cancel always bottom-right)
        connect_btn = QPushButton("Connect")
        connect_btn.setDefault(True)
        connect_btn.clicked.connect(connect)
        cancel_btn = QPushButton("Cancel")
        cancel_btn.clicked.connect(dialog.reject)
        add_dialog_button_row(
            layout, affirmative=connect_btn, cancel=cancel_btn
        )
        
        dialog.exec()
    
    def _show_sqlite_connection_dialog(self):
        """Deprecated SQLite open path — emergency use only."""
        warn = QMessageBox.warning(
            self,
            "SQLite is deprecated",
            "SQLite is deprecated. GURI has cut over to PostgreSQL.\n\n"
            "Only continue if you need emergency read access to an old .db file.\n\n"
            "Continue anyway?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if warn != QMessageBox.StandardButton.Yes:
            return
        filename, _ = QFileDialog.getOpenFileName(
            self,
            "Select SQLite Database (deprecated)",
            "C:/GeoFooter",
            "SQLite Database (*.db);;All Files (*.*)"
        )
        
        if filename:
            try:
                self.db = GURIDatabase(db_type="sqlite", db_path=filename)
                self.db_type = "sqlite"
                self._set_conn_status(f"SQLite (deprecated) — {filename}", ok=True)
                
                # Initialize decompositor and component library with database
                self.decompositor = GURIDecompositor(database=self.db)
                self.component_library = GURIComponentLibrary(database=self.db)
                
                self._load_records()
                self._refresh_welcome_dashboard()
                self._refresh_database_browser()
                self.status_bar.showMessage(
                    "Connected to deprecated SQLite — switch to PostgreSQL"
                )
            except Exception as e:
                QMessageBox.critical(self, "Connection Error", f"Failed to connect to SQLite:\n{str(e)}")
                self.logger.error(f"SQLite connection error: {e}")
    
    def _on_tree_header_clicked(self, column):
        """Handle tree header click for sorting."""
        if self.tree_sort_column == column:
            self.tree_sort_reverse = not self.tree_sort_reverse
        else:
            self.tree_sort_column = column
            self.tree_sort_reverse = False
        self._sort_tree_column(column, self.tree)
    
    def _on_search_tree_header_clicked(self, column):
        """Handle search tree header click for sorting."""
        if self.search_tree_sort_column == column:
            self.search_tree_sort_reverse = not self.search_tree_sort_reverse
        else:
            self.search_tree_sort_column = column
            self.search_tree_sort_reverse = False
        self._sort_tree_column(column, self.search_tree)
    
    def _load_records(self):
        """Load records from database."""
        if not self.db:
            QMessageBox.warning(self, "No Connection", "Please connect to a database first.")
            return
        
        try:
            # Clear existing items
            self.tree.clear()
            
            # Load records
            offset = self.current_page * self.records_per_page
            records = self.db.get_all_records(limit=self.records_per_page, offset=offset)
            self.current_records = records
            
            # Populate tree
            for record in records:
                doc_type_name = DOCUMENT_TYPES.get(record['document_type'], f"Unknown ({record['document_type']})")
                item = QTreeWidgetItem(self.tree)
                item.setText(0, str(record['id']))
                item.setText(1, str(record['guri']))
                item.setText(2, str(record['sender']))
                item.setText(3, str(record['recipients']))
                item.setText(4, str(record['subject']))
                # Convert datetime to string if it's a datetime object
                datetime_str = record['datetime']
                if datetime_str and not isinstance(datetime_str, str):
                    datetime_str = str(datetime_str)
                item.setText(5, datetime_str or '')
                item.setText(6, str(record['avg_risk']))
                item.setText(7, doc_type_name)
                # Convert created_at to string if it's a datetime object
                created_at_str = record['created_at']
                if created_at_str and not isinstance(created_at_str, str):
                    created_at_str = str(created_at_str)
                item.setText(8, created_at_str or '')
                item.setData(0, Qt.ItemDataRole.UserRole, record)  # Store full record data
            
            # Update labels
            total_count = self.db.get_record_count()
            self.record_count_label.setText(f"Total Records: {total_count}")
            self.page_label.setText(f"Page: {self.current_page + 1}")
            
            self.status_bar.showMessage(f"Loaded {len(records)} records")
            
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to load records:\n{str(e)}")
            self.logger.error(f"Error loading records: {e}")
    
    def _previous_page(self):
        """Go to previous page."""
        if self.current_page > 0:
            self.current_page -= 1
            self._load_records()
    
    def _next_page(self):
        """Go to next page."""
        self.current_page += 1
        self._load_records()
        # If no records loaded, go back
        if not self.current_records:
            self.current_page -= 1
            self._load_records()
    
    def _on_doc_type_radio_changed(self, doc_type):
        """Handle document type radio button change."""
        self.doc_type_var = doc_type
        self._on_doc_type_change()
    
    def _on_doc_type_change(self):
        """Handle document type change."""
        if not hasattr(self, 'doc_type_var'):
            return
        
        doc_type = self.doc_type_var
        
        if doc_type == "01":  # Email
            # Change label to Risk Level
            if hasattr(self, 'risk_label'):
                self.risk_label.setText("Risk Level:")
            if hasattr(self, 'sender_entry'):
                self.sender_entry.clear()
            if hasattr(self, 'recipients_combobox'):
                self.recipients_combobox.setCurrentText("")
            if hasattr(self, 'risk_entry'):
                self.risk_entry.clear()
        else:  # Other documents
            # Change label to Document Location
            if hasattr(self, 'risk_label'):
                self.risk_label.setText("Document Location:")
            # Auto-fill sender and recipients only
            self._autofill_non_email()
    
    def _autofill_non_email(self):
        """Auto-fill fields for non-email documents."""
        if not hasattr(self, 'doc_type_var'):
            return
        
        doc_type = self.doc_type_var
        
        if doc_type != "01":
            # Auto-fill sender
            if hasattr(self, 'sender_entry'):
                self.sender_entry.clear()
                self.sender_entry.setText(NON_EMAIL_SENDER)
            
            # Auto-fill recipients with placeholder
            if hasattr(self, 'recipients_combobox'):
                index = self.recipients_combobox.findText(NON_EMAIL_PLACEHOLDER)
                if index >= 0:
                    self.recipients_combobox.setCurrentIndex(index)
                else:
                    self.recipients_combobox.setCurrentText(NON_EMAIL_PLACEHOLDER)
            
            # DO NOT auto-fill location - user must enter it
            # Clear it instead
            if hasattr(self, 'risk_entry'):
                self.risk_entry.clear()
    
    def _set_current_datetime(self):
        """Set current date and time."""
        if hasattr(self, 'datetime_entry'):
            self.datetime_entry.clear()
            self.datetime_entry.setText(datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
    
    def _set_sender(self, email):
        """Set sender email address."""
        if hasattr(self, 'sender_entry'):
            self.sender_entry.clear()
            self.sender_entry.setText(email)
        self.status_bar.showMessage(f"Sender set to: {email}")
    
    def _generate_guri(self):
        """Generate a new GURI."""
        if not self.db:
            QMessageBox.warning(self, "No Connection", "Please connect to a database first.")
            return
        
        # Get values
        sender = self.sender_entry.text().strip()
        recipients = self.recipients_combobox.currentText().strip()
        subject = self.subject_combobox.currentText().strip()
        dt_str = self.datetime_entry.text().strip()
        avg_risk = self.risk_entry.text().strip()
        doc_type = self.doc_type_var
        
        # Determine field name for validation message
        risk_field_name = "Document Location" if doc_type != "01" else "Risk Level"
        
        # Validate
        if not all([sender, recipients, subject, dt_str, avg_risk]):
            missing_fields = []
            if not sender: missing_fields.append("Sender")
            if not recipients: missing_fields.append("Recipients")
            if not subject: missing_fields.append("Subject/Name")
            if not dt_str: missing_fields.append("Date/Time")
            if not avg_risk: missing_fields.append(risk_field_name)
            
            QMessageBox.warning(self, "Validation Error", 
                                 f"The following fields are required:\n\n" + "\n".join(f"• {field}" for field in missing_fields))
            return
        
        try:
            # Generate GURI
            guri = self.db.generate_guri(sender, recipients, subject, dt_str, avg_risk, doc_type)
            
            # Display result
            doc_type_name = DOCUMENT_TYPES.get(doc_type, f"Unknown ({doc_type})")
            result = f"GURI Generated Successfully!\n\n"
            result += f"GURI: {guri}\n"
            result += f"Saved to database: {self.db_type.upper()}\n"
            result += f"Document Type: {doc_type_name}\n"
            result += f"Sender: {sender}\n"
            result += f"Recipients: {recipients}\n"
            result += f"Subject: {subject}\n"
            result += f"DateTime: {dt_str}\n"
            
            # Show appropriate field name
            if doc_type == "01":
                result += f"Risk Level: {avg_risk}\n"
            else:
                result += f"Document Location: {avg_risk}\n"
            
            if hasattr(self, 'guri_result_text'):
                self.guri_result_text.clear()
                self.guri_result_text.setPlainText(result)
            
            self.status_bar.showMessage(f"GURI saved to database: {guri}")
            
            # Refresh view tab
            self._load_records()
            
            QMessageBox.information(self, "Success", f"GURI saved to the database.\n\n{guri}")
            
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to generate GURI:\n{str(e)}")
            self.logger.error(f"Error generating GURI: {e}")
    
    def _clear_create_fields(self):
        """Clear all create fields."""
        if hasattr(self, 'sender_entry'):
            self.sender_entry.clear()
        if hasattr(self, 'recipients_combobox'):
            self.recipients_combobox.setCurrentText("")
        if hasattr(self, 'subject_combobox'):
            self.subject_combobox.setCurrentText("")
        if hasattr(self, 'risk_entry'):
            self.risk_entry.clear()
        if hasattr(self, 'guri_result_text'):
            self.guri_result_text.clear()
        
        # Reset to email type
        self.doc_type_var = "01"
        # Find and check the email radio button
        for widget in self.findChildren(QRadioButton):
            if widget.property("doc_type") == "01":
                widget.setChecked(True)
                break
        
        self._set_current_datetime()
    
    def _perform_search(self):
        """Perform search."""
        if not self.db:
            QMessageBox.warning(self, "No Connection", "Please connect to a database first.")
            return
        
        sender = self.search_sender_entry.text().strip()
        
        if not sender:
            QMessageBox.warning(self, "Validation Error", "Please enter a sender to search for.")
            return
        
        try:
            # Clear existing items
            self.search_tree.clear()
            
            # Search
            records = self.db.search_by_sender(sender, limit=100)
            
            # Populate tree
            for record in records:
                doc_type_name = DOCUMENT_TYPES.get(record['document_type'], f"Unknown ({record['document_type']})")
                item = QTreeWidgetItem(self.search_tree)
                item.setText(0, str(record['id']))
                item.setText(1, str(record['guri']))
                item.setText(2, str(record['sender']))
                item.setText(3, str(record['recipients']))
                item.setText(4, str(record['subject']))
                # Convert datetime to string if it's a datetime object
                datetime_str = record['datetime']
                if datetime_str and not isinstance(datetime_str, str):
                    datetime_str = str(datetime_str)
                item.setText(5, datetime_str or '')
                item.setText(6, str(record['avg_risk']))
                item.setText(7, doc_type_name)
                # Convert created_at to string if it's a datetime object
                created_at_str = record['created_at']
                if created_at_str and not isinstance(created_at_str, str):
                    created_at_str = str(created_at_str)
                item.setText(8, created_at_str or '')
                item.setData(0, Qt.ItemDataRole.UserRole, record)  # Store full record data
            
            # Update column headers for search tree
            self._update_tree_column_headers(self.search_tree)
            
            self.status_bar.showMessage(f"Found {len(records)} matching records")
            
            if not records:
                QMessageBox.information(self, "No Results", "No records found matching the search criteria.")
            
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Search failed:\n{str(e)}")
            self.logger.error(f"Search error: {e}")
    
    def _clear_search(self):
        """Clear search results."""
        self.search_sender_entry.clear()
        self.search_tree.clear()
        self.status_bar.showMessage("Search cleared")
    
    def _focus_search(self):
        """Focus on search tab."""
        for i in range(self.notebook.count()):
            if self.notebook.tabText(i) == "Search":
                self.notebook.setCurrentIndex(i)
                break
        self.search_sender_entry.setFocus()
    
    def _view_record_details(self, item, column):
        """View detailed record information."""
        if not item:
            return
        
        # Get values from item
        values = [item.text(i) for i in range(item.columnCount())]
        # Get full record data if stored
        record_data = item.data(0, Qt.ItemDataRole.UserRole)
        if record_data:
            values = [
                str(record_data.get('id', '')),
                record_data.get('guri', ''),
                record_data.get('sender', ''),
                record_data.get('recipients', ''),
                record_data.get('subject', ''),
                record_data.get('datetime', ''),
                record_data.get('avg_risk', ''),
                DOCUMENT_TYPES.get(record_data.get('document_type', ''), 'Unknown'),
                record_data.get('created_at', '')
            ]
        
        self._show_record_details_dialog(values)
    
    def _view_search_record_details(self, item, column):
        """View detailed record information from search results."""
        if not item:
            return
        
        # Get values from item
        values = [item.text(i) for i in range(item.columnCount())]
        # Get full record data if stored
        record_data = item.data(0, Qt.ItemDataRole.UserRole)
        if record_data:
            values = [
                str(record_data.get('id', '')),
                record_data.get('guri', ''),
                record_data.get('sender', ''),
                record_data.get('recipients', ''),
                record_data.get('subject', ''),
                record_data.get('datetime', ''),
                record_data.get('avg_risk', ''),
                DOCUMENT_TYPES.get(record_data.get('document_type', ''), 'Unknown'),
                record_data.get('created_at', '')
            ]
        
        self._show_record_details_dialog(values)
    
    def _show_record_details_dialog(self, values):
        """Show record details in a dialog."""
        dialog = QDialog(self)
        dialog.setWindowTitle("Record Details")
        dialog.resize(600, 400)
        dialog.setModal(True)
        
        layout = QVBoxLayout()
        dialog.setLayout(layout)
        
        # Details text
        text = QTextEdit()
        text.setReadOnly(True)
        text.setFont(QFont("Courier", 10))
        
        # Determine if this is an email or document based on the doc type string
        doc_type_str = str(values[7]) if len(values) > 7 else ""
        is_email = "Email" in doc_type_str or "01" in doc_type_str
        risk_field_name = "Risk Level" if is_email else "Document Location"
        
        details = f"""
GURI Record Details
{'=' * 60}

ID: {values[0] if len(values) > 0 else ''}
GURI: {values[1] if len(values) > 1 else ''}

Sender: {values[2] if len(values) > 2 else ''}
Recipients: {values[3] if len(values) > 3 else ''}
Subject: {values[4] if len(values) > 4 else ''}

Date/Time: {values[5] if len(values) > 5 else ''}
{risk_field_name}: {values[6] if len(values) > 6 else ''}
Document Type: {values[7] if len(values) > 7 else ''}

Created At: {values[8] if len(values) > 8 else ''}
"""
        
        text.setPlainText(details)
        layout.addWidget(text)
        
        # Close alone — still bottom-right
        close_btn = QPushButton("Close")
        close_btn.clicked.connect(dialog.accept)
        add_dialog_button_row(layout, cancel=close_btn)
        
        dialog.exec()
    
    def _export_records(self):
        """Export records to JSON file."""
        if not self.db:
            QMessageBox.warning(self, "No Connection", "Please connect to a database first.")
            return
        
        filename, _ = QFileDialog.getSaveFileName(
            self,
            "Export Records",
            "",
            "JSON Files (*.json);;All Files (*.*)"
        )
        
        if filename:
            try:
                records = self.db.get_all_records(limit=10000)  # Export up to 10000 records
                
                with open(filename, 'w') as f:
                    json.dump(records, f, indent=2, default=str)
                
                QMessageBox.information(self, "Success", f"Exported {len(records)} records to:\n{filename}")
                self.status_bar.showMessage(f"Exported {len(records)} records")
                
            except Exception as e:
                QMessageBox.critical(self, "Export Error", f"Failed to export records:\n{str(e)}")
                self.logger.error(f"Export error: {e}")
    
    def _show_statistics(self):
        """Show database statistics."""
        if not self.db:
            QMessageBox.warning(self, "No Connection", "Please connect to a database first.")
            return
        
        try:
            total = self.db.get_record_count()
            
            stats = f"""
Database Statistics
{'=' * 40}

Total Records: {total}
Database Type: {self.db_type.upper()}
Records Per Page: {self.records_per_page}
Current Page: {self.current_page + 1}
"""
            
            QMessageBox.information(self, "Database Statistics", stats)
            
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to get statistics:\n{str(e)}")
    
    def _show_about(self):
        """Show about dialog (same content as About tab)."""
        QMessageBox.information(self, "About", about_summary() + "\n\nSee the About tab for release notes.")
    
    def _show_doc_types(self):
        """Show document type codes."""
        doc_types_text = "Document Type Codes\n" + "=" * 40 + "\n\n"
        for code, name in sorted(DOCUMENT_TYPES.items()):
            doc_types_text += f"{code}: {name}\n"
        
        QMessageBox.information(self, "Document Type Codes", doc_types_text)
    
    def _sort_tree_column(self, col, tree):
        """Sort tree contents when a column header is clicked."""
        # QTreeWidget handles sorting automatically, but we can customize
        if col == 0:  # ID column - numeric sort
            tree.sortItems(col, Qt.SortOrder.AscendingOrder if not self.tree_sort_reverse else Qt.SortOrder.DescendingOrder)
        else:
            tree.sortItems(col, Qt.SortOrder.AscendingOrder if not self.tree_sort_reverse else Qt.SortOrder.DescendingOrder)
        
        # Update sort state
        if tree == self.tree:
            self.tree_sort_column = col
            self.tree_sort_reverse = not self.tree_sort_reverse
        else:
            self.search_tree_sort_column = col
            self.search_tree_sort_reverse = not self.search_tree_sort_reverse
    
    def _update_tree_column_headers(self, tree=None):
        """Update column headers to show sort arrows."""
        if tree is None:
            tree = self.tree
        
        # Determine which tree and its sort state
        if tree == self.tree:
            sort_column = self.tree_sort_column
            sort_reverse = self.tree_sort_reverse
        else:
            sort_column = self.search_tree_sort_column
            sort_reverse = self.search_tree_sort_reverse
        
        # Column display names by index
        column_names = [
            "ID",
            "GURI",
            "Sender",
            "Recipients",
            "Subject",
            "Date/Time",
            "Risk/Location",
            "Doc Type",
            "Created At",
        ]
        
        header = tree.headerItem()
        if header is None:
            return
        
        # Update all column headers
        for col_index, display_name in enumerate(column_names):
            if col_index == sort_column:
                arrow = " ▼" if sort_reverse else " ▲"
                header.setText(col_index, display_name + arrow)
            else:
                header.setText(col_index, display_name)
    
    def _decompose_guri(self):
        """Decompose a GURI into its components and show the actual data they represent."""
        guri = self.decomp_guri_entry.text().strip()
        
        if not guri:
            QMessageBox.warning(self, "Input Required", "Please enter a GURI to decompose.")
            return
        
        # Clear previous results
        self.decomp_summary_text.clear()
        self.decomp_tree.clear()
        
        # Initialize decompositor and component library if needed
        if not self.decompositor or self.decompositor.database != self.db:
            self.decompositor = GURIDecompositor(database=self.db)
        if not self.component_library or (hasattr(self.component_library, 'database') and self.component_library.database != self.db):
            self.component_library = GURIComponentLibrary(database=self.db)
        
        try:
            # Decompose using the module
            result = self.decompositor.decompose(guri)
            
            # Display summary
            self.decomp_summary_text.append(f"Original GURI: {result['guri']}")
            self.decomp_summary_text.append(f"Total Length: {len(result['guri'])} characters")
            self.decomp_summary_text.append(f"Components Found: {len(result['components'])}")
            
            if result['global_type']:
                self.decomp_summary_text.append(f"📋 {result['global_type']}")
            
            if result['record_data']:
                self.decomp_summary_text.append(f"✓ Record found in database (ID: {result['record_data']['id']})")
            else:
                self.decomp_summary_text.append("⚠ No matching record found in database")
            
            # Display errors if any
            if result['errors']:
                for error in result['errors']:
                    self.decomp_summary_text.append(f"❌ {error}")
            
            # Display validation result
            if result['is_valid']:
                self.decomp_summary_text.append("✓ VALID GURI - All checks passed")
                self.status_bar.showMessage("Valid GURI decomposed successfully")
            else:
                self.decomp_summary_text.append("❌ INVALID GURI - See errors above")
                self.status_bar.showMessage("Invalid GURI format")
            
            # Populate component tree
            for comp_data in result['component_data']:
                item = QTreeWidgetItem(self.decomp_tree)
                item.setText(0, comp_data['name'])
                
                # Show real value if available, otherwise show decoded value or component
                if comp_data.get('real_value'):
                    item.setText(1, comp_data['real_value'])
                elif comp_data.get('decoded_value'):
                    item.setText(1, comp_data['decoded_value'])
                else:
                    item.setText(1, f"{comp_data['component']} (hashed)")
                
                item.setText(2, comp_data['component'])
                item.setText(3, str(comp_data['expected_length']))
                item.setText(4, str(comp_data['actual_length']))
                item.setText(5, "✓" if comp_data['is_length_valid'] else "❌")
                item.setText(6, "✓" if comp_data['is_hex'] else "❌")
                item.setText(7, comp_data['status'])
                
                # Color code based on status
                if comp_data['status'] == "Valid":
                    item.setForeground(0, QColor('#2E7D32'))
                elif comp_data['status'] == "Invalid Hex":
                    item.setForeground(0, QColor('#C62828'))
                else:
                    item.setForeground(0, QColor('#F57C00'))
            
        except Exception as e:
            self.decomp_summary_text.append(f"❌ ERROR: {str(e)}")
            self.logger.error(f"Error decomposing GURI: {e}")
            self.status_bar.showMessage("Error decomposing GURI")
    
    def _clear_decompositor(self):
        """Clear the decompositor fields."""
        self.decomp_guri_entry.clear()
        self.decomp_summary_text.clear()
        self.decomp_tree.clear()
        self.status_bar.showMessage("Decompositor cleared")
    
    def _paste_guri(self):
        """Paste GURI from clipboard."""
        try:
            clipboard = QApplication.clipboard()
            clipboard_content = clipboard.text()
            self.decomp_guri_entry.clear()
            self.decomp_guri_entry.setText(clipboard_content.strip())
            self.status_bar.showMessage("GURI pasted from clipboard")
        except Exception as e:
            QMessageBox.warning(self, "Clipboard Error", "Could not read from clipboard.")
            self.logger.error(f"Clipboard error: {e}")
    
    def _decomp_load_records(self):
        """Load records into decompositor database viewer."""
        if not self.db:
            return
        
        try:
            # Clear existing items
            self.decomp_db_tree.clear()
            
            # Load records
            offset = self.decomp_current_page * self.records_per_page
            records = self.db.get_all_records(limit=self.records_per_page, offset=offset)
            
            # Populate tree with decomposed GURI components
            for record in records:
                doc_type_name = DOCUMENT_TYPES.get(record['document_type'], f"Unknown ({record['document_type']})")
                subject = record['subject'][:40] + "..." if len(record['subject']) > 40 else record['subject']
                
                # Decompose GURI into components
                guri = record['guri']
                components = guri.split('x')
                
                # Ensure we have 6 components (pad with empty if needed)
                while len(components) < 6:
                    components.append("")
                
                # Column order: GURI, Subject, Doc Type, C1-C6
                item = QTreeWidgetItem(self.decomp_db_tree)
                item.setText(0, guri)
                item.setText(1, subject)
                item.setText(2, doc_type_name)
                item.setText(3, components[0])  # Component 1 (5 hex)
                item.setText(4, components[1])  # Component 2 (5 hex)
                item.setText(5, components[2])  # Component 3 (8 hex)
                item.setText(6, components[3])  # Component 4 (8 hex)
                item.setText(7, components[4])  # Component 5 (3 hex)
                item.setText(8, components[5])  # Component 6 (2 hex)
            
            # Update labels
            total_count = self.db.get_record_count()
            self.decomp_record_count_label.setText(f"Total: {total_count}")
            self.decomp_page_label.setText(f"Page: {self.decomp_current_page + 1}")
            
        except Exception as e:
            self.logger.error(f"Error loading decompositor records: {e}")
    
    def _decomp_previous_page(self):
        """Go to previous page in decompositor database."""
        if self.decomp_current_page > 0:
            self.decomp_current_page -= 1
            self._decomp_load_records()
    
    def _decomp_next_page(self):
        """Go to next page in decompositor database."""
        if not self.db:
            return
        
        self.decomp_current_page += 1
        self._decomp_load_records()
        
        # If no records loaded, go back
        if self.decomp_db_tree.topLevelItemCount() == 0:
            self.decomp_current_page -= 1
            self._decomp_load_records()
    
    def _decomp_search(self):
        """Search database in decompositor."""
        if not self.db:
            QMessageBox.warning(self, "No Connection", "Please connect to a database first.")
            return
        
        search_term = self.decomp_search_entry.text().strip()
        
        if not search_term:
            QMessageBox.warning(self, "Validation Error", "Please enter a search term.")
            return
        
        try:
            # Clear existing items
            self.decomp_db_tree.clear()
            
            # Search by sender
            records = self.db.search_by_sender(search_term, limit=1000)
            
            # Populate tree with decomposed GURI components
            for record in records:
                doc_type_name = DOCUMENT_TYPES.get(record['document_type'], f"Unknown ({record['document_type']})")
                subject = record['subject'][:40] + "..." if len(record['subject']) > 40 else record['subject']
                
                # Decompose GURI into components
                guri = record['guri']
                components = guri.split('x')
                
                # Ensure we have 6 components (pad with empty if needed)
                while len(components) < 6:
                    components.append("")
                
                # Column order: GURI, Subject, Doc Type, C1-C6
                item = QTreeWidgetItem(self.decomp_db_tree)
                item.setText(0, guri)
                item.setText(1, subject)
                item.setText(2, doc_type_name)
                item.setText(3, components[0])  # Component 1 (5 hex)
                item.setText(4, components[1])  # Component 2 (5 hex)
                item.setText(5, components[2])  # Component 3 (8 hex)
                item.setText(6, components[3])  # Component 4 (8 hex)
                item.setText(7, components[4])  # Component 5 (3 hex)
                item.setText(8, components[5])  # Component 6 (2 hex)
            
            # Update labels
            self.decomp_record_count_label.setText(f"Found: {len(records)}")
            self.decomp_page_label.setText("Search Results")
            
            if not records:
                QMessageBox.information(self, "No Results", "No records found matching the search criteria.")
        
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Search failed:\n{str(e)}")
            self.logger.error(f"Decompositor search error: {e}")
    
    def _decomp_clear_search(self):
        """Clear search in decompositor."""
        self.decomp_search_entry.clear()
        self.decomp_current_page = 0
        self._decomp_load_records()
    
    def _decomp_load_guri_from_db(self, item, column):
        """Load GURI from database into decompositor when double-clicked."""
        if not item:
            return
        
        guri = item.text(0)  # First column is GURI
        
        # Load into decompositor entry
        self.decomp_guri_entry.clear()
        self.decomp_guri_entry.setText(guri)
        
        # Automatically decompose
        self._decompose_guri()
        
        self.status_bar.showMessage(f"Loaded GURI from database: {guri}")
    
    def _decomp_preview_guri_from_db(self, item, column):
        """Preview GURI decomposition when single-clicked (non-intrusive)."""
        if not item:
            return
        
        guri = item.text(0)  # First column is GURI
        
        # Only update if the GURI field is empty or different
        current_guri = self.decomp_guri_entry.text().strip()
        if current_guri != guri:
            # Load into decompositor entry
            self.decomp_guri_entry.clear()
            self.decomp_guri_entry.setText(guri)
            
            # Automatically decompose
            self._decompose_guri()
            
            self.status_bar.showMessage(f"Preview: {guri}")
    
    def _load_recipients(self):
        """Load recipients list from file."""
        recipients_file = os.path.join(os.path.dirname(__file__), "recipients.txt")
        if os.path.exists(recipients_file):
            try:
                with open(recipients_file, 'r', encoding='utf-8') as f:
                    self.recipients_list = [line.strip() for line in f if line.strip()]
                self.logger.info(f"Loaded {len(self.recipients_list)} recipients from file")
            except Exception as e:
                self.logger.error(f"Error loading recipients: {e}")
                self.recipients_list = []
        else:
            # Create with some defaults
            self.recipients_list = [
                NON_EMAIL_PLACEHOLDER,
                "julian.garrett@aliniant.com",
                "team@aliniant.com"
            ]
            self._save_recipients()
    
    def _save_recipients(self):
        """Save recipients list to file."""
        recipients_file = os.path.join(os.path.dirname(__file__), "recipients.txt")
        try:
            with open(recipients_file, 'w', encoding='utf-8') as f:
                for recipient in sorted(set(self.recipients_list)):
                    f.write(f"{recipient}\n")
            self.logger.info(f"Saved {len(self.recipients_list)} recipients to file")
        except Exception as e:
            self.logger.error(f"Error saving recipients: {e}")
    
    def _check_new_recipient(self, event=None):
        """Check if the entered recipient is new and offer to add it."""
        recipient = self.recipients_combobox.currentText().strip()
        
        if not recipient:
            return
        
        # If it's already in the list, no action needed
        if recipient in self.recipients_list:
            return
        
        # Ask if user wants to add it
        response = QMessageBox.question(
            self,
            "New Recipient", 
            f"'{recipient}' is not in the recipients list.\n\nWould you like to add it?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.Yes
        )
        
        if response == QMessageBox.StandardButton.Yes:
            self.recipients_list.append(recipient)
            self.recipients_combobox.clear()
            self.recipients_combobox.addItems(sorted(self.recipients_list))
            self.recipients_combobox.setCurrentText(recipient)
            self._save_recipients()
            self.status_bar.showMessage(f"Added '{recipient}' to recipients list")
    
    def _add_new_recipient(self):
        """Manually add a new recipient."""
        # Get current value
        current = self.recipients_combobox.currentText().strip()
        
        if current and current not in self.recipients_list:
            self.recipients_list.append(current)
            self.recipients_combobox.clear()
            self.recipients_combobox.addItems(sorted(self.recipients_list))
            self.recipients_combobox.setCurrentText(current)
            self._save_recipients()
            QMessageBox.information(self, "Success", f"Added '{current}' to recipients list")
            self.status_bar.showMessage(f"Added '{current}' to recipients list")
        elif current in self.recipients_list:
            QMessageBox.information(self, "Already Exists", f"'{current}' is already in the recipients list")
        else:
            QMessageBox.warning(self, "No Recipient", "Please enter a recipient first")
    
    def _on_recipient_change(self, event=None):
        """Handle recipient change - check for new recipient and update subject dropdown."""
        self._check_new_recipient(event)
        self._update_subject_dropdown(event)
    
    def _update_subject_dropdown(self, event=None):
        """Update subject dropdown based on selected recipient."""
        if not self.db:
            return
        
        recipient = self.recipients_combobox.currentText().strip()
        
        if not recipient:
            # Clear subject dropdown if no recipient
            self.subject_combobox.clear()
            return
        
        try:
            # Get previous subjects for this recipient
            subjects = self.db.get_subjects_by_recipient(recipient, limit=100)
            
            # Update the combobox values
            current_text = self.subject_combobox.currentText()
            self.subject_combobox.clear()
            self.subject_combobox.addItems(subjects)
            if current_text in subjects:
                self.subject_combobox.setCurrentText(current_text)
            
            # Update status bar
            if subjects:
                self.status_bar.showMessage(f"Found {len(subjects)} previous subject(s) for {recipient}")
            else:
                self.status_bar.showMessage(f"No previous subjects found for {recipient}")
                
        except Exception as e:
            self.logger.error(f"Error updating subject dropdown: {e}")
            self.subject_combobox.clear()
    
    def _create_suppliers_tab(self):
        """Create the suppliers and contract codes tab."""
        suppliers_frame = QWidget()
        self.notebook.addTab(suppliers_frame, "Suppliers")
        
        main_layout = QHBoxLayout()
        suppliers_frame.setLayout(main_layout)
        main_layout.setContentsMargins(10, 10, 10, 10)
        
        # Left: supplier form
        form_group = QGroupBox("Supplier Details")
        form_layout = QGridLayout()
        form_group.setLayout(form_layout)
        
        form_layout.addWidget(QLabel("Extract from Email:"), 0, 0)
        self.extract_email_entry = QLineEdit()
        self.extract_email_entry.setPlaceholderText("name@company.com")
        form_layout.addWidget(self.extract_email_entry, 0, 1)
        extract_btn = QPushButton("Extract")
        extract_btn.clicked.connect(self._extract_company_from_email)
        form_layout.addWidget(extract_btn, 0, 2)
        
        form_layout.addWidget(QLabel("Company Name:"), 1, 0)
        self.supplier_company_entry = QLineEdit()
        form_layout.addWidget(self.supplier_company_entry, 1, 1, 1, 2)
        
        form_layout.addWidget(QLabel("Company Domain:"), 2, 0)
        self.supplier_domain_entry = QLineEdit()
        form_layout.addWidget(self.supplier_domain_entry, 2, 1, 1, 2)
        
        form_layout.addWidget(QLabel("Supplier Code:"), 3, 0)
        self.supplier_code_entry = QLineEdit()
        form_layout.addWidget(self.supplier_code_entry, 3, 1, 1, 2)
        
        form_layout.addWidget(QLabel("Contract Code:"), 4, 0)
        self.contract_code_entry = QLineEdit()
        form_layout.addWidget(self.contract_code_entry, 4, 1, 1, 2)
        
        form_layout.addWidget(QLabel("Contact Person:"), 5, 0)
        self.contact_person_entry = QLineEdit()
        form_layout.addWidget(self.contact_person_entry, 5, 1, 1, 2)
        
        form_layout.addWidget(QLabel("Contact Email:"), 6, 0)
        self.contact_email_entry = QLineEdit()
        form_layout.addWidget(self.contact_email_entry, 6, 1, 1, 2)
        
        form_layout.addWidget(QLabel("Notes:"), 7, 0)
        self.supplier_notes_text = QTextEdit()
        self.supplier_notes_text.setMaximumHeight(120)
        form_layout.addWidget(self.supplier_notes_text, 7, 1, 1, 2)
        
        buttons = QWidget()
        buttons_layout = QHBoxLayout()
        buttons.setLayout(buttons_layout)
        buttons_layout.setContentsMargins(0, 0, 0, 0)
        
        save_btn = QPushButton("Save Supplier")
        save_btn.clicked.connect(self._save_supplier)
        buttons_layout.addWidget(save_btn)
        
        clear_btn = QPushButton("Clear")
        clear_btn.clicked.connect(self._clear_supplier_form)
        buttons_layout.addWidget(clear_btn)
        
        delete_btn = QPushButton("Delete Selected")
        delete_btn.clicked.connect(self._delete_supplier)
        buttons_layout.addWidget(delete_btn)
        
        buttons_layout.addStretch()
        form_layout.addWidget(buttons, 8, 0, 1, 3)
        form_layout.setColumnStretch(1, 1)
        
        main_layout.addWidget(form_group, 1)
        
        # Right: suppliers list
        list_group = QGroupBox("Suppliers")
        list_layout = QVBoxLayout()
        list_group.setLayout(list_layout)
        
        search_row = QWidget()
        search_layout = QHBoxLayout()
        search_row.setLayout(search_layout)
        search_layout.setContentsMargins(0, 0, 0, 0)
        
        search_layout.addWidget(QLabel("Search:"))
        self.supplier_search_entry = QLineEdit()
        self.supplier_search_entry.setPlaceholderText("Filter by company, domain, code, or contact...")
        self.supplier_search_entry.textChanged.connect(self._filter_suppliers)
        search_layout.addWidget(self.supplier_search_entry)
        
        refresh_btn = QPushButton("Refresh")
        refresh_btn.clicked.connect(self._load_suppliers)
        search_layout.addWidget(refresh_btn)
        
        list_layout.addWidget(search_row)
        
        columns = ("Company", "Domain", "Supplier Code", "Contract Code", "Contact")
        self.suppliers_tree = QTreeWidget()
        self.suppliers_tree.setAlternatingRowColors(True)
        self.suppliers_tree.setRootIsDecorated(False)
        self.suppliers_tree.setHeaderLabels(columns)
        self.suppliers_tree.setColumnCount(len(columns))
        self.suppliers_tree.setColumnWidth(0, 160)
        self.suppliers_tree.setColumnWidth(1, 140)
        self.suppliers_tree.setColumnWidth(2, 120)
        self.suppliers_tree.setColumnWidth(3, 120)
        self.suppliers_tree.setColumnWidth(4, 140)
        self.suppliers_tree.setSortingEnabled(True)
        self.suppliers_tree.itemClicked.connect(self._on_supplier_select)
        self.suppliers_tree.itemDoubleClicked.connect(self._on_supplier_select)
        list_layout.addWidget(self.suppliers_tree)
        
        main_layout.addWidget(list_group, 2)
    
    def _extract_company_from_email(self):
        """Extract company name from email address."""
        email = self.extract_email_entry.text().strip()
        
        if not email:
            QMessageBox.warning(self, "No Email", "Please enter an email address")
            return
        
        company = GURIDatabase.extract_company_from_domain(email)
        
        if company:
            self.supplier_company_entry.clear()
            self.supplier_company_entry.setText(company)
            
            # Also extract domain
            if '@' in email:
                domain = email.split('@')[1]
                self.supplier_domain_entry.clear()
                self.supplier_domain_entry.setText(domain)
            
            self.status_bar.showMessage(f"Extracted company: {company}")
        else:
            QMessageBox.warning(self, "Extraction Failed", "Could not extract company from email")
    
    def _save_supplier(self):
        """Save supplier information."""
        if not self.db:
            QMessageBox.warning(self, "No Connection", "Please connect to a database first.")
            return
        
        company_name = self.supplier_company_entry.text().strip()
        
        if not company_name:
            QMessageBox.warning(self, "Validation Error", "Company name is required")
            return
        
        company_domain = self.supplier_domain_entry.text().strip()
        supplier_code = self.supplier_code_entry.text().strip()
        contract_code = self.contract_code_entry.text().strip()
        contact_person = self.contact_person_entry.text().strip()
        contact_email = self.contact_email_entry.text().strip()
        notes = self.supplier_notes_text.toPlainText().strip()
        
        try:
            success = self.db.add_supplier_code(
                company_name, company_domain, supplier_code, contract_code,
                contact_person, contact_email, notes
            )
            
            if success:
                QMessageBox.information(self, "Success", f"Supplier '{company_name}' saved successfully!")
                self._load_suppliers()
                self._clear_supplier_form()
                self.status_bar.showMessage(f"Saved supplier: {company_name}")
            else:
                QMessageBox.critical(self, "Error", "Failed to save supplier")
                
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to save supplier:\n{str(e)}")
            self.logger.error(f"Error saving supplier: {e}")
    
    def _load_suppliers(self):
        """Load all suppliers from database."""
        if not self.db:
            return
        
        try:
            # Clear existing items
            self.suppliers_tree.clear()
            
            # Get all suppliers
            suppliers = self.db.get_all_suppliers()
            
            # Populate tree
            for supplier in suppliers:
                item = QTreeWidgetItem(self.suppliers_tree)
                item.setText(0, supplier['company_name'])
                item.setText(1, supplier['company_domain'] or "")
                item.setText(2, supplier['supplier_code'] or "")
                item.setText(3, supplier['contract_code'] or "")
                item.setText(4, supplier['contact_person'] or "")
                item.setData(0, Qt.ItemDataRole.UserRole, supplier)  # Store full supplier data
            
            self.status_bar.showMessage(f"Loaded {len(suppliers)} suppliers")
            
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to load suppliers:\n{str(e)}")
            self.logger.error(f"Error loading suppliers: {e}")
    
    def _filter_suppliers(self):
        """Filter suppliers based on search text."""
        search_text = self.supplier_search_entry.text().lower()
        
        # For QTreeWidget, we need to filter by hiding items
        # Iterate through all items and show/hide based on match
        for i in range(self.suppliers_tree.topLevelItemCount()):
            item = self.suppliers_tree.topLevelItem(i)
            if item is None:
                continue
            # Search in all columns
            match = any(search_text in item.text(col).lower() for col in range(item.columnCount()))
            
            if match or not search_text:
                item.setHidden(False)
            else:
                item.setHidden(True)
    
    def _on_supplier_select(self, item, column):
        """Handle supplier selection."""
        if not item:
            return
        
        company_name = item.text(0)
        
        # Get full supplier data if stored
        supplier_data = item.data(0, Qt.ItemDataRole.UserRole)
        if supplier_data:
            self._populate_supplier_form(supplier_data)
        elif self.db:
            # Fallback: load from database
            supplier = self.db.get_supplier_by_company(company_name)
            if supplier:
                self._populate_supplier_form(supplier)
    
    def _on_supplier_double_click(self, item, column):
        """Handle double-click on supplier."""
        self._on_supplier_select(item, column)
    
    def _populate_supplier_form(self, supplier):
        """Populate form with supplier data."""
        self.supplier_company_entry.setText(supplier['company_name'])
        self.supplier_domain_entry.setText(supplier['company_domain'] or "")
        self.supplier_code_entry.setText(supplier['supplier_code'] or "")
        self.contract_code_entry.setText(supplier['contract_code'] or "")
        self.contact_person_entry.setText(supplier['contact_person'] or "")
        self.contact_email_entry.setText(supplier['contact_email'] or "")
        self.supplier_notes_text.setPlainText(supplier['notes'] or "")
    
    def _clear_supplier_form(self):
        """Clear all supplier form fields."""
        self.supplier_company_entry.clear()
        self.extract_email_entry.clear()
        self.supplier_domain_entry.clear()
        self.supplier_code_entry.clear()
        self.contract_code_entry.clear()
        self.contact_person_entry.clear()
        self.contact_email_entry.clear()
        self.supplier_notes_text.clear()
        self.suppliers_tree.clearSelection()
    
    def _delete_supplier(self):
        """Delete selected supplier."""
        if not self.db:
            QMessageBox.warning(self, "No Connection", "Please connect to a database first.")
            return
        
        selected_items = self.suppliers_tree.selectedItems()
        if not selected_items:
            QMessageBox.warning(self, "No Selection", "Please select a supplier to delete")
            return
        
        item = selected_items[0]
        company_name = item.text(0)
        
        response = QMessageBox.question(
            self,
            "Confirm Deletion", 
            f"Are you sure you want to delete supplier '{company_name}'?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No
        )
        if response == QMessageBox.StandardButton.Yes:
            try:
                success = self.db.delete_supplier_code(company_name)
                
                if success:
                    QMessageBox.information(self, "Success", f"Supplier '{company_name}' deleted successfully!")
                    self._load_suppliers()
                    self._clear_supplier_form()
                    self.status_bar.showMessage(f"Deleted supplier: {company_name}")
                else:
                    QMessageBox.critical(self, "Error", "Failed to delete supplier")
                    
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to delete supplier:\n{str(e)}")
                self.logger.error(f"Error deleting supplier: {e}")
    
    def _create_deadlines_tab(self):
        """Deadlines extracted from scraped Outlook mail for selected accounts."""
        frame = QWidget()
        self.notebook.addTab(frame, "Deadlines")
        root = QVBoxLayout(frame)
        root.setContentsMargins(12, 12, 12, 12)
        root.setSpacing(10)

        head = QLabel("Deadlines")
        head.setStyleSheet(
            f"font-size: 14px; font-weight: 700; color: {PALETTE['accent']};"
        )
        root.addWidget(head)

        self.deadlines_hint = QLabel(
            "Deadlines found during Outlook scrape for the accounts you select below."
        )
        self.deadlines_hint.setWordWrap(True)
        self.deadlines_hint.setStyleSheet(f"color: {PALETTE['muted']};")
        root.addWidget(self.deadlines_hint)

        toolbar = QHBoxLayout()
        accounts_btn = QPushButton("Select accounts…")
        accounts_btn.setToolTip("Choose which mailboxes contribute deadlines to this tab.")
        accounts_btn.clicked.connect(self._show_deadline_accounts_dialog)
        add_deadline_btn = QPushButton("Add deadline…")
        add_deadline_btn.setToolTip("Create your own deadline (not from email).")
        add_deadline_btn.clicked.connect(lambda: self._show_add_deadline_dialog())
        refresh_btn = QPushButton("Refresh")
        refresh_btn.clicked.connect(self._refresh_deadlines_panel)
        scrape_btn = QPushButton("Scrape Outlook")
        scrape_btn.setToolTip("Run the live scrape; deadline cues are collected into this list.")
        scrape_btn.clicked.connect(lambda: self._start_outlook_scrape(manual=True))
        open_btn = QPushButton("Open in Outlook")
        open_btn.clicked.connect(self._open_selected_deadline_in_outlook)
        identity_btn = QPushButton("My identity…")
        identity_btn.setToolTip(
            "Your names and emails — used to hide deadlines not addressed to you."
        )
        identity_btn.clicked.connect(self._show_identity_dialog)
        toolbar.addWidget(accounts_btn)
        toolbar.addWidget(add_deadline_btn)
        toolbar.addWidget(identity_btn)
        toolbar.addWidget(refresh_btn)
        toolbar.addWidget(scrape_btn)
        toolbar.addStretch(1)
        toolbar.addWidget(open_btn)
        root.addLayout(toolbar)

        self.deadlines_accounts_label = QLabel("")
        self.deadlines_accounts_label.setStyleSheet(f"color: {PALETTE['muted']};")
        self.deadlines_accounts_label.setWordWrap(True)
        root.addWidget(self.deadlines_accounts_label)

        split = QSplitter(Qt.Orientation.Horizontal)
        self.deadlines_split = split

        self.deadlines_tree = QTreeWidget()
        self.deadlines_tree.setAlternatingRowColors(True)
        self.deadlines_tree.setRootIsDecorated(False)
        self.deadlines_tree.setUniformRowHeights(True)
        self.deadlines_tree.setHeaderLabels(
            ("Due", "Account", "From", "Subject", "Signal", "Received", "Status")
        )
        self.deadlines_tree.setColumnCount(7)
        self.deadlines_tree.setColumnWidth(0, 100)
        self.deadlines_tree.setColumnWidth(1, 140)
        self.deadlines_tree.setColumnWidth(2, 150)
        self.deadlines_tree.setColumnWidth(3, 220)
        self.deadlines_tree.setColumnWidth(4, 140)
        self.deadlines_tree.setColumnWidth(5, 130)
        self.deadlines_tree.setColumnWidth(6, 80)
        self.deadlines_tree.itemClicked.connect(self._on_deadline_item_selected)
        self.deadlines_tree.itemDoubleClicked.connect(self._open_deadline_item_in_outlook)
        self.deadlines_tree.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.deadlines_tree.customContextMenuRequested.connect(
            self._show_deadline_context_menu
        )
        split.addWidget(self.deadlines_tree)

        # Right-hand deadline + email preview
        detail = QGroupBox("Deadline detail")
        detail_layout = QVBoxLayout(detail)
        detail_layout.setContentsMargins(8, 8, 8, 8)
        detail_layout.setSpacing(8)

        self.deadlines_preview_due = QLabel("Select a row to see the deadline.")
        due_font = QFont("Segoe UI", 14)
        due_font.setBold(True)
        self.deadlines_preview_due.setFont(due_font)
        self.deadlines_preview_due.setStyleSheet(f"color: {PALETTE['accent']};")
        self.deadlines_preview_due.setWordWrap(True)
        detail_layout.addWidget(self.deadlines_preview_due)

        self.deadlines_preview_signal = QLabel("")
        self.deadlines_preview_signal.setWordWrap(True)
        self.deadlines_preview_signal.setStyleSheet(
            "color: #8a4b08; background: #fff6e5; padding: 6px 8px; border-radius: 4px;"
        )
        detail_layout.addWidget(self.deadlines_preview_signal)

        self.deadlines_preview_meta = QLabel("")
        self.deadlines_preview_meta.setWordWrap(True)
        self.deadlines_preview_meta.setTextFormat(Qt.TextFormat.RichText)
        self.deadlines_preview_meta.setStyleSheet("color: #4a6570; font-size: 11px;")
        detail_layout.addWidget(self.deadlines_preview_meta)

        self.deadlines_preview_body = QTextEdit()
        self.deadlines_preview_body.setReadOnly(True)
        self.deadlines_preview_body.setPlaceholderText(
            "Email body appears here when you select a deadline row."
        )
        detail_layout.addWidget(self.deadlines_preview_body, stretch=1)

        split.addWidget(detail)
        split.setStretchFactor(0, 3)
        split.setStretchFactor(1, 2)
        split.setSizes([700, 420])
        root.addWidget(split, stretch=1)

        QTimer.singleShot(400, self._refresh_deadlines_panel)

    def _create_deadline_timeline_tab(self):
        """Visual timeline of extracted + manual deadlines by due date."""
        frame = QWidget()
        self.notebook.addTab(frame, "Deadline Timeline")
        root = QVBoxLayout(frame)
        root.setContentsMargins(12, 12, 12, 12)
        root.setSpacing(10)

        head = QLabel("Deadline Timeline")
        head.setStyleSheet(
            f"font-size: 14px; font-weight: 700; color: {PALETTE['accent']};"
        )
        root.addWidget(head)

        self.deadline_timeline_hint = QLabel(
            "Extracted and manually added deadlines plotted by due date."
        )
        self.deadline_timeline_hint.setWordWrap(True)
        self.deadline_timeline_hint.setStyleSheet(f"color: {PALETTE['muted']};")
        root.addWidget(self.deadline_timeline_hint)

        toolbar = QHBoxLayout()
        self.deadline_timeline_radio_14 = QRadioButton("Next 14 days")
        self.deadline_timeline_radio_30 = QRadioButton("Next 30 days")
        self.deadline_timeline_radio_30.setChecked(True)
        self.deadline_timeline_mode_group = QButtonGroup(self)
        self.deadline_timeline_mode_group.addButton(self.deadline_timeline_radio_14)
        self.deadline_timeline_mode_group.addButton(self.deadline_timeline_radio_30)
        self.deadline_timeline_radio_14.toggled.connect(self._on_deadline_timeline_mode)
        self.deadline_timeline_radio_30.toggled.connect(self._on_deadline_timeline_mode)
        refresh_btn = QPushButton("Refresh")
        refresh_btn.clicked.connect(self._refresh_deadline_timeline)
        add_btn = QPushButton("Add deadline…")
        add_btn.clicked.connect(lambda: self._show_add_deadline_dialog())
        scrape_btn = QPushButton("Scrape Outlook")
        scrape_btn.clicked.connect(lambda: self._start_outlook_scrape(manual=True))
        toolbar.addWidget(self.deadline_timeline_radio_14)
        toolbar.addWidget(self.deadline_timeline_radio_30)
        toolbar.addWidget(refresh_btn)
        toolbar.addWidget(add_btn)
        toolbar.addWidget(scrape_btn)
        toolbar.addStretch(1)
        root.addLayout(toolbar)

        self.deadline_timeline_count_label = QLabel("")
        self.deadline_timeline_count_label.setStyleSheet(f"color: {PALETTE['muted']};")
        root.addWidget(self.deadline_timeline_count_label)

        canvas_wrap = QGroupBox("Due-date timeline")
        canvas_layout = QVBoxLayout(canvas_wrap)
        canvas_layout.setContentsMargins(8, 8, 8, 8)
        self.deadline_timeline_canvas = DeadlineTimelineCanvas()
        self.deadline_timeline_canvas.cell_clicked.connect(
            self._on_deadline_timeline_cell_clicked
        )
        canvas_layout.addWidget(self.deadline_timeline_canvas)
        root.addWidget(canvas_wrap)

        split = QSplitter(Qt.Orientation.Horizontal)
        self.deadline_timeline_list = QTreeWidget()
        self.deadline_timeline_list.setAlternatingRowColors(True)
        self.deadline_timeline_list.setRootIsDecorated(False)
        self.deadline_timeline_list.setUniformRowHeights(True)
        self.deadline_timeline_list.setHeaderLabels(
            ("Due", "Source", "From", "Subject", "Status")
        )
        self.deadline_timeline_list.setColumnWidth(0, 90)
        self.deadline_timeline_list.setColumnWidth(1, 90)
        self.deadline_timeline_list.setColumnWidth(2, 150)
        self.deadline_timeline_list.setColumnWidth(3, 260)
        self.deadline_timeline_list.setColumnWidth(4, 80)
        self.deadline_timeline_list.itemClicked.connect(
            self._on_deadline_timeline_item_selected
        )
        self.deadline_timeline_list.itemDoubleClicked.connect(
            self._open_deadline_item_in_outlook
        )
        self.deadline_timeline_list.setContextMenuPolicy(
            Qt.ContextMenuPolicy.CustomContextMenu
        )
        self.deadline_timeline_list.customContextMenuRequested.connect(
            self._show_deadline_timeline_context_menu
        )
        split.addWidget(self.deadline_timeline_list)

        detail = QGroupBox("Deadline detail")
        detail_layout = QVBoxLayout(detail)
        self.deadline_timeline_preview_due = QLabel("Click a day circle or list row.")
        due_font = QFont("Segoe UI", 13)
        due_font.setBold(True)
        self.deadline_timeline_preview_due.setFont(due_font)
        self.deadline_timeline_preview_due.setStyleSheet(f"color: {PALETTE['accent']};")
        self.deadline_timeline_preview_due.setWordWrap(True)
        detail_layout.addWidget(self.deadline_timeline_preview_due)
        self.deadline_timeline_preview_meta = QLabel("")
        self.deadline_timeline_preview_meta.setWordWrap(True)
        self.deadline_timeline_preview_meta.setTextFormat(Qt.TextFormat.RichText)
        self.deadline_timeline_preview_meta.setStyleSheet(
            "color: #4a6570; font-size: 11px;"
        )
        detail_layout.addWidget(self.deadline_timeline_preview_meta)
        self.deadline_timeline_preview_body = QTextEdit()
        self.deadline_timeline_preview_body.setReadOnly(True)
        detail_layout.addWidget(self.deadline_timeline_preview_body, stretch=1)
        split.addWidget(detail)
        split.setStretchFactor(0, 3)
        split.setStretchFactor(1, 2)
        split.setSizes([680, 400])
        root.addWidget(split, stretch=1)

        QTimer.singleShot(450, self._refresh_deadline_timeline)

    def _on_deadline_timeline_mode(self, checked: bool) -> None:
        if not checked:
            return
        self._deadline_timeline_days = (
            14 if self.deadline_timeline_radio_14.isChecked() else 30
        )
        self._refresh_deadline_timeline()

    def _save_learning(self) -> None:
        try:
            save_learning(getattr(self, "learning", {}) or {})
        except Exception as exc:
            self.logger.warning("Failed saving learning store: %s", exc)

    def _important_senders_list(self) -> List[str]:
        learning = getattr(self, "learning", None) or {}
        return list(learning.get("important_senders") or [])

    def _learning_deadline_cues(self) -> List[str]:
        return cue_patterns_from_learning(getattr(self, "learning", {}) or {})

    def _seed_identity_defaults(self) -> None:
        """Seed my_emails / my_names from known accounts if empty."""
        extras = [MY_ALINIANT_EMAIL, NON_EMAIL_SENDER]
        try:
            for cfg in load_aes_account_configs() or []:
                smtp = str(getattr(cfg, "smtp", "") or "").strip()
                if smtp:
                    extras.append(smtp)
        except Exception:
            pass
        names = ["Julian", "Julian Garrett"]
        before = json.dumps(
            {
                "e": (getattr(self, "learning", {}) or {}).get("my_emails"),
                "n": (getattr(self, "learning", {}) or {}).get("my_names"),
            },
            sort_keys=True,
        )
        self.learning = ensure_identity_defaults(
            getattr(self, "learning", {}) or {},
            extra_emails=extras,
            extra_names=names,
        )
        after = json.dumps(
            {
                "e": self.learning.get("my_emails"),
                "n": self.learning.get("my_names"),
            },
            sort_keys=True,
        )
        if before != after:
            self._save_learning()

    def _deadline_account_match_keys(self, mail: Dict[str, Any]) -> Set[str]:
        keys: Set[str] = set()
        store_id = str((mail or {}).get("store_id") or "").strip()
        smtp = str((mail or {}).get("account_smtp") or "").strip()
        if store_id:
            keys.add(store_id)
        if smtp:
            keys.add(smtp)
            keys.add(smtp.lower())
        return keys

    def _mail_in_deadline_accounts(self, mail: Dict[str, Any]) -> bool:
        # File-location scrape hits are not Outlook accounts — always eligible
        if str((mail or {}).get("store_id") or "").lower() == "file":
            return True
        if str((mail or {}).get("account_smtp") or "").lower() == "file":
            return True
        selected = {
            str(k).strip()
            for k in (getattr(self, "deadline_account_keys", None) or [])
            if str(k).strip()
        }
        if not selected:
            # No filter yet — show all deadline hits so the tab isn't empty
            return True
        selected_l = {k.lower() for k in selected}
        mail_keys = self._deadline_account_match_keys(mail)
        return any(k in selected or k.lower() in selected_l for k in mail_keys)

    def _normalize_deadline_subject(self, subject: str) -> str:
        text = str(subject or "").strip().lower()
        # Strip repeated Re:/Fw:/Fwd: so thread copies collapse together
        while True:
            cleaned = re.sub(r"^(?:re|fw|fwd)\s*:\s*", "", text, flags=re.IGNORECASE)
            if cleaned == text:
                break
            text = cleaned.strip()
        return text

    def _deadline_dedupe_key(self, mail: Dict[str, Any]) -> str:
        entry_id = str((mail or {}).get("entry_id") or "").strip()
        if entry_id.startswith("manual:"):
            return entry_id
        if entry_id:
            return f"id:{entry_id}"
        account = str(
            (mail or {}).get("account_smtp") or (mail or {}).get("store_id") or ""
        ).lower()
        sender = str((mail or {}).get("sender") or "").lower()
        subject = self._normalize_deadline_subject(str((mail or {}).get("subject") or ""))
        return f"thread:{account}|{sender}|{subject}"

    def _manual_deadline_to_row(self, entry: Dict[str, Any]) -> Dict[str, Any]:
        """Convert a saved manual deadline into the shared {det, deadline} row shape."""
        mid = str(entry.get("id") or "")
        title = str(entry.get("title") or "Untitled")
        due = str(entry.get("due") or "")[:10]
        notes = str(entry.get("notes") or "")
        created = str(entry.get("created") or "")
        status = str(entry.get("status") or "—")
        linked_sender = str(entry.get("linked_sender") or "").strip()
        linked_subject = str(entry.get("linked_subject") or "").strip()
        deadline = {
            "cue": "manual",
            "label": title,
            "due": due,
            "snippet": notes or title,
        }
        det = {
            "mail": {
                "sender": linked_sender or "Me (manual)",
                "subject": linked_subject or title,
                "body_preview": notes,
                "account_smtp": "manual",
                "store_id": "manual",
                "entry_id": f"manual:{mid}",
                "received": created,
            },
            "actions": ["deadline"],
            "reasons": ["Manual deadline"],
            "importance": "high",
            "score": 60,
            "status": status,
            "deadlines": [deadline],
            "manual_id": mid,
            "mail_key": str(entry.get("mail_key") or ""),
        }
        return {"det": det, "deadline": deadline, "manual": True}

    def _collect_deadline_rows(self) -> List[Dict[str, Any]]:
        """Build deduped deadline rows from scrape + manual entries + learning."""
        items = self._feedback_items()
        learning = getattr(self, "learning", {}) or {}
        extra_cues = self._learning_deadline_cues()
        rows: List[Dict[str, Any]] = []
        for det in items:
            mail = det.get("mail") or {}
            if not self._mail_in_deadline_accounts(mail):
                continue
            key = mail_feedback_key(mail, det)
            if is_deadline_rejected(learning, key):
                continue
            if det.get("rule_skip_deadline"):
                continue
            confirmed = (learning.get("confirmed_deadline_mails") or {}).get(key)
            if not should_include_deadline_mail(
                mail,
                learning,
                confirmed=bool(confirmed),
                manual=False,
            ):
                continue
            if not det.get("deadlines") and "deadline" in [
                str(a).lower() for a in (det.get("actions") or [])
            ]:
                blob = f"{mail.get('subject') or ''}\n{mail.get('body_preview') or ''}"
                try:
                    det = dict(det)
                    det["deadlines"] = extract_deadlines(blob, extra_cues=extra_cues)
                except Exception:
                    pass
            elif not det.get("deadlines") and extra_cues:
                blob = f"{mail.get('subject') or ''}\n{mail.get('body_preview') or ''}"
                try:
                    learned = extract_deadlines(blob, extra_cues=extra_cues)
                    if learned:
                        det = dict(det)
                        det["deadlines"] = learned
                        actions = list(det.get("actions") or [])
                        if "deadline" not in [str(a).lower() for a in actions]:
                            actions.append("deadline")
                            det["actions"] = actions
                except Exception:
                    pass
            # Confirmed-by-user deadlines that heuristics missed
            if confirmed and not det.get("deadlines"):
                det = dict(det)
                due = str(confirmed.get("due") or "")[:10] or None
                title = str(confirmed.get("title") or mail.get("subject") or "Confirmed deadline")
                det["deadlines"] = [
                    {
                        "cue": "confirmed",
                        "label": title,
                        "due": due,
                        "snippet": title,
                    }
                ]
                actions = list(det.get("actions") or [])
                if "deadline" not in [str(a).lower() for a in actions]:
                    actions.append("deadline")
                    det["actions"] = actions
            rows.extend(self._item_deadline_rows(det))

        # File-location scrape: include deadline cues found in readable files
        for det in getattr(self, "file_scrape_items", None) or []:
            if not isinstance(det, dict):
                continue
            if not (det.get("deadlines") or []):
                continue
            mail = det.get("mail") or {}
            key = mail_feedback_key(mail, det)
            if is_deadline_rejected(learning, key):
                continue
            rows.extend(self._item_deadline_rows(det))

        # Always include user-created deadlines
        for entry in getattr(self, "manual_deadlines", None) or []:
            rows.append(self._manual_deadline_to_row(entry))

        by_key: Dict[str, Dict[str, Any]] = {}
        for row in rows:
            mail = (row.get("det") or {}).get("mail") or {}
            key = self._deadline_dedupe_key(mail)
            prev = by_key.get(key)
            if prev is None:
                by_key[key] = row
                continue
            prev_rec = str(((prev.get("det") or {}).get("mail") or {}).get("received") or "")
            cur_rec = str(mail.get("received") or "")
            if cur_rec >= prev_rec:
                by_key[key] = row
        rows = list(by_key.values())

        def sort_key(row: Dict[str, Any]):
            dl = row.get("deadline") or {}
            due = str(dl.get("due") or "")
            received = str((row.get("det") or {}).get("mail", {}).get("received") or "")
            return (0 if due else 1, due or "9999", received)

        rows.sort(key=sort_key)
        self._deadline_rows_cache = rows
        return rows

    def _show_add_deadline_dialog(self, preset_date=None) -> None:
        """Popup to create a personal deadline (Calendar / Deadlines / Timeline)."""
        dialog = QDialog(self)
        dialog.setWindowTitle("Add deadline")
        dialog.resize(480, 420)
        layout = QVBoxLayout(dialog)

        intro = QLabel(
            "Add a deadline — it appears on Deadlines, Deadline Timeline, and Calendar."
        )
        intro.setWordWrap(True)
        intro.setStyleSheet(f"color: {PALETTE['muted']};")
        layout.addWidget(intro)

        form = QFormLayout()
        title_edit = QLineEdit()
        title_edit.setPlaceholderText("What is due?")
        form.addRow("Title", title_edit)

        due_edit = QDateEdit()
        due_edit.setCalendarPopup(True)
        due_edit.setDisplayFormat("yyyy-MM-dd")
        if preset_date is not None:
            due_edit.setDate(QDate(preset_date.year, preset_date.month, preset_date.day))
        else:
            due_edit.setDate(QDate.currentDate())
        form.addRow("Due date", due_edit)

        notes_edit = QTextEdit()
        notes_edit.setPlaceholderText("Optional notes…")
        notes_edit.setMaximumHeight(80)
        form.addRow("Notes", notes_edit)

        candidates = candidate_deadline_emails(self._feedback_items(), limit=15)
        link_combo = QComboBox()
        link_combo.addItem("(none — not from an email)", None)
        for cand in candidates:
            link_combo.addItem(cand["label"], cand)
        form.addRow("From email?", link_combo)
        layout.addLayout(form)

        link_hint = QLabel(
            "Linking teaches GURI which messages carry deadlines for you."
        )
        link_hint.setWordWrap(True)
        link_hint.setStyleSheet(f"color: {PALETTE['muted']}; font-size: 11px;")
        layout.addWidget(link_hint)

        save_btn = QPushButton("Save")
        save_btn.setDefault(True)
        cancel_btn = QPushButton("Cancel")
        cancel_btn.clicked.connect(dialog.reject)
        add_dialog_button_row(layout, affirmative=save_btn, cancel=cancel_btn)

        def on_save() -> None:
            title = title_edit.text().strip()
            if not title:
                QMessageBox.warning(dialog, "Add deadline", "Please enter a title.")
                title_edit.setFocus()
                return
            due = due_edit.date().toString("yyyy-MM-dd")
            notes = notes_edit.toPlainText().strip()
            cand = link_combo.currentData()
            mail_key = ""
            sender = ""
            subject = ""
            if isinstance(cand, dict):
                # Confirm intent when an email is selected
                reply = QMessageBox.question(
                    dialog,
                    "Confirm email link",
                    f"Is this the email you are extracting this deadline from?\n\n"
                    f"From: {cand.get('sender') or '—'}\n"
                    f"Subject: {cand.get('subject') or '—'}",
                    QMessageBox.StandardButton.Yes
                    | QMessageBox.StandardButton.No
                    | QMessageBox.StandardButton.Cancel,
                    QMessageBox.StandardButton.Yes,
                )
                if reply == QMessageBox.StandardButton.Cancel:
                    return
                if reply == QMessageBox.StandardButton.Yes:
                    mail_key = str(cand.get("mail_key") or "")
                    sender = str(cand.get("sender") or "")
                    subject = str(cand.get("subject") or "")
            try:
                before_ids = {
                    str(x.get("id") or "")
                    for x in (getattr(self, "manual_deadlines", []) or [])
                }
                self.manual_deadlines = add_manual_deadline(
                    getattr(self, "manual_deadlines", []) or [],
                    title=title,
                    due=due,
                    notes=notes,
                    mail_key=mail_key,
                    sender=sender,
                    subject=subject,
                )
            except Exception as exc:
                QMessageBox.critical(dialog, "Add deadline", f"Could not save:\n{exc}")
                return

            # Learning: confirm mail, store cue phrase, optional important sender
            new_entry = None
            for entry in self.manual_deadlines or []:
                eid = str(entry.get("id") or "")
                if eid and eid not in before_ids:
                    new_entry = entry
                    break
            if mail_key:
                self.learning = confirm_deadline_mail(
                    self.learning,
                    mail_key=mail_key,
                    due=due,
                    title=title,
                    sender=sender,
                    subject=subject,
                )
                if new_entry:
                    self.learning = link_manual_deadline(
                        self.learning,
                        manual_id=str(new_entry.get("id") or ""),
                        mail_key=mail_key,
                        sender=sender,
                        subject=subject,
                    )
                # Learn a phrase from the title if it looks like deadline language
                self.learning = learn_deadline_cue(self.learning, title)
                if sender and should_ask_important_sender(self.learning, sender):
                    ask = QMessageBox.question(
                        dialog,
                        "Important sender?",
                        f"Did you mean to treat this sender as important?\n\n{sender}",
                        QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                        QMessageBox.StandardButton.No,
                    )
                    if ask == QMessageBox.StandardButton.Yes:
                        self.learning = remember_important_sender(self.learning, sender)
                        domain = domain_from_sender(sender)
                        if domain and domain not in self.important_domains:
                            add_dom = QMessageBox.question(
                                dialog,
                                "Important Emails domain?",
                                f"Also add @{domain} to Important Emails domains?",
                                QMessageBox.StandardButton.Yes
                                | QMessageBox.StandardButton.No,
                                QMessageBox.StandardButton.Yes,
                            )
                            if add_dom == QMessageBox.StandardButton.Yes:
                                self.important_domains.append(domain)
                                self._save_important_domains()
                                if hasattr(self, "domains_listbox"):
                                    self._refresh_domains_list()
                    else:
                        self.learning = mark_asked_important_sender(
                            self.learning, sender
                        )
                self._save_learning()
            elif candidates and not (
                (getattr(self, "learning", {}) or {}).get("suppress_email_link_prompt")
            ):
                top = candidates[0]
                ask = QMessageBox.question(
                    self,
                    "Did you mean this email?",
                    f"You did not link an email. Is this the message the deadline came from?\n\n"
                    f"From: {top.get('sender') or '—'}\n"
                    f"Subject: {top.get('subject') or '—'}",
                    QMessageBox.StandardButton.Yes
                    | QMessageBox.StandardButton.No
                    | QMessageBox.StandardButton.Ignore,
                    QMessageBox.StandardButton.No,
                )
                if ask == QMessageBox.StandardButton.Yes and new_entry:
                    mail_key = str(top.get("mail_key") or "")
                    sender = str(top.get("sender") or "")
                    subject = str(top.get("subject") or "")
                    # Patch the saved manual entry
                    updated = []
                    for entry in self.manual_deadlines or []:
                        e = dict(entry)
                        if str(e.get("id") or "") == str(new_entry.get("id") or ""):
                            e["mail_key"] = mail_key
                            e["linked_sender"] = sender
                            e["linked_subject"] = subject
                        updated.append(e)
                    self.manual_deadlines = updated
                    try:
                        save_manual_deadlines(updated)
                    except Exception:
                        pass
                    self.learning = confirm_deadline_mail(
                        self.learning,
                        mail_key=mail_key,
                        due=due,
                        title=title,
                        sender=sender,
                        subject=subject,
                    )
                    self.learning = link_manual_deadline(
                        self.learning,
                        manual_id=str(new_entry.get("id") or ""),
                        mail_key=mail_key,
                        sender=sender,
                        subject=subject,
                    )
                    self.learning = learn_deadline_cue(self.learning, title)
                    self._save_learning()
                elif ask == QMessageBox.StandardButton.Ignore:
                    self.learning = dict(self.learning or {})
                    self.learning["suppress_email_link_prompt"] = True
                    self._save_learning()

            dialog.accept()
            self._refresh_deadlines_panel()
            self.status_bar.showMessage(f"Added deadline: {title} ({due})")

        save_btn.clicked.connect(on_save)
        title_edit.setFocus()
        dialog.exec()

    def _add_deadline_for_selected_day(self) -> None:
        day = getattr(self, "_calendar_selected_date", None)
        self._show_add_deadline_dialog(preset_date=day)

    def _delete_selected_manual_deadline(self, item) -> None:
        row = self._deadline_row_payload(item)
        det = row.get("det") or {}
        mid = str(det.get("manual_id") or "")
        if not mid:
            mail = det.get("mail") or {}
            entry_id = str(mail.get("entry_id") or "")
            if entry_id.startswith("manual:"):
                mid = entry_id.split(":", 1)[-1]
        if not mid:
            QMessageBox.information(
                self, "Delete deadline", "Only manually added deadlines can be deleted here."
            )
            return
        title = str((det.get("mail") or {}).get("subject") or "this deadline")
        reply = QMessageBox.question(
            self,
            "Delete deadline",
            f"Delete manual deadline “{title}”?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if reply != QMessageBox.StandardButton.Yes:
            return
        self.manual_deadlines = delete_manual_deadline(
            getattr(self, "manual_deadlines", []) or [], mid
        )
        self._refresh_deadlines_panel()
        self.status_bar.showMessage(f"Deleted deadline: {title}")

    def _item_deadline_rows(self, det: Dict[str, Any]) -> List[Dict[str, Any]]:
        """One display row per email (deadline cues merged, not exploded)."""
        mail = det.get("mail") or {}
        actions = [str(a).lower() for a in (det.get("actions") or [])]
        deadlines = list(det.get("deadlines") or [])
        extra_cues = self._learning_deadline_cues()
        if not deadlines and "deadline" in actions:
            blob = f"{mail.get('subject') or ''}\n{mail.get('body_preview') or ''}"
            try:
                deadlines = extract_deadlines(blob, extra_cues=extra_cues)
            except Exception:
                deadlines = []
        if not deadlines and "deadline" not in actions:
            return []
        if not deadlines:
            deadlines = [
                {
                    "cue": "deadline",
                    "label": "Deadline / urgency",
                    "due": None,
                    "snippet": str(mail.get("subject") or "")[:120],
                }
            ]
        merged = merge_deadline_hits(deadlines)
        if not merged:
            return []
        return [{"det": det, "deadline": merged}]

    def _deadline_due_date(self, row: Dict[str, Any]):
        """Parse ISO due date from a deadline row, or None."""
        from datetime import date as date_cls

        dl = row.get("deadline") or {}
        due = str(dl.get("due") or "").strip()
        if not due:
            return None
        try:
            return date_cls.fromisoformat(due[:10])
        except ValueError:
            return None

    def _refresh_deadlines_panel(self) -> None:
        if not hasattr(self, "deadlines_tree"):
            return
        try:
            self.deadlines_tree.clear()
            keys = list(getattr(self, "deadline_account_keys", None) or [])
            if keys:
                self.deadlines_accounts_label.setText(
                    "Accounts: " + ", ".join(keys[:8]) + ("…" if len(keys) > 8 else "")
                )
            else:
                self.deadlines_accounts_label.setText(
                    "Accounts: all scraped mailboxes (use Select accounts… to narrow)."
                )

            rows = self._collect_deadline_rows()

            for row in rows[:200]:
                det = row.get("det") or {}
                dl = row.get("deadline") or {}
                mail = det.get("mail") or {}
                item = QTreeWidgetItem(self.deadlines_tree)
                due = str(dl.get("due") or dl.get("label") or "—")
                item.setText(0, due)
                account = str(
                    mail.get("account_smtp") or mail.get("store_id") or ""
                )[:60]
                item.setText(1, account)
                item.setText(2, str(mail.get("sender") or "")[:80])
                subject = str(mail.get("subject") or "")
                item.setText(3, subject[:70] + ("..." if len(subject) > 70 else ""))
                cue = str(dl.get("cue") or "")
                snippet = str(dl.get("snippet") or "")
                # Prefer cue over a snippet that just repeats the subject
                subj_norm = self._normalize_deadline_subject(subject)
                snip_norm = self._normalize_deadline_subject(snippet)
                if snip_norm and snip_norm != subj_norm and subj_norm not in snip_norm:
                    signal = snippet
                else:
                    signal = cue or snippet or "—"
                item.setText(4, signal[:80])
                item.setText(5, str(mail.get("received") or ""))
                item.setText(6, str(det.get("status") or "—"))
                item.setData(0, Qt.ItemDataRole.UserRole, row)

            n = len(rows)
            self.deadlines_hint.setText(
                f"{n} unique deadline email(s) from scrape"
                + (
                    f" across {len(keys)} selected account(s)."
                    if keys
                    else " (all accounts — narrow with Select accounts…)."
                )
            )
            self._refresh_calendar_panel(rows=rows)
            self._refresh_deadline_timeline(rows=rows)
        except Exception as e:
            self.logger.error("Error refreshing deadlines panel: %s", e)

    def _refresh_deadline_timeline(
        self, rows: Optional[List[Dict[str, Any]]] = None
    ) -> None:
        """Populate Deadline Timeline canvas + list from due-dated rows."""
        if not hasattr(self, "deadline_timeline_canvas"):
            return
        try:
            if rows is None:
                rows = self._collect_deadline_rows()
            n_days = int(getattr(self, "_deadline_timeline_days", 30) or 30)
            start = datetime.now().date()
            self._deadline_timeline_start = start
            end = start + timedelta(days=n_days - 1)
            self.deadline_timeline_canvas.set_range(start, n_days)

            day_counts: Dict[int, Dict[str, int]] = {}
            visible: List[Dict[str, Any]] = []
            overdue = 0
            for row in rows:
                due = self._deadline_due_date(row)
                if due is None:
                    continue
                is_manual = bool(
                    row.get("manual") or (row.get("det") or {}).get("manual_id")
                )
                category = DL_CAT_MANUAL if is_manual else DL_CAT_EXTRACTED
                if due < start:
                    overdue += 1
                    continue
                if due > end:
                    continue
                bucket = (due - start).days
                counts = day_counts.setdefault(
                    bucket, {c: 0 for c in DEADLINE_TIMELINE_CATEGORIES}
                )
                counts[category] += 1
                visible.append({**row, "dl_category": category, "day_bucket": bucket})

            self._deadline_timeline_items = visible
            self.deadline_timeline_canvas.update_day_data(day_counts)

            if hasattr(self, "deadline_timeline_list"):
                self.deadline_timeline_list.clear()
                for row in visible:
                    det = row.get("det") or {}
                    dl = row.get("deadline") or {}
                    mail = det.get("mail") or {}
                    item = QTreeWidgetItem(self.deadline_timeline_list)
                    item.setText(0, str(dl.get("due") or "—"))
                    item.setText(
                        1,
                        "Manual"
                        if row.get("dl_category") == DL_CAT_MANUAL
                        else "Extracted",
                    )
                    item.setText(2, str(mail.get("sender") or "")[:80])
                    subject = str(mail.get("subject") or "")
                    item.setText(
                        3, subject[:70] + ("..." if len(subject) > 70 else "")
                    )
                    item.setText(4, str(det.get("status") or "—"))
                    item.setData(0, Qt.ItemDataRole.UserRole, row)

            extracted_n = sum(
                1 for r in visible if r.get("dl_category") == DL_CAT_EXTRACTED
            )
            manual_n = sum(
                1 for r in visible if r.get("dl_category") == DL_CAT_MANUAL
            )
            if hasattr(self, "deadline_timeline_count_label"):
                extra = f"  ·  {overdue} overdue (before today)" if overdue else ""
                self.deadline_timeline_count_label.setText(
                    f"{len(visible)} due in window  ·  Extracted {extracted_n}  ·  "
                    f"Manual {manual_n}{extra}"
                )
            if hasattr(self, "deadline_timeline_hint"):
                self.deadline_timeline_hint.setText(
                    f"{start.isoformat()} → {end.isoformat()} — "
                    "click a circle to filter the list; right-click to teach corrections."
                )
        except Exception as e:
            self.logger.error("Error refreshing deadline timeline: %s", e)

    def _deadline_timeline_matches(
        self, day_bucket: int, category: str
    ) -> List[Dict[str, Any]]:
        return [
            r
            for r in (getattr(self, "_deadline_timeline_items", None) or [])
            if r.get("day_bucket") == day_bucket
            and r.get("dl_category") == category
        ]

    def _on_deadline_timeline_cell_clicked(
        self, day_bucket: int, category: str, _anchor
    ) -> None:
        if day_bucket < 0 or not category:
            return
        matches = self._deadline_timeline_matches(day_bucket, category)
        if not hasattr(self, "deadline_timeline_list"):
            return
        self.deadline_timeline_list.clear()
        for row in matches:
            det = row.get("det") or {}
            dl = row.get("deadline") or {}
            mail = det.get("mail") or {}
            item = QTreeWidgetItem(self.deadline_timeline_list)
            item.setText(0, str(dl.get("due") or "—"))
            item.setText(
                1, "Manual" if row.get("dl_category") == DL_CAT_MANUAL else "Extracted"
            )
            item.setText(2, str(mail.get("sender") or "")[:80])
            subject = str(mail.get("subject") or "")
            item.setText(3, subject[:70] + ("..." if len(subject) > 70 else ""))
            item.setText(4, str(det.get("status") or "—"))
            item.setData(0, Qt.ItemDataRole.UserRole, row)
        if matches:
            first = self.deadline_timeline_list.topLevelItem(0)
            if first is not None:
                self.deadline_timeline_list.setCurrentItem(first)
                self._on_deadline_timeline_item_selected(first)
            day = self._deadline_timeline_start + timedelta(days=day_bucket)
            label = DEADLINE_TIMELINE_LABELS.get(category, category)
            self.status_bar.showMessage(
                f"{len(matches)} deadline(s) on {day.isoformat()} ({label})"
            )

    def _on_deadline_timeline_item_selected(self, item, _column=None) -> None:
        if item is None:
            return
        row = item.data(0, Qt.ItemDataRole.UserRole) or {}
        if not isinstance(row, dict):
            return
        det = row.get("det") or {}
        deadline = row.get("deadline") or {}
        mail = det.get("mail") or {}
        due = str(deadline.get("due") or deadline.get("label") or "—")
        if hasattr(self, "deadline_timeline_preview_due"):
            src = (
                "Manual"
                if row.get("dl_category") == DL_CAT_MANUAL or row.get("manual")
                else "Extracted"
            )
            self.deadline_timeline_preview_due.setText(f"Due: {due}  ({src})")
        if hasattr(self, "deadline_timeline_preview_meta"):
            self.deadline_timeline_preview_meta.setText(
                f"<b>{mail.get('subject') or ''}</b><br>"
                f"From: {mail.get('sender') or ''}<br>"
                f"Signal: {deadline.get('cue') or '—'}"
            )
        body = str(mail.get("body_preview") or "")
        subject = str(mail.get("subject") or "")
        if hasattr(self, "deadline_timeline_preview_body"):
            self.deadline_timeline_preview_body.setHtml(
                highlight_action_html(f"Subject: {subject}\n\n{body}")
            )
        self._show_email_preview(det if isinstance(det, dict) else {})

    def _show_deadline_timeline_context_menu(self, position) -> None:
        item = self.deadline_timeline_list.itemAt(position)
        if not item:
            return
        self.deadline_timeline_list.setCurrentItem(item)
        row = self._deadline_row_payload(item)
        is_manual = bool(row.get("manual") or (row.get("det") or {}).get("manual_id"))
        menu = QMenu(self)
        if is_manual:
            menu.addAction(
                "Delete manual deadline",
                lambda: self._delete_selected_manual_deadline(item),
            )
        else:
            menu.addAction(
                "Open in Outlook", lambda: self._open_deadline_item_in_outlook(item)
            )
            menu.addAction(
                "Not a deadline (teach GURI)",
                lambda: self._reject_deadline_learning(item),
            )
            menu.addAction(
                "Not for me (teach GURI)",
                lambda: self._reject_deadline_not_for_me(item),
            )
            menu.addAction(
                "Confirm this deadline",
                lambda: self._confirm_deadline_learning(item),
            )
        menu.addAction("View Details", lambda: self._view_deadline_details(item))
        menu.exec(self.deadline_timeline_list.mapToGlobal(position))

    def _show_identity_dialog(self) -> None:
        """Edit names/emails used for 'is this deadline for me?' filtering."""
        learning = getattr(self, "learning", {}) or {}
        dialog = QDialog(self)
        dialog.setWindowTitle("My identity")
        dialog.resize(480, 360)
        layout = QVBoxLayout(dialog)

        intro = QLabel(
            "GURI uses these names and emails to hide deadline items that are "
            "clearly addressed to someone else (e.g. “Hi Lee,”)."
        )
        intro.setWordWrap(True)
        intro.setStyleSheet(f"color: {PALETTE['muted']};")
        layout.addWidget(intro)

        form = QFormLayout()
        names_edit = QTextEdit()
        names_edit.setPlaceholderText("One name per line — e.g. Julian\nJulian Garrett")
        names_edit.setPlainText("\n".join(learning.get("my_names") or []))
        names_edit.setMaximumHeight(90)
        form.addRow("My names", names_edit)

        emails_edit = QTextEdit()
        emails_edit.setPlaceholderText("One email per line")
        emails_edit.setPlainText("\n".join(learning.get("my_emails") or []))
        emails_edit.setMaximumHeight(110)
        form.addRow("My emails", emails_edit)

        filter_cb = QCheckBox("Hide deadlines not addressed to me")
        filter_cb.setChecked(bool(learning.get("filter_deadlines_not_for_me", True)))
        form.addRow("", filter_cb)
        layout.addLayout(form)

        others = learning.get("other_names") or []
        if others:
            other_lbl = QLabel(
                "Learned as not you: " + ", ".join(str(n) for n in others[:12])
                + ("…" if len(others) > 12 else "")
            )
            other_lbl.setWordWrap(True)
            other_lbl.setStyleSheet(f"color: {PALETTE['muted']}; font-size: 11px;")
            layout.addWidget(other_lbl)

        save_btn = QPushButton("Save")
        save_btn.setDefault(True)
        cancel_btn = QPushButton("Cancel")
        cancel_btn.clicked.connect(dialog.reject)
        add_dialog_button_row(layout, affirmative=save_btn, cancel=cancel_btn)

        def on_save() -> None:
            names = [
                ln.strip()
                for ln in names_edit.toPlainText().splitlines()
                if ln.strip()
            ]
            emails = [
                ln.strip()
                for ln in emails_edit.toPlainText().splitlines()
                if ln.strip()
            ]
            self.learning = set_my_identity(
                self.learning,
                names=names,
                emails=emails,
                filter_not_for_me=filter_cb.isChecked(),
            )
            self._save_learning()
            dialog.accept()
            self._refresh_deadlines_panel()
            self.status_bar.showMessage(
                f"Identity saved — {len(names)} name(s), {len(emails)} email(s)"
            )

        save_btn.clicked.connect(on_save)
        dialog.exec()

    def _reject_deadline_learning(self, item) -> None:
        row = self._deadline_row_payload(item)
        det = row.get("det") or {}
        mail = det.get("mail") or {}
        key = mail_feedback_key(mail, det)
        if not key:
            return
        greeting = extract_greeting_name(str(mail.get("body_preview") or ""))
        extra = f'\n\nDetected greeting: “Hi {greeting},”' if greeting else ""
        reply = QMessageBox.question(
            self,
            "Not a deadline",
            "Mark this email as not carrying a deadline?\n"
            "GURI will hide it from deadline lists going forward."
            + extra,
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.Yes,
        )
        if reply != QMessageBox.StandardButton.Yes:
            return
        self.learning = reject_deadline_mail(self.learning, key)
        if greeting:
            self.learning = remember_other_name(self.learning, greeting)
        self._save_learning()
        self._refresh_deadlines_panel()
        self.status_bar.showMessage("Learned: this email is not a deadline")

    def _reject_deadline_not_for_me(self, item) -> None:
        row = self._deadline_row_payload(item)
        det = row.get("det") or {}
        mail = det.get("mail") or {}
        key = mail_feedback_key(mail, det)
        if not key:
            return
        greeting = extract_greeting_name(str(mail.get("body_preview") or ""))
        verdict = mail_is_for_me(mail, getattr(self, "learning", {}) or {})
        detail = str(verdict.get("reason") or "")
        if greeting:
            detail = f"Greeting looks addressed to {greeting}." + (
                f" ({detail})" if detail else ""
            )
        reply = QMessageBox.question(
            self,
            "Not for me",
            "Hide this deadline because it is not addressed to you?\n\n"
            + (detail or "GURI will remember this and similar greetings."),
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.Yes,
        )
        if reply != QMessageBox.StandardButton.Yes:
            return
        self.learning = reject_deadline_mail(self.learning, key)
        if greeting:
            self.learning = remember_other_name(self.learning, greeting)
        self.learning = set_my_identity(self.learning, filter_not_for_me=True)
        self._save_learning()
        self._refresh_deadlines_panel()
        self.status_bar.showMessage(
            "Learned: not for you"
            + (f" (greeting {greeting})" if greeting else "")
        )

    def _confirm_deadline_learning(self, item) -> None:
        row = self._deadline_row_payload(item)
        det = row.get("det") or {}
        dl = row.get("deadline") or {}
        mail = det.get("mail") or {}
        key = mail_feedback_key(mail, det)
        if not key:
            return
        subject = str(mail.get("subject") or "")
        sender = str(mail.get("sender") or "")
        due = str(dl.get("due") or "")[:10]
        reply = QMessageBox.question(
            self,
            "Confirm deadline",
            f"Confirm this email as a deadline source?\n\n"
            f"From: {sender}\nSubject: {subject}\nDue: {due or '—'}",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.Yes,
        )
        if reply != QMessageBox.StandardButton.Yes:
            return
        self.learning = confirm_deadline_mail(
            self.learning,
            mail_key=key,
            due=due,
            title=subject,
            sender=sender,
            subject=subject,
        )
        if dl.get("cue"):
            self.learning = learn_deadline_cue(self.learning, str(dl.get("cue")))
        if sender and should_ask_important_sender(self.learning, sender):
            ask = QMessageBox.question(
                self,
                "Important sender?",
                f"Also treat {sender} as an important sender?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No,
            )
            if ask == QMessageBox.StandardButton.Yes:
                self.learning = remember_important_sender(self.learning, sender)
            else:
                self.learning = mark_asked_important_sender(self.learning, sender)
        self._save_learning()
        self._refresh_deadlines_panel()
        self.status_bar.showMessage("Learned: confirmed deadline email")

    def _create_calendar_tab(self):
        """30-day calendar of scraped deadlines with day list + detail."""
        frame = QWidget()
        self.notebook.addTab(frame, "Calendar")
        root = QVBoxLayout(frame)
        root.setContentsMargins(12, 12, 12, 12)
        root.setSpacing(10)

        head = QLabel("Deadline calendar")
        head.setStyleSheet(
            f"font-size: 14px; font-weight: 700; color: {PALETTE['accent']};"
        )
        root.addWidget(head)

        self.calendar_hint = QLabel(
            "Next 30 days — scraped deadlines with a due date, plus any you add yourself."
        )
        self.calendar_hint.setWordWrap(True)
        self.calendar_hint.setStyleSheet(f"color: {PALETTE['muted']};")
        root.addWidget(self.calendar_hint)

        toolbar = QHBoxLayout()
        refresh_btn = QPushButton("Refresh")
        refresh_btn.clicked.connect(lambda: self._refresh_calendar_panel())
        add_deadline_btn = QPushButton("Add deadline…")
        add_deadline_btn.setToolTip("Create your own deadline for a chosen date.")
        add_deadline_btn.clicked.connect(lambda: self._show_add_deadline_dialog())
        scrape_btn = QPushButton("Scrape Outlook")
        scrape_btn.clicked.connect(lambda: self._start_outlook_scrape(manual=True))
        accounts_btn = QPushButton("Select accounts…")
        accounts_btn.clicked.connect(self._show_deadline_accounts_dialog)
        identity_btn = QPushButton("My identity…")
        identity_btn.setToolTip(
            "Your names and emails — used to hide deadlines not addressed to you."
        )
        identity_btn.clicked.connect(self._show_identity_dialog)
        toolbar.addWidget(refresh_btn)
        toolbar.addWidget(add_deadline_btn)
        toolbar.addWidget(scrape_btn)
        toolbar.addWidget(accounts_btn)
        toolbar.addWidget(identity_btn)
        toolbar.addStretch(1)
        root.addLayout(toolbar)

        split = QSplitter(Qt.Orientation.Horizontal)
        self.calendar_split = split

        # Left: 30-day grid
        cal_wrap = QWidget()
        cal_layout = QVBoxLayout(cal_wrap)
        cal_layout.setContentsMargins(0, 0, 0, 0)
        cal_layout.setSpacing(6)

        self.calendar_range_label = QLabel("")
        self.calendar_range_label.setStyleSheet(f"color: {PALETTE['muted']};")
        cal_layout.addWidget(self.calendar_range_label)

        weekday_row = QHBoxLayout()
        weekday_row.setSpacing(4)
        for name in ("Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"):
            lbl = QLabel(name)
            lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
            lbl.setStyleSheet(
                f"color: {PALETTE['muted']}; font-weight: 600; font-size: 10px;"
            )
            weekday_row.addWidget(lbl)
        cal_layout.addLayout(weekday_row)

        self.calendar_grid_host = QWidget()
        self.calendar_grid = QGridLayout(self.calendar_grid_host)
        self.calendar_grid.setContentsMargins(0, 0, 0, 0)
        self.calendar_grid.setSpacing(4)
        cal_layout.addWidget(self.calendar_grid_host, stretch=1)

        self._calendar_day_buttons: List[QPushButton] = []
        self._calendar_selected_date = None
        self._calendar_by_date: Dict[str, List[Dict[str, Any]]] = {}

        split.addWidget(cal_wrap)

        # Right: day list + detail
        right = QWidget()
        right_layout = QVBoxLayout(right)
        right_layout.setContentsMargins(0, 0, 0, 0)
        right_layout.setSpacing(8)

        self.calendar_day_title = QLabel("Select a day")
        day_font = QFont("Segoe UI", 12)
        day_font.setBold(True)
        self.calendar_day_title.setFont(day_font)
        self.calendar_day_title.setStyleSheet(f"color: {PALETTE['accent']};")
        day_title_row = QHBoxLayout()
        day_title_row.addWidget(self.calendar_day_title, stretch=1)
        self.calendar_add_for_day_btn = QPushButton("Add for this day…")
        self.calendar_add_for_day_btn.clicked.connect(self._add_deadline_for_selected_day)
        day_title_row.addWidget(self.calendar_add_for_day_btn)
        right_layout.addLayout(day_title_row)

        self.calendar_day_tree = QTreeWidget()
        self.calendar_day_tree.setAlternatingRowColors(True)
        self.calendar_day_tree.setRootIsDecorated(False)
        self.calendar_day_tree.setUniformRowHeights(True)
        self.calendar_day_tree.setHeaderLabels(("From", "Subject", "Signal", "Status"))
        self.calendar_day_tree.setColumnWidth(0, 140)
        self.calendar_day_tree.setColumnWidth(1, 200)
        self.calendar_day_tree.setColumnWidth(2, 120)
        self.calendar_day_tree.setColumnWidth(3, 70)
        self.calendar_day_tree.itemClicked.connect(self._on_calendar_deadline_selected)
        self.calendar_day_tree.itemDoubleClicked.connect(
            self._open_deadline_item_in_outlook
        )
        self.calendar_day_tree.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.calendar_day_tree.customContextMenuRequested.connect(
            self._show_calendar_deadline_context_menu
        )
        right_layout.addWidget(self.calendar_day_tree, stretch=1)

        detail = QGroupBox("Deadline detail")
        detail_layout = QVBoxLayout(detail)
        self.calendar_preview_due = QLabel("")
        due_font = QFont("Segoe UI", 12)
        due_font.setBold(True)
        self.calendar_preview_due.setFont(due_font)
        self.calendar_preview_due.setStyleSheet(f"color: {PALETTE['accent']};")
        self.calendar_preview_due.setWordWrap(True)
        detail_layout.addWidget(self.calendar_preview_due)

        self.calendar_preview_signal = QLabel("")
        self.calendar_preview_signal.setWordWrap(True)
        self.calendar_preview_signal.setStyleSheet(
            "color: #8a4b08; background: #fff6e5; padding: 6px 8px; border-radius: 4px;"
        )
        detail_layout.addWidget(self.calendar_preview_signal)

        self.calendar_preview_meta = QLabel("")
        self.calendar_preview_meta.setWordWrap(True)
        self.calendar_preview_meta.setTextFormat(Qt.TextFormat.RichText)
        self.calendar_preview_meta.setStyleSheet("color: #4a6570; font-size: 11px;")
        detail_layout.addWidget(self.calendar_preview_meta)

        self.calendar_preview_body = QTextEdit()
        self.calendar_preview_body.setReadOnly(True)
        self.calendar_preview_body.setPlaceholderText(
            "Select a deadline on this day to read the email."
        )
        detail_layout.addWidget(self.calendar_preview_body, stretch=1)
        right_layout.addWidget(detail, stretch=2)

        split.addWidget(right)
        split.setStretchFactor(0, 3)
        split.setStretchFactor(1, 2)
        split.setSizes([560, 480])
        root.addWidget(split, stretch=1)

        QTimer.singleShot(500, self._refresh_calendar_panel)

    def _refresh_calendar_panel(self, rows: Optional[List[Dict[str, Any]]] = None) -> None:
        if not hasattr(self, "calendar_grid"):
            return
        try:
            # QPushButton.clicked may pass a bool — ignore non-list values
            if not isinstance(rows, list):
                cached = getattr(self, "_deadline_rows_cache", None)
                if isinstance(cached, list):
                    rows = cached
                else:
                    rows = self._collect_deadline_rows()
            row_list: List[Dict[str, Any]] = rows

            today = datetime.now().date()
            end = today + timedelta(days=29)
            self.calendar_range_label.setText(
                f"{today.strftime('%d %b %Y')} → {end.strftime('%d %b %Y')}"
            )

            by_date: Dict[str, List[Dict[str, Any]]] = {}
            dated_in_window = 0
            for row in row_list:
                due_d = self._deadline_due_date(row)
                if due_d is None:
                    continue
                if due_d < today or due_d > end:
                    continue
                key = due_d.isoformat()
                by_date.setdefault(key, []).append(row)
                dated_in_window += 1
            self._calendar_by_date = by_date

            # Clear previous day buttons
            while self.calendar_grid.count():
                item = self.calendar_grid.takeAt(0)
                if item is not None:
                    w = item.widget()
                    if w is not None:
                        w.deleteLater()
            self._calendar_day_buttons = []

            # Align grid to Monday of the week containing today
            start = today - timedelta(days=today.weekday())
            # Enough cells to cover through end (ceil to full weeks)
            total_days = (end - start).days + 1
            weeks = (total_days + 6) // 7

            selected = getattr(self, "_calendar_selected_date", None)
            if selected is not None and (selected < today or selected > end):
                selected = None
                self._calendar_selected_date = None

            for i in range(weeks * 7):
                day = start + timedelta(days=i)
                r, c = divmod(i, 7)
                btn = QPushButton()
                btn.setMinimumHeight(64)
                btn.setCursor(Qt.CursorShape.PointingHandCursor)
                in_window = today <= day <= end
                count = len(by_date.get(day.isoformat(), [])) if in_window else 0
                day_name = day.strftime("%d")
                month_tag = day.strftime("%b") if day.day == 1 or day == today else ""
                if count:
                    label = f"{day_name}\n{month_tag}\n{count} due".strip()
                else:
                    label = f"{day_name}\n{month_tag}".strip()
                btn.setText(label)

                if not in_window:
                    btn.setEnabled(False)
                    btn.setStyleSheet(
                        f"QPushButton {{ background: {PALETTE['bg']}; color: #b0bac2; "
                        f"border: 1px solid {PALETTE['border']}; border-radius: 6px; "
                        f"text-align: left; padding: 6px; }}"
                    )
                elif day == today:
                    bg = PALETTE["accent_soft"] if count == 0 else "#dff3f6"
                    btn.setStyleSheet(
                        f"QPushButton {{ background: {bg}; color: {PALETTE['text']}; "
                        f"border: 2px solid {PALETTE['accent']}; border-radius: 6px; "
                        f"text-align: left; padding: 6px; font-weight: 600; }}"
                        f"QPushButton:hover {{ background: #d0ebf0; }}"
                    )
                elif count:
                    btn.setStyleSheet(
                        "QPushButton { background: #fff6e5; color: #1a2e35; "
                        "border: 1px solid #e6c989; border-radius: 6px; "
                        "text-align: left; padding: 6px; font-weight: 600; }"
                        "QPushButton:hover { background: #ffefd0; }"
                    )
                else:
                    btn.setStyleSheet(
                        f"QPushButton {{ background: {PALETTE['surface']}; color: {PALETTE['text']}; "
                        f"border: 1px solid {PALETTE['border']}; border-radius: 6px; "
                        f"text-align: left; padding: 6px; }}"
                        f"QPushButton:hover {{ background: {PALETTE['accent_soft']}; }}"
                    )

                if selected is not None and day == selected:
                    btn.setStyleSheet(
                        btn.styleSheet()
                        + f"QPushButton {{ outline: none; border: 2px solid {PALETTE['accent']}; }}"
                    )

                if in_window:
                    btn.clicked.connect(
                        lambda _checked=False, d=day: self._on_calendar_day_clicked(d)
                    )
                self.calendar_grid.addWidget(btn, r, c)
                self._calendar_day_buttons.append(btn)

            self.calendar_hint.setText(
                f"{dated_in_window} deadline(s) with dates in the next 30 days "
                f"({len(by_date)} day(s) with activity)."
            )

            # Keep / restore day selection
            if selected is not None:
                self._on_calendar_day_clicked(selected)
            else:
                self._on_calendar_day_clicked(today)
        except Exception as e:
            self.logger.error("Error refreshing calendar panel: %s", e)

    def _on_calendar_day_clicked(self, day) -> None:
        self._calendar_selected_date = day
        key = day.isoformat()
        rows = list((getattr(self, "_calendar_by_date", {}) or {}).get(key) or [])
        self.calendar_day_title.setText(
            f"{day.strftime('%A %d %b %Y')} — {len(rows)} deadline(s)"
        )
        if hasattr(self, "calendar_day_tree"):
            self.calendar_day_tree.clear()
            for row in rows:
                det = row.get("det") or {}
                dl = row.get("deadline") or {}
                mail = det.get("mail") or {}
                item = QTreeWidgetItem(self.calendar_day_tree)
                item.setText(0, str(mail.get("sender") or "")[:60])
                subject = str(mail.get("subject") or "")
                item.setText(1, subject[:70] + ("..." if len(subject) > 70 else ""))
                item.setText(2, str(dl.get("cue") or dl.get("label") or "")[:60])
                item.setText(3, str(det.get("status") or "—"))
                item.setData(0, Qt.ItemDataRole.UserRole, row)
            if rows:
                first = self.calendar_day_tree.topLevelItem(0)
                if first is not None:
                    self.calendar_day_tree.setCurrentItem(first)
                    self._on_calendar_deadline_selected(first)
            else:
                self._clear_calendar_detail()

    def _clear_calendar_detail(self) -> None:
        if hasattr(self, "calendar_preview_due"):
            self.calendar_preview_due.setText("No deadlines on this day.")
        if hasattr(self, "calendar_preview_signal"):
            self.calendar_preview_signal.setText("")
            self.calendar_preview_signal.setVisible(False)
        if hasattr(self, "calendar_preview_meta"):
            self.calendar_preview_meta.setText("")
        if hasattr(self, "calendar_preview_body"):
            self.calendar_preview_body.clear()

    def _on_calendar_deadline_selected(self, item, _column=None) -> None:
        if item is None:
            return
        payload = item.data(0, Qt.ItemDataRole.UserRole) or {}
        if not isinstance(payload, dict):
            return
        if "det" in payload:
            det = payload.get("det") or {}
            deadline = payload.get("deadline") or {}
        else:
            det = payload
            deadline = merge_deadline_hits(det.get("deadlines") or []) or {}

        mail = det.get("mail") or {}
        due = str(deadline.get("due") or deadline.get("label") or "No date parsed")
        cue = str(deadline.get("cue") or "deadline")
        snippet = str(deadline.get("snippet") or "").strip()

        if hasattr(self, "calendar_preview_due"):
            self.calendar_preview_due.setText(f"Due: {due}")
        if hasattr(self, "calendar_preview_signal"):
            bits = [f"Signal: {cue}"]
            if snippet:
                bits.append(snippet)
            self.calendar_preview_signal.setText(" — ".join(bits))
            self.calendar_preview_signal.setVisible(True)

        sender = str(mail.get("sender") or "")
        subject = str(mail.get("subject") or "")
        when = str(mail.get("received") or "")
        account = str(mail.get("account_smtp") or "")
        if hasattr(self, "calendar_preview_meta"):
            self.calendar_preview_meta.setText(
                f"<b>{subject}</b><br>"
                f"From: {sender}<br>"
                f"When: {when} &nbsp;|&nbsp; Account: {account}"
            )
        body = str(mail.get("body_preview") or "")
        if hasattr(self, "calendar_preview_body"):
            self.calendar_preview_body.setHtml(
                highlight_action_html(f"Subject: {subject}\n\n{body}")
            )
        self._show_email_preview(det if isinstance(det, dict) else {})

    def _show_calendar_deadline_context_menu(self, position) -> None:
        item = self.calendar_day_tree.itemAt(position)
        if not item:
            return
        self.calendar_day_tree.setCurrentItem(item)
        row = self._deadline_row_payload(item)
        is_manual = bool(row.get("manual") or (row.get("det") or {}).get("manual_id"))
        menu = QMenu(self)
        if is_manual:
            menu.addAction("Delete manual deadline", lambda: self._delete_selected_manual_deadline(item))
        else:
            menu.addAction("Open in Outlook", lambda: self._open_deadline_item_in_outlook(item))
            menu.addAction(
                "Not a deadline (teach GURI)",
                lambda: self._reject_deadline_learning(item),
            )
            menu.addAction(
                "Not for me (teach GURI)",
                lambda: self._reject_deadline_not_for_me(item),
            )
            menu.addAction(
                "Confirm this deadline",
                lambda: self._confirm_deadline_learning(item),
            )
        menu.addAction("View Details", lambda: self._view_deadline_details(item))
        menu.exec(self.calendar_day_tree.mapToGlobal(position))

    def _on_deadline_item_selected(self, item, _column=None) -> None:
        """Show deadline summary + email body in the right-hand pane."""
        if item is None:
            return
        payload = item.data(0, Qt.ItemDataRole.UserRole) or {}
        if not isinstance(payload, dict):
            return
        # Support both new {det, deadline} rows and legacy bare det
        if "det" in payload:
            det = payload.get("det") or {}
            deadline = payload.get("deadline") or {}
        else:
            det = payload
            deadline = merge_deadline_hits(det.get("deadlines") or []) or {}
            if not deadline and "deadline" in [
                str(a).lower() for a in (det.get("actions") or [])
            ]:
                deadline = {
                    "cue": "deadline",
                    "label": "Deadline / urgency",
                    "due": None,
                    "snippet": "",
                }

        mail = det.get("mail") or {}
        due = str(deadline.get("due") or deadline.get("label") or "No date parsed")
        cue = str(deadline.get("cue") or "deadline")
        snippet = str(deadline.get("snippet") or "").strip()

        if hasattr(self, "deadlines_preview_due"):
            self.deadlines_preview_due.setText(f"Due: {due}")
        if hasattr(self, "deadlines_preview_signal"):
            bits = [f"Signal: {cue}"]
            if snippet:
                bits.append(snippet)
            self.deadlines_preview_signal.setText(" — ".join(bits))
            self.deadlines_preview_signal.setVisible(True)

        sender = str(mail.get("sender") or "")
        subject = str(mail.get("subject") or "")
        when = str(mail.get("received") or "")
        account = str(mail.get("account_smtp") or "")
        if hasattr(self, "deadlines_preview_meta"):
            self.deadlines_preview_meta.setText(
                f"<b>{subject}</b><br>"
                f"From: {sender}<br>"
                f"When: {when} &nbsp;|&nbsp; Account: {account}"
            )

        body = str(mail.get("body_preview") or "")
        blob = f"Subject: {subject}\n\n{body}"
        if hasattr(self, "deadlines_preview_body"):
            self.deadlines_preview_body.setHtml(highlight_action_html(blob))

        # Also drive Welcome preview so feedback buttons stay usable
        self._show_email_preview(det if isinstance(det, dict) else {})

    def _show_deadline_accounts_dialog(self) -> None:
        """Popup to choose which accounts feed the Deadlines tab."""
        dialog = QDialog(self)
        dialog.setWindowTitle("Deadline accounts")
        dialog.resize(480, 420)
        layout = QVBoxLayout(dialog)

        intro = QLabel(
            "Select Outlook accounts whose scraped mail should appear on the "
            "Deadlines tab. Selections are saved for next time."
        )
        intro.setWordWrap(True)
        layout.addWidget(intro)

        status = QLabel("Loading accounts…")
        status.setStyleSheet(f"color: {PALETTE['muted']};")
        layout.addWidget(status)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        host = QWidget()
        host_layout = QVBoxLayout(host)
        host_layout.setContentsMargins(4, 4, 4, 4)
        host_layout.setSpacing(6)
        scroll.setWidget(host)
        layout.addWidget(scroll, stretch=1)

        checks: List[tuple] = []  # (QCheckBox, key, label)

        all_btn = QPushButton("Select all")
        none_btn = QPushButton("Clear")
        save_btn = QPushButton("Save")
        save_btn.setDefault(True)
        cancel_btn = QPushButton("Cancel")
        cancel_btn.clicked.connect(dialog.reject)
        add_dialog_button_row(
            layout,
            leading=[all_btn, none_btn],
            affirmative=save_btn,
            cancel=cancel_btn,
        )

        selected = {
            str(k).strip()
            for k in (getattr(self, "deadline_account_keys", None) or [])
            if str(k).strip()
        }
        selected_l = {k.lower() for k in selected}

        def populate(configs: List[Any]) -> None:
            while host_layout.count():
                item = host_layout.takeAt(0)
                if item is not None:
                    w = item.widget()
                    if w is not None:
                        w.deleteLater()
            checks.clear()
            if not configs:
                status.setText("No Outlook accounts found.")
                return
            status.setText(f"{len(configs)} account(s) — tick those to include.")
            for cfg in configs:
                store_id = str(getattr(cfg, "store_id", "") or "")
                smtp = str(getattr(cfg, "smtp", "") or "")
                display = str(getattr(cfg, "display", "") or smtp or store_id[:24])
                key = store_id or smtp
                if not key:
                    continue
                label = display
                if smtp and smtp.lower() not in display.lower():
                    label = f"{display}  <{smtp}>"
                cb = QCheckBox(label)
                enabled = bool(getattr(cfg, "enabled", True))
                pre = (
                    key in selected
                    or key.lower() in selected_l
                    or (smtp and smtp.lower() in selected_l)
                    or (not selected and enabled)
                )
                cb.setChecked(pre)
                host_layout.addWidget(cb)
                checks.append((cb, key, smtp))
            host_layout.addStretch(1)

        def on_all() -> None:
            for cb, _k, _s in checks:
                cb.setChecked(True)

        def on_none() -> None:
            for cb, _k, _s in checks:
                cb.setChecked(False)

        def on_save() -> None:
            keys: List[str] = []
            for cb, key, smtp in checks:
                if cb.isChecked():
                    keys.append(key)
                    if smtp and smtp not in keys:
                        keys.append(smtp)
            self.deadline_account_keys = keys
            try:
                save_deadline_account_keys(keys)
            except Exception as exc:
                self.logger.warning("Failed saving deadline accounts: %s", exc)
            self._refresh_deadlines_panel()
            dialog.accept()

        all_btn.clicked.connect(on_all)
        none_btn.clicked.connect(on_none)
        cancel_btn.clicked.connect(dialog.reject)
        save_btn.clicked.connect(on_save)

        # Prefer live discovery; fall back to saved AES configs
        try:
            configs = discover_outlook_accounts()
            if not configs:
                configs = load_aes_account_configs()
            populate(configs)
        except Exception as exc:
            status.setText(f"Outlook discovery failed — using saved configs ({exc})")
            try:
                populate(load_aes_account_configs())
            except Exception:
                populate([])

        dialog.exec()

    def _deadline_row_payload(self, item) -> Dict[str, Any]:
        """Normalize tree UserRole to {det, deadline}."""
        payload = item.data(0, Qt.ItemDataRole.UserRole) if item is not None else None
        if not isinstance(payload, dict):
            return {"det": {}, "deadline": {}}
        if "det" in payload:
            return {
                "det": payload.get("det") or {},
                "deadline": payload.get("deadline") or {},
            }
        return {"det": payload, "deadline": {}}

    def _show_deadline_context_menu(self, position) -> None:
        item = self.deadlines_tree.itemAt(position)
        if not item:
            return
        self.deadlines_tree.setCurrentItem(item)
        row = self._deadline_row_payload(item)
        is_manual = bool(row.get("manual") or (row.get("det") or {}).get("manual_id"))
        menu = QMenu(self)
        menu.addAction(
            "Mark as Done",
            lambda: self._mark_deadline_status(item, "Done"),
        )
        menu.addAction(
            "Mark as Need to Reply",
            lambda: self._mark_deadline_status(item, "Need Reply"),
        )
        menu.addSeparator()
        if is_manual:
            menu.addAction("Delete manual deadline", lambda: self._delete_selected_manual_deadline(item))
        else:
            menu.addAction("Open in Outlook", lambda: self._open_deadline_item_in_outlook(item))
            menu.addAction(
                "Not a deadline (teach GURI)",
                lambda: self._reject_deadline_learning(item),
            )
            menu.addAction(
                "Not for me (teach GURI)",
                lambda: self._reject_deadline_not_for_me(item),
            )
            menu.addAction(
                "Confirm this deadline",
                lambda: self._confirm_deadline_learning(item),
            )
        menu.addAction("View Details", lambda: self._view_deadline_details(item))
        menu.exec(self.deadlines_tree.mapToGlobal(position))

    def _mark_deadline_status(self, item, status: str) -> None:
        """Mark status using the unwrapped detection dict stored on the row."""
        row = self._deadline_row_payload(item)
        det = row.get("det") or {}
        mid = str(det.get("manual_id") or "")
        if mid:
            # Persist status on the manual deadline record
            updated = []
            for entry in getattr(self, "manual_deadlines", []) or []:
                e = dict(entry)
                if str(e.get("id") or "") == mid:
                    e["status"] = status
                updated.append(e)
            self.manual_deadlines = updated
            try:
                save_manual_deadlines(updated)
            except Exception as exc:
                self.logger.warning("Failed saving manual deadline status: %s", exc)
            det["status"] = status
            row["det"] = det
            item.setData(0, Qt.ItemDataRole.UserRole, row)
            item.setText(6, status)
            self.status_bar.showMessage(f"Marked as '{status}'")
            return
        # Temporarily put det on the item so shared marker works
        item.setData(0, Qt.ItemDataRole.UserRole, det)
        self._mark_item_status(item, status, actions_tree=True)
        # Restore full row payload with updated status
        det["status"] = status
        row["det"] = det
        item.setData(0, Qt.ItemDataRole.UserRole, row)
        item.setText(6, status)

    def _view_deadline_details(self, item) -> None:
        row = self._deadline_row_payload(item)
        # Reuse details dialog by briefly swapping UserRole to the det
        det = row.get("det") or {}
        item.setData(0, Qt.ItemDataRole.UserRole, det)
        try:
            self._view_email_details(item)
        finally:
            item.setData(0, Qt.ItemDataRole.UserRole, row)

    def _open_deadline_item_in_outlook(self, item, _column=None) -> None:
        row = self._deadline_row_payload(item)
        det = row.get("det") or {}
        if row.get("manual") or det.get("manual_id"):
            # Show detail pane instead — no Outlook item for manual deadlines
            tree = item.treeWidget()
            if tree is getattr(self, "calendar_day_tree", None):
                self._on_calendar_deadline_selected(item)
            else:
                self._on_deadline_item_selected(item)
            self.status_bar.showMessage("Manual deadline — no Outlook message to open")
            return
        mail = det.get("mail") or det
        entry_id = str(mail.get("entry_id") or "")
        store_id = str(mail.get("store_id") or "")
        if tree := item.treeWidget():
            if tree is getattr(self, "calendar_day_tree", None):
                self._on_calendar_deadline_selected(item)
            else:
                self._on_deadline_item_selected(item)
        if not entry_id:
            QMessageBox.information(
                self,
                "Open in Outlook",
                "This row has no Outlook EntryID (GURI DB-only item).",
            )
            return
        if not open_mail_in_outlook(entry_id, store_id):
            QMessageBox.warning(self, "Open in Outlook", "Could not open the message in Outlook.")

    def _open_selected_deadline_in_outlook(self) -> None:
        item = self.deadlines_tree.currentItem() if hasattr(self, "deadlines_tree") else None
        if item is None:
            self.status_bar.showMessage("Select a deadline row first")
            return
        self._open_deadline_item_in_outlook(item)

    def _create_accounts_tab(self):
        """Email accounts — what to scan (enabled / Responses / In Cc)."""
        frame = QWidget()
        self.notebook.addTab(frame, "Accounts")
        root = QVBoxLayout(frame)
        root.setContentsMargins(12, 12, 12, 12)
        root.setSpacing(10)

        head = QLabel("Email accounts")
        head.setStyleSheet(
            f"font-size: 14px; font-weight: 700; color: {PALETTE['accent']};"
        )
        root.addWidget(head)

        sub = QLabel(
            "Choose which Outlook mailboxes GURI and AES should process, and which "
            "message types count for each account. Settings are shared with AES "
            "(saved to LocalAppData\\GeoFooter)."
        )
        sub.setWordWrap(True)
        sub.setStyleSheet(f"color: {PALETTE['muted']};")
        root.addWidget(sub)

        legend = QGroupBox("What each option means")
        legend_layout = QVBoxLayout(legend)
        for title, blurb in (
            (
                "Scan account",
                "Include this mailbox’s Inbox when scraping and when AES watches mail.",
            ),
            (
                "Responses",
                "Include reply / forward threads (Re:/Fw:). Off demotes response noise "
                "for this account.",
            ),
            (
                "In Cc",
                "Include messages where this account is only on Cc (not To). "
                "Off skips Cc-only mail for this account.",
            ),
        ):
            row = QLabel(f"<b>{title}</b> — {blurb}")
            row.setWordWrap(True)
            row.setTextFormat(Qt.TextFormat.RichText)
            legend_layout.addWidget(row)
        root.addWidget(legend)

        toolbar = QHBoxLayout()
        refresh_btn = QPushButton("Refresh from Outlook")
        refresh_btn.setToolTip("Re-read Outlook accounts and merge with saved flags.")
        refresh_btn.clicked.connect(self._accounts_refresh)
        self.accounts_refresh_btn = refresh_btn
        all_on_btn = QPushButton("Enable all")
        all_on_btn.clicked.connect(lambda: self._accounts_set_all(True))
        all_off_btn = QPushButton("Disable all")
        all_off_btn.clicked.connect(lambda: self._accounts_set_all(False))
        save_btn = QPushButton("Save")
        save_btn.clicked.connect(self._accounts_save)
        toolbar.addWidget(refresh_btn)
        toolbar.addWidget(all_on_btn)
        toolbar.addWidget(all_off_btn)
        toolbar.addStretch(1)
        toolbar.addWidget(save_btn)
        root.addLayout(toolbar)

        self.accounts_status = QLabel("Loading accounts…")
        self.accounts_status.setStyleSheet(f"color: {PALETTE['muted']};")
        root.addWidget(self.accounts_status)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)
        self.accounts_list_host = QWidget()
        self.accounts_list_layout = QVBoxLayout(self.accounts_list_host)
        self.accounts_list_layout.setContentsMargins(0, 0, 8, 0)
        self.accounts_list_layout.setSpacing(8)
        self.accounts_list_layout.addStretch(1)
        scroll.setWidget(self.accounts_list_host)
        root.addWidget(scroll, stretch=1)

        QTimer.singleShot(300, self._accounts_refresh)

    def _accounts_clear_rows(self) -> None:
        while self.accounts_list_layout.count():
            item = self.accounts_list_layout.takeAt(0)
            if item is not None:
                w = item.widget()
                if w is not None:
                    w.deleteLater()
        self._account_rows = []

    def _accounts_refresh(self) -> None:
        if self._accounts_busy:
            return
        self._accounts_busy = True
        if hasattr(self, "accounts_refresh_btn"):
            self.accounts_refresh_btn.setEnabled(False)
        self.accounts_status.setText("Reading Outlook accounts…")
        self.accounts_status.setStyleSheet(f"color: {PALETTE['muted']};")

        def work() -> None:
            try:
                configs = discover_outlook_accounts()
                self._accounts_bridge.ready.emit(configs, "")
            except Exception as exc:
                # Fall back to saved config if Outlook is unavailable
                try:
                    saved = load_aes_account_configs()
                    if saved:
                        self._accounts_bridge.ready.emit(
                            saved, f"Outlook unavailable ({exc}); showing saved settings"
                        )
                    else:
                        self._accounts_bridge.ready.emit(None, str(exc))
                except Exception as exc2:
                    self._accounts_bridge.ready.emit(None, f"{exc}; {exc2}")

        threading.Thread(target=work, daemon=True).start()

    def _on_accounts_discovered(self, configs_obj, error: str) -> None:
        self._accounts_busy = False
        if hasattr(self, "accounts_refresh_btn"):
            self.accounts_refresh_btn.setEnabled(True)
        if configs_obj is None:
            self.accounts_status.setText(f"Could not load accounts: {error}")
            self.accounts_status.setStyleSheet(f"color: {PALETTE['err']};")
            return

        configs: List[AesAccountScanConfig] = list(configs_obj or [])
        self._accounts_clear_rows()

        if not configs:
            empty = QLabel(
                "No Outlook accounts found. Open Outlook, then click Refresh from Outlook."
            )
            empty.setStyleSheet(f"color: {PALETTE['muted']};")
            self.accounts_list_layout.addWidget(empty)
            self.accounts_list_layout.addStretch(1)
            self.accounts_status.setText("No accounts")
            return

        for cfg in configs:
            card = QGroupBox()
            card_layout = QVBoxLayout(card)
            card_layout.setSpacing(6)

            title = cfg.display or cfg.smtp or "Account"
            smtp_line = cfg.smtp or "(no SMTP address)"
            name_lbl = QLabel(
                f"<b>{title}</b><br>"
                f"<span style='color:{PALETTE['muted']}'>{smtp_line}</span>"
            )
            name_lbl.setTextFormat(Qt.TextFormat.RichText)
            name_lbl.setWordWrap(True)
            card_layout.addWidget(name_lbl)

            opts = QHBoxLayout()
            enabled_cb = QCheckBox("Scan account")
            enabled_cb.setChecked(bool(cfg.enabled))
            responses_cb = QCheckBox("Responses")
            responses_cb.setChecked(bool(cfg.responses))
            responses_cb.setToolTip("Include Re:/Fw: threads for this account.")
            in_cc_cb = QCheckBox("In Cc")
            in_cc_cb.setChecked(bool(cfg.in_cc))
            in_cc_cb.setToolTip("Include messages where this account is only on Cc.")

            def _sync(checked: bool, r_cb=responses_cb, c_cb=in_cc_cb) -> None:
                r_cb.setEnabled(checked)
                c_cb.setEnabled(checked)

            enabled_cb.toggled.connect(_sync)
            _sync(enabled_cb.isChecked())

            opts.addWidget(enabled_cb)
            opts.addWidget(responses_cb)
            opts.addWidget(in_cc_cb)
            opts.addStretch(1)
            card_layout.addLayout(opts)
            self.accounts_list_layout.addWidget(card)

            self._account_rows.append(
                {
                    "store_id": cfg.store_id,
                    "smtp": cfg.smtp,
                    "display": cfg.display,
                    "enabled": enabled_cb,
                    "responses": responses_cb,
                    "in_cc": in_cc_cb,
                }
            )

        self.accounts_list_layout.addStretch(1)

        n_on = sum(1 for c in configs if c.enabled)
        msg = f"{len(configs)} account(s) · {n_on} enabled for scanning"
        if error:
            msg += f" — {error}"
            self.accounts_status.setStyleSheet(f"color: {PALETTE['accent']};")
        else:
            self.accounts_status.setStyleSheet(f"color: {PALETTE['ok']};")
        self.accounts_status.setText(msg)
        self._inbox_search_fill_accounts(configs)

    def _accounts_set_all(self, enabled: bool) -> None:
        for row in self._account_rows:
            row["enabled"].setChecked(enabled)

    def _accounts_collect(self) -> List[AesAccountScanConfig]:
        out: List[AesAccountScanConfig] = []
        for row in self._account_rows:
            out.append(
                AesAccountScanConfig(
                    store_id=str(row.get("store_id") or ""),
                    smtp=str(row.get("smtp") or ""),
                    display=str(row.get("display") or ""),
                    enabled=bool(row["enabled"].isChecked()),
                    responses=bool(row["responses"].isChecked()),
                    in_cc=bool(row["in_cc"].isChecked()),
                )
            )
        return out

    def _accounts_save(self) -> None:
        configs = self._accounts_collect()
        if not configs:
            QMessageBox.information(
                self,
                "Accounts",
                "No accounts to save. Refresh from Outlook first.",
            )
            return
        try:
            path = save_aes_account_configs(configs)
            n_on = sum(1 for c in configs if c.enabled)
            self.accounts_status.setText(
                f"Saved {len(configs)} account(s) ({n_on} enabled) → {path}"
            )
            self.accounts_status.setStyleSheet(f"color: {PALETTE['ok']};")
            self.status_bar.showMessage(
                f"Account scan settings saved ({n_on}/{len(configs)} enabled)"
            )
            self._inbox_search_fill_accounts(configs)
        except Exception as exc:
            QMessageBox.critical(self, "Accounts", f"Failed to save:\n{exc}")
            self.logger.error("Account settings save failed: %s", exc)

    def _focus_inbox_search(self) -> None:
        for i in range(self.notebook.count()):
            if self.notebook.tabText(i) == "Inbox Search":
                self.notebook.setCurrentIndex(i)
                break
        if hasattr(self, "inbox_search_query"):
            self.inbox_search_query.setFocus()

    def _create_inbox_search_tab(self) -> None:
        """Scan one Outlook inbox for conversational theme, tone, and subject matter."""
        frame = QWidget()
        self.inbox_search_tab_index = self.notebook.addTab(frame, "Inbox Search")
        root = QVBoxLayout(frame)
        root.setContentsMargins(12, 12, 12, 12)
        root.setSpacing(8)
        cfg = getattr(self, "_inbox_search_settings", {}) or {}

        head = QLabel("Inbox Search")
        head.setStyleSheet(
            f"font-size: 14px; font-weight: 700; color: {PALETTE['accent']};"
        )
        root.addWidget(head)
        sub = QLabel(
            "Scan a chosen mailbox for <b>conversations</b>, not isolated messages. "
            "Describe the material: a topic, a tone (frustrated, positive, urgent), "
            "or both — e.g. “frustrated suppliers about invoices” or "
            "“positive feedback on the O-RAN trial”."
        )
        sub.setWordWrap(True)
        sub.setTextFormat(Qt.TextFormat.RichText)
        sub.setStyleSheet(f"color: {PALETTE['muted']}; font-size: 11px;")
        root.addWidget(sub)

        form = QGridLayout()
        form.setHorizontalSpacing(8)
        form.setVerticalSpacing(6)

        form.addWidget(QLabel("Inbox"), 0, 0)
        self.inbox_search_account = QComboBox()
        self.inbox_search_account.setMinimumWidth(280)
        form.addWidget(self.inbox_search_account, 0, 1, 1, 3)
        refresh_acc_btn = QPushButton("Refresh inboxes")
        refresh_acc_btn.setToolTip("Re-read Outlook accounts for this picker.")
        refresh_acc_btn.clicked.connect(self._accounts_refresh)
        form.addWidget(refresh_acc_btn, 0, 4)

        form.addWidget(QLabel("Look for"), 1, 0)
        self.inbox_search_query = QTextEdit()
        self.inbox_search_query.setPlaceholderText(
            "e.g. worried about the Surrey timeline — or angry unpaid-invoice threads"
        )
        self.inbox_search_query.setMaximumHeight(72)
        self.inbox_search_query.setPlainText(str(cfg.get("last_query") or ""))
        form.addWidget(self.inbox_search_query, 1, 1, 1, 4)

        opts = QHBoxLayout()
        opts.setSpacing(8)
        self.inbox_search_sent_cb = QCheckBox("Include Sent (full threads)")
        self.inbox_search_sent_cb.setChecked(bool(cfg.get("include_sent", True)))
        self.inbox_search_sent_cb.setToolTip(
            "Pull Sent Items as well as Inbox so your replies sit in the conversation."
        )
        opts.addWidget(self.inbox_search_sent_cb)

        self.inbox_search_ollama_cb = QCheckBox("Rank with Ollama")
        self.inbox_search_ollama_cb.setChecked(bool(cfg.get("use_ollama", True)))
        self.inbox_search_ollama_cb.setToolTip(
            "After heuristic scoring, ask the configured Ollama model to rank "
            "the strongest threads as conversations."
        )
        opts.addWidget(self.inbox_search_ollama_cb)

        opts.addWidget(QLabel("Days"))
        self.inbox_search_days = QSpinBox()
        self.inbox_search_days.setRange(1, 180)
        self.inbox_search_days.setValue(int(cfg.get("days") or 30))
        opts.addWidget(self.inbox_search_days)

        opts.addWidget(QLabel("Max messages"))
        self.inbox_search_max = QSpinBox()
        self.inbox_search_max.setRange(20, 400)
        self.inbox_search_max.setValue(int(cfg.get("max_messages") or 120))
        self.inbox_search_max.setToolTip(
            "Cap per mailbox. Bodies are read, so keep this modest to avoid freezing Outlook."
        )
        opts.addWidget(self.inbox_search_max)
        opts.addStretch(1)
        form.addLayout(opts, 2, 1, 1, 4)

        root.addLayout(form)

        btn_row = QHBoxLayout()
        self.inbox_search_status = QLabel("Pick an inbox, describe the material, then Scan.")
        self.inbox_search_status.setStyleSheet(f"color: {PALETTE['muted']};")
        self.inbox_search_status.setWordWrap(True)
        btn_row.addWidget(self.inbox_search_status, stretch=1)
        scan_btn = QPushButton("Scan inbox")
        scan_btn.setDefault(True)
        scan_btn.clicked.connect(self._start_inbox_search)
        self.inbox_search_btn = scan_btn
        clear_btn = QPushButton("Clear")
        clear_btn.clicked.connect(self._clear_inbox_search)
        open_btn = QPushButton("Open in Outlook")
        open_btn.clicked.connect(self._open_inbox_search_in_outlook)
        btn_row.addWidget(scan_btn)
        btn_row.addWidget(open_btn)
        btn_row.addWidget(clear_btn)
        root.addLayout(btn_row)

        splitter = QSplitter(Qt.Orientation.Vertical)
        self.inbox_search_tree = QTreeWidget()
        self.inbox_search_tree.setAlternatingRowColors(True)
        self.inbox_search_tree.setRootIsDecorated(True)
        self.inbox_search_tree.setHeaderLabels(
            ("Conversation / message", "Tone", "People", "Msgs", "Latest", "Score", "Why")
        )
        self.inbox_search_tree.setColumnCount(7)
        self.inbox_search_tree.setColumnWidth(0, 280)
        self.inbox_search_tree.setColumnWidth(1, 90)
        self.inbox_search_tree.setColumnWidth(2, 160)
        self.inbox_search_tree.setColumnWidth(3, 48)
        self.inbox_search_tree.setColumnWidth(4, 118)
        self.inbox_search_tree.setColumnWidth(5, 56)
        hdr = self.inbox_search_tree.header()
        hdr.setStretchLastSection(True)
        hdr.setSectionResizeMode(0, QHeaderView.ResizeMode.Interactive)
        self.inbox_search_tree.itemClicked.connect(self._on_inbox_search_item_selected)
        self.inbox_search_tree.itemDoubleClicked.connect(
            lambda item, _col: self._open_inbox_search_item(item)
        )
        splitter.addWidget(self.inbox_search_tree)

        self.inbox_search_preview = QTextEdit()
        self.inbox_search_preview.setReadOnly(True)
        self.inbox_search_preview.setHtml(
            "<i style='color:#5a7280'>Select a conversation to read the thread "
            "as a whole — tone, subject matter, and the messages in order.</i>"
        )
        splitter.addWidget(self.inbox_search_preview)
        splitter.setStretchFactor(0, 3)
        splitter.setStretchFactor(1, 2)
        root.addWidget(splitter, stretch=1)

        QTimer.singleShot(500, self._inbox_search_load_accounts)

    def _inbox_search_load_accounts(self) -> None:
        try:
            configs = load_aes_account_configs()
        except Exception:
            configs = []
        if configs:
            self._inbox_search_fill_accounts(configs)

    def _inbox_search_fill_accounts(self, configs: List[Any]) -> None:
        combo = getattr(self, "inbox_search_account", None)
        if combo is None:
            return
        saved = getattr(self, "_inbox_search_settings", {}) or {}
        want_store = str(saved.get("store_id") or "")
        want_smtp = str(saved.get("smtp") or "").lower()
        current = combo.currentData()
        if isinstance(current, dict):
            want_store = want_store or str(current.get("store_id") or "")
            want_smtp = want_smtp or str(current.get("smtp") or "").lower()

        combo.blockSignals(True)
        combo.clear()
        select_idx = 0
        for i, cfg in enumerate(configs or []):
            store_id = str(getattr(cfg, "store_id", "") or "")
            smtp = str(getattr(cfg, "smtp", "") or "")
            display = str(getattr(cfg, "display", "") or smtp or store_id[:24])
            enabled = bool(getattr(cfg, "enabled", True))
            label = display
            if smtp and smtp.lower() not in display.lower():
                label = f"{display}  <{smtp}>"
            if not enabled:
                label += "  (scan off)"
            combo.addItem(
                label,
                {"store_id": store_id, "smtp": smtp, "display": display},
            )
            if (want_store and store_id == want_store) or (
                want_smtp and smtp.lower() == want_smtp
            ):
                select_idx = i
        combo.setCurrentIndex(select_idx if combo.count() else -1)
        combo.blockSignals(False)
        if not combo.count() and hasattr(self, "inbox_search_status"):
            self.inbox_search_status.setText(
                "No Outlook accounts yet — open Outlook, then Refresh inboxes."
            )

    def _inbox_search_selected_account(self) -> Dict[str, str]:
        data = self.inbox_search_account.currentData() if hasattr(self, "inbox_search_account") else None
        if isinstance(data, dict):
            return {
                "store_id": str(data.get("store_id") or ""),
                "smtp": str(data.get("smtp") or ""),
                "display": str(data.get("display") or ""),
            }
        return {"store_id": "", "smtp": "", "display": ""}

    def _start_inbox_search(self) -> None:
        if self._inbox_search_busy:
            self.status_bar.showMessage("Inbox search already running…")
            return
        acc = self._inbox_search_selected_account()
        if not acc.get("store_id") and not acc.get("smtp"):
            QMessageBox.information(
                self,
                "Inbox Search",
                "Pick an Outlook inbox first. Use Refresh inboxes if the list is empty.",
            )
            return
        query = self.inbox_search_query.toPlainText().strip()
        if not query:
            QMessageBox.information(
                self,
                "Inbox Search",
                "Describe the material to look for — a topic, a tone, or both.",
            )
            self.inbox_search_query.setFocus()
            return

        days = int(self.inbox_search_days.value())
        max_messages = int(self.inbox_search_max.value())
        include_sent = bool(self.inbox_search_sent_cb.isChecked())
        use_ollama = bool(self.inbox_search_ollama_cb.isChecked())
        settings = {
            "store_id": acc.get("store_id") or "",
            "smtp": acc.get("smtp") or "",
            "days": days,
            "max_messages": max_messages,
            "include_sent": include_sent,
            "use_ollama": use_ollama,
            "last_query": query,
        }
        self._inbox_search_settings = settings
        try:
            save_inbox_search_settings(settings)
        except Exception as exc:
            self.logger.warning("Could not save inbox search settings: %s", exc)

        ollama_cfg = {}
        if use_ollama:
            try:
                ollama_cfg = (
                    self._ollama_collect_cfg()
                    if hasattr(self, "ollama_url_edit")
                    else dict(self.ollama_cfg or {})
                )
            except Exception:
                ollama_cfg = dict(self.ollama_cfg or {})

        self._inbox_search_busy = True
        self.inbox_search_btn.setEnabled(False)
        self.inbox_search_status.setText("Scanning inbox…")
        self.inbox_search_status.setStyleSheet(f"color: {PALETTE['muted']};")
        self.status_bar.showMessage(
            f"Scanning {acc.get('display') or acc.get('smtp') or 'inbox'}…"
        )

        store_id = acc.get("store_id") or ""
        smtp = acc.get("smtp") or ""

        def work() -> None:
            try:
                def progress(msg: str) -> None:
                    self._inbox_search_bridge.progress.emit(msg)

                payload = search_inbox_conversations(
                    store_id=store_id,
                    smtp=smtp,
                    query_text=query,
                    days=days,
                    max_messages=max_messages,
                    include_sent=include_sent,
                    use_ollama=use_ollama,
                    ollama_cfg=ollama_cfg,
                    progress=progress,
                )
                self._inbox_search_bridge.finished.emit(payload, "")
            except Exception as exc:
                self._inbox_search_bridge.finished.emit(None, str(exc))

        threading.Thread(target=work, daemon=True).start()

    def _on_inbox_search_progress(self, message: str) -> None:
        text = (message or "").strip()
        if not text:
            return
        if hasattr(self, "inbox_search_status"):
            self.inbox_search_status.setText(text)
        self.status_bar.showMessage(text)

    def _on_inbox_search_finished(self, payload_obj, error: str) -> None:
        self._inbox_search_busy = False
        if hasattr(self, "inbox_search_btn"):
            self.inbox_search_btn.setEnabled(True)
        if error:
            self.inbox_search_status.setText(f"Scan failed: {error}")
            self.inbox_search_status.setStyleSheet(f"color: {PALETTE['err']};")
            self.status_bar.showMessage(f"Inbox search failed: {error}")
            self.logger.error("Inbox search failed: %s", error)
            return
        payload = payload_obj if isinstance(payload_obj, dict) else {}
        hits = list(payload.get("hits") or [])
        self._inbox_search_hits = hits
        self._populate_inbox_search_tree(payload)
        n = len(hits)
        mails = int(payload.get("mail_count") or 0)
        threads = int(payload.get("thread_count") or 0)
        note = str(payload.get("note") or "").strip()
        q = payload.get("query") or {}
        bits = [
            f"{n} conversation(s)",
            f"{mails} message(s) scraped",
            f"{threads} thread(s) grouped",
        ]
        if payload.get("used_ollama"):
            bits.append("Ollama ranked")
        sent = ", ".join(q.get("sentiments") or []) or "any tone"
        topics = ", ".join(q.get("topics") or []) or "open topic"
        summary = f"{' · '.join(bits)} — looking for {sent}; {topics}"
        if note:
            summary += f" — {note}"
        self.inbox_search_status.setText(summary)
        color = PALETTE["ok"] if n else PALETTE["muted"]
        self.inbox_search_status.setStyleSheet(f"color: {color};")
        self.status_bar.showMessage(summary)

    def _clear_inbox_search(self) -> None:
        if hasattr(self, "inbox_search_query"):
            self.inbox_search_query.clear()
        if hasattr(self, "inbox_search_tree"):
            self.inbox_search_tree.clear()
        self._inbox_search_hits = []
        if hasattr(self, "inbox_search_preview"):
            self.inbox_search_preview.setHtml(
                "<i style='color:#5a7280'>Cleared.</i>"
            )
        if hasattr(self, "inbox_search_status"):
            self.inbox_search_status.setText("Cleared.")
            self.inbox_search_status.setStyleSheet(f"color: {PALETTE['muted']};")

    def _inbox_search_tone_color(self, sentiment: str) -> QColor:
        key = (sentiment or "neutral").lower()
        mapping = {
            "angry": PALETTE["err"],
            "frustrated": "#b54708",
            "negative": PALETTE["err"],
            "positive": PALETTE["ok"],
            "grateful": PALETTE["ok"],
            "urgent": "#b54708",
            "worried": "#7a5b16",
            "confused": PALETTE["muted"],
            "mixed": PALETTE["accent"],
            "neutral": PALETTE["muted"],
            "polite": PALETTE["accent"],
        }
        return QColor(mapping.get(key, PALETTE["muted"]))

    def _populate_inbox_search_tree(self, payload: Dict[str, Any]) -> None:
        tree = self.inbox_search_tree
        tree.clear()
        hits = list(payload.get("hits") or [])
        for hit in hits:
            parent = QTreeWidgetItem(tree)
            people = ", ".join(hit.get("participants") or [])
            why = "; ".join(hit.get("reasons") or [])[:180]
            parent.setText(0, str(hit.get("topic") or "(no subject)"))
            parent.setText(1, str(hit.get("sentiment") or "neutral"))
            parent.setText(2, people)
            parent.setText(3, str(hit.get("message_count") or len(hit.get("messages") or [])))
            parent.setText(4, str(hit.get("latest") or ""))
            parent.setText(5, f"{float(hit.get('score') or 0) * 100:.0f}")
            parent.setText(6, why)
            parent.setData(0, Qt.ItemDataRole.UserRole, {"type": "conversation", "hit": hit})
            tone = self._inbox_search_tone_color(str(hit.get("sentiment") or ""))
            parent.setForeground(1, tone)
            for mail in hit.get("messages") or []:
                child = QTreeWidgetItem(parent)
                sender = str(mail.get("sender") or "")
                subject = str(mail.get("subject") or "")
                folder = str(mail.get("folder") or "")
                unread = " • unread" if mail.get("unread") else ""
                child.setText(0, f"{sender} — {subject}")
                child.setText(1, folder + unread)
                child.setText(2, str(mail.get("to") or "")[:80])
                child.setText(3, "")
                child.setText(4, str(mail.get("received") or ""))
                child.setText(5, "")
                child.setText(6, "")
                child.setData(
                    0,
                    Qt.ItemDataRole.UserRole,
                    {"type": "message", "hit": hit, "mail": mail},
                )
        if hits:
            first = tree.topLevelItem(0)
            if first is not None:
                tree.setCurrentItem(first)
                first.setExpanded(True)
                self._on_inbox_search_item_selected(first, 0)

    def _inbox_search_item_payload(self, item) -> Dict[str, Any]:
        data = item.data(0, Qt.ItemDataRole.UserRole) if item is not None else None
        return data if isinstance(data, dict) else {}

    def _on_inbox_search_item_selected(self, item, _column=None) -> None:
        payload = self._inbox_search_item_payload(item)
        hit = payload.get("hit") or {}
        mail = payload.get("mail") if payload.get("type") == "message" else None
        self.inbox_search_preview.setHtml(self._inbox_search_preview_html(hit, mail))

    def _inbox_search_preview_html(
        self,
        hit: Dict[str, Any],
        mail: Optional[Dict[str, Any]] = None,
    ) -> str:
        import html as html_mod

        if not hit:
            return "<i style='color:#5a7280'>No conversation selected.</i>"
        topic = html_mod.escape(str(hit.get("topic") or ""))
        sentiment = html_mod.escape(str(hit.get("sentiment") or "neutral"))
        people = html_mod.escape(", ".join(hit.get("participants") or []))
        matter = html_mod.escape(", ".join(hit.get("subject_matter") or []))
        reasons = "".join(
            f"<li>{html_mod.escape(str(r))}</li>" for r in (hit.get("reasons") or [])
        )
        score = float(hit.get("score") or 0) * 100
        parts = [
            f"<div style='font-size:14px;font-weight:700;color:{PALETTE['accent']}'>"
            f"{topic}</div>",
            f"<div style='color:{PALETTE['muted']};margin:4px 0 8px'>"
            f"Tone: <b>{sentiment}</b> · Score {score:.0f} · "
            f"{int(hit.get('message_count') or 0)} message(s)<br>"
            f"People: {people or '—'}<br>"
            f"Subject matter: {matter or '—'}"
            "</div>",
        ]
        if reasons:
            parts.append(f"<div><b>Why this thread</b><ul>{reasons}</ul></div>")
        parts.append("<div><b>Conversation</b></div>")
        highlight_id = str((mail or {}).get("entry_id") or "")
        for msg in hit.get("messages") or []:
            eid = str(msg.get("entry_id") or "")
            border = (
                f"border-left:3px solid {PALETTE['accent']};"
                if highlight_id and eid == highlight_id
                else f"border-left:3px solid {PALETTE['border']};"
            )
            body = html_mod.escape(
                re.sub(r"\s+", " ", str(msg.get("body_preview") or "")).strip()[:900]
            )
            sender = html_mod.escape(str(msg.get("sender") or ""))
            when = html_mod.escape(str(msg.get("received") or ""))
            subj = html_mod.escape(str(msg.get("subject") or ""))
            folder = html_mod.escape(str(msg.get("folder") or ""))
            parts.append(
                "<div style='margin:8px 0;padding:6px 8px;"
                f"{border}background:{PALETTE['surface']}'>"
                f"<div style='color:{PALETTE['muted']};font-size:11px'>"
                f"{when} · {folder} · {sender}</div>"
                f"<div style='font-weight:600'>{subj}</div>"
                f"<div style='margin-top:4px;white-space:pre-wrap'>{body or '<i>(no body)</i>'}</div>"
                "</div>"
            )
        return "".join(parts)

    def _open_inbox_search_item(self, item) -> None:
        payload = self._inbox_search_item_payload(item)
        mail = payload.get("mail")
        if not isinstance(mail, dict):
            hit = payload.get("hit") or {}
            messages = list(hit.get("messages") or [])
            mail = messages[-1] if messages else {}
        entry_id = str((mail or {}).get("entry_id") or "")
        store_id = str((mail or {}).get("store_id") or "")
        if not entry_id:
            QMessageBox.information(
                self,
                "Open in Outlook",
                "This row has no Outlook EntryID.",
            )
            return
        if not open_mail_in_outlook(entry_id, store_id):
            QMessageBox.warning(self, "Open in Outlook", "Could not open the message in Outlook.")

    def _open_inbox_search_in_outlook(self) -> None:
        item = self.inbox_search_tree.currentItem() if hasattr(self, "inbox_search_tree") else None
        if item is None:
            self.status_bar.showMessage("Select a conversation or message first")
            return
        self._open_inbox_search_item(item)

    # ------------------------------------------------------------------ #
    # AES Scores tab
    # ------------------------------------------------------------------ #

    _AES_SCORE_COLUMNS = ("Sender", "Domain", "Stored score", "Emails", "First seen", "Last seen")

    def _create_aes_scores_tab(self) -> None:
        """Per-sender smoothed threat scores the AES scanner remembers."""
        frame = QWidget()
        self.aes_scores_tab_index = self.notebook.addTab(frame, "AES Scores")
        root = QVBoxLayout(frame)
        root.setContentsMargins(12, 12, 12, 12)
        root.setSpacing(10)

        head = QLabel("AES sender threat scores")
        head.setStyleSheet(f"font-size: 14px; font-weight: 700; color: {PALETTE['accent']};")
        root.addWidget(head)

        sub = QLabel(
            "Each scan's risk score is blended with the score stored here for the sender "
            "(scores of 70+ always show as-is). If a bad scan inflated a sender, set a "
            "corrected score, or reset it so the next scan starts fresh from its raw score. "
            "Resetting a domain also clears cached threat-intel verdicts for it."
        )
        sub.setWordWrap(True)
        sub.setStyleSheet(f"color: {PALETTE['muted']};")
        root.addWidget(sub)

        toolbar = QHBoxLayout()
        self.aes_scores_filter = QLineEdit()
        self.aes_scores_filter.setPlaceholderText("Filter by sender or domain…")
        self.aes_scores_filter.setClearButtonEnabled(True)
        self.aes_scores_filter.textChanged.connect(self._aes_scores_apply_filter)
        toolbar.addWidget(self.aes_scores_filter, stretch=1)
        refresh_btn = QPushButton("Refresh")
        refresh_btn.clicked.connect(self._aes_scores_reload)
        toolbar.addWidget(refresh_btn)
        root.addLayout(toolbar)

        self.aes_scores_tree = QTreeWidget()
        self.aes_scores_tree.setHeaderLabels(self._AES_SCORE_COLUMNS)
        self.aes_scores_tree.setRootIsDecorated(False)
        self.aes_scores_tree.setSortingEnabled(True)
        self.aes_scores_tree.setSelectionMode(QTreeWidget.SelectionMode.ExtendedSelection)
        self.aes_scores_tree.setColumnWidth(0, 260)
        self.aes_scores_tree.setColumnWidth(1, 180)
        self.aes_scores_tree.setColumnWidth(2, 100)
        self.aes_scores_tree.setColumnWidth(3, 70)
        self.aes_scores_tree.setColumnWidth(4, 140)
        self.aes_scores_tree.itemDoubleClicked.connect(lambda *_: self._aes_scores_set_selected())
        root.addWidget(self.aes_scores_tree, stretch=1)

        actions = QHBoxLayout()
        set_btn = QPushButton("Set score…")
        set_btn.setToolTip("Override the stored score for the selected senders.")
        set_btn.clicked.connect(self._aes_scores_set_selected)
        reset_btn = QPushButton("Reset score")
        reset_btn.setToolTip(
            "Clear the stored score so the next scan shows its raw score unsmoothed."
        )
        reset_btn.clicked.connect(self._aes_scores_reset_selected)
        domain_btn = QPushButton("Reset domain…")
        domain_btn.setToolTip(
            "Reset every sender at the selected sender's domain and clear its "
            "cached threat-intel verdicts."
        )
        domain_btn.clicked.connect(self._aes_scores_reset_domain)
        forget_btn = QPushButton("Forget sender…")
        forget_btn.setToolTip(
            "Delete the selected senders' history entirely (email count and score); "
            "they are treated as unknown senders again."
        )
        forget_btn.clicked.connect(self._aes_scores_forget_selected)
        actions.addWidget(set_btn)
        actions.addWidget(reset_btn)
        actions.addWidget(domain_btn)
        actions.addStretch(1)
        actions.addWidget(forget_btn)
        root.addLayout(actions)

        self.aes_scores_status = QLabel("")
        self.aes_scores_status.setStyleSheet(f"color: {PALETTE['muted']};")
        root.addWidget(self.aes_scores_status)

        self._aes_scores_reload()

    def _aes_scores_reload(self) -> None:
        tree = self.aes_scores_tree
        try:
            rows = aes_score_history.load_sender_scores()
        except Exception as exc:
            tree.clear()
            self.aes_scores_status.setText(f"Could not read {aes_score_history.history_path()}: {exc}")
            return

        sort_col = tree.sortColumn()
        sort_order = tree.header().sortIndicatorOrder()
        tree.setSortingEnabled(False)
        tree.clear()
        for row in rows:
            item = _SortableTreeItem(
                [
                    row.email,
                    row.domain,
                    "—" if row.score is None else str(row.score),
                    str(row.count),
                    row.first_seen,
                    row.last_seen,
                ]
            )
            item.setData(0, Qt.ItemDataRole.UserRole, row.email)
            item.setData(2, _SortableTreeItem.SORT_ROLE, -1 if row.score is None else row.score)
            item.setData(3, _SortableTreeItem.SORT_ROLE, row.count)
            item.setTextAlignment(2, Qt.AlignmentFlag.AlignCenter)
            item.setTextAlignment(3, Qt.AlignmentFlag.AlignCenter)
            if row.score is not None:
                colour = (
                    PALETTE["err"] if row.score >= 70
                    else "#b26a00" if row.score >= 30
                    else PALETTE["ok"]
                )
                item.setForeground(2, QBrush(QColor(colour)))
                font = item.font(2)
                font.setBold(True)
                item.setFont(2, font)
            tree.addTopLevelItem(item)
        tree.setSortingEnabled(True)
        if sort_col < 0:
            tree.sortItems(2, Qt.SortOrder.DescendingOrder)
        else:
            tree.sortItems(sort_col, sort_order)

        scored = sum(1 for r in rows if r.score is not None)
        self.aes_scores_status.setText(
            f"{len(rows)} sender(s), {scored} with a stored score · "
            f"{aes_score_history.history_path()}"
        )
        self._aes_scores_apply_filter(self.aes_scores_filter.text())

    def _aes_scores_apply_filter(self, text: str) -> None:
        needle = (text or "").strip().lower()
        tree = self.aes_scores_tree
        for i in range(tree.topLevelItemCount()):
            item = tree.topLevelItem(i)
            hay = f"{item.text(0)} {item.text(1)}".lower()
            item.setHidden(bool(needle) and needle not in hay)

    def _aes_scores_selected_emails(self) -> List[str]:
        emails = [
            str(item.data(0, Qt.ItemDataRole.UserRole) or "")
            for item in self.aes_scores_tree.selectedItems()
        ]
        emails = [e for e in emails if e]
        if not emails:
            self.status_bar.showMessage("Select one or more senders first")
        return emails

    def _aes_scores_run(self, fn, *args):
        try:
            return fn(*args)
        except aes_score_history.HistoryLockedError as exc:
            QMessageBox.warning(self, "AES Scores", str(exc))
        except Exception as exc:
            QMessageBox.critical(self, "AES Scores", f"Update failed: {exc}")
        return None

    def _aes_scores_set_selected(self) -> None:
        emails = self._aes_scores_selected_emails()
        if not emails:
            return
        current = self.aes_scores_tree.selectedItems()[0].text(2)
        start = int(current) if current.isdigit() else 0
        label = emails[0] if len(emails) == 1 else f"{len(emails)} senders"
        value, ok = QInputDialog.getInt(
            self, "Set stored score", f"Stored score for {label} (0–100):", start, 0, 100, 1
        )
        if not ok:
            return
        changed = self._aes_scores_run(aes_score_history.set_scores, emails, value)
        if changed is not None:
            self.status_bar.showMessage(f"Set stored score to {value} for {changed} sender(s)")
            self._aes_scores_reload()

    def _aes_scores_reset_selected(self) -> None:
        emails = self._aes_scores_selected_emails()
        if not emails:
            return
        changed = self._aes_scores_run(aes_score_history.reset_scores, emails)
        if changed is not None:
            self.status_bar.showMessage(f"Reset stored score for {changed} sender(s)")
            self._aes_scores_reload()

    def _aes_scores_reset_domain(self) -> None:
        items = self.aes_scores_tree.selectedItems()
        default = items[0].text(1) if items else self.aes_scores_filter.text().strip()
        domain, ok = QInputDialog.getText(
            self, "Reset domain", "Domain to reset (includes subdomains):", text=default
        )
        domain = (domain or "").strip().lower().lstrip("@")
        if not ok or not domain:
            return
        emails = self._aes_scores_run(aes_score_history.senders_at_domain, domain)
        if emails is None:
            return
        confirm = QMessageBox.question(
            self,
            "Reset domain",
            f"Reset stored scores for {len(emails)} sender(s) at {domain} and clear "
            "its cached threat-intel verdicts?",
        )
        if confirm != QMessageBox.StandardButton.Yes:
            return
        changed = self._aes_scores_run(aes_score_history.reset_scores, emails) if emails else 0
        cleared = self._aes_scores_run(aes_score_history.clear_threat_intel_for_domain, domain)
        if changed is None or cleared is None:
            return
        self.status_bar.showMessage(
            f"{domain}: reset {changed} sender score(s), cleared {cleared} threat-intel entr(ies)"
        )
        self._aes_scores_reload()

    def _aes_scores_forget_selected(self) -> None:
        emails = self._aes_scores_selected_emails()
        if not emails:
            return
        confirm = QMessageBox.question(
            self,
            "Forget sender",
            f"Delete all history for {len(emails)} sender(s)? Their email count is lost "
            "and they will be treated as unknown senders again.",
        )
        if confirm != QMessageBox.StandardButton.Yes:
            return
        changed = self._aes_scores_run(aes_score_history.forget_senders, emails)
        if changed is not None:
            self.status_bar.showMessage(f"Forgot {changed} sender(s)")
            self._aes_scores_reload()

    def _create_data_brokers_tab(self) -> None:
        """Incogni-style data-broker opt-out / erasure request tracking."""
        panel = DataBrokersPanel(
            self,
            status_fn=lambda m: self.status_bar.showMessage(m),
            accent=PALETTE.get("accent", "#0f6b7c"),
            muted=PALETTE.get("muted", "#5a7280"),
        )
        self.data_brokers_tab_index = self.notebook.addTab(panel, "Aura")
        self.data_brokers_panel = panel

    def show_aura_tab(self) -> None:
        """Select the Aura (data-broker removal) tab and bring GURI forward."""
        try:
            idx = int(getattr(self, "data_brokers_tab_index", -1))
            if idx >= 0:
                self.notebook.setCurrentIndex(idx)
            panel = getattr(self, "data_brokers_panel", None)
            if panel is not None and hasattr(panel, "reload_all"):
                panel.reload_all()
        except Exception:
            pass
        try:
            self.expand_to_screen()
        except Exception:
            pass

    def _focus_locations_tab(self) -> None:
        if hasattr(self, "locations_tab_index"):
            self.notebook.setCurrentIndex(self.locations_tab_index)
            return
        for i in range(self.notebook.count()):
            if self.notebook.tabText(i) == "Locations":
                self.notebook.setCurrentIndex(i)
                break

    def _create_locations_tab(self):
        """Filesystem roots GURI scrapes for document inventory + deadline cues."""
        frame = QWidget()
        self.locations_tab_index = self.notebook.addTab(frame, "Locations")
        root = QVBoxLayout(frame)
        root.setContentsMargins(12, 12, 12, 12)
        root.setSpacing(10)

        head = QLabel("Scan locations")
        head.setStyleSheet(
            f"font-size: 14px; font-weight: 700; color: {PALETTE['accent']};"
        )
        root.addWidget(head)

        sub = QLabel(
            "Add folders GURI should inventory. Scraping is read-only: GURI never "
            "writes to, renames, or stamps your files. Tagging creates an "
            "internal GURI record in the database that references the path only. "
            "Office/PDF are listed by type; readable text (.txt, .md, .csv, …) "
            "is also checked for deadline cues."
        )
        sub.setWordWrap(True)
        sub.setStyleSheet(f"color: {PALETTE['muted']};")
        root.addWidget(sub)

        split = QSplitter(Qt.Orientation.Horizontal)
        root.addWidget(split, stretch=1)

        left = QWidget()
        left_layout = QVBoxLayout(left)
        left_layout.setContentsMargins(0, 0, 0, 0)
        left_layout.setSpacing(8)

        loc_toolbar = QHBoxLayout()
        add_btn = QPushButton("Add folder…")
        add_btn.clicked.connect(self._locations_add_folder)
        remove_btn = QPushButton("Remove")
        remove_btn.clicked.connect(self._locations_remove_selected)
        save_btn = QPushButton("Save")
        save_btn.clicked.connect(self._locations_save)
        loc_toolbar.addWidget(add_btn)
        loc_toolbar.addWidget(remove_btn)
        loc_toolbar.addStretch(1)
        loc_toolbar.addWidget(save_btn)
        left_layout.addLayout(loc_toolbar)

        self.locations_tree = QTreeWidget()
        self.locations_tree.setHeaderLabels(("Scan", "Recursive", "Folder"))
        self.locations_tree.setRootIsDecorated(False)
        self.locations_tree.setColumnWidth(0, 50)
        self.locations_tree.setColumnWidth(1, 80)
        left_layout.addWidget(self.locations_tree, stretch=1)

        scrape_row = QHBoxLayout()
        scrape_btn = QPushButton("Scrape Files")
        scrape_btn.setToolTip(
            "Inventory enabled folders (read-only — does not modify any file)."
        )
        scrape_btn.clicked.connect(self._start_file_scrape)
        self.file_scrape_btn = scrape_btn
        open_btn = QPushButton("Open folder")
        open_btn.clicked.connect(self._locations_open_selected_folder)
        guri_btn = QPushButton("Tag with GURI")
        guri_btn.setToolTip(
            "Create an internal GURI database tag for the selected files. "
            "Does not touch or modify the files on disk."
        )
        guri_btn.clicked.connect(self._locations_create_guris)
        scrape_row.addWidget(scrape_btn)
        scrape_row.addWidget(open_btn)
        scrape_row.addStretch(1)
        scrape_row.addWidget(guri_btn)
        left_layout.addLayout(scrape_row)

        self.locations_status = QLabel("No locations yet — add a folder to begin.")
        self.locations_status.setStyleSheet(f"color: {PALETTE['muted']};")
        left_layout.addWidget(self.locations_status)

        split.addWidget(left)

        right = QWidget()
        right_layout = QVBoxLayout(right)
        right_layout.setContentsMargins(0, 0, 0, 0)
        results_lbl = QLabel("Last scrape")
        results_lbl.setStyleSheet(f"font-weight: 600; color: {PALETTE['text']};")
        right_layout.addWidget(results_lbl)

        self.file_scrape_tree = QTreeWidget()
        self.file_scrape_tree.setHeaderLabels(
            ("Name", "Type", "Size", "Modified", "Deadline", "Path")
        )
        self.file_scrape_tree.setRootIsDecorated(False)
        self.file_scrape_tree.setSelectionMode(
            QTreeWidget.SelectionMode.ExtendedSelection
        )
        self.file_scrape_tree.setColumnWidth(0, 180)
        self.file_scrape_tree.setColumnWidth(1, 120)
        self.file_scrape_tree.setColumnWidth(5, 280)
        right_layout.addWidget(self.file_scrape_tree, stretch=1)
        split.addWidget(right)
        split.setStretchFactor(0, 2)
        split.setStretchFactor(1, 3)

        self._locations_reload_tree()
        self._locations_populate_results(self.file_scrape_items)

    def _locations_reload_tree(self) -> None:
        if not hasattr(self, "locations_tree"):
            return
        self.locations_tree.clear()
        locs = list(getattr(self, "_scan_locations", None) or [])
        for loc in locs:
            item = QTreeWidgetItem()
            item.setFlags(
                item.flags()
                | Qt.ItemFlag.ItemIsUserCheckable
                | Qt.ItemFlag.ItemIsSelectable
                | Qt.ItemFlag.ItemIsEnabled
            )
            item.setCheckState(
                0,
                Qt.CheckState.Checked if loc.enabled else Qt.CheckState.Unchecked,
            )
            item.setCheckState(
                1,
                Qt.CheckState.Checked if loc.recursive else Qt.CheckState.Unchecked,
            )
            item.setText(2, loc.path)
            item.setData(0, Qt.ItemDataRole.UserRole, loc.to_dict())
            self.locations_tree.addTopLevelItem(item)
        n_on = sum(1 for loc in locs if loc.enabled)
        if hasattr(self, "locations_status"):
            self.locations_status.setText(
                f"{len(locs)} location(s) · {n_on} enabled"
                if locs
                else "No locations yet — add a folder to begin."
            )

    def _locations_collect_from_tree(self) -> List[ScanLocation]:
        out: List[ScanLocation] = []
        if not hasattr(self, "locations_tree"):
            return list(getattr(self, "_scan_locations", None) or [])
        for i in range(self.locations_tree.topLevelItemCount()):
            raw = self.locations_tree.topLevelItem(i)
            if not isinstance(raw, QTreeWidgetItem):
                continue
            item = raw
            path = (item.text(2) or "").strip()
            if not path:
                continue
            out.append(
                ScanLocation(
                    path=path,
                    enabled=item.checkState(0) == Qt.CheckState.Checked,
                    recursive=item.checkState(1) == Qt.CheckState.Checked,
                )
            )
        return out

    def _locations_add_folder(self) -> None:
        folder = QFileDialog.getExistingDirectory(
            self, "Add scan location", str(Path.home())
        )
        if not folder:
            return
        locs = self._locations_collect_from_tree()
        key = os.path.normcase(os.path.abspath(folder))
        if any(os.path.normcase(os.path.abspath(l.path)) == key for l in locs):
            QMessageBox.information(self, "Locations", "That folder is already listed.")
            return
        locs.append(ScanLocation(path=folder, enabled=True, recursive=True))
        self._scan_locations = locs
        self._locations_reload_tree()
        self.locations_status.setText(f"Added {folder} (not saved yet)")

    def _locations_remove_selected(self) -> None:
        items = self.locations_tree.selectedItems()
        if not items:
            return
        for item in items:
            idx = self.locations_tree.indexOfTopLevelItem(item)
            if idx >= 0:
                self.locations_tree.takeTopLevelItem(idx)
        self._scan_locations = self._locations_collect_from_tree()
        self.locations_status.setText("Removed — click Save to persist")

    def _locations_save(self) -> None:
        locs = self._locations_collect_from_tree()
        try:
            path = save_scan_locations(locs)
            self._scan_locations = locs
            n_on = sum(1 for l in locs if l.enabled)
            self.locations_status.setText(
                f"Saved {len(locs)} location(s) ({n_on} enabled) → {path}"
            )
            self.locations_status.setStyleSheet(f"color: {PALETTE['ok']};")
            self.status_bar.showMessage(
                f"Scan locations saved ({n_on}/{len(locs)} enabled)"
            )
        except Exception as exc:
            QMessageBox.critical(self, "Locations", f"Failed to save:\n{exc}")

    def _locations_open_selected_folder(self) -> None:
        items = self.locations_tree.selectedItems()
        path = ""
        if items:
            path = (items[0].text(2) or "").strip()
        if not path:
            sel = self.file_scrape_tree.selectedItems() if hasattr(self, "file_scrape_tree") else []
            if sel:
                path = str(sel[0].data(0, Qt.ItemDataRole.UserRole) or "")
                if path:
                    path = str(Path(path).parent)
        if not path or not os.path.isdir(path):
            QMessageBox.information(self, "Locations", "Select a folder or scraped file first.")
            return
        try:
            os.startfile(path)  # type: ignore[attr-defined]
        except Exception as exc:
            QMessageBox.warning(self, "Locations", f"Could not open:\n{exc}")

    def _locations_populate_results(self, items: Optional[List[Dict[str, Any]]]) -> None:
        if not hasattr(self, "file_scrape_tree"):
            return
        self.file_scrape_tree.clear()
        deadline_n = 0
        for det in items or []:
            if not isinstance(det, dict):
                continue
            file_meta = det.get("file") or {}
            mail = det.get("mail") or {}
            name = str(file_meta.get("name") or mail.get("subject") or "")
            doc_code = str(
                file_meta.get("document_type") or mail.get("document_type") or "99"
            )
            doc_name = DOCUMENT_TYPES.get(doc_code, doc_code)
            size = int(file_meta.get("size") or 0)
            if size >= 1024 * 1024:
                size_s = f"{size / (1024 * 1024):.1f} MB"
            elif size >= 1024:
                size_s = f"{size / 1024:.1f} KB"
            else:
                size_s = f"{size} B"
            modified = str(file_meta.get("modified") or mail.get("received") or "")
            dls = det.get("deadlines") or []
            due = ""
            if dls:
                deadline_n += 1
                due = str((dls[0] or {}).get("due") or (dls[0] or {}).get("cue") or "yes")
            path = str(file_meta.get("path") or mail.get("file_path") or "")
            row = QTreeWidgetItem(
                (name, doc_name, size_s, modified[:19], due, path)
            )
            row.setData(0, Qt.ItemDataRole.UserRole, path)
            row.setData(1, Qt.ItemDataRole.UserRole, det)
            self.file_scrape_tree.addTopLevelItem(row)
        scraped_at = ""
        try:
            scraped_at = str(load_file_scrape_cache().get("scraped_at") or "")
        except Exception:
            pass
        n = len(items or [])
        msg = f"Last scrape: {n} file(s)"
        if deadline_n:
            msg += f" · {deadline_n} with deadline cues"
        if scraped_at:
            msg += f" · {scraped_at}"
        if hasattr(self, "locations_status") and not self._file_scrape_busy:
            # Keep save/add messages unless scrape just finished (caller updates)
            pass
        self.status_bar.showMessage(msg)

    def _on_file_scrape_progress(self, message: str) -> None:
        if hasattr(self, "locations_status"):
            self.locations_status.setText(message)
            self.locations_status.setStyleSheet(f"color: {PALETTE['muted']};")
        self.status_bar.showMessage(message)

    def _start_file_scrape(self) -> None:
        if self._file_scrape_busy:
            self.status_bar.showMessage("File scrape already running…")
            return
        locs = self._locations_collect_from_tree()
        if not locs:
            try:
                locs = load_scan_locations()
            except Exception:
                locs = []
        enabled = [l for l in locs if l.enabled]
        if not enabled:
            QMessageBox.information(
                self,
                "Locations",
                "Add and enable at least one folder on the Locations tab, then Save.",
            )
            self._focus_locations_tab()
            return
        # Persist current tree so scrape matches what the user sees
        try:
            save_scan_locations(locs)
            self._scan_locations = locs
        except Exception as exc:
            self.logger.warning("Could not save locations before scrape: %s", exc)

        self._file_scrape_busy = True
        if hasattr(self, "file_scrape_btn"):
            self.file_scrape_btn.setEnabled(False)
        self.status_bar.showMessage("Scraping file locations…")
        if hasattr(self, "locations_status"):
            self.locations_status.setText("Scraping…")
        cues = list(self._learning_deadline_cues())

        def work() -> None:
            try:
                def progress(msg: str) -> None:
                    self._file_scrape_bridge.progress.emit(msg)

                files = scrape_file_locations(
                    enabled,
                    extract_text=True,
                    extra_deadline_cues=cues,
                    progress=progress,
                )
                items = scraped_files_to_detection_items(files)
                roots = [l.path for l in enabled]
                try:
                    save_file_scrape_cache(items, roots=roots)
                except Exception as exc:
                    self.logger.warning("Failed saving file scrape cache: %s", exc)
                self._file_scrape_bridge.finished.emit(items, "")
            except Exception as exc:
                self._file_scrape_bridge.finished.emit(None, str(exc))

        threading.Thread(target=work, daemon=True).start()

    def _on_file_scrape_finished(self, items_obj, error: str) -> None:
        self._file_scrape_busy = False
        if hasattr(self, "file_scrape_btn"):
            self.file_scrape_btn.setEnabled(True)
        if error:
            self.status_bar.showMessage(f"File scrape failed: {error}")
            if hasattr(self, "locations_status"):
                self.locations_status.setText(f"Scrape failed: {error}")
                self.locations_status.setStyleSheet(f"color: {PALETTE['err']};")
            self.logger.error("File scrape failed: %s", error)
            return
        items: List[Dict[str, Any]] = list(items_obj or [])
        self.file_scrape_items = items
        self._locations_populate_results(items)
        deadline_n = sum(1 for it in items if it.get("deadlines"))
        msg = f"File scrape: {len(items)} file(s)"
        if deadline_n:
            msg += f", {deadline_n} with deadline cues"
        if hasattr(self, "locations_status"):
            self.locations_status.setText(msg)
            self.locations_status.setStyleSheet(f"color: {PALETTE['ok']};")
        self.status_bar.showMessage(msg)
        try:
            self._refresh_deadlines_panel()
        except Exception:
            pass

    def _locations_create_guris(self) -> None:
        """Create internal GURI DB tags for selected files — never modifies files."""
        if not self.db:
            QMessageBox.warning(self, "Locations", "Connect to the database first.")
            return
        selected = (
            self.file_scrape_tree.selectedItems()
            if hasattr(self, "file_scrape_tree")
            else []
        )
        if not selected:
            QMessageBox.information(
                self, "Locations", "Select one or more scraped files in the results list."
            )
            return
        reply = QMessageBox.question(
            self,
            "Tag with GURI",
            f"Create internal GURI tags for {len(selected)} selected file(s)?\n\n"
            "This writes only to the GURI database.\n"
            "It does not modify, rename, or stamp the files on disk.",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.Yes,
        )
        if reply != QMessageBox.StandardButton.Yes:
            return
        created = 0
        errors = 0
        for item in selected:
            det = item.data(1, Qt.ItemDataRole.UserRole) or {}
            file_meta = (det.get("file") if isinstance(det, dict) else None) or {}
            path = str(file_meta.get("path") or item.data(0, Qt.ItemDataRole.UserRole) or "")
            name = str(file_meta.get("name") or Path(path).name)
            # Subject stores the path reference so the tag points at the file
            # without writing anything into the file itself.
            subject = path or name
            doc_type = str(file_meta.get("document_type") or "99")
            modified = str(file_meta.get("modified") or datetime.now().isoformat(timespec="seconds"))
            dt_str = modified.replace("T", " ")[:19] if modified else datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            try:
                guri = self.db.generate_guri(
                    NON_EMAIL_SENDER,
                    NON_EMAIL_PLACEHOLDER,
                    subject,
                    dt_str,
                    "0",
                    document_type=doc_type,
                )
                if guri:
                    created += 1
            except Exception as exc:
                errors += 1
                self.logger.warning("GURI tag failed for %s: %s", path, exc)
        self.status_bar.showMessage(
            f"Tagged {created} file(s) with GURI (DB only)"
            + (f" · {errors} error(s)" if errors else "")
        )
        if created:
            try:
                self._load_records()
            except Exception:
                pass
            QMessageBox.information(
                self,
                "Locations",
                f"Created {created} internal GURI tag(s) in the database.\n"
                "Source files were not modified."
                + (f"\n{errors} failed." if errors else ""),
            )

    def _create_ollama_tab(self):
        """Ollama local LLM: connection, models, prompt, email analysis."""
        frame = QWidget()
        self.notebook.addTab(frame, "Ollama")
        root = QVBoxLayout(frame)
        root.setContentsMargins(12, 12, 12, 12)
        root.setSpacing(8)

        cfg = self.ollama_cfg

        head = QLabel("Ollama")
        head.setStyleSheet("font-size: 14px; font-weight: 700; color: #0f6b7c;")
        root.addWidget(head)
        sub = QLabel(
            "Talk to a local Ollama server for email summaries and action extraction. "
            "Default endpoint: http://localhost:11434"
        )
        sub.setWordWrap(True)
        sub.setStyleSheet("color: #4a6570; font-size: 11px;")
        root.addWidget(sub)

        form = QGridLayout()
        form.setHorizontalSpacing(8)
        form.setVerticalSpacing(6)

        form.addWidget(QLabel("Base URL"), 0, 0)
        self.ollama_url_edit = QLineEdit(str(cfg.get("base_url") or "http://localhost:11434"))
        form.addWidget(self.ollama_url_edit, 0, 1, 1, 3)

        form.addWidget(QLabel("Model"), 1, 0)
        self.ollama_model_combo = QComboBox()
        self.ollama_model_combo.setEditable(True)
        if cfg.get("model"):
            self.ollama_model_combo.addItem(str(cfg.get("model")))
            self.ollama_model_combo.setCurrentText(str(cfg.get("model")))
        form.addWidget(self.ollama_model_combo, 1, 1, 1, 2)

        refresh_models_btn = QPushButton("Refresh models")
        refresh_models_btn.clicked.connect(self._ollama_refresh_models)
        form.addWidget(refresh_models_btn, 1, 3)

        form.addWidget(QLabel("System"), 2, 0)
        self.ollama_system_edit = QTextEdit()
        self.ollama_system_edit.setPlainText(str(cfg.get("system") or ""))
        self.ollama_system_edit.setMaximumHeight(72)
        form.addWidget(self.ollama_system_edit, 2, 1, 1, 3)

        root.addLayout(form)

        status_row = QHBoxLayout()
        self.ollama_status = QLabel("Status: not checked")
        self.ollama_status.setStyleSheet("color: #5a7280;")
        test_btn = QPushButton("Test connection")
        test_btn.clicked.connect(self._ollama_test_connection)
        gpu_btn = QPushButton("Test GPU / CUDA")
        gpu_btn.setToolTip(
            "Checks NVIDIA GPU + CUDA driver (nvidia-smi), optional nvcc toolkit, "
            "then warms the selected model and verifies Ollama VRAM via /api/ps."
        )
        gpu_btn.clicked.connect(self._ollama_test_gpu)
        self.ollama_gpu_btn = gpu_btn
        save_btn = QPushButton("Save settings")
        save_btn.clicked.connect(self._ollama_save_settings)
        status_row.addWidget(self.ollama_status, stretch=1)
        status_row.addWidget(test_btn)
        status_row.addWidget(gpu_btn)
        status_row.addWidget(save_btn)
        root.addLayout(status_row)

        root.addWidget(QLabel("Prompt"))
        self.ollama_prompt = QTextEdit()
        self.ollama_prompt.setPlaceholderText(
            "Ask Ollama anything, or use “Analyze selected email” from the Welcome preview."
        )
        self.ollama_prompt.setMinimumHeight(100)
        root.addWidget(self.ollama_prompt)

        btn_row = QHBoxLayout()
        rules_btn = QPushButton("Show saved rules")
        rules_btn.setToolTip(
            "List natural-language rules saved from Welcome → Describe a rule…"
        )
        rules_btn.clicked.connect(self._ollama_show_saved_rules)
        analyze_btn = QPushButton("Analyze selected email")
        analyze_btn.setToolTip("Uses the email currently shown in Welcome → Email preview.")
        analyze_btn.clicked.connect(self._ollama_analyze_selected_email)
        run_btn = QPushButton("Run")
        run_btn.clicked.connect(self._ollama_run_prompt)
        clear_btn = QPushButton("Clear")
        clear_btn.clicked.connect(lambda: (self.ollama_prompt.clear(), self.ollama_output.clear()))
        self.ollama_run_btn = run_btn
        btn_row.addWidget(rules_btn)
        btn_row.addWidget(analyze_btn)
        btn_row.addStretch(1)
        btn_row.addWidget(clear_btn)
        btn_row.addWidget(run_btn)
        root.addLayout(btn_row)

        root.addWidget(QLabel("Response"))
        self.ollama_output = QTextEdit()
        self.ollama_output.setReadOnly(True)
        self.ollama_output.setMinimumHeight(180)
        root.addWidget(self.ollama_output, stretch=1)

        # Initial connection probe (async)
        QTimer.singleShot(400, self._ollama_refresh_models)

    def _ollama_collect_cfg(self) -> Dict[str, Any]:
        return {
            "base_url": self.ollama_url_edit.text().strip() or "http://localhost:11434",
            "model": self.ollama_model_combo.currentText().strip(),
            "system": self.ollama_system_edit.toPlainText().strip(),
            "timeout": int(self.ollama_cfg.get("timeout") or 120),
        }

    def _ollama_save_settings(self) -> None:
        self.ollama_cfg = self._ollama_collect_cfg()
        ollama_save_config(self.ollama_cfg)
        self.ollama_status.setText("Status: settings saved")
        self.ollama_status.setStyleSheet("color: #1b7a3d;")
        self.status_bar.showMessage("Ollama settings saved")

    def _ollama_test_connection(self) -> None:
        cfg = self._ollama_collect_cfg()

        def work() -> None:
            ok, msg = ollama_ping(cfg["base_url"], timeout=5)
            if ok:
                try:
                    models = ollama_list_models(cfg["base_url"], timeout=10)
                    self._ollama_bridge.models_ready.emit(models, "")
                except Exception as exc:
                    self._ollama_bridge.models_ready.emit(None, str(exc))
            else:
                self._ollama_bridge.models_ready.emit(None, msg)

        self.ollama_status.setText("Status: checking…")
        self.ollama_status.setStyleSheet("color: #5a7280;")
        threading.Thread(target=work, daemon=True).start()

    def _ollama_refresh_models(self) -> None:
        self._ollama_test_connection()

    def _ollama_test_gpu(self) -> None:
        """Async NVIDIA / CUDA / Ollama VRAM connection test."""
        if self._ollama_gpu_busy:
            self.status_bar.showMessage("GPU test already running…")
            return
        cfg = self._ollama_collect_cfg()
        self._ollama_gpu_busy = True
        if hasattr(self, "ollama_gpu_btn"):
            self.ollama_gpu_btn.setEnabled(False)
        self.ollama_status.setText("Status: testing GPU / CUDA…")
        self.ollama_status.setStyleSheet(f"color: {PALETTE['muted']};")
        self.ollama_output.setPlainText(
            "Running GPU / CUDA test…\n"
            "· nvidia-smi (driver + CUDA)\n"
            "· nvcc toolkit (optional)\n"
            "· Ollama warm + /api/ps VRAM check\n"
        )

        def work() -> None:
            try:
                report = ollama_probe_gpu(
                    cfg["base_url"],
                    cfg.get("model") or "",
                    warm=True,
                    timeout=max(60, int(cfg.get("timeout") or 120)),
                )
                self._ollama_bridge.gpu_ready.emit(report, "")
            except Exception as exc:
                self._ollama_bridge.gpu_ready.emit(None, str(exc))

        threading.Thread(target=work, daemon=True).start()

    def _on_ollama_gpu_ready(self, report_obj, error: str) -> None:
        self._ollama_gpu_busy = False
        if hasattr(self, "ollama_gpu_btn"):
            self.ollama_gpu_btn.setEnabled(True)
        if error:
            self.ollama_output.setPlainText(f"GPU test error: {error}")
            self.ollama_status.setText(f"Status: GPU test failed — {error}")
            self.ollama_status.setStyleSheet(f"color: {PALETTE['err']};")
            return
        report = report_obj if isinstance(report_obj, dict) else {}
        text = ollama_format_gpu_report(report)
        self.ollama_output.setPlainText(text)
        summary = str(report.get("summary") or "GPU test done")
        if report.get("ollama_gpu") and report.get("gpu_ok"):
            color = PALETTE["ok"]
        elif report.get("gpu_ok"):
            color = PALETTE["accent"]
        else:
            color = PALETTE["err"]
        self.ollama_status.setText(f"Status: {summary}")
        self.ollama_status.setStyleSheet(f"color: {color};")
        self.status_bar.showMessage(summary)

    def _on_ollama_models_ready(self, models_obj, error: str) -> None:
        if error:
            self.ollama_status.setText(f"Status: {error}")
            self.ollama_status.setStyleSheet("color: #b42318;")
            return
        models = list(models_obj or [])
        current = self.ollama_model_combo.currentText().strip()
        self.ollama_model_combo.blockSignals(True)
        self.ollama_model_combo.clear()
        self.ollama_model_combo.addItems(models)
        if current:
            idx = self.ollama_model_combo.findText(current)
            if idx >= 0:
                self.ollama_model_combo.setCurrentIndex(idx)
            else:
                self.ollama_model_combo.setEditText(current)
        elif models:
            self.ollama_model_combo.setCurrentIndex(0)
        self.ollama_model_combo.blockSignals(False)
        self.ollama_status.setText(f"Status: connected — {len(models)} model(s)")
        self.ollama_status.setStyleSheet("color: #1b7a3d;")

    def _ollama_show_saved_rules(self) -> None:
        """Show user rules saved via Welcome → Describe a rule…"""
        learning = getattr(self, "learning", {}) or {}
        rules = list(learning.get("user_rules") or [])
        active = [r for r in rules if isinstance(r, dict) and r.get("active", True)]

        dialog = QDialog(self)
        dialog.setWindowTitle("Show saved rules")
        dialog.resize(560, 420)
        layout = QVBoxLayout(dialog)

        intro = QLabel(
            f"{len(active)} active rule(s) "
            f"({len(rules)} total). Compiled rules enforce locally; "
            "all rules are also passed to Ollama."
        )
        intro.setWordWrap(True)
        intro.setStyleSheet(f"color: {PALETTE['muted']};")
        layout.addWidget(intro)

        body = QTextEdit()
        body.setReadOnly(True)
        if not active:
            body.setPlainText(
                "No saved rules yet.\n\n"
                "Add one from Welcome → Email preview → Describe a rule…"
            )
        else:
            lines: List[str] = []
            for i, rule in enumerate(active, 1):
                scope = str(rule.get("scope") or "always")
                text = str(rule.get("text") or "").strip()
                sender = str(rule.get("sender") or "")
                at = str(rule.get("at") or "")
                state = "COMPILED" if rule.get("compiled") else "text-only"
                lines.append(f"{i}. [{scope}|{state}] {text}")
                meta = []
                if sender:
                    meta.append(f"sender={sender}")
                if at:
                    meta.append(f"saved={at}")
                if rule.get("compiled"):
                    meta.append(summarize_rule_effects(rule.get("effects")))
                if meta:
                    lines.append("   " + " · ".join(meta))
                lines.append("")
            body.setPlainText("\n".join(lines).rstrip())
        layout.addWidget(body, stretch=1)

        self.ollama_output.setPlainText(body.toPlainText())

        close_btn = QPushButton("Close")
        close_btn.setDefault(True)
        close_btn.clicked.connect(dialog.accept)
        add_dialog_button_row(layout, cancel=close_btn)
        dialog.exec()

    def _ollama_system_with_rules(
        self,
        base_system: str = "",
        *,
        mail_key: str = "",
        sender: str = "",
    ) -> str:
        """Append active user rules to the Ollama system prompt."""
        sys_text = (base_system or "").strip() or (
            "You are an assistant for Aliniant Smart Ass Email / GURI."
        )
        rules = format_user_rules_for_llm(
            getattr(self, "learning", {}) or {},
            mail_key=mail_key,
            sender=sender,
        )
        if rules:
            return f"{sys_text}\n\n{rules}"
        return sys_text

    def _ollama_analyze_selected_email(self) -> None:
        det = getattr(self, "_preview_det", None) or {}
        mail = det.get("mail") or {}
        if not mail.get("subject") and not mail.get("body_preview"):
            QMessageBox.information(
                self,
                "Ollama",
                "Select an email on the Welcome tab first (Important or Actions list).",
            )
            return
        key = mail_feedback_key(mail, det)
        rules = format_user_rules_for_llm(
            getattr(self, "learning", {}) or {},
            mail_key=key,
            sender=str(mail.get("sender") or ""),
        )
        prompt = build_email_analysis_prompt(
            mail, det.get("reasons") or [], user_rules=rules
        )
        self.ollama_prompt.setPlainText(prompt)
        self._ollama_run_prompt(
            mail_key=key, sender=str(mail.get("sender") or "")
        )

    def _ollama_run_prompt(self, checked: bool = False, *, mail_key: str = "", sender: str = "") -> None:
        if self._ollama_busy:
            self.status_bar.showMessage("Ollama is already running…")
            return
        cfg = self._ollama_collect_cfg()
        prompt = self.ollama_prompt.toPlainText().strip()
        if not cfg.get("model"):
            QMessageBox.warning(self, "Ollama", "Select or type a model name first.")
            return
        if not prompt:
            QMessageBox.warning(self, "Ollama", "Enter a prompt (or analyze a selected email).")
            return

        system = self._ollama_system_with_rules(
            cfg.get("system") or "",
            mail_key=mail_key,
            sender=sender,
        )

        self._ollama_busy = True
        self.ollama_run_btn.setEnabled(False)
        self.ollama_output.setPlainText("Thinking…")
        self.ollama_status.setText("Status: generating…")
        self.ollama_status.setStyleSheet("color: #5a7280;")

        def work() -> None:
            try:
                text = ollama_chat(
                    prompt,
                    base_url=cfg["base_url"],
                    model=cfg["model"],
                    system=system,
                    timeout=float(cfg.get("timeout") or 120),
                )
                self._ollama_bridge.chat_ready.emit(text, "")
            except Exception as exc:
                self._ollama_bridge.chat_ready.emit("", str(exc))

        threading.Thread(target=work, daemon=True).start()

    def _on_ollama_chat_ready(self, text: str, error: str) -> None:
        self._ollama_busy = False
        self.ollama_run_btn.setEnabled(True)
        if error:
            self.ollama_output.setPlainText(f"Error: {error}")
            self.ollama_status.setText(f"Status: error — {error}")
            self.ollama_status.setStyleSheet("color: #b42318;")
            return
        self.ollama_output.setPlainText(text)
        self.ollama_status.setText("Status: done")
        self.ollama_status.setStyleSheet("color: #1b7a3d;")
        self.status_bar.showMessage("Ollama response ready")

    def _create_library_tab(self):
        """Create the GURI Library tab with component library viewer."""
        library_frame = QWidget()
        self.notebook.addTab(library_frame, "Library")
        
        main_layout = QVBoxLayout()
        library_frame.setLayout(main_layout)
        main_layout.setContentsMargins(10, 10, 10, 10)
        
        # Statistics section
        stats_group = QGroupBox("Component Library Statistics")
        stats_layout = QVBoxLayout()
        stats_group.setLayout(stats_layout)
        
        stats_info = QWidget()
        stats_info_layout = QHBoxLayout()
        stats_info.setLayout(stats_info_layout)
        
        self.library_stats_label = QLabel("Loading statistics...")
        self.library_stats_label.setFont(QFont('Segoe UI', 10))
        stats_info_layout.addWidget(self.library_stats_label)
        
        refresh_stats_btn = QPushButton("Refresh Statistics")
        refresh_stats_btn.clicked.connect(self._refresh_library_stats)
        stats_info_layout.addWidget(refresh_stats_btn)

        repair_btn = QPushButton("Repair stable tags…")
        repair_btn.setToolTip(
            "Same real value must share one hex. Dedupes the library and "
            "rewrites GURI strings to the canonical tags."
        )
        repair_btn.clicked.connect(self._repair_guri_components)
        stats_info_layout.addWidget(repair_btn)
        
        stats_info_layout.addStretch()
        
        stats_layout.addWidget(stats_info)
        main_layout.addWidget(stats_group)
        
        # Component selection and search
        controls_group = QGroupBox("Browse Component Mappings")
        controls_layout = QVBoxLayout()
        controls_group.setLayout(controls_layout)
        
        # Component type selection
        component_controls = QWidget()
        component_controls_layout = QHBoxLayout()
        component_controls.setLayout(component_controls_layout)
        
        component_label = QLabel("Component Type:")
        component_controls_layout.addWidget(component_label)
        
        self.library_component_combo = QComboBox()
        self.library_component_combo.addItems([
            "All Components",
            "Component 1 (Sender)",
            "Component 2 (Recipients)",
            "Component 3 (Subject)",
            "Component 5 (Risk/Location)"
        ])
        self.library_component_combo.currentIndexChanged.connect(self._load_library_mappings)
        component_controls_layout.addWidget(self.library_component_combo)
        
        search_label = QLabel("Search:")
        component_controls_layout.addWidget(search_label)
        
        self.library_search_entry = QLineEdit()
        self.library_search_entry.setPlaceholderText("Search hex or real value...")
        self.library_search_entry.textChanged.connect(self._filter_library_mappings)
        component_controls_layout.addWidget(self.library_search_entry)
        
        refresh_mappings_btn = QPushButton("Refresh")
        refresh_mappings_btn.clicked.connect(self._load_library_mappings)
        component_controls_layout.addWidget(refresh_mappings_btn)
        
        component_controls_layout.addStretch()
        
        controls_layout.addWidget(component_controls)
        main_layout.addWidget(controls_group)
        
        # Component mappings tree
        mappings_group = QGroupBox("Component Mappings")
        mappings_layout = QVBoxLayout()
        mappings_group.setLayout(mappings_layout)
        
        columns = ("Component Type", "Hex Value", "Real Value", "GURI ID", "Last Updated")
        self.library_tree = QTreeWidget()
        self.library_tree.setAlternatingRowColors(True)
        self.library_tree.setRootIsDecorated(False)
        self.library_tree.setHeaderLabels(columns)
        self.library_tree.setColumnCount(len(columns))
        
        # Set column widths
        self.library_tree.setColumnWidth(0, 180)  # Component Type
        self.library_tree.setColumnWidth(1, 120)  # Hex Value
        self.library_tree.setColumnWidth(2, 400)  # Real Value
        self.library_tree.setColumnWidth(3, 100)  # GURI ID
        self.library_tree.setColumnWidth(4, 150)  # Last Updated
        
        # Enable sorting
        self.library_tree.setSortingEnabled(True)
        
        # Context menu for viewing GURI details
        self.library_tree.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.library_tree.customContextMenuRequested.connect(self._show_library_context_menu)
        
        mappings_layout.addWidget(self.library_tree)
        main_layout.addWidget(mappings_group)
        
        # Info section
        info_group = QGroupBox("About Component Library")
        info_layout = QVBoxLayout()
        info_group.setLayout(info_layout)
        
        info_text = QLabel(
            "The Component Library stores translations from GURI component hex values to their real values.\n\n"
            "Rule: the same real value always uses the same hex tag "
            "(e.g. every “FFFFF” recipients field shares one component hex).\n\n"
            "• Component 1 (Sender): 5 hex characters → Email sender address\n"
            "• Component 2 (Recipients): 5 hex characters → Recipient email addresses\n"
            "• Component 3 (Subject): 8 hex characters → Email subject or document name\n"
            "• Component 5 (Risk/Location): 3 hex characters → Risk level or document location\n\n"
            "Note: Component 4 (DateTime) is decodable and Component 6 (Document Type) is the actual code, "
            "so they are not stored in the library. The GURI ID column is only a last-seen record "
            "reference — the stable tag is the Hex Value."
        )
        info_text.setWordWrap(True)
        info_text.setStyleSheet("color: #666666; padding: 10px;")
        info_layout.addWidget(info_text)
        
        main_layout.addWidget(info_group)
        
        # Load initial data
        self._refresh_library_stats()
        self._load_library_mappings()
    
    def _refresh_library_stats(self):
        """Refresh component library statistics."""
        if not self.component_library:
            self.library_stats_label.setText("No database connection")
            return
        
        try:
            stats = self.component_library.get_component_statistics()
            
            stats_text = "Component Library Statistics: "
            stats_parts = []
            for component_name, count in stats.items():
                stats_parts.append(f"{component_name}: {count}")
            
            if stats_parts:
                stats_text += " | ".join(stats_parts)
            else:
                stats_text += "No mappings stored yet"
            
            self.library_stats_label.setText(stats_text)
            self.status_bar.showMessage("Library statistics refreshed")
            
        except Exception as e:
            self.library_stats_label.setText(f"Error loading statistics: {str(e)}")
            self.logger.error(f"Error refreshing library stats: {e}")

    def _repair_guri_components(self) -> None:
        """Dedupe component library and rewrite GURI strings to stable hex tags."""
        if not self.db:
            QMessageBox.warning(self, "Repair GURI tags", "Connect to the database first.")
            return
        if not self.component_library or self.component_library.database != self.db:
            self.component_library = GURIComponentLibrary(database=self.db)

        reply = QMessageBox.question(
            self,
            "Repair GURI component tags",
            "This enforces: same real value → same hex tag.\n\n"
            "1. Collapse duplicate library rows (keep oldest hex per value)\n"
            "2. Rewrite guri_records so sender/recipients/subject/risk use those tags\n\n"
            "Continue?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.Yes,
        )
        if reply != QMessageBox.StandardButton.Yes:
            return

        self.status_bar.showMessage("Repairing GURI component tags…")
        try:
            report = self.component_library.repair_stable_components(rewrite_records=True)
            # Retry unique index now that duplicates are gone
            try:
                self.component_library._init_library_table()
            except Exception:
                pass
            dedupe = report.get("dedupe") or {}
            rewrite = report.get("rewrite") or {}
            msg = (
                f"Library: {dedupe.get('groups', 0)} duplicate group(s), "
                f"removed {dedupe.get('deleted', 0)} row(s).\n"
                f"Records: scanned {rewrite.get('scanned', 0)}, "
                f"updated {rewrite.get('updated', 0)}, "
                f"unchanged {rewrite.get('skipped', 0)}, "
                f"errors {rewrite.get('errors', 0)}."
            )
            self.status_bar.showMessage("GURI component repair finished")
            QMessageBox.information(self, "Repair GURI tags", msg)
            self._refresh_library_stats()
            self._load_library_mappings()
            try:
                self._load_records()
            except Exception:
                pass
        except Exception as exc:
            self.logger.error("GURI component repair failed: %s", exc)
            QMessageBox.critical(self, "Repair GURI tags", f"Repair failed:\n{exc}")
    
    def _load_library_mappings(self):
        """Load component mappings into the library tree."""
        if not self.component_library:
            self.library_tree.clear()
            return
        
        try:
            self.library_tree.clear()
            
            # Get selected component type
            selected_index = self.library_component_combo.currentIndex()
            
            # Component index mapping: 0=All, 1=Component1, 2=Component2, 3=Component3, 4=Component5
            component_index_map = {
                0: None,  # All
                1: 0,     # Component 1 (Sender)
                2: 1,     # Component 2 (Recipients)
                3: 2,     # Component 3 (Subject)
                4: 4      # Component 5 (Risk/Location)
            }
            
            component_index = component_index_map.get(selected_index)
            
            if component_index is None:
                # Load all components
                for comp_idx in [0, 1, 2, 4]:
                    mappings = self.component_library.get_all_mappings_for_component(comp_idx, limit=1000)
                    self._add_mappings_to_tree(mappings, comp_idx)
            else:
                # Load specific component
                mappings = self.component_library.get_all_mappings_for_component(component_index, limit=1000)
                self._add_mappings_to_tree(mappings, component_index)
            
            # Apply search filter if active
            self._filter_library_mappings()
            
            self.status_bar.showMessage(f"Loaded component mappings")
            
        except Exception as e:
            self.logger.error(f"Error loading library mappings: {e}")
            QMessageBox.critical(self, "Error", f"Failed to load component mappings:\n{str(e)}")
    
    def _add_mappings_to_tree(self, mappings: list, component_index: int):
        """Add mappings to the library tree."""
        component_names = {
            0: "Component 1 (Sender)",
            1: "Component 2 (Recipients)",
            2: "Component 3 (Subject)",
            4: "Component 5 (Risk/Location)"
        }
        
        component_name = component_names.get(component_index, f"Component {component_index}")
        
        for mapping in mappings:
            item = QTreeWidgetItem(self.library_tree)
            item.setText(0, component_name)
            item.setText(1, mapping['component_hex'])
            
            # Truncate long real values for display
            real_value = mapping['real_value']
            if len(real_value) > 100:
                display_value = real_value[:100] + "..."
            else:
                display_value = real_value
            item.setText(2, display_value)
            item.setToolTip(2, real_value)  # Full value in tooltip
            
            item.setText(3, str(mapping['guri_id']) if mapping['guri_id'] else "")
            item.setText(4, str(mapping.get('updated_at', mapping.get('created_at', ''))))
            
            # Store full data for context menu
            item.setData(0, Qt.ItemDataRole.UserRole, mapping)
    
    def _filter_library_mappings(self):
        """Filter library mappings based on search text."""
        search_text = self.library_search_entry.text().lower()
        
        if not search_text:
            # Show all items
            for i in range(self.library_tree.topLevelItemCount()):
                item = self.library_tree.topLevelItem(i)
                if item is not None:
                    item.setHidden(False)
            return
        
        # Filter items
        for i in range(self.library_tree.topLevelItemCount()):
            item = self.library_tree.topLevelItem(i)
            if item is None:
                continue
            # Search in component type, hex value, and real value
            match = (
                search_text in item.text(0).lower() or
                search_text in item.text(1).lower() or
                search_text in item.text(2).lower()
            )
            
            item.setHidden(not match)
    
    def _show_library_context_menu(self, position):
        """Show context menu for library item actions."""
        item = self.library_tree.itemAt(position)
        if not item:
            return
        
        # Create context menu
        menu = QMenu(self)
        
        # View full real value
        view_value_action = menu.addAction("View Full Value", lambda: self._view_full_library_value(item))
        
        # View GURI if available
        guri_id = item.text(3)
        if guri_id:
            menu.addSeparator()
            view_guri_action = menu.addAction("View GURI Record", lambda: self._view_library_guri(guri_id))
            copy_guri_action = menu.addAction("Copy GURI ID", lambda: self._copy_to_clipboard(guri_id))
        
        # Copy actions
        menu.addSeparator()
        copy_hex_action = menu.addAction("Copy Hex Value", lambda: self._copy_to_clipboard(item.text(1)))
        copy_real_action = menu.addAction("Copy Real Value", lambda: self._copy_to_clipboard(item.toolTip(2) or item.text(2)))
        
        # Show menu at cursor position
        menu.exec(self.library_tree.mapToGlobal(position))
    
    def _view_full_library_value(self, item):
        """View full real value in a dialog."""
        real_value = item.toolTip(2) or item.text(2)
        
        dialog = QDialog(self)
        dialog.setWindowTitle("Full Component Value")
        dialog.resize(600, 300)
        dialog.setModal(True)
        
        layout = QVBoxLayout()
        dialog.setLayout(layout)
        
        text = QTextEdit()
        text.setReadOnly(True)
        text.setFont(QFont("Courier", 10))
        text.setPlainText(real_value)
        layout.addWidget(text)
        
        close_btn = QPushButton("Close")
        close_btn.clicked.connect(dialog.accept)
        add_dialog_button_row(layout, cancel=close_btn)
        
        dialog.exec()
    
    def _view_library_guri(self, guri_id: str):
        """View GURI record details from library."""
        if not self.db:
            return
        
        try:
            record_id = int(guri_id)
            
            # Search through records to find by ID
            records = self.db.get_all_records(limit=10000)
            record = None
            for rec in records:
                if rec.get('id') == record_id:
                    record = rec
                    break
            
            if record:
                values = [
                    str(record.get('id', '')),
                    record.get('guri', ''),
                    record.get('sender', ''),
                    record.get('recipients', ''),
                    record.get('subject', ''),
                    record.get('datetime', ''),
                    record.get('avg_risk', ''),
                    DOCUMENT_TYPES.get(record.get('document_type', ''), 'Unknown'),
                    record.get('created_at', '')
                ]
                self._show_record_details_dialog(values)
            else:
                QMessageBox.information(self, "Not Found", f"GURI record with ID {guri_id} not found.")
                
        except Exception as e:
            self.logger.error(f"Error viewing library GURI: {e}")
            QMessageBox.critical(self, "Error", f"Failed to load GURI record:\n{str(e)}")
    
    def _copy_to_clipboard(self, text: str):
        """Copy text to clipboard."""
        clipboard = QApplication.clipboard()
        clipboard.setText(text)
        self.status_bar.showMessage(f"Copied to clipboard: {text[:50]}...")
    
    def _load_important_domains(self):
        """Load important domains list from file."""
        domains_file = os.path.join(os.path.dirname(__file__), "important_domains.txt")
        if os.path.exists(domains_file):
            try:
                with open(domains_file, 'r', encoding='utf-8') as f:
                    self.important_domains = [
                        line.strip().lower()
                        for line in f
                        if line.strip() and not line.strip().startswith('#')
                    ]
                self.logger.info(f"Loaded {len(self.important_domains)} important domains from file")
            except Exception as e:
                self.logger.error(f"Error loading important domains: {e}")
                self.important_domains = []
        else:
            # Create with some defaults
            self.important_domains = []
            self._save_important_domains()
    
    def _load_email_status(self):
        """Load persisted email reply-status labels keyed by GURI."""
        status_file = os.path.join(os.path.dirname(__file__), "email_status.json")
        if os.path.exists(status_file):
            try:
                with open(status_file, 'r', encoding='utf-8') as f:
                    self.email_status = json.load(f)
                self.logger.info(f"Loaded {len(self.email_status)} email status entries")
            except Exception as e:
                self.logger.error(f"Error loading email status: {e}")
                self.email_status = {}
        else:
            self.email_status = {}
    
    def _save_email_status(self):
        """Save email reply-status labels to file."""
        status_file = os.path.join(os.path.dirname(__file__), "email_status.json")
        try:
            with open(status_file, 'w', encoding='utf-8') as f:
                json.dump(self.email_status, f, indent=2)
        except Exception as e:
            self.logger.error(f"Error saving email status: {e}")
    
    def _save_important_domains(self):
        """Save important domains list to file."""
        domains_file = os.path.join(os.path.dirname(__file__), "important_domains.txt")
        try:
            with open(domains_file, 'w', encoding='utf-8') as f:
                for domain in sorted(set(self.important_domains)):
                    f.write(f"{domain}\n")
            self.logger.info(f"Saved {len(self.important_domains)} important domains to file")
        except Exception as e:
            self.logger.error(f"Error saving important domains: {e}")
    
    def _add_important_domain(self):
        """Add a new important domain."""
        domain = self.domain_entry.text().strip().lower()
        
        if not domain:
            QMessageBox.warning(self, "No Domain", "Please enter a domain to add")
            return
        
        if domain in self.important_domains:
            QMessageBox.information(self, "Already Exists", f"'{domain}' is already in the important domains list")
            return
        
        self.important_domains.append(domain)
        self._save_important_domains()
        self._refresh_domains_list()
        self.domain_entry.clear()
        self.status_bar.showMessage(f"Added '{domain}' to important domains")
        self._refresh_important_emails()
    
    def _remove_important_domain(self):
        """Remove selected important domain."""
        current_item = self.domains_listbox.currentItem()
        if not current_item:
            QMessageBox.warning(self, "No Selection", "Please select a domain to remove")
            return
        
        domain = current_item.text()
        
        response = QMessageBox.question(
            self,
            "Confirm Removal",
            f"Remove '{domain}' from important domains?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No
        )
        if response == QMessageBox.StandardButton.Yes:
            self.important_domains.remove(domain)
            self._save_important_domains()
            self._refresh_domains_list()
            self.status_bar.showMessage(f"Removed '{domain}' from important domains")
            self._refresh_important_emails()
    
    def _refresh_domains_list(self):
        """Refresh the domains listbox."""
        self.domains_listbox.clear()
        for domain in sorted(self.important_domains):
            self.domains_listbox.addItem(domain)
        self._refresh_vip_list()

    def _refresh_vip_list(self) -> None:
        """Refresh the VIP senders listbox from learning."""
        box = getattr(self, "vip_listbox", None)
        if box is None:
            return
        box.clear()
        for addr in sorted(self._important_senders_list()):
            box.addItem(addr)

    def _add_vip_sender(self) -> None:
        entry = getattr(self, "vip_entry", None)
        raw = (entry.text() if entry is not None else "").strip()
        addr = normalize_sender_address(raw) or raw.lower().strip()
        if not addr or "@" not in addr:
            QMessageBox.information(
                self,
                "VIP sender",
                "Enter a full email address (e.g. boss@acme.com).",
            )
            return
        if addr in {a.lower() for a in self._important_senders_list()}:
            self.status_bar.showMessage(f"{addr} is already a VIP")
            return
        self.learning = remember_important_sender(self.learning, addr)
        self._save_learning()
        if entry is not None:
            entry.clear()
        self._refresh_vip_list()
        self._refresh_important_emails(allow_db_fallback=False, refresh_actions=True)
        self.status_bar.showMessage(f"Marked VIP: {addr}")

    def _remove_vip_sender(self) -> None:
        box = getattr(self, "vip_listbox", None)
        if box is None:
            return
        item = box.currentItem()
        if item is None:
            QMessageBox.information(self, "VIP sender", "Select a VIP to remove.")
            return
        addr = (item.text() or "").strip()
        if not addr:
            return
        self.learning = forget_important_sender(self.learning, addr)
        self._save_learning()
        self._refresh_vip_list()
        self._refresh_important_emails(allow_db_fallback=False, refresh_actions=True)
        self.status_bar.showMessage(f"Removed VIP: {addr}")

    def _preview_toggle_vip_sender(self) -> None:
        ctx = self._preview_mail_context()
        if not ctx:
            QMessageBox.information(self, "VIP sender", "Select an email in the preview first.")
            return
        sender = str(ctx["mail"].get("sender") or "")
        addr = normalize_sender_address(sender)
        if not addr:
            QMessageBox.information(self, "VIP sender", "Could not read a sender address.")
            return
        from guri_learning import is_important_sender

        if is_important_sender(self.learning, addr):
            self.learning = forget_important_sender(self.learning, addr)
            self._save_learning()
            self._refresh_vip_list()
            self._refresh_important_emails(allow_db_fallback=False, refresh_actions=True)
            self._show_email_preview(ctx["det"])
            self.status_bar.showMessage(f"Removed VIP: {addr}")
        else:
            self.learning = remember_important_sender(self.learning, addr)
            self._save_learning()
            domain = domain_from_sender(addr)
            if domain and domain not in self.important_domains:
                add_dom = QMessageBox.question(
                    self,
                    "Important domain?",
                    f"Also add @{domain} to Important Emails domains?",
                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                    QMessageBox.StandardButton.No,
                )
                if add_dom == QMessageBox.StandardButton.Yes:
                    self.important_domains.append(domain)
                    self._save_important_domains()
                    self._refresh_domains_list()
            self._refresh_vip_list()
            self._refresh_important_emails(allow_db_fallback=False, refresh_actions=True)
            self._show_email_preview(ctx["det"])
            self.status_bar.showMessage(f"Marked VIP: {addr}")
    
    def _refresh_welcome_dashboard(self):
        """Refresh all welcome/dashboard panels after database connection."""
        if self.db:
            # Stats are expensive (10k rows) — keep off the auto-refresh path
            self._refresh_welcome_stats()
        else:
            self._refresh_welcome_stats_from_scrape_cache()
        self._refresh_welcome_panels(allow_db_fallback=True)
    
    def _my_email_addresses(self) -> Set[str]:
        """Lowercased addresses that count as 'me' for Welcome sent/received stats."""
        emails: Set[str] = {MY_ALINIANT_EMAIL.lower()}
        try:
            for addr in (getattr(self, "learning", {}) or {}).get("my_emails") or []:
                a = str(addr or "").strip().lower()
                if a and "@" in a:
                    emails.add(a)
        except Exception:
            pass
        try:
            for cfg in load_aes_account_configs() or []:
                smtp = str(getattr(cfg, "smtp", "") or "").strip().lower()
                if smtp and "@" in smtp:
                    emails.add(smtp)
        except Exception:
            pass
        return emails

    def _refresh_welcome_stats_from_scrape_cache(self) -> None:
        """Cheap Today/Yesterday counts from the live Outlook scrape cache."""
        if not hasattr(self, "today_sent_label"):
            return
        try:
            today = datetime.now().date()
            yesterday = today - timedelta(days=1)
            mine = self._my_email_addresses()
            today_sent = today_received = yesterday_sent = yesterday_received = 0

            for item in self.scrape_cache_items or []:
                mail = item.get("mail") or {}
                record_dt = _parse_record_datetime(mail.get("received"))
                if not record_dt:
                    continue
                record_date = record_dt.date()
                if record_date not in (today, yesterday):
                    continue

                sender = str(mail.get("sender") or "").lower()
                to_cc = " ".join(
                    [
                        str(mail.get("to") or ""),
                        str(mail.get("cc") or ""),
                    ]
                ).lower()
                is_sent = any(m in sender for m in mine)
                is_received = (not is_sent) and (
                    any(m in to_cc for m in mine)
                    or str(mail.get("recipient_role") or "").lower()
                    in ("to", "cc", "received", "inbox")
                )

                if record_date == today:
                    if is_sent:
                        today_sent += 1
                    if is_received:
                        today_received += 1
                else:
                    if is_sent:
                        yesterday_sent += 1
                    if is_received:
                        yesterday_received += 1

            self.today_sent_label.setText(f"Sent: {today_sent}")
            self.today_received_label.setText(f"Received: {today_received}")
            self.yesterday_sent_label.setText(f"Sent: {yesterday_sent}")
            self.yesterday_received_label.setText(f"Received: {yesterday_received}")
            self._touch_last_refresh_label()
        except Exception as e:
            self.logger.debug("Welcome stats from scrape cache failed: %s", e)

    def _refresh_welcome_stats(self):
        """Refresh email statistics on welcome tab."""
        if not self.db:
            self._refresh_welcome_stats_from_scrape_cache()
            return
        
        try:
            from datetime import timedelta
            
            # Get today's date
            today = datetime.now().date()
            yesterday = today - timedelta(days=1)
            
            # Query for emails (document_type = '01')
            all_records = self.db.get_all_records(limit=10000)
            
            mine = self._my_email_addresses()
            
            # Count today's emails
            today_sent = 0
            today_received = 0
            yesterday_sent = 0
            yesterday_received = 0
            
            for record in all_records:
                if record['document_type'] != '01':  # Only count emails
                    continue
                
                record_dt = _parse_record_datetime(record['datetime'])
                if not record_dt:
                    continue
                record_date = record_dt.date()
                
                # Check sender/recipient
                sender = str(record['sender']).lower()
                recipients = str(record['recipients']).lower()
                is_sent = any(m in sender for m in mine)
                is_received = (not is_sent) and any(m in recipients for m in mine)
                
                # Count by date
                if record_date == today:
                    if is_sent:
                        today_sent += 1
                    if is_received:
                        today_received += 1
                elif record_date == yesterday:
                    if is_sent:
                        yesterday_sent += 1
                    if is_received:
                        yesterday_received += 1
            
            # Update labels
            self.today_sent_label.setText(f"Sent: {today_sent}")
            self.today_received_label.setText(f"Received: {today_received}")
            self.yesterday_sent_label.setText(f"Sent: {yesterday_sent}")
            self.yesterday_received_label.setText(f"Received: {yesterday_received}")
            
            # Update last refresh time
            self._touch_last_refresh_label()
            
            self.status_bar.showMessage("Statistics refreshed")
            
        except Exception as e:
            self.logger.error(f"Error refreshing welcome stats: {e}")
            self._refresh_welcome_stats_from_scrape_cache()
            QMessageBox.critical(self, "Error", f"Failed to refresh statistics:\n{str(e)}")
    
    def _load_scrape_settings(self) -> None:
        try:
            cfg = load_scrape_settings()
        except Exception:
            cfg = {"lookback_months": 0, "lookback_locked": False, "unread_only": False}
        months = int(cfg.get("lookback_months") or 0)
        locked = bool(cfg.get("lookback_locked")) and months > 0
        self.scrape_lookback_months = months if locked else months
        self.scrape_lookback_locked = locked
        self.scrape_unread_only = bool(cfg.get("unread_only"))
        if locked and months > 0:
            self.scrape_days = lookback_months_to_days(months)
        elif not locked:
            # Keep default week until the user locks a lookback
            self.scrape_days = max(7, int(getattr(self, "scrape_days", 7) or 7))

    def _persist_scrape_settings(self) -> None:
        try:
            save_scrape_settings(
                {
                    "lookback_months": int(getattr(self, "scrape_lookback_months", 0) or 0),
                    "lookback_locked": bool(getattr(self, "scrape_lookback_locked", False)),
                    "unread_only": bool(getattr(self, "scrape_unread_only", False)),
                }
            )
        except Exception as exc:
            self.logger.warning("Failed saving scrape settings: %s", exc)

    def _update_lookback_save_button_style(self) -> None:
        btn = getattr(self, "scrape_lookback_save_btn", None)
        if btn is None:
            return
        locked = bool(getattr(self, "scrape_lookback_locked", False))
        if locked:
            btn.setText("Saved")
            btn.setStyleSheet(
                "QPushButton {"
                "background-color: #1b7a3d; color: #ffffff; font-weight: 600;"
                "border: 1px solid #145c2e; border-radius: 4px; padding: 4px 8px;"
                "}"
                "QPushButton:hover { background-color: #166533; }"
            )
        else:
            btn.setText("Save")
            btn.setStyleSheet("")

    def _on_lookback_months_edited(self, _value: int = 0) -> None:
        # Editing unlocks until Save is clicked again
        if getattr(self, "scrape_lookback_locked", False):
            self.scrape_lookback_locked = False
            self._update_lookback_save_button_style()

    def _save_scrape_lookback_months(self) -> None:
        spin = getattr(self, "scrape_lookback_spin", None)
        months = int(spin.value()) if spin is not None else 1
        months = max(1, min(60, months))
        self.scrape_lookback_months = months
        self.scrape_lookback_locked = True
        self.scrape_days = lookback_months_to_days(months)
        if spin is not None:
            spin.setValue(months)
        self._persist_scrape_settings()
        self._update_lookback_save_button_style()
        self.status_bar.showMessage(
            f"Scrape lookback locked: {months} month(s) (~{self.scrape_days} days)"
        )

    def _load_scrape_cache_into_memory(self) -> None:
        cache = load_scrape_cache()
        items = cache.get("items") if isinstance(cache, dict) else None
        self.scrape_cache_items = list(items or [])
        if isinstance(cache, dict):
            # Prefer locked lookback months over stale cache days
            if getattr(self, "scrape_lookback_locked", False) and self.scrape_lookback_months > 0:
                self.scrape_days = lookback_months_to_days(self.scrape_lookback_months)
            else:
                self.scrape_days = int(cache.get("days") or self.scrape_days)
            if "unread_only" in cache and not hasattr(self, "unread_only_toggle"):
                self.scrape_unread_only = bool(cache.get("unread_only") or False)

    def _on_unread_only_toggled(self, checked: bool) -> None:
        self.scrape_unread_only = bool(checked)
        self.unread_only_toggle.setText(
            "Unread only: ON" if checked else "Unread only: OFF"
        )
        self._persist_scrape_settings()

    def _link_guris_for_items(self, items: List[Dict[str, Any]]) -> None:
        """Match scraped mails to GURI records using one DB fetch (not per-mail)."""
        if not self.db or not items:
            return
        try:
            records = self.db.get_all_records(limit=3000)
        except Exception as exc:
            self.logger.debug("GURI batch link load failed: %s", exc)
            return

        index: Dict[tuple, List[Dict[str, Any]]] = {}
        for record in records:
            if str(record.get("document_type") or "") != "01":
                continue
            key = (
                str(record.get("sender") or "").lower(),
                str(record.get("subject") or "").strip().lower(),
            )
            if not key[0] or not key[1]:
                continue
            index.setdefault(key, []).append(record)

        for item in items:
            mail = item.get("mail") or {}
            sender = str(mail.get("sender") or "").lower()
            subject = str(mail.get("subject") or "").strip().lower()
            if not sender or not subject:
                continue
            received = _parse_record_datetime(mail.get("received"))
            for record in index.get((sender, subject), []):
                rec_dt = _parse_record_datetime(record.get("datetime"))
                if (
                    received
                    and rec_dt
                    and abs((received - rec_dt).total_seconds()) > 48 * 3600
                ):
                    continue
                guri = str(record.get("guri") or "") or None
                if guri:
                    mail["guri"] = guri
                    item["guri"] = guri
                break

    def _link_guri_to_mail(self, mail: Dict[str, Any]) -> Optional[str]:
        """Best-effort match one scraped mail to an existing GURI DB record."""
        item = {"mail": mail or {}}
        self._link_guris_for_items([item])
        return str((item.get("mail") or {}).get("guri") or "") or None

    def _start_outlook_scrape(self, *, manual: bool = False) -> None:
        """Kick off background Outlook scrape + action detection.

        Auto scrapes are light (headers only, capped lookback/count) so Outlook
        stays responsive. Manual Scrape Outlook uses the full saved lookback.
        """
        if self._scrape_busy:
            self.status_bar.showMessage("Outlook scrape already running…")
            return

        # AES account settings exist but every account is switched off: scraping
        # "all stores" as a fallback is what froze Outlook. Skip and say so.
        try:
            cfg_preview = load_aes_account_configs()
        except Exception:
            cfg_preview = []
        if cfg_preview and not any(c.enabled for c in cfg_preview):
            warn = (
                "All AES email accounts are disabled — Outlook scrape skipped. "
                "Enable at least one account in the Accounts tab (or AES Settings)."
            )
            self.status_bar.showMessage(warn)
            self.logger.warning(warn)
            if hasattr(self, "actions_hint"):
                self.actions_hint.setText(warn)
            return

        self._scrape_busy = True
        self._scrape_was_manual = bool(manual)
        if hasattr(self, "scrape_btn"):
            self.scrape_btn.setEnabled(False)
        self.status_bar.showMessage("Scraping Outlook inboxes…")

        full_days = max(1, int(self.scrape_days or 7))
        if manual:
            days = full_days
            max_per_account = 200
            include_body = False
        else:
            # Incremental auto: recent window only — full 6‑month body scrape freezes Outlook
            days = min(full_days, 14)
            max_per_account = 60
            include_body = False

        unread_only = self.scrape_unread_only
        domains = list(self.important_domains)
        vip_senders = list(self._important_senders_list())
        extra_cues = list(self._learning_deadline_cues())
        # The targeted important-sender pass uses non-indexed DASL LIKE queries
        # against the full lookback — expensive enough to freeze Outlook's UI.
        # Run it only on explicit manual scrapes, never on the 5-minute timer.
        targeted_domains = domains if manual else []
        targeted_senders = vip_senders if manual else []

        def work() -> None:
            try:
                configs = load_aes_account_configs()
                response_flags = {
                    (c.smtp or c.store_id).lower(): bool(c.responses) for c in configs
                }
                mails = scrape_outlook_inboxes(
                    days=days,
                    max_per_account=max_per_account,
                    unread_only=unread_only,
                    account_configs=configs or None,
                    include_body=include_body,
                    # Targeted full-lookback pass so important mail is never
                    # pushed out by the per-account recency cap (manual only).
                    important_domains=targeted_domains,
                    important_senders=targeted_senders,
                )
                detected = detect_batch(
                    mails,
                    important_domains=domains,
                    important_senders=vip_senders,
                    extra_deadline_cues=extra_cues,
                    account_response_flags=response_flags,
                )
                items = [d.to_dict() for d in detected]
                self._scrape_bridge.finished.emit(items, "")
            except Exception as exc:
                self._scrape_bridge.finished.emit(None, str(exc))

        threading.Thread(target=work, daemon=True).start()

    def _mail_item_key(self, item: Dict[str, Any]) -> str:
        mail = item.get("mail") or {}
        key = str(mail.get("entry_id") or item.get("entry_id") or "").strip()
        if key:
            return key
        return "|".join(
            [
                str(mail.get("sender") or ""),
                str(mail.get("subject") or ""),
                str(mail.get("received") or ""),
            ]
        )

    def _seed_seen_mail_keys(self, items: Optional[List[Dict[str, Any]]]) -> None:
        if not hasattr(self, "_seen_mail_keys"):
            self._seen_mail_keys = set()
        for it in items or []:
            key = self._mail_item_key(it)
            if key:
                self._seen_mail_keys.add(key)

    def _parse_vip_received(self, value: Any) -> Optional[datetime]:
        if isinstance(value, datetime):
            return value.replace(tzinfo=None) if value.tzinfo else value
        text = str(value or "").strip()
        if not text:
            return None
        for fmt, n in (
            ("%Y-%m-%d %H:%M:%S", 19),
            ("%Y-%m-%d %H:%M", 16),
            ("%Y-%m-%dT%H:%M:%S", 19),
        ):
            try:
                return datetime.strptime(text[:n], fmt)
            except ValueError:
                continue
        try:
            return datetime.fromisoformat(text.replace("Z", "+00:00")).replace(tzinfo=None)
        except ValueError:
            return None

    def _collect_recent_new_vip_mails(
        self,
        previous_items: List[Dict[str, Any]],
        incoming: List[Dict[str, Any]],
    ) -> List[Dict[str, Any]]:
        """New VIP-sender messages not seen before and received in the last ~15 min."""
        from guri_learning import is_important_sender

        learning = getattr(self, "learning", {}) or {}
        prev_keys = {self._mail_item_key(it) for it in previous_items or [] if self._mail_item_key(it)}
        seen = getattr(self, "_seen_mail_keys", set()) or set()
        cold_start = not prev_keys and not seen
        cutoff = datetime.now() - timedelta(minutes=15)
        arrivals: List[Dict[str, Any]] = []
        for it in incoming or []:
            key = self._mail_item_key(it)
            if not key or key in seen or key in prev_keys:
                continue
            mail = it.get("mail") or {}
            sender = str(mail.get("sender") or "")
            if not is_important_sender(learning, sender):
                continue
            received = self._parse_vip_received(mail.get("received"))
            if received is not None and received < cutoff:
                continue
            # Cold start with unparseable dates: don't siren for a whole lookback dump
            if received is None and cold_start:
                continue
            arrivals.append(it)
        return arrivals

    def _flash_vip_police_light(self, arrivals: List[Dict[str, Any]]) -> None:
        light = getattr(self, "_vip_police_light", None)
        if light is None or not arrivals:
            return
        light.flash(2000)
        mail = arrivals[0].get("mail") or {}
        sender = str(mail.get("sender") or "VIP")
        if len(arrivals) == 1:
            self.status_bar.showMessage(f"VIP mail from {sender}", 3000)
        else:
            self.status_bar.showMessage(
                f"VIP mail: {len(arrivals)} new messages (incl. {sender})",
                3000,
            )
        self.logger.info("VIP police light: %d new VIP message(s)", len(arrivals))

    def _on_scrape_finished(self, items_obj, error: str) -> None:
        self._scrape_busy = False
        if hasattr(self, "scrape_btn"):
            self.scrape_btn.setEnabled(True)
        if error:
            self.status_bar.showMessage(f"Outlook scrape failed: {error}")
            self.logger.error("Outlook scrape failed: %s", error)
            self._schedule_welcome_refresh(allow_db_fallback=False)
            return

        incoming: List[Dict[str, Any]] = list(items_obj or [])
        previous_items = list(self.scrape_cache_items or [])
        vip_arrivals = self._collect_recent_new_vip_mails(previous_items, incoming)

        items: List[Dict[str, Any]] = list(incoming)
        if not items and self.scrape_cache_items:
            # Don't wipe a good Welcome cache when Outlook returns nothing
            # (e.g. locale Restrict bug / transient COM glitch).
            self.status_bar.showMessage(
                "Outlook scrape returned 0 messages — keeping previous results. "
                "Try Scrape Outlook again."
            )
            self.logger.warning(
                "Outlook scrape returned 0 messages; preserving %d cached item(s)",
                len(self.scrape_cache_items),
            )
            if hasattr(self, "actions_hint"):
                scraped_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                self.actions_hint.setText(
                    f"Live Outlook scrape at {scraped_at} — 0 new messages "
                    f"(kept {len(self.scrape_cache_items)} cached)."
                )
            self._touch_last_refresh_label()
            self._schedule_welcome_refresh(allow_db_fallback=False)
            return

        # One DB fetch for all mails — never N× get_all_records(3000) on the UI thread
        self._link_guris_for_items(items)
        apply_statuses(items, {**self.email_status, **self.action_status})

        # Auto/incremental scrapes are intentionally shallow — merge into cache
        # instead of wiping older lookback results from a prior full scrape.
        if not getattr(self, "_scrape_was_manual", False) and self.scrape_cache_items:
            by_key: Dict[str, Dict[str, Any]] = {}
            for old in self.scrape_cache_items:
                key = self._mail_item_key(old)
                if key:
                    by_key[key] = old
            for new in items:
                key = self._mail_item_key(new)
                if key:
                    by_key[key] = new
            items = list(by_key.values())
            items.sort(
                key=lambda x: str((x.get("mail") or {}).get("received") or ""),
                reverse=True,
            )
            if len(items) > 800:
                # Keep ALL important-domain/VIP mail; trim only ordinary items,
                # so auto merges never evict what a manual 6-month scrape found.
                keep = [
                    it for it in items
                    if self._mail_is_important_domain(it.get("mail") or {})
                ]
                rest = [
                    it for it in items
                    if not self._mail_is_important_domain(it.get("mail") or {})
                ]
                items = keep + rest[: max(0, 800 - len(keep))]
                items.sort(
                    key=lambda x: str((x.get("mail") or {}).get("received") or ""),
                    reverse=True,
                )

        self.scrape_cache_items = items
        self._last_outlook_scrape_at = datetime.now()
        self._scrape_was_manual = False
        self._seed_seen_mail_keys(items)

        if vip_arrivals:
            self._flash_vip_police_light(vip_arrivals)

        days = self.scrape_days
        unread_only = self.scrape_unread_only

        def _save() -> None:
            try:
                save_scrape_cache(items, days=days, unread_only=unread_only)
            except Exception as exc:
                self.logger.warning("Failed saving scrape cache: %s", exc)

        threading.Thread(target=_save, daemon=True).start()

        action_count = sum(1 for it in items if it.get("actions"))
        deadline_count = sum(
            1
            for it in items
            if it.get("deadlines")
            or "deadline" in [str(a).lower() for a in (it.get("actions") or [])]
        )
        self.status_bar.showMessage(
            f"Outlook scrape: {len(items)} messages, {action_count} with actions, "
            f"{deadline_count} with deadlines"
        )
        self._touch_last_refresh_label()
        if hasattr(self, "actions_hint"):
            scraped_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            self.actions_hint.setText(
                f"Live Outlook scrape at {scraped_at} — "
                f"{action_count} action item(s) from {len(items)} message(s)."
            )
        # Coalesce panel rebuilds so maximize/tray clicks aren't buried in work
        self._schedule_welcome_refresh(allow_db_fallback=False)

    def _schedule_welcome_refresh(self, *, allow_db_fallback: bool = False) -> None:
        self._pending_welcome_db_fallback = bool(allow_db_fallback)
        if getattr(self, "_welcome_refresh_scheduled", False):
            return
        self._welcome_refresh_scheduled = True
        QTimer.singleShot(50, self._run_scheduled_welcome_refresh)

    def _run_scheduled_welcome_refresh(self) -> None:
        self._welcome_refresh_scheduled = False
        if self._scrape_busy or getattr(self, "_expanding_to_screen", False):
            self._welcome_refresh_scheduled = True
            QTimer.singleShot(200, self._run_scheduled_welcome_refresh)
            return
        allow_db = bool(getattr(self, "_pending_welcome_db_fallback", False))
        self._refresh_welcome_panels(allow_db_fallback=allow_db)

    def _refresh_welcome_panels(self, *, allow_db_fallback: bool = False) -> None:
        """Rebuild Welcome/Deadlines lists without nested duplicate refreshes."""
        if getattr(self, "_welcome_refresh_busy", False):
            return
        self._welcome_refresh_busy = True
        try:
            self._refresh_welcome_stats_from_scrape_cache()
            self._refresh_email_timeline(allow_db_fallback=allow_db_fallback)
            self._refresh_important_emails(
                allow_db_fallback=allow_db_fallback, refresh_actions=False
            )
            self._refresh_actions_panel()
            self._refresh_deadlines_panel()
            self._touch_last_refresh_label()
        except Exception as e:
            self.logger.error("Welcome panel refresh failed: %s", e)
        finally:
            self._welcome_refresh_busy = False

    def _request_ui_refresh_from_raise(self) -> None:
        """AES / second-instance RAISE: prove the UI is alive and redraw Welcome."""
        try:
            # Clear sticky gates that can leave Welcome looking frozen.
            self._expanding_to_screen = False
            if getattr(self, "_welcome_refresh_busy", False):
                self._welcome_refresh_busy = False
            self._touch_last_refresh_label()
            self._schedule_welcome_refresh(allow_db_fallback=False)
            # Nudge Outlook if the last scrape is stale (>2 min) or never ran.
            last = getattr(self, "_last_outlook_scrape_at", None)
            stale = last is None or (datetime.now() - last).total_seconds() > 120
            if stale and not getattr(self, "_scrape_busy", False):
                QTimer.singleShot(250, lambda: self._start_outlook_scrape(manual=False))
        except Exception as exc:
            self.logger.debug("raise refresh failed: %s", exc)
    def _mail_is_important_domain(self, mail: Dict[str, Any]) -> bool:
        """True when sender matches Important Domains or a learned VIP sender."""
        from guri_learning import is_important_sender

        sender = str((mail or {}).get("sender") or "")
        if sender_matches_important_domain(sender, self.important_domains):
            return True
        return is_important_sender(getattr(self, "learning", {}) or {}, sender)

    def _timeline_importance_tier(self, det: Optional[Dict[str, Any]]) -> str:
        """vip | domain | elevated | normal — drives timeline circle colour."""
        from guri_learning import is_important_sender

        det = det if isinstance(det, dict) else {}
        if det.get("rule_force_vip"):
            return "vip"
        mail = det.get("mail") or {}
        if not mail and det.get("sender"):
            mail = det
        sender = str(mail.get("sender") or "")

        if is_important_sender(getattr(self, "learning", {}) or {}, sender):
            return "vip"
        if sender_matches_important_domain(sender, self.important_domains):
            return "domain"

        scores = compute_relevance_scores(
            {
                "mail": mail,
                "score": det.get("score") or 0,
                "actions": det.get("actions") or [],
            },
            getattr(self, "relevance_feedback", {}) or {},
            important_domains=self.important_domains,
            important_senders=self._important_senders_list(),
        )
        boost = int(det.get("rule_boost_sender") or 0)
        sender_rel = int(scores.get("sender_relevance") or 0) + boost
        importance = str(det.get("importance") or "").lower()
        if (
            importance == "high"
            or sender_rel >= 70
            or int(scores.get("email_relevance") or 0) >= 55
            or int(det.get("score") or 0) >= 40
        ):
            return "elevated"
        return "normal"

    def _important_mail_dedupe_key(self, det: Dict[str, Any]) -> str:
        mail = det.get("mail") or {}
        guri = str(mail.get("guri") or det.get("guri") or "").strip()
        if guri:
            return f"guri:{guri}"
        entry = str(mail.get("entry_id") or "").strip()
        if entry:
            return f"entry:{entry}"
        return "|".join(
            [
                str(mail.get("sender") or "").strip().lower(),
                str(mail.get("subject") or "").strip().lower(),
                str(mail.get("received") or mail.get("datetime") or "").strip(),
            ]
        )

    def _db_important_email_candidates(self) -> List[Dict[str, Any]]:
        """Targeted GURI DB rows for Important domains / VIP senders (not a 10k dump)."""
        if not self.db:
            return []
        records: List[Dict[str, Any]] = []
        seen: set = set()
        domains = list(self.important_domains or [])
        if domains and hasattr(self.db, "search_by_sender_domains"):
            try:
                for record in self.db.search_by_sender_domains(
                    domains, limit_per_domain=40
                ):
                    key = str(record.get("guri") or record.get("id") or "")
                    if key and key in seen:
                        continue
                    if key:
                        seen.add(key)
                    records.append(record)
            except Exception as exc:
                self.logger.warning("Important-domain DB search failed: %s", exc)
        for vip in self._important_senders_list()[:30]:
            addr = str(vip or "").strip()
            if not addr:
                continue
            try:
                for record in self.db.search_by_sender(addr, limit=20):
                    if str(record.get("document_type") or "") != "01":
                        continue
                    if not self._mail_is_important_domain(
                        {"sender": record.get("sender")}
                    ):
                        continue
                    key = str(record.get("guri") or record.get("id") or "")
                    if key and key in seen:
                        continue
                    if key:
                        seen.add(key)
                    records.append(record)
            except Exception as exc:
                self.logger.warning("VIP sender DB search failed: %s", exc)

        candidates: List[Dict[str, Any]] = []
        for record in records:
            candidates.append(
                {
                    "mail": {
                        "sender": record.get("sender"),
                        "to": record.get("recipients"),
                        "subject": record.get("subject"),
                        "received": record.get("datetime"),
                        "guri": record.get("guri"),
                        "entry_id": "",
                        "store_id": "",
                    },
                    "actions": [],
                    "reasons": ["Important domain / VIP (GURI DB)"],
                    "importance": "medium",
                    "score": 20,
                    "status": self.email_status.get(
                        str(record.get("guri") or ""), "—"
                    ),
                    "guri": record.get("guri"),
                }
            )
        return candidates

    def _refresh_important_emails(
        self, *, allow_db_fallback: bool = True, refresh_actions: bool = True
    ):
        """Refresh Important Emails — scrape matches + targeted DB supplement."""
        try:
            self.important_tree.clear()
            if not self.important_domains and not self._important_senders_list():
                self.status_bar.showMessage(
                    "Important emails: add a domain above (or endorse a sender) to filter this list"
                )
                if refresh_actions:
                    self._refresh_actions_panel()
                return

            items = self._feedback_items()
            # Domain-based + learned important senders from live scrape
            important = [
                it
                for it in items
                if self._mail_is_important_domain(it.get("mail") or {})
                and not it.get("rule_hide_from_important")
            ]
            seen_keys = {self._important_mail_dedupe_key(it) for it in important}

            # Targeted domain/VIP DB supplement (cheap). The old path only ran when
            # scrape_cache was empty, so Matching emails went blank after any scrape.
            # allow_db_fallback kept for API compat; this query is always safe to run.
            _ = allow_db_fallback
            if self.db:
                db_items = self._feedback_items(self._db_important_email_candidates())
                for it in db_items:
                    if it.get("rule_hide_from_important"):
                        continue
                    if not self._mail_is_important_domain(it.get("mail") or {}):
                        continue
                    key = self._important_mail_dedupe_key(it)
                    if key in seen_keys:
                        continue
                    seen_keys.add(key)
                    important.append(it)

            important.sort(
                key=lambda x: (
                    int(x.get("score") or 0),
                    str((x.get("mail") or {}).get("received") or ""),
                ),
                reverse=True,
            )
            for det in important[:250]:
                mail = det.get("mail") or {}
                item = QTreeWidgetItem(self.important_tree)
                item.setText(0, str(mail.get("sender") or ""))
                item.setText(1, str(mail.get("to") or mail.get("cc") or ""))
                subject = str(mail.get("subject") or "")
                item.setText(2, subject[:120] + "..." if len(subject) > 120 else subject)
                item.setText(3, str(mail.get("received") or ""))
                item.setText(4, str(det.get("status") or "—"))
                item.setData(0, Qt.ItemDataRole.UserRole, det)

            self.status_bar.showMessage(
                f"Important emails: {len(important)} from configured domains"
            )
            if refresh_actions:
                self._refresh_actions_panel()
        except Exception as e:
            self.logger.error(f"Error refreshing important emails: {e}")

    def _refresh_actions_panel(self) -> None:
        """Populate Actions tree from scrape cache."""
        if not hasattr(self, "actions_tree"):
            return
        try:
            self.actions_tree.clear()
            items = self._feedback_items()
            actions = [
                it
                for it in items
                if it.get("actions") and not it.get("rule_hide_from_actions")
            ]
            actions.sort(key=lambda x: int(x.get("score") or 0), reverse=True)
            for det in actions[:250]:
                mail = det.get("mail") or {}
                item = QTreeWidgetItem(self.actions_tree)
                item.setText(0, str(mail.get("sender") or "")[:80])
                subject = str(mail.get("subject") or "")
                item.setText(1, subject[:120] + ("..." if len(subject) > 120 else ""))
                item.setText(2, ", ".join(det.get("actions") or []))
                item.setText(3, str(det.get("importance") or ""))
                item.setText(4, str(mail.get("received") or ""))
                item.setText(5, str(det.get("status") or "—"))
                item.setData(0, Qt.ItemDataRole.UserRole, det)
            if hasattr(self, "actions_hint") and actions:
                self.actions_hint.setText(
                    f"{len(actions)} action item(s) — double-click or use Open in Outlook."
                )
        except Exception as e:
            self.logger.error(f"Error refreshing actions panel: {e}")

    def _show_email_context_menu(self, position):
        """Show context menu for email actions."""
        item = self.important_tree.itemAt(position)
        if item:
            self.important_tree.setCurrentItem(item)
            menu = QMenu(self)
            menu.addAction("Mark as Need to Reply", lambda: self._mark_item_status(item, "Need Reply"))
            menu.addAction("Mark as Replied", lambda: self._mark_item_status(item, "Replied"))
            menu.addAction("Mark as Done", lambda: self._mark_item_status(item, "Done"))
            menu.addSeparator()
            det = item.data(0, Qt.ItemDataRole.UserRole) or {}
            mail = det.get("mail") or {}
            sender = str(mail.get("sender") or "")
            from guri_learning import is_important_sender

            if sender and is_important_sender(self.learning, sender):
                menu.addAction(
                    "Unmark VIP sender",
                    lambda: self._toggle_vip_for_sender(sender),
                )
            else:
                menu.addAction(
                    "Mark VIP sender",
                    lambda: self._toggle_vip_for_sender(sender),
                )
            menu.addSeparator()
            menu.addAction("Open in Outlook", lambda: self._open_detected_in_outlook(item))
            menu.addAction("View Details", lambda: self._view_email_details(item))
            menu.exec(self.important_tree.mapToGlobal(position))

    def _toggle_vip_for_sender(self, sender: str) -> None:
        addr = normalize_sender_address(sender)
        if not addr:
            return
        from guri_learning import is_important_sender

        if is_important_sender(self.learning, addr):
            self.learning = forget_important_sender(self.learning, addr)
            msg = f"Removed VIP: {addr}"
        else:
            self.learning = remember_important_sender(self.learning, addr)
            msg = f"Marked VIP: {addr}"
        self._save_learning()
        self._refresh_vip_list()
        self._refresh_important_emails(allow_db_fallback=False, refresh_actions=True)
        preview = self._preview_det
        if preview is not None:
            self._show_email_preview(preview)
        self.status_bar.showMessage(msg)

    def _show_action_context_menu(self, position):
        item = self.actions_tree.itemAt(position)
        if not item:
            return
        self.actions_tree.setCurrentItem(item)
        menu = QMenu(self)
        menu.addAction("Mark as Need to Reply", lambda: self._mark_item_status(item, "Need Reply", actions_tree=True))
        menu.addAction("Mark as Done", lambda: self._mark_item_status(item, "Done", actions_tree=True))
        menu.addAction("Mark as Replied", lambda: self._mark_item_status(item, "Replied", actions_tree=True))
        menu.addSeparator()
        menu.addAction("Open in Outlook", lambda: self._open_detected_in_outlook(item))
        menu.addAction("View Details", lambda: self._view_email_details(item))
        menu.exec(self.actions_tree.mapToGlobal(position))

    def _mark_item_status(self, item, status: str, actions_tree: bool = False) -> None:
        det = item.data(0, Qt.ItemDataRole.UserRole) or {}
        mail = det.get("mail") or det
        entry_id = str(mail.get("entry_id") or "")
        guri_key = str(mail.get("guri") or det.get("guri") or "")
        if entry_id:
            self.action_status[entry_id] = status
            save_action_status(self.action_status)
        if guri_key:
            self.email_status[guri_key] = status
            self._save_email_status()
        # Status column index differs by tree
        tree = item.treeWidget()
        if tree is getattr(self, "deadlines_tree", None):
            col = 6
        elif actions_tree or tree is getattr(self, "actions_tree", None):
            col = 5
        else:
            col = 4
        item.setText(col, status)
        det["status"] = status
        item.setData(0, Qt.ItemDataRole.UserRole, det)
        self.status_bar.showMessage(f"Marked as '{status}'")

    def _mark_need_reply(self, item):
        self._mark_item_status(item, "Need Reply")

    def _mark_replied(self, item):
        self._mark_item_status(item, "Replied")

    def _on_mail_item_selected(self, item, _column=None) -> None:
        """Show selected mail on the right with action-language highlights."""
        if item is None:
            return
        det = item.data(0, Qt.ItemDataRole.UserRole) or {}
        self._show_email_preview(det if isinstance(det, dict) else {})

    def _feedback_items(
        self,
        items: Optional[List[Dict[str, Any]]] = None,
        *,
        drop_ignored: bool = True,
    ) -> List[Dict[str, Any]]:
        """Status + relevance feedback + compiled user-rule enforcement."""
        source = items if items is not None else self.scrape_cache_items
        annotated = apply_statuses(
            [dict(x) for x in (source or [])],
            {**self.email_status, **self.action_status},
        )
        annotated = apply_relevance_feedback(
            annotated,
            getattr(self, "relevance_feedback", {}),
            important_domains=self.important_domains,
            important_senders=self._important_senders_list(),
            drop_ignored=drop_ignored,
        )
        return apply_user_rules_to_items(
            annotated, getattr(self, "learning", {}) or {}
        )

    def _save_relevance_feedback(self) -> None:
        try:
            save_relevance_feedback(self.relevance_feedback)
        except Exception as exc:
            self.logger.warning("Failed saving relevance feedback: %s", exc)

    def _preview_mail_context(self) -> Optional[Dict[str, Any]]:
        det = self._preview_det or {}
        if not det:
            return None
        mail = det.get("mail") or {}
        if not mail and det.get("sender"):
            mail = det
        if not mail:
            return None
        return {"det": det, "mail": mail}

    def _clear_email_preview(self) -> None:
        self._preview_det = None
        if hasattr(self, "preview_meta"):
            self.preview_meta.setText("Select an email from Important or Actions.")
        if hasattr(self, "preview_reasons"):
            self.preview_reasons.setText("")
        if hasattr(self, "email_preview"):
            self.email_preview.clear()
        if hasattr(self, "preview_sender_score"):
            self.preview_sender_score.setText("—")
        if hasattr(self, "preview_email_score"):
            self.preview_email_score.setText("—")
        self._sync_preview_popout()

    def _toggle_preview_maximize(self) -> None:
        """Expand / restore the Welcome email preview pane."""
        if not hasattr(self, "preview_frame"):
            return
        maximized = bool(getattr(self, "_preview_maximized", False))
        if not maximized:
            split = getattr(self, "welcome_main_split", None)
            if split is not None:
                self._preview_split_sizes = list(split.sizes())
            left = getattr(self, "welcome_left_widget", None)
            stats = getattr(self, "welcome_stats_frame", None)
            if left is not None:
                left.hide()
            if stats is not None:
                stats.hide()
            self._preview_maximized = True
            self.preview_max_btn.setText("Minimize")
            self.preview_max_btn.setToolTip("Restore the Welcome layout")
            self.status_bar.showMessage("Email preview maximized")
        else:
            left = getattr(self, "welcome_left_widget", None)
            stats = getattr(self, "welcome_stats_frame", None)
            if stats is not None:
                stats.show()
            if left is not None:
                left.show()
            split = getattr(self, "welcome_main_split", None)
            sizes = getattr(self, "_preview_split_sizes", None)
            if split is not None and sizes:
                split.setSizes(sizes)
            self._preview_maximized = False
            self.preview_max_btn.setText("Maximize")
            self.preview_max_btn.setToolTip("Expand email preview over the Welcome page")
            self.status_bar.showMessage("Email preview restored")

    def _popout_email_preview(self) -> None:
        """Spawn (or focus) a separate window for the email preview."""
        existing = getattr(self, "_preview_popout", None)
        if existing is not None and existing.isVisible():
            existing.raise_()
            existing.activateWindow()
            self._sync_preview_popout()
            return

        dialog = QDialog(self)
        dialog.setWindowTitle("Email preview")
        dialog.setWindowFlag(Qt.WindowType.Window, True)
        dialog.resize(720, 640)
        layout = QVBoxLayout(dialog)

        meta = QLabel()
        meta.setWordWrap(True)
        meta.setTextFormat(Qt.TextFormat.RichText)
        meta.setStyleSheet("color: #4a6570; font-size: 11px;")
        layout.addWidget(meta)

        reasons = QLabel()
        reasons.setWordWrap(True)
        reasons.setStyleSheet("color: #0f6b7c; font-size: 11px;")
        layout.addWidget(reasons)

        body = QTextEdit()
        body.setReadOnly(True)
        layout.addWidget(body, stretch=1)

        btn_row = QHBoxLayout()
        for label, handler in (
            ("Open in Outlook", self._open_preview_in_outlook),
            ("Ignore Email", self._preview_ignore_email),
            ("Ignore Recipient", self._preview_ignore_recipient),
            ("Deprecate Email", self._preview_deprecate_email),
            ("Deprecate Sender", self._preview_deprecate_sender),
        ):
            btn = QPushButton(label)
            btn.clicked.connect(handler)
            btn_row.addWidget(btn)
        btn_row.addStretch(1)
        layout.addLayout(btn_row)

        score_row = QHBoxLayout()
        sender_lbl = QLabel("Sender Relevance")
        sender_val = QLabel("—")
        sender_val.setStyleSheet("font-weight: 700; color: #0f6b7c;")
        email_lbl = QLabel("Email Relevance")
        email_val = QLabel("—")
        email_val.setStyleSheet("font-weight: 700; color: #0f6b7c;")
        score_row.addWidget(sender_lbl)
        score_row.addWidget(sender_val)
        tick_s = QPushButton("✓")
        tick_s.setObjectName("scoreTick")
        tick_s.setFixedWidth(28)
        tick_s.setToolTip("Endorse Sender Ranking (gradually increase)")
        tick_s.clicked.connect(self._preview_endorse_sender_relevance)
        cross_s = QPushButton("✗")
        cross_s.setObjectName("scoreCross")
        cross_s.setFixedWidth(28)
        cross_s.clicked.connect(self._preview_cross_relevance)
        score_row.addWidget(tick_s)
        score_row.addWidget(cross_s)
        score_row.addSpacing(16)
        score_row.addWidget(email_lbl)
        score_row.addWidget(email_val)
        tick_e = QPushButton("✓")
        tick_e.setObjectName("scoreTick")
        tick_e.setFixedWidth(28)
        tick_e.setToolTip("Endorse Email Ranking (gradually increase)")
        tick_e.clicked.connect(self._preview_endorse_email_relevance)
        cross_e = QPushButton("✗")
        cross_e.setObjectName("scoreCross")
        cross_e.setFixedWidth(28)
        cross_e.clicked.connect(self._preview_cross_relevance)
        score_row.addWidget(tick_e)
        score_row.addWidget(cross_e)
        score_row.addStretch(1)
        rule_btn = QPushButton("Describe a rule…")
        rule_btn.setToolTip(
            "Write a natural-language rule for Ollama to apply."
        )
        rule_btn.clicked.connect(self._show_describe_rule_dialog)
        score_row.addWidget(rule_btn)
        layout.addLayout(score_row)

        dialog._pop_meta = meta  # type: ignore[attr-defined]
        dialog._pop_reasons = reasons  # type: ignore[attr-defined]
        dialog._pop_body = body  # type: ignore[attr-defined]
        dialog._pop_sender_score = sender_val  # type: ignore[attr-defined]
        dialog._pop_email_score = email_val  # type: ignore[attr-defined]

        def _on_close(_event=None) -> None:
            if getattr(self, "_preview_popout", None) is dialog:
                self._preview_popout = None

        dialog.finished.connect(lambda *_: _on_close())
        self._preview_popout = dialog
        self._sync_preview_popout()
        dialog.show()
        dialog.raise_()
        dialog.activateWindow()

    def _sync_preview_popout(self) -> None:
        """Mirror docked preview content into the pop-out window when open."""
        dialog = getattr(self, "_preview_popout", None)
        if dialog is None or not dialog.isVisible():
            return
        meta = getattr(dialog, "_pop_meta", None)
        reasons = getattr(dialog, "_pop_reasons", None)
        body = getattr(dialog, "_pop_body", None)
        sender_score = getattr(dialog, "_pop_sender_score", None)
        email_score = getattr(dialog, "_pop_email_score", None)
        if meta is not None and hasattr(self, "preview_meta"):
            meta.setText(self.preview_meta.text())
        if reasons is not None and hasattr(self, "preview_reasons"):
            reasons.setText(self.preview_reasons.text())
        if body is not None and hasattr(self, "email_preview"):
            body.setHtml(self.email_preview.toHtml())
        if sender_score is not None and hasattr(self, "preview_sender_score"):
            sender_score.setText(self.preview_sender_score.text())
        if email_score is not None and hasattr(self, "preview_email_score"):
            email_score.setText(self.preview_email_score.text())
        det = getattr(self, "_preview_det", None) or {}
        mail = det.get("mail") or det
        subject = str(mail.get("subject") or "Email preview")
        dialog.setWindowTitle(f"Email preview — {subject[:80]}")

    def _refresh_after_feedback(self, *, clear_preview: bool = False) -> None:
        self._save_relevance_feedback()
        if clear_preview:
            self._clear_email_preview()
        self._schedule_welcome_refresh(allow_db_fallback=False)

    def _preview_ignore_email(self) -> None:
        ctx = self._preview_mail_context()
        if not ctx:
            QMessageBox.information(self, "Ignore Email", "Select an email in the preview first.")
            return
        key = mail_feedback_key(ctx["mail"], ctx["det"])
        self.relevance_feedback = ignore_email(self.relevance_feedback, key)
        self._refresh_after_feedback(clear_preview=True)
        self.status_bar.showMessage("Ignored email — removed from prioritisation")

    def _preview_ignore_recipient(self) -> None:
        """Ignore counterparty (From:) — removes them from prioritisation lists."""
        ctx = self._preview_mail_context()
        if not ctx:
            QMessageBox.information(
                self, "Ignore Recipient", "Select an email in the preview first."
            )
            return
        sender = str(ctx["mail"].get("sender") or "")
        addr = normalize_sender_address(sender)
        if not addr:
            QMessageBox.information(
                self, "Ignore Recipient", "Could not read a recipient address."
            )
            return
        self.relevance_feedback = ignore_recipient(self.relevance_feedback, sender)
        self._refresh_after_feedback(clear_preview=True)
        self.status_bar.showMessage(f"Ignored recipient {addr} — removed from prioritisation")

    def _preview_ignore_sender(self) -> None:
        # Back-compat alias for older menu/call sites
        self._preview_ignore_recipient()

    def _preview_deprecate_email(self) -> None:
        ctx = self._preview_mail_context()
        if not ctx:
            QMessageBox.information(
                self, "Deprecate Email", "Select an email in the preview first."
            )
            return
        key = mail_feedback_key(ctx["mail"], ctx["det"])
        self.relevance_feedback = deprecate_email(self.relevance_feedback, key, penalty=25)
        self._refresh_after_feedback()
        self._show_email_preview(ctx["det"])
        self.status_bar.showMessage("Deprecated email — scored lower, still listed")

    def _preview_deprecate_sender(self) -> None:
        ctx = self._preview_mail_context()
        if not ctx:
            QMessageBox.information(
                self, "Deprecate Sender", "Select an email in the preview first."
            )
            return
        sender = str(ctx["mail"].get("sender") or "")
        addr = normalize_sender_address(sender)
        if not addr:
            QMessageBox.information(
                self, "Deprecate Sender", "Could not read a sender address."
            )
            return
        self.relevance_feedback = deprecate_sender(
            self.relevance_feedback, sender, penalty=30
        )
        self._refresh_after_feedback()
        self._show_email_preview(ctx["det"])
        self.status_bar.showMessage(f"Deprecated sender {addr} — scored lower, still listed")

    def _preview_endorse_email_relevance(self) -> None:
        ctx = self._preview_mail_context()
        if not ctx:
            QMessageBox.information(self, "Email Relevance", "Select an email in the preview first.")
            return
        key = mail_feedback_key(ctx["mail"], ctx["det"])
        self.relevance_feedback = endorse_email_score(self.relevance_feedback, key)
        self._refresh_after_feedback()
        scores = compute_relevance_scores(
            ctx["det"],
            self.relevance_feedback,
            important_domains=self.important_domains,
            important_senders=self._important_senders_list(),
        )
        self._show_email_preview(ctx["det"])
        self.status_bar.showMessage(
            f"Endorsed email ranking — relevance now {scores['email_relevance']}"
        )

    def _preview_endorse_sender_relevance(self) -> None:
        ctx = self._preview_mail_context()
        if not ctx:
            QMessageBox.information(self, "Sender Relevance", "Select an email in the preview first.")
            return
        sender = str(ctx["mail"].get("sender") or "")
        self.relevance_feedback = endorse_sender_score(self.relevance_feedback, sender)
        # Occasional learning prompt: promote to important sender
        if should_ask_important_sender(getattr(self, "learning", {}), sender):
            ask = QMessageBox.question(
                self,
                "Important sender?",
                f"Did you mean to treat this sender as important going forward?\n\n{sender}",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.Yes,
            )
            if ask == QMessageBox.StandardButton.Yes:
                self.learning = remember_important_sender(self.learning, sender)
                domain = domain_from_sender(sender)
                if domain and domain not in self.important_domains:
                    add_dom = QMessageBox.question(
                        self,
                        "Important Emails domain?",
                        f"Also add @{domain} to Important Emails domains?",
                        QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                        QMessageBox.StandardButton.Yes,
                    )
                    if add_dom == QMessageBox.StandardButton.Yes:
                        self.important_domains.append(domain)
                        self._save_important_domains()
                        if hasattr(self, "domains_listbox"):
                            self._refresh_domains_list()
            else:
                self.learning = mark_asked_important_sender(self.learning, sender)
            self._save_learning()
        self._refresh_after_feedback()
        scores = compute_relevance_scores(
            ctx["det"],
            self.relevance_feedback,
            important_domains=self.important_domains,
            important_senders=self._important_senders_list(),
        )
        self._show_email_preview(ctx["det"])
        self.status_bar.showMessage(
            f"Endorsed sender ranking — relevance now {scores['sender_relevance']}"
        )

    def _preview_cross_relevance(self) -> None:
        """Cross on either score — choose how to demote / remove."""
        ctx = self._preview_mail_context()
        if not ctx:
            QMessageBox.information(self, "De-rate", "Select an email in the preview first.")
            return

        dialog = QDialog(self)
        dialog.setWindowTitle("De-rate or remove")
        dialog.setModal(True)
        layout = QVBoxLayout(dialog)
        layout.addWidget(
            QLabel(
                "How should GURI treat this message?\n"
                "Choose one option:"
            )
        )

        choice = {"value": ""}

        def _pick(value: str) -> None:
            choice["value"] = value
            dialog.accept()

        btn_content = QPushButton("De-Rate content type")
        btn_content.clicked.connect(lambda: _pick("content"))
        layout.addWidget(btn_content)

        btn_sender = QPushButton("De-rate Sender")
        btn_sender.clicked.connect(lambda: _pick("sender"))
        layout.addWidget(btn_sender)

        btn_spam = QPushButton("This is Spam, take it away!")
        btn_spam.clicked.connect(lambda: _pick("spam"))
        layout.addWidget(btn_spam)

        cancel = QPushButton("Cancel")
        cancel.clicked.connect(dialog.reject)
        add_dialog_button_row(layout, cancel=cancel)

        if dialog.exec() != QDialog.DialogCode.Accepted or not choice["value"]:
            return

        mail = ctx["mail"]
        det = ctx["det"]
        key = mail_feedback_key(mail, det)
        sender = str(mail.get("sender") or "")

        if choice["value"] == "content":
            actions = det.get("actions") or []
            if not actions:
                # Fall back to a generic content demotion bucket
                actions = ["general"]
            self.relevance_feedback = derate_content_types(
                self.relevance_feedback, actions, penalty=20
            )
            self._refresh_after_feedback()
            self._show_email_preview(det)
            self.status_bar.showMessage(
                "De-rated content type(s): " + ", ".join(str(a) for a in actions)
            )
        elif choice["value"] == "sender":
            self.relevance_feedback = derate_sender(
                self.relevance_feedback, sender, penalty=30
            )
            self._refresh_after_feedback()
            self._show_email_preview(det)
            self.status_bar.showMessage(
                f"De-rated sender {normalize_sender_address(sender) or sender}"
            )
        else:
            self.relevance_feedback = mark_spam(
                self.relevance_feedback, mail_key=key, sender=sender
            )
            self._refresh_after_feedback(clear_preview=True)
            self.status_bar.showMessage("Marked as spam and removed from Welcome lists")

    def _show_describe_rule_dialog(self) -> None:
        """Capture a natural-language rule for Ollama (bottom-right of preview)."""
        det = getattr(self, "_preview_det", None) or {}
        mail = det.get("mail") if isinstance(det, dict) else {}
        mail = mail or {}
        key = mail_feedback_key(mail, det) if mail else ""
        sender = str(mail.get("sender") or "")
        subject = str(mail.get("subject") or "")

        dialog = QDialog(self)
        dialog.setWindowTitle("Describe a rule")
        dialog.resize(520, 420)
        layout = QVBoxLayout(dialog)

        intro = QLabel(
            "Describe how GURI should treat this kind of email. "
            "Save keeps the text for Ollama. "
            "Save & compile asks Ollama to turn it into local match/effects "
            "so Timeline / Important / Actions / Deadlines can enforce it."
        )
        intro.setWordWrap(True)
        intro.setStyleSheet(f"color: {PALETTE['muted']};")
        layout.addWidget(intro)

        if subject or sender:
            ctx = QLabel(
                f"<b>Context</b><br/>From: {sender or '—'}<br/>"
                f"Subject: {subject or '—'}"
            )
            ctx.setWordWrap(True)
            ctx.setTextFormat(Qt.TextFormat.RichText)
            layout.addWidget(ctx)

        form = QFormLayout()
        scope_combo = QComboBox()
        scope_combo.addItem("Always (all mail)", "always")
        scope_combo.addItem("This sender only", "sender")
        scope_combo.addItem("This email only", "email")
        form.addRow("Apply to", scope_combo)

        rule_edit = QTextEdit()
        rule_edit.setPlaceholderText(
            "Examples:\n"
            "• Treat catch-up / networking emails as relevant but not actionable\n"
            "• Surrey O-RAN threads are high priority — always surface reply\n"
            "• Ignore calendar FYIs that only say “for information”"
        )
        rule_edit.setMinimumHeight(140)
        form.addRow("Rule", rule_edit)
        layout.addLayout(form)

        existing = format_user_rules_for_llm(
            getattr(self, "learning", {}) or {},
            mail_key=key,
            sender=sender,
        )
        if existing:
            prev = QLabel(existing.replace("\n", "<br/>"))
            prev.setWordWrap(True)
            prev.setTextFormat(Qt.TextFormat.RichText)
            prev.setStyleSheet(f"color: {PALETTE['muted']}; font-size: 11px;")
            layout.addWidget(prev)

        save_btn = QPushButton("Save rule")
        compile_btn = QPushButton("Save & compile")
        compile_btn.setDefault(True)
        compile_btn.setToolTip(
            "Save, then ask Ollama to compile matchers + effects for local enforcement."
        )
        apply_btn = QPushButton("Save & ask Ollama")
        apply_btn.setEnabled(bool(mail.get("subject") or mail.get("body_preview")))
        cancel_btn = QPushButton("Cancel")
        cancel_btn.clicked.connect(dialog.reject)
        add_dialog_button_row(
            layout,
            affirmative=[save_btn, compile_btn, apply_btn],
            cancel=cancel_btn,
        )

        def _persist(*, compiled_bits: Optional[Dict[str, Any]] = None) -> Optional[str]:
            text = rule_edit.toPlainText().strip()
            if len(text) < 3:
                QMessageBox.information(
                    dialog, "Describe a rule", "Write a short rule first."
                )
                return None
            scope = str(scope_combo.currentData() or "always")
            if scope == "email" and not key:
                QMessageBox.information(
                    dialog,
                    "Describe a rule",
                    "Select an email in the preview before scoping to this email.",
                )
                return None
            if scope == "sender" and not sender:
                QMessageBox.information(
                    dialog,
                    "Describe a rule",
                    "Select an email with a sender, or choose Always.",
                )
                return None
            match = (compiled_bits or {}).get("match")
            effects = (compiled_bits or {}).get("effects")
            self.learning = add_user_rule(
                getattr(self, "learning", {}) or {},
                text,
                scope=scope,
                mail_key=key if scope == "email" else "",
                sender=sender if scope in {"sender", "email"} else "",
                match=match,
                effects=effects,
                compiled=bool(compiled_bits),
            )
            self._save_learning()
            return text

        def on_save() -> None:
            if _persist() is None:
                return
            dialog.accept()
            self.status_bar.showMessage("Rule saved (text for Ollama; not compiled yet)")

        def on_apply() -> None:
            if _persist() is None:
                return
            dialog.accept()
            for i in range(self.notebook.count()):
                if self.notebook.tabText(i) == "Ollama":
                    self.notebook.setCurrentIndex(i)
                    break
            self._ollama_analyze_selected_email()
            self.status_bar.showMessage("Rule saved — asking Ollama…")

        def on_compile() -> None:
            text = rule_edit.toPlainText().strip()
            if len(text) < 3:
                QMessageBox.information(
                    dialog, "Describe a rule", "Write a short rule first."
                )
                return
            scope = str(scope_combo.currentData() or "always")
            dialog.accept()
            self._start_rule_compile(
                text=text,
                scope=scope,
                mail_key=key if scope == "email" else "",
                sender=sender if scope in {"sender", "email"} else sender,
                subject=subject,
                body=str(mail.get("body_preview") or ""),
            )

        save_btn.clicked.connect(on_save)
        compile_btn.clicked.connect(on_compile)
        apply_btn.clicked.connect(on_apply)
        dialog.exec()

    def _start_rule_compile(
        self,
        *,
        text: str,
        scope: str,
        mail_key: str,
        sender: str,
        subject: str,
        body: str,
    ) -> None:
        if self._rule_compile_busy or self._ollama_busy:
            self.status_bar.showMessage("Ollama is busy — try compile again in a moment")
            return
        cfg = self.ollama_cfg if isinstance(getattr(self, "ollama_cfg", None), dict) else {}
        try:
            cfg = self._ollama_collect_cfg() if hasattr(self, "ollama_url_edit") else ollama_load_config()
        except Exception:
            cfg = ollama_load_config()
        if not cfg.get("model"):
            QMessageBox.warning(
                self,
                "Compile rule",
                "Set an Ollama model on the Ollama tab first, then compile again.",
            )
            # Still save text-only so the rule isn't lost
            self.learning = add_user_rule(
                getattr(self, "learning", {}) or {},
                text,
                scope=scope,
                mail_key=mail_key,
                sender=sender,
            )
            self._save_learning()
            return

        self._rule_compile_busy = True
        self._rule_compile_pending = {
            "text": text,
            "scope": scope,
            "mail_key": mail_key,
            "sender": sender,
        }
        prompt = build_rule_compile_prompt(
            text,
            scope=scope,
            sender=sender,
            subject=subject,
            body_preview=body,
        )
        self.status_bar.showMessage("Compiling rule with Ollama…")

        def work() -> None:
            try:
                raw = ollama_chat(
                    prompt,
                    base_url=cfg["base_url"],
                    model=cfg["model"],
                    system=(
                        "You compile GURI triage rules to JSON only. "
                        "No prose outside the JSON object."
                    ),
                    timeout=float(cfg.get("timeout") or 120),
                )
                parsed = parse_compiled_rule_json(raw)
                if not parsed:
                    self._rule_compile_bridge.finished.emit(None, raw or "empty response")
                else:
                    self._rule_compile_bridge.finished.emit(parsed, "")
            except Exception as exc:
                self._rule_compile_bridge.finished.emit(None, str(exc))

        threading.Thread(target=work, daemon=True).start()

    def _on_rule_compile_finished(self, parsed_obj, error: str) -> None:
        self._rule_compile_busy = False
        pending = getattr(self, "_rule_compile_pending", None) or {}
        self._rule_compile_pending = None
        text = str(pending.get("text") or "")
        if not text:
            return
        if error or not isinstance(parsed_obj, dict):
            self.learning = add_user_rule(
                getattr(self, "learning", {}) or {},
                text,
                scope=str(pending.get("scope") or "always"),
                mail_key=str(pending.get("mail_key") or ""),
                sender=str(pending.get("sender") or ""),
            )
            self._save_learning()
            QMessageBox.warning(
                self,
                "Compile rule",
                "Could not compile with Ollama — saved as text-only.\n\n"
                + str(error or "")[:500],
            )
            return

        match = normalize_rule_match(parsed_obj.get("match"))
        effects = normalize_rule_effects(parsed_obj.get("effects"))
        notes = str(parsed_obj.get("notes") or "").strip()
        summary = summarize_rule_effects(effects)
        detail = (
            f"Rule:\n{text}\n\n"
            f"Local effects:\n{summary}\n\n"
            f"Matchers:\n"
            f"  domains: {', '.join(match.get('domains') or []) or '—'}\n"
            f"  senders: {', '.join(match.get('senders') or []) or '—'}\n"
            f"  subject: {', '.join(match.get('subject_contains') or []) or '—'}\n"
            f"  body: {', '.join(match.get('body_contains') or []) or '—'}\n"
            f"  keywords: {', '.join(match.get('any_keywords') or []) or '—'}\n"
        )
        if notes:
            detail += f"\nNotes: {notes}"
        reply = QMessageBox.question(
            self,
            "Accept compiled rule?",
            detail + "\n\nLock this in for local enforcement?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.Yes,
        )
        compiled = reply == QMessageBox.StandardButton.Yes and rule_has_enforceable_effects(
            effects
        )
        self.learning = add_user_rule(
            getattr(self, "learning", {}) or {},
            text,
            scope=str(pending.get("scope") or "always"),
            mail_key=str(pending.get("mail_key") or ""),
            sender=str(pending.get("sender") or ""),
            match=match if compiled else {},
            effects=effects if compiled else {},
            compiled=compiled,
        )
        self._save_learning()
        if compiled:
            self.status_bar.showMessage("Compiled rule locked — refreshing triage")
            self._schedule_welcome_refresh(allow_db_fallback=False)
            try:
                self._refresh_deadlines_panel()
            except Exception:
                pass
        else:
            self.status_bar.showMessage("Rule saved as text-only (not compiled)")

    def _show_email_preview(self, det: Dict[str, Any]) -> None:
        if not hasattr(self, "email_preview"):
            return
        # Re-apply compiled rules so explanations stay fresh
        try:
            det = apply_user_rules_to_item(
                dict(det or {}), getattr(self, "learning", {}) or {}
            )
        except Exception:
            det = det if isinstance(det, dict) else {}
        self._preview_det = det
        mail = det.get("mail") or {}
        if not mail and det.get("sender"):
            mail = det

        sender = str(mail.get("sender") or "")
        subject = str(mail.get("subject") or "")
        when = str(mail.get("received") or "")
        account = str(mail.get("account_smtp") or "")
        role = str(mail.get("recipient_role") or "")
        actions = ", ".join(det.get("actions") or []) or "—"
        importance = str(det.get("importance") or "—")
        scores = compute_relevance_scores(
            det,
            getattr(self, "relevance_feedback", {}),
            important_domains=self.important_domains,
            important_senders=self._important_senders_list(),
        )
        score = scores["email_relevance"]
        if det.get("rule_boost_sender"):
            scores["sender_relevance"] = max(
                0,
                min(100, int(scores["sender_relevance"]) + int(det.get("rule_boost_sender") or 0)),
            )

        if hasattr(self, "preview_meta"):
            self.preview_meta.setText(
                f"<b>{subject}</b><br>"
                f"From: {sender}<br>"
                f"When: {when} &nbsp;|&nbsp; Account: {account} &nbsp;|&nbsp; Role: {role}<br>"
                f"Actions: {actions} &nbsp;|&nbsp; Signal: {importance} &nbsp;|&nbsp; Score: {score}"
            )
        reasons = list(det.get("reasons") or [])
        for expl in det.get("rule_explanations") or []:
            reasons.append(expl)
        if hasattr(self, "preview_reasons"):
            if reasons:
                self.preview_reasons.setText(
                    "Why flagged: " + " · ".join(str(r) for r in reasons[:10])
                )
            else:
                self.preview_reasons.setText("")

        if hasattr(self, "preview_sender_score"):
            self.preview_sender_score.setText(str(scores["sender_relevance"]))
        if hasattr(self, "preview_email_score"):
            self.preview_email_score.setText(str(scores["email_relevance"]))

        if hasattr(self, "preview_vip_btn"):
            from guri_learning import is_important_sender

            if sender and is_important_sender(
                getattr(self, "learning", {}) or {}, sender
            ):
                self.preview_vip_btn.setText("Unmark VIP")
            else:
                self.preview_vip_btn.setText("Mark VIP")

        body = str(mail.get("body_preview") or "").strip()
        entry_id = str(mail.get("entry_id") or "")
        store_id = str(mail.get("store_id") or "")
        # Scrapes skip Body for Outlook responsiveness — load on demand for preview.
        if not body and entry_id:
            if hasattr(self, "email_preview"):
                self.email_preview.setPlainText("Loading message body from Outlook…")
            try:
                body = fetch_mail_body(entry_id, store_id, limit=20000)
            except Exception as exc:
                self.logger.debug("Preview body fetch failed: %s", exc)
                body = ""
            if body:
                mail["body_preview"] = body
                det["mail"] = mail
                self._preview_det = det
                # Keep the in-memory scrape cache warm for this entry
                try:
                    for it in self.scrape_cache_items or []:
                        m = it.get("mail") or {}
                        if str(m.get("entry_id") or "") == entry_id:
                            m["body_preview"] = body
                            it["mail"] = m
                            break
                except Exception:
                    pass
        if not body:
            body = "(No body text available — open in Outlook to read the full message.)"
        # Prefixed subject line helps highlight subject cues too.
        blob = f"Subject: {subject}\n\n{body}"
        self.email_preview.setHtml(highlight_action_html(blob))
        self._sync_preview_popout()

    def _open_preview_in_outlook(self) -> None:
        det = self._preview_det or {}
        mail = det.get("mail") or det
        entry_id = str(mail.get("entry_id") or "")
        store_id = str(mail.get("store_id") or "")
        if not entry_id:
            QMessageBox.information(
                self,
                "Open in Outlook",
                "Select a scraped email first (DB-only rows have no Outlook link).",
            )
            return
        if not open_mail_in_outlook(entry_id, store_id):
            QMessageBox.warning(self, "Open in Outlook", "Could not open the message in Outlook.")

    def _open_detected_in_outlook(self, item) -> None:
        det = item.data(0, Qt.ItemDataRole.UserRole) or {}
        self._show_email_preview(det if isinstance(det, dict) else {})
        mail = det.get("mail") or det
        entry_id = str(mail.get("entry_id") or "")
        store_id = str(mail.get("store_id") or "")
        if not entry_id:
            QMessageBox.information(
                self,
                "Open in Outlook",
                "This row has no Outlook EntryID (GURI DB-only item).",
            )
            return
        ok = open_mail_in_outlook(entry_id, store_id)
        if not ok:
            QMessageBox.warning(self, "Open in Outlook", "Could not open the message in Outlook.")

    def _open_selected_action_in_outlook(self) -> None:
        item = self.actions_tree.currentItem() if hasattr(self, "actions_tree") else None
        if not item:
            QMessageBox.information(self, "Open in Outlook", "Select an action item first.")
            return
        self._open_detected_in_outlook(item)

    def _open_action_item_in_outlook(self, item, _column=None) -> None:
        self._open_detected_in_outlook(item)

    def _view_email_details(self, item):
        """View full email / detection details."""
        det = item.data(0, Qt.ItemDataRole.UserRole) or {}
        mail = det.get("mail") or {}
        if not mail and det.get("sender"):
            mail = det

        dialog = QDialog(self)
        dialog.setWindowTitle("Email Details")
        dialog.resize(640, 480)
        dialog.setModal(True)

        layout = QVBoxLayout()
        dialog.setLayout(layout)

        text = QTextEdit()
        text.setReadOnly(True)
        text.setFont(QFont("Courier", 10))
        reasons = "\n".join(f"  - {r}" for r in (det.get("reasons") or []))
        actions = ", ".join(det.get("actions") or []) or "—"
        preview = str(mail.get("body_preview") or "")[:1200]
        details = f"""Email Details
{'=' * 60}

From: {mail.get('sender', item.text(0))}
To: {mail.get('to', '')}
Cc: {mail.get('cc', '')}
Account: {mail.get('account_smtp', '')}
Role: {mail.get('recipient_role', '')}
Subject: {mail.get('subject', item.text(2) if item.columnCount() > 2 else '')}
When: {mail.get('received', '')}
Unread: {mail.get('unread', '')}
Importance (Outlook): {mail.get('importance', '')}
Detected importance: {det.get('importance', '')}
Score: {det.get('score', '')}
Actions: {actions}
Status: {det.get('status', item.text(item.columnCount()-1))}
GURI: {mail.get('guri') or det.get('guri') or '—'}

Reasons:
{reasons or '  (none)'}

Preview:
{preview}
"""
        text.setPlainText(details)
        layout.addWidget(text)

        open_btn = QPushButton("Open in Outlook")
        open_btn.clicked.connect(lambda: self._open_detected_in_outlook(item))
        close_btn = QPushButton("Close")
        close_btn.clicked.connect(dialog.accept)
        add_dialog_button_row(layout, affirmative=open_btn, cancel=close_btn)

        dialog.exec()
    
    def _on_timeline_mode_toggled(self, checked: bool) -> None:
        if not checked:
            return
        self._timeline_mode = (
            "today" if self.timeline_radio_today.isChecked() else "rolling24"
        )
        self._refresh_email_timeline()

    def _on_timeline_account_filter_toggled(self, _checked: bool = False) -> None:
        self._refresh_email_timeline()

    def _timeline_account_bucket(self, mail: Optional[Dict[str, Any]]) -> str:
        """Map a mailbox SMTP/store to Aliniant / Coleago / TXO / others."""
        mail = mail or {}
        smtp = str(
            mail.get("account_smtp")
            or mail.get("account")
            or mail.get("mailbox")
            or ""
        ).strip().lower()
        # Prefer explicit Aliniant addresses; also accept the whole domain
        if smtp in {
            "julian.garrett@aliniant.com",
            "administrator@aliniant.com",
        } or smtp.endswith("@aliniant.com") or "aliniant.com" in smtp:
            return "aliniant"
        if "coleago.com" in smtp or smtp.endswith("@coleago.com"):
            return "coleago"
        if "txo.com" in smtp or smtp.endswith("@txo.com"):
            return "txo"
        return "others"

    def _timeline_account_allowed(self, mail: Optional[Dict[str, Any]]) -> bool:
        checks = getattr(self, "_timeline_account_checks", None) or {}
        if not checks:
            return True
        bucket = self._timeline_account_bucket(mail)
        cb = checks.get(bucket)
        if cb is None:
            cb = checks.get("others")
        return bool(cb.isChecked()) if cb is not None else True

    def _refresh_email_timeline(self, *, allow_db_fallback: bool = True):
        """Refresh email timeline (Today or last 24h; prefer live scrape)."""
        try:
            now = datetime.now()
            mode = getattr(self, "_timeline_mode", "rolling24")
            today = now.date()
            feedback = getattr(self, "relevance_feedback", {}) or {}
            domains = list(self.important_domains or [])
            timeline_emails = []

            # Prefer live scrape (include junked so they land in Deprecated)
            for det in self._feedback_items(drop_ignored=False):
                mail = det.get("mail") or {}
                if not self._timeline_account_allowed(mail):
                    continue
                email_time = _parse_record_datetime(mail.get("received"))
                if not email_time:
                    continue
                if mode == "today":
                    if email_time.date() != today:
                        continue
                    hour_bucket = int(email_time.hour)
                else:
                    if (now - email_time).total_seconds() > 24 * 3600:
                        continue
                    time_diff = (now - email_time).total_seconds() / 3600
                    if not (0 <= time_diff < 24):
                        continue
                    hour_bucket = int(23 - time_diff)

                if is_mail_junked(feedback, mail, det):
                    category = CATEGORY_DEPRECATED
                else:
                    category = resolve_timeline_category(
                        det,
                        getattr(self, "learning", {}) or {},
                        junked=False,
                        important_domains=domains,
                    )

                timeline_emails.append(
                    {
                        "time": email_time,
                        "sender": mail.get("sender"),
                        "subject": mail.get("subject"),
                        "datetime_str": email_time.strftime("%Y-%m-%d %H:%M:%S"),
                        "category": category,
                        "hour_bucket": hour_bucket,
                        "importance_tier": self._timeline_importance_tier(det),
                        "account_bucket": self._timeline_account_bucket(mail),
                        "det": det,
                    }
                )

            # Fallback to GURI DB only when allowed and scrape cache is empty
            if (
                not timeline_emails
                and self.db
                and allow_db_fallback
                and not self.scrape_cache_items
            ):
                all_records = self.db.get_all_records(limit=10000)
                my_email = MY_ALINIANT_EMAIL
                for record in all_records:
                    if record.get("document_type") != "01":
                        continue
                    if my_email.lower() not in str(record.get("recipients") or "").lower():
                        continue
                    email_time = _parse_record_datetime(record.get("datetime"))
                    if not email_time:
                        continue
                    if mode == "today":
                        if email_time.date() != today:
                            continue
                        hour_bucket = int(email_time.hour)
                    else:
                        if (now - email_time).total_seconds() > 24 * 3600:
                            continue
                        time_diff = (now - email_time).total_seconds() / 3600
                        if not (0 <= time_diff < 24):
                            continue
                        hour_bucket = int(23 - time_diff)

                    mail = {
                        "sender": record.get("sender"),
                        "subject": record.get("subject"),
                        "received": email_time.strftime("%Y-%m-%d %H:%M:%S"),
                        "guri": record.get("guri"),
                        "account_smtp": MY_ALINIANT_EMAIL,
                    }
                    if not self._timeline_account_allowed(mail):
                        continue
                    det = {
                        "mail": mail,
                        "actions": [],
                        "score": 20
                        if self._mail_is_important_domain(mail)
                        else 0,
                        "importance": "medium"
                        if self._mail_is_important_domain(mail)
                        else "low",
                    }
                    category = resolve_timeline_category(
                        det,
                        getattr(self, "learning", {}) or {},
                        junked=is_mail_junked(feedback, mail),
                        important_domains=domains,
                    )
                    timeline_emails.append(
                        {
                            "time": email_time,
                            "sender": record.get("sender"),
                            "subject": record.get("subject"),
                            "datetime_str": email_time.strftime("%Y-%m-%d %H:%M:%S"),
                            "category": category,
                            "hour_bucket": hour_bucket,
                            "importance_tier": self._timeline_importance_tier(det),
                            "account_bucket": self._timeline_account_bucket(mail),
                            "det": det,
                        }
                    )

            for email in timeline_emails:
                cat = email.get("category") or CATEGORY_LESS
                if cat not in TIMELINE_CATEGORIES:
                    email["category"] = CATEGORY_LESS

            self._timeline_emails = timeline_emails

            visible_hours = (now.hour + 1) if mode == "today" else 24
            if hasattr(self, "timeline_canvas"):
                self.timeline_canvas.set_timeline_mode(mode, visible_hours)

            title = (
                "Email Timeline — Today"
                if mode == "today"
                else "Email Timeline — Last 24 Hours"
            )
            if hasattr(self, "timeline_frame"):
                self.timeline_frame.setTitle(title)

            totals = {c: 0 for c in TIMELINE_CATEGORIES}
            for email in timeline_emails:
                totals[email.get("category") or CATEGORY_LESS] = (
                    totals.get(email.get("category") or CATEGORY_LESS, 0) + 1
                )
            self.timeline_count_label.setText(
                f"Emails: {len(timeline_emails)}  ·  "
                f"Important {totals.get(CATEGORY_IMPORTANT, 0)}  ·  "
                f"Actionable {totals[CATEGORY_ACTIONABLE]}  ·  "
                f"Relevant {totals[CATEGORY_RELEVANT]}  ·  "
                f"Less {totals[CATEGORY_LESS]}  ·  "
                f"Deprecated {totals[CATEGORY_DEPRECATED]}"
            )
            self._draw_timeline_grid(timeline_emails, now)
            scope = "today" if mode == "today" else "last 24 hours"
            self.status_bar.showMessage(
                f"Timeline refreshed: {len(timeline_emails)} emails {scope} "
                f"(I:{totals.get(CATEGORY_IMPORTANT, 0)} "
                f"A:{totals[CATEGORY_ACTIONABLE]} R:{totals[CATEGORY_RELEVANT]} "
                f"L:{totals[CATEGORY_LESS]} D:{totals[CATEGORY_DEPRECATED]})"
            )
        except Exception as e:
            self.logger.error(f"Error refreshing email timeline: {e}")
            if hasattr(self, "status_bar"):
                self.status_bar.showMessage(f"Timeline refresh failed: {e}")

    def _draw_timeline_grid(self, emails, current_time):
        """Draw the hour grid with email indicators in category rows."""
        hour_counts: Dict[int, Dict[str, Any]] = {}
        mode = getattr(self, "_timeline_mode", "rolling24")
        rank = TimelineCanvas.IMPORTANCE_RANK
        for email in emails:
            hour_bucket = email.get("hour_bucket")
            if hour_bucket is None:
                email_time = email.get("time")
                if not email_time:
                    continue
                if mode == "today":
                    hour_bucket = int(email_time.hour)
                else:
                    time_diff = (current_time - email_time).total_seconds() / 3600
                    if not (0 <= time_diff < 24):
                        continue
                    hour_bucket = int(23 - time_diff)
            category = email.get("category") or CATEGORY_LESS
            if category not in TIMELINE_CATEGORIES:
                category = CATEGORY_LESS
            tier = str(email.get("importance_tier") or "normal")
            if tier not in rank:
                tier = self._timeline_importance_tier(email.get("det") or {})
            bucket = hour_counts.setdefault(hour_bucket, {})
            cell = bucket.get(category)
            if not isinstance(cell, dict):
                cell = {"count": 0, "importance": "normal"}
                bucket[category] = cell
            cell["count"] = int(cell.get("count") or 0) + 1
            prev = str(cell.get("importance") or "normal")
            if rank.get(tier, 0) > rank.get(prev, 0):
                cell["importance"] = tier

        self.timeline_canvas.update_email_data(hour_counts)

    def _timeline_matches(self, hour_bucket: int, category: str) -> List[Dict[str, Any]]:
        return [
            e
            for e in (self._timeline_emails or [])
            if e.get("hour_bucket") == hour_bucket
            and (e.get("category") or CATEGORY_LESS) == category
        ]

    def _show_timeline_popup(
        self, hour_bucket: int, category: str, anchor: QPoint, *, pinned: bool = False
    ) -> None:
        """Show sender dropdown under a timeline circle."""
        if getattr(self, "_timeline_popup_busy", False):
            return
        self._timeline_popup_busy = True
        try:
            matches = self._timeline_matches(hour_bucket, category)
            popup = self._timeline_popup
            if not matches:
                popup.unpin()
                popup.hide()
                self._timeline_hover_key = None
                return
            key = (hour_bucket, category)
            # Same cell already open — just keep it (avoid hide/show churn)
            if popup.isVisible() and self._timeline_hover_key == key and not pinned:
                popup.cancel_hide()
                return

            label = TIMELINE_CATEGORY_LABELS.get(category, category)
            if getattr(self, "_timeline_mode", "rolling24") == "today":
                hour_label = f"{int(hour_bucket):02d}:00"
            else:
                hour_label = (
                    datetime.now() - timedelta(hours=23 - int(hour_bucket))
                ).strftime("%H:00")
            title = (
                f"{label} · ~{hour_label} · {len(matches)} email(s) — click to preview"
            )
            popup.cancel_hide()
            popup.populate(title, matches)
            self._timeline_hover_key = key

            popup.adjustSize()
            # Position in parent (timeline group) coordinates from a global anchor
            parent = popup.parentWidget()
            if parent is not None:
                local = parent.mapFromGlobal(anchor)
                x = int(local.x() - popup.width() / 2)
                y = int(local.y() + 4)
                # Keep inside parent widget as much as possible
                x = max(4, min(x, max(4, parent.width() - popup.width() - 4)))
                if y + popup.height() > parent.height() - 4:
                    y = int(local.y() - popup.height() - 8)
                y = max(4, y)
                popup.move(x, y)
            else:
                popup.move(int(anchor.x() - popup.width() / 2), int(anchor.y() + 4))

            if pinned:
                popup.pin()
            else:
                popup.unpin()
            popup.show()
            popup.raise_()
        finally:
            self._timeline_popup_busy = False

    def _on_timeline_cell_hovered(
        self, hour_bucket: int, category: str, anchor
    ) -> None:
        try:
            # While a click-pinned list is open, ignore hover switches (prevents
            # thrashing / freezes when moving toward another circle).
            popup = getattr(self, "_timeline_popup", None)
            if popup is not None and popup.isVisible() and getattr(popup, "_pinned", False):
                return
            if not isinstance(anchor, QPoint):
                anchor = QPoint(int(anchor.x()), int(anchor.y()))
            self._show_timeline_popup(hour_bucket, category, anchor, pinned=False)
        except Exception as e:
            self.logger.error("Timeline hover popup failed: %s", e)

    def _on_timeline_cell_left(self) -> None:
        if hasattr(self, "_timeline_popup"):
            self._timeline_popup.schedule_hide()

    def _on_timeline_cell_clicked(
        self, hour_bucket: int, category: str, anchor
    ) -> None:
        """Click pins the dropdown for that circle (replaces any previous pin)."""
        try:
            popup = self._timeline_popup
            # Sentinel from empty-canvas click
            if hour_bucket < 0 or not category:
                popup.unpin()
                popup.hide()
                self._timeline_hover_key = None
                return
            if not isinstance(anchor, QPoint):
                anchor = QPoint(int(anchor.x()), int(anchor.y()))
            key = (hour_bucket, category)
            # Toggle closed if the same pinned circle is clicked again
            if (
                popup.isVisible()
                and getattr(popup, "_pinned", False)
                and self._timeline_hover_key == key
            ):
                popup.unpin()
                popup.hide()
                self._timeline_hover_key = None
                self.status_bar.showMessage("Timeline list closed")
                return
            self._show_timeline_popup(hour_bucket, category, anchor, pinned=True)
            matches = self._timeline_matches(hour_bucket, category)
            label = TIMELINE_CATEGORY_LABELS.get(category, category)
            self.status_bar.showMessage(
                f"Timeline: {len(matches)} {label.lower()} — click a sender to preview"
            )
        except Exception as e:
            self.logger.error("Timeline cell click failed: %s", e)
            self.status_bar.showMessage(f"Timeline click failed: {e}")

    def _on_timeline_email_chosen(self, det) -> None:
        """Load chosen timeline email into the Welcome preview panel."""
        if not isinstance(det, dict):
            return
        self._show_email_preview(det)
        mail = det.get("mail") or {}
        sender = str(mail.get("sender") or "")[:60]
        self.status_bar.showMessage(f"Preview: {sender or 'selected email'}")
        # Ensure Welcome tab is visible so the right-hand panel is on screen
        try:
            if hasattr(self, "notebook"):
                for i in range(self.notebook.count()):
                    if self.notebook.tabText(i) == "Welcome":
                        self.notebook.setCurrentIndex(i)
                        break
        except Exception:
            pass

    def _on_timeline_wrong_category(self, entry) -> None:
        """Right-click on a timeline list item → wrong categorisation dialog."""
        if not isinstance(entry, dict):
            return
        # Keep the dropdown open while the dialog is up
        popup = getattr(self, "_timeline_popup", None)
        if popup is not None:
            popup.pin()
            popup.cancel_hide()
        self._show_wrong_categorisation_dialog(entry)

    def _show_wrong_categorisation_dialog(self, entry: Dict[str, Any]) -> None:
        det = entry.get("det") if isinstance(entry.get("det"), dict) else None
        if not det:
            det = {
                "mail": {
                    "sender": entry.get("sender"),
                    "subject": entry.get("subject"),
                    "received": entry.get("datetime_str"),
                },
                "actions": [],
                "reasons": [],
                "score": 0,
            }
        mail = det.get("mail") or {}
        sender = str(mail.get("sender") or entry.get("sender") or "")
        subject = str(mail.get("subject") or entry.get("subject") or "")
        from_category = str(entry.get("category") or CATEGORY_ACTIONABLE)
        from_label = TIMELINE_CATEGORY_LABELS.get(from_category, from_category)

        dialog = QDialog(self)
        dialog.setWindowTitle("Wrong categorisation")
        dialog.setModal(True)
        dialog.resize(440, 420)
        layout = QVBoxLayout(dialog)

        intro = QLabel(
            f"<b>Current:</b> {from_label}<br>"
            f"<b>From:</b> {sender}<br>"
            f"<b>Subject:</b> {subject[:120]}"
        )
        intro.setWordWrap(True)
        intro.setTextFormat(Qt.TextFormat.RichText)
        layout.addWidget(intro)

        override_box = QGroupBox("Instant override (skips gradual nudge)")
        override_layout = QVBoxLayout(override_box)
        force_important_email_cb = QCheckBox("This is an important email")
        force_important_email_cb.setToolTip(
            "Jump this message straight to Important — no one-step ladder climb."
        )
        force_important_sender_cb = QCheckBox("This is an important sender")
        force_important_sender_cb.setToolTip(
            "Mark the sender as VIP and jump this message to Important immediately."
        )
        override_font = QFont("Segoe UI", 9)
        override_font.setBold(True)
        force_important_email_cb.setFont(override_font)
        force_important_sender_cb.setFont(override_font)
        override_layout.addWidget(force_important_email_cb)
        override_layout.addWidget(force_important_sender_cb)
        layout.addWidget(override_box)

        why_box = QGroupBox("Why was this wrongly categorised?")
        why_layout = QVBoxLayout(why_box)
        reason_checks: Dict[str, QCheckBox] = {}
        for group_title, reason_rows in CATEGORY_CORRECTION_REASON_GROUPS:
            group_lbl = QLabel(group_title)
            group_lbl.setStyleSheet(
                f"color: {PALETTE['muted']}; font-size: 11px; font-weight: 600;"
            )
            why_layout.addWidget(group_lbl)
            for key, label in reason_rows:
                cb = QCheckBox(label)
                reason_checks[key] = cb
                why_layout.addWidget(cb)
        layout.addWidget(why_box)

        # If user ticks a positive reason, nudge the default radio toward its prefer
        def _nudge_target_from_reasons() -> None:
            if force_important_email_cb.isChecked() or force_important_sender_cb.isChecked():
                rb = target_radios.get(CATEGORY_IMPORTANT)
                if rb is not None:
                    rb.setChecked(True)
                return
            prefers: List[str] = []
            for key, cb in reason_checks.items():
                if not cb.isChecked():
                    continue
                hint = POSITIVE_REASON_LEARNING.get(key) or {}
                prefer = str(hint.get("prefer") or "").strip().lower()
                if prefer in target_radios:
                    prefers.append(prefer)
            if not prefers:
                return
            # Strongest prefer wins: important > actionable > relevant > less
            rank = {
                "important": 4,
                "actionable": 3,
                "relevant": 2,
                "less": 1,
                "deprecated": 0,
            }
            best = max(prefers, key=lambda p: rank.get(p, 0))
            rb = target_radios.get(best)
            if rb is not None:
                rb.setChecked(True)

        def _on_instant_override_toggled(_checked: bool = False) -> None:
            if force_important_email_cb.isChecked() or force_important_sender_cb.isChecked():
                rb = target_radios.get(CATEGORY_IMPORTANT)
                if rb is not None:
                    rb.setChecked(True)

        force_important_email_cb.toggled.connect(_on_instant_override_toggled)
        force_important_sender_cb.toggled.connect(_on_instant_override_toggled)

        recat_box = QGroupBox("Re-categorise to")
        recat_layout = QVBoxLayout(recat_box)
        target_group = QButtonGroup(dialog)
        target_radios: Dict[str, QRadioButton] = {}
        # Default: one rung toward a more useful bucket (or one rung down from top)
        default_target = step_category_toward(from_category, CATEGORY_ACTIONABLE)
        if from_category == CATEGORY_IMPORTANT:
            default_target = CATEGORY_ACTIONABLE
        elif from_category == CATEGORY_ACTIONABLE:
            default_target = CATEGORY_RELEVANT
        elif from_category == CATEGORY_DEPRECATED:
            default_target = CATEGORY_LESS
        for key, label in CATEGORY_RECAT_TARGETS:
            rb = QRadioButton(label)
            target_radios[key] = rb
            target_group.addButton(rb)
            recat_layout.addWidget(rb)
            if key == default_target:
                rb.setChecked(True)
        if not any(rb.isChecked() for rb in target_radios.values()):
            target_radios[CATEGORY_RELEVANT].setChecked(True)
        layout.addWidget(recat_box)

        for cb in reason_checks.values():
            cb.toggled.connect(lambda _checked: _nudge_target_from_reasons())

        hint = QLabel(
            "Use Instant override when you’re sure — jumps straight to Important "
            "(and VIP the sender if chosen). Otherwise reasons nudge one step "
            "(Less → Relevant → Actionable → Important); Actionable intent will "
            "not be parked in “Relevant but not actionable”."
        )
        hint.setWordWrap(True)
        hint.setStyleSheet(f"color: {PALETTE['muted']}; font-size: 11px;")
        layout.addWidget(hint)

        save_btn = QPushButton("Save")
        save_btn.setDefault(True)
        cancel_btn = QPushButton("Cancel")
        cancel_btn.clicked.connect(dialog.reject)
        add_dialog_button_row(layout, affirmative=save_btn, cancel=cancel_btn)

        def on_save() -> None:
            reasons = [k for k, cb in reason_checks.items() if cb.isChecked()]
            instant_email = force_important_email_cb.isChecked()
            instant_sender = force_important_sender_cb.isChecked()
            to_category = ""
            for key, rb in target_radios.items():
                if rb.isChecked():
                    to_category = key
                    break
            if instant_email or instant_sender:
                to_category = CATEGORY_IMPORTANT
            if not to_category:
                QMessageBox.warning(
                    dialog, "Wrong categorisation", "Choose a re-categorise target."
                )
                return
            if not reasons and not instant_email and not instant_sender:
                ask = QMessageBox.question(
                    dialog,
                    "No reason selected",
                    "No why-checkboxes were selected. Save the re-categorisation anyway?",
                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                    QMessageBox.StandardButton.Yes,
                )
                if ask != QMessageBox.StandardButton.Yes:
                    return
            self._apply_category_correction(
                det=det,
                entry=entry,
                from_category=from_category,
                to_category=to_category,
                reasons=reasons,
                instant_important_email=instant_email,
                instant_important_sender=instant_sender,
            )
            dialog.accept()

        save_btn.clicked.connect(on_save)
        dialog.resize(460, 520)
        dialog.exec()

    def _apply_category_correction(
        self,
        *,
        det: Dict[str, Any],
        entry: Dict[str, Any],
        from_category: str,
        to_category: str,
        reasons: List[str],
        instant_important_email: bool = False,
        instant_important_sender: bool = False,
    ) -> None:
        mail = det.get("mail") or {}
        sender = str(mail.get("sender") or entry.get("sender") or "")
        key = mail_feedback_key(mail, det)
        actions = list(det.get("actions") or [])
        instant = bool(instant_important_email or instant_important_sender)
        intended = CATEGORY_IMPORTANT if instant else to_category
        if instant:
            applied = CATEGORY_IMPORTANT
        else:
            applied = step_category_toward(from_category, intended, reasons=reasons)

        if instant_important_sender and sender:
            self.learning = remember_important_sender(
                getattr(self, "learning", {}) or {}, sender
            )

        self.learning = record_category_correction(
            getattr(self, "learning", {}) or {},
            mail_key=key,
            sender=sender,
            from_category=from_category,
            to_category=intended,
            reasons=reasons,
            actions=actions,
            instant=instant,
        )
        # Prefer the category stored on the override (instant or stepped)
        ov = (self.learning.get("category_overrides") or {}).get(key) or {}
        applied = str(ov.get("category") or applied)
        if instant:
            applied = CATEGORY_IMPORTANT
            if key and isinstance(self.learning.get("category_overrides"), dict):
                ov = dict(self.learning["category_overrides"].get(key) or {})
                ov["category"] = CATEGORY_IMPORTANT
                ov["intended_category"] = CATEGORY_IMPORTANT
                self.learning["category_overrides"][key] = ov
        self._save_learning()
        if instant_important_sender and hasattr(self, "_refresh_vip_list"):
            try:
                self._refresh_vip_list()
            except Exception:
                pass

        # Soft, one-step relevance nudges (no full knee-jerk pin)
        from guri_learning import category_ladder_index

        moving_up = category_ladder_index(applied) > category_ladder_index(from_category)
        moving_down = category_ladder_index(applied) < category_ladder_index(from_category)

        if moving_down:
            penalty = 8 if applied != CATEGORY_DEPRECATED else 14
            self.relevance_feedback = deprecate_email(
                self.relevance_feedback, key, penalty=penalty
            )
            if "not_important" in reasons or "marketing" in reasons:
                self.relevance_feedback = derate_sender(
                    self.relevance_feedback, sender, penalty=max(6, penalty - 2)
                )
            if (
                any(
                    r in {"not_actionable", "no_action_needed", "wrong_deadline"}
                    for r in reasons
                )
                and actions
            ):
                self.relevance_feedback = derate_content_types(
                    self.relevance_feedback, actions, penalty=8
                )
        elif moving_up:
            # Gentle boost toward the next rung
            bump = (
                12
                if applied == CATEGORY_IMPORTANT
                else 10
                if applied == CATEGORY_ACTIONABLE
                else 8
            )
            self.relevance_feedback = endorse_email_score(
                self.relevance_feedback, key, step=bump
            )
            if applied in {CATEGORY_ACTIONABLE, CATEGORY_IMPORTANT}:
                from guri_action_store import endorse_sender_score

                self.relevance_feedback = endorse_sender_score(
                    self.relevance_feedback, sender, step=6
                )

        # Positive why-reasons always feed relevance learning (even if already on rung)
        positive_hit = [r for r in reasons if r in POSITIVE_REASON_LEARNING]
        if positive_hit:
            from guri_action_store import endorse_sender_score

            extra = 6 + (2 * len(positive_hit))
            self.relevance_feedback = endorse_email_score(
                self.relevance_feedback, key, step=min(14, extra)
            )
            if any(
                (POSITIVE_REASON_LEARNING.get(r) or {}).get("important_sender")
                for r in positive_hit
            ):
                self.relevance_feedback = endorse_sender_score(
                    self.relevance_feedback, sender, step=8
                )
            elif any(
                (POSITIVE_REASON_LEARNING.get(r) or {}).get("friend_sender")
                for r in positive_hit
            ):
                # Friends: mild endorse, not VIP hammer
                self.relevance_feedback = endorse_sender_score(
                    self.relevance_feedback, sender, step=4
                )

        if "wrong_deadline" in reasons and key and moving_down:
            self.learning = reject_deadline_mail(self.learning, key)
            self._save_learning()

        try:
            save_relevance_feedback(self.relevance_feedback)
        except Exception as exc:
            self.logger.warning("Failed saving relevance after recategorise: %s", exc)

        # Hide popup and refresh timeline so the circle moves immediately
        popup = getattr(self, "_timeline_popup", None)
        if popup is not None:
            popup.unpin()
            popup.hide()
        self._timeline_hover_key = None
        self._refresh_email_timeline(allow_db_fallback=False)
        if hasattr(self, "_refresh_important_emails"):
            self._refresh_important_emails(allow_db_fallback=False, refresh_actions=True)
        if hasattr(self, "_refresh_actions_panel"):
            self._refresh_actions_panel()

        applied_label = TIMELINE_CATEGORY_LABELS.get(applied, applied)
        intended_label = TIMELINE_CATEGORY_LABELS.get(intended, intended)
        if instant:
            vip_note = " (sender marked VIP)" if instant_important_sender else ""
            self.status_bar.showMessage(
                f"Instant Important{vip_note} — override applied"
            )
        elif applied == intended:
            self.status_bar.showMessage(
                f"Recategorised to {applied_label} — GURI will remember this"
            )
        else:
            self.status_bar.showMessage(
                f"Moved toward {intended_label} → now {applied_label} "
                f"(gradual learning — correct again to climb further)"
            )

    def _touch_last_refresh_label(self) -> None:
        """Update the Welcome 'Last refresh' stamp (proof auto-refresh is alive)."""
        if hasattr(self, "last_refresh_label"):
            self.last_refresh_label.setText(
                datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            )

    def _start_auto_refresh(self):
        """Start Welcome UI refresh (60s) and lighter Outlook scrape (5 min)."""
        self._auto_scrape_enabled = True
        self.auto_refresh_timer.start(60_000)
        if hasattr(self, "scrape_timer"):
            interval = int(getattr(self, "_auto_scrape_interval_ms", 5 * 60 * 1000))
            self.scrape_timer.start(interval)
        # Lightweight first paint refresh — no 10k DB fallbacks
        self._touch_last_refresh_label()
        self._schedule_welcome_refresh(allow_db_fallback=False)
        # First Outlook scrape shortly after startup (then via scrape_timer)
        QTimer.singleShot(8000, lambda: self._start_outlook_scrape(manual=False))
    
    def _auto_refresh_welcome(self):
        """Auto-refresh Welcome panels every 60 seconds (no Outlook COM)."""
        try:
            self._touch_last_refresh_label()
            if getattr(self, "_expanding_to_screen", False):
                return
            # Cheap stats even when a full panel rebuild is already in flight.
            self._refresh_welcome_stats_from_scrape_cache()
            if not getattr(self, "_welcome_refresh_busy", False):
                self._schedule_welcome_refresh(allow_db_fallback=False)
        except Exception as e:
            self.logger.error(f"Error in auto-refresh: {e}")

    def _auto_scrape_outlook(self):
        """Periodic light Outlook scrape (default every 5 minutes)."""
        try:
            if not getattr(self, "_auto_scrape_enabled", True):
                return
            if self._scrape_busy or getattr(self, "_expanding_to_screen", False):
                return
            self._start_outlook_scrape(manual=False)
        except Exception as e:
            self.logger.error(f"Error in scrape auto-refresh: {e}")
    
    def _stop_auto_refresh(self):
        """Stop the auto-refresh timer."""
        if self.auto_refresh_timer:
            self.auto_refresh_timer.stop()
            self.logger.info("Auto-refresh stopped")
        if hasattr(self, "scrape_timer") and self.scrape_timer:
            self.scrape_timer.stop()
        self._auto_scrape_enabled = False
    
    @staticmethod
    def _pixmap_knockout_background(pixmap: QPixmap, threshold: int = 248) -> QPixmap:
        """Make near-white JPEG backgrounds transparent for tray icons."""
        if pixmap.isNull():
            return pixmap
        image = pixmap.toImage().convertToFormat(QImage.Format.Format_ARGB32)
        for y in range(image.height()):
            for x in range(image.width()):
                c = image.pixelColor(x, y)
                # Near-white / light gray plate → fully transparent
                if c.red() >= threshold and c.green() >= threshold and c.blue() >= threshold:
                    c.setAlpha(0)
                    image.setPixelColor(x, y, c)
                # Soft edge: fade remaining pale pixels so the cutout isn't jagged
                elif min(c.red(), c.green(), c.blue()) >= 220 and (max(c.red(), c.green(), c.blue()) - min(c.red(), c.green(), c.blue())) < 25:
                    fade = (threshold - min(c.red(), c.green(), c.blue())) / max(1, threshold - 220)
                    c.setAlpha(max(0, min(255, int(255 * fade))))
                    image.setPixelColor(x, y, c)
        return QPixmap.fromImage(image)

    def _apply_app_icon(self) -> None:
        """Set taskbar / window icon from abyitself.ico (with SVG logo fallback)."""
        icon = QIcon()
        if os.path.isfile(GURI_APP_ICON_PATH):
            icon = QIcon(GURI_APP_ICON_PATH)
        if icon.isNull():
            logo = _load_guri_logo_pixmap(32)
            if logo is not None and not logo.isNull():
                icon = QIcon(logo)
        if not icon.isNull():
            self.setWindowIcon(icon)

    def _create_system_tray(self):
        """Create system tray icon and menu."""
        if not self.tray_available:
            return
        
        # Create system tray icon
        self.tray_icon = QSystemTrayIcon(self)
        
        # Prefer the official taskbar ICO; fall back to PNG/JPEG assets
        icon_candidates = (
            GURI_APP_ICON_PATH,
            r"C:\GeoFooter\A_by_istself.png",
            r"C:\GeoFooter\A_by_istself.jpg",
        )
        icon_path = next((p for p in icon_candidates if os.path.exists(p)), None)
        if icon_path:
            try:
                if icon_path.lower().endswith(".ico"):
                    icon = QIcon(icon_path)
                    if not icon.isNull():
                        self.tray_icon.setIcon(icon)
                        self.logger.info("Loaded system tray icon from: %s", icon_path)
                    else:
                        self._create_fallback_tray_icon()
                else:
                    pixmap = QPixmap(icon_path)
                    if not pixmap.isNull():
                        if not pixmap.hasAlphaChannel() or icon_path.lower().endswith(
                            (".jpg", ".jpeg")
                        ):
                            pixmap = self._pixmap_knockout_background(pixmap)
                        icon = QIcon()
                        for size in (16, 22, 32):
                            icon.addPixmap(
                                pixmap.scaled(
                                    size,
                                    size,
                                    Qt.AspectRatioMode.KeepAspectRatio,
                                    Qt.TransformationMode.SmoothTransformation,
                                ),
                                QIcon.Mode.Normal,
                            )
                        self.tray_icon.setIcon(icon)
                        self.logger.info("Loaded system tray icon from: %s", icon_path)
                    else:
                        self.logger.error(
                            "Failed to load image from %s: Invalid image format",
                            icon_path,
                        )
                        self._create_fallback_tray_icon()
            except Exception as e:
                self.logger.error("Error loading tray icon from %s: %s", icon_path, e)
                self._create_fallback_tray_icon()
        else:
            self.logger.warning("Tray icon file not found. Using fallback icon.")
            self._create_fallback_tray_icon()
        
        # Tray preferences (Outlook-style checkable options)
        self.tray_show_notifications = True
        self.tray_hide_when_minimized = True

        # Create tray menu — must NOT be parented to the main window.
        # When GURI is hidden to the tray, a menu owned by a hidden QMainWindow
        # often fails to open on Windows right-click.
        tray_menu = QMenu()
        tray_menu.setObjectName("trayMenu")
        tray_font = QFont("Segoe UI", 9)
        tray_menu.setFont(tray_font)
        self.tray_menu = tray_menu

        # Section 1 — notification toggles (checkable)
        notify_action = QAction("Show Desktop Notifications", tray_menu)
        notify_action.setCheckable(True)
        notify_action.setChecked(True)
        notify_action.setFont(tray_font)
        notify_action.toggled.connect(self._tray_set_notifications)
        tray_menu.addAction(notify_action)
        self.tray_notify_action = notify_action

        # Section 2
        tray_menu.addSeparator()

        auto_scrape_action = QAction("Auto-scrape Outlook", tray_menu)
        auto_scrape_action.setCheckable(True)
        auto_scrape_action.setChecked(True)
        auto_scrape_action.setFont(tray_font)
        auto_scrape_action.toggled.connect(self._tray_set_auto_scrape)
        tray_menu.addAction(auto_scrape_action)
        self.tray_auto_scrape_action = auto_scrape_action

        # Section 3 — window behaviour + bold default open
        tray_menu.addSeparator()

        hide_min_action = QAction("Hide When Minimized", tray_menu)
        hide_min_action.setCheckable(True)
        hide_min_action.setChecked(True)
        hide_min_action.setFont(tray_font)
        hide_min_action.toggled.connect(self._tray_set_hide_when_minimized)
        tray_menu.addAction(hide_min_action)
        self.tray_hide_min_action = hide_min_action

        open_action = QAction("Open GURI", tray_menu)
        open_font = QFont("Segoe UI", 9)
        open_font.setBold(True)
        open_action.setFont(open_font)
        open_action.triggered.connect(self._tray_open_guri)
        tray_menu.addAction(open_action)
        self.tray_open_action = open_action

        # Section 4 — actions / quit
        tray_menu.addSeparator()

        refresh_action = QAction("Refresh Records", tray_menu)
        refresh_action.setFont(tray_font)
        refresh_action.triggered.connect(
            lambda: QTimer.singleShot(0, self._load_records)
        )
        tray_menu.addAction(refresh_action)

        scrape_action = QAction("Scrape Outlook Now", tray_menu)
        scrape_action.setFont(tray_font)
        scrape_action.triggered.connect(
            lambda: QTimer.singleShot(0, lambda: self._start_outlook_scrape(manual=True))
        )
        tray_menu.addAction(scrape_action)

        tray_menu.addSeparator()

        quit_action = QAction("Quit", tray_menu)
        quit_action.setFont(tray_font)
        quit_action.triggered.connect(
            lambda: QTimer.singleShot(0, self._quit_application)
        )
        tray_menu.addAction(quit_action)

        # Do NOT also call setContextMenu — on Windows that double-fires with our
        # Context handler and makes the menu sluggish or appear then vanish.
        self.tray_icon.activated.connect(self._tray_icon_activated)
        self.tray_icon.setToolTip("GURI Database Viewer & Manager")
        self.tray_icon.show()

        if self.tray_show_notifications:
            self.tray_icon.showMessage(
                "GURI Database Manager",
                "Application is running in the system tray",
                QSystemTrayIcon.MessageIcon.Information,
                2000,
            )

    def _show_tray_menu(self) -> None:
        """Show the tray context menu at the cursor (Windows-reliable path)."""
        menu = getattr(self, "tray_menu", None)
        if menu is None:
            return
        # Always use cursor position. QSystemTrayIcon.geometry() is often wrong
        # on modern Windows (overflow chevron / multi-monitor).
        try:
            if menu.isVisible():
                menu.hide()
            menu.popup(QCursor.pos())
        except Exception as exc:
            self.logger.warning("Tray menu popup failed: %s", exc)

    def _create_fallback_tray_icon(self):
        """Create a fallback tray icon if the image file is not available."""
        icon = QIcon()
        pixmap = QPixmap(16, 16)
        pixmap.fill(Qt.GlobalColor.transparent)
        painter = QPainter(pixmap)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing, True)
        painter.setPen(QColor("#FF7E2E"))
        painter.setFont(QFont("Segoe UI", 10, QFont.Weight.Bold))
        painter.drawText(pixmap.rect(), Qt.AlignmentFlag.AlignCenter, "A")
        painter.end()
        icon.addPixmap(pixmap)
        if self.tray_icon is not None:
            self.tray_icon.setIcon(icon)

    def _tray_set_notifications(self, checked: bool) -> None:
        self.tray_show_notifications = bool(checked)

    def _tray_set_hide_when_minimized(self, checked: bool) -> None:
        self.tray_hide_when_minimized = bool(checked)

    def _tray_set_auto_scrape(self, checked: bool) -> None:
        self._auto_scrape_enabled = bool(checked)
        if checked:
            if hasattr(self, "auto_refresh_timer") and not self.auto_refresh_timer.isActive():
                self.auto_refresh_timer.start(60_000)
            if hasattr(self, "scrape_timer"):
                interval = int(getattr(self, "_auto_scrape_interval_ms", 5 * 60 * 1000))
                self.scrape_timer.start(interval)
            QTimer.singleShot(0, lambda: self._start_outlook_scrape(manual=False))
        else:
            if hasattr(self, "scrape_timer"):
                self.scrape_timer.stop()
        # When unchecked, auto-refresh still redraws panels; only live Outlook scrape pauses

    def _tray_notify(self, title: str, message: str, msec: int = 2000) -> None:
        """Show a tray balloon only when notifications are enabled."""
        if (
            getattr(self, "tray_show_notifications", True)
            and self.tray_icon
            and self.tray_icon.isVisible()
        ):
            self.tray_icon.showMessage(
                title,
                message,
                QSystemTrayIcon.MessageIcon.Information,
                msec,
            )

    def _tray_open_guri(self) -> None:
        """Open/maximize from the tray after the menu has fully dismissed."""
        # Longer defer than a single event-loop tick — Windows still holds
        # foreground ownership for a beat after the tray menu closes.
        QTimer.singleShot(150, self.expand_to_screen)

    def expand_to_screen(self) -> None:
        """Restore and maximize the main window (tray / launch)."""
        if getattr(self, "_expanding_to_screen", False):
            return
        self._expanding_to_screen = True
        try:
            # Avoid setGeometry(full screen) before maximize — that double-layouts
            # the heavy Welcome UI and feels like the tray action hung.
            self.setWindowState(
                (self.windowState() & ~Qt.WindowState.WindowMinimized)
                | Qt.WindowState.WindowMaximized
            )
            self.show()
            self.showMaximized()
            self.raise_()
            self.activateWindow()
            # Soft foreground nudge only — no AttachThreadInput (can deadlock).
            QTimer.singleShot(0, self._force_window_foreground)
            # AES ribbon / second launch: refresh Welcome so it does not look stuck.
            QTimer.singleShot(200, self._request_ui_refresh_from_raise)
        finally:
            QTimer.singleShot(
                300, lambda: setattr(self, "_expanding_to_screen", False)
            )

    def _force_window_foreground(self) -> None:
        """Bring the main window to the front (safe Windows tray restore)."""
        try:
            self.raise_()
            self.activateWindow()
            if sys.platform != "win32":
                return
            import ctypes
            from ctypes import wintypes

            user32 = ctypes.windll.user32
            hwnd = int(self.winId())
            if not hwnd:
                return

            SW_RESTORE = 9
            SW_SHOW = 5
            user32.ShowWindow(hwnd, SW_RESTORE if user32.IsIconic(hwnd) else SW_SHOW)
            user32.BringWindowToTop(hwnd)
            if user32.SetForegroundWindow(hwnd):
                return

            # Fallback when another app owns foreground (e.g. Outlook after AES click):
            # brief taskbar flash so the user notices GURI if focus steal is blocked.
            class FLASHWINFO(ctypes.Structure):
                _fields_ = [
                    ("cbSize", wintypes.UINT),
                    ("hwnd", wintypes.HWND),
                    ("dwFlags", wintypes.DWORD),
                    ("uCount", wintypes.UINT),
                    ("dwTimeout", wintypes.DWORD),
                ]

            FLASHW_TRAY = 0x00000002
            FLASHW_TIMERNOFG = 0x0000000C
            info = FLASHWINFO(
                cbSize=ctypes.sizeof(FLASHWINFO),
                hwnd=hwnd,
                dwFlags=FLASHW_TRAY | FLASHW_TIMERNOFG,
                uCount=3,
                dwTimeout=0,
            )
            user32.FlashWindowEx(ctypes.byref(info))
        except Exception:
            pass

    def _tray_icon_activated(self, reason):
        """Handle system tray icon activation."""
        if reason == QSystemTrayIcon.ActivationReason.Context:
            self._show_tray_menu()
            return
        if reason == QSystemTrayIcon.ActivationReason.DoubleClick:
            self._tray_open_guri()
            return
        if reason == QSystemTrayIcon.ActivationReason.Trigger:
            # Left-click restores when hidden. Do not toggle-hide here — on
            # Windows a right-click can also emit Trigger and would hide the
            # window while the context menu opens.
            if (not self.isVisible()) or bool(
                self.windowState() & Qt.WindowState.WindowMinimized
            ):
                self._tray_open_guri()
    
    def _quit_application(self):
        """Quit the application completely."""
        self._stop_auto_refresh()
        if self.tray_icon:
            self.tray_icon.hide()
        QApplication.quit()
    
    def closeEvent(self, event):
        """Handle application closing — hide to tray when that option is on."""
        if (
            self.tray_available
            and self.tray_icon
            and getattr(self, "tray_hide_when_minimized", True)
        ):
            self.hide()
            self._tray_notify(
                "GURI Database Manager",
                "Application minimized to system tray",
            )
            event.ignore()
        else:
            self._stop_auto_refresh()
            event.accept()


def main():
    """Main function to run the GUI application.

    AES / second launches pass ``--raise`` (or any second start): if GURI is
    already running, bring it to the foreground and exit; otherwise start it.
    ``--aura`` opens the Aura (data-broker removal) tab.
    """
    from PySide6.QtNetwork import QLocalServer, QLocalSocket

    instance_key = "GeoFooter_GURI_GUI_v1"
    want_raise = "--raise" in sys.argv or "--foreground" in sys.argv
    want_aura = "--aura" in sys.argv or "--brokers" in sys.argv

    def _ping_existing() -> bool:
        sock = QLocalSocket()
        sock.connectToServer(instance_key)
        if not sock.waitForConnected(400):
            return False
        msg = b"AURA\n" if want_aura else b"RAISE\n"
        sock.write(msg)
        sock.flush()
        sock.waitForBytesWritten(800)
        sock.disconnectFromServer()
        return True

    # Always prefer a single instance — Outlook AES button and tray double-start.
    if _ping_existing():
        return

    app = QApplication(sys.argv)
    # Required on some Windows setups so the tray icon identity is stable
    app.setApplicationName(APP_NAME)
    app.setOrganizationName(APP_ORG)
    app.setQuitOnLastWindowClosed(False)
    if os.path.isfile(GURI_APP_ICON_PATH):
        app.setWindowIcon(QIcon(GURI_APP_ICON_PATH))

    # Don't quit when window is closed if system tray is available
    if not QSystemTrayIcon.isSystemTrayAvailable():
        # Show warning but allow app to run
        print("Warning: System tray is not available. Application will close when window is closed.")
        app.setQuitOnLastWindowClosed(True)

    # Set application style + design system
    app.setStyle('Fusion')
    app.setFont(QFont("Segoe UI", 9))
    app.setStyleSheet(GURI_STYLESHEET)

    QLocalServer.removeServer(instance_key)
    server = QLocalServer()
    if not server.listen(instance_key):
        # Lost a race — raise the winner.
        if _ping_existing():
            return

    window = GURIViewerGUI()
    window._instance_server = server  # keep alive

    def _on_instance_connection() -> None:
        conn = server.nextPendingConnection()
        if conn is None:
            return

        def _handle() -> None:
            try:
                if conn.bytesAvailable() == 0:
                    conn.waitForReadyRead(250)
                raw = bytes(conn.readAll()).decode("utf-8", "ignore")
            except Exception:
                raw = "RAISE"
            upper = raw.upper()
            if "AURA" in upper or "BROKER" in upper:
                window.show_aura_tab()
            elif "RAISE" in upper or not raw.strip():
                window.expand_to_screen()

        conn.readyRead.connect(_handle)
        # Peer often writes+disconnects before readyRead is wired; short defer.
        QTimer.singleShot(50, _handle)

    server.newConnection.connect(_on_instance_connection)

    # Always launch filling the screen (availableGeometry + maximized)
    window.expand_to_screen()
    if want_raise:
        QTimer.singleShot(0, window.expand_to_screen)
    if want_aura:
        QTimer.singleShot(0, window.show_aura_tab)

    sys.exit(app.exec())


if __name__ == "__main__":
    main()

