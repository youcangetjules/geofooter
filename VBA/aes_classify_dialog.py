#!/usr/bin/env python3
"""AES security classification picker for Outlook VBA (PySide6).

CLI:
  pythonw aes_classify_dialog.py --subject "..." --default 1 --aliniant true --out result.json

JSON result includes ReadNotify options. AES stamps To/CC/BCC with the computed
suffix on Aliniant sends (do not also use ActiveTracker auto-rewrite on the same send).
"""

from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path
from typing import Any, Dict, List, Tuple

# Must match MSCANCore.bas tag list (1-based indexes in VBA).
CLASSIFICATIONS: List[Tuple[str, str, bool]] = [
    ("[NR/E]", "Minimal / Aliniant Internal", False),
    ("[SEC1: (C) NOT RATED /EXTERNAL]", "Not rated, external", False),
    ("[SEC2: (C) COMMERCIAL-IN-CONFIDENCE /UNENCRYPTED]", "Commercial-in-Confidence / Unencrypted", False),
    (
        "[SEC3: (C) COMMERCIAL-IN-CONFIDENCE-SENSITIVE/SIGNED/EDIT-ENCRYPTED]",
        "CIC-Sensitive / Signed / Edit-Encrypted",
        False,
    ),
    (
        "[SEC4: (C) COMMERCIAL-IN-CONFIDENCE-SENSITIVE/SIGNED/ENCRYPTED/TRACKED]",
        "CIC-Sensitive / Signed / Encrypted / Tracked",
        True,
    ),
    (
        "[SEC5: (D) OFFICIAL-SENSITIVE/SIGNED/ENCRYPTED/TRACKED]",
        "Official-Sensitive / Signed / Encrypted / Tracked",
        True,
    ),
    (
        "[SEC6: (D) CLASSIFIED/SIGNED/ENCRYPTED/TRACKED]",
        "Classified / Signed / Encrypted / Tracked",
        True,
    ),
    (
        "[SEC7: (D) SOA/SIGNED/ENCRYPTED/TRACKED]",
        "Secret or Above / Signed / Encrypted / Tracked",
        True,
    ),
]

# ReadNotify option → address-suffix map (for UI hints only; ActiveTracker rewrites recipients).
RN_DOMAIN = "read-notify.com"
READNOTIFY_SETTINGS_URL = "https://www.readnotify.com/readnotify/login.asp"

STYLESHEET = """
QDialog {
    background: qlineargradient(x1:0, y1:0, x2:1, y2:1,
        stop:0 #e8eef2, stop:0.45 #f3f6f8, stop:1 #dde7ec);
    color: #102a32;
    font-family: "Segoe UI", "Candara", sans-serif;
    font-size: 13px;
}
QLabel#brandLogo {
    background: transparent;
    padding: 0;
}
QLabel#subtitle { color: #4a6670; font-size: 12px; margin-bottom: 2px; }
QLabel#fieldLabel { font-weight: 600; color: #163943; }
QLineEdit#subjectEdit {
    background: #ffffff; border: 1px solid #9bb4be; border-radius: 4px;
    padding: 7px 10px; selection-background-color: #0f6b7c;
}
QWidget#classList { background: transparent; }
QFrame#optionCard {
    background: #ffffff; border: 1px solid #c5d5db; border-radius: 5px;
}
QFrame#optionCard[selected="true"] {
    background: #e6f3f5; border: 1px solid #0f6b7c;
}
QRadioButton { font-weight: 600; color: #12343d; spacing: 8px; }
QLabel#tagLine { color: #5a7380; font-size: 11px; margin-left: 26px; }
QLabel#trackedBadge { color: #0f6b7c; font-size: 11px; font-weight: 700; margin-left: 26px; }
QFrame#readNotifyPanel {
    background: #12343d; border-radius: 8px;
}
QLabel#rnTitle { color: #d7eef2; font-weight: 700; font-size: 13px; }
QCheckBox#rnCheck, QCheckBox#rnOpt {
    color: #f2fbfc;
    font-weight: 600;
    spacing: 10px;
}
QCheckBox#rnCheck::indicator, QCheckBox#rnOpt::indicator {
    width: 15px;
    height: 15px;
    border: 1px solid #ffffff;
    border-radius: 2px;
    background-color: transparent;
}
QCheckBox#rnCheck::indicator:checked, QCheckBox#rnOpt::indicator:checked {
    border: 1px solid #ffffff;
    background-color: #1e6bff;
}
QCheckBox#rnCheck::indicator:disabled, QCheckBox#rnOpt::indicator:disabled {
    border: 1px solid #8aa4ac;
    background-color: transparent;
}
QCheckBox#rnCheck::indicator:checked:disabled, QCheckBox#rnOpt::indicator:checked:disabled {
    border: 1px solid #8aa4ac;
    background-color: #3d5f8a;
}
QCheckBox#rnOpt:disabled { color: #8aa4ac; }
QLabel#rnHelp { color: #a9c8d0; font-size: 11px; }
QPushButton { min-width: 96px; padding: 8px 16px; border-radius: 4px; font-weight: 600; }
QPushButton#cancelBtn { background: #eef3f5; border: 1px solid #9bb4be; color: #234650; }
QPushButton#cancelBtn:hover { background: #e2eaee; }
QPushButton#settingsBtn { background: #ffffff; border: 1px solid #0f6b7c; color: #0f6b7c; }
QPushButton#settingsBtn:hover { background: #e8f5f7; }
QPushButton#applyBtn { background: #0f6b7c; border: 1px solid #0c5663; color: #ffffff; }
QPushButton#applyBtn:hover { background: #128399; }
QPushButton#applyBtn:pressed { background: #0c5663; }
"""


def _resolve_logo_path() -> Path | None:
    candidates: list[Path] = []
    try:
        from geofooter_paths import brand_path

        for name in ("smart-ass.svg", "smart-ass-email.svg", "AES.png"):
            candidates.append(brand_path(name))
    except Exception:
        pass
    parent = Path(__file__).resolve().parent
    candidates.extend(
        [
            parent.parent / "assets" / "brand" / "smart-ass.svg",
            parent.parent / "assets" / "brand" / "AES.png",
            parent / "smart-ass.svg",
            parent / "AES.png",
            Path.home() / "AppData" / "Local" / "GeoFooter" / "AES.png",
        ]
    )
    for path in candidates:
        if path.is_file():
            return path
    return None


def _load_logo_pixmap(height: int = 120):
    """Load brand logo as a QPixmap (SVG uses its native tight viewBox)."""
    from PySide6.QtCore import QRectF, Qt
    from PySide6.QtGui import QPainter, QPixmap

    path = _resolve_logo_path()
    if path is None:
        return None
    if path.suffix.lower() == ".svg":
        try:
            from PySide6.QtSvg import QSvgRenderer
        except ImportError:
            return None
        renderer = QSvgRenderer(str(path))
        if not renderer.isValid():
            return None
        vb = renderer.viewBoxF()
        if vb.width() <= 0 or vb.height() <= 0:
            size = renderer.defaultSize()
            aspect = (size.width() / size.height()) if size.height() else 5.0
        else:
            aspect = vb.width() / vb.height()
        width = max(1, int(round(height * aspect)))
        pix = QPixmap(width, height)
        pix.fill(Qt.GlobalColor.transparent)
        painter = QPainter(pix)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing, True)
        renderer.render(painter, QRectF(0, 0, width, height))
        painter.end()
        return pix
    pix = QPixmap(str(path))
    if pix.isNull():
        return None
    return pix.scaledToHeight(height, Qt.TransformationMode.SmoothTransformation)


def _apply_tag(subject: str, tag: str) -> str:
    cleaned = subject or ""
    for existing, _, _ in CLASSIFICATIONS:
        if cleaned.upper().startswith(existing.upper()):
            cleaned = cleaned[len(existing) :].lstrip(" -:")
            break
    else:
        cleaned = re.sub(
            r"^\[(?:NR/E|SEC\d:[^\]]+)\]\s*",
            "",
            cleaned,
            count=1,
            flags=re.IGNORECASE,
        )
    cleaned = cleaned.strip()
    return f"{tag} {cleaned}" if cleaned else tag


def _write_result(path: Path, payload: Dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, indent=2), encoding="utf-8")


def _parse_bool(value: str) -> bool:
    return str(value).strip().lower() in {"1", "true", "yes", "y", "on"}


def build_read_notify_suffix(options: Dict[str, bool]) -> str:
    """Map dialog options to the ReadNotify address suffix AES will apply."""
    if not options.get("activate"):
        return ""

    # Invisible tracking cannot be combined with other features.
    if options.get("invisible"):
        return f".silent.{RN_DOMAIN}"

    parts: List[str] = []
    if options.get("block_print"):
        parts.append("noprint")
    if options.get("translate"):
        parts.append("translate")
    if options.get("certified"):
        parts.append("certified")
    if options.get("self_destruct"):
        parts.append("selfdestruct")
    elif options.get("ensured"):
        parts.append("ensured")

    if parts:
        return "." + ".".join(parts) + f".{RN_DOMAIN}"
    return f".{RN_DOMAIN}"


def empty_rn_options() -> Dict[str, bool]:
    return {
        "activate": False,
        "invisible": False,
        "ensured": False,
        "self_destruct": False,
        "block_print": False,
        "certified": False,
        "translate": False,
    }


def run_dialog(subject: str, default_index: int, out_path: Path, aliniant: bool) -> int:
    try:
        from PySide6.QtCore import Qt, QPropertyAnimation, QEasingCurve
        from PySide6.QtGui import QPixmap
        from PySide6.QtWidgets import (
            QApplication,
            QButtonGroup,
            QCheckBox,
            QDialog,
            QFrame,
            QGraphicsOpacityEffect,
            QHBoxLayout,
            QLabel,
            QLineEdit,
            QPushButton,
            QRadioButton,
            QVBoxLayout,
            QWidget,
        )
    except ImportError:
        print("PySide6 is required for aes_classify_dialog.py", file=sys.stderr)
        return 2

    app = QApplication.instance() or QApplication(sys.argv)

    dialog = QDialog()
    dialog.setWindowTitle("AES Security Classification")
    dialog.setModal(True)
    dialog.setWindowFlag(Qt.WindowType.WindowStaysOnTopHint, True)
    dialog.setStyleSheet(STYLESHEET)

    root = QVBoxLayout(dialog)
    root.setContentsMargins(18, 14, 18, 12)
    root.setSpacing(6)

    brand = QLabel()
    brand.setObjectName("brandLogo")
    brand.setAlignment(Qt.AlignmentFlag.AlignCenter)
    pixmap = _load_logo_pixmap(96)
    if pixmap is not None and not pixmap.isNull():
        brand.setPixmap(pixmap)
        brand.setFixedSize(pixmap.width(), pixmap.height())
    else:
        brand.setText("AES")
    root.addWidget(brand, 0, Qt.AlignmentFlag.AlignHCenter)

    subtitle = QLabel("Security Classification")
    subtitle.setObjectName("subtitle")
    root.addWidget(subtitle)

    subject_label = QLabel("Subject")
    subject_label.setObjectName("fieldLabel")
    root.addWidget(subject_label)

    subject_edit = QLineEdit(subject or "")
    subject_edit.setObjectName("subjectEdit")
    subject_edit.setPlaceholderText("Email subject")
    root.addWidget(subject_edit)

    hint = QLabel("Choose a classification. Optional: enable ReadNotify (AES stamps recipients).")
    hint.setObjectName("subtitle")
    root.addWidget(hint)

    class_list = QWidget()
    class_list.setObjectName("classList")
    class_layout = QVBoxLayout(class_list)
    class_layout.setContentsMargins(0, 2, 0, 2)
    class_layout.setSpacing(4)

    group = QButtonGroup(dialog)
    cards: List[QFrame] = []
    radios: List[QRadioButton] = []
    default_index = max(1, min(len(CLASSIFICATIONS), default_index))

    for i, (tag, desc, tracked) in enumerate(CLASSIFICATIONS, start=1):
        card = QFrame()
        card.setObjectName("optionCard")
        card.setProperty("selected", "true" if i == default_index else "false")
        card_layout = QVBoxLayout(card)
        card_layout.setContentsMargins(8, 5, 8, 5)
        card_layout.setSpacing(1)

        title = f"{i}. {desc}"
        if tracked:
            title += "  [TRACKED]"
        radio = QRadioButton(title)
        radio.setToolTip(tag)
        if i == default_index:
            radio.setChecked(True)
        group.addButton(radio, i)
        radios.append(radio)
        card_layout.addWidget(radio)

        tag_line = QLabel(tag)
        tag_line.setObjectName("tagLine")
        tag_line.setWordWrap(True)
        card_layout.addWidget(tag_line)

        class_layout.addWidget(card)
        cards.append(card)

    root.addWidget(class_list)

    rn_panel = QFrame()
    rn_panel.setObjectName("readNotifyPanel")
    rn_layout = QVBoxLayout(rn_panel)
    rn_layout.setContentsMargins(14, 10, 14, 10)
    rn_layout.setSpacing(3)

    rn_title = QLabel("ReadNotify tracking (AES applies on send)")
    rn_title.setObjectName("rnTitle")
    rn_layout.addWidget(rn_title)

    rn_activate = QCheckBox("Activate Tracking")
    rn_activate.setObjectName("rnCheck")
    rn_layout.addWidget(rn_activate)

    rn_invisible = QCheckBox("Invisible Tracking")
    rn_invisible.setObjectName("rnOpt")
    rn_layout.addWidget(rn_invisible)

    rn_ensured = QCheckBox("Ensured + Retractable")
    rn_ensured.setObjectName("rnOpt")
    rn_layout.addWidget(rn_ensured)

    rn_self_destruct = QCheckBox("Self-Destructing")
    rn_self_destruct.setObjectName("rnOpt")
    rn_layout.addWidget(rn_self_destruct)

    rn_block_print = QCheckBox("Block Print & Copy")
    rn_block_print.setObjectName("rnOpt")
    rn_layout.addWidget(rn_block_print)

    rn_certified = QCheckBox("Certified + Notarized")
    rn_certified.setObjectName("rnOpt")
    rn_layout.addWidget(rn_certified)

    rn_translate = QCheckBox("Translate Language")
    rn_translate.setObjectName("rnOpt")
    rn_layout.addWidget(rn_translate)

    rn_help = QLabel("")
    rn_help.setObjectName("rnHelp")
    rn_help.setWordWrap(True)
    rn_layout.addWidget(rn_help)

    feature_checks = [
        rn_invisible,
        rn_ensured,
        rn_self_destruct,
        rn_block_print,
        rn_certified,
        rn_translate,
    ]

    opacity = QGraphicsOpacityEffect(rn_panel)
    rn_panel.setGraphicsEffect(opacity)
    opacity.setOpacity(0.95)
    fade = QPropertyAnimation(opacity, b"opacity", dialog)
    fade.setDuration(160)
    fade.setEasingCurve(QEasingCurve.Type.OutCubic)

    root.addWidget(rn_panel)

    def tracked_default_for(idx: int) -> bool:
        return 1 <= idx <= len(CLASSIFICATIONS) and CLASSIFICATIONS[idx - 1][2]

    def current_options() -> Dict[str, bool]:
        return {
            "activate": rn_activate.isChecked(),
            "invisible": rn_invisible.isChecked(),
            "ensured": rn_ensured.isChecked(),
            "self_destruct": rn_self_destruct.isChecked(),
            "block_print": rn_block_print.isChecked(),
            "certified": rn_certified.isChecked(),
            "translate": rn_translate.isChecked(),
        }

    def update_cards(selected_id: int) -> None:
        for i, card in enumerate(cards, start=1):
            card.setProperty("selected", "true" if i == selected_id else "false")
            card.style().unpolish(card)
            card.style().polish(card)

    def update_read_notify_ui() -> None:
        active = rn_activate.isChecked()
        for chk in feature_checks:
            chk.setEnabled(active)

        if not active:
            for chk in feature_checks:
                chk.blockSignals(True)
                chk.setChecked(False)
                chk.blockSignals(False)
        elif rn_invisible.isChecked():
            # Invisible cannot combine with other features.
            for chk in (rn_ensured, rn_self_destruct, rn_block_print, rn_certified, rn_translate):
                chk.blockSignals(True)
                chk.setChecked(False)
                chk.setEnabled(False)
                chk.blockSignals(False)

        if rn_ensured.isChecked() and rn_self_destruct.isChecked():
            # Prefer the one just toggled; handled in handlers.
            pass

        opts = current_options()
        suffix = build_read_notify_suffix(opts)
        if not opts["activate"]:
            rn_help.setText("Tracking off — AES will leave To/CC/BCC unchanged.")
        elif not aliniant:
            rn_help.setText(
                "Not sending from Aliniant — AES will not stamp recipients. "
                "Switch From to the Aliniant account, or use ReadNotify separately."
            )
        else:
            rn_help.setText(
                "AES will rewrite To/CC/BCC to …"
                + (suffix if suffix else ".read-notify.com")
                + ". Leave ReadNotify ActiveTracker auto-track OFF for this send "
                "so addresses are not stamped twice."
            )

        fade.stop()
        fade.setStartValue(0.6)
        fade.setEndValue(1.0)
        fade.start()

    def on_activate_toggled(_checked: bool) -> None:
        update_read_notify_ui()

    def on_invisible_toggled(checked: bool) -> None:
        if checked:
            for chk in (rn_ensured, rn_self_destruct, rn_block_print, rn_certified, rn_translate):
                chk.blockSignals(True)
                chk.setChecked(False)
                chk.blockSignals(False)
        update_read_notify_ui()

    def on_ensured_toggled(checked: bool) -> None:
        if checked:
            rn_self_destruct.blockSignals(True)
            rn_self_destruct.setChecked(False)
            rn_self_destruct.blockSignals(False)
            rn_invisible.blockSignals(True)
            rn_invisible.setChecked(False)
            rn_invisible.blockSignals(False)
        update_read_notify_ui()

    def on_self_destruct_toggled(checked: bool) -> None:
        if checked:
            rn_ensured.blockSignals(True)
            rn_ensured.setChecked(False)
            rn_ensured.blockSignals(False)
            rn_invisible.blockSignals(True)
            rn_invisible.setChecked(False)
            rn_invisible.blockSignals(False)
        update_read_notify_ui()

    def on_feature_toggled(checked: bool) -> None:
        if checked:
            rn_invisible.blockSignals(True)
            rn_invisible.setChecked(False)
            rn_invisible.blockSignals(False)
        update_read_notify_ui()

    def on_selection_changed() -> None:
        idx = group.checkedId()
        update_cards(idx)
        if aliniant and tracked_default_for(idx):
            rn_activate.setChecked(True)
        update_read_notify_ui()

    # Defaults
    rn_activate.setChecked(aliniant and tracked_default_for(default_index))
    update_cards(default_index)
    update_read_notify_ui()

    for radio in radios:
        radio.toggled.connect(lambda checked: on_selection_changed() if checked else None)

    rn_activate.toggled.connect(on_activate_toggled)
    rn_invisible.toggled.connect(on_invisible_toggled)
    rn_ensured.toggled.connect(on_ensured_toggled)
    rn_self_destruct.toggled.connect(on_self_destruct_toggled)
    rn_block_print.toggled.connect(on_feature_toggled)
    rn_certified.toggled.connect(on_feature_toggled)
    rn_translate.toggled.connect(on_feature_toggled)

    # Windows dialog order: primary action left of Cancel; Cancel bottom-right.
    buttons = QHBoxLayout()
    settings_btn = QPushButton("ReadNotify Settings")
    settings_btn.setObjectName("settingsBtn")
    settings_btn.setToolTip("Open the ReadNotify account / settings page in your browser")
    buttons.addWidget(settings_btn)
    buttons.addStretch(1)
    apply_btn = QPushButton("Apply")
    apply_btn.setObjectName("applyBtn")
    apply_btn.setDefault(True)
    cancel_btn = QPushButton("Cancel")
    cancel_btn.setObjectName("cancelBtn")
    buttons.addWidget(apply_btn)
    buttons.addWidget(cancel_btn)
    root.addLayout(buttons)

    # Size to fit all classifications + ReadNotify panel without a scrollbar.
    dialog.adjustSize()
    hint = dialog.sizeHint()
    dialog.setMinimumWidth(640)
    dialog.resize(max(680, hint.width()), max(hint.height() + 12, 900))

    result: Dict[str, Any] = {
        "cancelled": True,
        "index": 0,
        "tag": "",
        "subject": subject or "",
        "read_notify": False,
        "read_notify_suffix": "",
        "read_notify_options": empty_rn_options(),
    }

    def on_cancel() -> None:
        result["cancelled"] = True
        result["index"] = 0
        result["tag"] = ""
        result["subject"] = subject_edit.text().strip() or subject or ""
        result["read_notify"] = False
        result["read_notify_suffix"] = ""
        result["read_notify_options"] = empty_rn_options()
        dialog.reject()

    def on_apply() -> None:
        idx = group.checkedId()
        if idx < 1 or idx > len(CLASSIFICATIONS):
            return
        tag = CLASSIFICATIONS[idx - 1][0]
        base_subject = subject_edit.text().strip()
        opts = current_options()
        suffix = build_read_notify_suffix(opts)
        result["cancelled"] = False
        result["index"] = idx
        result["tag"] = tag
        result["subject"] = _apply_tag(base_subject, tag)
        result["read_notify"] = bool(opts["activate"])
        result["read_notify_suffix"] = suffix
        result["read_notify_options"] = opts
        dialog.accept()

    def on_open_settings() -> None:
        import webbrowser

        webbrowser.open(READNOTIFY_SETTINGS_URL)

    cancel_btn.clicked.connect(on_cancel)
    apply_btn.clicked.connect(on_apply)
    settings_btn.clicked.connect(on_open_settings)

    dialog.exec()
    _write_result(out_path, result)
    return 0


def main() -> int:
    parser = argparse.ArgumentParser(description="AES classification dialog for Outlook")
    parser.add_argument("--subject", default="", help="Original email subject")
    parser.add_argument("--default", type=int, default=1, help="1-based default classification index")
    parser.add_argument(
        "--aliniant",
        default="false",
        help="true if send account is Aliniant (affects ReadNotify defaults/help)",
    )
    parser.add_argument(
        "--out",
        default="",
        help="Output JSON path (default: %%LOCALAPPDATA%%/GeoFooter/classify_result.json)",
    )
    args = parser.parse_args()

    if args.out:
        out_path = Path(args.out)
    else:
        out_path = Path.home() / "AppData" / "Local" / "GeoFooter" / "classify_result.json"

    try:
        return run_dialog(args.subject, args.default, out_path, _parse_bool(args.aliniant))
    except Exception as exc:  # noqa: BLE001
        print(f"aes_classify_dialog failed: {exc}", file=sys.stderr)
        try:
            _write_result(
                out_path,
                {
                    "cancelled": True,
                    "index": 0,
                    "tag": "",
                    "subject": args.subject or "",
                    "read_notify": False,
                    "read_notify_suffix": "",
                    "read_notify_options": empty_rn_options(),
                    "error": str(exc),
                },
            )
        except Exception:
            pass
        return 1


if __name__ == "__main__":
    sys.exit(main())
