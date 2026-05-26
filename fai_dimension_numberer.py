#!/usr/bin/env python3
"""
fai_dimension_numberer.py
=========================
FAI Dimension Numberer — engineering drawing callout annotation and
measurement record population tool.

Workflow:
  1. App loads measurement_template.xlsx from its own directory on startup.
  2. User opens a PDF engineering drawing.
  3. Regex engine detects dimensions; callout circles are overlaid on the drawing.
  4. User adjusts which callouts to include via the right-hand checklist.
  5. User exports an annotated PDF (moveable FreeText annotations) and/or
     a filled copy of the measurement template (formulas intact).
"""

# ---------------------------------------------------------------------------
# Sys-path setup — must run before any local imports so both dev and frozen
# PyInstaller exe can find dimension_detector.py in the same directory.
# ---------------------------------------------------------------------------
import sys
import os

_APP_DIR: str = (
    os.path.dirname(sys.executable)
    if getattr(sys, "frozen", False)
    else os.path.dirname(os.path.abspath(__file__))
)
if _APP_DIR not in sys.path:
    sys.path.insert(0, _APP_DIR)

# ---------------------------------------------------------------------------
# Standard library
# ---------------------------------------------------------------------------
import json
import math
import re
import shutil
from typing import Any, Callable, Optional

# ---------------------------------------------------------------------------
# PyQt6
# ---------------------------------------------------------------------------
from PyQt6.QtCore import (
    QEvent,
    QPointF,
    QRectF,
    QSize,
    Qt,
    QThread,
    QTimer,
    pyqtSignal,
    pyqtSlot,
)
from PyQt6.QtGui import (
    QBrush,
    QColor,
    QFont,
    QFontMetricsF,
    QImage,
    QPainter,
    QPainterPath,
    QPen,
    QPalette,
    QPixmap,
)
from PyQt6.QtWidgets import (
    QApplication,
    QColorDialog,
    QDialog,
    QDialogButtonBox,
    QFileDialog,
    QFormLayout,
    QFrame,
    QGraphicsItem,
    QGraphicsScene,
    QGraphicsView,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMenu,
    QMessageBox,
    QPushButton,
    QScrollArea,
    QSizePolicy,
    QSlider,
    QSpinBox,
    QSplitter,
    QStackedWidget,
    QStatusBar,
    QVBoxLayout,
    QWidget,
)

# ---------------------------------------------------------------------------
# Third-party — non-fatal import guards (errors shown in main())
# ---------------------------------------------------------------------------
try:
    import fitz  # PyMuPDF
except ImportError:
    fitz = None  # type: ignore[assignment]

try:
    from openpyxl import load_workbook  # type: ignore[import]
    from openpyxl.utils import get_column_letter  # type: ignore[import]
except ImportError:
    load_workbook = None  # type: ignore[assignment]

# ---------------------------------------------------------------------------
# Local module — dimension_detector.py (same directory)
# ---------------------------------------------------------------------------
_DETECTOR_IMPORT_ERROR: Optional[str] = None
detect_dimensions: Optional[Callable] = None
parse_spreadsheet_template: Optional[Callable] = None

try:
    from dimension_detector import (  # type: ignore[import]
        detect_dimensions,
        parse_spreadsheet_template,
    )
except ImportError as _ie:
    _DETECTOR_IMPORT_ERROR = str(_ie)

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------
TEMPLATE_FILENAME = "MSI-VSA_Template_Dim_Check.xlsx"
CONFIG_FILENAME = "config.json"
SIDECAR_EXT = ".autodim"   # Session sidecar saved alongside every exported PDF
RENDER_ZOOM = 2.0          # Render PDF at 144 DPI (2× 72 pt/inch)
DEFAULT_CALLOUT_RADIUS = 14  # Scene pixels
DEFAULT_CALLOUT_COLOR = "#E74C3C"

# MSI-VSA measurement template fixed column layout
_COL_DIM_NUM    = 2   # Column B — Callout / Balloon number
_COL_NOMINAL    = 3   # Column C — Nominal value
_COL_PLUS_TOL   = 4   # Column D — Plus tolerance (positive)
_COL_MINUS_TOL  = 5   # Column E — Minus tolerance (positive)
_DATA_START_ROW = 16  # First data entry row in the MSI-VSA template

CATEGORY_COLORS: dict[str, str] = {
    "linear":   "#2980B9",
    "angular":  "#8E44AD",
    "GD&T":     "#D35400",
    "thread":   "#27AE60",
    "finish":   "#B7950B",
    "diameter": "#C0392B",
    "radius":   "#C2185B",
    "datum":    "#5D4037",
    "other":    "#7F8C8D",
}


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _nominal_text(text: str) -> str:
    """Strip count prefix and tolerance suffixes; return only the nominal value."""
    t = text.strip()
    # Remove leading count multiplier: "3X ", "2x ", "4× ", etc.
    t = re.sub(r"^\d+\s*[xX×]\s*", "", t)
    # Remove trailing tolerance suffixes
    t = re.sub(r"\s*[±]\s*[\d.]+.*$", "", t)
    t = re.sub(r"\s*\+[\d.]+\s*/?\s*-[\d.]+.*$", "", t)
    t = re.sub(r"\s*\+[\d.]+\s+-[\d.]+.*$", "", t)
    return t.strip() or text


def _parse_tolerances(
    text: str,
) -> tuple[Optional[float], Optional[float], Optional[float]]:
    """Return (nominal, plus_tol, minus_tol) from a dimension text string.

    All returned values are positive floats, or None when not parseable.
    """
    t = text.strip()
    # Strip count prefix: "2X", "4x", "3×", "2X " etc.
    t = re.sub(r"^\d+\s*[xX×]\s*", "", t)
    t = re.sub(r"^[Øøâ£⌀∅]\s*", "", t)
    t = re.sub(r"^[Rr]\s*", "", t)

    # "25.4 ±0.2"
    m = re.match(r"^([+-]?\d+\.?\d*)\s*[±]\s*([0-9.]+)", t)
    if m:
        return float(m.group(1)), float(m.group(2)), float(m.group(2))

    # "25.4 +0.3/-0.1" or "25.4 +0.3 -0.1"
    m = re.match(r"^([+-]?\d+\.?\d*)\s*\+([0-9.]+)\s*/?\s*-([0-9.]+)", t)
    if m:
        return float(m.group(1)), float(m.group(2)), float(m.group(3))

    # Plain number
    m = re.match(r"^([+-]?\d+\.?\d*)\s*[°%]?\s*$", t)
    if m:
        return float(m.group(1)), None, None

    # Fractional inch "1/4" → 0.25
    m = re.match(r"^(\d+)/(\d+)$", t)
    if m and int(m.group(2)) != 0:
        return int(m.group(1)) / int(m.group(2)), None, None

    return None, None, None


def _count_decimal_places(text: str) -> int:
    """Return the number of decimal digits in a dimension text's nominal value.

    Strips common prefixes (count multiplier, Ø, R, ±, …) before inspecting
    the number.
    Examples: "25" → 0,  "25.4" → 1,  "2X 16.00" → 2,  "4x Ø5.5" → 1.
    """
    t = text.strip()
    t = re.sub(r'^\d+\s*[xX×]\s*', '', t)   # strip "2X ", "4x " etc.
    t = re.sub(r'^[ØøRr⌀∅±+\-\s]*', '', t)
    m = re.match(r'^[\d]+\.([\d]+)', t)
    return len(m.group(1)) if m else 0


def _parse_title_block_tolerances(text: str) -> dict:
    """Parse general-tolerance declarations from a title-block text blob.

    Handles both signed and unsigned formats, e.g.:

        X                    NO DECIMALS 1         ← bare number (no ±)
        .X                   DECIMAL  .3
        .XX                  DECIMAL  .13
        ANGULAR OR BEND      1/2                   ← fraction = 0.5°

        X NO DECIMALS ±1                           ← signed format
        .X DECIMAL +/- 0.3
        ANGULAR OR BEND +/- 1.2 DEGREE

    Strategy: process line-by-line (the natural format for title blocks).
    For each line that matches a format marker (X, .X, .XX, ANGULAR/BEND),
    extract the tolerance value with _tol_value(), which first looks for an
    explicit ± / +/- sign and falls back to the last numeric token on the
    line (bare number or N/M fraction).

    Returns a dict keyed by decimal-place count (int) or "angular" (str),
    e.g. {0: 1.0, 1: 0.3, 2: 0.13, "angular": 0.5}.
    Returns an empty dict when no tolerance declarations are found.
    """
    result: dict = {}
    # Normalize horizontal whitespace only; preserve newlines so line-by-line
    # parsing works on multi-line title blocks.
    upper = re.sub(r'[ \t]+', ' ', text.upper())

    def _tol_value(segment: str) -> Optional[float]:
        """Extract a tolerance magnitude from a text segment.

        Tries explicit-sign formats first (±N, +/-N, +N/-N), then falls back
        to the last numeric token on the line — either a plain decimal number
        or an N/M fraction (e.g. "1/2" → 0.5).
        """
        # --- Explicit sign formats ---
        for pat in (
            r'[±]\s*(\d+\.?\d*)',                         # ±0.13
            r'[+]\s*/\s*[-]\s*(\d+\.?\d*)',               # +/- 0.13
            r'[+](\d+\.?\d*)\s*/?\s*-\d+\.?\d*',         # +0.13/-0.13
        ):
            m = re.search(pat, segment)
            if m:
                try:
                    return float(m.group(1))
                except ValueError:
                    pass

        # --- Bare number / fraction fallback ---
        # Find every numeric token; take the last one as the tolerance value.
        # Fractions like "1/2" are resolved to 0.5.
        tokens = list(re.finditer(r'(\d+)\s*/\s*(\d+)|(\d*\.?\d+)', segment))
        if tokens:
            last = tokens[-1]
            if last.group(1) is not None:       # N/M fraction
                den = int(last.group(2))
                return int(last.group(1)) / den if den else None
            try:
                return float(last.group(3))
            except (ValueError, TypeError):
                pass
        return None

    for raw_line in upper.splitlines():
        line = raw_line.strip()
        if not line:
            continue

        # Angular / bend (check before X so "BEND" lines are not caught by X)
        if re.search(r'\bANGULAR\b|\bBEND\b', line):
            t = _tol_value(line)
            if t is not None:
                result["angular"] = t
            continue

        # Whole-number: line starts with standalone "X" (not part of ".X…")
        if re.match(r'^X\b', line):
            t = _tol_value(line)
            if t is not None:
                result[0] = t
            continue

        # .X, .XX, .XXX … — line starts with ".X"
        m = re.match(r'^(\.X+)\b', line)
        if m:
            decimals = len(m.group(1)) - 1   # .X→1, .XX→2, .XXX→3
            t = _tol_value(line)
            if t is not None:
                result[decimals] = t

    # ------------------------------------------------------------------ #
    # Single-line fallback: if line-by-line found nothing, try splitting  #
    # on commas/semicolons (some title blocks pack everything on one line) #
    # ------------------------------------------------------------------ #
    if not result:
        for seg in re.split(r'[,;]', upper):
            seg = seg.strip()
            if not seg:
                continue
            if re.search(r'\bANGULAR\b|\bBEND\b', seg):
                t = _tol_value(seg)
                if t is not None:
                    result["angular"] = t
            elif re.match(r'^X\b', seg):
                t = _tol_value(seg)
                if t is not None:
                    result[0] = t
            else:
                dm = re.match(r'^(\.X+)\b', seg)
                if dm:
                    decimals = len(dm.group(1)) - 1
                    t = _tol_value(seg)
                    if t is not None:
                        result[decimals] = t

    return result


def _smart_callout_pos(
    block: dict,
    placed: list[tuple[float, float]],
    r_pt: float,
) -> tuple[float, float]:
    """Pick the callout centre with fewest overlaps from 8 candidate positions.

    Candidates are tried in preference order: right, left, above, below,
    then the four diagonals.  The first candidate with zero overlaps wins;
    otherwise the one with the fewest overlaps is returned.
    """
    bx = block["x"]
    by = block["y"]
    bw = block["width"]
    bh = block["height"]
    bcx = bx + bw / 2.0
    bcy = by + bh / 2.0
    gap = r_pt + 4.0

    candidates = [
        (bx + bw + gap,  bcy),           # right (preferred)
        (bx - gap,       bcy),           # left
        (bcx,            by - gap),      # above
        (bcx,            by + bh + gap), # below
        (bx + bw + gap,  by - gap),      # upper-right
        (bx - gap,       by - gap),      # upper-left
        (bx + bw + gap,  by + bh + gap), # lower-right
        (bx - gap,       by + bh + gap), # lower-left
    ]

    best = candidates[0]
    min_overlaps: Optional[int] = None
    diameter = r_pt * 2.0

    for cx, cy in candidates:
        overlaps = sum(
            1 for px, py in placed if math.hypot(cx - px, cy - py) < diameter
        )
        if min_overlaps is None or overlaps < min_overlaps:
            min_overlaps = overlaps
            best = (cx, cy)
            if overlaps == 0:
                break

    return best


def _find_column(headers: list[dict], keywords: list[str]) -> Optional[int]:
    """Return the column index of the first header whose text contains any keyword."""
    for h in headers:
        val = str(h.get("value", "")).lower().strip()
        for kw in keywords:
            if kw in val:
                return int(h["column_index"])
    return None


# ---------------------------------------------------------------------------
# Config — reads / writes config.json next to the executable
# ---------------------------------------------------------------------------

class Config:
    def __init__(self) -> None:
        self._path = os.path.join(_APP_DIR, CONFIG_FILENAME)
        self._data: dict[str, Any] = {
            "callout_color": DEFAULT_CALLOUT_COLOR,
            "callout_radius": DEFAULT_CALLOUT_RADIUS,
        }
        self._load()

    def _load(self) -> None:
        if os.path.isfile(self._path):
            try:
                with open(self._path, "r", encoding="utf-8") as fh:
                    saved = json.load(fh)
                self._data.update(saved)
            except Exception:
                pass

    def save(self) -> None:
        try:
            with open(self._path, "w", encoding="utf-8") as fh:
                json.dump(self._data, fh, indent=2)
        except Exception:
            pass

    @property
    def callout_color(self) -> str:
        return self._data.get("callout_color", DEFAULT_CALLOUT_COLOR)

    @callout_color.setter
    def callout_color(self, v: str) -> None:
        self._data["callout_color"] = v

    @property
    def callout_radius(self) -> int:
        return int(self._data.get("callout_radius", DEFAULT_CALLOUT_RADIUS))

    @callout_radius.setter
    def callout_radius(self, v: int) -> None:
        self._data["callout_radius"] = int(v)


# ---------------------------------------------------------------------------
# Background worker thread — runs detect_dimensions() without blocking the UI
# ---------------------------------------------------------------------------

class DetectionWorker(QThread):
    finished = pyqtSignal(list)
    error = pyqtSignal(str)

    def __init__(self, pdf_path: str, page_num: int) -> None:
        super().__init__()
        self.pdf_path = pdf_path
        self.page_num = page_num

    def run(self) -> None:
        try:
            result = detect_dimensions(self.pdf_path, self.page_num)
            self.finished.emit(result)
        except RuntimeError as exc:
            self.error.emit(str(exc))
        except Exception as exc:
            self.error.emit(f"{type(exc).__name__}: {exc}")


# ---------------------------------------------------------------------------
# Spinner widget — animated loading indicator drawn with QPainter
# ---------------------------------------------------------------------------

class SpinnerWidget(QWidget):
    def __init__(self, parent: Optional[QWidget] = None) -> None:
        super().__init__(parent)
        self._angle = 0
        self._timer = QTimer(self)
        self._timer.timeout.connect(self._tick)

    def start(self) -> None:
        self._timer.start(40)

    def stop(self) -> None:
        self._timer.stop()

    def _tick(self) -> None:
        self._angle = (self._angle + 12) % 360
        self.update()

    def paintEvent(self, _event) -> None:  # noqa: N802
        w, h = self.width(), self.height()
        r = min(w, h) // 2 - 4
        cx, cy = w // 2, h // 2
        n = 10

        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)

        for i in range(n):
            angle_deg = (self._angle + i * (360 // n)) % 360
            alpha = int(255 * (i + 1) / n)
            color = QColor(231, 76, 60, alpha)
            pen = QPen(color, 3, Qt.PenStyle.SolidLine, Qt.PenCapStyle.RoundCap)
            painter.setPen(pen)

            rad = math.radians(angle_deg)
            x1 = cx + (r - 9) * math.cos(rad)
            y1 = cy - (r - 9) * math.sin(rad)
            x2 = cx + r * math.cos(rad)
            y2 = cy - r * math.sin(rad)
            painter.drawLine(int(x1), int(y1), int(x2), int(y2))

        painter.end()


# ---------------------------------------------------------------------------
# CalloutItem — draggable numbered circle overlay on the PDF scene
# ---------------------------------------------------------------------------

class CalloutItem(QGraphicsItem):
    def __init__(
        self,
        number: int,
        radius: int,
        color: str,
        on_moved: Optional[Callable] = None,
        on_drag_complete: Optional[Callable] = None,
        on_context_menu: Optional[Callable] = None,
    ) -> None:
        super().__init__()
        self.number = number
        self.radius = radius
        self.color = QColor(color)
        self._on_moved = on_moved
        self._on_drag_complete = on_drag_complete
        self._on_context_menu = on_context_menu
        self._drag_start_pos: Optional[QPointF] = None

        self.setFlag(QGraphicsItem.GraphicsItemFlag.ItemIsMovable, True)
        self.setFlag(QGraphicsItem.GraphicsItemFlag.ItemSendsGeometryChanges, True)
        self.setFlag(QGraphicsItem.GraphicsItemFlag.ItemIsSelectable, True)
        self.setZValue(10)
        self.setCursor(Qt.CursorShape.SizeAllCursor)

    def set_number(self, n: int) -> None:
        self.number = n
        self.update()

    def set_color(self, color_hex: str) -> None:
        self.color = QColor(color_hex)
        self.update()

    def set_radius(self, r: int) -> None:
        self.prepareGeometryChange()
        self.radius = r
        self.update()

    def boundingRect(self) -> QRectF:
        # Extra space accommodates the selection ring (r+5 + 1.5 px pen)
        r = self.radius + 9.0
        return QRectF(-r, -r, 2 * r, 2 * r)

    def shape(self) -> QPainterPath:
        # Hit-test area stays tight so only the circle itself is clickable
        path = QPainterPath()
        r = float(self.radius)
        path.addEllipse(QRectF(-r, -r, 2 * r, 2 * r))
        return path

    def paint(self, painter: QPainter, _option, _widget) -> None:
        r = float(self.radius)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)

        # Selection ring — drawn first so it sits behind the main circle
        if self.isSelected():
            sel_pen = QPen(QColor(52, 152, 219), 2.5, Qt.PenStyle.DashLine)
            painter.setPen(sel_pen)
            painter.setBrush(Qt.BrushStyle.NoBrush)
            rr = r + 5.0
            painter.drawEllipse(QRectF(-rr, -rr, 2 * rr, 2 * rr))

        # White fill + colored border
        painter.setPen(QPen(self.color, 2.0))
        painter.setBrush(QBrush(QColor(255, 255, 255, 230)))
        painter.drawEllipse(QRectF(-r, -r, 2 * r, 2 * r))

        # Number label — font shrinks automatically so it always fits the circle
        num_str = str(self.number)
        base_pt = max(6.0, r * 0.72)
        font = QFont("Arial", 1, QFont.Weight.Bold)
        font.setPointSizeF(base_pt)
        # Available horizontal space = diameter minus 2 px margin each side
        available_w = max(1.0, 2.0 * r - 4.0)
        text_w = QFontMetricsF(font).horizontalAdvance(num_str)
        if text_w > available_w:
            font.setPointSizeF(max(5.0, base_pt * available_w / text_w))
        painter.setFont(font)
        painter.setPen(QPen(self.color))
        painter.drawText(
            QRectF(-r, -r, 2 * r, 2 * r),
            Qt.AlignmentFlag.AlignCenter,
            num_str,
        )

    def mousePressEvent(self, event) -> None:
        if event.button() == Qt.MouseButton.LeftButton:
            self._drag_start_pos = self.pos()
        super().mousePressEvent(event)

    def mouseReleaseEvent(self, event) -> None:
        super().mouseReleaseEvent(event)
        if (
            event.button() == Qt.MouseButton.LeftButton
            and self._drag_start_pos is not None
            and self._on_drag_complete is not None
        ):
            end_pos = self.pos()
            # Only record if actually moved (not just a click-to-select)
            if (abs(end_pos.x() - self._drag_start_pos.x()) > 0.5 or
                    abs(end_pos.y() - self._drag_start_pos.y()) > 0.5):
                self._on_drag_complete(self, self._drag_start_pos, end_pos)
        self._drag_start_pos = None

    def itemChange(self, change, value):
        if (
            change == QGraphicsItem.GraphicsItemChange.ItemPositionHasChanged
            and self._on_moved is not None
        ):
            self._on_moved(self, value)
        return super().itemChange(change, value)

    def contextMenuEvent(self, event) -> None:
        """Right-click on a callout circle — fire the context-menu callback."""
        if self._on_context_menu is not None:
            self._on_context_menu(self, event.screenPos())
        event.accept()


# ---------------------------------------------------------------------------
# PDFViewer — zoomable QGraphicsView with add-callout mode
# ---------------------------------------------------------------------------

class PDFViewer(QGraphicsView):
    callout_place_requested  = pyqtSignal(float, float)  # scene x, y
    callout_delete_requested = pyqtSignal(object)        # selected CalloutItem

    def __init__(self, parent: Optional[QWidget] = None) -> None:
        super().__init__(parent)
        self._scene = QGraphicsScene()
        self.setScene(self._scene)
        self._add_mode = False

        self.setDragMode(QGraphicsView.DragMode.ScrollHandDrag)
        self.setTransformationAnchor(QGraphicsView.ViewportAnchor.AnchorUnderMouse)
        self.setResizeAnchor(QGraphicsView.ViewportAnchor.AnchorUnderMouse)
        self.setRenderHint(QPainter.RenderHint.Antialiasing)
        self.setRenderHint(QPainter.RenderHint.SmoothPixmapTransform)
        self.setBackgroundBrush(QBrush(QColor("#1A1A2E")))
        self.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        self.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        self.setMinimumSize(400, 300)
        # Allow the view to receive key events after clicking a callout
        self.setFocusPolicy(Qt.FocusPolicy.StrongFocus)

    # --- Public API --------------------------------------------------------

    def set_add_mode(self, enabled: bool) -> None:
        self._add_mode = enabled
        if enabled:
            self.setDragMode(QGraphicsView.DragMode.NoDrag)
            self.viewport().setCursor(Qt.CursorShape.CrossCursor)
        else:
            self.setDragMode(QGraphicsView.DragMode.ScrollHandDrag)
            self.viewport().setCursor(Qt.CursorShape.ArrowCursor)

    def display_pixmap(self, pixmap: QPixmap) -> None:
        self._scene.clear()
        self._scene.addPixmap(pixmap)
        self._scene.setSceneRect(QRectF(pixmap.rect()))

    def fit_page(self) -> None:
        if self._scene.sceneRect().isValid():
            self.fitInView(self._scene.sceneRect(), Qt.AspectRatioMode.KeepAspectRatio)

    def add_callout_item(self, item: CalloutItem) -> None:
        self._scene.addItem(item)

    def remove_callout_item(self, item: CalloutItem) -> None:
        if item.scene() == self._scene:
            self._scene.removeItem(item)

    def clear_callout_items(self) -> None:
        for item in list(self._scene.items()):
            if isinstance(item, CalloutItem):
                self._scene.removeItem(item)

    def zoom_in(self) -> None:
        self.scale(1.25, 1.25)

    def zoom_out(self) -> None:
        self.scale(0.8, 0.8)

    def zoom_reset(self) -> None:
        self.resetTransform()
        self.fit_page()

    # --- Qt overrides -------------------------------------------------------

    def wheelEvent(self, event) -> None:
        factor = 1.15 if event.angleDelta().y() > 0 else 1.0 / 1.15
        self.scale(factor, factor)

    def mousePressEvent(self, event) -> None:
        if self._add_mode and event.button() == Qt.MouseButton.LeftButton:
            pos = self.mapToScene(event.pos())
            self.callout_place_requested.emit(pos.x(), pos.y())
            return
        super().mousePressEvent(event)
        # Keep keyboard focus on the view after any click so Delete works
        self.setFocus()

    def keyPressEvent(self, event) -> None:
        """Delete or Backspace unchecks any currently selected callout circle."""
        if event.key() in (Qt.Key.Key_Delete, Qt.Key.Key_Backspace):
            for item in self._scene.selectedItems():
                if isinstance(item, CalloutItem):
                    self.callout_delete_requested.emit(item)
            return
        super().keyPressEvent(event)


# ---------------------------------------------------------------------------
# DimensionRow — one item in the right-hand checklist
# ---------------------------------------------------------------------------

class DimensionRow(QWidget):
    check_changed  = pyqtSignal(int, bool)  # (dim_index, is_checked)
    edit_requested = pyqtSignal(int)         # (dim_index)

    def __init__(
        self,
        dim_index: int,
        text: str,
        category: str,           # kept for API compat — no longer displayed
        confidence: float,       # kept for API compat — no longer displayed
        callout_color: str = DEFAULT_CALLOUT_COLOR,
        applied_plus_tol: Optional[float] = None,
        applied_minus_tol: Optional[float] = None,
        tol_inferred: bool = False,
        initial_enabled: bool = True,
        parent: Optional[QWidget] = None,
    ) -> None:
        super().__init__(parent)
        self.dim_index = dim_index
        self._base_bg    = "#F8F8F8" if dim_index % 2 == 0 else "#FFFFFF"
        self._tol_inferred = tol_inferred
        self._focused = False

        self.setMinimumHeight(36)
        self.setFocusPolicy(Qt.FocusPolicy.ClickFocus)
        self.setMaximumHeight(44)

        layout = QHBoxLayout(self)
        layout.setContentsMargins(6, 3, 6, 3)
        layout.setSpacing(6)

        # ── Toggle button (replaces bare QCheckBox) ──────────────────────────
        self.toggle_btn = QPushButton()
        self.toggle_btn.setCheckable(True)
        self.toggle_btn.setChecked(initial_enabled)
        self.toggle_btn.setFixedSize(58, 26)
        self.toggle_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.toggle_btn.toggled.connect(self._on_toggled)
        layout.addWidget(self.toggle_btn)

        # ── Callout number badge ──────────────────────────────────────────────
        self.num_label = QLabel("1")
        self.num_label.setFixedSize(22, 22)
        self.num_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self._badge_color = callout_color
        self._apply_badge_style(callout_color, active=initial_enabled)
        layout.addWidget(self.num_label)

        # ── Nominal text label ────────────────────────────────────────────────
        display = _nominal_text(text)
        if len(display) > 28:
            display = display[:26] + "…"
        self.text_lbl = QLabel(display)
        self.text_lbl.setToolTip(text)
        self.text_lbl.setSizePolicy(
            QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred
        )
        layout.addWidget(self.text_lbl)

        # ── Tolerance label ───────────────────────────────────────────────────
        if applied_plus_tol is not None:
            if applied_minus_tol is not None and applied_minus_tol != applied_plus_tol:
                tol_str = f"+{applied_plus_tol:g}/−{applied_minus_tol:g}"
            else:
                tol_str = f"±{applied_plus_tol:g}"
            if tol_inferred:
                tol_str += "*"          # asterisk = inferred from title block
            self.tol_lbl = QLabel(tol_str)
            self.tol_lbl.setFixedWidth(68)
            self.tol_lbl.setAlignment(
                Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter
            )
            tip = (
                "Inferred from title-block general tolerances"
                if tol_inferred else
                "Explicit tolerance from dimension text"
            )
            self.tol_lbl.setToolTip(tip)
        else:
            self.tol_lbl = QLabel("")
            self.tol_lbl.setFixedWidth(68)
        layout.addWidget(self.tol_lbl)

        # ── Edit button — wider, blue ─────────────────────────────────────────
        self.edit_btn = QPushButton("✎  Edit")
        self.edit_btn.setFixedHeight(24)
        self.edit_btn.setMinimumWidth(64)
        self.edit_btn.setToolTip("Edit this dimension")
        self.edit_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.edit_btn.setStyleSheet(
            "QPushButton{background:#3498DB;color:white;border:none;"
            "border-radius:4px;font-size:11px;font-weight:bold;padding:2px 7px;}"
            "QPushButton:hover{background:#2980B9;}"
            "QPushButton:pressed{background:#1F618D;}"
        )
        self.edit_btn.clicked.connect(lambda: self.edit_requested.emit(self.dim_index))
        layout.addWidget(self.edit_btn)

        # Clicking any label on the row focuses it so Delete/Backspace work
        for _lbl in (self.num_label, self.text_lbl, self.tol_lbl):
            _lbl.installEventFilter(self)

        # Apply initial visual state (background, text colour, toggle label)
        self._refresh_row_style(initial_enabled)

    # --- Public API --------------------------------------------------------

    def set_callout_number(self, n: int) -> None:
        self.num_label.setText(str(n))
        self._apply_badge_style(self._badge_color, active=True)
        # Sync toggle without re-emitting signal
        self.toggle_btn.blockSignals(True)
        self.toggle_btn.setChecked(True)
        self.toggle_btn.blockSignals(False)
        self._refresh_row_style(True)

    def set_color(self, color_hex: str) -> None:
        self._badge_color = color_hex
        self._apply_badge_style(color_hex, active=self.toggle_btn.isChecked())

    def mark_excluded(self) -> None:
        self.num_label.setText("—")
        self._apply_badge_style("#AAAAAA", active=False)
        # Sync toggle without re-emitting signal
        self.toggle_btn.blockSignals(True)
        self.toggle_btn.setChecked(False)
        self.toggle_btn.blockSignals(False)
        self._refresh_row_style(False)

    # --- Private -----------------------------------------------------------

    def _apply_badge_style(self, color: str, active: bool) -> None:
        if active:
            self.num_label.setStyleSheet(
                f"background:{color};color:white;border-radius:11px;"
                f"font-weight:bold;font-size:11px;"
            )
        else:
            self.num_label.setStyleSheet(
                "background:#AAAAAA;color:white;border-radius:11px;"
                "font-weight:bold;font-size:11px;"
            )

    def _refresh_toggle_label(self, enabled: bool) -> None:
        """Update toggle button text and colour to reflect enabled/disabled state."""
        if enabled:
            self.toggle_btn.setText("✓  On")
            self.toggle_btn.setStyleSheet(
                "QPushButton{background:#27AE60;color:white;border:none;"
                "border-radius:5px;font-weight:bold;font-size:11px;}"
                "QPushButton:hover{background:#219A52;}"
                "QPushButton:checked{background:#27AE60;}"
            )
        else:
            self.toggle_btn.setText("✗  Off")
            self.toggle_btn.setStyleSheet(
                "QPushButton{background:#95A5A6;color:white;border:none;"
                "border-radius:5px;font-weight:bold;font-size:11px;}"
                "QPushButton:hover{background:#7F8C8D;}"
                "QPushButton:checked{background:#95A5A6;}"
            )

    def _refresh_row_style(self, enabled: bool) -> None:
        """Dim the entire row when a dimension is toggled off."""
        if enabled:
            bg = "#D6EAF8" if self._focused else self._base_bg
            self.setStyleSheet(f"background:{bg};")
            self.text_lbl.setStyleSheet(
                "background:transparent;font-size:12px;color:#2C3E50;"
            )
            if self._tol_inferred:
                self.tol_lbl.setStyleSheet(
                    "background:transparent;font-size:10px;"
                    "color:#888;font-style:italic;"
                )
            else:
                self.tol_lbl.setStyleSheet(
                    "background:transparent;font-size:10px;"
                    "color:#2C3E50;font-weight:bold;"
                )
        else:
            bg = "#C5D8E0" if self._focused else "#E8E8E8"
            self.setStyleSheet(f"background:{bg};")
            self.text_lbl.setStyleSheet(
                "background:transparent;font-size:12px;color:#AAAAAA;"
            )
            self.tol_lbl.setStyleSheet(
                "background:transparent;font-size:10px;color:#BBBBBB;"
            )
        self._refresh_toggle_label(enabled)

    def _on_toggled(self, checked: bool) -> None:
        self._refresh_row_style(checked)
        self._apply_badge_style(
            self._badge_color if checked else "#AAAAAA", active=checked
        )
        self.check_changed.emit(self.dim_index, checked)

    # --- Keyboard delete support -------------------------------------------

    def mousePressEvent(self, event) -> None:
        """Clicking the row background focuses it for keyboard handling."""
        self.setFocus()
        super().mousePressEvent(event)

    def eventFilter(self, obj, event) -> bool:
        """Forward mouse clicks on child labels so the row gets focus."""
        if event.type() == QEvent.Type.MouseButtonPress:
            self.setFocus()
        return super().eventFilter(obj, event)

    def keyPressEvent(self, event) -> None:
        """Delete or Backspace unchecks (turns off) the focused dimension."""
        if event.key() in (Qt.Key.Key_Delete, Qt.Key.Key_Backspace):
            if self.toggle_btn.isChecked():
                self.toggle_btn.setChecked(False)
            return
        super().keyPressEvent(event)

    def focusInEvent(self, event) -> None:
        self._focused = True
        self._refresh_row_style(self.toggle_btn.isChecked())
        super().focusInEvent(event)

    def focusOutEvent(self, event) -> None:
        self._focused = False
        self._refresh_row_style(self.toggle_btn.isChecked())
        super().focusOutEvent(event)


# ---------------------------------------------------------------------------
# ChecklistPanel — scrollable container for DimensionRow widgets
# ---------------------------------------------------------------------------

class ChecklistPanel(QWidget):
    def __init__(self, parent: Optional[QWidget] = None) -> None:
        super().__init__(parent)
        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(0)

        header_bar = QWidget()
        header_bar.setFixedHeight(36)
        header_bar.setStyleSheet("background:#2C3E50;")
        hbl = QHBoxLayout(header_bar)
        hbl.setContentsMargins(10, 0, 10, 0)
        self.header_lbl = QLabel("Detected Dimensions")
        self.header_lbl.setStyleSheet("color:white;font-weight:bold;font-size:13px;")
        hbl.addWidget(self.header_lbl)
        hbl.addStretch()
        outer.addWidget(header_bar)

        self.scroll = QScrollArea()
        self.scroll.setWidgetResizable(True)
        self.scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.scroll.setStyleSheet("QScrollArea{border:none;}")

        self._container = QWidget()
        self._rows_layout = QVBoxLayout(self._container)
        self._rows_layout.setContentsMargins(0, 0, 0, 0)
        self._rows_layout.setSpacing(1)
        self._rows_layout.addStretch()

        self.scroll.setWidget(self._container)
        outer.addWidget(self.scroll)

    def set_count(self, n: int) -> None:
        self.header_lbl.setText(f"Detected Dimensions  ({n})")

    def clear_rows(self) -> None:
        while self._rows_layout.count() > 1:  # keep the trailing stretch
            item = self._rows_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()
        self.header_lbl.setText("Detected Dimensions")

    def add_row(self, row: DimensionRow) -> None:
        self._rows_layout.insertWidget(self._rows_layout.count() - 1, row)

    def reorder_rows(self, ordered_rows: list) -> None:
        """Rearrange existing row widgets to match ordered_rows (no rebuild).

        Widgets not present in ordered_rows are silently ignored so excluded
        rows (hidden by mark_excluded) end up at the bottom naturally.
        """
        # Detach every row from the layout without deleting it
        while self._rows_layout.count() > 1:          # keep trailing stretch
            self._rows_layout.takeAt(0)
        # Re-insert in the requested order
        for i, row in enumerate(ordered_rows):
            self._rows_layout.insertWidget(i, row)


# ---------------------------------------------------------------------------
# SettingsDialog
# ---------------------------------------------------------------------------

class SettingsDialog(QDialog):
    def __init__(self, config: Config, parent: Optional[QWidget] = None) -> None:
        super().__init__(parent)
        self.config = config
        self.setWindowTitle("Settings")
        self.setMinimumWidth(420)
        self.setModal(True)

        layout = QVBoxLayout(self)
        layout.setSpacing(12)

        grp = QGroupBox("Callout Appearance")
        grp_layout = QFormLayout(grp)

        self._chosen_color = config.callout_color
        self.color_btn = QPushButton()
        self.color_btn.setFixedSize(80, 26)
        self._refresh_color_btn()
        self.color_btn.clicked.connect(self._pick_color)
        grp_layout.addRow("Callout Color:", self.color_btn)

        radius_row = QHBoxLayout()
        self.radius_slider = QSlider(Qt.Orientation.Horizontal)
        self.radius_slider.setRange(8, 26)
        self.radius_slider.setValue(config.callout_radius)
        self.radius_val_lbl = QLabel(str(config.callout_radius))
        self.radius_val_lbl.setFixedWidth(28)
        self.radius_slider.valueChanged.connect(
            lambda v: self.radius_val_lbl.setText(str(v))
        )
        radius_row.addWidget(self.radius_slider)
        radius_row.addWidget(self.radius_val_lbl)
        grp_layout.addRow("Callout Size:", radius_row)
        layout.addWidget(grp)

        btns = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok
            | QDialogButtonBox.StandardButton.Cancel
        )
        btns.accepted.connect(self._accept)
        btns.rejected.connect(self.reject)
        layout.addWidget(btns)

    def _refresh_color_btn(self) -> None:
        self.color_btn.setStyleSheet(
            f"background-color:{self._chosen_color};"
            f"border:1px solid #666;border-radius:3px;"
        )

    def _pick_color(self) -> None:
        color = QColorDialog.getColor(QColor(self._chosen_color), self, "Callout Color")
        if color.isValid():
            self._chosen_color = color.name()
            self._refresh_color_btn()

    def _accept(self) -> None:
        self.config.callout_color = self._chosen_color
        self.config.callout_radius = self.radius_slider.value()
        self.config.save()
        self.accept()


# ---------------------------------------------------------------------------
# LabelDialog — shown when user places a manual callout
# ---------------------------------------------------------------------------

class LabelDialog(QDialog):
    """Three-field callout entry: nominal, + tolerance, − tolerance.

    Tolerances are optional.  If only one is given it is used for both sides
    (symmetric ±).  If both are given and equal they are also written as ±.
    A live preview shows the formatted string that will be stored.
    """

    def __init__(
        self,
        parent: Optional[QWidget] = None,
        *,
        title: str = "Add Callout",
        initial_nominal: str = "",
        initial_plus_tol: str = "",
        initial_minus_tol: str = "",
    ) -> None:
        super().__init__(parent)
        self.setWindowTitle(title)
        self.setMinimumWidth(360)
        self.setModal(True)

        layout = QVBoxLayout(self)
        layout.setSpacing(10)

        # ── Input fields ──────────────────────────────────────────────────
        form = QFormLayout()
        form.setSpacing(8)

        def _field(placeholder: str, initial: str = "") -> QLineEdit:
            e = QLineEdit()
            e.setPlaceholderText(placeholder)
            e.setStyleSheet("color: black; background: white;")
            if initial:
                e.setText(initial)
            return e

        self.nominal_edit = _field("e.g.  25.4",                         initial_nominal)
        self.plus_edit    = _field("e.g.  0.13  (leave blank if none)",   initial_plus_tol)
        self.minus_edit   = _field("e.g.  0.10  (leave blank to match +)", initial_minus_tol)

        form.addRow("Nominal:", self.nominal_edit)
        form.addRow("+ Tolerance:", self.plus_edit)
        form.addRow("− Tolerance:", self.minus_edit)
        layout.addLayout(form)

        # ── Live preview ──────────────────────────────────────────────────
        preview_row = QHBoxLayout()
        preview_row.addWidget(QLabel("Preview:"))
        self.preview_lbl = QLabel("")
        self.preview_lbl.setStyleSheet(
            "font-weight: bold; color: #2C3E50; font-size: 13px;"
        )
        preview_row.addWidget(self.preview_lbl)
        preview_row.addStretch()
        layout.addLayout(preview_row)

        # ── Buttons ───────────────────────────────────────────────────────
        self.btns = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok
            | QDialogButtonBox.StandardButton.Cancel
        )
        self.btns.accepted.connect(self._accept)
        self.btns.rejected.connect(self.reject)
        layout.addWidget(self.btns)

        # Wire up live preview and Tab-to-accept on the last field
        for field in (self.nominal_edit, self.plus_edit, self.minus_edit):
            field.textChanged.connect(self._update_preview)
        self.minus_edit.returnPressed.connect(self.btns.accepted)
        self.nominal_edit.returnPressed.connect(self.plus_edit.setFocus)
        self.plus_edit.returnPressed.connect(self.minus_edit.setFocus)

        self._update_preview()

    # ── Helpers ───────────────────────────────────────────────────────────

    def _build_label(self) -> str:
        """Combine the three fields into a single dimension string."""
        nom   = self.nominal_edit.text().strip()
        plus  = self.plus_edit.text().strip()
        minus = self.minus_edit.text().strip()

        if not nom:
            return ""

        # Symmetrical ± when both sides are the same (or only one is given)
        if plus and minus:
            if plus == minus:
                return f"{nom} ±{plus}"
            return f"{nom} +{plus}/-{minus}"
        if plus:
            return f"{nom} ±{plus}"
        if minus:
            return f"{nom} ±{minus}"
        return nom

    def _update_preview(self) -> None:
        label = self._build_label()
        self.preview_lbl.setText(label if label else "—")
        ok_btn = self.btns.button(QDialogButtonBox.StandardButton.Ok)
        if ok_btn:
            ok_btn.setEnabled(bool(label))

    # ── Public API ────────────────────────────────────────────────────────

    @property
    def label_text(self) -> str:
        return self._build_label()

    def _accept(self) -> None:
        if self.label_text:
            self.accept()
        else:
            self.nominal_edit.setPlaceholderText("Nominal cannot be empty!")


# ---------------------------------------------------------------------------
# MainWindow
# ---------------------------------------------------------------------------

class MainWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("FAI Dimension Numberer")
        self.setMinimumSize(1200, 740)

        self.config = Config()
        self.template_structure: Optional[dict] = None
        self.dimensions: list[dict] = []       # global list across all pages
        self._pages_done: set[int] = set()     # pages already detected
        self.pdf_doc = None                    # fitz.Document
        self.current_page: int = 0
        self.current_pdf_path: str = ""
        self._add_callout_mode: bool = False
        self._worker: Optional[DetectionWorker] = None
        # General tolerances parsed from the title block of page 0.
        # Keys: int (decimal places) → float tolerance, "angular" → float.
        # e.g. {0: 1.0, 1: 0.3, 2: 0.13, "angular": 1.2}
        self.general_tolerances: dict = {}
        # Positions of callout circles that were burned into the PDF when it
        # was last exported.  Used to redact them before re-exporting so that
        # old numbers don't appear behind the new ones.
        self._session_burned_positions: list[dict] = []
        # Undo stack — each entry is a dict describing a reversible action.
        self._undo_stack: list[dict] = []

        self._build_ui()
        self._load_template()

    # =======================================================================
    # UI construction
    # =======================================================================

    def _build_ui(self) -> None:
        central = QWidget()
        self.setCentralWidget(central)
        root = QVBoxLayout(central)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)

        root.addWidget(self._make_toolbar())

        self.splitter = QSplitter(Qt.Orientation.Horizontal)

        left_widget = QWidget()
        left_layout = QVBoxLayout(left_widget)
        left_layout.setContentsMargins(0, 0, 0, 0)
        left_layout.setSpacing(0)

        # Create PDFViewer BEFORE _make_viewer_controls so its methods can be bound
        self.pdf_viewer = PDFViewer()
        self.pdf_viewer.callout_place_requested.connect(self._on_callout_place_requested)
        self.pdf_viewer.callout_delete_requested.connect(self._on_callout_delete_requested)

        left_layout.addWidget(self._make_viewer_controls())
        left_layout.addWidget(self.pdf_viewer)

        self.right_stack = QStackedWidget()
        self.right_stack.setMinimumWidth(320)
        self.right_stack.setMaximumWidth(500)

        # Stack page 0: checklist
        self.checklist = ChecklistPanel()
        self.right_stack.addWidget(self.checklist)

        # Stack page 1: loading state
        loading_page = QWidget()
        loading_page.setStyleSheet("background:#F0F0F0;")
        ll = QVBoxLayout(loading_page)
        ll.setAlignment(Qt.AlignmentFlag.AlignCenter)
        ll.setSpacing(12)
        self.spinner = SpinnerWidget()
        self.spinner.setFixedSize(64, 64)
        ll.addWidget(self.spinner, alignment=Qt.AlignmentFlag.AlignHCenter)
        loading_lbl = QLabel("Analyzing drawing…")
        loading_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        loading_lbl.setStyleSheet("font-size:13px;color:#444;")
        ll.addWidget(loading_lbl)
        loading_sub = QLabel("Detecting dimensions via pattern matching.")
        loading_sub.setAlignment(Qt.AlignmentFlag.AlignCenter)
        loading_sub.setStyleSheet("font-size:11px;color:#888;")
        ll.addWidget(loading_sub)
        self.right_stack.addWidget(loading_page)

        self.splitter.addWidget(left_widget)
        self.splitter.addWidget(self.right_stack)
        self.splitter.setSizes([820, 380])
        root.addWidget(self.splitter)

        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        self.status_bar.showMessage("Ready — open a drawing to begin.")

    def _make_toolbar(self) -> QWidget:
        bar = QWidget()
        bar.setFixedHeight(52)
        bar.setStyleSheet(
            "background:#1E2132;"
            "border-bottom:1px solid #3A3A5C;"
        )
        layout = QHBoxLayout(bar)
        layout.setContentsMargins(10, 6, 10, 6)
        layout.setSpacing(6)

        def btn(text: str, tip: str = "", checkable: bool = False) -> QPushButton:
            b = QPushButton(text)
            b.setToolTip(tip)
            b.setCheckable(checkable)
            b.setFixedHeight(34)
            b.setStyleSheet(
                "QPushButton{"
                "background:#3A3D5C;color:white;border:none;"
                "border-radius:5px;padding:0 14px;font-size:12px;}"
                "QPushButton:hover{background:#5A5D8C;}"
                "QPushButton:pressed{background:#2A2D4C;}"
                "QPushButton:checked{background:#C0392B;}"
                "QPushButton:disabled{background:#2A2A3A;color:#555;}"
            )
            return b

        def sep() -> QFrame:
            f = QFrame()
            f.setFrameShape(QFrame.Shape.VLine)
            f.setStyleSheet("color:#444;")
            return f

        self.open_btn = btn("📂  Open Drawing", "Open a PDF engineering drawing (Ctrl+O)")
        self.open_btn.clicked.connect(self.open_drawing)

        self.add_callout_btn = btn(
            "✚  Add Callout",
            "Click on drawing to place a manual callout (Esc to cancel)",
            checkable=True,
        )
        self.add_callout_btn.toggled.connect(self._toggle_add_callout_mode)

        self.export_pdf_btn = btn("📄  Export PDF", "Save annotated PDF")
        self.export_pdf_btn.clicked.connect(self.export_pdf)
        self.export_pdf_btn.setEnabled(False)

        self.export_record_btn = btn("📊  Export Measurement Record", "Save filled copy of measurement template")
        self.export_record_btn.clicked.connect(self.export_measurement_record)
        self.export_record_btn.setEnabled(False)

        settings_btn = btn("⚙", "Settings")
        settings_btn.setFixedWidth(40)
        settings_btn.clicked.connect(self._open_settings)

        layout.addWidget(self.open_btn)
        layout.addWidget(sep())
        layout.addWidget(self.add_callout_btn)
        layout.addWidget(sep())
        layout.addWidget(self.export_pdf_btn)
        layout.addWidget(self.export_record_btn)
        layout.addStretch()
        layout.addWidget(sep())
        layout.addWidget(settings_btn)

        return bar

    def _make_viewer_controls(self) -> QWidget:
        bar = QWidget()
        bar.setFixedHeight(36)
        bar.setStyleSheet("background:#F4F4F4;border-bottom:1px solid #DDD;")
        layout = QHBoxLayout(bar)
        layout.setContentsMargins(6, 3, 6, 3)
        layout.setSpacing(4)

        def small_btn(text: str, tip: str = "") -> QPushButton:
            b = QPushButton(text)
            b.setToolTip(tip)
            b.setFixedSize(28, 26)
            b.setStyleSheet(
                "QPushButton{background:#DDD;border:1px solid #BBB;"
                "border-radius:3px;font-size:13px;}"
                "QPushButton:hover{background:#CCC;}"
            )
            return b

        zi = small_btn("+", "Zoom in")
        zi.clicked.connect(self.pdf_viewer.zoom_in)
        zo = small_btn("−", "Zoom out")
        zo.clicked.connect(self.pdf_viewer.zoom_out)
        fit = QPushButton("Fit")
        fit.setFixedHeight(26)
        fit.setToolTip("Fit page to window")
        fit.setStyleSheet(
            "QPushButton{background:#DDD;border:1px solid #BBB;"
            "border-radius:3px;font-size:11px;padding:0 8px;}"
            "QPushButton:hover{background:#CCC;}"
        )
        fit.clicked.connect(self.pdf_viewer.zoom_reset)

        layout.addWidget(zi)
        layout.addWidget(zo)
        layout.addWidget(fit)
        layout.addStretch()

        # Page navigation — Prev / label / Next
        self.prev_btn = small_btn("◀", "Previous page")
        self.prev_btn.clicked.connect(self._go_prev_page)
        self.prev_btn.setEnabled(False)

        self.page_lbl = QLabel("Page 1 / 1")
        self.page_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.page_lbl.setStyleSheet("color:#444; font-size:12px;")
        self.page_lbl.setFixedWidth(90)

        self.next_btn = small_btn("▶", "Next page")
        self.next_btn.clicked.connect(self._go_next_page)
        self.next_btn.setEnabled(False)

        layout.addWidget(self.prev_btn)
        layout.addWidget(self.page_lbl)
        layout.addWidget(self.next_btn)

        return bar

    # =======================================================================
    # Template loading
    # =======================================================================

    def _load_template(self) -> None:
        template_path = os.path.join(_APP_DIR, TEMPLATE_FILENAME)
        if not os.path.isfile(template_path):
            QMessageBox.critical(
                self,
                "Measurement Template Not Found",
                f"<b>{TEMPLATE_FILENAME}</b> was not found in the application folder.<br><br>"
                f"Please place this file here and relaunch:<br>"
                f"<tt>{_APP_DIR}</tt><br><br>"
                f"The application will now exit.",
            )
            QTimer.singleShot(0, QApplication.instance().quit)
            return

        if parse_spreadsheet_template is None:
            return

        try:
            self.template_structure = parse_spreadsheet_template(template_path)
            n_sheets = len(self.template_structure.get("sheet_names", []))
            self.status_bar.showMessage(
                f"Template loaded: {TEMPLATE_FILENAME}  ({n_sheets} sheet{'s' if n_sheets != 1 else ''})"
            )
        except Exception as exc:
            QMessageBox.warning(
                self,
                "Template Warning",
                f"Could not fully parse <b>{TEMPLATE_FILENAME}</b>:<br>{exc}<br><br>"
                f"Measurement record export will be limited.",
            )

    # =======================================================================
    # Opening a drawing
    # =======================================================================

    def open_drawing(self) -> None:
        path, _ = QFileDialog.getOpenFileName(
            self, "Open Engineering Drawing", "", "PDF Files (*.pdf)"
        )
        if not path:
            return

        try:
            if self.pdf_doc is not None:
                self.pdf_doc.close()
            self.pdf_doc = fitz.open(path)
        except Exception as exc:
            QMessageBox.critical(self, "Cannot Open PDF", f"Failed to open the drawing:\n\n{exc}")
            return

        self.current_pdf_path = path
        self.dimensions = []
        self._pages_done = set()
        self._session_burned_positions = []
        self._undo_stack = []

        # ── Session restore ──────────────────────────────────────────────────
        # Check whether a saved .autodim sidecar exists for this PDF.
        # If so, offer to restore the previous session instead of re-detecting.
        session = self._load_session_sidecar(path)
        if session:
            reply = QMessageBox.question(
                self,
                "Previous Session Found",
                "A saved session file was found for this drawing.\n\n"
                "Restore the previous callouts?\n\n"
                "• Yes — load the saved callout positions and skip re-detection.\n"
                "  When you re-export, old burned-in numbers will be removed\n"
                "  automatically before the updated ones are drawn.\n\n"
                "• No — run dimension detection from scratch (session ignored).",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.Yes,
            )
            if reply == QMessageBox.StandardButton.Yes:
                self._restore_from_session(session)
                return

        # ── Normal flow — detect from scratch ───────────────────────────────
        self.general_tolerances = self._detect_general_tolerances()
        if self.general_tolerances:
            parts: list[str] = []
            for k in sorted(k for k in self.general_tolerances if isinstance(k, int)):
                x_str = "X" if k == 0 else "." + "X" * k
                parts.append(f"{x_str} ±{self.general_tolerances[k]}")
            if "angular" in self.general_tolerances:
                parts.append(f"∠ ±{self.general_tolerances['angular']}°")
            self.status_bar.showMessage(
                "General tolerances detected: " + ",  ".join(parts)
            )

        n_pages = len(self.pdf_doc)
        self._update_page_nav(0, n_pages)
        self._go_to_page(0)

        self.export_pdf_btn.setEnabled(False)
        self.export_record_btn.setEnabled(False)

    def _update_page_nav(self, current: int, total: int) -> None:
        self.page_lbl.setText(f"Page {current + 1} / {total}")
        self.prev_btn.setEnabled(current > 0)
        self.next_btn.setEnabled(current < total - 1)

    def _go_prev_page(self) -> None:
        if self.pdf_doc and self.current_page > 0:
            self._go_to_page(self.current_page - 1)

    def _go_next_page(self) -> None:
        if self.pdf_doc and self.current_page < len(self.pdf_doc) - 1:
            self._go_to_page(self.current_page + 1)

    def _go_to_page(self, new_page: int) -> None:
        if self.pdf_doc is None:
            return
        self.current_page = new_page
        self._update_page_nav(new_page, len(self.pdf_doc))
        self._render_page()

        if new_page in self._pages_done:
            self._show_current_page_callouts()
            self._rebuild_checklist_for_page(new_page)
            self.right_stack.setCurrentIndex(0)
        else:
            self._run_detection()

    # =======================================================================
    # General-tolerance detection (title block, page 0)
    # =======================================================================

    def _detect_general_tolerances(self) -> dict:
        """Scan page 0 for a general-tolerance block.

        Tries the bottom-right corner first (where title blocks live on
        virtually every drawing), then falls back to the full page so the
        detection works even for non-standard layouts.
        """
        if self.pdf_doc is None or len(self.pdf_doc) == 0:
            return {}

        page = self.pdf_doc[0]
        pr = page.rect

        # Progressively larger crop windows anchored to the bottom-right,
        # then a full-page fallback.
        regions: list[Optional[fitz.Rect]] = [
            fitz.Rect(pr.width * 0.55, pr.height * 0.70, pr.width, pr.height),
            fitz.Rect(pr.width * 0.45, pr.height * 0.60, pr.width, pr.height),
            fitz.Rect(pr.width * 0.30, pr.height * 0.50, pr.width, pr.height),
            None,   # None → full page
        ]

        for clip in regions:
            text = page.get_text("text", clip=clip) if clip else page.get_text("text")
            tols = _parse_title_block_tolerances(text)
            if tols:
                label = (
                    f"bottom-right clip ({clip.x0:.0f},{clip.y0:.0f})"
                    if clip else "full page"
                )
                print(f"[tolerance-detect] found via {label}: {tols}")
                return tols

        # Nothing found — dump the raw bottom-right text so it can be debugged
        raw = page.get_text(
            "text",
            clip=fitz.Rect(pr.width * 0.30, pr.height * 0.50, pr.width, pr.height),
        )
        print(
            "[tolerance-detect] no tolerances found.\n"
            "--- bottom-right text ---\n"
            + raw[:800]
        )
        return {}

    def _apply_general_tolerances(self, dims: list) -> None:
        """For dims that carry no explicit tolerance, apply the matching
        general tolerance from the title block by decimal-place count.

        Stores results in ``dim["applied_plus_tol"]`` and
        ``dim["applied_minus_tol"]`` (both positive floats).  Dims that
        already have explicit tolerances parsed from their text are still
        populated so that the export path has a single field to consult.
        """
        if not self.general_tolerances:
            print("[tolerance-apply] skipped — no general_tolerances loaded")
            return

        print(f"[tolerance-apply] applying to {len(dims)} dim(s) using {self.general_tolerances}")
        for dim in dims:
            _nom, plus_t, minus_t = _parse_tolerances(dim["text"])

            if plus_t is not None:
                # Explicit tolerance in the text — store it and move on
                dim["applied_plus_tol"]  = plus_t
                dim["applied_minus_tol"] = minus_t if minus_t is not None else plus_t
                print(f"  explicit  {dim['text']!r:30} -> ±{plus_t}")
                continue

            # No explicit tolerance — look up by type / decimal count
            cat = dim.get("category", "")
            if cat == "angular":
                tol = self.general_tolerances.get("angular")
            else:
                decimals = _count_decimal_places(dim["text"])
                tol = self.general_tolerances.get(decimals)
                if tol is None:
                    # Fall back to the next-coarser entry available
                    for d in range(decimals - 1, -1, -1):
                        if d in self.general_tolerances:
                            tol = self.general_tolerances[d]
                            break

            if tol is not None:
                dim["applied_plus_tol"]  = tol
                dim["applied_minus_tol"] = tol
                print(f"  general   {dim['text']!r:30} -> ±{tol}")
            else:
                print(f"  NO MATCH  {dim['text']!r:30} (cat={cat!r}, decimals={_count_decimal_places(dim['text'])})")

    # =======================================================================
    # Session sidecar — save / restore previously annotated drawings
    # =======================================================================

    @staticmethod
    def _sidecar_path(pdf_path: str) -> str:
        """Return the .autodim sidecar path for a given PDF path."""
        return os.path.splitext(pdf_path)[0] + SIDECAR_EXT

    def _load_session_sidecar(self, pdf_path: str) -> Optional[dict]:
        """Return parsed sidecar JSON if one exists next to pdf_path, else None."""
        sp = self._sidecar_path(pdf_path)
        if not os.path.isfile(sp):
            return None
        try:
            with open(sp, "r", encoding="utf-8") as fh:
                return json.load(fh)
        except Exception:
            return None

    def _save_session_sidecar(self, pdf_path: str) -> None:
        """Persist all dimension data to a .autodim sidecar next to pdf_path."""
        # JSON requires string keys; convert int tolerance-level keys → str.
        tols_serialisable = {str(k): v for k, v in self.general_tolerances.items()}

        session: dict[str, Any] = {
            "version": "1.0",
            "callout_color": self.config.callout_color,
            "callout_radius": self.config.callout_radius,
            "general_tolerances": tols_serialisable,
            "dimensions": [
                {
                    "text":             d["text"],
                    "page":             d.get("page", 0),
                    "x":                d.get("x", 0.0),
                    "y":                d.get("y", 0.0),
                    "width":            d.get("width", 0.0),
                    "height":           d.get("height", 0.0),
                    "callout_x":        d["callout_x"],
                    "callout_y":        d["callout_y"],
                    "callout_num":      d.get("callout_num", 0),
                    "included":         d.get("included", True),
                    "category":         d.get("category", "other"),
                    "confidence":       d.get("confidence", 1.0),
                    "applied_plus_tol": d.get("applied_plus_tol"),
                    "applied_minus_tol":d.get("applied_minus_tol"),
                }
                for d in self.dimensions
            ],
        }
        sp = self._sidecar_path(pdf_path)
        try:
            with open(sp, "w", encoding="utf-8") as fh:
                json.dump(session, fh, indent=2)
        except Exception as exc:
            print(f"[session] could not save sidecar: {exc}")

    def _restore_from_session(self, session: dict) -> None:
        """Populate self.dimensions from a loaded sidecar and skip detection."""
        # Restore appearance config
        if "callout_color" in session:
            self.config.callout_color = session["callout_color"]
        if "callout_radius" in session:
            self.config.callout_radius = int(session["callout_radius"])

        # Restore general tolerances — JSON turns int keys into strings
        raw = session.get("general_tolerances", {})
        self.general_tolerances = {
            (int(k) if k.lstrip("-").isdigit() else k): v
            for k, v in raw.items()
        }

        # Rebuild dimensions list (no callout_item / row_widget yet)
        self.dimensions = []
        for d in session.get("dimensions", []):
            entry = {k: v for k, v in d.items()}
            entry["callout_item"] = None
            entry["row_widget"]   = None
            self.dimensions.append(entry)

        # Record which circles were burned into this PDF so _write_annotated_pdf
        # can redact them before drawing the updated ones.
        radius_pt = self.config.callout_radius / RENDER_ZOOM
        self._session_burned_positions = [
            {
                "page":   d.get("page", 0),
                "x":      d["callout_x"],
                "y":      d["callout_y"],
                "r_pt":   radius_pt,
            }
            for d in session.get("dimensions", [])
            if d.get("included", True) and d.get("callout_num", 0) > 0
        ]

        # ── Erase burned-in circles from the in-memory document ──────────────
        # Apply the same redactions to self.pdf_doc RIGHT NOW so the viewer
        # renders clean pages immediately.  The on-disk file is NOT modified;
        # _write_annotated_pdf opens it fresh and redacts independently.
        if self._session_burned_positions:
            pages_touched: set[int] = set()
            for bp in self._session_burned_positions:
                pg = bp["page"]
                if pg >= len(self.pdf_doc):
                    continue
                r = bp["r_pt"]
                cx, cy = bp["x"], bp["y"]
                try:
                    self.pdf_doc[pg].add_redact_annot(
                        fitz.Rect(cx - r, cy - r, cx + r, cy + r),
                        fill=(1, 1, 1),
                    )
                    pages_touched.add(pg)
                except Exception:
                    pass
            for pg in pages_touched:
                try:
                    self.pdf_doc[pg].apply_redactions()
                except Exception:
                    pass

        # Mark every page as already processed — no detection will run
        n_pages = len(self.pdf_doc)
        self._pages_done = set(range(n_pages))

        self._update_page_nav(0, n_pages)
        self._go_to_page(0)          # renders the now-clean page + shows callouts

        self.export_pdf_btn.setEnabled(True)
        self.export_record_btn.setEnabled(self.template_structure is not None)

        n_active = sum(1 for d in self.dimensions if d.get("included", True))
        n_total  = len(self.dimensions)
        self.status_bar.showMessage(
            f"Session restored — {n_active} active callout{'s' if n_active != 1 else ''} "
            f"({n_total} total).  Move or toggle callouts, then re-export to apply changes."
        )

    def _render_page(self) -> None:
        if self.pdf_doc is None:
            return
        # Remove callout items from scene BEFORE scene.clear() inside display_pixmap.
        # This prevents Qt from deleting C++ objects we still hold references to.
        self.pdf_viewer.clear_callout_items()
        for dim in self.dimensions:
            dim["callout_item"] = None

        page = self.pdf_doc[self.current_page]

        # Render at native physical-pixel resolution so the image is sharp on
        # HiDPI / scaled displays (e.g. Windows 125 %, 150 %, 200 % DPI).
        # dpr == 1.0 on a standard display → no change in behaviour.
        dpr: float = self.devicePixelRatio()
        mat = fitz.Matrix(RENDER_ZOOM * dpr, RENDER_ZOOM * dpr)
        pix = page.get_pixmap(matrix=mat)
        img = QImage(
            pix.samples, pix.width, pix.height,
            pix.stride, QImage.Format.Format_RGB888,
        )
        pixmap = QPixmap.fromImage(img)
        # Tell Qt the logical size is still RENDER_ZOOM × PDF-pts; coordinate
        # conversions that rely on RENDER_ZOOM are unaffected.
        pixmap.setDevicePixelRatio(dpr)
        self.pdf_viewer.display_pixmap(pixmap)
        self.pdf_viewer.zoom_reset()

    # =======================================================================
    # Dimension detection
    # =======================================================================

    def _run_detection(self) -> None:
        self.right_stack.setCurrentIndex(1)
        self.spinner.start()
        self.status_bar.showMessage(
            f"Analyzing page {self.current_page + 1} of '{os.path.basename(self.current_pdf_path)}'…"
        )

        self._worker = DetectionWorker(self.current_pdf_path, self.current_page)
        self._worker.finished.connect(self._on_detection_done)
        self._worker.error.connect(self._on_detection_error)
        self._worker.start()

    @pyqtSlot(list)
    def _on_detection_done(self, dims: list) -> None:
        self.spinner.stop()

        page_num = self.current_page
        r_pt = self.config.callout_radius / RENDER_ZOOM

        # Build list of already-placed callout centres for this page
        placed: list[tuple[float, float]] = [
            (d["callout_x"], d["callout_y"])
            for d in self.dimensions
            if d.get("page") == page_num and d.get("included", True)
        ]

        for raw in dims:
            cx, cy = _smart_callout_pos(raw, placed, r_pt)
            placed.append((cx, cy))

            entry: dict[str, Any] = dict(raw)
            entry.update({
                "page":       page_num,
                "callout_x":  cx,
                "callout_y":  cy,
                "included":   True,
                "callout_num": 0,
                "callout_item": None,
                "row_widget": None,
            })
            self.dimensions.append(entry)

        self._pages_done.add(page_num)
        # Fill in general tolerances for dims that have none explicitly
        self._apply_general_tolerances(
            [d for d in self.dimensions if d.get("page") == page_num]
        )
        self._update_callout_numbers()
        self._show_current_page_callouts()
        self._rebuild_checklist_for_page(page_num)

        self.right_stack.setCurrentIndex(0)
        self.export_pdf_btn.setEnabled(True)
        self.export_record_btn.setEnabled(self.template_structure is not None)

        n_page = sum(1 for d in self.dimensions if d.get("page") == page_num)
        n_total = len(self.dimensions)
        self.status_bar.showMessage(
            f"Page {page_num + 1}: {n_page} dimension{'s' if n_page != 1 else ''} found  "
            f"({n_total} total across {len(self._pages_done)} page{'s' if len(self._pages_done) != 1 else ''})."
        )

    @pyqtSlot(str)
    def _on_detection_error(self, msg: str) -> None:
        self.spinner.stop()
        self.right_stack.setCurrentIndex(0)
        QMessageBox.critical(
            self,
            "Detection Error",
            f"Dimension detection failed:\n\n{msg}\n\n"
            "If the PDF is scanned, run OCR on it first (e.g. ocrmypdf).",
        )
        self.status_bar.showMessage("Detection failed — see error dialog.")

    # =======================================================================
    # Checklist & callout item management
    # =======================================================================

    def _rebuild_checklist_for_page(self, page_num: int) -> None:
        """Rebuild the right-hand checklist showing only dims for page_num."""
        # Clear row_widget references for all dims (old rows are about to be deleted)
        for dim in self.dimensions:
            dim["row_widget"] = None

        self.checklist.clear_rows()

        page_dims = [d for d in self.dimensions if d.get("page") == page_num]
        for dim in page_dims:
            # Determine whether the tolerance is explicit (in the text) or
            # inferred from the title-block general tolerances.
            _, explicit_plus, _ = _parse_tolerances(dim["text"])
            tol_inferred = (
                explicit_plus is None
                and dim.get("applied_plus_tol") is not None
            )
            row = DimensionRow(
                dim_index=self.dimensions.index(dim),
                text=dim["text"],
                category=dim["category"],
                confidence=dim["confidence"],
                callout_color=self.config.callout_color,
                applied_plus_tol=dim.get("applied_plus_tol"),
                applied_minus_tol=dim.get("applied_minus_tol"),
                tol_inferred=tol_inferred,
                initial_enabled=dim.get("included", True),
            )
            row.check_changed.connect(self._on_check_changed)
            row.edit_requested.connect(self._on_edit_requested)
            dim["row_widget"] = row
            self.checklist.add_row(row)

        self.checklist.set_count(len(page_dims))
        # Sync badge numbers on the newly created rows
        self._update_callout_numbers()

    def _show_current_page_callouts(self) -> None:
        """Recreate callout items in the scene for the current page's dims."""
        # Items were already cleared from scene in _render_page; callout_item = None
        for dim in self.dimensions:
            if dim.get("page") != self.current_page:
                continue
            item = self._make_callout_item(dim)
            dim["callout_item"] = item
            self.pdf_viewer.add_callout_item(item)
        self._update_callout_numbers()

    def _make_callout_item(self, dim: dict) -> CalloutItem:
        """Create and position a CalloutItem for a dimension dict."""
        def on_moved(item: CalloutItem, pos: QPointF, d: dict = dim) -> None:
            d["callout_x"] = pos.x() / RENDER_ZOOM
            d["callout_y"] = pos.y() / RENDER_ZOOM

        def on_drag_complete(
            item: CalloutItem,
            old_scene: QPointF,
            new_scene: QPointF,
            d: dict = dim,
        ) -> None:
            self._push_undo({
                "type":      "move",
                "dim_index": self.dimensions.index(d),
                "old_x":     old_scene.x() / RENDER_ZOOM,
                "old_y":     old_scene.y() / RENDER_ZOOM,
                "new_x":     new_scene.x() / RENDER_ZOOM,
                "new_y":     new_scene.y() / RENDER_ZOOM,
            })
            # Renumber in new reading order now that the position has changed
            self._update_callout_numbers()

        item = CalloutItem(
            number=1,
            radius=self.config.callout_radius,
            color=self.config.callout_color,
            on_moved=on_moved,
            on_drag_complete=on_drag_complete,
            on_context_menu=self._on_callout_context_menu,
        )
        item.setPos(dim["callout_x"] * RENDER_ZOOM, dim["callout_y"] * RENDER_ZOOM)
        return item

    def _update_callout_numbers(self) -> None:
        """Reassign sequential callout numbers in reading order.

        Included callouts are sorted page-first, then top-to-bottom in
        bands (callouts within one circle-diameter vertically are treated as
        the same 'row'), then left-to-right within each band.  This means
        inserting or moving a callout anywhere on the page automatically
        renumbers everything else to maintain correct reading order.
        """
        r_pt  = self.config.callout_radius / RENDER_ZOOM
        band  = max(r_pt * 2.0, 15.0)   # vertical tolerance for same-row grouping

        def _reading_key(d: dict) -> tuple:
            return (
                d.get("page", 0),
                round(d.get("callout_y", 0.0) / band),  # row band
                d.get("callout_x", 0.0),                # left→right within row
            )

        # Assign numbers to included dims in reading order
        included_sorted = sorted(
            (d for d in self.dimensions if d.get("included", True)),
            key=_reading_key,
        )
        for num, dim in enumerate(included_sorted, start=1):
            dim["callout_num"] = num
            row: Optional[DimensionRow] = dim.get("row_widget")
            item: Optional[CalloutItem] = dim.get("callout_item")
            if row is not None:
                row.set_callout_number(num)
                row.set_color(self.config.callout_color)
            if item is not None:
                item.set_number(num)
                item.set_color(self.config.callout_color)
                item.setVisible(True)

        # Zero-out and hide excluded dims
        for dim in self.dimensions:
            if dim.get("included", True):
                continue
            dim["callout_num"] = 0
            row = dim.get("row_widget")
            item = dim.get("callout_item")
            if row is not None:
                row.mark_excluded()
            if item is not None:
                item.setVisible(False)

        # Reorder the checklist panel so rows appear numerically 1 → N,
        # with any excluded rows (callout_num == 0) pushed to the bottom.
        visible_rows = [
            (dim["callout_num"], dim["row_widget"])
            for dim in self.dimensions
            if dim.get("row_widget") is not None
        ]
        visible_rows.sort(key=lambda pair: pair[0] if pair[0] > 0 else float("inf"))
        ordered = [row for _, row in visible_rows]
        if ordered:
            self.checklist.reorder_rows(ordered)

    @pyqtSlot(int, bool)
    def _on_check_changed(self, dim_index: int, checked: bool) -> None:
        if 0 <= dim_index < len(self.dimensions):
            old_state = self.dimensions[dim_index].get("included", True)
            self.dimensions[dim_index]["included"] = checked
            self._push_undo({
                "type":      "toggle",
                "dim_index": dim_index,
                "old_state": old_state,
                "new_state": checked,
            })
            self._update_callout_numbers()

    @pyqtSlot(int)
    def _on_edit_requested(self, dim_index: int) -> None:
        """Open a pre-filled LabelDialog so the user can correct a dimension."""
        if not (0 <= dim_index < len(self.dimensions)):
            return

        dim = self.dimensions[dim_index]

        # Pre-populate from the current text
        nom_str = _nominal_text(dim["text"])
        _, plus_t, minus_t = _parse_tolerances(dim["text"])
        plus_str  = f"{plus_t:g}"  if plus_t  is not None else ""
        minus_str = f"{minus_t:g}" if minus_t is not None else ""

        dlg = LabelDialog(
            self,
            title="Edit Dimension",
            initial_nominal=nom_str,
            initial_plus_tol=plus_str,
            initial_minus_tol=minus_str,
        )
        if dlg.exec() != QDialog.DialogCode.Accepted:
            return

        # Update the dimension text and re-derive applied tolerances
        dim["text"] = dlg.label_text
        self._apply_general_tolerances([dim])

        # Rebuild the checklist so the row reflects the new values
        self._rebuild_checklist_for_page(self.current_page)
        self.status_bar.showMessage(
            f"Dimension #{dim_index + 1} updated to: {dim['text']}"
        )

    # =======================================================================
    # Keyboard shortcuts
    # =======================================================================

    def keyPressEvent(self, event) -> None:
        if event.key() == Qt.Key.Key_Escape and self._add_callout_mode:
            self.add_callout_btn.setChecked(False)
        elif (
            event.key() == Qt.Key.Key_Z
            and event.modifiers() & Qt.KeyboardModifier.ControlModifier
        ):
            self._undo_last()
        super().keyPressEvent(event)

    # =======================================================================
    # Undo
    # =======================================================================

    _MAX_UNDO = 50

    def _push_undo(self, action: dict) -> None:
        self._undo_stack.append(action)
        if len(self._undo_stack) > self._MAX_UNDO:
            self._undo_stack.pop(0)

    def _undo_last(self) -> None:
        if not self._undo_stack:
            self.status_bar.showMessage("Nothing to undo.")
            return
        action = self._undo_stack.pop()
        if action["type"] == "move":
            self._undo_move(action)
        elif action["type"] == "toggle":
            self._undo_toggle(action)

    def _undo_move(self, action: dict) -> None:
        idx = action["dim_index"]
        if not (0 <= idx < len(self.dimensions)):
            return
        dim = self.dimensions[idx]
        old_x, old_y = action["old_x"], action["old_y"]
        item: Optional[CalloutItem] = dim.get("callout_item")
        if item is not None:
            # setPos triggers on_moved which updates dim["callout_x/y"] for us
            item.setPos(old_x * RENDER_ZOOM, old_y * RENDER_ZOOM)
        else:
            dim["callout_x"] = old_x
            dim["callout_y"] = old_y
        # Renumber in reading order after the position is restored
        self._update_callout_numbers()
        self.status_bar.showMessage(
            f"Undone: callout move  (now #{dim.get('callout_num', '?')})."
        )

    def _undo_toggle(self, action: dict) -> None:
        idx = action["dim_index"]
        if not (0 <= idx < len(self.dimensions)):
            return
        old_state: bool = action["old_state"]
        dim = self.dimensions[idx]
        row: Optional[DimensionRow] = dim.get("row_widget")
        if row is not None:
            # blockSignals prevents _on_check_changed from pushing another undo entry
            row.toggle_btn.blockSignals(True)
            row.toggle_btn.setChecked(old_state)
            row.toggle_btn.blockSignals(False)
            row._refresh_row_style(old_state)
            row._apply_badge_style(
                row._badge_color if old_state else "#AAAAAA", active=old_state
            )
        dim["included"] = old_state
        self._update_callout_numbers()
        self.status_bar.showMessage(
            f"Undone: toggle for dimension #{idx + 1} "
            f"→ {'On' if old_state else 'Off'}."
        )

    # =======================================================================
    # Add-callout mode
    # =======================================================================

    def _toggle_add_callout_mode(self, checked: bool) -> None:
        self._add_callout_mode = checked
        self.pdf_viewer.set_add_mode(checked)
        if checked:
            self.status_bar.showMessage(
                "Add Callout mode active — click on the drawing to place a callout.  Press Esc to cancel."
            )
        else:
            self.status_bar.showMessage("Normal mode.")

    @pyqtSlot(float, float)
    def _on_callout_place_requested(self, scene_x: float, scene_y: float) -> None:
        # Exit add mode immediately so the user isn't stuck in crosshair mode
        self.add_callout_btn.setChecked(False)

        dlg = LabelDialog(self)
        if dlg.exec() != QDialog.DialogCode.Accepted:
            return
        label = dlg.label_text
        if not label:
            return

        pdf_x = scene_x / RENDER_ZOOM
        pdf_y = scene_y / RENDER_ZOOM
        idx = len(self.dimensions)

        dim: dict[str, Any] = {
            "index":       idx,
            "text":        label,
            "x":           pdf_x,
            "y":           pdf_y,
            "width":       0.0,
            "height":      0.0,
            "confidence":  1.0,
            "category":    "other",
            "page":        self.current_page,
            "included":    True,
            "callout_x":   pdf_x,
            "callout_y":   pdf_y,
            "callout_num": 0,
            "callout_item": None,
            "row_widget":   None,
        }
        self.dimensions.append(dim)
        self._apply_general_tolerances([dim])   # fill tolerances from title block

        _, explicit_plus, _ = _parse_tolerances(label)
        tol_inferred = (
            explicit_plus is None and dim.get("applied_plus_tol") is not None
        )
        row = DimensionRow(
            dim_index=idx,
            text=label,
            category="other",
            confidence=1.0,
            callout_color=self.config.callout_color,
            applied_plus_tol=dim.get("applied_plus_tol"),
            applied_minus_tol=dim.get("applied_minus_tol"),
            tol_inferred=tol_inferred,
            initial_enabled=True,
        )
        row.check_changed.connect(self._on_check_changed)
        row.edit_requested.connect(self._on_edit_requested)
        dim["row_widget"] = row
        self.checklist.add_row(row)
        self.checklist.set_count(
            sum(1 for d in self.dimensions if d.get("page") == self.current_page)
        )

        item = self._make_callout_item(dim)
        dim["callout_item"] = item
        self.pdf_viewer.add_callout_item(item)

        self._update_callout_numbers()
        self.export_pdf_btn.setEnabled(True)
        if self.template_structure is not None:
            self.export_record_btn.setEnabled(True)

    @pyqtSlot(object)
    def _on_callout_delete_requested(self, item: "CalloutItem") -> None:
        """Delete/Backspace pressed while a callout circle is selected — uncheck it."""
        for dim in self.dimensions:
            if dim.get("callout_item") is item:
                row: Optional[DimensionRow] = dim.get("row_widget")
                if row is not None:
                    # Let the toggle drive everything (visual update + signal chain)
                    if row.toggle_btn.isChecked():
                        row.toggle_btn.setChecked(False)
                else:
                    # Dimension is on a different page — update state directly
                    dim["included"] = False
                    self._update_callout_numbers()
                return

    def _on_callout_context_menu(self, item: "CalloutItem", screen_pos) -> None:
        """Show a right-click context menu on a callout circle."""
        # Find which dimension this item belongs to
        dim_index: Optional[int] = None
        for i, dim in enumerate(self.dimensions):
            if dim.get("callout_item") is item:
                dim_index = i
                break
        if dim_index is None:
            return

        menu = QMenu(self)
        menu.setStyleSheet("""
            QMenu {
                background-color: #2B2D42;
                border: 1px solid #555;
                border-radius: 6px;
                padding: 4px 0px;
                color: #EEEEEE;
                font-size: 13px;
            }
            QMenu::item {
                padding: 7px 20px 7px 12px;
                border-radius: 4px;
                margin: 2px 4px;
            }
            QMenu::item:selected {
                background-color: #4A90D9;
                color: white;
            }
        """)
        edit_action   = menu.addAction("✏   Edit Dimension")
        menu.addSeparator()
        delete_action = menu.addAction("🗑   Delete Dimension")

        chosen = menu.exec(screen_pos)
        if chosen is edit_action:
            self._on_edit_requested(dim_index)
        elif chosen is delete_action:
            self._on_callout_delete_requested(item)

    # =======================================================================
    # Settings
    # =======================================================================

    def _open_settings(self) -> None:
        dlg = SettingsDialog(self.config, self)
        if dlg.exec() == QDialog.DialogCode.Accepted:
            for dim in self.dimensions:
                if (item := dim.get("callout_item")) is not None:
                    item.set_color(self.config.callout_color)
                    item.set_radius(self.config.callout_radius)
                if (row := dim.get("row_widget")) is not None and dim.get("included"):
                    row.set_color(self.config.callout_color)

    # =======================================================================
    # Detect remaining pages (called before export)
    # =======================================================================

    def _detect_all_remaining_pages(self) -> None:
        """Synchronously detect dimensions on any pages not yet visited."""
        if self.pdf_doc is None or detect_dimensions is None:
            return
        r_pt = self.config.callout_radius / RENDER_ZOOM

        for page_num in range(len(self.pdf_doc)):
            if page_num in self._pages_done:
                continue
            try:
                raw_dims = detect_dimensions(self.current_pdf_path, page_num)
            except Exception:
                self._pages_done.add(page_num)
                continue

            placed: list[tuple[float, float]] = [
                (d["callout_x"], d["callout_y"])
                for d in self.dimensions
                if d.get("page") == page_num and d.get("included", True)
            ]

            for raw in raw_dims:
                cx, cy = _smart_callout_pos(raw, placed, r_pt)
                placed.append((cx, cy))

                entry = dict(raw)
                entry.update({
                    "page":        page_num,
                    "callout_x":   cx,
                    "callout_y":   cy,
                    "included":    True,
                    "callout_num": 0,
                    "callout_item": None,
                    "row_widget":   None,
                })
                self.dimensions.append(entry)

            self._pages_done.add(page_num)
            self._apply_general_tolerances(
                [d for d in self.dimensions if d.get("page") == page_num]
            )

        self._update_callout_numbers()

    # =======================================================================
    # Export — annotated PDF
    # =======================================================================

    def export_pdf(self) -> None:
        if not self.current_pdf_path or self.pdf_doc is None:
            return

        self._detect_all_remaining_pages()

        base = os.path.splitext(os.path.basename(self.current_pdf_path))[0]
        default_out = os.path.join(
            os.path.dirname(self.current_pdf_path), f"{base}_annotated.pdf"
        )
        out_path, _ = QFileDialog.getSaveFileName(
            self, "Export Annotated PDF", default_out, "PDF Files (*.pdf)"
        )
        if not out_path:
            return

        active = [d for d in self.dimensions if d.get("included")]
        if not active:
            QMessageBox.information(self, "Nothing to Export", "No callouts are currently included.")
            return

        try:
            self._write_annotated_pdf(out_path, active)
            self.status_bar.showMessage(f"PDF exported → {out_path}")
            QMessageBox.information(self, "Export Complete", f"Annotated PDF saved:\n{out_path}")
        except Exception as exc:
            QMessageBox.critical(self, "Export Failed", f"Could not write PDF:\n\n{exc}")

    # =======================================================================
    # Shared annotation helper
    # =======================================================================

    def _annotate_doc(
        self, doc: "fitz.Document", active_dims: list
    ) -> "dict[int, list]":
        """Draw callout circles on *doc* in-place.

        Returns a ``dims_by_page`` dict mapping page index → list of dims.
        Mutates the fitz.Document so caller can either save it or render
        pages to images.
        """
        r_pt = self.config.callout_radius / RENDER_ZOOM

        qt_color = QColor(self.config.callout_color)
        stroke_rgb = (qt_color.redF(), qt_color.greenF(), qt_color.blueF())

        # Base font size — will be scaled per number to always fit the circle.
        # diameter = 2*r_pt; target ≈ 60 % of diameter → r_pt * 1.2
        base_font_sz = max(8.0, r_pt * 1.2)
        # Available width inside the circle border (2 pt margin each side)
        _available_w = max(1.0, 2.0 * r_pt - 4.0)

        dims_by_page: dict[int, list] = {}
        for dim in active_dims:
            pg = dim.get("page", self.current_page)
            dims_by_page.setdefault(pg, []).append(dim)

        for page_num, page_dims in dims_by_page.items():
            if page_num >= len(doc):
                continue
            page = doc[page_num]
            for dim in page_dims:
                cx  = dim["callout_x"]
                cy  = dim["callout_y"]
                num_str = str(dim["callout_num"])

                circle_rect = fitz.Rect(cx - r_pt, cy - r_pt, cx + r_pt, cy + r_pt)

                # ── Font size: shrink to fit wide numbers inside the circle ──
                # Measure at the default size, then scale down proportionally
                # so the text always fits within the available diameter.
                try:
                    tw_default = fitz.get_text_length(
                        num_str, fontname="hebo", fontsize=base_font_sz
                    )
                except AttributeError:
                    tw_default = base_font_sz * 0.60 * len(num_str)
                if tw_default > _available_w and tw_default > 0:
                    font_sz = max(6.0, base_font_sz * _available_w / tw_default)
                else:
                    font_sz = base_font_sz

                # ── Circle: white fill + coloured border ─────────────────
                page.draw_oval(
                    circle_rect,
                    color=stroke_rgb, fill=(1.0, 1.0, 1.0),
                    width=2.0,
                )

                # ── Vertically centred number ─────────────────────────────
                # Use insert_text with manually computed centre so the
                # number always renders regardless of circle size or
                # PyMuPDF version (insert_textbox can silently swallow
                # text when the rect is just a little too tight).
                if dim["callout_num"] > 0:
                    # Horizontal: measure exact glyph width at final font_sz.
                    try:
                        tw = fitz.get_text_length(
                            num_str, fontname="hebo", fontsize=font_sz
                        )
                    except AttributeError:
                        tw = font_sz * 0.60 * len(num_str)
                    tx = cx - tw / 2.0
                    # Vertical: insert_text places the baseline at the given
                    # y-coordinate.  For digits the ascender ≈ 0.72×font_sz,
                    # so the visual centre sits 0.36×font_sz above baseline.
                    # Solve for baseline: baseline = cy + 0.36×font_sz.
                    ty = cy + font_sz * 0.36
                    page.insert_text(
                        fitz.Point(tx, ty),
                        num_str,
                        fontsize=font_sz,
                        fontname="hebo",
                        color=stroke_rgb,
                    )

        return dims_by_page

    # =======================================================================
    # Export — annotated PDF
    # =======================================================================

    def _write_annotated_pdf(self, out_path: str, active_dims: list) -> None:
        doc = fitz.open(self.current_pdf_path)

        # If we restored a session from a previously annotated PDF, the old
        # callout circles are burned into the content.  Redact (white-fill)
        # those exact positions before drawing the updated ones so that no
        # stale numbers show through underneath.
        if self._session_burned_positions:
            pages_touched: set[int] = set()
            for bp in self._session_burned_positions:
                pg = bp["page"]
                if pg >= len(doc):
                    continue
                r = bp["r_pt"]
                cx, cy = bp["x"], bp["y"]
                doc[pg].add_redact_annot(
                    fitz.Rect(cx - r, cy - r, cx + r, cy + r),
                    fill=(1, 1, 1),
                )
                pages_touched.add(pg)
            for pg in pages_touched:
                doc[pg].apply_redactions()

        self._annotate_doc(doc, active_dims)
        doc.save(out_path, garbage=4, deflate=True)
        doc.close()

        # Save / update the session sidecar next to the exported PDF so the
        # user can re-open this annotated file and resume editing later.
        self._save_session_sidecar(out_path)

        # Update burned positions to reflect what was just drawn, so a
        # subsequent re-export from this same session redacts correctly.
        r_pt = self.config.callout_radius / RENDER_ZOOM
        self._session_burned_positions = [
            {"page": d.get("page", 0), "x": d["callout_x"], "y": d["callout_y"], "r_pt": r_pt}
            for d in self.dimensions
            if d.get("included", True) and d.get("callout_num", 0) > 0
        ]

    # =======================================================================
    # Embed annotated pages into an Excel sheet
    # =======================================================================

    @staticmethod
    def _make_xl_image(XlImage, img_bytes: "io.BytesIO",
                       raw_width: int, raw_height: int):
        """Construct an openpyxl Image without requiring a Pillow install.

        openpyxl uses Pillow solely to read the pixel dimensions of an image.
        Because we already know the exact dimensions from the PyMuPDF pixmap,
        we can inject a minimal in-process shim that satisfies openpyxl's
        ``from PIL.Image import open`` call and returns those known dimensions.
        The real image bytes are stored normally; Pillow is never needed.
        """
        import sys

        img_bytes.seek(0)
        try:
            return XlImage(img_bytes)           # fast path — Pillow installed
        except ImportError:
            pass                                # Pillow absent — use shim

        # Minimal shim: looks like PIL.Image to openpyxl.
        # openpyxl only calls PIL.Image.open(fp) and reads .size / .format.
        class _Img:
            size   = (raw_width, raw_height)
            format = "PNG"

        class _PIL:
            @staticmethod
            def open(_fp):
                return _Img()

        shim = _PIL()
        prev_pil       = sys.modules.get("PIL")
        prev_pil_image = sys.modules.get("PIL.Image")
        if prev_pil is None:
            sys.modules["PIL"] = shim
        if prev_pil_image is None:
            sys.modules["PIL.Image"] = shim
        try:
            img_bytes.seek(0)
            return XlImage(img_bytes)
        finally:
            if prev_pil is None:
                sys.modules.pop("PIL", None)
            if prev_pil_image is None:
                sys.modules.pop("PIL.Image", None)

    def _embed_annotated_pages(
        self, wb, active_dims: list, data_sheet_title: str = ""
    ) -> None:
        """Render each annotated PDF page and embed it in an Excel sheet.

        Uses the first existing sheet that is *not* the data sheet
        (``data_sheet_title``).  If no such sheet exists, a new sheet called
        "Annotated Drawing" is appended.  Images are sized to ~1 100 px wide
        so callout numbers are clearly visible, stacked top-to-bottom.
        """
        try:
            from openpyxl.drawing.image import Image as XlImage  # type: ignore[import]
        except ImportError:
            return   # openpyxl is old or partial — skip silently
        import io
        import math

        # Find a sheet to write images into — prefer the existing secondary
        # sheet in the template over creating a brand-new one.
        ws = None
        for name in wb.sheetnames:
            if name != data_sheet_title:
                ws = wb[name]
                break
        if ws is None:
            ws = wb.create_sheet("Annotated Drawing")

        # Annotate a temporary in-memory copy of the PDF
        doc = fitz.open(self.current_pdf_path)
        dims_by_page = self._annotate_doc(doc, active_dims)

        # ── Image sizing ─────────────────────────────────────────────────────
        # Target ≈ 1 100 px wide so numbers are clearly readable.
        TARGET_W_PX = 1100

        # ── Excel geometry constants ──────────────────────────────────────────
        # Column width is measured in "character units" (≈7.5 px/unit at Calibri 11).
        # Row height is measured in points (1 pt ≈ 1.333 px at 96 dpi).
        PX_PER_COL_UNIT  = 7.5
        PT_PER_ROW       = 1.333   # px per point
        # Max Excel row height is 409 pt.  We split the image over multiple
        # rows so Excel renders them all at their full height.
        MAX_ROW_HEIGHT_PT = 409.0
        # Give each row a height equal to MAX_ROW_HEIGHT_PT so we need the
        # minimum number of rows and the image scales correctly.
        ROW_HEIGHT_PT    = MAX_ROW_HEIGHT_PT
        ROW_HEIGHT_PX    = ROW_HEIGHT_PT * PT_PER_ROW   # ≈ 545 px

        # Make column A wide enough to hold the image
        ws.column_dimensions["A"].width = math.ceil(TARGET_W_PX / PX_PER_COL_UNIT)

        current_row = 1
        for page_num in sorted(dims_by_page.keys()):
            if page_num >= len(doc):
                continue
            page = doc[page_num]

            # Render the annotated page at 2× (144 dpi) for crisp images
            mat = fitz.Matrix(2.0, 2.0)
            pix = page.get_pixmap(matrix=mat, alpha=False)

            img_bytes = io.BytesIO(pix.tobytes("png"))
            xl_img = self._make_xl_image(XlImage, img_bytes, pix.width, pix.height)

            scale = TARGET_W_PX / pix.width
            xl_img.width  = TARGET_W_PX
            xl_img.height = int(pix.height * scale)

            # Anchor the image at column A of the current row
            ws.add_image(xl_img, f"A{current_row}")

            # Set row heights so cells visually match the image height,
            # keeping each row ≤ MAX_ROW_HEIGHT_PT.
            rows_for_image = math.ceil(xl_img.height / ROW_HEIGHT_PX)
            for r in range(current_row, current_row + rows_for_image):
                ws.row_dimensions[r].height = ROW_HEIGHT_PT

            # Advance past this image plus a small gap row
            current_row += rows_for_image + 2

        doc.close()

    # =======================================================================
    # Export — measurement record
    # =======================================================================

    def export_measurement_record(self) -> None:
        if load_workbook is None:
            QMessageBox.critical(self, "Missing Library", "openpyxl is not installed.")
            return

        self._detect_all_remaining_pages()

        # Sort by callout_num so spreadsheet rows match the numbered drawing
        active = sorted(
            (d for d in self.dimensions if d.get("included")),
            key=lambda d: d.get("callout_num", 0),
        )
        if not active:
            QMessageBox.information(self, "Nothing to Export", "No callouts are currently included.")
            return

        base = (
            os.path.splitext(os.path.basename(self.current_pdf_path))[0]
            if self.current_pdf_path else "drawing"
        )
        default_dir = os.path.dirname(self.current_pdf_path) if self.current_pdf_path else _APP_DIR
        default_out = os.path.join(default_dir, f"{base}_measurement_record.xlsx")

        out_path, _ = QFileDialog.getSaveFileName(
            self, "Export Measurement Record", default_out, "Excel Files (*.xlsx)"
        )
        if not out_path:
            return

        try:
            self._write_measurement_record(out_path, active)
            self.status_bar.showMessage(f"Measurement record exported → {out_path}")
            QMessageBox.information(
                self, "Export Complete",
                f"Measurement record saved:\n{out_path}\n\n"
                "All original formulas, formatting, and conditional rules are preserved.",
            )
        except Exception as exc:
            QMessageBox.critical(self, "Export Failed", f"Could not write measurement record:\n\n{exc}")

    def _write_measurement_record(self, out_path: str, active_dims: list) -> None:
        template_path = os.path.join(_APP_DIR, TEMPLATE_FILENAME)
        shutil.copy2(template_path, out_path)

        wb = load_workbook(out_path, data_only=False)

        # Use the "Dimensions" sheet (first sheet with data entry rows)
        ws = None
        for name in wb.sheetnames:
            sheet = wb[name]
            if sheet.max_row >= _DATA_START_ROW:
                ws = sheet
                break
        if ws is None:
            ws = wb.active

        for i, dim in enumerate(active_dims):
            row_idx = _DATA_START_ROW + i
            nominal, plus_tol, minus_tol = _parse_tolerances(dim["text"])
            # Fall back to title-block general tolerances when the dimension
            # text carries no explicit tolerance annotation.
            if plus_tol is None:
                plus_tol  = dim.get("applied_plus_tol")
            if minus_tol is None:
                minus_tol = dim.get("applied_minus_tol")

            def safe_write(col: int, value: Any) -> None:
                if value is None:
                    return
                cell = ws.cell(row=row_idx, column=col)
                existing = cell.value
                if isinstance(existing, str) and existing.startswith("="):
                    return
                cell.value = value

            safe_write(_COL_DIM_NUM,   dim["callout_num"])
            safe_write(_COL_NOMINAL,   nominal)
            safe_write(_COL_PLUS_TOL,  plus_tol)
            safe_write(_COL_MINUS_TOL, minus_tol)

        # Embed annotated PDF pages in the existing secondary sheet so the
        # user can cross-reference dimension numbers without leaving the workbook.
        if fitz is not None:
            self._embed_annotated_pages(wb, active_dims, data_sheet_title=ws.title)

        wb.save(out_path)
        wb.close()


# ---------------------------------------------------------------------------
# Application entry point
# ---------------------------------------------------------------------------

def main() -> None:
    missing: list[str] = []
    if fitz is None:
        missing.append("PyMuPDF  (pip install pymupdf)")
    if load_workbook is None:
        missing.append("openpyxl  (pip install openpyxl)")

    app = QApplication(sys.argv)
    app.setApplicationName("FAI Dimension Numberer")
    app.setApplicationVersion("1.1.0")
    app.setStyle("Fusion")

    palette = QPalette()
    palette.setColor(QPalette.ColorRole.Window,          QColor("#F5F5F5"))
    palette.setColor(QPalette.ColorRole.WindowText,      QColor("#1A1A1A"))
    palette.setColor(QPalette.ColorRole.Base,            QColor("#FFFFFF"))
    palette.setColor(QPalette.ColorRole.AlternateBase,   QColor("#F0F0F0"))
    palette.setColor(QPalette.ColorRole.Text,            QColor("#1A1A1A"))
    palette.setColor(QPalette.ColorRole.Highlight,       QColor("#2980B9"))
    palette.setColor(QPalette.ColorRole.HighlightedText, QColor("#FFFFFF"))
    app.setPalette(palette)

    if missing:
        QMessageBox.critical(
            None,
            "Missing Dependencies",
            "The following required packages are not installed:\n\n"
            + "\n".join(f"  • {m}" for m in missing)
            + "\n\nInstall them and relaunch the application.",
        )
        sys.exit(1)

    if _DETECTOR_IMPORT_ERROR:
        QMessageBox.critical(
            None,
            "Missing Module: dimension_detector.py",
            f"Could not import dimension_detector.py:\n\n{_DETECTOR_IMPORT_ERROR}\n\n"
            f"Ensure dimension_detector.py is in the same folder as this executable:\n"
            f"{_APP_DIR}",
        )
        sys.exit(1)

    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
