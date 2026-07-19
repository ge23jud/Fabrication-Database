"""
PyQt5 GUI front-end for base.py.

Loads the Qt-Designer-authored base_gui.ui layout and wires it up to the
existing command logic in base.py. Read-only commands (ls, display, inspect,
info, tags, untagged, goto, sync, sync_all, tag, update_readme, edit_readme)
call straight into base.py and capture their console output for display.

Commands that rely on interactive input() prompts in base.py (add/create/
new_sample via entry(), delete, update, checkall, untag, comment) are
reimplemented here with Qt dialogs (QMessageBox / QInputDialog) instead of
console input, but otherwise mirror base.py's logic and read/write the same
pickle database so the CLI and GUI stay fully interchangeable.

Run with: python nosferatool.py
"""

import base as core

import html
import io
import os
import re
import shutil
import sys
from contextlib import redirect_stdout
from datetime import datetime

import numpy as np
from PyQt5 import uic
from PyQt5.QtCore import QSize, Qt
from PyQt5.QtGui import QIcon, QImage, QPixmap
from PyQt5.QtWidgets import (
    QApplication,
    QDialog,
    QFileDialog,
    QFileIconProvider,
    QFileSystemModel,
    QInputDialog,
    QListView,
    QListWidgetItem,
    QMainWindow,
    QMessageBox,
    QStyle,
)
import qt_material  # must be imported after PyQt5

UI_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "base_gui.ui")
ADD_CREATE_DIALOG_UI_PATH = os.path.join(
    os.path.dirname(os.path.abspath(__file__)), "add_create_dialog.ui"
)
MANAGE_ENTRY_DIALOG_UI_PATH = os.path.join(
    os.path.dirname(os.path.abspath(__file__)), "manage_entry_dialog.ui"
)
THEME_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "dark_orange.xml")

ANSI_COLORS = {
    core.RED: "#e05252",
    core.GREEN: "#3fae5a",
    core.BLUE: "#4a90d9",
    core.YELLOW: "#c99a1e",
    core.MAGENTA: "#b155c9",
}
_ANSI_PATTERN = re.compile(
    "(" + "|".join(re.escape(c) for c in list(ANSI_COLORS) + [core.RESET]) + ")"
)

NAME_COLORS = {"red": "#e05252", "green": "#3fae5a", "blue": "#4a90d9",
               "yellow": "#c99a1e", "magenta": "#b155c9"}

EXCLUDED_TAG_TYPES = {"des", "sim", "scr", "ana", "spl"}
SAMPLES_BASE_DIR = r"I:\e24\SQN\Researchers\Haubmann Benjamin\01_PhD\11_Samples"
GROWTH_INITIAL_DIR = r"I:\e24\SQN\Researchers\Haubmann Benjamin\01_PhD"

IMAGE_EXTENSIONS = {".png", ".jpg", ".jpeg", ".tif", ".tiff", ".bmp", ".gif"}
THUMBNAIL_SIZE = QSize(256, 256)

FILE_ICON_SIZES = {
    "Small Icons": QSize(48, 48),
    "Medium Icons": QSize(128, 128),
    "Large Icons": QSize(256, 256),
}


def _auto_contrast_grayscale16(image):
    """Min/max-stretch a Format_Grayscale16 QImage to the full 16-bit range.

    Real scientific 16-bit TIFFs (SEM/microscope tools) often only use a
    narrow slice of the 0-65535 range; without this the thumbnail renders as
    a near-featureless gray blob."""
    w, h = image.width(), image.height()
    bytes_per_line = image.bytesPerLine()
    ptr = image.bits()
    ptr.setsize(bytes_per_line * h)
    arr = np.frombuffer(ptr, dtype=np.uint16).reshape(h, bytes_per_line // 2)[:, :w]

    lo, hi = int(arr.min()), int(arr.max())
    if hi <= lo:
        return image

    stretched = ((arr.astype(np.float32) - lo) * (65535.0 / (hi - lo))).clip(0, 65535).astype(np.uint16)
    return QImage(stretched.tobytes(), w, h, w * 2, QImage.Format_Grayscale16).copy()


class ThumbnailIconProvider(QFileIconProvider):
    """Extends the default icon provider: recognized image files get a real
    content thumbnail instead of a generic file-type icon, so the Browse
    tab's Large Icons view looks like Explorer's icon view for images.
    Unsupported/undecodable images (e.g. 32-bit float TIFFs) and non-image
    files fall through to the default icon."""

    def __init__(self):
        super().__init__()
        self._cache = {}  # path -> (mtime, QIcon)

    def icon(self, file_info):
        path = file_info.filePath()
        ext = os.path.splitext(path)[1].lower()
        if ext in IMAGE_EXTENSIONS and file_info.isFile():
            thumbnail = self._thumbnail(path)
            if thumbnail is not None:
                return thumbnail
        return super().icon(file_info)

    def _thumbnail(self, path):
        try:
            mtime = os.path.getmtime(path)
        except OSError:
            return None

        cached = self._cache.get(path)
        if cached is not None and cached[0] == mtime:
            return cached[1]

        image = QImage(path)
        if image.isNull():
            return None
        if image.format() == QImage.Format_Grayscale16:
            image = _auto_contrast_grayscale16(image)

        pixmap = QPixmap.fromImage(image).scaled(
            THUMBNAIL_SIZE, Qt.KeepAspectRatio, Qt.SmoothTransformation
        )
        icon = QIcon(pixmap)
        self._cache[path] = (mtime, icon)
        return icon


_ID_TYPE_CODES = sorted(core.IDdir_dic.keys() | EXCLUDED_TAG_TYPES)
# The 3-letter type code isn't always followed by pure digits -- e.g. "des"
# entries in this lab are seen as "20251129-desel01" (code "des" + "el01").
# No trailing \b: the character class itself stops at the first non-alnum
# character (crucially including "_", which \b would NOT treat as a
# boundary against a preceding alnum -- folder names like
# "20251201-elx2534_spl2534" need the match to end right before the "_").
# Any over-matched candidate is harmless since it's only linked if it's an
# exact key in the loaded database.
_FULL_ID_RE = re.compile(r"\b\d{8}-?[a-zA-Z]{3}[a-zA-Z0-9]*")
_SHORT_TOKEN_RE = re.compile(
    r"\b(?:" + "|".join(_ID_TYPE_CODES) + r")-?[a-zA-Z]*\d+", re.IGNORECASE
)
LINK_COLOR = "#ffb74d"  # dark_orange.xml's primaryLightColor


def _resolve_short_token(token, base):
    """Reverse-lookup a short reference like 'spl2407' or 'EPI-1780' against
    the full 16-char dict keys. Only link on an unambiguous match."""
    normalized = token.lower().replace("-", "")
    matches = [k for k in base if normalized in k.lower().replace("-", "")]
    return matches[0] if len(matches) == 1 else None


def _linkify(escaped_text, base):
    """Wrap recognized ID/sample-name tokens in an escaped HTML segment with
    <a href="entry:{ID}"> links, resolved against the loaded database dict.

    Matches are found once against the original (unmodified) text and then
    assembled into the output in a single linear pass. Dashed sample IDs
    (e.g. "20251201-spl2534") contain a word boundary right before the type
    code, so the short-token pattern also matches inside a full-ID span —
    re-running regexes on already-substituted HTML would nest/corrupt the
    inserted <a> tag, so overlapping short-token matches are discarded
    instead of substituted sequentially."""
    spans = []  # (start, end, target_id, display_text)

    for m in _FULL_ID_RE.finditer(escaped_text):
        token = m.group(0)
        if token in base:
            spans.append((m.start(), m.end(), token, token))

    def _overlaps(start, end):
        return any(s < end and start < e for s, e, _, _ in spans)

    for m in _SHORT_TOKEN_RE.finditer(escaped_text):
        start, end = m.start(), m.end()
        if _overlaps(start, end):
            continue
        token = m.group(0)
        target = _resolve_short_token(token, base)
        if target:
            spans.append((start, end, target, token))

    if not spans:
        return escaped_text

    spans.sort(key=lambda s: s[0])

    out = []
    pos = 0
    for start, end, target, display in spans:
        if start < pos:
            continue  # defensive: skip anything left overlapping after sort
        out.append(escaped_text[pos:start])
        out.append(f'<a href="entry:{target}" style="color:{LINK_COLOR}; text-decoration:underline;">{display}</a>')
        pos = end
    out.append(escaped_text[pos:])
    return "".join(out)


def ansi_to_html(text, base=None):
    """Convert base.py's ANSI-colored console output to HTML for QTextBrowser/QTextEdit.

    If base (the loaded ID database dict) is provided, recognized ID/sample-name
    tokens are turned into "entry:{ID}" links resolved against it."""
    parts = _ANSI_PATTERN.split(text)
    out = []
    current_color = None
    for part in parts:
        if part in ANSI_COLORS:
            current_color = ANSI_COLORS[part]
        elif part == core.RESET:
            current_color = None
        elif part:
            escaped = html.escape(part).replace("\n", "<br>")
            if base is not None:
                escaped = _linkify(escaped, base)
            if current_color:
                out.append(f'<span style="color:{current_color}">{escaped}</span>')
            else:
                out.append(escaped)
    return "".join(out)


class AddCreateDialog(QDialog):
    """Standalone "New" window: the same add / create / new_sample actions as
    the Add / Create tab, opened via the Browse tab's "New" button. Delegates
    the actual work to the owning MainWindow's gui_add_entry/gui_create/
    gui_new_sample so there is a single source of truth for that logic."""

    def __init__(self, main_window):
        super().__init__(main_window)
        uic.loadUi(ADD_CREATE_DIALOG_UI_PATH, self)
        self.main_window = main_window

        self.addBrowseBtn.clicked.connect(self._on_add_browse)
        self.addSubmitBtn.clicked.connect(self._on_add_submit)
        self.createSubmitBtn.clicked.connect(self._on_create_submit)
        self.newSampleSubmitBtn.clicked.connect(self._on_new_sample_submit)
        self.closeBtn.clicked.connect(self.close)

    def _on_add_browse(self):
        folder = QFileDialog.getExistingDirectory(self, "Select Folder", GROWTH_INITIAL_DIR)
        if folder:
            self.addPathEdit.setText(os.path.normpath(folder))

    def _on_add_submit(self):
        path = self.addPathEdit.text().strip()
        if not path:
            QMessageBox.warning(self, "Missing path", "Please provide a folder path.")
            return
        self.main_window.gui_add_entry(path)
        self.addPathEdit.clear()

    def _on_create_submit(self):
        new_name = self.createNameEdit.text().strip()
        if not new_name:
            QMessageBox.warning(self, "Missing name", "Please provide the new folder name.")
            return
        self.main_window.gui_create(new_name)
        self.createNameEdit.clear()

    def _on_new_sample_submit(self):
        spl_name = self.newSampleEdit.text().strip()
        if not spl_name:
            QMessageBox.warning(self, "Missing name", "Please provide a sample name.")
            return
        self.main_window.gui_new_sample(spl_name)
        self.newSampleEdit.clear()


class ManageEntryDialog(QDialog):
    """Standalone "Edit" window: the same manage-entry actions as the Manage
    Entry tab, opened via the Browse tab's "Edit" button. Delegates to the
    owning MainWindow's gui_delete/gui_comment/gui_update_path and shared
    helpers so there is a single source of truth for that logic."""

    def __init__(self, main_window):
        super().__init__(main_window)
        uic.loadUi(MANAGE_ENTRY_DIALOG_UI_PATH, self)
        self.main_window = main_window

        self.manageInfoView.setOpenLinks(False)
        self.manageInfoView.setOpenExternalLinks(False)
        self.manageInfoView.anchorClicked.connect(self._on_manage_link_clicked)

        self.manageLoadBtn.clicked.connect(self._on_manage_load)
        self.commentBtn.clicked.connect(self._on_comment)
        self.editReadmeBtn.clicked.connect(self._on_edit_readme)
        self.deleteBtn.clicked.connect(self._on_delete)
        self.newPathBrowseBtn.clicked.connect(self._on_new_path_browse)
        self.updatePathSubmitBtn.clicked.connect(self._on_update_path_submit)
        self.closeBtn.clicked.connect(self.close)

    def _on_manage_load(self):
        ID = self.manageIdEdit.text().strip()
        if not ID:
            QMessageBox.warning(self, "Missing ID", "Please enter an entry ID.")
            return
        out = self.main_window._call_captured(core.info, self.main_window._to_info_query(ID))
        try:
            base = self.main_window._load_base()
        except Exception:
            base = None
        self.manageInfoView.setHtml(ansi_to_html(out, base) or "No matching entry.")

    def _on_manage_link_clicked(self, url):
        self.manageIdEdit.setText(url.path())
        self._on_manage_load()

    def _on_comment(self):
        ID = self.manageIdEdit.text().strip()
        if not ID:
            QMessageBox.warning(self, "Missing ID", "Please enter an entry ID.")
            return
        self.main_window.gui_comment(ID)

    def _on_edit_readme(self):
        ID = self.manageIdEdit.text().strip()
        if not ID:
            QMessageBox.warning(self, "Missing ID", "Please enter an entry ID.")
            return
        out = self.main_window._call_captured(core.edit_readme, ID)
        self.main_window._append_log_ansi(out)

    def _on_delete(self):
        ID = self.manageIdEdit.text().strip()
        if not ID:
            QMessageBox.warning(self, "Missing ID", "Please enter an entry ID.")
            return
        self.main_window.gui_delete(ID)

    def _on_new_path_browse(self):
        folder = QFileDialog.getExistingDirectory(self, "Select New Folder", GROWTH_INITIAL_DIR)
        if folder:
            self.newPathEdit.setText(os.path.normpath(folder))

    def _on_update_path_submit(self):
        path = self.newPathEdit.text().strip()
        if not path:
            QMessageBox.warning(self, "Missing path", "Please provide the new folder path.")
            return
        self.main_window.gui_update_path(path)
        self.newPathEdit.clear()


class MainWindow(QMainWindow):
    def __init__(self, access_level):
        super().__init__()
        uic.loadUi(UI_PATH, self)
        self.access_level = access_level

        self.setWindowTitle(f"Fabrication Database — {access_level.upper()}")
        self.accessLabel.setText(f"Access: {access_level.upper()}")

        self.fsModel = QFileSystemModel(self)
        self.fsModel.setReadOnly(True)
        self.fsModel.setIconProvider(ThumbnailIconProvider())

        self.fileTreeView.setModel(self.fsModel)
        self.fileTreeView.setColumnWidth(0, 250)
        self.fileTreeView.setExpandsOnDoubleClick(False)

        self.fileIconView.setModel(self.fsModel)
        self.fileIconView.setViewMode(QListView.IconMode)
        self.fileIconView.setResizeMode(QListView.Adjust)
        self.fileIconView.setWrapping(True)

        self._file_view_top_root = None  # this entry's own folder; Up won't go above it
        self.fileUpBtn.setIcon(self.style().standardIcon(QStyle.SP_ArrowBack))
        self.fileUpBtn.setToolTip("Up one folder")
        self.fileUpBtn.setEnabled(False)

        self.detailView.setOpenLinks(False)
        self.detailView.setOpenExternalLinks(False)
        self.manageInfoView.setOpenLinks(False)
        self.manageInfoView.setOpenExternalLinks(False)

        self._apply_access_control()
        self._connect_signals()
        self.refresh_id_list()

    # ------------------------------------------------------------------ #
    # Setup helpers
    # ------------------------------------------------------------------ #

    def _apply_access_control(self):
        is_owner = self.access_level == "owner"
        for tab_widget_name in ("tabAddCreate", "tabManage", "tabTagsSync", "tabMaintenance"):
            widget = getattr(self, tab_widget_name)
            idx = self.tabWidget.indexOf(widget)
            self.tabWidget.setTabEnabled(idx, is_owner)
            if not is_owner:
                self.tabWidget.setTabToolTip(idx, "Read-only access: owner-only tab")

        self.newBtn.setEnabled(is_owner)
        self.editBtn.setEnabled(is_owner)
        if not is_owner:
            self.newBtn.setToolTip("Read-only access: adding entries requires owner access")
            self.editBtn.setToolTip("Read-only access: editing entries requires owner access")

    def _connect_signals(self):
        self.actionRefresh.triggered.connect(self.refresh_id_list)
        self.actionExit.triggered.connect(self.close)
        self.actionAbout.triggered.connect(self._show_about)

        self.refreshBtn.clicked.connect(self.refresh_id_list)
        self.newBtn.clicked.connect(self._on_new_clicked)
        self.editBtn.clicked.connect(self._on_edit_clicked)
        self.searchEdit.textChanged.connect(self._filter_id_list)
        self.idListWidget.itemDoubleClicked.connect(lambda item: self._show_info(item.text()))

        self.gotoBtn.clicked.connect(self._on_goto)
        self.fileTreeView.doubleClicked.connect(self._on_file_tree_double_clicked)
        self.fileIconView.doubleClicked.connect(self._on_file_icon_double_clicked)
        self.fileViewModeCombo.currentIndexChanged.connect(self._on_file_view_mode_changed)
        self.fileUpBtn.clicked.connect(self._on_file_view_up)
        self.detailView.anchorClicked.connect(self._on_detail_link_clicked)

        self.addBrowseBtn.clicked.connect(self._on_add_browse)
        self.addSubmitBtn.clicked.connect(self._on_add_submit)
        self.createSubmitBtn.clicked.connect(self._on_create_submit)
        self.newSampleSubmitBtn.clicked.connect(self._on_new_sample_submit)

        self.manageLoadBtn.clicked.connect(self._on_manage_load)
        self.manageInfoView.anchorClicked.connect(self._on_manage_link_clicked)
        self.commentBtn.clicked.connect(self._on_comment)
        self.editReadmeBtn.clicked.connect(self._on_edit_readme)
        self.deleteBtn.clicked.connect(self._on_delete)
        self.newPathBrowseBtn.clicked.connect(self._on_new_path_browse)
        self.updatePathSubmitBtn.clicked.connect(self._on_update_path_submit)

        self.tagBtn.clicked.connect(self._on_tag)
        self.untagBtn.clicked.connect(self._on_untag)
        self.syncBtn.clicked.connect(self._on_sync)
        self.syncAllBtn.clicked.connect(self._on_sync_all)
        self.untaggedBtn.clicked.connect(self._on_untagged)

        self.checkAllBtn.clicked.connect(self._on_checkall)
        self.updateReadmeAllBtn.clicked.connect(self._on_update_readme_all)

        self.clearLogBtn.clicked.connect(self.logView.clear)

    def _show_about(self):
        QMessageBox.information(
            self,
            "About",
            "Fabrication Database GUI\n\n"
            "A PyQt5 front-end for base.py, covering all CLI commands:\n"
            "add, goto, ls, delete, checkall, update, display, update_readme,\n"
            "comment, inspect, edit_readme, create, new_sample, tag, untag,\n"
            "sync, sync_all, tags, untagged, info.",
        )

    # ------------------------------------------------------------------ #
    # Low-level helpers: DB access, output capture/logging
    # ------------------------------------------------------------------ #

    def _load_base(self):
        try:
            import pickle
            with open(core.IDbase_dir, "rb") as f:
                return pickle.load(f)
        except Exception as e:
            QMessageBox.critical(self, "Database unavailable", f"Could not read the database:\n{e}")
            raise

    def _save_base(self, base):
        try:
            import pickle
            with open(core.IDbase_dir, "wb") as f:
                pickle.dump(base, f)
        except Exception as e:
            QMessageBox.critical(self, "Database unavailable", f"Could not write the database:\n{e}")
            raise

    def _call_captured(self, func, *args, **kwargs):
        """Call a base.py function, capturing its stdout, returning the raw text."""
        buf = io.StringIO()
        try:
            with redirect_stdout(buf):
                func(*args, **kwargs)
        except Exception as e:
            buf.write(f"{core.RED}Error: {e}{core.RESET}\n")
        return buf.getvalue()

    def _log(self, message, color=None):
        if color:
            hexcolor = NAME_COLORS.get(color, color)
            self.logView.append(f'<span style="color:{hexcolor}">{html.escape(message)}</span>')
        else:
            self.logView.append(html.escape(message))

    def _append_log_ansi(self, raw_text):
        if raw_text.strip():
            self.logView.append(ansi_to_html(raw_text))

    def _selected_id(self):
        item = self.idListWidget.currentItem()
        if item is None:
            QMessageBox.warning(self, "No selection", "Please select an entry in the list first.")
            return None
        return item.text()

    # ------------------------------------------------------------------ #
    # Browse tab: ls, display, inspect, info, tags, goto
    # ------------------------------------------------------------------ #

    def refresh_id_list(self):
        try:
            base = self._load_base()
        except Exception:
            return
        self.idListWidget.clear()
        for ID in sorted(base.keys()):
            self.idListWidget.addItem(QListWidgetItem(ID))
        self._filter_id_list(self.searchEdit.text())

    def _filter_id_list(self, text):
        text = text.strip()
        for i in range(self.idListWidget.count()):
            item = self.idListWidget.item(i)
            item.setHidden(bool(text) and text not in item.text())

    def _to_info_query(self, ID):
        """core.info() matches a sample against tagged processes by checking
        whether the query is a substring of the short sample name used when
        tagging (e.g. "spl2534"), not the full dict key (e.g.
        "20251201-spl2534"). Shorten sample IDs so tagged processes are
        found, matching what CLI users normally type."""
        if "spl" in ID:
            m = re.search(r"spl\d+", ID, re.IGNORECASE)
            if m:
                return m.group(0)
        return ID

    def _show_info(self, ID):
        self.activeIdLabel.setText(ID)
        out = self._call_captured(core.info, self._to_info_query(ID))
        try:
            base = self._load_base()
        except Exception:
            base = None
        self.detailView.setHtml(ansi_to_html(out, base) or "No output.")
        self._show_file_tree(ID, base)

    def _on_detail_link_clicked(self, url):
        ID = url.path()
        for i in range(self.idListWidget.count()):
            item = self.idListWidget.item(i)
            if item.text() == ID:
                self.idListWidget.setCurrentItem(item)
                break
        self._show_info(ID)

    def _show_file_tree(self, ID, base=None):
        if base is None:
            try:
                base = self._load_base()
            except Exception:
                return

        entry_path = base.get(ID, {}).get("path")
        if not entry_path or not os.path.isdir(entry_path):
            self._log(f'Path not available for "{ID}"')
            empty_index = self.fsModel.index("")
            self.fileTreeView.setRootIndex(empty_index)
            self.fileIconView.setRootIndex(empty_index)
            self._file_view_top_root = None
            self.fileUpBtn.setEnabled(False)
            return

        self._navigate_file_view(self.fileTreeView, entry_path)
        self._navigate_file_view(self.fileIconView, entry_path)
        self._file_view_top_root = entry_path
        self.fileUpBtn.setEnabled(False)

    def _current_file_view(self):
        return self.fileTreeView if self.fileViewStack.currentWidget() is self.fileTreePage else self.fileIconView

    def _navigate_file_view(self, view, path):
        # setRootPath (not just setRootIndex) is needed so QFileSystemModel
        # actually scans/watches this directory - without it, a freshly
        # drilled-into subfolder can appear empty until something else
        # happens to trigger a fetch.
        root_index = self.fsModel.setRootPath(path)
        view.setRootIndex(root_index)

    def _on_file_tree_double_clicked(self, index):
        path = self.fsModel.filePath(index)
        if not path:
            return
        if os.path.isdir(path):
            self._navigate_file_view(self.fileTreeView, path)
            self._update_file_up_enabled()
        elif os.path.isfile(path):
            try:
                os.startfile(path)
            except OSError as e:
                QMessageBox.critical(self, "Could not open file", str(e))

    def _on_file_icon_double_clicked(self, index):
        path = self.fsModel.filePath(index)
        if not path:
            return
        if os.path.isdir(path):
            self._navigate_file_view(self.fileIconView, path)
            self._update_file_up_enabled()
        elif os.path.isfile(path):
            try:
                os.startfile(path)
            except OSError as e:
                QMessageBox.critical(self, "Could not open file", str(e))

    def _on_file_view_mode_changed(self, index):
        mode = self.fileViewModeCombo.itemText(index)
        old_view = self._current_file_view()
        if mode == "List":
            self.fileViewStack.setCurrentWidget(self.fileTreePage)
        else:
            size = FILE_ICON_SIZES[mode]
            self.fileIconView.setIconSize(size)
            self.fileIconView.setGridSize(QSize(size.width() + 30, size.height() + 30))
            self.fileViewStack.setCurrentWidget(self.fileIconPage)
        new_view = self._current_file_view()
        if new_view is not old_view:
            self._navigate_file_view(new_view, self.fsModel.filePath(old_view.rootIndex()))
        self._update_file_up_enabled()

    def _on_file_view_up(self):
        if self._file_view_top_root is None:
            return
        view = self._current_file_view()
        current_root = view.rootIndex()
        current_path = os.path.normcase(os.path.normpath(self.fsModel.filePath(current_root)))
        top_path = os.path.normcase(os.path.normpath(self._file_view_top_root))
        if current_path == top_path:
            return
        self._navigate_file_view(view, self.fsModel.filePath(current_root.parent()))
        self._update_file_up_enabled()

    def _update_file_up_enabled(self):
        if self._file_view_top_root is None:
            self.fileUpBtn.setEnabled(False)
            return
        current_path = os.path.normcase(os.path.normpath(self.fsModel.filePath(self._current_file_view().rootIndex())))
        top_path = os.path.normcase(os.path.normpath(self._file_view_top_root))
        self.fileUpBtn.setEnabled(current_path != top_path)

    def _on_goto(self):
        ID = self._selected_id()
        if ID is None:
            return
        out = self._call_captured(core.goto, ID)
        self._append_log_ansi(out)

    def _on_new_clicked(self):
        if getattr(self, "_add_create_dialog", None) is not None and self._add_create_dialog.isVisible():
            self._add_create_dialog.raise_()
            self._add_create_dialog.activateWindow()
            return
        self._add_create_dialog = AddCreateDialog(self)
        self._add_create_dialog.show()

    def _on_edit_clicked(self):
        if getattr(self, "_manage_entry_dialog", None) is not None and self._manage_entry_dialog.isVisible():
            dialog = self._manage_entry_dialog
        else:
            dialog = ManageEntryDialog(self)
            self._manage_entry_dialog = dialog

        selected_item = self.idListWidget.currentItem()
        if selected_item is not None:
            dialog.manageIdEdit.setText(selected_item.text())
            dialog._on_manage_load()

        dialog.show()
        dialog.raise_()
        dialog.activateWindow()

    # ------------------------------------------------------------------ #
    # Add / Create tab: add, create, new_sample (via gui_add_entry)
    # ------------------------------------------------------------------ #

    def _on_add_browse(self):
        folder = QFileDialog.getExistingDirectory(self, "Select Folder", GROWTH_INITIAL_DIR)
        if folder:
            self.addPathEdit.setText(os.path.normpath(folder))

    def _on_add_submit(self):
        path = self.addPathEdit.text().strip()
        if not path:
            QMessageBox.warning(self, "Missing path", "Please provide a folder path.")
            return
        self.gui_add_entry(path)
        self.addPathEdit.clear()

    def _on_create_submit(self):
        new_name = self.createNameEdit.text().strip()
        if not new_name:
            QMessageBox.warning(self, "Missing name", "Please provide the new folder name.")
            return
        self.gui_create(new_name)
        self.createNameEdit.clear()

    def _on_new_sample_submit(self):
        spl_name = self.newSampleEdit.text().strip()
        if not spl_name:
            QMessageBox.warning(self, "Missing name", "Please provide a sample name.")
            return
        self.gui_new_sample(spl_name)
        self.newSampleEdit.clear()

    def gui_add_entry(self, path):
        """Reimplementation of base.entry() / base.add() using Qt dialogs instead of input()."""
        if not os.path.exists(path):
            QMessageBox.critical(self, "Error", "Path does not exist")
            return

        ID, path = core.extract_ID_from_path(path)

        try:
            base = self._load_base()
        except Exception:
            return

        info_text, ok = QInputDialog.getMultiLineText(
            self, "New Entry", f'Short description for "{ID}":'
        )
        if not ok:
            return

        separator = "#" * 70
        with open(f"{path}\\{ID}_readme.txt", "a") as readme:
            readme.write(f"{info_text}\n\n{separator}\n\n")

        base[ID] = {"path": path, "info": info_text, "comments": "", "tags": {}}
        self._save_base(base)
        self._log(f'Entry "{ID}" has been added', "green")

        if "des" in ID:
            QMessageBox.information(
                self,
                "Creation code",
                "Paste the creation code into the Notepad window and save when done.",
            )
            import subprocess
            gen_code_path = f"{path}\\Generation_Code.txt"
            if not os.path.exists(gen_code_path):
                open(gen_code_path, "w").close()
            subprocess.run(["notepad", gen_code_path])

        if not any(ex in ID for ex in EXCLUDED_TAG_TYPES):
            while True:
                sample_input, ok = QInputDialog.getText(
                    self,
                    "Tag Samples",
                    "Which samples are involved? Enter sample names separated by commas\n"
                    "(e.g. spl01,spl02), or leave blank and confirm to skip:",
                )
                if not ok:
                    break
                sample_input = sample_input.strip()
                if sample_input:
                    for spl_name in [s.strip() for s in sample_input.split(",") if s.strip()]:
                        out = self._call_captured(core.tag, ID, spl_name)
                        self._append_log_ansi(out)
                    break
                else:
                    confirm = QMessageBox.question(
                        self,
                        "Confirm",
                        "No samples entered. Confirm no sample was involved?",
                    )
                    if confirm == QMessageBox.Yes:
                        break

        self.refresh_id_list()

    def gui_create(self, new_name):
        """Reimplementation of base.create() using QFileDialog instead of tkinter."""
        folder_path = QFileDialog.getExistingDirectory(self, "Select a Folder", GROWTH_INITIAL_DIR)
        if not folder_path:
            return
        folder_path = os.path.normpath(folder_path)

        parent_dir = os.path.dirname(folder_path)
        renamed_folder_path = os.path.join(parent_dir, new_name)
        try:
            os.rename(folder_path, renamed_folder_path)
        except OSError as e:
            QMessageBox.critical(self, "Rename failed", str(e))
            return

        ID = new_name[:16]
        new_parent_dir = None
        for key in core.IDdir_dic:
            if key in ID:
                new_parent_dir = core.IDdir_dic[key]
                break

        if new_parent_dir is None:
            QMessageBox.critical(
                self, "Unknown process type",
                f'Could not determine a target process directory for "{ID}".\n'
                f'The folder was renamed to "{new_name}" but not moved.',
            )
            return

        try:
            new_path = shutil.move(renamed_folder_path, new_parent_dir)
        except shutil.Error as e:
            QMessageBox.critical(self, "Move failed", str(e))
            return

        self.gui_add_entry(new_path)

    def gui_new_sample(self, spl_name):
        """Reimplementation of base.new_sample() calling gui_add_entry instead of add()."""
        try:
            base = self._load_base()
        except Exception:
            return

        samples = sorted(key for key in base.keys() if "spl" in key)
        if any(spl_name in x for x in samples):
            QMessageBox.warning(self, "Already exists", "Sample already exists.")
            return

        date = datetime.now().strftime("%Y%m%d")
        path = SAMPLES_BASE_DIR + "\\" + date + "-" + spl_name
        os.makedirs(path)
        for key in core.Sampledir_dic:
            os.makedirs(path + "\\" + core.Sampledir_dic[key])

        self.gui_add_entry(path)

    # ------------------------------------------------------------------ #
    # Manage Entry tab: update_readme_single (view), comment, edit_readme,
    # delete, update
    # ------------------------------------------------------------------ #

    def _on_manage_load(self):
        ID = self.manageIdEdit.text().strip()
        if not ID:
            QMessageBox.warning(self, "Missing ID", "Please enter an entry ID.")
            return
        out = self._call_captured(core.info, self._to_info_query(ID))
        try:
            base = self._load_base()
        except Exception:
            base = None
        self.manageInfoView.setHtml(ansi_to_html(out, base) or "No matching entry.")

    def _on_manage_link_clicked(self, url):
        self.manageIdEdit.setText(url.path())
        self._on_manage_load()

    def _on_comment(self):
        ID = self.manageIdEdit.text().strip()
        if not ID:
            QMessageBox.warning(self, "Missing ID", "Please enter an entry ID.")
            return
        self.gui_comment(ID)

    def _on_edit_readme(self):
        ID = self.manageIdEdit.text().strip()
        if not ID:
            QMessageBox.warning(self, "Missing ID", "Please enter an entry ID.")
            return
        out = self._call_captured(core.edit_readme, ID)
        self._append_log_ansi(out)

    def _on_delete(self):
        ID = self.manageIdEdit.text().strip()
        if not ID:
            QMessageBox.warning(self, "Missing ID", "Please enter an entry ID.")
            return
        self.gui_delete(ID)

    def _on_new_path_browse(self):
        folder = QFileDialog.getExistingDirectory(self, "Select New Folder", GROWTH_INITIAL_DIR)
        if folder:
            self.newPathEdit.setText(os.path.normpath(folder))

    def _on_update_path_submit(self):
        path = self.newPathEdit.text().strip()
        if not path:
            QMessageBox.warning(self, "Missing path", "Please provide the new folder path.")
            return
        self.gui_update_path(path)
        self.newPathEdit.clear()

    def gui_comment(self, ID):
        """Reimplementation of base.comment() using QInputDialog instead of input()."""
        try:
            base = self._load_base()
        except Exception:
            return
        if ID not in base:
            QMessageBox.critical(self, "Error", "Invalid ID")
            return

        text, ok = QInputDialog.getMultiLineText(self, "Add Comment", "Type your comment:")
        if not ok or not text.strip():
            return

        full_date = datetime.now().strftime("%A, %B %d, %Y")
        comment_text = f"{full_date}:\n{text}\n\n"
        with open(base[ID]["path"] + "\\" + ID + "_readme.txt", "a") as readme:
            readme.write(comment_text)

        core.update_readme_single(ID)
        self._log("Comment added", "green")

    def gui_delete(self, ID):
        """Reimplementation of base.delete() using QMessageBox instead of input()."""
        try:
            base = self._load_base()
        except Exception:
            return
        if not core.ID_exists(ID, base):
            QMessageBox.critical(self, "Error", "Invalid ID")
            return

        tag_dict = base[ID].get("tags", {})

        if QMessageBox.question(
            self, "Confirm Delete", f'Do you really want to delete entry "{ID}"?'
        ) != QMessageBox.Yes:
            return

        deleted_copy_paths = set()
        if tag_dict:
            if QMessageBox.question(
                self, "Delete Copies",
                f"Also delete {len(tag_dict)} sample folder copy/copies?",
            ) == QMessageBox.Yes:
                for spl_name, copy_path in tag_dict.items():
                    if os.path.exists(copy_path):
                        shutil.rmtree(copy_path)
                        deleted_copy_paths.add(copy_path)
                        self._log(f'Deleted copy for "{spl_name}"', "green")

        if QMessageBox.question(
            self, "Delete Readme", "Also delete the readme file?"
        ) == QMessageBox.Yes:
            readme_path = base[ID]["path"] + "\\" + ID + "_readme.txt"
            if os.path.exists(readme_path):
                os.remove(readme_path)
                self._log("Deleted readme from main path", "green")
            for spl_name, copy_path in tag_dict.items():
                if copy_path in deleted_copy_paths:
                    continue
                tagged_readme = copy_path + "\\" + ID + "_readme.txt"
                if os.path.exists(tagged_readme):
                    os.remove(tagged_readme)
                    self._log(f'Deleted readme from "{spl_name}" copy', "green")

        del base[ID]
        self._save_base(base)
        self._log(f'Entry "{ID}" has been deleted', "green")
        self.refresh_id_list()

    def gui_update_path(self, path):
        """Reimplementation of base.update() using QMessageBox instead of input()."""
        if not os.path.exists(path):
            QMessageBox.critical(self, "Error", "Path does not exist")
            return

        ID, path = core.extract_ID_from_path(path)

        try:
            base = self._load_base()
        except Exception:
            return
        if ID not in base:
            QMessageBox.critical(self, "Error", "Path exists, but ID has not been added yet")
            return

        oldpath = base[ID]["path"]
        if QMessageBox.question(
            self, "Confirm Update",
            f'Change the path of entry "{ID}"\nfrom:\n{oldpath}\nto:\n{path}?',
        ) != QMessageBox.Yes:
            return

        base[ID]["path"] = path
        self._save_base(base)
        self._log("Path has been updated", "green")

        out = self._call_captured(core.sync_folder, ID)
        self._append_log_ansi(out)
        self.refresh_id_list()

    # ------------------------------------------------------------------ #
    # Tags & Sync tab: tag, untag, sync, sync_all, untagged
    # ------------------------------------------------------------------ #

    def _on_tag(self):
        ID = self.tagIdEdit.text().strip()
        spl_name = self.tagSplEdit.text().strip()
        if not ID or not spl_name:
            QMessageBox.warning(self, "Missing input", "Please provide both an ID and a sample name.")
            return
        out = self._call_captured(core.tag, ID, spl_name)
        self._append_log_ansi(out)
        self.refresh_id_list()

    def _on_untag(self):
        ID = self.tagIdEdit.text().strip()
        spl_name = self.tagSplEdit.text().strip()
        if not ID or not spl_name:
            QMessageBox.warning(self, "Missing input", "Please provide both an ID and a sample name.")
            return
        self.gui_untag(ID, spl_name)

    def _on_sync(self):
        ID = self.tagIdEdit.text().strip()
        if not ID:
            QMessageBox.warning(self, "Missing ID", "Please provide an ID to sync.")
            return
        out = self._call_captured(core.sync_folder, ID)
        self._append_log_ansi(out)

    def _on_sync_all(self):
        out = self._call_captured(core.sync_all)
        self._append_log_ansi(out)

    def _on_untagged(self):
        out = self._call_captured(core.untagged)
        stripped = re.sub(r"\x1b\[[0-9;]*m", "", out)
        self.untaggedListWidget.clear()
        for line in stripped.splitlines():
            line = line.strip()
            if line:
                self.untaggedListWidget.addItem(QListWidgetItem(line))

    def gui_untag(self, ID, spl_name):
        """Reimplementation of base.untag() using QMessageBox instead of input()."""
        try:
            base = self._load_base()
        except Exception:
            return
        if not core.ID_exists(ID, base):
            QMessageBox.critical(self, "Error", "Invalid ID")
            return

        tag_dict = base[ID].get("tags", {})
        if spl_name not in tag_dict:
            QMessageBox.critical(self, "Error", f'Tag "{spl_name}" not found for "{ID}"')
            return

        copy_path = tag_dict[spl_name]
        if QMessageBox.question(
            self, "Delete copy?", f'Delete the copy at\n{copy_path}?'
        ) == QMessageBox.Yes:
            if os.path.isdir(copy_path):
                shutil.rmtree(copy_path)
                self._log("Deleted copy", "green")
            elif os.path.isfile(copy_path):
                os.remove(copy_path)
                self._log("Deleted file", "green")
            else:
                self._log("Copy path does not exist, skipping deletion", "yellow")

        del base[ID]["tags"][spl_name]
        self._save_base(base)
        self._log(f'Tag "{spl_name}" removed from "{ID}"', "green")

    # ------------------------------------------------------------------ #
    # Maintenance tab: checkall, update_readme
    # ------------------------------------------------------------------ #

    def _on_checkall(self):
        self.gui_checkall()

    def _on_update_readme_all(self):
        out = self._call_captured(core.update_readme)
        self._append_log_ansi(out)
        self.refresh_id_list()

    def gui_checkall(self):
        """Reimplementation of base.checkall() using QMessageBox instead of input()."""
        try:
            base = self._load_base()
        except Exception:
            return

        invalid_list = [ID for ID in base if not os.path.exists(base[ID]["path"])]
        if not invalid_list:
            self._log("Everything seems up to date", "green")
            return

        self._log(f"Invalid path found for {len(invalid_list)} entrie(s). Searching...", "red")

        for ID in invalid_list:
            expected_dir = next((core.IDdir_dic[k] for k in core.IDdir_dic if k in ID), None)
            if expected_dir is None or not os.path.exists(expected_dir):
                self._log(f"{ID} — could not determine expected directory, skipping", "red")
                continue

            matches = [
                f for f in os.listdir(expected_dir)
                if f[:16] == ID and os.path.isdir(os.path.join(expected_dir, f))
            ]
            if not matches:
                self._log(f"{ID} — not found in {expected_dir}, skipping", "red")
                continue

            for match in matches:
                new_path = os.path.join(expected_dir, match)
                if QMessageBox.question(
                    self, "Update path?", f"{ID} — found at:\n{new_path}\n\nUpdate path?"
                ) == QMessageBox.Yes:
                    base[ID]["path"] = new_path
                    self._save_base(base)
                    base = self._load_base()
                    self._log("Path updated", "green")
                else:
                    self._log("Skipped", "yellow")

        self.refresh_id_list()


def get_access_level_gui():
    """GUI-friendly variant of base.get_access_level(): reports errors via QMessageBox."""
    if not os.path.exists(core.CONFIG_PATH):
        QMessageBox.critical(
            None, "Configuration Error",
            "No config file found. Please contact the owner to obtain access.",
        )
        sys.exit(1)
    with open(core.CONFIG_PATH, "r") as f:
        for line in f:
            if line.strip().startswith("access:"):
                return line.split(":", 1)[1].strip().lower()
    QMessageBox.critical(None, "Configuration Error", "Invalid config file. Please contact the owner.")
    sys.exit(1)


def main():
    app = QApplication(sys.argv)
    qt_material.apply_stylesheet(app, theme=THEME_PATH)
    access_level = get_access_level_gui()
    window = MainWindow(access_level)
    window.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
