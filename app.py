from __future__ import annotations

import hashlib
import gc
import shutil
import tempfile
import traceback
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

import comtypes
from PySide6.QtCore import QEvent, QObject, QRunnable, QSize, Qt, QThreadPool, Signal, Slot
from PySide6.QtGui import QColor, QFont, QIcon, QPainter, QPen, QPixmap
from PySide6.QtWidgets import (
    QApplication,
    QAbstractItemView,
    QButtonGroup,
    QCheckBox,
    QFileDialog,
    QFrame,
    QGridLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QListView,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QPlainTextEdit,
    QScrollArea,
    QSizePolicy,
    QSpinBox,
    QStackedWidget,
    QVBoxLayout,
    QWidget,
)

from ppt2image import ppt_to_images
from ppt_export_media import export_media
from ppt_flatten import flatten_ppt
from ppt_tools import (
    INPUT_DIR,
    OUTPUT_DIR,
    PPT_EXTENSIONS,
    ROOT,
    clear_dir,
    export_images,
    open_presentation,
    powerpoint_app,
    safe_stem,
    unique_path,
)


CACHE_DIR = ROOT / ".cache" / "previews"


@dataclass(frozen=True)
class Feature:
    key: str
    title: str
    description: str
    accent: str


FEATURES = [
    Feature("images", "PPT 转图片", "将每一页导出为高清 PNG，可选择是否合并多个 PPT 的图片。", "#2563eb"),
    Feature("flatten", "PPT 扁平化", "将每页变成背景图，并按需保留视频或动图。", "#059669"),
    Feature("media", "PPT 素材导出", "导出 PPT 内嵌图片、视频和动图，可选择类型和合并方式。", "#d97706"),
]


class PreviewSignals(QObject):
    ready = Signal(str, str)
    failed = Signal(str)


class PreviewWorker(QRunnable):
    def __init__(self, ppt_path: Path):
        super().__init__()
        self.ppt_path = ppt_path
        self.signals = PreviewSignals()

    @Slot()
    def run(self):
        comtypes.CoInitialize()
        try:
            preview = first_slide_preview(self.ppt_path)
            self.signals.ready.emit(str(self.ppt_path), str(preview))
        except Exception:
            self.signals.failed.emit(str(self.ppt_path))
        finally:
            comtypes.CoUninitialize()


class JobSignals(QObject):
    log = Signal(str)
    finished = Signal(bool)


class JobWorker(QRunnable):
    def __init__(self, feature_key: str, files: list[Path], options: dict):
        super().__init__()
        self.feature_key = feature_key
        self.files = files
        self.options = options
        self.signals = JobSignals()
        self.batch_output_dir = self.build_batch_output_dir()

    def build_batch_output_dir(self) -> Path:
        base = self.files[0].parent if self.options["output_mode"] == "same" else self.options["output_path"]
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        return unique_path(base / f"{safe_stem(self.files[0])}_{timestamp}")

    @Slot()
    def run(self):
        comtypes.CoInitialize()
        ok = False
        try:
            self.batch_output_dir.mkdir(parents=True, exist_ok=True)
            self.signals.log.emit(f"输出目录：{self.batch_output_dir}")
            self.signals.log.emit(f"共 {len(self.files)} 个文件，按左侧顺序处理。")
            if self.feature_key == "images":
                self.run_images()
            elif self.feature_key == "flatten":
                self.run_flatten()
            elif self.feature_key == "media":
                self.run_media()
            ok = True
            self.signals.log.emit("全部处理完成。")
        except Exception:
            self.signals.log.emit(traceback.format_exc())
        finally:
            comtypes.CoUninitialize()
            gc.collect()
            self.signals.finished.emit(ok)

    def output_base_for(self, ppt_path: Path) -> Path:
        return self.batch_output_dir

    def common_output_base(self) -> Path:
        return self.batch_output_dir

    def run_images(self):
        width = self.options["width"]
        height = self.options["height"]
        if self.options["merge_images"]:
            target = self.common_output_base() / "merged_ppt_images"
            clear_dir(target)
            for ppt_path in self.files:
                self.signals.log.emit(f"导出图片：{ppt_path.name}")
                with tempfile.TemporaryDirectory() as temp:
                    ppt_to_images(ppt_path, Path(temp), width, height)
                    image_dir = next(Path(temp).iterdir())
                    for index, image in enumerate(png_files(image_dir), start=1):
                        shutil.copy2(image, target / f"{safe_stem(ppt_path)}_{index:03d}.png")
        else:
            for ppt_path in self.files:
                self.signals.log.emit(f"导出图片：{ppt_path.name}")
                ppt_to_images(ppt_path, self.output_base_for(ppt_path), width, height)

    def run_flatten(self):
        for ppt_path in self.files:
            self.signals.log.emit(f"扁平化：{ppt_path.name}")
            flatten_ppt(
                ppt_path,
                self.output_base_for(ppt_path),
                self.options["width"],
                self.options["height"],
                preserve_video=self.options["preserve_video"],
                preserve_gif=self.options["preserve_gif"],
            )

    def run_media(self):
        include_video = self.options["export_video"]
        include_gif = self.options["export_gif"]
        include_image = self.options["export_image"]
        if self.options["merge_media"]:
            target = self.common_output_base() / "merged_ppt_media"
            clear_dir(target)
            for ppt_path in self.files:
                self.signals.log.emit(f"导出素材：{ppt_path.name}")
                export_media(
                    ppt_path,
                    self.common_output_base(),
                    include_video=include_video,
                    include_gif=include_gif,
                    include_image=include_image,
                    target_dir=target,
                    clean_target=False,
                    filename_prefix=f"{safe_stem(ppt_path)}_",
                )
        else:
            for ppt_path in self.files:
                self.signals.log.emit(f"导出素材：{ppt_path.name}")
                export_media(
                    ppt_path,
                    self.output_base_for(ppt_path),
                    include_video=include_video,
                    include_gif=include_gif,
                    include_image=include_image,
                )


def first_slide_preview(ppt_path: Path) -> Path:
    CACHE_DIR.mkdir(parents=True, exist_ok=True)
    stat = ppt_path.stat()
    key = hashlib.sha1(f"{ppt_path.resolve()}:{stat.st_mtime_ns}:{stat.st_size}".encode("utf-8")).hexdigest()
    target = CACHE_DIR / f"{key}.png"
    if target.exists():
        return target

    work_dir = CACHE_DIR / key
    clear_dir(work_dir)
    with powerpoint_app(visible=False) as app:
        deck = open_presentation(app, ppt_path, read_only=True, with_window=False)
        try:
            export_images(deck, work_dir, width=480, height=270)
        finally:
            deck.Close()
            del deck
    images = png_files(work_dir)
    if not images:
        raise RuntimeError(f"Cannot create preview for {ppt_path}")
    shutil.copy2(images[0], target)
    shutil.rmtree(work_dir, ignore_errors=True)
    return target


def png_files(path: Path) -> list[Path]:
    return sorted({p.resolve() for p in list(path.glob("*.PNG")) + list(path.glob("*.png"))})


PREVIEW_ICON_SIZE = QSize(132, 74)
PREVIEW_GRID_SIZE = QSize(150, 130)


def placeholder_icon(color: str = "#e2e8f0") -> QIcon:
    pixmap = QPixmap(PREVIEW_ICON_SIZE)
    pixmap.fill(Qt.transparent)
    pixmap.fill(Qt.GlobalColor.white)
    return QIcon(pixmap)


class FileQueue(QListWidget):
    files_changed = Signal()

    def __init__(self, pool: QThreadPool):
        super().__init__()
        self.pool = pool
        self.drag_start_pos = None
        self.drag_start_index = None
        self.drop_index = None
        self.dragging_item = False
        self.dragging_files = False
        self.setDragEnabled(False)
        self.setDragDropMode(QAbstractItemView.NoDragDrop)
        self.setDefaultDropAction(Qt.CopyAction)
        self.setSelectionMode(QAbstractItemView.SingleSelection)
        self.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.setViewMode(QListWidget.IconMode)
        self.setFlow(QListView.LeftToRight)
        self.setWrapping(True)
        self.setResizeMode(QListWidget.Adjust)
        self.setMovement(QListView.Static)
        self.setSpacing(10)
        self.setIconSize(PREVIEW_ICON_SIZE)
        self.setGridSize(PREVIEW_GRID_SIZE)
        self.setWordWrap(True)
        self.setDropIndicatorShown(False)
        self.setAcceptDrops(True)
        self.viewport().setAcceptDrops(True)
        self.viewport().installEventFilter(self)

    def eventFilter(self, source, event):
        if source is self.viewport() and event.type() in {
            QEvent.Type.DragEnter,
            QEvent.Type.DragMove,
            QEvent.Type.DragLeave,
            QEvent.Type.Drop,
        }:
            return self.handle_external_drag_event(event, source)
        return super().eventFilter(source, event)

    def viewportEvent(self, event):
        if event.type() in {
            QEvent.Type.DragEnter,
            QEvent.Type.DragMove,
            QEvent.Type.DragLeave,
            QEvent.Type.Drop,
        }:
            return self.handle_external_drag_event(event, self.viewport())
        return super().viewportEvent(event)

    def handle_external_drag_event(self, event, source=None) -> bool:
        if event.type() == QEvent.Type.DragLeave:
            self.reset_external_drag_state()
            event.accept()
            return True

        paths = self.paths_from_mime(event.mimeData())
        if not paths:
            event.ignore()
            return True

        if event.type() in {QEvent.Type.DragEnter, QEvent.Type.DragMove}:
            self.dragging_files = True
            self.drop_index = self.insertion_index_at(self.event_pos_in_viewport(event, source))
            self.viewport().update()
            self.accept_file_drag(event)
            return True

        if event.type() == QEvent.Type.Drop:
            insertion_index = self.drop_index
            self.add_paths(paths, insertion_index=insertion_index)
            self.reset_external_drag_state()
            self.accept_file_drag(event)
            return True

        return False

    def dragEnterEvent(self, event):
        if self.paths_from_mime(event.mimeData()):
            self.dragging_files = True
            self.drop_index = self.insertion_index_at(self.event_pos(event))
            self.viewport().update()
            self.accept_file_drag(event)
        else:
            event.ignore()

    def dragMoveEvent(self, event):
        if self.paths_from_mime(event.mimeData()):
            self.dragging_files = True
            self.drop_index = self.insertion_index_at(self.event_pos(event))
            self.viewport().update()
            self.accept_file_drag(event)
        else:
            event.ignore()

    def dragLeaveEvent(self, event):
        self.reset_external_drag_state()
        event.accept()

    def dropEvent(self, event):
        paths = self.paths_from_mime(event.mimeData())
        if paths:
            self.add_paths(paths, insertion_index=self.drop_index)
            self.reset_external_drag_state()
            self.accept_file_drag(event)
        else:
            event.ignore()

    def accept_file_drag(self, event):
        event.setDropAction(Qt.CopyAction)
        event.accept()

    def paths_from_mime(self, mime_data) -> list[Path]:
        paths: list[Path] = []
        if mime_data.hasUrls():
            for url in mime_data.urls():
                local = url.toLocalFile()
                if local:
                    paths.append(Path(local))

        if not paths and mime_data.hasText():
            for line in mime_data.text().splitlines():
                text = line.strip().strip('"')
                if text.startswith("file:///"):
                    text = text.removeprefix("file:///")
                if text:
                    paths.append(Path(text))

        return paths

    def event_pos_in_viewport(self, event, source=None):
        pos = self.event_pos(event)
        if source is None or source is self.viewport():
            return pos
        return self.viewport().mapFromGlobal(source.mapToGlobal(pos))

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            item = self.itemAt(self.event_pos(event))
            self.drag_start_pos = self.event_pos(event)
            self.drag_start_index = self.row(item) if item else None
            self.drop_index = self.drag_start_index
        super().mousePressEvent(event)

    def mouseMoveEvent(self, event):
        if self.drag_start_index is None or not event.buttons() & Qt.LeftButton:
            super().mouseMoveEvent(event)
            return

        pos = self.event_pos(event)
        distance = (pos - self.drag_start_pos).manhattanLength()
        if distance < QApplication.startDragDistance() and not self.dragging_item:
            super().mouseMoveEvent(event)
            return

        self.dragging_item = True
        self.drop_index = self.insertion_index_at(pos)
        self.viewport().update()
        event.accept()

    def mouseReleaseEvent(self, event):
        if self.dragging_item and event.button() == Qt.LeftButton:
            self.move_item(self.drag_start_index, self.drop_index)
            self.reset_drag_state()
            event.accept()
            return

        self.reset_drag_state()
        super().mouseReleaseEvent(event)

    def paintEvent(self, event):
        super().paintEvent(event)
        if not (self.dragging_item or self.dragging_files) or self.drop_index is None:
            return

        line = self.insertion_line(self.drop_index)
        if not line:
            return

        painter = QPainter(self.viewport())
        pen = QPen(QColor("#2563eb"), 4)
        pen.setCapStyle(Qt.RoundCap)
        painter.setPen(pen)
        painter.drawLine(line[0], line[1], line[0], line[2])

    def event_pos(self, event):
        try:
            return event.position().toPoint()
        except AttributeError:
            return event.pos()

    def reset_drag_state(self):
        self.drag_start_pos = None
        self.drag_start_index = None
        self.drop_index = None
        self.dragging_item = False
        self.viewport().update()

    def reset_external_drag_state(self):
        self.drop_index = None
        self.dragging_files = False
        self.viewport().update()

    def insertion_index_at(self, pos) -> int:
        if self.count() == 0:
            return 0

        best_index = self.count()
        best_score = float("inf")
        for index in range(self.count() + 1):
            line = self.insertion_line(index)
            if not line:
                continue
            x, top, bottom = line
            center_y = (top + bottom) / 2
            vertical_penalty = 0 if top <= pos.y() <= bottom else abs(pos.y() - center_y)
            score = abs(pos.x() - x) + vertical_penalty * 3
            if score < best_score:
                best_score = score
                best_index = index
        return best_index

    def insertion_line(self, index: int):
        viewport_width = self.viewport().width()
        if self.count() == 0:
            x = max(8, min(viewport_width - 8, self.viewport().width() // 2))
            return x, 16, min(self.viewport().height() - 16, PREVIEW_GRID_SIZE.height() + 16)

        margin = 6
        if index <= 0:
            rect = self.visualItemRect(self.item(0))
            if not rect.isValid():
                return None
            x = max(4, rect.left() - margin)
            return x, rect.top(), rect.bottom()

        if index >= self.count():
            rect = self.visualItemRect(self.item(self.count() - 1))
            if not rect.isValid():
                return None
            x = min(viewport_width - 4, rect.right() + margin)
            return x, rect.top(), rect.bottom()

        previous_rect = self.visualItemRect(self.item(index - 1))
        next_rect = self.visualItemRect(self.item(index))
        if not previous_rect.isValid() or not next_rect.isValid():
            return None

        new_row = abs(previous_rect.top() - next_rect.top()) > previous_rect.height() // 2
        if new_row:
            x = min(viewport_width - 4, previous_rect.right() + margin)
            return x, previous_rect.top(), previous_rect.bottom()

        x = max(4, next_rect.left() - margin)
        return x, next_rect.top(), next_rect.bottom()

    def move_item(self, old_index: int | None, insertion_index: int | None):
        if old_index is None or insertion_index is None:
            return
        if insertion_index == old_index or insertion_index == old_index + 1:
            return

        item = self.takeItem(old_index)
        if insertion_index > old_index:
            insertion_index -= 1
        insertion_index = max(0, min(insertion_index, self.count()))
        self.insertItem(insertion_index, item)
        self.setCurrentItem(item)
        self.files_changed.emit()

    def add_paths(self, paths: list[Path], insertion_index: int | None = None):
        existing = {item.data(Qt.UserRole) for item in self.items()}
        added = False
        insert_at = self.count() if insertion_index is None else max(0, min(insertion_index, self.count()))
        for path in paths:
            if path.is_dir():
                candidates = sorted(p for p in path.iterdir() if p.suffix.lower() in PPT_EXTENSIONS)
            else:
                candidates = [path]
            for candidate in candidates:
                candidate = candidate.resolve()
                if candidate.suffix.lower() not in PPT_EXTENSIONS or str(candidate) in existing:
                    continue
                item = QListWidgetItem(placeholder_icon(), candidate.name)
                item.setData(Qt.UserRole, str(candidate))
                item.setToolTip(str(candidate))
                self.insertItem(insert_at, item)
                insert_at += 1
                existing.add(str(candidate))
                added = True
                self.load_preview(candidate)
        if added:
            self.files_changed.emit()

    def keyPressEvent(self, event):
        if event.key() == Qt.Key_Delete:
            for item in self.selectedItems():
                self.takeItem(self.row(item))
            self.files_changed.emit()
            return
        super().keyPressEvent(event)

    def load_preview(self, path: Path):
        worker = PreviewWorker(path)
        worker.signals.ready.connect(self.set_preview)
        self.pool.start(worker)

    @Slot(str, str)
    def set_preview(self, path_text: str, image_text: str):
        for item in self.items():
            if item.data(Qt.UserRole) == path_text:
                pixmap = QPixmap(image_text)
                if not pixmap.isNull():
                    item.setIcon(QIcon(pixmap.scaled(PREVIEW_ICON_SIZE, Qt.KeepAspectRatio, Qt.SmoothTransformation)))
                break

    def items(self):
        return [self.item(i) for i in range(self.count())]

    def paths(self) -> list[Path]:
        return [Path(item.data(Qt.UserRole)) for item in self.items()]


class CardButton(QFrame):
    def __init__(self, feature: Feature, on_click):
        super().__init__()
        self.feature = feature
        self.on_click = on_click
        self.setObjectName("featureCard")
        self.setCursor(Qt.PointingHandCursor)
        self.setMinimumSize(210, 150)
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(18, 16, 18, 16)
        layout.setSpacing(10)
        stripe = QFrame()
        stripe.setFixedHeight(4)
        stripe.setStyleSheet(f"background: {feature.accent}; border-radius: 2px;")
        layout.addWidget(stripe)
        title = QLabel(feature.title)
        title.setObjectName("cardTitle")
        desc = QLabel(feature.description)
        desc.setObjectName("cardDescription")
        desc.setWordWrap(True)
        layout.addWidget(title)
        layout.addWidget(desc)
        layout.addStretch()
        cta = QLabel("打开处理界面")
        cta.setObjectName("cardCta")
        layout.addWidget(cta)

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            self.on_click(self.feature)


class HomePage(QWidget):
    def __init__(self, open_feature):
        super().__init__()
        root = QVBoxLayout(self)
        root.setContentsMargins(28, 28, 28, 28)
        root.setSpacing(20)
        title = QLabel("Doc Solver")
        title.setObjectName("pageTitle")
        subtitle = QLabel("选择一个功能开始处理文件。新增功能会继续按同一框架出现在这里。")
        subtitle.setObjectName("pageSubtitle")
        root.addWidget(title)
        root.addWidget(subtitle)
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)
        content = QWidget()
        grid = QGridLayout(content)
        grid.setSpacing(16)
        for i, feature in enumerate(FEATURES):
            grid.addWidget(CardButton(feature, open_feature), i // 3, i % 3)
        for col in range(3):
            grid.setColumnStretch(col, 1)
        scroll.setWidget(content)
        root.addWidget(scroll, 1)


class FeaturePage(QWidget):
    back_requested = Signal()

    def __init__(self, feature: Feature, pool: QThreadPool):
        super().__init__()
        self.feature = feature
        self.pool = pool
        self.running = False
        self.external_drop_targets = []
        self.build()

    def build(self):
        root = QVBoxLayout(self)
        root.setContentsMargins(24, 20, 24, 24)
        root.setSpacing(14)
        top = QHBoxLayout()
        back = QPushButton("返回首页")
        back.clicked.connect(self.back_requested.emit)
        title = QLabel(self.feature.title)
        title.setObjectName("pageTitleSmall")
        top.addWidget(back)
        top.addWidget(title)
        top.addStretch()
        root.addLayout(top)

        body = QHBoxLayout()
        body.setSpacing(18)
        root.addLayout(body, 1)

        left = QFrame()
        left.setObjectName("panel")
        self.register_external_drop_target(left)
        left_layout = QVBoxLayout(left)
        left_layout.setContentsMargins(16, 16, 16, 16)
        left_layout.setSpacing(12)
        hint = QLabel("拖入 PPT 文件，或拖动下方小框改变处理顺序。")
        hint.setObjectName("fieldHelp")
        self.register_external_drop_target(hint)
        self.queue = FileQueue(self.pool)
        add_button = QPushButton("+")
        add_button.setObjectName("addButton")
        add_button.setToolTip("添加文件")
        add_button.clicked.connect(self.choose_files)
        self.register_external_drop_target(add_button)
        left_layout.addWidget(hint)
        left_layout.addWidget(self.queue, 1)
        left_layout.addWidget(add_button)
        body.addWidget(left, 3)

        right = QFrame()
        right.setObjectName("panel")
        right_layout = QVBoxLayout(right)
        right_layout.setContentsMargins(18, 18, 18, 18)
        right_layout.setSpacing(14)
        body.addWidget(right, 2)

        self.build_options(right_layout)
        right_layout.addStretch()
        self.start_button = QPushButton("开始")
        self.start_button.setObjectName("startButton")
        self.start_button.clicked.connect(self.start)
        right_layout.addWidget(self.start_button)

        self.log = QPlainTextEdit()
        self.log.setObjectName("logBox")
        self.log.setReadOnly(True)
        self.log.setMaximumHeight(150)
        root.addWidget(self.log)

    def register_external_drop_target(self, widget):
        widget.setAcceptDrops(True)
        widget.installEventFilter(self)
        self.external_drop_targets.append(widget)

    def eventFilter(self, source, event):
        if source in self.external_drop_targets and event.type() in {
            QEvent.Type.DragEnter,
            QEvent.Type.DragMove,
            QEvent.Type.DragLeave,
            QEvent.Type.Drop,
        }:
            return self.queue.handle_external_drag_event(event, source)
        return super().eventFilter(source, event)

    def build_options(self, layout: QVBoxLayout):
        layout.addWidget(section_title("输出路径"))
        self.same_output = QCheckBox("与源文件相同")
        self.custom_output = QCheckBox("指定路径")
        self.same_output.setChecked(True)
        self.output_group = QButtonGroup(self)
        self.output_group.setExclusive(True)
        self.output_group.addButton(self.same_output)
        self.output_group.addButton(self.custom_output)
        layout.addWidget(self.same_output)
        layout.addWidget(self.custom_output)

        row = QHBoxLayout()
        self.output_input = QLineEdit(str(OUTPUT_DIR))
        self.output_input.setEnabled(False)
        self.output_button = QPushButton("选择")
        self.output_button.setEnabled(False)
        self.output_button.clicked.connect(self.choose_output)
        row.addWidget(self.output_input, 1)
        row.addWidget(self.output_button)
        layout.addLayout(row)
        self.custom_output.toggled.connect(self.output_input.setEnabled)
        self.custom_output.toggled.connect(self.output_button.setEnabled)

        if self.feature.key in {"images", "flatten"}:
            layout.addWidget(section_title("图片尺寸"))
            size = QHBoxLayout()
            self.width_spin = QSpinBox()
            self.width_spin.setRange(1, 16000)
            self.width_spin.setValue(3840)
            self.height_spin = QSpinBox()
            self.height_spin.setRange(1, 16000)
            self.height_spin.setValue(2160)
            size.addWidget(QLabel("宽"))
            size.addWidget(self.width_spin)
            size.addWidget(QLabel("高"))
            size.addWidget(self.height_spin)
            layout.addLayout(size)
        else:
            self.width_spin = None
            self.height_spin = None

        layout.addWidget(section_title("任务选项"))
        if self.feature.key == "images":
            self.merge_images = QCheckBox("将多个 PPT 导出的图片合并")
            layout.addWidget(self.merge_images)
        elif self.feature.key == "flatten":
            self.preserve_video = QCheckBox("保留视频")
            self.preserve_video.setChecked(True)
            self.preserve_gif = QCheckBox("保留动图")
            self.preserve_gif.setChecked(True)
            layout.addWidget(self.preserve_video)
            layout.addWidget(self.preserve_gif)
        elif self.feature.key == "media":
            self.export_video = QCheckBox("导出视频")
            self.export_video.setChecked(True)
            self.export_gif = QCheckBox("导出动图")
            self.export_gif.setChecked(True)
            self.export_image = QCheckBox("导出图片")
            self.export_image.setChecked(True)
            self.merge_media = QCheckBox("将多个 PPT 的导出合并")
            layout.addWidget(self.export_video)
            layout.addWidget(self.export_gif)
            layout.addWidget(self.export_image)
            layout.addWidget(self.merge_media)

    def choose_files(self):
        files, _ = QFileDialog.getOpenFileNames(
            self,
            "选择 PPT 文件",
            str(INPUT_DIR),
            "PowerPoint (*.ppt *.pptx *.pptm *.pps *.ppsx *.pot *.potx);;All files (*.*)",
        )
        self.queue.add_paths([Path(file) for file in files])

    def choose_output(self):
        path = QFileDialog.getExistingDirectory(self, "选择输出目录", self.output_input.text())
        if path:
            self.output_input.setText(path)

    def options(self) -> dict:
        data = {
            "output_mode": "custom" if self.custom_output.isChecked() else "same",
            "output_path": Path(self.output_input.text()).expanduser(),
            "width": self.width_spin.value() if self.width_spin else 0,
            "height": self.height_spin.value() if self.height_spin else 0,
        }
        if self.feature.key == "images":
            data["merge_images"] = self.merge_images.isChecked()
        elif self.feature.key == "flatten":
            data["preserve_video"] = self.preserve_video.isChecked()
            data["preserve_gif"] = self.preserve_gif.isChecked()
        elif self.feature.key == "media":
            data["export_video"] = self.export_video.isChecked()
            data["export_gif"] = self.export_gif.isChecked()
            data["export_image"] = self.export_image.isChecked()
            data["merge_media"] = self.merge_media.isChecked()
        return data

    def start(self):
        if self.running:
            return
        files = self.queue.paths()
        if not files:
            QMessageBox.warning(self, "缺少文件", "请先添加至少一个 PPT 文件。")
            return
        options = self.options()
        if options["output_mode"] == "custom":
            options["output_path"].mkdir(parents=True, exist_ok=True)
        if self.feature.key == "media" and not any([options["export_video"], options["export_gif"], options["export_image"]]):
            QMessageBox.warning(self, "缺少选项", "请至少选择一种要导出的资源类型。")
            return

        self.running = True
        self.start_button.setEnabled(False)
        self.log.clear()
        self.append_log("任务启动。")
        worker = JobWorker(self.feature.key, files, options)
        worker.signals.log.connect(self.append_log)
        worker.signals.finished.connect(self.finish)
        self.pool.start(worker)

    @Slot(str)
    def append_log(self, message: str):
        self.log.appendPlainText(message)

    @Slot(bool)
    def finish(self, ok: bool):
        self.running = False
        self.start_button.setEnabled(True)


def section_title(text: str) -> QLabel:
    label = QLabel(text)
    label.setObjectName("sectionTitle")
    return label


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Doc Solver")
        self.resize(1120, 760)
        self.setMinimumSize(960, 640)
        INPUT_DIR.mkdir(exist_ok=True)
        OUTPUT_DIR.mkdir(exist_ok=True)
        CACHE_DIR.mkdir(parents=True, exist_ok=True)
        self.pool = QThreadPool.globalInstance()
        self.stack = QStackedWidget()
        self.home = HomePage(self.open_feature)
        self.stack.addWidget(self.home)
        self.pages: dict[str, FeaturePage] = {}
        self.setCentralWidget(self.stack)

    def open_feature(self, feature: Feature):
        page = self.pages.get(feature.key)
        if page is None:
            page = FeaturePage(feature, self.pool)
            page.back_requested.connect(self.show_home)
            self.pages[feature.key] = page
            self.stack.addWidget(page)
        self.stack.setCurrentWidget(page)

    def show_home(self):
        self.stack.setCurrentWidget(self.home)


def apply_style(app: QApplication):
    app.setFont(QFont("Microsoft YaHei UI", 10))
    app.setStyleSheet(
        """
        QWidget { background: #f6f7fb; color: #172033; }
        QLabel#pageTitle { font-size: 30px; font-weight: 700; color: #111827; }
        QLabel#pageTitleSmall { font-size: 22px; font-weight: 700; color: #111827; }
        QLabel#pageSubtitle, QLabel#fieldHelp { font-size: 14px; color: #64748b; }
        QLabel#sectionTitle { font-size: 14px; font-weight: 700; color: #0f172a; margin-top: 8px; }
        QFrame#featureCard, QFrame#panel {
            background: #ffffff;
            border: 1px solid #e2e8f0;
            border-radius: 8px;
        }
        QFrame#featureCard:hover { border-color: #94a3b8; background: #fbfdff; }
        QLabel#cardTitle { font-size: 17px; font-weight: 700; color: #0f172a; }
        QLabel#cardDescription { color: #64748b; }
        QLabel#cardCta { color: #2563eb; font-weight: 600; }
        QPushButton {
            background: #ffffff;
            border: 1px solid #cbd5e1;
            border-radius: 6px;
            padding: 8px 14px;
            color: #1e293b;
            font-weight: 600;
        }
        QPushButton:hover { background: #f8fafc; border-color: #94a3b8; }
        QPushButton:disabled { color: #94a3b8; background: #f1f5f9; }
        QPushButton#startButton {
            background: #2563eb;
            border-color: #2563eb;
            color: white;
            font-size: 20px;
            padding: 14px 18px;
        }
        QPushButton#startButton:hover { background: #1d4ed8; }
        QPushButton#addButton {
            font-size: 28px;
            min-height: 42px;
            border-style: dashed;
        }
        QLineEdit, QSpinBox {
            background: #ffffff;
            border: 1px solid #cbd5e1;
            border-radius: 6px;
            padding: 8px 10px;
            min-height: 22px;
        }
        QListWidget {
            background: #f8fafc;
            border: 1px dashed #cbd5e1;
            border-radius: 8px;
            padding: 12px;
        }
        QListWidget::item {
            background: #ffffff;
            border: 1px solid #dbe3ef;
            border-radius: 8px;
            padding: 8px;
        }
        QListWidget::item:selected { border: 2px solid #2563eb; color: #0f172a; }
        QPlainTextEdit#logBox {
            background: #0f172a;
            color: #dbeafe;
            border: 1px solid #1e293b;
            border-radius: 8px;
            padding: 10px;
            font-family: Consolas, Microsoft YaHei UI;
            font-size: 12px;
        }
        QScrollArea { background: transparent; }
        QCheckBox { color: #334155; }
        """
    )


def main():
    app = QApplication([])
    apply_style(app)
    window = MainWindow()
    window.show()
    app.exec()


if __name__ == "__main__":
    main()
