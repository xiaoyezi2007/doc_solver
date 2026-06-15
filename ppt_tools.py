from __future__ import annotations

import argparse
import gc
import re
import shutil
import sys
import time
from contextlib import contextmanager
from pathlib import Path
from typing import Iterable

import comtypes.client


def app_root() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


ROOT = app_root()
INPUT_DIR = ROOT / "input"
OUTPUT_DIR = ROOT / "output"
PPT_EXTENSIONS = {".ppt", ".pptx", ".pptm", ".pps", ".ppsx", ".pot", ".potx"}


def safe_stem(path: Path) -> str:
    name = re.sub(r'[<>:"/\\|?*\x00-\x1f]+', "_", path.stem).strip(" .")
    return name or "presentation"


def iter_presentations(source: Path | None) -> list[Path]:
    if source:
        source = source.resolve()
        if source.is_dir():
            files = [p for p in source.iterdir() if p.suffix.lower() in PPT_EXTENSIONS]
        else:
            files = [source]
    else:
        INPUT_DIR.mkdir(exist_ok=True)
        files = [p for p in INPUT_DIR.iterdir() if p.suffix.lower() in PPT_EXTENSIONS]
    return sorted(files)


def build_parser(description: str) -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description=description)
    parser.add_argument(
        "source",
        nargs="?",
        type=Path,
        help="PPT/PPTX file or a directory. Defaults to all presentations in ./input.",
    )
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        default=OUTPUT_DIR,
        help="Output root directory. Defaults to ./output.",
    )
    return parser


@contextmanager
def powerpoint_app(visible: bool = False):
    app = comtypes.client.CreateObject("PowerPoint.Application")
    try:
        app.Visible = 1 if visible else 0
    except Exception:
        if visible:
            app.Visible = 1
    try:
        yield app
    finally:
        try:
            app.Quit()
        finally:
            del app
            gc.collect()


def open_presentation(app, path: Path, read_only: bool = True, with_window: bool = False):
    return app.Presentations.Open(
        str(path.resolve()),
        ReadOnly=-1 if read_only else 0,
        Untitled=0,
        WithWindow=-1 if with_window else 0,
    )


def export_images(deck, output_dir: Path, width: int = 3840, height: int = 2160) -> Path:
    output_dir.mkdir(parents=True, exist_ok=True)
    deck.Export(str(output_dir.resolve()), "PNG", width, height)
    return output_dir


def clear_dir(path: Path):
    if path.exists():
        shutil.rmtree(path)
    path.mkdir(parents=True, exist_ok=True)


def unique_path(path: Path) -> Path:
    if not path.exists():
        return path
    stem = path.stem
    suffix = path.suffix
    parent = path.parent
    for i in range(1, 10_000):
        candidate = parent / f"{stem}_{i:03d}{suffix}"
        if not candidate.exists():
            return candidate
    raise RuntimeError(f"Cannot create a unique path for {path}")


def wait_for_files(paths: Iterable[Path], timeout: float = 30.0):
    deadline = time.time() + timeout
    paths = list(paths)
    while time.time() < deadline:
        if all(path.exists() and path.stat().st_size > 0 for path in paths):
            return
        time.sleep(0.2)
