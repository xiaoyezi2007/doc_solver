from __future__ import annotations

import shutil
import tempfile
import zipfile
from pathlib import Path

from ppt_tools import (
    build_parser,
    clear_dir,
    iter_presentations,
    open_presentation,
    powerpoint_app,
    safe_stem,
    unique_path,
)


PP_SAVE_AS_OPEN_XML_PRESENTATION = 24
MEDIA_PREFIX = "ppt/media/"
IMAGE_EXTENSIONS = {".apng", ".bmp", ".emf", ".gif", ".jpeg", ".jpg", ".png", ".svg", ".tif", ".tiff", ".webp", ".wmf"}
VIDEO_EXTENSIONS = {".avi", ".m4v", ".mov", ".mp4", ".mpeg", ".mpg", ".wmv"}
ANIMATED_EXTENSIONS = {".gif", ".apng"}


def ensure_openxml_pptx(ppt_path: Path) -> tuple[Path, tempfile.TemporaryDirectory | None]:
    if ppt_path.suffix.lower() in {".pptx", ".pptm", ".ppsx", ".potx"}:
        return ppt_path, None

    temp_dir = tempfile.TemporaryDirectory()
    converted = Path(temp_dir.name) / f"{safe_stem(ppt_path)}.pptx"
    with powerpoint_app(visible=False) as app:
        deck = open_presentation(app, ppt_path, read_only=True, with_window=False)
        try:
            deck.SaveAs(str(converted.resolve()), PP_SAVE_AS_OPEN_XML_PRESENTATION)
        finally:
            deck.Close()
            del deck
    return converted, temp_dir


def should_export_media(filename: str, include_video: bool, include_gif: bool, include_image: bool) -> bool:
    suffix = Path(filename).suffix.lower()
    if suffix in VIDEO_EXTENSIONS:
        return include_video
    if suffix in ANIMATED_EXTENSIONS:
        return include_gif
    if suffix in IMAGE_EXTENSIONS:
        return include_image
    return include_video


def export_media(
    ppt_path: Path,
    output_root: Path,
    include_video: bool = True,
    include_gif: bool = True,
    include_image: bool = True,
    target_dir: Path | None = None,
    clean_target: bool = True,
    filename_prefix: str = "",
):
    target_dir = target_dir or output_root / f"{safe_stem(ppt_path)}_media"
    if clean_target:
        clear_dir(target_dir)
    else:
        target_dir.mkdir(parents=True, exist_ok=True)

    archive_path, temp_dir = ensure_openxml_pptx(ppt_path)
    count = 0
    try:
        with zipfile.ZipFile(archive_path) as archive:
            media_members = [
                item for item in archive.infolist()
                if item.filename.startswith(MEDIA_PREFIX) and not item.is_dir()
            ]
            for item in media_members:
                source_name = Path(item.filename).name
                if not should_export_media(source_name, include_video, include_gif, include_image):
                    continue
                target = unique_path(target_dir / f"{filename_prefix}{source_name}")
                with archive.open(item) as src, target.open("wb") as dst:
                    shutil.copyfileobj(src, dst)
                count += 1
    finally:
        if temp_dir:
            temp_dir.cleanup()

    print(f"[OK] {ppt_path.name} -> {target_dir} ({count} files)")


def main():
    parser = build_parser("Export original images, videos, audio, and other media from PowerPoint files.")
    args = parser.parse_args()

    files = iter_presentations(args.source)
    if not files:
        print("No PowerPoint files found. Put files in ./input or pass a file path.")
        return

    for ppt_path in files:
        export_media(ppt_path, args.output)


if __name__ == "__main__":
    main()
