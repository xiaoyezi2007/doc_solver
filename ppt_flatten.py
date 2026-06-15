from __future__ import annotations

import re
from pathlib import Path

from ppt_tools import (
    build_parser,
    clear_dir,
    export_images,
    iter_presentations,
    open_presentation,
    powerpoint_app,
    safe_stem,
)


PP_LAYOUT_BLANK = 12
PP_SAVE_AS_OPEN_XML_PRESENTATION = 24
MSO_MEDIA = 16
MSO_PICTURE = 13
MSO_LINKED_PICTURE = 11


def is_video_shape(shape) -> bool:
    try:
        return shape.Type == MSO_MEDIA
    except Exception:
        return False


def is_gif_shape(shape) -> bool:
    try:
        if shape.Type in (MSO_PICTURE, MSO_LINKED_PICTURE):
            filename = str(shape.Name).lower()
            return filename.endswith(".gif") or ".gif" in filename
    except Exception:
        return False


def is_overlay_shape(shape, preserve_video: bool = True, preserve_gif: bool = True) -> bool:
    """Return True for media that should be copied over the flattened background."""
    return (preserve_video and is_video_shape(shape)) or (preserve_gif and is_gif_shape(shape))


def exported_slide_images(image_dir: Path, expected_count: int) -> list[Path]:
    images = list(image_dir.glob("*.PNG")) + list(image_dir.glob("*.png"))
    if len(images) < expected_count:
        raise RuntimeError(f"Expected {expected_count} exported slide images, found {len(images)} in {image_dir}")

    def slide_number(path: Path) -> int:
        match = re.search(r"(\d+)(?=\.png$)", path.name, re.IGNORECASE)
        return int(match.group(1)) if match else 0

    return sorted(images, key=slide_number)[:expected_count]


def flatten_ppt(
    ppt_path: Path,
    output_root: Path,
    width: int,
    height: int,
    preserve_video: bool = True,
    preserve_gif: bool = True,
):
    name = safe_stem(ppt_path)
    work_dir = output_root / f"{name}_flatten_assets"
    image_dir = work_dir / "slides"
    clear_dir(image_dir)

    flattened_path = output_root / f"{name}_flattened.pptx"
    output_root.mkdir(parents=True, exist_ok=True)

    with powerpoint_app(visible=True) as app:
        src = open_presentation(app, ppt_path, read_only=True, with_window=False)
        dst = app.Presentations.Add()
        try:
            dst.PageSetup.SlideWidth = src.PageSetup.SlideWidth
            dst.PageSetup.SlideHeight = src.PageSetup.SlideHeight

            export_images(src, image_dir, width=width, height=height)
            image_paths = exported_slide_images(image_dir, src.Slides.Count)

            for index in range(1, src.Slides.Count + 1):
                slide = dst.Slides.Add(index, PP_LAYOUT_BLANK)
                image_path = image_paths[index - 1]

                slide.Shapes.AddPicture(
                    str(image_path.resolve()),
                    LinkToFile=0,
                    SaveWithDocument=-1,
                    Left=0,
                    Top=0,
                    Width=dst.PageSetup.SlideWidth,
                    Height=dst.PageSetup.SlideHeight,
                )

                src_slide = src.Slides(index)
                for shape_index in range(1, src_slide.Shapes.Count + 1):
                    shape = src_slide.Shapes(shape_index)
                    if not is_overlay_shape(shape, preserve_video=preserve_video, preserve_gif=preserve_gif):
                        continue
                    try:
                        shape.Copy()
                        pasted = slide.Shapes.Paste()
                        overlay = pasted.Item(1)
                        overlay.Left = shape.Left
                        overlay.Top = shape.Top
                        overlay.Width = shape.Width
                        overlay.Height = shape.Height
                    except Exception as exc:
                        print(f"[WARN] Cannot copy media on slide {index}: {exc}")

            if flattened_path.exists():
                flattened_path.unlink()
            dst.SaveAs(str(flattened_path.resolve()), PP_SAVE_AS_OPEN_XML_PRESENTATION)
        finally:
            dst.Close()
            src.Close()
            del dst
            del src

    print(f"[OK] {ppt_path.name} -> {flattened_path}")


def main():
    parser = build_parser("Flatten PowerPoint slides into image-backed PPTX files.")
    parser.add_argument("--width", type=int, default=3840, help="Background image width. Default: 3840.")
    parser.add_argument("--height", type=int, default=2160, help="Background image height. Default: 2160.")
    args = parser.parse_args()

    files = iter_presentations(args.source)
    if not files:
        print("No PowerPoint files found. Put files in ./input or pass a file path.")
        return

    for ppt_path in files:
        flatten_ppt(ppt_path, args.output, args.width, args.height)


if __name__ == "__main__":
    main()
