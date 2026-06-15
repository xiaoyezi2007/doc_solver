from __future__ import annotations

import re
import shutil
import zipfile
import xml.etree.ElementTree as ET
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
GIF_EXTENSIONS = {".gif", ".apng"}
P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
XML_NS = {"p": P_NS, "a": A_NS, "r": R_NS}


def is_video_shape(shape) -> bool:
    try:
        return shape.Type == MSO_MEDIA
    except Exception:
        return False


def is_gif_shape(shape, gif_shape_ids: set[int] | None = None, gif_shape_names: set[str] | None = None) -> bool:
    gif_shape_ids = gif_shape_ids or set()
    gif_shape_names = gif_shape_names or set()
    try:
        if int(shape.Id) in gif_shape_ids:
            return True
    except Exception:
        pass

    try:
        if str(shape.Name) in gif_shape_names:
            return True
    except Exception:
        pass

    try:
        if shape.Type in (MSO_PICTURE, MSO_LINKED_PICTURE):
            filename = str(shape.Name).lower()
            return any(filename.endswith(ext) or ext in filename for ext in GIF_EXTENSIONS)
    except Exception:
        return False


def is_overlay_shape(
    shape,
    preserve_video: bool = True,
    preserve_gif: bool = True,
    gif_shape_ids: set[int] | None = None,
    gif_shape_names: set[str] | None = None,
) -> bool:
    """Return True for media that should be copied over the flattened background."""
    return (preserve_video and is_video_shape(shape)) or (
        preserve_gif and is_gif_shape(shape, gif_shape_ids=gif_shape_ids, gif_shape_names=gif_shape_names)
    )


def media_extension(target: str) -> str:
    target = target.split("#", 1)[0].split("?", 1)[0]
    return Path(target).suffix.lower()


def slide_number_from_name(name: str) -> int:
    match = re.search(r"slide(\d+)\.xml$", name, re.IGNORECASE)
    return int(match.group(1)) if match else 0


def gif_shape_refs_by_slide(ppt_path: Path) -> dict[int, tuple[set[int], set[str]]]:
    """Return slide index -> GIF picture shape ids/names by reading PPTX XML relationships."""
    if not zipfile.is_zipfile(ppt_path):
        return {}

    refs: dict[int, tuple[set[int], set[str]]] = {}
    try:
        with zipfile.ZipFile(ppt_path) as archive:
            names = set(archive.namelist())
            slide_names = sorted(
                [
                    name
                    for name in names
                    if re.fullmatch(r"ppt/slides/slide\d+\.xml", name, flags=re.IGNORECASE)
                ],
                key=slide_number_from_name,
            )

            for slide_index, slide_name in enumerate(slide_names, start=1):
                rel_name = f"ppt/slides/_rels/{Path(slide_name).name}.rels"
                if rel_name not in names:
                    continue

                rel_root = ET.fromstring(archive.read(rel_name))
                gif_rel_ids = {
                    rel.attrib.get("Id")
                    for rel in rel_root.findall(f"{{{REL_NS}}}Relationship")
                    if media_extension(rel.attrib.get("Target", "")) in GIF_EXTENSIONS
                }
                gif_rel_ids.discard(None)
                if not gif_rel_ids:
                    continue

                slide_root = ET.fromstring(archive.read(slide_name))
                shape_ids: set[int] = set()
                shape_names: set[str] = set()
                for pic in slide_root.findall(".//p:pic", XML_NS):
                    blip = pic.find(".//a:blip", XML_NS)
                    rel_id = None
                    if blip is not None:
                        rel_id = blip.attrib.get(f"{{{R_NS}}}embed") or blip.attrib.get(f"{{{R_NS}}}link")
                    if rel_id not in gif_rel_ids:
                        continue

                    c_nv_pr = pic.find("./p:nvPicPr/p:cNvPr", XML_NS)
                    if c_nv_pr is None:
                        continue
                    try:
                        shape_ids.add(int(c_nv_pr.attrib.get("id", "0")))
                    except ValueError:
                        pass
                    name = c_nv_pr.attrib.get("name")
                    if name:
                        shape_names.add(name)

                if shape_ids or shape_names:
                    refs[slide_index] = (shape_ids, shape_names)
    except Exception as exc:
        print(f"[WARN] Cannot inspect GIF shapes in {ppt_path.name}: {exc}")

    return refs


def exported_slide_images(image_dir: Path, expected_count: int) -> list[Path]:
    images = list({path.resolve() for path in list(image_dir.glob("*.PNG")) + list(image_dir.glob("*.png"))})
    if len(images) < expected_count:
        raise RuntimeError(f"Expected {expected_count} exported slide images, found {len(images)} in {image_dir}")

    def slide_number(path: Path) -> int:
        match = re.search(r"(\d+)(?=\.png$)", path.name, re.IGNORECASE)
        return int(match.group(1)) if match else 0

    return sorted(images, key=lambda path: (slide_number(path), path.name.lower()))[:expected_count]


def delete_all_slides(deck):
    for index in range(deck.Slides.Count, 0, -1):
        deck.Slides(index).Delete()


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
    gif_refs = gif_shape_refs_by_slide(ppt_path) if preserve_gif else {}

    with powerpoint_app(visible=True) as app:
        src = open_presentation(app, ppt_path, read_only=True, with_window=False)
        dst = app.Presentations.Add()
        try:
            dst.PageSetup.SlideWidth = src.PageSetup.SlideWidth
            dst.PageSetup.SlideHeight = src.PageSetup.SlideHeight
            delete_all_slides(dst)

            export_images(src, image_dir, width=width, height=height)
            image_paths = exported_slide_images(image_dir, src.Slides.Count)

            for index in range(1, src.Slides.Count + 1):
                slide = dst.Slides.Add(dst.Slides.Count + 1, PP_LAYOUT_BLANK)
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
                gif_shape_ids, gif_shape_names = gif_refs.get(index, (set(), set()))
                for shape_index in range(1, src_slide.Shapes.Count + 1):
                    shape = src_slide.Shapes(shape_index)
                    if not is_overlay_shape(
                        shape,
                        preserve_video=preserve_video,
                        preserve_gif=preserve_gif,
                        gif_shape_ids=gif_shape_ids,
                        gif_shape_names=gif_shape_names,
                    ):
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
            shutil.rmtree(work_dir, ignore_errors=True)

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
