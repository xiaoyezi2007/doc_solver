from __future__ import annotations

from pathlib import Path

from ppt_tools import build_parser, export_images, iter_presentations, open_presentation, powerpoint_app, safe_stem


def ppt_to_images(ppt_path: Path, output_root: Path, width: int, height: int):
    target = output_root / f"{safe_stem(ppt_path)}_images"
    with powerpoint_app(visible=False) as app:
        deck = open_presentation(app, ppt_path)
        try:
            export_images(deck, target, width=width, height=height)
        finally:
            deck.Close()
            del deck
    print(f"[OK] {ppt_path.name} -> {target}")


def main():
    parser = build_parser("Convert PowerPoint slides to high-resolution PNG images.")
    parser.add_argument("--width", type=int, default=3840, help="Export image width. Default: 3840.")
    parser.add_argument("--height", type=int, default=2160, help="Export image height. Default: 2160.")
    args = parser.parse_args()

    files = iter_presentations(args.source)
    if not files:
        print("No PowerPoint files found. Put files in ./input or pass a file path.")
        return

    args.output.mkdir(parents=True, exist_ok=True)
    for ppt_path in files:
        ppt_to_images(ppt_path, args.output, args.width, args.height)


if __name__ == "__main__":
    main()
