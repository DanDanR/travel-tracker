#!/usr/bin/env python3
"""
make_gif.py

Combine a series of PNG images (of any size / aspect ratio) into a single
animated GIF. Each frame is scaled down (never up) to fit within the
largest width/height found across all inputs, preserving its own aspect
ratio, and centered on a canvas of that size.

Usage:
    python make_gif.py frame1.png frame2.png frame3.png -o output.gif
    python make_gif.py /path/to/*.png -o output.gif --duration 150 --bg 0 0 0
    python make_gif.py img1.png img2.png --loop 1   # play once instead of forever
"""

import argparse
import sys
from pathlib import Path

from PIL import Image


def build_frames(paths, bg_color=(255, 255, 255, 255)):
    """
    Load images from `paths`, resize each to fit within the max
    width/height (preserving aspect ratio), center it on a canvas of
    that size, and flatten onto `bg_color`.

    Returns a list of Pillow Image objects in "P" (palette) mode,
    ready to be saved as GIF frames.
    """
    images = []
    for p in paths:
        try:
            images.append(Image.open(p))
        except Exception as e:
            raise SystemExit(f"Error opening '{p}': {e}")

    if not images:
        raise SystemExit("No images to process.")

    max_width = max(img.width for img in images)
    max_height = max(img.height for img in images)
    target_size = (max_width, max_height)

    frames = []
    for img in images:
        img = img.convert("RGBA")

        scale = min(target_size[0] / img.width, target_size[1] / img.height)
        new_size = (max(1, round(img.width * scale)), max(1, round(img.height * scale)))
        resized = img.resize(new_size, Image.LANCZOS)

        canvas = Image.new("RGBA", target_size, (0, 0, 0, 0))
        offset = (
            (target_size[0] - new_size[0]) // 2,
            (target_size[1] - new_size[1]) // 2,
        )
        canvas.paste(resized, offset, resized)

        bg = Image.new("RGBA", target_size, bg_color)
        composited = Image.alpha_composite(bg, canvas).convert("P", palette=Image.ADAPTIVE)
        frames.append(composited)

    return frames


def save_gif(frames, output_path, duration=200, loop=0):
    frames[0].save(
        output_path,
        save_all=True,
        append_images=frames[1:],
        duration=duration,
        loop=loop,
        disposal=2,
    )


def main(argv=None):
    parser = argparse.ArgumentParser(
        description="Combine PNG images of different sizes into a centered, "
        "non-stretched animated GIF."
    )
    parser.add_argument(
        "images",
        nargs="+",
        help="Paths to the input PNG files, in the order they should appear "
        "in the animation (shell globs like *.png also work).",
    )
    parser.add_argument(
        "-o", "--output",
        default="output.gif",
        help="Path to the output GIF file (default: output.gif)",
    )
    parser.add_argument(
        "-d", "--duration",
        type=int,
        default=200,
        help="Duration of each frame in milliseconds (default: 200)",
    )
    parser.add_argument(
        "-l", "--loop",
        type=int,
        default=0,
        help="Number of times to loop (0 = infinite, default: 0)",
    )
    parser.add_argument(
        "--bg",
        type=int,
        nargs=3,
        metavar=("R", "G", "B"),
        default=(255, 255, 255),
        help="Background/padding color as three RGB values 0-255 "
        "(default: 255 255 255, i.e. white)",
    )

    args = parser.parse_args(argv)

    paths = [Path(p) for p in args.images]
    missing = [str(p) for p in paths if not p.exists()]
    if missing:
        raise SystemExit(f"File(s) not found: {', '.join(missing)}")

    bg_color = (*args.bg, 255)

    frames = build_frames(paths, bg_color=bg_color)
    save_gif(frames, args.output, duration=args.duration, loop=args.loop)

    print(f"Saved {len(frames)} frame(s) to '{args.output}' "
          f"(canvas size: {frames[0].size[0]}x{frames[0].size[1]})")


if __name__ == "__main__":
    main()
