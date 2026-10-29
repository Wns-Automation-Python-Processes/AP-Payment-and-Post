"""Build a clean, high-quality transparent logo from the raw image.png capture.

The source is a low-resolution (228x121) screenshot of the TiAuto mark:
a triangular 'A' emblem above a 'TiAuto' wordmark, light-on-dark. This script
keys out the dark background to transparency, tight-crops, upscales with a
high-quality resampler and sharpens the result for crisp edges.
"""
import os
import numpy as np
from PIL import Image, ImageFilter

BASE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(BASE, "image.png")
OUT = os.path.join(BASE, "assets")
os.makedirs(OUT, exist_ok=True)


def build():
    im = Image.open(SRC).convert("RGB")
    a = np.asarray(im).astype(np.float32)
    lum = 0.299 * a[:, :, 0] + 0.587 * a[:, :, 1] + 0.114 * a[:, :, 2]

    # Soft luminance key: dark background -> 0 alpha, bright mark -> 1 alpha.
    lo, hi = 42.0, 190.0
    alpha = np.clip((lum - lo) / (hi - lo), 0.0, 1.0)

    # Foreground is near-white; keep the original tint so anti-aliased pixels
    # blend smoothly instead of turning gray.
    fg = np.array([240.0, 241.0, 243.0], dtype=np.float32)
    rgba = np.zeros((a.shape[0], a.shape[1], 4), dtype=np.uint8)
    rgba[:, :, 0] = fg[0]
    rgba[:, :, 1] = fg[1]
    rgba[:, :, 2] = fg[2]
    rgba[:, :, 3] = (alpha * 255).astype(np.uint8)

    logo = Image.fromarray(rgba, "RGBA")

    # Tight crop to the visible mark.
    bbox = logo.getchannel("A").point(lambda v: 255 if v > 40 else 0).getbbox()
    if bbox:
        logo = logo.crop(bbox)

    # Upscale with Lanczos, then gently sharpen to recover crisp edges.
    scale = 6
    w, h = logo.size
    big = logo.resize((w * scale, h * scale), Image.LANCZOS)
    big = big.filter(ImageFilter.UnsharpMask(radius=2.0, percent=110, threshold=2))

    full = os.path.join(OUT, "tiauto_logo.png")
    big.save(full)

    # A monochrome white version useful for dark backgrounds (the UI header).
    white = Image.new("RGBA", big.size, (255, 255, 255, 0))
    white.putalpha(big.getchannel("A"))
    white.save(os.path.join(OUT, "tiauto_logo_white.png"))

    # A dark version for light backgrounds (the UI header).
    dark = Image.new("RGBA", big.size, (0, 0, 0, 0))
    dark.putalpha(big.getchannel("A"))
    dark.paste((0x1f, 0x24, 0x30), (0, 0), big.getchannel("A"))
    dark.save(os.path.join(OUT, "tiauto_logo_dark.png"))

    # A red version for light backgrounds.
    red = Image.new("RGBA", big.size, (0, 0, 0, 0))
    red.putalpha(big.getchannel("A"))
    red.paste((0xd9, 0x23, 0x2e), (0, 0), big.getchannel("A"))
    red.save(os.path.join(OUT, "tiauto_logo_red.png"))

    # Small header-sized render for the Tk UI.
    hh = 44
    hw = int(big.size[0] * hh / big.size[1])
    big.resize((hw, hh), Image.LANCZOS).save(os.path.join(OUT, "tiauto_logo_header.png"))

    # Multi-size .ico for the executable / window icon (square-padded).
    side = max(big.size)
    sq = Image.new("RGBA", (side, side), (0, 0, 0, 0))
    sq.paste(red, ((side - red.width) // 2, (side - red.height) // 2), red)
    ico_path = os.path.join(OUT, "tiauto.ico")
    sq.save(ico_path, sizes=[(16, 16), (32, 32), (48, 48), (64, 64),
                             (128, 128), (256, 256)])

    print(f"Built {full}  {big.size}  from crop {bbox}")
    print(f"Built {ico_path}")
    return full


if __name__ == "__main__":
    build()
