#!/usr/bin/env python3
"""Build a double-clickable "<COURSE_CODE> Dashboard.app" for macOS.

macOS has no .desktop files, so the Linux launcher setup.sh writes has no
equivalent here. This makes the equivalent: a minimal .app bundle whose
executable is a two-line shell script exec'ing launch_dashboard.sh, plus a
generated .icns so it looks like a real app in Finder, Spotlight and the Dock.

The bundle goes in ~/Applications, deliberately *outside* the Dropbox folder
-- a synced .app would be pushed to the Windows/Linux machines as a junk
directory, and macOS re-signs bundles in place, which fights the sync client.
It holds no logic of its own; it just runs the launcher script where the
project actually lives, so editing the project needs no rebuild of the app.

Run directly, or let setup.sh call it:  python3 make_mac_app.py
"""
from __future__ import annotations

import os
import plistlib
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

import config as C

APP_NAME = C.APP_NAME
BUNDLE_ID = f"local.coursedashboard.{C.COURSE_CODE.lower()}"
PROJECT = Path(__file__).resolve().parent
APPS_DIR = Path.home() / "Applications"

# Same palette as the dashboard's hall plans: dark slate ground, amber seat.
BG_TOP = (24, 26, 33)
BG_BOTTOM = (38, 41, 52)
SEAT_ON = (232, 168, 66)
SEAT_OFF = (74, 80, 96)
TEXT = (240, 240, 245)


def draw_icon(size: int):
    """A 1024-safe icon: seat grid over a dark rounded square, 'PHL' beneath.

    Drawn rather than shipped as a binary blob so the icon lives in the repo
    as readable source, and so it can be re-rendered at any size iconutil
    asks for without a resampling step.
    """
    from PIL import Image, ImageDraw, ImageFont

    S = 1024  # draw big, downsample once -- cheap and much cleaner than
    img = Image.new("RGBA", (S, S), (0, 0, 0, 0))  # hinting tiny primitives
    d = ImageDraw.Draw(img)

    # Vertical gradient inside a rounded-square mask.
    grad = Image.new("RGB", (1, S))
    for y in range(S):
        t = y / (S - 1)
        grad.putpixel((0, y), tuple(round(a + (b - a) * t)
                                    for a, b in zip(BG_TOP, BG_BOTTOM)))
    grad = grad.resize((S, S))
    mask = Image.new("L", (S, S), 0)
    ImageDraw.Draw(mask).rounded_rectangle([0, 0, S - 1, S - 1], radius=int(S * 0.22), fill=255)
    img.paste(grad, (0, 0), mask)

    # Seat grid: 6 columns x 4 rows, a scaled-down hall plan. The filled ones
    # spell out nothing in particular -- they just read as "seats allocated".
    cols, rows = 6, 4
    pad_x, top, gap = S * 0.16, S * 0.20, S * 0.035
    cell = (S - 2 * pad_x - gap * (cols - 1)) / cols
    filled = {(0, 0), (1, 0), (2, 0), (4, 0),
              (0, 1), (2, 1), (3, 1), (5, 1),
              (1, 2), (2, 2), (4, 2), (5, 2),
              (0, 3), (3, 3), (4, 3)}
    for r in range(rows):
        for c in range(cols):
            x0 = pad_x + c * (cell + gap)
            y0 = top + r * (cell + gap)
            d.rounded_rectangle([x0, y0, x0 + cell, y0 + cell],
                                radius=cell * 0.28,
                                fill=SEAT_ON if (c, r) in filled else SEAT_OFF)

    # Wordmark. DejaVuSans ships with Pillow; fall back to the default bitmap
    # font on the odd install that lacks it rather than failing the build.
    label = C.COURSE_CODE[:8]
    try:
        font = ImageFont.truetype("DejaVuSans-Bold.ttf", int(S * 0.115))
    except OSError:
        try:
            font = ImageFont.truetype("/System/Library/Fonts/Helvetica.ttc", int(S * 0.115))
        except OSError:
            font = ImageFont.load_default()
    box = d.textbbox((0, 0), label, font=font)
    d.text(((S - (box[2] - box[0])) / 2 - box[0], S * 0.775), label, font=font, fill=TEXT)

    return img.resize((size, size), Image.LANCZOS) if size != S else img


def build_icns(dest: Path) -> bool:
    """Render an .icns via iconutil. Returns False if it couldn't be made."""
    try:
        from PIL import Image  # noqa: F401
    except ImportError:
        print("Pillow not available -- app will use the generic icon.")
        return False
    if not shutil.which("iconutil"):
        print("iconutil not found (needs Xcode command line tools) -- generic icon.")
        return False

    with tempfile.TemporaryDirectory() as tmp:
        iconset = Path(tmp) / "AppIcon.iconset"
        iconset.mkdir()
        # The exact filename set `iconutil` expects; @2x entries are the same
        # pixel size as the next tier up, just labelled for Retina.
        for px in (16, 32, 128, 256, 512):
            draw_icon(px).save(iconset / f"icon_{px}x{px}.png")
            draw_icon(px * 2).save(iconset / f"icon_{px}x{px}@2x.png")
        dest.parent.mkdir(parents=True, exist_ok=True)
        subprocess.run(["iconutil", "-c", "icns", str(iconset), "-o", str(dest)], check=True)
    return True


def build_app() -> Path:
    app = APPS_DIR / f"{APP_NAME}.app"
    macos = app / "Contents" / "MacOS"
    resources = app / "Contents" / "Resources"

    # Rebuild from scratch: a half-updated bundle (old plist, new binary) is
    # the kind of thing macOS caches and then refuses to launch.
    if app.exists():
        shutil.rmtree(app)
    macos.mkdir(parents=True)
    resources.mkdir(parents=True)

    launcher = macos / APP_NAME
    # Hand the script to Terminal.app instead of running it here. The project
    # lives under ~/Library/CloudStorage (Dropbox), which macOS protects with
    # TCC: an app launched from Finder gets `deny file-read-data` on the
    # launcher script -- silently, with no consent prompt, because this bundle
    # has no Full Disk Access and an ad-hoc signature doesn't earn one.
    # Terminal already holds that access, so it can read what this app cannot.
    # (If you'd rather have no Terminal window: grant this .app Full Disk
    # Access in System Settings > Privacy & Security, then change the line
    # below to exec launch_dashboard.sh directly.)
    launcher.write_text(
        "#!/bin/sh\n"
        "# Generated by make_mac_app.py -- edit the project, not this file.\n"
        f'exec /usr/bin/open -a Terminal "{PROJECT}/launch_dashboard.sh"\n'
    )
    launcher.chmod(0o755)

    has_icon = build_icns(resources / "AppIcon.icns")

    info = {
        "CFBundleName": APP_NAME,
        "CFBundleDisplayName": APP_NAME,
        "CFBundleIdentifier": BUNDLE_ID,
        "CFBundleVersion": "1.0",
        "CFBundleShortVersionString": "1.0",
        "CFBundlePackageType": "APPL",
        "CFBundleExecutable": APP_NAME,
        "LSMinimumSystemVersion": "10.13",
        # This app hands off to Terminal and exits immediately; the dashboard
        # server outlives it (nohup'd inside launch_dashboard.sh), the same as
        # on Linux and Windows.
        "LSBackgroundOnly": False,
        "NSHighResolutionCapable": True,
    }
    if has_icon:
        info["CFBundleIconFile"] = "AppIcon"
    (app / "Contents" / "Info.plist").write_bytes(plistlib.dumps(info))

    # Nudge Launch Services so the icon and Spotlight entry appear now rather
    # than whenever it next rescans ~/Applications.
    lsregister = ("/System/Library/Frameworks/CoreServices.framework/Frameworks"
                  "/LaunchServices.framework/Support/lsregister")
    if os.path.exists(lsregister):
        subprocess.run([lsregister, "-f", str(app)], check=False)

    return app


def main() -> int:
    if sys.platform != "darwin":
        print("make_mac_app.py is macOS-only; nothing to do here.")
        return 0
    if not (PROJECT / "launch_dashboard.sh").exists():
        print("launch_dashboard.sh missing next to this script.", file=sys.stderr)
        return 1
    app = build_app()
    print(f"Created {app}")
    print(f"  - Spotlight: type '{C.COURSE_CODE}'")
    print("  - Finder: Go > Applications (or ~/Applications) -- drag it to the Dock to pin it")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
