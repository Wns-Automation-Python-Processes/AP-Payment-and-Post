"""Build a standalone Windows executable for the AP Payment Process UI.

Run from the project folder with the project virtual environment:

    .\\venv\\Scripts\\python.exe build_exe.py

Produces  dist\\AP Payment Process\\AP Payment Process.exe  (folder build,
faster to start) plus a zip-friendly folder you can copy to any PC.
"""
import os
import shutil
import subprocess
import sys

BASE = os.path.dirname(os.path.abspath(__file__))
APP_NAME = "AP Payment Process"
ENTRY = os.path.join(BASE, "sagex3_ui.py")
ICON = os.path.join(BASE, "assets", "tiauto.ico")
ASSETS = os.path.join(BASE, "assets")
DIST = os.path.join(BASE, "dist")
BUILD = os.path.join(BASE, "build")


def clean():
    for d in (DIST, BUILD):
        if os.path.isdir(d):
            shutil.rmtree(d, ignore_errors=True)


def main():
    if not os.path.exists(ENTRY):
        sys.exit(f"Entry script not found: {ENTRY}")

    # Make sure the logo/icon assets exist before bundling.
    if not os.path.exists(ICON):
        print("[build] assets/tiauto.ico missing - generating logos first...")
        subprocess.check_call([sys.executable, os.path.join(BASE, "make_logo.py")])

    clean()

    args = [
        sys.executable, "-m", "PyInstaller",
        "--noconfirm",
        "--clean",
        "--onefile",                      # single file - works better from network/local
        "--windowed",                     # GUI app: no console window
        "--name", APP_NAME,
        "--icon", ICON,
        "--add-data", f"{ASSETS}{os.pathsep}assets",
        # Selenium loads many submodules lazily (via __getattr__), so pull the
        # whole package in, not just a hand-picked list.
        "--collect-all", "selenium",
        "--hidden-import", "selenium",
        "--hidden-import", "selenium.webdriver",
        "--hidden-import", "selenium.webdriver.chrome.webdriver",
        "--hidden-import", "selenium.webdriver.chrome.service",
        "--hidden-import", "selenium.webdriver.common.by",
        "--hidden-import", "selenium.webdriver.common.keys",
        "--hidden-import", "selenium.webdriver.support.ui",
        "--hidden-import", "selenium.webdriver.support.expected_conditions",
        "--hidden-import", "selenium.common.exceptions",
        "--collect-all", "webdriver_manager",
        "--hidden-import", "openpyxl",
        "--hidden-import", "openpyxl.styles",
        "--hidden-import", "PIL",
        "--hidden-import", "PIL.ImageTk",
        # Trim obvious heavyweight junk we do not use.
        "--exclude-module", "matplotlib",
        "--exclude-module", "scipy",
        "--exclude-module", "PyQt5",
        "--exclude-module", "PySide2",
        "--exclude-module", "pandas",
        ENTRY,
    ]

    print("[build] Running PyInstaller (onefile)...")
    subprocess.check_call(args, cwd=BASE)

    exe = os.path.join(DIST, f"{APP_NAME}.exe")
    if os.path.exists(exe):
        size = os.path.getsize(exe) / 1024 / 1024
        print(f"\n[build] Done -> {exe}  ({size:.1f} MB)")
        print("[build] This is a single self-contained file. Copy just this .exe to the other PC.")
        print("[build] Put 'InPut_Parameter_File.xlsx' next to the .exe if you "
              "want it picked up on startup.")
    else:
        sys.exit("[build] FAILED: executable was not produced.")


if __name__ == "__main__":
    main()
