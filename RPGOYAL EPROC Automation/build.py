"""
Build script for RPGOYAL eProcurement Tool.

Usage (from within the project venv):
    .venv/bin/python build.py          # Mac / Linux
    .venv\\Scripts\\python build.py     # Windows

Produces a single executable in dist/:
    Mac   - dist/RPGOYAL_eProcurement_Tool
    Win   - dist/RPGOYAL_eProcurement_Tool.exe

Prerequisites:
    pip install -r requirements.txt
"""

import os
import sys
import platform
import subprocess

HERE = os.path.dirname(os.path.abspath(__file__))


def _find_python():
    """Return the best available Python — prefer the project venv."""
    if platform.system() == "Windows":
        venv_py = os.path.join(HERE, ".venv", "Scripts", "python.exe")
    else:
        venv_py = os.path.join(HERE, ".venv", "bin", "python3")
    if os.path.isfile(venv_py):
        return venv_py
    return sys.executable


def main():
    python = _find_python()

    # Pre-check: make sure PyInstaller is available in the chosen interpreter
    check = subprocess.run(
        [python, "-c", "import PyInstaller"],
        capture_output=True,
    )
    if check.returncode != 0:
        print("ERROR: PyInstaller is not installed in the target Python.")
        print(f"  Python: {python}")
        print()
        print("Fix: install dependencies into the project venv first:")
        if platform.system() == "Windows":
            print("  .venv\\Scripts\\pip install -r requirements.txt")
        else:
            print("  .venv/bin/pip install -r requirements.txt")
        sys.exit(1)

    sep = ";" if platform.system() == "Windows" else ":"

    pin_codes = os.path.join(HERE, "List of Pin Codes of Rajasthan.xlsx")
    add_data = []
    if os.path.isfile(pin_codes):
        add_data.append(f"--add-data={pin_codes}{sep}.")

    # CustomTkinter stores assets (themes, icons) that PyInstaller misses
    ctk_check = subprocess.run(
        [python, "-c",
         "import customtkinter, os;"
         "print(os.path.dirname(customtkinter.__file__))"],
        capture_output=True, text=True,
    )
    if ctk_check.returncode == 0:
        ctk_path = ctk_check.stdout.strip()
        add_data.append(f"--add-data={ctk_path}{sep}customtkinter")
    else:
        print("WARNING: customtkinter not found; build may fail.")

    is_mac = platform.system() == "Darwin"
    pack_mode = "--onedir" if is_mac else "--onefile"

    cmd = [
        python, "-m", "PyInstaller",
        pack_mode,
        "--windowed",
        "--name", "RPGOYAL_eProcurement_Tool",
        "--hidden-import", "customtkinter",
        "--hidden-import", "openpyxl",
        "--hidden-import", "xlrd",
        "--hidden-import", "reportlab",
        "--hidden-import", "webdriver_manager",
        *add_data,
        os.path.join(HERE, "launcher.py"),
    ]

    print("Running PyInstaller:")
    print(f"  Python: {python}")
    print(f"  Command: {' '.join(cmd)}")
    print()

    subprocess.check_call(cmd, cwd=HERE)

    print()
    print("Build complete.")
    if is_mac:
        print("  App bundle: dist/RPGOYAL_eProcurement_Tool.app")
        print("  (double-click the .app to launch)")
    elif platform.system() == "Windows":
        print("  Executable: dist/RPGOYAL_eProcurement_Tool.exe")
    else:
        print("  Executable: dist/RPGOYAL_eProcurement_Tool")


if __name__ == "__main__":
    main()
