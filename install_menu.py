"""Add/remove the Explorer right-click 'Convert to' menu (current user only, HKCU).
Usage: python install_menu.py            install
       python install_menu.py --uninstall
"""
import sys
import winreg
from pathlib import Path

from ezconvert import IMG, TXT, targets_for

SCRIPT = Path(__file__).with_name("ezconvert.py").resolve()
PYW = Path(sys.executable).with_name("pythonw.exe")
BASE = r"Software\Classes\SystemFileAssociations"
EXTS = sorted(IMG | TXT | {".pdf", ".docx"})


def set_val(path, value, name=""):
    with winreg.CreateKey(winreg.HKEY_CURRENT_USER, path) as k:
        winreg.SetValueEx(k, name, 0, winreg.REG_SZ, value)


def delete_tree(path):
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, path, 0, winreg.KEY_ALL_ACCESS) as k:
            while True:
                try:
                    delete_tree(path + "\\" + winreg.EnumKey(k, 0))
                except OSError:
                    break
        winreg.DeleteKey(winreg.HKEY_CURRENT_USER, path)
    except FileNotFoundError:
        pass


for ext in EXTS:
    root = f"{BASE}\\{ext}\\shell\\EZConvert"
    delete_tree(root)
    if "--uninstall" in sys.argv:
        continue
    set_val(root, "Convert to", "MUIVerb")
    set_val(root, "", "SubCommands")
    for fmt in targets_for(ext):
        sub = f"{root}\\shell\\{fmt}"
        set_val(sub, fmt.upper(), "MUIVerb")
        set_val(sub + "\\command", f'"{PYW}" "{SCRIPT}" --to {fmt} "%1"')
print("removed" if "--uninstall" in sys.argv else f"installed for {len(EXTS)} file types")
