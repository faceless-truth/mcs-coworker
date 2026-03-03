"""
MC & S Desktop Agent — Desktop Shortcut Creator
Creates a shortcut on the Windows desktop with the MC & S icon.
Run once after installation: python create_shortcut.py
"""
import os
import sys


def create_shortcut():
    try:
        import winshell
        from win32com.client import Dispatch
    except ImportError:
        # Install required packages
        import subprocess
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pywin32", "winshell", "--quiet"])
        import winshell
        from win32com.client import Dispatch

    desktop = winshell.desktop()
    shortcut_path = os.path.join(desktop, "MC & S Coworker.lnk")
    app_dir = os.path.dirname(os.path.abspath(__file__))
    icon_path = os.path.join(app_dir, "assets", "icon.ico")
    vbs_path = os.path.join(app_dir, "launch_silent.vbs")

    shell = Dispatch("WScript.Shell")
    shortcut = shell.CreateShortCut(shortcut_path)
    shortcut.Targetpath = vbs_path
    shortcut.WorkingDirectory = app_dir
    shortcut.IconLocation = icon_path
    shortcut.Description = "MC & S Coworker — Desktop Agent"
    shortcut.save()

    print(f"✓ Desktop shortcut created: {shortcut_path}")
    print(f"  Icon: {icon_path}")
    print(f"  Target: {vbs_path}")


if __name__ == "__main__":
    create_shortcut()
