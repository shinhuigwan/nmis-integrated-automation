"""Create portable app shortcuts once, including after the folder is moved."""
import json
from pathlib import Path

SHORTCUT_NAMES = ("통합 자동화 시스템.lnk", "NMIS Integrated Automation.lnk")


def ensure_desktop_shortcuts(app_root, *, desktop=None, shell=None, force=False):
    root = Path(app_root).resolve()
    exe = root / "NMIS-Automation.exe"
    if not exe.is_file():
        raise FileNotFoundError("패키징 실행 파일을 찾지 못했습니다.")
    marker = root / "data/desktop_shortcuts.json"
    if not force and marker.is_file():
        try:
            if json.loads(marker.read_text(encoding="utf-8")).get("exe") == str(exe):
                return False
        except (ValueError, OSError):
            pass
    if shell is None:
        import win32com.client
        shell = win32com.client.Dispatch("WScript.Shell")
    desktop = Path(desktop or shell.SpecialFolders("Desktop"))
    for name in SHORTCUT_NAMES:
        link = shell.CreateShortcut(str(desktop / name))
        link.TargetPath = str(exe)
        link.Arguments = ""
        link.WorkingDirectory = str(root)
        link.IconLocation = str(exe) + ",0"
        link.Description = "NMIS 통합 자동화 시스템"
        link.Save()
    marker.parent.mkdir(parents=True, exist_ok=True)
    marker.write_text(json.dumps({"exe": str(exe)}, ensure_ascii=False), encoding="utf-8")
    return True
