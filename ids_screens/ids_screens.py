from pathlib import Path
import win32api

from pywinauto import Application
import yaml

from .apps.ids import shoot_ids
from .apps.rr_viewer import shoot_rr_viewer
from ._pywinauto.patcher import patch as pywinauto_patch


def main():
    pywinauto_patch()

    # удаляем старые файлы (все bmp-файлы в текущей папке)
    for p in Path(".").glob("*.bmp"):
        p.unlink()

    with open("ids_screens.yaml", encoding="UTF-8") as f:
        conf = yaml.safe_load(f)

    start_mouse_pos = win32api.GetCursorPos()
    # для белого фона на скриншотах в Win7
    bg = Application(backend="win32").start("notepad.exe")
    bg.Notepad.maximize()

    shoot_ids(conf["ids"])
    shoot_rr_viewer(conf["rr_viewer"])

    bg.Notepad.close()
    win32api.SetCursorPos(start_mouse_pos)
