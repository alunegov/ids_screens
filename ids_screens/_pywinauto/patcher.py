import time

from pywinauto.controls.menuwrapper import Menu
from pywinauto.controls.win32_controls import PopupMenuWrapper
from pywinauto.mouse import move as mouse_move
from pywinauto.win32structures import RECT


def _menu_highlight_and_capture(self: Menu, path: str, filename: str = None, rect: RECT = None) -> None:
    items = self.get_menu_path(path)
    items[0].click_input()  # чтобы открылось меню
    for i in range(1, len(items)):
        r = items[i].rectangle()
        mouse_move((r.left, r.top))
        time.sleep(.5)
    if filename is not None:
        self.ctrl.capture_as_image(rect).save(filename)


def _popup_highlight_and_capture(self: PopupMenuWrapper, path: str, filename: str = None, rect: RECT = None) -> None:
    items = self.menu().get_menu_path(path)
    for i in range(0, len(items)):
        r = items[i].rectangle()
        mouse_move((r.left, r.top))
        time.sleep(.5)
    if filename is not None:
        self.capture_as_image(rect).save(filename)


def patch():
    Menu.highlight_and_capture = _menu_highlight_and_capture
    PopupMenuWrapper.highlight_and_capture = _popup_highlight_and_capture
