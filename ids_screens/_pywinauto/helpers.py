def treeview_element_colapse(element) -> None:
    # element: _treeview_element
    # если вызывать просто colapse, м. не перестроится контекстное меню
    element.click_input()
    element.collapse()
