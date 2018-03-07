import win32api

import time

from pywinauto.application import Application, ProcessNotFoundError
from pywinauto.controls.menuwrapper import Menu
from pywinauto.controls.win32_controls import PopupMenuWrapper
from pywinauto.keyboard import SendKeys
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


def _patch():
    # pywinauto patching
    Menu.highlight_and_capture = _menu_highlight_and_capture
    PopupMenuWrapper.highlight_and_capture = _popup_highlight_and_capture


def _treeview_element_colapse(element) -> None:
    # element: _treeview_element
    # если вызывать просто colapse, м. не перестроится контекстное меню
    element.click_input()
    element.collapse()


def _get_corepopup_rect(popup) -> RECT:
    # popup: WindowSpecification
    rect = popup.rectangle()
    if popup.menu().item_count() > 17:
        r = popup.menu().item(17).rectangle()  # Диагностика
        rect.bottom = r.bottom + 5
    return rect


def shoot_ids():
    # conf
    delay_win_active = .5

    ids_path = "C:\Work\IDS-l10n\.out\Debug\IDS.exe"
    ids_svk_import_file = "D:\Pub\__IDS master image\IDS\Предприятие Подшипники Демо 01.03.2010.rar"
    expert_system = "_SVK"

    screen_rect = RECT(0, 0, win32api.GetSystemMetrics(0), win32api.GetSystemMetrics(1))
    app_created = False

    try:
        app = Application(backend="win32").connect(path=ids_path)
        app.TApplication.set_focus() 
    except ProcessNotFoundError:
        app = Application(backend="win32").start(ids_path)
        app.wait_cpu_usage_lower()
        app_created = True

    ids = app.window(class_name="TIDSMainForm")

    ids.move_window(200, 200, 600, 400)
    app.wait_cpu_usage_lower()

    # определяем индекс пункта меню Сервис->Параметры
    params_index = ids.menu_item("#1").sub_menu().item_count() - 1

    # MainWindow.bmp
    _treeview_element_colapse(ids.TreeView.tree_root())
    ids.capture_as_image().save("MainWindow{}.bmp".format(expert_system))

    # MainWindow_MainMenu.bmp
    _treeview_element_colapse(ids.TreeView.tree_root())
    # Сервис->Параметры->xxx
    ids.menu().highlight_and_capture("#1->#{}".format(params_index), "MainWindow_MainMenu{}.bmp".format(expert_system))
    SendKeys("{ESC}{ESC}{ESC}")  # закрываем меню

    # MainWindow_NewObject_Call.bmp
    ids.TreeView.get_item((0, 0)).click_input(button="right")
    r = _get_corepopup_rect(app.PopupMenu)
    app.PopupMenu.wrapper_object().highlight_and_capture("#0", "MainWindow_NewObject_Call.bmp", r)  # Добавить объект
    SendKeys("{ESC}")  # закрываем меню

    # MainWindow_NewObject.bmp
    ids.TreeView.get_item((0, 0)).click_input(button="right")
    app.PopupMenu.menu_item("#0").click_input()  # Добавить объект
    time.sleep(delay_win_active)
    app.TStructureAddForm.capture_as_image().save("MainWindow_NewObject.bmp")
    app.TStructureAddForm.close_alt_f4()

    # MainWindow_NewZamer_Call.bmp
    ids.TreeView.get_item((0, "Подшипники Демо", "Внешняя обойма")).click_input(button="right")
    r = _get_corepopup_rect(app.PopupMenu)
    app.PopupMenu.wrapper_object().highlight_and_capture("#2", "MainWindow_NewZamer_Call.bmp", r)  # Добавить измерение
    SendKeys("{ESC}")  # закрываем меню

    # ZamerReg_SVK.bmp
    ids.TreeView.get_item((0, "Подшипники Демо", "Внешняя обойма")).click_input(button="right")
    app.PopupMenu.menu_item("#2").click_input()  # Добавить измерение
    time.sleep(7)  # идёт инициализация АЦП
    if app.TMessageForm.exists():
        # нет АЦП, условий в МВИ и т.п.
        print("skip ZamerReg_SVK.bmp")
        SendKeys("{ESC}")  # закрываем сообщение
    else:
        # ищем без best_match, на случай если юзер сам закрыл окно предупреждения
        app.window(class_name="TSVKSignalForm").capture_as_image().save("ZamerReg_SVK.bmp")
        app.TSVKSignalForm.close_alt_f4()
        app.wait_cpu_usage_lower()

    # MainWindow_ZamerFormirov_SVK.bmp
    # включение ручного формирования
    ids.menu_select("#1->#{}->#0".format(params_index))  # Сервис->Параметры->Общие
    app.TParamsForm["ручное"].check_by_click()  # TODO: l10n
    app.TParamsForm.OkButton.click()
    #
    ids.TreeView.get_item((0, "Подшипники Демо", "Внешняя обойма")).click_input(button="right")
    app.PopupMenu.menu_item("#2").click_input()  # Добавить измерение
    time.sleep(delay_win_active)
    app.TZamerFormirovForm.capture_as_image().save("MainWindow_ZamerFormirov_SVK.bmp")
    app.TZamerFormirovForm.close_alt_f4()
    # включение автоматического формирования
    ids.menu_select("#1->#{}->#0".format(params_index))  # Сервис->Параметры->Общие
    app.TParamsForm["автоматическое"].check_by_click()  # TODO: l10n
    app.TParamsForm.OkButton.click()

    # Params_RB.bmp
    # TODO: l10n
    ids.menu_select("#1->#{}->Подшипники качения".format(params_index))  # Сервис->Параметры->Подшипники качения
    # TODO: дефолтные параметры
    # не находит по TSysParamsEditingForm, используем top_window()
    app.top_window().capture_as_image().save("Params_RB.bmp")
    app.top_window().close_alt_f4()

    # Params_RB_SVK.bmp
    ids.menu_select("#1->#{}->#0".format(params_index))  # Сервис->Параметры->Общие
    SendKeys("{TAB}{TAB}{RIGHT}{RIGHT}")  # переключаемся на закладку Подшипники
    # TODO: дефолтные параметры
    app.TParamsForm.capture_as_image().save("Params_RB_SVK.bmp")
    app.TParamsForm.close_alt_f4()

    # MainWindow_PSP_Call.bmp
    ids.TreeView.get_item((0, "Подшипники Демо", "Внешняя обойма")).click_input(button="right")
    r = _get_corepopup_rect(app.PopupMenu)
    app.PopupMenu.wrapper_object().highlight_and_capture("#12", "MainWindow_PSP_Call.bmp", r)  # Паспорт
    SendKeys("{ESC}")  # закрываем меню

    # PSP_RB.bmp
    ids.TreeView.get_item((0, "Подшипники Демо", "Внешняя обойма")).click_input(button="right")
    app.PopupMenu.menu_item("#12").click_input()  # Паспорт
    time.sleep(delay_win_active)
    # не находит по TPSP_RollingBearingForm, используем top_window()
    app.top_window().capture_as_image().save("PSP_RB.bmp")
    app.top_window().close_alt_f4()

    # паспорт
    ids.TreeView.get_item((0, "Подшипники Демо", "Внешняя обойма")).click_input(button="right")
    app.PopupMenu.menu_item("#12").click_input()  # Паспорт
    time.sleep(delay_win_active)
    # не находит по TPSP_RollingBearingForm, используем top_window()
    app.top_window().CheckBox1.click()  # инвертируем состояние Подшипник из нержавеющей...
    app.top_window().TBitBtn3.click()  # Применить
    time.sleep(delay_win_active)
    # подтверждение сохранения паспорта
    app.TMessageForm.capture_as_image().save("PSP_SaveConfirm.bmp")
    app.TMessageForm.Button3.click()  # Да
    time.sleep(delay_win_active)
    # подтверждение сохранения марки
    app.TMessageForm.capture_as_image().save("PSP_MarkaSaveConfirm.bmp")
    # TODO: в окне выводится название марки
    app.TMessageForm.Button2.click()  # Да
    time.sleep(delay_win_active)
    # инфо о существовании похожей марки (opt)
    if app.TMessageForm.exists() and app.TMessageForm.control_count() == 1:
        SendKeys("{ESC}")  # закрываем сообщение
        time.sleep(delay_win_active)
    # подтверждение изменения марок в замерах (opt)
    if app.TMessageForm.exists():
        app.TMessageForm.capture_as_image().save("PSP_MarkaSaveInZamersConfirm.bmp")
        SendKeys("{ESC}")  # закрываем сообщение
    else:
        print("skip PSP_MarkaSaveInZamersConfirm.bmp")

    # PSP_Marka_RB.bmp
    ids.menu_select("#1->#0")  # Сервис->Редактор БД
    # не находит по TDBEditForm, используем top_window()
    app.top_window().TreeView.select((0, "Подшипник качения"))  # TODO: l10n
    app.wait_cpu_usage_lower()
    app.top_window().TListView.double_click(coords=(5, 5))  # открываем окно первой марки в списке
    time.sleep(2 * delay_win_active)
    app.top_window().capture_as_image().save("PSP_Marka_RB.bmp")
    app.top_window().close_alt_f4()  # окно марки
    time.sleep(delay_win_active)
    app.top_window().close_alt_f4()  # окно Редактор БД
    app.wait_cpu_usage_lower()

    # Zamer_Show_SVK.bmp
    ids.TreeView.get_item((0, "Подшипники Демо", "Внешняя обойма", 0, 0)).click_input(button="right")
    r = _get_corepopup_rect(app.PopupMenu)
    app.PopupMenu.wrapper_object().highlight_and_capture("#14", "Zamer_Show_SVK.bmp", r)  # Просмотр
    SendKeys("{ESC}")  # закрываем меню

    # Zamer_Show_RollingBearing.bmp
    ids.TreeView.get_item((0, "Подшипники Демо", "Внешняя обойма", 0, 0)).click_input(double=True)
    time.sleep(delay_win_active)
    # подгоняем размер чтобы не было верт. скрола и окно целиком влазило на экран
    app.TSignalShow_RollingBearingForm.move_window(100, 0, 800, 970)
    # TODO: дефолтные параметры
    app.TSignalShow_RollingBearingForm.capture_as_image().save("Zamer_Show_RollingBearing.bmp")
    app.TSignalShow_RollingBearingForm.close_alt_f4()

    # Zamer_Cascade_Call.bmp
    ids.TreeView.get_item((0, "Подшипники Демо", "Внешняя обойма")).click_input(button="right")
    r = _get_corepopup_rect(app.PopupMenu)
    app.PopupMenu.wrapper_object().highlight_and_capture("#15", "Zamer_Cascade_Call.bmp", r)  # Просмотр каскада
    SendKeys("{ESC}")  # закрываем меню

    # Zamer_Cascade_SelectFiles{}.bmp
    # Zamer_Cascade{}.bmp
    ids.TreeView.get_item((0, "Подшипники Демо", "Внешняя обойма")).click_input(button="right")
    app.PopupMenu.menu_item("#15").click_input()  # Просмотр каскада
    time.sleep(delay_win_active)
    app.TShowCascadeInitForm.capture_as_image().save("Zamer_Cascade_SelectFiles{}.bmp".format(expert_system))
    # TODO: выбрать замеры
    app.TShowCascadeInitForm.Button2.click()  # Просмотр
    if app.TMessageForm.exists():
        print("skip Zamer_Cascade{}.bmp".format(expert_system))
        SendKeys("{ESC}")  # закрываем сообщение
        app.TShowCascadeInitForm.close_alt_f4()
    else:
        app.TShowCascadeForm.capture_as_image().save("Zamer_Cascade{}.bmp".format(expert_system))
        app.TShowCascadeForm.close_alt_f4()

    # Zamer_ShowMarka_Call.bmp
    ids.TreeView.get_item((0, "Подшипники Демо", "Внешняя обойма", 0, 0)).click_input(button="right")
    r = _get_corepopup_rect(app.PopupMenu)
    app.PopupMenu.wrapper_object().highlight_and_capture("#16", "Zamer_ShowMarka_Call.bmp", r)  # Измерение. Просмотр марки
    SendKeys("{ESC}")  # закрываем меню

    # Zamer_Diagnos_Call.bmp
    ids.TreeView.get_item((0, "Подшипники Демо", "Внешняя обойма", 0, 0)).click_input(button="right")
    r = _get_corepopup_rect(app.PopupMenu)
    app.PopupMenu.wrapper_object().highlight_and_capture("#17", "Zamer_Diagnos_Call.bmp", r)  # Диагностика
    SendKeys("{ESC}")  # закрываем меню

    # Diagnoz_RB.bmp
    ids.TreeView.get_item((0, "Подшипники Демо", "Внешняя обойма", 0, 0)).click_input(button="right")
    app.PopupMenu.menu_item("#17").click_input()  # Диагностика
    time.sleep(delay_win_active)
    if app.TMessageForm.exists():
        if app.TMessageForm.control_count() == 2:
            # предупреждение о нормах и т.п.
            app.TMessageForm.Button2.click()  # Да
            time.sleep(delay_win_active)
        else:
            # ошибка диагностики
            SendKeys("{ESC}")  # закрываем сообщение
    if app.TRollingBearing_DiagnosForm.exists():
        app.TRollingBearing_DiagnosForm.capture_as_image().save("Diagnoz_RB.bmp")
        # TODO: картинки примеров отчётов (из браузера)
        # app.TRollingBearing_DiagnosForm.TBitBtn1.click()  # Отчёт
        app.TRollingBearing_DiagnosForm.close_alt_f4()
    else:
        print("skip Diagnoz_RB.bmp")

    # TODO: спец. функции. Желательно, чтобы меню "уходило" вверх - м.б. сместить окно в низ экрана
    ids.TreeView.get_item((0, "Подшипники Демо", "Внешняя обойма", 0, 0)).click_input(button="right")
    # Подшипники качения->Фактический зазор
    if app.PopupMenu.menu().item_count() > 21:
        app.PopupMenu.wrapper_object().highlight_and_capture("#21->#2", "EDIT_SF_RB_Call.bmp", screen_rect)
        # TODO: обрезать лишнее
        SendKeys("{ESC}{ESC}")  # закрываем меню
    else:
        print("skip SF_RB_Call.bmp")
        SendKeys("{ESC}")  # закрываем меню

    # Zamer_Extra_SVK_CommonDiag.bmp
    ids.TreeView.get_item((0, "Подшипники Демо", "Внешняя обойма")).click_input(button="right")
    app.PopupMenu.menu_item("#19->#0").click_input()
    time.sleep(delay_win_active)
    app.TRollingBearing_CommonDiagIntro_Form.capture_as_image().save("Zamer_Extra_SVK_CommonDiag.bmp")
    # TODO: картинка примера отчёта (из браузера)
    # app.TRollingBearing_CommonDiagIntro_Form.Button2.click()  # Ok
    app.TRollingBearing_CommonDiagIntro_Form.close_alt_f4()

    # MainWindow_Export_Call.bmp
    _treeview_element_colapse(ids.TreeView.tree_root())
    ids.menu().highlight_and_capture("#1->#4", "MainWindow_Export_Call.bmp")  # Сервис->Экспорт
    SendKeys("{ESC}{ESC}")  # закрываем меню

    # Export_Params.bmp
    _treeview_element_colapse(ids.TreeView.tree_root())
    ids.menu_select("#1->#4")  # Сервис->Экспорт
    time.sleep(delay_win_active)
    app.TExportInitForm.capture_as_image().save("Export_Params.bmp")
    app.TExportInitForm.close_alt_f4()

    # MainWindow_Import_Call.bmp
    _treeview_element_colapse(ids.TreeView.tree_root())
    ids.menu().highlight_and_capture("#1->#5", "MainWindow_Import_Call.bmp")  # Сервис->Импорт
    SendKeys("{ESC}{ESC}")  # закрываем меню

    # Import_Params.bmp
    _treeview_element_colapse(ids.TreeView.tree_root())
    ids.menu_select("#1->#5")  # Сервис->Импорт
    time.sleep(delay_win_active)
    app.TImportInitForm.capture_as_image().save("Import_Params.bmp")
    app.TImportInitForm.close_alt_f4()

    # Import_Manual{}.bmp
    _treeview_element_colapse(ids.TreeView.tree_root())
    ids.menu_select("#1->#5")  # Сервис->Импорт
    app.TImportInitForm["...Button"].click()  # выбор файла
    app.DialogEdit.Edit.set_edit_text(ids_svk_import_file)
    app.DialogEdit.Button1.click_input()  # Открыть
    if app.DialogEdit.exists() or app.Dialog.exists():
        # ошибка открытия файла (иногда окно предупреждения не появляется)
        print("skip Import_Manual{}.bmp".format(expert_system))
        SendKeys("{ESC}{ESC}{ESC}")  # закрываем все окна вплоть до IDS
    else:
        app.TImportInitForm.Button3.click()  # Ok
        app.wait_cpu_usage_lower()
        app.TImportTreeForm.move_window(450, 270, 330, 300)  # в границах IDS, справа
        ids.capture_as_image().save("Import_Manual{}.bmp".format(expert_system))
        app.TImportTreeForm.close_alt_f4()

    # Import_Params_Extra.bmp
    _treeview_element_colapse(ids.TreeView.tree_root())
    ids.menu_select("#1->#5")  # Сервис->Импорт
    app.TImportInitForm.Button1.click()  # Настройки
    time.sleep(delay_win_active)
    app.TImportSettingsForm.capture_as_image().save("Import_Params_Extra.bmp")
    SendKeys("{ESC}{ESC}")  # закрываем меню

    # DBEditor_Markas{}.bmp
    ids.menu_select("#1->#0")  # Сервис->Редактор БД
    # не находит по TDBEditForm, используем top_window()
    app.top_window().capture_as_image().save("DBEditor_Markas{}.bmp".format(expert_system))
    app.top_window().close_alt_f4()

    # Params_Common{}.bmp
    # Params_Common2{}.bmp
    # Params_RB{}.bmp
    # Params_UserPointNames{}.bmp
    # TODO: дефолтные параметры
    ids.menu_select("#1->#{}->#0".format(params_index))  # Сервис->Параметры->Общие
    app.TParamsForm.capture_as_image().save("Params_Common{}.bmp".format(expert_system))
    SendKeys("{TAB}{TAB}{RIGHT}")  # переключаемся на закладку Общие2
    app.TParamsForm.capture_as_image().save("Params_Common2{}.bmp".format(expert_system))
    # TODO: третьей закладки м. не быть, также это м.б. Формирование измерения
    SendKeys("{RIGHT}")  # переключаемся на закладку Подшипники
    app.TParamsForm.capture_as_image().save("Params_RB{}.bmp".format(expert_system))
    # TODO: закладки Формирование измерения м. не быть
    SendKeys("{RIGHT}")  # переключаемся на закладку Формирование измерения
    time.sleep(delay_win_active)
    app.TParamsForm.capture_as_image().save("Params_UserPointNames{}.bmp".format(expert_system))
    app.TParamsForm.close_alt_f4()

    # Params_RB_ExtraVibroLevels.bmp
    # TODO: закладки Подшипники м. не быть
    ids.menu_select("#1->#{}->#0".format(params_index))  # Сервис->Параметры->Общие
    SendKeys("{TAB}{TAB}{RIGHT}{RIGHT}")  # переключаемся на закладку Подшипники
    app.TParamsForm["...TBitBtn"].click()
    time.sleep(delay_win_active)
    # TODO: дефолтные параметры
    app.TParams_RB_PowerInBandsForm.capture_as_image().save("Params_RB_ExtraVibroLevels.bmp")
    app.TParams_RB_PowerInBandsForm.close_alt_f4()
    app.TParamsForm.close_alt_f4()

    # Params_SignalsShow.bmp
    # Params_SignalsShow_p2.bmp
    # TODO: дефолтные параметры
    ids.menu_select("#1->#{}->#1".format(params_index))  # Сервис->Параметры->Просмотр сигналов
    app.TGrafSetups_Params_Form.capture_as_image().save("Params_SignalsShow.bmp")
    SendKeys("{TAB}{TAB}{TAB}{TAB}{RIGHT}")  # переключаемся на закладку Параметры 2
    app.TGrafSetups_Params_Form.capture_as_image().save("Params_SignalsShow_p2.bmp")
    app.TGrafSetups_Params_Form.close_alt_f4()

    # Params_Reporter_p1.bmp
    # Params_Reporter_p2.bmp
    # Params_Reporter_p3.bmp
    # TODO: дефолтные параметры
    ids.menu_select("#1->#{}->#2".format(params_index))  # Сервис->Параметры->Формирование отчёта
    app.TSetupReporterForm.capture_as_image().save("Params_Reporter_p1.bmp")
    SendKeys("{TAB}{TAB}{RIGHT}")  # переключаемся на закладку Колонтитулы
    app.TSetupReporterForm.capture_as_image().save("Params_Reporter_p2.bmp")
    SendKeys("{RIGHT}")  # переключаемся на закладку Шрифт, поля
    app.TSetupReporterForm.capture_as_image().save("Params_Reporter_p3.bmp")
    app.TSetupReporterForm.close_alt_f4()

    # окна с параметрами экспертных систем
    # TODO: дефолтные параметры
    params_sub_count = ids.menu_item("#1->#{}".format(params_index)).sub_menu().item_count()
    for i in range(3, params_sub_count):
        ids.menu_select("#1->#{}->#{}".format(params_index, i))  # Сервис->Параметры->i
        if app.top_window().class_name() == "TSysParamsEditingForm":
            image_name = "Params_RB.bmp"
        else:
            print("unknown params window {}".format(app.top_window().class_name()))
            continue
        app.top_window().capture_as_image().save(image_name)
        if app.top_window().class_name() == "TSysParamsEditingForm":
            # не находит по TSysParamsEditingForm, используем top_window()
            app.top_window().Button5.click()  # Настройка каналов
            time.sleep(delay_win_active)
            app.TParams_RollingBearing_Channels_Form.capture_as_image().save("Params_RB_ChannsSetup_SVK.bmp")
            app.TParams_RollingBearing_Channels_Form.close_alt_f4()
        app.top_window().close_alt_f4()

    if app_created:
        ids.close()


def shoot_rr_viewer():
    # conf
    rr_viewer_path = "C:\Serv\ROSRRViewer.exe"

    app_created = False

    try:
        app = Application(backend="win32").connect(path=rr_viewer_path)
        app.TApplication.set_focus()
    except ProcessNotFoundError:
        app = Application(backend="win32").start(rr_viewer_path)
        app.wait_cpu_usage_lower()
        app_created = True

    rr_viewer = app.window(class_name="TMainForm")

    rr_viewer.move_window(200, 200, 600, 400)
    app.wait_cpu_usage_lower()

    # ROSRaveReportsViewer.bmp
    rr_viewer.capture_as_image().save("ROSRaveReportsViewer.bmp")

    if app_created:
        rr_viewer.close()


def main():
    _patch()

    start_mouse_pos = win32api.GetCursorPos()

    shoot_ids()
    shoot_rr_viewer()

    win32api.SetCursorPos(start_mouse_pos)


if __name__ == "__main__":
    main()
