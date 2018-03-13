from pathlib import Path
import shutil
import time
import win32api

from pywinauto.application import Application
from pywinauto.keyboard import SendKeys
from pywinauto.win32structures import RECT

from .._pywinauto.helpers import treeview_element_colapse


class IDS:
    def __init__(self, path, obj_tree_path, msr_tree_path):
        self.path = path
        self.obj_tree_path = obj_tree_path
        self.msr_tree_path = msr_tree_path
        self.screen_rect = RECT(0, 0, win32api.GetSystemMetrics(0), win32api.GetSystemMetrics(1))
        self.delay_win_active = .5
        self.app = None
        self.ids = None
        self.params_index = None
        self.do_common_MainWindow = True
        self.do_common_MainWindow_MainMenu = True

    @classmethod
    def get_corepopup_rect(cls, popup) -> RECT:
        # popup: WindowSpecification
        rect = popup.rectangle()
        if popup.menu().item_count() > 17:
            r = popup.menu().item(17).rectangle()  # Диагностика
            rect.bottom = r.bottom + 5
        return rect

    @classmethod
    def get_suffixed_name(cls, name: str, suffix: str = None) -> str:
        return name.format(suffix if suffix is not None else "")

    def shoot(self):
        pass

    def shoot_common(self):
        if self.do_common_MainWindow:
            self.MainWindow()
        if self.do_common_MainWindow_MainMenu:
            self.MainWindow_MainMenu("_Common")
        self.MainWindow_NewObject_Call()
        self.MainWindow_NewObject()
        self.MainWindow_PSP_Call()
        self.PSP()
        self.Zamer_Show_Call()
        self.Zamer_Cascade_Call()
        self.Zamer_ShowMarka_Call()
        self.Zamer_Diagnos_Call()
        self.MainWindow_Export_Call()
        self.Export_Params()
        self.MainWindow_Import_Call()
        self.Import_Params()
        self.Import_Params_Extra()
        self.Params_SignalsShow()
        self.Params_Reporter()
        self.Params_Systems()

    def start_app(self):
        self.app = Application(backend="win32").start(self.path)
        self.app.wait_cpu_usage_lower()

        self.ids = self.app.window(class_name="TIDSMainForm")

        self.ids.move_window(200, 200, 600, 400)
        self.app.wait_cpu_usage_lower()

        # определяем индекс пункта меню Сервис->Параметры
        self.params_index = self.ids.menu_item("#1").sub_menu().item_count() - 1

    def stop_app(self):
        if self.ids is not None and self.ids.exists():
            self.ids.close()
            time.sleep(1)

    def MainWindow(self, suffix: str=None):
        filename = self.get_suffixed_name("MainWindow{}.bmp", suffix)
        treeview_element_colapse(self.ids.TreeView.tree_root())
        self.ids.capture_as_image().save(filename)

    def MainWindow_MainMenu(self, suffix: str=None):
        filename = self.get_suffixed_name("MainWindow_MainMenu{}.bmp", suffix)
        treeview_element_colapse(self.ids.TreeView.tree_root())
        self.ids.menu().highlight_and_capture("#1->#{}".format(self.params_index), filename)  # Сервис->Параметры->xxx
        SendKeys("{ESC}{ESC}{ESC}")  # закрываем меню

    def MainWindow_NewObject_Call(self):
        self.capture_popupmenu_item([0, 0], "#0", "MainWindow_NewObject_Call.bmp")  # Добавить объект
        SendKeys("{ESC}")  # закрываем меню

    def MainWindow_NewObject(self):
        self.select_popupmenu_item([0, ], "#0")  # Добавить объект
        time.sleep(self.delay_win_active)
        self.app.TStructureAddForm.capture_as_image().save("MainWindow_NewObject.bmp")
        self.app.TStructureAddForm.close_alt_f4()

    def MainWindow_NewZamer_Call(self):
        self.capture_popupmenu_item(self.obj_tree_path, "#2", "MainWindow_NewZamer_Call.bmp")  # Добавить измерение
        SendKeys("{ESC}")  # закрываем меню

    def MainWindow_PSP_Call(self):
        self.capture_popupmenu_item(self.obj_tree_path, "#12", "MainWindow_PSP_Call.bmp")  # Паспорт
        SendKeys("{ESC}")  # закрываем меню

    def PSP(self):
        # паспорт
        self.select_popupmenu_item(self.obj_tree_path, "#12")  # Паспорт
        time.sleep(self.delay_win_active)
        # изменение марки
        self.psp_change_marka()
        # подтверждение сохранения паспорта
        self.app.TMessageForm.capture_as_image().save("PSP_SaveConfirm.bmp")
        self.app.TMessageForm.Button3.click()  # Да
        time.sleep(self.delay_win_active)
        # подтверждение сохранения марки
        self.app.TMessageForm.capture_as_image().save("PSP_MarkaSaveConfirm.bmp")
        # TODO: в окне выводится название марки
        self.app.TMessageForm.Button2.click()  # Да
        time.sleep(self.delay_win_active)
        # инфо о существовании похожей марки (opt)
        if self.app.TMessageForm.exists() and self.app.TMessageForm.control_count() == 1:
            SendKeys("{ESC}")  # закрываем сообщение
            time.sleep(self.delay_win_active)
        # подтверждение изменения марок в замерах (opt)
        if self.app.TMessageForm.exists():
            self.app.TMessageForm.capture_as_image().save("PSP_MarkaSaveInZamersConfirm.bmp")
            SendKeys("{ESC}")  # закрываем сообщение
        else:
            print("skip PSP_MarkaSaveInZamersConfirm.bmp")

    def psp_change_marka(self):
        pass

    def Zamer_Show_Call(self):
        self.capture_popupmenu_item(self.msr_tree_path, "#14", "Zamer_Show_Call.bmp")  # Просмотр
        SendKeys("{ESC}")  # закрываем меню

    def Zamer_Cascade_Call(self):
        self.capture_popupmenu_item(self.obj_tree_path, "#15", "Zamer_Cascade_Call.bmp")  # Просмотр каскада
        SendKeys("{ESC}")  # закрываем меню

    def Zamer_ShowMarka_Call(self):
        self.capture_popupmenu_item(self.msr_tree_path, "#16", "Zamer_ShowMarka_Call.bmp")  # Измерение. Просмотр марки
        SendKeys("{ESC}")  # закрываем меню

    def Zamer_Diagnos_Call(self):
        self.capture_popupmenu_item(self.msr_tree_path, "#17", "Zamer_Diagnos_Call.bmp")  # Диагностика
        SendKeys("{ESC}")  # закрываем меню

    def MainWindow_Export_Call(self):
        treeview_element_colapse(self.ids.TreeView.tree_root())
        self.ids.menu().highlight_and_capture("#1->#4", "MainWindow_Export_Call.bmp")  # Сервис->Экспорт
        SendKeys("{ESC}{ESC}")  # закрываем меню

    def Export_Params(self):
        treeview_element_colapse(self.ids.TreeView.tree_root())
        self.ids.menu_select("#1->#4")  # Сервис->Экспорт
        time.sleep(self.delay_win_active)
        self.app.TExportInitForm.capture_as_image().save("Export_Params.bmp")
        self.app.TExportInitForm.close_alt_f4()

    def MainWindow_Import_Call(self):
        treeview_element_colapse(self.ids.TreeView.tree_root())
        self.ids.menu().highlight_and_capture("#1->#5", "MainWindow_Import_Call.bmp")  # Сервис->Импорт
        SendKeys("{ESC}{ESC}")  # закрываем меню

    def Import_Params(self):
        treeview_element_colapse(self.ids.TreeView.tree_root())
        self.ids.menu_select("#1->#5")  # Сервис->Импорт
        time.sleep(self.delay_win_active)
        self.app.TImportInitForm.capture_as_image().save("Import_Params.bmp")
        self.app.TImportInitForm.close_alt_f4()

    def Import_Params_Extra(self):
        treeview_element_colapse(self.ids.TreeView.tree_root())
        self.ids.menu_select("#1->#5")  # Сервис->Импорт
        self.app.TImportInitForm.Button1.click()  # Настройки
        time.sleep(self.delay_win_active)
        self.app.TImportSettingsForm.capture_as_image().save("Import_Params_Extra.bmp")
        SendKeys("{ESC}{ESC}")  # закрываем меню

    def Params_SignalsShow(self):
        # TODO: дефолтные параметры
        self.ids.menu_select("#1->#{}->#1".format(self.params_index))  # Сервис->Параметры->Просмотр сигналов
        self.app.TGrafSetups_Params_Form.capture_as_image().save("Params_SignalsShow.bmp")
        SendKeys("{TAB}{TAB}{TAB}{TAB}{RIGHT}")  # переключаемся на закладку Параметры 2
        self.app.TGrafSetups_Params_Form.capture_as_image().save("Params_SignalsShow_p2.bmp")
        self.app.TGrafSetups_Params_Form.close_alt_f4()

    def Params_Reporter(self):
        # TODO: дефолтные параметры
        self.ids.menu_select("#1->#{}->#2".format(self.params_index))  # Сервис->Параметры->Формирование отчёта
        self.app.TSetupReporterForm.capture_as_image().save("Params_Reporter_p1.bmp")
        SendKeys("{TAB}{TAB}{RIGHT}")  # переключаемся на закладку Колонтитулы
        self.app.TSetupReporterForm.capture_as_image().save("Params_Reporter_p2.bmp")
        SendKeys("{RIGHT}")  # переключаемся на закладку Шрифт, поля
        self.app.TSetupReporterForm.capture_as_image().save("Params_Reporter_p3.bmp")
        self.app.TSetupReporterForm.close_alt_f4()

    def Params_Systems(self):
        """ окна с параметрами экспертных систем """

        # TODO: дефолтные параметры
        params_sub_count = self.ids.menu_item("#1->#{}".format(self.params_index)).sub_menu().item_count()
        for i in range(3, params_sub_count):
            self.ids.menu_select("#1->#{}->#{}".format(self.params_index, i))  # Сервис->Параметры->i
            if self.app.top_window().class_name() == "TSysParamsEditingForm":
                image_name = "Params_RB.bmp"
            else:
                print("unknown params window {}".format(self.app.top_window().class_name()))
                continue
            self.app.top_window().capture_as_image().save(image_name)
            if self.app.top_window().class_name() == "TSysParamsEditingForm":
                # не находит по TSysParamsEditingForm, используем top_window()
                self.app.top_window().Button5.click()  # Настройка каналов
                time.sleep(self.delay_win_active)
                self.app.TParams_RollingBearing_Channels_Form.capture_as_image().save("Params_RB_ChannsSetup_SVK.bmp")
                self.app.TParams_RollingBearing_Channels_Form.close_alt_f4()
            self.app.top_window().close_alt_f4()

    def capture_popupmenu_item(self, tree_path, menu_path, filename):
        self.ids.TreeView.get_item(tree_path).click_input(button="right")
        r = self.get_corepopup_rect(self.app.PopupMenu)
        self.app.PopupMenu.wrapper_object().highlight_and_capture(menu_path, filename, r)

    def select_popupmenu_item(self, tree_path, menu_path):
        self.ids.TreeView.get_item(tree_path).click_input(button="right")
        self.app.PopupMenu.menu_item(menu_path).click_input()


class ExpertSystem(IDS):
    KEY_NAME = "Opts.gkd"

    def __init__(self, path, obj_tree_path, msr_tree_path, conf):
        super().__init__(path, obj_tree_path, msr_tree_path)
        self.key = conf.get("key", None)
        self.key_origin = str(Path(self.path).with_name(self.KEY_NAME))
        self.key_bak = self.key_origin + ".bak"
        self.suffix = conf["suffix"]
        self.import_file = conf["import_file"]

    def start_app(self):
        if self.key is not None and self.key != "":
            self.key_backup()
            shutil.copy2(self.key, self.key_origin)
        super().start_app()

    def stop_app(self):
        super().stop_app()
        if self.key is not None and self.key != "":
            self.key_restore()

    def key_backup(self):
        shutil.copy2(self.key_origin, self.key_bak)

    def key_restore(self):
        shutil.copy2(self.key_bak, self.key_origin)


class KD(ExpertSystem):
    def shoot(self):
        pass


class RB(ExpertSystem):
    def __init__(self, path, obj_tree_path, msr_tree_path, conf):
        super().__init__(path, obj_tree_path, msr_tree_path, conf)
        self.rb_tree_path = conf["tree_path"]["obj"]
        self.rbmsr_tree_path = conf["tree_path"]["msr"]
        self.do_common_MainWindow = False
        self.do_common_MainWindow_MainMenu = False

    def shoot(self):
        self.MainWindow(self.suffix)
        self.MainWindow_MainMenu(self.suffix)
        self.ZamerReg_SVK()
        self.MainWindow_ZamerFormirov_SVK()
        self.PSP_RB()
        self.PSP_Marka_RB()
        self.Zamer_Show_RollingBearing()
        self.Zamer_Cascade()
        self.Diagnoz_RB()
        self.SF()
        self.Zamer_Extra_SVK_CommonDiag()
        self.Import_Manual()
        self.DBEditor_Markas()
        self.Params()
        self.Params_RB_ExtraVibroLevels()

    def psp_change_marka(self):
        # не находит по TPSP_RollingBearingForm, используем top_window()
        self.app.top_window().CheckBox1.click()  # инвертируем состояние Подшипник из нержавеющей...
        self.app.top_window().TBitBtn3.click()  # Применить
        time.sleep(self.delay_win_active)

    def ZamerReg_SVK(self):
        self.select_popupmenu_item(self.rb_tree_path, "#2")  # Добавить измерение
        time.sleep(7)  # идёт инициализация АЦП
        if self.app.TMessageForm.exists():
            # нет АЦП, условий в МВИ и т.п.
            print("skip ZamerReg_SVK.bmp")
            SendKeys("{ESC}")  # закрываем сообщение
        else:
            # ищем без best_match, на случай если юзер сам закрыл окно предупреждения
            self.app.window(class_name="TSVKSignalForm").capture_as_image().save("ZamerReg_SVK.bmp")
            self.app.TSVKSignalForm.close_alt_f4()
            self.app.wait_cpu_usage_lower()

    def MainWindow_ZamerFormirov_SVK(self):
        # включение ручного формирования
        self.change_add_msr_mode("ручное")  # TODO: l10n
        #
        self.select_popupmenu_item(self.rb_tree_path, "#2")  # Добавить измерение
        time.sleep(self.delay_win_active)
        self.app.TZamerFormirovForm.capture_as_image().save("MainWindow_ZamerFormirov_SVK.bmp")
        self.app.TZamerFormirovForm.close_alt_f4()
        # включение автоматического формирования
        self.change_add_msr_mode("автоматическое")  # TODO: l10n

    def change_add_msr_mode(self, mode: str):
        self.ids.menu_select("#1->#{}->#0".format(self.params_index))  # Сервис->Параметры->Общие
        self.app.TParamsForm[mode].check_by_click()
        self.app.TParamsForm.OkButton.click()

    def PSP_RB(self):
        self.select_popupmenu_item(self.rb_tree_path, "#12")  # Паспорт
        time.sleep(self.delay_win_active)
        # не находит по TPSP_RollingBearingForm, используем top_window()
        self.app.top_window().capture_as_image().save("PSP_RB.bmp")
        self.app.top_window().close_alt_f4()

    def PSP_Marka_RB(self):
        self.ids.menu_select("#1->#0")  # Сервис->Редактор БД
        # не находит по TDBEditForm, используем top_window()
        self.app.top_window().TreeView.select((0, "Подшипник качения"))  # TODO: l10n
        self.app.wait_cpu_usage_lower()
        self.app.top_window().TListView.double_click(coords=(5, 5))  # открываем окно первой марки в списке
        time.sleep(2 * self.delay_win_active)
        self.app.top_window().capture_as_image().save("PSP_Marka_RB.bmp")
        self.app.top_window().close_alt_f4()  # окно марки
        time.sleep(self.delay_win_active)
        self.app.top_window().close_alt_f4()  # окно Редактор БД
        self.app.wait_cpu_usage_lower()

    def Zamer_Show_RollingBearing(self):
        self.ids.TreeView.get_item(self.rbmsr_tree_path).click_input(double=True)
        time.sleep(self.delay_win_active)
        # подгоняем размер чтобы не было верт. скрола и окно целиком влазило на экран
        self.app.TSignalShow_RollingBearingForm.move_window(100, 0, 800, 970)
        # TODO: дефолтные параметры
        self.app.TSignalShow_RollingBearingForm.capture_as_image().save("Zamer_Show_RollingBearing.bmp")
        self.app.TSignalShow_RollingBearingForm.close_alt_f4()

    def Zamer_Cascade(self):
        self.select_popupmenu_item(self.rb_tree_path, "#15")  # Просмотр каскада
        time.sleep(self.delay_win_active)
        self.app.TShowCascadeInitForm.capture_as_image().save("Zamer_Cascade_SelectFiles{}.bmp".format(self.suffix))
        # TODO: выбрать замеры
        self.app.TShowCascadeInitForm.Button2.click()  # Просмотр
        if self.app.TMessageForm.exists():
            print("skip Zamer_Cascade{}.bmp".format(self.suffix))
            SendKeys("{ESC}")  # закрываем сообщение
            self.app.TShowCascadeInitForm.close_alt_f4()
        else:
            self.app.TShowCascadeForm.capture_as_image().save("Zamer_Cascade{}.bmp".format(self.suffix))
            self.app.TShowCascadeForm.close_alt_f4()

    def Diagnoz_RB(self):
        self.select_popupmenu_item(self.rbmsr_tree_path, "#17")  # Диагностика
        time.sleep(self.delay_win_active)
        if self.app.TMessageForm.exists():
            if self.app.TMessageForm.control_count() == 2:
                # предупреждение о нормах и т.п.
                self.app.TMessageForm.Button2.click()  # Да
                time.sleep(self.delay_win_active)
            else:
                # ошибка диагностики
                SendKeys("{ESC}")  # закрываем сообщение
        if self.app.TRollingBearing_DiagnosForm.exists():
            self.app.TRollingBearing_DiagnosForm.capture_as_image().save("Diagnoz_RB.bmp")
            # TODO: картинки примеров отчётов (из браузера)
            # self.app.TRollingBearing_DiagnosForm.TBitBtn1.click()  # Отчёт
            self.app.TRollingBearing_DiagnosForm.close_alt_f4()
        else:
            print("skip Diagnoz_RB.bmp")

    def SF(self):
        # TODO: спец. функции. Желательно, чтобы меню "уходило" вверх - м.б. сместить окно в низ экрана
        self.ids.TreeView.get_item(self.rbmsr_tree_path).click_input(button="right")
        # Подшипники качения->Фактический зазор
        if self.app.PopupMenu.menu().item_count() > 21:
            self.app.PopupMenu.wrapper_object().highlight_and_capture("#21->#2", "EDIT_SF_RB_Call.bmp", self.screen_rect)
            SendKeys("{ESC}{ESC}")  # закрываем меню
        else:
            print("skip SF_RB_Call.bmp")
            SendKeys("{ESC}")  # закрываем меню

    def Zamer_Extra_SVK_CommonDiag(self):
        self.select_popupmenu_item(self.rb_tree_path, "#19->#0")
        time.sleep(self.delay_win_active)
        self.app.TRollingBearing_CommonDiagIntro_Form.capture_as_image().save("Zamer_Extra_SVK_CommonDiag.bmp")
        # TODO: картинка примера отчёта (из браузера)
        # self.app.TRollingBearing_CommonDiagIntro_Form.Button2.click()  # Ok
        self.app.TRollingBearing_CommonDiagIntro_Form.close_alt_f4()

    def Import_Manual(self):
        treeview_element_colapse(self.ids.TreeView.tree_root())
        self.ids.menu_select("#1->#5")  # Сервис->Импорт
        self.app.TImportInitForm["...Button"].click()  # выбор файла
        self.app.DialogEdit.Edit.set_edit_text(self.import_file)
        self.app.DialogEdit.Button1.click_input()  # Открыть
        if self.app.DialogEdit.exists() or self.app.Dialog.exists():
            # ошибка открытия файла (иногда окно предупреждения не появляется)
            print("skip Import_Manual{}.bmp".format(self.suffix))
            SendKeys("{ESC}{ESC}{ESC}")  # закрываем все окна вплоть до IDS
        else:
            self.app.TImportInitForm.Button3.click()  # Ok
            self.app.wait_cpu_usage_lower()
            self.app.TImportTreeForm.move_window(450, 270, 330, 300)  # в границах IDS, справа
            self.ids.capture_as_image().save("Import_Manual{}.bmp".format(self.suffix))
            self.app.TImportTreeForm.close_alt_f4()

    def DBEditor_Markas(self):
        self.ids.menu_select("#1->#0")  # Сервис->Редактор БД
        # не находит по TDBEditForm, используем top_window()
        self.app.top_window().move_window(100, 100, 400, 300)
        self.app.top_window().capture_as_image().save("DBEditor_Markas{}.bmp".format(self.suffix))
        self.app.top_window().close_alt_f4()

    def Params(self):
        # Params_Common{}.bmp
        # Params_Common2{}.bmp
        # Params_RB{}.bmp
        # Params_UserPointNames{}.bmp
        # TODO: дефолтные параметры
        self.ids.menu_select("#1->#{}->#0".format(self.params_index))  # Сервис->Параметры->Общие
        self.app.TParamsForm.capture_as_image().save("Params_Common{}.bmp".format(self.suffix))
        SendKeys("{TAB}{TAB}{RIGHT}")  # переключаемся на закладку Общие2
        self.app.TParamsForm.capture_as_image().save("Params_Common2{}.bmp".format(self.suffix))
        # TODO: третьей закладки м. не быть, также это м.б. Формирование измерения
        SendKeys("{RIGHT}")  # переключаемся на закладку Подшипники
        self.app.TParamsForm.capture_as_image().save("Params_RB{}.bmp".format(self.suffix))
        # TODO: закладки Формирование измерения м. не быть
        SendKeys("{RIGHT}")  # переключаемся на закладку Формирование измерения
        time.sleep(self.delay_win_active)
        self.app.TParamsForm.capture_as_image().save("Params_UserPointNames{}.bmp".format(self.suffix))
        self.app.TParamsForm.close_alt_f4()

    def Params_RB_ExtraVibroLevels(self):
        self.ids.menu_select("#1->#{}->#0".format(self.params_index))  # Сервис->Параметры->Общие
        SendKeys("{TAB}{TAB}{RIGHT}{RIGHT}")  # переключаемся на закладку Подшипники
        self.app.TParamsForm["...TBitBtn"].click()
        time.sleep(self.delay_win_active)
        # TODO: дефолтные параметры
        self.app.TParams_RB_PowerInBandsForm.capture_as_image().save("Params_RB_ExtraVibroLevels.bmp")
        self.app.TParams_RB_PowerInBandsForm.close_alt_f4()
        self.app.TParamsForm.close_alt_f4()


def shoot_ids(conf):
    path = conf["path"]
    obj_tree_path = conf["tree_path"]["obj"]
    msr_tree_path = conf["tree_path"]["msr"]

    for i, s in enumerate(conf["systems"]):
        if s["type"] == "kd":
            o = KD(path, obj_tree_path, msr_tree_path, s)
        elif s["type"] == "rb":
            o = RB(path, obj_tree_path, msr_tree_path, s)
        else:
            continue

        try:
            o.start_app()
            if i == 0:
                o.shoot_common()
            o.shoot()
        finally:
            o.stop_app()
