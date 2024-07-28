import glob
import os
import sys

from kivy.core.window import Window
from kivy.lang import Builder
from kivy.properties import StringProperty
from kivy.resources import resource_add_path
from kivy.uix.boxlayout import BoxLayout
from kivymd.app import MDApp
from kivymd.material_resources import dp
from kivymd.uix.button import MDFlatButton
from kivymd.uix.dialog import MDDialog
from kivymd.uix.list import OneLineListItem
from kivymd.uix.menu import MDDropdownMenu
from kivymd.uix.pickers import MDDatePicker
import locale
import File_work
from kivy_deps import sdl2, glew


locale.setlocale(locale.LC_ALL, 'ru_RU')


class Content(BoxLayout):
    pass


class MainBoxLayout(MDApp):

    def write_file(self):
        text_to_display = ""
        if len(self.get_list_employees_to_display()) == 0:
            text_to_display += "Выберите техников\n"
        if self.dropdown_text == "Выбор по дате":
            text_to_display += "Выберите критерий по дате\n"
        if len(self.date_range) == 0:
            text_to_display += "Выберите период"

        if text_to_display != "":
            dialog = MDDialog(title='Ошибка данных', text=text_to_display, auto_dismiss=True, size_hint=(None, None),
                              size=('7cm', '4cm'), radius=[20, 7, 20, 7], )
            dialog.open()
        else:

            list_of_employees = []
            for employee in sorted(self.get_list_employees_to_display(), key=lambda x: x.id):
                list_of_employees.append(employee.id)
                # list_of_employees.append(employee.id)

            text_name = ""
            file_name = ""
            shorter_left = str(self.date_range[0]).split('-')
            if len(self.date_range) > 1:
                shorter_right = str(self.date_range[len(self.date_range) - 1]).split('-')
                file_name = (F"Отчет {shorter_left[2]}.{shorter_left[1]}.{shorter_left[0]} "
                                          F"- {shorter_right[2]}."
                                          F"{shorter_right[1]}."
                                          F"{shorter_right[0]}")
                text_name = resource_path(F"..\Отчеты\\{file_name}")
                if self.check_box_value:
                    text_name += " Полный.xls"
                    file_name += " Полный.xls"
                else:
                    text_name += " Упрощенный.xls"
                    file_name += " Упрощенный.xls"
            else:
                text_name = resource_path(F"..\Отчет за {shorter_left[2]}.{shorter_left[1]}.{shorter_left[0]}.xls")

            File_work.write_file(self.name_of_file, list_of_employees, self.date_range, self.dropdown_text,
                                 self.check_box_value, text_name)

            dialog = MDDialog(title='Готово!', text=F"Отчёт был успешно создан в папке Отчёты.\nИмя отчёта:\n{file_name}", auto_dismiss=True,
                              size_hint=(None, None), size=('15cm', '4cm'), radius=[20, 7, 20, 7], )
            dialog.open()

    global_list_of_employees = []
    global_dict_of_employees = {}
    height = 550
    width = 850
    scrollPart = 0.75
    buttonPart = 1 - scrollPart
    dates_label = StringProperty("")
    dropdown_text = StringProperty("Выбор по дате")
    preset_name = StringProperty("")
    group_button_text = StringProperty("Выбрать группу")

    def build(self):
        self.title = 'Артдент'
        self.theme_cls.primary_palette = "Amber"
        return self.screen

    def open_file(self):
        name_of_file = self.name_of_file
        self.global_list_of_employees = File_work.first_scan(name_of_file)
        self.set_list_employees_to_show_left(self.global_list_of_employees.copy())
        for employee in self.global_list_of_employees:
            self.global_dict_of_employees.update({employee.id: employee.name})
        self.scroll_view_to_click_update()

    '''
    def on_start(self):
        for employee in self.list_of_employees_test:
            self.screen.ids.container_all.add_widget(
                OneLineListItem(width='10dp', text=f"{employee}", id=f"{employee}",
                                on_press=lambda x: self.scroll_view_on_pressed(x))
            )
            # self.list_of_employees_ids.append(self.screen.ids.employee.id)

        '''
    '''
        for employee in self.list_employees_to_click:
            self.screen.ids.container_to_add.add_widget(
                OneLineListItem(width='10dp', text=f"{employee}", id=f"{employee}",
                                on_press=lambda x: self.scroll_view_on_pressed(x))
            )
        '''

    def test(self):
        pass

    def add_employees_to_display(self):
        for employee_to_add in self.get_list_employees_to_click():
            temp_bool = True
            for employee in self.get_list_employees_to_display():
                if employee.id == employee_to_add.id:
                    temp_bool = False
            if temp_bool:
                self.add_list_employees_to_display(employee_to_add)

        self.clear_scroll_view_to_click()
        self.scroll_view_to_display_update()

    def clear_scroll_view_to_click(self):
        for employee in self.get_list_employees_to_click():
            employee.bg_color = [0.98, 0.98, 0.98, 0.0]
        self.set_list_employees_to_click([])

    def scroll_view_delete_everyone(self):
        self.set_list_employees_to_display([])
        self.set_list_employees_to_display_clicked([])
        self.scroll_view_to_display_update()

    def scroll_view_add_everyone(self, list_to_use):
        for employee in list_to_use:
            widget = self.employees_to_display_create_widget(employee)
            self.add_list_employees_to_click(widget)
        self.set_list_employees_to_display([])
        self.add_employees_to_display()
        self.clear_scroll_view_to_click()
        self.scroll_view_to_display_update()

    def delete_employees(self):
        for employee_to_del in self.get_list_employees_to_display_clicked():
            for employee in self.get_list_employees_to_display():
                if employee_to_del.id == employee.id:
                    self.remove_list_employees_to_display(employee)
                    if employee in self.right_list_save:
                        self.right_list_save.remove(employee)

        self.set_list_employees_to_display_clicked([])
        self.scroll_view_to_display_update()

    def get_list_employees_to_click(self):
        return self.list_employees_to_click

    def set_list_employees_to_click(self, list_to_update):
        self.list_employees_to_click = list_to_update.copy()

    def get_list_employees_to_display(self):
        return self.list_employees_to_display

    def set_list_employees_to_display(self, list_to_update):
        self.list_employees_to_display = list_to_update.copy()

    def add_list_employees_to_click(self, employee):
        self.list_employees_to_click.append(employee)

    def add_list_employees_to_display(self, employee):
        self.list_employees_to_display.append(employee)

    def remove_list_employees_to_display(self, employee):
        self.list_employees_to_display.remove(employee)

    def remove_list_employees_to_click(self, employee):
        self.list_employees_to_click.remove(employee)

    def add_list_employees_to_display_clicked(self, employee):
        self.list_employees_to_display_clicked.append(employee)

    def set_list_employees_to_display_clicked(self, list_to_update):
        self.list_employees_to_display_clicked = list_to_update.copy()

    def remove_list_employees_to_display_clicked(self, employee):
        self.list_employees_to_display_clicked.remove(employee)

    def get_list_employees_to_display_clicked(self):
        return self.list_employees_to_display_clicked

    def add_list_employees_to_show_left(self, employee):
        self.list_employees_to_show_left.append(employee)

    def set_list_employees_to_show_left(self, list_to_update):
        self.list_employees_to_show_left = list_to_update.copy()

    def remove_list_employees_to_show_left(self, employee):
        self.list_employees_to_show_left.remove(employee)

    def get_list_employees_to_show_left(self):
        return self.list_employees_to_show_left

    def scroll_view_to_display_update(self):
        self.screen.ids.container_to_add.clear_widgets()
        for employee in sorted(self.get_list_employees_to_display(), key=lambda x: x.id):
            self.screen.ids.container_to_add.add_widget(self.employees_to_display_create_widget(employee))

    def employees_to_display_create_widget(self, employee):
        return OneLineListItem(width='10dp', text=f"{employee.id}", id=employee.id,
                               on_press=lambda x: self.scroll_view_display(x))

    def employees_to_click_create_widget(self, employee):
        return OneLineListItem(width='10dp', text=f"{employee.id}", id=employee.id,
                               on_press=lambda x: self.scroll_view_on_pressed(x))

    def scroll_view_to_click_update(self):
        self.screen.ids.container_all.clear_widgets()
        for employee in sorted(self.get_list_employees_to_show_left(), key=lambda x: x.id):
            self.screen.ids.container_all.add_widget(self.employees_to_click_create_widget(employee))

    def scroll_view_display(self, one_line_employee):
        if one_line_employee not in self.get_list_employees_to_display_clicked():
            self.add_list_employees_to_display_clicked(one_line_employee)
            one_line_employee.bg_color = [1, 0.745, 0, 0.7]
        else:
            self.remove_list_employees_to_display_clicked(one_line_employee)
            one_line_employee.bg_color = [0.98, 0.98, 0.98, 0.0]

    def scroll_view_on_pressed(self, one_line_employee):
        if one_line_employee not in self.get_list_employees_to_click():
            self.add_list_employees_to_click(one_line_employee)
            one_line_employee.bg_color = [1, 0.745, 0, 0.7]
        else:
            self.remove_list_employees_to_click(one_line_employee)
            one_line_employee.bg_color = [0.98, 0.98, 0.98, 0.0]

    def set_technic_list_left(self, text="", search=False):
        if len(self.left_list_save) != 0:
            self.set_list_employees_to_show_left(self.left_list_save.copy())
        if search:
            temp_show_list = []
            if len(self.left_list_save) == 0:
                self.left_list_save = self.get_list_employees_to_show_left().copy()
            for employee in self.get_list_employees_to_show_left():
                if employee.id.lower().startswith(text.lower()):
                    temp_show_list.append(employee)
            self.set_list_employees_to_show_left(temp_show_list)
            self.scroll_view_to_click_update()
        else:
            self.set_list_employees_to_show_left(self.left_list_save.copy())
            self.left_list_save = []
            self.scroll_view_to_click_update()

    def set_technic_list_right(self, text="", search=False):
        if len(self.right_list_save) != 0:
            self.set_list_employees_to_display(self.right_list_save.copy())
        if search:
            temp_show_list = []
            if len(self.right_list_save) == 0:
                self.right_list_save = self.get_list_employees_to_display().copy()
            for employee in self.get_list_employees_to_display():
                if employee.id.lower().startswith(text.lower()):
                    temp_show_list.append(employee)
            self.set_list_employees_to_display(temp_show_list)
        else:
            self.set_list_employees_to_display(self.right_list_save.copy())
            self.right_list_save = []
        self.scroll_view_to_display_update()
        self.scroll_view_to_display_update()

    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.check_box_value = False
        self.preset_list = []
        self.right_list_save = []
        self.left_list_save = []
        self.date_range = []
        self.name_of_file = resource_path("..\otchet.xls")
        Window.size = (self.width, self.height)
        self.list_employees_to_show_left = []
        self.list_employees_to_click = []
        self.list_employees_to_display = []
        self.list_employees_to_display_clicked = []
        self.screen = Builder.load_file(resource_path("TheLab.kv"))
        menu_items = [
            {
                "viewclass": "OneLineListItem",
                "height": dp(45),
                "text": "Дата отправки",
                "on_release": lambda x="Дата отправки": self.menu_callback(x),

            }, {
                "viewclass": "OneLineListItem",
                "height": dp(45),
                "text": "Дата сдачи",
                "on_release": lambda x="Дата сдачи": self.menu_callback(x),

            }, {
                "viewclass": "OneLineListItem",
                "height": dp(45),
                "text": "Дата техника",
                "on_release": lambda x="Дата техника": self.menu_callback(x),

            }
        ]
        self.menu = MDDropdownMenu(
            caller=self.screen.ids.button,
            items=menu_items,
            width_mult=2.5,
        )

        self.group_create()

    def group_create(self):
        preset_names = []

        for filename in glob.glob(resource_path('..\Группы') + "\*.txt"):
            with open(os.path.join(os.getcwd(), filename), 'r') as f:
                preset_names.append({
                    "viewclass": "OneLineListItem",
                    "height": dp(45),
                    "text": filename.split('\\')[-1][0:-4],
                    "on_release": lambda x=filename: self.choose_group(x),

                })

        preset_names.append({
                "viewclass": "OneLineListItem",
                "height": dp(45),
                "text": "Новая группа",
                "on_release": lambda x="Новая группа": self.create_new_group(x),

        })

        self.preset_list = MDDropdownMenu(
            caller=self.screen.ids.button_choose_group,
            items=preset_names,
            width_mult=3.5,
        )

    def create_new_group(self, name):
        self.preset_name = name
        self.group_button_text = name
        self.set_list_employees_to_display([])
        self.scroll_view_to_display_update()

        self.preset_list.dismiss()


    def on_save(self, instance, value, date_range):
        # self.root.ids.date_label.text = str(value)
        if len(date_range) == 1:
            self.dates_label = "Период:\n" + f' {str(date_range[0])}'
            self.date_range = date_range.copy()
        elif len(date_range) > 1:
            self.dates_label = "Период:\n" + f' {str(date_range[0])} -\n {str(date_range[-1])}'
            self.date_range = date_range.copy()
        else:
            self.dates_label = "Даты не выбраны"
            self.date_range = []

    # Click Cancel
    def on_cancel(self, instance, value):
        self.dates_label = "Даты не выбраны"
        self.date_range = []

    def show_date_picker(self, *args):
        date_dialog = MDDatePicker(mode="range")
        date_dialog.bind(on_save=self.on_save, on_cancel=self.on_cancel)
        date_dialog.open()

    def menu_callback(self, text_item):
        self.dropdown_text = text_item
        self.menu.dismiss()

    def choose_group(self, name):
        self.preset_name = name.split("\\")[-1][0:-4]
        self.group_button_text = self.preset_name
        temp_mass = []
        with open(resource_path(F"{name}"), "r") as f:
            temp_mass_cache = f.readlines()
            for line in temp_mass_cache:
                temp_mass.append(line.split("\n")[0])

        res = []
        for employee in temp_mass:
            res.append(File_work.Employee(employee))

        self.scroll_view_add_everyone(res)
        self.preset_list.dismiss()

    def open_popup_check(self):
        text_to_display = "Введите название группы"
        self.dialog = MDDialog(
            title=text_to_display,
            type="custom",
            content_cls=Content(),
            radius=[20, 7, 20, 7],
            buttons=[
                MDFlatButton(
                    text="Назад",
                    theme_text_color="Custom",
                    text_color=self.theme_cls.primary_color,
                    on_release=self.dialog_dismiss
                ),
                MDFlatButton(
                    text="ОК",
                    theme_text_color="Custom",
                    text_color=self.theme_cls.primary_color,
                    on_release=self.dialog_save
                ),
            ],
        )
        self.dialog.open()

    def dialog_dismiss(self, button):
        self.preset_name = ""
        self.dialog.dismiss(force=True)

    def dialog_save(self, button):

        self.save_new_preset()
        self.group_button_text = self.preset_name

        self.group_create()
        self.dialog.dismiss(force=True)

    def save_new_preset(self):
        with open(resource_path(F"..\Группы\{self.preset_name}.txt"), "w") as f:
            for i in range(len(self.get_list_employees_to_display())):
                employee = self.get_list_employees_to_display()[i]
                if i == len(self.get_list_employees_to_display()) - 1:
                    f.write(employee.id)
                else:
                    f.write(employee.id + "\n")

    def save_preset_name(self, text_to_save):
        self.preset_name = text_to_save

    def open_popup_delete(self):
        if self.preset_name != "":
            text_to_display = F"Вы уверены, что хотите удалить группу {self.preset_name}?"
            self.dialog_to_delete = MDDialog(
                title=text_to_display,
                radius=[20, 7, 20, 7],
                size_hint_y=None,
                size_hint_x=None,
                height=F"{self.height / 4}dp",
                width=F"{self.width / 1.85}dp",
                buttons=[
                    MDFlatButton(
                        text="Да",
                        theme_text_color="Custom",
                        text_color=self.theme_cls.primary_color,
                        on_release=self.dialog_to_delete_yes
                    ),
                    MDFlatButton(
                        text="Нет",
                        theme_text_color="Custom",
                        text_color=self.theme_cls.primary_color,
                        on_release=self.dialog_to_delete_no
                    ),
                ],
            )
            self.dialog_to_delete.open()

        else:
            self.dialog_to_delete = MDDialog(
                title="Вы не выбрали группу",
                radius=[20, 7, 20, 7],
                size_hint_y=None,
                size_hint_x=None,
                height=F"{self.height / 3}dp",
                width=F"{self.width / 3}dp",
            )
            self.dialog_to_delete.open()

    def dialog_to_delete_yes(self, button):
        path = resource_path(F'..\Группы\{self.preset_name}.txt')
        self.preset_name = ""
        os.remove(path)
        self.group_button_text = "Выбрать группу"
        self.set_list_employees_to_display([])
        self.scroll_view_to_display_update()
        self.group_create()
        self.dialog_to_delete.dismiss(force=True)

    def dialog_to_delete_no(self, button):
        self.dialog_to_delete.dismiss(force=True)

    def on_checkbox_active_left(self, checkbox, value):
        self.check_box_value = False

    def on_checkbox_active_right(self, checkbox, value):
        self.check_box_value = True


def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath('.')
    return os.path.join(base_path, relative_path)


if __name__ == "__main__":
    if hasattr(sys, '_MEIPASS'):
        resource_add_path((os.path.join(sys._MEIPASS)))
    MainBoxLayout().run()
