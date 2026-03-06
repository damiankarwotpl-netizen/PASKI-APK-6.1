import os
import threading
from pathlib import Path
from datetime import datetime

from kivy.app import App
from kivy.clock import Clock
from kivy.core.window import Window
from kivy.metrics import dp
from kivy.utils import platform

from kivy.uix.boxlayout import BoxLayout
from kivy.uix.button import Button
from kivy.uix.label import Label
from kivy.uix.scrollview import ScrollView
from kivy.uix.gridlayout import GridLayout
from kivy.uix.popup import Popup
from kivy.uix.screenmanager import ScreenManager, Screen
from kivy.uix.checkbox import CheckBox

from openpyxl import load_workbook, Workbook
from openpyxl.utils import get_column_letter


APP_TITLE = "Paski Future"


class HomeScreen(Screen):
    pass


class TableScreen(Screen):
    pass


class PremiumButton(Button):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.background_normal = ""
        self.background_color = (0.2, 0.4, 0.8, 1)
        self.color = (1, 1, 1, 1)
        self.size_hint_y = None
        self.height = dp(48)


class PaskiFutureApp(App):

    def build(self):

        self.title = APP_TITLE
        Window.clearcolor = (0.08, 0.1, 0.15, 1)

        self.full_data = []
        self.filtered_data = []
        self.current_file = None

        self.selected_columns = None
        self.email_column = None

        self.sm = ScreenManager()

        self.home = HomeScreen(name="home")
        self.table = TableScreen(name="table")

        self._build_home()
        self._build_table()

        self.sm.add_widget(self.home)
        self.sm.add_widget(self.table)

        return self.sm

# ---------------- HOME ----------------

    def _build_home(self):

        layout = BoxLayout(orientation="vertical", padding=30, spacing=20)

        title = Label(text=APP_TITLE, font_size=28)

        open_btn = PremiumButton(text="Otwórz Excel")
        open_btn.bind(on_press=self.open_excel_picker)

        load_btn = PremiumButton(text="Wczytaj dane")
        load_btn.bind(on_press=self.load_full_excel)

        self.home_status = Label(text="Gotowy")

        layout.add_widget(title)
        layout.add_widget(open_btn)
        layout.add_widget(load_btn)
        layout.add_widget(self.home_status)

        self.home.add_widget(layout)

# ---------------- FILE PICKER ----------------

    def open_excel_picker(self, _):

        if platform != "android":
            self.home_status.text = "Picker działa tylko na Android"
            return

        from jnius import autoclass
        from android import activity

        PythonActivity = autoclass("org.kivy.android.PythonActivity")
        Intent = autoclass("android.content.Intent")

        intent = Intent(Intent.ACTION_OPEN_DOCUMENT)
        intent.setType("*/*")
        intent.addCategory(Intent.CATEGORY_OPENABLE)

        activity.bind(on_activity_result=self._on_activity_result)
        PythonActivity.mActivity.startActivityForResult(intent, 999)

    def _on_activity_result(self, request_code, result_code, intent):

        if request_code != 999 or not intent:
            return

        from android import activity
        activity.unbind(on_activity_result=self._on_activity_result)

        from jnius import autoclass
        PythonActivity = autoclass("org.kivy.android.PythonActivity")

        resolver = PythonActivity.mActivity.getContentResolver()
        uri = intent.getData()
        stream = resolver.openInputStream(uri)

        local = Path(self.user_data_dir) / "selected.xlsx"

        with open(local, "wb") as out:

            buf = bytearray(4096)

            while True:
                r = stream.read(buf)
                if r == -1:
                    break
                out.write(buf[:r])

        stream.close()

        self.current_file = local
        self.home_status.text = "Wybrano plik"

# ---------------- LOAD EXCEL ----------------

    def load_full_excel(self, _):

        if not self.current_file:
            self.popup("Błąd", "Najpierw wybierz plik")
            return

        wb = load_workbook(str(self.current_file), data_only=True)
        sheet = wb.active

        self.full_data = [[str(v) if v else "" for v in r] for r in sheet.iter_rows(values_only=True)]
        wb.close()

        self.filtered_data = self.full_data

        self.display_table()

        self.sm.current = "table"

# ---------------- TABLE ----------------

    def _build_table(self):

        root = BoxLayout(orientation="vertical")

        top = BoxLayout(size_hint_y=None, height=50, spacing=5)

        col_btn = PremiumButton(text="Kolumny")
        col_btn.bind(on_press=lambda x: self.select_columns_popup())

        email_btn = PremiumButton(text="Kolumna Email")
        email_btn.bind(on_press=lambda x: self.select_email_popup())

        export_btn = PremiumButton(text="Export")
        export_btn.bind(on_press=lambda x: threading.Thread(target=self.export_excel).start())

        back_btn = PremiumButton(text="Powrót")
        back_btn.bind(on_press=lambda x: setattr(self.sm, "current", "home"))

        top.add_widget(col_btn)
        top.add_widget(email_btn)
        top.add_widget(export_btn)
        top.add_widget(back_btn)

        self.table_scroll = ScrollView()

        root.add_widget(top)
        root.add_widget(self.table_scroll)

        self.table.add_widget(root)

# ---------------- DISPLAY TABLE ----------------

    def display_table(self):

        if not self.filtered_data:
            return

        grid = GridLayout(cols=len(self.filtered_data[0]), size_hint_y=None)
        grid.bind(minimum_height=grid.setter("height"))

        for row in self.filtered_data:
            for cell in row:
                lbl = Label(text=str(cell), size_hint_y=None, height=40)
                grid.add_widget(lbl)

        self.table_scroll.clear_widgets()
        self.table_scroll.add_widget(grid)

# ---------------- SELECT EXPORT COLUMNS ----------------

    def select_columns_popup(self):

        header = self.full_data[0]

        root = BoxLayout(orientation="vertical", padding=10, spacing=10)

        scroll = ScrollView()

        grid = GridLayout(cols=1, size_hint_y=None, spacing=10)
        grid.bind(minimum_height=grid.setter("height"))

        self.column_checks = []

        for i, name in enumerate(header):

            row = BoxLayout(size_hint_y=None, height=50)

            lbl = Label(text=name)
            cb = CheckBox()

            row.add_widget(lbl)
            row.add_widget(cb)

            grid.add_widget(row)

            self.column_checks.append((i, cb))

        scroll.add_widget(grid)

        btn = PremiumButton(text="OK")

        root.add_widget(scroll)
        root.add_widget(btn)

        popup = Popup(title="Kolumny exportu", content=root, size_hint=(0.9, 0.9))

        def apply_cols(instance):

            cols = []

            for i, cb in self.column_checks:
                if cb.active:
                    cols.append(i)

            if cols:
                self.selected_columns = cols

            popup.dismiss()

        btn.bind(on_press=apply_cols)

        popup.open()

# ---------------- SELECT EMAIL COLUMN ----------------

    def select_email_popup(self):

        header = self.full_data[0]

        root = BoxLayout(orientation="vertical", padding=10, spacing=10)

        scroll = ScrollView()

        grid = GridLayout(cols=1, size_hint_y=None, spacing=10)
        grid.bind(minimum_height=grid.setter("height"))

        self.email_checks = []

        for i, name in enumerate(header):

            row = BoxLayout(size_hint_y=None, height=50)

            lbl = Label(text=name)
            cb = CheckBox(group="email")

            row.add_widget(lbl)
            row.add_widget(cb)

            grid.add_widget(row)

            self.email_checks.append((i, cb))

        scroll.add_widget(grid)

        btn = PremiumButton(text="OK")

        root.add_widget(scroll)
        root.add_widget(btn)

        popup = Popup(title="Kolumna Email", content=root, size_hint=(0.9, 0.9))

        def apply_email(instance):

            for i, cb in self.email_checks:
                if cb.active:
                    self.email_column = i

            popup.dismiss()

        btn.bind(on_press=apply_email)

        popup.open()

# ---------------- EXPORT ----------------

    def export_excel(self):

        if not self.filtered_data:
            return

        wb = Workbook()
        ws = wb.active

        data = self.filtered_data

        if self.selected_columns:
            data = [[row[i] for i in self.selected_columns] for row in data]

        for row in data:
            ws.append(row)

        self.auto_size(ws)

        path = Path(self.user_data_dir) / f"export_{datetime.now().strftime('%H%M%S')}.xlsx"

        wb.save(path)

        Clock.schedule_once(lambda x: self.popup("Export", str(path)))

# ---------------- AUTO WIDTH ----------------

    def auto_size(self, ws):

        for column_cells in ws.columns:

            length = max(len(str(cell.value)) if cell.value else 0 for cell in column_cells)

            ws.column_dimensions[get_column_letter(column_cells[0].column)].width = length + 3

# ---------------- POPUP ----------------

    def popup(self, title, msg):

        box = BoxLayout(orientation="vertical", padding=10, spacing=10)

        box.add_widget(Label(text=msg))

        btn = PremiumButton(text="OK")

        box.add_widget(btn)

        pop = Popup(title=title, content=box, size_hint=(0.8, 0.4))

        btn.bind(on_press=pop.dismiss)

        pop.open()


if __name__ == "__main__":
    PaskiFutureApp().run()
