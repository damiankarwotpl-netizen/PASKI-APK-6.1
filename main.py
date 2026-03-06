import os
import json
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
from kivy.uix.textinput import TextInput
from kivy.uix.popup import Popup
from kivy.uix.progressbar import ProgressBar
from kivy.uix.screenmanager import ScreenManager, Screen

APP_TITLE = "Paski Future 6.1 STABLE PREMIUM"
CONFIG_FILE = "smtp_config.json"
EXPORT_CONFIG = "export_folder.json"


class HomeScreen(Screen):
    pass


class TableScreen(Screen):
    pass


class EmailScreen(Screen):
    pass


class SMTPScreen(Screen):
    pass


class PremiumButton(Button):

    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.background_normal = ""
        self.background_color = (0.2, 0.4, 0.8, 1)
        self.color = (1, 1, 1, 1)
        self.font_size = 16
        self.size_hint_y = None
        self.height = dp(48)


class PaskiFutureApp(App):

    def build(self):

        self.title = APP_TITLE
        Window.clearcolor = (0.08, 0.1, 0.15, 1)

        self.full_data = []
        self.filtered_data = []
        self.current_file = None
        self.export_folder = None
        self.email_file = None
        self.email_dict = {}

        os.makedirs(self.user_data_dir, exist_ok=True)

        self.load_export_folder()

        self.sm = ScreenManager()

        self.home = HomeScreen(name="home")
        self.table = TableScreen(name="table")
        self.email = EmailScreen(name="email")
        self.smtp = SMTPScreen(name="smtp")

        self._build_home()
        self._build_table()
        self._build_email()
        self._build_smtp()

        self.sm.add_widget(self.home)
        self.sm.add_widget(self.table)
        self.sm.add_widget(self.email)
        self.sm.add_widget(self.smtp)

        return self.sm

    def save_export_folder(self):

        if not self.export_folder:
            return

        with open(EXPORT_CONFIG, "w") as f:
            json.dump({"folder": self.export_folder}, f)

    def load_export_folder(self):

        if os.path.exists(EXPORT_CONFIG):
            with open(EXPORT_CONFIG) as f:
                self.export_folder = json.load(f).get("folder")

    def _build_home(self):

        layout = BoxLayout(orientation="vertical", padding=dp(30), spacing=dp(20))

        title = Label(text=APP_TITLE, font_size=26, bold=True)

        self.home_status = Label(text="Gotowy", font_size=16)

        open_btn = PremiumButton(text="📂 Otwórz plik Excel")
        open_btn.bind(on_press=self.open_excel_picker)

        load_btn = PremiumButton(text="📊 Wczytaj dane")
        load_btn.bind(on_press=self.load_full_excel)

        smtp_btn = PremiumButton(text="⚙ Konfiguracja SMTP")
        smtp_btn.bind(on_press=lambda x: setattr(self.sm, "current", "smtp"))

        layout.add_widget(title)
        layout.add_widget(open_btn)
        layout.add_widget(load_btn)
        layout.add_widget(smtp_btn)
        layout.add_widget(self.home_status)

        self.home.add_widget(layout)

    def open_excel_picker(self, _):

        if platform != "android":
            self.home_status.text = "Picker działa tylko na Androidzie"
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
        input_stream = resolver.openInputStream(uri)

        local_file = Path(self.user_data_dir) / "selected.xlsx"

        with open(local_file, "wb") as out:

            buffer = bytearray(4096)

            while True:
                read = input_stream.read(buffer)
                if read == -1:
                    break
                out.write(buffer[:read])

        input_stream.close()

        self.current_file = local_file
        self.home_status.text = "Plik wybrany"

    def load_full_excel(self, _):

        if not self.current_file:
            self._popup("Błąd", "Najpierw wybierz plik")
            return

        from openpyxl import load_workbook

        wb = load_workbook(str(self.current_file), data_only=True)

        sheet = wb.active

        self.full_data = [
            ["" if v is None else str(v) for v in row]
            for row in sheet.iter_rows(values_only=True)
        ]

        wb.close()

        self.filtered_data = self.full_data

        self.display_table()

        self.sm.current = "table"

    def _build_table(self):

        layout = BoxLayout(orientation="vertical", padding=dp(10), spacing=dp(10))

        top = BoxLayout(size_hint=(1, 0.12), spacing=dp(10))

        self.search = TextInput(hint_text="🔎 Wyszukaj...", multiline=False)
        self.search.bind(text=self.filter_data)

        export_btn = PremiumButton(text="📦 Eksport")
        export_btn.bind(
            on_press=lambda x: threading.Thread(
                target=self._export_thread,
                daemon=True
            ).start()
        )

        back_btn = PremiumButton(text="⬅ Powrót")
        back_btn.bind(on_press=lambda x: setattr(self.sm, "current", "home"))

        top.add_widget(self.search)
        top.add_widget(export_btn)
        top.add_widget(back_btn)

        self.scroll = ScrollView()

        self.grid = GridLayout(size_hint=(None, None), spacing=dp(1))
        self.grid.bind(minimum_height=self.grid.setter('height'))
        self.grid.bind(minimum_width=self.grid.setter('width'))

        self.scroll.add_widget(self.grid)

        self.progress = ProgressBar(max=100)

        layout.add_widget(top)
        layout.add_widget(self.scroll)
        layout.add_widget(self.progress)

        self.table.add_widget(layout)

    def filter_data(self, instance, value):

        value = value.lower()

        self.filtered_data = [self.full_data[0]] + [
            row for row in self.full_data[1:]
            if any(value in str(cell).lower() for cell in row)
        ]

        self.display_table()

    def display_table(self):

        self.grid.clear_widgets()

        if not self.filtered_data:
            return

        rows = len(self.filtered_data)
        cols = len(self.filtered_data[0])

        self.grid.cols = cols

        self.grid.width = cols * dp(160)
        self.grid.height = rows * dp(40)

        for row in self.filtered_data:
            for cell in row:

                lbl = Label(
                    text=str(cell),
                    size_hint=(None, None),
                    size=(dp(160), dp(40))
                )

                self.grid.add_widget(lbl)

    def _export_thread(self):

        from openpyxl import Workbook

        wb = Workbook()
        ws = wb.active

        for row in self.filtered_data:
            ws.append(row)

        file = Path(self.user_data_dir) / f"export_{datetime.now().strftime('%H%M%S')}.xlsx"

        wb.save(file)

        Clock.schedule_once(
            lambda dt: self._popup("Eksport", f"Zapisano:\n{file}")
        )

    def _build_email(self):
        pass

    def _build_smtp(self):

        layout = BoxLayout(orientation="vertical", padding=dp(30), spacing=dp(15))

        self.s_server = TextInput(hint_text="SMTP server")
        self.s_port = TextInput(hint_text="Port")
        self.s_email = TextInput(hint_text="Email")
        self.s_pass = TextInput(hint_text="Hasło", password=True)

        save = PremiumButton(text="💾 Zapisz")
        back = PremiumButton(text="⬅ Powrót")

        save.bind(on_press=self.save_smtp)
        back.bind(on_press=lambda x: setattr(self.sm, "current", "home"))

        layout.add_widget(self.s_server)
        layout.add_widget(self.s_port)
        layout.add_widget(self.s_email)
        layout.add_widget(self.s_pass)
        layout.add_widget(save)
        layout.add_widget(back)

        self.smtp.add_widget(layout)

    def save_smtp(self, _):

        config = {
            "server": self.s_server.text.strip(),
            "port": self.s_port.text.strip(),
            "email": self.s_email.text.strip(),
            "password": self.s_pass.text.strip()
        }

        with open(CONFIG_FILE, "w") as f:
            json.dump(config, f)

        self._popup("Zapisano", "SMTP zapisany")

    def _popup(self, title, message):

        content = BoxLayout(orientation="vertical", padding=dp(10))

        content.add_widget(Label(text=message))

        btn = Button(text="OK", size_hint=(1, 0.3))

        popup = Popup(
            title=title,
            content=content,
            size_hint=(0.8, 0.5)
        )

        btn.bind(on_release=popup.dismiss)

        content.add_widget(btn)

        popup.open()


if __name__ == "__main__":
    PaskiFutureApp().run()
