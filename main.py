import os
import json
import smtplib
import threading
from pathlib import Path
from datetime import datetime
from email.message import EmailMessage

from kivy.app import App
from kivy.clock import Clock
from kivy.core.window import Window
from kivy.metrics import dp
from kivy.utils import platform
from kivy.graphics import Color, Rectangle

from kivy.uix.boxlayout import BoxLayout
from kivy.uix.button import Button
from kivy.uix.label import Label
from kivy.uix.scrollview import ScrollView
from kivy.uix.gridlayout import GridLayout
from kivy.uix.textinput import TextInput
from kivy.uix.popup import Popup
from kivy.uix.progressbar import ProgressBar
from kivy.uix.screenmanager import ScreenManager, Screen

from openpyxl import load_workbook, Workbook
from plyer import filechooser


APP_TITLE = "Paski Future 6.1 STABLE PREMIUM"
CONFIG_FILE = "smtp_config.json"
EMAIL_COLUMN_INDEX = 3


class HomeScreen(Screen): pass
class TableScreen(Screen): pass
class EmailScreen(Screen): pass
class SMTPScreen(Screen): pass


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
        self.export_dir = None

        self.sm = ScreenManager()

        self.home = HomeScreen(name="home")
        self.table = TableScreen(name="table")
        self.email = EmailScreen(name="email")
        self.smtp = SMTPScreen(name="smtp")

        self.sm.add_widget(self.home)
        self.sm.add_widget(self.table)
        self.sm.add_widget(self.email)
        self.sm.add_widget(self.smtp)

        self._build_home()
        self._build_table()
        self._build_email()
        self._build_smtp()

        return self.sm

    # ======================================================
    # HOME
    # ======================================================

    def _build_home(self):

        layout = BoxLayout(orientation="vertical", padding=dp(30), spacing=dp(20))

        load = PremiumButton(text="📂 Wczytaj Excel")
        load.bind(on_press=self.load_file)

        layout.add_widget(load)

        self.home.add_widget(layout)

    # ======================================================
    # LOAD FILE
    # ======================================================

    def load_file(self, _):
        filechooser.open_file(on_selection=self._file_selected)

    def _file_selected(self, selection):

        if not selection:
            return

        self.current_file = selection[0]

        wb = load_workbook(self.current_file)
        ws = wb.active

        self.full_data = []
        for row in ws.iter_rows(values_only=True):
            self.full_data.append([str(cell) if cell else "" for cell in row])

        self.filtered_data = self.full_data.copy()

        self.display_table()
        self.sm.current = "table"

    # ======================================================
    # TABLE
    # ======================================================

    def _build_table(self):

        root = BoxLayout(orientation="vertical", spacing=dp(10), padding=dp(10))

        buttons = BoxLayout(size_hint_y=None, height=dp(50), spacing=dp(10))

        export = PremiumButton(text="📦 Eksport")
        export.bind(on_press=self.export_files)

        email = PremiumButton(text="📧 Email")
        email.bind(on_press=lambda x: setattr(self.sm, "current", "email"))

        buttons.add_widget(export)
        buttons.add_widget(email)

        root.add_widget(buttons)

        self.progress = ProgressBar(max=100, value=0, size_hint_y=None, height=dp(20))
        root.add_widget(self.progress)

        self.scroll = ScrollView()
        self.grid = GridLayout(cols=1, size_hint_y=None)
        self.grid.bind(minimum_height=self.grid.setter("height"))

        self.scroll.add_widget(self.grid)
        root.add_widget(self.scroll)

        self.table.add_widget(root)

    def display_table(self):

        self.grid.clear_widgets()

        for row in self.filtered_data:

            txt = " | ".join(str(cell) for cell in row)
            lbl = Label(text=txt, size_hint_y=None, height=dp(30))

            self.grid.add_widget(lbl)

    # ======================================================
    # EXPORT
    # ======================================================

    def export_files(self, _):

        if not self.export_dir:
            filechooser.choose_dir(on_selection=self._set_export_dir)
        else:
            threading.Thread(target=self._export_thread).start()

    def _set_export_dir(self, selection):

        if selection:
            self.export_dir = selection[0]
            threading.Thread(target=self._export_thread).start()

    def _export_thread(self):

        if len(self.filtered_data) < 2:
            return

        documents = self.export_dir
        os.makedirs(documents, exist_ok=True)

        header = self.full_data[0]
        rows = self.filtered_data[1:]

        total = len(rows)
        done = 0

        for row in rows:

            wb = Workbook()
            ws = wb.active
            ws.append(header)
            ws.append(row)

            name = row[1] if len(row) > 1 else "brak"

            now = datetime.now().strftime("%Y%m%d_%H%M%S")

            filepath = os.path.join(documents, f"{name}_{now}.xlsx")
            wb.save(filepath)

            done += 1
            percent = int((done / total) * 100)

            Clock.schedule_once(
                lambda dt, p=percent: setattr(self.progress, "value", p)
            )

        Clock.schedule_once(
            lambda dt: self._popup("Sukces", f"Wyeksportowano {done} plików")
        )

    # ======================================================
    # EMAIL
    # ======================================================

    def _build_email(self):

        layout = BoxLayout(orientation="vertical", padding=dp(30), spacing=dp(20))

        send1 = PremiumButton(text="📧 Wyślij 1 rekord")
        sendAll = PremiumButton(text="📨 Wyślij hurtowo")
        back = PremiumButton(text="⬅ Powrót")

        send1.bind(on_press=self.send_single)
        sendAll.bind(on_press=self.send_bulk)
        back.bind(on_press=lambda x: setattr(self.sm, "current", "table"))

        layout.add_widget(send1)
        layout.add_widget(sendAll)
        layout.add_widget(back)

        self.email.add_widget(layout)

    def send_single(self, _):
        pass

    def send_bulk(self, _):
        pass

    # ======================================================
    # SMTP
    # ======================================================

    def _build_smtp(self):

        layout = BoxLayout(orientation="vertical", padding=dp(30), spacing=dp(20))

        back = PremiumButton(text="⬅ Powrót")
        back.bind(on_press=lambda x: setattr(self.sm, "current", "home"))

        layout.add_widget(back)

        self.smtp.add_widget(layout)

    # ======================================================
    # POPUP
    # ======================================================

    def _popup(self, title, text):

        box = BoxLayout(orientation="vertical", padding=dp(20), spacing=dp(20))
        box.add_widget(Label(text=text))

        btn = PremiumButton(text="OK")
        popup = Popup(title=title, content=box, size_hint=(0.7, 0.4))

        btn.bind(on_press=popup.dismiss)
        box.add_widget(btn)

        popup.open()


if __name__ == "__main__":
    PaskiFutureApp().run()
