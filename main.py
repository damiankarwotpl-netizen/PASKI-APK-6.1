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
from openpyxl.styles import Border, Side, Font, Alignment


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
        self.export_folder = None

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

    # ======================================================
    # HOME
    # ======================================================

    def _build_home(self):

        layout = BoxLayout(orientation="vertical", padding=dp(30), spacing=dp(20))

        title = Label(text=APP_TITLE, font_size=26, bold=True)

        open_btn = PremiumButton(text="📂 Otwórz plik Excel")
        open_btn.bind(on_press=self.open_excel_picker)

        load_btn = PremiumButton(text="📊 Wczytaj dane")
        load_btn.bind(on_press=self.load_full_excel)

        smtp_btn = PremiumButton(text="⚙ Konfiguracja SMTP")
        smtp_btn.bind(on_press=lambda x: setattr(self.sm, "current", "smtp"))

        self.home_status = Label(text="Gotowy", font_size=16)

        layout.add_widget(title)
        layout.add_widget(open_btn)
        layout.add_widget(load_btn)
        layout.add_widget(smtp_btn)
        layout.add_widget(self.home_status)

        self.home.add_widget(layout)

    # ======================================================
    # LOAD EXCEL
    # ======================================================

    def load_full_excel(self, _):

        if not self.current_file:
            self._popup("Błąd", "Najpierw wybierz plik")
            return

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

    # ======================================================
    # TABLE
    # ======================================================

    def _build_table(self):

        layout = BoxLayout(orientation="vertical", padding=dp(10), spacing=dp(10))

        top = BoxLayout(size_hint=(1, 0.12), spacing=dp(10))

        self.search = TextInput(
            hint_text="🔎 Wyszukaj...",
            multiline=False
        )
        self.search.bind(text=self.filter_data)

        export_btn = PremiumButton(text="📦 Eksport")
        export_btn.bind(on_press=self.export_files)

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

    # ======================================================
    # FILTER
    # ======================================================

    def filter_data(self, instance, value):
        value = value.lower()

        self.filtered_data = [
            row for row in self.full_data
            if any(value in str(cell).lower() for cell in row)
        ]

        self.display_table()

    # ======================================================
    # TABLE DISPLAY
    # ======================================================

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

    # ======================================================
    # EXPORT (POPRAWIONE RAMKI + NAGŁÓWEK)
    # ======================================================

    def export_files(self, _):
        threading.Thread(target=self._export_thread).start()

    def _export_thread(self):

        if len(self.filtered_data) < 2:
            return

        documents = "/storage/emulated/0/Documents/PaskiFuture"
        os.makedirs(documents, exist_ok=True)

        header = self.full_data[0]
        rows = self.filtered_data[1:]

        border = Border(
            left=Side(style="thin"),
            right=Side(style="thin"),
            top=Side(style="thin"),
            bottom=Side(style="thin")
        )

        total = len(rows)
        done = 0

        for row in rows:

            wb = Workbook()
            ws = wb.active

            ws.append(header)
            ws.append(row)

            for col in range(1, len(header)+1):

                cell = ws.cell(row=1, column=col)
                cell.font = Font(bold=True)
                cell.alignment = Alignment(horizontal="center")
                cell.border = border

            for col in range(1, len(header)+1):

                cell = ws.cell(row=2, column=col)
                cell.alignment = Alignment(horizontal="center")
                cell.border = border

            name = row[1] if len(row) > 1 else "brak"
            now = datetime.now().strftime("%Y%m%d_%H%M%S")

            filepath = os.path.join(documents, f"{name}_{now}.xlsx")
            wb.save(filepath)

            done += 1

            percent = int((done / total) * 100)
            Clock.schedule_once(lambda dt, p=percent: setattr(self.progress, "value", p))

        Clock.schedule_once(lambda dt: self._popup("Sukces", f"Wyeksportowano {done} plików"))

    # ======================================================
    # POPUP
    # ======================================================

    def _popup(self, title, text):
        Popup(
            title=title,
            content=Label(text=text),
            size_hint=(0.8, 0.4)
        ).open()


if __name__ == "__main__":
    PaskiFutureApp().run()
