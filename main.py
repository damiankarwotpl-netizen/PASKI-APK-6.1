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
from kivy.uix.screenmanager import ScreenManager, Screen
from kivy.uix.checkbox import CheckBox

from openpyxl import load_workbook, Workbook
from openpyxl.utils import get_column_letter


APP_TITLE = "Paski Future"
CONFIG_FILE = "smtp_config.json"


class HomeScreen(Screen):
    pass


class TableScreen(Screen):
    pass


class SMTPScreen(Screen):
    pass


class PremiumButton(Button):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.background_normal = ""
        self.background_color = (0.2, 0.4, 0.8, 1)
        self.color = (1, 1, 1, 1)
        self.size_hint_y = None
        self.height = dp(50)
        self.font_size = 16


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
        self.smtp = SMTPScreen(name="smtp")

        self._build_home()
        self._build_table()
        self._build_smtp()

        self.sm.add_widget(self.home)
        self.sm.add_widget(self.table)
        self.sm.add_widget(self.smtp)

        return self.sm


# ---------------- HOME ----------------

    def _build_home(self):

        layout = BoxLayout(orientation="vertical", padding=30, spacing=20)

        title = Label(text=APP_TITLE, font_size=28)

        open_btn = PremiumButton(text="Otwórz Excel")
        open_btn.bind(on_press=self.open_excel_picker)

        load_btn = PremiumButton(text="Wczytaj dane")
        load_btn.bind(on_press=self.load_full_excel)

        smtp_btn = PremiumButton(text="Ustawienia SMTP")
        smtp_btn.bind(on_press=lambda x: setattr(self.sm, "current", "smtp"))

        self.home_status = Label(text="Gotowy")

        layout.add_widget(title)
        layout.add_widget(open_btn)
        layout.add_widget(load_btn)
        layout.add_widget(smtp_btn)
        layout.add_widget(self.home_status)

        self.home.add_widget(layout)


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


# ---------------- DISPLAY TABLE (FIXED FOR PHONE) ----------------

    def display_table(self):

        if not self.filtered_data:
            return

        header = self.filtered_data[0]

        grid = GridLayout(
            cols=len(header),
            spacing=2,
            size_hint=(None, None)
        )

        grid.bind(minimum_height=grid.setter("height"))

        column_width = dp(200)

        for row in self.filtered_data:

            for cell in row:

                lbl = Label(
                    text=str(cell),
                    size_hint=(None, None),
                    width=column_width,
                    height=40,
                    text_size=(column_width - 10, None),
                    halign="left",
                    valign="middle"
                )

                grid.add_widget(lbl)

        grid.width = column_width * len(header)

        scroll = ScrollView(do_scroll_x=True, do_scroll_y=True)

        scroll.add_widget(grid)

        self.table_scroll.clear_widgets()
        self.table_scroll.add_widget(scroll)


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
        def show_column_selector(self):
        from kivy.uix.popup import Popup
        from kivy.uix.boxlayout import BoxLayout
        from kivy.uix.checkbox import CheckBox
        from kivy.uix.label import Label
        from kivy.uix.button import Button

        layout = BoxLayout(orientation="vertical", padding=10, spacing=10)

        self.column_checks = {}

        for col in self.df.columns:
            row = BoxLayout(size_hint_y=None, height=40)

            chk = CheckBox(active=True)
            self.column_checks[col] = chk

            row.add_widget(chk)
            row.add_widget(Label(text=str(col)))

            layout.add_widget(row)

        btn = Button(text="Zatwierdź", size_hint_y=None, height=50)
        btn.bind(on_press=self.save_column_selection)

        layout.add_widget(btn)

        self.col_popup = Popup(
            title="Wybierz kolumny do eksportu",
            content=layout,
            size_hint=(0.9,0.9)
        )

        self.col_popup.open()


    def save_column_selection(self, instance):
        self.selected_columns = [
            col for col, chk in self.column_checks.items() if chk.active
        ]

        self.col_popup.dismiss()


    def show_email_selector(self):
        from kivy.uix.popup import Popup
        from kivy.uix.boxlayout import BoxLayout
        from kivy.uix.checkbox import CheckBox
        from kivy.uix.label import Label
        from kivy.uix.button import Button

        layout = BoxLayout(orientation="vertical", padding=10, spacing=10)

        self.email_checks = {}

        for col in self.df.columns:
            row = BoxLayout(size_hint_y=None, height=40)

            chk = CheckBox(group="email")
            self.email_checks[col] = chk

            row.add_widget(chk)
            row.add_widget(Label(text=str(col)))

            layout.add_widget(row)

        btn = Button(text="Zatwierdź", size_hint_y=None, height=50)
        btn.bind(on_press=self.save_email_column)

        layout.add_widget(btn)

        self.email_popup = Popup(
            title="Wybierz kolumnę email",
            content=layout,
            size_hint=(0.9,0.9)
        )

        self.email_popup.open()


    def save_email_column(self, instance):
        for col, chk in self.email_checks.items():
            if chk.active:
                self.email_column = col

        self.email_popup.dismiss()


    def export_excel(self):
        import pandas as pd
        from openpyxl import load_workbook
        from openpyxl.styles import Font, Border, Side
        from openpyxl.utils import get_column_letter

        if not hasattr(self, "selected_columns"):
            self.selected_columns = self.df.columns

        export_df = self.df[self.selected_columns]

        path = "export.xlsx"
        export_df.to_excel(path, index=False)

        wb = load_workbook(path)
        ws = wb.active

        bold = Font(bold=True)
        border = Border(
            left=Side(style="thin"),
            right=Side(style="thin"),
            top=Side(style="thin"),
            bottom=Side(style="thin")
        )

        for cell in ws[1]:
            cell.font = bold
            cell.border = border

        for row in ws.iter_rows():
            for cell in row:
                cell.border = border

        for col in ws.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)

            for cell in col:
                if cell.value:
                    max_len = max(max_len, len(str(cell.value)))

            ws.column_dimensions[col_letter].width = max_len + 4

        wb.save(path)


    def send_emails(self):
        import smtplib
        from email.mime.text import MIMEText

        if not hasattr(self, "email_column"):
            return

        smtp_server = self.smtp_server
        smtp_user = self.smtp_user
        smtp_pass = self.smtp_pass

        server = smtplib.SMTP_SSL(smtp_server, 465)
        server.login(smtp_user, smtp_pass)

        for _, row in self.df.iterrows():
            email = row[self.email_column]

            msg = MIMEText("Test message")
            msg["Subject"] = "Mail"
            msg["From"] = smtp_user
            msg["To"] = email

            try:
                server.sendmail(smtp_user, email, msg.as_string())
            except:
                pass

        server.quit()


class FuturePaskiApp(App):

    def build(self):

        root = BoxLayout(
            orientation="vertical",
            padding=10,
            spacing=10
        )

        table = ScrollView()

        root.add_widget(table)

        return root


if __name__ == "__main__":
    FuturePaskiApp().run()
