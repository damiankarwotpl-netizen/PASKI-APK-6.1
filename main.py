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
from kivy.uix.togglebutton import ToggleButton


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
        self.selected_columns = []

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
# FORMAT EXCEL
# ======================================================

    def format_excel(self, ws):

        from openpyxl.styles import Border, Side, Font, Alignment, PatternFill

        thin = Side(style="thin")
        border = Border(left=thin, right=thin, top=thin, bottom=thin)

        header_font = Font(bold=True, size=12)
        header_fill = PatternFill("solid", fgColor="D9E1F2")
        center = Alignment(horizontal="center", vertical="center")

        max_row = ws.max_row
        max_col = ws.max_column

        for col in range(1, max_col + 1):

            cell = ws.cell(row=1, column=col)

            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = center

        for row in ws.iter_rows(min_row=1, max_row=max_row, min_col=1, max_col=max_col):

            for cell in row:
                cell.border = border


# ======================================================
# COLUMN SELECTOR
# ======================================================

    def show_column_selector(self):

        if not self.full_data:
            return

        header = self.full_data[0]

        layout = GridLayout(cols=1, spacing=10, padding=10, size_hint_y=None)
        layout.bind(minimum_height=layout.setter("height"))

        scroll = ScrollView()
        scroll.add_widget(layout)

        for i, name in enumerate(header):

            btn = ToggleButton(
                text=name,
                size_hint_y=None,
                height=dp(40)
            )

            if i in self.selected_columns:
                btn.state = "down"

            btn.bind(on_press=lambda x, idx=i: self.toggle_column(idx))

            layout.add_widget(btn)

        close = Button(text="OK", size_hint_y=None, height=dp(50))
        layout.add_widget(close)

        popup = Popup(
            title="Wybierz kolumny",
            content=scroll,
            size_hint=(0.8, 0.8)
        )

        close.bind(on_press=popup.dismiss)

        popup.open()


    def toggle_column(self, index):

        if index in self.selected_columns:
            self.selected_columns.remove(index)
        else:
            self.selected_columns.append(index)


# ======================================================
# GET EMAIL COLUMN
# ======================================================

    def get_email_column(self):

        if not os.path.exists(CONFIG_FILE):
            return EMAIL_COLUMN_INDEX

        with open(CONFIG_FILE) as f:
            cfg = json.load(f)

        return int(cfg.get("email_column", EMAIL_COLUMN_INDEX))


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
# ANDROID FILE PICKER
# ======================================================

    def open_excel_picker(self, _):

        if platform != "android":
            self.home_status.text = "Picker działa tylko Android"
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

        with open(local_file, "wb") as output:

            buffer = bytearray(4096)

            while True:

                bytes_read = input_stream.read(buffer)

                if bytes_read == -1:
                    break

                output.write(buffer[:bytes_read])

        input_stream.close()

        self.current_file = local_file
        self.home_status.text = "Plik wybrany"


# ======================================================
# LOAD EXCEL
# ======================================================

    def load_full_excel(self, _):

        from openpyxl import load_workbook

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

        if not self.selected_columns:
            self.selected_columns = list(range(len(self.full_data[0])))

        self.filtered_data = self.full_data

        self.display_table()
        self.sm.current = "table"


# ======================================================
# TABLE
# ======================================================

    def _build_table(self):

        layout = BoxLayout(orientation="vertical", padding=dp(10), spacing=dp(10))

        top = BoxLayout(size_hint=(1, 0.12), spacing=dp(10))

        self.search = TextInput(hint_text="🔎 Wyszukaj...", multiline=False)
        self.search.bind(text=self.filter_data)

        column_btn = PremiumButton(text="📑 Kolumny")
        column_btn.bind(on_press=lambda x: self.show_column_selector())

        export_btn = PremiumButton(text="📦 Eksport")
        export_btn.bind(on_press=self.export_files)

        email_btn = PremiumButton(text="📬 Email")
        email_btn.bind(on_press=lambda x: setattr(self.sm, "current", "email"))

        back_btn = PremiumButton(text="⬅ Powrót")
        back_btn.bind(on_press=lambda x: setattr(self.sm, "current", "home"))

        top.add_widget(self.search)
        top.add_widget(column_btn)
        top.add_widget(export_btn)
        top.add_widget(email_btn)
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
# EXPORT
# ======================================================

    def export_files(self, _):

        threading.Thread(target=self._export_thread).start()


    def _export_thread(self):

        from openpyxl import Workbook

        if len(self.filtered_data) < 2:
            return

        documents = "/storage/emulated/0/Documents/PaskiFuture"
        os.makedirs(documents, exist_ok=True)

        header = [self.full_data[0][i] for i in self.selected_columns]

        rows = self.filtered_data[1:]

        for row in rows:

            data = [row[i] for i in self.selected_columns]

            wb = Workbook()
            ws = wb.active

            ws.append(header)
            ws.append(data)

            self.format_excel(ws)

            name = row[1] if len(row) > 1 else "brak"
            now = datetime.now().strftime("%Y%m%d_%H%M%S")

            filepath = os.path.join(documents, f"{name}_{now}.xlsx")

            wb.save(filepath)


# ======================================================
# EMAIL
# ======================================================

    def _send_email_row(self, row):

        import smtplib
        from openpyxl import Workbook
        from email.message import EmailMessage

        if not os.path.exists(CONFIG_FILE):
            return

        with open(CONFIG_FILE) as f:
            config = json.load(f)

        email_col = self.get_email_column()

        header = [self.full_data[0][i] for i in self.selected_columns]
        data = [row[i] for i in self.selected_columns]

        wb = Workbook()
        ws = wb.active

        ws.append(header)
        ws.append(data)

        self.format_excel(ws)

        file = Path(self.user_data_dir) / "temp.xlsx"
        wb.save(file)

        msg = EmailMessage()

        msg["Subject"] = "Dane"
        msg["From"] = config["email"]
        msg["To"] = row[email_col]

        msg.set_content("W załączniku dane.")

        with open(file, "rb") as f:

            msg.add_attachment(
                f.read(),
                maintype="application",
                subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                filename="dane.xlsx"
            )

        server = smtplib.SMTP(config["server"], int(config["port"]))

        server.starttls()
        server.login(config["email"], config["password"])
        server.send_message(msg)
        server.quit()


# ======================================================
# SMTP
# ======================================================

    def _build_smtp(self):

        layout = BoxLayout(orientation="vertical", padding=dp(30), spacing=dp(15))

        self.s_server = TextInput(hint_text="SMTP Server")
        self.s_port = TextInput(hint_text="Port")
        self.s_email = TextInput(hint_text="Email")
        self.s_pass = TextInput(hint_text="Hasło", password=True)
        self.s_email_col = TextInput(hint_text="Kolumna Email")

        save = PremiumButton(text="💾 Zapisz")
        back = PremiumButton(text="⬅ Powrót")

        save.bind(on_press=self.save_smtp)
        back.bind(on_press=lambda x: setattr(self.sm, "current", "home"))

        layout.add_widget(self.s_server)
        layout.add_widget(self.s_port)
        layout.add_widget(self.s_email)
        layout.add_widget(self.s_pass)
        layout.add_widget(self.s_email_col)
        layout.add_widget(save)
        layout.add_widget(back)

        self.smtp.add_widget(layout)


    def save_smtp(self, _):

        data = {
            "server": self.s_server.text,
            "port": self.s_port.text,
            "email": self.s_email.text,
            "password": self.s_pass.text,
            "email_column": self.s_email_col.text
        }

        with open(CONFIG_FILE, "w") as f:
            json.dump(data, f)

        self._popup("Sukces", "SMTP zapisane")


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
