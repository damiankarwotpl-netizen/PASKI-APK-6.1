import os
import threading
from datetime import datetime
from pathlib import Path

from kivy.app import App
from kivy.metrics import dp
from kivy.clock import Clock
from kivy.utils import platform
from kivy.core.window import Window

from kivy.uix.boxlayout import BoxLayout
from kivy.uix.button import Button
from kivy.uix.label import Label
from kivy.uix.popup import Popup
from kivy.uix.scrollview import ScrollView
from kivy.uix.gridlayout import GridLayout
from kivy.uix.textinput import TextInput
from kivy.uix.checkbox import CheckBox
from kivy.uix.progressbar import ProgressBar
from kivy.uix.screenmanager import ScreenManager, Screen


APP_TITLE = "Future 9.0 ULTRA PRO"


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
        self.background_color = (0.2, 0.4, 0.9, 1)
        self.color = (1, 1, 1, 1)
        self.size_hint_y = None
        self.height = dp(50)


class FutureApp(App):

    def build(self):

        Window.clearcolor = (0.08, 0.1, 0.15, 1)

        self.full_data = []
        self.filtered_data = []
        self.current_file = None

        self.export_columns = []
        self.email_columns = []

        self.sm = ScreenManager()

        self.home = HomeScreen(name="home")
        self.table = TableScreen(name="table")
        self.email = EmailScreen(name="email")
        self.smtp = SMTPScreen(name="smtp")

        self.build_home()
        self.build_table()
        self.build_email()
        self.build_smtp()

        self.sm.add_widget(self.home)
        self.sm.add_widget(self.table)
        self.sm.add_widget(self.email)
        self.sm.add_widget(self.smtp)

        return self.sm


# -----------------------------
# HOME
# -----------------------------

    def build_home(self):

        layout = BoxLayout(orientation="vertical", padding=dp(25), spacing=dp(20))

        title = Label(text=APP_TITLE, font_size=26)

        open_btn = PremiumButton(text="📂 Otwórz Excel")
        open_btn.bind(on_press=self.open_picker)

        load_btn = PremiumButton(text="📊 Wczytaj dane")
        load_btn.bind(on_press=self.load_excel)

        smtp_btn = PremiumButton(text="⚙ SMTP")
        smtp_btn.bind(on_press=lambda x: setattr(self.sm, "current", "smtp"))

        self.status = Label(text="Gotowy")

        layout.add_widget(title)
        layout.add_widget(open_btn)
        layout.add_widget(load_btn)
        layout.add_widget(smtp_btn)
        layout.add_widget(self.status)

        self.home.add_widget(layout)


# -----------------------------
# ANDROID PICKER
# -----------------------------

    def open_picker(self, _):

        if platform != "android":
            self.status.text = "Picker tylko Android"
            return

        from jnius import autoclass
        from android import activity

        PythonActivity = autoclass("org.kivy.android.PythonActivity")
        Intent = autoclass("android.content.Intent")

        intent = Intent(Intent.ACTION_OPEN_DOCUMENT)
        intent.setType("*/*")
        intent.addCategory(Intent.CATEGORY_OPENABLE)

        activity.bind(on_activity_result=self.picker_result)

        PythonActivity.mActivity.startActivityForResult(intent, 1001)


    def picker_result(self, request_code, result_code, intent):

        if request_code != 1001 or not intent:
            return

        from android import activity
        activity.unbind(on_activity_result=self.picker_result)

        from jnius import autoclass

        PythonActivity = autoclass("org.kivy.android.PythonActivity")

        resolver = PythonActivity.mActivity.getContentResolver()

        uri = intent.getData()

        stream = resolver.openInputStream(uri)

        local = Path(self.user_data_dir) / "excel.xlsx"

        with open(local, "wb") as f:

            buffer = bytearray(4096)

            while True:

                r = stream.read(buffer)

                if r == -1:
                    break

                f.write(buffer[:r])

        stream.close()

        self.current_file = local

        self.status.text = "Plik wybrany"


# -----------------------------
# LOAD EXCEL
# -----------------------------

    def load_excel(self, _):

        if not self.current_file:
            self.popup("Błąd", "Najpierw wybierz plik")
            return

        from openpyxl import load_workbook

        wb = load_workbook(self.current_file, data_only=True)

        sheet = wb.active

        self.full_data = [
            ["" if v is None else str(v) for v in row]
            for row in sheet.iter_rows(values_only=True)
        ]

        wb.close()

        self.filtered_data = self.full_data

        self.show_table()

        self.sm.current = "table"


# -----------------------------
# TABLE UI
# -----------------------------

    def build_table(self):

        layout = BoxLayout(orientation="vertical", padding=dp(10), spacing=dp(10))

        top = BoxLayout(size_hint=(1, 0.12), spacing=dp(8))

        self.search = TextInput(hint_text="Szukaj...", multiline=False)
        self.search.bind(text=self.filter_data)

        export_btn = PremiumButton(text="Export")
        export_btn.bind(on_press=self.export_popup)

        email_btn = PremiumButton(text="Email")
        email_btn.bind(on_press=lambda x: setattr(self.sm, "current", "email"))

        back_btn = PremiumButton(text="Powrót")
        back_btn.bind(on_press=lambda x: setattr(self.sm, "current", "home"))

        top.add_widget(self.search)
        top.add_widget(export_btn)
        top.add_widget(email_btn)
        top.add_widget(back_btn)

        self.scroll = ScrollView()

        self.grid = GridLayout(size_hint=(None, None))
        self.grid.bind(minimum_height=self.grid.setter("height"))
        self.grid.bind(minimum_width=self.grid.setter("width"))

        self.scroll.add_widget(self.grid)

        self.progress = ProgressBar(max=100, size_hint=(1, 0.05))

        layout.add_widget(top)
        layout.add_widget(self.scroll)
        layout.add_widget(self.progress)

        self.table.add_widget(layout)


# -----------------------------
# TABLE DISPLAY
# -----------------------------

    def show_table(self):

        self.grid.clear_widgets()

        if not self.filtered_data:
            return

        rows = len(self.filtered_data)
        cols = len(self.filtered_data[0])

        w = dp(160)
        h = dp(42)

        self.grid.cols = cols
        self.grid.width = cols * w
        self.grid.height = rows * h

        for row in self.filtered_data:

            for cell in row:

                self.grid.add_widget(Label(
                    text=str(cell),
                    size_hint=(None, None),
                    size=(w, h)
                ))


# -----------------------------
# SEARCH
# -----------------------------

    def filter_data(self, instance, value):

        value = value.lower()

        self.filtered_data = [

            row for row in self.full_data

            if any(value in str(cell).lower() for cell in row)

        ]

        self.show_table()


# -----------------------------
# EXPORT POPUP (CHECKBOX)
# -----------------------------

    def export_popup(self, _):

        if not self.full_data:
            return

        header = self.full_data[0]

        box = BoxLayout(orientation="vertical", spacing=dp(8))

        checks = []

        for i, name in enumerate(header):

            row = BoxLayout()

            cb = CheckBox()

            checks.append((i, cb))

            row.add_widget(cb)
            row.add_widget(Label(text=name))

            box.add_widget(row)

        btn = PremiumButton(text="Start export")

        def start(_):

            self.export_columns = [i for i, c in checks if c.active]

            popup.dismiss()

            threading.Thread(target=self.export_excel).start()

        btn.bind(on_press=start)

        box.add_widget(btn)

        popup = Popup(title="Wybierz kolumny", content=box, size_hint=(0.9, 0.9))

        popup.open()
        # -----------------------------
# EMAIL SCREEN
# -----------------------------

    def build_email(self):

        layout = BoxLayout(orientation="vertical", padding=dp(20), spacing=dp(15))

        title = Label(text="Email Sender", font_size=22)

        select_col = PremiumButton(text="Wybierz kolumnę email")
        select_col.bind(on_press=self.select_email_column)

        send_btn = PremiumButton(text="Wyślij emaile")
        send_btn.bind(on_press=self.send_emails)

        back = PremiumButton(text="Powrót")
        back.bind(on_press=lambda x: setattr(self.sm, "current", "table"))

        self.email_status = Label(text="")

        layout.add_widget(title)
        layout.add_widget(select_col)
        layout.add_widget(send_btn)
        layout.add_widget(back)
        layout.add_widget(self.email_status)

        self.email.add_widget(layout)


# -----------------------------
# EMAIL COLUMN CHECKBOX
# -----------------------------

    def select_email_column(self, _):

        if not self.full_data:
            self.popup("Błąd", "Brak danych")
            return

        header = self.full_data[0]

        box = BoxLayout(orientation="vertical", spacing=dp(10))

        checks = []

        for i, col in enumerate(header):

            row = BoxLayout()

            cb = CheckBox(group="email")

            checks.append((i, cb))

            row.add_widget(cb)
            row.add_widget(Label(text=str(col)))

            box.add_widget(row)

        btn = PremiumButton(text="OK")

        def save(_):

            for i, c in checks:
                if c.active:
                    self.email_columns = [i]

            popup.dismiss()

        btn.bind(on_press=save)

        box.add_widget(btn)

        popup = Popup(
            title="Wybierz kolumnę email",
            content=box,
            size_hint=(0.9, 0.9)
        )

        popup.open()


# -----------------------------
# SMTP SCREEN
# -----------------------------

    def build_smtp(self):

        layout = BoxLayout(orientation="vertical", padding=dp(20), spacing=dp(10))

        title = Label(text="SMTP", font_size=22)

        self.smtp_server = TextInput(hint_text="SMTP server", multiline=False)
        self.smtp_port = TextInput(hint_text="Port", multiline=False)
        self.smtp_user = TextInput(hint_text="Email", multiline=False)
        self.smtp_pass = TextInput(hint_text="Hasło", multiline=False, password=True)

        save = PremiumButton(text="Zapisz")

        save.bind(on_press=self.save_smtp)

        back = PremiumButton(text="Powrót")

        back.bind(on_press=lambda x: setattr(self.sm, "current", "home"))

        layout.add_widget(title)
        layout.add_widget(self.smtp_server)
        layout.add_widget(self.smtp_port)
        layout.add_widget(self.smtp_user)
        layout.add_widget(self.smtp_pass)
        layout.add_widget(save)
        layout.add_widget(back)

        self.smtp.add_widget(layout)


# -----------------------------
# SAVE SMTP
# -----------------------------

    def save_smtp(self, _):

        import json

        data = {

            "server": self.smtp_server.text,
            "port": self.smtp_port.text,
            "user": self.smtp_user.text,
            "pass": self.smtp_pass.text

        }

        path = Path(self.user_data_dir) / "smtp.json"

        with open(path, "w") as f:
            json.dump(data, f)

        self.popup("OK", "SMTP zapisane")


# -----------------------------
# LOAD SMTP
# -----------------------------

    def load_smtp(self):

        import json

        path = Path(self.user_data_dir) / "smtp.json"

        if not path.exists():
            return None

        with open(path) as f:
            return json.load(f)


# -----------------------------
# SEND EMAILS
# -----------------------------

    def send_emails(self, _):

        if not self.email_columns:

            self.popup("Błąd", "Wybierz kolumnę email")

            return

        threading.Thread(target=self._email_thread).start()


    def _email_thread(self):

        import smtplib
        from email.message import EmailMessage

        smtp = self.load_smtp()

        if not smtp:

            Clock.schedule_once(
                lambda dt: self.popup("Błąd", "SMTP nie skonfigurowane")
            )

            return

        try:

            server = smtplib.SMTP(
                smtp["server"],
                int(smtp["port"])
            )

            server.starttls()

            server.login(
                smtp["user"],
                smtp["pass"]
            )

        except Exception as e:

            Clock.schedule_once(
                lambda dt: self.popup("SMTP error", str(e))
            )

            return

        col = self.email_columns[0]

        rows = self.full_data[1:]

        total = len(rows)

        sent = 0

        for i, row in enumerate(rows):

            if col >= len(row):
                continue

            email = row[col]

            if not email or "@" not in str(email):
                continue

            msg = EmailMessage()

            msg["Subject"] = "Informacja"
            msg["From"] = smtp["user"]
            msg["To"] = email

            msg.set_content("Wiadomość wygenerowana automatycznie")

            try:

                server.send_message(msg)

                sent += 1

            except:
                pass

            progress = int((i + 1) / total * 100)

            Clock.schedule_once(
                lambda dt, p=progress: setattr(self.progress, "value", p)
            )

        server.quit()

        Clock.schedule_once(
            lambda dt: self.popup("Gotowe", f"Wysłano {sent} emaili")
        )


# -----------------------------
# EXPORT THREAD
# -----------------------------

    def export_excel(self):

        from openpyxl import Workbook

        rows = self.filtered_data

        if not rows:
            return

        folder = Path("/storage/emulated/0/Documents/FutureExport")

        folder.mkdir(parents=True, exist_ok=True)

        header = rows[0]

        if self.export_columns:
            header = [header[i] for i in self.export_columns]

        total = len(rows) - 1

        for i, row in enumerate(rows[1:]):

            wb = Workbook()

            ws = wb.active

            ws.append(header)

            if self.export_columns:

                row = [row[i] for i in self.export_columns]

            ws.append(row)

            # autosize kolumn

            for col in ws.columns:

                length = 0

                for cell in col:

                    if cell.value:

                        length = max(
                            length,
                            len(str(cell.value))
                        )

                ws.column_dimensions[
                    col[0].column_letter
                ].width = length + 4

            name = row[0] if row else "file"

            now = datetime.now().strftime("%Y%m%d_%H%M%S")

            file = folder / f"{name}_{now}.xlsx"

            wb.save(file)

            percent = int((i + 1) / total * 100)

            Clock.schedule_once(
                lambda dt, p=percent: setattr(self.progress, "value", p)
            )

        Clock.schedule_once(
            lambda dt: self.popup("Export", "Zakończony")
        )


# -----------------------------
# POPUP
# -----------------------------

    def popup(self, title, text):

        box = BoxLayout(orientation="vertical", padding=dp(20))

        box.add_widget(Label(text=text))

        btn = PremiumButton(text="OK")

        box.add_widget(btn)

        popup = Popup(
            title=title,
            content=box,
            size_hint=(0.7, 0.4)
        )

        btn.bind(on_press=popup.dismiss)

        popup.open()


# -----------------------------
# APP START
# -----------------------------

from ultra_patch import apply_patch

if __name__ == "__main__":

    app = FutureApp()

    apply_patch(__import__(__name__), app)

    app.run()
