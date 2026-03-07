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

        # folder eksportu
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

        open_btn = PremiumButton(text="ðŸ“‚ OtwÃ³rz plik Excel")
        open_btn.bind(on_press=self.open_excel_picker)

        load_btn = PremiumButton(text="ðŸ“Š Wczytaj dane")
        load_btn.bind(on_press=self.load_full_excel)

        smtp_btn = PremiumButton(text="âš™ Konfiguracja SMTP")
        smtp_btn.bind(on_press=lambda x: setattr(self.sm, "current", "smtp"))

        self.home_status = Label(text="Gotowy", font_size=16)

        layout.add_widget(title)
        layout.add_widget(open_btn)
        layout.add_widget(load_btn)
        layout.add_widget(smtp_btn)
        layout.add_widget(self.home_status)

        self.home.add_widget(layout)


    # ======================================================
    # PICKER EXCEL
    # ======================================================

    def open_excel_picker(self, _):

        if platform != "android":
            self.home_status.text = "Picker dziaÅ‚a tylko Android"
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

        if not self.current_file:
            self._popup("BÅ‚Ä…d", "Najpierw wybierz plik")
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
            hint_text="ðŸ”Ž Wyszukaj...",
            multiline=False
        )
        self.search.bind(text=self.filter_data)

        folder_btn = PremiumButton(text="ðŸ“ Folder")
        folder_btn.bind(on_press=self.pick_export_folder)

        export_btn = PremiumButton(text="ðŸ“¦ Eksport")
        export_btn.bind(on_press=self.export_files)

        email_btn = PremiumButton(text="ðŸ“¬ Email")
        email_btn.bind(on_press=lambda x: setattr(self.sm, "current", "email"))

        back_btn = PremiumButton(text="â¬… PowrÃ³t")
        back_btn.bind(on_press=lambda x: setattr(self.sm, "current", "home"))

        top.add_widget(self.search)
        top.add_widget(folder_btn)
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
    # PICKER FOLDER
    # ======================================================

    def pick_export_folder(self, _):

        if platform != "android":
            return

        from jnius import autoclass
        from android import activity

        PythonActivity = autoclass("org.kivy.android.PythonActivity")
        Intent = autoclass("android.content.Intent")

        intent = Intent(Intent.ACTION_OPEN_DOCUMENT_TREE)

        activity.bind(on_activity_result=self._on_folder_result)
        PythonActivity.mActivity.startActivityForResult(intent, 777)


    def _on_folder_result(self, request_code, result_code, intent):

        if request_code != 777 or not intent:
            return

        from android import activity
        activity.unbind(on_activity_result=self._on_folder_result)

        uri = intent.getData()

        if uri:
            self.export_folder = "/storage/emulated/0/Documents/PaskiFuture"
            self._popup("Folder", "Folder eksportu ustawiony")


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

        for r_index, row in enumerate(self.filtered_data):
            for cell in row:

                lbl = Label(
                    text=str(cell),
                    size_hint=(None, None),
                    size=(dp(160), dp(40))
                )

                self.grid.add_widget(lbl)


    # ======================================================
    # EXPORT
    # ======================================================

    def export_files(self, _):
        threading.Thread(target=self._export_thread).start()


    def _export_thread(self):

        if len(self.filtered_data) < 2:
            return

        documents = self.export_folder or "/storage/emulated/0/Documents/PaskiFuture"
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
            Clock.schedule_once(lambda dt, p=percent: setattr(self.progress, "value", p))

        Clock.schedule_once(lambda dt: self._popup("Sukces", f"Wyeksportowano {done} plikÃ³w"))


    # ======================================================
    # EMAIL
    # ======================================================

    def _build_email(self):

        layout = BoxLayout(orientation="vertical", padding=dp(30), spacing=dp(20))

        send1 = PremiumButton(text="ðŸ“§ WyÅ›lij 1 rekord")
        sendAll = PremiumButton(text="ðŸ“¨ WyÅ›lij hurtowo")
        back = PremiumButton(text="â¬… PowrÃ³t")

        send1.bind(on_press=self.send_single)
        sendAll.bind(on_press=self.send_bulk)
        back.bind(on_press=lambda x: setattr(self.sm, "current", "table"))

        self.email_status = Label(text="Gotowy")

        layout.add_widget(send1)
        layout.add_widget(sendAll)
        layout.add_widget(self.email_status)
        layout.add_widget(back)

        self.email.add_widget(layout)


    def send_single(self, _):
        threading.Thread(target=self._send_one).start()


    def send_bulk(self, _):
        threading.Thread(target=self._send_all).start()


    def _send_one(self):
        if len(self.filtered_data) < 2:
            return
        self._send_email_row(self.filtered_data[1])


    def _send_all(self):
        for row in self.filtered_data[1:]:
            self._send_email_row(row)


    def _send_email_row(self, row):

        if not os.path.exists(CONFIG_FILE):
            return

        with open(CONFIG_FILE) as f:
            config = json.load(f)

        wb = Workbook()
        ws = wb.active
        ws.append(self.full_data[0])
        ws.append(row)

        file = Path(self.user_data_dir) / "temp.xlsx"
        wb.save(file)

        msg = EmailMessage()
        msg["Subject"] = "Dane"
        msg["From"] = config["email"]
        msg["To"] = row[EMAIL_COLUMN_INDEX]
        msg.set_content("W zaÅ‚Ä…czniku dane.")

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
        self.s_pass = TextInput(hint_text="HasÅ‚o", password=True)

        save = PremiumButton(text="ðŸ’¾ Zapisz")
        back = PremiumButton(text="â¬… PowrÃ³t")

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

        data = {
            "server": self.s_server.text,
            "port": self.s_port.text,
            "email": self.s_email.text,
            "password": self.s_pass.text
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
