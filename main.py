import os
import json
import threading
from pathlib import Path
from datetime import datetime
import smtplib
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
EXPORT_CONFIG = "export_folder.json"

# -------------------------------
# Screens
# -------------------------------
class HomeScreen(Screen): pass
class TableScreen(Screen): pass
class EmailScreen(Screen): pass
class SMTPScreen(Screen): pass

# -------------------------------
# Styled Button
# -------------------------------
class PremiumButton(Button):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.background_normal = ""
        self.background_color = (0.2, 0.4, 0.8, 1)
        self.color = (1, 1, 1, 1)
        self.font_size = 16
        self.size_hint_y = None
        self.height = dp(48)

# -------------------------------
# Main App
# -------------------------------
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

        # Screen Manager
        self.sm = ScreenManager()
        self.home = HomeScreen(name="home")
        self.table = TableScreen(name="table")
        self.email = EmailScreen(name="email")
        self.smtp = SMTPScreen(name="smtp")

        # Build screens
        self._build_home()
        self._build_table()
        self._build_email()
        self._build_smtp()

        self.sm.add_widget(self.home)
        self.sm.add_widget(self.table)
        self.sm.add_widget(self.email)
        self.sm.add_widget(self.smtp)

        return self.sm

    # ---------------------------
    # Export folder save/load
    # ---------------------------
    def save_export_folder(self):
        if not self.export_folder:
            return
        with open(EXPORT_CONFIG, "w") as f:
            json.dump({"folder": self.export_folder}, f)

    def load_export_folder(self):
        if os.path.exists(EXPORT_CONFIG):
            with open(EXPORT_CONFIG) as f:
                self.export_folder = json.load(f).get("folder")

    # ---------------------------
    # Home Screen
    # ---------------------------
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

        email_excel_btn = PremiumButton(text="📧 Wybierz Email Excel")
        email_excel_btn.bind(on_press=self.select_email_excel)

        layout.add_widget(title)
        layout.add_widget(open_btn)
        layout.add_widget(load_btn)
        layout.add_widget(email_excel_btn)
        layout.add_widget(smtp_btn)
        layout.add_widget(self.home_status)

        self.home.add_widget(layout)

    # ---------------------------
    # Android Excel Picker
    # ---------------------------
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
        os.makedirs(self.user_data_dir, exist_ok=True)
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

    # ---------------------------
    # Load Excel
    # ---------------------------
    def load_full_excel(self, _):
        if not self.current_file:
            self._popup("Błąd", "Najpierw wybierz plik")
            return
        wb = load_workbook(str(self.current_file), data_only=True)
        sheet = wb.active
        self.full_data = [["" if v is None else str(v) for v in row] for row in sheet.iter_rows(values_only=True)]
        wb.close()
        self.filtered_data = self.full_data
        self.display_table()
        self.sm.current = "table"

    # ---------------------------
    # Table Screen
    # ---------------------------
    def _build_table(self):
        layout = BoxLayout(orientation="vertical", padding=dp(10), spacing=dp(10))
        top = BoxLayout(size_hint=(1, 0.12), spacing=dp(10))

        self.search = TextInput(hint_text="🔎 Wyszukaj...", multiline=False)
        self.search.bind(text=self.filter_data)

        folder_btn = PremiumButton(text="📁 Folder")
        folder_btn.bind(on_press=self.pick_export_folder)

        export_btn = PremiumButton(text="📦 Eksport")
        export_btn.bind(on_press=lambda x: threading.Thread(target=self._export_thread, daemon=True).start())

        email_btn = PremiumButton(text="📬 Email")
        email_btn.bind(on_press=lambda x: setattr(self.sm, "current", "email"))

        back_btn = PremiumButton(text="⬅ Powrót")
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

    # ---------------------------
    # Pick Export Folder
    # ---------------------------
    def pick_export_folder(self, _):
        if platform != "android":
            self._popup("Błąd", "Folder picker działa tylko na Androidzie")
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
        if not uri:
            return
        uri_string = uri.toString()
        if "primary:" in uri_string:
            folder = uri_string.split("primary:")[1]
            path = f"/storage/emulated/0/{folder}"
        else:
            path = "/storage/emulated/0"
        os.makedirs(path, exist_ok=True)
        self.export_folder = path
        self.save_export_folder()
        self._popup("Folder ustawiony", path)

    # ---------------------------
    # Filter Table
    # ---------------------------
    def filter_data(self, instance, value):
        value = value.lower()
        self.filtered_data = [self.full_data[0]] + [
            row for row in self.full_data[1:]
            if any(value in str(cell).lower() for cell in row)
        ]
        self.display_table()

    # ---------------------------
    # Display Table
    # ---------------------------
    def display_table(self):
        self.grid.clear_widgets()
        if not self.filtered_data:
            return
        rows, cols = len(self.filtered_data), len(self.filtered_data[0])
        self.grid.cols = cols
        self.grid.width = cols * dp(160)
        self.grid.height = rows * dp(40)
        for row in self.filtered_data:
            for cell in row:
                lbl = Label(text=str(cell), size_hint=(None, None), size=(dp(160), dp(40)))
                self.grid.add_widget(lbl)

    # ---------------------------
    # Export Excel
    # ---------------------------
    def _export_thread(self):
        if len(self.filtered_data) < 2:
            return
        folder = self.export_folder or "/storage/emulated/0/Documents/PaskiFuture"
        os.makedirs(folder, exist_ok=True)
        header = self.full_data[0]
        rows = self.filtered_data[1:]
        border = Border(left=Side(style="thin"), right=Side(style="thin"),
                        top=Side(style="thin"), bottom=Side(style="thin"))
        total = len(rows)
        done = 0
        for row in rows:
            wb = Workbook()
            ws = wb.active
            ws.append(header)
            ws.append(row)
            for col in range(1, len(header)+1):
                c = ws.cell(row=1, column=col)
                c.font = Font(bold=True)
                c.alignment = Alignment(horizontal="center")
                c.border = border
                c = ws.cell(row=2, column=col)
                c.alignment = Alignment(horizontal="center")
                c.border = border
            name = row[1] if len(row) > 1 else "brak"
            now = datetime.now().strftime("%Y%m%d_%H%M%S")
            filepath = os.path.join(folder, f"{name}_{now}.xlsx")
            wb.save(filepath)
            done += 1
            percent = int((done / total) * 100)
            Clock.schedule_once(lambda dt, p=percent: setattr(self.progress, "value", p))
        Clock.schedule_once(lambda dt: self._popup("Sukces", f"Wyeksportowano {done} plików"))

    # ---------------------------
    # Email Screen
    # ---------------------------
    layout = BoxLayout(orientation="vertical", padding=dp(30), spacing=dp(20))
        self.email_status = Label(text="Gotowy")
        self.email_progress = ProgressBar(max=100, value=0)
        send1 = PremiumButton(text="📧 Wyślij 1 rekord")
        sendAll = PremiumButton(text="📨 Wyślij hurtowo")
        back = PremiumButton(text="⬅ Powrót")

        send1.bind(on_press=lambda x: threading.Thread(target=self._send_one, daemon=True).start())
        sendAll.bind(on_press=lambda x: threading.Thread(target=self._send_all, daemon=True).start())
        back.bind(on_press=lambda x: setattr(self.sm, "current", "table"))

        layout.add_widget(send1)
        layout.add_widget(sendAll)
        layout.add_widget(self.email_status)
        layout.add_widget(self.email_progress)
        layout.add_widget(back)
        self.email.add_widget(layout)

    def _send_one(self):
        if len(self.filtered_data) < 2:
            self._popup("Błąd", "Brak danych do wysyłki")
            return
        self._send_email_row(self.filtered_data[1])

    def _send_all(self):
        if len(self.filtered_data) < 2:
            self._popup("Błąd", "Brak danych do wysyłki")
            return
        threading.Thread(target=self._send_all_thread, daemon=True).start()

    def _send_all_thread(self):
        total = len(self.filtered_data[1:])
        done = 0
        for row in self.filtered_data[1:]:
            self._send_email_row(row)
            done += 1
            percent = int((done / total) * 100)
            Clock.schedule_once(lambda dt, p=percent: setattr(self.email_progress, "value", p))
        Clock.schedule_once(lambda dt: self._popup("Sukces", f"Wysłano {done} maili"))

    def _send_email_row(self, row):
        if not os.path.exists(CONFIG_FILE):
            Clock.schedule_once(lambda dt: self._popup("Błąd", "Najpierw skonfiguruj SMTP"))
            return False
        email = self._get_email_for_row(row)
        if not email:
            Clock.schedule_once(lambda dt: setattr(self.email_status, "text", f"Nie znaleziono email dla: {row[0]} {row[1]}"))
            return False
        try:
            with open(CONFIG_FILE) as f:
                config = json.load(f)
            wb = Workbook()
            ws = wb.active
            ws.append(self.full_data[0])
            ws.append(row)
            file = Path(self.user_data_dir)/"temp.xlsx"
            wb.save(file)

            msg = EmailMessage()
            msg["Subject"] = "Dane"
            msg["From"] = config["email"]
            msg["To"] = email
            msg.set_content("W załączniku dane.")

            with open(file, "rb") as f:
                msg.add_attachment(f.read(),
                                   maintype="application",
                                   subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                   filename="dane.xlsx")

            server = smtplib.SMTP(config["server"], int(config["port"]))
            server.starttls()
            server.login(config["email"], config["password"])
            server.send_message(msg)
            server.quit()

            Clock.schedule_once(lambda dt: setattr(self.email_status, "text", f"Wysłano: {row[0]} {row[1]}"))
            return True
        except Exception as e:
            Clock.schedule_once(lambda dt: self._popup("Błąd wysyłki", str(e)))
            return False

    # ---------------------------
    # Select Email Excel
    # ---------------------------
    def select_email_excel(self, _):
        if platform != "android":
            self._popup("Błąd", "Picker działa tylko na Androidzie")
            return
        from jnius import autoclass
        from android import activity
        PythonActivity = autoclass("org.kivy.android.PythonActivity")
        Intent = autoclass("android.content.Intent")
        intent = Intent(Intent.ACTION_OPEN_DOCUMENT)
        intent.setType("*/*")
        intent.addCategory(Intent.CATEGORY_OPENABLE)
        activity.bind(on_activity_result=self._on_email_excel_result)
        PythonActivity.mActivity.startActivityForResult(intent, 555)

    def _on_email_excel_result(self, request_code, result_code, intent):
        if request_code != 555 or not intent:
            return
        from android import activity
        activity.unbind(on_activity_result=self._on_email_excel_result)
        from jnius import autoclass
        PythonActivity = autoclass("org.kivy.android.PythonActivity")
        resolver = PythonActivity.mActivity.getContentResolver()
        uri = intent.getData()
        input_stream = resolver.openInputStream(uri)
        local_file = Path(self.user_data_dir) / "email_list.xlsx"
        os.makedirs(self.user_data_dir, exist_ok=True)
        with open(local_file, "wb") as out:
            buffer = bytearray(4096)
            while True:
                read = input_stream.read(buffer)
                if read == -1:
                    break
                out.write(buffer[:read])
        input_stream.close()
        self.email_file = local_file
        self._popup("Email Excel wybrany", str(local_file))
        self._load_email_file()

    def _load_email_file(self):
        if not self.email_file:
            return
        wb = load_workbook(str(self.email_file), data_only=True)
        sheet = wb.active
        self.email_dict = {}
        header = [str(c) for c in next(sheet.iter_rows(values_only=True))]
        try:
            name_idx = header.index("Name")
            surname_idx = header.index("Surname")
            email_idx = header.index("Email")
        except ValueError:
            self._popup("Błąd", "Plik email musi zawierać kolumny: Name, Surname, Email")
            wb.close()
            return
        for row in sheet.iter_rows(min_row=2, values_only=True):
            key = f"{row[name_idx].strip().lower()} {row[surname_idx].strip().lower()}"
            self.email_dict[key] = row[email_idx]
        wb.close()

    def _get_email_for_row(self, row):
        name = str(row[0]).strip().lower()
        surname = str(row[1]).strip().lower()
        return self.email_dict.get(f"{name} {surname}")

    # ---------------------------
    # SMTP Screen
    # ---------------------------
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
        if not all(config.values()):
            self._popup("Błąd", "Wypełnij wszystkie pola SMTP")
            return
        with open(CONFIG_FILE, "w") as f:
            json.dump(config, f)
        self._popup("Zapisano", "Konfiguracja SMTP została zapisana")

    # ---------------------------
    # Popup helper
    # ---------------------------
    def _popup(self, title, message):
        content = BoxLayout(orientation="vertical", padding=dp(10))
        content.add_widget(Label(text=message))
        btn = Button(text="OK", size_hint=(1, 0.3))
        popup = Popup(title=title, content=content, size_hint=(0.8, 0.5))
        btn.bind(on_release=popup.dismiss)
        content.add_widget(btn)
        popup.open()

if __name__ == "__main__":
    PaskiFutureApp().run()
