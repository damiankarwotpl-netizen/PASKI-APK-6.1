# === Kod części 1/2 ===
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
from kivy.uix.checkbox import CheckBox
from kivy.uix.label import Label
from kivy.uix.scrollview import ScrollView
from kivy.uix.gridlayout import GridLayout
from kivy.uix.textinput import TextInput
from kivy.uix.popup import Popup
from kivy.uix.progressbar import ProgressBar
from kivy.uix.screenmanager import ScreenManager, Screen

from openpyxl import load_workbook, Workbook
from openpyxl.utils import get_column_letter

APP_TITLE = "Future 7.0 PRO"
CONFIG_FILE = "smtp_config.json"

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

        self.full_data = []        # wszystkie dane z pliku
        self.filtered_data = []    # dane po filtrze
        self.current_file = None
        self.email_column = None   # indeks kolumny z e-mailami
        self.export_columns = None # lista indeksów kolumn do eksportu

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
        intent.addFlags(Intent.FLAG_GRANT_READ_URI_PERMISSION)
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
        if not self.full_data:
            self._popup("Błąd", "Plik pusty")
            return

        # Ustawiamy domyślnie: kolumna e-mail = 0, eksport wszystkich kolumn
        self.email_column = 0
        self.export_columns = list(range(len(self.full_data[0])))

        self.filtered_data = self.full_data
        self.display_table()
        self.sm.current = "table"
        # Po wczytaniu danych pokazujemy popupy wyboru kolumn
        Clock.schedule_once(lambda dt: self._popup_email_column(), 0)

    def _popup_email_column(self, *args):
        header = self.full_data[0]
        popup = Popup(title="Wybierz kolumnę e-mail", size_hint=(0.9, 0.9))
        layout = BoxLayout(orientation="vertical", spacing=dp(10), padding=dp(10))
        scroll = ScrollView()
        grid = GridLayout(cols=1, spacing=dp(5), size_hint_y=None)
        grid.bind(minimum_height=grid.setter('height'))
        self.email_checkbox_list = []
        for i, col in enumerate(header):
            row = BoxLayout(size_hint_y=None, height=dp(40))
            cb = CheckBox(group='email', active=(i == self.email_column))
            lbl = Label(text=col, valign='middle')
            row.add_widget(cb); row.add_widget(lbl)
            grid.add_widget(row)
            self.email_checkbox_list.append((cb, i))
        scroll.add_widget(grid)
        layout.add_widget(scroll)
        btn_layout = BoxLayout(size_hint_y=None, height=dp(40), spacing=dp(10))
        ok = PremiumButton(text="OK"); cancel = PremiumButton(text="Anuluj")
        btn_layout.add_widget(ok); btn_layout.add_widget(cancel)
        layout.add_widget(btn_layout)
        popup.content = layout

        def on_ok(instance):
            for cb, idx in self.email_checkbox_list:
                if cb.active:
                    self.email_column = idx
                    break
            popup.dismiss()
            # Po wyborze kolumny email otwieramy popup wyboru kolumn do eksportu
            Clock.schedule_once(lambda dt: self._popup_export_columns(), 0)
        ok.bind(on_press=on_ok)
        cancel.bind(on_press=lambda x: popup.dismiss())
        popup.open()

    def _popup_export_columns(self, *args):
        header = self.full_data[0]
        popup = Popup(title="Wybierz kolumny do eksportu", size_hint=(0.9, 0.9))
        layout = BoxLayout(orientation="vertical", spacing=dp(10), padding=dp(10))
        scroll = ScrollView()
        grid = GridLayout(cols=1, spacing=dp(5), size_hint_y=None)
        grid.bind(minimum_height=grid.setter('height'))
        self.export_checkbox_list = []
        for i, col in enumerate(header):
            row = BoxLayout(size_hint_y=None, height=dp(40))
            cb = CheckBox(active=(i in self.export_columns))
            lbl = Label(text=col, valign='middle')
            row.add_widget(cb); row.add_widget(lbl)
            grid.add_widget(row)
            self.export_checkbox_list.append((cb, i))
        scroll.add_widget(grid)
        layout.add_widget(scroll)
        btn_layout = BoxLayout(size_hint_y=None, height=dp(40), spacing=dp(10))
        ok = PremiumButton(text="OK"); cancel = PremiumButton(text="Anuluj")
        btn_layout.add_widget(ok); btn_layout.add_widget(cancel)
        layout.add_widget(btn_layout)
        popup.content = layout

        def on_ok(instance):
            self.export_columns = [idx for cb, idx in self.export_checkbox_list if cb.active]
            if not self.export_columns:
                self._popup("Błąd", "Wybierz przynajmniej jedną kolumnę")
                return
            popup.dismiss()
        ok.bind(on_press=on_ok)
        cancel.bind(on_press=lambda x: popup.dismiss())
        popup.open()

    def display_table(self):
        # Tworzy widok tabeli (nagłówki + dane) na ekranie TableScreen
        self.table.clear_widgets()
        layout = BoxLayout(orientation="vertical", padding=dp(10), spacing=dp(10))

        top = BoxLayout(size_hint=(1, 0.12), spacing=dp(10))
        self.search = TextInput(
            hint_text="🔎 Wyszukaj...", multiline=False,
            background_color=(0.15, 0.18, 0.25, 1),
            foreground_color=(1, 1, 1, 1)
        )
        self.search.bind(text=self.filter_data)
        export_btn = PremiumButton(text="📦 Eksport"); export_btn.bind(on_press=self.export_files)
        email_btn = PremiumButton(text="📬 Email");  email_btn.bind(on_press=lambda x: setattr(self.sm, "current", "email"))
        back_btn = PremiumButton(text="⬅ Powrót"); back_btn.bind(on_press=lambda x: setattr(self.sm, "current", "home"))
        top.add_widget(self.search); top.add_widget(export_btn)
        top.add_widget(email_btn); top.add_widget(back_btn)
        layout.add_widget(top)

        # Nagłówki kolumn
        header = self.filtered_data[0]
        header_layout = GridLayout(cols=len(header), size_hint_y=None, height=dp(30))
        for col in header:
            header_layout.add_widget(Label(text=str(col), color=(1,1,1,1), bold=True))
        layout.add_widget(header_layout)
        # === Kod części 2/2 ===
        # Dane wierszy (dostęp do przewijania, jeśli dużo kolumn/wierszy)
        scroll = ScrollView(do_scroll_x=True, do_scroll_y=True)
        grid = GridLayout(cols=len(header), spacing=dp(5), size_hint=(None, None))
        grid.bind(minimum_height=grid.setter('height'), minimum_width=grid.setter('width'))
        for row in self.filtered_data[1:]:
            for cell in row:
                grid.add_widget(Label(text=str(cell)))
        scroll.add_widget(grid)
        layout.add_widget(scroll)

        # Pasek postępu przy eksporcie
        self.progress = ProgressBar(max=100, size_hint=(1, 0.05))
        layout.add_widget(self.progress)
        self.table.add_widget(layout)

    def filter_data(self, instance, value):
        value = value.lower()
        self.filtered_data = [
            row for row in self.full_data
            if any(value in str(cell).lower() for cell in row)
        ]
        if self.filtered_data:
            self.display_table()

    def export_files(self, _):
        threading.Thread(target=self._export_thread).start()

    def _export_thread(self):
        if not self.filtered_data or len(self.filtered_data) < 2:
            return
        documents = "/storage/emulated/0/Documents/PaskiFuture"
        os.makedirs(documents, exist_ok=True)
        header = self.full_data[0]
        rows = self.filtered_data[1:]
        total = len(rows); done = 0
        for row in rows:
            wb = Workbook(); ws = wb.active
            # Tylko wybrane kolumny
            selected_header = [header[i] for i in self.export_columns]
            selected_row = [row[i] for i in self.export_columns]
            ws.append(selected_header); ws.append(selected_row)
            # Dopasowanie szerokości kolumn
            for col_pos, orig_idx in enumerate(self.export_columns, start=1):
                col_letter = get_column_letter(col_pos)
                max_len = max(len(str(selected_header[col_pos-1])), len(str(selected_row[col_pos-1])))
                ws.column_dimensions[col_letter].width = max_len + 2
            name = row[1] if len(row) > 1 else "brak"
            now = datetime.now().strftime("%Y%m%d_%H%M%S")
            filepath = os.path.join(documents, f"{name}_{now}.xlsx")
            wb.save(filepath)
            done += 1
            percent = int((done / total) * 100)
            Clock.schedule_once(lambda dt, p=percent: setattr(self.progress, "value", p))
        Clock.schedule_once(lambda dt: self._popup("Sukces", f"Wyeksportowano {done} plików"))

    def _build_email(self):
        layout = BoxLayout(orientation="vertical", padding=dp(30), spacing=dp(20))
        send1 = PremiumButton(text="📧 Wyślij 1 rekord")
        sendAll = PremiumButton(text="📨 Wyślij hurtowo")
        back = PremiumButton(text="⬅ Powrót")
        send1.bind(on_press=self.send_single)
        sendAll.bind(on_press=self.send_bulk)
        back.bind(on_press=lambda x: setattr(self.sm, "current", "table"))
        self.email_status = Label(text="Gotowy")
        layout.add_widget(send1); layout.add_widget(sendAll)
        layout.add_widget(self.email_status); layout.add_widget(back)
        self.email.add_widget(layout)

    def send_single(self, _):
        threading.Thread(target=self._send_one).start()
    def send_bulk(self, _):
        threading.Thread(target=self._send_all).start()
    def _send_one(self):
        if not self.filtered_data or len(self.filtered_data) < 2: return
        self._send_email_row(self.filtered_data[1])
    def _send_all(self):
        for row in self.filtered_data[1:]:
            self._send_email_row(row)

    def _send_email_row(self, row):
        if not os.path.exists(CONFIG_FILE): return
        with open(CONFIG_FILE) as f:
            config = json.load(f)
        wb = Workbook(); ws = wb.active
        header = self.full_data[0]
        selected_header = [header[i] for i in self.export_columns]
        selected_row = [row[i] for i in self.export_columns]
        ws.append(selected_header); ws.append(selected_row)
        file = Path(self.user_data_dir) / "temp.xlsx"
        wb.save(file)

        msg = EmailMessage()
        msg["Subject"] = "Dane"; msg["From"] = config["email"]
        msg["To"] = row[self.email_column]
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

    def _build_smtp(self):
        layout = BoxLayout(orientation="vertical", padding=dp(30), spacing=dp(15))
        self.s_server = TextInput(hint_text="SMTP Server")
        self.s_port = TextInput(hint_text="Port")
        self.s_email = TextInput(hint_text="Email")
        self.s_pass = TextInput(hint_text="Hasło", password=True)
        save = PremiumButton(text="💾 Zapisz")
        back = PremiumButton(text="⬅ Powrót")
        save.bind(on_press=self.save_smtp)
        back.bind(on_press=lambda x: setattr(self.sm, "current", "home"))
        layout.add_widget(self.s_server); layout.add_widget(self.s_port)
        layout.add_widget(self.s_email); layout.add_widget(self.s_pass)
        layout.add_widget(save);     layout.add_widget(back)
        self.smtp.add_widget(layout)

    def save_smtp(self, _):
        data = {"server": self.s_server.text, "port": self.s_port.text,
                "email": self.s_email.text, "password": self.s_pass.text}
        with open(CONFIG_FILE, "w") as f: json.dump(data, f)
        self._popup("Sukces", "SMTP zapisane")

    def _popup(self, title, text):
        Popup(title=title, content=Label(text=text), size_hint=(0.8, 0.4)).open()

if __name__ == "__main__":
    PaskiFutureApp().run()
