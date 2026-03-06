# main.py – część 1/6

from pathlib import Path
from datetime import datetime
import os
import json
import threading

# ------------------ Importy Kivy ------------------
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
from kivy.uix.checkbox import CheckBox

# ------------------ Importy Excel i Mail ------------------
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Border, Side, Font, Alignment
from openpyxl.utils import get_column_letter
import smtplib
from email.message import EmailMessage

APP_TITLE = "Paski Future 6.5 PREMIUM"
CONFIG_FILE = "smtp_config.json"
EXPORT_CONFIG = "export_folder.json"
EMAILS_FILE = "emails.xlsx"

# ------------------ Ekrany ------------------
class HomeScreen(Screen): pass
class TableScreen(Screen): pass
class EmailScreen(Screen): pass
class SMTPScreen(Screen): pass

# ------------------ Przyciski ------------------
class PremiumButton(Button):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.background_normal = ""
        self.background_color = (0.2, 0.4, 0.8, 1)
        self.color = (1,1,1,1)
        self.font_size = 16
        self.size_hint_y = None
        self.height = dp(48)

# ------------------ Popup wyboru kolumn ------------------
class ColumnSelectionPopup(Popup):
    def __init__(self, headers, callback, **kwargs):
        super().__init__(**kwargs)
        self.title = "Wybierz kolumny do eksportu"
        self.size_hint = (0.9,0.8)
        self.callback = callback
        self.selected_indices = []

        layout = BoxLayout(orientation="vertical", spacing=dp(10), padding=dp(10))
        grid = GridLayout(cols=2, spacing=dp(10), size_hint_y=None)
        grid.bind(minimum_height=grid.setter('height'))
        self.checkboxes = []

        for idx, header in enumerate(headers):
            cb = CheckBox(active=True)
            self.checkboxes.append((idx, cb))
            grid.add_widget(cb)
            grid.add_widget(Label(text=str(header)))

        scroll = ScrollView(size_hint=(1,1))
        scroll.add_widget(grid)

        btn_layout = BoxLayout(size_hint=(1,None), height=dp(50), spacing=dp(10))
        btn_ok = PremiumButton(text="OK")
        btn_ok.bind(on_press=self.confirm)
        btn_layout.add_widget(btn_ok)

        layout.add_widget(scroll)
        layout.add_widget(btn_layout)
        self.add_widget(layout)

    def confirm(self, _):
        self.selected_indices = [idx for idx, cb in self.checkboxes if cb.active]
        self.dismiss()
        self.callback(self.selected_indices)

# ------------------ Główna aplikacja – część 1 ------------------
class PaskiFutureApp(App):
    def build(self):
        self.title = APP_TITLE
        Window.clearcolor = (0.08, 0.1, 0.15, 1)

        # Dane i konfiguracje
        self.full_data = []
        self.filtered_data = []
        self.current_file = None
        self.export_folder = None
        self.email_dict = {}
        self.selected_columns = None
        self.progress_value = 0

        os.makedirs(self.user_data_dir, exist_ok=True)
        self.load_export_folder()
        self.load_email_file()

        # Ekrany
        self.sm = ScreenManager()
        self.home = HomeScreen(name="home")
        self.table = TableScreen(name="table")
        self.email = EmailScreen(name="email")
        self.smtp = SMTPScreen(name="smtp")

        self._build_home()

        self.sm.add_widget(self.home)
        self.sm.add_widget(self.table)
        self.sm.add_widget(self.email)
        self.sm.add_widget(self.smtp)

        return self.sm

    # ------------------ Export folder ------------------
    def save_export_folder(self):
        if not self.export_folder:
            return
        with open(EXPORT_CONFIG,"w") as f:
            json.dump({"folder": self.export_folder}, f)

    def load_export_folder(self):
        if os.path.exists(EXPORT_CONFIG):
            with open(EXPORT_CONFIG) as f:
                self.export_folder = json.load(f).get("folder")
        else:
            self.export_folder = str(Path.home() / "Documents")
            os.makedirs(self.export_folder, exist_ok=True)

    # ------------------ Wczytanie pliku z e-mailami ------------------
    def load_email_file(self):
        path = Path(self.user_data_dir)/EMAILS_FILE
        if path.exists():
            wb = load_workbook(path,data_only=True)
            sheet = wb.active
            for row in sheet.iter_rows(min_row=2, values_only=True):
                if len(row)>=2 and row[0] and row[1]:
                    name = str(row[0]).strip()
                    email = str(row[1]).strip()
                    self.email_dict[name] = email
            wb.close()

    # ------------------ HomeScreen UI ------------------
    def _build_home(self):
        layout = BoxLayout(orientation="vertical",padding=dp(30),spacing=dp(20))
        title = Label(text=APP_TITLE,font_size=26,bold=True)
        self.home_status = Label(text="Gotowy",font_size=16)
        open_btn = PremiumButton(text="📂 Otwórz plik Excel")
        open_btn.bind(on_press=self.open_excel_picker)
        load_btn = PremiumButton(text="📊 Wczytaj dane")
        load_btn.bind(on_press=self.load_full_excel)
        smtp_btn = PremiumButton(text="⚙ Konfiguracja SMTP")
        smtp_btn.bind(on_press=lambda x:setattr(self.sm,"current","smtp"))
        layout.add_widget(title)
        layout.add_widget(open_btn)
        layout.add_widget(load_btn)
        layout.add_widget(smtp_btn)
        layout.add_widget(self.home_status)
        self.home.add_widget(layout)
        # main.py – część 2/6

# ------------------ TableScreen UI ------------------
def _build_table_ui(self):
    layout = BoxLayout(orientation="vertical", padding=dp(10), spacing=dp(10))

    # Górny panel: wyszukiwanie + przyciski
    top = BoxLayout(size_hint=(1, 0.12), spacing=dp(10))

    self.search = TextInput(hint_text="🔎 Wyszukaj...", multiline=False)
    self.search.bind(text=self.filter_data)

    export_btn = PremiumButton(text="📦 Eksport")
    export_btn.bind(on_press=lambda x: threading.Thread(target=self.export_data_thread, daemon=True).start())

    back_btn = PremiumButton(text="⬅ Powrót")
    back_btn.bind(on_press=lambda x: setattr(self.sm, "current", "home"))

    top.add_widget(self.search)
    top.add_widget(export_btn)
    top.add_widget(back_btn)

    # Scroll z tabelą
    self.scroll = ScrollView()

    self.grid = GridLayout(size_hint=(None, None), spacing=dp(1))
    self.grid.bind(minimum_height=self.grid.setter('height'))
    self.grid.bind(minimum_width=self.grid.setter('width'))

    self.scroll.add_widget(self.grid)

    # Progress bar
    self.progress = ProgressBar(max=100)

    layout.add_widget(top)
    layout.add_widget(self.scroll)
    layout.add_widget(self.progress)

    self.table.clear_widgets()
    self.table.add_widget(layout)

    # Wyświetlenie tabeli
    self.display_table()

# ------------------ Filtr danych ------------------
def filter_data(self, instance, value):
    value = value.lower()
    self.filtered_data = [self.full_data[0]] + [
        row for row in self.full_data[1:]
        if any(value in str(cell).lower() for cell in row)
    ]
    self.display_table()

# ------------------ Wyświetlenie tabeli ------------------
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
                size=(dp(160), dp(40)),
            )
            self.grid.add_widget(lbl)
         # main.py – część 3/6

# ------------------ Funkcja eksportu w osobnym wątku ------------------
def export_data_thread(self):
    if not self.filtered_data or len(self.filtered_data) < 2:
        self._popup("Błąd", "Brak danych do eksportu")
        return

    # Jeśli nie wybrano kolumn, wybierz wszystkie
    if self.selected_columns is None:
        self.selected_columns = list(range(len(self.filtered_data[0])))

    # Tworzenie nowego workbooka
    wb = Workbook()
    ws = wb.active

    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )

    # Dodanie nagłówków i ramki
    headers = [self.filtered_data[0][i] for i in self.selected_columns]
    ws.append(headers)
    for col_idx, header in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col_idx)
        cell.font = Font(bold=True)
        cell.border = thin_border
        cell.alignment = Alignment(horizontal='center', vertical='center')
        ws.column_dimensions[get_column_letter(col_idx)].width = 20

    # Dodanie wierszy
    for row in self.filtered_data[1:]:
        row_values = [row[i] for i in self.selected_columns]
        ws.append(row_values)
        row_num = ws.max_row
        for col_idx in range(1, len(row_values)+1):
            ws.cell(row=row_num, column=col_idx).border = thin_border

    # Zapis do folderu Documents / export_folder
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    filename = f"export_{timestamp}.xlsx"
    filepath = Path(self.export_folder)/filename
    wb.save(filepath)

    Clock.schedule_once(lambda dt: self._popup("Eksport", f"Zapisano plik:\n{filepath}"))
    # main.py – część 4/6

# ------------------ EmailScreen UI ------------------
def _build_email(self):
    layout = BoxLayout(orientation="vertical", padding=dp(30), spacing=dp(15))

    self.email_status = Label(text="Gotowy", font_size=16)

    send_btn = PremiumButton(text="📧 Wyślij pliki")
    send_btn.bind(on_press=lambda x: threading.Thread(target=self.send_emails_thread, daemon=True).start())

    back_btn = PremiumButton(text="⬅ Powrót")
    back_btn.bind(on_press=lambda x: setattr(self.sm, "current", "home"))

    layout.add_widget(send_btn)
    layout.add_widget(back_btn)
    layout.add_widget(self.email_status)

    self.email.clear_widgets()
    self.email.add_widget(layout)

# ------------------ Wysyłka maili w osobnym wątku ------------------
def send_emails_thread(self):
    if not self.filtered_data or len(self.filtered_data) < 2:
        self._popup("Błąd", "Brak danych do wysłania")
        return

    if not self.selected_columns:
        self.selected_columns = list(range(len(self.filtered_data[0])))

    # Wczytaj konfigurację SMTP
    if not os.path.exists(CONFIG_FILE):
        self._popup("Błąd", "Brak konfiguracji SMTP")
        return

    with open(CONFIG_FILE) as f:
        smtp_config = json.load(f)

    try:
        server = smtplib.SMTP(smtp_config["server"], int(smtp_config["port"]))
        server.starttls()
        server.login(smtp_config["email"], smtp_config["password"])
        server.quit()
        Clock.schedule_once(lambda dt: self._popup("SMTP", "Połączenie SMTP OK"))
    except Exception as e:
        Clock.schedule_once(lambda dt: self._popup("Błąd SMTP", str(e)))
        return

    # Wysyłka indywidualna
    for row in self.filtered_data[1:]:
        # Spróbuj dopasować adres e-mail po imieniu i nazwisku
        key = f"{row[0].strip()} {row[1].strip()}" if len(row)>=2 else str(row[0]).strip()
        recipient = self.email_dict.get(key)
        if not recipient:
            continue  # pomiń jeśli brak adresu

        # Tworzenie tymczasowego pliku dla jednej osoby
        wb = Workbook()
        ws = wb.active
        thin_border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )

        headers = [self.filtered_data[0][i] for i in self.selected_columns]
        ws.append(headers)
        for col_idx, header in enumerate(headers, start=1):
            cell = ws.cell(row=1, column=col_idx)
            cell.font = Font(bold=True)
            cell.border = thin_border
            cell.alignment = Alignment(horizontal='center', vertical='center')
            ws.column_dimensions[get_column_letter(col_idx)].width = 20

        row_values = [row[i] for i in self.selected_columns]
        ws.append(row_values)
        row_num = ws.max_row
        for col_idx in range(1, len(row_values)+1):
            ws.cell(row=row_num, column=col_idx).border = thin_border

        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        filename = f"{key.replace(' ','_')}_{timestamp}.xlsx"
        filepath = Path(self.export_folder)/filename
        wb.save(filepath)

        # Wysyłka maila z załącznikiem
        try:
            msg = EmailMessage()
            msg['Subject'] = "Twój plik Paski Future"
            msg['From'] = smtp_config["email"]
            msg['To'] = recipient
            msg.set_content("W załączeniu Twój plik.")

            with open(filepath, 'rb') as f:
                file_data = f.read()
                file_name = f.name

            msg.add_attachment(file_data, maintype='application', subtype='vnd.openxmlformats-officedocument.spreadsheetml.sheet', filename=file_name)

            with smtplib.SMTP(smtp_config["server"], int(smtp_config["port"])) as s:
                s.starttls()
                s.login(smtp_config["email"], smtp_config["password"])
                s.send_message(msg)

        except Exception as e:
            Clock.schedule_once(lambda dt: self._popup("Błąd wysyłki", f"{recipient}: {e}"))

    Clock.schedule_once(lambda dt: self._popup("Wysyłka", "Wysyłanie zakończone"))
    # main.py – część 5/6

# ------------------ SMTPScreen UI ------------------
def _build_smtp(self):
    layout = BoxLayout(orientation="vertical", padding=dp(30), spacing=dp(15))

    self.s_server = TextInput(hint_text="SMTP server")
    self.s_port = TextInput(hint_text="Port")
    self.s_email = TextInput(hint_text="Email")
    self.s_pass = TextInput(hint_text="Hasło", password=True)

    save_btn = PremiumButton(text="💾 Zapisz")
    save_btn.bind(on_press=self.save_smtp)

    back_btn = PremiumButton(text="⬅ Powrót")
    back_btn.bind(on_press=lambda x: setattr(self.sm,"current","home"))

    layout.add_widget(self.s_server)
    layout.add_widget(self.s_port)
    layout.add_widget(self.s_email)
    layout.add_widget(self.s_pass)
    layout.add_widget(save_btn)
    layout.add_widget(back_btn)

    self.smtp.clear_widgets()
    self.smtp.add_widget(layout)

# ------------------ Zapis konfiguracji SMTP ------------------
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
    # main.py – część 6/6

# ------------------ Popupy ------------------
def _popup(self, title, message):
    content = BoxLayout(orientation="vertical", padding=dp(10))
    content.add_widget(Label(text=message))
    btn = Button(text="OK", size_hint=(1,0.3))
    popup = Popup(
        title=title,
        content=content,
        size_hint=(0.8,0.5)
    )
    btn.bind(on_release=popup.dismiss)
    content.add_widget(btn)
    popup.open()

# ------------------ Funkcje pomocnicze do Excel ------------------
# (opcjonalnie można dodać tu dodatkowe funkcje np. automatyczne dopasowanie kolumn itp.)

# ------------------ Uruchomienie aplikacji ------------------
if __name__ == "__main__":
    PaskiFutureApp().run()
