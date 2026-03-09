import os
import json
import sqlite3
import threading
import smtplib
from datetime import datetime
from pathlib import Path
from email.message import EmailMessage

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

from openpyxl import load_workbook, Workbook


APP_TITLE = "Future 9.0 ULTRA PRO"


class HomeScreen(Screen):
    pass


class TableScreen(Screen):
    pass


class EmailScreen(Screen):
    pass


class SMTPScreen(Screen):
    pass


class TemplateScreen(Screen):
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
        self.global_attachments = []

        # Baza Danych
        self.init_db()

        self.sm = ScreenManager()

        self.home = HomeScreen(name="home")
        self.table = TableScreen(name="table")
        self.email = EmailScreen(name="email")
        self.smtp = SMTPScreen(name="smtp")
        self.tmpl = TemplateScreen(name="tmpl")

        self.build_home()
        self.build_table()
        self.build_email()
        self.build_smtp()
        self.build_tmpl()

        self.sm.add_widget(self.home)
        self.sm.add_widget(self.table)
        self.sm.add_widget(self.email)
        self.sm.add_widget(self.smtp)
        self.sm.add_widget(self.tmpl)

        return self.sm

    def init_db(self):
        db_path = Path(self.user_data_dir) / "app_v9.db"
        self.conn = sqlite3.connect(str(db_path), check_same_thread=False)
        self.conn.execute("CREATE TABLE IF NOT EXISTS contacts (name TEXT, surname TEXT, email TEXT, PRIMARY KEY(name, surname))")
        self.conn.execute("CREATE TABLE IF NOT EXISTS settings (key TEXT PRIMARY KEY, val TEXT)")
        self.conn.execute("CREATE TABLE IF NOT EXISTS logs (id INTEGER PRIMARY KEY, msg TEXT, date TEXT)")
        
        # Domyślne wartości szablonu
        if not self.conn.execute("SELECT val FROM settings WHERE key='t_body'").fetchone():
            self.conn.execute("INSERT OR REPLACE INTO settings VALUES ('t_sub', 'Informacja dla {Imię}')")
            self.conn.execute("INSERT OR REPLACE INTO settings VALUES ('t_body', 'Witaj {Imię},\n\nW załączniku przesyłamy dane.')")
        self.conn.commit()

# -----------------------------
# HOME
# -----------------------------

    def build_home(self):

        layout = BoxLayout(orientation="vertical", padding=dp(25), spacing=dp(15))

        title = Label(text=APP_TITLE, font_size=26, bold=True)

        open_btn = PremiumButton(text="📂 Otwórz Excel Płac")
        open_btn.bind(on_press=self.open_picker_data)

        load_btn = PremiumButton(text="📊 Podgląd Tabeli")
        load_btn.bind(on_press=self.load_excel)

        email_btn = PremiumButton(text="✉ Centrum Mailingu")
        email_btn.bind(on_press=lambda x: setattr(self.sm, "current", "email"))

        settings_btn = PremiumButton(text="⚙ Ustawienia SMTP")
        settings_btn.bind(on_press=lambda x: setattr(self.sm, "current", "smtp"))

        self.status = Label(text="Gotowy", color=(0.7, 0.7, 0.7, 1))

        layout.add_widget(title)
        layout.add_widget(open_btn)
        layout.add_widget(load_btn)
        layout.add_widget(email_btn)
        layout.add_widget(settings_btn)
        layout.add_widget(self.status)

        self.home.add_widget(layout)


# -----------------------------
# ANDROID PICKER LOGIC
# -----------------------------

    def open_picker_data(self, _):
        self.open_picker(mode="data")

    def open_picker(self, mode="data"):

        if platform != "android":
            self.status.text = "Picker tylko Android"
            return

        from jnius import autoclass
        from android import activity

        PythonActivity = autoclass("org.kivy.android.PythonActivity")
        Intent = autoclass("android.content.Intent")

        intent = Intent(Intent.ACTION_GET_CONTENT)
        intent.setType("*/*")
        intent.addCategory(Intent.CATEGORY_OPENABLE)

        def callback(request_code, result_code, intent_data):
            if request_code != 1001 or not intent_data:
                return
            
            activity.unbind(on_activity_result=callback)
            
            uri = intent_data.getData()
            resolver = PythonActivity.mActivity.getContentResolver()
            stream = resolver.openInputStream(uri)
            
            # Unikalna nazwa dla plików tymczasowych
            suffix = datetime.now().strftime("%H%M%S")
            filename = f"file_{mode}_{suffix}.xlsx"
            local = Path(self.user_data_dir) / filename

            with open(local, "wb") as f:
                buffer = bytearray(4096)
                while True:
                    r = stream.read(buffer)
                    if r == -1: break
                    f.write(buffer[:r])
            stream.close()

            if mode == "data":
                self.current_file = local
                self.status.text = "Plik Excel załadowany"
            elif mode == "book":
                self.import_contacts_to_db(local)
            elif mode == "attachment":
                if str(local) not in self.global_attachments:
                    self.global_attachments.append(str(local))
                self.update_email_ui_labels()

        activity.bind(on_activity_result=callback)
        PythonActivity.mActivity.startActivityForResult(intent, 1001)

# -----------------------------
# LOAD EXCEL
# -----------------------------

    def load_excel(self, _):

        if not self.current_file:
            self.popup("Błąd", "Najpierw wybierz plik")
            return

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

        back_btn = PremiumButton(text="Powrót")
        back_btn.bind(on_press=lambda x: setattr(self.sm, "current", "home"))

        top.add_widget(self.search)
        top.add_widget(export_btn)
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
        
        mass_export_btn = Button(text="EKSPORTUJ WSZYSTKIE OSOBNO", size_hint_y=None, height=dp(45))
        mass_export_btn.bind(on_press=self.mass_export_start)
        layout.add_widget(mass_export_btn)

        self.table.add_widget(layout)


    def show_table(self):

        self.grid.clear_widgets()
        if not self.filtered_data: return

        rows = len(self.filtered_data)
        cols = len(self.filtered_data[0])

        w = dp(160)
        h = dp(42)

        # +1 kolumna na przycisk akcji
        self.grid.cols = cols + 1
        self.grid.width = (cols + 1) * w
        self.grid.height = rows * h

        # Header
        for cell in self.filtered_data[0]:
            self.grid.add_widget(Label(text=str(cell), size_hint=(None, None), size=(w, h), bold=True))
        self.grid.add_widget(Label(text="Akcja", size_hint=(None, None), size=(w, h), bold=True))

        # Rows
        for row in self.filtered_data[1:]:
            for cell in row:
                self.grid.add_widget(Label(text=str(cell), size_hint=(None, None), size=(w, h)))
            
            btn = Button(text="Zapisz ten", size_hint=(None, None), size=(w, h))
            btn.bind(on_press=lambda x, r=row: self.single_export(r))
            self.grid.add_widget(btn)

# -----------------------------
# LOGIKA FILTRA I EKSPORTU
# -----------------------------

    def filter_data(self, instance, value):
        v = value.lower()
        self.filtered_data = [self.full_data[0]] + [
            row for row in self.full_data[1:]
            if any(v in str(cell).lower() for cell in row)
        ]
        self.show_table()

    def mass_export_start(self, _):
        if not self.full_data: return
        threading.Thread(target=self._mass_export_thread).start()

    def _mass_export_thread(self):
        folder = Path("/storage/emulated/0/Documents/FutureExport")
        folder.mkdir(parents=True, exist_ok=True)
        rows = self.filtered_data[1:] if len(self.filtered_data) > 1 else self.full_data[1:]
        
        for i, row in enumerate(rows):
            self.single_export(row, silent=True)
            percent = int((i + 1) / len(rows) * 100)
            Clock.schedule_once(lambda dt, p=percent: setattr(self.progress, "value", p))
        
        Clock.schedule_once(lambda dt: self.popup("Sukces", "Wszystkie pliki zapisane w Documents/FutureExport"))

    def single_export(self, row, silent=False):
        folder = Path("/storage/emulated/0/Documents/FutureExport")
        folder.mkdir(parents=True, exist_ok=True)
        
        wb = Workbook()
        ws = wb.active
        ws.append(self.full_data[0])
        ws.append(row)
        
        name = f"{row[0]}_{row[1]}" if len(row) > 1 else "raport"
        wb.save(folder / f"Raport_{name}.xlsx")
        if not silent: self.popup("Zapisano", f"Plik dla {name} gotowy.")


# -----------------------------
# EMAIL SCREEN
# -----------------------------

    def build_email(self):

        layout = BoxLayout(orientation="vertical", padding=dp(20), spacing=dp(10))

        layout.add_widget(Label(text="Centrum Wysyłki", font_size=22, bold=True))

        self.email_info = Label(text="Baza GMAIL: niezaładowana", size_hint_y=None, height=dp(30))
        self.att_info = Label(text="Załączniki: 0", size_hint_y=None, height=dp(30))
        
        layout.add_widget(self.email_info)
        layout.add_widget(self.att_info)

        btns = [
            ("📁 Wczytaj Bazę Kontaktów", lambda x: self.open_picker(mode="book")),
            ("📝 Edytuj Treść Maila", lambda x: setattr(self.sm, "current", "tmpl")),
            ("📎 Dodaj Załącznik PDF", lambda x: self.open_picker(mode="attachment")),
            ("⚡ Test Mail (do siebie)", self.send_test_email),
            ("🚀 Uruchom Mailing Masowy", self.send_emails),
            ("📜 Historia", self.show_history),
            ("Powrót", lambda x: setattr(self.sm, "current", "home"))
        ]

        for txt, cmd in btns:
            btn = PremiumButton(text=txt)
            btn.bind(on_press=cmd)
            layout.add_widget(btn)

        self.email_status = Label(text="")
        layout.add_widget(self.email_status)

        self.email.add_widget(layout)

    def update_email_ui_labels(self):
        count = self.conn.execute("SELECT count(*) FROM contacts").fetchone()[0]
        self.email_info.text = f"Baza GMAIL: {count} osób"
        self.att_info.text = f"Załączniki: {len(self.global_attachments)}"

# -----------------------------
# TEMPLATE SCREEN
# -----------------------------

    def build_tmpl(self):
        layout = BoxLayout(orientation="vertical", padding=dp(20), spacing=dp(10))
        
        layout.add_widget(Label(text="Szablon Wiadomości", bold=True))
        
        self.tmpl_sub = TextInput(hint_text="Temat maila", size_hint_y=None, height=dp(45))
        self.tmpl_body = TextInput(hint_text="Treść (użyj {Imię})", multiline=True)
        
        # Wczytaj z DB
        res = self.conn.execute("SELECT val FROM settings WHERE key='t_sub'").fetchone()
        if res: self.tmpl_sub.text = res[0]
        res = self.conn.execute("SELECT val FROM settings WHERE key='t_body'").fetchone()
        if res: self.tmpl_body.text = res[0]

        save_btn = PremiumButton(text="Zapisz Szablon")
        save_btn.bind(on_press=self.save_template)
        
        back_btn = PremiumButton(text="Wróć")
        back_btn.bind(on_press=lambda x: setattr(self.sm, "current", "email"))

        layout.add_widget(self.tmpl_sub)
        layout.add_widget(self.tmpl_body)
        layout.add_widget(save_btn)
        layout.add_widget(back_btn)
        
        self.tmpl.add_widget(layout)

    def save_template(self, _):
        self.conn.execute("UPDATE settings SET val=? WHERE key='t_sub'", (self.tmpl_sub.text,))
        self.conn.execute("UPDATE settings SET val=? WHERE key='t_body'", (self.tmpl_body.text,))
        self.conn.commit()
        self.popup("OK", "Szablon zapisany")

# -----------------------------
# SMTP SETTINGS
# -----------------------------

    def build_smtp(self):

        layout = BoxLayout(orientation="vertical", padding=dp(20), spacing=dp(10))

        layout.add_widget(Label(text="Ustawienia SMTP (Gmail)", font_size=20))

        self.smtp_user = TextInput(hint_text="Email", multiline=False)
        self.smtp_pass = TextInput(hint_text="Hasło Aplikacji (16 znaków)", multiline=False, password=True)

        # Load existing
        path = Path(self.user_data_dir) / "smtp.json"
        if path.exists():
            data = json.load(open(path))
            self.smtp_user.text = data.get("user", "")
            self.smtp_pass.text = data.get("pass", "")

        save = PremiumButton(text="Zapisz Konfigurację")
        save.bind(on_press=self.save_smtp)

        back = PremiumButton(text="Powrót")
        back.bind(on_press=lambda x: setattr(self.sm, "current", "home"))

        layout.add_widget(self.smtp_user)
        layout.add_widget(self.smtp_pass)
        layout.add_widget(save)
        layout.add_widget(back)

        self.smtp.add_widget(layout)

    def save_smtp(self, _):
        data = {"user": self.smtp_user.text, "pass": self.smtp_pass.text}
        with open(Path(self.user_data_dir) / "smtp.json", "w") as f:
            json.dump(data, f)
        self.popup("OK", "SMTP zapisane")

# -----------------------------
# MAILING LOGIC
# -----------------------------

    def import_contacts_to_db(self, path):
        try:
            wb = load_workbook(path, data_only=True)
            ws = wb.active
            rows = list(ws.iter_rows(values_only=True))
            count = 0
            for r in rows[1:]:
                if r[0] and r[2]: # Imię i Email
                    self.conn.execute("INSERT OR REPLACE INTO contacts VALUES (?,?,?)", (str(r[0]).lower().strip(), str(r[1]).lower().strip(), str(r[2]).strip()))
                    count += 1
            self.conn.commit()
            Clock.schedule_once(lambda dt: self.popup("Import", f"Dodano {count} kontaktów"))
            self.update_email_ui_labels()
        except Exception as e:
            Clock.schedule_once(lambda dt: self.popup("Błąd", str(e)))

    def send_test_email(self, _):
        if not self.full_data: 
            self.popup("!", "Najpierw wczytaj Excel!")
            return
        threading.Thread(target=self._mailing_thread, args=(True,)).start()

    def send_emails(self, _):
        if not self.full_data:
            self.popup("!", "Najpierw wczytaj Excel!")
            return
        threading.Thread(target=self._mailing_thread, args=(False,)).start()

    def _mailing_thread(self, is_test):
        # Load SMTP
        smtp_path = Path(self.user_data_dir) / "smtp.json"
        if not smtp_path.exists():
            Clock.schedule_once(lambda dt: self.popup("Błąd", "Skonfiguruj SMTP!"))
            return
        smtp_cfg = json.load(open(smtp_path))
        
        try:
            server = smtplib.SMTP("smtp.gmail.com", 587, timeout=15)
            server.starttls()
            server.login(smtp_cfg["user"], smtp_cfg["pass"])
        except Exception as e:
            Clock.schedule_once(lambda dt: self.popup("SMTP Error", str(e)))
            return

        header = self.full_data[0]
        data_rows = self.full_data[1:2] if is_test else self.full_data[1:]
        sent_count = 0

        for i, row in enumerate(data_rows):
            name, surname = str(row[0]).lower().strip(), str(row[1]).lower().strip()
            
            target_email = smtp_cfg["user"] if is_test else ""
            if not is_test:
                res = self.conn.execute("SELECT email FROM contacts WHERE name=? AND surname=?", (name, surname)).fetchone()
                if res: target_email = res[0]
            
            if target_email:
                try:
                    msg = EmailMessage()
                    dat_str = datetime.now().strftime("%d.%m.%Y")
                    
                    msg["Subject"] = self.tmpl_sub.text.replace("{Imię}", str(row[0]))
                    msg["From"] = smtp_cfg["user"]
                    msg["To"] = target_email
                    msg.set_content(self.tmpl_body.text.replace("{Imię}", str(row[0])).replace("{Data}", dat_str))

                    # Załącznik automatyczny (Excel tej osoby)
                    tmp_p = Path(self.user_data_dir) / "tmp_att.xlsx"
                    wb = Workbook(); ws = wb.active
                    ws.append(header); ws.append(row)
                    wb.save(tmp_p)
                    with open(tmp_p, "rb") as f:
                        msg.add_attachment(f.read(), maintype="application", subtype="xlsx", filename=f"Raport_{row[0]}.xlsx")
                    
                    # Załączniki dodatkowe
                    for ap in self.global_attachments:
                        if os.path.exists(ap):
                            with open(ap, "rb") as f:
                                msg.add_attachment(f.read(), maintype="application", subtype="octet-stream", filename=os.path.basename(ap))
                    
                    server.send_message(msg)
                    sent_count += 1
                    self.conn.execute("INSERT INTO logs (msg, date) VALUES (?,?)", (f"Wysłano do: {target_email}", dat_str))
                except:
                    pass

            # Progress
            perc = int((i + 1) / len(data_rows) * 100)
            Clock.schedule_once(lambda dt, p=perc: setattr(self.progress, "value", p))

        server.quit()
        self.conn.commit()
        Clock.schedule_once(lambda dt: self.popup("Mailing", f"Ukończono! Wysłano: {sent_count}"))

    def show_history(self, _):
        logs = self.conn.execute("SELECT msg, date FROM logs ORDER BY id DESC LIMIT 10").fetchall()
        content = "\n".join([f"{d}: {m}" for m, d in logs])
        self.popup("Historia (Ostatnie 10)", content if content else "Brak wysyłek")

# -----------------------------
# EXPORT POPUP (CHECKBOX)
# -----------------------------

    def export_popup(self, _):
        if not self.full_data: return
        header = self.full_data[0]
        box = BoxLayout(orientation="vertical", spacing=dp(8))
        checks = []
        for i, name in enumerate(header):
            row = BoxLayout()
            cb = CheckBox(active=True)
            checks.append((i, cb))
            row.add_widget(cb)
            row.add_widget(Label(text=name))
            box.add_widget(row)

        btn = PremiumButton(text="Zatwierdź kolumny")
        def start(_):
            self.export_columns = [i for i, c in checks if c.active]
            popup.dismiss()
            self.popup("OK", "Kolumny eksportu zostały ustawione")
        btn.bind(on_press=start)
        box.add_widget(btn)
        popup = Popup(title="Widoczność kolumn", content=box, size_hint=(0.9, 0.9))
        popup.open()

# -----------------------------
# POPUP
# -----------------------------

    def popup(self, title, text):
        box = BoxLayout(orientation="vertical", padding=dp(20))
        box.add_widget(Label(text=text, halign="center"))
        btn = PremiumButton(text="Okej")
        box.add_widget(btn)
        popup = Popup(title=title, content=box, size_hint=(0.8, 0.4))
        btn.bind(on_press=popup.dismiss)
        popup.open()


# -----------------------------
# APP START
# -----------------------------

if __name__ == "__main__":
    FutureApp().run()
