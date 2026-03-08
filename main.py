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

class HomeScreen(Screen): pass
class TableScreen(Screen): pass
class EmailScreen(Screen): pass
class SMTPScreen(Screen): pass

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
        self.home, self.table = HomeScreen(name="home"), TableScreen(name="table")
        self.email, self.smtp = EmailScreen(name="email"), SMTPScreen(name="smtp")

        self.build_home()
        self.build_table()
        self.build_email()
        self.build_smtp()

        for s in [self.home, self.table, self.email, self.smtp]: self.sm.add_widget(s)
        
        apply_mail_patch(self)
        return self.sm

    def build_home(self):
        layout = BoxLayout(orientation="vertical", padding=dp(25), spacing=dp(20))
        layout.add_widget(Label(text=APP_TITLE, font_size=26))
        
        btns = [
            ("📂 Otwórz Excel", self.open_picker),
            ("📊 Wczytaj dane", self.load_excel),
            ("⚙ SMTP", lambda x: setattr(self.sm, "current", "smtp"))
        ]
        for txt, cmd in btns:
            b = PremiumButton(text=txt)
            b.bind(on_press=cmd)
            layout.add_widget(b)

        self.status = Label(text="Gotowy")
        layout.add_widget(self.status)
        self.home.add_widget(layout)

    def open_picker(self, _):
        if platform != "android":
            self.status.text = "Picker tylko Android"; return
        from jnius import autoclass
        from android import activity
        Intent = autoclass("android.content.Intent")
        PythonActivity = autoclass("org.kivy.android.PythonActivity")
        intent = Intent(Intent.ACTION_OPEN_DOCUMENT)
        intent.setType("*/*")
        intent.addCategory(Intent.CATEGORY_OPENABLE)
        activity.bind(on_activity_result=self.picker_result)
        PythonActivity.mActivity.startActivityForResult(intent, 1001)

    def picker_result(self, request_code, result_code, intent):
        if request_code != 1001 or not intent: return
        from pathlib import Path
        from jnius import autoclass
        from android import activity
        activity.unbind(on_activity_result=self.picker_result)
        resolver = autoclass("org.kivy.android.PythonActivity").mActivity.getContentResolver()
        stream = resolver.openInputStream(intent.getData())
        local = Path(self.user_data_dir) / "excel.xlsx"
        with open(local, "wb") as f:
            buf = bytearray(4096)
            while True:
                r = stream.read(buf)
                if r == -1: break
                f.write(buf[:r])
        stream.close()
        self.current_file = local
        self.status.text = "Plik wybrany"

    def load_excel(self, _):
        if not self.current_file: self.popup("Błąd", "Wybierz plik"); return
        from openpyxl import load_workbook
        wb = load_workbook(self.current_file, data_only=True)
        sheet = wb.active
        self.full_data = [["" if v is None else str(v) for v in row] for row in sheet.iter_rows(values_only=True)]
        wb.close()
        self.filtered_data = self.full_data
        self.show_table()
        self.sm.current = "table"

    def build_table(self):
        layout = BoxLayout(orientation="vertical", padding=dp(10), spacing=dp(10))
        top = BoxLayout(size_hint=(1, 0.12), spacing=dp(8))
        self.search = TextInput(hint_text="Szukaj...", multiline=False)
        self.search.bind(text=self.filter_data)
        
        for t, c in [("Export", self.export_popup), ("Email", lambda x: setattr(self.sm, "current", "email")), ("Powrót", lambda x: setattr(self.sm, "current", "home"))]:
            b = PremiumButton(text=t); b.bind(on_press=c); top.add_widget(b)

        top.add_widget(self.search, index=4)
        self.scroll = ScrollView()
        self.grid = GridLayout(size_hint=(None, None))
        self.grid.bind(minimum_height=self.grid.setter("height"), minimum_width=self.grid.setter("width"))
        self.scroll.add_widget(self.grid)
        self.progress = ProgressBar(max=100, size_hint=(1, 0.05))
        layout.add_widget(top); layout.add_widget(self.scroll); layout.add_widget(self.progress)
        self.table.add_widget(layout)

    def show_table(self):
        self.grid.clear_widgets()
        if not self.filtered_data: return
        rows, cols = len(self.filtered_data), len(self.filtered_data[0])
        w, h = dp(160), dp(42)
        self.grid.cols = cols
        self.grid.width, self.grid.height = cols * w, rows * h
        for row in self.filtered_data:
            for cell in row: self.grid.add_widget(Label(text=str(cell), size_hint=(None, None), size=(w, h)))

    def filter_data(self, instance, value):
        v = value.lower()
        self.filtered_data = [r for r in self.full_data if any(v in str(c).lower() for c in r)]
        self.show_table()

    def export_popup(self, _):
        if not self.full_data: return
        box = BoxLayout(orientation="vertical", spacing=dp(8))
        checks = []
        for i, name in enumerate(self.full_data[0]):
            rl = BoxLayout(); cb = CheckBox(); checks.append((i, cb))
            rl.add_widget(cb); rl.add_widget(Label(text=str(name))); box.add_widget(rl)
        
        def start(_):
            import threading
            self.export_columns = [i for i, c in checks if c.active]
            p.dismiss()
            threading.Thread(target=self.export_excel).start()
        
        btn = PremiumButton(text="Start"); btn.bind(on_press=start); box.add_widget(btn)
        p = Popup(title="Kolumny", content=box, size_hint=(0.9, 0.9)); p.open()

    def build_email(self):
        layout = BoxLayout(orientation="vertical", padding=dp(20), spacing=dp(15))
        self.email_status = Label(text="")
        btns = [
            ("📚 Import adresów", lambda x: import_address_excel(self, self.current_file) if self.current_file else self.popup("Błąd", "Brak pliku")),
            ("Wybierz kolumnę email", self.select_email_column),
            ("Wyślij emaile", self.send_emails),
            ("Powrót", lambda x: setattr(self.sm, "current", "table"))
        ]
        layout.add_widget(Label(text="Email Sender", font_size=22))
        for t, c in btns:
            b = PremiumButton(text=t); b.bind(on_press=c); layout.add_widget(b)
        layout.add_widget(self.email_status); self.email.add_widget(layout)

    def select_email_column(self, _):
        if not self.full_data: return
        box = BoxLayout(orientation="vertical", spacing=dp(10))
        checks = []
        for i, col in enumerate(self.full_data[0]):
            rl = BoxLayout(); cb = CheckBox(group="email"); checks.append((i, cb))
            rl.add_widget(cb); rl.add_widget(Label(text=str(col))); box.add_widget(rl)
        def save(_):
            for i, c in checks:
                if c.active: self.email_columns = [i]
            p.dismiss()
        btn = PremiumButton(text="OK"); btn.bind(on_press=save); box.add_widget(btn)
        p = Popup(title="Wybierz email", content=box, size_hint=(0.9, 0.9)); p.open()

    def build_smtp(self):
        layout = BoxLayout(orientation="vertical", padding=dp(20), spacing=dp(10))
        self.smtp_server = TextInput(hint_text="SMTP server")
        self.smtp_port = TextInput(hint_text="Port")
        self.smtp_user = TextInput(hint_text="Email")
        self.smtp_pass = TextInput(hint_text="Hasło", password=True)
        save = PremiumButton(text="Zapisz"); save.bind(on_press=self.save_smtp)
        back = PremiumButton(text="Powrót"); back.bind(on_press=lambda x: setattr(self.sm, "current", "home"))
        for w in [self.smtp_server, self.smtp_port, self.smtp_user, self.smtp_pass, save, back]: layout.add_widget(w)
        self.smtp.add_widget(layout)

    def save_smtp(self, _):
        import json
        from pathlib import Path
        data = {"server": self.smtp_server.text, "port": self.smtp_port.text, "user": self.smtp_user.text, "pass": self.smtp_pass.text}
        with open(Path(self.user_data_dir) / "smtp.json", "w") as f: json.dump(data, f)
        self.popup("OK", "Zapisano")

    def load_smtp(self):
        import json
        from pathlib import Path
        p = Path(self.user_data_dir) / "smtp.json"
        if p.exists():
            with open(p) as f: return json.load(f)
        return None

    def send_emails(self, _):
        if not self.email_columns: self.popup("Błąd", "Wybierz kolumnę email")
        else:
            import threading
            threading.Thread(target=self._email_thread).start()

    def export_excel(self):
        from openpyxl import Workbook
        from pathlib import Path
        from datetime import datetime
        if not self.filtered_data: return
        folder = Path("/storage/emulated/0/Documents/FutureExport"); folder.mkdir(parents=True, exist_ok=True)
        h_orig = self.filtered_data[0]
        h = [h_orig[i] for i in self.export_columns] if self.export_columns else h_orig
        rows = self.filtered_data[1:]
        for i, row in enumerate(rows):
            wb = Workbook(); ws = wb.active; ws.append(h)
            ws.append([row[j] for j in self.export_columns] if self.export_columns else row)
            wb.save(folder / f"{str(row[0])}_{datetime.now().strftime('%H%M%S')}.xlsx")
            Clock.schedule_once(lambda dt, p=int((i+1)/len(rows)*100): setattr(self.progress, "value", p))
        Clock.schedule_once(lambda dt: self.popup("Export", "Gotowe"))

    def popup(self, title, text):
        box = BoxLayout(orientation="vertical", padding=dp(20))
        box.add_widget(Label(text=text))
        btn = PremiumButton(text="OK"); box.add_widget(btn)
        p = Popup(title=title, content=box, size_hint=(0.7, 0.4))
        btn.bind(on_press=p.dismiss); p.open()

def apply_mail_patch(app):
    import sqlite3
    from pathlib import Path
    db_path = Path(app.user_data_dir) / "mail.db"
    conn = sqlite3.connect(db_path)
    conn.execute("CREATE TABLE IF NOT EXISTS address_book(id INTEGER PRIMARY KEY, name TEXT, surname TEXT, email TEXT)")
    conn.commit()

    def mail_thread():
        import smtplib
        from email.message import EmailMessage
        from openpyxl import Workbook
        from datetime import datetime
        smtp = app.load_smtp()
        if not smtp: Clock.schedule_once(lambda dt: app.popup("SMTP", "Brak konfiguracji")); return
        try:
            srv = smtplib.SMTP(smtp["server"], int(smtp["port"]))
            srv.starttls(); srv.login(smtp["user"], smtp["pass"])
        except Exception as e: Clock.schedule_once(lambda dt: app.popup("Error", str(e))); return
        
        rows = app.full_data[1:]; h = app.full_data[0]; sent = 0
        for i, row in enumerate(rows):
            email = None
            if len(row) >= 2:
                r = conn.execute("SELECT email FROM address_book WHERE name=? AND surname=?", (str(row[0]).lower(), str(row[1]).lower())).fetchone()
                if r: email = r[0]
            if not email and app.email_columns: email = row[app.email_columns[0]]
            
            if email and "@" in str(email):
                try:
                    msg = EmailMessage(); msg["Subject"] = "Info"; msg["From"] = smtp["user"]; msg["To"] = email
                    msg.set_content("Plik w załączniku")
                    wb = Workbook(); ws = wb.active; ws.append(h); ws.append(row)
                    path = Path(app.user_data_dir) / "temp.xlsx"; wb.save(path)
                    with open(path, "rb") as f: msg.add_attachment(f.read(), maintype="application", subtype="xlsx", filename="dane.xlsx")
                    srv.send_message(msg); sent += 1
                except: pass
            Clock.schedule_once(lambda dt, p=int((i+1)/len(rows)*100): setattr(app.progress, "value", p))
        srv.quit(); Clock.schedule_once(lambda dt: app.popup("Koniec", f"Wysłano {sent}"))

    app._email_thread = mail_thread
    app.mail_db_conn = conn

def import_address_excel(app, path):
    from openpyxl import load_workbook
    try:
        wb = load_workbook(path, data_only=True); sheet = wb.active
        rows = list(sheet.iter_rows(values_only=True))
        h = [str(x).lower() for x in rows[0]]
        ni, si, mi = h.index("imię"), h.index("nazwisko"), h.index("email")
        added = 0
        for r in rows[1:]:
            if r[mi]:
                app.mail_db_conn.execute("INSERT INTO address_book(name,surname,email) VALUES(?,?,?)", (str(r[ni]).lower(), str(r[si]).lower(), str(r[mi])))
                added += 1
        app.mail_db_conn.commit(); app.popup("Sukces", f"Dodano {added} adresów")
    except Exception as e: app.popup("Błąd", str(e))

if __name__ == "__main__":
    FutureApp().run()
