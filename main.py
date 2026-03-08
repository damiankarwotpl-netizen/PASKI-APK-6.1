import os
import sqlite3
import json
import threading
from pathlib import Path
from datetime import datetime

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

APP_TITLE = "Future 10.1 ULTIMATE"

class PremiumButton(Button):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.background_normal = ""
        self.background_color = (0.1, 0.4, 0.7, 1)
        self.color = (1, 1, 1, 1)
        self.size_hint_y = None
        self.height = dp(50)
        self.bold = True

class FutureApp(App):
    def build(self):
        Window.clearcolor = (0.05, 0.07, 0.1, 1)
        self.full_data = [] 
        self.current_file = None
        self.global_attachments = []
        self.export_col_indices = []
        
        self.init_db()
        self.sm = ScreenManager()
        
        self.home_scr = Screen(name="home"); self.setup_home()
        self.table_scr = Screen(name="table"); self.setup_table()
        self.email_scr = Screen(name="email"); self.setup_email_center()
        self.smtp_scr = Screen(name="smtp"); self.setup_smtp_ui()
        self.template_scr = Screen(name="template"); self.setup_template_ui()
        self.log_scr = Screen(name="logs"); self.setup_logs_ui()

        for s in [self.home_scr, self.table_scr, self.email_scr, self.smtp_scr, self.template_scr, self.log_scr]:
            self.sm.add_widget(s)
        
        if platform == 'android':
            from android.permissions import request_permissions, Permission
            request_permissions([Permission.READ_EXTERNAL_STORAGE, Permission.WRITE_EXTERNAL_STORAGE, Permission.INTERNET])

        return self.sm

    def init_db(self):
        db_p = Path(self.user_data_dir) / "app_v10.db"
        self.conn = sqlite3.connect(str(db_p), check_same_thread=False)
        self.conn.execute("CREATE TABLE IF NOT EXISTS contacts (name TEXT, surname TEXT, email TEXT, PRIMARY KEY(name, surname))")
        self.conn.execute("CREATE TABLE IF NOT EXISTS logs (id INTEGER PRIMARY KEY, recipient TEXT, status TEXT, date TEXT)")
        self.conn.execute("CREATE TABLE IF NOT EXISTS settings (key TEXT PRIMARY KEY, val TEXT)")
        if not self.conn.execute("SELECT val FROM settings WHERE key='t_body'").fetchone():
            self.conn.execute("INSERT INTO settings VALUES ('t_sub', 'Raport: {Imię} {Nazwisko}')")
            self.conn.execute("INSERT INTO settings VALUES ('t_body', 'Dzień dobry {Imię},\nPrzesyłamy raport.')")
        self.conn.commit()

    # --- UI SETUP ---
    def setup_home(self):
        l = BoxLayout(orientation="vertical", padding=dp(30), spacing=dp(15))
        l.add_widget(Label(text=APP_TITLE, font_size=dp(26), bold=True))
        l.add_widget(PremiumButton(text="\ud83d\udcc2 WCZYTAJ DANE", on_press=lambda x: self.pick_file("data")))
        l.add_widget(PremiumButton(text="\ud83d\udcca OTWÓRZ TABELĘ", on_press=self.go_to_table))
        l.add_widget(PremiumButton(text="\u2699 USTAWIENIA GMAIL", on_press=lambda x: setattr(self.sm, "current", "smtp")))
        self.h_stat = Label(text="Wybierz plik .xlsx", color=(0.7,0.7,0.7,1))
        l.add_widget(self.h_stat); self.home_scr.add_widget(l)

    def setup_table(self):
        l = BoxLayout(orientation="vertical", padding=dp(5))
        top = BoxLayout(size_hint_y=None, height=dp(55), spacing=dp(5))
        self.search = TextInput(hint_text="Szukaj...", multiline=False); self.search.bind(text=self.filter_data)
        top.add_widget(self.search)
        top.add_widget(Button(text="KOLUMNY", size_hint_x=0.3, on_press=self.column_picker_popup))
        top.add_widget(Button(text="WYŚLIJ", size_hint_x=0.3, on_press=lambda x: setattr(self.sm, "current", "email")))
        self.grid = GridLayout(cols=1, size_hint_y=None, spacing=dp(2)); self.grid.bind(minimum_height=self.grid.setter('height'))
        sv = ScrollView(); sv.add_widget(self.grid)
        self.prog = ProgressBar(max=100, size_hint_y=None, height=dp(10))
        l.add_widget(top); l.add_widget(sv); l.add_widget(self.prog)
        l.add_widget(Button(text="EKSPORTUJ PLIKI", size_hint_y=None, height=dp(45), on_press=self.mass_export_files))
        l.add_widget(Button(text="COFNIJ", size_hint_y=None, height=dp(45), on_press=lambda x: setattr(self.sm, "current", "home")))
        self.table_scr.add_widget(l)

    def setup_email_center(self):
        le = BoxLayout(orientation="vertical", padding=dp(20), spacing=dp(10))
        self.size_lbl = Label(text="Załączniki: 0 MB", size_hint_y=None, height=dp(30))
        btns = [
            ("\ud83d\udce5 IMPORTUJ KONTAKTY", lambda x: self.pick_file("book")),
            ("\u270f\ufe0f EDYTUJ SZABLON", lambda x: setattr(self.sm, "current", "template")),
            ("\ud83d\udcc1 INNE ZAŁĄCZNIKI", self.pop_attachments),
            ("\ud83d\udcc4 LOGI", self.load_logs_ui),
            ("\u26a1 TEST MAILA", self.run_test),
            ("\ud83d\ude80 WYŚLIJ MASOWO", self.run_mass_mail),
            ("POWRÓT", lambda x: setattr(self.sm, "current", "table"))
        ]
        le.add_widget(Label(text="Centrum Operacyjne", font_size=22, bold=True))
        le.add_widget(self.size_lbl)
        for t, c in btns: le.add_widget(PremiumButton(text=t, on_press=c))
        self.email_scr.add_widget(le)

    # --- LOGIKA OPERACYJNA (IMPORTY WEWNĄTRZ DLA STABILNOŚCI) ---
    def go_to_table(self, _):
        from openpyxl import load_workbook
        if not self.current_file: self.msg("Błąd", "Wczytaj Excel!"); return
        wb = load_workbook(str(self.current_file), data_only=True); ws = wb.active
        self.full_data = [[("" if v is None else str(v)) for v in r] for r in ws.iter_rows(values_only=True)]
        self.update_grid(self.full_data); self.sm.current = "table"

    def mass_export_files(self, _):
        if not self.full_data: return
        threading.Thread(target=self._export_worker).start()

    def _export_worker(self):
        from openpyxl import Workbook
        from openpyxl.styles import PatternFill, Border, Side, Font
        h, rows = self.full_data[0], self.full_data[1:]
        ni, si = self.get_indices(h)
        folder = Path("/storage/emulated/0/Documents/FutureExport"); folder.mkdir(parents=True, exist_ok=True)
        idxs = self.export_col_indices if self.export_col_indices else list(range(len(h)))
        
        for i, r in enumerate(rows):
            wb = Workbook(); ws = wb.active
            ws.append([h[k] for k in idxs]); ws.append([r[k] for k in idxs])
            name = f"{r[ni]}_{r[si]}".replace(" ", "_")
            wb.save(str(folder / f"Raport_{name}.xlsx"))
            Clock.schedule_once(lambda dt, p=int(((i+1)/len(rows))*100): setattr(self.prog, "value", p))
        Clock.schedule_once(lambda x: self.msg("Gites", "Zapisano w Documents/FutureExport"))

    def run_mass_mail(self, _):
        if not self.full_data: return
        threading.Thread(target=self._mail_worker).start()

    def _mail_worker(self):
        import smtplib
        from email.message import EmailMessage
        from openpyxl import Workbook
        
        p = Path(self.user_data_dir) / "smtp.json"
        if not p.exists(): Clock.schedule_once(lambda x: self.msg("!", "Brak SMTP")); return
        with open(p, "r") as f: cfg = json.load(f)
        
        try:
            srv = smtplib.SMTP("smtp.gmail.com", 587, timeout=15); srv.starttls(); srv.login(cfg['u'], cfg['p'])
        except Exception as e: Clock.schedule_once(lambda x: self.msg("Błąd", str(e))); return

        h, rows = self.full_data[0], self.full_data[1:]; ni, si = self.get_indices(h)
        for i, r in enumerate(rows):
            name, sur = str(r[ni]).strip(), str(r[si]).strip()
            res = self.conn.execute("SELECT email FROM contacts WHERE name=? AND surname=?", (name.lower(), sur.lower())).fetchone()
            if res:
                msg = self.build_email(cfg['u'], res[0], name, sur, h, r)
                try: srv.send_message(msg); self.log_to_db(res[0], "OK")
                except: self.log_to_db(res[0], "Fail")
            Clock.schedule_once(lambda dt, p=int(((i+1)/len(rows))*100): setattr(self.prog, "value", p))
        srv.quit(); Clock.schedule_once(lambda x: self.msg("Sukces", "Wysłano."))

    def build_email(self, sender, to, n, s, h, r):
        from email.message import EmailMessage
        from openpyxl import Workbook
        msg = EmailMessage(); today = datetime.now().strftime("%d.%m.%Y")
        rs = self.conn.execute("SELECT val FROM settings WHERE key='t_sub'").fetchone()
        rb = self.conn.execute("SELECT val FROM settings WHERE key='t_body'").fetchone()
        msg["Subject"] = rs[0].replace("{Imię}", n).replace("{Nazwisko}", s)
        msg["From"], msg["To"] = sender, to
        msg.set_content(rb[0].replace("{Imię}", n).replace("{Nazwisko}", s).replace("{Data}", today))
        
        # Załącznik
        tmp = Path(self.user_data_dir) / "m.xlsx"; wb = Workbook(); ws = wb.active
        ws.append(h); ws.append(r); wb.save(str(tmp))
        with open(tmp, "rb") as f: msg.add_attachment(f.read(), maintype="application", subtype="xlsx", filename=f"Raport_{n}.xlsx")
        return msg

    # --- POMOCNICZE ---
    def pick_file(self, mode):
        if platform != 'android': return
        from jnius import autoclass; from android import activity
        Intent = autoclass("android.content.Intent")
        intent = Intent(Intent.ACTION_OPEN_DOCUMENT); intent.setType("*/*"); intent.addCategory(Intent.CATEGORY_OPENABLE)
        def on_res(req, res, dt):
            if dt:
                try:
                    uri = dt.getData(); ctx = autoclass("org.kivy.android.PythonActivity").mActivity
                    stream = ctx.getContentResolver().openInputStream(uri)
                    dest = Path(self.user_data_dir) / (f"extra_{datetime.now().microsecond}.tmp" if mode=="extra" else f"{mode}.xlsx")
                    with open(dest, "wb") as f:
                        j_buf = autoclass('[B')(16384)
                        while True:
                            r = stream.read(j_buf)
                            if r <= 0: break
                            f.write(bytes(j_buf)[:r])
                    if mode == "data": self.current_file = dest; Clock.schedule_once(lambda x: setattr(self.h_stat, "text", "Załadowano."))
                    elif mode == "book": self.import_contacts(dest)
                    elif mode == "extra": self.global_attachments.append(str(dest)); self.update_att_txt()
                except: pass
            activity.unbind(on_activity_result=on_res)
        activity.bind(on_activity_result=on_res); autoclass("org.kivy.android.PythonActivity").mActivity.startActivityForResult(intent, 1001)

    def import_contacts(self, path):
        from openpyxl import load_workbook
        wb = load_workbook(str(path), data_only=True); ws = wb.active; rows = list(ws.iter_rows(values_only=True))
        h = [str(x).lower() for x in rows[0]]; ni, si = self.get_indices(h)
        mi = next((i for i, v in enumerate(h) if "mail" in v), 2)
        for r in rows[1:]:
            if r[mi]: self.conn.execute("INSERT OR REPLACE INTO contacts VALUES(?,?,?)", (str(r[ni]).lower().strip(), str(r[si]).lower().strip(), str(r[mi]).strip()))
        self.conn.commit(); self.msg("OK", "Kontakty wczytane.")

    def update_grid(self, data):
        self.grid.clear_widgets()
        for r_data in data[:50]:
            row = BoxLayout(size_hint_y=None, height=dp(30))
            for cell in r_data[:3]: row.add_widget(Label(text=str(cell)[:15], font_size=11))
            self.grid.add_widget(row)

    def get_indices(self, h):
        ni, si = 0, 1
        for i, v in enumerate(h):
            val = str(v).lower()
            if "imi" in val: ni = i
            if "nazw" in val: si = i
        return ni, si

    def log_to_db(self, mail, status):
        self.conn.execute("INSERT INTO logs (recipient, status, date) VALUES (?,?,?)", (mail, status, datetime.now().strftime("%H:%M")))
        self.conn.commit()

    def msg(self, t, txt): Popup(title=t, content=Label(text=txt, halign="center"), size_hint=(0.8, 0.4)).open()

    def update_att_txt(self):
        sz = sum(os.path.getsize(p) for p in self.global_attachments) / (1024*1024)
        self.size_lbl.text = f"Załączniki: {sz:.2f} MB"

    def column_picker_popup(self, _):
        if not self.full_data: return
        box = BoxLayout(orientation="vertical", padding=10)
        grid = GridLayout(cols=1, size_hint_y=None, spacing=5); grid.bind(minimum_height=grid.setter('height'))
        checks = []
        for i, h in enumerate(self.full_data[0]):
            r = BoxLayout(size_hint_y=None, height=dp(40))
            cb = CheckBox(size_hint_x=0.2); cb.active = True
            r.add_widget(cb); r.add_widget(Label(text=str(h))); grid.add_widget(r); checks.append((i, cb))
        def apply(_): self.export_col_indices = [idx for idx, c in checks if c.active]; p.dismiss()
        sv = ScrollView(); sv.add_widget(grid); box.add_widget(sv)
        box.add_widget(Button(text="ZAPISZ", size_hint_y=None, height=dp(50), on_press=apply))
        p = Popup(title="Kolumny", content=box, size_hint=(0.9, 0.8)); p.open()

    def setup_smtp_ui(self):
        l = BoxLayout(orientation="vertical", padding=dp(25), spacing=dp(10))
        self.u_in = TextInput(hint_text="Gmail"); self.p_in = TextInput(hint_text="Hasło App", password=True)
        def save(_):
            with open(Path(self.user_data_dir) / "smtp.json", "w") as f: json.dump({'u':self.u_in.text,'p':self.p_in.text}, f)
            self.msg("OK", "Zapisano."); setattr(self.sm, "current", "home")
        l.add_widget(Label(text="Ustawienia Gmail")); l.add_widget(self.u_in); l.add_widget(self.p_in)
        l.add_widget(Button(text="ZAPISZ", on_press=save)); l.add_widget(Button(text="COFNIJ", on_press=lambda x: setattr(self.sm, "current", "home")))
        p = Path(self.user_data_dir) / "smtp.json"
        if p.exists(): d = json.load(open(p)); self.u_in.text, self.p_in.text = d['u'], d['p']
        self.smtp_scr.add_widget(l)

    def setup_template_ui(self):
        l = BoxLayout(orientation="vertical", padding=dp(20), spacing=dp(10))
        self.ts = TextInput(size_hint_y=None, height=dp(45)); self.tb = TextInput()
        def save(_):
            self.conn.execute("UPDATE settings SET val=? WHERE key='t_sub'", (self.ts.text,))
            self.conn.execute("UPDATE settings SET val=? WHERE key='t_body'", (self.tb.text,)); self.conn.commit(); self.msg("OK", "Zapisano.")
        l.add_widget(self.ts); l.add_widget(self.tb); l.add_widget(Button(text="ZAPISZ", on_press=save)); l.add_widget(Button(text="COFNIJ", on_press=lambda x: setattr(self.sm, "current", "email")))
        self.template_scr.add_widget(l)

    def pop_attachments(self, _):
        box = BoxLayout(orientation="vertical", padding=10)
        box.add_widget(Button(text="DODAJ PLIK", on_press=lambda x: self.pick_file("extra")))
        box.add_widget(Button(text="ZAMKNIJ", on_press=lambda x: self.at_pop.dismiss()))
        self.at_pop = Popup(title="Załączniki", content=box, size_hint=(0.8, 0.5)); self.at_pop.open()

    def filter_data(self, ins, val):
        if not self.full_data: return
        f = [self.full_data[0]] + [r for r in self.full_data[1:] if val.lower() in str(r).lower()]
        self.update_grid(f)

    def run_test(self, _):
        if self.full_data: threading.Thread(target=self._test_task).start()

    def _test_task(self):
        try:
            p = Path(self.user_data_dir) / "smtp.json"; cfg = json.load(open(p))
            import smtplib; srv = smtplib.SMTP("smtp.gmail.com", 587, timeout=10); srv.starttls(); srv.login(cfg['u'], cfg['p'])
            msg = self.build_email(cfg['u'], cfg['u'], "TEST", "URUCHOMIONY", self.full_data[0], self.full_data[1])
            srv.send_message(msg); srv.quit(); Clock.schedule_once(lambda x: self.msg("OK", "Test wysłany."))
        except Exception as e: Clock.schedule_once(lambda x: self.msg("Błąd", str(e)))

    def setup_logs_ui(self): pass
    def load_logs_ui(self, _):
        box = BoxLayout(orientation="vertical", padding=10)
        logs = self.conn.execute("SELECT recipient, status, date FROM logs ORDER BY id DESC LIMIT 20").fetchall()
        for r, s, d in logs: box.add_widget(Label(text=f"{d} | {r} | {s}", font_size=10))
        box.add_widget(Button(text="ZAMKNIJ", on_press=lambda x: p.dismiss()))
        p = Popup(title="Logi", content=box, size_hint=(0.9, 0.9)); p.open()

if __name__ == "__main__":
    try:
        FutureApp().run()
    except Exception as e:
        # Próba zapisania błędu do pliku, jeśli aplikacja padnie przed startem UI
        import traceback
        with open("/sdcard/Documents/crash_log.txt", "w") as f:
            f.write(str(e))
            f.write("\n\nFull Traceback:\n")
            f.write(traceback.format_exc())
