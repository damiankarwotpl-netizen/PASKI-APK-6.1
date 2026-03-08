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

APP_TITLE = "Future 10.0 FINAL ULTIMATE"

class PremiumButton(Button):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.background_normal = ""
        self.background_color = (0.1, 0.35, 0.65, 1)
        self.color = (1, 1, 1, 1)
        self.size_hint_y = None
        self.height = dp(50)
        self.bold = True

class FutureApp(App):
    def build(self):
        # Importy systemowe wewnątrz build
        import os
        import sqlite3
        from pathlib import Path

        Window.clearcolor = (0.04, 0.06, 0.1, 1)
        self.full_data = [] # Wszystkie dane z głównego Excela
        self.current_file = None
        self.global_attachments = []
        self.export_col_indices = [] # Indeksy wybranych kolumn
        
        self.init_db()
        self.sm = ScreenManager()
        
        # Inicjalizacja wszystkich ekranów
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
        import sqlite3
        from pathlib import Path
        db_p = Path(self.user_data_dir) / "app_v10_final.db"
        self.conn = sqlite3.connect(str(db_p), check_same_thread=False)
        self.conn.execute("CREATE TABLE IF NOT EXISTS contacts (name TEXT, surname TEXT, email TEXT, PRIMARY KEY(name, surname))")
        self.conn.execute("CREATE TABLE IF NOT EXISTS logs (id INTEGER PRIMARY KEY, recipient TEXT, status TEXT, date TEXT)")
        self.conn.execute("CREATE TABLE IF NOT EXISTS settings (key TEXT PRIMARY KEY, val TEXT)")
        
        if not self.conn.execute("SELECT val FROM settings WHERE key='t_body'").fetchone():
            self.conn.execute("INSERT INTO settings VALUES ('t_sub', 'Raport: {Imię} {Nazwisko}')")
            self.conn.execute("INSERT INTO settings VALUES ('t_body', 'Dzień dobry {Imię},\n\nPrzesyłamy raport miesięczny.')")
        self.conn.commit()

    # --- UI HOME ---
    def setup_home(self):
        l = BoxLayout(orientation="vertical", padding=dp(30), spacing=dp(15))
        l.add_widget(Label(text=APP_TITLE, font_size=dp(28), bold=True))
        b1 = PremiumButton(text="\ud83d\udcc2 WCZYTAJ DANE GŁÓWNE"); b1.bind(on_press=lambda x: self.pick_file("data"))
        b2 = PremiumButton(text="\ud83d\udcca OTWÓRZ TABELĘ"); b2.bind(on_press=self.go_to_table)
        b3 = PremiumButton(text="\u2699 USTAWIENIA GMAIL"); b3.bind(on_press=lambda x: setattr(self.sm, "current", "smtp"))
        self.h_stat = Label(text="Oczekiwanie na plik Excel...", color=(0.6,0.6,0.6,1))
        for w in [b1, b2, b3, self.h_stat]: l.add_widget(w)
        self.home_scr.add_widget(l)

    # --- UI TABELA ---
    def setup_table(self):
        l = BoxLayout(orientation="vertical", padding=dp(5))
        top = BoxLayout(size_hint_y=None, height=dp(55), spacing=dp(2))
        self.search = TextInput(hint_text="Szukaj osoby...", multiline=False); self.search.bind(text=self.filter_data)
        top.add_widget(self.search)
        top.add_widget(Button(text="KOLUMNY", size_hint_x=0.25, background_color=(0.5,0.4,0,1), on_press=self.column_picker_popup))
        top.add_widget(Button(text="MAIL-CENTER", size_hint_x=0.25, background_color=(0,0.5,0,1), on_press=lambda x: setattr(self.sm, "current", "email")))
        
        self.grid = GridLayout(cols=1, size_hint_y=None, spacing=dp(2)); self.grid.bind(minimum_height=self.grid.setter('height'))
        sv = ScrollView(); sv.add_widget(self.grid)
        self.prog = ProgressBar(max=100, size_hint_y=None, height=dp(10))
        l.add_widget(top); l.add_widget(sv); l.add_widget(self.prog)
        l.add_widget(Button(text="EKSPORTUJ WSZYSTKO NA TELEFON", size_hint_y=None, height=dp(45), on_press=self.mass_export_files))
        l.add_widget(Button(text="COFNIJ", size_hint_y=None, height=dp(45), on_press=lambda x: setattr(self.sm, "current", "home")))
        self.table_scr.add_widget(l)

    # --- UI EMAIL CENTER ---
    def setup_email_center(self):
        le = BoxLayout(orientation="vertical", padding=dp(20), spacing=dp(8))
        self.size_lbl = Label(text="Załączniki ogólne: 0.00 MB", size_hint_y=None, height=dp(30))
        btns = [
            ("\ud83d\udce5 IMPORTUJ BAZĘ KONTAKTÓW", lambda x: self.pick_file("book")),
            ("\u270f\ufe0f EDYTUJ TREŚĆ MAILA", lambda x: setattr(self.sm, "current", "template")),
            ("\ud83d\udcc1 DODAJ PLIKI OGÓLNE", self.pop_attachments),
            ("\ud83d\udcc4 HISTORIA (LOGI)", self.load_logs_into_ui),
            ("\u26a1 TEST - WYŚLIJ DO SIEBIE", self.run_test),
            ("\ud83d\ude80 URUCHOM MASOWĄ WYSYŁKĘ", self.run_mass_mail),
            ("POWRÓT DO TABELI", lambda x: setattr(self.sm, "current", "table"))
        ]
        le.add_widget(Label(text="Centrum Operacyjne", font_size=22, bold=True))
        le.add_widget(self.size_lbl)
        for t, c in btns: le.add_widget(PremiumButton(text=t, on_press=c))
        self.email_scr.add_widget(le)

    # --- LOGIKA IDENTYFIKACJI ---
    def find_name_indices(self, h):
        ni, si = 0, 1
        for i, v in enumerate(h):
            val = str(v).lower()
            if "imi" in val or "name" in val: ni = i
            if "nazw" in val or "sur" in val: si = i
        return ni, si

    # --- EKSPORT PLIKÓW NA TELEFON ---
    def mass_export_files(self, _):
        import threading
        if not self.full_data: return
        threading.Thread(target=self._mass_export_worker).start()

    def _mass_export_worker(self):
        from openpyxl import Workbook
        from openpyxl.styles import PatternFill, Border, Side, Font
        import os
        from pathlib import Path
        from datetime import datetime

        h, rows = self.full_data[0], self.full_data[1:]
        ni, si = self.find_name_indices(h)
        folder = Path("/storage/emulated/0/Documents/FutureExport")
        try: folder.mkdir(parents=True, exist_ok=True)
        except: pass
        
        idxs = self.export_col_indices if self.export_col_indices else list(range(len(h)))
        filtered_h = [h[i] for i in idxs]

        for i, row in enumerate(rows):
            wb = Workbook(); ws = wb.active
            blue = PatternFill(start_color='CFE2F3', end_color='CFE2F3', fill_type='solid')
            bd = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
            
            ws.append(filtered_h)
            for cell in ws[1]: cell.fill, cell.font, cell.border = blue, Font(bold=True), bd
            ws.append([row[k] for k in idxs])
            for cell in ws[2]: cell.border = bd
            for col in ws.columns: ws.column_dimensions[col[0].column_letter].width = 22

            safe_name = f"{row[ni]}_{row[si]}".replace(" ", "_").replace("/", "-")
            wb.save(str(folder / f"Raport_{safe_name}_{datetime.now().strftime('%H%M%S')}.xlsx"))
            Clock.schedule_once(lambda dt, p=int(((i+1)/len(rows))*100): setattr(self.prog, "value", p))
        Clock.schedule_once(lambda x: self.msg("Sukces", f"Zapisano w Documents/FutureExport"))

    # --- WYSYŁKA MASOWA ---
    def run_mass_mail(self, _):
        import threading
        if not self.full_data: return
        threading.Thread(target=self._mailing_worker).start()

    def _mailing_worker(self):
        import smtplib
        import json
        from pathlib import Path
        from datetime import datetime
        
        p_cfg = Path(self.user_data_dir) / "smtp.json"
        if not p_cfg.exists(): Clock.schedule_once(lambda x: self.msg("Błąd", "Skonfiguruj ustawienia SMTP!")); return
        with open(p_cfg, "r") as f: cfg = json.load(f)
        
        try:
            srv = smtplib.SMTP("smtp.gmail.com", 587, timeout=15); srv.starttls(); srv.login(cfg['u'], cfg['p'])
        except Exception as e: Clock.schedule_once(lambda x: self.msg("Błąd SMTP", str(e))); return

        h, rows = self.full_data[0], self.full_data[1:]
        ni, si = self.find_name_indices(h)
        sent = 0

        for i, r in enumerate(rows):
            name, sur = str(r[ni]).strip(), str(r[si]).strip()
            res = self.conn.execute("SELECT email FROM contacts WHERE name=? AND surname=?", (name.lower(), sur.lower())).fetchone()
            if res:
                try:
                    msg = self.create_email_obj(cfg['u'], res[0], name, sur, h, r)
                    srv.send_message(msg); sent += 1
                    self.conn.execute("INSERT INTO logs (recipient, status, date) VALUES (?,?,?)", (res[0], "Wysłano", datetime.now().strftime("%H:%M")))
                except: pass
            Clock.schedule_once(lambda dt, p=int(((i+1)/len(rows))*100): setattr(self.prog, "value", p))
        
        self.conn.commit(); srv.quit()
        Clock.schedule_once(lambda x: self.msg("Finał", f"Pomyślnie wysłano {sent} maili."))

    def create_email_obj(self, sender, to, name, sur, h, r):
        from email.message import EmailMessage
        from openpyxl import Workbook
        from pathlib import Path
        import mimetypes
        import os
        from datetime import datetime

        msg = EmailMessage(); today = datetime.now().strftime("%d.%m.%Y")
        res_s = self.conn.execute("SELECT val FROM settings WHERE key='t_sub'").fetchone()
        res_b = self.conn.execute("SELECT val FROM settings WHERE key='t_body'").fetchone()
        
        sub = res_s[0].replace("{Imię}", name).replace("{Nazwisko}", sur).replace("{Data}", today)
        body = res_b[0].replace("{Imię}", name).replace("{Nazwisko}", sur).replace("{Data}", today)
        
        msg["Subject"], msg["From"], msg["To"] = sub, sender, to
        msg.set_content(body)

        # Spersonalizowany arkusz
        idxs = self.export_col_indices if self.export_col_indices else list(range(len(h)))
        tmp = Path(self.user_data_dir) / "mail.xlsx"; wb = Workbook(); ws = wb.active
        ws.append([h[k] for k in idxs]); ws.append([r[k] for k in idxs])
        wb.save(str(tmp))
        with open(tmp, "rb") as f:
            msg.add_attachment(f.read(), maintype="application", subtype="xlsx", filename=f"Raport_{name}_{sur}.xlsx")
        
        # Pozostałe załączniki
        for p in self.global_attachments:
            if os.path.exists(p):
                ct, _ = mimetypes.guess_type(p)
                m, s = (ct or "application/octet-stream").split("/", 1)
                with open(p, "rb") as f: msg.add_attachment(f.read(), maintype=m, subtype=s, filename=os.path.basename(p))
        return msg

    # --- WYBÓR PLIKÓW ANDROID ---
    def pick_file(self, mode):
        if platform != "android": self.msg("Błąd", "Dostępne tylko na Androidzie."); return
        from jnius import autoclass; from android import activity
        from pathlib import Path
        from datetime import datetime
        
        Intent = autoclass("android.content.Intent")
        intent = Intent(Intent.ACTION_OPEN_DOCUMENT); intent.setType("*/*"); intent.addCategory(Intent.CATEGORY_OPENABLE)
        
        def on_res(req, res, dt):
            if dt:
                try:
                    uri = dt.getData(); ctx = autoclass("org.kivy.android.PythonActivity").mActivity
                    stream = ctx.getContentResolver().openInputStream(uri)
                    dest = Path(self.user_data_dir) / (f"ex_{datetime.now().microsecond}.tmp" if mode=="extra" else f"{mode}.xlsx")
                    with open(dest, "wb") as f:
                        j_buf = autoclass('[B')(16384)
                        while True:
                            r = stream.read(j_buf)
                            if r <= 0: break
                            f.write(bytes(j_buf)[:r])
                    
                    if mode == "data": 
                        self.current_file = dest; Clock.schedule_once(lambda x: setattr(self.h_stat, "text", "Załadowano Excel."))
                    elif mode == "book": 
                        self.import_contacts_db(dest)
                    elif mode == "extra": 
                        self.global_attachments.append(str(dest)); self.update_att_txt()
                except Exception as e: Clock.schedule_once(lambda x: self.msg("Błąd pliku", str(e)))
            activity.unbind(on_activity_result=on_res)
        activity.bind(on_activity_result=on_res); autoclass("org.kivy.android.PythonActivity").mActivity.startActivityForResult(intent, 1001)

    # --- POPUPY I POMOCNICZE ---
    def column_picker_popup(self, _):
        if not self.full_data: return
        box = BoxLayout(orientation="vertical", padding=10)
        grid = GridLayout(cols=1, size_hint_y=None, spacing=5); grid.bind(minimum_height=grid.setter('height'))
        checks = []
        for i, h in enumerate(self.full_data[0]):
            r = BoxLayout(size_hint_y=None, height=dp(40))
            cb = CheckBox(size_hint_x=0.2); cb.active = (i in self.export_col_indices or not self.export_col_indices)
            r.add_widget(cb); r.add_widget(Label(text=str(h))); grid.add_widget(r); checks.append((i, cb))
        def apply(_): self.export_col_indices = [idx for idx, c in checks if c.active]; p.dismiss()
        sv = ScrollView(); sv.add_widget(grid); box.add_widget(sv)
        b = Button(text="ZAPISZ WYBÓR", size_hint_y=None, height=dp(50)); b.bind(on_press=apply)
        box.add_widget(b); p = Popup(title="Wybór kolumn", content=box, size_hint=(0.9, 0.8)); p.open()

    def go_to_table(self, _):
        from openpyxl import load_workbook
        if not self.current_file: self.msg("Błąd", "Najpierw wczytaj Excel!"); return
        wb = load_workbook(str(self.current_file), data_only=True); ws = wb.active
        self.full_data = [[("" if v is None else str(v)) for v in r] for r in ws.iter_rows(values_only=True)]
        self.update_grid_ui(self.full_data); self.sm.current = "table"

    def update_grid_ui(self, data):
        self.grid.clear_widgets()
        for i, row in enumerate(data[:100]):
            br = BoxLayout(size_hint_y=None, height=dp(35))
            for cell in row[:4]: br.add_widget(Label(text=str(cell)[:15], font_size=dp(11)))
            self.grid.add_widget(br)

    def filter_data(self, ins, val):
        if not self.full_data: return
        f = [self.full_data[0]] + [r for r in self.full_data[1:] if val.lower() in str(r).lower()]
        self.update_grid_ui(f)

    def run_test(self, _):
        import threading
        if not self.full_data: return
        threading.Thread(target=self._test_task).start()

    def _test_task(self):
        import smtplib, json
        from pathlib import Path
        try:
            p = Path(self.user_data_dir) / "smtp.json"
            cfg = json.load(open(p))
            srv = smtplib.SMTP("smtp.gmail.com", 587, timeout=10); srv.starttls(); srv.login(cfg['u'], cfg['p'])
            msg = self.create_email_obj(cfg['u'], cfg['u'], "TEST", "URUCHOMIONY", self.full_data[0], self.full_data[1])
            msg["Subject"] = "[TEST] " + msg["Subject"]
            srv.send_message(msg); srv.quit(); Clock.schedule_once(lambda x: self.msg("OK", "Test wysłany!"))
        except Exception as e: Clock.schedule_once(lambda x: self.msg("Błąd", str(e)))

    def import_contacts_db(self, p_path):
        from openpyxl import load_workbook
        wb = load_workbook(str(p_path), data_only=True); ws = wb.active; rows = list(ws.iter_rows(values_only=True))
        h = [str(x).lower() for x in rows[0]]
        ni, si = self.find_name_indices(h)
        mi = next((i for i, v in enumerate(h) if "mail" in v), 0)
        for r in rows[1:]:
            if r[mi]: self.conn.execute("INSERT OR REPLACE INTO contacts VALUES(?,?,?)", (str(r[ni]).lower().strip(), str(r[si]).lower().strip(), str(r[mi]).strip()))
        self.conn.commit(); self.msg("OK", "Baza kontaktów wczytana.")

    def update_att_txt(self):
        import os
        sz = sum(os.path.getsize(p) for p in self.global_attachments) / (1024*1024)
        self.size_lbl.text = f"Załączniki ogólne: {sz:.2f} MB"

    def pop_attachments(self, _):
        import os
        box = BoxLayout(orientation="vertical", padding=10)
        for p in self.global_attachments:
            r = BoxLayout(size_hint_y=None, height=dp(40))
            r.add_widget(Label(text=os.path.basename(p)[:20], font_size=10))
            btn = Button(text="USUŃ", size_hint_x=0.3, on_press=lambda x, path=p: self.remove_att(path))
            r.add_widget(btn); box.add_widget(r)
        box.add_widget(Button(text="DODAJ PLIK", size_hint_y=None, height=dp(50), on_press=lambda x: self.pick_file("extra")))
        b_cls = Button(text="ZAMKNIJ", size_hint_y=None, height=dp(50), on_press=lambda x: self.pop.dismiss())
        box.add_widget(b_cls); self.pop = Popup(title="Zarządzaj załącznikami", content=box, size_hint=(0.9, 0.8)); self.pop.open()

    def remove_att(self, path):
        if path in self.global_attachments: self.global_attachments.remove(path)
        self.pop.dismiss(); self.update_att_txt()

    def setup_smtp_ui(self):
        import json, os
        from pathlib import Path
        l = BoxLayout(orientation="vertical", padding=dp(25), spacing=dp(10))
        self.u_in = TextInput(hint_text="Twój Gmail", multiline=False)
        self.p_in = TextInput(hint_text="Hasło Aplikacji (16 znaków)", password=True, multiline=False)
        def save(_):
            with open(Path(self.user_data_dir) / "smtp.json", "w") as f: json.dump({'u':self.u_in.text,'p':self.p_in.text}, f)
            self.msg("OK", "Ustawienia zapisane."); setattr(self.sm, "current", "home")
        l.add_widget(Label(text="Konfiguracja Gmail", font_size=20)); l.add_widget(self.u_in); l.add_widget(self.p_in)
        l.add_widget(Button(text="ZAPISZ", on_press=save, background_color=(0,0.5,0,1)))
        l.add_widget(Button(text="POWRÓT", on_press=lambda x: setattr(self.sm, "current", "home")))
        p = Path(self.user_data_dir) / "smtp.json"
        if p.exists(): d = json.load(open(p)); self.u_in.text, self.p_in.text = d.get('u',''), d.get('p','')
        self.smtp_scr.add_widget(l)

    def setup_template_ui(self):
        l = BoxLayout(orientation="vertical", padding=dp(20), spacing=dp(10))
        self.ts = TextInput(hint_text="Temat maila", size_hint_y=None, height=dp(45))
        self.tb = TextInput(hint_text="Treść maila...")
        def save(_):
            self.conn.execute("UPDATE settings SET val=? WHERE key='t_sub'", (self.ts.text,))
            self.conn.execute("UPDATE settings SET val=? WHERE key='t_body'", (self.tb.text,)); self.conn.commit(); self.msg("OK", "Szablon gotowy.")
        l.add_widget(Label(text="Szablon (Tagi: {Imię}, {Nazwisko}, {Data})")); l.add_widget(self.ts); l.add_widget(self.tb)
        l.add_widget(Button(text="ZAPISZ", on_press=save)); l.add_widget(Button(text="POWRÓT", on_press=lambda x: setattr(self.sm, "current", "email")))
        rs = self.conn.execute("SELECT val FROM settings WHERE key='t_sub'").fetchone()
        rb = self.conn.execute("SELECT val FROM settings WHERE key='t_body'").fetchone()
        if rs: self.ts.text, self.tb.text = rs[0], rb[0]
        self.template_scr.add_widget(l)

    def setup_logs_ui(self):
        self.log_layout = BoxLayout(orientation="vertical", padding=5)
        self.log_grid = GridLayout(cols=1, size_hint_y=None, spacing=2); self.log_grid.bind(minimum_height=self.log_grid.setter('height'))
        sv = ScrollView(); sv.add_widget(self.log_grid)
        self.log_layout.add_widget(Label(text="Historia ostatnich wysyłek", size_hint_y=None, height=dp(40)))
        self.log_layout.add_widget(sv)
        self.log_layout.add_widget(Button(text="POWRÓT", size_hint_y=None, height=dp(45), on_press=lambda x: setattr(self.sm, "current", "email")))
        self.log_scr.add_widget(self.log_layout)

    def load_logs_into_ui(self, _):
        self.log_grid.clear_widgets()
        logs = self.conn.execute("SELECT recipient, status, date FROM logs ORDER BY id DESC LIMIT 50").fetchall()
        for r, s, d in logs:
            self.log_grid.add_widget(Label(text=f"{d} | {r} | {s}", font_size=dp(10), size_hint_y=None, height=dp(30)))
        self.sm.current = "logs"

    def msg(self, t, txt):
        Popup(title=t, content=Label(text=txt, halign="center"), size_hint=(0.8, 0.4)).open()

if __name__ == "__main__":
    FutureApp().run()
