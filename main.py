from kivy.app import App
from kivy.metrics import dp
from kivy.clock import Clock
from kivy.utils import platform
from kivy.uix.screenmanager import ScreenManager, Screen
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.button import Button
from kivy.uix.label import Label
from kivy.uix.popup import Popup
from kivy.uix.scrollview import ScrollView
from kivy.uix.gridlayout import GridLayout
from kivy.uix.textinput import TextInput
from kivy.uix.checkbox import CheckBox
from kivy.uix.progressbar import ProgressBar
import os

# Symbole tekstowe zamiast problematycznych kodów \ud...
ICO_FOLDER = "📁"
ICO_GEAR = "⚙️"
ICO_MAIL = "✉️"
ICO_SEND = "🚀"
ICO_TABLE = "📊"

class PremiumButton(Button):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.background_normal = ""
        self.background_color = (0.1, 0.4, 0.7, 1)
        self.color = (1, 1, 1, 1)
        self.size_hint_y = None
        self.height = dp(52)
        self.bold = True

class FutureApp(App):
    def build(self):
        from kivy.core.window import Window
        Window.clearcolor = (0.05, 0.08, 0.1, 1)
        
        self.full_data = [] 
        self.current_file = None
        self.global_attachments = []
        self.selected_col_indices = []
        
        self.init_db()
        self.sm = ScreenManager()
        
        # Inicjalizacja wszystkich ekranów
        self.screens = {
            "home": Screen(name="home"),
            "table": Screen(name="table"),
            "email": Screen(name="email"),
            "smtp": Screen(name="smtp"),
            "tmpl": Screen(name="tmpl"),
            "logs": Screen(name="logs")
        }

        self.setup_home()
        self.setup_table()
        self.setup_email_center()
        self.setup_smtp_ui()
        self.setup_tmpl_ui()
        self.setup_logs_ui()

        for s in self.screens.values(): self.sm.add_widget(s)
        
        if platform == 'android':
            from android.permissions import request_permissions, Permission
            request_permissions([Permission.READ_EXTERNAL_STORAGE, Permission.WRITE_EXTERNAL_STORAGE, Permission.INTERNET])

        return self.sm

    def init_db(self):
        import sqlite3
        from pathlib import Path
        db_p = Path(self.user_data_dir) / "app_v10_4.db"
        self.conn = sqlite3.connect(str(db_p), check_same_thread=False)
        self.conn.execute("CREATE TABLE IF NOT EXISTS contacts (name TEXT, surname TEXT, email TEXT, PRIMARY KEY(name, surname))")
        self.conn.execute("CREATE TABLE IF NOT EXISTS settings (key TEXT PRIMARY KEY, val TEXT)")
        self.conn.execute("CREATE TABLE IF NOT EXISTS logs (id INTEGER PRIMARY KEY, msg TEXT, date TEXT)")
        
        if not self.conn.execute("SELECT val FROM settings WHERE key='t_body'").fetchone():
            self.conn.execute("INSERT OR REPLACE INTO settings VALUES ('t_sub', 'Raport miesięczny: {Imię} {Nazwisko}')")
            self.conn.execute("INSERT OR REPLACE INTO settings VALUES ('t_body', 'Dzień dobry {Imię},\n\nPrzesyłamy raport za ostatni miesiąc.')")
        self.conn.commit()

    # --- EKRAN GŁÓWNY ---
    def setup_home(self):
        l = BoxLayout(orientation="vertical", padding=dp(30), spacing=dp(15))
        l.add_widget(Label(text="FUTURE HR 10.4", font_size=dp(30), bold=True))
        l.add_widget(PremiumButton(text=f"{ICO_FOLDER} WCZYTAJ DANE Z EXCELA", on_press=lambda x: self.pick_file("data")))
        l.add_widget(PremiumButton(text=f"{ICO_TABLE} ZOBACZ I FILTRUJ TABELĘ", on_press=self.go_to_table))
        l.add_widget(PremiumButton(text=f"{ICO_GEAR} KONFIGURACJA GMAIL", on_press=lambda x: setattr(self.sm, "current", "smtp")))
        self.h_stat = Label(text="Oczekiwanie na plik...", color=(0.7,0.7,0.7,1))
        l.add_widget(self.h_stat); self.screens["home"].add_widget(l)

    # --- TABELA I EKSPORT ---
    def setup_table(self):
        l = BoxLayout(orientation="vertical", padding=dp(5))
        top = BoxLayout(size_hint_y=None, height=dp(55), spacing=dp(5))
        self.search = TextInput(hint_text="Szukaj pracownika...", multiline=False); self.search.bind(text=self.filter_data)
        top.add_widget(self.search)
        top.add_widget(Button(text="KOLUMNY", size_hint_x=0.3, on_press=self.column_picker_popup))
        top.add_widget(Button(text="MAILING", size_hint_x=0.3, on_press=lambda x: setattr(self.sm, "current", "email")))
        
        self.grid = GridLayout(cols=1, size_hint_y=None, spacing=dp(2)); self.grid.bind(minimum_height=self.grid.setter('height'))
        sv = ScrollView(); sv.add_widget(self.grid)
        self.prog = ProgressBar(max=100, size_hint_y=None, height=dp(10))
        l.add_widget(top); l.add_widget(sv); l.add_widget(self.prog)
        l.add_widget(Button(text="EKSPORTUJ WSZYSTKIE PLIKI NA TELEFON", size_hint_y=None, height=dp(48), on_press=self.mass_export_files))
        l.add_widget(Button(text="POWRÓT", size_hint_y=None, height=dp(45), on_press=lambda x: setattr(self.sm, "current", "home")))
        self.screens["table"].add_widget(l)

    # --- CENTRUM MAILINGOWE ---
    def setup_email_center(self):
        le = BoxLayout(orientation="vertical", padding=dp(20), spacing=dp(8))
        self.att_info = Label(text="Dodatkowe załączniki: 0", size_hint_y=None, height=dp(30))
        btns = [
            (f"{ICO_FOLDER} WCZYTAJ BAZĘ E-MAIL", lambda x: self.pick_file("book")),
            (f"{ICO_MAIL} EDYTUJ TREŚĆ MAILA", lambda x: setattr(self.sm, "current", "tmpl")),
            (f"📎 DODAJ INNE PLIKI (PDF/FOTO)", self.manage_attachments_popup),
            (f"📜 HISTORIA WYSYŁKI", self.show_logs_popup),
            (f"⚡ TEST MAILA DO SIEBIE", self.run_mail_test),
            (f"{ICO_SEND} URUCHOM WYSYŁKĘ MASOWĄ", self.run_mass_mailing),
            ("POWRÓT", lambda x: setattr(self.sm, "current", "table"))
        ]
        le.add_widget(Label(text="Centrum Operacyjne", font_size=22, bold=True))
        le.add_widget(self.att_info)
        for t, c in btns: le.add_widget(PremiumButton(text=t, on_press=c))
        self.screens["email"].add_widget(le)

    # --- LOGIKA IMPORTU I EXCELA ---
    def go_to_table(self, _):
        from openpyxl import load_workbook
        if not self.current_file: self.msg("Błąd", "Najpierw wczytaj Excel!"); return
        try:
            wb = load_workbook(str(self.current_file), data_only=True); ws = wb.active
            self.full_data = [[("" if v is None else str(v)) for v in r] for r in ws.iter_rows(values_only=True)]
            self.update_table_grid(self.full_data); self.sm.current = "table"
        except Exception as e: self.msg("Błąd pliku", str(e))

    def mass_export_files(self, _):
        import threading
        if self.full_data: threading.Thread(target=self._export_thread).start()

    def _export_thread(self):
        from openpyxl import Workbook
        from pathlib import Path
        h, rows = self.full_data[0], self.full_data[1:]; ni, si = self.find_name_sur_idxs(h)
        folder = Path("/storage/emulated/0/Documents/FutureExport"); folder.mkdir(parents=True, exist_ok=True)
        idxs = self.selected_col_indices if self.selected_col_indices else list(range(len(h)))
        
        for i, r in enumerate(rows):
            wb = Workbook(); ws = wb.active
            ws.append([h[k] for k in idxs]); ws.append([r[k] for k in idxs])
            fname = f"Raport_{r[ni]}_{r[si]}".replace(" ", "_")
            wb.save(str(folder / f"{fname}.xlsx"))
            Clock.schedule_once(lambda dt, p=int(((i+1)/len(rows))*100): setattr(self.prog, "value", p))
        Clock.schedule_once(lambda x: self.msg("Sukces", "Zapisano w Documents/FutureExport"))

    # --- LOGIKA MAILINGU ---
    def run_mass_mailing(self, _):
        import threading
        if self.full_data: threading.Thread(target=self._mailing_thread).start()

    def _mailing_thread(self):
        import smtplib, json
        from pathlib import Path
        from datetime import datetime
        
        p = Path(self.user_data_dir) / "smtp.json"
        if not p.exists(): Clock.schedule_once(lambda x: self.msg("Blad", "Ustaw SMTP!")); return
        with open(p, "r") as f: cfg = json.load(f)
        
        try:
            srv = smtplib.SMTP("smtp.gmail.com", 587, timeout=12); srv.starttls(); srv.login(cfg['u'], cfg['p'])
        except Exception as e: Clock.schedule_once(lambda x: self.msg("Błąd logowania", str(e))); return

        h, rows = self.full_data[0], self.full_data[1:]; ni, si = self.find_name_sur_idxs(h)
        cnt = 0
        for i, r in enumerate(rows):
            n, s = str(r[ni]).strip(), str(r[si]).strip()
            res = self.conn.execute("SELECT email FROM contacts WHERE name=? AND surname=?", (n.lower(), s.lower())).fetchone()
            if res:
                try:
                    msg = self.create_mail(cfg['u'], res[0], n, s, h, r)
                    srv.send_message(msg); cnt += 1
                    self.conn.execute("INSERT INTO logs (msg, date) VALUES (?,?)", (f"Wysłano do: {res[0]}", datetime.now().strftime("%H:%M")))
                except: pass
            Clock.schedule_once(lambda dt, p=int(((i+1)/len(rows))*100): setattr(self.prog, "value", p))
        srv.quit(); self.conn.commit()
        Clock.schedule_once(lambda x: self.msg("Koniec", f"Wysłano {cnt} maili."))

    def create_mail(self, sender, to, n, s, h, r):
        from email.message import EmailMessage
        from openpyxl import Workbook
        from pathlib import Path
        import mimetypes
        from datetime import datetime

        msg = EmailMessage(); dat = datetime.now().strftime("%d.%m.%Y")
        sub_raw = self.conn.execute("SELECT val FROM settings WHERE key='t_sub'").fetchone()[0]
        bod_raw = self.conn.execute("SELECT val FROM settings WHERE key='t_body'").fetchone()[0]
        
        msg["Subject"] = sub_raw.replace("{Imię}", n).replace("{Nazwisko}", s).replace("{Data}", dat)
        msg["From"], msg["To"] = sender, to
        msg.set_content(bod_raw.replace("{Imię}", n).replace("{Nazwisko}", s).replace("{Data}", dat))
        
        # Załącznik Excel
        idxs = self.selected_col_indices if self.selected_col_indices else list(range(len(h)))
        tmp = Path(self.user_data_dir) / "mail_tmp.xlsx"; wb = Workbook(); ws = wb.active
        ws.append([h[k] for k in idxs]); ws.append([r[k] for k in idxs]); wb.save(str(tmp))
        with open(tmp, "rb") as f: msg.add_attachment(f.read(), maintype="application", subtype="xlsx", filename=f"Raport_{n}.xlsx")
        
        # Załączniki dodatkowe
        for ap in self.global_attachments:
            if os.path.exists(ap):
                ctype, _ = mimetypes.guess_type(ap)
                main, sub = (ctype or "application/octet-stream").split("/", 1)
                with open(ap, "rb") as f: msg.add_attachment(f.read(), maintype=main, subtype=sub, filename=os.path.basename(ap))
        return msg

    # --- POMOCNICZE / UI ---
    def pick_file(self, mode):
        if platform != 'android': return
        from jnius import autoclass; from android import activity
        from pathlib import Path
        Intent = autoclass("android.content.Intent")
        intent = Intent(Intent.ACTION_OPEN_DOCUMENT); intent.setType("*/*"); intent.addCategory(Intent.CATEGORY_OPENABLE)
        def on_res(req, res, dt):
            if dt:
                try:
                    uri = dt.getData(); ctx = autoclass("org.kivy.android.PythonActivity").mActivity
                    stream = ctx.getContentResolver().openInputStream(uri)
                    dest = Path(self.user_data_dir) / (f"extra_{os.urandom(4).hex()}.tmp" if mode=="extra" else f"{mode}.xlsx")
                    with open(dest, "wb") as f:
                        buf = autoclass('[B')(16384)
                        while True:
                            r = stream.read(buf)
                            if r <= 0: break
                            f.write(bytes(buf)[:r])
                    if mode == "data": self.current_file = dest; Clock.schedule_once(lambda x: setattr(self.h_stat, "text", "Załadowano GŁÓWNY."))
                    elif mode == "book": self.import_contacts_logic(dest)
                    elif mode == "extra": self.global_attachments.append(str(dest)); self.update_att_info()
                except: pass
            activity.unbind(on_activity_result=on_res)
        activity.bind(on_activity_result=on_res); autoclass("org.kivy.android.PythonActivity").mActivity.startActivityForResult(intent, 1001)

    def update_table_grid(self, data):
        self.grid.clear_widgets()
        for r_data in data[:60]:
            row = BoxLayout(size_hint_y=None, height=dp(32))
            for cell in r_data[:3]: row.add_widget(Label(text=str(cell)[:15], font_size=11))
            self.grid.add_widget(row)

    def find_name_sur_idxs(self, h):
        ni, si = 0, 1
        for i, v in enumerate(h):
            if "imi" in str(v).lower(): ni = i
            if "nazw" in str(v).lower(): si = i
        return ni, si

    def import_contacts_logic(self, p):
        from openpyxl import load_workbook
        wb = load_workbook(str(p), data_only=True); ws = wb.active; rows = list(ws.iter_rows(values_only=True))
        h = [str(x).lower() for x in rows[0]]; ni, si = self.find_name_sur_idxs(h)
        mi = next((i for i, v in enumerate(h) if "mail" in v), 2)
        for r in rows[1:]:
            if r[mi]: self.conn.execute("INSERT OR REPLACE INTO contacts VALUES(?,?,?)", (str(r[ni]).lower().strip(), str(r[si]).lower().strip(), str(r[mi]).strip()))
        self.conn.commit(); self.msg("OK", "Baza e-mail wczytana.")

    def update_att_info(self): self.att_info.text = f"Dodatkowe załączniki: {len(self.global_attachments)}"

    def column_picker_popup(self, _):
        if not self.full_data: return
        box = BoxLayout(orientation="vertical", padding=10)
        grid = GridLayout(cols=1, size_hint_y=None, spacing=5); grid.bind(minimum_height=grid.setter('height'))
        chks = []
        for i, h in enumerate(self.full_data[0]):
            r = BoxLayout(size_hint_y=None, height=dp(40)); cb = CheckBox(size_hint_x=0.2); cb.active = True
            r.add_widget(cb); r.add_widget(Label(text=str(h))); grid.add_widget(r); chks.append((i, cb))
        def save(_): self.selected_col_indices = [idx for idx, c in chks if c.active]; p.dismiss()
        sv = ScrollView(); sv.add_widget(grid); box.add_widget(sv)
        box.add_widget(Button(text="ZAPISZ WYBÓR", size_hint_y=None, height=dp(50), on_press=save))
        p = Popup(title="Wybierz kolumny do raportu", content=box, size_hint=(0.9, 0.8)); p.open()

    def manage_attachments_popup(self, _):
        box = BoxLayout(orientation="vertical", padding=10, spacing=10)
        for path in self.global_attachments:
            r = BoxLayout(size_hint_y=None, height=dp(40))
            r.add_widget(Label(text=os.path.basename(path)[:20]))
            btn = Button(text="USUŃ", size_hint_x=0.3, on_press=lambda x, p=path: self.remove_att(p))
            r.add_widget(btn); box.add_widget(r)
        box.add_widget(Button(text="DODAJ PLIK", on_press=lambda x: self.pick_file("extra")))
        b_cls = Button(text="ZAMKNIJ", on_press=lambda x: self.at_p.dismiss())
        box.add_widget(b_cls); self.at_p = Popup(title="Załączniki", content=box, size_hint=(0.8, 0.6)); self.at_p.open()

    def remove_att(self, p):
        if p in self.global_attachments: self.global_attachments.remove(p)
        self.at_p.dismiss(); self.update_att_info()

    def show_logs_popup(self, _):
        box = BoxLayout(orientation="vertical", padding=10)
        logs = self.conn.execute("SELECT msg, date FROM logs ORDER BY id DESC LIMIT 30").fetchall()
        l_grid = GridLayout(cols=1, size_hint_y=None); l_grid.bind(minimum_height=l_grid.setter('height'))
        for m, d in logs: l_grid.add_widget(Label(text=f"{d}: {m}", font_size=10, size_hint_y=None, height=dp(25)))
        sv = ScrollView(); sv.add_widget(l_grid); box.add_widget(sv)
        box.add_widget(Button(text="ZAMKNIJ", size_hint_y=None, height=dp(45), on_press=lambda x: p.dismiss()))
        p = Popup(title="Logi wysyłki", content=box, size_hint=(0.9, 0.8)); p.open()

    def setup_smtp_ui(self):
        import json; from pathlib import Path
        l = BoxLayout(orientation="vertical", padding=dp(25), spacing=dp(10))
        self.u_in = TextInput(hint_text="Gmail", multiline=False); self.p_in = TextInput(hint_text="Hasło Aplikacji", password=True, multiline=False)
        def save(_):
            with open(Path(self.user_data_dir) / "smtp.json", "w") as f: json.dump({'u':self.u_in.text,'p':self.p_in.text}, f)
            setattr(self.sm, "current", "home")
        l.add_widget(Label(text="Ustawienia Gmail", font_size=20)); l.add_widget(self.u_in); l.add_widget(self.p_in)
        l.add_widget(Button(text="ZAPISZ", on_press=save)); l.add_widget(Button(text="WRÓĆ", on_press=lambda x: setattr(self.sm, "current", "home")))
        p = Path(self.user_data_dir) / "smtp.json"
        if p.exists(): d = json.load(open(p)); self.u_in.text, self.p_in.text = d.get('u',''), d.get('p','')
        self.screens["smtp"].add_widget(l)

    def setup_tmpl_ui(self):
        l = BoxLayout(orientation="vertical", padding=dp(20), spacing=dp(10))
        self.t_sub = TextInput(size_hint_y=None, height=dp(45)); self.t_bod = TextInput()
        def save(_):
            self.conn.execute("UPDATE settings SET val=? WHERE key='t_sub'", (self.t_sub.text,))
            self.conn.execute("UPDATE settings SET val=? WHERE key='t_body'", (self.t_bod.text,)); self.conn.commit(); self.msg("OK", "Zapisano szablon.")
        l.add_widget(Label(text="Tagi: {Imię}, {Nazwisko}, {Data}")); l.add_widget(self.t_sub); l.add_widget(self.t_bod)
        l.add_widget(Button(text="ZAPISZ", on_press=save)); l.add_widget(Button(text="POWRÓT", on_press=lambda x: setattr(self.sm, "current", "email")))
        rs = self.conn.execute("SELECT val FROM settings WHERE key='t_sub'").fetchone()
        rb = self.conn.execute("SELECT val FROM settings WHERE key='t_body'").fetchone()
        if rs: self.t_sub.text = rs[0]; self.t_bod.text = rb[0]
        self.screens["tmpl"].add_widget(l)

    def run_mail_test(self, _):
        import threading
        if self.full_data: threading.Thread(target=self._test_thread).start()

    def _test_thread(self):
        try:
            import json; from pathlib import Path; import smtplib
            cfg = json.load(open(Path(self.user_data_dir) / "smtp.json"))
            srv = smtplib.SMTP("smtp.gmail.com", 587, timeout=10); srv.starttls(); srv.login(cfg['u'], cfg['p'])
            msg = self.create_mail(cfg['u'], cfg['u'], "TEST", "URUCHOMIONY", self.full_data[0], self.full_data[1])
            msg["Subject"] = "[TEST] " + msg["Subject"]; srv.send_message(msg); srv.quit()
            Clock.schedule_once(lambda x: self.msg("OK", "Test wysłany!"))
        except Exception as e: Clock.schedule_once(lambda x: self.msg("Blad", str(e)))

    def filter_data(self, ins, val):
        if not self.full_data: return
        f = [self.full_data[0]] + [r for r in self.full_data[1:] if val.lower() in str(r).lower()]
        self.update_table_grid(f)

    def setup_logs_ui(self): pass
    def msg(self, t, txt): Popup(title=t, content=Label(text=txt, halign="center"), size_hint=(0.8, 0.4)).open()

if __name__ == "__main__":
    try:
        FutureApp().run()
    except Exception as e:
        import traceback
        with open("crash_log.txt", "w") as f: f.write(traceback.format_exc())
