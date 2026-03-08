import os
import json
import sqlite3
import threading
import mimetypes
from datetime import datetime
from pathlib import Path

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

# Używamy bezpiecznych symboli tekstowych (unikamy błędów Unicode)
ICO_DATA = "📂"
ICO_TABLE = "📊"
ICO_GEAR = "⚙️"
ICO_MAIL = "✉️"
ICO_SEND = "🚀"
ICO_CLIP = "📎"

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
        self.export_columns = []
        
        self.init_db()
        self.sm = ScreenManager()
        
        # Definicja ekranów
        self.home_scr = Screen(name="home")
        self.table_scr = Screen(name="table")
        self.email_scr = Screen(name="email")
        self.smtp_scr = Screen(name="smtp")
        self.tmpl_scr = Screen(name="tmpl")

        self.setup_ui()

        self.sm.add_widget(self.home_scr)
        self.sm.add_widget(self.table_scr)
        self.sm.add_widget(self.email_scr)
        self.sm.add_widget(self.smtp_scr)
        self.sm.add_widget(self.tmpl_scr)
        
        if platform == 'android':
            from android.permissions import request_permissions, Permission
            request_permissions([Permission.READ_EXTERNAL_STORAGE, Permission.WRITE_EXTERNAL_STORAGE, Permission.INTERNET])

        return self.sm

    def init_db(self):
        db_p = Path(self.user_data_dir) / "future_v10.db"
        self.conn = sqlite3.connect(str(db_p), check_same_thread=False)
        self.conn.execute("CREATE TABLE IF NOT EXISTS contacts (name TEXT, surname TEXT, email TEXT, PRIMARY KEY(name, surname))")
        self.conn.execute("CREATE TABLE IF NOT EXISTS settings (key TEXT PRIMARY KEY, val TEXT)")
        self.conn.execute("CREATE TABLE IF NOT EXISTS logs (id INTEGER PRIMARY KEY, msg TEXT, date TEXT)")
        
        if not self.conn.execute("SELECT val FROM settings WHERE key='t_body'").fetchone():
            self.conn.execute("INSERT INTO settings VALUES ('t_sub', 'Raport: {Imię} {Nazwisko}')")
            self.conn.execute("INSERT INTO settings VALUES ('t_body', 'Witaj {Imię},\n\nPrzesyłamy raport z dnia {Data}.')")
        self.conn.commit()

    def setup_ui(self):
        # --- HOME ---
        l_home = BoxLayout(orientation="vertical", padding=dp(25), spacing=dp(15))
        l_home.add_widget(Label(text="FUTURE 10.5 PRO", font_size=dp(28), bold=True))
        l_home.add_widget(PremiumButton(text=f"{ICO_DATA} WCZYTAJ EXCEL PŁAC", on_press=lambda x: self.pick_file("data")))
        l_home.add_widget(PremiumButton(text=f"{ICO_TABLE} OTWÓRZ TABELĘ", on_press=self.go_to_table))
        l_home.add_widget(PremiumButton(text=f"{ICO_GEAR} USTAWIENIA GMAIL", on_press=lambda x: setattr(self.sm, "current", "smtp")))
        self.h_stat = Label(text="Wybierz plik .xlsx", color=(0.6, 0.6, 0.6, 1))
        l_home.add_widget(self.h_stat); self.home_scr.add_widget(l_home)

        # --- TABLE ---
        l_tab = BoxLayout(orientation="vertical", padding=dp(5))
        top = BoxLayout(size_hint_y=None, height=dp(55), spacing=dp(5))
        self.search = TextInput(hint_text="Szukaj...", multiline=False); self.search.bind(text=self.filter_data)
        top.add_widget(self.search)
        top.add_widget(Button(text="KOLUMNY", size_hint_x=0.3, on_press=self.column_popup))
        top.add_widget(Button(text="MAILING", size_hint_x=0.3, on_press=lambda x: setattr(self.sm, "current", "email")))
        self.grid = GridLayout(cols=1, size_hint_y=None, spacing=dp(2)); self.grid.bind(minimum_height=self.grid.setter('height'))
        sv = ScrollView(); sv.add_widget(self.grid)
        self.prog = ProgressBar(max=100, size_hint_y=None, height=dp(10))
        l_tab.add_widget(top); l_tab.add_widget(sv); l_tab.add_widget(self.prog)
        l_tab.add_widget(Button(text="EKSPORTUJ WSZYSTKO (DOCUMENTS)", size_hint_y=None, height=dp(48), on_press=self.mass_export))
        l_tab.add_widget(Button(text="POWRÓT", size_hint_y=None, height=dp(45), on_press=lambda x: setattr(self.sm, "current", "home")))
        self.table_scr.add_widget(l_tab)

        # --- EMAIL CENTER ---
        le = BoxLayout(orientation="vertical", padding=dp(20), spacing=dp(10))
        le.add_widget(Label(text="CENTRUM MAILINGOWE", font_size=dp(22), bold=True))
        self.att_lbl = Label(text="Dodatkowe załączniki: 0", size_hint_y=None, height=dp(30))
        le.add_widget(self.att_lbl)
        btns = [
            (f"{ICO_DATA} WCZYTAJ BAZĘ E-MAIL", lambda x: self.pick_file("book")),
            (f"{ICO_MAIL} EDYTUJ TREŚĆ", lambda x: setattr(self.sm, "current", "tmpl")),
            (f"{ICO_CLIP} DODAJ PDF/ZAŁĄCZNIK", self.att_manager_popup),
            ("📜 LOGI WYSYŁKI", self.show_logs),
            (f"{ICO_SEND} WYŚLIJ DO WSZYSTKICH", self.start_mailing),
            ("POWRÓT", lambda x: setattr(self.sm, "current", "table"))
        ]
        for t, c in btns: le.add_widget(PremiumButton(text=t, on_press=c))
        self.email_scr.add_widget(le)

        # --- SMTP ---
        ls = BoxLayout(orientation="vertical", padding=dp(25), spacing=dp(10))
        self.u_in = TextInput(hint_text="Twój Gmail", multiline=False)
        self.p_in = TextInput(hint_text="Hasło Aplikacji", password=True, multiline=False)
        def save_smtp(_):
            with open(Path(self.user_data_dir) / "smtp.json", "w") as f: json.dump({'u':self.u_in.text,'p':self.p_in.text}, f)
            self.msg("OK", "Zapisano."); setattr(self.sm, "current", "home")
        ls.add_widget(Label(text="Konfiguracja Gmail SMTP")); ls.add_widget(self.u_in); ls.add_widget(self.p_in)
        ls.add_widget(Button(text="ZAPISZ", on_press=save_smtp))
        ls.add_widget(Button(text="WRÓĆ", on_press=lambda x: setattr(self.sm, "current", "home")))
        self.smtp_scr.add_widget(ls); self.load_smtp_info()

        # --- TEMPLATE ---
        lt = BoxLayout(orientation="vertical", padding=dp(20), spacing=dp(10))
        self.t_sub = TextInput(size_hint_y=None, height=dp(45)); self.t_bod = TextInput()
        def save_tmpl(_):
            self.conn.execute("UPDATE settings SET val=? WHERE key='t_sub'", (self.t_sub.text,))
            self.conn.execute("UPDATE settings SET val=? WHERE key='t_body'", (self.t_bod.text,)); self.conn.commit()
            self.msg("OK", "Zapisano szablon.")
        lt.add_widget(Label(text="Tagi: {Imię}, {Nazwisko}, {Data}")); lt.add_widget(self.t_sub); lt.add_widget(self.t_bod)
        lt.add_widget(Button(text="ZAPISZ", on_press=save_tmpl))
        lt.add_widget(Button(text="WRÓĆ", on_press=lambda x: setattr(self.sm, "current", "email")))
        self.tmpl_scr.add_widget(lt); self.load_tmpl_info()

    # --- POPRAWIONY PICKER (ROZWIĄZUJE PROBLEM WCZYTYWANIA) ---
    def pick_file(self, mode):
        if platform != 'android': self.msg("Błąd", "Dostępne tylko na Androidzie"); return
        from jnius import autoclass
        from android import activity
        Intent = autoclass("android.content.Intent")
        intent = Intent(Intent.ACTION_OPEN_DOCUMENT); intent.setType("*/*"); intent.addCategory(Intent.CATEGORY_OPENABLE)

        def on_res(req, res, dt):
            if dt:
                try:
                    uri = dt.getData(); ctx = autoclass("org.kivy.android.PythonActivity").mActivity
                    stream = ctx.getContentResolver().openInputStream(uri)
                    dest = Path(self.user_data_dir) / (f"extra_{os.urandom(4).hex()}" if mode == "extra" else f"{mode}_local.xlsx")
                    with open(dest, "wb") as f:
                        j_buf = autoclass('[B')(16384)
                        while True:
                            r = stream.read(j_buf)
                            if r <= 0: break
                            f.write(bytes(j_buf)[:r])
                    stream.close()
                    if mode == "data": self.current_file = dest; Clock.schedule_once(lambda x: setattr(self.h_stat, "text", "Załadowano. Otwórz tabelę."))
                    elif mode == "book": self.import_book(dest)
                    elif mode == "extra": self.global_attachments.append(str(dest)); self.update_att_lbl()
                except Exception as e: self.msg("Błąd pliku", str(e))
            activity.unbind(on_activity_result=on_res)
        activity.bind(on_activity_result=on_res); ctx = autoclass("org.kivy.android.PythonActivity").mActivity; ctx.startActivityForResult(intent, 1001)

    def import_book(self, p):
        from openpyxl import load_workbook
        try:
            wb = load_workbook(str(p), data_only=True); ws = wb.active; rows = list(ws.iter_rows(values_only=True))
            h = [str(x).lower() for x in rows[0]]; ni, si = self.get_name_idxs(h)
            mi = next((i for i, v in enumerate(h) if "mail" in v), 2)
            for r in rows[1:]:
                if r[mi]: self.conn.execute("INSERT OR REPLACE INTO contacts VALUES(?,?,?)", (str(r[ni]).lower().strip(), str(r[si]).lower().strip(), str(r[mi]).strip()))
            self.conn.commit(); self.msg("Sukces", "Baza e-mail gotowa!")
        except Exception as e: self.msg("Błąd Excel", str(e))

    def go_to_table(self, _):
        from openpyxl import load_workbook
        if not self.current_file or not os.path.exists(str(self.current_file)): self.msg("Błąd", "Wczytaj plik!"); return
        try:
            wb = load_workbook(str(self.current_file), data_only=True); ws = wb.active
            self.full_data = [[("" if v is None else str(v)) for v in r] for r in ws.iter_rows(values_only=True)]
            self.draw_table(self.full_data); self.sm.current = "table"
        except Exception as e: self.msg("Błąd", str(e))

    def draw_table(self, data):
        self.grid.clear_widgets()
        for r in data[:50]:
            row = BoxLayout(size_hint_y=None, height=dp(32))
            for c in r[:3]: row.add_widget(Label(text=str(c)[:15], font_size=11))
            self.grid.add_widget(row)

    # --- MAILING I EKSPORT ---
    def start_mailing(self, _):
        if self.full_data: threading.Thread(target=self._mail_task).start()

    def _mail_task(self):
        import smtplib, mimetypes
        from email.message import EmailMessage
        from openpyxl import Workbook
        
        cfg_p = Path(self.user_data_dir) / "smtp.json"
        if not cfg_p.exists(): Clock.schedule_once(lambda x: self.msg("!", "Ustaw Gmail!")); return
        cfg = json.load(open(cfg_p))
        
        try:
            srv = smtplib.SMTP("smtp.gmail.com", 587, timeout=12); srv.starttls(); srv.login(cfg['u'], cfg['p'])
        except Exception as e: Clock.schedule_once(lambda x: self.msg("Błąd login", str(e))); return

        h, rows = self.full_data[0], self.full_data[1:]; ni, si = self.get_name_idxs(h)
        sub = self.t_sub.text; bod = self.t_bod.text; sent = 0
        
        for i, r in enumerate(rows):
            n, s = str(r[ni]).strip(), str(r[si]).strip()
            res = self.conn.execute("SELECT email FROM contacts WHERE name=? AND surname=?", (n.lower(), s.lower())).fetchone()
            if res:
                msg = EmailMessage(); dat = datetime.now().strftime("%d.%m.%Y")
                msg["Subject"] = sub.replace("{Imię}", n).replace("{Nazwisko}", s).replace("{Data}", dat)
                msg["From"], msg["To"] = cfg['u'], res[0]
                msg.set_content(bod.replace("{Imię}", n).replace("{Nazwisko}", s).replace("{Data}", dat))
                
                # Załącznik personalny
                idxs = self.export_columns if self.export_columns else list(range(len(h)))
                tmp_p = Path(self.user_data_dir) / "temp.xlsx"; wb = Workbook(); ws = wb.active
                ws.append([h[k] for k in idxs]); ws.append([r[k] for k in idxs]); wb.save(str(tmp_p))
                with open(tmp_p, "rb") as f: msg.add_attachment(f.read(), maintype="application", subtype="xlsx", filename=f"Raport_{n}.xlsx")
                
                # Załączniki dodatkowe
                for ap in self.global_attachments:
                    if os.path.exists(ap):
                        ct, _ = mimetypes.guess_type(ap); m, s_t = (ct or "application/octet-stream").split("/", 1)
                        with open(ap, "rb") as f: msg.add_attachment(f.read(), maintype=m, subtype=s_t, filename=os.path.basename(ap))
                
                try: 
                    srv.send_message(msg); sent += 1
                    self.conn.execute("INSERT INTO logs (msg, date) VALUES (?,?)", (f"Wysłano: {res[0]}", dat))
                except: pass
            Clock.schedule_once(lambda dt, p=int(((i+1)/len(rows))*100): setattr(self.prog, "value", p))
        
        srv.quit(); self.conn.commit()
        Clock.schedule_once(lambda x: self.msg("Koniec", f"Wysłano {sent} maili."))

    def mass_export(self, _):
        if self.full_data: threading.Thread(target=self._export_task).start()

    def _export_task(self):
        from openpyxl import Workbook
        h, rows = self.full_data[0], self.full_data[1:]; ni, si = self.get_name_idxs(h)
        folder = Path("/storage/emulated/0/Documents/FutureExport"); folder.mkdir(parents=True, exist_ok=True)
        idxs = self.export_columns if self.export_columns else list(range(len(h)))
        
        for i, r in enumerate(rows):
            wb = Workbook(); ws = wb.active; ws.append([h[k] for k in idxs]); ws.append([r[k] for k in idxs])
            wb.save(str(folder / f"Raport_{r[ni]}_{r[si]}.xlsx"))
            Clock.schedule_once(lambda dt, p=int(((i+1)/len(rows))*100): setattr(self.prog, "value", p))
        Clock.schedule_once(lambda x: self.msg("OK", "Zapisano w Documents/FutureExport"))

    # --- POMOCNICZE ---
    def get_name_idxs(self, h):
        ni, si = 0, 1
        for i, v in enumerate(h):
            if "imi" in str(v).lower(): ni = i
            if "nazw" in str(v).lower(): si = i
        return ni, si

    def update_att_lbl(self): self.att_lbl.text = f"Dodatkowe załączniki: {len(self.global_attachments)}"

    def column_popup(self, _):
        if not self.full_data: return
        box = BoxLayout(orientation="vertical", padding=10); grid = GridLayout(cols=1, size_hint_y=None); grid.bind(minimum_height=grid.setter('height'))
        chks = []
        for i, h in enumerate(self.full_data[0]):
            r = BoxLayout(size_hint_y=None, height=dp(40)); cb = CheckBox(size_hint_x=0.2); cb.active = True
            r.add_widget(cb); r.add_widget(Label(text=str(h))); grid.add_widget(r); chks.append((i, cb))
        def save(_): self.export_columns = [idx for idx, c in chks if c.active]; p.dismiss()
        sv = ScrollView(); sv.add_widget(grid); box.add_widget(sv); box.add_widget(Button(text="ZAPISZ", size_hint_y=None, height=dp(50), on_press=save))
        p = Popup(title="Wybierz kolumny", content=box, size_hint=(0.9, 0.8)); p.open()

    def att_manager_popup(self, _):
        box = BoxLayout(orientation="vertical", padding=10, spacing=10)
        for ap in self.global_attachments:
            r = BoxLayout(size_hint_y=None, height=dp(40))
            r.add_widget(Label(text=os.path.basename(ap)[:15])); btn = Button(text="USUŃ", size_hint_x=0.3, on_press=lambda x, p=ap: self.remove_att(p))
            r.add_widget(btn); box.add_widget(r)
        box.add_widget(Button(text="DODAJ PLIK", on_press=lambda x: self.pick_file("extra")))
        box.add_widget(Button(text="ZAMKNIJ", on_press=lambda x: self.at_p.dismiss()))
        self.at_p = Popup(title="Załączniki", content=box, size_hint=(0.8, 0.6)); self.at_p.open()

    def remove_att(self, p):
        if p in self.global_attachments: self.global_attachments.remove(p)
        self.at_p.dismiss(); self.update_att_lbl()

    def show_logs(self, _):
        box = BoxLayout(orientation="vertical", padding=10)
        logs = self.conn.execute("SELECT msg, date FROM logs ORDER BY id DESC LIMIT 20").fetchall()
        l_grid = GridLayout(cols=1, size_hint_y=None); l_grid.bind(minimum_height=l_grid.setter('height'))
        for m, d in logs: l_grid.add_widget(Label(text=f"{d}: {m}", font_size=11, size_hint_y=None, height=dp(25)))
        sv = ScrollView(); sv.add_widget(l_grid); box.add_widget(sv); box.add_widget(Button(text="OK", size_hint_y=None, height=dp(45), on_press=lambda x: p.dismiss()))
        p = Popup(title="Historia", content=box, size_hint=(0.9, 0.8)); p.open()

    def filter_data(self, ins, val):
        if self.full_data: self.draw_table([self.full_data[0]] + [r for r in self.full_data[1:] if val.lower() in str(r).lower()])

    def load_smtp_info(self):
        p = Path(self.user_data_dir) / "smtp.json"
        if p.exists(): d = json.load(open(p)); self.u_in.text, self.p_in.text = d.get('u',''), d.get('p','')

    def load_tmpl_info(self):
        rs = self.conn.execute("SELECT val FROM settings WHERE key='t_sub'").fetchone()
        rb = self.conn.execute("SELECT val FROM settings WHERE key='t_body'").fetchone()
        if rs: self.t_sub.text, self.t_bod.text = rs[0], rb[0]

    def msg(self, t, txt): Popup(title=t, content=Label(text=txt, halign="center"), size_hint=(0.8, 0.4)).open()

if __name__ == "__main__":
    try:
        FutureApp().run()
    except Exception as e:
        import traceback
        # Zapis błędu do bezpiecznej lokalizacji wewnętrznej
        with open(os.path.join(os.getcwd(), "crash_report.txt"), "w") as f:
            f.write(traceback.format_exc())
