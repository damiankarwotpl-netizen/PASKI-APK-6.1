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

# SYMBOLE (Bezpieczne dla Androida - eliminują błąd Unicode)
ICO_DATA = "📂"
ICO_TABLE = "📊"
ICO_GEAR = "⚙"
ICO_MAIL = "✉"
ICO_SEND = "🚀"
ICO_CLIP = "📎"
ICO_TEST = "⚡"

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
        
        # Inicjalizacja stanów
        self.full_data = [] 
        self.current_file = None
        self.global_attachments = []
        self.export_columns = []
        
        self.init_db()
        self.sm = ScreenManager()
        
        # Definicja 5 ekranów
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
        db_p = Path(self.user_data_dir) / "future_final_v9.db"
        self.conn = sqlite3.connect(str(db_p), check_same_thread=False)
        self.conn.execute("CREATE TABLE IF NOT EXISTS contacts (name TEXT, surname TEXT, email TEXT, PRIMARY KEY(name, surname))")
        self.conn.execute("CREATE TABLE IF NOT EXISTS settings (key TEXT PRIMARY KEY, val TEXT)")
        self.conn.execute("CREATE TABLE IF NOT EXISTS logs (id INTEGER PRIMARY KEY, msg TEXT, date TEXT)")
        
        if not self.conn.execute("SELECT val FROM settings WHERE key='t_body'").fetchone():
            self.conn.execute("INSERT OR REPLACE INTO settings VALUES ('t_sub', 'Raport: {Imię} {Nazwisko}')")
            self.conn.execute("INSERT OR REPLACE INTO settings VALUES ('t_body', 'Witaj {Imię},\n\nPrzesyłamy raport miesięczny z dnia {Data}.')")
        self.conn.commit()

    def setup_ui(self):
        # --- EKRAN STARTOWY (Umożliwia wejście wszędzie) ---
        l_home = BoxLayout(orientation="vertical", padding=dp(25), spacing=dp(12))
        l_home.add_widget(Label(text="FUTURE 9.0 ULTIMATE", font_size=dp(28), bold=True))
        l_home.add_widget(PremiumButton(text=f"{ICO_DATA} WCZYTAJ EXCEL PŁAC", on_press=lambda x: self.pick_file("data")))
        l_home.add_widget(PremiumButton(text=f"{ICO_TABLE} PODGLĄD TABELI", on_press=self.go_to_table))
        l_home.add_widget(PremiumButton(text=f"{ICO_SEND} CENTRUM MAILINGOWE", on_press=lambda x: setattr(self.sm, "current", "email")))
        l_home.add_widget(PremiumButton(text=f"{ICO_GEAR} USTAWIENIA GMAIL", on_press=lambda x: setattr(self.sm, "current", "smtp")))
        self.h_stat = Label(text="Aplikacja gotowa", color=(0.6, 0.6, 0.6, 1))
        l_home.add_widget(self.h_stat); self.home_scr.add_widget(l_home)

        # --- EKRAN TABELI ---
        l_tab = BoxLayout(orientation="vertical", padding=dp(5))
        top = BoxLayout(size_hint_y=None, height=dp(55), spacing=dp(5))
        self.search = TextInput(hint_text="Szukaj osoby...", multiline=False); self.search.bind(text=self.filter_data)
        top.add_widget(self.search)
        top.add_widget(Button(text="KOLUMNY", size_hint_x=0.3, on_press=self.column_popup))
        self.grid = GridLayout(cols=1, size_hint_y=None, spacing=dp(2)); self.grid.bind(minimum_height=self.grid.setter('height'))
        sv = ScrollView(); sv.add_widget(self.grid)
        self.prog = ProgressBar(max=100, size_hint_y=None, height=dp(10))
        l_tab.add_widget(top); l_tab.add_widget(sv); l_tab.add_widget(self.prog)
        l_tab.add_widget(Button(text="EKSPORTUJ WSZYSTKO (TELEFON)", size_hint_y=None, height=dp(48), on_press=self.mass_export))
        l_tab.add_widget(Button(text="POWRÓT", size_hint_y=None, height=dp(45), on_press=lambda x: setattr(self.sm, "current", "home")))
        self.table_scr.add_widget(l_tab)

        # --- EKRAN MAILINGU ---
        le = BoxLayout(orientation="vertical", padding=dp(20), spacing=dp(8))
        le.add_widget(Label(text="MENU OPERACYJNE", font_size=dp(22), bold=True))
        self.att_lbl = Label(text="Załączniki dodatkowe: 0", size_hint_y=None, height=dp(30))
        le.add_widget(self.att_lbl)
        btns = [
            (f"{ICO_DATA} WCZYTAJ BAZĘ E-MAIL", lambda x: self.pick_file("book")),
            (f"{ICO_MAIL} EDYTUJ SZABLON WIADOMOŚCI", lambda x: setattr(self.sm, "current", "tmpl")),
            (f"{ICO_CLIP} DODAJ PLIKI (PDF/FOTO)", self.att_manager_popup),
            (f"{ICO_TEST} TEST DO SIEBIE", self.run_test_mail),
            ("📜 HISTORIA WYSYŁEK", self.show_logs),
            (f"{ICO_SEND} URUCHOM WYSYŁKĘ MASOWĄ", self.start_mailing),
            ("POWRÓT", lambda x: setattr(self.sm, "current", "home"))
        ]
        for t, c in btns: le.add_widget(PremiumButton(text=t, on_press=c))
        self.email_scr.add_widget(le)

        # Inicjalizacja UI Ustawień
        self.setup_settings_uis()

    def setup_settings_uis(self):
        # UI SMTP
        ls = BoxLayout(orientation="vertical", padding=dp(25), spacing=dp(10))
        self.u_in = TextInput(hint_text="Email Gmail"); self.p_in = TextInput(hint_text="Hasło Aplikacji", password=True)
        p = Path(self.user_data_dir) / "smtp.json"
        if p.exists(): d = json.load(open(p)); self.u_in.text = d.get('u',''); self.p_in.text = d.get('p','')
        def save_smtp(_):
            with open(Path(self.user_data_dir) / "smtp.json", "w") as f: json.dump({'u':self.u_in.text,'p':self.p_in.text}, f)
            self.msg("OK", "Zapisano."); setattr(self.sm, "current", "home")
        ls.add_widget(Label(text="Konfiguracja Gmail SMTP")); ls.add_widget(self.u_in); ls.add_widget(self.p_in)
        ls.add_widget(Button(text="ZAPISZ", on_press=save_smtp)); ls.add_widget(Button(text="WRÓĆ", on_press=lambda x: setattr(self.sm, "current", "home")))
        self.smtp_scr.add_widget(ls)

        # UI SZABLON
        lt = BoxLayout(orientation="vertical", padding=dp(20), spacing=dp(10))
        self.ts = TextInput(); self.tb = TextInput()
        rs = self.conn.execute("SELECT val FROM settings WHERE key='t_sub'").fetchone()
        rb = self.conn.execute("SELECT val FROM settings WHERE key='t_body'").fetchone()
        if rs: self.ts.text, self.tb.text = rs[0], rb[0]
        def save_tmpl(_):
            self.conn.execute("UPDATE settings SET val=? WHERE key='t_sub'", (self.ts.text,))
            self.conn.execute("UPDATE settings SET val=? WHERE key='t_body'", (self.tb.text,)); self.conn.commit()
            self.msg("OK", "Zapisano."); setattr(self.sm, "current", "email")
        lt.add_widget(Label(text="Szablon (Tagi: {Imię}, {Data})")); lt.add_widget(self.ts); lt.add_widget(self.tb)
        lt.add_widget(Button(text="ZAPISZ", on_press=save_tmpl)); lt.add_widget(Button(text="WRÓĆ", on_press=lambda x: setattr(self.sm, "current", "email")))
        self.tmpl_scr.add_widget(lt)

    # --- ZINTEGROWANY PATCH PICKER (Fix: Nie wczytuje / Scoped Storage) ---
    def pick_file(self, mode):
        if platform != "android": return
        from jnius import autoclass; from android import activity
        Intent = autoclass("android.content.Intent"); PythonActivity = autoclass("org.kivy.android.PythonActivity")
        intent = Intent(Intent.ACTION_GET_CONTENT); intent.setType("*/*"); intent.addCategory(Intent.CATEGORY_OPENABLE)

        def on_res(req, res, data):
            if data:
                try:
                    uri = data.getData(); ctx = PythonActivity.mActivity; stream = ctx.getContentResolver().openInputStream(uri)
                    dest = Path(self.user_data_dir) / (f"extra_{os.urandom(2).hex()}" if mode == "extra" else f"{mode}_doc.xlsx")
                    with open(dest, "wb") as f:
                        j_buf = autoclass('[B')(16384)
                        while True:
                            r = stream.read(j_buf)
                            if r <= 0: break
                            f.write(bytes(j_buf)[:r])
                    stream.close()
                    if mode == "data": 
                        self.current_file = dest; Clock.schedule_once(lambda x: setattr(self.h_stat, "text", "✔ Excel załadowany. Otwórz tabelę."))
                    elif mode == "book": self.import_book(dest)
                    elif mode == "extra": self.global_attachments.append(str(dest)); self.update_att_lbl()
                except Exception as e: self.msg("Błąd pliku", str(e))
            activity.unbind(on_activity_result=on_res)
        activity.bind(on_activity_result=on_res); PythonActivity.mActivity.startActivityForResult(intent, 1001)

    # --- OBSŁUGA DANYCH ---
    def go_to_table(self, _):
        if not self.current_file: self.msg("Błąd", "Wczytaj najpierw plik Excel Płac!"); return
        from openpyxl import load_workbook
        try:
            wb = load_workbook(str(self.current_file), data_only=True); ws = wb.active
            self.full_data = [[("" if v is None else str(v)) for v in r] for r in ws.iter_rows(values_only=True)]
            self.draw_table(self.full_data); self.sm.current = "table"
        except Exception as e: self.msg("Błąd Excela", str(e))

    def draw_table(self, data):
        self.grid.clear_widgets(); self.grid.cols = 4
        for r in data[1:100]:
            for cell in r[:3]: self.grid.add_widget(Label(text=str(cell)[:12], font_size=11, size_hint_y=None, height=dp(42)))
            btn = Button(text="EKSPORT", size_hint=(None,None), size=(dp(80),dp(42)), background_color=(0,0.6,0.2,1))
            btn.bind(on_press=lambda x, row=r: self.single_export(row)); self.grid.add_widget(btn)

    # --- WYSYŁKA ---
    def start_mailing(self, _):
        if not self.full_data: self.msg("Błąd", "Brak danych z Excela!"); return
        threading.Thread(target=self._mail_task, args=(False,)).start()

    def run_test_mail(self, _):
        if not self.full_data: self.msg("Błąd", "Najpierw wgraj Excel Płac na ekranie startowym."); return
        threading.Thread(target=self._mail_task, args=(True,)).start()

    def _mail_task(self, is_test):
        import smtplib, mimetypes; from email.message import EmailMessage; from openpyxl import Workbook
        p = Path(self.user_data_dir) / "smtp.json"
        if not p.exists(): Clock.schedule_once(lambda x: self.msg("!", "Ustaw Gmail w ustawieniach!")); return
        cfg = json.load(open(p))
        try:
            srv = smtplib.SMTP("smtp.gmail.com", 587, timeout=12); srv.starttls(); srv.login(cfg['u'], cfg['p'])
        except Exception as e: Clock.schedule_once(lambda x: self.msg("Błąd logowania SMTP", str(e))); return

        h, rows = self.full_data[0], (self.full_data[1:2] if is_test else self.full_data[1:])
        ni, si = self.get_name_idxs(h); sent = 0
        for i, r in enumerate(rows):
            target = cfg['u'] if is_test else ""
            if not is_test:
                res = self.conn.execute("SELECT email FROM contacts WHERE name=? AND surname=?", (str(r[ni]).lower().strip(), str(r[si]).lower().strip())).fetchone()
                if res: target = res[0]
            if target:
                msg = EmailMessage(); dat = datetime.now().strftime("%d.%m.%Y")
                msg["Subject"] = self.ts.text.replace("{Imię}", str(r[ni])).replace("{Data}", dat)
                msg["From"], msg["To"] = cfg['u'], target
                msg.set_content(self.tb.text.replace("{Imię}", str(r[ni])).replace("{Data}", dat))
                
                # Excel Report
                tmp = Path(self.user_data_dir) / "tmp_rep.xlsx"; wb = Workbook(); ws = wb.active
                idxs = self.export_columns if self.export_columns else list(range(len(h)))
                ws.append([h[k] for k in idxs]); ws.append([r[k] for k in idxs]); wb.save(str(tmp))
                with open(tmp, "rb") as f: msg.add_attachment(f.read(), maintype="application", subtype="xlsx", filename="Raport.xlsx")
                
                # Global attachments
                for ap in self.global_attachments:
                    if os.path.exists(ap):
                        ct, _ = mimetypes.guess_type(ap); m, st = (ct or "application/octet-stream").split("/",1)
                        with open(ap, "rb") as f: msg.add_attachment(f.read(), maintype=m, subtype=st, filename=os.path.basename(ap))
                
                try: srv.send_message(msg); sent += 1; self.conn.execute("INSERT INTO logs (msg, date) VALUES (?,?)", (f"Wysłano: {target}", dat))
                except: pass
            Clock.schedule_once(lambda dt, p=int(((i+1)/len(rows))*100): setattr(self.prog, "value", p))
        srv.quit(); self.conn.commit(); Clock.schedule_once(lambda x: self.msg("Mailing", f"Zakończono. Wysłano maili: {sent}"))

    # --- POMOCNICZE / POPUPY ---
    def import_book(self, p):
        from openpyxl import load_workbook
        try:
            wb = load_workbook(str(p), data_only=True); ws = wb.active; rows = list(ws.iter_rows(values_only=True))
            h = [str(col).lower() for col in rows[0]]; ni, si = self.get_name_idxs(h)
            mi = next((i for i, v in enumerate(h) if "mail" in v), 2)
            for r in rows[1:]:
                if r[mi]: self.conn.execute("INSERT OR REPLACE INTO contacts VALUES(?,?,?)", (str(r[ni]).lower().strip(), str(r[si]).lower().strip(), str(r[mi]).strip()))
            self.conn.commit(); self.msg("OK", "Baza e-mail zaimportowana.")
        except Exception as e: self.msg("Błąd importu", str(e))

    def mass_export(self, _):
        if not self.full_data: self.msg("!", "Brak danych!"); return
        threading.Thread(target=self._mass_task).start()

    def _mass_task(self):
        from openpyxl import Workbook
        folder = Path("/storage/emulated/0/Documents/FutureExport"); folder.mkdir(parents=True, exist_ok=True)
        h, rows = self.full_data[0], self.full_data[1:]
        for i, r in enumerate(rows):
            wb = Workbook(); ws = wb.active; ws.append(h); ws.append(r); wb.save(str(folder / f"Raport_{r[0]}_{r[1]}.xlsx"))
            Clock.schedule_once(lambda dt, p=int(((i+1)/len(rows))*100): setattr(self.prog, "value", p))
        self.msg("Sukces", "Eksport zakończony. Pliki są w Documents/FutureExport.")

    def single_export(self, r):
        from openpyxl import Workbook
        folder = Path("/storage/emulated/0/Documents/FutureExport"); folder.mkdir(parents=True, exist_ok=True)
        wb = Workbook(); ws = wb.active; ws.append(self.full_data[0]); ws.append(r)
        wb.save(str(folder / f"Pojedynczy_{r[0]}_{r[1]}.xlsx")); self.msg("OK", "Zapisano raport.")

    def att_manager_popup(self, _):
        box = BoxLayout(orientation="vertical", padding=10, spacing=10)
        for ap in self.global_attachments:
            row = BoxLayout(size_hint_y=None, height=dp(40))
            row.add_widget(Label(text=os.path.basename(ap)[:20])); btn = Button(text="USUŃ", on_press=lambda x, p=ap: self.remove_att(p))
            row.add_widget(btn); box.add_widget(row)
        box.add_widget(Button(text="DODAJ PLIK PDF/FOTO", on_press=lambda x: self.pick_file("extra")))
        box.add_widget(Button(text="ZAMKNIJ", on_press=lambda x: p.dismiss())); p = Popup(title="Załączniki", content=box, size_hint=(0.8, 0.6)); p.open()

    def remove_att(self, p):
        if p in self.global_attachments: self.global_attachments.remove(p)
        self.update_att_lbl()

    def update_att_lbl(self): self.att_lbl.text = f"Załączniki dodatkowe: {len(self.global_attachments)}"

    def column_popup(self, _):
        if not self.full_data: return
        box = BoxLayout(orientation="vertical", padding=10); grid = GridLayout(cols=1, size_hint_y=None); grid.bind(minimum_height=grid.setter('height'))
        chks = []
        for i, h in enumerate(self.full_data[0]):
            r = BoxLayout(size_hint_y=None, height=dp(40)); cb = CheckBox(size_hint_x=0.2); cb.active = True
            r.add_widget(cb); r.add_widget(Label(text=str(h))); grid.add_widget(r); chks.append((i, cb))
        def save(_): self.export_columns = [idx for idx, c in chks if c.active]; p.dismiss()
        sv = ScrollView(); sv.add_widget(grid); box.add_widget(sv); box.add_widget(Button(text="ZAPISZ", size_hint_y=None, height=dp(50), on_press=save))
        p = Popup(title="Kolumny raportu", content=box, size_hint=(0.9, 0.8)); p.open()

    def show_logs(self, _):
        logs = self.conn.execute("SELECT msg, date FROM logs ORDER BY id DESC LIMIT 20").fetchall()
        txt = "\n".join([f"{d}: {m}" for m, d in logs])
        self.msg("Ostatnie wysyłki", txt if txt else "Brak historii.")

    def get_name_idxs(self, h):
        ni, si = 0, 1
        for i, v in enumerate(h):
            if "imi" in str(v).lower(): ni = i
            if "nazw" in str(v).lower(): si = i
        return ni, si

    def filter_data(self, ins, val):
        if self.full_data: self.draw_table([self.full_data[0]] + [r for r in self.full_data[1:] if val.lower() in str(r).lower()])

    def msg(self, t, txt): Popup(title=t, content=Label(text=txt, halign="center"), size_hint=(0.8, 0.4)).open()


# =========================================
# MEGA PATCH v15 - DATABASE ENGINE
# =========================================

from kivy.clock import Clock
from kivy.app import App
from kivy.uix.popup import Popup
from kivy.uix.label import Label
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.button import Button
from kivy.uix.textinput import TextInput
from kivy.uix.scrollview import ScrollView
from kivy.uix.gridlayout import GridLayout
from kivy.uix.screenmanager import Screen

import csv
import os
import sqlite3


# =====================
# SAFE POPUP
# =====================

def popup(title, text):

    box = BoxLayout(orientation="vertical", padding=10, spacing=10)

    box.add_widget(Label(text=str(text)))

    btn = Button(text="OK", size_hint_y=None, height=50)

    p = Popup(
        title=title,
        content=box,
        size_hint=(0.8,0.6)
    )

    btn.bind(on_press=p.dismiss)

    box.add_widget(btn)

    p.open()


# =====================
# DEBUG LOGGER
# =====================

def log_error(e):

    try:

        with open("debug_log.txt","a") as f:
            f.write(str(e)+"\n")

    except:
        pass


# =====================
# LOAD DATABASE FILE
# =====================

def load_database_file(self, path):

    try:

        if path.endswith(".csv"):

            data=[]

            with open(path,encoding="utf8") as f:
                reader=csv.reader(f)

                for r in reader:
                    data.append(r)

            self.full_data=data

            popup("CSV","Załadowano plik CSV")

            return

        if path.endswith(".db"):

            conn=sqlite3.connect(path)

            c=conn.cursor()

            tables=c.execute(
                "SELECT name FROM sqlite_master WHERE type='table'"
            ).fetchall()

            if not tables:

                popup("SQLite","Brak tabel")
                return

            table=tables[0][0]

            rows=c.execute(f"SELECT * FROM {table}").fetchall()

            self.full_data=[list(r) for r in rows]

            popup("SQLite",f"Tabela: {table}")

            return

        if path.endswith(".xlsx"):

            popup(
                "XLSX",
                "Obsługa XLSX wymaga openpyxl"
            )

    except Exception as e:

        log_error(e)

        popup("Błąd",str(e))


# =====================
# TABLE VIEW
# =====================

def draw_table(self,data):

    try:

        self.table_scr.clear_widgets()

        root=BoxLayout(orientation="vertical")

        search=TextInput(
            hint_text="Szukaj...",
            size_hint_y=None,
            height=50
        )

        root.add_widget(search)

        scroll=ScrollView()

        grid=GridLayout(
            cols=len(data[0]) if data else 1,
            size_hint_y=None
        )

        grid.bind(minimum_height=grid.setter("height"))

        current=data[:]

        def render(rows):

            grid.clear_widgets()

            if not rows:
                grid.add_widget(Label(text="Brak danych"))
                return

            for r in rows:

                for c in r:

                    grid.add_widget(Label(
                        text=str(c),
                        size_hint_y=None,
                        height=40
                    ))

        render(current)

        def search_fn(instance,value):

            if not value:

                render(data)

                return

            res=[]

            for r in data:

                if any(value.lower() in str(c).lower() for c in r):

                    res.append(r)

            render(res)

        search.bind(text=search_fn)

        scroll.add_widget(grid)

        root.add_widget(scroll)

        bottom=BoxLayout(size_hint_y=None,height=60)

        back=Button(text="Powrót")

        back.bind(
            on_press=lambda x:
            setattr(self.sm,"current","home")
        )

        export=Button(text="Eksport CSV")

        def export_csv(x):

            try:

                with open("export.csv","w",newline="",encoding="utf8") as f:

                    w=csv.writer(f)

                    for r in data:
                        w.writerow(r)

                popup("Eksport","Zapisano export.csv")

            except Exception as e:

                popup("Błąd",str(e))

        export.bind(on_press=export_csv)

        bottom.add_widget(back)
        bottom.add_widget(export)

        root.add_widget(bottom)

        self.table_scr.add_widget(root)

    except Exception as e:

        log_error(e)

        popup("Crash tabeli",str(e))


# =====================
# OPEN TABLE
# =====================

def go_to_table(self,*a):

    try:

        if not hasattr(self,"full_data") or not self.full_data:

            popup(
                "Brak danych",
                "Najpierw wczytaj plik"
            )

            return

        self.draw_table(self.full_data)

        self.sm.current="table"

    except Exception as e:

        log_error(e)

        popup("Błąd",str(e))


# =====================
# DEBUG SCREEN
# =====================

def debug_screen(self):

    if self.sm.has_screen("debug"):
        return

    s=Screen(name="debug")

    root=BoxLayout(
        orientation="vertical",
        padding=20,
        spacing=20
    )

    root.add_widget(Label(text="DEBUG PANEL"))

    btn=Button(text="Info o bazie")

    def info(x):

        if hasattr(self,"full_data") and self.full_data:

            r=len(self.full_data)
            c=len(self.full_data[0])

            popup(
                "Baza",
                f"Rekordy: {r}\nKolumny: {c}"
            )

        else:

            popup("Baza","Brak danych")

    btn.bind(on_press=info)

    root.add_widget(btn)

    back=Button(text="Powrót")

    back.bind(
        on_press=lambda x:
        setattr(self.sm,"current","home")
    )

    root.add_widget(back)

    s.add_widget(root)

    self.sm.add_widget(s)


# =====================
# TEST SCREEN
# =====================

def test_screen(self):

    if self.sm.has_screen("test"):
        return

    s=Screen(name="test")

    root=BoxLayout(
        orientation="vertical",
        padding=20,
        spacing=20
    )

    root.add_widget(Label(text="TEST PANEL"))

    btn=Button(text="Test Popup")

    btn.bind(
        on_press=lambda x:
        popup("Test","Patch działa")
    )

    root.add_widget(btn)

    debug=Button(text="DEBUG")

    debug.bind(
        on_press=lambda x:
        setattr(self.sm,"current","debug")
    )

    root.add_widget(debug)

    back=Button(text="Powrót")

    back.bind(
        on_press=lambda x:
        setattr(self.sm,"current","home")
    )

    root.add_widget(back)

    s.add_widget(root)

    self.sm.add_widget(s)


# =====================
# ADD BUTTONS
# =====================

def add_buttons(self):

    try:

        layout=self.home_scr.children[0]

        test=Button(text="TEST",size_hint_y=None,height=60)

        test.bind(
            on_press=lambda x:
            setattr(self.sm,"current","test")
        )

        debug=Button(text="DEBUG",size_hint_y=None,height=60)

        debug.bind(
            on_press=lambda x:
            setattr(self.sm,"current","debug")
        )

        layout.add_widget(test)
        layout.add_widget(debug)

    except:
        pass


# =====================
# APPLY PATCH
# =====================

def apply_patch(dt):

    try:

        FutureApp.draw_table=draw_table
        FutureApp.go_to_table=go_to_table
        FutureApp.load_database_file=load_database_file

        app=App.get_running_app()

        if app:

            debug_screen(app)
            test_screen(app)
            add_buttons(app)

        print("MEGA PATCH v15 aktywny")

    except Exception as e:

        log_error(e)


Clock.schedule_once(apply_patch,1)

if __name__ == "__main__":
    try:
        FutureApp().run()
    except Exception as e:
        import traceback
        # Zapis błędu do bezpiecznej lokalizacji wewnętrznej (fix crash log)
        with open(os.path.join(os.getcwd(), "last_crash.txt"), "w") as f:
            f.write(traceback.format_exc())
