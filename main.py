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

# SYMBOLE (Bezpieczne dla Androida)
ICO_DATA = "FOLDER"
ICO_TABLE = "TABELA"
ICO_GEAR = "Opcje"
ICO_MAIL = "Mail"
ICO_SEND = "START"
ICO_CLIP = "Spinacz"

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
        
        # Ekrany
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
        db_p = Path(self.user_data_dir) / "future_data_v11.db"
        self.conn = sqlite3.connect(str(db_p), check_same_thread=False)
        self.conn.execute("CREATE TABLE IF NOT EXISTS contacts (name TEXT, surname TEXT, email TEXT, PRIMARY KEY(name, surname))")
        self.conn.execute("CREATE TABLE IF NOT EXISTS settings (key TEXT PRIMARY KEY, val TEXT)")
        self.conn.execute("CREATE TABLE IF NOT EXISTS logs (id INTEGER PRIMARY KEY, msg TEXT, date TEXT)")
        
        if not self.conn.execute("SELECT val FROM settings WHERE key='t_body'").fetchone():
            self.conn.execute("INSERT OR REPLACE INTO settings VALUES ('t_sub', 'Raport: {Imię} {Nazwisko}')")
            self.conn.execute("INSERT OR REPLACE INTO settings VALUES ('t_body', 'Witaj {Imię},\n\nPrzesyłamy raport z dnia {Data}.')")
        self.conn.commit()

    def setup_ui(self):
        # HOME
        l_home = BoxLayout(orientation="vertical", padding=dp(25), spacing=dp(12))
        l_home.add_widget(Label(text="FUTURE 11.0", font_size=dp(28), bold=True))
        l_home.add_widget(PremiumButton(text=f"{ICO_DATA} WCZYTAJ EXCEL PŁAC", on_press=lambda x: self.pick_file("data")))
        l_home.add_widget(PremiumButton(text=f"{ICO_TABLE} PODGLĄD TABELI", on_press=self.go_to_table))
        l_home.add_widget(PremiumButton(text=f"{ICO_SEND} CENTRUM MAILINGOWE", on_press=lambda x: setattr(self.sm, "current", "email")))
        l_home.add_widget(PremiumButton(text=f"{ICO_GEAR} USTAWIENIA GMAIL", on_press=lambda x: setattr(self.sm, "current", "smtp")))
        self.h_stat = Label(text="System gotowy", color=(0.6, 0.6, 0.6, 1))
        l_home.add_widget(self.h_stat); self.home_scr.add_widget(l_home)

        # MAILING CENTER
        le = BoxLayout(orientation="vertical", padding=dp(20), spacing=dp(8))
        le.add_widget(Label(text="CENTRUM OPERACJI", font_size=dp(22), bold=True))
        self.att_lbl = Label(text="Załączniki: 0", size_hint_y=None, height=dp(30))
        le.add_widget(self.att_lbl)
        btns = [
            (f"{ICO_DATA} WCZYTAJ KONTAKTY", lambda x: self.pick_file("book")),
            (f"{ICO_MAIL} EDYTUJ TREŚĆ", lambda x: setattr(self.sm, "current", "tmpl")),
            (f"{ICO_CLIP} DODAJ PDF/ZAŁĄCZNIK", self.att_manager_popup),
            ("TEST MAILA DO SIEBIE", self.run_test_mail),
            ("📜 LOGI WYSYŁKI", self.show_logs),
            (f"{ICO_SEND} URUCHOM WYSYŁKĘ", self.start_mailing),
            ("POWRÓT", lambda x: setattr(self.sm, "current", "home"))
        ]
        for t, c in btns: le.add_widget(PremiumButton(text=t, on_press=c))
        self.email_scr.add_widget(le)

        # SMTP & TEMPLATE
        self.setup_settings_uis()

    def setup_settings_uis(self):
        # UI SMTP
        ls = BoxLayout(orientation="vertical", padding=dp(25), spacing=dp(10))
        self.u_in = TextInput(hint_text="Email Gmail"); self.p_in = TextInput(hint_text="Hasło Aplikacji", password=True)
        p = Path(self.user_data_dir) / "smtp.json"
        if p.exists(): d = json.load(open(p)); self.u_in.text, self.p_in.text = d.get('u',''), d.get('p','')
        def save_smtp(_):
            with open(Path(self.user_data_dir) / "smtp.json", "w") as f: json.dump({'u':self.u_in.text,'p':self.p_in.text}, f)
            self.msg("OK", "Zapisano."); setattr(self.sm, "current", "home")
        ls.add_widget(Label(text="Ustawienia SMTP Gmail")); ls.add_widget(self.u_in); ls.add_widget(self.p_in)
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
        lt.add_widget(Label(text="Treść (Tagi: {Imię}, {Nazwisko}, {Data})")); lt.add_widget(self.ts); lt.add_widget(self.tb)
        lt.add_widget(Button(text="ZAPISZ", on_press=save_tmpl)); lt.add_widget(Button(text="WRÓĆ", on_press=lambda x: setattr(self.sm, "current", "email")))
        self.tmpl_scr.add_widget(lt)

    # --- ZINTEGROWANY PICKER Z PATCHA (NAPRAWIA WCZYTYWANIE) ---
    def pick_file(self, mode):
        if platform != "android": self.msg("Błąd", "Działa tylko na Android"); return
        from jnius import autoclass; from android import activity
        Intent = autoclass("android.content.Intent")
        intent = Intent(Intent.ACTION_GET_CONTENT); intent.setType("*/*"); intent.addCategory(Intent.CATEGORY_OPENABLE)

        def on_res(req, res, data):
            if data:
                try:
                    uri = data.getData(); ctx = autoclass("org.kivy.android.PythonActivity").mActivity
                    stream = ctx.getContentResolver().openInputStream(uri)
                    dest = Path(self.user_data_dir) / f"{mode}_temp.xlsx"
                    with open(dest, "wb") as f:
                        j_buf = autoclass('[B')(16384)
                        while True:
                            r = stream.read(j_buf)
                            if r <= 0: break
                            f.write(bytes(j_buf)[:r])
                    stream.close()
                    
                    if mode == "data": 
                        self.current_file = dest
                        # Natychmiastowe ładowanie danych do pamięci
                        self.load_excel_to_memory(dest)
                        Clock.schedule_once(lambda x: setattr(self.h_stat, "text", "✔ Excel załadowany poprawnie!"))
                    elif mode == "book": 
                        self.import_contacts(dest)
                    elif mode == "extra": 
                        self.global_attachments.append(str(dest)); self.update_att_lbl()
                except Exception as e: self.msg("Błąd pliku", str(e))
            activity.unbind(on_activity_result=on_res)
        activity.bind(on_activity_result=on_res); autoclass("org.kivy.android.PythonActivity").mActivity.startActivityForResult(intent, 1001)

    # --- ŁADOWANIE I TABELA (BEZPOŚREDNIO W KLASIE) ---
    def load_excel_to_memory(self, path):
        from openpyxl import load_workbook
        wb = load_workbook(str(path), data_only=True); ws = wb.active
        self.full_data = [[("" if v is None else str(v)) for v in r] for r in ws.iter_rows(values_only=True)]

    def go_to_table(self, _):
        if not self.full_data: self.msg("Błąd", "Wczytaj najpierw plik Excel!"); return
        self.render_table_view(self.full_data); self.sm.current = "table"

    def render_table_view(self, data):
        self.table_scr.clear_widgets()
        root = BoxLayout(orientation="vertical", padding=dp(5))
        
        # Wyszukiwarka
        search = TextInput(hint_text="Szukaj osoby...", size_hint_y=None, height=dp(50), multiline=False)
        root.add_widget(search)
        
        # Grid
        self.grid = GridLayout(cols=4, size_hint_y=None, spacing=dp(2))
        self.grid.bind(minimum_height=self.grid.setter('height'))
        
        def populate_grid(rows):
            self.grid.clear_widgets()
            for r in rows[:100]: # Prędkość renderowania
                for cell in r[:3]: self.grid.add_widget(Label(text=str(cell)[:13], font_size=11, size_hint_y=None, height=dp(40)))
                btn = Button(text="ZAPISZ", size_hint=(None,None), size=(dp(80),dp(40)), background_color=(0,0.6,0.2,1))
                btn.bind(on_press=lambda x, row=r: self.single_export(row))
                self.grid.add_widget(btn)

        populate_grid(data[1:])
        search.bind(text=lambda i, v: populate_grid([r for r in data[1:] if any(v.lower() in str(c).lower() for c in r)]))

        sv = ScrollView(); sv.add_widget(self.grid); root.add_widget(sv)
        
        # Progress Bar i przyciski dolne
        self.prog = ProgressBar(max=100, size_hint_y=None, height=dp(10)); root.add_widget(self.prog)
        
        bot = BoxLayout(size_hint_y=None, height=dp(100), orientation="vertical", spacing=dp(2))
        bot.add_widget(Button(text="EKSPORTUJ WSZYSTKO", on_press=self.mass_export))
        bot.add_widget(Button(text="WYBIERZ KOLUMNY", on_press=self.column_popup))
        bot.add_widget(Button(text="POWRÓT", on_press=lambda x: setattr(self.sm, "current", "home")))
        root.add_widget(bot)
        
        self.table_scr.add_widget(root)

    # --- WYSYŁKA MAILI ---
    def start_mailing(self, _):
        if not self.full_data: self.msg("!", "Brak danych!"); return
        threading.Thread(target=self._mailing_logic, args=(False,)).start()

    def run_test_mail(self, _):
        if not self.full_data: self.msg("!", "Wgraj Excel!"); return
        threading.Thread(target=self._mailing_logic, args=(True,)).start()

    def _mailing_logic(self, is_test):
        import smtplib, mimetypes; from email.message import EmailMessage; from openpyxl import Workbook
        p = Path(self.user_data_dir) / "smtp.json"
        if not p.exists(): Clock.schedule_once(lambda x: self.msg("!", "Ustaw SMTP!")); return
        cfg = json.load(open(p))
        try:
            srv = smtplib.SMTP("smtp.gmail.com", 587, timeout=12); srv.starttls(); srv.login(cfg['u'], cfg['p'])
        except Exception as e: Clock.schedule_once(lambda x: self.msg("SMTP Error", str(e))); return

        h, rows = self.full_data[0], (self.full_data[1:2] if is_test else self.full_data[1:])
        sent = 0
        for i, r in enumerate(rows):
            target = cfg['u'] if is_test else ""
            if not is_test:
                res = self.conn.execute("SELECT email FROM contacts WHERE name=? AND surname=?", (str(r[0]).lower().strip(), str(r[1]).lower().strip())).fetchone()
                if res: target = res[0]
            
            if target:
                msg = EmailMessage(); dat = datetime.now().strftime("%d.%m.%Y")
                msg["Subject"] = self.ts.text.replace("{Imię}", str(r[0])).replace("{Data}", dat)
                msg["From"], msg["To"] = cfg['u'], target
                msg.set_content(self.tb.text.replace("{Imię}", str(r[0])).replace("{Data}", dat))
                
                # Excel
                tmp = Path(self.user_data_dir) / "tmp.xlsx"; wb = Workbook(); ws = wb.active
                idxs = self.export_columns if self.export_columns else list(range(len(h)))
                ws.append([h[k] for k in idxs]); ws.append([r[k] for k in idxs]); wb.save(str(tmp))
                with open(tmp, "rb") as f: msg.add_attachment(f.read(), maintype="application", subtype="xlsx", filename="Raport.xlsx")
                
                # Załączniki
                for ap in self.global_attachments:
                    if os.path.exists(ap):
                        ct, _ = mimetypes.guess_type(ap); m, st = (ct or "application/octet-stream").split("/",1)
                        with open(ap, "rb") as f: msg.add_attachment(f.read(), maintype=m, subtype=st, filename=os.path.basename(ap))
                
                try: srv.send_message(msg); sent += 1; self.conn.execute("INSERT INTO logs (msg, date) VALUES (?,?)", (f"Wysłano: {target}", dat))
                except: pass
            
            # Progress bar update w tabeli jeśli jest otwarta
            if hasattr(self, 'prog'): Clock.schedule_once(lambda dt, p=int(((i+1)/len(rows))*100): setattr(self.prog, "value", p))
            
        srv.quit(); self.conn.commit(); Clock.schedule_once(lambda x: self.msg("Koniec", f"Wysłano {sent} maili."))

    # --- POMOCNICZE / EKSPORT ---
    def import_contacts(self, p):
        from openpyxl import load_workbook
        wb = load_workbook(str(p), data_only=True); ws = wb.active; rows = list(ws.iter_rows(values_only=True))
        for r in rows[1:]:
            if r[2]: self.conn.execute("INSERT OR REPLACE INTO contacts VALUES(?,?,?)", (str(r[0]).lower().strip(), str(r[1]).lower().strip(), str(r[2]).strip()))
        self.conn.commit(); self.msg("OK", "Zaimportowano bazę maili.")

    def mass_export(self, _):
        threading.Thread(target=self._mass_task).start()

    def _mass_task(self):
        from openpyxl import Workbook
        folder = Path("/storage/emulated/0/Documents/FutureExport"); folder.mkdir(parents=True, exist_ok=True)
        h, rows = self.full_data[0], self.full_data[1:]
        for i, r in enumerate(rows):
            wb = Workbook(); ws = wb.active; ws.append(h); ws.append(r); wb.save(str(folder / f"Raport_{r[0]}_{r[1]}.xlsx"))
            Clock.schedule_once(lambda dt, p=int(((i+1)/len(rows))*100): setattr(self.prog, "value", p))
        self.msg("OK", "Zapisano w Documents/FutureExport")

    def single_export(self, r):
        from openpyxl import Workbook
        folder = Path("/storage/emulated/0/Documents/FutureExport"); folder.mkdir(parents=True, exist_ok=True)
        wb = Workbook(); ws = wb.active; ws.append(self.full_data[0]); ws.append(r)
        wb.save(str(folder / f"Pojedynczy_{r[0]}_{r[1]}.xlsx")); self.msg("OK", "Zapisano.")

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

    def att_manager_popup(self, _):
        box = BoxLayout(orientation="vertical", padding=10, spacing=10)
        for ap in self.global_attachments:
            row = BoxLayout(size_hint_y=None, height=dp(40))
            row.add_widget(Label(text=os.path.basename(ap)[:20])); btn = Button(text="USUŃ", on_press=lambda x, p=ap: self.remove_att(p))
            row.add_widget(btn); box.add_widget(row)
        box.add_widget(Button(text="DODAJ DOKUMENT", on_press=lambda x: self.pick_file("extra")))
        box.add_widget(Button(text="ZAMKNIJ", on_press=lambda x: p.dismiss())); p = Popup(title="Załączniki", content=box, size_hint=(0.8, 0.6)); p.open()

    def remove_att(self, p):
        if p in self.global_attachments: self.global_attachments.remove(p)
        self.update_att_lbl()

    def update_att_lbl(self): self.att_lbl.text = f"Załączniki: {len(self.global_attachments)}"
    def show_logs(self, _):
        logs = self.conn.execute("SELECT msg, date FROM logs ORDER BY id DESC LIMIT 20").fetchall()
        txt = "\n".join([f"{d}: {m}" for m, d in logs]); self.msg("Historia", txt if txt else "Brak logów.")
    def msg(self, t, txt): Popup(title=t, content=Label(text=txt, halign="center"), size_hint=(0.8, 0.4)).open()


# =========================
# FUTURE APP STABILITY PATCH
# =========================

def _safe_render_table_view(self, data):
    from kivy.uix.boxlayout import BoxLayout
    from kivy.uix.textinput import TextInput
    from kivy.uix.scrollview import ScrollView
    from kivy.uix.gridlayout import GridLayout
    from kivy.uix.label import Label
    from kivy.uix.button import Button
    from kivy.metrics import dp

    if not data or len(data) < 2:
        self.msg("Błąd", "Excel nie zawiera danych.")
        return

    self.table_scr.clear_widgets()

    root = BoxLayout(orientation="vertical", padding=dp(5))

    search = TextInput(
        hint_text="Szukaj osoby...",
        size_hint_y=None,
        height=dp(50),
        multiline=False
    )

    root.add_widget(search)

    self.grid = GridLayout(
        cols=4,
        size_hint_y=None,
        spacing=dp(2)
    )

    self.grid.bind(minimum_height=self.grid.setter('height'))

    def populate(rows):
        self.grid.clear_widgets()

        for r in rows[:200]:

            row = list(r) + ["", "", ""]

            for cell in row[:3]:
                self.grid.add_widget(
                    Label(
                        text=str(cell)[:15],
                        font_size=11,
                        size_hint_y=None,
                        height=dp(40)
                    )
                )

            btn = Button(
                text="ZAPISZ",
                size_hint=(None, None),
                size=(dp(80), dp(40)),
                background_color=(0,0.6,0.2,1)
            )

            btn.bind(on_press=lambda x, row=r: self.single_export(row))
            self.grid.add_widget(btn)

    populate(data[1:])

    def do_search(instance, value):

        value = value.lower()

        filtered = [
            r for r in data[1:]
            if any(value in str(c).lower() for c in r if c)
        ]

        populate(filtered)

    search.bind(text=do_search)

    sv = ScrollView()
    sv.add_widget(self.grid)

    root.add_widget(sv)

    self.prog = ProgressBar(
        max=100,
        size_hint_y=None,
        height=dp(10)
    )

    root.add_widget(self.prog)

    bot = BoxLayout(
        size_hint_y=None,
        height=dp(100),
        orientation="vertical",
        spacing=dp(2)
    )

    bot.add_widget(Button(text="EKSPORTUJ WSZYSTKO", on_press=self.mass_export))
    bot.add_widget(Button(text="WYBIERZ KOLUMNY", on_press=self.column_popup))
    bot.add_widget(Button(text="POWRÓT", on_press=lambda x: setattr(self.sm,"current","home")))

    root.add_widget(bot)

    self.table_scr.add_widget(root)


def _safe_att_popup(self, _):

    from kivy.uix.boxlayout import BoxLayout
    from kivy.uix.button import Button
    from kivy.uix.label import Label
    from kivy.uix.popup import Popup
    from kivy.metrics import dp

    box = BoxLayout(
        orientation="vertical",
        padding=10,
        spacing=10
    )

    popup = Popup(
        title="Załączniki",
        content=box,
        size_hint=(0.8,0.6)
    )

    for ap in self.global_attachments:

        row = BoxLayout(size_hint_y=None,height=dp(40))

        row.add_widget(Label(text=os.path.basename(ap)[:25]))

        btn = Button(text="USUŃ")

        btn.bind(on_press=lambda x,p=ap: self.remove_att(p))

        row.add_widget(btn)

        box.add_widget(row)

    box.add_widget(Button(
        text="DODAJ DOKUMENT",
        on_press=lambda x: self.pick_file("extra")
    ))

    box.add_widget(Button(
        text="ZAMKNIJ",
        on_press=lambda x: popup.dismiss()
    ))

    popup.open()


def _safe_import_contacts(self, p):

    try:
        from openpyxl import load_workbook

        wb = load_workbook(str(p), data_only=True)

        ws = wb.active

        rows = list(ws.iter_rows(values_only=True))

        added = 0

        for r in rows[1:]:

            if not r:
                continue

            name = str(r[0]).lower().strip() if len(r) > 0 else ""
            surname = str(r[1]).lower().strip() if len(r) > 1 else ""
            email = str(r[2]).strip() if len(r) > 2 else ""

            if email:

                self.conn.execute(
                    "INSERT OR REPLACE INTO contacts VALUES(?,?,?)",
                    (name,surname,email)
                )

                added += 1

        self.conn.commit()

        self.msg(
            "Baza zaimportowana",
            f"Dodano kontaktów: {added}"
        )

    except Exception as e:

        self.msg("Błąd importu", str(e))


def _safe_load_excel(self, path):

    try:

        from openpyxl import load_workbook

        wb = load_workbook(str(path), data_only=True)

        ws = wb.active

        self.full_data = [
            [("" if v is None else str(v)) for v in r]
            for r in ws.iter_rows(values_only=True)
        ]

        rows = max(0,len(self.full_data)-1)
        cols = len(self.full_data[0]) if self.full_data else 0

        self.msg(
            "Excel wczytany",
            f"Wiersze: {rows}\nKolumny: {cols}"
        )

    except Exception as e:

        self.msg("Błąd Excela", str(e))


# NADPISANIE FUNKCJI APLIKACJI
FutureApp.render_table_view = _safe_render_table_view
FutureApp.att_manager_popup = _safe_att_popup
FutureApp.import_contacts = _safe_import_contacts
FutureApp.load_excel_to_memory = _safe_load_excel
if __name__ == "__main__":
    try:
        FutureApp().run()
    except Exception as e:
        import traceback
        with open("critical_error.txt", "w") as f: f.write(traceback.format_exc())
