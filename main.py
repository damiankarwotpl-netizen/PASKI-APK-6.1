import os
import sqlite3
import threading
import json
import smtplib
import re
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
from kivy.uix.recycleview import RecycleView

from openpyxl import load_workbook, Workbook
from openpyxl.styles import PatternFill, Border, Side, Alignment, Font

APP_TITLE = "Future 9.1 ULTRA PRO"

class PremiumButton(Button):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.background_normal = ""
        self.background_color = (0.2, 0.4, 0.9, 1)
        self.color = (1, 1, 1, 1)
        self.size_hint_y = None
        self.height = dp(50)

# --- KLASY DO WYDAJNEJ TABELI (RecycleView) ---
class TableRow(BoxLayout):
    def __init__(self, **kwargs):
        data = kwargs.pop('row_data', [])
        is_header = kwargs.pop('is_header', False)
        callback = kwargs.pop('callback', None)
        super().__init__(**kwargs)
        self.orientation = 'horizontal'
        self.size_hint_y = None
        self.height = dp(45)
        
        for text in data:
            self.add_widget(Label(
                text=str(text),
                size_hint_x=None,
                width=dp(160),
                bold=is_header
            ))
        
        if not is_header and callback:
            btn = Button(text="EKSPORT", size_hint_x=None, width=dp(100), background_color=(0, 0.6, 0, 1))
            btn.bind(on_press=lambda x: callback(data))
            self.add_widget(btn)
        elif is_header:
            self.add_widget(Label(text="AKCJA", size_hint_x=None, width=dp(100), bold=True))

class RVTable(RecycleView):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.data = []

class FutureApp(App):
    def build(self):
        Window.clearcolor = (0.08, 0.1, 0.15, 1)
        self.full_data = []  # Dane surowe
        self.current_file = None
        self.export_columns = []
        
        # Zapytanie o uprawnienia na Android
        if platform == 'android':
            from android.permissions import request_permissions, Permission
            request_permissions([Permission.READ_EXTERNAL_STORAGE, Permission.WRITE_EXTERNAL_STORAGE])

        self.sm = ScreenManager()
        self.init_ui()
        self.init_db()
        return self.sm

    def init_db(self):
        db_path = Path(self.user_data_dir) / "app_v9_final.db"
        self.conn = sqlite3.connect(str(db_path), check_same_thread=False)
        self.conn.execute("CREATE TABLE IF NOT EXISTS contacts (name TEXT, surname TEXT, email TEXT, PRIMARY KEY(name, surname))")
        self.conn.commit()

    def init_ui(self):
        # --- HOME ---
        self.home_scr = Screen(name="home")
        l_home = BoxLayout(orientation="vertical", padding=dp(25), spacing=dp(20))
        l_home.add_widget(Label(text=APP_TITLE, font_size=26, bold=True))
        
        btns = [
            ("\ud83d\udcc2 WCZYTAJ DANE", lambda x: self.pick_file(mode="data")),
            ("\ud83d\udcca OTW\u00d3RZ TABEL\u0118", self.go_to_table),
            ("\u2699 USTAWIENIA GMAIL", lambda x: setattr(self.sm, "current", "smtp"))
        ]
        for text, cmd in btns:
            b = PremiumButton(text=text)
            b.bind(on_press=cmd)
            l_home.add_widget(b)
            
        self.home_status = Label(text="Zacznij od wczytania Excela", color=(0.7, 0.7, 0.7, 1))
        l_home.add_widget(self.home_status)
        self.home_scr.add_widget(l_home)

        # --- TABLE (Zoptymalizowana) ---
        self.table_scr = Screen(name="table")
        lt = BoxLayout(orientation="vertical", padding=dp(5))
        top = BoxLayout(size_hint_y=None, height=dp(60), spacing=dp(5))
        self.search = TextInput(hint_text="Szukaj...", multiline=False)
        self.search.bind(text=self.filter_data)
        
        b_next = Button(text="WYSY\u0141KA", size_hint_x=0.3, background_color=(0.2, 0.6, 0.2, 1))
        b_next.bind(on_press=lambda x: setattr(self.sm, "current", "email"))
        b_back = Button(text="COFNIJ", size_hint_x=0.2)
        b_back.bind(on_press=lambda x: setattr(self.sm, "current", "home"))
        
        top.add_widget(self.search); top.add_widget(b_next); top.add_widget(b_back)
        
        self.rv = RVTable()
        self.rv_layout = GridLayout(cols=1, size_hint_y=None, spacing=dp(2))
        self.rv_layout.bind(minimum_height=self.rv_layout.setter('height'))
        self.rv.add_widget(self.rv_layout)
        
        self.progress = ProgressBar(max=100, size_hint_y=None, height=dp(10))
        lt.add_widget(top); lt.add_widget(self.rv); lt.add_widget(self.progress)
        self.table_scr.add_widget(lt)

        # --- EMAIL & SMTP (Logika pozosta\u0142a podobna, ale z ulepszonymi funkcjami) ---
        self.email_scr = Screen(name="email")
        self.setup_email_ui()
        
        self.smtp_scr = Screen(name="smtp")
        self.setup_smtp_ui()

        self.sm.add_widget(self.home_scr)
        self.sm.add_widget(self.table_scr)
        self.sm.add_widget(self.email_scr)
        self.sm.add_widget(self.smtp_scr)

    def setup_email_ui(self):
        le = BoxLayout(orientation="vertical", padding=dp(20), spacing=dp(10))
        btn_data = [
            ("\ud83d\udce5 WCZYTAJ KONTAKTY", lambda x: self.pick_file(mode="book")),
            ("\ud83d\udccb WYBIERZ KOLUMNY", self.column_popup),
            ("\ud83d\udcbe EKSPORTUJ WSZYSTKO", self.start_export_all_thread),
            ("\ud83d\ude80 WY\u015aLIJ MAILE", self.start_mailing),
            ("POWR\u00d3T", lambda x: setattr(self.sm, "current", "table"))
        ]
        for t, c in btn_data:
            btn = PremiumButton(text=t); btn.bind(on_press=c); le.add_widget(btn)
        self.email_status = Label(text="Gotowy do dzia\u0142ania")
        le.add_widget(self.email_status)
        self.email_scr.add_widget(le)

    def setup_smtp_ui(self):
        ls = BoxLayout(orientation="vertical", padding=dp(20), spacing=dp(10))
        self.s_user = TextInput(hint_text="Gmail", multiline=False)
        self.s_pass = TextInput(hint_text="Has\u0142o Aplikacji", password=True, multiline=False)
        b_test = PremiumButton(text="TESTUJ PO\u0141\u0104CZENIE"); b_test.bind(on_press=self.test_smtp)
        b_save = PremiumButton(text="ZAPISZ"); b_save.bind(on_press=self.save_smtp)
        ls.add_widget(Label(text="Konfiguracja SMTP", font_size=20))
        ls.add_widget(self.s_user); ls.add_widget(self.s_pass)
        ls.add_widget(b_test); ls.add_widget(b_save)
        ls.add_widget(PremiumButton(text="POWR\u00d3T", on_press=lambda x: setattr(self.sm, "current", "home")))
        self.smtp_scr.add_widget(ls)
        self.load_smtp()

    def sanitize_filename(self, name):
        return re.sub(r'[\\/*?:"<>|]', "", str(name))

    def pick_file(self, mode):
        if platform != "android":
            self.msg("B\u0142\u0105d", "Tylko na Android.")
            return
        from jnius import autoclass; from android import activity
        Intent = autoclass("android.content.Intent")
        intent = Intent(Intent.ACTION_OPEN_DOCUMENT)
        intent.setType("*/*")
        intent.addCategory(Intent.CATEGORY_OPENABLE)
        
        def on_res(req, res, dt):
            if dt:
                try:
                    uri = dt.getData()
                    ctx = autoclass("org.kivy.android.PythonActivity").mActivity
                    stream = ctx.getContentResolver().openInputStream(uri)
                    local = Path(self.user_data_dir) / ("data.xlsx" if mode == "data" else "book.xlsx")
                    
                    j_buf = autoclass('[B')(16384)
                    with open(local, "wb") as f:
                        while True:
                            r = stream.read(j_buf)
                            if r <= 0: break
                            f.write(bytes(j_buf)[:r])
                    stream.close()
                    
                    if mode == "data": 
                        self.current_file = local
                        Clock.schedule_once(lambda x: setattr(self.home_status, "text", "Za\u0142adowano plik."))
                    else: 
                        self.import_contacts_to_db(local)
                except Exception as e:
                    Clock.schedule_once(lambda x: self.msg("B\u0142\u0105d", str(e)))
            activity.unbind(on_activity_result=on_res)
        activity.bind(on_activity_result=on_res)
        autoclass("org.kivy.android.PythonActivity").mActivity.startActivityForResult(intent, 1001)

    def import_contacts_to_db(self, path):
        try:
            wb = load_workbook(str(path), data_only=True); ws = wb.active
            rows = list(ws.iter_rows(values_only=True))
            if not rows: return
            h = [str(x).lower() for x in rows[0]]
            # Szukanie indeks\u00f3w kolumn
            id_mi = next((i for i, v in enumerate(h) if "mail" in v), None)
            id_ni = next((i for i, v in enumerate(h) if "imi" in v), 0)
            id_si = next((i for i, v in enumerate(h) if "nazw" in v), 1)

            if id_mi is None: 
                self.msg("B\u0142\u0105d", "Nie znaleziono kolumny Email!")
                return
                
            for r in rows[1:]:
                if r[id_mi]:
                    self.conn.execute("INSERT OR REPLACE INTO contacts VALUES(?,?,?)", 
                                     (str(r[id_ni]).lower().strip(), str(r[id_si]).lower().strip(), str(r[id_mi]).strip()))
            self.conn.commit()
            self.msg("Sukces", "Baza kontakt\u00f3w zaktualizowana.")
        except Exception as e: self.msg("B\u0142\u0105d", str(e))

    def go_to_table(self, _):
        if not self.current_file: return
        try:
            wb = load_workbook(str(self.current_file), data_only=True)
            ws = wb.active
            self.full_data = [[("" if v is None else str(v)) for v in r] for r in ws.iter_rows(values_only=True)]
            self.update_table_view(self.full_data)
            self.sm.current = "table"
        except Exception as e: self.msg("B\u0142\u0105d", str(e))

    def update_table_view(self, data_list):
        self.rv_layout.clear_widgets()
        if not data_list: return
        
        # Nag\u0142\u00f3wek
        self.rv_layout.add_widget(TableRow(row_data=data_list[0], is_header=True))
        
        # Wiersze (RecycleView manualne w tym przypadku dla zachowania prostoty GridLayout wewn\u0105trz ScrollView)
        # Przy bardzo du\u017cych plikach (5000+ wierszy) nale\u017cy u\u017cy\u0107 pe\u0142nego dict-data RV.
        for row in data_list[1:]:
            self.rv_layout.add_widget(TableRow(row_data=row, callback=self.export_single))

    def filter_data(self, ins, val):
        if not self.full_data: return
        filtered = [self.full_data[0]] + [r for r in self.full_data[1:] if any(val.lower() in str(c).lower() for c in r)]
        self.update_table_view(filtered)

    def export_single(self, row):
        threading.Thread(target=self._run_export, args=(self.full_data[0], row, False)).start()

    def _run_export(self, header, row, mass=False):
        try:
            # Scoped Storage / Documents
            folder = Path("/storage/emulated/0/Documents/FutureExport")
            folder.mkdir(parents=True, exist_ok=True)
            
            wb = Workbook(); ws = wb.active
            idxs = self.export_columns if self.export_columns else list(range(len(header)))
            
            ws.append([header[i] for i in idxs])
            ws.append([row[i] for i in idxs])
            
            # Formatowanie
            blue_fill = PatternFill(start_color='CFE2F3', end_color='CFE2F3', fill_type='solid')
            for cell in ws[1]: cell.fill, cell.font = blue_fill, Font(bold=True)
            
            name = self.sanitize_filename(row[0])[:20]
            filename = f"Raport_{name}_{datetime.now().strftime('%H%M%S')}.xlsx"
            dst = folder / filename
            wb.save(str(dst))
            
            if not mass: 
                Clock.schedule_once(lambda x: self.msg("OK", f"Zapisano w Documents/FutureExport"))
        except Exception as e: 
            Clock.schedule_once(lambda x: self.msg("B\u0142\u0105d Zapisu", str(e)))

    def start_mailing(self, _):
        threading.Thread(target=self._mailing_process).start()

    def _mailing_process(self):
        p_cfg = Path(self.user_data_dir) / "smtp.json"
        if not p_cfg.exists(): 
            Clock.schedule_once(lambda x: self.msg("B\u0142\u0105d", "Skonfiguruj SMTP!")); return
            
        with open(p_cfg, "r") as f: cfg = json.load(f)
        
        try:
            srv = smtplib.SMTP("smtp.gmail.com", 587, timeout=15)
            srv.starttls(); srv.login(cfg['u'], cfg['p'])
        except Exception as e:
            Clock.schedule_once(lambda x: self.msg("B\u0142\u0105d Logowania", str(e))); return

        h = self.full_data[0]
        # Pr\u00f3ba znalezienia kolumn ID
        ni = next((i for i, v in enumerate(h) if "imi" in str(v).lower()), 0)
        si = next((i for i, v in enumerate(h) if "nazw" in str(v).lower()), 1)
        
        sent = 0
        rows = self.full_data[1:]
        for i, r in enumerate(rows):
            name, sur = str(r[ni]).lower().strip(), str(r[si]).lower().strip()
            res = self.conn.execute("SELECT email FROM contacts WHERE name=? AND surname=?", (name, sur)).fetchone()
            
            if res:
                # Plik tymczasowy
                tmp = Path(self.user_data_dir) / "temp.xlsx"
                wb_t = Workbook(); ws_t = wb_t.active
                ws_t.append(h); ws_t.append(r)
                wb_t.save(str(tmp))
                
                msg = EmailMessage()
                msg["Subject"] = "Tw\u00f3j Raport Future"
                msg["From"] = cfg['u']
                msg["To"] = res[0]
                msg.set_content(f"Dzie\u0144 dobry {r[ni]}, w za\u0142\u0105czniku przesy\u0142amy raport.")
                
                with open(tmp, "rb") as f_att:
                    msg.add_attachment(f_att.read(), maintype="application", subtype="xlsx", filename=f"Raport_{name}.xlsx")
                
                srv.send_message(msg); sent += 1
            
            Clock.schedule_once(lambda dt, p=int(((i+1)/len(rows))*100): setattr(self.progress, "value", p))

        srv.quit()
        Clock.schedule_once(lambda x: self.msg("Koniec", f"Wys\u0142ano {sent} maili."))

    # --- HELPERS ---
    def msg(self, t, txt):
        b = BoxLayout(orientation="vertical", padding=dp(10))
        b.add_widget(Label(text=txt, halign="center"))
        btn = Button(text="OK", size_hint_y=None, height=dp(50))
        p = Popup(title=t, content=b, size_hint=(0.8, 0.4))
        btn.bind(on_press=p.dismiss); b.add_widget(btn); p.open()

    def test_smtp(self, _):
        def _t():
            try:
                srv = smtplib.SMTP("smtp.gmail.com", 587, timeout=10)
                srv.starttls(); srv.login(self.s_user.text, self.s_pass.text); srv.quit()
                Clock.schedule_once(lambda x: self.msg("Sukces", "Po\u0142\u0104czenie dzia\u0142a!"))
            except Exception as e: Clock.schedule_once(lambda x: self.msg("B\u0142\u0105d", str(e)))
        threading.Thread(target=_t).start()

    def save_smtp(self, _):
        with open(Path(self.user_data_dir) / "smtp.json", "w") as f:
            json.dump({'u': self.s_user.text, 'p': self.s_pass.text}, f)
        self.msg("OK", "Zapisano.")

    def load_smtp(self):
        p = Path(self.user_data_dir) / "smtp.json"
        if p.exists():
            with open(p, "r") as f:
                d = json.load(f); self.s_user.text = d['u']; self.s_pass.text = d['p']

    def column_popup(self, _):
        if not self.full_data: return
        box = BoxLayout(orientation="vertical", padding=dp(10))
        scroll = ScrollView(); grid = GridLayout(cols=1, size_hint_y=None); grid.bind(minimum_height=grid.setter('height'))
        checks = []
        for i, h in enumerate(self.full_data[0]):
            r = BoxLayout(size_hint_y=None, height=dp(40))
            cb = CheckBox(size_hint_x=0.2, active=True)
            r.add_widget(cb); r.add_widget(Label(text=str(h))); grid.add_widget(r); checks.append((i, cb))
        
        def apply(_): self.export_columns = [idx for idx, c in checks if c.active]; p.dismiss()
        scroll.add_widget(grid); box.add_widget(scroll)
        btn = PremiumButton(text="ZATWIERD\u0179", on_press=apply); box.add_widget(btn)
        p = Popup(title="Kolumny raportu", content=box, size_hint=(0.9, 0.9)); p.open()

    def start_export_all_thread(self, _):
        def _task():
            h, r = self.full_data[0], self.full_data[1:]
            for i, row in enumerate(r):
                self._run_export(h, row, mass=True)
                Clock.schedule_once(lambda dt, p=int(((i+1)/len(r))*100): setattr(self.progress, "value", p))
            Clock.schedule_once(lambda x: self.msg("Koniec", "Wszystkie raporty zapisane w Documents."))
        threading.Thread(target=_task).start()

if __name__ == "__main__":
    FutureApp().run()
