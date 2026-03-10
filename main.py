import os
import json
import threading
import time
from datetime import datetime
from pathlib import Path

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
from kivy.uix.screenmanager import ScreenManager, Screen
from kivy.uix.progressbar import ProgressBar
from kivy.graphics import Color, Rectangle

# --- KOLORY WERSJI 2.0 ---
COLOR_PRIMARY = (0.1, 0.5, 0.9, 1)
COLOR_BG = (0.08, 0.1, 0.15, 1)
COLOR_HEADER = (0.9, 0.9, 0.95, 1)
COLOR_ROW_A = (1, 1, 1, 1)
COLOR_ROW_B = (0.94, 0.97, 1, 1)

class PremiumButton(Button):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.background_normal = ""
        self.background_color = COLOR_PRIMARY
        self.height = dp(53)
        self.size_hint_y = None
        self.bold = True

class ColorSafeLabel(Label):
    def __init__(self, bg_color=(1,1,1,1), text_color=(0,0,0,1), **kwargs):
        super().__init__(**kwargs)
        self.color = text_color
        self.halign = 'center'
        self.valign = 'middle'
        with self.canvas.before:
            Color(*bg_color)
            self.rect = Rectangle(size=self.size, pos=self.pos)
        self.bind(size=self._update, pos=self._update)

    def _update(self, inst, val):
        self.rect.size = self.size
        self.rect.pos = self.pos
        self.text_size = (self.width - dp(10), None)

class HomeScreen(Screen): pass
class TableScreen(Screen): pass
class EmailScreen(Screen): pass
class SMTPScreen(Screen): pass
class TemplateScreen(Screen): pass
class ContactsScreen(Screen): pass
class ReportScreen(Screen): pass

class FutureApp(App):
    def build(self):
        # FIX DLA KLAWIATURY
        Window.softinput_mode = "below_target"
        Window.clearcolor = COLOR_BG
        
        self.full_data = []; self.filtered_data = []; self.export_indices = []
        self.global_attachments = []; self.selected_emails = []; self.queue = []
        self.stats = {"ok": 0, "fail": 0, "skip": 0, "auto": 0}
        self.idx_name, self.idx_surname, self.idx_pesel = 0, 1, -1
        self.auto_send_mode = False
        
        if not os.path.exists(self.user_data_dir): 
            os.makedirs(self.user_data_dir, exist_ok=True)
            
        self.init_db() # Lazy Import wewnątrz
        
        self.sm = ScreenManager()
        self.add_screens()
        return self.sm

    # --- LEAZY IMPORTS ---
    def init_db(self):
        import sqlite3
        db_p = Path(self.user_data_dir) / "future_ultimate_v2.db"
        self.conn = sqlite3.connect(str(db_p), check_same_thread=False)
        self.conn.execute("CREATE TABLE IF NOT EXISTS contacts (name TEXT, surname TEXT, email TEXT, pesel TEXT, phone TEXT, PRIMARY KEY(name, surname))")
        self.conn.execute("CREATE TABLE IF NOT EXISTS settings (key TEXT PRIMARY KEY, val TEXT)")
        self.conn.execute("CREATE TABLE IF NOT EXISTS reports (id INTEGER PRIMARY KEY AUTOINCREMENT, date TEXT, ok INTEGER, fail INTEGER, skip INTEGER, details TEXT)")
        self.conn.commit()

    def add_screens(self):
        self.screens = {
            "home": HomeScreen(name="home"), "table": TableScreen(name="table"),
            "email": EmailScreen(name="email"), "smtp": SMTPScreen(name="smtp"),
            "tmpl": TemplateScreen(name="tmpl"), "contacts": ContactsScreen(name="contacts"),
            "report": ReportScreen(name="report")
        }
        self.setup_ui()
        for s in self.screens.values(): self.sm.add_widget(s)

    # --- ANDROID PICKER LOGIC (LAZY JNIUS) ---
    def open_picker(self, mode):
        if platform != "android": self.msg("!", "Funkcja dostępna tylko na Android"); return
        from jnius import autoclass
        from android import activity
        PA = autoclass("org.kivy.android.PythonActivity"); Intent = autoclass("android.content.Intent")
        intent = Intent(Intent.ACTION_GET_CONTENT); intent.setType("*/*")
        
        def cb(req, res, dt):
            if req != 1001: return
            activity.unbind(on_activity_result=cb)
            if res == -1 and dt:
                uri = dt.getData(); resolver = PA.mActivity.getContentResolver()
                d_name = f"plik_{datetime.now().strftime('%H%M%S')}.xlsx"
                try:
                    cur = resolver.query(uri, None, None, None, None)
                    if cur and cur.moveToFirst():
                        idx = cur.getColumnIndex("_display_name")
                        if idx != -1: d_name = cur.getString(idx)
                        cur.close()
                    stream = resolver.openInputStream(uri); loc = Path(self.user_data_dir) / d_name
                    with open(loc, "wb") as f:
                        buf = bytearray(16384)
                        while True:
                            n = stream.read(buf)
                            if n <= 0: break
                            f.write(buf[:n])
                    stream.close()
                    if mode == "data": self.process_excel(loc)
                    elif mode == "book": self.process_book(loc)
                    elif mode == "attachment": self.global_attachments.append(str(loc)); self.update_stats()
                    elif mode == "special_send": Clock.schedule_once(lambda dt: self.special_send_step_2_recipients(str(loc)))
                except: self.msg("Błąd", "Wystąpił błąd pliku.")
        
        activity.bind(on_activity_result=cb)
        PA.mActivity.startActivityForResult(intent, 1001)

    # --- MASOWA WYSYŁKA 2.0 (LAZY SMTP + OPENPYXL) ---
    def start_mass_mailing(self, _):
        if not self.full_data: self.msg("!", "Wczytaj arkusz płac!"); return
        self.stats = {"ok": 0, "fail": 0, "skip": 0, "auto": 0}
        self.queue = list(self.full_data[1:])
        self.total_q = len(self.queue)
        self.pb.value = 0
        Clock.schedule_once(self.process_mailing_queue)

    def process_mailing_queue(self, *args):
        done = self.total_q - len(self.queue)
        self.pb.value = int((done/self.total_q)*100) if self.total_q > 0 else 100
        self.pb_label.text = f"Postęp: {self.pb.value}% ({done}/{self.total_q})"
        
        if not self.queue:
            self.conn.execute("INSERT INTO reports (date, ok, fail, skip) VALUES (?,?,?,?)", 
                              (datetime.now().strftime("%Y-%m-%d %H:%M"), self.stats['ok'], self.stats['fail'], self.stats['skip']))
            self.conn.commit()
            self.msg("Koniec", "Operacja zakończona."); return

        row = self.queue.pop(0)
        try:
            n, s = str(row[self.idx_name]).strip(), str(row[self.idx_surname])
            p = str(row[self.idx_pesel]).strip() if self.idx_pesel != -1 else ""
        except: self.stats["skip"] += 1; Clock.schedule_once(self.process_mailing_queue); return
        
        # Szukanie kontaktu w bazie
        res = self.conn.execute("SELECT email FROM contacts WHERE name=? AND surname=? COLLATE NOCASE", (n.lower(), s.lower())).fetchone()
        if not res and p: res = self.conn.execute("SELECT email FROM contacts WHERE pesel=?", (p,)).fetchone()
        
        if res:
            if self.auto_send_mode:
                self.send_email_engine(row, res[0])
            else:
                self.ask_before_send(row, res[0], n, s)
        else:
            self.stats["skip"] += 1
            Clock.schedule_once(self.process_mailing_queue)

    def send_email_engine(self, row_data, target):
        def thread_task():
            import smtplib
            import mimetypes
            from openpyxl import Workbook
            from email.message import EmailMessage
            
            cfg_p = Path(self.user_data_dir) / "smtp.json"
            if not cfg_p.exists(): return
            cfg = json.load(open(cfg_p))
            
            try:
                srv = smtplib.SMTP("smtp.gmail.com", 587, timeout=15); srv.starttls(); srv.login(cfg['u'], cfg['p'])
                nx, sx = str(row_data[self.idx_name]).title(), str(row_data[self.idx_surname]).title()
                
                # Szablon
                ts = self.conn.execute("SELECT val FROM settings WHERE key='t_sub'").fetchone()
                tb = self.conn.execute("SELECT val FROM settings WHERE key='t_body'").fetchone()
                
                msg = EmailMessage()
                msg["Subject"] = (ts[0] if ts else "Wiadomość").replace("{Imię}", nx)
                msg["From"], msg["To"] = cfg['u'], target
                msg.set_content((tb[0] if tb else "Dzień dobry").replace("{Imię}", nx))
                
                # Załącznik Excel (Pasek)
                tmp = Path(self.user_data_dir) / "r_tmp.xlsx"
                wb = Workbook(); ws = wb.active
                ws.append([self.full_data[0][k] for k in self.export_indices])
                ws.append([row_data[k] for k in self.export_indices])
                self.style_xlsx(ws); wb.save(tmp)
                
                with open(tmp, "rb") as f:
                    msg.add_attachment(f.read(), maintype="application", subtype="xlsx", filename=f"Raport_{nx}.xlsx")
                
                # Załączniki globalne
                for path in self.global_attachments:
                    if os.path.exists(path):
                        ct, _ = mimetypes.guess_type(path)
                        mn, sb = (ct or 'application/octet-stream').split('/', 1)
                        with open(path, "rb") as f:
                            msg.add_attachment(f.read(), maintype=mn, subtype=sb, filename=os.path.basename(path))
                
                srv.send_message(msg); srv.quit()
                Clock.schedule_once(lambda d: [self.update_stat("ok"), self.process_mailing_queue()])
            except:
                Clock.schedule_once(lambda d: [self.update_stat("fail"), self.process_mailing_queue()])
        
        threading.Thread(target=thread_task, daemon=True).start()

    # --- STYLE & UTILS (2.0) ---
    def style_xlsx(self, ws):
        from openpyxl.styles import Border, Side, Font, Alignment, PatternFill
        s, c = Side(style='thin'), Alignment(horizontal='center', vertical='center')
        for ri, row in enumerate(ws.iter_rows(), 1):
            for cell in row:
                cell.border = Border(top=s, left=s, right=s, bottom=s); cell.alignment = c
                if ri == 1: 
                    cell.font = Font(bold=True); cell.fill = PatternFill("solid", start_color="DDEBF7")
                elif ri % 2 == 0:
                    cell.fill = PatternFill("solid", start_color="F7F7F7")
        for col in ws.columns: ws.column_dimensions[col[0].column_letter].width = 18

    # --- MAIN UI SETUP (STYL WERSJI 2.0) ---
    def setup_ui(self):
        l = BoxLayout(orientation="vertical", padding=dp(30), spacing=dp(15))
        l.add_widget(Label(text="FUTURE 10.1 ULTIMATE", font_size='26sp', bold=True, color=COLOR_PRIMARY))
        btn = lambda t, c: l.add_widget(PremiumButton(text=t, on_press=c))
        btn("WCZYTAJ ARKUSZ PŁAC", lambda x: self.open_picker("data"))
        btn("PODGLĄD I EKSPORT", lambda x: [self.refresh_table(), setattr(self.sm, 'current', 'table')] if self.full_data else self.msg("!", "Brak danych"))
        btn("CENTRUM MAILINGOWE", lambda x: setattr(self.sm, 'current', 'email'))
        btn("RAPORTY WYSYŁEK", lambda x: [self.refresh_reports(), setattr(self.sm, 'current', 'report')])
        btn("USTAWIENIA SMTP", lambda x: setattr(self.sm, 'current', 'smtp'))
        self.screens["home"].add_widget(l)
        self.setup_table_ui(); self.setup_email_ui(); self.setup_smtp_ui(); self.setup_tmpl_ui(); self.setup_contacts_ui(); self.setup_report_ui()

    def setup_email_ui(self):
        l = BoxLayout(orientation="vertical", padding=dp(25), spacing=dp(10))
        l.add_widget(Label(text="CENTRUM MAILINGOWE", font_size='22sp', bold=True))
        auto_box = BoxLayout(size_hint_y=None, height=dp(45), spacing=dp(10))
        self.cb_auto = CheckBox(size_hint_x=None, width=dp(45))
        self.cb_auto.bind(active=lambda i, v: setattr(self, 'auto_send_mode', v))
        auto_box.add_widget(self.cb_auto)
        auto_box.add_widget(Label(text="AUTO-WYSYŁKA", halign="left", text_size=(dp(280), None)))
        l.add_widget(auto_box)
        self.lbl_stats = Label(text="Baza: 0", size_hint_y=None, height=dp(30)); l.add_widget(self.lbl_stats)
        self.pb_label = Label(text="Gotowy", size_hint_y=None, height=dp(25)); self.pb = ProgressBar(max=100, size_hint_y=None, height=dp(20))
        l.add_widget(self.pb_label); l.add_widget(self.pb)
        btn = lambda t, c: l.add_widget(PremiumButton(text=t, on_press=c))
        btn("IMPORT KONTAKTÓW (EXCEL)", lambda x: self.open_picker("book"))
        btn("ZARZĄDZAJ BAZĄ", lambda x: [self.refresh_contacts_list(), setattr(self.sm, 'current', 'contacts')])
        btn("EDYTUJ SZABLON WIADOMOŚCI", lambda x: setattr(self.sm, 'current', 'tmpl'))
        btn("DODAJ ZAŁĄCZNIK GLOBALNY", lambda x: self.open_picker("attachment"))
        btn("START MASOWA WYSYŁKA", self.start_mass_mailing)
        btn("POWRÓT", lambda x: setattr(self.sm, 'current', 'home'))
        self.screens["email"].add_widget(l); self.update_stats()

    # --- REZULTAT ZE ZDJĘCIA (NAPRAWA PÓL FORMULARZA) ---
    def form_contact(self, n="", s="", e="", pes="", ph=""):
        # Używamy ScrollView, aby klawiatura nie zasłaniała pól
        root = BoxLayout(orientation="vertical", padding=dp(15), spacing=dp(10))
        sc = ScrollView()
        box = BoxLayout(orientation="vertical", size_hint_y=None, spacing=dp(10), padding=[0,0,0,dp(20)])
        box.bind(minimum_height=box.setter('height'))
        
        # FIX: Dodajemy stałą wysokość dp(45) dla każdego pola
        self.flds = []
        hints = ["Imię", "Nazwisko", "Email", "PESEL", "Telefon"]
        vals = [n, s, e, pes, ph]
        
        for i in range(5):
            ti = TextInput(text=vals[i], hint_text=hints[i], multiline=False, size_hint_y=None, height=dp(45))
            self.flds.append(ti)
            box.add_widget(ti)
            
        sc.add_widget(box)
        root.add_widget(sc)
        
        btn_save = PremiumButton(text="ZAPISZ")
        btn_save.bind(on_press=lambda x: self.save_contact_db(px))
        root.add_widget(btn_save)
        
        px = Popup(title="Dane Kontaktu", content=root, size_hint=(0.95, 0.9))
        px.open()

    def save_contact_db(self, popup):
        d = [f.text.strip() for f in self.flds]
        self.conn.execute("INSERT OR REPLACE INTO contacts (name, surname, email, pesel, phone) VALUES (?,?,?,?,?)", 
                          (d[0].lower(), d[1].lower(), d[2], d[3], d[4]))
        self.conn.commit(); popup.dismiss(); self.refresh_contacts_list(); self.update_stats()

    # --- POZOSTAŁE FUNKCJE (2.0) ---
    def process_excel(self, path):
        from openpyxl import load_workbook
        try:
            wb = load_workbook(path, data_only=True); ws = wb.active
            raw = [["" if v is None else str(v).strip() for v in r] for r in ws.iter_rows(values_only=True)]
            self.full_data = raw; self.filtered_data = raw
            self.export_indices = list(range(len(raw[0])))
            h = [str(x).lower() for x in raw[0]]
            for i,v in enumerate(h):
                if "imi" in v: self.idx_name = i
                if "naz" in v: self.idx_surname = i
                if "pesel" in v: self.idx_pesel = i
            self.msg("OK", "Wczytano arkusz.")
        except Exception as e: self.msg("Błąd", str(e))

    def setup_smtp_ui(self):
        l = BoxLayout(orientation="vertical", padding=dp(25), spacing=dp(10))
        self.ti_su = TextInput(hint_text="Gmail (np. biuro@gmail.com)", multiline=False, size_hint_y=None, height=dp(45))
        self.ti_sp = TextInput(hint_text="Hasło aplikacji", password=True, size_hint_y=None, height=dp(45))
        p = Path(self.user_data_dir) / "smtp.json"
        if p.exists():
            d = json.load(open(p)); self.ti_su.text, self.ti_sp.text = d.get('u',''), d.get('p','')
        sv = lambda x: [json.dump({'u':self.ti_su.text, 'p':self.ti_sp.text}, open(p, "w")), self.msg("OK", "Zapisano")]
        l.add_widget(Label(text="USTAWIENIA GMAIL", bold=True))
        l.add_widget(self.ti_su); l.add_widget(self.ti_sp)
        l.add_widget(PremiumButton(text="ZAPISZ", on_press=sv))
        l.add_widget(PremiumButton(text="POWRÓT", on_press=lambda x: setattr(self.sm, 'current', 'home')))
        self.screens["smtp"].add_widget(l)

    def refresh_contacts_list(self, *args):
        self.c_list.clear_widgets(); rows = self.conn.execute("SELECT name, surname, email, pesel, phone FROM contacts ORDER BY surname ASC").fetchall()
        for n, s, e, p, ph in rows:
            row = BoxLayout(size_hint_y=None, height=dp(80), padding=dp(5)); info = Label(text=f"{n.title()} {s.title()}\n{e}", font_size='13sp')
            row.add_widget(info); btn_ed = Button(text="Edytuj", size_hint_x=0.2); btn_ed.bind(on_press=lambda x, d=(n,s,e,p,ph): self.form_contact(*d))
            row.add_widget(btn_ed); self.c_list.add_widget(row)

    def setup_contacts_ui(self):
        l = BoxLayout(orientation="vertical", padding=dp(10)); self.c_list = GridLayout(cols=1, size_hint_y=None, spacing=dp(10))
        self.c_list.bind(minimum_height=self.c_list.setter('height')); sc = ScrollView(); sc.add_widget(self.c_list)
        l.add_widget(PremiumButton(text="+ DODAJ KONTAKT", on_press=lambda x: self.form_contact()))
        l.add_widget(sc); l.add_widget(PremiumButton(text="WRÓĆ", on_press=lambda x: setattr(self.sm, 'current', 'email')))
        self.screens["contacts"].add_widget(l)
    
    def setup_tmpl_ui(self):
        l = BoxLayout(orientation="vertical", padding=dp(25), spacing=dp(10))
        self.ti_ts = TextInput(hint_text="Temat {Imię}", size_hint_y=None, height=dp(45))
        self.ti_tb = TextInput(hint_text="Treść maila...", multiline=True)
        ts = self.conn.execute("SELECT val FROM settings WHERE key='t_sub'").fetchone(); tb = self.conn.execute("SELECT val FROM settings WHERE key='t_body'").fetchone()
        self.ti_ts.text = ts[0] if ts else ""; self.ti_tb.text = tb[0] if tb else ""
        save = lambda x: [self.conn.execute("INSERT OR REPLACE INTO settings VALUES (?,?)", ('t_sub', self.ti_ts.text)), self.conn.execute("INSERT OR REPLACE INTO settings VALUES (?,?)", ('t_body', self.ti_tb.text)), self.conn.commit(), self.msg("OK", "Zapisano")]
        l.add_widget(Label(text="SZABLON MAILA", bold=True)); l.add_widget(self.ti_ts); l.add_widget(self.ti_tb)
        l.add_widget(PremiumButton(text="ZAPISZ", on_press=save)); l.add_widget(PremiumButton(text="POWRÓT", on_press=lambda x: setattr(self.sm, 'current', 'email')))
        self.screens["tmpl"].add_widget(l)

    def msg(self, tit, txt):
        b = BoxLayout(orientation="vertical", padding=dp(20)); b.add_widget(Label(text=txt)); btn = Button(text="OK", size_hint_y=None, height=dp(50), on_press=lambda x: p.dismiss()); b.add_widget(btn); p = Popup(title=tit, content=b, size_hint=(0.85, 0.45)); p.open()
    def update_stat(self, k): self.stats[k]+=1
    def update_stats(self, *a): 
        try: 
            c = self.conn.execute("SELECT count(*) FROM contacts").fetchone()[0]
            self.lbl_stats.text = f"Baza: {c} kontaktów | Załączniki: {len(self.global_attachments)}"
        except: pass

    # Stuby dla Table i Report (aby kod był kompletny)
    def setup_table_ui(self):
        s = self.screens["table"]; l = BoxLayout(orientation="vertical")
        self.grid = GridLayout(rows=1, size_hint=(None, None), height=dp(50)); sc = ScrollView(); sc.add_widget(self.grid)
        l.add_widget(sc); l.add_widget(PremiumButton(text="WRÓĆ", on_press=lambda x: setattr(self.sm, 'current', 'home')))
        s.add_widget(l)

    def refresh_table(self): pass
    def setup_report_ui(self):
        s = self.screens["report"]; l = BoxLayout(orientation="vertical")
        l.add_widget(Label(text="Historia raportów (Wkrótce)")); l.add_widget(PremiumButton(text="WRÓĆ", on_press=lambda x: setattr(self.sm, 'current', 'home')))
        s.add_widget(l)
    def refresh_reports(self): pass

if __name__ == "__main__":
    FutureApp().run()
