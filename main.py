import os
import json
import sqlite3
import threading
import smtplib
import mimetypes
import shutil
import time
import random
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
from kivy.uix.screenmanager import ScreenManager, Screen, SlideTransition # Dodano SlideTransition
from kivy.uix.progressbar import ProgressBar
from kivy.graphics import Color, Rectangle, RoundedRectangle # Dodano RoundedRectangle

# Obsługa Excel (openpyxl + xlrd)
try:
    from openpyxl import load_workbook, Workbook
    from openpyxl.styles import Border, Side, Font, Alignment, PatternFill
except ImportError:
    load_workbook = Workbook = None

try:
    import xlrd
except ImportError:
    xlrd = None

# Paleta kolorów
COLOR_PRIMARY = (0.1, 0.5, 0.9, 1)
COLOR_BG = (0.08, 0.1, 0.15, 1) # Zmieniono na ciemniejszy kolor tła
COLOR_CARD = (0.12, 0.15, 0.2, 1) # Nowy kolor dla kart/kontaktów
COLOR_TEXT = (0.95, 0.95, 0.95, 1) # Jasny tekst
COLOR_ROW_A = (0.08, 0.1, 0.15, 1) # Kolor wiersza A (taki jak tło)
COLOR_ROW_B = (0.13, 0.16, 0.22, 1) # Kolor wiersza B
COLOR_HEADER = (0.1, 0.2, 0.35, 1) # Kolor nagłówka tabeli

class ModernButton(Button): # Nowa klasa przycisku z zaokrąglonymi rogami
    def __init__(self, bg_color=COLOR_PRIMARY, **kwargs):
        super().__init__(**kwargs)
        self.background_normal = ""
        self.background_color = (0,0,0,0) # Przezroczyste tło, rysujemy własne
        self.color = COLOR_TEXT
        self.bold, self.radius = True, [dp(12)]
        with self.canvas.before:
            Color(*bg_color)
            self.rect = RoundedRectangle(pos=self.pos, size=self.size, radius=self.radius)
        self.bind(pos=self._update, size=self._update)
    def _update(self, *args):
        self.rect.pos, self.rect.size = self.pos, self.size

class ModernInput(TextInput): # Nowa klasa TextInput
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.background_normal = self.background_active = ""
        self.background_color = (0.15, 0.18, 0.25, 1) # Ciemniejsze tło inputa
        self.foreground_color = COLOR_TEXT
        self.padding = [dp(12), dp(12)] # Większy padding

class ColorSafeLabel(Label):
    def __init__(self, bg_color=(1,1,1,1), text_color=(1,1,1,1), **kwargs): # Domyślny text_color zmieniony na biały
        super().__init__(**kwargs)
        self.color = text_color
        self.halign = 'left' # Domyślne wyrównanie do lewej
        self.valign = 'middle'
        with self.canvas.before:
            Color(*bg_color)
            self.rect = Rectangle(size=self.size, pos=self.pos)
        self.bind(size=self._update, pos=self._update)

    def _update(self, inst, val):
        self.rect.size = self.size
        self.rect.pos = self.pos
        self.text_size = (self.width - dp(10), None)

# Usunięto PremiumButton i dedykowane klasy Screen, używamy Screen(name=...)

class FutureApp(App):
    def build(self):
        Window.clearcolor = COLOR_BG
        self.full_data = []; self.filtered_data = []; self.export_indices = []
        self.global_attachments = []; self.selected_emails = []; self.queue = []
        self.stats = {"ok": 0, "fail": 0, "skip": 0, "auto": 0}
        self.idx_name, self.idx_surname, self.idx_pesel = 0, 1, -1
        self.auto_send_mode = False
        
        if not os.path.exists(self.user_data_dir): 
            os.makedirs(self.user_data_dir, exist_ok=True)
        self.init_db()
        
        if platform == "android":
            from android.permissions import request_permissions, Permission
            request_permissions([Permission.READ_EXTERNAL_STORAGE, Permission.WRITE_EXTERNAL_STORAGE])

        self.sm = ScreenManager(transition=SlideTransition()) # Dodano SlideTransition
        self.add_screens()
        return self.sm

    def init_db(self):
        db_p = Path(self.user_data_dir) / "future_ultimate_v2.db"
        self.conn = sqlite3.connect(str(db_p), check_same_thread=False)
        self.conn.execute("CREATE TABLE IF NOT EXISTS contacts (name TEXT, surname TEXT, email TEXT, pesel TEXT, phone TEXT, PRIMARY KEY(name, surname))")
        self.conn.execute("CREATE TABLE IF NOT EXISTS settings (key TEXT PRIMARY KEY, val TEXT)")
        self.conn.execute("CREATE TABLE IF NOT EXISTS reports (id INTEGER PRIMARY KEY AUTOINCREMENT, date TEXT, ok INTEGER, fail INTEGER, skip INTEGER, auto INTEGER, details TEXT)") # Dodano details
        self.conn.commit()

    def add_screens(self):
        # Rozszerzono listę ekranów
        self.sc_ref = {name: Screen(name=name) for name in ["home", "table", "email", "smtp", "tmpl", "contacts", "report", "paski", "settings"]}
        self.setup_ui_all(); [self.sm.add_widget(s) for s in self.sc_ref.values()]

    # --- ANDROID PICKER LOGIC ---
    def open_picker(self, mode):
        if platform != "android": self.msg("!", "Funkcja dostępna tylko na Android"); return
        from jnius import autoclass; from android import activity
        PA = autoclass("org.kivy.android.PythonActivity"); Intent = autoclass("android.content.Intent")
        intent = Intent(Intent.ACTION_GET_CONTENT); intent.setType("*/*")
        def cb(req, res, dt):
            if req != 1001: return
            activity.unbind(on_activity_result=cb)
            if res == -1 and dt:
                uri = dt.getData(); resolver = PA.mActivity.getContentResolver()
                cur = resolver.query(uri, None, None, None, None); d_name = f"plik_{datetime.now().strftime('%H%M%S')}.xlsx"
                if cur and cur.moveToFirst():
                    idx = cur.getColumnIndex("_display_name")
                    if idx != -1: d_name = cur.getString(idx)
                    cur.close()
                try:
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
                except Exception as e: self.msg("Błąd", f"Wystąpił błąd pliku: {e}")
        activity.bind(on_activity_result=cb); PA.mActivity.startActivityForResult(intent, 1001)

    # --- KREATOR WYSYŁKI SPECJALNEJ (4 KROKI) ---
    def start_special_send_flow(self, _):
        self.open_picker("special_send")

    def special_send_step_2_recipients(self, file_path):
        self.selected_emails = []
        box = BoxLayout(orientation="vertical", padding=dp(15), spacing=dp(10))
        box.add_widget(Label(text="KROK 2: WYBIERZ ODBIORCÓW", bold=True, color=COLOR_PRIMARY, size_hint_y=None, height=dp(40)))
        ti = ModernInput(hint_text="Szukaj kontaktu..."); box.add_widget(ti) # Użycie ModernInput
        sc = ScrollView(); gl = GridLayout(cols=1, size_hint_y=None, spacing=dp(5)); gl.bind(minimum_height=gl.setter('height')); sc.add_widget(gl); box.add_widget(sc)
        def refresh(v=""):
            gl.clear_widgets(); rows = self.conn.execute("SELECT name, surname, email FROM contacts").fetchall()
            for n, s, e in rows:
                if v and v.lower() not in f"{n} {s} {e}".lower(): continue
                r = BoxLayout(size_hint_y=None, height=dp(55))
                cb = CheckBox(size_hint_x=None, width=dp(50), active=(e in self.selected_emails))
                def update_sel(inst, val, mail=e):
                    if val: 
                        if mail not in self.selected_emails: self.selected_emails.append(mail)
                    else:
                        if mail in self.selected_emails: self.selected_emails.remove(mail)
                cb.bind(active=update_sel)
                r.add_widget(cb); r.add_widget(Label(text=f"{n.title()} {s.title()}\n{e}", halign="left", text_size=(dp(250), None), font_size='12sp')); gl.add_widget(r)
        ti.bind(text=lambda i,v: refresh(v)); refresh()
        btn = ModernButton(text="DALEJ (KROK 3)"); btn.bind(on_press=lambda x: [p.dismiss(), self.special_send_step_3_msg(file_path)] if self.selected_emails else self.msg("!", "Wybierz kogoś!")); box.add_widget(btn); p = Popup(title="Lista odbiorców", content=box, size_hint=(0.95,0.9)); p.open()

    def special_send_step_3_msg(self, file_path):
        box = BoxLayout(orientation="vertical", padding=dp(15), spacing=dp(10))
        box.add_widget(Label(text=f"KROK 3: TREŚĆ MAILA ({len(self.selected_emails)})", color=COLOR_PRIMARY, size_hint_y=None, height=dp(40)))
        ti_s = ModernInput(hint_text="Temat maila"); ti_b = ModernInput(hint_text="Treść wiadomości...", multiline=True) # Użycie ModernInput
        box.add_widget(ti_s); box.add_widget(ti_b); btn = ModernButton(text="WYŚLIJ (KROK 4)")
        btn.bind(on_press=lambda x: [p.dismiss(), self.special_send_step_4_progress(file_path, self.selected_emails, ti_s.text, ti_b.text)] if ti_s.text and ti_b.text else self.msg("!", "Wymagane dane!")); box.add_widget(btn); p = Popup(title="Edycja wiadomości", content=box, size_hint=(0.95,0.85)); p.open()

    def special_send_step_4_progress(self, file_path, target_list, subject, body):
        box = BoxLayout(orientation="vertical", padding=dp(20), spacing=dp(15))
        lbl = Label(text="Inicjalizacja..."); pb = ProgressBar(max=len(target_list), value=0, size_hint_y=None, height=dp(30))
        box.add_widget(lbl); box.add_widget(pb); btn_c = ModernButton(text="ZAMKNIJ", size_hint_y=None, height=dp(50), disabled=True, bg_color=(0.3,0.3,0.3,1)) # Użycie ModernButton
        p = Popup(title="Proces wysyłki", content=box, size_hint=(0.85, 0.45), auto_dismiss=False); btn_c.bind(on_press=p.dismiss); box.add_widget(btn_c); p.open()
        def run():
            cfg_p = Path(self.user_data_dir) / "smtp.json"
            if not cfg_p.exists(): Clock.schedule_once(lambda dt: self.msg("!", "Brak SMTP"), 0); p.dismiss(); return
            cfg = json.load(open(cfg_p)); ok, err = 0, 0
            exp_dir = Path("/storage/emulated/0/Documents/FutureExport") if platform=="android" else Path("./exports")
            exp_dir.mkdir(parents=True, exist_ok=True)
            try:
                srv = smtplib.SMTP("smtp.gmail.com", 587, timeout=15); srv.starttls(); srv.login(cfg['u'], cfg['p'])
                for i, email in enumerate(target_list):
                    try:
                        Clock.schedule_once(lambda dt: setattr(lbl, 'text', f"Wysyłka: {email}"), 0)
                        info = self.conn.execute("SELECT name, surname FROM contacts WHERE email=?", (email,)).fetchone()
                        base = f"{info[0].title()} {info[1].title()}" if info else email.split('@')[0]
                        shutil.copy(file_path, exp_dir / f"Spec [{base}] {os.path.basename(file_path)}")
                        msg = EmailMessage(); msg["Subject"] = subject; msg["From"], msg["To"] = cfg['u'], email; msg.set_content(body)
                        with open(file_path, "rb") as f:
                            ct, _ = mimetypes.guess_type(file_path); mn, sb = (ct or 'application/octet-stream').split('/', 1)
                            msg.add_attachment(f.read(), maintype=mn, subtype=sb, filename=os.path.basename(file_path))
                        srv.send_message(msg); ok += 1
                    except Exception as e: 
                        err += 1
                        Clock.schedule_once(lambda dt: self.msg("Błąd wysyłki", f"Do: {email}\nBłąd: {e}"), 0)
                    Clock.schedule_once(lambda dt, idx=i+1: setattr(pb, 'value', idx), 0)
                srv.quit()
                Clock.schedule_once(lambda dt: [setattr(lbl, 'text', f"KONIEC\\nSukces: {ok} | Błąd: {err}"), setattr(btn_c, 'disabled', False)], 0)
                self.conn.execute("INSERT INTO reports (date, ok, fail, skip, auto, details) VALUES (?,?,?,?,?,?)", (f"{datetime.now().strftime('%Y-%m-%d %H:%M')} (PLIK)", ok, err, 0, 0, f"Wysłano {ok} z {len(target_list)}")); self.conn.commit()
            except Exception as e: Clock.schedule_once(lambda dt: [setattr(lbl, 'text', f"Serwer: {str(e)}"), setattr(btn_c, 'disabled', False)], 0)
        threading.Thread(target=run, daemon=True).start()

    # --- MAIN UI SETUP ---
    def setup_ui_all(self): # Zmieniono nazwę na setup_ui_all
        # HOME SCREEN
        l = BoxLayout(orientation="vertical", padding=dp(30), spacing=dp(15))
        l.add_widget(Label(text="FUTURE ULTIMATE v15", font_size='34sp', bold=True, color=COLOR_PRIMARY))
        grid = GridLayout(cols=2, spacing=dp(15), padding=dp(10))
        grid.add_widget(ModernButton(text="Kontakty", on_press=lambda x: [self.refresh_contacts_list(), setattr(self.sm, 'current', 'contacts')], size_hint=(1,1)))
        grid.add_widget(ModernButton(text="Paski", on_press=lambda x: setattr(self.sm, 'current', 'paski'), size_hint=(1,1))) # Nowy przycisk Paski
        grid.add_widget(ModernButton(text="Centrum Mailingowe", on_press=lambda x: setattr(self.sm, 'current', 'email'), size_hint=(1,1)))
        grid.add_widget(ModernButton(text="Raporty", on_press=lambda x: [self.refresh_reports(), setattr(self.sm, 'current', 'report')], size_hint=(1,1)))
        grid.add_widget(ModernButton(text="Ustawienia", on_press=lambda x: setattr(self.sm, 'current', 'settings'), size_hint=(1,1))) # Nowy przycisk Ustawienia
        grid.add_widget(ModernButton(text="Wyjście", on_press=lambda x: App.get_running_app().stop(), size_hint=(1,1), bg_color=(0.6,0.1,0.1,1)))
        l.add_widget(grid)
        self.sc_ref["home"].add_widget(l)

        # Inicjalizacja wszystkich ekranów
        self.setup_table_ui()
        self.setup_email_ui()
        self.setup_smtp_ui()
        self.setup_tmpl_ui()
        self.setup_contacts_ui()
        self.setup_report_ui()
        self.setup_paski_ui() # Nowa inicjalizacja ekranu Paski
        self.setup_settings_ui() # Nowa inicjalizacja ekranu Ustawienia

    def setup_email_ui(self):
        l = BoxLayout(orientation="vertical", padding=dp(25), spacing=dp(10))
        l.add_widget(Label(text="CENTRUM MAILINGOWE", font_size='22sp', bold=True))
        auto_box = BoxLayout(size_hint_y=None, height=dp(45), spacing=dp(10))
        self.cb_auto = CheckBox(size_hint_x=None, width=dp(45))
        self.cb_auto.bind(active=lambda i, v: setattr(self, 'auto_send_mode', v))
        auto_box.add_widget(self.cb_auto)
        auto_box.add_widget(Label(text="AUTO-WYSYŁKA (POMIŃ PESEL I POTWIERDZENIA)", halign="left", text_size=(dp(280), None), font_size='12sp', bold=True))
        l.add_widget(auto_box)
        self.lbl_stats = Label(text="Baza: 0 | Załączniki: 0", size_hint_y=None, height=dp(30)); l.add_widget(self.lbl_stats) # Ustawiono domyślny tekst
        self.pb_label = Label(text="Gotowy", size_hint_y=None, height=dp(25)); self.pb = ProgressBar(max=100, size_hint_y=None, height=dp(20))
        l.add_widget(self.pb_label); l.add_widget(self.pb)
        btn = lambda t, c: l.add_widget(ModernButton(text=t, on_press=c)) # Użycie ModernButton
        btn("IMPORT KONTAKTÓW (EXCEL)", lambda x: self.open_picker("book"))
        btn("ZARZĄDZAJ BAZĄ", lambda x: [self.refresh_contacts_list(), setattr(self.sm, 'current', 'contacts')])
        btn("EDYTUJ SZABLON WIADOMOŚCI", lambda x: setattr(self.sm, 'current', 'tmpl'))
        btn("DODAJ ZAŁĄCZNIK GLOBALNY", lambda x: self.open_picker("attachment"))
        btn("WYŚLIJ PLIK", self.start_special_send_flow)
        btn("START MASOWA WYSYŁKA", self.start_mass_mailing)
        btn("POWRÓT", lambda x: setattr(self.sm, 'current', 'home'), bg_color=(0.3,0.3,0.3,1)) # Użycie ModernButton
        self.sc_ref["email"].add_widget(l); self.update_stats()

    # --- MASOWA WYSYŁKA (AUTO TRYB) ---
    def start_mass_mailing(self, _):
        if not self.full_data: self.msg("!", "Wczytaj arkusz płac!"); return
        self.stats = {"ok": 0, "fail": 0, "skip": 0, "auto": 0}; self.queue = list(self.full_data[1:]); self.total_q = len(self.queue)
        self.pb.value = 0; Clock.schedule_once(self.process_mailing_queue, 0)

    def process_mailing_queue(self, *args):
        self.update_progress() # Wywołanie uogólnionej funkcji
        if not self.queue:
            self.conn.execute("INSERT INTO reports (date, ok, fail, skip, auto, details) VALUES (?,?,?,?,?,?)", (datetime.now().strftime("%Y-%m-%d %H:%M"), self.stats['ok'], self.stats['fail'], self.stats['skip'], self.stats['auto'], f"Wysłano {self.stats['ok']} z {self.total_q}")); self.conn.commit()
            self.msg("Koniec", "Operacja zakończona."); return
        row = self.queue.pop(0)
        try:
            n, s = str(row[self.idx_name]).strip(), str(row[self.idx_surname]).strip(); p = str(row[self.idx_pesel]).strip() if self.idx_pesel != -1 else ""
        except: self.stats["skip"] += 1; Clock.schedule_once(self.process_mailing_queue, 0); return
        
        if self.auto_send_mode:
            res = self.conn.execute("SELECT email FROM contacts WHERE name=? AND surname=?", (n.lower(), s.lower())).fetchone()
            if res: self.stats["auto"] += 1; self.send_email_engine(row, res[0]); return
            else: self.stats["skip"] += 1; Clock.schedule_once(self.process_mailing_queue, 0); return

        if p and len(p) > 5:
            res = self.conn.execute("SELECT email FROM contacts WHERE pesel=?", (p,)).fetchone()
            if res: self.stats["auto"] += 1; self.send_email_engine(row, res[0]); return
        
        res = self.conn.execute("SELECT email FROM contacts WHERE name=? AND surname=?", (n.lower(), s.lower())).fetchone()
        if res: self.ask_before_send(row, res[0], n, s, p)
        else: self.stats["skip"] += 1; Clock.schedule_once(self.process_mailing_queue, 0)

    def ask_before_send(self, row, email, n, s, p_file):
        box = BoxLayout(orientation="vertical", padding=dp(20), spacing=dp(10)); box.add_widget(Label(text=f"POTWIERDŹ:\n[b]{n} {s}[/b]\n{email}", markup=True, halign="center")); btns = BoxLayout(size_hint_y=None, height=dp(55), spacing=dp(10))
        def dec(v):
            px.dismiss()
            if v: self.send_email_engine(row, email)
            else: self.stats["skip"] += 1; Clock.schedule_once(self.process_mailing_queue, 0)
        btns.add_widget(ModernButton(text="WYŚLIJ", on_press=lambda x: dec(True), bg_color=(0,0.7,0,1))); btns.add_widget(ModernButton(text="POMIŃ", on_press=lambda x: dec(False), bg_color=(0.7,0,0,1))) # Użycie ModernButton
        box.add_widget(btns); px = Popup(title="Weryfikacja", content=box, size_hint=(0.9, 0.45)); px.open()

    def send_email_engine(self, row_data, target, fast_mode=False):
        def thread_task():
            cfg_p = Path(self.user_data_dir) / "smtp.json"
            if not cfg_p.exists(): return
            cfg = json.load(open(cfg_p))
            try:
                srv = smtplib.SMTP("smtp.gmail.com", 587, timeout=15); srv.starttls(); srv.login(cfg['u'], cfg['p'])
                msg = EmailMessage(); dat = datetime.now().strftime("%d.%m.%Y"); ts = self.conn.execute("SELECT val FROM settings WHERE key='t_sub'").fetchone(); tb = self.conn.execute("SELECT val FROM settings WHERE key='t_body'").fetchone()
                nx, sx = str(row_data[self.idx_name]).title(), str(row_data[self.idx_surname]).title()
                msg["Subject"] = (ts[0] if ts else "Wiadomość").replace("{Imię}", nx); msg["From"], msg["To"] = cfg['u'], target; msg.set_content((tb[0] if tb else "Dzień dobry").replace("{Imię}", nx).replace("{Data}", dat))
                if not fast_mode and self.full_data and Workbook:
                    tmp = Path(self.user_data_dir) / "r_tmp.xlsx"; wb = Workbook(); ws = wb.active; ws.append([self.full_data[0][k] for k in self.export_indices]); ws.append([row_data[k] for k in self.export_indices]); self.style_xlsx(ws); wb.save(tmp)
                    msg.add_attachment(open(tmp, "rb").read(), maintype="application", subtype="xlsx", filename=f"Raport_{nx}_{sx}.xlsx")
                for path in self.global_attachments:
                    if os.path.exists(path):
                        ct, _ = mimetypes.guess_type(path); mn, sb = (ct or 'application/octet-stream').split('/', 1); msg.add_attachment(open(path, "rb").read(), maintype=mn, subtype=sb, filename=os.path.basename(path))
                srv.send_message(msg); srv.quit(); Clock.schedule_once(lambda d: [self.update_stat("ok"), self.process_mailing_queue()], 0)
            except Exception as e: Clock.schedule_once(lambda d: [self.update_stat("fail"), self.process_mailing_queue()], 0)
        threading.Thread(target=thread_task, daemon=True).start()

    # --- ADVANCED LOGIC (STYLE, SMTP, REPORTS) ---
    def style_xlsx(self, ws):
        if not Workbook: return
        s, c = Side(style='thin'), Alignment(horizontal='center', vertical='center')
        for ri, row in enumerate(ws.iter_rows(), 1):
            for cell in row:
                cell.border = Border(top=s, left=s, right=s, bottom=s); cell.alignment = c
                if ri == 1: cell.font = Font(bold=True); cell.fill = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")
                elif ri % 2 == 0: cell.fill = PatternFill(start_color="F7F7F7", end_color="F7F7F7", fill_type="solid")
        for col in ws.columns: ws.column_dimensions[col[0].column_letter].width = 18

    def test_smtp(self, _):
        p = Path(self.user_data_dir) / "smtp.json"; cfg = json.load(open(p)) if p.exists() else None
        if not cfg: self.msg("!", "Skonfiguruj SMTP"); return
        def tk():
            try: srv = smtplib.SMTP("smtp.gmail.com", 587, timeout=10); srv.starttls(); srv.login(cfg['u'], cfg['p']); srv.quit(); Clock.schedule_once(lambda dt: self.msg("OK", "Połączenie serwera OK"), 0)
            except Exception as e: Clock.schedule_once(lambda dt: self.msg("Błąd", str(e)), 0)
        threading.Thread(target=tk, daemon=True).start()

    def setup_smtp_ui(self):
        l = BoxLayout(orientation="vertical", padding=dp(25), spacing=dp(10)); self.ti_h = ModernInput(hint_text="Host", text="smtp.gmail.com"); self.ti_pt = ModernInput(hint_text="Port", text="587") # Użycie ModernInput
        self.ti_u = ModernInput(hint_text="Email/Login"); self.ti_p = ModernInput(hint_text="Hasło/Klucz", password=True) # Użycie ModernInput
        p = Path(self.user_data_dir) / "smtp.json"; d = json.load(open(p)) if p.exists() else {}; self.ti_u.text, self.ti_p.text = d.get('u',''), d.get('p','')
        self.ti_h.text = d.get('h', 'smtp.gmail.com')
        self.ti_pt.text = str(d.get('port', '587'))

        sv = lambda x: [json.dump({'h':self.ti_h.text, 'port':self.ti_pt.text, 'u':self.ti_u.text, 'p':self.ti_p.text}, open(p, "w")), self.msg("OK", "Zapisano")]
        l.add_widget(Label(text="USTAWIENIA GMAIL", bold=True)); l.add_widget(self.ti_h); l.add_widget(self.ti_pt); l.add_widget(self.ti_u); l.add_widget(self.ti_p)
        l.add_widget(ModernButton(text="ZAPISZ", on_press=sv)); l.add_widget(ModernButton(text="TESTUJ", on_press=self.test_smtp)); l.add_widget(ModernButton(text="POWRÓT", on_press=lambda x: setattr(self.sm, 'current', 'home'), bg_color=(0.3, 0.3, 0.3, 1))) # Użycie ModernButton
        self.sc_ref["smtp"].add_widget(l)

    def setup_report_ui(self):
        l = BoxLayout(orientation="vertical", padding=dp(15), spacing=dp(10)); l.add_widget(Label(text="HISTORIA WYDANYCH RAPORTÓW", bold=True, size_hint_y=None, height=dp(40)))
        self.report_grid = GridLayout(cols=1, size_hint_y=None, spacing=dp(12)); self.report_grid.bind(minimum_height=self.report_grid.setter('height')); sc = ScrollView(); sc.add_widget(self.report_grid); l.add_widget(sc); l.add_widget(ModernButton(text="POWRÓT", on_press=lambda x: setattr(self.sm, 'current', 'home'), bg_color=(0.3,0.3,0.3,1))) # Użycie ModernButton
        self.sc_ref["report"].add_widget(l)

    def refresh_reports(self, *args):
        self.report_grid.clear_widgets(); rows = self.conn.execute("SELECT date, ok, fail, details FROM reports ORDER BY id DESC").fetchall() # Dodano details
        for d, ok, fl, det in rows:
            row = BoxLayout(orientation="vertical", size_hint_y=None, height=dp(110), padding=dp(10)) # Zwiększono wysokość
            with row.canvas.before:
                Color(0.15, 0.2, 0.25, 1)
                r = Rectangle(pos=row.pos, size=row.size)
            row.bind(pos=lambda i,v,rect=r: setattr(rect, 'pos', v), size=lambda i,v,rect=r: setattr(rect, 'size', v))
            row.add_widget(Label(text=f"Sesja: {d}", bold=True, color=COLOR_PRIMARY, halign="left", text_size=(dp(300), None)))
            row.add_widget(Label(text=f"Sukces: {ok} | Błąd: {fl}", font_size='13sp', halign="left", text_size=(dp(300), None)))
            row.add_widget(ModernButton(text="Pokaż logi", size_hint_y=None, height=dp(35), on_press=lambda x, t=det: self.show_details(t))) # Przycisk do pokazywania szczegółów
            self.report_grid.add_widget(row)

    def show_details(self, t): # Nowa funkcja do pokazywania szczegółów raportu
        box = BoxLayout(orientation="vertical", padding=dp(10)); ti = TextInput(text=str(t), readonly=True, font_size='11sp', background_color=(0.1,0.1,0.1,1), foreground_color=COLOR_TEXT)
        box.add_widget(ti); box.add_widget(ModernButton(text="ZAMKNIJ", size_hint_y=None, height=dp(50), on_press=lambda x: p.dismiss(), bg_color=(0.3,0.3,0.3,1)))
        p = Popup(title="Logi", content=box, size_hint=(.9,.8)); p.open()

    def setup_table_ui(self):
        root = BoxLayout(orientation="vertical"); menu = BoxLayout(size_hint_y=None, height=dp(55), spacing=dp(5), padding=dp(5)); self.ti_search = ModernInput(hint_text="Szukaj..."); self.ti_search.bind(text=self.filter_table); menu.add_widget(self.ti_search) # Użycie ModernInput
        menu.add_widget(ModernButton(text="Opcje Kolumn", size_hint_x=0.2, on_press=self.popup_columns)); menu.add_widget(ModernButton(text="WRÓĆ", size_hint_x=0.2, on_press=lambda x: setattr(self.sm, 'current', 'home'), bg_color=(0.3,0.3,0.3,1))) # Użycie ModernButton
        hs = ScrollView(size_hint_y=None, height=dp(55), do_scroll_y=False); self.table_header_layout = GridLayout(rows=1, size_hint=(None, None), height=dp(55)); hs.add_widget(self.table_header_layout); ds = ScrollView(do_scroll_x=True, do_scroll_y=True); self.table_content_layout = GridLayout(size_hint=(None, None)); self.table_content_layout.bind(minimum_height=self.table_content_layout.setter('height'), minimum_width=self.table_content_layout.setter('width')); ds.add_widget(self.table_content_layout); ds.bind(scroll_x=lambda inst, val: setattr(hs, 'scroll_x', val)); root.add_widget(menu); root.add_widget(hs); root.add_widget(ds); self.sc_ref["table"].add_widget(root)

    def refresh_table(self):
        self.table_content_layout.clear_widgets(); self.table_header_layout.clear_widgets()
        if not self.filtered_data or not self.export_indices: return
        w, h = dp(170), dp(55); headers = [self.full_data[0][i] for i in self.export_indices]; total_w = (len(headers) * w) + dp(220) # Zwiększono szerokość akcji
        self.table_header_layout.cols = len(headers) + 1; self.table_header_layout.width = total_w
        for head in headers: self.table_header_layout.add_widget(ColorSafeLabel(text=str(head), bg_color=COLOR_HEADER, bold=True, size=(w,h), size_hint=(None,None), text_color=(0,0,0,1))) # Zmieniono kolor tekstu nagłówka
        self.table_header_layout.add_widget(ColorSafeLabel(text="AKCJE", bg_color=COLOR_HEADER, bold=True, size=(dp(220),h), size_hint=(None,None), text_color=(0,0,0,1))) # Zmieniono kolor tekstu nagłówka
        self.table_content_layout.cols = len(headers) + 1; self.table_content_layout.width = total_w; 
        
        for r_idx, row in enumerate(self.filtered_data[1:]):
            row_bg = COLOR_ROW_A if r_idx % 2 == 0 else COLOR_ROW_B
            for c_idx in self.export_indices: val = row[c_idx] if c_idx < len(row) else ""; self.table_content_layout.add_widget(ColorSafeLabel(text=str(val), bg_color=row_bg, size=(w,h), size_hint=(None,None)))
            act_box = BoxLayout(size=(dp(220), h), size_hint=(None,None), spacing=dp(4), padding=dp(4)) # Szerokość akcji
            act_box.add_widget(ModernButton(text="ZAPISZ", on_press=lambda x, r=row: self.export_xlsx(r), bg_color=(0.2, 0.6, 0.2, 1))) # Użycie ModernButton
            act_box.add_widget(ModernButton(text="WYŚLIJ", on_press=lambda x, r=row: self.send_individual_from_table(r), bg_color=(0.1, 0.5, 0.9, 1))) # Użycie ModernButton
            self.table_content_layout.add_widget(act_box)

    def send_individual_from_table(self, row): # Nowa funkcja do wysyłki z tabeli
        name, sur = str(row[self.idx_name]).strip(), str(row[self.idx_surname]).strip()
        pes = str(row[self.idx_pesel]).strip() if self.idx_pesel != -1 else ""
        res = self.conn.execute("SELECT email FROM contacts WHERE pesel=? AND pesel != ''", (pes,)).fetchone() if pes else None
        if not res: res = self.conn.execute("SELECT email FROM contacts WHERE name=? AND surname=? COLLATE NOCASE", (name.lower(), sur.lower())).fetchone()
        if not res: return self.msg("Błąd", f"Brak maila dla: {name}")
        def task():
            cfg_p = Path(self.user_data_dir)/"smtp.json"
            if not cfg_p.exists(): return Clock.schedule_once(lambda d: self.msg("!", "Brak SMTP"), 0)
            cfg = json.load(open(cfg_p)); srv = self.connect_smtp(cfg)
            if self.send_single_email(srv, cfg, row, res[0]): Clock.schedule_once(lambda d: self.msg("OK", f"Wysłano do: {name}"), 0)
            srv.quit()
        threading.Thread(target=task, daemon=True).start()

    def connect_smtp(self, cfg): # Nowa funkcja do łączenia z SMTP
        s = smtplib.SMTP(cfg.get('h','smtp.gmail.com'), int(cfg.get('port',587)), timeout=25); s.starttls(); s.login(cfg['u'], cfg['p']); return s

    def send_single_email(self, srv, cfg, row_data, target): # Nowa funkcja do wysyłki pojedynczego maila
        try:
            nx, sx = str(row_data[self.idx_name]).title(), str(row_data[self.idx_surname]).title()
            msg = EmailMessage(); ts, tb = self.conn.execute("SELECT val FROM settings WHERE key='t_sub'").fetchone(), self.conn.execute("SELECT val FROM settings WHERE key='t_body'").fetchone()
            msg["Subject"] = (ts[0] if ts else "Raport").replace("{Imię}", nx); msg["From"], msg["To"] = cfg['u'], target
            msg.set_content((tb[0] if tb else "Dzień dobry").replace("{Imię}", nx).replace("{Data}", datetime.now().strftime("%d.%m.%Y")))
            t_f = Path(self.user_data_dir)/f"r_{nx}.xlsx"; wb = Workbook(); ws = wb.active
            ws.append([self.full_data[0][k] for k in self.export_indices]); ws.append([str(row_data[k]) if (str(row_data[k]).strip()!="") else "0" for k in self.export_indices])
            self.style_xlsx(ws); wb.save(t_f)
            with open(t_f, "rb") as f: msg.add_attachment(f.read(), maintype="application", subtype="xlsx", filename=f"Raport_{nx}_{sx}.xlsx")
            for p in self.global_attachments:
                if os.path.exists(p):
                    ct, _ = mimetypes.guess_type(p); mn, sb = (ct or 'application/octet-stream').split('/', 1)
                    with open(p,"rb") as f: msg.add_attachment(f.read(), maintype=mn, subtype=sb, filename=os.path.basename(p))
            srv.send_message(msg); return True
        except: return False

    def setup_tmpl_ui(self):
        l, ti_s, ti_b = BoxLayout(orientation="vertical", padding=dp(25), spacing=dp(10)), ModernInput(hint_text="Temat {Imię}"), ModernInput(hint_text="Treść...", multiline=True) # Użycie ModernInput
        ts, tb = self.conn.execute("SELECT val FROM settings WHERE key='t_sub'").fetchone(), self.conn.execute("SELECT val FROM settings WHERE key='t_body'").fetchone()
        ti_s.text, ti_b.text = (ts[0] if ts else ""), (tb[0] if tb else "")
        l.add_widget(Label(text="SZABLON EMAIL", bold=True)); l.add_widget(ti_s); l.add_widget(ti_b)
        l.add_widget(ModernButton(text="ZAPISZ", on_press=lambda x: [self.conn.execute("INSERT OR REPLACE INTO settings VALUES (?,?)", ('t_sub',ti_s.text)), self.conn.execute("INSERT OR REPLACE INTO settings VALUES (?,?)", ('t_body',ti_b.text)), self.conn.commit(), self.msg("OK","Wzór zapisany")])); l.add_widget(ModernButton(text="POWRÓT", on_press=lambda x: setattr(self.sm, 'current', 'email'), bg_color=(0.3,0.3,0.3,1))) # Użycie ModernButton
        self.sc_ref["tmpl"].add_widget(l)

    def setup_contacts_ui(self):
        l, top = BoxLayout(orientation="vertical", padding=dp(10)), BoxLayout(size_hint_y=None, height=dp(55), spacing=dp(5))
        self.ti_csearch = ModernInput(hint_text="Szukaj..."); self.ti_csearch.bind(text=self.refresh_contacts_list); top.add_widget(self.ti_csearch) # Użycie ModernInput
        top.add_widget(ModernButton(text="+", size_hint_x=0.15, on_press=lambda x: self.form_contact())); top.add_widget(ModernButton(text="Wróć", size_hint_x=0.2, on_press=lambda x: setattr(self.sm, 'current', 'home'), bg_color=(0.3,0.3,0.3,1))) # Użycie ModernButton
        self.c_list = GridLayout(cols=1, size_hint_y=None, spacing=dp(10)); self.c_list.bind(minimum_height=self.c_list.setter('height'))
        sc = ScrollView(); sc.add_widget(self.c_list); l.add_widget(top); l.add_widget(sc); self.sc_ref["contacts"].add_widget(l)

    def refresh_contacts_list(self, *args):
        self.c_list.clear_widgets(); sv = self.ti_csearch.text.lower(); rows = self.conn.execute("SELECT name, surname, email, pesel, phone FROM contacts ORDER BY surname ASC").fetchall()
        for n, s, e, p, ph in rows:
            if sv and sv not in f"{n} {s} {e} {p} {ph}".lower(): continue
            row = BoxLayout(size_hint_y=None, height=dp(125), padding=dp(10)) # Zwiększono wysokość dla lepszego wyglądu
            with row.canvas.before: Color(*COLOR_CARD); Rectangle(pos=row.pos, size=row.size, radius=[dp(12)]) # Użycie RoundedRectangle
            inf, acts = BoxLayout(orientation="vertical"), BoxLayout(size_hint_x=0.3, orientation="vertical", spacing=dp(4))
            inf.add_widget(Label(text=f"{n} {s}".title(), bold=True, halign="left", text_size=(dp(250),None), color=COLOR_TEXT))
            inf.add_widget(Label(text=f"E: {e}\nP: {p if p else '-'}\nT: {ph if ph else '-'}", font_size='11sp', halign="left", text_size=(dp(250),None), color=(0.7,0.7,0.7,1)))
            row.add_widget(inf); acts.add_widget(ModernButton(text="Edytuj", on_press=lambda x, data=(n,s,e,p,ph): self.form_contact(*data))); acts.add_widget(ModernButton(text="Usuń", bg_color=(0.8,0.2,0.2,1), on_press=lambda x, name=n, sur=s: self.delete_contact(name,sur))) # Użycie ModernButton
            row.add_widget(acts); self.c_list.add_widget(row)

    def process_excel(self, path):
        try:
            if str(path).endswith(".xls") and xlrd:
                wb = xlrd.open_workbook(path); ws = wb.sheet_by_index(0); raw = [[str(ws.cell_value(r,c)).strip() for c in range(ws.ncols)] for r in range(ws.nrows)]
            else:
                wb = load_workbook(path, data_only=True); ws = wb.active; raw = [["" if v is None else str(v).strip() for v in r] for r in ws.iter_rows(values_only=True)]
            h_idx = 0;
            for i, row in enumerate(raw[:15]):
                ln = " ".join([str(x) for x in row]).lower()
                if any(x in ln for x in ["imię", "imie", "nazwisko", "pesel"]): h_idx = i; break
            self.full_data = raw[h_idx:]; self.filtered_data = self.full_data; self.export_indices = list(range(len(self.full_data[0]))); h = [str(x).lower() for x in self.full_data[0]]
            for i,v in enumerate(h):
                if "imi" in v: self.idx_name = i
                if "naz" in v: self.idx_surname = i
                if "pesel" in v: self.idx_pesel = i
            self.msg("OK", "Wczytano arkusz.")
        except Exception as e: self.msg("Błąd", str(e))

    def process_book(self, path):
        try:
            wb = load_workbook(path, data_only=True); ws = wb.active; raw = [["" if v is None else str(v).strip() for v in r] for r in ws.iter_rows(values_only=True)]
            h = [str(x).lower() for x in raw[0]]; iN, iS, iE, iP = 0, 1, 2, -1
            for i,v in enumerate(h):
                if "imi" in v: iN=i
                elif "naz" in v: iS=i
                elif "@" in v or "mail" in v: iE=i
                elif "pesel" in v: iP=i
            for r in raw[1:]:
                if len(r) > iE and "@" in str(r[iE]):
                    pes_v = str(r[iP]) if (iP != -1 and len(r) > iP) else ""
                    self.conn.execute("INSERT OR REPLACE INTO contacts (name, surname, email, pesel, phone) VALUES (?,?,?,?,?)", (r[iN].lower(), r[iS].lower(), str(r[iE]).strip(), pes_v, ""))
            self.conn.commit(); self.update_stats(); self.msg("OK", "Skompletowano bazę.")
        except Exception as e: self.msg("Błąd", f"Nie wczytano bazy: {e}")

    def export_xlsx(self, r):
        p = Path("/storage/emulated/0/Documents/FutureExport") if platform=="android" else Path("./exports"); p.mkdir(parents=True, exist_ok=True)
        nx, sx = str(r[self.idx_name]).title(), str(r[self.idx_surname]).title(); wb = Workbook(); ws = wb.active; ws.append([self.full_data[0][k] for k in self.export_indices]); ws.append([r[k] for k in self.export_indices]); self.style_xlsx(ws); wb.save(p/f"Raport_{nx}_{sx}.xlsx"); self.msg("OK", f"Zapisano: {nx}")

    def filter_table(self, ins, val):
        v = val.lower(); self.filtered_data = [self.full_data[0]] + [r for r in self.full_data[1:] if any(v in str(c).lower() for c in r)]; self.refresh_table()

    def popup_columns(self, _):
        if not self.full_data: return
        box = BoxLayout(orientation="vertical", padding=dp(15)); gr = GridLayout(cols=1, size_hint_y=None, spacing=dp(8)); gr.bind(minimum_height=gr.setter('height')); checks = []
        for i, h in enumerate(self.full_data[0]):
            r = BoxLayout(size_hint_y=None, height=dp(50)); cb = CheckBox(active=(i in self.export_indices), size_hint_x=None, width=dp(50)); checks.append((i, cb)); r.add_widget(cb); r.add_widget(Label(text=str(h))); gr.add_widget(r)
        sc = ScrollView(); sc.add_widget(gr); box.add_widget(sc); box.add_widget(ModernButton(text="OK", on_press=lambda x: [setattr(self, 'export_indices', [idx for idx, c in checks if c.active]), p.dismiss(), self.refresh_table()], bg_color=(0.3,0.3,0.3,1))) # Użycie ModernButton
        p = Popup(title="Widoczność kolumn", content=box, size_hint=(0.9, 0.9)); p.open()

    def form_contact(self, n="", s="", e="", pes="", ph=""):
        b = BoxLayout(orientation="vertical", padding=dp(15), spacing=dp(10)); flds = [ModernInput(text=n, hint_text="Imię"), ModernInput(text=s, hint_text="Nazwisko"), ModernInput(text=e, hint_text="Email"), ModernInput(text=pes, hint_text="PESEL"), ModernInput(text=ph, hint_text="Telefon")] # Użycie ModernInput
        for f in flds: b.add_widget(f)
        def save(_): [self.conn.execute("INSERT OR REPLACE INTO contacts VALUES (?,?,?,?,?)", [f.text.strip().lower() if i<2 else f.text.strip() for i,f in enumerate(flds)]), self.conn.commit(), px.dismiss(), self.refresh_contacts_list(), self.update_stats()]
        b.add_widget(ModernButton(text="ZAPISZ", on_press=save)); px = Popup(title="Kontakt", content=b, size_hint=(0.9, 0.85)); px.open() # Użycie ModernButton

    def delete_contact(self, n, s):
        def pr(_): [self.conn.execute("DELETE FROM contacts WHERE name=? AND surname=?", (n, s)), self.conn.commit(), px.dismiss(), self.refresh_contacts_list(), self.update_stats()]; px = Popup(title="Usuń?", content=ModernButton(text="USUŃ", on_press=pr, bg_color=(1,0,0,1)), size_hint=(0.7,0.3)); px.open() # Użycie ModernButton

    def update_stat(self, k): self.stats[k]+=1
    def update_stats(self, *a):
        try:
            c = self.conn.execute("SELECT count(*) FROM contacts").fetchone()[0]
            if hasattr(self, 'lbl_stats'):
                self.lbl_stats.text = f"Baza: {c} | Załączniki: {len(self.global_attachments)}"
            if hasattr(self, 'lbl_stats_paski'): # Aktualizacja licznika dla ekranu Paski
                self.lbl_stats_paski.text = f"Baza: {c} | Załączniki: {len(self.global_attachments)}"
        except: pass

    def update_progress(self, *args): # Uogólniona funkcja aktualizacji paska postępu
        try:
            val = int((self.total_q - len(self.queue)) / self.total_q * 100) if self.total_q else 0
            if hasattr(self, 'pb'):
                self.pb.value = val
            if hasattr(self, 'pb_paski'): # Aktualizacja paska dla ekranu Paski
                self.pb_paski.value = val
            if hasattr(self, 'pb_label'):
                self.pb_label.text = f"Postęp: {self.total_q - len(self.queue)}/{self.total_q}"
            if hasattr(self, 'pb_label_paski'): # Aktualizacja etykiety dla ekranu Paski
                self.pb_label_paski.text = f"Postęp: {self.total_q - len(self.queue)}/{self.total_q}"
        except:
            pass

    def msg(self, tit, txt):
        b = BoxLayout(orientation="vertical", padding=dp(20)); b.add_widget(Label(text=txt, halign="center")); btn = ModernButton(text="OK", size_hint_y=None, height=dp(50), on_press=lambda x: p.dismiss()); b.add_widget(btn); p = Popup(title=tit, content=b, size_hint=(0.85, 0.45)); p.open()

    # --- NOWY MODUŁ PASKI ---
    def setup_paski_ui(self):
        l = BoxLayout(orientation="vertical", padding=dp(15), spacing=dp(10))
        header = BoxLayout(size_hint_y=None, height=dp(40))
        header.add_widget(Label(text="Moduł Paski", bold=True, font_size='22sp', color=COLOR_PRIMARY))
        l.add_widget(header)

        ab = BoxLayout(size_hint_y=None, height=dp(45), spacing=dp(10))
        self.cb_paski_auto = CheckBox(size_hint_x=None, width=dp(45))
        self.cb_paski_auto.bind(active=lambda i, v: setattr(self, 'auto_send_mode', v))
        ab.add_widget(self.cb_paski_auto); ab.add_widget(Label(text="AUTO-WYSYŁKA (BEZ POTWIERDZEŃ)", bold=True, font_size='12sp'))
        l.add_widget(ab)

        self.lbl_stats_paski = Label(text="Baza: 0 | Załączniki: 0", height=dp(30)); l.add_widget(self.lbl_stats_paski) # Licznik załączników
        self.pb_label_paski = Label(text="Gotowy", height=dp(25)); self.pb_paski = ProgressBar(max=100, height=dp(20)); l.add_widget(self.pb_label_paski); l.add_widget(self.pb_paski)

        l.add_widget(ModernButton(text="Wczytaj arkusz płac", on_press=lambda x: self.open_picker("data"), height=dp(50), size_hint_y=None))
        l.add_widget(ModernButton(text="Podgląd i eksport", on_press=lambda x: [self.refresh_table(), setattr(self.sm, 'current', 'table')] if self.full_data else self.msg("!", "Wczytaj arkusz!"), height=dp(50), size_hint_y=None))
        l.add_widget(ModernButton(text="Edytuj szablon", on_press=lambda x: setattr(self.sm, 'current', 'tmpl'), height=dp(50), size_hint_y=None))
        l.add_widget(ModernButton(text="Dołącz załącznik", on_press=lambda x: self.open_picker("attachment"), height=dp(50), size_hint_y=None))
        l.add_widget(ModernButton(text="Wyślij jeden plik", on_press=self.start_special_send_flow, height=dp(50), size_hint_y=None))
        l.add_widget(ModernButton(text="Start masowa wysyłka", on_press=self.start_mass_mailing, height=dp(50), size_hint_y=None))
        l.add_widget(ModernButton(text="Raporty sesji", on_press=lambda x: [self.refresh_reports(), setattr(self.sm, 'current', 'report')], height=dp(50), size_hint_y=None))
        l.add_widget(ModernButton(text="Wyczyść załączniki", on_press=lambda x: [self.global_attachments.clear(), self.update_stats()], height=dp(50), size_hint_y=None, bg_color=(0.7, 0.1, 0.1, 1)))
        l.add_widget(ModernButton(text="Powrót", on_press=lambda x: setattr(self.sm, 'current', 'home'), height=dp(55), size_hint_y=None, bg_color=(0.3,0.3,0.3,1)))
        self.sc_ref["paski"].add_widget(l)

    # --- NOWY MODUŁ USTAWIEŃ ---
    def setup_settings_ui(self):
        l = BoxLayout(orientation="vertical", padding=dp(15), spacing=dp(10))
        l.add_widget(Label(text="Ustawienia Aplikacji", bold=True, font_size='22sp', color=COLOR_PRIMARY))
        
        l.add_widget(ModernButton(text="Dodaj bazę danych kontaktów", on_press=lambda x: self.open_picker("book"), height=dp(50), size_hint_y=None))
        l.add_widget(ModernButton(text="Ustawienia SMTP", on_press=lambda x: setattr(self.sm, 'current', 'smtp'), height=dp(50), size_hint_y=None))
        # Usunięto: "Edytuj szablon email" i "Wczytaj arkusz płac" - są w innych miejscach
        
        l.add_widget(ModernButton(text="Powrót", on_press=lambda x: setattr(self.sm, 'current', 'home'), height=dp(55), size_hint_y=None, bg_color=(0.3,0.3,0.3,1)))
        self.sc_ref["settings"].add_widget(l)

if __name__ == "__main__":
    FutureApp().run()
