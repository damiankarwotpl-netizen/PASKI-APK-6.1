import os
import json
import sqlite3
import threading
import smtplib
import mimetypes
import shutil
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
from kivy.uix.screenmanager import ScreenManager, Screen, SlideTransition
from kivy.uix.progressbar import ProgressBar
from kivy.graphics import Color, Rectangle, RoundedRectangle

# Obsługa Excel
try:
    from openpyxl import load_workbook, Workbook
    from openpyxl.styles import Border, Side, Font, Alignment, PatternFill
except ImportError:
    load_workbook = Workbook = None

try:
    import xlrd
except ImportError:
    xlrd = None

# Nowa Paleta Kolorów (Modern Dark)
COLOR_PRIMARY = (0.12, 0.58, 0.95, 1)  # Vivid Blue
COLOR_ACCENT = (0.0, 0.8, 0.6, 1)     # Teal
COLOR_BG = (0.05, 0.07, 0.1, 1)       # Dark Navy
COLOR_CARD = (0.1, 0.13, 0.18, 1)     # Lighter Navy
COLOR_TEXT = (0.95, 0.95, 0.98, 1)
COLOR_DANGER = (0.9, 0.2, 0.2, 1)

class ModernButton(Button):
    def __init__(self, bg_color=COLOR_PRIMARY, **kwargs):
        super().__init__(**kwargs)
        self.background_normal = ""
        self.background_down = ""
        self.background_color = (0,0,0,0)
        self.color = COLOR_TEXT
        self.bold = True
        self.font_size = '15sp'
        self.radius = [dp(12)]
        self.inner_color = bg_color
        with self.canvas.before:
            Color(*self.inner_color)
            self.rect = RoundedRectangle(pos=self.pos, size=self.size, radius=self.radius)
        self.bind(pos=self._update, size=self._update)

    def _update(self, *args):
        self.rect.pos = self.pos
        self.rect.size = self.size

class ModernInput(TextInput):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.background_normal = ""
        self.background_active = ""
        self.background_color = (0.15, 0.18, 0.25, 1)
        self.foreground_color = COLOR_TEXT
        self.cursor_color = COLOR_PRIMARY
        self.padding = [dp(12), dp(12)]
        self.hint_text_color = (0.5, 0.5, 0.6, 1)
        with self.canvas.after:
            Color(*COLOR_PRIMARY[:3], 0.3)
            self.line = RoundedRectangle(pos=self.pos, size=self.size, radius=[dp(8)], size_hint=(None, None))
        self.bind(pos=self._update, size=self._update)

    def _update(self, *args):
        self.line.pos = self.pos
        self.line.size = self.size

class CardRow(BoxLayout):
    def __init__(self, bg_color=COLOR_CARD, **kwargs):
        super().__init__(**kwargs)
        with self.canvas.before:
            Color(*bg_color)
            self.rect = RoundedRectangle(pos=self.pos, size=self.size, radius=[dp(15)])
        self.bind(pos=self._update, size=self._update)

    def _update(self, *args):
        self.rect.pos = (self.pos[0] + dp(5), self.pos[1] + dp(5))
        self.rect.size = (self.size[0] - dp(10), self.size[1] - dp(10))

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
        Window.clearcolor = COLOR_BG
        self.full_data = []; self.filtered_data = []; self.export_indices = []
        self.global_attachments = []; self.selected_emails = []; self.queue = []
        self.session_details = []
        self.stats = {"ok": 0, "fail": 0, "skip": 0, "auto": 0}
        self.idx_name, self.idx_surname, self.idx_pesel = 0, 1, -1
        self.auto_send_mode = False
        
        if not os.path.exists(self.user_data_dir): 
            os.makedirs(self.user_data_dir, exist_ok=True)
        self.init_db()
        
        if platform == "android":
            from android.permissions import request_permissions, Permission
            request_permissions([Permission.READ_EXTERNAL_STORAGE, Permission.WRITE_EXTERNAL_STORAGE])

        self.sm = ScreenManager(transition=SlideTransition())
        self.add_screens()
        return self.sm

    def init_db(self):
        db_p = Path(self.user_data_dir) / "future_ultimate_v2.db"
        self.conn = sqlite3.connect(str(db_p), check_same_thread=False)
        self.conn.execute("CREATE TABLE IF NOT EXISTS contacts (name TEXT, surname TEXT, email TEXT, pesel TEXT, phone TEXT, PRIMARY KEY(name, surname))")
        self.conn.execute("CREATE TABLE IF NOT EXISTS settings (key TEXT PRIMARY KEY, val TEXT)")
        self.conn.execute("CREATE TABLE IF NOT EXISTS reports (id INTEGER PRIMARY KEY AUTOINCREMENT, date TEXT, ok INTEGER, fail INTEGER, skip INTEGER, auto INTEGER, details TEXT)")
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

    def setup_ui(self):
        # HOME SCREEN
        l = BoxLayout(orientation="vertical", padding=dp(25), spacing=dp(18))
        l.add_widget(Label(text="FUTURE 22.4", font_size='32sp', bold=True, color=COLOR_PRIMARY, size_hint_y=None, height=dp(60)))
        l.add_widget(Label(text="ULTIMATE MAILING SYSTEM", font_size='14sp', color=(0.5, 0.5, 0.6, 1), size_hint_y=None, height=dp(20)))
        
        btn_grid = GridLayout(cols=1, spacing=dp(12))
        btns = [
            ("📊 WCZYTAJ ARKUSZ PŁAC", lambda x: self.open_picker("data")),
            ("👁️ PODGLĄD I EKSPORT", lambda x: [self.refresh_table(), setattr(self.sm, 'current', 'table')] if self.full_data else self.msg("!", "Brak danych")),
            ("📧 CENTRUM MAILINGOWE", lambda x: setattr(self.sm, 'current', 'email')),
            ("📜 RAPORTY WYSYŁEK", lambda x: [self.refresh_reports(), setattr(self.sm, 'current', 'report')]),
            ("⚙️ USTAWIENIA GMAIL", lambda x: setattr(self.sm, 'current', 'smtp'))
        ]
        for t, c in btns:
            btn_grid.add_widget(ModernButton(text=t, on_press=c, height=dp(58), size_hint_y=None))
        
        l.add_widget(btn_grid)
        self.screens["home"].add_widget(l)
        
        # Inicjalizacja reszty UI
        self.setup_table_ui(); self.setup_email_ui(); self.setup_smtp_ui(); self.setup_tmpl_ui(); self.setup_contacts_ui(); self.setup_report_ui()

    def setup_email_ui(self):
        l = BoxLayout(orientation="vertical", padding=dp(25), spacing=dp(12))
        
        header = CardRow(size_hint_y=None, height=dp(80), orientation="vertical", padding=dp(10))
        auto_box = BoxLayout(spacing=dp(10))
        self.cb_auto = CheckBox(size_hint_x=None, width=dp(45))
        self.cb_auto.bind(active=lambda i, v: setattr(self, 'auto_send_mode', v))
        auto_box.add_widget(self.cb_auto)
        auto_box.add_widget(Label(text="TRYB AUTOMATYCZNY", bold=True))
        header.add_widget(auto_box)
        self.lbl_stats = Label(text="Baza: 0", font_size='12sp', color=COLOR_ACCENT)
        header.add_widget(self.lbl_stats)
        l.add_widget(header)

        prog_box = BoxLayout(orientation="vertical", size_hint_y=None, height=dp(60), spacing=dp(5))
        self.pb_label = Label(text="Gotowy", font_size='13sp')
        self.pb = ProgressBar(max=100, height=dp(15))
        prog_box.add_widget(self.pb_label); prog_box.add_widget(self.pb)
        l.add_widget(prog_box)

        btn_grid = GridLayout(cols=2, spacing=dp(10), size_hint_y=None, height=dp(200))
        btns = [
            ("👥 BAZA", lambda x: [self.refresh_contacts_list(), setattr(self.sm, 'current', 'contacts')]),
            ("📝 SZABLON", lambda x: setattr(self.sm, 'current', 'tmpl')),
            ("📎 ZAŁĄCZNIK", lambda x: self.open_picker("attachment")),
            ("📁 WYŚLIJ PLIK", self.start_special_send_flow)
        ]
        for t, c in btns: btn_grid.add_widget(ModernButton(text=t, on_press=c, bg_color=COLOR_CARD))
        l.add_widget(btn_grid)

        l.add_widget(ModernButton(text="🚀 START MASOWA WYSYŁKA", on_press=self.start_mass_mailing, bg_color=COLOR_ACCENT, height=dp(60), size_hint_y=None))
        l.add_widget(ModernButton(text="POWRÓT", on_press=lambda x: setattr(self.sm, 'current', 'home'), bg_color=(0.3, 0.3, 0.3, 1), height=dp(50), size_hint_y=None))
        
        self.screens["email"].add_widget(l); self.update_stats()

    def setup_contacts_ui(self):
        l = BoxLayout(orientation="vertical", padding=dp(12), spacing=dp(10))
        top = BoxLayout(size_hint_y=None, height=dp(55), spacing=dp(8))
        self.ti_csearch = ModernInput(hint_text="Szukaj kontaktu...")
        self.ti_csearch.bind(text=self.refresh_contacts_list)
        top.add_widget(self.ti_csearch)
        top.add_widget(ModernButton(text="+", size_hint_x=None, width=dp(55), on_press=lambda x: self.form_contact()))
        
        self.c_list = GridLayout(cols=1, size_hint_y=None, spacing=dp(8))
        self.c_list.bind(minimum_height=self.c_list.setter('height'))
        sc = ScrollView(); sc.add_widget(self.c_list)
        
        l.add_widget(top); l.add_widget(sc)
        
        bot = BoxLayout(size_hint_y=None, height=dp(55), spacing=dp(10))
        bot.add_widget(ModernButton(text="CZYŚĆ WYBÓR", on_press=lambda x: [setattr(self, 'selected_emails', []), self.refresh_contacts_list()], bg_color=(0.4, 0.4, 0.4, 1)))
        bot.add_widget(ModernButton(text="POWRÓT", on_press=lambda x: setattr(self.sm, 'current', 'email')))
        l.add_widget(bot)
        
        self.screens["contacts"].add_widget(l)

    def refresh_contacts_list(self, *args):
        self.c_list.clear_widgets()
        sv = self.ti_csearch.text.lower()
        rows = self.conn.execute("SELECT name, surname, email, pesel, phone FROM contacts ORDER BY surname ASC").fetchall()
        for n, s, e, p, ph in rows:
            if sv and sv not in f"{n} {s} {e} {p} {ph}".lower(): continue
            
            row = CardRow(size_hint_y=None, height=dp(110), padding=dp(12))
            cb = CheckBox(size_hint_x=None, width=dp(40), active=(e in self.selected_emails))
            cb.bind(active=lambda i, v, m=e: self.selected_emails.append(m) if v else self.selected_emails.remove(m))
            row.add_widget(cb)
            
            main_info = BoxLayout(orientation="vertical", spacing=dp(2))
            main_info.add_widget(Label(text=f"{n} {s}".upper(), bold=True, halign="left", text_size=(dp(200), None), color=COLOR_PRIMARY))
            main_info.add_widget(Label(text=f"✉ {e}", font_size='11sp', halign="left", text_size=(dp(200), None)))
            main_info.add_widget(Label(text=f"ID: {p if p else '-'}", font_size='11sp', color=(0.6, 0.6, 0.6, 1), halign="left", text_size=(dp(200), None)))
            row.add_widget(main_info)
            
            acts = BoxLayout(size_hint_x=None, width=dp(80), orientation="vertical", spacing=dp(5))
            acts.add_widget(Button(text="EDYTUJ", font_size='10sp', background_color=(0.2,0.6,1,1), on_press=lambda x, d=(n,s,e,p,ph): self.form_contact(*d)))
            # NAPRAWIONY GUZIK USUWANIA:
            acts.add_widget(Button(text="USUŃ", font_size='10sp', background_color=COLOR_DANGER, on_press=lambda x, na=n, su=s: self.delete_contact(na,su)))
            row.add_widget(acts)
            
            self.c_list.add_widget(row)

    def delete_contact(self, n, s):
        content = BoxLayout(orientation="vertical", padding=dp(20), spacing=dp(15))
        content.add_widget(Label(text=f"Czy na pewno usunąć:\n{n.title()} {s.title()}?", halign="center"))
        
        def confirm_delete(_):
            self.conn.execute("DELETE FROM contacts WHERE name=? AND surname=?", (n, s)) # Poprawiona krotka parametrów
            self.conn.commit()
            px.dismiss()
            self.refresh_contacts_list()
            self.update_stats()

        btn_box = BoxLayout(spacing=dp(10), size_hint_y=None, height=dp(50))
        btn_box.add_widget(ModernButton(text="TAK", on_press=confirm_delete, bg_color=COLOR_DANGER))
        btn_box.add_widget(ModernButton(text="NIE", on_press=lambda x: px.dismiss(), bg_color=(0.4, 0.4, 0.4, 1)))
        
        content.add_widget(btn_box)
        px = Popup(title="Potwierdzenie", content=content, size_hint=(0.85, 0.4), background_color=(0.1, 0.1, 0.1, 0.9))
        px.open()

    def setup_report_ui(self):
        l = BoxLayout(orientation="vertical", padding=dp(15), spacing=dp(10))
        l.add_widget(Label(text="HISTORIA OPERACJI", bold=True, height=dp(40), size_hint_y=None, color=COLOR_PRIMARY))
        
        sc = ScrollView()
        self.report_grid = GridLayout(cols=1, size_hint_y=None, spacing=dp(10))
        self.report_grid.bind(minimum_height=self.report_grid.setter('height'))
        sc.add_widget(self.report_grid)
        l.add_widget(sc)
        
        l.add_widget(ModernButton(text="POWRÓT", on_press=lambda x: setattr(self.sm, 'current', 'home'), height=dp(55), size_hint_y=None))
        self.screens["report"].add_widget(l)

    def refresh_reports(self, *args):
        self.report_grid.clear_widgets()
        rows = self.conn.execute("SELECT date, ok, fail, skip, details FROM reports ORDER BY id DESC LIMIT 50").fetchall()
        for d, ok, fl, sk, dt_text in rows:
            row = CardRow(orientation="vertical", size_hint_y=None, height=dp(100), padding=dp(12))
            row.add_widget(Label(text=f"📅 {d}", bold=True, color=COLOR_ACCENT, halign="left", text_size=(dp(300), None)))
            row.add_widget(Label(text=f"Sukces: {ok} | Błąd: {fl} | Pominięto: {sk}", font_size='12sp', halign="left", text_size=(dp(300), None)))
            
            btn = Button(text="SZCZEGÓŁY LOGÓW", size_hint_y=None, height=dp(30), background_color=(1,1,1,0.1), font_size='10sp')
            btn.bind(on_press=lambda x, t=dt_text: self.show_report_details(t))
            row.add_widget(btn)
            self.report_grid.add_widget(row)

    # --- Reszta metod (mechanika bez zmian, tylko minimalne poprawki UI) ---

    def open_picker(self, mode):
        if platform != "android": 
            # Fallback dla PC (do testów)
            if mode == "attachment": 
                self.global_attachments.append("test_file.pdf")
                self.update_stats()
            self.msg("!", "Pełny picker działa na systemie Android")
            return
            
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
                except: self.msg("Błąd", "Błąd odczytu pliku.")
        activity.bind(on_activity_result=cb); PA.mActivity.startActivityForResult(intent, 1001)

    def start_special_send_flow(self, _): 
        if not self.conn.execute("SELECT count(*) FROM contacts").fetchone()[0]:
            return self.msg("!", "Baza kontaktów jest pusta!")
        self.open_picker("special_send")

    def special_send_step_2_recipients(self, file_path):
        self.selected_emails = []
        box = BoxLayout(orientation="vertical", padding=dp(15), spacing=dp(10))
        box.add_widget(Label(text="KROK 2: WYBIERZ ODBIORCÓW", bold=True, color=COLOR_PRIMARY, size_hint_y=None, height=dp(40)))
        ti = ModernInput(hint_text="Szukaj..."); box.add_widget(ti)
        sc = ScrollView(); gl = GridLayout(cols=1, size_hint_y=None, spacing=dp(5)); gl.bind(minimum_height=gl.setter('height')); sc.add_widget(gl); box.add_widget(sc)
        def refresh(v=""):
            gl.clear_widgets(); rows = self.conn.execute("SELECT name, surname, email FROM contacts").fetchall()
            for n, s, e in rows:
                if v and v.lower() not in f"{n} {s} {e}".lower(): continue
                r = CardRow(size_hint_y=None, height=dp(55))
                cb = CheckBox(size_hint_x=None, width=dp(50), active=(e in self.selected_emails))
                cb.bind(active=lambda i, val, mail=e: self.selected_emails.append(mail) if val else self.selected_emails.remove(mail))
                r.add_widget(cb); r.add_widget(Label(text=f"{n.title()} {s.title()}\\n{e}", font_size='11sp')); gl.add_widget(r)
        ti.bind(text=lambda i,v: refresh(v)); refresh()
        btn = ModernButton(text="DALEJ"); btn.bind(on_press=lambda x: [p.dismiss(), self.special_send_step_3_msg(file_path)] if self.selected_emails else self.msg("!", "Wybierz kogoś!")); box.add_widget(btn); p = Popup(title="Odbiorcy", content=box, size_hint=(0.95,0.9)); p.open()

    def special_send_step_3_msg(self, file_path):
        box = BoxLayout(orientation="vertical", padding=dp(15), spacing=dp(10))
        ti_s = ModernInput(hint_text="Temat wiadomości", size_hint_y=None, height=dp(55))
        ti_b = ModernInput(hint_text="Treść...", multiline=True)
        box.add_widget(ti_s); box.add_widget(ti_b); btn = ModernButton(text="WYŚLIJ DO WYBRANYCH", bg_color=COLOR_ACCENT)
        btn.bind(on_press=lambda x: [p.dismiss(), self.special_send_step_4_progress(file_path, self.selected_emails, ti_s.text, ti_b.text)] if ti_s.text and ti_b.text else self.msg("!", "Wypełnij dane!")); box.add_widget(btn); p = Popup(title="Wiadomość", content=box, size_hint=(0.95,0.85)); p.open()

    def special_send_step_4_progress(self, file_path, target_list, subject, body):
        box = BoxLayout(orientation="vertical", padding=dp(25), spacing=dp(15))
        lbl = Label(text="Inicjalizacja wysyłki...", bold=True); pb = ProgressBar(max=len(target_list), value=0, height=dp(30))
        box.add_widget(lbl); box.add_widget(pb); btn_c = ModernButton(text="ZAMKNIJ", height=dp(55), disabled=True, bg_color=(0.3,0.3,0.3,1))
        p = Popup(title="Status wysyłki", content=box, size_hint=(0.9, 0.45), auto_dismiss=False); btn_c.bind(on_press=p.dismiss); box.add_widget(btn_c); p.open()
        def run():
            cfg_p = Path(self.user_data_dir) / "smtp.json"
            if not cfg_p.exists(): Clock.schedule_once(lambda dt: self.msg("!", "Brak SMTP")); p.dismiss(); return
            cfg = json.load(open(cfg_p)); ok, err = 0, 0
            try:
                srv = smtplib.SMTP("smtp.gmail.com", 587, timeout=15); srv.starttls(); srv.login(cfg['u'], cfg['p'])
                for i, email in enumerate(target_list):
                    try:
                        Clock.schedule_once(lambda dt: setattr(lbl, 'text', f"Wysyłka: {email}"))
                        msg = EmailMessage(); msg["Subject"], msg["From"], msg["To"] = subject, cfg["u"], email; msg.set_content(body)
                        with open(file_path, "rb") as f:
                            ct, _ = mimetypes.guess_type(file_path); mn, sb = (ct or 'application/octet-stream').split('/', 1)
                            msg.add_attachment(f.read(), maintype=mn, subtype=sb, filename=os.path.basename(file_path))
                        srv.send_message(msg); ok += 1
                    except: err += 1
                    Clock.schedule_once(lambda dt: setattr(pb, 'value', i+1))
                srv.quit()
                Clock.schedule_once(lambda dt: [setattr(lbl, 'text', f"UKOŃCZONO\nSukces: {ok} | Błąd: {err}"), setattr(btn_c, 'disabled', False)])
                self.conn.execute("INSERT INTO reports (date, ok, fail, skip, auto, details) VALUES (?,?,?,?,?,?)", (f"{datetime.now().strftime('%m-%d %H:%M')} (PLIK)", ok, err, 0, 0, f"Wysyłka specjalna pliku: {os.path.basename(file_path)}")); self.conn.commit()
            except Exception as e: Clock.schedule_once(lambda dt: [setattr(lbl, 'text', f"Błąd krytyczny: {str(e)[:50]}..."), setattr(btn_c, 'disabled', False)])
        threading.Thread(target=run, daemon=True).start()

    def start_mass_mailing(self, _):
        if not self.full_data: self.msg("!", "Danych z arkusza brak!"); return
        self.stats = {"ok": 0, "fail": 0, "skip": 0, "auto": 0}; self.session_details = []
        self.queue = list(self.full_data[1:]); self.total_q = len(self.queue)
        self.pb.value = 0; Clock.schedule_once(self.process_mailing_queue)

    def process_mailing_queue(self, *args):
        done = self.total_q - len(self.queue)
        self.pb.value = int((done/self.total_q)*100) if self.total_q > 0 else 100
        self.pb_label.text = f"Postęp: {done} / {self.total_q}"
        if not self.queue:
            det = "\n".join(self.session_details)
            self.conn.execute("INSERT INTO reports (date, ok, fail, skip, auto, details) VALUES (?,?,?,?,?,?)", (datetime.now().strftime("%Y-%m-%d %H:%M"), self.stats['ok'], self.stats['fail'], self.stats['skip'], self.stats['auto'], det)); self.conn.commit()
            self.msg("Koniec", f"Zakończono wysyłkę!\nSukces: {self.stats['ok']}\nPominięto: {self.stats['skip']}"); return
        row = self.queue.pop(0)
        try:
            n, s = str(row[self.idx_name]).strip(), str(row[self.idx_surname]).strip(); p = str(row[self.idx_pesel]).strip() if self.idx_pesel != -1 else ""
            res = None
            if self.auto_send_mode: res = self.conn.execute("SELECT email FROM contacts WHERE name=? AND surname=? COLLATE NOCASE", (n.lower(), s.lower())).fetchone()
            elif p and len(p) > 5:
                res = self.conn.execute("SELECT email FROM contacts WHERE pesel=?", (p,)).fetchone()
                if not res: res = self.conn.execute("SELECT email FROM contacts WHERE name=? AND surname=? COLLATE NOCASE", (n.lower(), s.lower())).fetchone()
            else: res = self.conn.execute("SELECT email FROM contacts WHERE name=? AND surname=? COLLATE NOCASE", (n.lower(), s.lower())).fetchone()
            
            if res:
                if self.auto_send_mode: self.stats["auto"] += 1; self.send_email_engine(row, res[0])
                else: self.ask_before_send(row, res[0], n, s, p)
            else:
                self.stats["skip"] += 1; self.session_details.append(f"SKIP: {n} {s} (Brak w bazie)"); Clock.schedule_once(self.process_mailing_queue)
        except: self.stats["skip"] += 1; Clock.schedule_once(self.process_mailing_queue)

    def ask_before_send(self, row, email, n, s, p_file):
        box = BoxLayout(orientation="vertical", padding=dp(20), spacing=dp(15))
        box.add_widget(Label(text=f"WYSŁAĆ DO:\n[b]{n} {s}[/b]\n[color=1eeaff]{email}[/color]", markup=True, halign="center"))
        btns = BoxLayout(size_hint_y=None, height=dp(55), spacing=dp(10))
        def dec(v):
            px.dismiss()
            if v: self.send_email_engine(row, email)
            else: self.stats["skip"] += 1; self.session_details.append(f"SKIP: {n} {s} (Odmowa)"); Clock.schedule_once(self.process_mailing_queue)
        btns.add_widget(ModernButton(text="WYŚLIJ", on_press=lambda x: dec(True), bg_color=COLOR_ACCENT))
        btns.add_widget(ModernButton(text="POMIŃ", on_press=lambda x: dec(False), bg_color=COLOR_DANGER))
        box.add_widget(btns); px = Popup(title="Weryfikacja", content=box, size_hint=(0.9, 0.4), auto_dismiss=False); px.open()

    def send_email_engine(self, row_data, target):
        def thread_task():
            cfg_p = Path(self.user_data_dir) / "smtp.json"
            if not cfg_p.exists(): return
            cfg = json.load(open(cfg_p))
            nx, sx = str(row_data[self.idx_name]).title(), str(row_data[self.idx_surname]).title()
            try:
                srv = smtplib.SMTP("smtp.gmail.com", 587, timeout=15); srv.starttls(); srv.login(cfg['u'], cfg['p'])
                msg = EmailMessage(); dat = datetime.now().strftime("%d.%m.%Y")
                ts = self.conn.execute("SELECT val FROM settings WHERE key='t_sub'").fetchone()
                tb = self.conn.execute("SELECT val FROM settings WHERE key='t_body'").fetchone()
                msg["Subject"] = (ts[0] if ts else "Raport").replace("{Imię}", nx); msg["From"], msg["To"] = cfg['u'], target
                msg.set_content((tb[0] if tb else "Dzień dobry").replace("{Imię}", nx).replace("{Data}", dat))
                tmp = Path(self.user_data_dir) / "r_tmp.xlsx"; wb = Workbook(); ws = wb.active
                ws.append([self.full_data[0][k] for k in self.export_indices])
                ws.append([str(row_data[k]) if (k < len(row_data) and str(row_data[k]).strip() != "") else "0" for k in self.export_indices])
                self.style_xlsx(ws); wb.save(tmp)
                msg.add_attachment(open(tmp, "rb").read(), maintype="application", subtype="xlsx", filename=f"Raport_{nx}_{sx}.xlsx")
                for path in self.global_attachments:
                    if os.path.exists(path):
                        ct, _ = mimetypes.guess_type(path); mn, sb = (ct or 'application/octet-stream').split('/', 1)
                        msg.add_attachment(open(path, "rb").read(), maintype=mn, subtype=sb, filename=os.path.basename(path))
                srv.send_message(msg); srv.quit(); self.session_details.append(f"OK: {nx} {sx}")
                Clock.schedule_once(lambda d: [self.update_stat("ok"), self.process_mailing_queue()])
            except Exception as e:
                self.session_details.append(f"ERR: {nx} {sx} ({str(e)[:20]})")
                Clock.schedule_once(lambda d: [self.update_stat("fail"), self.process_mailing_queue()])
        threading.Thread(target=thread_task, daemon=True).start()

    def style_xlsx(self, ws):
        s, c = Side(style='thin'), Alignment(horizontal='center', vertical='center')
        for ri, row in enumerate(ws.iter_rows(), 1):
            for cell in row:
                cell.border = Border(top=s, left=s, right=s, bottom=s); cell.alignment = c
                if ri == 1: cell.font = Font(bold=True); cell.fill = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")
                elif ri % 2 == 0: cell.fill = PatternFill(start_color="F7F7F7", end_color="F7F7F7", fill_type="solid")
        for col in ws.columns:
            max_l = 0; column_letter = col[0].column_letter
            for cell in col:
                if cell.value: max_l = max(max_l, len(str(cell.value)))
            ws.column_dimensions[column_letter].width = (max_l * 1.3) + 6

    def show_report_details(self, text):
        box = BoxLayout(orientation="vertical", padding=dp(10), spacing=dp(5))
        ti = TextInput(text=str(text) if text else "Brak logów", readonly=True, font_size='11sp', size_hint_y=0.8, background_color=(0.1, 0.1, 0.1, 1), foreground_color=(0.8, 0.8, 0.8, 1))
        btn = ModernButton(text="ZAMKNIJ", size_hint_y=0.2, on_press=lambda x: p.dismiss())
        box.add_widget(ti); box.add_widget(btn); p = Popup(title="Szczegóły sesji", content=box, size_hint=(0.95, 0.8)); p.open()

    def setup_table_ui(self):
        root = BoxLayout(orientation="vertical"); menu = BoxLayout(size_hint_y=None, height=dp(60), spacing=dp(5), padding=dp(8))
        self.ti_search = ModernInput(hint_text="Filtruj tabelę...", multiline=False); self.ti_search.bind(text=self.filter_table)
        menu.add_widget(self.ti_search); menu.add_widget(ModernButton(text="⚙", size_hint_x=None, width=dp(55), on_press=self.popup_columns))
        menu.add_widget(ModernButton(text="↩", size_hint_x=None, width=dp(55), on_press=lambda x: setattr(self.sm, 'current', 'home')))
        hs = ScrollView(size_hint_y=None, height=dp(55), do_scroll_y=False); self.table_header_layout = GridLayout(rows=1, size_hint=(None, None), height=dp(55))
        hs.add_widget(self.table_header_layout); ds = ScrollView(do_scroll_x=True, do_scroll_y=True); self.table_content_layout = GridLayout(size_hint=(None, None))
        self.table_content_layout.bind(minimum_height=self.table_content_layout.setter('height'), minimum_width=self.table_content_layout.setter('width'))
        ds.add_widget(self.table_content_layout); ds.bind(scroll_x=lambda inst, val: setattr(hs, 'scroll_x', val))
        root.add_widget(menu); root.add_widget(hs); root.add_widget(ds); self.screens["table"].add_widget(root)

    def refresh_table(self):
        self.table_content_layout.clear_widgets(); self.table_header_layout.clear_widgets()
        if not self.filtered_data or not self.export_indices: return
        w, h = dp(180), dp(55); headers = [self.full_data[0][i] for i in self.export_indices]
        self.table_header_layout.cols = len(headers) + 1; self.table_header_layout.width = (len(headers) + 1) * w
        for head in headers: self.table_header_layout.add_widget(ColorSafeLabel(text=str(head), bg_color=COLOR_CARD, text_color=COLOR_ACCENT, bold=True, size=(w,h), size_hint=(None,None)))
        self.table_header_layout.add_widget(ColorSafeLabel(text="OPCJE", bg_color=COLOR_CARD, text_color=COLOR_ACCENT, bold=True, size=(w,h), size_hint=(None,None)))
        self.table_content_layout.cols = len(headers) + 1; self.table_content_layout.width = (len(headers)+1)*w
        for r_idx, row in enumerate(self.filtered_data[1:]):
            row_bg = (0.08, 0.1, 0.15, 1) if r_idx % 2 == 0 else (0.12, 0.15, 0.2, 1)
            for c_idx in self.export_indices: 
                val = row[c_idx] if c_idx < len(row) else ""
                self.table_content_layout.add_widget(ColorSafeLabel(text=str(val), bg_color=row_bg, text_color=COLOR_TEXT, size=(w,h), size_hint=(None,None)))
            bt = Button(text="EKSPORT", size=(w,h), size_hint=(None,None), background_color=(0.1, 0.5, 0.3, 1)); bt.bind(on_press=lambda x, r=row: self.export_xlsx(r)); self.table_content_layout.add_widget(bt)

    def setup_smtp_ui(self):
        l = BoxLayout(orientation="vertical", padding=dp(25), spacing=dp(12)); self.ti_su = ModernInput(hint_text="Email Gmail"); self.ti_sp = ModernInput(hint_text="Hasło aplikacji", password=True)
        p = Path(self.user_data_dir) / "smtp.json"; d = json.load(open(p)) if p.exists() else {}; self.ti_su.text, self.ti_sp.text = d.get('u',''), d.get('p','')
        l.add_widget(Label(text="KONFIGURACJA GMAIL", bold=True, color=COLOR_PRIMARY)); l.add_widget(self.ti_su); l.add_widget(self.ti_sp)
        l.add_widget(ModernButton(text="ZAPISZ USTAWIENIA", on_press=lambda x: [json.dump({'u':self.ti_su.text, 'p':self.ti_sp.text}, open(p, "w")), self.msg("OK", "Zapisano!")]))
        l.add_widget(ModernButton(text="TESTUJ POŁĄCZENIE", on_press=self.test_smtp, bg_color=COLOR_ACCENT))
        l.add_widget(ModernButton(text="POWRÓT", on_press=lambda x: setattr(self.sm, 'current', 'home'), bg_color=(0.3,0.3,0.3,1))); self.screens["smtp"].add_widget(l)

    def setup_tmpl_ui(self):
        l = BoxLayout(orientation="vertical", padding=dp(25), spacing=dp(10)); self.ti_ts = ModernInput(hint_text="Temat (użyj {Imię})"); self.ti_tb = ModernInput(hint_text="Treść (użyj {Imię}, {Data})", multiline=True)
        ts = self.conn.execute("SELECT val FROM settings WHERE key='t_sub'").fetchone(); tb = self.conn.execute("SELECT val FROM settings WHERE key='t_body'").fetchone(); self.ti_ts.text = ts[0] if ts else ""; self.ti_tb.text = tb[0] if tb else ""
        sv = lambda x: [self.conn.execute("INSERT OR REPLACE INTO settings VALUES (?,?)", ('t_sub', self.ti_ts.text)), self.conn.execute("INSERT OR REPLACE INTO settings VALUES (?,?)", ('t_body', self.ti_tb.text)), self.conn.commit(), self.msg("OK", "Szablon zapisany")]
        l.add_widget(Label(text="SZABLON WIADOMOŚCI", bold=True, color=COLOR_PRIMARY)); l.add_widget(self.ti_ts); l.add_widget(self.ti_tb); l.add_widget(ModernButton(text="ZAPISZ SZABLON", on_press=sv)); l.add_widget(ModernButton(text="POWRÓT", on_press=lambda x: setattr(self.sm, 'current', 'email'))); self.screens["tmpl"].add_widget(l)

    def process_excel(self, path):
        try:
            if str(path).endswith(".xls") and xlrd:
                wb = xlrd.open_workbook(path); ws = wb.sheet_by_index(0); raw = [[str(ws.cell_value(r,c)).strip() for c in range(ws.ncols)] for r in range(ws.nrows)]
            else:
                wb = load_workbook(path, data_only=True); ws = wb.active; raw = [["" if v is None else str(v).strip() for v in r] for r in ws.iter_rows(values_only=True)]
            h_idx = 0
            for i, row in enumerate(raw[:15]):
                ln = " ".join([str(x) for x in row]).lower()
                if any(x in ln for x in ["imię", "imie", "nazwisko", "pesel"]): h_idx = i; break
            self.full_data = raw[h_idx:]; self.filtered_data = self.full_data; self.export_indices = list(range(len(self.full_data[0])))
            for i,v in enumerate([str(x).lower() for x in self.full_data[0]]):
                if "imi" in v: self.idx_name = i
                if "naz" in v: self.idx_surname = i
                if "pesel" in v: self.idx_pesel = i
            self.msg("Sukces", "Arkusz wczytany poprawnie.")
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
                    self.conn.execute("INSERT OR REPLACE INTO contacts (name, surname, email, pesel, phone) VALUES (?,?,?,?,?)", (r[iN].lower(), r[iS].lower(), str(r[iE]).strip(), str(r[iP]) if (iP != -1 and len(r) > iP) else "", ""))
            self.conn.commit(); self.update_stats(); self.msg("Sukces", "Baza kontaktów zaktualizowana.")
        except: self.msg("Błąd", "Nieprawidłowy format pliku bazy.")

    def export_xlsx(self, r):
        p = Path("/storage/emulated/0/Documents/FutureExport") if platform=="android" else Path("./exports"); p.mkdir(parents=True, exist_ok=True)
        nx, sx = str(r[self.idx_name]).title(), str(r[self.idx_surname]).title(); wb = Workbook(); ws = wb.active
        ws.append([self.full_data[0][k] for k in self.export_indices])
        ws.append([str(r[k]) if (k < len(r) and str(r[k]).strip() != "") else "0" for k in self.export_indices])
        self.style_xlsx(ws); wb.save(p/f"Raport_{nx}_{sx}.xlsx"); self.msg("OK", f"Zapisano PDF dla {nx}")

    def filter_table(self, ins, val):
        v = val.lower(); self.filtered_data = [self.full_data[0]] + [r for r in self.full_data[1:] if any(v in str(c).lower() for c in r)]; self.refresh_table()

    def popup_columns(self, _):
        if not self.full_data: return
        box = BoxLayout(orientation="vertical", padding=dp(15)); gr = GridLayout(cols=1, size_hint_y=None, spacing=dp(8)); gr.bind(minimum_height=gr.setter('height')); checks = []
        for i, h in enumerate(self.full_data[0]):
            r = BoxLayout(size_hint_y=None, height=dp(50)); cb = CheckBox(active=(i in self.export_indices), size_hint_x=None, width=dp(50)); checks.append((i, cb)); r.add_widget(cb); r.add_widget(Label(text=str(h))); gr.add_widget(r)
        sc = ScrollView(); sc.add_widget(gr); box.add_widget(sc); box.add_widget(ModernButton(text="ZASTOSUJ", on_press=lambda x: [setattr(self, 'export_indices', [idx for idx, c in checks if c.active]), p.dismiss(), self.refresh_table()])); p = Popup(title="Widoczne kolumny", content=box, size_hint=(0.9, 0.9)); p.open()

    def form_contact(self, n="", s="", e="", pes="", ph=""):
        b = BoxLayout(orientation="vertical", padding=dp(15), spacing=dp(10)); flds = [ModernInput(text=n, hint_text="Imię"), ModernInput(text=s, hint_text="Nazwisko"), ModernInput(text=e, hint_text="Email"), ModernInput(text=pes, hint_text="PESEL"), ModernInput(text=ph, hint_text="Telefon")]
        for f in flds: b.add_widget(f)
        def save(_): [self.conn.execute("INSERT OR REPLACE INTO contacts VALUES (?,?,?,?,?)", [f.text.strip().lower() if i<2 else f.text.strip() for i,f in enumerate(flds)]), self.conn.commit(), px.dismiss(), self.refresh_contacts_list(), self.update_stats()]
        b.add_widget(ModernButton(text="ZAPISZ KONTAKT", on_press=save, bg_color=COLOR_ACCENT)); px = Popup(title="Dane kontaktu", content=b, size_hint=(0.9, 0.85)); px.open()

    def update_stat(self, k): self.stats[k]+=1
    def update_stats(self, *a):
        try: self.lbl_stats.text = f"BAZA: {self.conn.execute('SELECT count(*) FROM contacts').fetchone()[0]} | ZAŁĄCZNIKI: {len(self.global_attachments)}"
        except: pass

    def test_smtp(self, _):
        p = Path(self.user_data_dir) / "smtp.json"; cfg = json.load(open(p)) if p.exists() else None
        def tk():
            try: srv = smtplib.SMTP("smtp.gmail.com", 587, timeout=10); srv.starttls(); srv.login(cfg['u'], cfg['p']); srv.quit(); Clock.schedule_once(lambda d: self.msg("OK", "Połączenie z Gmail prawidłowe!"))
            except Exception as e: Clock.schedule_once(lambda d: self.msg("BŁĄD", f"Błąd logowania: {str(e)[:40]}"))
        if cfg: threading.Thread(target=tk, daemon=True).start()
        else: self.msg("!", "Najpierw zapisz ustawienia")

    def msg(self, tit, txt):
        b = BoxLayout(orientation="vertical", padding=dp(20), spacing=dp(15))
        b.add_widget(Label(text=txt, halign="center", font_size='14sp'))
        btn = ModernButton(text="ROZUMIEM", height=dp(50), size_hint_y=None, on_press=lambda x: p.dismiss())
        b.add_widget(btn); p = Popup(title=tit, content=b, size_hint=(0.8, 0.35)); p.open()

if __name__ == "__main__":
    FutureApp().run()
