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

# Paleta Kolorów
COLOR_PRIMARY = (0.1, 0.55, 0.95, 1)  # Electric Blue
COLOR_SUCCESS = (0.15, 0.75, 0.5, 1)  # Emerald
COLOR_BG = (0.05, 0.06, 0.09, 1)     # Deep Navy
COLOR_CARD = (0.12, 0.14, 0.2, 1)     # Surface
COLOR_TEXT = (0.95, 0.95, 0.98, 1)
COLOR_DANGER = (0.9, 0.25, 0.25, 1)

# --- KOMPONENTY UI ---

class ModernButton(Button):
    def __init__(self, bg_color=COLOR_PRIMARY, **kwargs):
        super().__init__(**kwargs)
        self.background_normal = ""
        self.background_down = ""
        self.background_color = (0,0,0,0)
        self.color = COLOR_TEXT
        self.bold = True
        self.font_size = '15sp'
        self.radius = [dp(14)]
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
            self.line = RoundedRectangle(pos=self.pos, size=self.size, radius=[dp(10)])
        self.bind(pos=self._update, size=self._update)
    def _update(self, *args):
        self.line.pos = self.pos
        self.line.size = self.size

class CardRow(BoxLayout):
    def __init__(self, bg_color=COLOR_CARD, **kwargs):
        super().__init__(**kwargs)
        with self.canvas.before:
            Color(*bg_color)
            self.rect = RoundedRectangle(pos=self.pos, size=self.size, radius=[dp(16)])
        self.bind(pos=self._update, size=self._update)
    def _update(self, *args):
        m = dp(5)
        self.rect.pos = (self.pos[0] + m, self.pos[1] + m)
        self.rect.size = (self.size[0] - m*2, self.size[1] - m*2)

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

# --- EKRANY ---
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
        self.full_data, self.filtered_data, self.export_indices = [], [], []
        self.global_attachments, self.selected_emails, self.queue = [], [], []
        self.stats = {"ok": 0, "fail": 0, "skip": 0, "auto": 0}
        self.idx_name, self.idx_surname, self.idx_pesel = 0, 1, -1
        self.auto_send_mode = False
        
        self.init_db()
        self.sm = ScreenManager(transition=SlideTransition())
        self.add_all_screens()
        return self.sm

    def init_db(self):
        p = Path(self.user_data_dir) / "future_v3_core.db"
        self.conn = sqlite3.connect(str(p), check_same_thread=False)
        self.conn.execute("CREATE TABLE IF NOT EXISTS contacts (name TEXT, surname TEXT, email TEXT, pesel TEXT, phone TEXT, PRIMARY KEY(name, surname))")
        self.conn.execute("CREATE TABLE IF NOT EXISTS settings (key TEXT PRIMARY KEY, val TEXT)")
        self.conn.execute("CREATE TABLE IF NOT EXISTS reports (id INTEGER PRIMARY KEY AUTOINCREMENT, date TEXT, ok INTEGER, fail INTEGER, skip INTEGER, auto INTEGER, details TEXT)")
        self.conn.commit()

    def add_all_screens(self):
        self.sc_map = {
            "home": HomeScreen(name="home"), "table": TableScreen(name="table"),
            "email": EmailScreen(name="email"), "smtp": SMTPScreen(name="smtp"),
            "tmpl": TemplateScreen(name="tmpl"), "contacts": ContactsScreen(name="contacts"),
            "report": ReportScreen(name="report")
        }
        self.ui_home(); self.ui_email(); self.ui_contacts(); self.ui_report(); self.ui_smtp(); self.ui_tmpl(); self.ui_table()
        for s in self.sc_map.values(): self.sm.add_widget(s)

    # --- UI DESIGNS ---

    def ui_home(self):
        l = BoxLayout(orientation="vertical", padding=dp(30), spacing=dp(20))
        l.add_widget(Label(text="FUTURE 24", font_size='40sp', bold=True, color=COLOR_PRIMARY, size_hint_y=None, height=dp(100)))
        
        grid = GridLayout(cols=1, spacing=dp(15))
        btns = [
            ("📊 WCZYTAJ LISTĘ PŁAC", lambda x: self.open_picker("data"), COLOR_PRIMARY),
            ("👁️ PODGLĄD TABELI", lambda x: [self.refresh_table(), setattr(self.sm, 'current', 'table')] if self.full_data else self.msg("!", "Brak danych"), COLOR_CARD),
            ("📧 CENTRUM MAILINGU", lambda x: setattr(self.sm, 'current', 'email'), COLOR_SUCCESS),
            ("⚙️ USTAWIENIA", lambda x: setattr(self.sm, 'current', 'smtp'), (0.4, 0.4, 0.4, 1))
        ]
        for t, c, clr in btns: grid.add_widget(ModernButton(text=t, on_press=c, height=dp(65), size_hint_y=None, bg_color=clr))
        l.add_widget(grid)
        self.sc_map["home"].add_widget(l)

    def ui_email(self):
        l = BoxLayout(orientation="vertical", padding=dp(20), spacing=dp(12))
        
        stat_card = CardRow(size_hint_y=None, height=dp(110), orientation="vertical", padding=dp(15))
        self.lbl_stats = Label(text="Baza: 0 kontaktów", bold=True, color=COLOR_SUCCESS)
        stat_card.add_widget(self.lbl_stats)
        
        auto_box = BoxLayout(spacing=dp(10), size_hint_y=None, height=dp(40))
        self.cb_auto = CheckBox(size_hint_x=None, width=dp(40))
        self.cb_auto.bind(active=lambda i,v: setattr(self, 'auto_send_mode', v))
        auto_box.add_widget(self.cb_auto); auto_box.add_widget(Label(text="TRYB AUTOMATYCZNY (BEZ PYTANIA)", font_size='12sp'))
        stat_card.add_widget(auto_box)
        l.add_widget(stat_card)

        self.pb_label = Label(text="Gotowy", font_size='12sp', size_hint_y=None, height=dp(20))
        self.pb = ProgressBar(max=100, height=dp(10), size_hint_y=None)
        l.add_widget(self.pb_label); l.add_widget(self.pb)

        grid = GridLayout(cols=2, spacing=dp(10))
        acts = [
            ("👥 BAZA", lambda x: [self.refresh_contacts_list(), setattr(self.sm, 'current', 'contacts')]),
            ("📥 IMPORT BAZY", lambda x: self.open_picker("book")),
            ("📝 SZABLON", lambda x: setattr(self.sm, 'current', 'tmpl')),
            ("📜 RAPORTY", lambda x: [self.refresh_reports(), setattr(self.sm, 'current', 'report')]),
            ("📎 ZAŁĄCZNIK", lambda x: self.open_picker("attachment")),
            ("📁 WYŚLIJ PLIK", self.start_special_send_flow)
        ]
        for t, c in acts: grid.add_widget(ModernButton(text=t, on_press=c, bg_color=COLOR_CARD, height=dp(55), size_hint_y=None))
        l.add_widget(grid)
        
        l.add_widget(ModernButton(text="🚀 START WYSYŁKI", on_press=self.start_mass_mailing, bg_color=COLOR_PRIMARY, height=dp(65), size_hint_y=None))
        l.add_widget(ModernButton(text="POWRÓT", on_press=lambda x: setattr(self.sm, 'current', 'home'), bg_color=(0.3, 0.3, 0.3, 1), height=dp(50), size_hint_y=None))
        
        self.sc_map["email"].add_widget(l); self.update_stats()

    def ui_contacts(self):
        l = BoxLayout(orientation="vertical", padding=dp(15), spacing=dp(10))
        top = BoxLayout(size_hint_y=None, height=dp(55), spacing=dp(10))
        self.ti_csearch = ModernInput(hint_text="Szukaj...")
        self.ti_csearch.bind(text=self.refresh_contacts_list)
        top.add_widget(self.ti_csearch)
        top.add_widget(ModernButton(text="+", size_hint_x=None, width=dp(55), on_press=lambda x: self.form_contact()))
        
        self.c_list = GridLayout(cols=1, size_hint_y=None, spacing=dp(8))
        self.c_list.bind(minimum_height=self.c_list.setter('height'))
        sc = ScrollView(); sc.add_widget(self.c_list)
        l.add_widget(top); l.add_widget(sc)
        l.add_widget(ModernButton(text="POWRÓT", on_press=lambda x: setattr(self.sm, 'current', 'email'), height=dp(55), size_hint_y=None))
        self.sc_map["contacts"].add_widget(l)

    def ui_smtp(self):
        l = BoxLayout(orientation="vertical", padding=dp(30), spacing=dp(15))
        self.ti_su = ModernInput(hint_text="Email Gmail")
        self.ti_sp = ModernInput(hint_text="Hasło aplikacji", password=True)
        p = Path(self.user_data_dir) / "smtp.json"
        if p.exists():
            d = json.load(open(p)); self.ti_su.text = d.get('u',''); self.ti_sp.text = d.get('p','')
        
        l.add_widget(Label(text="USTAWIENIA GMAIL", bold=True, color=COLOR_PRIMARY, font_size='20sp'))
        l.add_widget(self.ti_su); l.add_widget(self.ti_sp)
        l.add_widget(ModernButton(text="ZAPISZ", on_press=lambda x: [json.dump({'u':self.ti_su.text,'p':self.ti_sp.text}, open(p, "w")), self.msg("OK", "Zapisano")]))
        l.add_widget(ModernButton(text="TESTUJ POŁĄCZENIE", bg_color=COLOR_SUCCESS, on_press=self.test_smtp))
        l.add_widget(ModernButton(text="POWRÓT", bg_color=(0.3,0.3,0.3,1), on_press=lambda x: setattr(self.sm, 'current', 'home')))
        self.sc_map["smtp"].add_widget(l)

    # --- LOGIKA BAZY I IMPORTU ---

    def refresh_contacts_list(self, *args):
        self.c_list.clear_widgets()
        sv = self.ti_csearch.text.lower()
        rows = self.conn.execute("SELECT name, surname, email, pesel, phone FROM contacts ORDER BY surname ASC").fetchall()
        for n, s, e, p, ph in rows:
            if sv and sv not in f"{n} {s} {e} {p}".lower(): continue
            card = CardRow(size_hint_y=None, height=dp(100), padding=dp(12))
            
            cb = CheckBox(size_hint_x=None, width=dp(40), active=(e in self.selected_emails))
            cb.bind(active=lambda inst, v, m=e: self.selected_emails.append(m) if v else self.selected_emails.remove(m))
            card.add_widget(cb)
            
            txt = BoxLayout(orientation="vertical")
            txt.add_widget(Label(text=f"{n} {s}".upper(), bold=True, halign="left", text_size=(dp(200),None)))
            txt.add_widget(Label(text=e, font_size='11sp', color=COLOR_PRIMARY, halign="left", text_size=(dp(200),None)))
            card.add_widget(txt)
            
            btns = BoxLayout(orientation="vertical", size_hint_x=None, width=dp(80), spacing=dp(4))
            btns.add_widget(Button(text="EDYTUJ", font_size='10sp', background_color=(0.2,0.6,1,1), on_press=lambda x, d=(n,s,e,p,ph): self.form_contact(*d)))
            # NAPRAWIONA FUNKCJA USUWANIA:
            btns.add_widget(Button(text="USUŃ", font_size='10sp', background_color=COLOR_DANGER, on_press=lambda x, na=n, su=s: self.delete_contact(na, su)))
            card.add_widget(btns)
            self.c_list.add_widget(card)

    def delete_contact(self, n, s):
        def confirm(_):
            self.conn.execute("DELETE FROM contacts WHERE name=? AND surname=?", (n, s))
            self.conn.commit()
            p.dismiss(); self.refresh_contacts_list(); self.update_stats()
        content = BoxLayout(orientation="vertical", padding=dp(15), spacing=dp(10))
        content.add_widget(Label(text=f"Usunąć {n} {s}?"))
        btns = BoxLayout(size_hint_y=None, height=dp(50), spacing=dp(10))
        btns.add_widget(ModernButton(text="TAK", bg_color=COLOR_DANGER, on_press=confirm))
        btns.add_widget(ModernButton(text="NIE", on_press=lambda x: p.dismiss()))
        content.add_widget(btns)
        p = Popup(title="Usuwanie", content=content, size_hint=(0.8, 0.3)); p.open()

    def process_book(self, path):
        try:
            wb = load_workbook(path, data_only=True); ws = wb.active
            rows = list(ws.iter_rows(values_only=True))
            h = [str(x).lower() for x in rows[0]]
            iN, iS, iE, iP = 0, 1, 2, -1
            for i,v in enumerate(h):
                if "imi" in v: iN=i
                elif "naz" in v: iS=i
                elif "@" in v or "mail" in v: iE=i
                elif "pesel" in v: iP=i
            c = 0
            for r in rows[1:]:
                if r[iN] and r[iE]:
                    self.conn.execute("INSERT OR REPLACE INTO contacts VALUES (?,?,?,?,?)", (str(r[iN]).strip().lower(), str(r[iS]).strip().lower(), str(r[iE]).strip(), str(r[iP]) if (iP!=-1) else "", ""))
                    c += 1
            self.conn.commit(); self.update_stats(); self.msg("Sukces", f"Zaimportowano {c} osób.")
        except Exception as e: self.msg("Błąd", str(e))

    # --- WYSYŁKA ---

    def start_mass_mailing(self, _):
        if not self.full_data: self.msg("!", "Brak danych z arkusza!"); return
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
            self.msg("Koniec", f"Wysłano: {self.stats['ok']}\nBłędy: {self.stats['fail']}"); return
        
        row = self.queue.pop(0)
        n, s = str(row[self.idx_name]).strip(), str(row[self.idx_surname]).strip()
        p = str(row[self.idx_pesel]).strip() if self.idx_pesel != -1 else ""
        
        res = self.conn.execute("SELECT email FROM contacts WHERE (name=? AND surname=?) OR pesel=?", (n.lower(), s.lower(), p)).fetchone()
        if res:
            if self.auto_send_mode: self.send_email_engine(row, res[0])
            else: self.ask_send(row, res[0], n, s)
        else:
            self.stats["skip"] += 1; self.process_mailing_queue()

    def ask_send(self, row, email, n, s):
        b = BoxLayout(orientation="vertical", padding=dp(20), spacing=dp(15))
        b.add_widget(Label(text=f"Wysłać do:\n{n} {s}\n({email})", halign="center"))
        btns = BoxLayout(size_hint_y=None, height=dp(55), spacing=dp(10))
        def dec(v):
            p.dismiss()
            if v: self.send_email_engine(row, email)
            else: self.stats["skip"] += 1; self.process_mailing_queue()
        btns.add_widget(ModernButton(text="WYŚLIJ", on_press=lambda x: dec(True), bg_color=COLOR_SUCCESS))
        btns.add_widget(ModernButton(text="POMIŃ", on_press=lambda x: dec(False), bg_color=COLOR_DANGER))
        b.add_widget(btns); p = Popup(title="Potwierdzenie", content=b, size_hint=(.9, .4)); p.open()

    def send_email_engine(self, row_data, target):
        def run():
            cfg_p = Path(self.user_data_dir) / "smtp.json"
            if not cfg_p.exists(): return
            cfg = json.load(open(cfg_p))
            try:
                srv = smtplib.SMTP("smtp.gmail.com", 587, timeout=15); srv.starttls(); srv.login(cfg['u'], cfg['p'])
                msg = EmailMessage(); nx = str(row_data[self.idx_name]).title()
                ts = self.conn.execute("SELECT val FROM settings WHERE key='t_sub'").fetchone()
                tb = self.conn.execute("SELECT val FROM settings WHERE key='t_body'").fetchone()
                msg["Subject"] = (ts[0] if ts else "Raport").replace("{Imię}", nx)
                msg["From"], msg["To"] = cfg['u'], target
                msg.set_content((tb[0] if tb else "Witaj").replace("{Imię}", nx))
                
                # Załącznik Excel personalizowany
                tmp = Path(self.user_data_dir) / "send.xlsx"; wb = Workbook(); ws = wb.active
                ws.append([self.full_data[0][k] for k in self.export_indices])
                ws.append([str(row_data[k]) for k in self.export_indices])
                wb.save(tmp)
                msg.add_attachment(open(tmp, "rb").read(), maintype="application", subtype="xlsx", filename=f"Raport_{nx}.xlsx")
                
                srv.send_message(msg); srv.quit()
                self.stats["ok"] += 1; self.session_details.append(f"OK: {target}")
            except: self.stats["fail"] += 1
            Clock.schedule_once(lambda d: self.process_mailing_queue())
        threading.Thread(target=run, daemon=True).start()

    # --- PICKERY I POMOCNICZE ---

    def open_picker(self, mode):
        if platform != "android":
            self.msg("PC", "Działa tylko na Android")
            return
        from jnius import autoclass; from android import activity
        PA = autoclass("org.kivy.android.PythonActivity"); Intent = autoclass("android.content.Intent")
        intent = Intent(Intent.ACTION_GET_CONTENT); intent.setType("*/*")
        def cb(req, res, dt):
            if req != 77: return
            activity.unbind(on_activity_result=cb)
            if res == -1 and dt:
                uri = dt.getData(); reslvr = PA.mActivity.getContentResolver()
                dest = Path(self.user_data_dir) / f"file_{mode}.xlsx"
                try:
                    stream = reslvr.openInputStream(uri)
                    with open(dest, "wb") as f: f.write(stream.readAllBytes() if hasattr(stream, 'readAllBytes') else stream.read())
                    stream.close()
                    if mode == "data": self.process_excel(dest)
                    elif mode == "book": self.process_book(dest)
                    elif mode == "attachment": self.global_attachments.append(str(dest)); self.update_stats()
                except: self.msg("!", "Błąd pliku")
        activity.bind(on_activity_result=cb); PA.mActivity.startActivityForResult(intent, 77)

    def process_excel(self, p):
        try:
            wb = load_workbook(p, data_only=True); ws = wb.active
            self.full_data = [["" if v is None else str(v).strip() for v in r] for r in ws.iter_rows(values_only=True)]
            for i,v in enumerate([str(x).lower() for x in self.full_data[0]]):
                if "imi" in v: self.idx_name = i
                if "naz" in v: self.idx_surname = i
                if "pesel" in v: self.idx_pesel = i
            self.export_indices = list(range(len(self.full_data[0])))
            self.msg("OK", "Arkusz wczytany.")
        except Exception as e: self.msg("Błąd", str(e))

    def test_smtp(self, _):
        p = Path(self.user_data_dir) / "smtp.json"
        if not p.exists(): return
        cfg = json.load(open(p))
        def tk():
            try:
                s = smtplib.SMTP("smtp.gmail.com", 587, timeout=10); s.starttls(); s.login(cfg['u'], cfg['p']); s.quit()
                Clock.schedule_once(lambda d: self.msg("OK", "Połączono!"))
            except Exception as e: Clock.schedule_once(lambda d: self.msg("Błąd", str(e)[:40]))
        threading.Thread(target=tk, daemon=True).start()

    def update_stats(self, *a):
        try:
            c = self.conn.execute("SELECT count(*) FROM contacts").fetchone()[0]
            self.lbl_stats.text = f"BAZA: {c} kontaktów | ZAŁĄCZNIKI: {len(self.global_attachments)}"
        except: pass

    def msg(self, tit, txt):
        b = BoxLayout(orientation="vertical", padding=dp(20), spacing=dp(15))
        b.add_widget(Label(text=txt, halign="center"))
        b.add_widget(ModernButton(text="OK", on_press=lambda x: p.dismiss(), height=dp(50), size_hint_y=None))
        p = Popup(title=tit, content=b, size_hint=(0.8, 0.35)); p.open()

    # --- BRAKUJĄCE UI ---
    def ui_report(self):
        l = BoxLayout(orientation="vertical", padding=dp(15), spacing=dp(10))
        l.add_widget(Label(text="HISTORIA", bold=True, color=COLOR_PRIMARY, size_hint_y=None, height=dp(40)))
        self.r_grid = GridLayout(cols=1, size_hint_y=None, spacing=dp(10)); self.r_grid.bind(minimum_height=self.r_grid.setter('height'))
        sc = ScrollView(); sc.add_widget(self.r_grid); l.add_widget(sc)
        l.add_widget(ModernButton(text="POWRÓT", on_press=lambda x: setattr(self.sm, 'current', 'email'), height=dp(55), size_hint_y=None))
        self.sc_map["report"].add_widget(l)

    def refresh_reports(self):
        self.r_grid.clear_widgets()
        rows = self.conn.execute("SELECT date, ok, fail FROM reports ORDER BY id DESC LIMIT 20").fetchall()
        for d, ok, fl in rows:
            c = CardRow(size_hint_y=None, height=dp(70), padding=dp(10))
            c.add_widget(Label(text=d, bold=True, font_size='12sp')); c.add_widget(Label(text=f"S:{ok} F:{fl}"))
            self.r_grid.add_widget(c)

    def ui_tmpl(self):
        l = BoxLayout(orientation="vertical", padding=dp(25), spacing=dp(10))
        self.ti_ts = ModernInput(hint_text="Temat {Imię}")
        self.ti_tb = ModernInput(hint_text="Treść {Imię}", multiline=True)
        l.add_widget(Label(text="SZABLON", bold=True, color=COLOR_PRIMARY))
        l.add_widget(self.ti_ts); l.add_widget(self.ti_tb)
        sv = lambda x: [self.conn.execute("INSERT OR REPLACE INTO settings VALUES (?,?)",('t_sub',self.ti_ts.text)),self.conn.execute("INSERT OR REPLACE INTO settings VALUES (?,?)",('t_body',self.ti_tb.text)),self.conn.commit(),self.msg("OK","Zapisano")]
        l.add_widget(ModernButton(text="ZAPISZ", on_press=sv))
        l.add_widget(ModernButton(text="POWRÓT", on_press=lambda x: setattr(self.sm, 'current', 'email')))
        self.sc_map["tmpl"].add_widget(l)

    def ui_table(self):
        l = BoxLayout(orientation="vertical")
        self.table_content = GridLayout(size_hint=(None, None), cols=1)
        self.table_content.bind(minimum_height=self.table_content.setter('height'), minimum_width=self.table_content.setter('width'))
        sc = ScrollView(do_scroll_x=True, do_scroll_y=True); sc.add_widget(self.table_content)
        l.add_widget(sc); l.add_widget(ModernButton(text="POWRÓT", on_press=lambda x: setattr(self.sm, 'current', 'home'), size_hint_y=None, height=dp(55)))
        self.sc_map["table"].add_widget(l)

    def refresh_table(self):
        self.table_content.clear_widgets()
        if not self.full_data: return
        self.table_content.cols = len(self.full_data[0])
        self.table_content.width = len(self.full_data[0]) * dp(150)
        for row in self.full_data[:100]:
            for cell in row: self.table_content.add_widget(ColorSafeLabel(text=str(cell), size=(dp(150), dp(40)), size_hint=(None,None), bg_color=COLOR_CARD, text_color=COLOR_TEXT))

    def start_special_send_flow(self, _): self.msg("Info", "Wybierz plik przez 'Załącznik' i użyj startu.")
    def form_contact(self, n="", s="", e="", pes="", ph=""):
        b = BoxLayout(orientation="vertical", padding=dp(15), spacing=dp(10))
        ins = [ModernInput(text=n, hint_text="Imię"), ModernInput(text=s, hint_text="Nazwisko"), ModernInput(text=e, hint_text="Email"), ModernInput(text=pes, hint_text="PESEL")]
        for i in ins: b.add_widget(i)
        def save(_):
            self.conn.execute("INSERT OR REPLACE INTO contacts VALUES (?,?,?,?,?)", (ins[0].text.lower(), ins[1].text.lower(), ins[2].text, ins[3].text, ""))
            self.conn.commit(); p.dismiss(); self.refresh_contacts_list(); self.update_stats()
        b.add_widget(ModernButton(text="Zapisz", on_press=save)); p = Popup(title="Kontakt", content=b, size_hint=(.9, .8)); p.open()

if __name__ == "__main__":
    FutureApp().run()
