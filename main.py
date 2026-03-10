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

# Paleta kolorów (zostawiamy UI które Ci się spodobało)
COLOR_PRIMARY = (0.1, 0.55, 0.95, 1)
COLOR_SUCCESS = (0.15, 0.75, 0.5, 1)
COLOR_BG = (0.05, 0.06, 0.09, 1)
COLOR_CARD = (0.12, 0.14, 0.2, 1)
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
        self.rect.size, self.rect.pos = self.size, self.pos
        self.text_size = (self.width - dp(10), None)

# --- EKRANY ---
class HomeScreen(Screen): pass
class TableScreen(Screen): pass
class EmailScreen(Screen): pass
class SMTPScreen(Screen): pass
class TemplateScreen(Screen): pass
class ContactsScreen(Screen): pass
class ReportScreen(Screen): pass

# --- GŁÓWNA APLIKACJA ---

class FutureApp(App):
    def build(self):
        Window.clearcolor = COLOR_BG
        # Inicjalizacja zmiennych
        self.full_data, self.filtered_data, self.export_indices = [], [], []
        self.global_attachments, self.selected_emails, self.queue = [], [], []
        self.session_details = []
        self.stats = {"ok": 0, "fail": 0, "skip": 0, "auto": 0}
        self.idx_name, self.idx_surname, self.idx_pesel = 0, 1, -1
        self.auto_send_mode = False
        
        self.init_db()
        self.sm = ScreenManager(transition=SlideTransition())
        self.add_all_screens()
        return self.sm

    def init_db(self):
        db_p = Path(self.user_data_dir) / "future_v3_master.db"
        self.conn = sqlite3.connect(str(db_p), check_same_thread=False)
        self.conn.execute("CREATE TABLE IF NOT EXISTS contacts (name TEXT, surname TEXT, email TEXT, pesel TEXT, phone TEXT, PRIMARY KEY(name, surname))")
        self.conn.execute("CREATE TABLE IF NOT EXISTS settings (key TEXT PRIMARY KEY, val TEXT)")
        self.conn.execute("CREATE TABLE IF NOT EXISTS reports (id INTEGER PRIMARY KEY AUTOINCREMENT, date TEXT, ok INTEGER, fail INTEGER, skip INTEGER, auto INTEGER, details TEXT)")
        self.conn.commit()

    def add_all_screens(self):
        self.sc = {
            "home": HomeScreen(name="home"), "table": TableScreen(name="table"),
            "email": EmailScreen(name="email"), "smtp": SMTPScreen(name="smtp"),
            "tmpl": TemplateScreen(name="tmpl"), "contacts": ContactsScreen(name="contacts"),
            "report": ReportScreen(name="report")
        }
        self.setup_home_ui()
        self.setup_email_ui()
        self.setup_table_ui()
        self.setup_contacts_ui()
        self.setup_smtp_ui()
        self.setup_tmpl_ui()
        self.setup_report_ui()
        for s in self.sc.values(): self.sm.add_widget(s)

    # --- UI SETUP ---

    def setup_home_ui(self):
        l = BoxLayout(orientation="vertical", padding=dp(30), spacing=dp(20))
        l.add_widget(Label(text="FUTURE APK", font_size='38sp', bold=True, color=COLOR_PRIMARY, size_hint_y=None, height=dp(100)))
        grid = GridLayout(cols=1, spacing=dp(15))
        btns = [
            ("📊 WCZYTAJ LISTĘ PŁAC", lambda x: self.open_picker("data"), COLOR_PRIMARY),
            ("👁️ PODGLĄD I EKSPORT", lambda x: [self.refresh_table(), setattr(self.sm, 'current', 'table')] if self.full_data else self.msg("!", "Brak danych"), COLOR_CARD),
            ("📧 CENTRUM WYSYŁKI", lambda x: setattr(self.sm, 'current', 'email'), COLOR_SUCCESS),
            ("⚙️ USTAWIENIA GMAIL", lambda x: setattr(self.sm, 'current', 'smtp'), (0.3, 0.3, 0.3, 1))
        ]
        for t, c, clr in btns: grid.add_widget(ModernButton(text=t, on_press=c, height=dp(65), size_hint_y=None, bg_color=clr))
        l.add_widget(grid)
        self.sc["home"].add_widget(l)

    def setup_email_ui(self):
        l = BoxLayout(orientation="vertical", padding=dp(20), spacing=dp(12))
        info = CardRow(size_hint_y=None, height=dp(100), orientation="vertical", padding=dp(15))
        self.lbl_stats = Label(text="Baza: 0", bold=True, color=COLOR_SUCCESS)
        info.add_widget(self.lbl_stats)
        ab = BoxLayout(spacing=dp(10), size_hint_y=None, height=dp(30))
        self.cb_auto = CheckBox(size_hint_x=None, width=dp(40))
        self.cb_auto.bind(active=lambda i,v: setattr(self, 'auto_send_mode', v))
        ab.add_widget(self.cb_auto); ab.add_widget(Label(text="WYSYŁKA BEZ PYTANIA", font_size='12sp'))
        info.add_widget(ab)
        l.add_widget(info)

        prog_box = BoxLayout(orientation="vertical", size_hint_y=None, height=dp(50))
        self.pb_label = Label(text="Gotowy", font_size='12sp')
        self.pb = ProgressBar(max=100, height=dp(10)); prog_box.add_widget(self.pb_label); prog_box.add_widget(self.pb)
        l.add_widget(prog_box)

        grid = GridLayout(cols=2, spacing=dp(10))
        actions = [
            ("👥 BAZA", lambda x: [self.refresh_contacts_list(), setattr(self.sm, 'current', 'contacts')]),
            ("📥 IMPORT BAZY", lambda x: self.open_picker("book")),
            ("📝 SZABLON", lambda x: setattr(self.sm, 'current', 'tmpl')),
            ("📜 HISTORIA", lambda x: [self.refresh_reports(), setattr(self.sm, 'current', 'report')]),
            ("📎 ZAŁĄCZNIK", lambda x: self.open_picker("attachment")),
            ("📁 WYŚLIJ PLIK", self.start_special_send_flow)
        ]
        for t, c in actions: grid.add_widget(ModernButton(text=t, on_press=c, bg_color=COLOR_CARD, height=dp(55), size_hint_y=None))
        l.add_widget(grid)
        l.add_widget(ModernButton(text="🚀 URUCHOM MASOWĄ WYSYŁKĘ", on_press=self.start_mass_mailing, bg_color=COLOR_PRIMARY, height=dp(65), size_hint_y=None))
        l.add_widget(ModernButton(text="POWRÓT", on_press=lambda x: setattr(self.sm, 'current', 'home'), bg_color=(0.3, 0.3, 0.3, 1), height=dp(50), size_hint_y=None))
        self.sc["email"].add_widget(l); self.update_stats()

    # --- LOGIKA PICKERA (NAPRAWIONA) ---

    def open_picker(self, mode):
        if platform != "android":
            self.msg("Info", "Picker działa na Androidzie")
            return
        from jnius import autoclass; from android import activity
        PA = autoclass("org.kivy.android.PythonActivity"); Intent = autoclass("android.content.Intent")
        intent = Intent(Intent.ACTION_GET_CONTENT); intent.setType("*/*")
        
        def on_res(req, res, data):
            if req != 101: return
            activity.unbind(on_activity_result=on_res)
            if res == -1 and data:
                uri = data.getData(); resolver = PA.mActivity.getContentResolver()
                f_name = f"pick_{datetime.now().strftime('%M%S')}.xlsx"
                try:
                    stream = resolver.openInputStream(uri)
                    target = Path(self.user_data_dir) / f_name
                    with open(target, "wb") as f:
                        buf = bytearray(16384)
                        while True:
                            n = stream.read(buf)
                            if n <= 0: break
                            f.write(buf[:n])
                    stream.close()
                    # Przekierowanie logiki
                    if mode == "data": self.process_excel(target)
                    elif mode == "book": self.process_book(target)
                    elif mode == "attachment": 
                        self.global_attachments.append(str(target))
                        self.update_stats()
                    elif mode == "special_send":
                        Clock.schedule_once(lambda dt: self.special_send_step_2_recipients(str(target)))
                except: self.msg("Błąd", "Nie udało się odczytać pliku")
        
        activity.bind(on_activity_result=on_res)
        PA.mActivity.startActivityForResult(intent, 101)

    # --- WYSYŁKA SPECJALNA (PRZYWRÓCONA) ---

    def start_special_send_flow(self, _):
        if self.conn.execute("SELECT count(*) FROM contacts").fetchone()[0] == 0:
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
                r.add_widget(cb); r.add_widget(Label(text=f"{n} {s}\\n{e}", font_size='11sp')); gl.add_widget(r)
        ti.bind(text=lambda i,v: refresh(v)); refresh()
        btn = ModernButton(text="DALEJ"); btn.bind(on_press=lambda x: [p.dismiss(), self.special_send_step_3_msg(file_path)] if self.selected_emails else self.msg("!", "Wybierz kogoś!")); box.add_widget(btn); p = Popup(title="Odbiorcy", content=box, size_hint=(0.95,0.9)); p.open()

    def special_send_step_3_msg(self, file_path):
        box = BoxLayout(orientation="vertical", padding=dp(15), spacing=dp(10))
        ti_s = ModernInput(hint_text="Temat wiadomości", size_hint_y=None, height=dp(55))
        ti_b = ModernInput(hint_text="Treść...", multiline=True)
        box.add_widget(ti_s); box.add_widget(ti_b); btn = ModernButton(text="WYŚLIJ", bg_color=COLOR_SUCCESS)
        btn.bind(on_press=lambda x: [p.dismiss(), self.special_send_step_4_run(file_path, self.selected_emails, ti_s.text, ti_b.text)] if ti_s.text else self.msg("!", "Wpisz temat")); box.add_widget(btn); p = Popup(title="Wiadomość", content=box, size_hint=(0.95,0.85)); p.open()

    def special_send_step_4_run(self, file_path, target_list, subject, body):
        def thread():
            cfg_p = Path(self.user_data_dir) / "smtp.json"
            if not cfg_p.exists(): return
            cfg = json.load(open(cfg_p)); ok, err = 0, 0
            try:
                srv = smtplib.SMTP("smtp.gmail.com", 587, timeout=15); srv.starttls(); srv.login(cfg['u'], cfg['p'])
                for email in target_list:
                    try:
                        msg = EmailMessage(); msg["Subject"], msg["From"], msg["To"] = subject, cfg["u"], email; msg.set_content(body)
                        with open(file_path, "rb") as f:
                            ct, _ = mimetypes.guess_type(file_path); mn, sb = (ct or 'application/octet-stream').split('/', 1)
                            msg.add_attachment(f.read(), maintype=mn, subtype=sb, filename=os.path.basename(file_path))
                        srv.send_message(msg); ok += 1
                    except: err += 1
                srv.quit()
                Clock.schedule_once(lambda d: self.msg("Koniec", f"Wysłano: {ok}, Błędy: {err}"))
            except Exception as e: Clock.schedule_once(lambda d: self.msg("Błąd krytyczny", str(e)))
        threading.Thread(target=thread, daemon=True).start()

    # --- MASOWA WYSYŁKA ---

    def start_mass_mailing(self, _):
        if not self.full_data: return self.msg("!", "Brak arkusza!")
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
            return self.msg("Koniec", f"Sukces: {self.stats['ok']}, Błędy: {self.stats['fail']}")
        
        row = self.queue.pop(0)
        n, s = str(row[self.idx_name]).strip(), str(row[self.idx_surname]).strip()
        p = str(row[self.idx_pesel]).strip() if self.idx_pesel != -1 else ""
        res = self.conn.execute("SELECT email FROM contacts WHERE (name=? AND surname=?) OR pesel=?", (n.lower(), s.lower(), p)).fetchone()
        
        if res:
            if self.auto_send_mode: self.send_engine(row, res[0])
            else: self.ask_send(row, res[0], n, s)
        else:
            self.stats["skip"] += 1; self.process_mailing_queue()

    def ask_send(self, row, email, n, s):
        b = BoxLayout(orientation="vertical", padding=dp(20), spacing=dp(15))
        b.add_widget(Label(text=f"WYSŁAĆ DO:\\n{n} {s}\\n{email}", halign="center"))
        btns = BoxLayout(size_hint_y=None, height=dp(55), spacing=dp(10))
        btns.add_widget(ModernButton(text="TAK", on_press=lambda x: [p.dismiss(), self.send_engine(row, email)], bg_color=COLOR_SUCCESS))
        btns.add_widget(ModernButton(text="NIE", on_press=lambda x: [p.dismiss(), setattr(self.stats, 'skip', self.stats['skip']+1), self.process_mailing_queue()]))
        b.add_widget(btns); p = Popup(title="Weryfikacja", content=b, size_hint=(0.85, 0.35)); p.open()

    def send_engine(self, row_data, target):
        def run():
            cfg_p = Path(self.user_data_dir) / "smtp.json"
            if not cfg_p.exists(): return
            cfg = json.load(open(cfg_p))
            try:
                srv = smtplib.SMTP("smtp.gmail.com", 587, timeout=12); srv.starttls(); srv.login(cfg['u'], cfg['p'])
                nx = str(row_data[self.idx_name]).title()
                ts = self.conn.execute("SELECT val FROM settings WHERE key='t_sub'").fetchone()
                tb = self.conn.execute("SELECT val FROM settings WHERE key='t_body'").fetchone()
                msg = EmailMessage()
                msg["Subject"] = (ts[0] if ts else "Raport").replace("{Imię}", nx)
                msg["From"], msg["To"] = cfg['u'], target
                msg.set_content((tb[0] if tb else "Witaj").replace("{Imię}", nx))
                
                # Plik Excel indywidualny
                tmp = Path(self.user_data_dir) / "r.xlsx"; wb = Workbook(); ws = wb.active
                ws.append([self.full_data[0][k] for k in self.export_indices])
                ws.append([str(row_data[k]) for k in self.export_indices])
                self.style_xlsx(ws); wb.save(tmp)
                msg.add_attachment(open(tmp, "rb").read(), maintype="application", subtype="xlsx", filename=f"Raport_{nx}.xlsx")
                
                for path in self.global_attachments:
                    if os.path.exists(path):
                        with open(path, "rb") as af:
                            ct, _ = mimetypes.guess_type(path); mn, sb = (ct or 'application/octet-stream').split('/', 1)
                            msg.add_attachment(af.read(), maintype=mn, subtype=sb, filename=os.path.basename(path))
                
                srv.send_message(msg); srv.quit()
                self.stats["ok"] += 1; self.session_details.append(f"OK: {target}")
            except: self.stats["fail"] += 1
            Clock.schedule_once(lambda d: self.process_mailing_queue())
        threading.Thread(target=run, daemon=True).start()

    # --- POMOCNICZE ---

    def style_xlsx(self, ws):
        side = Side(style='thin')
        for row in ws.iter_rows():
            for cell in row:
                cell.border = Border(top=side, left=side, right=side, bottom=side)
                cell.alignment = Alignment(horizontal='center')

    def process_excel(self, p):
        try:
            wb = load_workbook(p, data_only=True); ws = wb.active
            raw = [["" if v is None else str(v).strip() for v in r] for r in ws.iter_rows(values_only=True)]
            h_idx = 0
            for i, r in enumerate(raw[:10]):
                if any(x in " ".join(r).lower() for x in ["imię", "imie", "nazwisko"]): h_idx = i; break
            self.full_data = raw[h_idx:]; self.filtered_data = self.full_data; self.export_indices = list(range(len(self.full_data[0])))
            for i, v in enumerate(self.full_data[0]):
                v = v.lower()
                if "imi" in v: self.idx_name = i
                if "naz" in v: self.idx_surname = i
                if "pesel" in v: self.idx_pesel = i
            self.msg("OK", "Arkusz wczytany")
        except: self.msg("Błąd", "Nieprawidłowy plik")

    def process_book(self, p):
        try:
            wb = load_workbook(p, data_only=True); ws = wb.active
            raw = list(ws.iter_rows(values_only=True)); h = [str(x).lower() for x in raw[0]]
            inm, isur, imail, ipes = 0, 1, 2, -1
            for i, v in enumerate(h):
                if "imi" in v: inm=i
                elif "naz" in v: isur=i
                elif "@" in v or "mail" in v: imail=i
                elif "pesel" in v: ipes=i
            c = 0
            for r in raw[1:]:
                if r[imail] and "@" in str(r[imail]):
                    self.conn.execute("INSERT OR REPLACE INTO contacts VALUES (?,?,?,?,?)", (str(r[inm]).strip().lower(), str(r[isur]).strip().lower(), str(r[imail]).strip(), str(r[ipes]) if ipes!=-1 else "", ""))
                    c += 1
            self.conn.commit(); self.update_stats(); self.msg("Import", f"Dodano {c} osób")
        except: self.msg("!", "Błąd importu")

    def refresh_contacts_list(self, *a):
        self.c_list.clear_widgets(); sv = self.ti_csearch.text.lower()
        rows = self.conn.execute("SELECT name, surname, email, pesel, phone FROM contacts ORDER BY surname ASC").fetchall()
        for n, s, e, p, ph in rows:
            if sv and sv not in f"{n} {s} {e}".lower(): continue
            cr = CardRow(size_hint_y=None, height=dp(90), padding=dp(10))
            cb = CheckBox(size_hint_x=None, width=dp(40), active=(e in self.selected_emails))
            cb.bind(active=lambda i,v,m=e: self.selected_emails.append(m) if v else self.selected_emails.remove(m))
            cr.add_widget(cb); tc = BoxLayout(orientation="vertical")
            tc.add_widget(Label(text=f"{n} {s}".upper(), bold=True, halign="left", text_size=(dp(200),None)))
            tc.add_widget(Label(text=e, font_size='11sp', color=COLOR_PRIMARY, halign="left", text_size=(dp(200),None)))
            cr.add_widget(tc)
            btns = BoxLayout(orientation="vertical", size_hint_x=None, width=dp(70), spacing=dp(2))
            btns.add_widget(Button(text="USUŃ", background_color=COLOR_DANGER, font_size='10sp', on_press=lambda x, n=n, s=s: self.delete_contact(n, s)))
            cr.add_widget(btns); self.c_list.add_widget(cr)

    def delete_contact(self, n, s):
        def go(_):
            self.conn.execute("DELETE FROM contacts WHERE name=? AND surname=?", (n, s))
            self.conn.commit(); px.dismiss(); self.refresh_contacts_list(); self.update_stats()
        b = BoxLayout(orientation="vertical", padding=dp(20), spacing=dp(10))
        b.add_widget(Label(text=f"Usunąć {n} {s}?"))
        btns = BoxLayout(size_hint_y=None, height=dp(50), spacing=dp(10))
        btns.add_widget(ModernButton(text="TAK", bg_color=COLOR_DANGER, on_press=go))
        btns.add_widget(ModernButton(text="NIE", on_press=lambda x: px.dismiss()))
        b.add_widget(btns); px = Popup(title="Usuń", content=b, size_hint=(0.8, 0.35)); px.open()

    def setup_table_ui(self):
        root = BoxLayout(orientation="vertical")
        menu = BoxLayout(size_hint_y=None, height=dp(60), spacing=dp(5), padding=dp(5))
        self.ti_tab_search = ModernInput(hint_text="Szukaj..."); self.ti_tab_search.bind(text=self.filter_table)
        menu.add_widget(self.ti_tab_search)
        menu.add_widget(ModernButton(text="POWRÓT", on_press=lambda x: setattr(self.sm, 'current', 'home'), size_hint_x=None, width=dp(100)))
        root.add_widget(menu)
        self.table_view = GridLayout(size_hint=(None, None), cols=1)
        self.table_view.bind(minimum_height=self.table_view.setter('height'), minimum_width=self.table_view.setter('width'))
        sc = ScrollView(do_scroll_x=True, do_scroll_y=True); sc.add_widget(self.table_view)
        root.add_widget(sc); self.sc["table"].add_widget(root)

    def refresh_table(self):
        self.table_view.clear_widgets(); self.table_view.cols = len(self.full_data[0])
        self.table_view.width = len(self.full_data[0]) * dp(150)
        for row in self.filtered_data[:150]:
            for cell in row:
                self.table_view.add_widget(ColorSafeLabel(text=str(cell), size=(dp(150), dp(45)), size_hint=(None,None), bg_color=COLOR_BG, text_color=COLOR_TEXT))

    def filter_table(self, i, v):
        v = v.lower()
        self.filtered_data = [self.full_data[0]] + [r for r in self.full_data[1:] if v in " ".join([str(x) for x in r]).lower()]
        self.refresh_table()

    def update_stats(self, *a):
        try:
            c = self.conn.execute("SELECT count(*) FROM contacts").fetchone()[0]
            self.lbl_stats.text = f"BAZA: {c} | ZAŁĄCZNIKI: {len(self.global_attachments)}"
        except: pass

    def msg(self, tit, txt):
        b = BoxLayout(orientation="vertical", padding=dp(20), spacing=dp(15))
        b.add_widget(Label(text=txt, halign="center")); b.add_widget(ModernButton(text="OK", on_press=lambda x: p.dismiss(), height=dp(50), size_hint_y=None))
        p = Popup(title=tit, content=b, size_hint=(0.85, 0.4)); p.open()

    # --- POZOSTAŁE EKRANY (USTAWIENIA I RAPORTY) ---

    def setup_smtp_ui(self):
        l = BoxLayout(orientation="vertical", padding=dp(30), spacing=dp(15))
        self.ti_su = ModernInput(hint_text="Email Gmail"); self.ti_sp = ModernInput(hint_text="Hasło aplikacji", password=True)
        p = Path(self.user_data_dir) / "smtp.json"
        if p.exists(): d = json.load(open(p)); self.ti_su.text, self.ti_sp.text = d.get('u',''), d.get('p','')
        l.add_widget(Label(text="GMAIL SMTP", bold=True, color=COLOR_PRIMARY))
        l.add_widget(self.ti_su); l.add_widget(self.ti_sp)
        l.add_widget(ModernButton(text="ZAPISZ", on_press=lambda x: [json.dump({'u':self.ti_su.text, 'p':self.ti_sp.text}, open(p, "w")), self.msg("OK", "Zapisano")]))
        l.add_widget(ModernButton(text="POWRÓT", bg_color=(0.3, 0.3, 0.3, 1), on_press=lambda x: setattr(self.sm, 'current', 'home')))
        self.sc["smtp"].add_widget(l)

    def setup_tmpl_ui(self):
        l = BoxLayout(orientation="vertical", padding=dp(30), spacing=dp(15))
        self.ti_ts = ModernInput(hint_text="Temat {Imię}"); self.ti_tb = ModernInput(hint_text="Treść {Imię}", multiline=True)
        ts = self.conn.execute("SELECT val FROM settings WHERE key='t_sub'").fetchone()
        tb = self.conn.execute("SELECT val FROM settings WHERE key='t_body'").fetchone()
        if ts: self.ti_ts.text = ts[0]; 
        if tb: self.ti_tb.text = tb[0]
        l.add_widget(Label(text="SZABLON EMAIL", bold=True, color=COLOR_PRIMARY))
        l.add_widget(self.ti_ts); l.add_widget(self.ti_tb)
        def sv(_):
            self.conn.execute("INSERT OR REPLACE INTO settings VALUES (?,?)", ('t_sub', self.ti_ts.text))
            self.conn.execute("INSERT OR REPLACE INTO settings VALUES (?,?)", ('t_body', self.ti_tb.text))
            self.conn.commit(); self.msg("OK", "Szablon zapisany")
        l.add_widget(ModernButton(text="ZAPISZ SZABLON", on_press=sv))
        l.add_widget(ModernButton(text="POWRÓT", on_press=lambda x: setattr(self.sm, 'current', 'email')))
        self.sc["tmpl"].add_widget(l)

    def setup_report_ui(self):
        l = BoxLayout(orientation="vertical", padding=dp(15), spacing=dp(10))
        l.add_widget(Label(text="HISTORIA OPERACJI", bold=True, color=COLOR_PRIMARY, size_hint_y=None, height=dp(40)))
        self.r_grid = GridLayout(cols=1, size_hint_y=None, spacing=dp(8)); self.r_grid.bind(minimum_height=self.r_grid.setter('height'))
        sc = ScrollView(); sc.add_widget(self.r_grid); l.add_widget(sc)
        l.add_widget(ModernButton(text="POWRÓT", on_press=lambda x: setattr(self.sm, 'current', 'email'), size_hint_y=None, height=dp(55)))
        self.sc["report"].add_widget(l)

    def refresh_reports(self):
        self.r_grid.clear_widgets()
        rows = self.conn.execute("SELECT date, ok, fail FROM reports ORDER BY id DESC LIMIT 25").fetchall()
        for d, o, f in rows:
            c = CardRow(size_hint_y=None, height=dp(60), padding=dp(10))
            c.add_widget(Label(text=d, bold=True, font_size='12sp')); c.add_widget(Label(text=f"OK: {o} | ERR: {f}"))
            self.r_grid.add_widget(c)

    def setup_contacts_ui(self):
        l = BoxLayout(orientation="vertical", padding=dp(15), spacing=dp(10))
        h = BoxLayout(size_hint_y=None, height=dp(55), spacing=dp(10))
        self.ti_csearch = ModernInput(hint_text="Szukaj..."); h.add_widget(self.ti_csearch)
        h.add_widget(ModernButton(text="+", size_hint_x=None, width=dp(55), on_press=lambda x: self.form_contact()))
        self.c_list = GridLayout(cols=1, size_hint_y=None, spacing=dp(5)); self.c_list.bind(minimum_height=self.c_list.setter('height'))
        sc = ScrollView(); sc.add_widget(self.c_list); l.add_widget(h); l.add_widget(sc)
        l.add_widget(ModernButton(text="POWRÓT", on_press=lambda x: setattr(self.sm, 'current', 'email'), height=dp(55), size_hint_y=None))
        self.sc["contacts"].add_widget(l)

    def form_contact(self, n="", s="", e="", p=""):
        b = BoxLayout(orientation="vertical", padding=dp(15), spacing=dp(10))
        ins = [ModernInput(text=n, hint_text="Imię"), ModernInput(text=s, hint_text="Nazwisko"), ModernInput(text=e, hint_text="Email")]
        for i in ins: b.add_widget(i)
        def save(_):
            self.conn.execute("INSERT OR REPLACE INTO contacts (name, surname, email) VALUES (?,?,?)", (ins[0].text.lower(), ins[1].text.lower(), ins[2].text))
            self.conn.commit(); pf.dismiss(); self.refresh_contacts_list(); self.update_stats()
        b.add_widget(ModernButton(text="Zapisz", on_press=save)); pf = Popup(title="Kontakt", content=b, size_hint=(0.9, 0.6)); pf.open()

if __name__ == "__main__":
    FutureApp().run()
