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
from kivy.utils import platform, get_color_from_hex
from kivy.core.window import Window
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.button import Button
from kivy.uix.label import Label
from kivy.uix.popup import Popup
from kivy.uix.scrollview import ScrollView
from kivy.uix.gridlayout import GridLayout
from kivy.uix.textinput import TextInput
from kivy.uix.checkbox import CheckBox
from kivy.uix.screenmanager import ScreenManager, Screen, FadeTransition
from kivy.uix.progressbar import ProgressBar
from kivy.graphics import Color, RoundedRectangle

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

# PALETA KOLORÓW MODERN DARK MATERIAL
CLR_PRIMARY = get_color_from_hex("#3B82F6")   # Aktywny niebieski
CLR_BG = get_color_from_hex("#0F172A")        # Tło główne
CLR_CARD = get_color_from_hex("#1E293B")       # Tło kart/elementów
CLR_TEXT = get_color_from_hex("#F8FAFC")       # Główny tekst
CLR_SUBTEXT = get_color_from_hex("#94A3B8")    # Tekst pomocniczy
CLR_DANGER = get_color_from_hex("#EF4444")     # Czerwony (Usuń/Błąd)
CLR_SUCCESS = get_color_from_hex("#10B981")    # Zielony (Wyślij/Ok)

class StyledButton(Button):
    def __init__(self, bg_color=CLR_PRIMARY, **kwargs):
        super().__init__(**kwargs)
        self.background_normal = ""
        self.background_down = ""
        self.background_color = (0,0,0,0)
        self.bold = True
        self.font_size = '14sp'
        self.real_bg = bg_color
        self.bind(pos=self._update, size=self._update)
    def _update(self, *args):
        self.canvas.before.clear()
        with self.canvas.before:
            Color(*(self.real_bg if self.state == 'normal' else [c*0.8 for c in self.real_bg[:3]] + [1]))
            RoundedRectangle(pos=self.pos, size=self.size, radius=[dp(10)])

class StyledInput(TextInput):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.background_normal = ""
        self.background_active = ""
        self.background_color = (0.1, 0.15, 0.25, 1)
        self.foreground_color = CLR_TEXT
        self.cursor_color = CLR_PRIMARY
        self.padding = [dp(12), dp(12)]
        self.font_size = '15sp'
        self.hint_text_color = CLR_SUBTEXT

class ContactCard(BoxLayout):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.orientation = 'horizontal'
        self.size_hint_y = None
        self.height = dp(120)
        self.padding = dp(12)
        self.spacing = dp(12)
        self.bind(pos=self._update, size=self._update)
    def _update(self, *args):
        self.canvas.before.clear()
        with self.canvas.before:
            Color(*CLR_CARD)
            RoundedRectangle(pos=(self.x + dp(5), self.y + dp(5)), size=(self.width - dp(10), self.height - dp(10)), radius=[dp(12)])

class HomeScreen(Screen): pass
class TableScreen(Screen): pass
class EmailScreen(Screen): pass
class SMTPScreen(Screen): pass
class TemplateScreen(Screen): pass
class ContactsScreen(Screen): pass
class ReportScreen(Screen): pass

class FutureApp(App):
    def build(self):
        self.title = "PASKI-FUTURE 2.0"
        Window.softinput_mode = "below_target"
        Window.clearcolor = CLR_BG
        self.full_data = []; self.filtered_data = []; self.export_indices = []
        self.global_attachments = []; self.selected_emails = []; self.queue = []
        self.session_details = []; self.stats = {"ok": 0, "fail": 0, "skip": 0, "auto": 0}
        self.idx_name, self.idx_surname, self.idx_pesel = 0, 1, -1
        self.auto_send_mode = False
        
        self.init_db()
        self.sm = ScreenManager(transition=FadeTransition())
        self.add_screens()
        return self.sm

    def init_db(self):
        db_p = Path(self.user_data_dir) / "future_ultimate_v2.db"
        self.conn = sqlite3.connect(str(db_p), check_same_thread=False)
        self.conn.execute("CREATE TABLE IF NOT EXISTS contacts (name TEXT, surname TEXT, email TEXT, pesel TEXT, phone TEXT, PRIMARY KEY(name, surname))")
        self.conn.execute("CREATE TABLE IF NOT EXISTS settings (key TEXT PRIMARY KEY, val TEXT)")
        self.conn.execute("CREATE TABLE IF NOT EXISTS reports (id INTEGER PRIMARY KEY AUTOINCREMENT, date TEXT, ok INTEGER, fail INTEGER, skip INTEGER, auto INTEGER, details TEXT)")
        try: self.conn.execute("ALTER TABLE reports ADD COLUMN details TEXT")
        except: pass
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
        l = BoxLayout(orientation="vertical", padding=dp(40), spacing=dp(15))
        l.add_widget(Label(text="PASKI-FUTURE 2.0", font_size='32sp', bold=True, color=CLR_PRIMARY, size_hint_y=None, height=dp(80)))
        l.add_widget(Label(text="System Zarządzania Paskami Płac", color=CLR_SUBTEXT, size_hint_y=None, height=dp(20)))
        btn_box = BoxLayout(orientation="vertical", spacing=dp(10), padding=[0, dp(20)])
        b = lambda t, c, clr=CLR_PRIMARY: btn_box.add_widget(StyledButton(text=t.upper(), on_press=c, real_bg=clr, size_hint_y=None, height=dp(55)))
        b("Wczytaj Arkusz Płac", lambda x: self.open_picker("data"))
        b("Podgląd i Eksport", lambda x: [self.refresh_table(), setattr(self.sm, 'current', 'table')] if self.full_data else self.msg("!", "Brak danych"))
        b("Centrum Mailingowe", lambda x: setattr(self.sm, 'current', 'email'))
        b("Historia Wysyłek", lambda x: [self.refresh_reports(), setattr(self.sm, 'current', 'report')], CLR_CARD)
        b("Ustawienia SMTP", lambda x: setattr(self.sm, 'current', 'smtp'), CLR_CARD)
        l.add_widget(btn_box); self.screens["home"].add_widget(l)
        
        self.setup_email_ui(); self.setup_contacts_ui(); self.setup_table_ui(); self.setup_smtp_ui(); self.setup_tmpl_ui(); self.setup_report_ui()

    def setup_email_ui(self):
        l = BoxLayout(orientation="vertical", padding=dp(20), spacing=dp(10))
        l.add_widget(Label(text="MAILING DASHBOARD", font_size='22sp', bold=True, height=dp(50), size_hint_y=None))
        ab = BoxLayout(size_hint_y=None, height=dp(45), spacing=dp(10), padding=[dp(10), 0])
        self.cb_auto = CheckBox(size_hint_x=None, width=dp(40)); self.cb_auto.bind(active=lambda i, v: setattr(self, 'auto_send_mode', v))
        ab.add_widget(self.cb_auto); ab.add_widget(Label(text="AUTOMATYCZNE POTWIERDZANIE ODBIORCÓW", color=CLR_SUBTEXT, halign='left', text_size=(dp(250), None)))
        l.add_widget(ab)
        self.lbl_stats = Label(text="Baza: -", color=CLR_PRIMARY, height=dp(30), size_hint_y=None, bold=True); l.add_widget(self.lbl_stats)
        self.pb_label = Label(text="System Gotowy", font_size='13sp', color=CLR_SUBTEXT); l.add_widget(self.pb_label)
        self.pb = ProgressBar(max=100, height=dp(12), size_hint_y=None); l.add_widget(self.pb)
        g = GridLayout(cols=1, spacing=dp(8), padding=[0, dp(10)])
        bb = lambda t, c, clr=CLR_PRIMARY: g.add_widget(StyledButton(text=t, on_press=c, real_bg=clr, size_hint_y=None, height=dp(52)))
        bb("Importuj Kontakty z Excel", lambda x: self.open_picker("book"), CLR_CARD)
        bb("Zarządzaj Bazą Kontaktów", lambda x: [self.refresh_contacts_list(), setattr(self.sm, 'current', 'contacts')], CLR_CARD)
        bb("Edytuj Treść Wiadomości", lambda x: setattr(self.sm, 'current', 'tmpl'), CLR_CARD)
        bb("WYŚLIJ HURTOWO (Safe Engine)", self.start_mass_mailing_safe, CLR_SUCCESS)
        bb("POWRÓT", lambda x: setattr(self.sm, 'current', 'home'), get_color_from_hex("#475569"))
        l.add_widget(g); self.screens["email"].add_widget(l)

    # --- SAFE MAILING ENGINE ---
    def start_mass_mailing_safe(self, _):
        if not self.full_data: self.msg("!", "Pierwszy krok: Wczytaj arkusz płac!"); return
        if not (Path(self.user_data_dir) / "smtp.json").exists(): self.msg("!", "Brak konfiguracji Gmail!"); return
        self.stats = {"ok": 0, "fail": 0, "skip": 0, "auto": 0}; self.session_details = []
        self.queue = list(self.full_data[1:]); self.total_q = len(self.queue); self.pb.value = 0
        threading.Thread(target=self.safe_mailing_worker, daemon=True).start()

    def safe_mailing_worker(self):
        cfg = json.load(open(Path(self.user_data_dir) / "smtp.json"))
        ts = self.conn.execute("SELECT val FROM settings WHERE key='t_sub'").fetchone()
        tb = self.conn.execute("SELECT val FROM settings WHERE key='t_body'").fetchone()
        t_sub = ts[0] if ts else "Raport"; t_body = tb[0] if tb else "Dzień dobry"
        server = None
        for idx, row in enumerate(self.full_data[1:]):
            n, s = str(row[self.idx_name]).strip(), str(row[self.idx_surname]).strip()
            p = str(row[self.idx_pesel]).strip() if self.idx_pesel != -1 else ""
            Clock.schedule_once(lambda dt, i=idx+1: self.update_progress_ui(i))
            res = self.conn.execute("SELECT email FROM contacts WHERE name=? AND surname=? COLLATE NOCASE", (n.lower(), s.lower())).fetchone()
            if not res and p and len(p) > 5: res = self.conn.execute("SELECT email FROM contacts WHERE pesel=?", (p,)).fetchone()
            if not res:
                self.stats["skip"] += 1; self.session_details.append(f"POMINIĘTO: {n} {s} (Brak maila w bazie)")
                continue
            try:
                if server is None:
                    server = smtplib.SMTP("smtp.gmail.com", 587, timeout=20); server.starttls(); server.login(cfg['u'], cfg['p'])
                nx, sx = n.title(), s.title(); msg = EmailMessage(); dat = datetime.now().strftime("%d.%m.%Y")
                msg["Subject"] = t_sub.replace("{Imię}", nx); msg["From"] = cfg['u']; msg["To"] = res[0]
                msg.set_content(t_body.replace("{Imię}", nx).replace("{Data}", dat))
                tmp_p = Path(self.user_data_dir) / f"temp_paski.xlsx"; wb = Workbook(); ws = wb.active
                ws.append([self.full_data[0][k] for k in self.export_indices])
                ws.append([str(row[k]) if (k < len(row) and str(row[k]).strip() != "") else "0" for k in self.export_indices])
                self.style_xlsx(ws); wb.save(tmp_p)
                with open(tmp_p, "rb") as f: msg.add_attachment(f.read(), maintype="application", subtype="xlsx", filename=f"Raport_{nx}_{sx}.xlsx")
                for ga in self.global_attachments:
                    if os.path.exists(ga):
                        ct, _ = mimetypes.guess_type(ga); mn, sb = (ct or 'app/oct').split('/', 1)
                        msg.add_attachment(open(ga, "rb").read(), maintype=mn, subtype=sb, filename=os.path.basename(ga))
                server.send_message(msg); self.stats["ok"] += 1; self.session_details.append(f"WYŁANO: {nx} {sx} ({res[0]})")
                time.sleep(random.uniform(1.3, 2.8)) # Ochrona przed banem
            except Exception as e:
                self.stats["fail"] += 1; self.session_details.append(f"BŁĄD: {n} {s} ({str(e)})"); server = None
        if server: server.quit()
        det = "\n".join(self.session_details); self.conn.execute("INSERT INTO reports (date, ok, fail, skip, auto, details) VALUES (?,?,?,?,?,?)", (datetime.now().strftime("%Y-%m-%d %H:%M"), self.stats['ok'], self.stats['fail'], self.stats['skip'], 0, det)); self.conn.commit()
        Clock.schedule_once(lambda dt: self.msg("Koniec Serii", f"Sukcesy: {self.stats['ok']}\nBłędy: {self.stats['fail']}"))

    def update_progress_ui(self, val):
        self.pb.value = int((val/self.total_q)*100); self.pb_label.text = f"Pakiety danych: {val}/{self.total_q}"

    def form_contact(self, n="", s="", e="", pes="", ph=""):
        root = BoxLayout(orientation="vertical", padding=dp(15), spacing=dp(10))
        sc = ScrollView(size_hint_y=0.85); b = BoxLayout(orientation="vertical", size_hint_y=None, spacing=dp(10)); b.bind(minimum_height=b.setter('height'))
        self.f_inputs = []
        for lab, val in [("Imię", n), ("Nazwisko", s), ("Email", e), ("PESEL", pes), ("Telefon", ph)]:
            box = BoxLayout(orientation="vertical", size_hint_y=None, height=dp(65), spacing=dp(4))
            box.add_widget(Label(text=lab.upper(), font_size='11sp', color=CLR_PRIMARY, bold=True, halign='left', text_size=(dp(280), None)))
            ti = StyledInput(text=val, multiline=False); box.add_widget(ti); self.f_inputs.append(ti); b.add_widget(box)
        sc.add_widget(b); root.add_widget(sc)
        btns = BoxLayout(size_hint_y=None, height=dp(50), spacing=dp(10))
        btns.add_widget(StyledButton(text="Anuluj", on_press=lambda x: px.dismiss(), real_bg=CLR_CARD))
        def sv(_):
            d = [f.text.strip().lower() if i<2 else f.text.strip() for i,f in enumerate(self.f_inputs)]
            self.conn.execute("INSERT OR REPLACE INTO contacts VALUES (?,?,?,?,?)", d); self.conn.commit(); px.dismiss(); self.refresh_contacts_list(); self.update_stats()
        btns.add_widget(StyledButton(text="Zapisz", on_press=sv, real_bg=CLR_SUCCESS))
        root.add_widget(btns); px = Popup(title="Karta Kontaktu", content=root, size_hint=(0.95, 0.9)); px.open()

    def refresh_contacts_list(self, *a):
        self.c_list.clear_widgets(); rows = self.conn.execute("SELECT name, surname, email, pesel, phone FROM contacts ORDER BY surname ASC").fetchall()
        for n, s, e, p, ph in rows:
            if self.ti_csearch.text.lower() and self.ti_csearch.text.lower() not in f"{n} {s} {e} {p} {ph}".lower(): continue
            card = ContactCard(); info = BoxLayout(orientation="vertical", padding=[dp(5), dp(5)])
            info.add_widget(Label(text=f"{n.upper()} {s.upper()}", bold=True, halign="left", text_size=(dp(220),None)))
            info.add_widget(Label(text=f"{e}", font_size='12sp', color=CLR_PRIMARY, halign="left", text_size=(dp(220),None)))
            info.add_widget(Label(text=f"PESEL: {p} | Tel: {ph if ph else '-'}", font_size='11sp', color=CLR_SUBTEXT, halign="left", text_size=(dp(220),None)))
            card.add_widget(info); acts = BoxLayout(orientation="vertical", size_hint_x=None, width=dp(75), spacing=dp(5))
            acts.add_widget(StyledButton(text="Edytuj", on_press=lambda x, d=(n,s,e,p,ph): self.form_contact(*d), real_bg=CLR_CARD))
            acts.add_widget(StyledButton(text="Usuń", on_press=lambda x, na=n, su=s: self.delete_contact(na,su), real_bg=CLR_DANGER))
            card.add_widget(acts); self.c_list.add_widget(card)

    def style_xlsx(self, ws):
        s = Side(style='thin'); c = Alignment(horizontal='center', vertical='center')
        for ri, row in enumerate(ws.iter_rows(), 1):
            for cell in row:
                cell.border = Border(top=s, left=s, right=s, bottom=s); cell.alignment = c
                if ri == 1: cell.font = Font(bold=True); cell.fill = PatternFill("solid", start_color="DDEBF7")
                elif ri % 2 == 0: cell.fill = PatternFill("solid", start_color="F9FAFB")
        for col in ws.columns:
            m_l = max([len(str(cell.value or "")) for cell in col])
            ws.column_dimensions[col[0].column_letter].width = (m_l * 1.3) + 6

    def open_picker(self, mode):
        if platform != "android": self.msg("!", "Dostępne tylko na urządzeniu mobilnym"); return
        from jnius import autoclass; from android import activity
        PA = autoclass("org.kivy.android.PythonActivity"); Intent = autoclass("android.content.Intent")
        intent = Intent(Intent.ACTION_GET_CONTENT); intent.setType("*/*")
        def cb(req, res, dt):
            if req == 1001 and res == -1 and dt:
                uri = dt.getData(); reslv = PA.mActivity.getContentResolver()
                loc = Path(self.user_data_dir) / f"imp_{int(time.time())}.xlsx"
                with open(loc, "wb") as f:
                    stream = reslv.openInputStream(uri); b = bytearray(16384)
                    while True:
                        n = stream.read(b)
                        if n <= 0: break
                        f.write(b[:n])
                    stream.close()
                if mode == "data": self.process_excel(loc)
                else: self.process_book(loc)
        activity.bind(on_activity_result=cb); PA.mActivity.startActivityForResult(intent, 1001)

    def process_excel(self, p):
        try:
            wb = load_workbook(p, data_only=True); ws = wb.active
            self.full_data = [["" if v is None else str(v).strip() for v in r] for r in ws.iter_rows(values_only=True)]
            self.export_indices = list(range(len(self.full_data[0]))); h = [str(x).lower() for x in self.full_data[0]]
            for i,v in enumerate(h):
                if "imi" in v: self.idx_name = i
                elif "naz" in v: self.idx_surname = i
                elif "pes" in v: self.idx_pesel = i
            self.msg("OK", f"Wczytano {len(self.full_data)-1} pracowników.")
        except Exception as e: self.msg("Błąd", str(e))

    def setup_contacts_ui(self):
        l = BoxLayout(orientation="vertical", padding=dp(12), spacing=dp(10))
        t = BoxLayout(size_hint_y=None, height=dp(50), spacing=dp(6))
        self.ti_csearch = StyledInput(hint_text="Szukaj..."); self.ti_csearch.bind(text=self.refresh_contacts_list)
        t.add_widget(self.ti_csearch); t.add_widget(StyledButton(text="+", width=dp(50), size_hint_x=None, on_press=lambda x: self.form_contact()))
        t.add_widget(StyledButton(text="WRÓĆ", width=dp(70), size_hint_x=None, on_press=lambda x: setattr(self.sm, 'current', 'email'), real_bg=CLR_CARD))
        self.c_list = GridLayout(cols=1, size_hint_y=None, spacing=dp(5)); self.c_list.bind(minimum_height=self.c_list.setter('height'))
        sc = ScrollView(); sc.add_widget(self.c_list); l.add_widget(t); l.add_widget(sc); self.screens["contacts"].add_widget(l)

    def msg(self, tit, txt):
        b = BoxLayout(orientation="vertical", padding=dp(20), spacing=dp(15))
        b.add_widget(Label(text=txt, halign="center", font_size='15sp'))
        b.add_widget(StyledButton(text="ROZUMIEM", height=dp(50), size_hint_y=None, on_press=lambda x: p.dismiss()))
        p = Popup(title=tit, content=b, size_hint=(0.85, 0.4)); p.open()

    def delete_contact(self, n, s):
        def pr(_): self.conn.execute("DELETE FROM contacts WHERE name=? AND surname=?", (n, s)); self.conn.commit(); dx.dismiss(); self.refresh_contacts_list()
        btn = StyledButton(text="USUŃ NA ZAWSZE", real_bg=CLR_DANGER, on_press=pr); dx = Popup(title="Usunięcie", content=btn, size_hint=(0.7, 0.25)); dx.open()

    def update_stats(self, *a):
        try: c = self.conn.execute("SELECT count(*) FROM contacts").fetchone()[0]; self.lbl_stats.text = f"BAZA KONTAKTÓW: {c}"
        except: pass

    def setup_table_ui(self):
        root = BoxLayout(orientation="vertical")
        menu = BoxLayout(size_hint_y=None, height=dp(55), padding=dp(6), spacing=dp(6))
        ti = StyledInput(hint_text="Filtruj tabelę..."); ti.bind(text=self.filter_table); menu.add_widget(ti)
        menu.add_widget(StyledButton(text="WRÓĆ", size_hint_x=0.25, on_press=lambda x: setattr(self.sm, 'current', 'home'), real_bg=CLR_CARD))
        hs = ScrollView(size_hint_y=None, height=dp(50), do_scroll_y=False); self.table_header_layout = GridLayout(rows=1, size_hint=(None,None), height=dp(50))
        ds = ScrollView(); self.table_content_layout = GridLayout(size_hint=(None,None)); self.table_content_layout.bind(minimum_height=self.table_content_layout.setter('height'), minimum_width=self.table_content_layout.setter('width'))
        ds.bind(scroll_x=lambda i,v: setattr(hs, 'scroll_x', v)); hs.add_widget(self.table_header_layout); ds.add_widget(self.table_content_layout)
        root.add_widget(menu); root.add_widget(hs); root.add_widget(ds); self.screens["table"].add_widget(root)

    def setup_smtp_ui(self):
        l = BoxLayout(orientation="vertical", padding=dp(30), spacing=dp(15))
        l.add_widget(Label(text="KONFIGURACJA GMAIL", font_size='20sp', bold=True))
        l.add_widget(Label(text="Pamiętaj o haśle aplikacji Google", color=CLR_SUBTEXT, font_size='12sp'))
        self.ti_su = StyledInput(hint_text="Adres Gmail"); self.ti_sp = StyledInput(hint_text="Hasło aplikacji", password=True)
        p = Path(self.user_data_dir) / "smtp.json"; d = json.load(open(p)) if p.exists() else {}
        self.ti_su.text, self.ti_sp.text = d.get('u',''), d.get('p','')
        l.add_widget(self.ti_su); l.add_widget(self.ti_sp)
        l.add_widget(StyledButton(text="ZAPISZ USTAWIENIA", on_press=lambda x: [json.dump({'u':self.ti_su.text, 'p':self.ti_sp.text}, open(p, 'w')), self.msg("OK", "Konfiguracja zapisana")], real_bg=CLR_SUCCESS))
        l.add_widget(StyledButton(text="POWRÓT DO MENU", on_press=lambda x: setattr(self.sm, 'current', 'home'), real_bg=CLR_CARD))
        self.screens["smtp"].add_widget(l)

    def setup_tmpl_ui(self):
        l = BoxLayout(orientation="vertical", padding=dp(20), spacing=dp(10))
        l.add_widget(Label(text="SZABLON POWIADOMIENIA", bold=True, size_hint_y=None, height=dp(40)))
        self.ti_ts = StyledInput(hint_text="Temat (użyj {Imię})"); self.ti_tb = StyledInput(multiline=True, hint_text="Treść (użyj {Imię} i {Data})")
        ts = self.conn.execute("SELECT val FROM settings WHERE key='t_sub'").fetchone(); tb = self.conn.execute("SELECT val FROM settings WHERE key='t_body'").fetchone()
        self.ti_ts.text, self.ti_tb.text = (ts[0] if ts else ""), (tb[0] if tb else "")
        l.add_widget(self.ti_ts); l.add_widget(self.ti_tb)
        l.add_widget(StyledButton(text="ZAPISZ SZABLON", on_press=lambda x: [self.conn.execute("INSERT OR REPLACE INTO settings VALUES (?,?)", ('t_sub', self.ti_ts.text)), self.conn.execute("INSERT OR REPLACE INTO settings VALUES (?,?)", ('t_body', self.ti_tb.text)), self.conn.commit(), self.msg("Zapisano", "Zaktualizowano treść")], real_bg=CLR_SUCCESS))
        l.add_widget(StyledButton(text="ANULUJ", on_press=lambda x: setattr(self.sm, 'current', 'email'), real_bg=CLR_CARD))
        self.screens["tmpl"].add_widget(l)

    def setup_report_ui(self):
        l = BoxLayout(orientation="vertical", padding=dp(15)); self.report_grid = GridLayout(cols=1, size_hint_y=None, spacing=dp(10))
        self.report_grid.bind(minimum_height=self.report_grid.setter('height')); sc = ScrollView(); sc.add_widget(self.report_grid)
        l.add_widget(sc); l.add_widget(StyledButton(text="WRÓĆ DO MENU", height=dp(55), size_hint_y=None, on_press=lambda x: setattr(self.sm, 'current', 'home'), real_bg=CLR_CARD))
        self.screens["report"].add_widget(l)

    def refresh_reports(self, *a):
        self.report_grid.clear_widgets(); rows = self.conn.execute("SELECT date, ok, fail, details FROM reports ORDER BY id DESC").fetchall()
        for d, ok, fl, dt in rows:
            c = ContactCard(height=dp(100)); info = BoxLayout(orientation="vertical", padding=[dp(10), dp(5)])
            info.add_widget(Label(text=f"SESJA: {d}", bold=True)); info.add_widget(Label(text=f"Wysłano: {ok} | Błędów: {fl}", color=CLR_SUBTEXT, font_size='13sp'))
            c.add_widget(info); btn = StyledButton(text="LOGI", width=dp(70), size_hint_x=None, on_press=lambda x, t=dt: self.show_report_details(t), real_bg=CLR_PRIMARY)
            c.add_widget(btn); self.report_grid.add_widget(c)

    def show_report_details(self, t):
        box = BoxLayout(orientation='vertical', padding=dp(10), spacing=dp(10))
        ti = StyledInput(text=str(t), readonly=True, size_hint_y=0.8); b = StyledButton(text="ZAMKNIJ SZCZEGÓŁY", size_hint_y=0.2, on_press=lambda x: p.dismiss(), real_bg=CLR_CARD)
        box.add_widget(ti); box.add_widget(b); p = Popup(title="Szczegóły Sesji", content=box, size_hint=(0.95, 0.85)); p.open()

    def filter_table(self, i, v):
        if not self.full_data: return
        self.filtered_data = [self.full_data[0]] + [r for r in self.full_data[1:] if any(v.lower() in str(c).lower() for c in r)]; self.refresh_table()

    def refresh_table(self):
        self.table_content_layout.clear_widgets(); self.table_header_layout.clear_widgets()
        if not self.filtered_data: return
        w, h = dp(180), dp(50); headers = [self.full_data[0][i] for i in self.export_indices]
        self.table_header_layout.cols = len(headers); self.table_header_layout.width = len(headers)*w
        for head in headers: self.table_header_layout.add_widget(Label(text=str(head), color=CLR_PRIMARY, bold=True))
        self.table_content_layout.cols = len(headers); self.table_content_layout.width = len(headers)*w
        for row in self.filtered_data[1:]:
            for idx in self.export_indices: 
                v = str(row[idx]) if (idx < len(row) and str(row[idx]).strip() != "") else "0"
                self.table_content_layout.add_widget(Label(text=v, size=(w,h), size_hint=(None,None)))

if __name__ == "__main__": FutureApp().run()
