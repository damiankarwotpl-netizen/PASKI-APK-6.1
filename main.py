import os
import json
import sqlite3
import threading
import smtplib
import mimetypes
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

# Paleta kolorów
COLOR_PRIMARY = (0.1, 0.5, 0.9, 1)
COLOR_BG = (0.05, 0.07, 0.1, 1)
COLOR_CARD = (0.12, 0.15, 0.2, 1)
COLOR_TEXT = (0.95, 0.95, 0.95, 1)
COLOR_ROW_A = (0.08, 0.1, 0.15, 1)
COLOR_ROW_B = (0.13, 0.16, 0.22, 1)
COLOR_HEADER = (0.1, 0.2, 0.35, 1)

# --- KOMPONENTY ---

class ModernButton(Button):
    def __init__(self, bg_color=COLOR_PRIMARY, **kwargs):
        super().__init__(**kwargs)
        self.background_normal = ""
        self.background_color = (0,0,0,0)
        self.color = COLOR_TEXT
        self.bold = True
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
        self.padding = [dp(12), dp(12)]

class ColorSafeLabel(Label):
    def __init__(self, bg_color=(1,1,1,1), text_color=(1,1,1,1), **kwargs):
        super().__init__(**kwargs)
        self.color = text_color
        with self.canvas.before:
            Color(*bg_color)
            self.rect = Rectangle(size=self.size, pos=self.pos)
        self.bind(size=self._update, pos=self._update)
    def _update(self, inst, val):
        self.rect.size, self.rect.pos = self.size, self.pos
        self.text_size = (self.width - dp(10), None)

# --- APLIKACJA ---

class FutureApp(App):
    def build(self):
        Window.clearcolor = COLOR_BG
        self.full_data, self.filtered_data, self.export_indices = [], [], []
        self.global_attachments, self.selected_emails, self.queue = [], [], []
        self.session_details = []
        self.stats = {"ok": 0, "fail": 0, "skip": 0}
        self.idx_name, self.idx_surname, self.idx_pesel = 0, 1, -1
        self.auto_send_mode = False
        self.is_mailing_running = False
        self.init_db()
        self.sm = ScreenManager(transition=SlideTransition())
        self.add_screens()
        return self.sm

    def init_db(self):
        db_p = Path(self.user_data_dir) / "future_ultimate_v10.db"
        self.conn = sqlite3.connect(str(db_p), check_same_thread=False)
        self.conn.execute("CREATE TABLE IF NOT EXISTS contacts (name TEXT, surname TEXT, email TEXT, pesel TEXT, phone TEXT, PRIMARY KEY(name, surname))")
        self.conn.execute("CREATE TABLE IF NOT EXISTS settings (key TEXT PRIMARY KEY, val TEXT)")
        self.conn.execute("CREATE TABLE IF NOT EXISTS reports (id INTEGER PRIMARY KEY AUTOINCREMENT, date TEXT, ok INTEGER, fail INTEGER, skip INTEGER, auto INTEGER, details TEXT)")
        self.conn.commit()

    def add_screens(self):
        self.sc_ref = {
            "home": Screen(name="home"), "table": Screen(name="table"),
            "email": Screen(name="email"), "smtp": Screen(name="smtp"),
            "tmpl": Screen(name="tmpl"), "contacts": Screen(name="contacts"),
            "report": Screen(name="report")
        }
        self.setup_all_ui()
        for s in self.sc_ref.values(): self.sm.add_widget(s)

    def setup_all_ui(self):
        l = BoxLayout(orientation="vertical", padding=dp(30), spacing=dp(15))
        l.add_widget(Label(text="FUTURE ULTIMATE v10", font_size='34sp', bold=True, color=COLOR_PRIMARY))
        btn = lambda t, c: l.add_widget(ModernButton(text=t, on_press=c, height=dp(55), size_hint_y=None))
        btn("WCZYTAJ ARKUSZ PŁAC", lambda x: self.open_picker("data"))
        btn("PODGLĄD I EKSPORT", lambda x: [self.refresh_table(), setattr(self.sm, 'current', 'table')] if self.full_data else self.msg("!", "Danych brak"))
        btn("CENTRUM MAILINGOWE", lambda x: setattr(self.sm, 'current', 'email'))
        btn("RAPORTY SESJI", lambda x: [self.refresh_reports(), setattr(self.sm, 'current', 'report')])
        btn("USTAWIENIA SMTP", lambda x: setattr(self.sm, 'current', 'smtp'))
        self.sc_ref["home"].add_widget(l)
        self.setup_table_ui(); self.setup_email_ui(); self.setup_smtp_ui(); self.setup_tmpl_ui(); self.setup_contacts_ui(); self.setup_report_ui()

    def setup_email_ui(self):
        l = BoxLayout(orientation="vertical", padding=dp(25), spacing=dp(10))
        ab = BoxLayout(size_hint_y=None, height=dp(45), spacing=dp(10))
        self.cb_auto = CheckBox(size_hint_x=None, width=dp(45)); self.cb_auto.bind(active=lambda i, v: setattr(self, 'auto_send_mode', v))
        ab.add_widget(self.cb_auto); ab.add_widget(Label(text="AUTOMATYCZNA WYSYŁKA", bold=True)); l.add_widget(ab)
        self.lbl_stats = Label(text="Baza: 0", height=dp(30)); l.add_widget(self.lbl_stats)
        
        l.add_widget(ModernButton(text="WYCZYŚĆ ZAŁĄCZNIKI", on_press=self.clear_all_attachments, height=dp(45), size_hint_y=None, bg_color=(0.7, 0.1, 0.1, 1)))
        self.pb_label = Label(text="Gotowy", height=dp(25)); self.pb = ProgressBar(max=100, height=dp(20)); l.add_widget(self.pb_label); l.add_widget(self.pb)
        
        btns = [("IMPORT KONTAKTÓW", lambda x: self.open_picker("book")), 
                ("ZARZĄDZAJ BAZĄ", lambda x: [self.refresh_contacts_list(), setattr(self.sm, 'current', 'contacts')]), 
                ("EDYTUJ SZABLON", lambda x: setattr(self.sm, 'current', 'tmpl')), 
                ("DODAJ ZAŁĄCZNIK", lambda x: self.open_picker("attachment")), 
                ("WYŚLIJ JEDEN PLIK", self.start_special_send_flow), 
                ("START MASOWA WYSYŁKA", self.start_mass_mailing)]
        
        for t, c in btns: l.add_widget(ModernButton(text=t, on_press=c, height=dp(50), size_hint_y=None))
        l.add_widget(ModernButton(text="POWRÓT", on_press=lambda x: setattr(self.sm, 'current', 'home'), bg_color=(0.3,0.3,0.3,1)))
        self.sc_ref["email"].add_widget(l); self.update_stats()

    def clear_all_attachments(self, _):
        self.global_attachments = []
        self.update_stats()

    def open_picker(self, mode):
        if platform != "android": return self.msg("!", "Tylko na Android")
        from jnius import autoclass; from android import activity
        PA, Intent = autoclass("org.kivy.android.PythonActivity"), autoclass("android.content.Intent")
        intent = Intent(Intent.ACTION_GET_CONTENT); intent.setType("*/*")
        if mode == "attachment": intent.putExtra(Intent.EXTRA_ALLOW_MULTIPLE, True)

        def cb(req, res, dt):
            if req != 1001: return
            activity.unbind(on_activity_result=cb)
            if res == -1 and dt:
                resolver = PA.mActivity.getContentResolver()
                files = []
                clip = dt.getClipData()
                if clip:
                    for i in range(clip.getItemCount()): files.append(clip.getItemAt(i).getUri())
                else: files.append(dt.getData())

                for uri in files:
                    cur = resolver.query(uri, None, None, None, None); name = f"pick_{random.randint(100,999)}.xlsx"
                    if cur and cur.moveToFirst():
                        idx = cur.getColumnIndex("_display_name"); name = cur.getString(idx) if idx != -1 else name; cur.close()
                    try:
                        stream, loc = resolver.openInputStream(uri), Path(self.user_data_dir) / name
                        with open(loc, "wb") as f:
                            buf = bytearray(16384)
                            while (n := stream.read(buf)) > 0: f.write(buf[:n])
                        stream.close()
                        if mode == "data": self.process_excel(loc)
                        elif mode == "book": self.process_book(loc)
                        elif mode == "attachment": self.global_attachments.append(str(loc))
                    except: pass
                self.update_stats()
        activity.bind(on_activity_result=cb); PA.mActivity.startActivityForResult(intent, 1001)

    # --- SILNIK CORE (1:1 - PESEL, BATCHING, RANDOM DELAY) ---

    def mailing_worker(self):
        cfg_p = Path(self.user_data_dir) / "smtp.json"
        if not cfg_p.exists(): return self.finish_mailing("Brak SMTP")
        cfg = json.load(open(cfg_p)); b_on = cfg.get('batch', True); b_sz = 30; proc = 0
        try:
            srv = self.connect_smtp(cfg)
            while self.queue:
                row = self.queue.pop(0); n, s = str(row[self.idx_name]).strip(), str(row[self.idx_surname]).strip()
                p_exc = str(row[self.idx_pesel]).strip() if self.idx_pesel != -1 else ""
                res_p = self.conn.execute("SELECT email FROM contacts WHERE pesel=? AND pesel != ''", (p_exc,)).fetchone() if p_exc else None
                target, verify = None, False
                if res_p: target = res_p[0]
                else:
                    res_n = self.conn.execute("SELECT email FROM contacts WHERE name=? AND surname=? COLLATE NOCASE", (n.lower(), s.lower())).fetchone()
                    if res_n: target = res_n[0]; verify = not self.auto_send_mode
                if target:
                    if verify:
                        self.wait_for_user = True; Clock.schedule_once(lambda dt: self.ask_before_send_worker(row, target, n, s))
                        while self.wait_for_user: time.sleep(0.5)
                        if self.user_decision == "skip": continue
                    if self.send_single_email(srv, cfg, row, target): self.stats["ok"] += 1; self.session_details.append(f"OK: {n} {s}")
                    else: self.stats["fail"] += 1; srv.quit(); srv = self.connect_smtp(cfg)
                    proc += 1
                    if self.queue:
                        if b_on and proc >= b_sz: srv.quit(); time.sleep(60); srv = self.connect_smtp(cfg); proc = 0
                        else: time.sleep(random.uniform(3, 7))
                else: self.stats["skip"] += 1; self.session_details.append(f"SKIP: {n} {s}")
                Clock.schedule_once(lambda dt, d=(self.total_q - len(self.queue)): self.update_progress(d))
            srv.quit(); self.finish_mailing("Koniec sesji")
        except Exception as e: self.finish_mailing(f"Błąd: {e}")

    def connect_smtp(self, cfg):
        srv = smtplib.SMTP(cfg.get('h', 'smtp.gmail.com'), int(cfg.get('port', 587)), timeout=25)
        srv.starttls(); srv.login(cfg['u'], cfg['p']); return srv

    def send_single_email(self, srv, cfg, row_data, target):
        try:
            nx, sx = str(row_data[self.idx_name]).title(), str(row_data[self.idx_surname]).title()
            msg = EmailMessage(); ts = self.conn.execute("SELECT val FROM settings WHERE key='t_sub'").fetchone(); tb = self.conn.execute("SELECT val FROM settings WHERE key='t_body'").fetchone()
            msg["Subject"] = (ts[0] if ts else "Raport").replace("{Imię}", nx); msg["From"], msg["To"] = cfg['u'], target
            msg.set_content((tb[0] if tb else "Witaj").replace("{Imię}", nx).replace("{Data}", datetime.now().strftime("%d.%m.%Y")))
            tmp = Path(self.user_data_dir) / f"r_{nx}.xlsx"; wb = Workbook(); ws = wb.active
            ws.append([self.full_data[0][k] for k in self.export_indices]); ws.append([str(row_data[k]) if (str(row_data[k]).strip()!="") else "0" for k in self.export_indices])
            self.style_xlsx(ws); wb.save(tmp)
            with open(tmp, "rb") as f: msg.add_attachment(f.read(), maintype="application", subtype="xlsx", filename=f"Raport_{nx}_{sx}.xlsx")
            for p in self.global_attachments:
                if os.path.exists(p):
                    ct, _ = mimetypes.guess_type(p); mn, sb = (ct or 'application/octet-stream').split('/', 1)
                    with open(p, "rb") as f: msg.add_attachment(f.read(), maintype=mn, subtype=sb, filename=os.path.basename(p))
            srv.send_message(msg); return True
        except: return False

    def style_xlsx(self, ws):
        s, c = Side(style='thin'), Alignment(horizontal='center', vertical='center')
        for ri, row in enumerate(ws.iter_rows(), 1):
            for cell in row:
                cell.border = Border(top=s, left=s, right=s, bottom=s); cell.alignment = c
                if ri == 1: cell.font = Font(bold=True); cell.fill = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")
                elif ri % 2 == 0: cell.fill = PatternFill(start_color="F7F7F7", end_color="F7F7F7", fill_type="solid")
        for col in ws.columns:
            m = 0; col_let = col[0].column_letter
            for cell in col:
                if cell.value: m = max(m, len(str(cell.value)))
            ws.column_dimensions[col_let].width = (m * 1.3) + 7

    def setup_smtp_ui(self):
        l = BoxLayout(orientation="vertical", padding=dp(25), spacing=dp(8)); p = Path(self.user_data_dir) / "smtp.json"; d = json.load(open(p)) if p.exists() else {}
        self.ti_shost = ModernInput(hint_text="Serwer", text=d.get('h','smtp.gmail.com')); self.ti_sport = ModernInput(hint_text="Port", text=str(d.get('port',587))); self.ti_suser = ModernInput(hint_text="Email", text=d.get('u','')); self.ti_spass = ModernInput(hint_text="Hasło", password=True, text=d.get('p',''))
        l.add_widget(Label(text="USTAWIENIA SMTP", bold=True)); l.add_widget(self.ti_shost); l.add_widget(self.ti_sport); l.add_widget(self.ti_suser); l.add_widget(self.ti_spass)
        bx = BoxLayout(size_hint_y=None, height=dp(45)); self.cb_batching = CheckBox(size_hint_x=None, width=dp(45), active=d.get('batch', True)); bx.add_widget(self.cb_batching); bx.add_widget(Label(text="Batching (przerwa 60s/30 maili)")); l.add_widget(bx)
        l.add_widget(ModernButton(text="ZAPISZ", on_press=lambda x: [json.dump({'h':self.ti_shost.text,'port':self.ti_sport.text,'u':self.ti_suser.text,'p':self.ti_spass.text,'batch':self.cb_batching.active}, open(p,"w")), self.msg("OK","Zapisano")]))
        l.add_widget(ModernButton(text="TEST", on_press=lambda x: self.test_smtp_direct(), bg_color=(.1,.7,.4,1))); l.add_widget(ModernButton(text="POWRÓT", on_press=lambda x: setattr(self.sm, 'current', 'home'), bg_color=(.3,.3,.3,1))); self.sc_ref["smtp"].add_widget(l)

    def test_smtp_direct(self):
        try: srv = self.connect_smtp({'h':self.ti_shost.text,'port':self.ti_sport.text,'u':self.ti_suser.text,'p':self.ti_spass.text}); srv.quit(); self.msg("OK", "Serwer SMTP OK")
        except Exception as e: self.msg("Błędne dane", str(e))

    def ask_before_send_worker(self, row, email, n, s):
        def dec(v): self.user_decision = "send" if v else "skip"; self.wait_for_user = False; px.dismiss()
        box = BoxLayout(orientation="vertical", padding=dp(20), spacing=dp(10))
        box.add_widget(Label(text=f"POTWIERDŹ:\n[b]{n} {s}[/b]\n{email}", markup=True, halign="center"))
        btns = BoxLayout(size_hint_y=None, height=dp(55), spacing=dp(10))
        btns.add_widget(Button(text="WYŚLIJ", on_press=lambda x: dec(True), background_color=(0,0.6,0,1)))
        btns.add_widget(Button(text="POMIŃ", on_press=lambda x: dec(False), background_color=(0.7,0,0,1)))
        box.add_widget(btns); px = Popup(title="Manualna weryfikacja", content=box, size_hint=(0.9, 0.45), auto_dismiss=False); px.open()

    def update_stats(self, *a):
        try: self.lbl_stats.text = f"Baza: {self.conn.execute('SELECT count(*) FROM contacts').fetchone()[0]} | Załączniki: {len(self.global_attachments)}"
        except: pass
    def update_progress(self, d): self.pb.value = int((d/self.total_q)*100); self.pb_label.text = f"Wysyłanie: {d}/{self.total_q}"
    def finish_mailing(self, s): 
        self.is_mailing_running = False; det = "\n".join(self.session_details); self.conn.execute("INSERT INTO reports (date, ok, fail, skip, auto, details) VALUES (?,?,?,?,?,?)", (datetime.now().strftime("%Y-%m-%d %H:%M"), self.stats['ok'], self.stats['fail'], self.stats['skip'], 0, det)); self.conn.commit()
        Clock.schedule_once(lambda dt: self.msg("Status", f"{s}\nSukces: {self.stats['ok']}"))

    def msg(self, tit, txt):
        b = BoxLayout(orientation="vertical", padding=dp(20)); b.add_widget(Label(text=txt, halign="center")); b.add_widget(ModernButton(text="OK", on_press=lambda x: p.dismiss(), height=dp(50), size_hint_y=None)); p = Popup(title=tit, content=b, size_hint=(0.8, 0.4)); p.open()

    # --- TABELA I KONTAKTY (1:1 CORE) ---

    def setup_table_ui(self):
        root = BoxLayout(orientation="vertical"); menu = BoxLayout(size_hint_y=None, height=dp(55), spacing=dp(5), padding=dp(5))
        self.ti_tab_search = ModernInput(hint_text="Szukaj..."); self.ti_tab_search.bind(text=self.filter_table)
        menu.add_widget(self.ti_tab_search); menu.add_widget(Button(text="Opcje", size_hint_x=0.2, on_press=self.popup_columns)); menu.add_widget(Button(text="Wróć", size_hint_x=0.2, on_press=lambda x: setattr(self.sm, 'current', 'home')))
        hs = ScrollView(size_hint_y=None, height=dp(55), do_scroll_y=False); self.table_header_layout = GridLayout(rows=1, size_hint=(None, None), height=dp(55)); hs.add_widget(self.table_header_layout)
        ds = ScrollView(do_scroll_x=True, do_scroll_y=True); self.table_content_layout = GridLayout(size_hint=(None, None)); self.table_content_layout.bind(minimum_height=self.table_content_layout.setter('height'), minimum_width=self.table_content_layout.setter('width'))
        ds.add_widget(self.table_content_layout); ds.bind(scroll_x=lambda inst, val: setattr(hs, 'scroll_x', val)); root.add_widget(menu); root.add_widget(hs); root.add_widget(ds); self.sc_ref["table"].add_widget(root)

    def refresh_table(self):
        self.table_content_layout.clear_widgets(); self.table_header_layout.clear_widgets()
        if not self.filtered_data: return
        w, h = dp(180), dp(55); headers = [self.full_data[0][i] for i in self.export_indices]
        self.table_header_layout.cols = len(headers) + 1; self.table_header_layout.width = (len(headers) + 1) * w
        for head in headers: self.table_header_layout.add_widget(ColorSafeLabel(text=str(head), bg_color=COLOR_HEADER, bold=True, size=(w,h), size_hint=(None,None)))
        self.table_header_layout.add_widget(ColorSafeLabel(text="Akcja", bg_color=COLOR_HEADER, bold=True, size=(w,h), size_hint=(None,None)))
        self.table_content_layout.cols = len(headers) + 1; self.table_content_layout.width = (len(headers)+1)*w
        for r_idx, row in enumerate(self.filtered_data[1:]):
            row_bg = COLOR_ROW_A if r_idx % 2 == 0 else COLOR_ROW_B
            for c_idx in self.export_indices:
                val = str(row[c_idx]) if c_idx < len(row) and str(row[c_idx]).strip() != "" else "0"
                self.table_content_layout.add_widget(ColorSafeLabel(text=val, bg_color=row_bg, size=(w,h), size_hint=(None,None)))
            bt = Button(text="Zapisz", size=(w,h), size_hint=(None,None)); bt.bind(on_press=lambda x, r=row: self.export_single_row(r)); self.table_content_layout.add_widget(bt)

    def export_single_row(self, r):
        p = Path("/storage/emulated/0/Documents/FutureExport") if platform=="android" else Path("./exports"); p.mkdir(parents=True, exist_ok=True)
        nx, sx = str(r[self.idx_name]).title(), str(r[self.idx_surname]).title(); wb = Workbook(); ws = wb.active
        ws.append([self.full_data[0][k] for k in self.export_indices]); ws.append([str(r[k]) if (k < len(r) and str(r[k]).strip() != "") else "0" for k in self.export_indices])
        self.style_xlsx(ws); wb.save(p/f"Raport_{nx}_{sx}.xlsx"); self.msg("OK", f"Zapisano PDF dla: {nx}")

    def setup_contacts_ui(self):
        l, top = BoxLayout(orientation="vertical", padding=dp(10)), BoxLayout(size_hint_y=None, height=dp(55), spacing=dp(5))
        self.ti_csearch = TextInput(hint_text="Szukaj..."); self.ti_csearch.bind(text=self.refresh_contacts_list)
        top.add_widget(self.ti_csearch); top.add_widget(Button(text="+", size_hint_x=0.15, on_press=lambda x: self.form_contact())); top.add_widget(Button(text="Wróć", size_hint_x=0.2, on_press=lambda x: setattr(self.sm, 'current', 'email')))
        self.c_list = GridLayout(cols=1, size_hint_y=None, spacing=dp(10)); self.c_list.bind(minimum_height=self.c_list.setter('height'))
        sc = ScrollView(); sc.add_widget(self.c_list); l.add_widget(top); l.add_widget(sc); self.sc_ref["contacts"].add_widget(l)

    def refresh_contacts_list(self, *args):
        self.c_list.clear_widgets(); sv = self.ti_csearch.text.lower(); rows = self.conn.execute("SELECT name, surname, email, pesel, phone FROM contacts ORDER BY surname ASC").fetchall()
        for d in rows:
            if sv and sv not in f"{d[0]} {d[1]} {d[2]}".lower(): continue
            row = BoxLayout(size_hint_y=None, height=dp(125), padding=dp(10))
            with row.canvas.before: Color(*COLOR_CARD); Rectangle(pos=row.pos, size=row.size)
            info, acts = BoxLayout(orientation="vertical"), BoxLayout(size_hint_x=0.3, orientation="vertical", spacing=dp(4))
            info.add_widget(Label(text=f"{d[0]} {d[1]}".title(), bold=True, halign="left", text_size=(dp(250),None)))
            info.add_widget(Label(text=f"E: {d[2]}\nP: {d[3]}\nT: {d[4] if d[4] else '-'}", font_size='11sp', halign="left", text_size=(dp(250),None), color=(0.7,0.7,0.7,1)))
            row.add_widget(info); acts.add_widget(Button(text="Edytuj", on_press=lambda x, data=d: self.form_contact(*data))); acts.add_widget(Button(text="Usuń", background_color=(0.8,0.2,0.2,1), on_press=lambda x, n=d[0], s=d[1]: self.delete_contact(n, s))); row.add_widget(acts); self.c_list.add_widget(row)

    def delete_contact(self, n, s):
        def pr(_): [self.conn.execute("DELETE FROM contacts WHERE name=? AND surname=?", (n, s)), self.conn.commit(), px.dismiss(), self.refresh_contacts_list(), self.update_stats()]
        px = Popup(title="Usuń?", content=Button(text="USUŃ KONTAKT", on_press=pr, background_color=(1,0,0,1)), size_hint=(0.7,0.3)); px.open()
    def form_contact(self, n="", s="", e="", pes="", ph=""):
        b, f_ins = BoxLayout(orientation="vertical", padding=dp(15), spacing=dp(10)), [TextInput(text=str(n), hint_text="Imię"), TextInput(text=str(s), hint_text="Nazwisko"), TextInput(text=str(e), hint_text="Email"), TextInput(text=str(pes), hint_text="PESEL"), TextInput(text=str(ph), hint_text="Telefon")]
        for f in f_ins: b.add_widget(f)
        def sv(_): [self.conn.execute("INSERT OR REPLACE INTO contacts VALUES (?,?,?,?,?)", (f_ins[0].text.lower(), f_ins[1].text.lower(), f_ins[2].text.strip(), f_ins[3].text.strip(), f_ins[4].text.strip())), self.conn.commit(), px.dismiss(), self.refresh_contacts_list(), self.update_stats()]
        b.add_widget(ModernButton(text="ZAPISZ", on_press=sv)); px = Popup(title="Kontakt", content=b, size_hint=(0.9, 0.85)); px.open()

    def filter_table(self, i, v): self.filtered_data = [self.full_data[0]] + [r for r in self.full_data[1:] if any(v.lower() in str(c).lower() for c in r)]; self.refresh_table()
    def start_special_send_flow(self, _): self.open_picker("special_send")
    def process_excel(self, path):
        try:
            if str(path).endswith(".xls") and xlrd:
                wb = xlrd.open_workbook(path); ws = wb.sheet_by_index(0); raw = [[str(ws.cell_value(r,c)).strip() for c in range(ws.ncols)] for r in range(ws.nrows)]
            else:
                wb = load_workbook(path, data_only=True); ws = wb.active; raw = [["" if v is None else str(v).strip() for v in r] for r in ws.iter_rows(values_only=True)]
            h_idx = 0
            for i, r in enumerate(raw[:15]):
                if any(x in " ".join([str(v) for v in r]).lower() for x in ["imię", "imie", "nazwisko"]): h_idx = i; break
            self.full_data, self.export_indices = raw[h_idx:], list(range(len(raw[h_idx][0])))
            for i,v in enumerate(self.full_data[0]):
                v = str(v).lower()
                if "imi" in v: self.idx_name = i
                if "naz" in v: self.idx_surname = i
                if "pesel" in v: self.idx_pesel = i
            self.msg("OK", "Arkusz wczytany")
        except: self.msg("BŁĄD", "Plik uszkodzony")
    def setup_tmpl_ui(self):
        l, ti_s, ti_b = BoxLayout(orientation="vertical", padding=dp(25), spacing=dp(10)), ModernInput(hint_text="Temat {Imię}"), ModernInput(hint_text="Treść...", multiline=True)
        ts, tb = self.conn.execute("SELECT val FROM settings WHERE key='t_sub'").fetchone(), self.conn.execute("SELECT val FROM settings WHERE key='t_body'").fetchone()
        ti_s.text, ti_b.text = (ts[0] if ts else ""), (tb[0] if tb else "")
        l.add_widget(Label(text="SZABLON EMAIL", bold=True)); l.add_widget(ti_s); l.add_widget(ti_b)
        l.add_widget(ModernButton(text="ZAPISZ", on_press=lambda x: [self.conn.execute("INSERT OR REPLACE INTO settings VALUES (?,?)", ('t_sub',ti_s.text)), self.conn.execute("INSERT OR REPLACE INTO settings VALUES (?,?)", ('t_body',ti_b.text)), self.conn.commit(), self.msg("OK","Zapamiętano")]))
        l.add_widget(ModernButton(text="POWRÓT", on_press=lambda x: setattr(self.sm, 'current', 'email'))); self.sc_ref["tmpl"].add_widget(l)
    def setup_report_ui(self):
        l, self.r_grid = BoxLayout(orientation="vertical", padding=dp(15), spacing=dp(10)), GridLayout(cols=1, size_hint_y=None, spacing=dp(10))
        self.r_grid.bind(minimum_height=self.r_grid.setter('height')); sc = ScrollView(); sc.add_widget(self.r_grid); l.add_widget(Label(text="HISTORIA SESJI", bold=True, height=dp(40), size_hint_y=None)); l.add_widget(sc); l.add_widget(ModernButton(text="POWRÓT", on_press=lambda x: setattr(self.sm, 'current', 'home'), height=dp(55), size_hint_y=None)); self.sc_ref["report"].add_widget(l)
    def refresh_reports(self, *a):
        self.r_grid.clear_widgets(); rows = self.conn.execute("SELECT date, ok, fail, skip, details FROM reports ORDER BY id DESC").fetchall()
        for d, ok, fl, sk, det in rows:
            row = BoxLayout(orientation="vertical", size_hint_y=None, height=dp(110), padding=dp(10))
            with row.canvas.before: Color(0.15, 0.2, 0.25, 1); Rectangle(pos=row.pos, size=row.size)
            row.add_widget(Label(text=f"Sesja: {d}", bold=True, color=COLOR_PRIMARY)); row.add_widget(Button(text="Pokaż logi", size_hint_y=None, height=dp(35), on_press=lambda x, t=det: self.show_details(t))); self.r_grid.add_widget(row)
    def show_details(self, t):
        b = BoxLayout(orientation="vertical", padding=dp(10)); ti = TextInput(text=str(t), readonly=True, font_size='11sp'); b.add_widget(ti); b.add_widget(Button(text="ZAMKNIJ", size_hint_y=0.2, on_press=lambda x: p.dismiss())); p = Popup(title="Logi", content=b, size_hint=(.9,.8)); p.open()
    def popup_columns(self, _):
        box, gr, checks = BoxLayout(orientation="vertical", padding=dp(10)), GridLayout(cols=1, size_hint_y=None, spacing=dp(5)), []
        gr.bind(minimum_height=gr.setter('height'))
        for i, h in enumerate(self.full_data[0]):
            r, cb = BoxLayout(size_hint_y=None, height=dp(45)), CheckBox(active=(i in self.export_indices), size_hint_x=None, width=dp(50)); checks.append((i, cb)); r.add_widget(cb); r.add_widget(Label(text=str(h))); gr.add_widget(r)
        sc = ScrollView(); sc.add_widget(gr); box.add_widget(sc); box.add_widget(ModernButton(text="ZASTOSUJ", on_press=lambda x: [setattr(self, 'export_indices', [idx for idx, c in checks if c.active]), p.dismiss(), self.refresh_table()], height=dp(50), size_hint_y=None)); p = Popup(title="Kolumny", content=box, size_hint=(.9,.9)); p.open()

if __name__ == "__main__":
    FutureApp().run()
