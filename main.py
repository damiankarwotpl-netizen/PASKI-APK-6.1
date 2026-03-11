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

# Paleta Kolorów
COLOR_PRIMARY = (0.1, 0.55, 0.95, 1)
COLOR_SUCCESS = (0.15, 0.75, 0.5, 1)
COLOR_BG = (0.05, 0.06, 0.09, 1)
COLOR_CARD = (0.12, 0.14, 0.2, 1)
COLOR_TEXT = (0.95, 0.95, 0.98, 1)
COLOR_DANGER = (0.9, 0.25, 0.25, 1)
COLOR_HEADER = (0.15, 0.2, 0.3, 1)
COLOR_ROW_A = (0.05, 0.06, 0.09, 1)
COLOR_ROW_B = (0.1, 0.12, 0.18, 1)

# --- KOMPONENTY UI ---

class ModernButton(Button):
    def __init__(self, bg_color=COLOR_PRIMARY, **kwargs):
        super().__init__(**kwargs)
        self.background_normal = ""
        self.background_down = ""
        self.background_color = (0,0,0,0)
        self.color = COLOR_TEXT
        self.bold = True
        self.font_size = '14sp'
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
            self.line = RoundedRectangle(pos=self.pos, size=self.size, radius=[dp(10)])
        self.bind(pos=self._update, size=self._update)
    def _update(self, *args):
        self.line.pos = self.pos
        self.line.size = self.size

class ColorSafeLabel(Label):
    def __init__(self, bg_color=(1,1,1,1), text_color=(1,1,1,1), **kwargs):
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

class FutureApp(App):
    def build(self):
        Window.clearcolor = COLOR_BG
        self.full_data, self.filtered_data, self.export_indices = [], [], []
        self.global_attachments, self.selected_emails, self.queue = [], [], []
        self.session_details = []
        self.stats = {"ok": 0, "fail": 0, "skip": 0, "auto": 0}
        self.idx_name, self.idx_surname, self.idx_pesel = 0, 1, -1
        self.auto_send_mode = False
        
        self.init_db()
        self.sm = ScreenManager(transition=SlideTransition())
        self.setup_screens()
        return self.sm

    def init_db(self):
        db_p = Path(self.user_data_dir) / "future_ultimate_v3.db"
        self.conn = sqlite3.connect(str(db_p), check_same_thread=False)
        self.conn.execute("CREATE TABLE IF NOT EXISTS contacts (name TEXT, surname TEXT, email TEXT, pesel TEXT, phone TEXT, PRIMARY KEY(name, surname))")
        self.conn.execute("CREATE TABLE IF NOT EXISTS settings (key TEXT PRIMARY KEY, val TEXT)")
        self.conn.execute("CREATE TABLE IF NOT EXISTS reports (id INTEGER PRIMARY KEY AUTOINCREMENT, date TEXT, ok INTEGER, fail INTEGER, skip INTEGER, auto INTEGER, details TEXT)")
        self.conn.commit()

    def setup_screens(self):
        self.sc_map = {
            "home": HomeScreen(name="home"), "table": TableScreen(name="table"),
            "email": EmailScreen(name="email"), "smtp": SMTPScreen(name="smtp"),
            "tmpl": TemplateScreen(name="tmpl"), "contacts": ContactsScreen(name="contacts"),
            "report": ReportScreen(name="report")
        }
        self.build_ui(); 
        for s in self.sc_map.values(): self.sm.add_widget(s)

    # --- UI DESIGNS ---

    def build_ui(self):
        # Home
        l = BoxLayout(orientation="vertical", padding=dp(30), spacing=dp(15))
        l.add_widget(Label(text="FUTURE ULTIMATE", font_size='32sp', bold=True, color=COLOR_PRIMARY, size_hint_y=None, height=dp(80)))
        btns = [
            ("📊 WCZYTAJ LISTĘ PŁAC", lambda x: self.open_picker("data")),
            ("👁️ PODGLĄD I EKSPORT", lambda x: [self.refresh_table(), setattr(self.sm, 'current', 'table')] if self.full_data else self.msg("!", "Brak danych")),
            ("📧 CENTRUM MAILINGU", lambda x: setattr(self.sm, 'current', 'email')),
            ("📜 RAPORTY WYSYŁEK", lambda x: [self.refresh_reports(), setattr(self.sm, 'current', 'report')]),
            ("⚙️ USTAWIENIA GMAIL", lambda x: setattr(self.sm, 'current', 'smtp'))
        ]
        for t, c in btns: l.add_widget(ModernButton(text=t, on_press=c, height=dp(60), size_hint_y=None))
        self.sc_map["home"].add_widget(l)

        # Inicjalizacja reszty UI
        self.setup_table_ui(); self.setup_email_ui(); self.setup_smtp_ui(); self.setup_tmpl_ui(); self.setup_contacts_ui(); self.setup_report_ui()

    def setup_email_ui(self):
        l = BoxLayout(orientation="vertical", padding=dp(25), spacing=dp(10))
        ab = BoxLayout(size_hint_y=None, height=dp(45), spacing=dp(10))
        self.cb_auto = CheckBox(size_hint_x=None, width=dp(45))
        self.cb_auto.bind(active=lambda i,v: setattr(self, 'auto_send_mode', v))
        ab.add_widget(self.cb_auto); ab.add_widget(Label(text="AUTOMATYCZNA WYSYŁKA", bold=True))
        l.add_widget(ab)
        self.lbl_stats = Label(text="Baza: 0", height=dp(30)); l.add_widget(self.lbl_stats)
        self.pb_label = Label(text="Gotowy", height=dp(25)); self.pb = ProgressBar(max=100, height=dp(20))
        l.add_widget(self.pb_label); l.add_widget(self.pb)
        
        btns = [
            ("IMPORT KONTAKTÓW", lambda x: self.open_picker("book")),
            ("ZARZĄDZAJ BAZĄ", lambda x: [self.refresh_contacts_list(), setattr(self.sm, 'current', 'contacts')]),
            ("EDYTUJ SZABLON", lambda x: setattr(self.sm, 'current', 'tmpl')),
            ("DODAJ ZAŁĄCZNIK", lambda x: self.open_picker("attachment")),
            ("WYŚLIJ JEDEN PLIK", self.start_special_send_flow),
            ("START MASOWA WYSYŁKA", self.start_mass_mailing)
        ]
        for t, c in btns: l.add_widget(ModernButton(text=t, on_press=c, height=dp(50), size_hint_y=None))
        l.add_widget(ModernButton(text="POWRÓT", on_press=lambda x: setattr(self.sm, 'current', 'home'), bg_color=(0.3,0.3,0.3,1)))
        self.sc_map["email"].add_widget(l); self.update_stats()

    # --- MAILING ENGINE (AUDIT 1:1) ---

    def start_mass_mailing(self, _):
        if not self.full_data: self.msg("!", "Danych brak!"); return
        self.stats = {"ok": 0, "fail": 0, "skip": 0, "auto": 0}; self.session_details = []
        self.queue = list(self.full_data[1:]); self.total_q = len(self.queue)
        Clock.schedule_once(self.process_mailing_queue)

    def process_mailing_queue(self, *args):
        done = self.total_q - len(self.queue)
        self.pb.value = int((done/self.total_q)*100) if self.total_q > 0 else 100
        self.pb_label.text = f"Postęp: {done}/{self.total_q}"
        if not self.queue:
            det = "\n".join(self.session_details)
            self.conn.execute("INSERT INTO reports (date, ok, fail, skip, auto, details) VALUES (?,?,?,?,?,?)", (datetime.now().strftime("%Y-%m-%d %H:%M"), self.stats['ok'], self.stats['fail'], self.stats['skip'], self.stats['auto'], det)); self.conn.commit()
            self.msg("Koniec", f"OK: {self.stats['ok']}\nPominięto: {self.stats['skip']}"); return
        
        row = self.queue.pop(0)
        try:
            n, s = str(row[self.idx_name]).strip(), str(row[self.idx_surname]).strip()
            p = str(row[self.idx_pesel]).strip() if self.idx_pesel != -1 else ""
            res = None
            # Inteligentne dopasowanie: PESEL -> Imię/Nazwisko
            if p and len(p) > 5:
                res = self.conn.execute("SELECT email FROM contacts WHERE pesel=?", (p,)).fetchone()
            if not res:
                res = self.conn.execute("SELECT email FROM contacts WHERE name=? AND surname=? COLLATE NOCASE", (n.lower(), s.lower())).fetchone()
            
            if res:
                if self.auto_send_mode: self.send_email_engine(row, res[0])
                else: self.ask_before_send(row, res[0], n, s)
            else:
                self.stats["skip"] += 1; self.session_details.append(f"SKIP: {n} {s} (Brak w bazie)"); Clock.schedule_once(self.process_mailing_queue)
        except: self.stats["skip"] += 1; Clock.schedule_once(self.process_mailing_queue)

    def send_email_engine(self, row_data, target):
        def task():
            cfg_p = Path(self.user_data_dir) / "smtp.json"
            if not cfg_p.exists(): return
            cfg = json.load(open(cfg_p))
            nx, sx = str(row_data[self.idx_name]).title(), str(row_data[self.idx_surname]).title()
            try:
                srv = smtplib.SMTP("smtp.gmail.com", 587, timeout=15); srv.starttls(); srv.login(cfg['u'], cfg['p'])
                msg = EmailMessage(); dat = datetime.now().strftime("%d.%m.%Y")
                ts = self.conn.execute("SELECT val FROM settings WHERE key='t_sub'").fetchone()
                tb = self.conn.execute("SELECT val FROM settings WHERE key='t_body'").fetchone()
                msg["Subject"] = (ts[0] if ts else "Raport").replace("{Imię}", nx)
                msg["From"], msg["To"] = cfg['u'], target
                msg.set_content((tb[0] if tb else "Dzień dobry").replace("{Imię}", nx).replace("{Data}", dat))
                
                # Generowanie Excel 1:1 z pierwotnym stylem
                tmp = Path(self.user_data_dir) / "r_tmp.xlsx"
                wb = Workbook(); ws = wb.active
                ws.append([self.full_data[0][k] for k in self.export_indices])
                # Wypełnianie 0 dla pustych komórek (Audit 1:1)
                ws.append([str(row_data[k]) if (k < len(row_data) and str(row_data[k]).strip() != "") else "0" for k in self.export_indices])
                self.style_xlsx(ws); wb.save(tmp)
                msg.add_attachment(open(tmp, "rb").read(), maintype="application", subtype="xlsx", filename=f"Raport_{nx}_{sx}.xlsx")
                
                for p in self.global_attachments:
                    if os.path.exists(p):
                        ct, _ = mimetypes.guess_type(p); mn, sb = (ct or 'application/octet-stream').split('/', 1)
                        msg.add_attachment(open(p, "rb").read(), maintype=mn, subtype=sb, filename=os.path.basename(p))
                
                srv.send_message(msg); srv.quit()
                self.session_details.append(f"OK: {nx} {sx}"); self.stats["ok"] += 1
            except Exception as e:
                self.session_details.append(f"ERR: {nx} {sx} ({str(e)[:20]})"); self.stats["fail"] += 1
            Clock.schedule_once(lambda d: self.process_mailing_queue())
        threading.Thread(target=task, daemon=True).start()

    # --- TABELA I EKSPORT 1:1 ---

    def style_xlsx(self, ws):
        s, c = Side(style='thin'), Alignment(horizontal='center', vertical='center')
        for ri, row in enumerate(ws.iter_rows(), 1):
            for cell in row:
                cell.border = Border(top=s, left=s, right=s, bottom=s); cell.alignment = c
                if ri == 1: 
                    cell.font = Font(bold=True); cell.fill = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")
                elif ri % 2 == 0: 
                    cell.fill = PatternFill(start_color="F7F7F7", end_color="F7F7F7", fill_type="solid")
        for col in ws.columns:
            max_l = 0; column_letter = col[0].column_letter
            for cell in col:
                if cell.value: max_l = max(max_l, len(str(cell.value)))
            ws.column_dimensions[column_letter].width = (max_l * 1.3) + 6

    def setup_table_ui(self):
        root = BoxLayout(orientation="vertical")
        menu = BoxLayout(size_hint_y=None, height=dp(55), spacing=dp(5), padding=dp(5))
        self.ti_search = ModernInput(hint_text="Szukaj..."); self.ti_search.bind(text=self.filter_table)
        menu.add_widget(self.ti_search)
        menu.add_widget(Button(text="Opcje", size_hint_x=0.2, on_press=self.popup_columns))
        menu.add_widget(Button(text="Wróć", size_hint_x=0.2, on_press=lambda x: setattr(self.sm, 'current', 'home')))
        # Nieruchomy nagłówek (Audit 1:1)
        hs = ScrollView(size_hint_y=None, height=dp(55), do_scroll_y=False)
        self.table_header_layout = GridLayout(rows=1, size_hint=(None, None), height=dp(55))
        hs.add_widget(self.table_header_layout)
        # Przewijana treść
        ds = ScrollView(do_scroll_x=True, do_scroll_y=True)
        self.table_content_layout = GridLayout(size_hint=(None, None))
        self.table_content_layout.bind(minimum_height=self.table_content_layout.setter('height'), minimum_width=self.table_content_layout.setter('width'))
        ds.add_widget(self.table_content_layout)
        ds.bind(scroll_x=lambda inst, val: setattr(hs, 'scroll_x', val))
        root.add_widget(menu); root.add_widget(hs); root.add_widget(ds)
        self.sc_map["table"].add_widget(root)

    def refresh_table(self):
        self.table_content_layout.clear_widgets(); self.table_header_layout.clear_widgets()
        if not self.filtered_data or not self.export_indices: return
        w, h = dp(180), dp(55); headers = [self.full_data[0][i] for i in self.export_indices]
        self.table_header_layout.cols = len(headers) + 1; self.table_header_layout.width = (len(headers) + 1) * w
        for head in headers: self.table_header_layout.add_widget(ColorSafeLabel(text=str(head), bg_color=COLOR_HEADER, bold=True, size=(w,h), size_hint=(None,None)))
        self.table_header_layout.add_widget(ColorSafeLabel(text="Akcja", bg_color=COLOR_HEADER, bold=True, size=(w,h), size_hint=(None,None)))
        
        self.table_content_layout.cols = len(headers) + 1; self.table_content_layout.width = (len(headers)+1)*w
        for r_idx, row in enumerate(self.filtered_data[1:]):
            # Naprzemienne kolory wierszy (Audit 1:1)
            row_bg = COLOR_ROW_A if r_idx % 2 == 0 else COLOR_ROW_B
            for c_idx in self.export_indices:
                val = str(row[c_idx]) if c_idx < len(row) and str(row[c_idx]).strip()!="" else "0"
                self.table_content_layout.add_widget(ColorSafeLabel(text=val, bg_color=row_bg, size=(w,h), size_hint=(None,None)))
            bt = Button(text="Eksport", size=(w,h), size_hint=(None,None)); bt.bind(on_press=lambda x, r=row: self.export_single_row(r))
            self.table_content_layout.add_widget(bt)

    def export_single_row(self, r):
        # Ścieżki zależne od platformy (Audit 1:1)
        p = Path("/storage/emulated/0/Documents/FutureExport") if platform=="android" else Path("./exports")
        p.mkdir(parents=True, exist_ok=True)
        nx, sx = str(r[self.idx_name]).title(), str(r[self.idx_surname]).title()
        wb = Workbook(); ws = wb.active
        ws.append([self.full_data[0][k] for k in self.export_indices])
        ws.append([str(r[k]) if (k < len(r) and str(r[k]).strip() != "") else "0" for k in self.export_indices])
        self.style_xlsx(ws); wb.save(p/f"Raport_{nx}_{sx}.xlsx")
        self.msg("OK", f"Zapisano plik: {nx}")

    # --- BAZA KONTAKTÓW I RAPORTY (AUDIT 1:1) ---

    def setup_contacts_ui(self):
        l = BoxLayout(orientation="vertical", padding=dp(10))
        top = BoxLayout(size_hint_y=None, height=dp(55), spacing=dp(5))
        self.ti_csearch = ModernInput(hint_text="Szukaj..."); self.ti_csearch.bind(text=self.refresh_contacts_list)
        top.add_widget(self.ti_csearch); 
        top.add_widget(Button(text="+", size_hint_x=0.15, on_press=lambda x: self.form_contact()))
        top.add_widget(Button(text="Wróć", size_hint_x=0.2, on_press=lambda x: setattr(self.sm, 'current', 'email')))
        self.c_list = GridLayout(cols=1, size_hint_y=None, spacing=dp(10)); self.c_list.bind(minimum_height=self.c_list.setter('height'))
        sc = ScrollView(); sc.add_widget(self.c_list); l.add_widget(top); l.add_widget(sc)
        self.sc_map["contacts"].add_widget(l)

    def refresh_contacts_list(self, *args):
        self.c_list.clear_widgets(); sv = self.ti_csearch.text.lower()
        rows = self.conn.execute("SELECT name, surname, email, pesel, phone FROM contacts ORDER BY surname ASC").fetchall()
        for d in rows:
            if sv and sv not in f"{d[0]} {d[1]} {d[2]}".lower(): continue
            row = BoxLayout(size_hint_y=None, height=dp(110), padding=dp(8))
            with row.canvas.before: Color(*COLOR_CARD); Rectangle(pos=row.pos, size=row.size)
            info = BoxLayout(orientation="vertical")
            info.add_widget(Label(text=f"{d[0]} {d[1]}".title(), bold=True, halign="left", text_size=(dp(200),None)))
            info.add_widget(Label(text=f"{d[2]}\nPESEL: {d[3]}", font_size='11sp', halign="left", text_size=(dp(200),None)))
            row.add_widget(info)
            acts = BoxLayout(size_hint_x=0.3, orientation="vertical", spacing=dp(4))
            # Przycisk EDYTUJ (Audit 1:1)
            acts.add_widget(Button(text="Edytuj", on_press=lambda x, data=d: self.form_contact(*data)))
            acts.add_widget(Button(text="Usuń", background_color=COLOR_DANGER, on_press=lambda x, n=d[0], s=d[1]: self.delete_contact(n, s)))
            row.add_widget(acts); self.c_list.add_widget(row)

    def setup_report_ui(self):
        l = BoxLayout(orientation="vertical", padding=dp(15), spacing=dp(10))
        self.r_grid = GridLayout(cols=1, size_hint_y=None, spacing=dp(12)); self.r_grid.bind(minimum_height=self.r_grid.setter('height'))
        sc = ScrollView(); sc.add_widget(self.r_grid)
        l.add_widget(Label(text="HISTORIA WYSYŁEK", bold=True, height=dp(40), size_hint_y=None))
        l.add_widget(sc)
        l.add_widget(ModernButton(text="POWRÓT", on_press=lambda x: setattr(self.sm, 'current', 'home'), height=dp(55), size_hint_y=None))
        self.sc_map["report"].add_widget(l)

    def refresh_reports(self, *args):
        self.r_grid.clear_widgets()
        rows = self.conn.execute("SELECT date, ok, fail, skip, details FROM reports ORDER BY id DESC").fetchall()
        for d, ok, fl, sk, det in rows:
            row = BoxLayout(orientation="vertical", size_hint_y=None, height=dp(110), padding=dp(10))
            with row.canvas.before: Color(0.15, 0.2, 0.25, 1); Rectangle(pos=row.pos, size=row.size)
            row.add_widget(Label(text=f"Sesja: {d}", bold=True, color=COLOR_PRIMARY))
            row.add_widget(Label(text=f"OK: {ok} | ERR: {fl} | SKIP: {sk}", font_size='13sp'))
            # Szczegóły raportu (Audit 1:1)
            btn = Button(text="Szczegóły logów", size_hint_y=None, height=dp(35), on_press=lambda x, t=det: self.show_details(t))
            row.add_widget(btn); self.r_grid.add_widget(row)

    def show_details(self, txt):
        b = BoxLayout(orientation="vertical", padding=dp(10)); ti = TextInput(text=str(txt), readonly=True, font_size='11sp')
        b.add_widget(ti); b.add_widget(Button(text="ZAMKNIJ", size_hint_y=0.2, on_press=lambda x: p.dismiss()))
        p = Popup(title="Logi sesji", content=b, size_hint=(.9,.8)); p.open()

    # --- PICKER I POMOCNICZE ---

    def open_picker(self, mode):
        if platform != "android": self.msg("!", "Tylko na Android"); return
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
                except: self.msg("!", "Błąd pliku")
        activity.bind(on_activity_result=cb); PA.mActivity.startActivityForResult(intent, 1001)

    def process_excel(self, path):
        try:
            if str(path).endswith(".xls") and xlrd:
                wb = xlrd.open_workbook(path); ws = wb.sheet_by_index(0); raw = [[str(ws.cell_value(r,c)).strip() for c in range(ws.ncols)] for r in range(ws.nrows)]
            else:
                wb = load_workbook(path, data_only=True); ws = wb.active; raw = [["" if v is None else str(v).strip() for v in r] for r in ws.iter_rows(values_only=True)]
            h_idx = 0
            for i, r in enumerate(raw[:15]):
                if any(x in " ".join([str(v) for v in r]).lower() for x in ["imię", "imie", "nazwisko"]): h_idx = i; break
            self.full_data = raw[h_idx:]; self.filtered_data = self.full_data; self.export_indices = list(range(len(self.full_data[0])))
            for i,v in enumerate([str(x).lower() for x in self.full_data[0]]):
                if "imi" in v: self.idx_name = i
                if "naz" in v: self.idx_surname = i
                if "pesel" in v: self.idx_pesel = i
            self.msg("OK", "Wczytano arkusz.")
        except Exception as e: self.msg("ERR", str(e))

    def process_book(self, path):
        try:
            wb = load_workbook(path, data_only=True); ws = wb.active; raw = list(ws.iter_rows(values_only=True))
            h = [str(x).lower() for x in raw[0]]; iN, iS, iE, iP = 0, 1, 2, -1
            for i,v in enumerate(h):
                if "imi" in v: iN=i
                elif "naz" in v: iS=i
                elif "@" in v or "mail" in v: iE=i
                elif "pesel" in v: iP=i
            for r in raw[1:]:
                if r[iE] and "@" in str(r[iE]):
                    self.conn.execute("INSERT OR REPLACE INTO contacts (name, surname, email, pesel, phone) VALUES (?,?,?,?,?)", (str(r[iN]).lower(), str(r[iS]).lower(), str(r[iE]).strip(), str(r[iP]) if (iP!=-1) else "", ""))
            self.conn.commit(); self.update_stats(); self.msg("OK", "Baza gotowa.")
        except: self.msg("ERR", "Błąd importu.")

    def filter_table(self, inst, val):
        v = val.lower(); self.filtered_data = [self.full_data[0]] + [r for r in self.full_data[1:] if any(v in str(c).lower() for c in r)]; self.refresh_table()

    def popup_columns(self, _):
        if not self.full_data: return
        box = BoxLayout(orientation="vertical", padding=dp(10)); gr = GridLayout(cols=1, size_hint_y=None, spacing=dp(5)); gr.bind(minimum_height=gr.setter('height')); checks = []
        for i, h in enumerate(self.full_data[0]):
            r = BoxLayout(size_hint_y=None, height=dp(45)); cb = CheckBox(active=(i in self.export_indices), size_hint_x=None, width=dp(50)); checks.append((i, cb)); r.add_widget(cb); r.add_widget(Label(text=str(h))); gr.add_widget(r)
        sc = ScrollView(); sc.add_widget(gr); box.add_widget(sc)
        box.add_widget(ModernButton(text="ZASTOSUJ", on_press=lambda x: [setattr(self, 'export_indices', [idx for idx, c in checks if c.active]), p.dismiss(), self.refresh_table()], height=dp(50), size_hint_y=None))
        p = Popup(title="Kolumny", content=box, size_hint=(.9,.9)); p.open()

    def delete_contact(self, n, s):
        def go(_): [self.conn.execute("DELETE FROM contacts WHERE name=? AND surname=?", (n, s)), self.conn.commit(), px.dismiss(), self.refresh_contacts_list(), self.update_stats()]
        px = Popup(title="Usuń?", content=Button(text="POTWIERDŹ USUNIĘCIE", on_press=go, background_color=COLOR_DANGER), size_hint=(.8,.3)); px.open()

    def form_contact(self, n="", s="", e="", pes="", ph=""):
        b = BoxLayout(orientation="vertical", padding=dp(15), spacing=dp(10))
        ins = [ModernInput(text=str(n), hint_text="Imię"), ModernInput(text=str(s), hint_text="Nazwisko"), ModernInput(text=str(e), hint_text="Email"), ModernInput(text=str(pes), hint_text="PESEL")]
        for i in ins: b.add_widget(i)
        def save(_): [self.conn.execute("INSERT OR REPLACE INTO contacts (name, surname, email, pesel, phone) VALUES (?,?,?,?,?)", (ins[0].text.lower(), ins[1].text.lower(), ins[2].text, ins[3].text, "")), self.conn.commit(), px.dismiss(), self.refresh_contacts_list(), self.update_stats()]
        b.add_widget(ModernButton(text="ZAPISZ", on_press=save)); px = Popup(title="Kontakt", content=b, size_hint=(.9, .75)); px.open()

    def start_special_send_flow(self, _): self.open_picker("special_send")
    def special_send_step_2_recipients(self, file_path):
        self.selected_emails = []
        box = BoxLayout(orientation="vertical", padding=dp(15), spacing=dp(10))
        ti = ModernInput(hint_text="Szukaj..."); box.add_widget(ti)
        sc = ScrollView(); gl = GridLayout(cols=1, size_hint_y=None, spacing=dp(5)); gl.bind(minimum_height=gl.setter('height')); sc.add_widget(gl); box.add_widget(sc)
        def rf(v=""):
            gl.clear_widgets(); rows = self.conn.execute("SELECT name, surname, email FROM contacts").fetchall()
            for r in rows:
                if v and v.lower() not in f"{r[0]} {r[1]} {r[2]}".lower(): continue
                bx = BoxLayout(size_hint_y=None, height=dp(50))
                cb = CheckBox(size_hint_x=None, width=dp(50)); cb.bind(active=lambda inst, val, m=r[2]: self.selected_emails.append(m) if val else self.selected_emails.remove(m))
                bx.add_widget(cb); bx.add_widget(Label(text=f"{r[0].title()} {r[1].title()}")); gl.add_widget(bx)
        ti.bind(text=lambda i,v: rf(v)); rf()
        btn = ModernButton(text="DALEJ", on_press=lambda x: [p.dismiss(), self.special_send_step_3(file_path)]); box.add_widget(btn)
        p = Popup(title="Odbiorcy", content=box, size_hint=(.95,.9)); p.open()

    def special_send_step_3(self, path):
        b = BoxLayout(orientation="vertical", padding=dp(15), spacing=dp(10))
        ti_s = ModernInput(hint_text="Temat"); ti_b = ModernInput(hint_text="Treść", multiline=True)
        b.add_widget(ti_s); b.add_widget(ti_b)
        def run(_):
            def thread_task():
                cfg = json.load(open(Path(self.user_data_dir)/"smtp.json"))
                srv = smtplib.SMTP("smtp.gmail.com", 587); srv.starttls(); srv.login(cfg['u'], cfg['p'])
                for m in self.selected_emails:
                    msg = EmailMessage(); msg["Subject"], msg["From"], msg["To"] = ti_s.text, cfg['u'], m; msg.set_content(ti_b.text)
                    with open(path, "rb") as f: msg.add_attachment(f.read(), maintype="application", subtype="octet-stream", filename=os.path.basename(path))
                    srv.send_message(msg)
                srv.quit(); Clock.schedule_once(lambda d: self.msg("OK", "Wysłano pomyślnie"))
            threading.Thread(target=thread_task, daemon=True).start(); p.dismiss()
        b.add_widget(ModernButton(text="WYŚLIJ", on_press=run)); p = Popup(title="Wiadomość", content=b, size_hint=(.9, .8)); p.open()

    def msg(self, tit, txt):
        b = BoxLayout(orientation="vertical", padding=dp(20)); b.add_widget(Label(text=txt, halign="center")); btn = Button(text="OK", height=dp(50), size_hint_y=None, on_press=lambda x: p.dismiss()); b.add_widget(btn); p = Popup(title=tit, content=b, size_hint=(0.8, 0.4)); p.open()
    def update_stats(self, *a):
        try: self.lbl_stats.text = f"Baza: {self.conn.execute('SELECT count(*) FROM contacts').fetchone()[0]} | Załączniki: {len(self.global_attachments)}"
        except: pass
    def setup_smtp_ui(self):
        l = BoxLayout(orientation="vertical", padding=dp(30), spacing=dp(10)); self.ti_su = ModernInput(hint_text="Gmail"); self.ti_sp = ModernInput(hint_text="Hasło", password=True)
        p = Path(self.user_data_dir) / "smtp.json"; d = json.load(open(p)) if p.exists() else {}; self.ti_su.text, self.ti_sp.text = d.get('u',''), d.get('p','')
        l.add_widget(Label(text="USTAWIENIA GMAIL", bold=True)); l.add_widget(self.ti_su); l.add_widget(self.ti_sp)
        l.add_widget(ModernButton(text="ZAPISZ", on_press=lambda x: [json.dump({'u':self.ti_su.text,'p':self.ti_sp.text}, open(p,"w")), self.msg("OK","Zapisano")]))
        l.add_widget(ModernButton(text="POWRÓT", on_press=lambda x: setattr(self.sm, 'current', 'home'), bg_color=(0.3,0.3,0.3,1))); self.sc_map["smtp"].add_widget(l)
    def setup_tmpl_ui(self):
        l = BoxLayout(orientation="vertical", padding=dp(25), spacing=dp(10)); self.ti_ts = ModernInput(hint_text="Temat {Imię}"); self.ti_tb = ModernInput(hint_text="Treść...", multiline=True)
        ts = self.conn.execute("SELECT val FROM settings WHERE key='t_sub'").fetchone(); tb = self.conn.execute("SELECT val FROM settings WHERE key='t_body'").fetchone(); self.ti_ts.text = ts[0] if ts else ""; self.ti_tb.text = tb[0] if tb else ""
        sv = lambda x: [self.conn.execute("INSERT OR REPLACE INTO settings VALUES (?,?)", ('t_sub',self.ti_ts.text)), self.conn.execute("INSERT OR REPLACE INTO settings VALUES (?,?)", ('t_body',self.ti_tb.text)), self.conn.commit(), self.msg("OK","Zapisano")]
        l.add_widget(Label(text="SZABLON EMAIL", bold=True)); l.add_widget(self.ti_ts); l.add_widget(self.ti_tb); l.add_widget(ModernButton(text="ZAPISZ", on_press=sv)); l.add_widget(ModernButton(text="POWRÓT", on_press=lambda x: setattr(self.sm, 'current', 'email'))); self.sc_map["tmpl"].add_widget(l)

if __name__ == "__main__":
    FutureApp().run()
