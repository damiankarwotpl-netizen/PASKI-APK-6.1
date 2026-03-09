import os
import json
import sqlite3
import threading
import smtplib
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

from openpyxl import load_workbook, Workbook
from openpyxl.styles import Border, Side, Font, Alignment

try:
    import xlrd
except ImportError:
    xlrd = None

# --- KONFIGURACJA UI ---
COLOR_PRIMARY = (0.12, 0.45, 0.9, 1)
COLOR_BG = (0.05, 0.07, 0.1, 1)
COLOR_TEXT = (0.9, 0.9, 0.9, 1)

class PremiumButton(Button):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.background_normal = ""
        self.background_color = COLOR_PRIMARY
        self.color = (1, 1, 1, 1)
        self.size_hint_y = None
        self.height = dp(52)
        self.bold = True

class SafeLabel(Label):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.halign = 'center'
        self.valign = 'middle'
        self.color = COLOR_TEXT
        self.bind(size=self._update_text_size)
    def _update_text_size(self, instance, value):
        self.text_size = (value[0] - dp(10), None)

# --- EKRANY ---
class HomeScreen(Screen): pass
class TableScreen(Screen): pass
class EmailScreen(Screen): pass
class SMTPScreen(Screen): pass
class TemplateScreen(Screen): pass
class ContactsScreen(Screen): pass

class FutureApp(App):
    def build(self):
        Window.clearcolor = COLOR_BG
        self.full_data = []      
        self.filtered_data = []  
        self.current_file = None
        self.global_attachments = []
        self.export_indices = [] 
        self.idx_name = 0
        self.idx_surname = 1
        
        self.init_db()
        self.sm = ScreenManager()
        self.screens = {
            "home": HomeScreen(name="home"),
            "table": TableScreen(name="table"),
            "email": EmailScreen(name="email"),
            "smtp": SMTPScreen(name="smtp"),
            "tmpl": TemplateScreen(name="tmpl"),
            "contacts": ContactsScreen(name="contacts")
        }
        self.setup_ui()
        for s in self.screens.values(): self.sm.add_widget(s)
        return self.sm

    def init_db(self):
        db_p = Path(self.user_data_dir) / "future_pro_v16.db"
        self.conn = sqlite3.connect(str(db_p), check_same_thread=False)
        self.conn.execute("CREATE TABLE IF NOT EXISTS contacts (name TEXT, surname TEXT, email TEXT, PRIMARY KEY(name, surname))")
        self.conn.execute("CREATE TABLE IF NOT EXISTS settings (key TEXT PRIMARY KEY, val TEXT)")
        self.conn.commit()

# -----------------------------
# ODCZYT EXCEL + AUTO-HEADER
# -----------------------------
    def read_excel(self, path):
        p = str(path).lower()
        raw_data = []
        try:
            if p.endswith(".xlsx"):
                wb = load_workbook(path, data_only=True); ws = wb.active
                raw_data = [["" if v is None else str(v).strip() for v in r] for r in ws.iter_rows(values_only=True)]
            elif p.endswith(".xls"):
                if not xlrd: return None, "Brak xlrd"
                wb = xlrd.open_workbook(path); ws = wb.sheet_by_index(0)
                for r in range(ws.nrows):
                    row = []
                    for c in range(ws.ncols):
                        val = ws.cell_value(r, c)
                        if isinstance(val, float) and val.is_integer(): val = int(val)
                        row.append(str(val).strip())
                    raw_data.append(row)
            
            # --- AUTO-DETEKCJA NAGŁÓWKA ---
            header_idx = 0
            for i, row in enumerate(raw_data[:10]): # Sprawdź pierwsze 10 wierszy
                row_str = " ".join(row).lower()
                if any(k in row_str for k in ["nazwisko", "imię", "spółka", "stawka"]):
                    header_idx = i
                    break
            
            cleaned_data = raw_data[header_idx:]
            return cleaned_data, None
        except Exception as e: return None, str(e)

# -----------------------------
# STYLIZACJA EXCEL (RAMKI I AUTO-FIT)
# -----------------------------
    def apply_expert_styling(self, ws):
        # 1. Auto-dopasowanie szerokości kolumn
        for col in ws.columns:
            m_len = 0
            c_name = col[0].column_letter
            for cell in col:
                if cell.value: m_len = max(m_len, len(str(cell.value)))
            ws.column_dimensions[c_name].width = m_len + 4

        # 2. Pogrubione ramki i formatowanie
        thick = Side(style='thick', color="000000")
        thin = Side(style='thin', color="000000")
        
        for r_idx, row in enumerate(ws.iter_rows()):
            for cell in row:
                # Ramki wewnętrzne (cienkie)
                cell.border = Border(top=thin, left=thin, right=thin, bottom=thin)
                cell.alignment = Alignment(horizontal='center', vertical='center')
                # Nagłówek
                if r_idx == 0:
                    cell.font = Font(bold=True, color="FFFFFF")
                    cell.fill = Font() # Tutaj można dodać wypełnienie
                    cell.border = Border(top=thick, left=thin, right=thin, bottom=thick)

# -----------------------------
# UI: PODGLĄD TABELI
# -----------------------------
    def setup_table_ui(self):
        root = BoxLayout(orientation="vertical", padding=dp(8), spacing=dp(8))
        menu = BoxLayout(size_hint_y=None, height=dp(55), spacing=dp(5))
        self.search_ti = TextInput(hint_text="Szukaj...", multiline=False, size_hint_x=0.5)
        self.search_ti.bind(text=self.apply_filter)
        btn_col = Button(text="Wybierz Kolumny", size_hint_x=0.3, on_press=self.popup_columns)
        btn_back = Button(text="Wróć", size_hint_x=0.2, on_press=lambda x: setattr(self.sm, "current", "home"))
        menu.add_widget(self.search_ti); menu.add_widget(btn_col); menu.add_widget(btn_back)

        self.table_scroll = ScrollView(do_scroll_x=True, do_scroll_y=True)
        self.table_grid = GridLayout(size_hint=(None, None), spacing=dp(2))
        self.table_grid.bind(minimum_height=self.table_grid.setter("height"), minimum_width=self.table_grid.setter("width"))
        self.table_scroll.add_widget(self.table_grid)
        
        self.table_progress = ProgressBar(max=100, size_hint_y=None, height=dp(10))
        root.add_widget(menu); root.add_widget(self.table_scroll); root.add_widget(self.table_progress)
        self.screens["table"].add_widget(root)

    def refresh_table_view(self):
        self.table_grid.clear_widgets()
        if not self.filtered_data or len(self.filtered_data) == 0: return
        
        cls = len(self.filtered_data[0])
        w_c, h_c = dp(180), dp(60)
        self.table_grid.cols = cls + 1
        self.table_grid.width = (cls + 1) * w_c
        self.table_grid.height = len(self.filtered_data) * h_c

        for h in self.filtered_data[0]:
            self.table_grid.add_widget(SafeLabel(text=str(h), bold=True, color=COLOR_PRIMARY, size_hint=(None, None), size=(w_c, h_c)))
        self.table_grid.add_widget(SafeLabel(text="Akcja", bold=True, size_hint=(None, None), size=(w_c, h_c)))

        for row_data in self.filtered_data[1:]:
            for cell in row_data:
                self.table_grid.add_widget(SafeLabel(text=str(cell), size_hint=(None, None), size=(w_c, h_c)))
            btn = Button(text="Zapisz", size_hint=(None, None), size=(w_c, h_c))
            btn.bind(on_press=lambda x, r=row_data: self.export_single(r))
            self.table_grid.add_widget(btn)

# -----------------------------
# ZARZĄDZANIE KONTAKTAMI
# -----------------------------
    def setup_contacts_ui(self):
        root = BoxLayout(orientation="vertical", padding=dp(10), spacing=dp(10))
        top = BoxLayout(size_hint_y=None, height=dp(55), spacing=dp(10))
        self.cont_search = TextInput(hint_text="Szukaj osoby...", multiline=False)
        self.cont_search.bind(text=self.refresh_contacts_list)
        btn_add = Button(text="+ Dodaj", size_hint_x=0.3, on_press=lambda x: self.form_contact())
        btn_back = Button(text="Wróć", size_hint_x=0.2, on_press=lambda x: setattr(self.sm, "current", "email"))
        top.add_widget(self.cont_search); top.add_widget(btn_add); top.add_widget(btn_back)
        
        header = BoxLayout(size_hint_y=None, height=dp(40))
        header.add_widget(Label(text="Dane / Nazwisko", bold=True, size_hint_x=0.4))
        header.add_widget(Label(text="Email", bold=True, size_hint_x=0.4))
        header.add_widget(SafeLabel(text="Akcje", bold=True, size_hint_x=0.2))
        
        self.cont_scroll = ScrollView(); self.cont_list = GridLayout(cols=1, size_hint_y=None, spacing=dp(10))
        self.cont_list.bind(minimum_height=self.cont_list.setter('height'))
        self.cont_scroll.add_widget(self.cont_list)
        root.add_widget(top); root.add_widget(header); root.add_widget(self.cont_scroll)
        self.screens["contacts"].add_widget(root)

    def refresh_contacts_list(self, *args):
        self.cont_list.clear_widgets()
        sv = self.cont_search.text.lower()
        rows = self.conn.execute("SELECT name, surname, email FROM contacts ORDER BY surname ASC").fetchall()
        for n, s, e in rows:
            name_full = f"{n} {s}".title()
            if sv and sv not in name_full.lower() and sv not in e.lower(): continue
            row = BoxLayout(size_hint_y=None, height=dp(65), padding=dp(5))
            row.add_widget(Label(text=name_full, size_hint_x=0.4, halign="left", text_size=(dp(150), None)))
            row.add_widget(Label(text=e, size_hint_x=0.4, color=(0.7,0.7,0.7,1), halign="left", text_size=(dp(150), None)))
            btns = BoxLayout(size_hint_x=0.2, spacing=dp(2))
            be = Button(text="✏"); be.bind(on_press=lambda x, n=n, s=s, e=e: self.form_contact(n,s,e))
            bd = Button(text="❌", background_color=(0.8,0.2,0.2,1))
            bd.bind(on_press=lambda x, n=n, s=s: self.delete_contact(n,s))
            btns.add_widget(be); btns.add_widget(bd); row.add_widget(btns)
            self.cont_list.add_widget(row)

# -----------------------------
# LOGIKA MAILINGU & EKSPORTU
# -----------------------------
    def start_mailing(self, _):
        if not self.full_data: self.popup_msg("!", "Wczytaj Płace!"); return
        threading.Thread(target=self._mailing_thread, daemon=True).start()

    def _mailing_thread(self):
        p = Path(self.user_data_dir) / "smtp.json"; cfg = json.load(open(p)) if p.exists() else None
        if not cfg: Clock.schedule_once(lambda d: self.popup_msg("!", "Ustaw SMTP!")); return
        try:
            srv = smtplib.SMTP("smtp.gmail.com", 587, timeout=15); srv.starttls(); srv.login(cfg['u'], cfg['p'])
        except Exception as e: Clock.schedule_once(lambda d: self.popup_msg("Błąd SMTP", str(e))); return
        
        rows = self.full_data[1:]; sent = 0
        for i, row in enumerate(rows):
            n, s = row[self.idx_name].lower(), row[self.idx_surname].lower()
            res = self.conn.execute("SELECT email FROM contacts WHERE name=? AND surname=?", (n, s)).fetchone()
            if res:
                try:
                    msg = EmailMessage(); dat = datetime.now().strftime("%d.%m.%Y")
                    ts = self.conn.execute("SELECT val FROM settings WHERE key='t_sub'").fetchone()
                    tb = self.conn.execute("SELECT val FROM settings WHERE key='t_body'").fetchone()
                    msg["Subject"] = (ts[0] if ts else "Raport").replace("{Imię}", row[self.idx_name])
                    msg["From"], msg["To"] = cfg['u'], res[0]
                    msg.set_content((tb[0] if tb else "Informacja").replace("{Imię}", row[self.idx_name]).replace("{Data}", dat))
                    
                    tmp = Path(self.user_data_dir) / f"r_{i}.xlsx"; wb = Workbook(); ws = wb.active
                    ws.append([self.full_data[0][k] for k in self.export_indices])
                    ws.append([row[k] for k in self.export_indices])
                    self.apply_expert_styling(ws); wb.save(tmp)
                    
                    with open(tmp, "rb") as f: msg.add_attachment(f.read(), maintype="application", subtype="xlsx", filename=f"Raport_{row[self.idx_name]}.xlsx")
                    for ex in self.global_attachments:
                        if os.path.exists(ex):
                            with open(ex, "rb") as f: msg.add_attachment(f.read(), maintype="application", subtype="octet-stream", filename=os.path.basename(ex))
                    srv.send_message(msg); sent += 1
                except: pass
        srv.quit(); Clock.schedule_once(lambda d: self.popup_msg("Raport", f"Wysłano {sent} maili."))

    def export_single(self, r):
        f = Path("/storage/emulated/0/Documents/FutureExport"); f.mkdir(parents=True, exist_ok=True)
        wb = Workbook(); ws = wb.active; ws.append([self.full_data[0][k] for k in self.export_indices]); ws.append([r[k] for k in self.export_indices])
        self.apply_expert_styling(ws)
        fname = f"Raport_{r[self.idx_name]}.xlsx"; wb.save(f / fname); self.popup_msg("Sukces", f"Zapisano:\n{fname}")

# -----------------------------
# POZOSTAŁE FUNKCJE UI
# -----------------------------
    def setup_ui(self):
        # Home
        l = BoxLayout(orientation="vertical", padding=dp(30), spacing=dp(15))
        l.add_widget(Label(text="FUTURE 16.0 GOLD", font_size='28sp', bold=True, color=COLOR_PRIMARY))
        l.add_widget(PremiumButton(text="📂 WCZYTAJ PLIK PŁAC", on_press=lambda x: self.open_picker("data")))
        l.add_widget(PremiumButton(text="📊 PODGLĄD TABELI", on_press=lambda x: [self.refresh_table_view(), setattr(self.sm, 'current', 'table')] if self.full_data else self.popup_msg("!","Wczytaj Excel")))
        l.add_widget(PremiumButton(text="✉ CENTRUM MAILINGU", on_press=lambda x: setattr(self.sm, "current", "email")))
        l.add_widget(PremiumButton(text="⚙ USTAWIENIA SMTP", on_press=lambda x: setattr(self.sm, "current", "smtp")))
        self.h_stat = Label(text="System Gotowy", color=(0.5,0.5,0.5,1)); l.add_widget(self.h_stat); self.screens["home"].add_widget(l)
        
        self.setup_table_ui(); self.setup_email_ui(); self.setup_smtp_ui(); self.setup_tmpl_ui(); self.setup_contacts_ui()

    def setup_email_ui(self):
        l = BoxLayout(orientation="vertical", padding=dp(25), spacing=dp(10))
        l.add_widget(Label(text="CENTRUM MAILINGOWE", font_size='22sp', bold=True))
        self.e_stat = Label(text="Baza: 0"); l.add_widget(self.e_stat)
        def btn(t, c): b = PremiumButton(text=t); b.bind(on_press=c); l.add_widget(b)
        btn("📁 WCZYTAJ KONTAKTY (EXCEL)", lambda x: self.open_picker("book"))
        btn("🔧 ZARZĄDZAJ BAZĄ", lambda x: [self.refresh_contacts_list(), setattr(self.sm, 'current', 'contacts')])
        btn("📝 EDYTUJ TREŚĆ MAILA", lambda x: setattr(self.sm, "current", "tmpl"))
        btn("📎 DODAJ ZAŁĄCZNIK", lambda x: self.open_picker("attachment"))
        btn("🚀 URUCHOM WYSYŁKĘ", self.start_mailing)
        btn("POWRÓT", lambda x: setattr(self.sm, "current", "home"))
        self.screens["email"].add_widget(l); self.update_stats()

    def popup_columns(self, _):
        if not self.full_data: return
        box = BoxLayout(orientation="vertical", padding=dp(15), spacing=dp(5)); sc = ScrollView(); gr = GridLayout(cols=1, size_hint_y=None); gr.bind(minimum_height=gr.setter('height'))
        checks = []
        for i, h in enumerate(self.full_data[0]):
            r = BoxLayout(size_hint_y=None, height=dp(45)); cb = CheckBox(active=(i in self.export_indices), size_hint_x=None, width=dp(50))
            checks.append((i,cb)); r.add_widget(cb); r.add_widget(Label(text=str(h), halign="left", text_size=(Window.width*0.65, None))); gr.add_widget(r)
        def sv(_): self.export_indices = [i for i,c in checks if c.active]; p.dismiss()
        sc.add_widget(gr); box.add_widget(sc); box.add_widget(PremiumButton(text="OK", on_press=sv))
        p = Popup(title="Widoczność danych", content=box, size_hint=(0.95, 0.9)); p.open()

    def update_stats(self, *args):
        c = self.conn.execute("SELECT count(*) FROM contacts").fetchone()[0]
        self.e_stat.text = f"Baza: {c} osób | Załączniki: {len(self.global_attachments)}"

    def setup_smtp_ui(self):
        l = BoxLayout(orientation="vertical", padding=dp(25), spacing=dp(10))
        self.su = TextInput(hint_text="Email Gmail", multiline=False); self.sp = TextInput(hint_text="Hasło App (16 znaków)", password=True, multiline=False)
        p = Path(self.user_data_dir) / "smtp.json"
        if p.exists(): d = json.load(open(p)); self.su.text, self.sp.text = d.get('u',''), d.get('p','')
        btn_sv = PremiumButton(text="ZAPISZ", on_press=lambda x: [json.dump({'u':self.su.text, 'p':self.sp.text}, open(p, "w")), self.popup_msg("OK","Zapisano")])
        l.add_widget(Label(text="USTAWIENIA SMTP", bold=True)); l.add_widget(self.su); l.add_widget(self.sp); l.add_widget(btn_sv)
        l.add_widget(PremiumButton(text="POWRÓT", on_press=lambda x: setattr(self.sm, 'current', 'home'), background_color=(0.4,0.4,0.4,1))); self.screens["smtp"].add_widget(l)

    def setup_tmpl_ui(self):
        l = BoxLayout(orientation="vertical", padding=dp(25), spacing=dp(10))
        self.ts = TextInput(hint_text="Temat {Imię}", size_hint_y=None, height=dp(45)); self.tb = TextInput(hint_text="Treść maila...", multiline=True)
        r = self.conn.execute("SELECT val FROM settings WHERE key='t_sub'").fetchone(); rb = self.conn.execute("SELECT val FROM settings WHERE key='t_body'").fetchone()
        if r: self.ts.text, self.tb.text = r[0], rb[0]
        def sv(_): self.conn.execute("INSERT OR REPLACE INTO settings VALUES (?,?)", ('t_sub', self.ts.text)); self.conn.execute("INSERT OR REPLACE INTO settings VALUES (?,?)", ('t_body', self.tb.text)); self.conn.commit(); self.popup_msg("OK","Zapisano")
        l.add_widget(Label(text="TREŚĆ MAILA", bold=True)); l.add_widget(self.ts); l.add_widget(self.tb); l.add_widget(PremiumButton(text="ZAPISZ", on_press=sv)); l.add_widget(PremiumButton(text="POWRÓT", on_press=lambda x: setattr(self.sm, 'current', 'email'), background_color=(0.4,0.4,0.4,1))); self.screens["tmpl"].add_widget(l)

    def open_picker(self, mode):
        if platform != "android": self.popup_msg("Info", "Tylko Android"); return
        from jnius import autoclass; from android import activity
        PA = autoclass("org.kivy.android.PythonActivity"); Intent = autoclass("android.content.Intent")
        intent = Intent(Intent.ACTION_GET_CONTENT); intent.setType("*/*")
        def cb(req, res, dt):
            if not dt: return
            activity.unbind(on_activity_result=cb)
            uri = dt.getData(); stream = PA.mActivity.getContentResolver().openInputStream(uri)
            loc = Path(self.user_data_dir) / f"f_{mode}_{os.urandom(2).hex()}.xlsx"
            with open(loc, "wb") as f:
                while True:
                    b = stream.read(bytearray(4096))
                    if b == -1: break
                    f.write(b)
            stream.close()
            if mode == "data": 
                d, e = self.read_excel(loc)
                if not e: self.full_data = d; self.filtered_data = d; self.h_stat.text = "Excel wczytany!"; self.export_indices = list(range(len(d[0])))
            elif mode == "book": self.import_book(loc)
            elif mode == "attachment": self.global_attachments.append(str(loc)); self.update_stats()
        activity.bind(on_activity_result=cb); PA.mActivity.startActivityForResult(intent, 1001)

    def import_book(self, path):
        d, e = self.read_excel(path); h = [str(x).lower().strip() for x in d[0]]
        in_n, in_s, in_e = 0, 1, 2
        for i,v in enumerate(h):
            if "imi" in v: in_n=i
            elif "naz" in v: in_s=i
            elif "@" in v or "mail" in v: in_e=i
        for r in d[header_idx+1 if 'header_idx' in locals() else 1:]:
            if len(r) > max(in_n, in_s, in_e) and r[in_e]:
                self.conn.execute("INSERT OR REPLACE INTO contacts VALUES (?,?,?)", (r[in_n].lower(), r[in_s].lower(), r[in_e]))
        self.conn.commit(); self.update_stats()

    def form_contact(self, n="", s="", e=""):
        b = BoxLayout(orientation="vertical", padding=dp(15), spacing=dp(10)); in_n = TextInput(text=n, hint_text="Imię"); in_s = TextInput(text=s, hint_text="Nazwisko"); in_e = TextInput(text=e, hint_text="Email")
        if n: in_n.readonly = True; in_s.readonly = True
        def sv(_): self.conn.execute("INSERT OR REPLACE INTO contacts VALUES (?,?,?)", (in_n.text.strip().lower(), in_s.text.strip().lower(), in_e.text.strip())); self.conn.commit(); p.dismiss(); self.refresh_contacts_list(); self.update_stats()
        b.add_widget(in_n); b.add_widget(in_s); b.add_widget(in_e); b.add_widget(PremiumButton(text="ZAPISZ", on_press=sv)); p = Popup(title="Kontakt", content=b, size_hint=(0.9, 0.6)); p.open()

    def delete_contact(self, n, s):
        def pr(_): self.conn.execute("DELETE FROM contacts WHERE name=? AND surname=?", (n, s)); self.conn.commit(); p.dismiss(); self.refresh_contacts_list(); self.update_stats()
        p = Popup(title="Usuń?", content=Button(text="USUŃ", on_press=pr, background_color=(1,0,0,1)), size_hint=(0.7,0.35)); p.open()

    def apply_filter(self, ins, val):
        v = val.lower(); self.filtered_data = [self.full_data[0]] + [r for r in self.full_data[1:] if any(v in str(c).lower() for c in r)]; self.refresh_table_view()

    def popup_msg(self, title, text):
        box = BoxLayout(orientation="vertical", padding=dp(20)); box.add_widget(Label(text=text, halign="center")); b = PremiumButton(text="ZAMKNIJ", on_press=lambda x: p.dismiss()); box.add_widget(b); p = Popup(title=title, content=box, size_hint=(0.85, 0.45)); p.open()

if __name__ == "__main__": FutureApp().run()
