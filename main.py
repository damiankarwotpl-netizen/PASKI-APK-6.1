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

# --- STYLIZACJA ---
COLOR_PRIMARY = (0.1, 0.5, 0.9, 1)
COLOR_BG = (0.08, 0.1, 0.15, 1)

class PremiumButton(Button):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.background_normal = ""
        self.background_color = COLOR_PRIMARY
        self.height = dp(52)
        self.size_hint_y = None
        self.bold = True

class ScreenTitle(Label):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.font_size = '22sp'
        self.bold = True
        self.size_hint_y = None
        self.height = dp(60)

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
        self.add_screens()
        return self.sm

    def init_db(self):
        db_p = Path(self.user_data_dir) / "future_data_v17.db"
        self.conn = sqlite3.connect(str(db_p), check_same_thread=False)
        self.conn.execute("CREATE TABLE IF NOT EXISTS contacts (name TEXT, surname TEXT, email TEXT, PRIMARY KEY(name, surname))")
        self.conn.execute("CREATE TABLE IF NOT EXISTS settings (key TEXT PRIMARY KEY, val TEXT)")
        self.conn.commit()

    def add_screens(self):
        self.screens = {
            "home": HomeScreen(name="home"), "table": TableScreen(name="table"),
            "email": EmailScreen(name="email"), "smtp": SMTPScreen(name="smtp"),
            "tmpl": TemplateScreen(name="tmpl"), "contacts": ContactsScreen(name="contacts")
        }
        self.setup_home_ui(); self.setup_table_ui(); self.setup_email_ui()
        self.setup_smtp_ui(); self.setup_tmpl_ui(); self.setup_contacts_ui()
        for s in self.screens.values(): self.sm.add_widget(s)

# -----------------------------
# PICKER PLIKÓW (STABILNY)
# -----------------------------
    def open_picker(self, mode):
        if platform != "android": self.popup("Błąd", "Picker dostępny tylko na Android"); return
        from jnius import autoclass; from android import activity
        PA = autoclass("org.kivy.android.PythonActivity"); Intent = autoclass("android.content.Intent")
        intent = Intent(Intent.ACTION_GET_CONTENT); intent.setType("*/*")
        
        def picker_callback(req, res, intent_data):
            if not intent_data: return
            activity.unbind(on_activity_result=picker_callback)
            uri = intent_data.getData()
            resolver = PA.mActivity.getContentResolver()
            stream = resolver.openInputStream(uri)
            
            ext = ".xlsx" if mode != "attachment" else ""
            local_path = Path(self.user_data_dir) / f"temp_{mode}{ext}"
            
            with open(local_path, "wb") as f:
                buffer = bytearray(1024 * 16)
                while True:
                    read = stream.read(buffer)
                    if read == -1: break
                    f.write(buffer[:read])
            stream.close()
            
            if mode == "data": self.load_data_file(local_path)
            elif mode == "book": self.load_contacts_file(local_path)
            elif mode == "attachment": 
                self.global_attachments.append(str(local_path))
                self.update_email_stats()

        activity.bind(on_activity_result=picker_callback)
        PA.mActivity.startActivityForResult(intent, 1001)

# -----------------------------
# ODCZYT I AUTO-NAGŁÓWEK
# -----------------------------
    def load_data_file(self, path):
        data, err = self.read_excel_universal(path)
        if err: self.popup("Błąd", err); return
        self.full_data = data; self.filtered_data = data
        self.idx_name = 0; self.idx_surname = 1
        h = [str(x).lower() for x in data[0]]
        for i, v in enumerate(h):
            if "imi" in v: self.idx_name = i
            if "naz" in v and i != self.idx_name: self.idx_surname = i
        self.export_indices = list(range(len(data[0])))
        self.popup("Sukces", "Załadowano arkusz płac")

    def read_excel_universal(self, path):
        raw = []
        try:
            if str(path).endswith(".xls") and xlrd:
                wb = xlrd.open_workbook(path); ws = wb.sheet_by_index(0)
                for r in range(ws.nrows): raw.append([str(ws.cell_value(r, c)).strip() for c in range(ws.ncols)])
            else:
                wb = load_workbook(path, data_only=True); ws = wb.active
                raw = [["" if v is None else str(v).strip() for v in r] for r in ws.iter_rows(values_only=True)]
            
            # Detekcja wiersza nagłówkowego
            h_idx = 0
            for i, row in enumerate(raw[:15]):
                txt = " ".join(row).lower()
                if any(k in txt for k in ["imię", "imie", "nazwisko", "stawka"]): h_idx = i; break
            return raw[h_idx:], None
        except Exception as e: return None, str(e)

# -----------------------------
# STYLIZACJA EXCEL (RAMKI I AUTO-FIT)
# -----------------------------
    def apply_style_and_save(self, wb, path):
        ws = wb.active
        thin = Side(style='thin'); thick = Side(style='thick')
        for row in ws.iter_rows():
            for cell in row:
                cell.border = Border(top=thin, left=thin, right=thin, bottom=thin)
                cell.alignment = Alignment(horizontal='center')
        # Header bold + ramka
        for cell in ws[1]:
            cell.font = Font(bold=True)
            cell.border = Border(top=thick, bottom=thick, left=thin, right=thin)
        # Auto-fit
        for col in ws.columns:
            dim = max(len(str(c.value or "")) for c in col)
            ws.column_dimensions[col[0].column_letter].width = dim + 3
        wb.save(path)

# -----------------------------
# UI: TABELA (NAPRAWA NACHODZENIA)
# -----------------------------
    def setup_table_ui(self):
        l = BoxLayout(orientation="vertical", padding=dp(8), spacing=dp(5))
        top = BoxLayout(size_hint_y=None, height=dp(55), spacing=dp(5))
        self.ti_search = TextInput(hint_text="Szukaj...", multiline=False)
        self.ti_search.bind(text=self.filter_table)
        btn_cols = Button(text="Kolumny", size_hint_x=0.3, on_press=self.show_column_popup)
        btn_back = Button(text="Wróć", size_hint_x=0.2, on_press=lambda x: setattr(self.sm, "current", "home"))
        top.add_widget(self.ti_search); top.add_widget(btn_cols); top.add_widget(btn_back)
        
        self.table_scroll = ScrollView(do_scroll_x=True, do_scroll_y=True)
        self.table_grid = GridLayout(size_hint=(None, None), spacing=dp(2))
        self.table_grid.bind(minimum_height=self.table_grid.setter("height"), minimum_width=self.table_grid.setter("width"))
        self.table_scroll.add_widget(self.table_grid)
        l.add_widget(top); l.add_widget(self.table_scroll); self.screens["table"].add_widget(l)

    def refresh_table(self):
        self.table_grid.clear_widgets()
        if not self.filtered_data: return
        rws, cls = len(self.filtered_data), len(self.filtered_data[0])
        w, h = dp(200), dp(55) # Stała szerokość chroni przed nachodzeniem
        self.table_grid.cols = cls + 1
        self.table_grid.width, self.table_grid.height = (cls+1)*w, rws*h
        
        for head in self.filtered_data[0]:
            self.table_grid.add_widget(Label(text=str(head), bold=True, color=COLOR_PRIMARY, size_hint=(None,None), size=(w,h), text_size=(w-dp(10), None), halign="center"))
        self.table_grid.add_widget(Label(text="Akcja", bold=True, size_hint=(None,None), size=(w,h)))
        
        for r in self.filtered_data[1:]:
            for c in r:
                self.table_grid.add_widget(Label(text=str(c), size_hint=(None,None), size=(w,h), text_size=(w-dp(10), None), halign="center", valign="middle"))
            b = Button(text="Eksport", size_hint=(None,None), size=(w,h)); b.bind(on_press=lambda x, row=r: self.export_row(row))
            self.table_grid.add_widget(b)

# -----------------------------
# UI: KONTAKTY (NAPRAWA NACHODZENIA)
# -----------------------------
    def setup_contacts_ui(self):
        l = BoxLayout(orientation="vertical", padding=dp(10), spacing=dp(10))
        top = BoxLayout(size_hint_y=None, height=dp(55), spacing=dp(5))
        self.ti_csearch = TextInput(hint_text="Szukaj osoby...", multiline=False)
        self.ti_csearch.bind(text=self.refresh_contacts)
        btn_add = Button(text="+ Dodaj", size_hint_x=0.3, on_press=lambda x: self.form_contact())
        btn_back = Button(text="Wróć", size_hint_x=0.2, on_press=lambda x: setattr(self.sm, "current", "email"))
        top.add_widget(self.ti_csearch); top.add_widget(btn_add); top.add_widget(btn_back)
        
        self.contacts_scroll = ScrollView(); self.contacts_list = GridLayout(cols=1, size_hint_y=None, spacing=dp(10))
        self.contacts_list.bind(minimum_height=self.contacts_list.setter("height"))
        self.contacts_scroll.add_widget(self.contacts_list)
        l.add_widget(top); l.add_widget(self.contacts_scroll); self.screens["contacts"].add_widget(l)

    def refresh_contacts(self, *args):
        self.contacts_list.clear_widgets()
        search_val = self.ti_csearch.text.lower()
        rows = self.conn.execute("SELECT name, surname, email FROM contacts ORDER BY surname ASC").fetchall()
        for n, s, e in rows:
            name = f"{n} {s}".title()
            if search_val and search_val not in name.lower() and search_val not in e.lower(): continue
            # Specjalny wiersz chroniący przed nakładaniem
            row = BoxLayout(orientation="vertical", size_hint_y=None, height=dp(80), padding=dp(10))
            line1 = BoxLayout(); line1.add_widget(Label(text=name, bold=True, halign="left", text_size=(dp(200),None)))
            
            actions = BoxLayout(size_hint_x=None, width=dp(100), spacing=dp(5))
            be = Button(text="✏", on_press=lambda x, n=n, s=s, e=e: self.form_contact(n,s,e))
            bd = Button(text="❌", background_color=(0.8,0.2,0.2,1), on_press=lambda x, n=n, s=s: self.delete_contact(n,s))
            actions.add_widget(be); actions.add_widget(bd); line1.add_widget(actions)
            
            row.add_widget(line1); row.add_widget(Label(text=e, color=(0.7,0.7,0.7,1), halign="left", text_size=(dp(300),None)))
            self.contacts_list.add_widget(row)

# -----------------------------
# LOGIKA MAILINGU
# -----------------------------
    def start_mailing(self, _):
        if not self.full_data: self.popup("!", "Wczytaj najpierw arkusz płac!"); return
        threading.Thread(target=self._mailing_process, daemon=True).start()

    def _mailing_process(self):
        p = Path(self.user_data_dir) / "smtp.json"
        if not p.exists(): Clock.schedule_once(lambda d: self.popup("!", "Skonfiguruj SMTP!")); return
        cfg = json.load(open(p))
        try:
            srv = smtplib.SMTP("smtp.gmail.com", 587, timeout=15); srv.starttls(); srv.login(cfg['u'], cfg['p'])
        except Exception as e: Clock.schedule_once(lambda d: self.popup("SMTP Error", str(e))); return
        
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
                    self.apply_style_and_save(wb, tmp)
                    with open(tmp, "rb") as f: msg.add_attachment(f.read(), maintype="application", subtype="xlsx", filename=f"Raport_{row[self.idx_name]}.xlsx")
                    for ex in self.global_attachments:
                        if os.path.exists(ex):
                            with open(ex, "rb") as f: msg.add_attachment(f.read(), maintype="application", subtype="octet-stream", filename=os.path.basename(ex))
                    srv.send_message(msg); sent += 1
                except: pass
        srv.quit(); Clock.schedule_once(lambda d: self.popup("Gotowe", f"Wysłano {sent} maili."))

# -----------------------------
# POZOSTAŁE FUNKCJE UI
# -----------------------------
    def setup_home_ui(self):
        l = BoxLayout(orientation="vertical", padding=dp(30), spacing=dp(15))
        l.add_widget(Label(text="FUTURE 17.0 PLATINUM", font_size='26sp', bold=True, color=COLOR_PRIMARY))
        btn = lambda t, c: l.add_widget(PremiumButton(text=t, on_press=c))
        btn("📂 WCZYTAJ DANE PŁAC", lambda x: self.open_picker("data"))
        btn("📊 PODGLĄD I EKSPORT", lambda x: [self.refresh_table(), setattr(self.sm, 'current', 'table')] if self.full_data else self.popup("!","Brak danych"))
        btn("✉ CENTRUM MAILINGOWE", lambda x: setattr(self.sm, 'current', 'email'))
        btn("⚙ USTAWIENIA SMTP", lambda x: setattr(self.sm, 'current', 'smtp'))
        self.screens["home"].add_widget(l)

    def setup_email_ui(self):
        l = BoxLayout(orientation="vertical", padding=dp(25), spacing=dp(10))
        l.add_widget(ScreenTitle(text="CENTRUM MAILINGOWE"))
        self.lbl_estat = Label(text="Baza: 0 osób"); l.add_widget(self.lbl_estat)
        btn = lambda t, c: l.add_widget(PremiumButton(text=t, on_press=c))
        btn("📁 IMPORTUJ BAZĘ GMAIL", lambda x: self.open_picker("book"))
        btn("🔧 ZARZĄDZAJ BAZĄ", lambda x: [self.refresh_contacts(), setattr(self.sm, 'current', 'contacts')])
        btn("📝 EDYTUJ TREŚĆ", lambda x: setattr(self.sm, 'current', 'tmpl'))
        btn("📎 DODAJ ZAŁĄCZNIK", lambda x: self.open_picker("attachment"))
        btn("🚀 URUCHOM WYSYŁKĘ", self.start_mailing)
        btn("POWRÓT", lambda x: setattr(self.sm, 'current', 'home'))
        self.screens["email"].add_widget(l); self.update_email_stats()

    def show_column_popup(self, _):
        if not self.full_data: return
        box = BoxLayout(orientation="vertical", padding=dp(10), spacing=dp(5)); sc = ScrollView(); gr = GridLayout(cols=1, size_hint_y=None); gr.bind(minimum_height=gr.setter('height'))
        checks = []
        for i, h in enumerate(self.full_data[0]):
            r = BoxLayout(size_hint_y=None, height=dp(45))
            txt = str(h) if h else f"Kolumna {i+1}"
            cb = CheckBox(active=(i in self.export_indices), size_hint_x=None, width=dp(50))
            checks.append((i,cb)); r.add_widget(cb); r.add_widget(Label(text=txt, halign="left", text_size=(dp(250), None))); gr.add_widget(r)
        sc.add_widget(gr); box.add_widget(sc); box.add_widget(Button(text="OK", size_hint_y=None, height=dp(50), on_press=lambda x: [setattr(self, 'export_indices', [i for i,c in checks if c.active]), p.dismiss()]))
        p = Popup(title="Wybierz kolumny do raportu", content=box, size_hint=(0.95, 0.9)); p.open()

    def setup_smtp_ui(self):
        l = BoxLayout(orientation="vertical", padding=dp(25), spacing=dp(10))
        self.ti_su = TextInput(hint_text="Gmail", multiline=False); self.ti_sp = TextInput(hint_text="Hasło App (16 znaków)", password=True, multiline=False)
        def sv(_): [json.dump({'u':self.ti_su.text, 'p':self.ti_sp.text}, open(Path(self.user_data_dir)/"smtp.json", "w")), self.popup("OK", "Zapisano")]
        l.add_widget(ScreenTitle(text="USTAWIENIA SMTP")); l.add_widget(self.ti_su); l.add_widget(self.ti_sp)
        l.add_widget(PremiumButton(text="ZAPISZ", on_press=sv)); l.add_widget(PremiumButton(text="POWRÓT", on_press=lambda x: setattr(self.sm, 'current', 'home'), background_color=(0.4,0.4,0.4,1))); self.screens["smtp"].add_widget(l)

    def setup_tmpl_ui(self):
        l = BoxLayout(orientation="vertical", padding=dp(25), spacing=dp(10))
        self.ti_ts = TextInput(hint_text="Temat {Imię}", size_hint_y=None, height=dp(45)); self.ti_tb = TextInput(hint_text="Treść maila...", multiline=True)
        def sv(_): [self.conn.execute("INSERT OR REPLACE INTO settings VALUES (?,?)", (k,v)) for k,v in [('t_sub', self.ti_ts.text),('t_body', self.ti_tb.text)]]; self.conn.commit(); self.popup("OK", "Zapisano")
        l.add_widget(ScreenTitle(text="SZABLON MAILA")); l.add_widget(self.ti_ts); l.add_widget(self.ti_tb)
        l.add_widget(PremiumButton(text="ZAPISZ", on_press=sv)); l.add_widget(PremiumButton(text="POWRÓT", on_press=lambda x: setattr(self.sm, 'current', 'email'), background_color=(0.4,0.4,0.4,1))); self.screens["tmpl"].add_widget(l)

    def load_contacts_file(self, path):
        d, _ = self.read_excel_universal(path); h = [str(x).lower() for x in d[0]]
        iN, iS, iE = 0, 1, 2
        for i,v in enumerate(h):
            if "imi" in v: iN=i
            elif "naz" in v: iS=i
            elif "@" in v or "mail" in v: iE=i
        for r in d[1:]:
            if len(r) > iE and "@" in str(r[iE]): self.conn.execute("INSERT OR REPLACE INTO contacts VALUES (?,?,?)", (r[iN].lower(), r[iS].lower(), str(r[iE])))
        self.conn.commit(); self.update_email_stats()

    def update_email_stats(self, *args):
        c = self.conn.execute("SELECT count(*) FROM contacts").fetchone()[0]
        self.lbl_estat.text = f"Baza kontaktów: {c} osób | Załączniki: {len(self.global_attachments)}"

    def form_contact(self, n="", s="", e=""):
        b = BoxLayout(orientation="vertical", padding=dp(15), spacing=dp(10))
        iN, iS, iE = TextInput(text=n, hint_text="Imię"), TextInput(text=s, hint_text="Nazwisko"), TextInput(text=e, hint_text="Email")
        def sv(_): [self.conn.execute("INSERT OR REPLACE INTO contacts VALUES (?,?,?)", (iN.text.strip().lower(), iS.text.strip().lower(), iE.text.strip())), self.conn.commit(), p.dismiss(), self.refresh_contacts(), self.update_email_stats()]
        b.add_widget(iN); b.add_widget(iS); b.add_widget(iE); b.add_widget(Button(text="ZAPISZ", on_press=sv)); p = Popup(title="Dane kontaktu", content=b, size_hint=(0.9, 0.6)); p.open()

    def delete_contact(self, n, s):
        def pr(_): [self.conn.execute("DELETE FROM contacts WHERE name=? AND surname=?", (n, s)), self.conn.commit(), p.dismiss(), self.refresh_contacts(), self.update_email_stats()]
        p = Popup(title="Usuń?", content=Button(text="USUŃ", on_press=pr, background_color=(1,0,0,1)), size_hint=(0.7,0.3)); p.open()

    def export_row(self, r):
        f = Path("/storage/emulated/0/Documents/Raporty"); f.mkdir(parents=True, exist_ok=True)
        wb = Workbook(); ws = wb.active; ws.append([self.full_data[0][k] for k in self.export_indices]); ws.append([r[k] for k in self.export_indices])
        path = f / f"Raport_{r[self.idx_name]}.xlsx"; self.apply_style_and_save(wb, path); self.popup("OK", f"Zapisano w Documents/Raporty")

    def filter_table(self, ins, val):
        v = val.lower(); self.filtered_data = [self.full_data[0]] + [r for r in self.full_data[1:] if any(v in str(c).lower() for c in r)]; self.refresh_table()

    def popup(self, title, text):
        box = BoxLayout(orientation="vertical", padding=dp(20)); box.add_widget(Label(text=text, halign="center")); box.add_widget(Button(text="OK", on_press=lambda x: p.dismiss())); p = Popup(title=title, content=box, size_hint=(0.85, 0.45)); p.open()

if __name__ == "__main__": FutureApp().run()
