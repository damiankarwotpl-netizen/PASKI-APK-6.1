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

# Obsługa formatu .xls
try:
    import xlrd
except ImportError:
    xlrd = None

# --- KONFIGURACJA ---
APP_TITLE = "Future 13.0 ULTRA FIX "

class HomeScreen(Screen): pass
class TableScreen(Screen): pass
class EmailScreen(Screen): pass
class SMTPScreen(Screen): pass
class TemplateScreen(Screen): pass
class ContactsMgrScreen(Screen): pass 

class PremiumButton(Button):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.background_normal = ""
        self.background_color = (0.2, 0.4, 0.9, 1)
        self.color = (1, 1, 1, 1)
        self.size_hint_y = None
        self.height = dp(55) # Nieco wyższy przycisk dla wygody
        self.bold = True

class FutureApp(App):
    def build(self):
        Window.clearcolor = (0.08, 0.1, 0.15, 1)
        self.full_data = []
        self.filtered_data = []
        self.current_file = None
        self.global_attachments = []
        self.export_col_indices = []
        self.idx_name = 0
        self.idx_surname = 1
        
        self.init_db()
        self.sm = ScreenManager()
        self.home = HomeScreen(name="home")
        self.table = TableScreen(name="table")
        self.email = EmailScreen(name="email")
        self.smtp = SMTPScreen(name="smtp")
        self.tmpl = TemplateScreen(name="tmpl")
        self.contacts_mgr = ContactsMgrScreen(name="contacts")

        self.build_home(); self.build_table(); self.build_email(); self.build_smtp(); self.build_tmpl(); self.build_contacts_mgr()
        for s in [self.home, self.table, self.email, self.smtp, self.tmpl, self.contacts_mgr]:
            self.sm.add_widget(s)
        return self.sm

    def init_db(self):
        db_p = Path(self.user_data_dir) / "app_v13_final.db"
        self.conn = sqlite3.connect(str(db_p), check_same_thread=False)
        self.conn.execute("CREATE TABLE IF NOT EXISTS contacts (name TEXT, surname TEXT, email TEXT, PRIMARY KEY(name, surname))")
        self.conn.execute("CREATE TABLE IF NOT EXISTS settings (key TEXT PRIMARY KEY, val TEXT)")
        self.conn.execute("CREATE TABLE IF NOT EXISTS logs (id INTEGER PRIMARY KEY, msg TEXT, date TEXT)")
        self.conn.commit()

# -----------------------------
# UNIWERSALNY ODCZYT I FIX XLS
# -----------------------------
    def read_any_excel(self, path):
        path_str = str(path).lower()
        data = []
        try:
            if path_str.endswith(".xlsx"):
                wb = load_workbook(path, data_only=True); ws = wb.active
                data = [["" if v is None else str(v).strip() for v in row] for row in ws.iter_rows(values_only=True)]
            elif path_str.endswith(".xls"):
                if not xlrd: return None, "Brak xlrd!"
                wb = xlrd.open_workbook(path); ws = wb.sheet_by_index(0)
                for r in range(ws.nrows):
                    row_vals = []
                    for c in range(ws.ncols):
                        val = ws.cell_value(r, c)
                        if isinstance(val, float) and val.is_integer(): val = int(val)
                        row_vals.append(str(val).strip())
                    data.append(row_vals)
            return data, None
        except Exception as e: return None, str(e)

# -----------------------------
# DYNAMICZNA TABELA (NAPRAWA OVERLAPU)
# -----------------------------
    def build_table(self):
        l = BoxLayout(orientation="vertical", padding=dp(10), spacing=dp(10))
        top = BoxLayout(size_hint_y=None, height=dp(55), spacing=5)
        self.search = TextInput(hint_text="Szukaj...", multiline=False)
        self.search.bind(text=self.filter_data)
        bc = Button(text="Wybierz Kolumny", size_hint_x=0.4); bc.bind(on_press=self.open_column_selector)
        bb = Button(text="Powrót", size_hint_x=0.25); bb.bind(on_press=lambda x: setattr(self.sm, "current", "home"))
        top.add_widget(self.search); top.add_widget(bc); top.add_widget(bb)
        
        self.scroll = ScrollView(do_scroll_x=True, do_scroll_y=True)
        self.grid = GridLayout(size_hint=(None, None), spacing=dp(2))
        self.grid.bind(minimum_height=self.grid.setter("height"), minimum_width=self.grid.setter("width"))
        self.scroll.add_widget(self.grid)
        
        self.progress = ProgressBar(max=100, size_hint_y=None, height=dp(15))
        l.add_widget(top); l.add_widget(self.scroll); l.add_widget(self.progress)
        self.table.add_widget(l)

    def show_table(self):
        self.grid.clear_widgets()
        if not self.filtered_data: return
        
        rws = len(self.filtered_data)
        cls = len(self.filtered_data[0])
        w_cell, h_cell = dp(180), dp(50) # Zwiększona szerokość, by dane nie nachodziły na siebie
        
        self.grid.cols = cls + 1
        self.grid.width = (cls + 1) * (w_cell + dp(2))
        self.grid.height = rws * (h_cell + dp(2))

        # Nagłówki
        for head in self.filtered_data[0]:
            self.grid.add_widget(Label(text=str(head), bold=True, size_hint=(None, None), size=(w_cell, h_cell),
                                     color=(0.2, 0.6, 1, 1), halign="center", valign="middle", text_size=(w_cell, None)))
        self.grid.add_widget(Label(text="Akcja", bold=True, size_hint=(None, None), size=(w_cell, h_cell)))

        # Wiersze
        for idx, r in enumerate(self.filtered_data[1:]):
            for c in r:
                self.grid.add_widget(Label(text=str(c), size_hint=(None, None), size=(w_cell, h_cell),
                                         halign="center", valign="middle", text_size=(w_cell-dp(10), None)))
            btn = Button(text="Zapisz", size_hint=(None, None), size=(w_cell, h_cell))
            btn.bind(on_press=lambda x, row=r: self.single_export(row))
            self.grid.add_widget(btn)

# -----------------------------
# WYBÓR KOLUMNY (NAPRAWA XLS)
# -----------------------------
    def open_column_selector(self, _):
        if not self.full_data: return
        box = BoxLayout(orientation="vertical", padding=20, spacing=10)
        box.add_widget(Label(text="Zaznacz dane do eksportu", bold=True, size_hint_y=None, height=dp(30)))
        
        sc = ScrollView(); gr = GridLayout(cols=1, size_hint_y=None, spacing=5); gr.bind(minimum_height=gr.setter('height'))
        checks = []
        # Używamy self.full_data[0] - nagłówki są pobierane bezpośrednio z arkusza
        for i, h in enumerate(self.full_data[0]):
            r = BoxLayout(size_hint_y=None, height=dp(45))
            txt = str(h) if h and str(h).strip() != "" else f"Kolumna {i+1}" # Fix dla brakujących opisów
            cb = CheckBox(active=(i in self.export_col_indices), size_hint_x=None, width=dp(50))
            checks.append((i, cb)); r.add_widget(cb); r.add_widget(Label(text=txt, halign="left", text_size=(Window.width*0.6, None)))
            gr.add_widget(r)
        
        def save(_): self.export_col_indices = [i for i, c in checks if c.active]; p.dismiss()
        sc.add_widget(gr); box.add_widget(sc)
        btn = PremiumButton(text="Zatwierdź Wybór"); btn.bind(on_press=save); box.add_widget(btn)
        p = Popup(title="Konfiguracja Raportu", content=box, size_hint=(0.95, 0.9)); p.open()

# -----------------------------
# ZARZĄDZANIE KONTAKTAMI
# -----------------------------
    def build_contacts_mgr(self):
        ly = BoxLayout(orientation="vertical", padding=10, spacing=10)
        tp = BoxLayout(size_hint_y=None, height=dp(55), spacing=5)
        self.c_search = TextInput(hint_text="Szukaj w bazie...", multiline=False)
        self.c_search.bind(text=self.refresh_contacts_list)
        ba = Button(text="+ Dodaj", size_hint_x=0.3); ba.bind(on_press=lambda x: self.open_contact_form())
        bb = Button(text="Wróć", size_hint_x=0.2); bb.bind(on_press=lambda x: setattr(self.sm, "current", "email"))
        tp.add_widget(self.c_search); tp.add_widget(ba); tp.add_widget(bb)
        self.c_scroll = ScrollView(); self.c_grid = GridLayout(cols=4, size_hint=(1, None), spacing=2)
        self.c_grid.bind(minimum_height=self.c_grid.setter('height')); self.c_scroll.add_widget(self.c_grid)
        ly.add_widget(tp); ly.add_widget(self.c_scroll); self.contacts_mgr.add_widget(ly)

    def refresh_contacts_list(self, *args):
        self.c_grid.clear_widgets()
        sv = self.c_search.text.lower()
        rows = self.conn.execute("SELECT name, surname, email FROM contacts ORDER BY surname ASC").fetchall()
        h = dp(50)
        for ht in ["Dane", "Email", "Edytuj", "Usuń"]:
            self.c_grid.add_widget(Label(text=ht, bold=True, size_hint_y=None, height=h, color=(0.2,0.6,1,1)))
        for n, s, e in rows:
            disp = f"{n} {s}"
            if sv and sv not in disp.lower() and sv not in e.lower(): continue
            self.c_grid.add_widget(Label(text=disp, size_hint_y=None, height=h))
            self.c_grid.add_widget(Label(text=e, size_hint_y=None, height=h))
            be = Button(text="✏", size_hint=(None, None), size=(dp(50), h)); be.bind(on_press=lambda x, n=n, s=s, e=e: self.open_contact_form(n,s,e))
            bd = Button(text="❌", size_hint=(None, None), size=(dp(50), h), background_color=(0.8,0.2,0.2,1)); bd.bind(on_press=lambda x, n=n, s=s: self.delete_contact(n,s))
            self.c_grid.add_widget(be); self.c_grid.add_widget(bd)

# -----------------------------
# ŚCIEŻKA KRYTYCZNA (ŁADOWANIE I MAILING)
# -----------------------------
    def build_home(self):
        l = BoxLayout(orientation="vertical", padding=dp(30), spacing=dp(20))
        l.add_widget(Label(text=APP_TITLE, font_size=28, bold=True, color=(0.2, 0.6, 1, 1)))
        def btn(t, c): b = PremiumButton(text=t); b.bind(on_press=c); l.add_widget(b)
        btn("📂 Wczytaj Płace (.xlsx / .xls)", lambda x: self.open_picker(mode="data"))
        btn("📊 Podgląd / Wybór Kolumn", self.load_excel)
        btn("✉ Centrum Mailingu", lambda x: setattr(self.sm, "current", "email"))
        btn("⚙ Ustawienia SMTP", lambda x: setattr(self.sm, "current", "smtp"))
        self.status = Label(text="Oczekiwanie na dane..."); l.add_widget(self.status); self.home.add_widget(l)

    def load_excel(self, _):
        if not self.current_file: self.popup("!", "Wybierz plik!"); return
        data, err = self.read_any_excel(self.current_file)
        if err: self.popup("Błąd", err); return
        self.full_data = data; self.filtered_data = data
        header = [str(x).lower().strip() for x in data[0]]
        self.idx_name = 0; self.idx_surname = 1
        for i, v in enumerate(header):
            if any(k in v for k in ["imię", "imie", "name"]): self.idx_name = i; break
        for i, v in enumerate(header):
            if any(k in v for k in ["nazwisko", "surname"]) and i != self.idx_name: self.idx_surname = i; break
        self.export_col_indices = list(range(len(data[0])))
        self.show_table(); self.sm.current = "table"

    def build_email(self):
        l = BoxLayout(orientation="vertical", padding=dp(20), spacing=dp(10))
        l.add_widget(Label(text="Centrum Mailingowe", font_size=24, bold=True))
        self.e_info = Label(text=""); self.a_info = Label(text="")
        l.add_widget(self.e_info); l.add_widget(self.a_info)
        def btn(t, c): b = PremiumButton(text=t); b.bind(on_press=c); l.add_widget(b)
        btn("📁 Wczytaj Bazę z Excel", lambda x: self.open_picker(mode="book"))
        btn("🔧 Zarządzaj Kontaktami", lambda x: [setattr(self.sm, "current", "contacts"), self.refresh_contacts_list()])
        btn("📝 Treść Maila", lambda x: setattr(self.sm, "current", "tmpl"))
        btn("📎 Dodaj Załącznik", lambda x: self.open_picker(mode="attachment"))
        btn("🚀 URUCHOM WYSYŁKĘ", self.send_emails_start)
        btn("Cofnij", lambda x: setattr(self.sm, "current", "home"))
        self.email.add_widget(l); self.update_email_ui_labels()

    def update_email_ui_labels(self, *args):
        c = self.conn.execute("SELECT count(*) FROM contacts").fetchone()[0]
        self.e_info.text = f"Osob w bazie: {c}"; self.a_info.text = f"Załączniki: {len(self.global_attachments)}"

    def send_emails_start(self, _):
        if not self.full_data: self.popup("!", "Wczytaj płace!"); return
        threading.Thread(target=self._mail_process, daemon=True).start()

    def _mail_process(self):
        p = Path(self.user_data_dir) / "smtp.json"
        if not p.exists(): Clock.schedule_once(lambda d: self.popup("!", "Skonfiguruj SMTP!")); return
        cfg = json.load(open(p))
        try: srv = smtplib.SMTP("smtp.gmail.com", 587, timeout=15); srv.starttls(); srv.login(cfg['u'], cfg['p'])
        except Exception as e: Clock.schedule_once(lambda d: self.popup("Błąd", str(e))); return
        
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
                    msg.set_content((tb[0] if tb else "Cześć").replace("{Imię}", row[self.idx_name]).replace("{Data}", dat))
                    
                    tmp = Path(self.user_data_dir) / f"out_{i}.xlsx"; wb = Workbook(); ws = wb.active
                    ws.append([self.full_data[0][k] for k in self.export_col_indices])
                    ws.append([row[k] for k in self.export_col_indices]); self.apply_excel_styling(ws); wb.save(tmp)
                    with open(tmp, "rb") as f: msg.add_attachment(f.read(), maintype="application", subtype="xlsx", filename=f"Raport_{row[self.idx_name]}.xlsx")
                    for ex in self.global_attachments:
                        if os.path.exists(ex):
                            with open(ex, "rb") as f: msg.add_attachment(f.read(), maintype="application", subtype="octet-stream", filename=os.path.basename(ex))
                    srv.send_message(msg); sent += 1
                except: pass
            Clock.schedule_once(lambda dt, p=int((i+1)/len(rows)*100): setattr(self.progress, "value", p))
        srv.quit(); Clock.schedule_once(lambda d: self.popup("Gotowe", f"Wysłano {sent} maili."))

# -----------------------------
# POMOCNICZE (SMTP, FORM, POPUP)
# -----------------------------
    def open_picker(self, mode="data"):
        if platform != "android": self.popup("Info", "Tylko Android"); return
        from jnius import autoclass; from android import activity
        PythonActivity = autoclass("org.kivy.android.PythonActivity"); Intent = autoclass("android.content.Intent")
        intent = Intent(Intent.ACTION_GET_CONTENT); intent.setType("*/*")
        def cb(req, res, intent_data):
            if not intent_data: return
            activity.unbind(on_activity_result=cb)
            uri = intent_data.getData(); stream = PythonActivity.mActivity.getContentResolver().openInputStream(uri)
            local = Path(self.user_data_dir) / f"f_{mode}_{os.urandom(2).hex()}.xlsx"
            with open(local, "wb") as f:
                buf = bytearray(4096)
                while True:
                    r = stream.read(buf); 
                    if r == -1: break
                    f.write(buf[:r])
            stream.close()
            if mode == "data": self.current_file = local; Clock.schedule_once(lambda d: setattr(self.status, "text", "Załadowano Nowy Arkusz"))
            elif mode == "book": self.import_contacts_to_db(local)
            elif mode == "attachment": self.global_attachments.append(str(local)); Clock.schedule_once(lambda d: self.update_email_ui_labels())
        activity.bind(on_activity_result=cb); PythonActivity.mActivity.startActivityForResult(intent, 1001)

    def import_contacts_to_db(self, path):
        d, e = self.read_any_excel(path)
        if e: self.popup("!", e); return
        h = [str(x).lower().strip() for x in d[0]]
        in_n, in_s, in_e = 0, 1, 2
        for i,v in enumerate(h):
            if "imi" in v: in_n=i
            elif "naz" in v: in_s=i
            elif "@" in v or "mail" in v: in_e=i
        for r in d[1:]:
            if len(r) > max(in_n, in_s, in_e) and r[in_e]:
                self.conn.execute("INSERT OR REPLACE INTO contacts VALUES (?,?,?)", (r[in_n].lower(), r[in_s].lower(), r[in_e]))
        self.conn.commit(); self.update_email_ui_labels()

    def open_contact_form(self, n="", s="", e=""):
        b = BoxLayout(orientation="vertical", padding=15, spacing=10)
        in_n = TextInput(text=n, hint_text="Imię", multiline=False)
        in_s = TextInput(text=s, hint_text="Nazwisko", multiline=False)
        in_e = TextInput(text=e, hint_text="E-mail", multiline=False)
        if n: in_n.readonly = True; in_s.readonly = True
        def sv(_):
            self.conn.execute("INSERT OR REPLACE INTO contacts VALUES (?,?,?)", (in_n.text.strip().lower(), in_s.text.strip().lower(), in_e.text.strip()))
            self.conn.commit(); p.dismiss(); self.refresh_contacts_list(); self.update_email_ui_labels()
        b.add_widget(in_n); b.add_widget(in_s); b.add_widget(in_e)
        bt = PremiumButton(text="Zapisz"); bt.bind(on_press=sv); b.add_widget(bt)
        p = Popup(title="Dane Kontaktu", content=b, size_hint=(0.9, 0.6)); p.open()

    def delete_contact(self, n, s):
        def pr(_): self.conn.execute("DELETE FROM contacts WHERE name=? AND surname=?", (n, s)); self.conn.commit(); p.dismiss(); self.refresh_contacts_list(); self.update_email_ui_labels()
        btn = Button(text="POTWIERDŹ USUNIĘCIE", background_color=(1,0,0,1)); btn.bind(on_press=pr); p = Popup(title="Usuwanie", content=btn, size_hint=(0.7,0.3)); p.open()

    def build_tmpl(self):
        l = BoxLayout(orientation="vertical", padding=20, spacing=10); self.ts = TextInput(hint_text="Temat {Imię}"); self.tb = TextInput(hint_text="Treść", multiline=True)
        r = self.conn.execute("SELECT val FROM settings WHERE key='t_sub'").fetchone(); rb = self.conn.execute("SELECT val FROM settings WHERE key='t_body'").fetchone()
        if r: self.ts.text, self.tb.text = r[0], rb[0]
        def sv(_): self.conn.execute("INSERT OR REPLACE INTO settings VALUES (?,?)", ('t_sub', self.ts.text)); self.conn.execute("INSERT OR REPLACE INTO settings VALUES (?,?)", ('t_body', self.tb.text)); self.conn.commit(); self.popup("OK", "Zapisano")
        b = PremiumButton(text="Zapisz"); b.bind(on_press=sv); l.add_widget(self.ts); l.add_widget(self.tb); l.add_widget(b); self.tmpl.add_widget(l)

    def build_smtp(self):
        l = BoxLayout(orientation="vertical", padding=20, spacing=10); self.su = TextInput(hint_text="Gmail"); self.sp = TextInput(hint_text="Hasło App", password=True)
        p = Path(self.user_data_dir) / "smtp.json"
        if p.exists(): d = json.load(open(p)); self.su.text, self.sp.text = d['u'], d['p']
        def sv(_): json.dump({'u':self.su.text, 'p':self.sp.text}, open(p, "w")); self.popup("OK", "Zapisano")
        b = PremiumButton(text="Zapisz"); b.bind(on_press=sv); l.add_widget(self.su); l.add_widget(self.sp); l.add_widget(b); self.smtp.add_widget(l)

    def single_export(self, row):
        f = Path("/storage/emulated/0/Documents/FutureExport"); f.mkdir(parents=True, exist_ok=True); wb = Workbook(); ws = wb.active
        ws.append([self.full_data[0][k] for k in self.export_col_indices]); ws.append([row[k] for k in self.export_col_indices])
        self.apply_excel_styling(ws); wb.save(f / f"Raport_{row[self.idx_name]}.xlsx"); self.popup("OK", "Zapisano w Documents.")

    def apply_excel_styling(self, ws):
        for col in ws.columns:
            m = 0
            for cell in col:
                if cell.value: m = max(m, len(str(cell.value)))
            ws.column_dimensions[col[0].column_letter].width = m + 5
        t = Side(style='thick'); th = Side(style='thin')
        for r in range(1, ws.max_row+1):
            for c in range(1, ws.max_column+1):
                cell = ws.cell(row=r, column=c)
                if r == 1: cell.font = Font(bold=True)
                cell.border = Border(left=th, right=th, top=th, bottom=th); cell.alignment = Alignment(horizontal='center')

    def filter_data(self, ins, val):
        v = val.lower(); self.filtered_data = [self.full_data[0]] + [r for r in self.full_data[1:] if any(v in str(c).lower() for c in r)]; self.show_table()

    def popup(self, title, text):
        box = BoxLayout(orientation="vertical", padding=20); box.add_widget(Label(text=text, halign="center"))
        b = PremiumButton(text="OK"); b.bind(on_press=lambda x: p.dismiss()); box.add_widget(b); p = Popup(title=title, content=box, size_hint=(0.8, 0.4)); p.open()

if __name__ == "__main__": FutureApp().run()
