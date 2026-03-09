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
APP_TITLE = "Future 12.0 ULTRA PRO "

class HomeScreen(Screen): pass
class TableScreen(Screen): pass
class EmailScreen(Screen): pass
class SMTPScreen(Screen): pass
class TemplateScreen(Screen): pass
class ContactsMgrScreen(Screen): pass # Nowy ekran bazy

class PremiumButton(Button):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.background_normal = ""
        self.background_color = (0.2, 0.4, 0.9, 1)
        self.color = (1, 1, 1, 1)
        self.size_hint_y = None
        self.height = dp(50)
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
        db_p = Path(self.user_data_dir) / "app_v12_contacts.db"
        self.conn = sqlite3.connect(str(db_p), check_same_thread=False)
        self.conn.execute("CREATE TABLE IF NOT EXISTS contacts (name TEXT, surname TEXT, email TEXT, PRIMARY KEY(name, surname))")
        self.conn.execute("CREATE TABLE IF NOT EXISTS settings (key TEXT PRIMARY KEY, val TEXT)")
        self.conn.execute("CREATE TABLE IF NOT EXISTS logs (id INTEGER PRIMARY KEY, msg TEXT, date TEXT)")
        self.conn.commit()

# -----------------------------
# ZARZĄDZANIE BAZĄ (NOWE)
# -----------------------------
    def build_contacts_mgr(self):
        layout = BoxLayout(orientation="vertical", padding=dp(10), spacing=dp(10))
        top = BoxLayout(size_hint_y=0.12, spacing=5)
        self.contact_search = TextInput(hint_text="Szukaj osoby...", multiline=False)
        self.contact_search.bind(text=self.refresh_contacts_list)
        btn_add = Button(text="+ Dodaj", size_hint_x=0.3); btn_add.bind(on_press=lambda x: self.open_contact_form())
        btn_back = Button(text="Wróć", size_hint_x=0.2); btn_back.bind(on_press=lambda x: setattr(self.sm, "current", "email"))
        top.add_widget(self.contact_search); top.add_widget(btn_add); top.add_widget(btn_back)

        self.contacts_scroll = ScrollView(); self.contacts_grid = GridLayout(cols=4, size_hint_y=None, spacing=2)
        self.contacts_grid.bind(minimum_height=self.contacts_grid.setter('height'))
        self.contacts_scroll.add_widget(self.contacts_grid)
        layout.add_widget(top); layout.add_widget(self.contacts_scroll)
        self.contacts_mgr.add_widget(layout)

    def refresh_contacts_list(self, *args):
        self.contacts_grid.clear_widgets()
        search_val = self.contact_search.text.lower()
        query = "SELECT name, surname, email FROM contacts ORDER BY surname ASC"
        data = self.conn.execute(query).fetchall()
        
        w_btn = dp(60)
        h = dp(45)
        # Nagłówki
        for h_txt in ["Dane", "Email", "Edytuj", "Usuń"]:
            self.contacts_grid.add_widget(Label(text=h_txt, bold=True, size_hint_y=None, height=h))

        for n, s, e in data:
            full_display = f"{n.capitalize()} {s.capitalize()}"
            if search_val and search_val not in full_display.lower() and search_val not in e.lower(): continue
            
            self.contacts_grid.add_widget(Label(text=full_display, size_hint_y=None, height=h))
            self.contacts_grid.add_widget(Label(text=str(e), size_hint_y=None, height=h))
            
            # Edytuj
            btn_e = Button(text="✏", size_hint=(None, None), size=(w_btn, h))
            btn_e.bind(on_press=lambda x, cn=n, cs=s, ce=e: self.open_contact_form(cn, cs, ce))
            # Usuń
            btn_d = Button(text="❌", size_hint=(None, None), size=(w_btn, h), background_color=(0.9, 0.2, 0.2, 1))
            btn_d.bind(on_press=lambda x, cn=n, cs=s: self.delete_contact(cn, cs))
            
            self.contacts_grid.add_widget(btn_e); self.contacts_grid.add_widget(btn_d)

    def open_contact_form(self, ename="", esurname="", eemail=""):
        box = BoxLayout(orientation="vertical", padding=15, spacing=10)
        ti_n = TextInput(text=ename, hint_text="Imię", multiline=False)
        ti_s = TextInput(text=esurname, hint_text="Nazwisko", multiline=False)
        ti_e = TextInput(text=eemail, hint_text="Adres Email", multiline=False)
        
        if ename: ti_n.readonly = True; ti_s.readonly = True # Klucz podstawowy nie do zmiany przy edycji maila

        def save_form(_):
            n, s, e = ti_n.text.strip().lower(), ti_s.text.strip().lower(), ti_e.text.strip()
            if not n or not s or "@" not in e: 
                self.popup("Błąd", "Wypełnij poprawnie wszystkie pola!"); return
            self.conn.execute("INSERT OR REPLACE INTO contacts VALUES (?,?,?)", (n, s, e))
            self.conn.commit(); p.dismiss(); self.refresh_contacts_list(); self.update_email_ui_labels()

        box.add_widget(ti_n); box.add_widget(ti_s); box.add_widget(ti_e)
        btn = PremiumButton(text="Zapisz Kontakt"); btn.bind(on_press=save_form); box.add_widget(btn)
        p = Popup(title="Dane Kontaktu", content=box, size_hint=(0.85, 0.6)); p.open()

    def delete_contact(self, n, s):
        def proceed(_):
            self.conn.execute("DELETE FROM contacts WHERE name=? AND surname=?", (n, s))
            self.conn.commit(); p.dismiss(); self.refresh_contacts_list(); self.update_email_ui_labels()
        btn = Button(text="TAK, USUŃ", background_color=(1,0,0,1)); btn.bind(on_press=proceed)
        p = Popup(title="Czy na pewno?", content=btn, size_hint=(0.6, 0.3)); p.open()

# -----------------------------
# TRZON APLIKACJI (Open, Load, SMTP)
# -----------------------------
    def read_any_excel(self, path):
        path_str = str(path).lower()
        data = []
        try:
            if path_str.endswith(".xlsx"):
                wb = load_workbook(path, data_only=True); ws = wb.active
                data = [["" if v is None else str(v) for v in row] for row in ws.iter_rows(values_only=True)]
            elif path_str.endswith(".xls") and xlrd:
                wb = xlrd.open_workbook(path); ws = wb.sheet_by_index(0)
                for r in range(ws.nrows): data.append([str(ws.cell_value(r, c)).strip() for c in range(ws.ncols)])
            return data, None
        except Exception as e: return None, str(e)

    def find_name_surname_indices(self, header_row):
        h = [str(x).lower().strip() for x in header_row]
        idx_n, idx_s, idx_e = 0, 1, -1
        for i, val in enumerate(h):
            if any(k in val for k in ["imię", "imie", "name"]): idx_n = i; break
        for i, val in enumerate(h):
            if any(k in val for k in ["nazwisko", "surname"]): idx_s = i; break
        for i, val in enumerate(h):
            if "@" in val or "email" in val: idx_e = i; break
        return idx_n, idx_s, idx_e

    def apply_excel_styling(self, ws):
        for col in ws.columns:
            max_len = 0
            for cell in col:
                if cell.value: max_len = max(max_len, len(str(cell.value)))
            ws.column_dimensions[col[0].column_letter].width = max_len + 5
        thick = Side(style='thick'); thin = Side(style='thin')
        for r in range(1, ws.max_row+1):
            for c in range(1, ws.max_column+1):
                cell = ws.cell(row=r, column=c)
                if r == 1: cell.font = Font(bold=True)
                cell.border = Border(left=thick if c==1 else thin, right=thick if c==ws.max_column else thin, top=thick if r==1 else thin, bottom=thick if r==ws.max_row else thin)
                cell.alignment = Alignment(horizontal='center')

    def build_home(self):
        l = BoxLayout(orientation="vertical", padding=dp(25), spacing=dp(15))
        l.add_widget(Label(text=APP_TITLE, font_size=26, bold=True))
        def btn(t, c): b = PremiumButton(text=t); b.bind(on_press=c); l.add_widget(b)
        btn("📂 Wczytaj Plik Płac (.xlsx / .xls)", lambda x: self.open_picker(mode="data"))
        btn("📊 Podgląd i Kolumny", self.load_excel)
        btn("✉ Centrum Mailingu", lambda x: setattr(self.sm, "current", "email"))
        btn("⚙ Ustawienia SMTP", lambda x: setattr(self.sm, "current", "smtp"))
        self.status = Label(text="Gotowy"); l.add_widget(self.status); self.home.add_widget(l)

    def open_picker(self, mode="data"):
        if platform != "android": self.popup("Błąd", "Funkcja tylko na Android"); return
        from jnius import autoclass; from android import activity
        PythonActivity = autoclass("org.kivy.android.PythonActivity"); Intent = autoclass("android.content.Intent")
        intent = Intent(Intent.ACTION_GET_CONTENT); intent.setType("*/*")
        def callback(req, res, intent_data):
            if not intent_data: return
            activity.unbind(on_activity_result=callback)
            uri = intent_data.getData(); stream = PythonActivity.mActivity.getContentResolver().openInputStream(uri)
            local = Path(self.user_data_dir) / f"tmp_{mode}_{os.urandom(2).hex()}.xlsx"
            with open(local, "wb") as f:
                buf = bytearray(4096)
                while True:
                    r = stream.read(buf); 
                    if r == -1: break
                    f.write(buf[:r])
            stream.close()
            if mode == "data": self.current_file = local; Clock.schedule_once(lambda d: setattr(self.status, "text", "Wczytano Płace"))
            elif mode == "book": self.import_contacts_to_db(local)
            elif mode == "attachment": self.global_attachments.append(str(local)); Clock.schedule_once(lambda d: self.update_email_ui_labels())
        activity.bind(on_activity_result=callback); PythonActivity.mActivity.startActivityForResult(intent, 1001)

    def load_excel(self, _):
        if not self.current_file: self.popup("!", "Wczytaj najpierw plik!"); return
        data, err = self.read_any_excel(self.current_file)
        if err: self.popup("Błąd", err); return
        self.full_data = data; self.filtered_data = self.full_data
        self.idx_name, self.idx_surname, _ = self.find_name_surname_indices(self.full_data[0])
        self.export_col_indices = list(range(len(self.full_data[0])))
        self.show_table(); self.sm.current = "table"

# -----------------------------
# TABELA
# -----------------------------
    def build_table(self):
        l = BoxLayout(orientation="vertical", padding=dp(10), spacing=dp(10))
        top = BoxLayout(size_hint_y=0.12, spacing=5)
        self.search = TextInput(hint_text="Szukaj..."); self.search.bind(text=self.filter_data)
        bc = Button(text="Wybierz Kolumny", size_hint_x=0.4); bc.bind(on_press=self.open_column_selector)
        bb = Button(text="Wróć", size_hint_x=0.25); bb.bind(on_press=lambda x: setattr(self.sm, "current", "home"))
        top.add_widget(self.search); top.add_widget(bc); top.add_widget(bb)
        self.scroll = ScrollView(); self.grid = GridLayout(size_hint=(None, None))
        self.grid.bind(minimum_height=self.grid.setter("height"), minimum_width=self.grid.setter("width"))
        self.scroll.add_widget(self.grid); self.progress = ProgressBar(max=100, size_hint_y=0.05)
        l.add_widget(top); l.add_widget(self.scroll); l.add_widget(self.progress); self.table.add_widget(l)

    def open_column_selector(self, _):
        box = BoxLayout(orientation="vertical", padding=10, spacing=5); sc = ScrollView(); gr = GridLayout(cols=1, size_hint_y=None); gr.bind(minimum_height=gr.setter('height'))
        checks = []
        for i, h in enumerate(self.full_data[0]):
            r = BoxLayout(size_hint_y=None, height=dp(40)); cb = CheckBox(active=(i in self.export_col_indices), size_hint_x=0.2)
            checks.append((i, cb)); r.add_widget(cb); r.add_widget(Label(text=str(h))); gr.add_widget(r)
        def save(_): self.export_col_indices = [i for i, c in checks if c.active]; p.dismiss()
        sc.add_widget(gr); box.add_widget(sc); btn = PremiumButton(text="Ok"); btn.bind(on_press=save); box.add_widget(btn)
        p = Popup(title="Widoczność danych", content=box, size_hint=(0.9, 0.9)); p.open()

    def show_table(self):
        self.grid.clear_widgets()
        if not self.filtered_data: return
        rws, cls = len(self.filtered_data), len(self.filtered_data[0]); w, h = dp(160), dp(42)
        self.grid.cols = cls+1; self.grid.width, self.grid.height = (cls+1)*w, rws*h
        for head in self.filtered_data[0]: self.grid.add_widget(Label(text=str(head), bold=True, size=(w,h), size_hint=(None,None)))
        self.grid.add_widget(Label(text="Akcja", bold=True, size=(w,h), size_hint=(None,None)))
        for r in self.filtered_data[1:]:
            for c in r: self.grid.add_widget(Label(text=str(c), size=(w,h), size_hint=(None,None)))
            b = Button(text="Zapisz", size=(w,h), size_hint=(None,None)); b.bind(on_press=lambda x, row=r: self.single_export(row)); self.grid.add_widget(b)

    def filter_data(self, ins, val):
        v = val.lower(); self.filtered_data = [self.full_data[0]] + [r for r in self.full_data[1:] if any(v in str(c).lower() for c in r)]; self.show_table()

# -----------------------------
# MAILING
# -----------------------------
    def build_email(self):
        l = BoxLayout(orientation="vertical", padding=dp(20), spacing=dp(10))
        l.add_widget(Label(text="Centrum Mailingowe", font_size=22, bold=True))
        self.email_info = Label(text="Baza: 0"); self.att_info = Label(text="Załączniki: 0")
        l.add_widget(self.email_info); l.add_widget(self.att_info)
        def btn(t, c): b = PremiumButton(text=t); b.bind(on_press=c); l.add_widget(b)
        btn("📁 Wczytaj Bazę z Excela", lambda x: self.open_picker(mode="book"))
        btn("🔧 ZARZĄDZAJ BAZĄ (Dodaj/Usuń)", lambda x: [setattr(self.sm, "current", "contacts"), self.refresh_contacts_list()])
        btn("📝 Treść Wiadomości", lambda x: setattr(self.sm, "current", "tmpl"))
        btn("📎 Dodaj PDF", lambda x: self.open_picker(mode="attachment"))
        btn("🚀 URUCHOM MAILING", self.send_emails_start)
        btn("Powrót", lambda x: setattr(self.sm, "current", "home"))
        self.email.add_widget(l); self.update_email_ui_labels()

    def update_email_ui_labels(self, *args):
        cnt = self.conn.execute("SELECT count(*) FROM contacts").fetchone()[0]
        self.email_info.text = f"Kontakty w bazie: {cnt}"; self.att_info.text = f"Załączniki: {len(self.global_attachments)}"

    def import_contacts_to_db(self, path):
        data, err = self.read_any_excel(path)
        if err: self.popup("Błąd", err); return
        idx_n, idx_s, idx_e = self.find_name_surname_indices(data[0])
        if idx_e == -1: idx_e = 2
        for r in data[1:]:
            if r and idx_e < len(r) and r[idx_e]:
                self.conn.execute("INSERT OR REPLACE INTO contacts VALUES (?,?,?)", (str(r[idx_n]).strip().lower(), str(r[idx_s]).strip().lower(), str(r[idx_e]).strip()))
        self.conn.commit(); Clock.schedule_once(lambda d: self.popup("Ok", "Baza zaktualizowana")); self.update_email_ui_labels()

    def send_emails_start(self, _): threading.Thread(target=self._mail_process, daemon=True).start()

    def _mail_process(self):
        p = Path(self.user_data_dir) / "smtp.json"
        if not p.exists(): Clock.schedule_once(lambda d: self.popup("!", "Ustaw SMTP!")); return
        cfg = json.load(open(p))
        try: srv = smtplib.SMTP("smtp.gmail.com", 587, timeout=12); srv.starttls(); srv.login(cfg['u'], cfg['p'])
        except Exception as e: Clock.schedule_once(lambda d: self.popup("Błąd", str(e))); return
        
        rows = self.full_data[1:]; sent = 0
        for i, row in enumerate(rows):
            n, s = str(row[self.idx_name]).strip().lower(), str(row[self.idx_surname]).strip().lower()
            res = self.conn.execute("SELECT email FROM contacts WHERE name=? AND surname=?", (n, s)).fetchone()
            if res:
                try:
                    msg = EmailMessage(); dat = datetime.now().strftime("%d.%m.%Y")
                    ts = self.conn.execute("SELECT val FROM settings WHERE key='t_sub'").fetchone()
                    tb = self.conn.execute("SELECT val FROM settings WHERE key='t_body'").fetchone()
                    msg["Subject"], msg["From"], msg["To"] = (ts[0] if ts else "Raport").replace("{Imię}", str(row[self.idx_name])), cfg['u'], res[0]
                    msg.set_content((tb[0] if tb else "Informacja").replace("{Imię}", str(row[self.idx_name])).replace("{Data}", dat))
                    tmp = Path(self.user_data_dir) / f"r_{i}.xlsx"; wb = Workbook(); ws = wb.active
                    ws.append([self.full_data[0][k] for k in self.export_col_indices])
                    ws.append([row[k] for k in self.export_col_indices]); self.apply_excel_styling(ws); wb.save(tmp)
                    with open(tmp, "rb") as f: msg.add_attachment(f.read(), maintype="application", subtype="xlsx", filename=f"Raport_{row[self.idx_name]}.xlsx")
                    for ex in self.global_attachments:
                        if os.path.exists(ex):
                            with open(ex, "rb") as f: msg.add_attachment(f.read(), maintype="application", subtype="octet-stream", filename=os.path.basename(ex))
                    srv.send_message(msg); sent += 1
                except: pass
            Clock.schedule_once(lambda dt, p=int((i+1)/len(rows)*100): setattr(self.progress, "value", p))
        srv.quit(); final = f"Wysłano: {sent}"; self.popup("Koniec", final)

# -----------------------------
# POZOSTAŁE (SMTP, Szablon, Eksport)
# -----------------------------
    def build_tmpl(self):
        l = BoxLayout(orientation="vertical", padding=dp(20), spacing=dp(10)); self.ts = TextInput(hint_text="Temat {Imię}"); self.tb = TextInput(hint_text="Treść", multiline=True)
        rs = self.conn.execute("SELECT val FROM settings WHERE key='t_sub'").fetchone()
        rb = self.conn.execute("SELECT val FROM settings WHERE key='t_body'").fetchone()
        if rs: self.ts.text, self.tb.text = rs[0], rb[0]
        def save(_): [self.conn.execute("INSERT OR REPLACE INTO settings VALUES (?,?)", (k, v)) for k, v in [('t_sub', self.ts.text), ('t_body', self.tb.text)]]; self.conn.commit(); self.popup("OK", "Zapisano")
        b = PremiumButton(text="Zapisz"); b.bind(on_press=save); l.add_widget(self.ts); l.add_widget(self.tb); l.add_widget(b); self.tmpl.add_widget(l)

    def build_smtp(self):
        l = BoxLayout(orientation="vertical", padding=dp(20), spacing=dp(10)); self.su = TextInput(hint_text="Gmail"); self.sp = TextInput(hint_text="Hasło API", password=True)
        p = Path(self.user_data_dir) / "smtp.json"
        if p.exists(): d = json.load(open(p)); self.su.text, self.sp.text = d.get('u',''), d.get('p','')
        def save(_): json.dump({'u':self.su.text, 'p':self.sp.text}, open(p, "w")); self.popup("OK", "Zapisano")
        b = PremiumButton(text="Zapisz"); b.bind(on_press=save); l.add_widget(self.su); l.add_widget(self.sp); l.add_widget(b); self.smtp.add_widget(l)

    def single_export(self, row):
        f = Path("/storage/emulated/0/Documents/FutureExport"); f.mkdir(parents=True, exist_ok=True); wb = Workbook(); ws = wb.active
        ws.append([self.full_data[0][k] for k in self.export_col_indices]); ws.append([row[k] for k in self.export_col_indices])
        self.apply_excel_styling(ws); wb.save(f / f"Raport_{row[self.idx_name]}.xlsx"); self.popup("OK", "Zapisano w Documents.")

    def popup(self, title, text):
        box = BoxLayout(orientation="vertical", padding=20)
        box.add_widget(Label(text=text, halign="center"))
        b = PremiumButton(text="OK"); b.bind(on_press=lambda x: p.dismiss()); box.add_widget(b)
        p = Popup(title=title, content=box, size_hint=(0.8, 0.45)); p.open()

if __name__ == "__main__": FutureApp().run()
