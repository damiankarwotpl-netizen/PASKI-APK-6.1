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

# --- KONFIGURACJA ---
APP_TITLE = "Future 11.0 ULTRA PRO"

class HomeScreen(Screen): pass
class TableScreen(Screen): pass
class EmailScreen(Screen): pass
class SMTPScreen(Screen): pass
class TemplateScreen(Screen): pass

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
        
        # Indeksy kolumn (Płace)
        self.idx_name = 0
        self.idx_surname = 1
        
        self.init_db()
        self.sm = ScreenManager()
        self.home = HomeScreen(name="home")
        self.table = TableScreen(name="table")
        self.email = EmailScreen(name="email")
        self.smtp = SMTPScreen(name="smtp")
        self.tmpl = TemplateScreen(name="tmpl")

        self.build_home(); self.build_table(); self.build_email(); self.build_smtp(); self.build_tmpl()
        for s in [self.home, self.table, self.email, self.smtp, self.tmpl]:
            self.sm.add_widget(s)
        return self.sm

    def init_db(self):
        db_p = Path(self.user_data_dir) / "app_v11_fix.db"
        self.conn = sqlite3.connect(str(db_p), check_same_thread=False)
        self.conn.execute("CREATE TABLE IF NOT EXISTS contacts (name TEXT, surname TEXT, email TEXT, PRIMARY KEY(name, surname))")
        self.conn.execute("CREATE TABLE IF NOT EXISTS settings (key TEXT PRIMARY KEY, val TEXT)")
        self.conn.execute("CREATE TABLE IF NOT EXISTS logs (id INTEGER PRIMARY KEY, msg TEXT, date TEXT)")
        self.conn.commit()

# -----------------------------
# INTELIGENTNE WYKRYWANIE KOLUMN
# -----------------------------
    def find_name_surname_indices(self, header_row):
        h = [str(x).lower().strip() for x in header_row]
        idx_n, idx_s, idx_e = 0, 1, -1
        
        # Szukaj Imię
        for i, val in enumerate(h):
            if any(k in val for k in ["imię", "imie", "first name", "name"]):
                idx_n = i
                break
        # Szukaj Nazwisko
        for i, val in enumerate(h):
            if any(k in val for k in ["nazwisko", "surname", "last name"]) and i != idx_n:
                idx_s = i
                break
        # Szukaj Email (dla bazy kontaktów)
        for i, val in enumerate(h):
            if "@" in val or "email" in val or "mail" in val:
                idx_e = i
                break
                
        return idx_n, idx_s, idx_e

# -----------------------------
# STYLIZACJA EXCEL
# -----------------------------
    def apply_excel_styling(self, ws):
        for col in ws.columns:
            max_len = 0
            for cell in col:
                if cell.value: max_len = max(max_len, len(str(cell.value)))
            ws.column_dimensions[col[0].column_letter].width = max_len + 5
        
        thick = Side(style='thick'); thin = Side(style='thin')
        for r in range(1, ws.max_row + 1):
            for c in range(1, ws.max_column + 1):
                cell = ws.cell(row=r, column=c)
                if r == 1: cell.font = Font(bold=True)
                cell.border = Border(left=thick if c==1 else thin, right=thick if c==ws.max_column else thin,
                                     top=thick if r==1 else thin, bottom=thick if r==ws.max_row else thin)
                cell.alignment = Alignment(horizontal='center')

# -----------------------------
# LOGIKA HOME & PICKER
# -----------------------------
    def build_home(self):
        l = BoxLayout(orientation="vertical", padding=dp(25), spacing=dp(15))
        l.add_widget(Label(text=APP_TITLE, font_size=26, bold=True))
        def btn(t, c): b = PremiumButton(text=t); b.bind(on_press=c); l.add_widget(b)
        btn("📂 Otwórz Excel Płac (Test.xlsx)", lambda x: self.open_picker(mode="data"))
        btn("📊 Podgląd Tabeli", self.load_excel)
        btn("✉ Centrum Mailingu", lambda x: setattr(self.sm, "current", "email"))
        btn("⚙ Ustawienia SMTP (Gmail)", lambda x: setattr(self.sm, "current", "smtp"))
        self.status = Label(text="Zacznij od pliku płac"); l.add_widget(self.status)
        self.home.add_widget(l)

    def open_picker(self, mode="data"):
        if platform != "android": self.popup("Błąd", "Funkcja tylko na Android"); return
        from jnius import autoclass; from android import activity
        PythonActivity = autoclass("org.kivy.android.PythonActivity")
        Intent = autoclass("android.content.Intent")
        intent = Intent(Intent.ACTION_GET_CONTENT); intent.setType("*/*")
        
        def callback(req, res, intent_data):
            if not intent_data: return
            activity.unbind(on_activity_result=callback)
            uri = intent_data.getData(); resolver = PythonActivity.mActivity.getContentResolver()
            stream = resolver.openInputStream(uri)
            suffix = os.urandom(2).hex()
            local = Path(self.user_data_dir) / f"tmp_{mode}_{suffix}.xlsx"
            with open(local, "wb") as f:
                buf = bytearray(4096)
                while True:
                    r = stream.read(buf)
                    if r == -1: break
                    f.write(buf[:r])
            stream.close()
            if mode == "data": 
                self.current_file = local
                Clock.schedule_once(lambda d: setattr(self.status, "text", "Plik Płac załadowany"))
            elif mode == "book": 
                self.import_contacts_to_db(local)
            elif mode == "attachment": 
                self.global_attachments.append(str(local))
                Clock.schedule_once(lambda d: self.update_email_ui_labels())

        activity.bind(on_activity_result=callback)
        PythonActivity.mActivity.startActivityForResult(intent, 1001)

    def load_excel(self, _):
        if not self.current_file: self.popup("!", "Wczytaj najpierw Excel!"); return
        wb = load_workbook(self.current_file, data_only=True); ws = wb.active
        self.full_data = [["" if v is None else str(v) for v in row] for row in ws.iter_rows(values_only=True)]
        self.filtered_data = self.full_data
        
        # Wykryj kolumny w Płacach
        self.idx_name, self.idx_surname, _ = self.find_name_surname_indices(self.full_data[0])
        self.popup("Wykryto kolumny", f"Imię: {self.full_data[0][self.idx_name]}\nNazwisko: {self.full_data[0][self.idx_surname]}")
        
        self.show_table(); self.sm.current = "table"

# -----------------------------
# TABELA
# -----------------------------
    def build_table(self):
        l = BoxLayout(orientation="vertical", padding=dp(10), spacing=dp(10))
        top = BoxLayout(size_hint_y=0.12, spacing=5)
        self.search = TextInput(hint_text="Szukaj..."); self.search.bind(text=self.filter_data)
        b_back = PremiumButton(text="Wróć"); b_back.bind(on_press=lambda x: setattr(self.sm, "current", "home"))
        top.add_widget(self.search); top.add_widget(b_back)
        self.scroll = ScrollView(); self.grid = GridLayout(size_hint=(None, None))
        self.grid.bind(minimum_height=self.grid.setter("height"), minimum_width=self.grid.setter("width"))
        self.scroll.add_widget(self.grid); self.progress = ProgressBar(max=100, size_hint_y=0.05)
        l.add_widget(top); l.add_widget(self.scroll); l.add_widget(self.progress); self.table.add_widget(l)

    def show_table(self):
        self.grid.clear_widgets()
        if not self.filtered_data: return
        rows, cols = len(self.filtered_data), len(self.filtered_data[0])
        w, h = dp(160), dp(42)
        self.grid.cols = cols + 1; self.grid.width, self.grid.height = (cols+1)*w, rows*h
        for head in self.filtered_data[0]: self.grid.add_widget(Label(text=str(head), size_hint=(None, None), size=(w,h), bold=True))
        self.grid.add_widget(Label(text="Akcja", size_hint=(None, None), size=(w,h), bold=True))
        for r in self.filtered_data[1:]:
            for c in r: self.grid.add_widget(Label(text=str(c), size_hint=(None, None), size=(w,h)))
            b = Button(text="Zapisz", size_hint=(None, None), size=(w,h)); b.bind(on_press=lambda x, row=r: self.single_export(row))
            self.grid.add_widget(b)

    def filter_data(self, ins, val):
        v = val.lower(); self.filtered_data = [self.full_data[0]] + [r for r in self.full_data[1:] if any(v in str(c).lower() for c in r)]
        self.show_table()

# -----------------------------
# MAILING
# -----------------------------
    def build_email(self):
        l = BoxLayout(orientation="vertical", padding=dp(20), spacing=dp(10))
        l.add_widget(Label(text="Centrum Mailingowe", font_size=22, bold=True))
        self.email_info = Label(text="Baza: 0 osób"); self.att_info = Label(text="Załączniki: 0")
        l.add_widget(self.email_info); l.add_widget(self.att_info)
        def btn(size_t, t, c): b = PremiumButton(text=t); b.bind(on_press=c); l.add_widget(b)
        btn(1, "📁 Wczytaj Bazę GMAIL (Baza.xlsx)", lambda x: self.open_picker(mode="book"))
        btn(1, "📝 Ustaw Treść Maila", lambda x: setattr(self.sm, "current", "tmpl"))
        btn(1, "📎 Dodaj PDF / Obraz", lambda x: self.open_picker(mode="attachment"))
        btn(1, "⚡ Test Wysyłki (Do siebie)", self.send_test_email)
        btn(1, "🚀 URUCHOM MAILING MASOWY", self.send_emails_start)
        btn(1, "📜 Historia / Błędy", self.show_history)
        btn(1, "Powrót", lambda x: setattr(self.sm, "current", "home"))
        self.email.add_widget(l); self.update_email_ui_labels()

    def update_email_ui_labels(self):
        count = self.conn.execute("SELECT count(*) FROM contacts").fetchone()[0]
        self.email_info.text = f"Baza kontaktów: {count} osób"; self.att_info.text = f"Dodatkowe załączniki: {len(self.global_attachments)}"

    def import_contacts_to_db(self, path):
        try:
            wb = load_workbook(path, data_only=True); ws = wb.active; count = 0
            rows = list(ws.iter_rows(values_only=True))
            
            # Wykryj kolumny w pliku Baza.xlsx
            idx_n, idx_s, idx_e = self.find_name_surname_indices(rows[0])
            if idx_e == -1: idx_e = 2 # Zabezpieczenie na 3 kolumnę
            
            for r in rows[1:]:
                if r and idx_e < len(r) and r[idx_e]:
                    name = str(r[idx_n]).strip().lower()
                    surname = str(r[idx_s]).strip().lower()
                    email = str(r[idx_e]).strip()
                    self.conn.execute("INSERT OR REPLACE INTO contacts VALUES (?,?,?)", (name, surname, email))
                    count += 1
            self.conn.commit()
            Clock.schedule_once(lambda d: self.popup("Baza", f"Wgrano pomyślnie {count} osób."))
            Clock.schedule_once(lambda d: self.update_email_ui_labels())
        except Exception as e: Clock.schedule_once(lambda d: self.popup("Błąd Bazy", str(e)))

    def send_emails_start(self, _):
        if not self.full_data: self.popup("!", "Wczytaj płace!"); return
        threading.Thread(target=self._mail_process, args=(False,), daemon=True).start()

    def send_test_email(self, _):
        if not self.full_data: self.popup("!", "Wczytaj płace!"); return
        threading.Thread(target=self._mail_process, args=(True,), daemon=True).start()

    def _mail_process(self, is_test):
        p = Path(self.user_data_dir) / "smtp.json"
        if not p.exists(): Clock.schedule_once(lambda d: self.popup("!", "Ustaw SMTP!")); return
        cfg = json.load(open(p))
        try:
            srv = smtplib.SMTP("smtp.gmail.com", 587, timeout=12); srv.starttls(); srv.login(cfg['u'], cfg['p'])
        except Exception as e: Clock.schedule_once(lambda d: self.popup("Błąd SMTP", str(e))); return

        rows = self.full_data[1:2] if is_test else self.full_data[1:]
        sent, err_list = 0, []
        
        for i, row in enumerate(rows):
            target = cfg['u'] if is_test else ""
            if not is_test:
                # Szukaj używając wykrytych indeksów z pliku Płac
                name = str(row[self.idx_name]).strip().lower()
                surname = str(row[self.idx_surname]).strip().lower()
                res = self.conn.execute("SELECT email FROM contacts WHERE name=? AND surname=?", (name, surname)).fetchone()
                if res: target = res[0]
                else: err_list.append(f"Brak w bazie GMAIL: {name} {surname}")

            if target:
                try:
                    msg = EmailMessage(); dat = datetime.now().strftime("%d.%m.%Y")
                    # Szablon
                    ts = self.conn.execute("SELECT val FROM settings WHERE key='t_sub'").fetchone()
                    tb = self.conn.execute("SELECT val FROM settings WHERE key='t_body'").fetchone()
                    sub = (ts[0] if ts else "Raport").replace("{Imię}", str(row[self.idx_name]))
                    bod = (tb[0] if tb else "Cześć").replace("{Imię}", str(row[self.idx_name])).replace("{Data}", dat)
                    
                    msg["Subject"], msg["From"], msg["To"] = sub, cfg['u'], target
                    msg.set_content(bod)

                    # Załącznik Raportu
                    tmp = Path(self.user_data_dir) / f"temp_r_{i}.xlsx"
                    wb = Workbook(); ws = wb.active; ws.append(self.full_data[0]); ws.append(row)
                    self.apply_excel_styling(ws); wb.save(tmp)
                    with open(tmp, "rb") as f: msg.add_attachment(f.read(), maintype="application", subtype="xlsx", filename=f"Raport_{row[self.idx_name]}.xlsx")

                    # Extra załączniki
                    for ex in self.global_attachments:
                        if os.path.exists(ex):
                            with open(ex, "rb") as f: msg.add_attachment(f.read(), maintype="application", subtype="octet-stream", filename=os.path.basename(ex))
                    
                    srv.send_message(msg); sent += 1
                except: pass
            
            Clock.schedule_once(lambda dt, p=int((i+1)/len(rows)*100): setattr(self.progress, "value", p))

        srv.quit()
        if err_list:
            for e in err_list: self.conn.execute("INSERT INTO logs (msg, date) VALUES (?,?)", (e, datetime.now().strftime("%H:%M")))
            self.conn.commit()
        Clock.schedule_once(lambda d: self.popup("Mailing", f"Wysłano: {sent}.\nBłędy dopasowania: {len(err_list)}"))

# -----------------------------
# USTAWIENIA I HISTORIA
# -----------------------------
    def build_tmpl(self):
        l = BoxLayout(orientation="vertical", padding=dp(20), spacing=dp(10))
        self.ts = TextInput(hint_text="Temat (użyj {Imię})"); self.tb = TextInput(hint_text="Treść maila", multiline=True)
        rs = self.conn.execute("SELECT val FROM settings WHERE key='t_sub'").fetchone()
        rb = self.conn.execute("SELECT val FROM settings WHERE key='t_body'").fetchone()
        if rs: self.ts.text, self.tb.text = rs[0], rb[0]
        def save(_):
            self.conn.execute("INSERT OR REPLACE INTO settings VALUES ('t_sub', ?)", (self.ts.text,))
            self.conn.execute("INSERT OR REPLACE INTO settings VALUES ('t_body', ?)", (self.tb.text,))
            self.conn.commit(); self.popup("OK", "Szablon zapisany.")
        b1 = PremiumButton(text="Zapisz"); b1.bind(on_press=save)
        b2 = PremiumButton(text="Cofnij"); b2.bind(on_press=lambda x: setattr(self.sm, "current", "email"))
        l.add_widget(self.ts); l.add_widget(self.tb); l.add_widget(b1); l.add_widget(b2); self.tmpl.add_widget(l)

    def build_smtp(self):
        l = BoxLayout(orientation="vertical", padding=dp(20), spacing=dp(10))
        self.su = TextInput(hint_text="Email Gmail"); self.sp = TextInput(hint_text="Hasło Aplikacji (16 znaków)", password=True)
        p = Path(self.user_data_dir) / "smtp.json"
        if p.exists(): d = json.load(open(p)); self.su.text, self.sp.text = d.get('u',''), d.get('p','')
        def save(_):
            with open(p, "w") as f: json.dump({'u':self.su.text, 'p':self.sp.text}, f)
            self.popup("OK", "SMTP zapisane.")
        b1 = PremiumButton(text="Zapisz"); b1.bind(on_press=save)
        b2 = PremiumButton(text="Cofnij"); b2.bind(on_press=lambda x: setattr(self.sm, "current", "home"))
        l.add_widget(self.su); l.add_widget(self.sp); l.add_widget(b1); l.add_widget(b2); self.smtp.add_widget(l)

    def single_export(self, row):
        f = Path("/storage/emulated/0/Documents/FutureExport"); f.mkdir(parents=True, exist_ok=True)
        wb = Workbook(); ws = wb.active; ws.append(self.full_data[0]); ws.append(row); self.apply_excel_styling(ws)
        fname = f"{row[self.idx_name]}_{row[self.idx_surname]}"
        wb.save(f / f"Raport_{fname}.xlsx"); self.popup("OK", f"Zapisano {fname}")

    def show_history(self, _):
        logs = self.conn.execute("SELECT msg, date FROM logs ORDER BY id DESC LIMIT 20").fetchall()
        txt = "\n".join([f"[{d}] {m}" for m, d in logs])
        self.popup("Historia / Błędy", txt if txt else "Brak błędów wysyłki.")

    def popup(self, title, text):
        box = BoxLayout(orientation="vertical", padding=20)
        box.add_widget(Label(text=text, halign="center"))
        b = PremiumButton(text="OK"); b.bind(on_press=lambda x: p.dismiss()); box.add_widget(b)
        p = Popup(title=title, content=box, size_hint=(0.85, 0.45)); p.open()

if __name__ == "__main__":
    FutureApp().run()
