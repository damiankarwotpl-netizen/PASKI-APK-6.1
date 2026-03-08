import os
import sqlite3
import threading
import json
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
from openpyxl.styles import PatternFill, Border, Side, Alignment, Font

APP_TITLE = "Future 9.0 ULTRA PRO"

class PremiumButton(Button):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.background_normal = ""
        self.background_color = (0.2, 0.4, 0.9, 1)
        self.color = (1, 1, 1, 1)
        self.size_hint_y = None
        self.height = dp(50)

class FutureApp(App):
    def build(self):
        Window.clearcolor = (0.08, 0.1, 0.15, 1)
        self.full_data = []
        self.filtered_data = []
        self.current_file = None
        self.export_columns = []
        
        self.sm = ScreenManager()
        self.home = Screen(name="home")
        self.table = Screen(name="table")
        self.email_scr = Screen(name="email")
        self.smtp_scr = Screen(name="smtp")

        self.init_ui()
        self.init_db()

        for s in [self.home, self.table, self.email_scr, self.smtp_scr]:
            self.sm.add_widget(s)
        return self.sm

    def init_db(self):
        db_path = Path(self.user_data_dir) / "app_data.db"
        self.conn = sqlite3.connect(str(db_path))
        self.conn.execute("CREATE TABLE IF NOT EXISTS contacts (name TEXT, surname TEXT, email TEXT, PRIMARY KEY(name, surname))")
        self.conn.commit()

    # --- UI BUILDER ---
    def init_ui(self):
        # Home
        l_home = BoxLayout(orientation="vertical", padding=dp(25), spacing=dp(20))
        l_home.add_widget(Label(text=APP_TITLE, font_size=26))
        btn_open = PremiumButton(text="📂 Wczytaj plik Excel"); btn_open.bind(on_press=self.open_picker)
        btn_table = PremiumButton(text="📊 Otwórz Tabelę"); btn_table.bind(on_press=self.go_to_table)
        btn_smtp = PremiumButton(text="⚙ Ustawienia SMTP"); btn_smtp.bind(on_press=lambda x: setattr(self.sm, "current", "smtp"))
        self.home_status = Label(text="Gotowy")
        for w in [btn_open, btn_table, btn_smtp, self.home_status]: l_home.add_widget(w)
        self.home.add_widget(l_home)

        # Table
        lt = BoxLayout(orientation="vertical", padding=dp(10), spacing=dp(10))
        top = BoxLayout(size_hint=(1, 0.12), spacing=dp(8))
        self.search = TextInput(hint_text="Szukaj...", multiline=False); self.search.bind(text=self.filter_data)
        btn_mail_scr = PremiumButton(text="E-mail"); btn_mail_scr.bind(on_press=lambda x: setattr(self.sm, "current", "email"))
        btn_back = PremiumButton(text="<-"); btn_back.bind(on_press=lambda x: setattr(self.sm, "current", "home"))
        top.add_widget(self.search); top.add_widget(btn_mail_scr); top.add_widget(btn_back)
        self.scroll = ScrollView(); self.grid = GridLayout(size_hint=(None, None))
        self.grid.bind(minimum_height=self.grid.setter("height"), minimum_width=self.grid.setter("width"))
        self.scroll.add_widget(self.grid); self.progress = ProgressBar(max=100, size_hint=(1, 0.05))
        lt.add_widget(top); lt.add_widget(self.scroll); lt.add_widget(self.progress)
        self.table.add_widget(lt)

        # Email Screen
        le = BoxLayout(orientation="vertical", padding=dp(20), spacing=dp(15))
        le.add_widget(Label(text="Centrum Wysyłki", font_size=22))
        btn_imp = PremiumButton(text="📥 Importuj Książkę Adresową"); btn_imp.bind(on_press=self.import_addresses)
        btn_cols = PremiumButton(text="📋 Wybierz kolumny raportu"); btn_cols.bind(on_press=self.column_popup)
        btn_send = PremiumButton(text="🚀 Wyślij Raporty"); btn_send.bind(on_press=self.start_mailing)
        btn_back_e = PremiumButton(text="Powrót"); btn_back_e.bind(on_press=lambda x: setattr(self.sm, "current", "table"))
        self.email_status = Label(text="Mail zostanie dołączony automatycznie z bazy.")
        for w in [btn_imp, btn_cols, btn_send, btn_back_e, self.email_status]: le.add_widget(w)
        self.email_scr.add_widget(le)

        # SMTP
        ls = BoxLayout(orientation="vertical", padding=dp(20), spacing=dp(10))
        self.s_host = TextInput(hint_text="Host SMTP (np. smtp.gmail.com)"); self.s_port = TextInput(hint_text="Port (np. 587)")
        self.s_user = TextInput(hint_text="E-mail"); self.s_pass = TextInput(hint_text="Hasło (Aplikacji)", password=True)
        btn_save = PremiumButton(text="Zapisz"); btn_save.bind(on_press=self.save_smtp)
        ls.add_widget(Label(text="Konfiguracja Poczty")); ls.add_widget(self.s_host); ls.add_widget(self.s_port)
        ls.add_widget(self.s_user); ls.add_widget(self.s_pass); ls.add_widget(btn_save)
        ls.add_widget(PremiumButton(text="Powrót", on_press=lambda x: setattr(self.sm, "current", "home")))
        self.smtp_scr.add_widget(ls); self.load_smtp()

    # --- LOGIKA PLIKÓW ---
    def open_picker(self, _):
        if platform != "android": self.msg("Błąd", "Funkcja dostępna na Androidzie"); return
        from jnius import autoclass; from android import activity
        Intent = autoclass("android.content.Intent")
        intent = Intent(Intent.ACTION_OPEN_DOCUMENT); intent.setType("*/*"); intent.addCategory(Intent.CATEGORY_OPENABLE)
        def on_res(req, res, dt):
            if dt:
                try:
                    uri = dt.getData(); resolver = autoclass("org.kivy.android.PythonActivity").mActivity.getContentResolver()
                    stream = resolver.openInputStream(uri); self.current_file = Path(self.user_data_dir) / "data.xlsx"
                    with open(self.current_file, "wb") as f:
                        buf = bytearray(4096)
                        # Fix Android Stream Read
                        j_Array = autoclass('java.lang.reflect.Array'); j_Byte = autoclass('java.lang.Byte')
                        j_buf = j_Array.newInstance(j_Byte.TYPE, 4096)
                        while True:
                            r = stream.read(j_buf)
                            if r <= 0: break
                            f.write(bytes(j_buf)[:r])
                    stream.close(); Clock.schedule_once(lambda x: setattr(self.home_status, "text", "Załadowano."))
                except Exception as e: Clock.schedule_once(lambda x: self.msg("Błąd", str(e)))
            activity.unbind(on_activity_result=on_res)
        activity.bind(on_activity_result=on_res); autoclass("org.kivy.android.PythonActivity").mActivity.startActivityForResult(intent, 1001)

    def go_to_table(self, _):
        if not self.current_file: self.msg("Błąd", "Wybierz plik!"); return
        try:
            wb = load_workbook(str(self.current_file), data_only=True); ws = wb.active
            self.full_data = [["" if v is None else str(v) for v in r] for r in ws.iter_rows(values_only=True)]
            self.filtered_data = self.full_data; self.show_table(); self.sm.current = "table"
        except Exception as e: self.msg("Błąd Excel", str(e))

    def show_table(self):
        self.grid.clear_widgets()
        if not self.filtered_data: return
        r, c = len(self.filtered_data), len(self.filtered_data[0])
        w, h = dp(160), dp(42)
        self.grid.cols = c; self.grid.width, self.grid.height = c * w, r * h
        for row in self.filtered_data:
            for cell in row: self.grid.add_widget(Label(text=str(cell), size_hint=(None, None), size=(w, h)))

    def filter_data(self, ins, val):
        self.filtered_data = [r for r in self.full_data if any(val.lower() in str(c).lower() for c in r)]
        self.show_table()

    # --- KONTAKTY & WYSYŁKA ---
    def import_addresses(self, _):
        if platform != "android": return
        from jnius import autoclass; from android import activity
        Intent = autoclass("android.content.Intent"); intent = Intent(Intent.ACTION_OPEN_DOCUMENT); intent.setType("*/*")
        def on_res(req, res, dt):
            if dt:
                try:
                    uri = dt.getData(); stream = autoclass("org.kivy.android.PythonActivity").mActivity.getContentResolver().openInputStream(uri)
                    p = Path(self.user_data_dir) / "book.xlsx"
                    with open(p, "wb") as f:
                        j_buf = autoclass('java.lang.reflect.Array').newInstance(autoclass('java.lang.Byte').TYPE, 4096)
                        while True:
                            r = stream.read(j_buf)
                            if r <= 0: break
                            f.write(bytes(j_buf)[:r])
                    wb = load_workbook(str(p), data_only=True); ws = wb.active; rows = list(ws.iter_rows(values_only=True))
                    # Inteligente szukanie kolumn
                    h = [str(x).lower().strip() for x in rows[0]]
                    def find_idx(keys):
                        for i, v in enumerate(h):
                            if any(k in v for k in keys): return i
                        return None
                    ni, si, mi = find_idx(["imi"]), find_idx(["nazw"]), find_idx(["mail", "email"])
                    if mi is None: self.msg("Błąd", "Nie znaleziono kolumny Email!"); return
                    added = 0
                    for r in rows[1:]:
                        if r[mi]:
                            self.conn.execute("INSERT OR REPLACE INTO contacts VALUES(?,?,?)", (str(r[ni or 0]).lower(), str(r[si or 1]).lower(), str(r[mi])))
                            added += 1
                    self.conn.commit(); self.msg("Baza", f"Zaimportowano {added} kontaktów.")
                except Exception as e: self.msg("Błąd", str(e))
            activity.unbind(on_activity_result=on_res)
        activity.bind(on_activity_result=on_res); autoclass("org.kivy.android.PythonActivity").mActivity.startActivityForResult(intent, 1002)

    def start_mailing(self, _):
        threading.Thread(target=self.mailing_process).start()

    def mailing_process(self):
        smtp = self.get_smtp()
        if not smtp: Clock.schedule_once(lambda x: self.msg("Błąd", "Skonfiguruj SMTP!")); return
        try:
            server = smtplib.SMTP(smtp['h'], int(smtp['p'])); server.starttls(); server.login(smtp['u'], smtp['p_s'])
        except Exception as e: Clock.schedule_once(lambda x: self.msg("Błąd SMTP", str(e))); return

        head = self.full_data[0]; rows = self.full_data[1:]; sent = 0
        idxs = self.export_columns if self.export_columns else list(range(len(head)))
        
        # Inteligente szukanie imienia i nazwiska w danych źródłowych
        h_low = [str(x).lower() for x in head]
        ni = next((i for i, x in enumerate(h_low) if "imi" in x), 0)
        si = next((i for i, x in enumerate(h_low) if "nazw" in x), 1)

        for i, row in enumerate(rows):
            name, sur = str(row[ni]).lower().strip(), str(row[si]).lower().strip()
            res = self.conn.execute("SELECT email FROM contacts WHERE name=? AND surname=?", (name, sur)).fetchone()
            if res:
                email = res[0]
                try:
                    # Tworzenie pliku dla pracownika
                    wb = Workbook(); ws = wb.active
                    blue = PatternFill(start_color='CFE2F3', end_color='CFE2F3', fill_type='solid')
                    border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
                    
                    # Nagłówki
                    exp_h = [head[idx] for idx in idxs]
                    ws.append(exp_h)
                    for cell in ws[1]:
                        cell.fill = blue; cell.font = Font(bold=True); cell.border = border; cell.alignment = Alignment(horizontal='center')
                    
                    # Dane (Wiersz osoby)
                    exp_r = [row[idx] for idx in idxs]
                    ws.append(exp_r)
                    for cell in ws[2]:
                        cell.border = border; cell.alignment = Alignment(horizontal='left')
                    
                    # Auto-Size kolumn
                    for col in ws.columns:
                        max_l = max(len(str(c.value or "")) for c in col)
                        ws.column_dimensions[col[0].column_letter].width = max_l + 3

                    p_attach = Path(self.user_data_dir) / f"Raport_{name}.xlsx"; wb.save(str(p_attach))
                    
                    msg = EmailMessage(); msg["Subject"] = "Pasek Wynagrodzeń - Future 9.0"; msg["From"] = smtp['u']; msg["To"] = email
                    msg.set_content("Dzień dobry,\nW załączniku przesyłamy raport miesięczny.")
                    with open(p_attach, "rb") as f:
                        msg.add_attachment(f.read(), maintype="application", subtype="xlsx", filename=p_attach.name)
                    server.send_message(msg); sent += 1
                except: pass
            Clock.schedule_once(lambda dt, p=int((i+1)/len(rows)*100): setattr(self.progress, "value", p))
        server.quit(); Clock.schedule_once(lambda x: self.msg("Koniec", f"Wysłano {sent} maili."))

    # --- HELPERS ---
    def column_popup(self, _):
        if not self.full_data: return
        box = BoxLayout(orientation="vertical", padding=dp(10)); scroll = ScrollView(); grid = GridLayout(cols=1, size_hint_y=None); grid.bind(minimum_height=grid.setter('height'))
        checks = []
        for i, h in enumerate(self.full_data[0]):
            r = BoxLayout(size_hint_y=None, height=dp(40)); cb = CheckBox(size_hint_x=0.2); cb.active = True
            r.add_widget(cb); r.add_widget(Label(text=str(h))); grid.add_widget(r); checks.append((i, cb))
        def apply(_): self.export_columns = [idx for idx, c in checks if c.active]; p.dismiss()
        scroll.add_widget(grid); box.add_widget(scroll); btn = PremiumButton(text="Zastosuj"); btn.bind(on_press=apply); box.add_widget(btn)
        p = Popup(title="Wybierz kolumny", content=box, size_hint=(0.9, 0.9)); p.open()

    def msg(self, t, txt):
        b = BoxLayout(orientation="vertical", padding=dp(20)); b.add_widget(Label(text=txt))
        btn = PremiumButton(text="OK"); b.add_widget(btn)
        p = Popup(title=t, content=b, size_hint=(0.8, 0.4)); btn.bind(on_press=p.dismiss); p.open()

    def save_smtp(self, _):
        d = {'h': self.s_host.text, 'p': self.s_port.text, 'u': self.s_user.text, 'p_s': self.s_pass.text}
        with open(Path(self.user_data_dir) / "smtp.json", "w") as f: json.dump(d, f)
        self.msg("OK", "Ustawienia zapisane.")

    def load_smtp(self):
        p = Path(self.user_data_dir) / "smtp.json"
        if p.exists():
            with open(p) as f:
                d = json.load(f); self.s_host.text = d['h']; self.s_port.text = d['p']; self.s_user.text = d['u']; self.s_pass.text = d['p_s']

    def get_smtp(self):
        p = Path(self.user_data_dir) / "smtp.json"
        return json.load(open(p)) if p.exists() else None

if __name__ == "__main__":
    FutureApp().run()
