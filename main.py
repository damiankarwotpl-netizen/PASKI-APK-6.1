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

# Tytuł aplikacji
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
        
        self.full_data = [] # Wszystkie dane z Excela
        self.filtered_data = [] # Dane po przefiltrowaniu szukajką
        self.current_data_file = None
        self.selected_export_cols = [] # Wybrane kolumny do wysyłki
        self.db_conn = None
        
        self.sm = ScreenManager()
        self.scr_home = Screen(name="home")
        self.scr_table = Screen(name="table")
        self.scr_email = Screen(name="email")
        self.scr_smtp = Screen(name="smtp")
        
        self.init_ui_home()
        self.init_ui_table()
        self.init_ui_email()
        self.init_ui_smtp()
        
        for s in [self.scr_home, self.scr_table, self.scr_email, self.scr_smtp]:
            self.sm.add_widget(s)
            
        self.init_database()
        return self.sm

    def init_database(self):
        import sqlite3
        from pathlib import Path
        db_path = Path(self.user_data_dir) / "contacts.db"
        self.db_conn = sqlite3.connect(str(db_path))
        self.db_conn.execute("CREATE TABLE IF NOT EXISTS address_book(name TEXT, surname TEXT, email TEXT, PRIMARY KEY(name, surname))")
        self.db_conn.commit()

    # ==========================================
    # UI - GŁÓWNE EKRANY
    # ==========================================
    def init_ui_home(self):
        l = BoxLayout(orientation="vertical", padding=dp(25), spacing=dp(20))
        l.add_widget(Label(text=APP_TITLE, font_size=26))
        b1 = PremiumButton(text="📂 Wczytaj plik z DANYMI"); b1.bind(on_press=lambda x: self.pick_file(mode="data"))
        b2 = PremiumButton(text="📊 Otwórz Tabelę"); b2.bind(on_press=self.go_to_table)
        b3 = PremiumButton(text="⚙ Ustawienia SMTP Google"); b3.bind(on_press=lambda x: setattr(self.sm, "current", "smtp"))
        self.home_status = Label(text="Gotowy")
        for w in [b1, b2, b3, self.home_status]: l.add_widget(w)
        self.scr_home.add_widget(l)

    def init_ui_table(self):
        lt = BoxLayout(orientation="vertical", padding=dp(10), spacing=dp(10))
        top = BoxLayout(size_hint=(1, 0.12), spacing=dp(8))
        self.search_input = TextInput(hint_text="Szukaj...", multiline=False); self.search_input.bind(text=self.filter_data)
        b_mail = PremiumButton(text="DALEJ"); b_mail.bind(on_press=lambda x: setattr(self.sm, "current", "email"))
        b_back = PremiumButton(text="<-"); b_back.bind(on_press=lambda x: setattr(self.sm, "current", "home"))
        top.add_widget(self.search_input); top.add_widget(b_mail); top.add_widget(b_back)
        self.scroll = ScrollView(); self.grid = GridLayout(size_hint=(None, None))
        self.grid.bind(minimum_height=self.grid.setter("height"), minimum_width=self.grid.setter("width"))
        self.scroll.add_widget(self.grid); self.progress = ProgressBar(max=100, size_hint=(1, 0.05))
        lt.add_widget(top); lt.add_widget(self.scroll); lt.add_widget(self.progress)
        self.scr_table.add_widget(lt)

    def init_ui_email(self):
        le = BoxLayout(orientation="vertical", padding=dp(20), spacing=dp(15))
        le.add_widget(Label(text="Poczta & Baza Kontaktów", font_size=22))
        b_book = PremiumButton(text="📥 Importuj KONTAKTY (Imię, Nazwisko, Email)"); b_book.bind(on_press=lambda x: self.pick_file(mode="book"))
        b_cols = PremiumButton(text="📋 Wybierz kolumny do raportu"); b_cols.bind(on_press=self.show_column_selector)
        b_send = PremiumButton(text="🚀 WYŚLIJ RAPORTY"); b_send.bind(on_press=self.start_mailing)
        b_back = PremiumButton(text="Wróć"); b_back.bind(on_press=lambda x: setattr(self.sm, "current", "table"))
        self.email_status = Label(text="Pamiętaj o konfiguracji SMTP")
        for w in [b_book, b_cols, b_send, b_back, self.email_status]: le.add_widget(w)
        self.scr_email.add_widget(le)

    def init_ui_smtp(self):
        ls = BoxLayout(orientation="vertical", padding=dp(20), spacing=dp(10))
        ls.add_widget(Label(text="SMTP Google", font_size=22))
        self.s_srv = TextInput(hint_text="Serwer", text="smtp.gmail.com")
        self.s_port = TextInput(hint_text="Port", text="587")
        self.s_user = TextInput(hint_text="Twój Gmail")
        self.s_pass = TextInput(hint_text="Hasło Aplikacji (16 znaków)", password=True)
        b_test = PremiumButton(text="⚡ Testuj Połączenie"); b_test.bind(on_press=self.test_smtp)
        b_save = PremiumButton(text="✅ Zapisz"); b_save.bind(on_press=self.save_smtp)
        b_back = PremiumButton(text="<- Wróć"); b_back.bind(on_press=lambda x: setattr(self.sm, "current", "home"))
        for w in [self.s_srv, self.s_port, self.s_user, self.s_pass, b_test, b_save, b_back]: ls.add_widget(w)
        self.scr_smtp.add_widget(ls)
        self.load_smtp_from_file()

    # ==========================================
    # WYBÓR KOLUMN (POPUP)
    # ==========================================
    def show_column_selector(self, _):
        if not self.full_data: return
        heads = self.full_data[0]
        box = BoxLayout(orientation="vertical", padding=dp(10), spacing=dp(5))
        scroll = ScrollView()
        grid = GridLayout(cols=1, size_hint_y=None, spacing=dp(8))
        grid.bind(minimum_height=grid.setter('height'))
        checks = []
        for i, h in enumerate(heads):
            row = BoxLayout(size_hint_y=None, height=dp(40))
            cb = CheckBox(size_hint_x=0.2); cb.active = True
            lbl = Label(text=str(h), halign="left"); lbl.bind(size=lbl.setter('text_size'))
            row.add_widget(cb); row.add_widget(lbl); grid.add_widget(row); checks.append((i, cb))
        def apply(_):
            self.selected_export_cols = [idx for idx, cb in checks if cb.active]
            self.email_status.text = f"Wybrano {len(self.selected_export_cols)} kolumn."
            p.dismiss()
        btn = PremiumButton(text="Zatwierdź"); btn.bind(on_press=apply)
        scroll.add_widget(grid); box.add_widget(scroll); box.add_widget(btn)
        p = Popup(title="Które dane wysłać?", content=box, size_hint=(0.9, 0.8)); p.open()

    # ==========================================
    # LOGIKA PLIKÓW I EXCELA
    # ==========================================
    def pick_file(self, mode):
        if platform != "android": self.msg("Błąd", "Funkcja dostępna na Android."); return
        from jnius import autoclass
        from android import activity
        Intent = autoclass("android.content.Intent")
        intent = Intent(Intent.ACTION_OPEN_DOCUMENT); intent.setType("*/*"); intent.addCategory(Intent.CATEGORY_OPENABLE)
        def on_res(req, res, dt):
            if dt:
                from pathlib import Path
                resolver = autoclass("org.kivy.android.PythonActivity").mActivity.getContentResolver()
                stream = resolver.openInputStream(dt.getData())
                fname = "d.xlsx" if mode == "data" else "b.xlsx"
                local = Path(self.user_data_dir) / fname
                with open(local, "wb") as f:
                    while True:
                        buf = stream.read(bytearray(4096))
                        if buf == -1: break
                        f.write(buf)
                stream.close()
                if mode == "data": self.current_data_file = local; self.home_status.text = "Plik wyczytany"
                else: self.import_book(local)
            activity.unbind(on_activity_result=on_res)
        activity.bind(on_activity_result=on_res)
        autoclass("org.kivy.android.PythonActivity").mActivity.startActivityForResult(intent, 1001)

    def import_book(self, path):
        from openpyxl import load_workbook
        try:
            wb = load_workbook(str(path), data_only=True); ws = wb.active
            rows = list(ws.iter_rows(values_only=True))
            h = [str(x).lower().strip() for x in rows[0]]
            def find_i(ks):
                for i, t in enumerate(h):
                    if any(k in t for k in ks): return i
                return None
            ni, si, mi = find_i(["imi"]), find_i(["nazw", "surn"]), find_i(["mail"])
            if mi is None: self.msg("Błąd", "Nie znaleziono kolumny 'Email'!"); return
            added = 0
            for r in rows[1:]:
                mail = str(r[mi]).strip()
                if "@" in mail:
                    n, s = str(r[ni]).lower().strip() if ni is not None else "", str(r[si]).lower().strip() if si is not None else ""
                    self.db_conn.execute("INSERT OR REPLACE INTO address_book VALUES(?,?,?)", (n, s, mail))
                    added += 1
            self.db_conn.commit(); self.msg("Książka", f"Zaktualizowano {added} adresów.")
        except Exception as e: self.msg("Błąd", str(e))

    def go_to_table(self, _):
        if not self.current_data_file: return
        from openpyxl import load_workbook
        wb = load_workbook(str(self.current_data_file), data_only=True); ws = wb.active
        self.full_data = [["" if v is None else str(v) for v in row] for row in ws.iter_rows(values_only=True)]
        self.filtered_data = self.full_data; self.show_table(); self.sm.current = "table"

    def show_table(self):
        self.grid.clear_widgets()
        if not self.filtered_data: return
        rows, cols = len(self.filtered_data), len(self.filtered_data[0])
        w, h = dp(160), dp(42)
        self.grid.cols = cols; self.grid.width, self.grid.height = cols * w, rows * h
        for row in self.filtered_data:
            for cell in row: self.grid.add_widget(Label(text=str(cell), size_hint=(None, None), size=(w, h)))

    def filter_data(self, ins, val):
        self.filtered_data = [r for r in self.full_data if any(val.lower() in str(c).lower() for c in r)]
        self.show_table()

    # ==========================================
    # SMTP I WYSYŁKA (Z FORMATOWANIEM EXCELA)
    # ==========================================
    def save_smtp(self, _):
        import json; from pathlib import Path
        d = {"srv": self.s_srv.text, "port": self.s_port.text, "user": self.s_user.text, "pass": self.s_pass.text}
        with open(Path(self.user_data_dir) / "smtp.json", "w") as f: json.dump(d, f)
        self.msg("OK", "Zapisano ustawienia.")

    def load_smtp_from_file(self):
        import json; from pathlib import Path
        p = Path(self.user_data_dir) / "smtp.json"
        if p.exists():
            with open(p) as f:
                d = json.load(f); self.s_srv.text = d.get("srv", "smtp.gmail.com")
                self.s_port.text = d.get("port", "587"); self.s_user.text = d.get("user", ""); self.s_pass.text = d.get("pass", "")

    def test_smtp(self, _):
        import threading
        def _test():
            import smtplib
            try:
                s = smtplib.SMTP(self.s_srv.text, int(self.s_port.text), timeout=10)
                s.starttls(); s.login(self.s_user.text, self.s_pass.text); s.quit()
                Clock.schedule_once(lambda dt: self.msg("Test OK", "Połączono!"))
            except Exception as e: Clock.schedule_once(lambda dt: self.msg("Błąd SMTP", str(e)))
        threading.Thread(target=_test).start()

    def start_mailing(self, _):
        import threading
        threading.Thread(target=self.mailing_thread).start()

    def mailing_thread(self):
        import smtplib; from email.message import EmailMessage; from openpyxl import Workbook; from pathlib import Path
        from openpyxl.styles import PatternFill, Border, Side, Alignment, Font

        try:
            srv = smtplib.SMTP(self.s_srv.text, int(self.s_port.text))
            srv.starttls(); srv.login(self.s_user.text, self.s_pass.text)
        except Exception as e: Clock.schedule_once(lambda dt: self.msg("Błąd", str(e))); return

        header = self.full_data[0]
        rows = self.full_data[1:]; sent = 0
        idxs = self.selected_export_cols if self.selected_export_cols else list(range(len(header)))
        
        for i, row in enumerate(rows):
            n, s = str(row[0]).lower().strip(), str(row[1]).lower().strip()
            res = self.db_conn.execute("SELECT email FROM address_book WHERE name=? AND surname=?", (n, s)).fetchone()
            
            if res:
                email = res[0]
                try:
                    msg = EmailMessage(); msg["Subject"] = "Twój Raport"; msg["From"] = self.s_user.text; msg["To"] = email
                    msg.set_content(f"Dzień dobry,\nPrzesyłamy raport dla {n} {s}.")
                    
                    # TWORZENIE PROFESJONALNEGO EXCELA
                    wb = Workbook(); ws = wb.active
                    
                    # 1.Style
                    fill = PatternFill(start_color='FFCFE2F3', end_color='FFCFE2F3', fill_type='solid')
                    thin = Side(border_style="thin", color="000000")
                    border = Border(top=thin, left=thin, right=thin, bottom=thin)
                    
                    # 2.Nagłówki
                    exp_h = [header[idx] for idx in idxs]
                    ws.append(exp_h)
                    for cell in ws[1]:
                        cell.fill = fill
                        cell.font = Font(bold=True)
                        cell.border = border
                        cell.alignment = Alignment(horizontal="center")
                    
                    # 3.Dane
                    exp_r = [row[idx] for idx in idxs]
                    ws.append(exp_r)
                    for cell in ws[2]:
                        cell.border = border
                        cell.alignment = Alignment(horizontal="left")

                    # 4.AUTO-DOPASOWANIE KOLUMN
                    for col in ws.columns:
                        max_len = 0
                        col_letter = col[0].column_letter
                        for cell in col:
                            if cell.value: max_len = max(max_len, len(str(cell.value)))
                        ws.column_dimensions[col_letter].width = max_len + 3

                    tmp = Path(self.user_data_dir) / "report.xlsx"; wb.save(str(tmp))
                    with open(tmp, "rb") as f:
                        msg.add_attachment(f.read(), maintype="application", subtype="xlsx", filename=f"Raport_{n}.xlsx")
                    srv.send_message(msg); sent += 1
                except: pass
            Clock.schedule_once(lambda dt, p=int((i+1)/len(rows)*100): setattr(self.progress, "value", p))

        srv.quit()
        Clock.schedule_once(lambda dt: self.msg("Koniec", f"Wysłano {sent} maili."))

    def msg(self, title, text):
        box = BoxLayout(orientation="vertical", padding=dp(20)); box.add_widget(Label(text=text))
        btn = PremiumButton(text="OK"); box.add_widget(btn)
        p = Popup(title=title, content=box, size_hint=(0.8, 0.4)); btn.bind(on_press=p.dismiss); p.open()

if __name__ == "__main__":
    FutureApp().run()
