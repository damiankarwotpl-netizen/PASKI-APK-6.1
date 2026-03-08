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

# Styl przycisku
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
        self.current_data_file = None
        self.selected_export_cols = []
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
        import sqlite3; from pathlib import Path
        db_path = Path(self.user_data_dir) / "contacts.db"
        self.db_conn = sqlite3.connect(str(db_path))
        self.db_conn.execute("CREATE TABLE IF NOT EXISTS address_book(name TEXT, surname TEXT, email TEXT, PRIMARY KEY(name, surname))")
        self.db_conn.commit()

    # ================= UI =================
    def init_ui_home(self):
        l = BoxLayout(orientation="vertical", padding=dp(25), spacing=dp(20))
        l.add_widget(Label(text="FUTURE 9.0 ULTRA PRO", font_size=26))
        b1 = PremiumButton(text="📂 Wczytaj plik z danymi (Płace)"); b1.bind(on_press=lambda x: self.pick_file(mode="data"))
        b2 = PremiumButton(text="📊 Otwórz Tabelę"); b2.bind(on_press=self.go_to_table)
        b3 = PremiumButton(text="⚙ Ustawienia SMTP (Gmail)"); b3.bind(on_press=lambda x: setattr(self.sm, "current", "smtp"))
        self.home_status = Label(text="Witaj! Wgraj dane, aby zacząć.")
        for w in [b1, b2, b3, self.home_status]: l.add_widget(w)
        self.scr_home.add_widget(l)

    def init_ui_table(self):
        lt = BoxLayout(orientation="vertical", padding=dp(10), spacing=dp(10))
        top = BoxLayout(size_hint=(1, 0.12), spacing=dp(8))
        self.search = TextInput(hint_text="Szukaj osoby...", multiline=False); self.search.bind(text=self.filter_data)
        b_next = PremiumButton(text="DALEJ"); b_next.bind(on_press=lambda x: setattr(self.sm, "current", "email"))
        b_back = PremiumButton(text="Wróć"); b_back.bind(on_press=lambda x: setattr(self.sm, "current", "home"))
        top.add_widget(self.search); top.add_widget(b_next); top.add_widget(b_back)
        self.scroll = ScrollView(); self.grid = GridLayout(size_hint=(None, None))
        self.grid.bind(minimum_height=self.grid.setter("height"), minimum_width=self.grid.setter("width"))
        self.scroll.add_widget(self.grid); self.progress = ProgressBar(max=100, size_hint=(1, 0.05))
        lt.add_widget(top); lt.add_widget(self.scroll); lt.add_widget(self.progress)
        self.scr_table.add_widget(lt)

    def init_ui_email(self):
        le = BoxLayout(orientation="vertical", padding=dp(20), spacing=dp(15))
        le.add_widget(Label(text="Poczta i Baza Adresowa", font_size=22))
        b_book = PremiumButton(text="📥 Importuj Książkę Adresową (Excel)"); b_book.bind(on_press=lambda x: self.pick_file(mode="book"))
        b_cols = PremiumButton(text="📋 Wybierz kolumny do raportu"); b_cols.bind(on_press=self.show_column_selector)
        b_send = PremiumButton(text="🚀 WYŚLIJ RAPORTY DO KAŻDEGO"); b_send.bind(on_press=self.start_mailing)
        b_back = PremiumButton(text="Powrót"); b_back.bind(on_press=lambda x: setattr(self.sm, "current", "table"))
        self.email_status = Label(text="Automatyczne dopasowanie maila po Imieniu i Nazwisku.")
        for w in [b_book, b_cols, b_send, b_back, self.email_status]: le.add_widget(w)
        self.scr_email.add_widget(le)

    def init_ui_smtp(self):
        ls = BoxLayout(orientation="vertical", padding=dp(20), spacing=dp(10))
        ls.add_widget(Label(text="Konfiguracja Gmail", font_size=22))
        self.s_user = TextInput(hint_text="Twój Gmail")
        self.s_pass = TextInput(hint_text="Hasło Aplikacji (16 znaków)", password=True)
        b_save = PremiumButton(text="✅ Zapisz i Testuj"); b_save.bind(on_press=self.save_smtp)
        b_back = PremiumButton(text="Wróć"); b_back.bind(on_press=lambda x: setattr(self.sm, "current", "home"))
        for w in [self.s_user, self.s_pass, b_save, b_back]: ls.add_widget(w)
        self.scr_smtp.add_widget(ls); self.load_smtp()

    # ================= LOGIKA =================
    def pick_file(self, mode):
        if platform != "android": self.msg("Błąd", "Picker działa tylko na Androidzie."); return
        from jnius import autoclass; from android import activity
        Intent = autoclass("android.content.Intent")
        intent = Intent(Intent.ACTION_OPEN_DOCUMENT); intent.setType("*/*"); intent.addCategory(Intent.CATEGORY_OPENABLE)
        def on_res(req, res, dt):
            if dt:
                try:
                    from pathlib import Path
                    resolver = autoclass("org.kivy.android.PythonActivity").mActivity.getContentResolver()
                    stream = resolver.openInputStream(dt.getData())
                    fname = "data.xlsx" if mode == "data" else "book.xlsx"
                    local = Path(self.user_data_dir) / fname
                    with open(local, "wb") as f:
                        j_buf = autoclass('java.lang.reflect.Array').newInstance(autoclass('java.lang.Byte').TYPE, 4096)
                        while True:
                            r = stream.read(j_buf)
                            if r <= 0: break
                            f.write(bytes(j_buf)[:r])
                    stream.close()
                    if mode == "data": self.current_data_file = local; self.home_status.text = "Plik danych gotowy."
                    else: self.import_book_excel(local)
                except Exception as e: self.msg("Błąd", str(e))
            activity.unbind(on_activity_result=on_res)
        activity.bind(on_activity_result=on_res)
        autoclass("org.kivy.android.PythonActivity").mActivity.startActivityForResult(intent, 1001)

    def import_book_excel(self, path):
        from openpyxl import load_workbook
        try:
            wb = load_workbook(str(path), data_only=True); ws = wb.active
            rows = list(ws.iter_rows(values_only=True))
            h = [str(x).lower().strip() for x in rows[0]]
            def find_idx(keys):
                for i, t in enumerate(h):
                    if any(k in t for k in keys): return i
                return None
            ni, si, mi = find_idx(["imi"]), find_idx(["nazw"]), find_idx(["mail", "email"])
            if mi is None: self.msg("Błąd", "Nie znaleziono kolumny Email!"); return
            for r in rows[1:]:
                email = str(r[mi]).strip()
                if "@" in email:
                    n, s = str(r[ni or 0]).lower().strip(), str(r[si or 1]).lower().strip()
                    self.db_conn.execute("INSERT OR REPLACE INTO address_book VALUES(?,?,?)", (n, s, email))
            self.db_conn.commit(); self.msg("OK", "Adresy zapisane w bazie.")
        except Exception as e: self.msg("Błąd", str(e))

    def mailing_thread(self):
        import smtplib; from email.message import EmailMessage; from openpyxl import Workbook; from pathlib import Path
        from openpyxl.styles import PatternFill, Border, Side, Alignment, Font
        
        try:
            srv = smtplib.SMTP("smtp.gmail.com", 587, timeout=15)
            srv.starttls(); srv.login(self.s_user.text, self.s_pass.text)
        except Exception as e: Clock.schedule_once(lambda dt: self.msg("Błąd SMTP", str(e))); return

        header = self.full_data[0]
        rows = self.full_data[1:]; sent = 0
        idxs = self.selected_export_cols if self.selected_export_cols else list(range(len(header)))
        
        # Znajdź kolumny Imię i Nazwisko w pliku danych
        h_low = [str(x).lower() for x in header]
        ni = next((i for i, x in enumerate(h_low) if "imi" in x), 0)
        si = next((i for i, x in enumerate(h_low) if "nazw" in x), 1)

        for i, row in enumerate(rows):
            name, sur = str(row[ni]).lower().strip(), str(row[si]).lower().strip()
            res = self.db_conn.execute("SELECT email FROM address_book WHERE name=? AND surname=?", (name, sur)).fetchone()
            if res:
                email = res[0]
                try:
                    msg = EmailMessage(); msg["Subject"] = "Raport Future"; msg["From"] = self.s_user.text; msg["To"] = email
                    msg.set_content(f"Dzień dobry,\nW załączniku przesyłamy raport.")
                    
                    # --- FORMATOWANY EXCEL ---
                    wb = Workbook(); ws = wb.active
                    fill = PatternFill(start_color='FFCFE2F3', end_color='FFCFE2F3', fill_type='solid')
                    thin = Side(border_style="thin", color="000000")
                    border = Border(top=thin, left=thin, right=thin, bottom=thin)
                    
                    ws.append([header[idx] for idx in idxs])
                    for cell in ws[1]:
                        cell.fill = fill; cell.font = Font(bold=True)
                        cell.border = border; cell.alignment = Alignment(horizontal="center")
                    
                    ws.append([row[idx] for idx in idxs])
                    for cell in ws[2]:
                        cell.border = border; cell.alignment = Alignment(horizontal="left")

                    for col in ws.columns:
                        max_l = max(len(str(cell.value or "")) for cell in col)
                        ws.column_dimensions[col[0].column_letter].width = max_l + 3

                    tmp = Path(self.user_data_dir) / "report.xlsx"; wb.save(str(tmp))
                    with open(tmp, "rb") as f:
                        msg.add_attachment(f.read(), maintype="application", subtype="xlsx", filename=f"Raport_{name}.xlsx")
                    srv.send_message(msg); sent += 1
                except: pass
            Clock.schedule_once(lambda dt, p=int((i+1)/len(rows)*100): setattr(self.progress, "value", p))
        srv.quit(); Clock.schedule_once(lambda dt: self.msg("Koniec", f"Wysłano {sent} maili."))

    # Funkcje pomocnicze...
    def go_to_table(self, _):
        if not self.current_data_file: self.msg("Błąd", "Wgraj najpierw plik!"); return
        from openpyxl import load_workbook
        wb = load_workbook(str(self.current_data_file), data_only=True); ws = wb.active
        self.full_data = [["" if v is None else str(v) for v in row] for row in ws.iter_rows(values_only=True)]
        self.show_table(); self.sm.current = "table"

    def show_table(self):
        self.grid.clear_widgets()
        if not self.full_data: return
        r, c = len(self.full_data), len(self.full_data[0])
        self.grid.cols = c; self.grid.width, self.grid.height = c * dp(160), r * dp(42)
        for row in self.full_data:
            for cell in row: self.grid.add_widget(Label(text=str(cell), size_hint=(None, None), size=(dp(160), dp(42))))

    def filter_data(self, ins, val):
        filtered = [r for r in self.full_data if any(val.lower() in str(c).lower() for c in r)]
        self.grid.clear_widgets()
        for row in filtered:
            for cell in row: self.grid.add_widget(Label(text=str(cell), size_hint=(None, None), size=(dp(160), dp(42))))

    def show_column_selector(self, _):
        if not self.full_data: return
        box = BoxLayout(orientation="vertical", padding=dp(10), spacing=dp(5))
        scroll = ScrollView(); grid = GridLayout(cols=1, size_hint_y=None, spacing=dp(8)); grid.bind(minimum_height=grid.setter('height'))
        checks = []
        for i, h in enumerate(self.full_data[0]):
            row = BoxLayout(size_hint_y=None, height=dp(40)); cb = CheckBox(size_hint_x=0.2); cb.active = True
            row.add_widget(cb); row.add_widget(Label(text=str(h))); grid.add_widget(row); checks.append((i, cb))
        def apply(_):
            self.selected_export_cols = [idx for idx, cb in checks if cb.active]; p.dismiss()
        btn = PremiumButton(text="Zatwierdź"); btn.bind(on_press=apply)
        scroll.add_widget(grid); box.add_widget(scroll); box.add_widget(btn)
        p = Popup(title="Zaznacz kolumny raportu", content=box, size_hint=(0.9, 0.8)); p.open()

    def save_smtp(self, _):
        import json; from pathlib import Path
        d = {"user": self.s_user.text, "pass": self.s_pass.text}
        with open(Path(self.user_data_dir) / "smtp.json", "w") as f: json.dump(d, f)
        self.test_smtp()

    def load_smtp(self):
        import json; from pathlib import Path
        p = Path(self.user_data_dir) / "smtp.json"
        if p.exists():
            with open(p) as f:
                d = json.load(f); self.s_user.text = d.get("user", ""); self.s_pass.text = d.get("pass", "")

    def test_smtp(self):
        import threading
        def _t():
            import smtplib
            try:
                s = smtplib.SMTP("smtp.gmail.com", 587, timeout=10); s.starttls(); s.login(self.s_user.text, self.s_pass.text); s.quit()
                Clock.schedule_once(lambda dt: self.msg("Test OK", "Połączono!"))
            except Exception as e: Clock.schedule_once(lambda dt: self.msg("Błąd SMTP", str(e)))
        threading.Thread(target=_t).start()

    def start_mailing(self, _):
        import threading
        threading.Thread(target=self.mailing_thread).start()

    def msg(self, title, text):
        box = BoxLayout(orientation="vertical", padding=dp(20)); box.add_widget(Label(text=text))
        btn = PremiumButton(text="OK"); box.add_widget(btn)
        p = Popup(title=title, content=box, size_hint=(0.8, 0.4)); btn.bind(on_press=p.dismiss); p.open()

if __name__ == "__main__":
    FutureApp().run()
