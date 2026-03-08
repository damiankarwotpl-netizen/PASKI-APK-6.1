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
        self.table_scr = Screen(name="table")
        self.email_scr = Screen(name="email")
        self.smtp_scr = Screen(name="smtp")

        self.init_ui()
        self.init_db()

        for s in [self.home, self.table_scr, self.email_scr, self.smtp_scr]:
            self.sm.add_widget(s)
        return self.sm

    def init_db(self):
        db_path = Path(self.user_data_dir) / "app_v9_final.db"
        self.conn = sqlite3.connect(str(db_path))
        self.conn.execute("CREATE TABLE IF NOT EXISTS contacts (name TEXT, surname TEXT, email TEXT, PRIMARY KEY(name, surname))")
        self.conn.commit()

    def init_ui(self):
        # --- HOME ---
        l_home = BoxLayout(orientation="vertical", padding=dp(25), spacing=dp(20))
        l_home.add_widget(Label(text=APP_TITLE, font_size=26))
        b1 = PremiumButton(text="📂 Wczytaj plik z wypłatami (Excel)"); b1.bind(on_press=lambda x: self.open_picker(mode="data"))
        b2 = PremiumButton(text="📊 Otwórz Tabelę"); b2.bind(on_press=self.go_to_table)
        b3 = PremiumButton(text="⚙ Konfiguracja Gmail"); b3.bind(on_press=lambda x: setattr(self.sm, "current", "smtp"))
        self.home_status = Label(text="Witaj! Wczytaj dane, aby zacząć.")
        for w in [b1, b2, b3, self.home_status]: l_home.add_widget(w)
        self.home.add_widget(l_home)

        # --- TABLE ---
        lt = BoxLayout(orientation="vertical", padding=dp(10), spacing=dp(10))
        top = BoxLayout(size_hint=(1, 0.12), spacing=dp(8))
        self.search = TextInput(hint_text="Szukaj osoby...", multiline=False); self.search.bind(text=self.filter_data)
        b_next = PremiumButton(text="Wysyłka/Eksport"); b_next.bind(on_press=lambda x: setattr(self.sm, "current", "email"))
        b_back = PremiumButton(text="Powrót"); b_back.bind(on_press=lambda x: setattr(self.sm, "current", "home"))
        top.add_widget(self.search); top.add_widget(b_next); top.add_widget(b_back)
        self.scroll = ScrollView(); self.grid = GridLayout(size_hint=(None, None))
        self.grid.bind(minimum_height=self.grid.setter("height"), minimum_width=self.grid.setter("width"))
        self.scroll.add_widget(self.grid); self.progress = ProgressBar(max=100, size_hint=(1, 0.05))
        lt.add_widget(top); lt.add_widget(self.scroll); lt.add_widget(self.progress)
        self.table_scr.add_widget(lt)

        # --- EMAIL & EXPORT ALL ---
        le = BoxLayout(orientation="vertical", padding=dp(20), spacing=dp(15))
        le.add_widget(Label(text="Centrum Operacyjne", font_size=22))
        b_book = PremiumButton(text="📥 Importuj Bazę Kontaktów"); b_book.bind(on_press=lambda x: self.open_picker(mode="book"))
        b_cols = PremiumButton(text="📋 Wybierz kolumny raportu"); b_cols.bind(on_press=self.column_popup)
        b_exp_all = PremiumButton(text="💾 Eksportuj wszystko (FOLDER)"); b_exp_all.bind(on_press=self.start_export_all_thread)
        b_send = PremiumButton(text="🚀 Wyślij Maila do WSZYSTKICH"); b_send.bind(on_press=self.start_mailing)
        b_prev = PremiumButton(text="Wróć"); b_prev.bind(on_press=lambda x: setattr(self.sm, "current", "table"))
        self.email_status = Label(text="Automatyczne dopasowanie po Imieniu i Nazwisku")
        for w in [b_book, b_cols, b_exp_all, b_send, b_prev, self.email_status]: le.add_widget(w)
        self.email_scr.add_widget(le)

        # --- SMTP ---
        ls = BoxLayout(orientation="vertical", padding=dp(20), spacing=dp(10))
        self.s_user = TextInput(hint_text="Twój Gmail"); self.s_pass = TextInput(hint_text="Hasło Aplikacji (16 znaków)", password=True)
        b_save = PremiumButton(text="Zapisz"); b_save.bind(on_press=self.save_smtp)
        ls.add_widget(Label(text="Konfiguracja Poczty")); ls.add_widget(self.s_user); ls.add_widget(self.s_pass); ls.add_widget(b_save)
        ls.add_widget(PremiumButton(text="Powrót", on_press=lambda x: setattr(self.sm, "current", "home")))
        self.smtp_scr.add_widget(ls); self.load_smtp()

    # --- LOGIKA NAPRAWCZA PLIKÓW (FIX ZIP ERROR) ---
    def open_picker(self, mode):
        if platform != "android": self.msg("Błąd", "Funkcja tylko na Android"); return
        from jnius import autoclass; from android import activity
        Intent = autoclass("android.content.Intent")
        intent = Intent(Intent.ACTION_OPEN_DOCUMENT); intent.setType("*/*"); intent.addCategory(Intent.CATEGORY_OPENABLE)
        def on_res(req, res, dt):
            if dt:
                try:
                    uri = dt.getData(); resolver = autoclass("org.kivy.android.PythonActivity").mActivity.getContentResolver()
                    stream = resolver.openInputStream(uri)
                    fname = "data_v9.xlsx" if mode == "data" else "book_v9.xlsx"
                    local = Path(self.user_data_dir) / fname
                    with open(local, "wb") as f:
                        # Fix Android Stream Read - Java Byte Array
                        j_buf = autoclass('java.lang.reflect.Array').newInstance(autoclass('java.lang.Byte').TYPE, 4096)
                        while True:
                            r = stream.read(j_buf)
                            if r <= 0: break
                            f.write(bytes(j_buf)[:r])
                    stream.close()
                    if mode == "data": 
                        self.current_file = local; setattr(self.home_status, "text", "Wczytano plik płac.")
                    else: 
                        self.import_contacts_to_db(local)
                except Exception as e: self.msg("Błąd pliku", str(e))
            activity.unbind(on_activity_result=on_res)
        activity.bind(on_activity_result=on_res); autoclass("org.kivy.android.PythonActivity").mActivity.startActivityForResult(intent, 1001 if mode == "data" else 1002)

    def go_to_table(self, _):
        if not self.current_file: self.msg("Błąd", "Wczytaj najpierw plik!"); return
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
        # Tabele + kolumna AKCJA
        self.grid.cols = c + 1
        self.grid.width, self.grid.height = (c + 1) * w, r * h
        # Nagłówek
        for v in self.filtered_data[0]: self.grid.add_widget(Label(text=str(v), size_hint=(None, None), size=(w, h), bold=True))
        self.grid.add_widget(Label(text="AKCJA", size_hint=(None, None), size=(w, h), bold=True))
        # Wiersze
        for row in self.filtered_data[1:]:
            for cell in row: self.grid.add_widget(Label(text=str(cell), size_hint=(None, None), size=(w, h)))
            btn = Button(text="EKSPORTUJ", size_hint=(None, None), size=(w, h), background_color=(0, 0.7, 0, 1))
            btn.bind(on_press=lambda x, r=row: self.export_styled_excel(self.full_data[0], r, mode="single"))
            self.grid.add_widget(btn)

    def filter_data(self, ins, val):
        self.filtered_data = [self.full_data[0]] + [r for r in self.full_data[1:] if any(val.lower() in str(c).lower() for c in r)]
        self.show_table()

    # --- FORMATOWANIE I EKSPORT ---
    def export_styled_excel(self, header, row, mode="single"):
        try:
            folder = Path("/storage/emulated/0/Documents/FutureExport")
            folder.mkdir(parents=True, exist_ok=True)
            wb = Workbook(); ws = wb.active
            idxs = self.export_columns if self.export_columns else list(range(len(header)))
            
            # Style
            blue = PatternFill(start_color='CFE2F3', end_color='CFE2F3', fill_type='solid')
            border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
            
            # Nagłówki
            ws.append([header[i] for i in idxs])
            for cell in ws[1]:
                cell.fill, cell.font, cell.border = blue, Font(bold=True), border
                cell.alignment = Alignment(horizontal='center')
            
            # Dane
            ws.append([row[i] for i in idxs])
            for cell in ws[2]: cell.border = border

            # Auto-dopasowanie szerokości
            for col in ws.columns:
                max_l = max(len(str(c.value or "")) for c in col)
                ws.column_dimensions[col[0].column_letter].width = max_l + 3
            
            name = str(row[0]).strip().replace(" ", "_")
            file_path = folder / f"Raport_{name}_{datetime.now().strftime('%H%M%S')}.xlsx"
            wb.save(str(file_path))
            if mode == "single": self.msg("Sukces", f"Zapisano w Documents/FutureExport")
        except Exception as e: self.msg("Błąd Eksportu", str(e))

    def start_export_all_thread(self, _):
        if not self.full_data: return
        threading.Thread(target=self._export_all).start()

    def _export_all(self):
        head, rows = self.full_data[0], self.full_data[1:]
        for i, r in enumerate(rows):
            self.export_styled_excel(head, r, mode="mass")
            Clock.schedule_once(lambda dt, p=int((i+1)/len(rows)*100): setattr(self.progress, "value", p))
        self.msg("Gotowe", f"Wyeksportowano {len(rows)} plików do Documents/FutureExport")

    # --- KONTAKTY I WYSYŁKA ---
    def import_contacts_to_db(self, path):
        try:
            wb = load_workbook(str(path), data_only=True); ws = wb.active; rows = list(ws.iter_rows(values_only=True))
            h = [str(x).lower().strip() for x in rows[0]]
            def find_idx(keys):
                for i, v in enumerate(h):
                    if any(k in v for k in keys): return i
                return None
            ni, si, mi = find_idx(["imi"]), find_idx(["nazw"]), find_idx(["mail", "email", "adres"])
            if mi is None: self.msg("Błąd", "Nie znaleziono kolumny Email!"); return
            for r in rows[1:]:
                if r[mi]: self.conn.execute("INSERT OR REPLACE INTO contacts VALUES(?,?,?)", (str(r[ni or 0]).lower(), str(r[si or 1]).lower(), str(r[mi])))
            self.conn.commit(); self.msg("Baza", "Zaimportowano kontakty.")
        except Exception as e: self.msg("Błąd", str(e))

    def start_mailing(self, _):
        threading.Thread(target=self._mailing_process).start()

    def _mailing_process(self):
        p_smtp = Path(self.user_data_dir) / "config.json"
        if not p_smtp.exists(): Clock.schedule_once(lambda x: self.msg("Błąd", "Ustaw Gmail SMTP!")); return
        conf = json.load(open(p_smtp))
        try:
            srv = smtplib.SMTP("smtp.gmail.com", 587, timeout=15); srv.starttls(); srv.login(conf['u'], conf['p'])
        except Exception as e: Clock.schedule_once(lambda x: self.msg("Błąd Poczty", str(e))); return

        head, rows = self.full_data[0], self.full_data[1:]; sent = 0
        h_low = [str(x).lower() for x in head]
        ni, si = next((i for i, x in enumerate(h_low) if "imi" in x), 0), next((i for i, x in enumerate(h_low) if "nazw" in x), 1)

        for i, r in enumerate(rows):
            db_res = self.conn.execute("SELECT email FROM contacts WHERE name=? AND surname=?", (str(r[ni]).lower().strip(), str(r[si]).lower().strip())).fetchone()
            if db_res:
                tmp_p = Path(self.user_data_dir) / "temp.xlsx"
                self.export_styled_excel(head, r, mode="mass") # tutaj wykorzystujemy tę samą logikę zapisu
                # (W kodzie wysyłki trzeba zmienić docelowy zapis na temp, żeby go wysłać)
                msg = EmailMessage(); msg["Subject"] = "Raport Future"; msg["From"] = conf['u']; msg["To"] = db_res[0]
                msg.set_content("W załączniku przesyłamy raport miesięczny."); msg.add_attachment(open(tmp_p, "rb").read(), maintype="application", subtype="xlsx", filename="Raport.xlsx")
                srv.send_message(msg); sent += 1
            Clock.schedule_once(lambda dt, p=int((i+1)/len(rows)*100): setattr(self.progress, "value", p))
        srv.quit(); self.msg("Koniec", f"Wysłano {sent} maili.")

    # --- HELPERS ---
    def column_popup(self, _):
        if not self.full_data: return
        box = BoxLayout(orientation="vertical", padding=dp(10)); scroll = ScrollView(); grid = GridLayout(cols=1, size_hint_y=None); grid.bind(minimum_height=grid.setter('height'))
        checks = []
        for i, h in enumerate(self.full_data[0]):
            r = BoxLayout(size_hint_y=None, height=dp(40)); cb = CheckBox(size_hint_x=0.2); cb.active = True
            r.add_widget(cb); r.add_widget(Label(text=str(h))); grid.add_widget(r); checks.append((i, cb))
        def apply(_): self.export_columns = [idx for idx, c in checks if c.active]; p.dismiss()
        scroll.add_widget(grid); box.add_widget(scroll); btn = PremiumButton(text="Zatwierdź"); btn.bind(on_press=apply); box.add_widget(btn)
        p = Popup(title="Wybierz kolumny raportu", content=box, size_hint=(0.9, 0.9)); p.open()

    def msg(self, t, txt):
        b = BoxLayout(orientation="vertical", padding=dp(20)); b.add_widget(Label(text=txt)); btn = Button(text="OK", size_hint_y=None, height=dp(50))
        p = Popup(title=t, content=b, size_hint=(0.8, 0.4)); btn.bind(on_press=p.dismiss); b.add_widget(btn); p.open()

    def save_smtp(self, _):
        with open(Path(self.user_data_dir) / "config.json", "w") as f: json.dump({'u': self.s_user.text, 'p': self.s_pass.text}, f)
        self.msg("OK", "Zapisano hasło.")

    def load_smtp(self):
        p = Path(self.user_data_dir) / "config.json"
        if p.exists():
            d = json.load(open(p)); self.s_user.text = d['u']; self.s_pass.text = d['p']

if __name__ == "__main__":
    FutureApp().run()
