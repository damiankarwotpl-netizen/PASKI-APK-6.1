import os
import json
import sqlite3
import threading
import mimetypes
from datetime import datetime
from pathlib import Path
from email.message import EmailMessage

from kivy.app import App
from kivy.metrics import dp
from kivy.clock import Clock
from kivy.utils import platform
from kivy.core.window import Window
from kivy.uix.screenmanager import ScreenManager, Screen
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.button import Button
from kivy.uix.label import Label
from kivy.uix.popup import Popup
from kivy.uix.scrollview import ScrollView
from kivy.uix.gridlayout import GridLayout
from kivy.uix.textinput import TextInput
from kivy.uix.checkbox import CheckBox
from kivy.uix.progressbar import ProgressBar

from openpyxl import load_workbook, Workbook

# --- SYMBOLE I STYL ---
APP_TITLE = "Future 11.0 ULTIMATE"

class PremiumButton(Button):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.background_normal = ""
        self.background_color = (0.2, 0.4, 0.9, 1)
        self.color = (1, 1, 1, 1)
        self.size_hint_y = None
        self.height = dp(52)
        self.bold = True

class FutureApp(App):
    def build(self):
        Window.clearcolor = (0.08, 0.1, 0.15, 1)
        self.full_data = [] 
        self.current_file = None
        self.global_attachments = []
        self.export_columns = []
        
        self.init_db()
        self.sm = ScreenManager()
        
        # Ekrany
        self.home_scr = Screen(name="home")
        self.table_scr = Screen(name="table")
        self.email_scr = Screen(name="email")
        self.smtp_scr = Screen(name="smtp")
        self.tmpl_scr = Screen(name="tmpl")

        self.init_ui()

        for s in [self.home_scr, self.table_scr, self.email_scr, self.smtp_scr, self.tmpl_scr]:
            self.sm.add_widget(s)
        return self.sm

    def init_db(self):
        db_p = Path(self.user_data_dir) / "app_v11_stable.db"
        self.conn = sqlite3.connect(str(db_p), check_same_thread=False)
        self.conn.execute("CREATE TABLE IF NOT EXISTS contacts (name TEXT, surname TEXT, email TEXT, PRIMARY KEY(name, surname))")
        self.conn.execute("CREATE TABLE IF NOT EXISTS settings (key TEXT PRIMARY KEY, val TEXT)")
        self.conn.execute("CREATE TABLE IF NOT EXISTS logs (id INTEGER PRIMARY KEY, msg TEXT, date TEXT)")
        
        if not self.conn.execute("SELECT val FROM settings WHERE key='t_body'").fetchone():
            self.conn.execute("INSERT OR REPLACE INTO settings VALUES ('t_sub', 'Raport: {Imię} {Nazwisko}')")
            self.conn.execute("INSERT OR REPLACE INTO settings VALUES ('t_body', 'Witaj {Imię},\n\nPrzesyłamy raport.')")
        self.conn.commit()

    def init_ui(self):
        # --- HOME ---
        l_home = BoxLayout(orientation="vertical", padding=dp(25), spacing=dp(15))
        l_home.add_widget(Label(text=APP_TITLE, font_size=28, bold=True))
        l_home.add_widget(PremiumButton(text="📂 WCZYTAJ EXCEL PŁAC", on_press=lambda x: self.pick_file(mode="data")))
        l_home.add_widget(PremiumButton(text="📊 PODGLĄD TABELI", on_press=self.go_to_table))
        l_home.add_widget(PremiumButton(text="✉ CENTRUM MAILINGU", on_press=lambda x: setattr(self.sm, "current", "email")))
        l_home.add_widget(PremiumButton(text="⚙ USTAWIENIA GMAIL", on_press=lambda x: setattr(self.sm, "current", "smtp")))
        self.home_status = Label(text="Zacznij od wczytania danych", color=(0.7, 0.7, 0.7, 1))
        l_home.add_widget(self.home_status)
        self.home_scr.add_widget(l_home)

        # --- TABLE ---
        lt = BoxLayout(orientation="vertical", padding=dp(10), spacing=dp(10))
        top = BoxLayout(size_hint_y=None, height=dp(55), spacing=dp(5))
        self.search = TextInput(hint_text="Szukaj osoby...", multiline=False); self.search.bind(text=self.filter_data)
        top.add_widget(self.search)
        top.add_widget(Button(text="KOLUMNY", size_hint_x=0.3, on_press=self.column_selection_popup))
        self.scroll = ScrollView(); self.grid = GridLayout(size_hint=(None, None))
        self.grid.bind(minimum_height=self.grid.setter("height"), minimum_width=self.grid.setter("width"))
        self.scroll.add_widget(self.grid)
        self.progress = ProgressBar(max=100, size_hint_y=None, height=dp(10))
        lt.add_widget(top); lt.add_widget(self.scroll); lt.add_widget(self.progress)
        lt.add_widget(Button(text="EKSPORTUJ WSZYSTKO", size_hint_y=None, height=dp(50), on_press=self.mass_export))
        lt.add_widget(Button(text="COFNIJ", size_hint_y=None, height=dp(45), on_press=lambda x: setattr(self.sm, "current", "home")))
        self.table_scr.add_widget(lt)

        # --- EMAIL CENTER ---
        le = BoxLayout(orientation="vertical", padding=dp(20), spacing=dp(8))
        le.add_widget(Label(text="OPERACJE MAILINGOWE", font_size=22, bold=True))
        self.att_status = Label(text="Załączniki: 0", size_hint_y=None, height=dp(30))
        le.add_widget(self.att_status)
        btns = [
            ("📂 WCZYTAJ BAZĘ GMAIL", lambda x: self.pick_file(mode="book")),
            ("📝 EDYTUJ TREŚĆ MAILA", lambda x: setattr(self.sm, "current", "tmpl")),
            ("📎 DODAJ PDF/FOTO", self.attachment_manager),
            ("⚡ TEST MAILA (DO SIEBIE)", self.run_test_mail),
            ("📜 HISTORIA WYSYŁEK", self.show_history),
            ("🚀 URUCHOM MAILING MASOWY", self.start_mailing_thread),
            ("POWRÓT", lambda x: setattr(self.sm, "current", "home"))
        ]
        for t, c in btns: le.add_widget(PremiumButton(text=t, on_press=c))
        self.email_scr.add_widget(le)

        # --- USTAWIENIA I SZABLONY ---
        self.setup_settings_uis()

    def setup_settings_uis(self):
        # SMTP
        ls = BoxLayout(orientation="vertical", padding=dp(25), spacing=dp(10))
        self.s_user = TextInput(hint_text="Twój Gmail")
        self.s_pass = TextInput(hint_text="Hasło Aplikacji (16 znaków)", password=True)
        p = Path(self.user_data_dir) / "smtp.json"
        if p.exists(): d = json.load(open(p)); self.s_user.text, self.s_pass.text = d.get('u',''), d.get('p','')
        ls.add_widget(Label(text="Konfiguracja SMTP")); ls.add_widget(self.s_user); ls.add_widget(self.s_pass)
        ls.add_widget(Button(text="ZAPISZ", on_press=self.save_smtp))
        ls.add_widget(Button(text="POWRÓT", on_press=lambda x: setattr(self.sm, "current", "home")))
        self.smtp_scr.add_widget(ls)

        # SZABLON
        lt = BoxLayout(orientation="vertical", padding=dp(20), spacing=dp(10))
        self.ts = TextInput(); self.tb = TextInput()
        rs = self.conn.execute("SELECT val FROM settings WHERE key='t_sub'").fetchone()
        rb = self.conn.execute("SELECT val FROM settings WHERE key='t_body'").fetchone()
        if rs: self.ts.text, self.tb.text = rs[0], rb[0]
        lt.add_widget(Label(text="Temat i Treść ({Imię}, {Data})")); lt.add_widget(self.ts); lt.add_widget(self.tb)
        lt.add_widget(Button(text="ZAPISZ", on_press=self.save_tmpl))
        lt.add_widget(Button(text="POWRÓT", on_press=lambda x: setattr(self.sm, "current", "email")))
        self.tmpl_scr.add_widget(lt)

    # ========================================================
    # STABILNY PICKER (Z TWOJEJ WERSJI - BUFOR BAJTOWY [B])
    # ========================================================
    def pick_file(self, mode):
        if platform != "android": self.msg("Błąd", "Funkcja tylko na Android."); return
        from jnius import autoclass; from android import activity
        Intent = autoclass("android.content.Intent")
        intent = Intent(Intent.ACTION_GET_CONTENT); intent.setType("*/*"); intent.addCategory(Intent.CATEGORY_OPENABLE)
        
        def on_activity_result(request_code, result_code, intent_data):
            if intent_data:
                try:
                    uri = intent_data.getData()
                    resolver = autoclass("org.kivy.android.PythonActivity").mActivity.getContentResolver()
                    stream = resolver.openInputStream(uri)
                    
                    filename = "data_v11.xlsx" if mode == "data" else ("book_v11.xlsx" if mode == "book" else f"extra_{os.urandom(2).hex()}.pdf")
                    local_path = Path(self.user_data_dir) / filename
                    
                    # --- KLUCZOWY PATCH: NATIVE BYTE ARRAY ---
                    j_buf = autoclass('[B')(16384) 
                    with open(local_path, "wb") as f:
                        while True:
                            r = stream.read(j_buf)
                            if r <= 0: break
                            f.write(bytes(j_buf)[:r])
                    stream.close()

                    if mode == "data": 
                        self.current_file = local_path; Clock.schedule_once(lambda x: setattr(self.home_status, "text", "Załadowano Excel."))
                    elif mode == "book": 
                        self.import_contacts_to_db(local_path)
                    elif mode == "extra": 
                        self.global_attachments.append(str(local_path)); self.update_att_status()
                except Exception as e: Clock.schedule_once(lambda x: self.msg("Błąd", str(e)))
            activity.unbind(on_activity_result=on_activity_result)
            
        activity.bind(on_activity_result=on_activity_result)
        autoclass("org.kivy.android.PythonActivity").mActivity.startActivityForResult(intent, 1001)

    # ========================================================
    # WGRYWANIE I TABELA (Z TWOJEJ STABILNEJ WERSJI)
    # ========================================================
    def go_to_table(self, _):
        if not self.current_file: self.msg("!", "Wczytaj Excel!"); return
        try:
            wb = load_workbook(str(self.current_file), data_only=True); ws = wb.active
            self.full_data = [["" if v is None else str(v) for v in row] for row in ws.iter_rows(values_only=True)]
            self.show_table(); self.sm.current = "table"
        except Exception as e: self.msg("Błąd ZIP", str(e))

    def show_table(self):
        self.grid.clear_widgets()
        if not self.full_data: return
        r_count, c_count = len(self.full_data), len(self.full_data[0])
        w, h = dp(160), dp(42)
        self.grid.cols = c_count + 1
        self.grid.width, self.grid.height = (c_count + 1) * w, r_count * h
        # Header
        for h_val in self.full_data[0]: self.grid.add_widget(Label(text=str(h_val), size_hint=(None, None), size=(w, h), bold=True))
        self.grid.add_widget(Label(text="AKCJA", size_hint=(None, None), size=(w, h), bold=True))
        # Dane
        for row in self.full_data[1:]:
            for cell in row: self.grid.add_widget(Label(text=str(cell), size_hint=(None, None), size=(w, h)))
            btn = Button(text="ZAPISZ", size_hint=(None, None), size=(w, h), background_color=(0, 0.7, 0, 1))
            btn.bind(on_press=lambda x, r=row: self.single_export(r)); self.grid.add_widget(btn)

    # ========================================================
    # LOGIKA MAILINGU I PLIKÓW
    # ========================================================
    def import_contacts_to_db(self, path):
        try:
            wb = load_workbook(str(path), data_only=True); ws = wb.active; rows = list(ws.iter_rows(values_only=True))
            count = 0
            for r in rows[1:]:
                if r[2]: # Założenie: Imię (0), Nazwisko (1), Mail (2)
                    self.conn.execute("INSERT OR REPLACE INTO contacts VALUES(?,?,?)", (str(r[0]).lower().strip(), str(r[1]).lower().strip(), str(r[2]).strip()))
                    count += 1
            self.conn.commit(); self.msg("OK", f"Zaimportowano {count} osób.")
        except Exception as e: self.msg("Błąd Excel", str(e))

    def start_mailing_thread(self, _):
        if not self.full_data: self.msg("!", "Brak danych!"); return
        threading.Thread(target=self._mailing_process, args=(False,)).start()

    def run_test_mail(self, _):
        if not self.full_data: self.msg("!", "Wczytaj najpierw Excel!"); return
        threading.Thread(target=self._mailing_process, args=(True,)).start()

    def _mailing_process(self, is_test):
        p_cfg = Path(self.user_data_dir) / "smtp.json"
        if not p_cfg.exists(): Clock.schedule_once(lambda x: self.msg("!", "Zapisz SMTP!")); return
        cfg = json.load(open(p_cfg))
        try:
            srv = smtplib.SMTP("smtp.gmail.com", 587, timeout=15); srv.starttls(); srv.login(cfg['u'], cfg['p'])
        except Exception as e: Clock.schedule_once(lambda x: self.msg("Błąd", str(e))); return

        h, rows = self.full_data[0], (self.full_data[1:2] if is_test else self.full_data[1:])
        sent = 0
        for i, r in enumerate(rows):
            target = cfg['u'] if is_test else ""
            if not is_test:
                res = self.conn.execute("SELECT email FROM contacts WHERE name=? AND surname=?", (str(r[0]).lower().strip(), str(r[1]).lower().strip())).fetchone()
                if res: target = res[0]
            
            if target:
                msg = EmailMessage(); dat = datetime.now().strftime("%d.%m.%Y")
                msg["Subject"] = self.ts.text.replace("{Imię}", str(r[0])).replace("{Data}", dat)
                msg["From"], msg["To"] = cfg['u'], target
                msg.set_content(self.tb.text.replace("{Imię}", str(r[0])).replace("{Data}", dat))
                
                # Załącznik Excela
                tmp_p = Path(self.user_data_dir) / "tmp.xlsx"; wb = Workbook(); ws = wb.active
                idxs = self.export_columns if self.export_columns else list(range(len(h)))
                ws.append([h[k] for k in idxs]); ws.append([r[k] for k in idxs]); wb.save(str(tmp_p))
                with open(tmp_p, "rb") as f: msg.add_attachment(f.read(), maintype="application", subtype="xlsx", filename="Raport.xlsx")
                
                # Załączniki dodatkowe
                for ap in self.global_attachments:
                    if os.path.exists(ap):
                        with open(ap, "rb") as f: msg.add_attachment(f.read(), maintype="application", subtype="pdf", filename=os.path.basename(ap))
                
                try: srv.send_message(msg); sent += 1; self.conn.execute("INSERT INTO logs (msg, date) VALUES (?,?)", (f"Wysłano: {target}", dat))
                except: pass
            
            Clock.schedule_once(lambda dt, p=int(((i+1)/len(rows))*100): setattr(self.progress, "value", p))
        
        srv.quit(); self.conn.commit()
        Clock.schedule_once(lambda x: self.msg("Mailing", f"Wysłano e-maili: {sent}"))

    # --- EKSPORTY I POMOCNICZE ---
    def mass_export(self, _):
        if not self.full_data: return
        def _task():
            folder = Path("/storage/emulated/0/Documents/FutureExport"); folder.mkdir(parents=True, exist_ok=True)
            for i, r in enumerate(self.full_data[1:]):
                self.single_export(r, mass=True); Clock.schedule_once(lambda dt, p=int(((i+1)/len(self.full_data))*100): setattr(self.progress, "value", p))
            self.msg("OK", "Wszystko w Documents/FutureExport")
        threading.Thread(target=_task).start()

    def single_export(self, r, mass=False):
        folder = Path("/storage/emulated/0/Documents/FutureExport"); folder.mkdir(parents=True, exist_ok=True)
        wb = Workbook(); ws = wb.active; ws.append(self.full_data[0]); ws.append(r)
        wb.save(str(folder / f"Raport_{r[0]}_{r[1]}.xlsx"))
        if not mass: self.msg("OK", "Zapisano raport.")

    def column_selection_popup(self, _):
        if not self.full_data: return
        box = BoxLayout(orientation="vertical", padding=10); grid = GridLayout(cols=1, size_hint_y=None); grid.bind(minimum_height=grid.setter('height'))
        chks = []
        for i, h_val in enumerate(self.full_data[0]):
            row = BoxLayout(size_hint_y=None, height=dp(40)); cb = CheckBox(size_hint_x=0.2); cb.active = True
            row.add_widget(cb); row.add_widget(Label(text=str(h_val))); grid.add_widget(row); chks.append((i, cb))
        def apply(_): self.export_columns = [idx for idx, c in chks if c.active]; p.dismiss()
        sv = ScrollView(); sv.add_widget(grid); box.add_widget(sv); box.add_widget(Button(text="ZATWIERDŹ", on_press=apply, size_hint_y=None, height=dp(50)))
        p = Popup(title="Wybierz kolumny raportu", content=box, size_hint=(0.9, 0.9)); p.open()

    def attachment_manager(self, _):
        box = BoxLayout(orientation="vertical", padding=10, spacing=10)
        for ap in self.global_attachments:
            r = BoxLayout(size_hint_y=None, height=dp(40))
            r.add_widget(Label(text=os.path.basename(ap)[:20])); btn = Button(text="USUŃ", on_press=lambda x, p=ap: self.remove_att(p))
            r.add_widget(btn); box.add_widget(r)
        box.add_widget(Button(text="DODAJ DOKUMENT", on_press=lambda x: self.pick_file("extra")))
        box.add_widget(Button(text="ZAMKNIJ", on_press=lambda x: p.dismiss())); p = Popup(title="Załączniki", content=box, size_hint=(0.8, 0.6)); p.open()

    def remove_att(self, p):
        if p in self.global_attachments: self.global_attachments.remove(p)
        self.update_att_status()

    def save_smtp(self, _):
        with open(Path(self.user_data_dir) / "smtp.json", "w") as f: json.dump({'u': self.s_user.text, 'p': self.s_pass.text}, f)
        self.msg("OK", "Zapisano SMTP.")

    def save_tmpl(self, _):
        self.conn.execute("UPDATE settings SET val=? WHERE key='t_sub'", (self.ts.text,))
        self.conn.execute("UPDATE settings SET val=? WHERE key='t_body'", (self.tb.text,)); self.conn.commit()
        self.msg("OK", "Zapisano szablon.")

    def filter_data(self, ins, val):
        filtered = [self.full_data[0]] + [r for r in self.full_data[1:] if val.lower() in str(r).lower()]
        self.grid.clear_widgets() # To można ulepszyć renderując tylko filtered
        
    def show_history(self, _):
        logs = self.conn.execute("SELECT msg, date FROM logs ORDER BY id DESC LIMIT 15").fetchall()
        txt = "\n".join([f"{d}: {m}" for m, d in logs]); self.msg("Logi", txt if txt else "Brak")
    def update_att_status(self): self.att_status.text = f"Załączniki: {len(self.global_attachments)}"
    def msg(self, t, txt): Popup(title=t, content=Label(text=txt, halign="center"), size_hint=(0.8, 0.4)).open()

if __name__ == "__main__":
    import smtplib
    try:
        FutureApp().run()
    except Exception as e:
        import traceback
        with open("critical_crash.txt", "w") as f: f.write(traceback.format_exc())
