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
APP_TITLE = "Future 9.0 ULTRA PRO"

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
        self.export_columns = []
        self.global_attachments = []
        
        self.init_db()
        self.sm = ScreenManager()
        
        self.home = HomeScreen(name="home")
        self.table = TableScreen(name="table")
        self.email = EmailScreen(name="email")
        self.smtp = SMTPScreen(name="smtp")
        self.tmpl = TemplateScreen(name="tmpl")

        self.build_home()
        self.build_table()
        self.build_email()
        self.build_smtp()
        self.build_tmpl()

        for s in [self.home, self.table, self.email, self.smtp, self.tmpl]:
            self.sm.add_widget(s)
            
        return self.sm

    def init_db(self):
        db_path = Path(self.user_data_dir) / "app_v9_core.db"
        self.conn = sqlite3.connect(str(db_path), check_same_thread=False)
        self.conn.execute("CREATE TABLE IF NOT EXISTS contacts (name TEXT, surname TEXT, email TEXT, PRIMARY KEY(name, surname))")
        self.conn.execute("CREATE TABLE IF NOT EXISTS settings (key TEXT PRIMARY KEY, val TEXT)")
        self.conn.execute("CREATE TABLE IF NOT EXISTS logs (id INTEGER PRIMARY KEY, msg TEXT, date TEXT)")
        
        if not self.conn.execute("SELECT val FROM settings WHERE key='t_body'").fetchone():
            self.conn.execute("INSERT OR REPLACE INTO settings VALUES ('t_sub', 'Raport dla {Imię}')")
            self.conn.execute("INSERT OR REPLACE INTO settings VALUES ('t_body', 'Witaj {Imię},\n\nPrzesyłamy Twój raport za bieżący miesiąc.')")
        self.conn.commit()

# -----------------------------
# STYLIZACJA EXCEL (RAMKI I DOPASOWANIE)
# -----------------------------
    def apply_excel_styling(self, ws):
        # 1. Autodopasowanie kolumn
        for col in ws.columns:
            max_length = 0
            column = col[0].column_letter
            for cell in col:
                try:
                    if cell.value:
                        max_length = max(max_length, len(str(cell.value)))
                except: pass
            ws.column_dimensions[column].width = max_length + 4

        # 2. Ramki i Align
        thick = Side(style='thick')
        thin = Side(style='thin')
        max_r = ws.max_row
        max_c = ws.max_column

        for r in range(1, max_r + 1):
            for c in range(1, max_c + 1):
                cell = ws.cell(row=r, column=c)
                if r == 1: cell.font = Font(bold=True)
                
                # Dynamiczna ramka zewnętrzna pogrubiona
                cell.border = Border(
                    left=thick if c == 1 else thin,
                    right=thick if c == max_c else thin,
                    top=thick if r == 1 else thin,
                    bottom=thick if r == max_r else thin
                )
                cell.alignment = Alignment(horizontal='center', vertical='center')

# -----------------------------
# EKRAN HOME
# -----------------------------
    def build_home(self):
        layout = BoxLayout(orientation="vertical", padding=dp(25), spacing=dp(15))
        layout.add_widget(Label(text=APP_TITLE, font_size=26, bold=True))
        
        btns = [
            ("📂 Otwórz Excel Płac", self.open_picker_data),
            ("📊 Podgląd Tabeli", self.load_excel),
            ("✉ Centrum Mailingu", lambda x: setattr(self.sm, "current", "email")),
            ("⚙ Ustawienia SMTP", lambda x: setattr(self.sm, "current", "smtp"))
        ]
        for t, c in btns:
            b = PremiumButton(text=t)
            b.bind(on_press=c)
            layout.add_widget(b)
            
        self.status = Label(text="Gotowy", color=(0.7, 0.7, 0.7, 1))
        layout.add_widget(self.status)
        self.home.add_widget(layout)

# -----------------------------
# ANDROID PICKER LOGIC
# -----------------------------
    def open_picker_data(self, _): self.open_picker(mode="data")

    def open_picker(self, mode="data"):
        if platform != "android":
            self.popup("Błąd", "Funkcja dostępna tylko na Androidzie")
            return
        from jnius import autoclass
        from android import activity
        PythonActivity = autoclass("org.kivy.android.PythonActivity")
        Intent = autoclass("android.content.Intent")
        intent = Intent(Intent.ACTION_GET_CONTENT)
        intent.setType("*/*")
        
        def callback(req, res, intent_data):
            if not intent_data: return
            activity.unbind(on_activity_result=callback)
            uri = intent_data.getData()
            resolver = PythonActivity.mActivity.getContentResolver()
            stream = resolver.openInputStream(uri)
            local = Path(self.user_data_dir) / f"tmp_{mode}_{os.urandom(2).hex()}.xlsx"
            with open(local, "wb") as f:
                buf = bytearray(4096)
                while True:
                    r = stream.read(buf)
                    if r == -1: break
                    f.write(buf[:r])
            stream.close()
            
            if mode == "data":
                self.current_file = local
                Clock.schedule_once(lambda dt: setattr(self.status, "text", "Wczytano Excel Główny"))
            elif mode == "book":
                self.import_contacts_to_db(local)
            elif mode == "attachment":
                self.global_attachments.append(str(local))
                Clock.schedule_once(lambda dt: self.update_email_ui_labels())

        activity.bind(on_activity_result=callback)
        PythonActivity.mActivity.startActivityForResult(intent, 1001)

# -----------------------------
# TABELA DANYCH
# -----------------------------
    def build_table(self):
        layout = BoxLayout(orientation="vertical", padding=dp(10), spacing=dp(10))
        top = BoxLayout(size_hint_y=0.12, spacing=dp(5))
        self.search = TextInput(hint_text="Szukaj...", multiline=False)
        self.search.bind(text=self.filter_data)
        col_btn = PremiumButton(text="Kolumny"); col_btn.bind(on_press=self.export_popup)
        back_btn = PremiumButton(text="Wróć"); back_btn.bind(on_press=lambda x: setattr(self.sm, "current", "home"))
        top.add_widget(self.search); top.add_widget(col_btn); top.add_widget(back_btn)
        
        self.scroll = ScrollView(); self.grid = GridLayout(size_hint=(None, None))
        self.grid.bind(minimum_height=self.grid.setter("height"), minimum_width=self.grid.setter("width"))
        self.scroll.add_widget(self.grid)
        self.progress = ProgressBar(max=100, size_hint_y=0.05)
        
        m_exp = Button(text="EKSPORTUJ WSZYSTKIE OSOBNO (RAMKI + AUTO)", size_hint_y=None, height=dp(45))
        m_exp.bind(on_press=self.mass_export_start)
        
        layout.add_widget(top); layout.add_widget(self.scroll); layout.add_widget(self.progress); layout.add_widget(m_exp)
        self.table.add_widget(layout)

    def load_excel(self, _):
        if not self.current_file:
            self.popup("Błąd", "Wybierz plik Excel!"); return
        wb = load_workbook(self.current_file, data_only=True)
        ws = wb.active
        self.full_data = [["" if v is None else str(v) for v in row] for row in ws.iter_rows(values_only=True)]
        self.filtered_data = self.full_data
        self.show_table(); self.sm.current = "table"

    def show_table(self):
        self.grid.clear_widgets()
        if not self.filtered_data: return
        rows, cols = len(self.filtered_data), len(self.filtered_data[0])
        w, h = dp(160), dp(42)
        self.grid.cols = cols + 1
        self.grid.width, self.grid.height = (cols + 1) * w, rows * h
        for head in self.filtered_data[0]:
            self.grid.add_widget(Label(text=str(head), size_hint=(None, None), size=(w, h), bold=True))
        self.grid.add_widget(Label(text="Akcja", size_hint=(None, None), size=(w, h), bold=True))
        for r in self.filtered_data[1:]:
            for c in r: self.grid.add_widget(Label(text=str(c), size_hint=(None, None), size=(w, h)))
            btn = Button(text="Zapisz", size_hint=(None, None), size=(w, h)); btn.bind(on_press=lambda x, row=r: self.single_export(row))
            self.grid.add_widget(btn)

    def filter_data(self, ins, val):
        v = val.lower()
        self.filtered_data = [self.full_data[0]] + [r for r in self.full_data[1:] if any(v in str(c).lower() for c in r)]
        self.show_table()

# -----------------------------
# CENTRUM EMAIL
# -----------------------------
    def build_email(self):
        layout = BoxLayout(orientation="vertical", padding=dp(20), spacing=dp(10))
        layout.add_widget(Label(text="Centrum Mailingowe", font_size=22, bold=True))
        self.email_info = Label(text="Kontakty: 0", size_hint_y=None, height=dp(30))
        self.att_info = Label(text="Załączniki: 0", size_hint_y=None, height=dp(30))
        layout.add_widget(self.email_info); layout.add_widget(self.att_info)
        
        btns = [
            ("📁 Wgraj Bazę Kontaktów", lambda x: self.open_picker(mode="book")),
            ("📝 Edytuj Szablon Wiadomości", lambda x: setattr(self.sm, "current", "tmpl")),
            ("📎 Dodaj Załącznik PDF", lambda x: self.open_picker(mode="attachment")),
            ("⚡ Test (Wysyłka do siebie)", self.send_test_email),
            ("🚀 URUCHOM MAILING MASOWY", self.send_emails_start),
            ("📜 Historia Wysyłek", self.show_history),
            ("Powrót", lambda x: setattr(self.sm, "current", "home"))
        ]
        for t, c in btns:
            b = PremiumButton(text=t); b.bind(on_press=c); layout.add_widget(b)
        self.update_email_ui_labels(); self.email.add_widget(layout)

    def update_email_ui_labels(self):
        count = self.conn.execute("SELECT count(*) FROM contacts").fetchone()[0]
        self.email_info.text = f"Baza GMAIL: {count} osób"
        self.att_info.text = f"Załączniki: {len(self.global_attachments)}"

# -----------------------------
# LOGIKA SZABLONU I SMTP
# -----------------------------
    def build_tmpl(self):
        layout = BoxLayout(orientation="vertical", padding=dp(20), spacing=dp(10))
        self.ts = TextInput(hint_text="Temat", size_hint_y=None, height=dp(45))
        self.tb = TextInput(hint_text="Treść (użyj {Imię})", multiline=True)
        rs = self.conn.execute("SELECT val FROM settings WHERE key='t_sub'").fetchone()
        rb = self.conn.execute("SELECT val FROM settings WHERE key='t_body'").fetchone()
        if rs: self.ts.text, self.tb.text = rs[0], rb[0]
        s_btn = PremiumButton(text="Zapisz"); s_btn.bind(on_press=self.save_tmpl)
        b_btn = PremiumButton(text="Cofnij"); b_btn.bind(on_press=lambda x: setattr(self.sm, "current", "email"))
        layout.add_widget(Label(text="Szablon Maila")); layout.add_widget(self.ts); layout.add_widget(self.tb); layout.add_widget(s_btn); layout.add_widget(b_btn)
        self.tmpl.add_widget(layout)

    def save_tmpl(self, _):
        self.conn.execute("UPDATE settings SET val=? WHERE key='t_sub'", (self.ts.text,))
        self.conn.execute("UPDATE settings SET val=? WHERE key='t_body'", (self.tb.text,))
        self.conn.commit(); self.popup("OK", "Szablon zapisany")

    def build_smtp(self):
        layout = BoxLayout(orientation="vertical", padding=dp(20), spacing=dp(10))
        self.su = TextInput(hint_text="Twój Gmail"); self.sp = TextInput(hint_text="Hasło App (16 znaków)", password=True)
        path = Path(self.user_data_dir) / "smtp.json"
        if path.exists():
            d = json.load(open(path)); self.su.text, self.sp.text = d.get('u',''), d.get('p','')
        sv = PremiumButton(text="Zapisz"); sv.bind(on_press=self.save_smtp)
        back = PremiumButton(text="Wróć"); back.bind(on_press=lambda x: setattr(self.sm, "current", "home"))
        layout.add_widget(Label(text="Ustawienia SMTP")); layout.add_widget(self.su); layout.add_widget(self.sp); layout.add_widget(sv); layout.add_widget(back)
        self.smtp.add_widget(layout)

    def save_smtp(self, _):
        with open(Path(self.user_data_dir) / "smtp.json", "w") as f:
            json.dump({'u': self.su.text, 'p': self.sp.text}, f)
        self.popup("OK", "SMTP zapisane")

# -----------------------------
# OPERACJE EXCEL I MAILING
# -----------------------------
    def import_contacts_to_db(self, path):
        try:
            wb = load_workbook(path, data_only=True); ws = wb.active
            rows = list(ws.iter_rows(values_only=True)); count = 0
            for r in rows[1:]:
                if r[0] and r[2]:
                    self.conn.execute("INSERT OR REPLACE INTO contacts VALUES (?,?,?)", (str(r[0]).lower().strip(), str(r[1]).lower().strip(), str(r[2]).strip()))
                    count += 1
            self.conn.commit()
            Clock.schedule_once(lambda dt: self.popup("Sukces", f"Pomyślnie wgrano {count} kontaktów do bazy."))
            Clock.schedule_once(lambda dt: self.update_email_ui_labels())
        except Exception as e:
            Clock.schedule_once(lambda dt: self.popup("Błąd Bazy", str(e)))

    def single_export(self, row, silent=False):
        f = Path("/storage/emulated/0/Documents/FutureExport"); f.mkdir(parents=True, exist_ok=True)
        wb = Workbook(); ws = wb.active
        ws.append(self.full_data[0]); ws.append(row)
        self.apply_excel_styling(ws)
        name = f"{row[0]}_{row[1]}" if len(row)>1 else "file"
        wb.save(f / f"Raport_{name}.xlsx")
        if not silent: self.popup("OK", f"Zapisano w Documents/FutureExport")

    def mass_export_start(self, _):
        if not self.full_data: return
        threading.Thread(target=self._mass_task).start()

    def _mass_task(self):
        rows = self.filtered_data[1:]
        for i, r in enumerate(rows):
            self.single_export(r, silent=True)
            perc = int((i+1)/len(rows)*100)
            Clock.schedule_once(lambda dt: setattr(self.progress, "value", perc))
        Clock.schedule_once(lambda dt: self.popup("Gotowe", "Eksport zakończony pomyślnie"))

    def send_emails_start(self, _): threading.Thread(target=self._mail_task, args=(False,)).start()
    def send_test_email(self, _): threading.Thread(target=self._mail_task, args=(True,)).start()

    def _mail_task(self, is_test):
        p = Path(self.user_data_dir) / "smtp.json"
        if not p.exists(): Clock.schedule_once(lambda dt: self.popup("!", "Ustaw SMTP!")); return
        cfg = json.load(open(p))
        try:
            srv = smtplib.SMTP("smtp.gmail.com", 587, timeout=15); srv.starttls(); srv.login(cfg['u'], cfg['p'])
        except Exception as e: Clock.schedule_once(lambda dt: self.popup("SMTP Error", str(e))); return
        
        data_rows = self.full_data[1:2] if is_test else self.full_data[1:]
        sent = 0
        for i, row in enumerate(data_rows):
            target = cfg['u'] if is_test else ""
            if not is_test:
                res = self.conn.execute("SELECT email FROM contacts WHERE name=? AND surname=?", (str(row[0]).lower().strip(), str(row[1]).lower().strip())).fetchone()
                if res: target = res[0]
            
            if target:
                try:
                    msg = EmailMessage(); dat = datetime.now().strftime("%d.%m.%Y")
                    msg["Subject"] = self.ts.text.replace("{Imię}", str(row[0]))
                    msg["From"], msg["To"] = cfg['u'], target
                    msg.set_content(self.tb.text.replace("{Imię}", str(row[0])).replace("{Data}", dat))
                    
                    # Załącznik Raportu + Stylizacja
                    tmp = Path(self.user_data_dir) / "tmp_email.xlsx"
                    wb = Workbook(); ws = wb.active; ws.append(self.full_data[0]); ws.append(row)
                    self.apply_excel_styling(ws); wb.save(tmp)
                    with open(tmp, "rb") as f: msg.add_attachment(f.read(), maintype="application", subtype="xlsx", filename=f"Raport_{row[0]}.xlsx")
                    
                    for att in self.global_attachments:
                        if os.path.exists(att):
                            with open(att, "rb") as f: msg.add_attachment(f.read(), maintype="application", subtype="octet-stream", filename=os.path.basename(att))
                    
                    srv.send_message(msg); sent += 1
                    self.conn.execute("INSERT INTO logs (msg, date) VALUES (?,?)", (f"Wysłano: {target}", dat))
                except: pass
            perc = int((i+1)/len(data_rows)*100)
            Clock.schedule_once(lambda dt: setattr(self.progress, "value", perc))
            
        srv.quit(); self.conn.commit()
        Clock.schedule_once(lambda dt: self.popup("Mailing", f"Wysłano: {sent} maili"))

    def show_history(self, _):
        logs = self.conn.execute("SELECT msg, date FROM logs ORDER BY id DESC LIMIT 10").fetchall()
        txt = "\n".join([f"{d}: {m}" for m, d in logs])
        self.popup("Historia", txt if txt else "Brak historii")

    def export_popup(self, _):
        if not self.full_data: return
        box = BoxLayout(orientation="vertical", spacing=dp(5))
        checks = []
        for i, h in enumerate(self.full_data[0]):
            r = BoxLayout(); cb = CheckBox(active=True); checks.append((i, cb))
            r.add_widget(cb); r.add_widget(Label(text=str(h))); box.add_widget(r)
        b = PremiumButton(text="OK"); b.bind(on_press=lambda x: popup.dismiss()); box.add_widget(b)
        popup = Popup(title="Widoczność kolumn", content=box, size_hint=(0.9, 0.9)); popup.open()

    def popup(self, title, text):
        box = BoxLayout(orientation="vertical", padding=dp(20))
        box.add_widget(Label(text=text, halign="center"))
        b = PremiumButton(text="OK"); b.bind(on_press=lambda x: popup.dismiss()); box.add_widget(b)
        popup = Popup(title=title, content=box, size_hint=(0.8, 0.4)); popup.open()

# =====================================================
# MEGA STABILIZATION PATCH v9.5
# TABLE PREVIEW + DATABASE INFO + SAFE EXCEL LOAD
# =====================================================

from openpyxl.styles import PatternFill

HEADER_FILL = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")


# -----------------------------------------------------
# SAFE POPUP
# -----------------------------------------------------
def popup(self, title, msg):

    box = BoxLayout(orientation="vertical", padding=20, spacing=15)

    box.add_widget(Label(text=msg))

    btn = PremiumButton(text="OK")
    box.add_widget(btn)

    p = Popup(
        title=title,
        content=box,
        size_hint=(0.8,0.5)
    )

    btn.bind(on_press=p.dismiss)
    p.open()


FutureApp.popup = popup


# -----------------------------------------------------
# DATABASE INFO POPUP
# -----------------------------------------------------
def show_db_info(self):

    try:
        cur = self.conn.cursor()

        contacts = cur.execute("SELECT COUNT(*) FROM contacts").fetchone()[0]
        settings = cur.execute("SELECT COUNT(*) FROM settings").fetchone()[0]
        logs = cur.execute("SELECT COUNT(*) FROM logs").fetchone()[0]

        msg = f"""
Baza danych załadowana poprawnie

Kontakty: {contacts}
Ustawienia: {settings}
Logi: {logs}

Plik bazy:
{self.conn}
"""

        self.popup("Status bazy danych", msg)

    except Exception as e:

        self.popup("Błąd bazy", str(e))


FutureApp.show_db_info = show_db_info


# -----------------------------------------------------
# SAFE EXCEL LOADER
# -----------------------------------------------------
def safe_load_excel(self):

    if not self.current_file:

        self.popup(
            "Brak pliku",
            "Najpierw wybierz plik Excel"
        )
        return

    if not os.path.exists(self.current_file):

        self.popup(
            "Błąd pliku",
            "Plik Excel nie istnieje"
        )
        return

    try:

        wb = load_workbook(self.current_file)
        ws = wb.active

        data = []

        for row in ws.iter_rows(values_only=True):
            data.append(list(row))

        if len(data) == 0:

            self.popup(
                "Błąd",
                "Plik Excel jest pusty"
            )
            return

        self.full_data = data
        self.filtered_data = data

        self.popup(
            "Excel wczytany",
            f"Wiersze: {len(data)-1}\nKolumny: {len(data[0])}"
        )

        self.build_table_view()

        self.sm.current = "table"

    except Exception as e:

        self.popup(
            "Crash Excel",
            str(e)
        )


FutureApp.load_excel = safe_load_excel


# -----------------------------------------------------
# TABLE VIEW BUILDER
# -----------------------------------------------------
def build_table_view(self):

    try:

        self.table.clear_widgets()

        root = BoxLayout(orientation="vertical")

        scroll = ScrollView()

        cols = len(self.filtered_data[0])

        grid = GridLayout(
            cols=cols,
            spacing=2,
            size_hint_y=None
        )

        grid.bind(minimum_height=grid.setter('height'))

        for r, row in enumerate(self.filtered_data):

            for cell in row:

                text = "" if cell is None else str(cell)

                lbl = Label(
                    text=text,
                    size_hint_y=None,
                    height=dp(40)
                )

                if r == 0:

                    lbl.bold = True
                    lbl.color = (1,1,1,1)
                    lbl.canvas.before.clear()

                    with lbl.canvas.before:
                        from kivy.graphics import Color, Rectangle
                        Color(0.26,0.44,0.78,1)
                        lbl.rect = Rectangle(pos=lbl.pos, size=lbl.size)

                    def update_rect(instance, value):
                        instance.rect.pos = instance.pos
                        instance.rect.size = instance.size

                    lbl.bind(pos=update_rect, size=update_rect)

                grid.add_widget(lbl)

        scroll.add_widget(grid)
        root.add_widget(scroll)

        btn_back = PremiumButton(text="⬅ Powrót")

        btn_back.bind(
            on_press=lambda x: setattr(self.sm, "current", "home")
        )

        root.add_widget(btn_back)

        self.table.add_widget(root)

    except Exception as e:

        self.popup(
            "Błąd tabeli",
            str(e)
        )


FutureApp.build_table_view = build_table_view


# -----------------------------------------------------
# EXCEL STYLE PATCH
# -----------------------------------------------------
def improved_excel_style(self, ws):

    thick = Side(style="thick")
    thin = Side(style="thin")

    max_r = ws.max_row
    max_c = ws.max_column

    for r in range(1, max_r+1):
        for c in range(1, max_c+1):

            cell = ws.cell(r,c)

            if r == 1:

                cell.font = Font(bold=True)
                cell.fill = HEADER_FILL

            cell.border = Border(
                left=thick if c==1 else thin,
                right=thick if c==max_c else thin,
                top=thick if r==1 else thin,
                bottom=thick if r==max_r else thin
            )

            cell.alignment = Alignment(
                horizontal="center",
                vertical="center"
            )

    # AUTO WIDTH

    for col in ws.columns:

        max_len = 0
        letter = col[0].column_letter

        for cell in col:

            if cell.value:
                max_len = max(max_len, len(str(cell.value)))

        ws.column_dimensions[letter].width = max_len + 4


FutureApp.apply_excel_styling = improved_excel_style


# -----------------------------------------------------
# SAFE CONTACT IMPORT
# -----------------------------------------------------
def safe_import_contacts(self, file):

    try:

        wb = load_workbook(file)
        ws = wb.active

        added = 0

        for row in ws.iter_rows(min_row=2, values_only=True):

            if not row:
                continue

            name = str(row[0])
            surname = str(row[1])
            email = str(row[2])

            try:

                self.conn.execute(
                    "INSERT OR IGNORE INTO contacts VALUES (?,?,?)",
                    (name,surname,email)
                )

                added += 1

            except:
                pass

        self.conn.commit()

        self.popup(
            "Import kontaktów",
            f"Dodano {added} rekordów"
        )

        self.show_db_info()

    except Exception as e:

        self.popup(
            "Błąd importu",
            str(e)
        )


FutureApp.import_contacts_to_db = safe_import_contacts


# -----------------------------------------------------
# FINAL START MESSAGE
# -----------------------------------------------------
print("PATCH v9.5 LOADED")

if __name__ == "__main__":
    FutureApp().run()
