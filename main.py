
import os
import json
import sqlite3
import threading
import smtplib
import mimetypes
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
# ... (rest of imports same as before)
from kivy.uix.label import Label
from kivy.uix.popup import Popup
from kivy.uix.scrollview import ScrollView
from kivy.uix.gridlayout import GridLayout
from kivy.uix.textinput import TextInput
from kivy.uix.checkbox import CheckBox
from kivy.uix.screenmanager import ScreenManager, Screen
from kivy.graphics import Color, Rectangle

from openpyxl import load_workbook, Workbook
from openpyxl.styles import Border, Side, Font, Alignment

try:
    import xlrd
except ImportError:
    xlrd = None

COLOR_PRIMARY = (0.1, 0.5, 0.9, 1)
COLOR_BG = (1, 1, 1, 1)
COLOR_HEADER = (0.9, 0.9, 0.95, 1)
COLOR_ROW_A = (1, 1, 1, 1)
COLOR_ROW_B = (0.94, 0.97, 1, 1)

class PremiumButton(Button):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.background_normal = ""
        self.background_color = COLOR_PRIMARY
        self.height = dp(53)
        self.size_hint_y = None
        self.bold = True

class ColorSafeLabel(Label):
    def __init__(self, bg_color=(1,1,1,1), text_color=(0,0,0,1), **kwargs):
        super().__init__(**kwargs)
        self.color = text_color
        self.halign = 'center'
        self.valign = 'middle'
        with self.canvas.before:
            Color(*bg_color)
            self.rect = Rectangle(size=self.size, pos=self.pos)
        self.bind(size=self._update, pos=self._update)

    def _update(self, inst, val):
        self.rect.size = self.size
        self.rect.pos = self.pos
        self.text_size = (self.width - dp(10), None)

class HomeScreen(Screen): pass
class TableScreen(Screen): pass
class EmailScreen(Screen): pass
class SMTPScreen(Screen): pass
class TemplateScreen(Screen): pass
class ContactsScreen(Screen): pass

class FutureApp(App):
    def build(self):
        Window.clearcolor = (0.08, 0.1, 0.15, 1)
        self.full_data = []
        self.filtered_data = []
        self.export_indices = []
        self.global_attachments = []
        self.selected_emails = []
        self.queue = []
        self.stats = {"ok": 0, "fail": 0, "skip": 0, "auto": 0}
        self.idx_name, self.idx_surname, self.idx_pesel = 0, 1, -1
        self.init_db()
        self.sm = ScreenManager()
        self.add_screens()
        return self.sm

    def init_db(self):
        db_p = Path(self.user_data_dir) / "future_ultimate_v2.db"
        self.conn = sqlite3.connect(str(db_p), check_same_thread=False)
        self.conn.execute("CREATE TABLE IF NOT EXISTS contacts (name TEXT, surname TEXT, email TEXT, pesel TEXT, phone TEXT, PRIMARY KEY(name, surname))")
        self.conn.execute("CREATE TABLE IF NOT EXISTS settings (key TEXT PRIMARY KEY, val TEXT)")
        self.conn.commit()

    def add_screens(self):
        self.screens = {
            "home": HomeScreen(name="home"), "table": TableScreen(name="table"),
            "email": EmailScreen(name="email"), "smtp": SMTPScreen(name="smtp"),
            "tmpl": TemplateScreen(name="tmpl"), "contacts": ContactsScreen(name="contacts")
        }
        self.setup_ui()
        for s in self.screens.values(): self.sm.add_widget(s)

    def start_special_send_flow(self, _):
        self.open_picker("special_send")

    def special_send_step_recipients(self, file_path):
        self.selected_emails = []
        box = BoxLayout(orientation="vertical", padding=dp(10), spacing=dp(10))
        box.add_widget(Label(text="KROK 2: WYBIERZ ODBIORCÓW", bold=True, color=(1,1,1,1), size_hint_y=None, height=dp(40)))
        ti_search = TextInput(hint_text="Szukaj...", size_hint_y=None, height=dp(45), multiline=False)
        list_layout = GridLayout(cols=1, size_hint_y=None, spacing=dp(5))
        list_layout.bind(minimum_height=list_layout.setter('height'))
        
        def refresh_sublist(val=""):
            list_layout.clear_widgets()
            rows = self.conn.execute("SELECT name, surname, email FROM contacts").fetchall()
            for n, s, e in rows:
                if val and val.lower() not in f"{n} {s} {e}".lower(): continue
                row = BoxLayout(size_hint_y=None, height=dp(50), padding=[dp(10), 0])
                cb = CheckBox(size_hint_x=None, width=dp(50))
                cb.bind(active=lambda inst, v, mail=e: [self.selected_emails.append(mail) if v else self.selected_emails.remove(mail)])
                row.add_widget(cb)
                row.add_widget(Label(text=f"{n.title()} {s.title()} ({e})", halign="left", text_size=(dp(250), None)))
                list_layout.add_widget(row)

        ti_search.bind(text=lambda i, v: refresh_sublist(v))
        refresh_sublist()
        scroll = ScrollView()
        scroll.add_widget(list_layout)
        box.add_widget(ti_search)
        box.add_widget(scroll)
        btn_next = PremiumButton(text="DALEJ")
        btn_next.bind(on_press=lambda x: [p.dismiss(), self.special_send_step_message(file_path)] if self.selected_emails else self.msg("!", "Wybierz kogoś!"))
        box.add_widget(btn_next)
        p = Popup(title="Wybór odbiorców", content=box, size_hint=(0.95, 0.9))
        p.open()

    def special_send_step_message(self, file_path):
        box = BoxLayout(orientation="vertical", padding=dp(15), spacing=dp(10))
        ti_sub = TextInput(hint_text="Temat wiadomości", size_hint_y=None, height=dp(50), multiline=False)
        ti_body = TextInput(hint_text="Treść wiadomości...", multiline=True)
        lbl_info = Label(text=f"Plik: {os.path.basename(file_path)}\nOdbiorców: {len(self.selected_emails)}", size_hint_y=None, height=dp(60))
        box.add_widget(lbl_info)
        box.add_widget(ti_sub)
        box.add_widget(ti_body)

        def final_send(_):
            subj, body = ti_sub.text, ti_body.text
            if not subj or not body:
                self.msg("!", "Wpisz temat i treść!")
                return
            p.dismiss()
            for mail in self.selected_emails:
                self.send_direct_email(mail, file_path, subj, body)
            self.msg("Start", "Rozpoczęto proces wysyłki...")

        btn_send = PremiumButton(text="WYŚLIJ TERAZ")
        btn_send.bind(on_press=final_send)
        box.add_widget(btn_send)
        p = Popup(title="Wpisz treść", content=box, size_hint=(0.95, 0.8))
        p.open()

    def send_direct_email(self, target, file_path, subject, body):
        def thread_task():
            p_smtp = Path(self.user_data_dir) / "smtp.json"
            if not p_smtp.exists():
                Clock.schedule_once(lambda dt: self.msg("Błąd", "Skonfiguruj SMTP!"))
                return
            cfg = json.load(open(p_smtp))
            try:
                srv = smtplib.SMTP("smtp.gmail.com", 587, timeout=15)
                srv.starttls()
                srv.login(cfg['u'], cfg['p'])
                msg = EmailMessage()
                msg["Subject"] = subject
                msg["From"], msg["To"] = cfg['u'], target
                msg.set_content(body)
                
                if os.path.exists(file_path):
                    ctype, _ = mimetypes.guess_type(file_path)
                    if not ctype: ctype = 'application/octet-stream'
                    main, sub = ctype.split('/', 1)
                    with open(file_path, "rb") as f:
                        msg.add_attachment(f.read(), maintype=main, subtype=sub, filename=os.path.basename(file_path))
                
                srv.send_message(msg)
                srv.quit()
                Clock.schedule_once(lambda dt: self.msg("Potwierdzenie", f"Wysłano pomyślnie do:\n{target}"))
            except Exception as e:
                Clock.schedule_once(lambda dt: self.msg("Błąd wysyłki", f"Adres: {target}\n\nSzczegóły: {str(e)}"))
        threading.Thread(target=thread_task, daemon=True).start()

    def refresh_table(self):
        self.table_content_layout.clear_widgets()
        self.table_header_layout.clear_widgets()
        if not self.filtered_data or not self.export_indices: return
        w, h = dp(200), dp(55)
        headers = [self.full_data[0][i] for i in self.export_indices]
        self.table_header_layout.cols = len(headers) + 1
        self.table_header_layout.width = (len(headers) + 1) * w
        for head in headers:
            self.table_header_layout.add_widget(ColorSafeLabel(text=str(head), bg_color=COLOR_HEADER, bold=True, size=(w, h), size_hint=(None, None)))
        self.table_header_layout.add_widget(ColorSafeLabel(text="Akcja", bg_color=COLOR_HEADER, bold=True, size=(w, h), size_hint=(None, None)))
        self.table_content_layout.cols = len(headers) + 1
        self.table_content_layout.width = (len(headers) + 1) * w
        self.table_content_layout.height = (len(self.filtered_data) - 1) * h
        for r_idx, row in enumerate(self.filtered_data[1:]):
            row_bg = COLOR_ROW_A if r_idx % 2 == 0 else COLOR_ROW_B
            for c_idx in self.export_indices:
                val = row[c_idx] if c_idx < len(row) else ""
                self.table_content_layout.add_widget(ColorSafeLabel(text=str(val), bg_color=row_bg, size=(w, h), size_hint=(None, None)))
            btn = Button(text="Pojedynczy", size=(w, h), size_hint=(None, None))
            btn.bind(on_press=lambda x, r=row: self.export_xlsx(r))
            self.table_content_layout.add_widget(btn)

    def setup_ui(self):
        l = BoxLayout(orientation="vertical", padding=dp(30), spacing=dp(15))
        l.add_widget(Label(text="FUTURE 22.4 ULTIMATE", font_size='26sp', bold=True, color=COLOR_PRIMARY))
        btn = lambda t, c: l.add_widget(PremiumButton(text=t, on_press=c))
        btn("WCZYTAJ ARKUSZ PŁAC", lambda x: self.open_picker("data"))
        btn("PODGLĄD I EKSPORT", lambda x: [self.refresh_table(), setattr(self.sm, 'current', 'table')] if self.full_data else self.msg("!", "Brak danych"))
        btn("CENTRUM MAILINGOWE", lambda x: setattr(self.sm, 'current', 'email'))
        btn("USTAWIENIA SMTP", lambda x: setattr(self.sm, 'current', 'smtp'))
        self.screens["home"].add_widget(l)
        self.setup_table_ui(); self.setup_email_ui(); self.setup_smtp_ui(); self.setup_tmpl_ui(); self.setup_contacts_ui()

    def setup_table_ui(self):
        root = BoxLayout(orientation="vertical")
        menu = BoxLayout(size_hint_y=None, height=dp(55), spacing=dp(5), padding=dp(5))
        self.ti_search = TextInput(hint_text="Szukaj...", multiline=False)
        self.ti_search.bind(text=self.filter_table)
        menu.add_widget(self.ti_search)
        menu.add_widget(Button(text="Opcje", size_hint_x=0.2, on_press=self.popup_columns))
        menu.add_widget(Button(text="Wróć", size_hint_x=0.2, on_press=lambda x: setattr(self.sm, 'current', 'home')))
        header_scroll = ScrollView(size_hint_y=None, height=dp(55), do_scroll_y=False)
        self.table_header_layout = GridLayout(rows=1, size_hint=(None, None), height=dp(55))
        header_scroll.add_widget(self.table_header_layout)
        data_scroll = ScrollView(do_scroll_x=True, do_scroll_y=True)
        self.table_content_layout = GridLayout(size_hint=(None, None))
        self.table_content_layout.bind(minimum_height=self.table_content_layout.setter('height'), minimum_width=self.table_content_layout.setter('width'))
        data_scroll.add_widget(self.table_content_layout)
        data_scroll.bind(scroll_x=lambda inst, val: setattr(header_scroll, 'scroll_x', val))
        root.add_widget(menu); root.add_widget(header_scroll); root.add_widget(data_scroll)
        self.screens["table"].add_widget(root)

    def setup_email_ui(self):
        l = BoxLayout(orientation="vertical", padding=dp(25), spacing=dp(10))
        l.add_widget(Label(text="CENTRUM MAILINGOWE", font_size='22sp', bold=True))
        self.lbl_stats = Label(text="Baza: 0"); l.add_widget(self.lbl_stats)
        btn = lambda t, c: l.add_widget(PremiumButton(text=t, on_press=c))
        btn("IMPORT KONTAKTÓW (EXCEL)", lambda x: self.open_picker("book"))
        btn("ZARZĄDZAJ BAZĄ", lambda x: [self.refresh_contacts_list(), setattr(self.sm, 'current', 'contacts')])
        btn("EDYTUJ TREŚĆ MAILA", lambda x: setattr(self.sm, 'current', 'tmpl'))
        btn("DODAJ ZAŁĄCZNIK GLOBALNY", lambda x: self.open_picker("attachment"))
        btn("WYŚLIJ ZAŁĄCZNIK (SPECJALNE)", self.start_special_send_flow)
        btn("START MASOWA WYSYŁKA", self.start_mass_mailing)
        btn("POWRÓT", lambda x: setattr(self.sm, 'current', 'home'))
        self.screens["email"].add_widget(l); self.update_stats()

    def open_picker(self, mode):
        if platform != "android":
            self.msg("!", "Tylko Android"); return
        from jnius import autoclass
        from android import activity
        PA = autoclass("org.kivy.android.PythonActivity"); Intent = autoclass("android.content.Intent")
        intent = Intent(Intent.ACTION_GET_CONTENT); intent.setType("*/*")
        def cb(req, res, dt):
            if not dt: return
            activity.unbind(on_activity_result=cb)
            uri = dt.getData()
            d_name = "plik.xlsx"
            cur = PA.mActivity.getContentResolver().query(uri, None, None, None, None)
            if cur and cur.moveToFirst():
                idx = cur.getColumnIndex("_display_name")
                if idx != -1: d_name = cur.getString(idx)
                cur.close()
            stream = PA.mActivity.getContentResolver().openInputStream(uri)
            loc = Path(self.user_data_dir) / d_name
            with open(loc, "wb") as f:
                buf = bytearray(16384)
                while True:
                    n = stream.read(buf)
                    if n == -1: break
                    f.write(buf[:n])
            stream.close()
            if mode == "data": self.process_excel(loc)
            elif mode == "book": self.process_book(loc)
            elif mode == "attachment": self.global_attachments.append(str(loc)); self.update_stats()
            elif mode == "special_send": self.special_send_step_recipients(str(loc))
        activity.bind(on_activity_result=cb); PA.mActivity.startActivityForResult(intent, 1001)

    def process_excel(self, path):
        try:
            if str(path).endswith(".xls") and xlrd:
                wb = xlrd.open_workbook(path); ws = wb.sheet_by_index(0)
                raw = [[str(ws.cell_value(r,c)).strip() for c in range(ws.ncols)] for r in range(ws.nrows)]
            else:
                wb = load_workbook(path, data_only=True); ws = wb.active
                raw = [["" if v is None else str(v).strip() for v in r] for r in ws.iter_rows(values_only=True)]
            h_idx = 0
            for i, row in enumerate(raw[:15]):
                line = " ".join(row).lower()
                if any(x in line for x in ["imię", "imie", "nazwisko", "pesel"]): h_idx = i; break
            self.full_data = raw[h_idx:]; self.filtered_data = self.full_data
            self.export_indices = list(range(len(self.full_data[0])))
            h = [x.lower() for x in self.full_data[0]]
            for i,v in enumerate(h):
                if "imi" in v: self.idx_name = i
                if "naz" in v: self.idx_surname = i
                if "pesel" in v: self.idx_pesel = i
            self.msg("OK", "Wczytano tabele.")
        except Exception as e: self.msg("Błąd", str(e))

    def process_book(self, path):
        try:
            wb = load_workbook(path, data_only=True); ws = wb.active
            raw = [["" if v is None else str(v).strip() for v in r] for r in ws.iter_rows(values_only=True)]
            h = [x.lower() for x in raw[0]]; iN, iS, iE, iP = 0, 1, 2, -1
            for i,v in enumerate(h):
                if "imi" in v: iN=i
                elif "naz" in v: iS=i
                elif "@" in v or "mail" in v: iE=i
                elif "pesel" in v: iP=i
            for r in raw[1:]:
                if len(r) > iE and "@" in str(r[iE]):
                    pes_v = str(r[iP]) if (iP != -1 and len(r) > iP) else ""
                    self.conn.execute("INSERT OR REPLACE INTO contacts (name, surname, email, pesel, phone) VALUES (?,?,?,?,?)", (r[iN].lower(), r[iS].lower(), str(r[iE]).strip(), pes_v, ""))
            self.conn.commit(); self.update_stats(); self.msg("OK", "Zaktualizowano baze.")
        except: self.msg("Błąd", "Błąd arkusza bazy.")

    def start_mass_mailing(self, _):
        if not self.full_data: self.msg("!", "Wczytaj arkusz płac!"); return
        self.stats = {"ok": 0, "fail": 0, "skip": 0, "auto": 0}
        self.queue = list(self.full_data[1:]); Clock.schedule_once(lambda dt: self.process_mailing_queue())

    def process_mailing_queue(self, *args):
        if not self.queue:
            self.msg("Raport Koncowy", f"Zakończono.\nPESEL: {self.stats['auto']}\nRęcznie: {self.stats['ok']-self.stats['auto']}")
            return
        row = self.queue.pop(0)
        try:
            n, s = str(row[self.idx_name]).strip(), str(row[self.idx_surname]).strip()
            p = str(row[self.idx_pesel]).strip() if self.idx_pesel != -1 else ""
        except: self.stats["skip"] += 1; Clock.schedule_once(lambda dt: self.process_mailing_queue()); return
        if p and len(p) > 5:
            res = self.conn.execute("SELECT email FROM contacts WHERE pesel=?", (p,)).fetchone()
            if res: self.stats["auto"] += 1; self.send_email_engine(row, res[0]); return
        res = self.conn.execute("SELECT email FROM contacts WHERE name=? AND surname=?", (n.lower(), s.lower())).fetchone()
        if res: self.ask_before_send(row, res[0], n, s, p)
        else: self.stats["skip"] += 1; Clock.schedule_once(lambda dt: self.process_mailing_queue())

    def ask_before_send(self, row, email, n, s, p_file):
        box = BoxLayout(orientation="vertical", padding=dp(20), spacing=dp(10))
        box.add_widget(Label(text=f"POTWIERDŹ:\n{n} {s}\nEmail: {email}\nPESEL: {p_file if p_file else 'BRAK'}", halign="center"))
        btns = BoxLayout(size_hint_y=None, height=dp(55), spacing=dp(10))
        def dec(v):
            pp.dismiss()
            if v: self.send_email_engine(row, email)
            else: self.stats["skip"] += 1; Clock.schedule_once(lambda dt: self.process_mailing_queue(), 0.1)
        btns.add_widget(Button(text="WYŚLIJ", on_press=lambda x: dec(True), background_color=(0,0.7,0,1))); btns.add_widget(Button(text="POMIŃ", on_press=lambda x: dec(False), background_color=(0.7,0,0,1)))
        box.add_widget(btns); pp = Popup(title="Weryfikacja", content=box, size_hint=(0.9, 0.5)); pp.open()

    def send_email_engine(self, row_data, target, fast_mode=False):
        def thread_task():
            p_conf = Path(self.user_data_dir) / "smtp.json"
            if not p_conf.exists(): return
            cfg = json.load(open(p_conf))
            try:
                srv = smtplib.SMTP("smtp.gmail.com", 587, timeout=15); srv.starttls(); srv.login(cfg['u'], cfg['p'])
                msg = EmailMessage(); dat = datetime.now().strftime("%d.%m.%Y"); ts = self.conn.execute("SELECT val FROM settings WHERE key='t_sub'").fetchone(); tb = self.conn.execute("SELECT val FROM settings WHERE key='t_body'").fetchone()
                name = row_data[self.idx_name] if (row_data and self.idx_name < len(row_data)) else "Użytkowniku"
                msg["Subject"] = (ts[0] if ts else "Informacja").replace("{Imię}", name); msg["From"], msg["To"] = cfg['u'], target; msg.set_content((tb[0] if tb else "Dzień dobry").replace("{Imię}", name).replace("{Data}", dat))
                if not fast_mode and self.full_data:
                    tmp = Path(self.user_data_dir) / "r_tmp.xlsx"; wb = Workbook(); ws = wb.active; ws.append([self.full_data[0][k] for k in self.export_indices]); ws.append([row_data[k] for k in self.export_indices]); self.style_xlsx(ws); wb.save(tmp)
                    with open(tmp, "rb") as f: msg.add_attachment(f.read(), maintype="application", subtype="xlsx", filename=f"Raport_{name}.xlsx")
                for path in self.global_attachments:
                    if os.path.exists(path):
                        ctype, _ = mimetypes.guess_type(path); main, sub = (ctype or 'application/octet-stream').split('/', 1)
                        with open(path, "rb") as f: msg.add_attachment(f.read(), maintype=main, subtype=sub, filename=os.path.basename(path))
                srv.send_message(msg); srv.quit()
                if not fast_mode: Clock.schedule_once(lambda d: [self.update_stat("ok"), self.process_mailing_queue()])
            except:
                if not fast_mode: Clock.schedule_once(lambda d: [self.update_stat("fail"), self.process_mailing_queue()])
        threading.Thread(target=thread_task, daemon=True).start()

    def setup_smtp_ui(self):
        l = BoxLayout(orientation="vertical", padding=dp(25), spacing=dp(10))
        self.ti_su = TextInput(hint_text="Gmail", multiline=False); self.ti_sp = TextInput(hint_text="Hasło", password=True)
        p = Path(self.user_data_dir) / "smtp.json"
        if p.exists():
            d = json.load(open(p)); self.ti_su.text, self.ti_sp.text = d.get('u',''), d.get('p','')
        sv = lambda x: [json.dump({'u':self.ti_su.text, 'p':self.ti_sp.text}, open(p, "w")), self.msg("OK", "Zapisano")]
        l.add_widget(Label(text="USTAWIENIA GMAIL", bold=True)); l.add_widget(self.ti_su); l.add_widget(self.ti_sp)
        l.add_widget(PremiumButton(text="ZAPISZ", on_press=sv)); l.add_widget(PremiumButton(text="POWRÓT", on_press=lambda x: setattr(self.sm, 'current', 'home'), background_color=(0.4,0.4,0.4,1)))
        self.screens["smtp"].add_widget(l)

    def setup_tmpl_ui(self):
        l = BoxLayout(orientation="vertical", padding=dp(25), spacing=dp(10))
        self.ti_ts = TextInput(hint_text="Temat {Imię}", size_hint_y=None, height=dp(45)); self.ti_tb = TextInput(hint_text="Treść...", multiline=True)
        ts = self.conn.execute("SELECT val FROM settings WHERE key='t_sub'").fetchone(); tb = self.conn.execute("SELECT val FROM settings WHERE key='t_body'").fetchone()
        if ts: self.ti_ts.text = ts[0]
        if tb: self.ti_tb.text = tb[0]
        def sv(_): [self.conn.execute("INSERT OR REPLACE INTO settings VALUES (?,?)", ('t_sub', self.ti_ts.text)), self.conn.execute("INSERT OR REPLACE INTO settings VALUES (?,?)", ('t_body', self.ti_tb.text)), self.conn.commit(), self.msg("OK", "Zapisano")]
        l.add_widget(Label(text="SZABLON MAILA", bold=True)); l.add_widget(self.ti_ts); l.add_widget(self.ti_tb)
        l.add_widget(PremiumButton(text="ZAPISZ", on_press=sv)); l.add_widget(PremiumButton(text="POWRÓT", on_press=lambda x: setattr(self.sm, 'current', 'email')))
        self.screens["tmpl"].add_widget(l)

    def setup_contacts_ui(self):
        l = BoxLayout(orientation="vertical", padding=dp(10), spacing=dp(5)); top = BoxLayout(size_hint_y=None, height=dp(55), spacing=dp(5)); self.ti_csearch = TextInput(hint_text="Szukaj...", multiline=False); self.ti_csearch.bind(text=self.refresh_contacts_list); top.add_widget(self.ti_csearch); top.add_widget(Button(text="+ Dodaj", size_hint_x=0.2, on_press=lambda x: self.form_contact())); top.add_widget(Button(text="Wróć", size_hint_x=0.2, on_press=lambda x: setattr(self.sm, 'current', 'email'))); self.c_scroll = ScrollView(); self.c_list = GridLayout(cols=1, size_hint_y=None, spacing=dp(10)); self.c_list.bind(minimum_height=self.c_list.setter('height')); self.c_scroll.add_widget(self.c_list); btn_clr = Button(text="Wyczyść wybranych", size_hint_y=None, height=dp(50), on_press=lambda x: [setattr(self, 'selected_emails', []), self.refresh_contacts_list()]); l.add_widget(top); l.add_widget(self.c_scroll); l.add_widget(btn_clr); self.screens["contacts"].add_widget(l)

    def refresh_contacts_list(self, *args):
        self.c_list.clear_widgets(); sv = self.ti_csearch.text.lower(); rows = self.conn.execute("SELECT name, surname, email, pesel, phone FROM contacts ORDER BY surname ASC").fetchall()
        for n, s, e, p, ph in rows:
            if sv and sv not in f"{n} {s} {e} {p}".lower(): continue
            row = BoxLayout(orientation="horizontal", size_hint_y=None, height=dp(115), padding=dp(8)); cb = CheckBox(size_hint_x=None, width=dp(50), active=(e in self.selected_emails)); cb.bind(active=lambda inst, v, m=e: [self.selected_emails.append(m) if v else self.selected_emails.remove(m)]); row.add_widget(cb); info = BoxLayout(orientation="vertical"); info.add_widget(Label(text=f"{n} {s}".title(), bold=True, color=(1,1,1,1), halign="left", text_size=(dp(200), None))); info.add_widget(Label(text=f"Email: {e}\nPESEL: {p if p else '---'} | Tel: {ph if ph else '---'}", font_size='12sp', color=(0.7,0.7,0.7,1), halign="left", text_size=(dp(200), None))); row.add_widget(info); acts = BoxLayout(size_hint_x=None, width=dp(90), orientation="vertical", spacing=dp(4)); acts.add_widget(Button(text="Edytuj", on_press=lambda x, d=(n,s,e,p,ph): self.form_contact(*d))); acts.add_widget(Button(text="Usuń", background_color=(0.7,0,0,1), on_press=lambda x, name=n, sur=s: self.delete_contact(name,sur))); row.add_widget(acts); self.c_list.add_widget(row)

    def form_contact(self, n="", s="", e="", pes="", ph=""):
        b = BoxLayout(orientation="vertical", padding=dp(15), spacing=dp(10)); fields = [TextInput(text=n, hint_text="Imię"), TextInput(text=s, hint_text="Nazwisko"), TextInput(text=e, hint_text="Email"), TextInput(text=pes, hint_text="PESEL"), TextInput(text=ph, hint_text="Tel")]
        for f in fields: b.add_widget(f)
        def save(_): [self.conn.execute("INSERT OR REPLACE INTO contacts VALUES (?,?,?,?,?)", [f.text.strip().lower() if i<2 else f.text.strip() for i,f in enumerate(fields)]), self.conn.commit(), p.dismiss(), self.refresh_contacts_list(), self.update_stats()]
        b.add_widget(PremiumButton(text="ZAPISZ", on_press=save)); p = Popup(title="Kontakt", content=b, size_hint=(0.9, 0.82)); p.open()

    def popup_columns(self, _):
        if not self.full_data: return
        box = BoxLayout(orientation="vertical", padding=dp(15), spacing=dp(5)); sc = ScrollView(); gr = GridLayout(cols=1, size_hint_y=None, spacing=dp(8)); gr.bind(minimum_height=gr.setter('height')); checks = []
        for i, h in enumerate(self.full_data[0]):
            r = BoxLayout(size_hint_y=None, height=dp(50)); cb = CheckBox(active=(i in self.export_indices), size_hint_x=None, width=dp(50)); checks.append((i,cb)); r.add_widget(cb); r.add_widget(Label(text=str(h) if h else f"Kol {i+1}")); gr.add_widget(r)
        def apply(_): [setattr(self, 'export_indices', [idx for idx, c in checks if c.active]), p.dismiss(), self.refresh_table()]
        sc.add_widget(gr); box.add_widget(sc); box.add_widget(PremiumButton(text="ZATWIERDŹ", on_press=apply)); p = Popup(title="Widoczność kolumn", content=box, size_hint=(0.95, 0.9)); p.open()

    def export_xlsx(self, r):
        f = Path("/storage/emulated/0/Documents/FutureExport"); f.mkdir(parents=True, exist_ok=True); wb = Workbook(); ws = wb.active; ws.append([self.full_data[0][k] for k in self.export_indices]); ws.append([r[k] for k in self.export_indices]); self.style_xlsx(ws); wb.save(f / f"Raport_{r[self.idx_name]}.xlsx"); self.msg("OK", "Zapisano raport.")

    def style_xlsx(self, ws):
        side = Side(style='thin')
        for r in ws.iter_rows():
            for c in r: c.border = Border(top=side, left=side, right=side, bottom=side); c.alignment = Alignment(horizontal='center')
        for c in ws[1]: c.font = Font(bold=True)
        for col in ws.columns: ws.column_dimensions[col[0].column_letter].width = max(len(str(x.value or "")) for x in col) + 5

    def filter_table(self, ins, val):
        v = val.lower(); self.filtered_data = [self.full_data[0]] + [r for r in self.full_data[1:] if any(v in str(c).lower() for c in r)]; self.refresh_table()

    def update_stat(self, k): self.stats[k]+=1
    def update_stats(self, *a):
        try:
            c = self.conn.execute("SELECT count(*) FROM contacts").fetchone()[0]
            self.lbl_stats.text = f"Baza: {c} | Załączniki: {len(self.global_attachments)}"
        except: pass

    def delete_contact(self, n, s):
        def pr(_): [self.conn.execute("DELETE FROM contacts WHERE name=? AND surname=?", (n, s)), self.conn.commit(), px.dismiss(), self.refresh_contacts_list(), self.update_stats()]
        px = Popup(title="Usuń?", content=Button(text="USUŃ", on_press=pr, background_color=(1,0,0,1)), size_hint=(0.7,0.3)); px.open()

    def msg(self, tit, txt):
        b = BoxLayout(orientation="vertical", padding=dp(20)); b.add_widget(Label(text=txt, halign="center")); btn = Button(text="OK", size_hint_y=None, height=dp(50), on_press=lambda x: p.dismiss()); b.add_widget(btn); p = Popup(title=tit, content=b, size_hint=(0.85, 0.45)); p.open()

if __name__ == "__main__":
    FutureApp().run()
