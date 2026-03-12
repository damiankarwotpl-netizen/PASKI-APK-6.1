import os
import json
import sqlite3
import threading
import smtplib
import mimetypes
import time
import random
from datetime import datetime
from pathlib import Path
from email.message import EmailMessage
import io
import csv
import urllib.request

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
from kivy.uix.screenmanager import ScreenManager, Screen, SlideTransition
from kivy.uix.progressbar import ProgressBar
from kivy.graphics import Color, Rectangle, RoundedRectangle

try:
    from openpyxl import load_workbook, Workbook
    from openpyxl.styles import Border, Side, Font, Alignment, PatternFill
except ImportError:
    load_workbook = Workbook = None
try:
    import xlrd
except ImportError:
    xlrd = None

COLOR_PRIMARY = (0.1, 0.5, 0.9, 1)
COLOR_BG = (0.05, 0.07, 0.1, 1)
COLOR_CARD = (0.12, 0.15, 0.2, 1)
COLOR_TEXT = (0.95, 0.95, 0.95, 1)
COLOR_ROW_A = (0.08, 0.1, 0.15, 1)
COLOR_ROW_B = (0.13, 0.16, 0.22, 1)
COLOR_HEADER = (0.1, 0.2, 0.35, 1)

class ModernButton(Button):
    def __init__(self, bg_color=COLOR_PRIMARY, **kwargs):
        super().__init__(**kwargs)
        self.background_normal = ""
        self.background_color = (0, 0, 0, 0)
        self.color = COLOR_TEXT
        self.bold = True
        self.font_size = '15sp'
        with self.canvas.before:
            Color(*bg_color)
            self.rect = RoundedRectangle(pos=self.pos, size=self.size, radius=[dp(12)])
        self.bind(pos=self._update, size=self._update)
    def _update(self, *args):
        self.rect.pos, self.rect.size = self.pos, self.size

class ModernInput(TextInput):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.background_normal = self.background_active = ""
        self.background_color = (0.15, 0.18, 0.25, 1)
        self.foreground_color = COLOR_TEXT
        self.padding = [dp(12), dp(12)]

class ColorSafeLabel(Label):
    def __init__(self, bg_color=(1,1,1,1), text_color=(1,1,1,1), **kwargs):
        super().__init__(**kwargs)
        self.color = text_color
        with self.canvas.before:
            Color(*bg_color)
            self.rect = Rectangle(size=self.size, pos=self.pos)
        self.bind(size=self._update, pos=self._update)
    def _update(self, *args):
        self.rect.size, self.rect.pos = self.size, self.pos
        self.text_size = (self.width - dp(10), None)

class HomeScreen(Screen): pass
class TableScreen(Screen): pass
class EmailScreen(Screen): pass
class SMTPScreen(Screen): pass
class TemplateScreen(Screen): pass
class ContactsScreen(Screen): pass
class ReportScreen(Screen): pass
class ClothesScreen(Screen): pass

class FutureApp(App):
    def build(self):
        Window.clearcolor = COLOR_BG
        if platform == "android":
            from android.permissions import request_permissions, Permission
            request_permissions([Permission.READ_EXTERNAL_STORAGE, Permission.WRITE_EXTERNAL_STORAGE, Permission.INTERNET])
        if not os.path.exists(self.user_data_dir): os.makedirs(self.user_data_dir, exist_ok=True)

        self.full_data, self.filtered_data, self.export_indices = [], [], []
        self.global_attachments, self.queue = [], []
        self.stats = {"ok": 0, "fail": 0, "skip": 0}
        self.idx_name, self.idx_surname, self.idx_pesel = 0, 1, -1
        self.auto_send_mode = self.is_mailing_running = False
        self.current_excel_file_path = None
        self.current_excel_filename_for_report = ""
        
        self.init_db()
        self.sm = ScreenManager(transition=SlideTransition())
        self.add_screens()
        return self.sm

    def init_db(self):
        db_p = Path(self.user_data_dir) / "future_v20.db"
        self.conn = sqlite3.connect(str(db_p), check_same_thread=False)
        self.conn.execute("CREATE TABLE IF NOT EXISTS contacts (name TEXT, surname TEXT, email TEXT, pesel TEXT, phone TEXT, company TEXT, PRIMARY KEY(name, surname))")
        try:
            self.conn.execute("ALTER TABLE contacts ADD COLUMN company TEXT")
        except sqlite3.OperationalError:
            pass
        self.conn.execute("CREATE TABLE IF NOT EXISTS settings (key TEXT PRIMARY KEY, val TEXT)")
        self.conn.execute("CREATE TABLE IF NOT EXISTS reports (id INTEGER PRIMARY KEY AUTOINCREMENT, date TEXT, ok INTEGER, fail INTEGER, skip INTEGER, auto INTEGER, original_filename TEXT, details TEXT)")
        self.conn.execute("CREATE TABLE IF NOT EXISTS clothes_types (id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT UNIQUE)")
        self.conn.execute("CREATE TABLE IF NOT EXISTS employee_clothes (contact_name TEXT, contact_surname TEXT, cloth_type_id INTEGER, size TEXT, PRIMARY KEY(contact_name, contact_surname, cloth_type_id), FOREIGN KEY(contact_name, contact_surname) REFERENCES contacts(name, surname), FOREIGN KEY(cloth_type_id) REFERENCES clothes_types(id))")
        self.conn.execute("CREATE TABLE IF NOT EXISTS clothes_issuance_log (id INTEGER PRIMARY KEY AUTOINCREMENT, contact_name TEXT, contact_surname TEXT, company TEXT, issuance_date TEXT, issued_by TEXT, issuance_type TEXT)")
        self.conn.commit()

        default_clothes = ["Koszulka", "Spodnie", "Bluza", "Kurtka", "Buty", "Czapka"]
        for item in default_clothes:
            try:
                self.conn.execute("INSERT INTO clothes_types (name) VALUES (?)", (item,))
            except sqlite3.IntegrityError:
                pass
        self.conn.commit()

    def add_screens(self):
        self.sc_ref = {name: Screen(name=name) for name in ["home", "table", "email", "smtp", "tmpl", "contacts", "report", "clothes"]}
        self.setup_ui_all()
        for s in self.sc_ref.values():
            self.sm.add_widget(s)

    def setup_ui_all(self):
        l = BoxLayout(orientation="vertical", padding=dp(30), spacing=dp(15))
        l.add_widget(Label(text="FUTURE ULTIMATE v20", font_size='34sp', bold=True, color=COLOR_PRIMARY))
        btn = lambda t, c: l.add_widget(ModernButton(text=t, on_press=c, height=dp(55), size_hint_y=None))
        btn("WCZYTAJ ARKUSZ PŁAC", lambda x: self.open_picker("data"))
        btn("PODGLĄD I EKSPORT", lambda x: [self.refresh_table(), setattr(self.sm, 'current', 'table')] if self.full_data else self.msg("!", "Wczytaj arkusz!"))
        btn("CENTRUM MAILINGOWE", lambda x: setattr(self.sm, 'current', 'email'))
        btn("RAPORTY SESJI", lambda x: [self.refresh_reports(), setattr(self.sm, 'current', 'report')])
        btn("USTAWIENIA SMTP", lambda x: setattr(self.sm, 'current', 'smtp'))
        btn("CIUCHY", lambda x: [self.refresh_employee_list_for_clothes(), setattr(self.sm, 'current', 'clothes')])
        self.sc_ref["home"].add_widget(l)
        self.setup_table_ui(); self.setup_email_ui(); self.setup_smtp_ui(); self.setup_tmpl_ui(); self.setup_contacts_ui(); self.setup_report_ui()
        self.setup_clothes_ui()

    def setup_table_ui(self):
        root = BoxLayout(orientation="vertical")
        menu = BoxLayout(size_hint_y=None, height=dp(55), spacing=dp(5), padding=dp(5))
        self.ti_tab_search = ModernInput(hint_text="Szukaj w tabeli..."); self.ti_tab_search.bind(text=self.filter_table)
        menu.add_widget(self.ti_tab_search)
        menu.add_widget(ModernButton(text="KOLUMNY", size_hint_x=0.2, on_press=self.popup_columns))
        menu.add_widget(ModernButton(text="WRÓĆ", size_hint_x=0.2, on_press=lambda x: setattr(self.sm, 'current', 'home')))
        
        hs = ScrollView(size_hint_y=None, height=dp(55), do_scroll_y=False)
        self.table_header_layout = GridLayout(rows=1, size_hint=(None, None), height=dp(55))
        hs.add_widget(self.table_header_layout)
        
        ds = ScrollView(do_scroll_x=True, do_scroll_y=True)
        self.table_content_layout = GridLayout(size_hint=(None, None))
        self.table_content_layout.bind(minimum_height=self.table_content_layout.setter('height'), minimum_width=self.table_content_layout.setter('width'))
        ds.add_widget(self.table_content_layout)
        ds.bind(scroll_x=lambda inst, val: setattr(hs, 'scroll_x', val))
        
        root.add_widget(menu); root.add_widget(hs); root.add_widget(ds)
        self.sc_ref["table"].add_widget(root)

    def refresh_table(self):
        self.table_content_layout.clear_widgets(); self.table_header_layout.clear_widgets()
        if not self.filtered_data: return
        w_cell, w_act, h = dp(170), dp(220), dp(55)
        headers = [self.full_data[0][i] for i in self.export_indices]
        
        total_w = (len(headers) * w_cell) + w_act
        self.table_header_layout.cols = self.table_content_layout.cols = len(headers) + 1
        self.table_header_layout.width = self.table_content_layout.width = total_w

        for head in headers:
            self.table_header_layout.add_widget(ColorSafeLabel(text=str(head), bg_color=COLOR_HEADER, bold=True, size=(w_cell, h), size_hint=(None,None), text_color=(0,0,0,1)))
        self.table_header_layout.add_widget(ColorSafeLabel(text="AKCJE", bg_color=COLOR_HEADER, bold=True, size=(w_act, h), size_hint=(None,None), text_color=(0,0,0,1)))

        for r_idx, row in enumerate(self.filtered_data[1:]):
            row_bg = COLOR_ROW_A if r_idx % 2 == 0 else COLOR_ROW_B
            for c_idx in self.export_indices:
                val = str(row[c_idx]) if c_idx < len(row) and str(row[c_idx]).strip() != "" else "0"
                self.table_content_layout.add_widget(ColorSafeLabel(text=val, bg_color=row_bg, size=(w_cell, h), size_hint=(None,None)))
            
            act_box = BoxLayout(size=(w_act, h), size_hint=(None,None), spacing=dp(4), padding=dp(4))
            act_box.add_widget(ModernButton(text="ZAPISZ", on_press=lambda x, r=row: self.export_single_row(r), height=dp(40), bg_color=(0.2, 0.6, 0.2, 1)))
            act_box.add_widget(ModernButton(text="WYŚLIJ", on_press=lambda x, r=row: self.send_individual_from_table(r), height=dp(40), bg_color=(0.1, 0.5, 0.9, 1)))
            self.table_content_layout.add_widget(act_box)

    def send_individual_from_table(self, row):
        name, sur = str(row[self.idx_name]).strip(), str(row[self.idx_surname]).strip()
        pes = str(row[self.idx_pesel]).strip() if self.idx_pesel != -1 else ""
        
        res = self.conn.execute("SELECT email FROM contacts WHERE pesel=? AND pesel != ''", (pes,)).fetchone() if pes else None
        if not res: res = self.conn.execute("SELECT email FROM contacts WHERE name=? AND surname=? COLLATE NOCASE", (name.lower(), sur.lower())).fetchone()
        
        if not res: return self.msg("Błąd", f"Nie znaleziono adresu dla: {name} {sur}")
        
        def task():
            cfg_p = Path(self.user_data_dir)/"smtp.json"
            if not cfg_p.exists(): return Clock.schedule_once(lambda d: self.msg("!", "Brak SMTP"))
            cfg = json.load(open(cfg_p)); srv = self.connect_smtp(cfg)
            
            session_detail_entry = {'name': name, 'surname': sur, 'email': res[0], 'status': 'OK', 'error': None}
            if not self.send_single_email(srv, cfg, row, res[0]):
                session_detail_entry['status'] = 'FAIL'
                session_detail_entry['error'] = "Błąd wysyłki indywidualnej"
            
            Clock.schedule_once(lambda d: self.msg("OK" if session_detail_entry['status'] == 'OK' else "Błąd", 
                                                f"Wysłano do: {name}" if session_detail_entry['status'] == 'OK' else f"Błąd wysyłki do: {name}"))
            srv.quit()
        threading.Thread(target=task, daemon=True).start()

    def setup_email_ui(self):
        l = BoxLayout(orientation="vertical", padding=dp(25), spacing=dp(10))
        ab = BoxLayout(size_hint_y=None, height=dp(45), spacing=dp(10))
        self.cb_auto = CheckBox(size_hint_x=None, width=dp(45)); self.cb_auto.bind(active=lambda i, v: setattr(self, 'auto_send_mode', v))
        ab.add_widget(self.cb_auto); ab.add_widget(Label(text="AUTOMATYCZNA WYSYŁKA", bold=True)); l.add_widget(ab)
        self.lbl_stats = Label(text="Baza: 0", height=dp(30)); l.add_widget(self.lbl_stats)
        l.add_widget(ModernButton(text="WYCZYŚĆ ZAŁĄCZNIKI", on_press=self.clear_all_attachments, height=dp(45), size_hint_y=None, bg_color=(0.7, 0.1, 0.1, 1)))
        self.pb_label = Label(text="Gotowy", height=dp(25)); self.pb = ProgressBar(max=100, height=dp(20)); l.add_widget(self.pb_label); l.add_widget(self.pb)
        
        btns = [("IMPORT KSIĄŻKI", lambda x: self.open_picker("book")), 
                ("AKTUALIZUJ Z GOOGLE SHEETS", self.open_google_sheet_url_popup),
                ("ZARZĄDZAJ BAZĄ", lambda x: [self.refresh_contacts_list(), setattr(self.sm, 'current', 'contacts')]), 
                ("EDYTUJ SZABLON", lambda x: setattr(self.sm, 'current', 'tmpl')), 
                ("DODAJ ZAŁĄCZNIK", lambda x: self.open_picker("attachment")), 
                ("WYŚLIJ JEDEN PLIK", self.start_special_send_flow), 
                ("START MASOWA WYSYŁKA", self.start_mass_mailing)]
        for t, c in btns: l.add_widget(ModernButton(text=t, on_press=c, height=dp(50), size_hint_y=None))
        l.add_widget(ModernButton(text="POWRÓT", on_press=lambda x: setattr(self.sm, 'current', 'home'), bg_color=(0.3,0.3,0.3,1))); self.sc_ref["email"].add_widget(l); self.update_stats()

    def process_book(self, path):
        try:
            wb = load_workbook(path, data_only=True); ws = wb.active; raw = list(ws.iter_rows(values_only=True))
            h = [str(x).lower() for x in raw[0]]; iN, iS, iE, iP, iC = 0, 1, 2, -1, -1
            for i,v in enumerate(h):
                if "imi" in v: iN=i
                elif "naz" in v: iS=i
                elif "@" in v or "mail" in v: iE=i
                elif "pesel" in v: iP=i
                elif "firma" in v or "company" in v: iC=i
            for r in raw[1:]:
                if len(r) > iE and r[iE] and "@" in str(r[iE]):
                    company_val = str(r[iC]).strip() if iC != -1 and len(r) > iC else ""
                    self.conn.execute("INSERT OR REPLACE INTO contacts VALUES (?,?,?,?,?,?)", (str(r[iN]).lower(), str(r[iS]).lower(), str(r[iE]).strip(), str(r[iP]) if iP!=-1 else "", "", company_val))
            self.conn.commit(); self.update_stats(); self.msg("OK", "Baza zaktualizowana")
        except Exception:
            self.msg("Błąd", "Nieudany import")

    def open_google_sheet_url_popup(self, _):
        box = BoxLayout(orientation="vertical", padding=dp(15), spacing=dp(10))
        url_input = ModernInput(hint_text="Wklej publiczny link do Google Sheet (CSV/XLSX)")
        box.add_widget(Label(text="LINK DO ARKUSZA GOOGLE", bold=True, size_hint_y=None, height=dp(40)))
        box.add_widget(url_input)

        def fetch_and_update(instance):
            px.dismiss()
            url = url_input.text.strip()
            if not url: return self.msg("Błąd", "URL nie może być pusty.")
            
            threading.Thread(target=self._fetch_and_process_google_sheet_thread, args=(url,), daemon=True).start()

        box.add_widget(ModernButton(text="POBIERZ I AKTUALIZUJ", on_press=fetch_and_update))
        px = Popup(title="Google Sheets Import", content=box, size_hint=(0.9, 0.5), auto_dismiss=False)
        px.open()

    def _fetch_and_process_google_sheet_thread(self, url):
        try:
            req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
            with urllib.request.urlopen(req, timeout=20) as resp:
                content_type = resp.headers.get('Content-Type', '').lower()
                data = resp.read()

            temp_path = Path(self.user_data_dir) / f"google_sheet_temp_{random.randint(1000, 9999)}"
            if url.lower().endswith((".xls", ".xlsx")) or "output=xlsx" in url.lower() or "spreadsheet" in url.lower() or "application/vnd.openxmlformats" in content_type or "application/vnd.ms-excel" in content_type:
                temp_path = temp_path.with_suffix('.xlsx')
                with open(temp_path, 'wb') as f:
                    f.write(data)
                Clock.schedule_once(lambda dt: self.process_excel(temp_path))
            else:
                temp_path = temp_path.with_suffix('.csv')
                text = data.decode('utf-8', errors='replace')
                with open(temp_path, 'w', encoding='utf-8') as f:
                    f.write(text)
                Clock.schedule_once(lambda dt: self._process_csv_from_google_sheet(temp_path))

            if temp_path.exists(): os.remove(temp_path)

        except Exception as e:
            Clock.schedule_once(lambda dt: self.msg("Błąd", f"Wystąpił błąd: {e}"))

    def _process_csv_from_google_sheet(self, path):
        try:
            with open(path, 'r', encoding='utf-8') as f:
                reader = csv.reader(f)
                raw_data = list(reader)
            
            if not raw_data:
                self.msg("Błąd", "Plik CSV jest pusty.")
                return

            h = [str(x).lower() for x in raw_data[0]]
            iN, iS, iE, iP, iC = 0, 1, 2, -1, -1
            for i,v in enumerate(h):
                if "imi" in v: iN=i
                elif "naz" in v: iS=i
                elif "@" in v or "mail" in v: iE=i
                elif "pesel" in v: iP=i
                elif "firma" in v or "company" in v: iC=i

            for r in raw_data[1:]:
                if len(r) > iE and r[iE] and "@" in str(r[iE]):
                    name = str(r[iN]).lower() if len(r) > iN else ""
                    surname = str(r[iS]).lower() if len(r) > iS else ""
                    email = str(r[iE]).strip()
                    pesel = str(r[iP]).strip() if iP != -1 and len(r) > iP else ""
                    phone = ""
                    company = str(r[iC]).strip() if iC != -1 and len(r) > iC else ""
                    self.conn.execute("INSERT OR REPLACE INTO contacts (name, surname, email, pesel, phone, company) VALUES (?,?,?,?,?,?)", (name, surname, email, pesel, phone, company))
            self.conn.commit()
            self.update_stats()
            self.msg("OK", "Kontakty z Google Sheets zaimportowane!")

        except Exception as e:
            self.msg("Błąd CSV", f"Problem z parsowaniem CSV: {e}")


    def mailing_worker(self):
        cfg_p = Path(self.user_data_dir)/"smtp.json"
        if not cfg_p.exists(): return self.finish_mailing("Brak SMTP")
        cfg = json.load(open(cfg_p)); b_on, b_sz, proc = cfg.get('batch', True), 30, 0
        
        self.session_details = [] 

        try:
            srv = self.connect_smtp(cfg)
            while self.queue:
                row_original_data = self.queue.pop(0)
                
                if isinstance(row_original_data, dict):
                    n, s, target_email = row_original_data['name'], row_original_data['surname'], row_original_data['email']
                    p_exc = row_original_data.get('pesel', '')
                    found_row = next((r for r in self.full_data[1:] if str(r[self.idx_name]).strip().lower() == n.lower() and str(r[self.idx_surname]).strip().lower() == s.lower()), None)
                    row_data_for_email = found_row if found_row else row_original_data
                else:
                    n, s = str(row_original_data[self.idx_name]).strip(), str(row_original_data[self.idx_surname]).strip()
                    p_exc = str(row_original_data[self.idx_pesel]).strip() if self.idx_pesel != -1 else ""
                    row_data_for_email = row_original_data

                    res_p = self.conn.execute("SELECT email FROM contacts WHERE pesel=? AND pesel != ''", (p_exc,)).fetchone() if p_exc else None
                    target_email, vrf = (res_p[0], False) if res_p else (None, False)
                    if not target_email:
                        res_n = self.conn.execute("SELECT email FROM contacts WHERE name=? AND surname=? COLLATE NOCASE", (n.lower(), s.lower())).fetchone()
                        if res_n: target_email, vrf = res_n[0], not self.auto_send_mode
                    
                
                if target_email:
                    if not isinstance(row_original_data, dict) and vrf:
                        self.wait_for_user = True; Clock.schedule_once(lambda dt: self.ask_before_send_worker(row_original_data, target_email, n, s))
                        while self.wait_for_user: time.sleep(0.5)
                        if self.user_decision == "skip":
                            self.stats["skip"] += 1
                            self.session_details.append({'name': n, 'surname': s, 'email': target_email, 'status': 'SKIP', 'error': 'Anulowano ręcznie'})
                            Clock.schedule_once(lambda dt: self.update_progress(self.total_q - len(self.queue)))
                            continue
                    
                    status_entry = {'name': n, 'surname': s, 'email': target_email, 'status': 'OK', 'error': None}
                    send_success, error_msg = self.send_single_email(srv, cfg, row_data_for_email, target_email)
                    
                    if send_success:
                        self.stats["ok"] += 1
                    else:
                        self.stats["fail"] += 1
                        status_entry['status'] = 'FAIL'
                        status_entry['error'] = error_msg if error_msg else "Nieznany błąd SMTP"
                        srv.quit(); srv = self.connect_smtp(cfg)
                    self.session_details.append(status_entry)
                    proc += 1
                    if self.queue:
                        if b_on and proc >= b_sz: srv.quit(); time.sleep(60); srv = self.connect_smtp(cfg); proc = 0
                        else: time.sleep(random.uniform(3, 7))
                else:
                    self.stats["skip"] += 1
                    self.session_details.append({'name': n, 'surname': s, 'email': None, 'status': 'SKIP', 'error': 'Brak w bazie kontaktów'})
                Clock.schedule_once(lambda dt: self.update_progress(self.total_q - len(self.queue)))
            srv.quit(); self.finish_mailing("Zakończono wysyłkę")
        except Exception as e: self.finish_mailing(f"Error: {e}")

    def connect_smtp(self, cfg):
        s = smtplib.SMTP(cfg.get('h','smtp.gmail.com'), int(cfg.get('port',587)), timeout=25); s.starttls(); s.login(cfg['u'], cfg['p']); return s

    def send_single_email(self, srv, cfg, row_data, target) -> (bool, str):
        try:
            nx, sx = str(row_data[self.idx_name]).title(), str(row_data[self.idx_surname]).title()
            msg = EmailMessage(); ts = self.conn.execute("SELECT val FROM settings WHERE key='t_sub'").fetchone(); tb = self.conn.execute("SELECT val FROM settings WHERE key='t_body'").fetchone()
            msg["Subject"] = (ts[0] if ts else "Raport").replace("{Imię}", nx); msg["From"], msg["To"] = cfg['u'], target
            msg.set_content((tb[0] if tb else "Dzień dobry").replace("{Imię}", nx).replace("{Data}", datetime.now().strftime("%d.%m.%Y")))
            t_f = Path(self.user_data_dir)/f"r_{nx}.xlsx"; wb = Workbook(); ws = wb.active
            ws.append([self.full_data[0][k] for k in self.export_indices]); ws.append([str(row_data[k]) if (k < len(row_data) and str(row_data[k]).strip()!="") else "0" for k in self.export_indices])
            self.style_xlsx(ws); wb.save(str(t_f))
            with open(t_f, "rb") as f: 
                ctype, _ = mimetypes.guess_type(str(t_f))
                if ctype:
                    maintype, subtype = ctype.split('/',1)
                else:
                    maintype, subtype = 'application', 'octet-stream'
                msg.add_attachment(f.read(), maintype=maintype, subtype=subtype, filename=f"Raport_{nx}_{sx}.xlsx")
            for p in self.global_attachments:
                if os.path.exists(p):
                    ct, _ = mimetypes.guess_type(p); mn, sb = (ct or 'application/octet-stream').split('/', 1)
                    with open(p,"rb") as f: msg.add_attachment(f.read(), maintype=mn, subtype=sb, filename=os.path.basename(p))
            srv.send_message(msg); return True, None
        except Exception as e: 
            return False, str(e)

    def style_xlsx(self, ws):
        s, c = Side(style='thin'), Alignment(horizontal='center', vertical='center')
        for ri, row in enumerate(ws.iter_rows(), 1):
            for cell in row:
                cell.border = Border(top=s, left=s, right=s, bottom=s); cell.alignment = c
                if ri == 1: cell.font = Font(bold=True); cell.fill = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")
                elif ri % 2 == 0: cell.fill = PatternFill(start_color="F7F7F7", end_color="F7F7F7", fill_type="solid")
        for col in ws.columns:
            m = 0; col_let = col[0].column_letter
            for cell in col:
                if cell.value: m = max(m, len(str(cell.value)))
            ws.column_dimensions[col_let].width = (m * 1.3) + 7

    def setup_smtp_ui(self):
        l = BoxLayout(orientation="vertical", padding=dp(25), spacing=dp(8)); p = Path(self.user_data_dir)/"smtp.json"; d = json.load(open(p)) if p.exists() else {}
        self.ti_h, self.ti_pt = ModernInput(hint_text="Host", text=d.get('h','')), ModernInput(hint_text="Port", text=str(d.get('port','587')))
        self.ti_u, self.ti_p = ModernInput(hint_text="Email/Login", text=d.get('u','')), ModernInput(hint_text="Hasło/Klucz", password=True, text=d.get('p',''))
        l.add_widget(Label(text="USTAWIENIA POCZTY", bold=True)); l.add_widget(self.ti_h); l.add_widget(self.ti_pt); l.add_widget(self.ti_u); l.add_widget(self.ti_p)
        bx = BoxLayout(size_hint_y=None, height=dp(45)); self.cb_b = CheckBox(size_hint_x=None, width=dp(45), active=d.get('batch', True)); bx.add_widget(self.cb_b); bx.add_widget(Label(text="Batching (przerwa 60s/30 maili)")); l.add_widget(bx)
        l.add_widget(ModernButton(text="ZAPISZ KONFIGURACJĘ", on_press=lambda x: [json.dump({'h':self.ti_h.text,'port':self.ti_pt.text,'u':self.ti_u.text,'p':self.ti_p.text,'batch':self.cb_b.active}, open(p,"w")), self.msg("OK","Zapisano")]))
        l.add_widget(ModernButton(text="TEST POŁĄCZENIA", on_press=lambda x: self.test_smtp_direct(), bg_color=(.1,.7,.4,1)))
        l.add_widget(ModernButton(text="POWRÓT", on_press=lambda x: setattr(self.sm,'current','home'), bg_color=(.3,.3,.3,1)))
        self.sc_ref["smtp"].add_widget(l)

    def test_smtp_direct(self):
        try: s = self.connect_smtp({'h':self.ti_h.text,'port':self.ti_pt.text,'u':self.ti_u.text,'p':self.ti_p.text}); s.quit(); self.msg("OK", "Serwer SMTP Działa!")
        except Exception as e: self.msg("BŁĄD", str(e)[:60])

    def clear_all_attachments(self, _=None): 
        self.global_attachments.clear(); self.update_stats()

    def start_special_send_flow(self, _): self.open_picker("special_send")
    def special_send_step_2(self, path):
        self.selected_emails = []
        box = BoxLayout(orientation="vertical", padding=dp(15), spacing=dp(10))
        ti = ModernInput(hint_text="Szukaj...")
        box.add_widget(ti)
        sc = ScrollView()
        gl = GridLayout(cols=1, size_hint_y=None, spacing=dp(5))
        gl.bind(minimum_height=gl.setter('height'))
        sc.add_widget(gl)
        box.add_widget(sc)
        def rf(v=""):
            gl.clear_widgets()
            rows = self.conn.execute("SELECT name, surname, email FROM contacts").fetchall()
            for r in rows:
                if v and v.lower() not in f"{r[0]} {r[1]} {r[2]}".lower(): continue
                bx = BoxLayout(size_hint_y=None, height=dp(50))
                cb = CheckBox(size_hint_x=None, width=dp(50))
                cb.bind(active=lambda inst, val, m=r[2]: self.selected_emails.append(m) if val else self.selected_emails.remove(m))
                bx.add_widget(cb)
                bx.add_widget(Label(text=f"{r[0].title()} {r[1].title()}"))
                gl.add_widget(bx)
        ti.bind(text=lambda i,v: rf(v))
        rf()
        btn = ModernButton(text="DALEJ", on_press=lambda x: [p.dismiss(), self.special_send_step_3(path)] if self.selected_emails else None)
        box.add_widget(btn)
        p = Popup(title="Odbiorcy", content=box, size_hint=(.95,.9))
        p.open()

    def special_send_step_3(self, path):
        b = BoxLayout(orientation="vertical", padding=dp(15), spacing=dp(10)); ti_s = ModernInput(hint_text="Temat"); ti_b = ModernInput(hint_text="Treść", multiline=True)
        b.add_widget(ti_s); b.add_widget(ti_b)
        def run(_):
            def task():
                cfg = json.load(open(Path(self.user_data_dir)/"smtp.json")); srv = self.connect_smtp(cfg)
                for m in self.selected_emails:
                    msg = EmailMessage(); msg["Subject"], msg["From"], msg["To"] = ti_s.text, cfg['u'], m; msg.set_content(ti_b.text)
                    with open(path, "rb") as f: msg.add_attachment(f.read(), maintype="application", subtype="octet-stream", filename=os.path.basename(path))
                    srv.send_message(msg)
                srv.quit(); Clock.schedule_once(lambda d: self.msg("OK", "Wysłano"))
            threading.Thread(target=task, daemon=True).start(); p.dismiss()
        b.add_widget(ModernButton(text="WYŚLIJ PLIK", on_press=run)); p = Popup(title="Wiadomość", content=b, size_hint=(.9, .8)); p.open()

    def filter_table(self, i, v): 
        self.filtered_data = [self.full_data[0]] + [r for r in self.full_data[1:] if any(v.lower() in str(c).lower() for c in r)]
        self.refresh_table()
    
    def start_mass_mailing(self, _):
        if self.is_mailing_running: return self.msg("!", "Wysyłka już trwa!")
        if not self.full_data: return self.msg("!", "Wczytaj arkusz płac najpierw!")

        self.stats, self.session_details, self.queue = {"ok": 0, "fail": 0, "skip": 0}, [], list(self.full_data[1:])
        self.total_q = len(self.queue)
        self.is_mailing_running = True
        self.current_excel_filename_for_report = Path(self.current_excel_file_path).name if self.current_excel_file_path else "Nieznany plik"
        
        threading.Thread(target=self.mailing_worker, daemon=True).start()

    def open_picker(self, mode):
        if platform != "android": return self.msg("!", "Tylko Android")
        from jnius import autoclass; from android import activity
        PA, Intent = autoclass("org.kivy.android.PythonActivity"), autoclass("android.content.Intent"); intent = Intent(Intent.ACTION_GET_CONTENT); intent.setType("*/*")
        if mode == "attachment": intent.putExtra(Intent.EXTRA_ALLOW_MULTIPLE, True)
        def cb(req, res, dt):
            if req != 1001: return
            activity.unbind(on_activity_result=cb)
            if res == -1 and dt:
                resolver = PA.mActivity.getContentResolver(); files = []
                clip = dt.getClipData()
                if clip: [files.append(clip.getItemAt(i).getUri()) for i in range(clip.getItemCount())]
                else: files.append(dt.getData())
                for uri in files:
                    cur = resolver.query(uri, None, None, None, None); name = f"f_{random.randint(10,99)}.xlsx"
                    if cur and cur.moveToFirst(): idx = cur.getColumnIndex("_display_name"); name = cur.getString(idx) if idx != -1 else name; cur.close()
                    try:
                        stream, loc = resolver.openInputStream(uri), Path(self.user_data_dir) / name
                        with open(loc, "wb") as f:
                            buf = bytearray(16384)
                            while True:
                                n = stream.read(buf)
                                if not n: break
                                f.write(buf[:n])
                        stream.close()
                        if mode == "data": self.process_excel(loc)
                        elif mode == "book": self.process_book(loc)
                        elif mode == "attachment": self.global_attachments.append(str(loc))
                        elif mode == "special_send": Clock.schedule_once(lambda dt: self.special_send_step_2(str(loc)))
                    except Exception as e:
                        Clock.schedule_once(lambda dt: self.msg("Błąd pliku", f"Nie udało się przetworzyć pliku: {e}"))
                self.update_stats()
        activity.bind(on_activity_result=cb); PA.mActivity.startActivityForResult(intent, 1001)

    def process_excel(self, path):
        try:
            self.current_excel_file_path = str(path)

            if str(path).endswith(".xls") and xlrd:
                wb = xlrd.open_workbook(path); ws = wb.sheet_by_index(0); raw = [[str(ws.cell_value(r,c)).strip() for c in range(ws.ncols)] for r in range(ws.nrows)]
            else:
                wb = load_workbook(path, data_only=True); ws = wb.active; raw = [["" if v is None else str(v).strip() for v in r] for r in ws.iter_rows(values_only=True)]
            h_idx = 0
            for i, r in enumerate(raw[:15]):
                row_text = " ".join([str(v) for v in r]).lower()
                if any(x in row_text for x in ["imię", "imie", "nazwisko"]):
                    h_idx = i
                    break
            self.full_data, self.export_indices = raw[h_idx:], list(range(len(raw[h_idx][0])))
            self.filtered_data = self.full_data

            filename = Path(path).name
            saved_indices_json = self.conn.execute("SELECT val FROM settings WHERE key=?", (f"columns_{filename}",)).fetchone()
            if saved_indices_json and saved_indices_json[0]:
                try:
                    loaded_indices = json.loads(saved_indices_json[0])
                    valid_indices = [idx for idx in loaded_indices if idx < len(self.full_data[0])]
                    if valid_indices:
                        self.export_indices = valid_indices
                except json.JSONDecodeError:
                    pass

            for i,v in enumerate(self.full_data[0]):
                v = str(v).lower()
                if "imi" in v: self.idx_name = i
                if "naz" in v: self.idx_surname = i
                if "pesel" in v: self.idx_pesel = i
            self.msg("OK", "Arkusz wczytany")
        except Exception as e:
            self.msg("BŁĄD", f"Plik uszkodzony lub problem z przetwarzaniem: {e}")

    def setup_tmpl_ui(self):
        l, ti_s, ti_b = BoxLayout(orientation="vertical", padding=dp(25), spacing=dp(10)), ModernInput(hint_text="Temat {Imię}"), ModernInput(hint_text="Treść...", multiline=True)
        ts = self.conn.execute("SELECT val FROM settings WHERE key='t_sub'").fetchone(); tb = self.conn.execute("SELECT val FROM settings WHERE key='t_body'").fetchone()
        ti_s.text, ti_b.text = (ts[0] if ts else ""), (tb[0] if tb else "")
        l.add_widget(Label(text="SZABLON EMAIL", bold=True)); l.add_widget(ti_s); l.add_widget(ti_b)
        l.add_widget(ModernButton(text="ZAPISZ", on_press=lambda x: [self.conn.execute("INSERT OR REPLACE INTO settings VALUES (?,?)", ('t_sub',ti_s.text)), self.conn.execute("INSERT OR REPLACE INTO settings VALUES (?,?)", ('t_body',ti_b.text)), self.conn.commit(), self.msg("OK","Wzór zapisany")]))
        l.add_widget(ModernButton(text="POWRÓT", on_press=lambda x: setattr(self.sm, 'current', 'email'))); self.sc_ref["tmpl"].add_widget(l)

    def setup_contacts_ui(self):
        l, top = BoxLayout(orientation="vertical", padding=dp(10)), BoxLayout(size_hint_y=None, height=dp(55), spacing=dp(5))
        self.ti_cs = ModernInput(hint_text="Szukaj..."); self.ti_cs.bind(text=self.refresh_contacts_list); top.add_widget(self.ti_cs)
        top.add_widget(ModernButton(text="+", size_hint_x=0.15, on_press=lambda x: self.form_contact())); top.add_widget(ModernButton(text="Wróć", size_hint_x=0.2, on_press=lambda x: setattr(self.sm, 'current', 'email')))
        self.c_ls = GridLayout(cols=1, size_hint_y=None, spacing=dp(10)); self.c_ls.bind(minimum_height=self.c_ls.setter('height'))
        sc = ScrollView(); sc.add_widget(self.c_ls); l.add_widget(top); l.add_widget(sc); self.sc_ref["contacts"].add_widget(l)

    def refresh_contacts_list(self, *args):
        self.c_ls.clear_widgets(); sv = self.ti_cs.text.lower() if self.ti_cs else ""
        rows = self.conn.execute("SELECT name, surname, email, pesel, phone, company FROM contacts ORDER BY surname ASC").fetchall()
        for d in rows:
            if sv and sv not in f"{d[0]} {d[1]} {d[2]} {d[5]}".lower(): continue
            r = BoxLayout(size_hint_y=None, height=dp(125), padding=dp(10))
            with r.canvas.before: Color(*COLOR_CARD); Rectangle(pos=r.pos, size=r.size)
            inf, acts = BoxLayout(orientation="vertical"), BoxLayout(size_hint_x=0.3, orientation="vertical", spacing=dp(4))
            inf.add_widget(Label(text=f"{d[0]} {d[1]}".title(), bold=True, halign="left", text_size=(dp(250),None)))
            inf.add_widget(Label(text=f"E: {d[2]}\nP: {d[3]}\nT: {d[4] if d[4] else '-'}\nFirma: {d[5] if d[5] else '-'}", font_size='11sp', halign="left", text_size=(dp(250),None), color=(0.7,0.7,0.7,1)))
            r.add_widget(inf); acts.add_widget(ModernButton(text="Edytuj", on_press=lambda x, data=d: self.form_contact(*data))); acts.add_widget(ModernButton(text="Usuń", background_color=(0.8,0.2,0.2,1), on_press=lambda x, n=d[0], s=d[1]: self.delete_contact(n, s))); r.add_widget(acts); self.c_ls.add_widget(r)

    def msg(self, tit, txt):
        b = BoxLayout(orientation="vertical", padding=dp(20)); b.add_widget(Label(text=txt, halign="center")); b.add_widget(ModernButton(text="OK", on_press=lambda x: p.dismiss(), height=dp(50), size_hint_y=None)); p = Popup(title=tit, content=b, size_hint=(0.85, 0.45)); p.open()
    def update_stats(self, *a):
        try: self.lbl_stats.text = f"Baza: {self.conn.execute('SELECT count(*) FROM contacts').fetchone()[0]} | Załączniki: {len(self.global_attachments)}"
        except: pass
    def update_progress(self, d): 
        try:
            self.pb.value = int((d/self.total_q)*100)
            self.pb_label.text = f"Postęp: {d}/{self.total_q}"
        except Exception:
            pass
    
    def finish_mailing(self, s): 
        self.is_mailing_running = False
        
        det_json = json.dumps(self.session_details)
        self.conn.execute("INSERT INTO reports (date, ok, fail, skip, auto, original_filename, details) VALUES (?,?,?,?,?,?,?)", 
                          (datetime.now().strftime("%Y-%m-%d %H:%M"), self.stats['ok'], self.stats['fail'], 
                           self.stats['skip'], 0, self.current_excel_filename_for_report, det_json))
        self.conn.commit()
        Clock.schedule_once(lambda dt: self.msg("Mailing", f"{s}\nSukces: {self.stats['ok']} | Błędy: {self.stats['fail']} | Pominięto: {self.stats['skip']}"))

    def popup_columns(self, _):
        if not self.full_data: return self.msg("!", "Wczytaj arkusz najpierw!")

        box, gr, checks = BoxLayout(orientation="vertical", padding=dp(10)), GridLayout(cols=1, size_hint_y=None, spacing=dp(5)), []
        gr.bind(minimum_height=gr.setter('height'))
        for i, h in enumerate(self.full_data[0]):
            r, cb = BoxLayout(size_hint_y=None, height=dp(45)), CheckBox(active=(i in self.export_indices), size_hint_x=None, width=dp(50)); checks.append((i, cb)); r.add_widget(cb); r.add_widget(Label(text=str(h))); gr.add_widget(r)
        sc = ScrollView(); sc.add_widget(gr); box.add_widget(sc)
        
        def ask_to_save_columns(instance_of_apply_button):
            new_export_indices = [idx for idx, c in checks if c.active]
            
            p.dismiss()

            save_box = BoxLayout(orientation="vertical", padding=dp(20), spacing=dp(10))
            save_box.add_widget(Label(text="Czy zapamiętać te kolumny dla tego pliku?", halign="center"))
            
            btns_layout = BoxLayout(size_hint_y=None, height=dp(55), spacing=dp(10))
            
            def save_and_apply(instance):
                if self.current_excel_file_path:
                    filename = Path(self.current_excel_file_path).name
                    self.conn.execute("INSERT OR REPLACE INTO settings (key, val) VALUES (?,?)", (f"columns_{filename}", json.dumps(new_export_indices)))
                    self.conn.commit()
                self.apply_columns_and_refresh(new_export_indices)
                save_popup.dismiss()

            def apply_only(instance):
                self.apply_columns_and_refresh(new_export_indices)
                save_popup.dismiss()

            btns_layout.add_widget(ModernButton(text="Tak, zapamiętaj", on_press=save_and_apply, height=dp(45), bg_color=(0,0.6,0,1)))
            btns_layout.add_widget(ModernButton(text="Nie, tylko zastosuj", on_press=apply_only, height=dp(45), bg_color=(0.7,0,0,1)))
            save_box.add_widget(btns_layout)
            
            save_popup = Popup(title="Zapisz kolumny?", content=save_box, size_hint=(0.9, 0.45), auto_dismiss=False)
            save_popup.open()

        box.add_widget(ModernButton(text="ZASTOSUJ", on_press=ask_to_save_columns, height=dp(50), size_hint_y=None)); 
        p = Popup(title="Kolumny", content=box, size_hint=(0.9, 0.9), auto_dismiss=False); p.open()

    def apply_columns_and_refresh(self, new_export_indices):
        self.export_indices = new_export_indices
        self.refresh_table()

    def setup_report_ui(self):
        l, self.r_grid = BoxLayout(orientation="vertical", padding=dp(15), spacing=dp(10)), GridLayout(cols=1, size_hint_y=None, spacing=dp(10))
        self.r_grid.bind(minimum_height=self.r_grid.setter('height')); sc = ScrollView(); sc.add_widget(self.r_grid); l.add_widget(Label(text="HISTORIA SESJI", bold=True, height=dp(40), size_hint_y=None)); l.add_widget(sc); l.add_widget(ModernButton(text="POWRÓT", on_press=lambda x: setattr(self.sm, 'current', 'home'), height=dp(55), size_hint_y=None)); self.sc_ref["report"].add_widget(l)
    
    def refresh_reports(self, *a):
        self.r_grid.clear_widgets(); 
        rows = self.conn.execute("SELECT id, date, ok, fail, skip, original_filename, details FROM reports ORDER BY id DESC").fetchall()
        for r_id, d, ok, fl, sk, filename, det_json in rows:
            row = BoxLayout(orientation="vertical", size_hint_y=None, height=dp(140), padding=dp(10))
            with row.canvas.before: Color(0.15, 0.2, 0.25, 1); Rectangle(pos=row.pos, size=row.size)
            row.add_widget(Label(text=f"Sesja: {d}\nPlik: {filename}", bold=True, color=COLOR_PRIMARY, halign="left", text_size=(dp(300), None)))
            row.add_widget(Label(text=f"OK: {ok} | Błędy: {fl} | Pominięto: {sk}", font_size='13sp', halign="left", text_size=(dp(300), None)))

            btn_box = BoxLayout(size_hint_y=None, height=dp(38), spacing=dp(5))
            btn_box.add_widget(ModernButton(text="Pokaż logi", on_press=lambda x, t=det_json: self.show_details(t)))
            
            has_fails = False
            if det_json:
                try:
                    details_list = json.loads(det_json)
                    if any(entry['status'] == 'FAIL' for entry in details_list):
                        has_fails = True
                except json.JSONDecodeError:
                    pass
            
            if has_fails:
                btn_box.add_widget(ModernButton(text="Ponów nieudane", background_color=(0.8, 0.4, 0.1, 1), on_press=lambda x, report_id=r_id, orig_filename=filename: self.start_retry_mailing(report_id, orig_filename)))
            else:
                btn_box.add_widget(ModernButton(text="Brak nieudanych", background_color=(0.5, 0.5, 0.5, 1), height=dp(38)))
            row.add_widget(btn_box)
            self.r_grid.add_widget(row)

    def show_details(self, t):
        b = BoxLayout(orientation="vertical", padding=dp(10)); ti = TextInput(text=str(t), readonly=True, font_size='11sp'); b.add_widget(ti); b.add_widget(ModernButton(text="ZAMKNIJ", on_press=lambda x: p.dismiss(), size_hint_y=0.2)); p = Popup(title="Logi", content=b, size_hint=(.9,.8)); p.open()
    
    def start_retry_mailing(self, report_id, original_filename):
        if self.is_mailing_running: return self.msg("!", "Wysyłka już trwa!")
        if not self.full_data: return self.msg("!", "Wczytaj oryginalny arkusz płac najpierw!")
        
        if not self.current_excel_file_path or Path(self.current_excel_file_path).name != original_filename:
            return self.msg("Błąd", f"Proszę wczytać oryginalny plik: [b]{original_filename}[/b], aby ponowić wysyłkę.")

        report = self.conn.execute("SELECT details FROM reports WHERE id=?", (report_id,)).fetchone()
        if not report or not report[0]:
            return self.msg("Błąd", "Nie znaleziono szczegółów raportu.")

        try:
            details_list = json.loads(report[0])
            failed_recipients = [entry for entry in details_list if entry['status'] == 'FAIL']

            if not failed_recipients:
                return self.msg("Informacja", "Brak nieudanych wysyłek do ponowienia w tym raporcie.")

            self.stats = {"ok": 0, "fail": 0, "skip": 0}
            self.session_details = []
            self.queue = list(failed_recipients)
            self.total_q = len(self.queue)
            self.is_mailing_running = True
            self.current_excel_filename_for_report = f"Ponowienie ({original_filename})"
            
            threading.Thread(target=self.mailing_worker, daemon=True).start()
            self.msg("OK", f"Rozpoczęto ponowną wysyłkę do {len(failed_recipients)} osób.")

        except json.JSONDecodeError:
            self.msg("Błąd", "Błąd w formacie logów raportu.")
        except Exception as e:
            self.msg("Błąd", f"Wystąpił błąd podczas ponawiania: {e}")

    def ask_before_send_worker(self, row, email, n, s):
        def dec(v): self.user_decision = "send" if v else "skip"; self.wait_for_user = False; px.dismiss()
        box = BoxLayout(orientation="vertical", padding=dp(20), spacing=dp(10)); box.add_widget(Label(text=f"POTWIERDŹ:\n[b]{n} {s}[/b]\n{email}", markup=True, halign="center"))
        btns = BoxLayout(size_hint_y=None, height=dp(55), spacing=dp(10)); btns.add_widget(ModernButton(text="WYŚLIJ", on_press=lambda x: dec(True), height=dp(45), bg_color=(0,0.6,0,1))); btns.add_widget(ModernButton(text="POMIŃ", on_press=lambda x: dec(False), height=dp(45), bg_color=(0.7,0,0,1)))
        box.add_widget(btns); px = Popup(title="Weryfikacja", content=box, size_hint=(0.9, 0.45), auto_dismiss=False); px.open()
    def export_single_row(self, r):
        p = Path("/storage/emulated/0/Documents/FutureExport") if platform=="android" else Path("./exports"); p.mkdir(parents=True, exist_ok=True)
        nx, sx = str(r[self.idx_name]).title(), str(r[self.idx_surname]).title(); wb = Workbook(); ws = wb.active
        ws.append([self.full_data[0][k] for k in self.export_indices]); ws.append([str(r[k]) if (k < len(r) and str(r[k]).strip() != "") else "0" for k in self.export_indices])
        self.style_xlsx(ws); wb.save(str(p/f"Raport_{nx}_{sx}.xlsx")); self.msg("OK", f"Zapisano PDF dla: {nx}")
    def delete_contact(self, n, s):
        def pr(_):
            self.conn.execute("DELETE FROM contacts WHERE name=? AND surname=?", (n, s))
            self.conn.execute("DELETE FROM employee_clothes WHERE contact_name=? AND contact_surname=?", (n, s))
            self.conn.commit(); px.dismiss(); self.refresh_contacts_list(); self.update_stats()
        px = Popup(title="Usuń?", content=ModernButton(text="USUŃ KONTAKT", on_press=pr, height=dp(50), bg_color=(1,0,0,1)), size_hint=(0.7,0.3)); px.open()
    def form_contact(self, n="", s="", e="", pes="", ph="", company=""):
        b = BoxLayout(orientation="vertical", padding=dp(15), spacing=dp(10))
        f_ins = [
            ModernInput(text=str(n), hint_text="Imię"),
            ModernInput(text=str(s), hint_text="Nazwisko"),
            ModernInput(text=str(e), hint_text="Email"),
            ModernInput(text=str(pes), hint_text="PESEL"),
            ModernInput(text=str(ph), hint_text="Telefon"),
            ModernInput(text=str(company), hint_text="Firma")
        ]
        for f in f_ins: b.add_widget(f)
        def save(_):
            self.conn.execute("INSERT OR REPLACE INTO contacts VALUES (?,?,?,?,?,?)", (
                f_ins[0].text.lower(),
                f_ins[1].text.lower(),
                f_ins[2].text.strip(),
                f_ins[3].text.strip(),
                f_ins[4].text.strip(),
                f_ins[5].text.strip()
            ))
            self.conn.commit(); px.dismiss(); self.refresh_contacts_list(); self.update_stats()
        b.add_widget(ModernButton(text="ZAPISZ", on_press=save)); px = Popup(title="Kontakt", content=b, size_hint=(0.9, 0.85)); px.open()

    def setup_clothes_ui(self):
        l = BoxLayout(orientation="vertical", padding=dp(10))
        top = BoxLayout(size_hint_y=None, height=dp(55), spacing=dp(5))
        self.ti_clothes_search = ModernInput(hint_text="Szukaj pracownika (imię, nazwisko, firma)...")
        self.ti_clothes_search.bind(text=self.refresh_employee_list_for_clothes)
        top.add_widget(self.ti_clothes_search)
        top.add_widget(ModernButton(text="WRÓĆ", size_hint_x=0.2, on_press=lambda x: setattr(self.sm, 'current', 'home')))
        l.add_widget(top)

        self.clothes_employee_list = GridLayout(cols=1, size_hint_y=None, spacing=dp(10))
        self.clothes_employee_list.bind(minimum_height=self.clothes_employee_list.setter('height'))
        sc = ScrollView(); sc.add_widget(self.clothes_employee_list); l.add_widget(sc)

        self.sc_ref["clothes"].add_widget(l)

    def refresh_employee_list_for_clothes(self, *args):
        self.clothes_employee_list.clear_widgets()
        search_val = (self.ti_clothes_search.text or "").lower()
        query = "SELECT name, surname, company FROM contacts ORDER BY surname ASC"
        employees = self.conn.execute(query).fetchall()

        for name, surname, company in employees:
            full_name = f"{name.title()} {surname.title()}"
            company_display = f"Firma: {company}" if company else "Brak firmy"
            if search_val and search_val not in f"{name} {surname} {company}".lower():
                continue

            row = BoxLayout(size_hint_y=None, height=dp(120), padding=dp(10))
            with row.canvas.before: Color(*COLOR_CARD); Rectangle(pos=row.pos, size=row.size)
            
            info_layout = BoxLayout(orientation="vertical")
            info_layout.add_widget(Label(text=full_name, bold=True, halign="left", text_size=(dp(200),None)))
            info_layout.add_widget(Label(text=company_display, font_size='11sp', halign="left", text_size=(dp(200),None), color=(0.7,0.7,0.7,1)))
            row.add_widget(info_layout)

            actions_layout = BoxLayout(size_hint_x=0.4, orientation="vertical", spacing=dp(5))
            actions_layout.add_widget(ModernButton(text="Edytuj Rozmiary", on_press=lambda x, n=name, s=surname: self.show_edit_clothes_popup(n, s), height=dp(44)))
            actions_layout.add_widget(ModernButton(text="Wydaj Ciuchy", on_press=lambda x, n=name, s=surname, c=company: self.show_issue_clothes_popup(n, s, c), height=dp(44)))
            row.add_widget(actions_layout)

            self.clothes_employee_list.add_widget(row)

    def show_edit_clothes_popup(self, employee_name, employee_surname):
        box = BoxLayout(orientation="vertical", padding=dp(15), spacing=dp(10))
        box.add_widget(Label(text=f"Edytuj rozmiary dla: [b]{employee_name.title()} {employee_surname.title()}[/b]", markup=True, size_hint_y=None, height=dp(40)))

        scroll_view = ScrollView()
        grid = GridLayout(cols=2, size_hint_y=None, spacing=dp(5))
        grid.bind(minimum_height=grid.setter('height'))
        
        cloth_inputs = {}

        clothes_types = self.conn.execute("SELECT id, name FROM clothes_types ORDER BY name").fetchall()
        
        current_sizes_rows = self.conn.execute("SELECT cloth_type_id, size FROM employee_clothes WHERE contact_name=? AND contact_surname=?", (employee_name, employee_surname)).fetchall()
        current_sizes = {row[0]: row[1] for row in current_sizes_rows}

        for c_id, c_name in clothes_types:
            grid.add_widget(Label(text=c_name, halign="left", text_size=(dp(150), None), size_hint_x=None, width=dp(150)))
            text_input = ModernInput(text=current_sizes.get(c_id, ''), hint_text="Rozmiar", multiline=False, size_hint_x=None, width=dp(150))
            cloth_inputs[c_id] = text_input
            grid.add_widget(text_input)
        
        scroll_view.add_widget(grid)
        box.add_widget(scroll_view)

        def save_sizes(instance):
            for c_id, text_input in cloth_inputs.items():
                size_val = text_input.text.strip()
                if size_val:
                    self.conn.execute("INSERT OR REPLACE INTO employee_clothes (contact_name, contact_surname, cloth_type_id, size) VALUES (?,?,?,?)",
                                      (employee_name, employee_surname, c_id, size_val))
                else:
                    self.conn.execute("DELETE FROM employee_clothes WHERE contact_name=? AND contact_surname=? AND cloth_type_id=?",
                                      (employee_name, employee_surname, c_id))
            self.conn.commit()
            self.msg("OK", "Rozmiary zapisane!")
            popup.dismiss()
            self.refresh_employee_list_for_clothes()

        box.add_widget(ModernButton(text="ZAPISZ", on_press=save_sizes, height=dp(50), size_hint_y=None))
        popup = Popup(title="Edycja rozmiarów", content=box, size_hint=(0.9, 0.85), auto_dismiss=False)
        popup.open()

    def show_issue_clothes_popup(self, employee_name, employee_surname, employee_company):
        box = BoxLayout(orientation="vertical", padding=dp(15), spacing=dp(10))
        box.add_widget(Label(text=f"Wydanie ciuchów dla: [b]{employee_name.title()} {employee_surname.title()}[/b]", markup=True, size_hint_y=None, height=dp(40)))
        
        box.add_widget(Label(text=f"Firma: [b]{employee_company if employee_company else 'Brak firmy'}[/b]", markup=True, size_hint_y=None, height=dp(30), halign="left", text_size=(dp(300), None)))

        issue_type_label = Label(text="Typ wydania:", size_hint_y=None, height=dp(30), halign="left", text_size=(dp(300), None))
        box.add_widget(issue_type_label)

        radio_layout = BoxLayout(size_hint_y=None, height=dp(40), spacing=dp(10))
        
        cb_person = CheckBox(group='issue_type', active=True, size_hint_x=None, width=dp(40))
        cb_person_label = Label(text="Dla osoby", halign="left", text_size=(dp(100), None), size_hint_x=None, width=dp(100))
        
        cb_company = CheckBox(group='issue_type', active=False, size_hint_x=None, width=dp(40))
        cb_company_label = Label(text="Dla firmy", halign="left", text_size=(dp(100), None), size_hint_x=None, width=dp(100))

        radio_layout.add_widget(cb_person)
        radio_layout.add_widget(cb_person_label)
        radio_layout.add_widget(cb_company)
        radio_layout.add_widget(cb_company_label)
        box.add_widget(radio_layout)

        issued_by_input = ModernInput(hint_text="Kto wydał (imię, nazwisko)", multiline=False)
        box.add_widget(Label(text="Kto wydał:", size_hint_y=None, height=dp(30), halign="left", text_size=(dp(300), None)))
        box.add_widget(issued_by_input)

        def confirm_issue(instance):
            issuance_type = 'person' if cb_person.active else 'company'
            issued_by_val = issued_by_input.text.strip()
            
            if not issued_by_val:
                self.msg("Błąd", "Wprowadź, kto wydał ubrania.")
                return

            self.issue_clothes(employee_name, employee_surname, employee_company, issuance_type, issued_by_val)
            popup.dismiss()

        box.add_widget(ModernButton(text="POTWIERDŹ WYDANIE", on_press=confirm_issue, height=dp(50), size_hint_y=None))
        popup = Popup(title="Wydanie ubrań", content=box, size_hint=(0.9, 0.7), auto_dismiss=False)
        popup.open()

    def issue_clothes(self, name, surname, company, issuance_type, issued_by):
        current_date = datetime.now().strftime("%Y-%m-%d %H:%M")
        
        if issuance_type == 'person':
            self.conn.execute("INSERT INTO clothes_issuance_log (contact_name, contact_surname, company, issuance_date, issued_by, issuance_type) VALUES (?,?,?,?,?,?)",
                              (name, surname, company, current_date, issued_by, 'person'))
            self.conn.commit()
            self.msg("OK", f"Wydano ubrania dla: {name.title()} {surname.title()}")
        elif issuance_type == 'company':
            self.conn.execute("INSERT INTO clothes_issuance_log (contact_name, contact_surname, company, issuance_date, issued_by, issuance_type) VALUES (?,?,?,?,?,?)",
                              (None, None, company, current_date, issued_by, 'company'))
            self.conn.commit()
            self.msg("OK", f"Wydano ubrania dla firmy: {company}")
        
        self.refresh_employee_list_for_clothes()

if __name__ == "__main__":
    FutureApp().run()
