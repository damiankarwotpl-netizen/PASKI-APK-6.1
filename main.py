import os, json, sqlite3, threading, smtplib, mimetypes
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
from kivy.uix.screenmanager import ScreenManager, Screen
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Border, Side, Font, Alignment

try: import xlrd
except ImportError: xlrd = None

# --- KONFIGURACJA WIZUALNA ---
COLOR_PRIMARY = (0.1, 0.5, 0.9, 1)
COLOR_BG = (0.08, 0.1, 0.15, 1)

class PremiumButton(Button):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.background_normal = ""
        self.background_color = COLOR_PRIMARY
        self.height = dp(53)
        self.size_hint_y = None
        self.bold = True

class SafeLabel(Label):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.halign = 'center'
        self.valign = 'middle'
        self.padding = (dp(5), dp(5))
        self.bind(size=self._update)
    def _update(self, inst, val):
        self.text_size = (val[0], None)

class FutureApp(App):
    def build(self):
        self.title = "FUTURE 22.0 ULTIMATE"
        Window.clearcolor = COLOR_BG
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
        self.pages = {
            "home": Screen(name="home"), "table": Screen(name="table"),
            "email": Screen(name="email"), "smtp": Screen(name="smtp"),
            "tmpl": Screen(name="tmpl"), "contacts": Screen(name="contacts")
        }
        self.setup_ui()
        for s in self.pages.values(): self.sm.add_widget(s)
        return self.sm

    def init_db(self):
        db_p = Path(self.user_data_dir) / "future_v22_ultimate.db"
        self.conn = sqlite3.connect(str(db_p), check_same_thread=False)
        self.conn.execute("CREATE TABLE IF NOT EXISTS contacts (name TEXT, surname TEXT, email TEXT, pesel TEXT, phone TEXT, PRIMARY KEY(name, surname))")
        self.conn.execute("CREATE TABLE IF NOT EXISTS settings (key TEXT PRIMARY KEY, val TEXT)")
        self.conn.commit()

    # --- LOGIKA MAILINGU ---
    def start_mass_mailing(self, _):
        if not self.full_data: 
            self.msg("!", "Brak danych! Wczytaj najpierw arkusz płac.")
            return
        self.stats = {"ok": 0, "fail": 0, "skip": 0, "auto": 0}
        self.queue = list(self.full_data[1:])
        self.process_mailing_queue()

    def process_mailing_queue(self):
        if not self.queue:
            self.msg("Raport Wysyłki", f"Zakończono.\nAutomatycznie (PESEL): {self.stats['auto']}\nRęcznie: {self.stats['ok']-self.stats['auto']}\nBłędy: {self.stats['fail']}\nPominięto: {self.stats['skip']}")
            return
        
        row = self.queue.pop(0)
        n, s = str(row[self.idx_name]).strip(), str(row[self.idx_surname]).strip()
        pesel = str(row[self.idx_pesel]).strip() if self.idx_pesel != -1 else ""
        
        if pesel and len(pesel) > 5:
            res = self.conn.execute("SELECT email FROM contacts WHERE pesel=?", (pesel,)).fetchone()
            if res and res[0]:
                self.stats["auto"] += 1
                self.send_email_engine(row, res[0])
                return
        
        res = self.conn.execute("SELECT email, pesel FROM contacts WHERE name=? AND surname=?", (n.lower(), s.lower())).fetchone()
        if res and res[0]:
            self.ask_before_send(row, res[0], n, s, pesel, res[1])
        else:
            self.stats["skip"] += 1
            self.process_mailing_queue()

    def ask_before_send(self, row, email, n, s, pesel_f, pesel_db):
        box = BoxLayout(orientation="vertical", padding=dp(20), spacing=dp(10))
        box.add_widget(Label(text="DOPASOWANIE RĘCZNE:", bold=True, color=(1, 0.6, 0.2, 1)))
        box.add_widget(SafeLabel(text=f"Osoba: {n} {s}\nEmail: {email}\nPESEL plik: {pesel_f}\nPESEL baza: {pesel_db}"))
        
        btns = BoxLayout(size_hint_y=None, height=dp(55), spacing=dp(10))
        def dec(choice):
            p.dismiss()
            if choice: self.send_email_engine(row, email)
            else: self.stats["skip"] += 1; self.process_mailing_queue()
            
        btns.add_widget(Button(text="WYŚLIJ", on_press=lambda x: dec(True), background_color=(0,0.6,0,1)))
        btns.add_widget(Button(text="POMIŃ", on_press=lambda x: dec(False), background_color=(0.6,0,0,1)))
        box.add_widget(btns)
        p = Popup(title="Weryfikacja adresata", content=box, size_hint=(0.95, 0.6)); p.open()

    def send_email_engine(self, row_data, target_email, fast_mode=False):
        def thread_task():
            p_smtp = Path(self.user_data_dir) / "smtp.json"
            if not p_smtp.exists(): return
            cfg = json.load(open(p_smtp))
            try:
                srv = smtplib.SMTP("smtp.gmail.com", 587, timeout=12); srv.starttls(); srv.login(cfg['u'], cfg['p'])
                msg = EmailMessage(); dat = datetime.now().strftime("%d.%m.%Y")
                ts = self.conn.execute("SELECT val FROM settings WHERE key='t_sub'").fetchone()
                tb = self.conn.execute("SELECT val FROM settings WHERE key='t_body'").fetchone()
                name_disp = row_data[self.idx_name] if not fast_mode else "Użytkowniku"
                
                msg["Subject"] = (ts[0] if ts else "Raport").replace("{Imię}", name_disp)
                msg["From"], msg["To"] = cfg['u'], target_email
                msg.set_content((tb[0] if tb else "Informacja").replace("{Imię}", name_disp).replace("{Data}", dat))
                
                if not fast_mode and self.full_data:
                    tmp = Path(self.user_data_dir) / f"temp_send.xlsx"
                    wb = Workbook(); ws = wb.active
                    ws.append([self.full_data[0][k] for k in self.export_indices])
                    ws.append([row_data[k] for k in self.export_indices])
                    self.style_xlsx(ws); wb.save(tmp)
                    with open(tmp, "rb") as f: msg.add_attachment(f.read(), maintype="application", subtype="xlsx", filename=f"Raport_{name_disp}.xlsx")
                
                for path in self.global_attachments:
                    if os.path.exists(path):
                        ctype, _ = mimetypes.guess_type(path)
                        main, sub = (ctype or 'application/octet-stream').split('/', 1)
                        with open(path, "rb") as f: msg.add_attachment(f.read(), maintype=main, subtype=sub, filename=os.path.basename(path))
                
                srv.send_message(msg); srv.quit()
                if not fast_mode: Clock.schedule_once(lambda d: [self.update_stat("ok"), self.process_mailing_queue()])
            except:
                if not fast_mode: Clock.schedule_once(lambda d: [self.update_stat("fail"), self.process_mailing_queue()])
        threading.Thread(target=thread_task, daemon=True).start()

    def update_stat(self, key): self.stats[key] += 1

    # --- UI SETUP ---
    def setup_ui(self):
        # HOME PAGE
        l = BoxLayout(orientation="vertical", padding=dp(25), spacing=dp(12))
        l.add_widget(Label(text="FUTURE 22.0 ULTIMATE", font_size='26sp', bold=True, color=COLOR_PRIMARY, size_hint_y=None, height=dp(80)))
        btn = lambda t, c: l.add_widget(PremiumButton(text=t, on_press=c))
        btn("[F1] WCZYTAJ ARKUSZ PŁAC", lambda x: self.open_picker("data"))
        btn("[F2] PODGLĄD I EKSPORT", lambda x: [self.refresh_table(), setattr(self.sm, 'current', 'table')] if self.full_data else self.msg("!", "Wczytaj plik!"))
        btn("[F3] CENTRUM MAILINGOWE", lambda x: setattr(self.sm, 'current', 'email'))
        btn("[F4] KONFIGURACJA SMTP", lambda x: setattr(self.sm, 'current', 'smtp'))
        self.pages["home"].add_widget(l)
        
        # TABLE PAGE (REFACTORING DLA POPRAWY WIDOKU)
        rt = BoxLayout(orientation="vertical", padding=dp(5), spacing=dp(5))
        mt = BoxLayout(size_hint_y=None, height=dp(55), spacing=dp(5))
        self.ti_search = TextInput(hint_text="Filtruj tabelę...", multiline=False); self.ti_search.bind(text=self.filter_table)
        mt.add_widget(self.ti_search)
        mt.add_widget(Button(text="Opcje", size_hint_x=None, width=dp(80), on_press=self.popup_columns))
        mt.add_widget(Button(text="Wróć", size_hint_x=None, width=dp(80), on_press=lambda x: setattr(self.sm, 'current', 'home')))
        
        self.table_scroll = ScrollView(do_scroll_x=True, do_scroll_y=True)
        self.table_grid = GridLayout(size_hint=(None, None), spacing=dp(1))
        self.table_grid.bind(minimum_height=self.table_grid.setter("height"), minimum_width=self.table_grid.setter("width"))
        self.table_scroll.add_widget(self.table_grid)
        rt.add_widget(mt); rt.add_widget(self.table_scroll)
        self.pages["table"].add_widget(rt)
        
        self.setup_email_ui(); self.setup_smtp(); self.setup_tmpl(); self.setup_contacts()

    def setup_email_ui(self):
        l = BoxLayout(orientation="vertical", padding=dp(25), spacing=dp(12))
        l.add_widget(Label(text="CENTRUM MAILINGOWE", font_size='22sp', bold=True, size_hint_y=None, height=dp(60)))
        self.lbl_stats = Label(text="Baza: 0", color=(0.7,0.8,1,1)); l.add_widget(self.lbl_stats)
        btn = lambda t, c: l.add_widget(PremiumButton(text=t, on_press=c))
        btn("IMPORTUJ KONTAKTY", lambda x: self.open_picker("book"))
        btn("ZARZĄDZAJ BAZĄ", lambda x: [self.refresh_contacts_list(), setattr(self.sm, 'current', 'contacts')])
        btn("EDYTUJ WIADOMOŚĆ", lambda x: setattr(self.sm, 'current', 'tmpl'))
        btn("DODAJ ZAŁĄCZNIK", lambda x: self.open_picker("attachment"))
        btn("URUCHOM WYSYŁKĘ", self.start_mass_mailing)
        btn("POWRÓT", lambda x: setattr(self.sm, 'current', 'home'))
        self.pages["email"].add_widget(l); self.update_stats()

    def setup_contacts(self):
        l = BoxLayout(orientation="vertical", padding=dp(8), spacing=dp(5))
        top = BoxLayout(size_hint_y=None, height=dp(55), spacing=dp(5))
        self.ti_csearch = TextInput(hint_text="Szukaj (Nazwisko/PESEL)...", multiline=False)
        self.ti_csearch.bind(text=self.refresh_contacts_list)
        top.add_widget(self.ti_csearch)
        top.add_widget(Button(text="+", size_hint_x=None, width=dp(55), on_press=lambda x: self.form_contact()))
        top.add_widget(Button(text="X", size_hint_x=None, width=dp(55), on_press=lambda x: setattr(self.sm, 'current', 'email')))
        
        self.c_scroll = ScrollView(); self.c_list = GridLayout(cols=1, size_hint_y=None, spacing=dp(5))
        self.c_list.bind(minimum_height=self.c_list.setter('height'))
        self.c_scroll.add_widget(self.c_list)
        
        foot = BoxLayout(size_hint_y=None, height=dp(55), spacing=dp(8))
        foot.add_widget(Button(text="Odznacz", on_press=lambda x: [setattr(self, 'selected_emails', []), self.refresh_contacts_list()]))
        foot.add_widget(PremiumButton(text="WYŚLIJ MULTI (SZYBKI)", on_press=self.run_quick_send))
        l.add_widget(top); l.add_widget(self.c_scroll); l.add_widget(foot)
        self.pages["contacts"].add_widget(l)

    def refresh_contacts_list(self, *args):
        self.c_list.clear_widgets(); sv = self.ti_csearch.text.lower()
        rows = self.conn.execute("SELECT name, surname, email, pesel, phone FROM contacts ORDER BY surname ASC").fetchall()
        for n, s, e, p, ph in rows:
            name_full = f"{n} {s}".title()
            if sv and (sv not in name_full.lower() and sv not in str(e).lower() and sv not in str(p)): continue
            
            row = BoxLayout(orientation="horizontal", size_hint_y=None, height=dp(85), padding=dp(5))
            cb = CheckBox(size_hint_x=None, width=dp(45), active=(e in self.selected_emails))
            cb.bind(active=lambda inst, val, mail=e: [self.selected_emails.append(mail) if val else self.selected_emails.remove(mail)])
            row.add_widget(cb)
            
            info = BoxLayout(orientation="vertical")
            info.add_widget(Label(text=name_full, bold=True, halign="left", text_size=(dp(220),None)))
            info.add_widget(Label(text=f"{e} | PESEL: {p}", font_size='11sp', color=(0.7,0.7,0.7,1), halign="left", text_size=(dp(220),None)))
            row.add_widget(info)
            
            acts = BoxLayout(size_hint_x=None, width=dp(70), orientation="vertical", spacing=dp(2))
            acts.add_widget(Button(text="Edytuj", font_size='10sp', on_press=lambda x, d=(n,s,e,p,ph): self.form_contact(*d)))
            acts.add_widget(Button(text="Usuń", font_size='10sp', background_color=(0.6,0,0,1), on_press=lambda x, n=n, s=s: self.delete_contact(n,s)))
            row.add_widget(acts); self.c_list.add_widget(row)

    # --- PICKER I EXCEL ---
    def open_picker(self, mode):
        if platform != "android": self.msg("!", "Opcja dostępna tylko na Androidzie"); return
        from jnius import autoclass; from android import activity
        PA = autoclass("org.kivy.android.PythonActivity"); Intent = autoclass("android.content.Intent")
        intent = Intent(Intent.ACTION_GET_CONTENT); intent.setType("*/*")
        def cb(req, res, dt):
            if not dt: return
            activity.unbind(on_activity_result=cb); 
            uri = dt.getData(); stream = PA.mActivity.getContentResolver().openInputStream(uri)
            ext = ".xlsx" if mode != "attachment" else "_" + os.urandom(2).hex()
            loc = Path(self.user_data_dir) / f"file_{mode}{ext}"
            with open(loc, "wb") as f:
                while True:
                    b = stream.read(bytearray(16384))
                    if b == -1: break
                    f.write(b)
            stream.close()
            if mode == "data": self.process_excel(loc)
            elif mode == "book": self.process_book(loc)
            elif mode == "attachment": self.global_attachments.append(str(loc)); self.update_stats()
        activity.bind(on_activity_result=cb); PA.mActivity.startActivityForResult(intent, 1001)

    def process_excel(self, path):
        try:
            if str(path).endswith(".xls") and xlrd:
                wb = xlrd.open_workbook(path); ws = wb.sheet_by_index(0); raw = [[str(ws.cell_value(r,c)).strip() for c in range(ws.ncols)] for r in range(ws.nrows)]
            else:
                wb = load_workbook(path, data_only=True); ws = wb.active; raw = [["" if v is None else str(v).strip() for v in r] for r in ws.iter_rows(values_only=True)]
            
            h_idx = 0
            for i, row in enumerate(raw[:15]):
                line = " ".join(row).lower()
                if any(x in line for x in ["imię", "imie", "nazwisko", "pesel"]): h_idx = i; break
            
            self.full_data = [r for r in raw[h_idx:] if any(c.strip() for c in r)]
            self.filtered_data = self.full_data
            self.export_indices = list(range(len(self.full_data[0])))
            h = [x.lower() for x in self.full_data[0]]
            for i,v in enumerate(h):
                if "imi" in v: self.idx_name = i
                if "naz" in v: self.idx_surname = i
                if "pesel" in v: self.idx_pesel = i
            self.msg("OK", "Arkusz wczytany.")
        except Exception as e: self.msg("Błąd", str(e))

    def refresh_table(self):
        self.table_grid.clear_widgets()
        if not self.filtered_data: return
        rws, cls = len(self.filtered_data), len(self.filtered_data[0])
        w, h = dp(180), dp(55)
        self.table_grid.cols = cls+1; self.table_grid.width, self.table_grid.height = (cls+1)*w, rws*h
        
        for head in self.filtered_data[0]:
            self.table_grid.add_widget(SafeLabel(text=str(head), bold=True, color=COLOR_PRIMARY, size=(w,h), size_hint=(None,None)))
        self.table_grid.add_widget(SafeLabel(text="AKCJA", bold=True, size=(w,h), size_hint=(None,None)))
        
        for row in self.filtered_data[1:]:
            for c in row: self.table_grid.add_widget(SafeLabel(text=str(c), size=(w,h), size_hint=(None,None)))
            btn = Button(text="Zapisz", size=(w,h), size_hint=(None,None))
            btn.bind(on_press=lambda x, r=row: self.export_xlsx(r))
            self.table_grid.add_widget(btn)

    def export_xlsx(self, r):
        doc_dir = Path("/storage/emulated/0/Documents/FutureExport")
        doc_dir.mkdir(parents=True, exist_ok=True)
        wb = Workbook(); ws = wb.active
        ws.append([self.full_data[0][k] for k in self.export_indices])
        ws.append([r[k] for k in self.export_indices])
        self.style_xlsx(ws)
        name = str(r[self.idx_name]).replace(" ","_")
        path = doc_dir / f"Raport_{name}.xlsx"
        wb.save(path); self.msg("OK", f"Zapisano: {path.name}")

    def style_xlsx(self, ws):
        side = Side(style='thin')
        for r in ws.iter_rows():
            for c in r: 
                c.border = Border(top=side, left=side, right=side, bottom=side)
                c.alignment = Alignment(horizontal='center')
        for c in ws[1]: c.font = Font(bold=True)
        for col in ws.columns: ws.column_dimensions[col[0].column_letter].width = 20

    def popup_columns(self, _):
        if not self.full_data: return
        box = BoxLayout(orientation="vertical", padding=dp(10), spacing=dp(5))
        sc = ScrollView(); gr = GridLayout(cols=1, size_hint_y=None, spacing=dp(5))
        gr.bind(minimum_height=gr.setter('height'))
        checks = []
        for i, h in enumerate(self.full_data[0]):
            r = BoxLayout(size_hint_y=None, height=dp(45))
            txt = str(h) if h else f"Kolumna {i+1}"
            cb = CheckBox(active=(i in self.export_indices), size_hint_x=None, width=dp(50)); checks.append((i,cb))
            r.add_widget(cb); r.add_widget(Label(text=txt, halign="left", text_size=(dp(220), None))); gr.add_widget(r)
        sc.add_widget(gr); box.add_widget(sc)
        box.add_widget(PremiumButton(text="ZASTOSUJ", on_press=lambda x: [setattr(self, 'export_indices', [idx for idx,c in checks if c.active]), p.dismiss()]))
        p = Popup(title="Widoczność kolumn", content=box, size_hint=(0.95, 0.85)); p.open()

    def update_stats(self, *args):
        c = self.conn.execute("SELECT count(*) FROM contacts").fetchone()[0]
        self.lbl_stats.text = f"Baza kontaktów: {c} | Załączniki: {len(self.global_attachments)}"

    def setup_smtp(self):
        l = BoxLayout(orientation="vertical", padding=dp(25), spacing=dp(10))
        self.ti_su = TextInput(hint_text="Gmail (np. biuro@gmail.com)", multiline=False)
        self.ti_sp = TextInput(hint_text="Hasło App (16 znaków)", password=True)
        p = Path(self.user_data_dir) / "smtp.json"
        if p.exists():
            d = json.load(open(p)); self.ti_su.text, self.ti_sp.text = d.get('u',''), d.get('p','')
        
        def sv(_): 
            json.dump({'u':self.ti_su.text, 'p':self.ti_sp.text}, open(p, "w"))
            self.msg("OK", "Ustawienia SMTP zapisane.")
        
        l.add_widget(Label(text="USTAWIENIA SMTP", bold=True)); l.add_widget(self.ti_su); l.add_widget(self.ti_sp)
        l.add_widget(PremiumButton(text="ZAPISZ", on_press=sv))
        l.add_widget(PremiumButton(text="POWRÓT", on_press=lambda x: setattr(self.sm, 'current', 'home'), background_color=(0.4,0.4,0.4,1)))
        self.pages["smtp"].add_widget(l)

    def setup_tmpl(self):
        l = BoxLayout(orientation="vertical", padding=dp(20), spacing=dp(8))
        self.ti_ts = TextInput(hint_text="Temat {Imię}", size_hint_y=None, height=dp(45))
        self.ti_tb = TextInput(hint_text="Treść maila...", multiline=True)
        r = self.conn.execute("SELECT val FROM settings WHERE key='t_sub'").fetchone()
        rb = self.conn.execute("SELECT val FROM settings WHERE key='t_body'").fetchone()
        if r: self.ti_ts.text, self.ti_tb.text = r[0], rb[0]
        
        def sv(_):
            self.conn.execute("INSERT OR REPLACE INTO settings VALUES (?,?)", ('t_sub', self.ti_ts.text))
            self.conn.execute("INSERT OR REPLACE INTO settings VALUES (?,?)", ('t_body', self.ti_tb.text))
            self.conn.commit(); self.msg("OK", "Szablon zapisany.")
            
        l.add_widget(Label(text="SZABLON MAILA", bold=True)); l.add_widget(self.ti_ts); l.add_widget(self.ti_tb)
        l.add_widget(PremiumButton(text="ZAPISZ", on_press=sv))
        l.add_widget(PremiumButton(text="POWRÓT", on_press=lambda x: setattr(self.sm, 'current', 'email'), background_color=(0.4,0.4,0.4,1)))
        self.pages["tmpl"].add_widget(l)

    def form_contact(self, n="", s="", e="", pes="", ph=""):
        b = BoxLayout(orientation="vertical", padding=dp(15), spacing=dp(8))
        ti_n = TextInput(text=n, hint_text="Imię"); ti_s = TextInput(text=s, hint_text="Nazwisko")
        ti_e = TextInput(text=e, hint_text="Email"); ti_p = TextInput(text=pes, hint_text="PESEL")
        ti_ph = TextInput(text=ph, hint_text="Telefon")
        if n: ti_n.readonly = True; ti_s.readonly = True
        
        def sv(_):
            self.conn.execute("INSERT OR REPLACE INTO contacts VALUES (?,?,?,?,?)", 
                              (ti_n.text.strip().lower(), ti_s.text.strip().lower(), ti_e.text.strip(), ti_p.text.strip(), ti_ph.text.strip()))
            self.conn.commit(); p.dismiss(); self.refresh_contacts_list(); self.update_stats()
            
        b.add_widget(ti_n); b.add_widget(ti_s); b.add_widget(ti_e); b.add_widget(ti_p); b.add_widget(ti_ph)
        b.add_widget(PremiumButton(text="ZAPISZ", on_press=sv))
        p = Popup(title="Kontakt", content=b, size_hint=(0.95, 0.8)); p.open()

    def delete_contact(self, n, s):
        def pr(_): 
            self.conn.execute("DELETE FROM contacts WHERE name=? AND surname=?", (n, s))
            self.conn.commit(); p.dismiss(); self.refresh_contacts_list(); self.update_stats()
        p = Popup(title="Usuń?", content=Button(text="USUŃ KONTAKT", on_press=pr, background_color=(1,0,0,1)), size_hint=(0.7,0.3)); p.open()

    def process_book(self, path):
        try:
            wb = load_workbook(path, data_only=True); ws = wb.active; raw = [r for r in ws.iter_rows(values_only=True)]
            h = [str(x).lower() for x in raw[0]]; iN, iS, iE, iP = 0, 1, 2, -1
            for i,v in enumerate(h):
                if "imi" in v: iN=i
                elif "naz" in v: iS=i
                elif "@" in v or "mail" in v: iE=i
                elif "pesel" in v: iP=i
            for r in raw[1:]:
                if len(r) > iE and r[iE] and "@" in str(r[iE]):
                    pv = str(r[iP]) if (iP != -1 and len(r) > iP and r[iP]) else ""
                    self.conn.execute("INSERT OR REPLACE INTO contacts (name, surname, email, pesel, phone) VALUES (?,?,?,?,?)", 
                                      (str(r[iN]).lower(), str(r[iS]).lower(), str(r[iE]).strip(), pv, ""))
            self.conn.commit(); self.update_stats(); self.msg("OK", "Baza zaimportowana.")
        except: self.msg("Błąd", "Nie udało się zaimportować kontaktów.")

    def run_quick_send(self, _):
        if not self.selected_emails: self.msg("!", "Wybierz kogoś z listy!"); return
        for mail in self.selected_emails: self.send_email_engine([], mail, fast_mode=True)
        self.msg("Mailing", f"Wysłano maile do {len(self.selected_emails)} osób.")

    def filter_table(self, ins, val):
        v = val.lower(); self.filtered_data = [self.full_data[0]] + [r for r in self.full_data[1:] if any(v in str(c).lower() for c in r)]
        self.refresh_table()

    def msg(self, title, text):
        box = BoxLayout(orientation="vertical", padding=dp(20))
        box.add_widget(SafeLabel(text=text))
        btn = Button(text="OK", size_hint_y=None, height=dp(50), on_press=lambda x: p.dismiss())
        box.add_widget(btn)
        p = Popup(title=title, content=box, size_hint=(0.85, 0.45)); p.open()

if __name__ == "__main__":
    FutureApp().run()
