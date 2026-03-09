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
from kivy.uix.label import Label
from kivy.uix.popup import Popup
from kivy.uix.scrollview import ScrollView
from kivy.uix.gridlayout import GridLayout
from kivy.uix.textinput import TextInput
from kivy.uix.checkbox import CheckBox
from kivy.uix.screenmanager import ScreenManager, Screen

from openpyxl import load_workbook, Workbook
from openpyxl.styles import Border, Side, Font, Alignment

try:
    import xlrd
except ImportError:
    xlrd = None

# --- STYLIZACJA ---
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
        self.bind(size=self._update)
    def _update(self, inst, val):
        self.text_size = (val[0] - dp(10), None)

class FutureApp(App):
    def build(self):
        Window.clearcolor = COLOR_BG
        self.full_data = [] 
        self.filtered_data = []
        self.export_indices = []
        self.global_attachments = [] 
        self.selected_contacts = [] # Lista maili do szybkiej wysyłki
        
        self.idx_name = 0
        self.idx_surname = 1
        self.idx_pesel = -1
        
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
        db_p = Path(self.user_data_dir) / "future_v21_elite.db"
        self.conn = sqlite3.connect(str(db_p), check_same_thread=False)
        self.conn.execute("CREATE TABLE IF NOT EXISTS contacts (name TEXT, surname TEXT, email TEXT, pesel TEXT, phone TEXT, PRIMARY KEY(name, surname))")
        self.conn.execute("CREATE TABLE IF NOT EXISTS settings (key TEXT PRIMARY KEY, val TEXT)")
        self.conn.commit()

# -----------------------------
# FUNKCJA: WYSYŁKA BEZ RAPORTU
# -----------------------------
    def fast_send_popup(self, _):
        if not self.selected_contacts:
            self.msg("!", "Zaznacz najpierw osoby w bazie danych!")
            return
        
        box = BoxLayout(orientation="vertical", padding=dp(20), spacing=dp(10))
        box.add_widget(Label(text=f"WYŚLIJ PLIK DO {len(self.selected_contacts)} OSÓB", bold=True))
        box.add_widget(Label(text="Bez raportu Excel - tylko Twoje załączniki.", color=(0.7,0.7,0.7,1)))
        
        btn_send = PremiumButton(text="🚀 URUCHOM SZYBKĄ WYSYŁKĘ")
        btn_send.bind(on_press=lambda x: [self.run_fast_send(), p.dismiss()])
        
        box.add_widget(btn_send)
        p = Popup(title="Szybka wysyłka", content=box, size_hint=(0.9, 0.4)); p.open()

    def run_fast_send(self):
        def job():
            p_smtp = Path(self.user_data_dir) / "smtp.json"
            if not p_smtp.exists(): return
            cfg = json.load(open(p_smtp))
            sent, err = 0, 0
            try:
                srv = smtplib.SMTP("smtp.gmail.com", 587, timeout=12); srv.starttls(); srv.login(cfg['u'], cfg['p'])
                ts = self.conn.execute("SELECT val FROM settings WHERE key='t_sub'").fetchone()
                tb = self.conn.execute("SELECT val FROM settings WHERE key='t_body'").fetchone()
                
                for email in self.selected_contacts:
                    try:
                        msg = EmailMessage()
                        msg["Subject"] = ts[0] if ts else "Wiadomość"
                        msg["From"], msg["To"] = cfg['u'], email
                        msg.set_content(tb[0] if tb else "Brak treści")
                        
                        for ex in self.global_attachments:
                            if os.path.exists(ex):
                                ct, _ = mimetypes.guess_type(ex)
                                if not ct: ct = 'application/octet-stream'
                                main, sub = ct.split('/', 1)
                                with open(ex, "rb") as f:
                                    msg.add_attachment(f.read(), maintype=main, subtype=sub, filename=os.path.basename(ex))
                        srv.send_message(msg); sent += 1
                    except: err += 1
                srv.quit()
                Clock.schedule_once(lambda d: self.msg("Gotowe", f"Wysłano: {sent}\nBłędy: {err}"))
            except Exception as e:
                Clock.schedule_once(lambda d: self.msg("Błąd SMTP", str(e)))
        threading.Thread(target=job, daemon=True).start()

# -----------------------------
# POPRAWIONE UI: BAZA DANYCH (Szerokie karty + Wybór)
# -----------------------------
    def setup_contacts(self):
        l = BoxLayout(orientation="vertical", padding=dp(10), spacing=dp(10))
        top = BoxLayout(size_hint_y=None, height=dp(55), spacing=dp(5))
        self.ti_csearch = TextInput(hint_text="Szukaj...", multiline=False)
        self.ti_csearch.bind(text=self.refresh_contacts_list)
        top.add_widget(self.ti_csearch)
        top.add_widget(Button(text="+ Dodaj", size_hint_x=0.2, on_press=lambda x: self.form_contact()))
        top.add_widget(Button(text="Wróć", size_hint_x=0.2, on_press=lambda x: setattr(self.sm, 'current', 'email')))
        
        self.c_scroll = ScrollView(); self.c_list = GridLayout(cols=1, size_hint_y=None, spacing=dp(12))
        self.c_list.bind(minimum_height=self.c_list.setter('height'))
        self.c_scroll.add_widget(self.c_list)
        
        # STOPKA Z AKCJAMI
        bot = BoxLayout(size_hint_y=None, height=dp(55), spacing=dp(10))
        bot.add_widget(Button(text="Wyczyść wybór", on_press=self.clear_selection))
        bot.add_widget(PremiumButton(text="WYŚLIJ PLIK DO ZAZNACZONYCH", on_press=self.fast_send_popup))
        
        l.add_widget(top); l.add_widget(self.c_scroll); l.add_widget(bot)
        self.pages["contacts"].add_widget(l)

    def refresh_contacts_list(self, *args):
        self.c_list.clear_widgets()
        sv = self.ti_csearch.text.lower()
        rows = self.conn.execute("SELECT name, surname, email, pesel, phone FROM contacts ORDER BY surname ASC").fetchall()
        for n, s, e, pes, ph in rows:
            name_full = f"{n} {s}".title()
            if sv and (sv not in name_full.lower() and sv not in str(e).lower() and sv not in str(pes)): continue
            
            # KARTA KONTAKTU (Fix Twojego 1. zdjęcia)
            card = BoxLayout(orientation="horizontal", size_hint_y=None, height=dp(95), padding=dp(10))
            
            # 1. Wybór (Checkbox)
            cb = CheckBox(size_hint_x=None, width=dp(50), active=(e in self.selected_contacts))
            cb.bind(active=lambda inst, val, email=e: self.toggle_contact(email, val))
            card.add_widget(cb)
            
            # 2. Dane (Box pionowy)
            info_box = BoxLayout(orientation="vertical")
            info_box.add_widget(Label(text=name_full, bold=True, halign="left", text_size=(dp(200),None)))
            info_box.add_widget(Label(text=e, color=(0.8,0.8,0.8,1), halign="left", font_size='13sp', text_size=(dp(200),None)))
            info_box.add_widget(Label(text=f"PESEL: {pes}", color=(0.6,0.6,0.6,1), halign="left", font_size='12sp', text_size=(dp(200),None)))
            card.add_widget(info_box)
            
            # 3. Przyciski
            acts = BoxLayout(size_hint_x=None, width=dp(90), orientation="vertical", spacing=dp(5))
            acts.add_widget(Button(text="Edytuj", on_press=lambda x, dt=(n,s,e,pes,ph): self.form_contact(*dt)))
            acts.add_widget(Button(text="Usuń", background_color=(0.8,0.1,0.1,1), on_press=lambda x, n=n, s=s: self.delete_contact(n,s)))
            card.add_widget(acts)
            
            self.c_list.add_widget(card)

    def toggle_contact(self, email, active):
        if active and email not in self.selected_contacts: self.selected_contacts.append(email)
        elif not active and email in self.selected_contacts: self.selected_contacts.remove(email)

    def clear_selection(self, _):
        self.selected_contacts = []
        self.refresh_contacts_list()

# -----------------------------
# FIX UI: PODGLĄD TABELI (Zrzut 3) I KOLUMNY (Zrzut 2)
# -----------------------------
    def setup_table(self):
        root = BoxLayout(orientation="vertical", padding=dp(8), spacing=dp(5))
        top = BoxLayout(size_hint_y=None, height=dp(55), spacing=dp(5))
        self.ti_search = TextInput(hint_text="Szukaj w arkuszu...", multiline=False)
        self.ti_search.bind(text=self.filter_table)
        top.add_widget(self.ti_search); top.add_widget(Button(text="Kolumny", size_hint_x=0.25, on_press=self.popup_columns))
        top.add_widget(Button(text="Wróć", size_hint_x=0.2, on_press=lambda x: setattr(self.sm, 'current', 'home')))
        self.table_scroll = ScrollView(do_scroll_x=True, do_scroll_y=True)
        self.table_grid = GridLayout(size_hint=(None, None), spacing=dp(3))
        self.table_grid.bind(minimum_height=self.table_grid.setter("height"), minimum_width=self.table_grid.setter("width"))
        self.table_scroll.add_widget(self.table_grid)
        root.add_widget(top); root.add_widget(self.table_scroll); self.pages["table"].add_widget(root)

    def refresh_table(self):
        self.table_grid.clear_widgets()
        if not self.filtered_data: return
        rws, cls = len(self.filtered_data), len(self.filtered_data[0])
        w, h = dp(220), dp(62) # SZEROKA KOLUMNA - Fix Twojego 3. zdjęcia
        self.table_grid.cols = cls+1; self.table_grid.width, self.table_grid.height = (cls+1)*w, rws*h
        for head in self.filtered_data[0]:
            self.table_grid.add_widget(SafeLabel(text=str(head), bold=True, color=COLOR_PRIMARY, size=(w,h), size_hint=(None,None)))
        self.table_grid.add_widget(SafeLabel(text="Akcja", bold=True, size=(w,h), size_hint=(None,None)))
        for row in self.filtered_data[1:]:
            for cell in row: self.table_grid.add_widget(SafeLabel(text=str(cell), size=(w,h), size_hint=(None,None)))
            self.table_grid.add_widget(Button(text="Zapisz", size=(w,h), size_hint=(None,None), on_press=lambda x, r=row: self.export_xlsx(r)))

    def popup_columns(self, _):
        if not self.full_data: return
        box = BoxLayout(orientation="vertical", padding=dp(15), spacing=dp(5))
        sc = ScrollView(); gr = GridLayout(cols=1, size_hint_y=None, spacing=dp(8)); gr.bind(minimum_height=gr.setter('height'))
        checks = []
        for i, h in enumerate(self.full_data[0]):
            r = BoxLayout(size_hint_y=None, height=dp(50))
            txt = str(h) if (h and str(h).strip()) else f"Kolumna {i+1}"
            cb = CheckBox(active=(i in self.export_indices), size_hint_x=None, width=dp(50))
            checks.append((i,cb)); r.add_widget(cb)
            # ETYKIETA Z KOLOREM - Fix Twojego 2. zdjęcia
            r.add_widget(Label(text=txt, halign="left", text_size=(dp(280), None), color=(1,1,1,1))) 
            gr.add_widget(r)
        sc.add_widget(gr); box.add_widget(sc)
        box.add_widget(PremiumButton(text="POTWIERDŹ", on_press=lambda x: [setattr(self, 'export_indices', [i for i,c in checks if c.active]), p.dismiss()]))
        p = Popup(title="Wybierz dane do raportu", content=box, size_hint=(0.95, 0.9)); p.open()

# -----------------------------
# POZOSTAŁE FUNKCJE SYSTEMOWE
# -----------------------------
    def setup_ui(self):
        # Home
        l = BoxLayout(orientation="vertical", padding=dp(30), spacing=dp(15))
        l.add_widget(Label(text="FUTURE 21.0 ELITE", font_size='26sp', bold=True, color=COLOR_PRIMARY))
        btn = lambda t, c: l.add_widget(PremiumButton(text=t, on_press=c))
        btn("📂 ARKUSZ PŁAC", lambda x: self.open_picker("data"))
        btn("📊 TABELA DANYCH", lambda x: [self.refresh_table(), setattr(self.sm, 'current', 'table')] if self.full_data else self.msg("!","Brak danych"))
        btn("✉ CENTRUM MAILINGU", lambda x: setattr(self.sm, 'current', 'email'))
        btn("⚙ SMTP", lambda x: setattr(self.sm, 'current', 'smtp'))
        self.pages["home"].add_widget(l); self.setup_contacts(); self.setup_table(); self.setup_email_ui(); self.setup_smtp(); self.setup_tmpl()

    def setup_email_ui(self):
        l = BoxLayout(orientation="vertical", padding=dp(25), spacing=dp(10))
        l.add_widget(Label(text="CENTRUM MAILINGOWE", font_size='22sp', bold=True))
        self.lbl_stats = Label(text="Baza: 0"); l.add_widget(self.lbl_stats)
        btn = lambda t, c: l.add_widget(PremiumButton(text=t, on_press=c))
        btn("📁 IMPORT BAZY GMAIL", lambda x: self.open_picker("book"))
        btn("🔧 BAZA (PESEL / SZYBKA WYSYŁKA)", lambda x: [self.refresh_contacts_list(), setattr(self.sm, 'current', 'contacts')])
        btn("📝 SZABLON MAILA", lambda x: setattr(self.sm, 'current', 'tmpl'))
        btn("📎 ZAŁĄCZNIK", lambda x: self.open_picker("attachment"))
        btn("🚀 MASOWA WYSYŁKA (Z RAPORTEM)", self.mass_mailing_start)
        btn("WSTECZ", lambda x: setattr(self.sm, 'current', 'home'))
        self.pages["email"].add_widget(l); self.update_stats()

    def open_picker(self, mode):
        if platform != "android": self.msg("!", "Tylko Android"); return
        from jnius import autoclass; from android import activity
        PA = autoclass("org.kivy.android.PythonActivity"); Intent = autoclass("android.content.Intent")
        intent = Intent(Intent.ACTION_GET_CONTENT); intent.setType("*/*")
        def cb(req, res, dt):
            if not dt: return
            activity.unbind(on_activity_result=cb); uri = dt.getData(); stream = PA.mActivity.getContentResolver().openInputStream(uri)
            ext = ".xlsx" if mode != "attachment" else "_" + os.urandom(2).hex()
            loc = Path(self.user_data_dir) / f"ptr_{mode}{ext}"
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
            self.full_data = raw[h_idx:]; self.filtered_data = self.full_data; self.export_indices = list(range(len(self.full_data[0])))
            h = [x.lower() for x in self.full_data[0]]
            for i,v in enumerate(h):
                if "imi" in v: self.idx_name = i
                if "naz" in v: self.idx_surname = i
                if "pesel" in v: self.idx_pesel = i
            self.msg("OK", "Wczytano arkusz płac.")
        except Exception as e: self.msg("Błąd", str(e))

    def mass_mailing_start(self, _):
        if not self.full_data: self.msg("!", "Wczytaj najpierw dane płac!"); return
        threading.Thread(target=self._mass_mail_worker, daemon=True).start()

    def _mass_mail_worker(self):
        # Tutaj logika wysyłki masowej z raportem Excel (jak w poprzedniej wersji Commander)
        pass 

    def update_stats(self, *args):
        c = self.conn.execute("SELECT count(*) FROM contacts").fetchone()[0]
        self.lbl_stats.text = f"Kontakty w bazie: {c} | Załączniki: {len(self.global_attachments)}"

    def setup_smtp(self):
        l = BoxLayout(orientation="vertical", padding=dp(25), spacing=dp(10))
        self.ti_su = TextInput(hint_text="Gmail", multiline=False); self.ti_sp = TextInput(hint_text="Hasło App", password=True)
        p = Path(self.user_data_dir) / "smtp.json"
        if p.exists(): d = json.load(open(p)); self.ti_su.text, self.ti_sp.text = d.get('u',''), d.get('p','')
        sv = lambda x: [json.dump({'u':self.ti_su.text, 'p':self.ti_sp.text}, open(p,"w")), self.msg("OK","Zapisano")]
        l.add_widget(Label(text="USTAWIENIA SMTP", bold=True)); l.add_widget(self.ti_su); l.add_widget(self.ti_sp)
        l.add_widget(PremiumButton(text="ZAPISZ", on_press=sv)); l.add_widget(PremiumButton(text="POWRÓT", on_press=lambda x: setattr(self.sm, 'current', 'home'), background_color=(0.4,0.4,0.4,1))); self.pages["smtp"].add_widget(l)

    def setup_tmpl(self):
        l = BoxLayout(orientation="vertical", padding=dp(25), spacing=dp(10))
        self.ti_ts = TextInput(hint_text="Temat {Imię}", size_hint_y=None, height=dp(45)); self.ti_tb = TextInput(hint_text="Treść...", multiline=True)
        r = self.conn.execute("SELECT val FROM settings WHERE key='t_sub'").fetchone(); rb = self.conn.execute("SELECT val FROM settings WHERE key='t_body'").fetchone()
        if r: self.ti_ts.text, self.ti_tb.text = r[0], rb[0]
        sv = lambda x: [self.conn.execute("INSERT OR REPLACE INTO settings VALUES (?,?)", (i,j)) for i,j in [('t_sub', self.ti_ts.text),('t_body', self.ti_tb.text)], self.conn.commit(), self.msg("OK","Zapisano")]
        l.add_widget(Label(text="TREŚĆ MAILA", bold=True)); l.add_widget(self.ti_ts); l.add_widget(self.ti_tb); l.add_widget(PremiumButton(text="ZAPISZ", on_press=sv)); l.add_widget(PremiumButton(text="POWRÓT", on_press=lambda x: setattr(self.sm, 'current', 'email'), background_color=(0.4,0.4,0.4,1))); self.pages["tmpl"].add_widget(l)

    def style_xlsx(self, ws):
        thin = Side(style='thin'); thick = Side(style='thick')
        for r in ws.iter_rows():
            for c in r: c.border = Border(top=thin, left=thin, right=thin, bottom=thin); c.alignment = Alignment(horizontal='center')
        for c in ws[1]: c.font = Font(bold=True); c.border = Border(top=thick, bottom=thick, left=thin, right=thin)
        for col in ws.columns: ws.column_dimensions[col[0].column_letter].width = max(len(str(x.value or "")) for x in col) + 4

    def export_xlsx(self, r):
        f = Path("/storage/emulated/0/Documents/FutureExport"); f.mkdir(parents=True, exist_ok=True)
        wb = Workbook(); ws = wb.active; ws.append([self.full_data[0][k] for k in self.export_indices]); ws.append([r[k] for k in self.export_indices])
        self.style_xlsx(ws); wb.save(f / f"Raport_{r[self.idx_name]}.xlsx"); self.msg("OK", "Zapisano raport.")

    def form_contact(self, n="", s="", e="", pes="", ph=""):
        b = BoxLayout(orientation="vertical", padding=dp(15), spacing=dp(10))
        ti_n, ti_s, ti_e = TextInput(text=n, hint_text="Imię"), TextInput(text=s, hint_text="Nazwisko"), TextInput(text=e, hint_text="Email")
        ti_p, ti_ph = TextInput(text=pes, hint_text="PESEL"), TextInput(text=ph, hint_text="Telefon")
        if n: ti_n.readonly = True; ti_s.readonly = True
        sv = lambda x: [self.conn.execute("INSERT OR REPLACE INTO contacts VALUES (?,?,?,?,?)", (ti_n.text.strip().lower(), ti_s.text.strip().lower(), ti_e.text.strip(), ti_p.text.strip(), ti_ph.text.strip())), self.conn.commit(), p.dismiss(), self.refresh_contacts_list(), self.update_stats()]
        b.add_widget(ti_n); b.add_widget(ti_s); b.add_widget(ti_e); b.add_widget(ti_p); b.add_widget(ti_ph); b.add_widget(PremiumButton(text="ZAPISZ", on_press=sv)); p = Popup(title="Kontakt", content=b, size_hint=(0.9, 0.75)); p.open()

    def delete_contact(self, n, s):
        def pr(_): [self.conn.execute("DELETE FROM contacts WHERE name=? AND surname=?", (n, s)), self.conn.commit(), p.dismiss(), self.refresh_contacts_list(), self.update_stats()]
        p = Popup(title="Usuń?", content=Button(text="USUŃ KONTAKT", on_press=pr, background_color=(1,0,0,1)), size_hint=(0.7,0.3)); p.open()

    def filter_table(self, ins, val):
        v = val.lower(); self.filtered_data = [self.full_data[0]] + [r for r in self.full_data[1:] if any(v in str(c).lower() for c in r)]; self.refresh_table()

    def msg(self, title, text):
        box = BoxLayout(orientation="vertical", padding=dp(20)); box.add_widget(Label(text=text, halign="center")); box.add_widget(Button(text="OK", on_press=lambda x: p.dismiss())); p = Popup(title=title, content=box, size_hint=(0.85, 0.45)); p.open()

    def process_book(self, path):
        try:
            raw = []; wb = load_workbook(path, data_only=True); ws = wb.active; raw = [["" if v is None else str(v).strip() for v in r] for r in ws.iter_rows(values_only=True)]
            h = [str(x).lower().strip() for x in raw[0]]
            iN, iS, iE, iP = 0, 1, 2, -1
            for i,v in enumerate(h):
                if "imi" in v: iN=i
                elif "naz" in v: iS=i
                elif "@" in v or "mail" in v: iE=i
                elif "pesel" in v: iP=i
            for r in raw[1:]:
                if len(r) > iE and "@" in str(r[iE]):
                    pes_val = str(r[iP]).strip() if (iP != -1 and len(r) > iP) else ""
                    self.conn.execute("INSERT OR REPLACE INTO contacts (name, surname, email, pesel, phone) VALUES (?,?,?,?,?)", (r[iN].lower(), r[iS].lower(), str(r[iE]).strip(), pes_val, ""))
            self.conn.commit(); self.update_stats(); self.msg("OK", "Zaimportowano bazę.")
        except: self.msg("Błąd", "Nie udało się zaimportować.")

if __name__ == "__main__": FutureApp().run()
