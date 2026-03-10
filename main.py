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

class HomeScreen(Screen): pass
class TableScreen(Screen): pass
class EmailScreen(Screen): pass
class SMTPScreen(Screen): pass
class TemplateScreen(Screen): pass
class ContactsScreen(Screen): pass

class FutureApp(App):
    def build(self):
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

    def start_mass_mailing(self, _):
        if not self.full_data:
            self.msg("!", "Wczytaj najpierw arkusz płac!")
            return
        self.stats = {"ok": 0, "fail": 0, "skip": 0, "auto": 0}
        self.queue = list(self.full_data[1:])
        Clock.schedule_once(lambda dt: self.process_mailing_queue())

    def process_mailing_queue(self, *args):
        if not self.queue:
            self.msg("Raport Koncowy", f"Zakończono.\nAuto (PESEL): {self.stats['auto']}\nRęcznie: {self.stats['ok']-self.stats['auto']}\nBłędy: {self.stats['fail']}\nPominięto: {self.stats['skip']}")
            return

        row = self.queue.pop(0)
        try:
            n = str(row[self.idx_name]).strip()
            s = str(row[self.idx_surname]).strip()
            pesel = str(row[self.idx_pesel]).strip() if self.idx_pesel != -1 else ""
        except:
            self.stats["skip"] += 1
            Clock.schedule_once(lambda dt: self.process_mailing_queue())
            return

        if pesel and len(pesel) > 5:
            res = self.conn.execute("SELECT email FROM contacts WHERE pesel=?", (pesel,)).fetchone()
            if res and res[0]:
                self.stats["auto"] += 1
                self.send_email_engine(row, res[0])
                return

        res = self.conn.execute("SELECT email FROM contacts WHERE name=? AND surname=?", (n.lower(), s.lower())).fetchone()
        if res and res[0]:
            self.ask_before_send(row, res[0], n, s, pesel)
        else:
            self.stats["skip"] += 1
            Clock.schedule_once(lambda dt: self.process_mailing_queue())

    def ask_before_send(self, row, email, n, s, pesel):
        box = BoxLayout(orientation="vertical", padding=dp(20), spacing=dp(10))
        box.add_widget(SafeLabel(text=f"POTWIERDZENIE (Brak PESEL):\n\nOsoba: {n} {s}\nEmail: {email}\nPESEL: {pesel if pesel else 'BRAK'}", bold=True))
        btns = BoxLayout(size_hint_y=None, height=dp(55), spacing=dp(10))
        def dec(choice):
            p.dismiss()
            if choice:
                self.send_email_engine(row, email)
            else:
                self.stats["skip"] += 1
                Clock.schedule_once(lambda dt: self.process_mailing_queue(), 0.1)
        btns.add_widget(Button(text="WYŚLIJ", on_press=lambda x: dec(True), background_color=(0, 0.7, 0, 1)))
        btns.add_widget(Button(text="POMIŃ", on_press=lambda x: dec(False), background_color=(0.7, 0, 0, 1)))
        box.add_widget(btns)
        p = Popup(title="Weryfikacja", content=box, size_hint=(0.9, 0.5), auto_dismiss=False)
        p.open()

    def send_email_engine(self, row_data, target_email, fast_mode=False):
        def thread_task():
            p_smtp = Path(self.user_data_dir) / "smtp.json"
            if not p_smtp.exists():
                Clock.schedule_once(lambda d: self.msg("Błąd", "Skonfiguruj SMTP!"))
                return
            cfg = json.load(open(p_smtp))
            try:
                srv = smtplib.SMTP("smtp.gmail.com", 587, timeout=15)
                srv.starttls()
                srv.login(cfg['u'], cfg['p'])
                msg = EmailMessage()
                dat = datetime.now().strftime("%d.%m.%Y")
                ts = self.conn.execute("SELECT val FROM settings WHERE key='t_sub'").fetchone()
                tb = self.conn.execute("SELECT val FROM settings WHERE key='t_body'").fetchone()
                name_disp = row_data[self.idx_name] if (row_data and self.idx_name < len(row_data)) else "Użytkowniku"
                msg["Subject"] = (ts[0] if ts else "Informacja").replace("{Imię}", name_disp)
                msg["From"], msg["To"] = cfg['u'], target_email
                msg.set_content((tb[0] if tb else "Dzień dobry").replace("{Imię}", name_disp).replace("{Data}", dat))
                
                if not fast_mode and self.full_data:
                    tmp = Path(self.user_data_dir) / "r_tmp.xlsx"
                    wb = Workbook()
                    ws = wb.active
                    ws.append([self.full_data[0][k] for k in self.export_indices])
                    ws.append([row_data[k] for k in self.export_indices])
                    self.style_xlsx(ws)
                    wb.save(tmp)
                    with open(tmp, "rb") as f:
                        msg.add_attachment(f.read(), maintype="application", subtype="xlsx", filename=f"Raport_{name_disp}.xlsx")
                
                for path in self.global_attachments:
                    if os.path.exists(path):
                        ctype, _ = mimetypes.guess_type(path)
                        if not ctype: ctype = 'application/octet-stream'
                        main, sub = ctype.split('/', 1)
                        with open(path, "rb") as f:
                            msg.add_attachment(f.read(), maintype=main, subtype=sub, filename=os.path.basename(path))
                
                srv.send_message(msg)
                srv.quit()
                if not fast_mode: 
                    Clock.schedule_once(lambda d: [self.update_stat("ok"), self.process_mailing_queue()])
            except Exception as e:
                if not fast_mode: 
                    Clock.schedule_once(lambda d: [self.update_stat("fail"), self.msg("Błąd SMTP", str(e)), self.process_mailing_queue()])
        threading.Thread(target=thread_task, daemon=True).start()

    def open_picker(self, mode):
        if platform != "android":
            self.msg("!", "Dostępne tylko na Android")
            return
        from jnius import autoclass
        from android import activity
        PA = autoclass("org.kivy.android.PythonActivity")
        Intent = autoclass("android.content.Intent")
        intent = Intent(Intent.ACTION_GET_CONTENT)
        intent.setType("*/*")
        def cb(req, res, dt):
            if not dt: return
            activity.unbind(on_activity_result=cb)
            uri = dt.getData()
            # Pobieranie oryginalnej nazwy pliku
            display_name = "plik.xlsx"
            cursor = PA.mActivity.getContentResolver().query(uri, None, None, None, None)
            if cursor and cursor.moveToFirst():
                name_index = cursor.getColumnIndex("_display_name")
                if name_index != -1:
                    display_name = cursor.getString(name_index)
                cursor.close()

            stream = PA.mActivity.getContentResolver().openInputStream(uri)
            if mode == "attachment":
                loc = Path(self.user_data_dir) / display_name
            else:
                loc = Path(self.user_data_dir) / f"temp_{mode}.xlsx"

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
        activity.bind(on_activity_result=cb)
        PA.mActivity.startActivityForResult(intent, 1001)

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
            self.full_data = raw[h_idx:]; self.filtered_data = self.full_data; self.export_indices = list(range(len(self.full_data[0])))
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

    def setup_ui(self):
        l = BoxLayout(orientation="vertical", padding=dp(30), spacing=dp(15))
        l.add_widget(Label(text="FUTURE 22.3 ULTIMATE", font_size='26sp', bold=True, color=COLOR_PRIMARY))
        btn = lambda t, c: l.add_widget(PremiumButton(text=t, on_press=c))
        btn("WCZYTAJ ARKUSZ PŁAC", lambda x: self.open_picker("data"))
        btn("PODGLĄD I EKSPORT", lambda x: [self.refresh_table(), setattr(self.sm, 'current', 'table')] if self.full_data else self.msg("!", "Drak danych"))
        btn("CENTRUM MAILINGOWE", lambda x: setattr(self.sm, 'current', 'email'))
        btn("USTAWIENIA SMTP", lambda x: setattr(self.sm, 'current', 'smtp'))
        self.screens["home"].add_widget(l)
        self.setup_table_ui(); self.setup_email_ui(); self.setup_smtp_ui(); self.setup_tmpl_ui(); self.setup_contacts_ui()

    def setup_smtp_ui(self):
        l = BoxLayout(orientation="vertical", padding=dp(25), spacing=dp(10))
        self.ti_su = TextInput(hint_text="Gmail", multiline=False)
        self.ti_sp = TextInput(hint_text="Hasło Aplikacji", password=True)
        p = Path(self.user_data_dir) / "smtp.json"
        if p.exists():
            d = json.load(open(p))
            self.ti_su.text, self.ti_sp.text = d.get('u',''), d.get('p','')
        sv = lambda x: [json.dump({'u':self.ti_su.text, 'p':self.ti_sp.text}, open(p, "w")), self.msg("OK", "Zapisano SMTP")]
        l.add_widget(Label(text="USTAWIENIA GMAIL", bold=True))
        l.add_widget(self.ti_su); l.add_widget(self.ti_sp)
        l.add_widget(PremiumButton(text="ZAPISZ", on_press=sv))
        l.add_widget(PremiumButton(text="POWRÓT", on_press=lambda x: setattr(self.sm, 'current', 'home'), background_color=(0.4, 0.4, 0.4, 1)))
        self.screens["smtp"].add_widget(l)

    def refresh_contacts_list(self, *args):
        self.c_list.clear_widgets()
        sv = self.ti_csearch.text.lower()
        rows = self.conn.execute("SELECT name, surname, email, pesel, phone FROM contacts ORDER BY surname ASC").fetchall()
        for n, s, e, p, ph in rows:
            name_full = f"{n} {s}".title()
            if sv and (sv not in name_full.lower() and sv not in str(e).lower() and sv not in str(p)): continue
            row = BoxLayout(orientation="horizontal", size_hint_y=None, height=dp(115), padding=dp(8))
            cb = CheckBox(size_hint_x=None, width=dp(50), active=(e in self.selected_emails))
            cb.bind(active=lambda inst, v, m=e: [self.selected_emails.append(m) if v else self.selected_emails.remove(m)])
            row.add_widget(cb)
            info = BoxLayout(orientation="vertical")
            info.add_widget(SafeLabel(text=name_full, bold=True, halign="left"))
            info.add_widget(SafeLabel(text=f"Email: {e}\nPESEL: {p if p else '---'} | Tel: {ph if ph else '---'}", font_size='12sp', color=(0.7,0.7,0.7,1), halign="left"))
            row.add_widget(info)
            acts = BoxLayout(size_hint_x=None, width=dp(90), orientation="vertical", spacing=dp(4))
            acts.add_widget(Button(text="Edytuj", font_size='12sp', on_press=lambda x, dt=(n,s,e,p,ph): self.form_contact(*dt)))
            acts.add_widget(Button(text="Usuń", font_size='12sp', background_color=(0.7,0,0,1), on_press=lambda x, n=n, s=s: self.delete_contact(n,s)))
            row.add_widget(acts); self.c_list.add_widget(row)

    def refresh_table(self):
        self.table_grid.clear_widgets()
        if not self.filtered_data: return
        rws, cls = len(self.filtered_data), len(self.filtered_data[0])
        w, h = dp(220), dp(62)
        self.table_grid.cols = cls+1; self.table_grid.width, self.table_grid.height = (cls+1)*w, rws*h
        for head in self.filtered_data[0]:
            self.table_grid.add_widget(SafeLabel(text=str(head), bold=True, color=COLOR_PRIMARY, size=(w,h), size_hint=(None,None)))
        self.table_grid.add_widget(SafeLabel(text="Akcja", bold=True, size=(w,h), size_hint=(None,None)))
        for row in self.filtered_data[1:]:
            for c in row: self.table_grid.add_widget(SafeLabel(text=str(c), size=(w,h), size_hint=(None,None)))
            btn = Button(text="Zapisz", size=(w,h), size_hint=(None,None), on_press=lambda x, r=row: self.export_xlsx(r))
            self.table_grid.add_widget(btn)

    def export_xlsx(self, r):
        f = Path("/storage/emulated/0/Documents/FutureExport")
        f.mkdir(parents=True, exist_ok=True)
        wb = Workbook(); ws = wb.active; ws.append([self.full_data[0][k] for k in self.export_indices]); ws.append([r[k] for k in self.export_indices])
        self.style_xlsx(ws); wb.save(f / f"Raport_{r[self.idx_name]}.xlsx"); self.msg("OK", "Zapisano raport.")

    def style_xlsx(self, ws):
        thin = Side(style='thin'); thick = Side(style='thick')
        for r in ws.iter_rows():
            for c in r:
                c.border = Border(top=thin, left=thin, right=thin, bottom=thin)
                c.alignment = Alignment(horizontal='center')
        for c in ws[1]:
            c.font = Font(bold=True)
            c.border = Border(top=thick, bottom=thick, left=thin, right=thin)
        for col in ws.columns: ws.column_dimensions[col[0].column_letter].width = max(len(str(x.value or "")) for x in col) + 5

    def popup_columns(self, _):
        if not self.full_data: return
        box = BoxLayout(orientation="vertical", padding=dp(15), spacing=dp(5))
        sc = ScrollView(); gr = GridLayout(cols=1, size_hint_y=None, spacing=dp(8)); gr.bind(minimum_height=gr.setter('height'))
        checks = []
        for i, h in enumerate(self.full_data[0]):
            r = BoxLayout(size_hint_y=None, height=dp(50))
            cb = CheckBox(active=(i in self.export_indices), size_hint_x=None, width=dp(50)); checks.append((i,cb))
            r.add_widget(cb); r.add_widget(Label(text=str(h) if h else f"Kol {i+1}")); gr.add_widget(r)
        sc.add_widget(gr); box.add_widget(sc)
        box.add_widget(PremiumButton(text="ZATWIERDŹ", on_press=lambda x: [setattr(self, 'export_indices', [idx for idx,c in checks if c.active]), p.dismiss()]))
        p = Popup(title="Widoczność kolumn", content=box, size_hint=(0.95, 0.9)); p.open()

    def setup_table_ui(self):
        rt = BoxLayout(orientation="vertical", padding=dp(8), spacing=dp(5))
        mt = BoxLayout(size_hint_y=None, height=dp(55), spacing=dp(5))
        self.ti_search = TextInput(hint_text="Szukaj...", multiline=False); self.ti_search.bind(text=self.filter_table)
        mt.add_widget(self.ti_search); mt.add_widget(Button(text="Opcje", size_hint_x=0.3, on_press=self.popup_columns))
        mt.add_widget(Button(text="Wróć", size_hint_x=0.2, on_press=lambda x: setattr(self.sm, 'current', 'home')))
        self.table_scroll = ScrollView(do_scroll_x=True, do_scroll_y=True)
        self.table_grid = GridLayout(size_hint=(None, None), spacing=dp(2))
        self.table_grid.bind(minimum_height=self.table_grid.setter("height"), minimum_width=self.table_grid.setter("width"))
        self.table_scroll.add_widget(self.table_grid); rt.add_widget(mt); rt.add_widget(self.table_scroll)
        self.screens["table"].add_widget(rt)

    def setup_tmpl_ui(self):
        l = BoxLayout(orientation="vertical", padding=dp(25), spacing=dp(10))
        self.ti_ts = TextInput(hint_text="Temat {Imię}", size_hint_y=None, height=dp(45)); self.ti_tb = TextInput(hint_text="Treść...", multiline=True)
        r = self.conn.execute("SELECT val FROM settings WHERE key='t_sub'").fetchone()
        rb = self.conn.execute("SELECT val FROM settings WHERE key='t_body'").fetchone()
        if r: self.ti_ts.text = r[0]
        if rb: self.ti_tb.text = rb[0]
        def sv_tmpl(_):
            self.conn.execute("INSERT OR REPLACE INTO settings VALUES (?,?)", ('t_sub', self.ti_ts.text))
            self.conn.execute("INSERT OR REPLACE INTO settings VALUES (?,?)", ('t_body', self.ti_tb.text))
            self.conn.commit(); self.msg("OK", "Szablon zapisany")
        l.add_widget(Label(text="TREŚĆ WIADOMOŚCI", bold=True)); l.add_widget(self.ti_ts); l.add_widget(self.ti_tb)
        l.add_widget(PremiumButton(text="ZAPISZ", on_press=sv_tmpl))
        l.add_widget(PremiumButton(text="POWRÓT", on_press=lambda x: setattr(self.sm, 'current', 'email')))
        self.screens["tmpl"].add_widget(l)

    def setup_contacts_ui(self):
        l = BoxLayout(orientation="vertical", padding=dp(10), spacing=dp(5))
        top = BoxLayout(size_hint_y=None, height=dp(55), spacing=dp(5))
        self.ti_csearch = TextInput(hint_text="Szukaj...", multiline=False); self.ti_csearch.bind(text=self.refresh_contacts_list)
        top.add_widget(self.ti_csearch); top.add_widget(Button(text="+ Dodaj", size_hint_x=0.2, on_press=lambda x: self.form_contact()))
        top.add_widget(Button(text="Wróć", size_hint_x=0.2, on_press=lambda x: setattr(self.sm, 'current', 'email')))
        self.c_scroll = ScrollView(); self.c_list = GridLayout(cols=1, size_hint_y=None, spacing=dp(10))
        self.c_list.bind(minimum_height=self.c_list.setter('height'))
        self.c_scroll.add_widget(self.c_list); foot = BoxLayout(size_hint_y=None, height=dp(55), spacing=dp(10))
        foot.add_widget(Button(text="Wyczyść", on_press=lambda x: [setattr(self, 'selected_emails', []), self.refresh_contacts_list()]))
        foot.add_widget(PremiumButton(text="MULTI SZYBKO", on_press=self.run_quick_send))
        l.add_widget(top); l.add_widget(self.c_scroll); l.add_widget(foot); self.screens["contacts"].add_widget(l)

    def filter_table(self, ins, val):
        v = val.lower(); self.filtered_data = [self.full_data[0]] + [r for r in self.full_data[1:] if any(v in str(c).lower() for c in r)]; self.refresh_table()

    def update_stat(self, key): self.stats[key] += 1
    def update_stats(self, *args):
        try:
            c = self.conn.execute("SELECT count(*) FROM contacts").fetchone()[0]
            self.lbl_stats.text = f"Baza: {c} | Załączniki: {len(self.global_attachments)}"
        except: pass

    def delete_contact(self, n, s):
        def pr(_): [self.conn.execute("DELETE FROM contacts WHERE name=? AND surname=?", (n, s)), self.conn.commit(), p.dismiss(), self.refresh_contacts_list(), self.update_stats()]
        p = Popup(title="Usuń?", content=Button(text="USUŃ KONTAKT", on_press=pr, background_color=(1,0,0,1)), size_hint=(0.7,0.3)); p.open()

    def run_quick_send(self, _):
        if not self.selected_emails: self.msg("!", "Zaznacz kontakty!"); return
        for e in self.selected_emails: self.send_email_engine([], e, fast_mode=True)
        self.msg("Start", "Rozpoczęto wysyłkę.")

    def form_contact(self, n="", s="", e="", pes="", ph=""):
        b = BoxLayout(orientation="vertical", padding=dp(15), spacing=dp(10))
        ti_n, ti_s, ti_e = TextInput(text=n, hint_text="Imię"), TextInput(text=s, hint_text="Nazwisko"), TextInput(text=e, hint_text="Email")
        ti_p, ti_ph = TextInput(text=pes, hint_text="PESEL"), TextInput(text=ph, hint_text="Telefon")
        def sv_contact(_):
            self.conn.execute("INSERT OR REPLACE INTO contacts VALUES (?,?,?,?,?)", (ti_n.text.strip().lower(), ti_s.text.strip().lower(), ti_e.text.strip(), ti_p.text.strip(), ti_ph.text.strip()))
            self.conn.commit(); p.dismiss(); self.refresh_contacts_list(); self.update_stats()
        for i in [ti_n, ti_s, ti_e, ti_p, ti_ph]: b.add_widget(i)
        b.add_widget(PremiumButton(text="ZAPISZ", on_press=sv_contact))
        p = Popup(title="Kontakt", content=b, size_hint=(0.9, 0.8)); p.open()

    def msg(self, title, text):
        box = BoxLayout(orientation="vertical", padding=dp(20))
        box.add_widget(Label(text=text, halign="center"))
        btn = Button(text="OK", size_hint_y=None, height=dp(50), on_press=lambda x: p.dismiss())
        box.add_widget(btn)
        p = Popup(title=title, content=box, size_hint=(0.85, 0.45)); p.open()

if __name__ == "__main__":
    FutureApp().run()
