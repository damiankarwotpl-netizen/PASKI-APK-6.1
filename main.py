import os
import sqlite3
import threading
import json
import smtplib
import mimetypes
from datetime import datetime
from pathlib import Path
from email.message import EmailMessage

from kivy.app import App
from kivy.metrics import dp
from kivy.clock import Clock
from kivy.utils import platform
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.button import Button
from kivy.uix.label import Label
from kivy.uix.popup import Popup
from kivy.uix.scrollview import ScrollView
from kivy.uix.gridlayout import GridLayout
from kivy.uix.textinput import TextInput
from kivy.uix.progressbar import ProgressBar
from kivy.uix.screenmanager import ScreenManager, Screen

from openpyxl import load_workbook, Workbook

APP_TITLE = "Future 9.4 ULTRA PRO"

class FutureApp(App):
    def build(self):
        self.full_data = []
        self.current_file = None
        self.global_attachments = [] 
        
        self.sm = ScreenManager()
        self.init_db()
        self.init_ui()
        
        for s in [self.home_scr, self.table_scr, self.email_scr, self.smtp_scr]:
            self.sm.add_widget(s)
        return self.sm

    def init_db(self):
        db_path = Path(self.user_data_dir) / "app_v9_core.db"
        self.conn = sqlite3.connect(str(db_path), check_same_thread=False)
        self.conn.execute("CREATE TABLE IF NOT EXISTS contacts (name TEXT, surname TEXT, email TEXT, PRIMARY KEY(name, surname))")
        self.conn.commit()

    def init_ui(self):
        self.home_scr = Screen(name="home"); self.setup_home()
        self.table_scr = Screen(name="table"); self.setup_table()
        self.email_scr = Screen(name="email"); self.setup_email()
        self.smtp_scr = Screen(name="smtp"); self.setup_smtp()

    def setup_email(self):
        le = BoxLayout(orientation="vertical", padding=dp(20), spacing=dp(10))
        # NOWO\u015a\u0106: Przycisk TESTU oraz licznik rozmiaru
        self.size_label = Label(text="Ca\u0142kowity rozmiar za\u0142\u0105cznik\u00f3w: 0 MB", font_size=12, color=(0.8, 0.8, 0.8, 1))
        
        btn_data = [
            ("\ud83d\udcc1 ZARZ\u0104DZAJ ZA\u0141\u0104CZNIKAMI", self.attachment_manager_popup),
            ("\u26a1 WY\u015aLIJ TEST DO SIEBIE", self.send_test_email),
            ("\ud83d\ude80 URUCHOM MASOW\u0104 WYSY\u0141K\u0118", self.start_mailing),
            ("POWR\u00d3T", lambda x: setattr(self.sm, "current", "table"))
        ]
        
        le.add_widget(Label(text="Centrum Wysy\u0142ki", font_size=20, bold=True))
        le.add_widget(self.size_label)
        
        for t, c in btn_data:
            btn = Button(text=t, size_hint_y=None, height=dp(50), 
                         background_color=(0.1, 0.4, 0.7, 1) if "TEST" in t else (0.2, 0.2, 0.2, 1))
            btn.bind(on_press=c); le.add_widget(btn)
            
        self.email_status = Label(text="System gotowy")
        le.add_widget(self.email_status); self.email_scr.add_widget(le)

    # --- LOGIKA LICZENIA ROZMIARU PLIK\u00d3W ---
    def get_total_attachments_size(self):
        total = 0
        for path in self.global_attachments:
            if os.path.exists(path):
                total += os.path.getsize(path)
        return total / (1024 * 1024) # Wynik w MB

    def update_size_display(self):
        size = self.get_total_attachments_size()
        self.size_label.text = f"Rozmiar za\u0142\u0105cznik\u00f3w: {size:.2f} MB"
        if size > 20:
            self.size_label.color = (1, 0, 0, 1)
        else:
            self.size_label.color = (0.8, 0.8, 0.8, 1)

    # --- NOWA FUNKCJA: TESTOWA WYSY\u0141KA ---
    def send_test_email(self, _):
        if not self.full_data or len(self.full_data) < 2:
            self.msg("B\u0142\u0105d", "Wczytaj najpierw dane z Excela!"); return
        
        threading.Thread(target=self._test_mailing_thread).start()

    def _test_mailing_thread(self):
        p_cfg = Path(self.user_data_dir) / "smtp.json"
        if not p_cfg.exists(): 
            Clock.schedule_once(lambda x: self.msg("B\u0142\u0105d", "Brak ustawie\u0144 SMTP!")); return
            
        with open(p_cfg, "r") as f: cfg = json.load(f)
        
        try:
            srv = smtplib.SMTP("smtp.gmail.com", 587, timeout=15); srv.starttls(); srv.login(cfg['u'], cfg['p'])
            
            # Budowanie maila na podstawie PIERWSZEGO wiersza danych
            h = self.full_data[0]
            r = self.full_data[1]
            
            msg = self.build_message(cfg['u'], cfg['u'], h, r) # Wysy\u0142amy do samego siebie
            msg["Subject"] = "[TEST] " + msg["Subject"]
            
            srv.send_message(msg)
            srv.quit()
            Clock.schedule_once(lambda x: self.msg("Test OK", f"Wiadomo\u015b\u0107 testowa wys\u0142ana na: {cfg['u']}"))
        except Exception as e:
            Clock.schedule_once(lambda x: self.msg("B\u0142\u0105d Testu", str(e)))

    # --- WSP\u00d3LNY BUDOWNICZY WIADOMO\u015aCI ---
    def build_message(self, sender, recipient, header, row):
        msg = EmailMessage()
        msg["Subject"] = f"Raport Future - {row[0]}"
        msg["From"] = sender
        msg["To"] = recipient
        msg.set_content(f"Dzie\u0144 dobry {row[0]},\n\nTo jest automatyczna wiadomo\u015b\u0107 systemowa z raportem miesi\u0119cznym.")
        
        # 1. Raport indywidualny
        tmp = Path(self.user_data_dir) / "temp.xlsx"
        wb = Workbook(); ws = wb.active; ws.append(header); ws.append(row); wb.save(str(tmp))
        with open(tmp, "rb") as f:
            msg.add_attachment(f.read(), maintype="application", subtype="xlsx", filename=f"Raport_{row[0]}.xlsx")
        
        # 2. Za\u0142\u0105czniki globalne
        for path in self.global_attachments:
            if os.path.exists(path):
                ctype, _ = mimetypes.guess_type(path)
                maintype, subtype = (ctype or 'application/octet-stream').split('/', 1)
                with open(path, "rb") as f:
                    msg.add_attachment(f.read(), maintype=maintype, subtype=subtype, filename=os.path.basename(path))
        return msg

    def start_mailing(self, _):
        if self.get_total_attachments_size() > 22:
            self.msg("UWAGA", "Przekroczono limit rozmiaru maila (22MB). Serwer mo\u017ce odrzuci\u0107 wysy\u0142k\u0119!"); return
        threading.Thread(target=self._mailing_process).start()

    def _mailing_process(self):
        # ... Logika masowej wysy\u0142ki z u\u017cyciem self.build_message() dla ka\u017cdego wiersza
        # Implementacja analogiczna do 9.3, wywo\u0142uj\u0105ca build_message
        pass

    # --- POPUP ZARZ\u0104DZANIA ---
    def attachment_manager_popup(self, _):
        box = BoxLayout(orientation="vertical", padding=dp(10), spacing=dp(10))
        self.att_list = GridLayout(cols=1, size_hint_y=None, spacing=dp(5))
        self.att_list.bind(minimum_height=self.att_list.setter('height'))
        
        self.refresh_att_list()
        
        sv = ScrollView(); sv.add_widget(self.att_list)
        box.add_widget(sv)
        
        btn_add = Button(text="+ DODAJ PLIK", size_hint_y=None, height=dp(50), background_color=(0, 0.6, 0, 1))
        btn_add.bind(on_press=lambda x: self.pick_file("extra"))
        
        btn_close = Button(text="ZATWIERD\u0179", size_hint_y=None, height=dp(50))
        box.add_widget(btn_add); box.add_widget(btn_close)
        
        self.pop_att = Popup(title="Zarz\u0105dzaj plikami", content=box, size_hint=(0.9, 0.8))
        btn_close.bind(on_press=self.pop_att.dismiss); self.pop_att.open()

    def refresh_att_list(self):
        self.att_list.clear_widgets()
        for p in self.global_attachments:
            row = BoxLayout(size_hint_y=None, height=dp(40))
            row.add_widget(Label(text=os.path.basename(p)[:20], font_size=12))
            btn = Button(text="X", size_hint_x=0.2, background_color=(1, 0, 0, 1))
            btn.bind(on_press=lambda x, path=p: self.remove_att(path))
            row.add_widget(btn); self.att_list.add_widget(row)
        self.update_size_display()

    def remove_att(self, path):
        if path in self.global_attachments: self.global_attachments.remove(path)
        self.refresh_att_list()

    # --- KONSTRUKCJA UI POZOSTA\u0141YCH (Jak w 9.2/9.3) ---
    def setup_home(self): pass 
    def setup_table(self): pass
    def setup_smtp(self): pass
    def pick_file(self, mode): pass # Kod mechanizmu jnius z wersji 9.1
    def msg(self, t, txt): Popup(title=t, content=Label(text=txt), size_hint=(0.8, 0.3)).open()

if __name__ == "__main__":
    FutureApp().run()
