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

from kivy.app import App
from kivy.metrics import dp
from kivy.clock import Clock
from kivy.utils import platform, get_color_from_hex
from kivy.core.window import Window
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.button import Button
from kivy.uix.label import Label
from kivy.uix.popup import Popup
from kivy.uix.scrollview import ScrollView
from kivy.uix.gridlayout import GridLayout
from kivy.uix.textinput import TextInput
from kivy.uix.checkbox import CheckBox
from kivy.uix.screenmanager import ScreenManager, Screen, FadeTransition
from kivy.uix.progressbar import ProgressBar
from kivy.graphics import Color, RoundedRectangle

# Obsługa Excel
try:
    from openpyxl import load_workbook, Workbook
    from openpyxl.styles import Border, Side, Font, Alignment, PatternFill
except ImportError:
    load_workbook = Workbook = None

# KOLORY FUTURE 2.0
CLR_PRIMARY = get_color_from_hex("#3B82F6")
CLR_BG = get_color_from_hex("#0F172A")
CLR_CARD = get_color_from_hex("#1E293B")
CLR_TEXT = get_color_from_hex("#F8FAFC")
CLR_SUBTEXT = get_color_from_hex("#94A3B8")
CLR_DANGER = get_color_from_hex("#EF4444")
CLR_SUCCESS = get_color_from_hex("#10B981")

class StyledButton(Button):
    def __init__(self, bg_color=CLR_PRIMARY, **kwargs):
        super().__init__(**kwargs)
        self.background_normal = ""
        self.background_color = (0,0,0,0)
        self.bold = True
        self.real_bg = bg_color
        self.bind(pos=self._update, size=self._update)
    def _update(self, *args):
        self.canvas.before.clear()
        with self.canvas.before:
            Color(*(self.real_bg if self.state == 'normal' else [c*0.8 for c in self.real_bg[:3]] + [1]))
            RoundedRectangle(pos=self.pos, size=self.size, radius=[dp(10)])

class StyledInput(TextInput):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.background_normal = ""
        self.background_active = ""
        self.background_color = (0.1, 0.15, 0.25, 1)
        self.foreground_color = CLR_TEXT
        self.cursor_color = CLR_PRIMARY
        self.padding = [dp(10), dp(10)]
        self.font_size = '16sp'
        self.size_hint_y = None
        self.height = dp(45) # NAPRAWA: To zapobiegnie wielkim polom ze zdjęcia
        self.hint_text_color = CLR_SUBTEXT

class FutureApp(App):
    def build(self):
        # NAPRAWA: Keyboard nie zasłania pól
        Window.softinput_mode = "below_target"
        Window.clearcolor = CLR_BG
        self.full_data = []; self.idx_name, self.idx_surname, self.idx_pesel = 0, 1, -1
        self.init_db()
        self.sm = ScreenManager(transition=FadeTransition())
        self.setup_screens()
        return self.sm

    def init_db(self):
        try:
            db_p = Path(self.user_data_dir) / "future_v3.db"
            self.conn = sqlite3.connect(str(db_p), check_same_thread=False)
            self.conn.execute("CREATE TABLE IF NOT EXISTS contacts (name TEXT, surname TEXT, email TEXT, pesel TEXT, phone TEXT, PRIMARY KEY(name, surname))")
            self.conn.execute("CREATE TABLE IF NOT EXISTS reports (id INTEGER PRIMARY KEY AUTOINCREMENT, date TEXT, ok INTEGER, fail INTEGER, skip INTEGER, details TEXT)")
            self.conn.execute("CREATE TABLE IF NOT EXISTS settings (key TEXT PRIMARY KEY, val TEXT)")
            # Bezpieczna migracja
            try: self.conn.execute("ALTER TABLE reports ADD COLUMN details TEXT")
            except: pass
            self.conn.commit()
        except Exception as e:
            print(f"DB Error: {e}")

    def setup_screens(self):
        # EKRAN GŁÓWNY (UPROSZCZONY)
        s_home = Screen(name="home"); l = BoxLayout(orientation="vertical", padding=dp(30), spacing=dp(15))
        l.add_widget(Label(text="PASKI-FUTURE 2.0", font_size='28sp', bold=True, color=CLR_PRIMARY))
        l.add_widget(StyledButton(text="WCZYTAJ EXCEL", on_press=lambda x: self.open_picker("data")))
        l.add_widget(StyledButton(text="KONTAKTY", on_press=lambda x: setattr(self.sm, 'current', 'contacts')))
        l.add_widget(StyledButton(text="WYSYŁKA", on_press=lambda x: setattr(self.sm, 'current', 'email')))
        l.add_widget(StyledButton(text="GMAIL", on_press=lambda x: setattr(self.sm, 'current', 'smtp'), real_bg=CLR_CARD))
        s_home.add_widget(l); self.sm.add_widget(s_home)

        # EKRAN KONTAKTÓW (Z TWOJEGO ZDJĘCIA)
        self.s_contacts = Screen(name="contacts"); scroll = ScrollView()
        self.c_list = GridLayout(cols=1, size_hint_y=None, spacing=dp(5), padding=dp(10))
        self.c_list.bind(minimum_height=self.c_list.setter('height'))
        cl = BoxLayout(orientation="vertical")
        top = BoxLayout(size_hint_y=None, height=dp(55), padding=dp(5), spacing=dp(5))
        top.add_widget(StyledButton(text="+ DODAJ", on_press=lambda x: self.form_contact()))
        top.add_widget(StyledButton(text="COFNIJ", on_press=lambda x: setattr(self.sm, 'current', 'home'), real_bg=CLR_CARD))
        cl.add_widget(top); scroll.add_widget(self.c_list); cl.add_widget(scroll)
        self.s_contacts.add_widget(cl); self.sm.add_widget(self.s_contacts)
        
        # DODATKOWE EKRANY (GMAIL / SMTP / WYSYŁKA) - uproszczone stuby dla stabilności startu
        self.sm.add_widget(Screen(name="email"))
        self.sm.add_widget(Screen(name="smtp"))

    def form_contact(self, n="", s="", e="", pes="", ph=""):
        # POPUP NAPRAWIONY (ZGODNIE ZE ZDJĘCIEM)
        root = BoxLayout(orientation="vertical", padding=dp(15), spacing=dp(10))
        scroll = ScrollView(size_hint_y=0.8); b = BoxLayout(orientation="vertical", size_hint_y=None, spacing=dp(8))
        b.bind(minimum_height=b.setter('height'))
        inputs = []
        for label_text, value in [("IMIĘ", n), ("NAZWISKO", s), ("EMAIL", e), ("PESEL", pes), ("TEL", ph)]:
            box = BoxLayout(orientation="vertical", size_hint_y=None, height=dp(65))
            box.add_widget(Label(text=label_text, font_size='11sp', color=CLR_PRIMARY, bold=True, halign='left', text_size=(dp(250), None)))
            ti = StyledInput(text=value); box.add_widget(ti); inputs.append(ti); b.add_widget(box)
        scroll.add_widget(b); root.add_widget(scroll)
        btns = BoxLayout(size_hint_y=None, height=dp(50), spacing=dp(10))
        btns.add_widget(StyledButton(text="ZAPISZ", on_press=lambda x: self.save_contact(inputs, p), real_bg=CLR_SUCCESS))
        btns.add_widget(StyledButton(text="ANULUJ", on_press=lambda x: p.dismiss(), real_bg=CLR_CARD))
        root.add_widget(btns); p = Popup(title="Karta Kontaktu", content=root, size_hint=(0.95, 0.9)); p.open()

    def save_contact(self, inputs, popup):
        d = [i.text.strip() for i in inputs]
        self.conn.execute("INSERT OR REPLACE INTO contacts VALUES (?,?,?,?,?)", d); self.conn.commit()
        popup.dismiss(); self.refresh_contacts_list()

    def refresh_contacts_list(self):
        self.c_list.clear_widgets()
        rows = self.conn.execute("SELECT name, surname, email FROM contacts").fetchall()
        for n, s, e in rows:
            btn = Button(text=f"{n} {s}\n{e}", size_hint_y=None, height=dp(60), background_color=CLR_CARD)
            self.c_list.add_widget(btn)

    def open_picker(self, mode):
        # Wybierak plików Android
        if platform == 'android':
            from jnius import autoclass; from android import activity
            PA = autoclass("org.kivy.android.PythonActivity"); Intent = autoclass("android.content.Intent")
            intent = Intent(Intent.ACTION_GET_CONTENT); intent.setType("*/*")
            PA.mActivity.startActivityForResult(intent, 1001)

if __name__ == "__main__":
    FutureApp().run()
