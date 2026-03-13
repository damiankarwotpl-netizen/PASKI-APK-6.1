# v19 — 2026-03-13
import os
import json
import sqlite3
import threading
import smtplib
import mimetypes
import time
import random
import traceback
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

try:
    from reportlab.pdfgen import canvas
except:
    canvas = None

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
        self.background_color = (0,0,0,0)
        self.color = COLOR_TEXT
        self.bold, self.radius = True, [dp(12)]
        with self.canvas.before:
            Color(*bg_color)
            self.rect = RoundedRectangle(pos=self.pos, size=self.size, radius=self.radius)
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
    def _update(self, inst, val):
        self.rect.size, self.rect.pos = self.size, self.pos
        self.text_size = (self.width - dp(10), None)

class ClothesSizesScreen(Screen):
    def on_enter(self):
        if not hasattr(self, 'built'):
            self.build_ui()
        self.refresh()

    def build_ui(self):
        try:
            self.clear_widgets()
        except:
            pass
        root = BoxLayout(orientation='vertical')
        top = BoxLayout(size_hint_y=None, height=dp(50), spacing=dp(8), padding=dp(8))
        top.add_widget(Label(text="Rozmiary pracowników", bold=True))
        top.add_widget(Button(text="Import Excel", size_hint_x=None, width=dp(140), on_press=lambda x: self.open_import()))
        root.add_widget(top)
        sc = ScrollView()
        self.list_layout = GridLayout(cols=1, size_hint_y=None, spacing=dp(6), padding=dp(6))
        self.list_layout.bind(minimum_height=self.list_layout.setter('height'))
        sc.add_widget(self.list_layout)
        root.add_widget(sc)
        self.add_widget(root)
        self.built = True

    def refresh(self):
        self.list_layout.clear_widgets()
        rows = App.get_running_app().conn.execute(
            "SELECT id, name, surname, plant, shirt, hoodie, pants, jacket, shoes FROM clothes_sizes ORDER BY surname"
        ).fetchall()
        for r in rows:
            box = BoxLayout(size_hint_y=None,height=dp(80), padding=dp(6))
            txt = f"{r[1]} {r[2]} ({r[3]})  K:{r[4]} B:{r[5]} S:{r[6]} KUR:{r[7]} BUT:{r[8]}"
            box.add_widget(Label(text=txt))
            box.add_widget(Button(
                text="Edytuj",
                size_hint_x=0.2,
                on_press=lambda x,data=r:self.edit(data)
            ))
            self.list_layout.add_widget(box)

    def edit(self,row):
        box = BoxLayout(orientation="vertical", spacing=dp(6), padding=dp(6))
        fields = []
        labels = ["Imię","Nazwisko","Zakład","Koszulka","Bluza","Spodnie","Kurtka","Buty"]
        for i in range(1,9):
            box.add_widget(Label(text=labels[i-1], size_hint_y=None, height=dp(24)))
            ti = TextInput(text=str(row[i]), multiline=False)
            fields.append(ti)
            box.add_widget(ti)
        def save(_):
            App.get_running_app().conn.execute("""
            UPDATE clothes_sizes
            SET name=?,surname=?,plant=?,shirt=?,hoodie=?,pants=?,jacket=?,shoes=?
            WHERE id=?
            """,(
                fields[0].text,
                fields[1].text,
                fields[2].text,
                fields[3].text,
                fields[4].text,
                fields[5].text,
                fields[6].text,
                fields[7].text,
                row[0]
            ))
            App.get_running_app().conn.commit()
            popup.dismiss()
            self.refresh()
        box.add_widget(Button(text="ZAPISZ", size_hint_y=None, height=dp(44), on_press=save))
        popup = Popup(title="Edycja",content=box,size_hint=(0.9,0.9))
        popup.open()

    def open_import(self):
        def pick_file(_):
            popup.dismiss()
            self.show_input_for_import()
        box = BoxLayout(orientation='vertical')
        box.add_widget(Label(text="Kliknij aby podać nazwę pliku Excel w katalogu aplikacji"))
        box.add_widget(Button(text="OK", on_press=pick_file))
        popup = Popup(title="Import", content=box, size_hint=(0.8,0.3))
        popup.open()

    def show_input_for_import(self):
        box = BoxLayout(orientation='vertical', spacing=dp(6))
        ti = TextInput(hint_text="nazwa pliku (np. rozmiary.xlsx)")
        box.add_widget(ti)
        def run(_):
            name = ti.text.strip()
            if not name: return
            path = Path(App.get_running_app().user_data_dir)/name
            if path.exists():
                import_clothes_excel(path)
                popup.dismiss()
                self.refresh()
                App.get_running_app().log(f"Imported clothes excel: {path}")
            else:
                App.get_running_app().msg("Błąd", "Plik nie istnieje w katalogu aplikacji")
        box.add_widget(Button(text="IMPORT", on_press=run))
        popup = Popup(title="Import Excel", content=box, size_hint=(0.9,0.4))
        popup.open()

class ClothesOrdersScreen(Screen):
    def on_enter(self):
        if not hasattr(self, 'built'):
            self.build_ui()
        self.refresh()

    def build_ui(self):
        try:
            self.clear_widgets()
        except:
            pass
        root = BoxLayout(orientation='vertical', spacing=dp(6))
        header = BoxLayout(size_hint_y=None, height=dp(50))
        header.add_widget(Label(text="Zamówienia", bold=True))
        header.add_widget(Button(text="Nowe zamówienie", size_hint_x=None, width=dp(180), on_press=lambda x: self.create_order()))
        root.add_widget(header)
        sc = ScrollView()
        self.list_layout = GridLayout(cols=1, size_hint_y=None, spacing=dp(6), padding=dp(6))
        self.list_layout.bind(minimum_height=self.list_layout.setter('height'))
        sc.add_widget(self.list_layout)
        root.add_widget(sc)
        self.add_widget(root)
        self.built = True

    def create_order(self):
        db = App.get_running_app()
        db.conn.execute("""
        INSERT INTO clothes_orders(date,plant,status)
        VALUES (?,?,?)
        """,(datetime.now().strftime("%Y-%m-%d"),"Zakład","Do zamówienia"))
        db.conn.commit()
        self.refresh()

    def refresh(self):
        self.list_layout.clear_widgets()
        rows = App.get_running_app().conn.execute("""
        SELECT id,date,plant,status FROM clothes_orders ORDER BY id DESC
        """).fetchall()
        for r in rows:
            box = BoxLayout(size_hint_y=None,height=dp(70), padding=dp(6))
            box.add_widget(Label(text=f"Zamówienie #{r[0]}  {r[2]}  {r[3]}"))
            box.add_widget(Button(
                text="Zmień",
                size_hint_x=0.2,
                on_press=lambda x,i=r[0]:self.change(i)
            ))
            self.list_layout.add_widget(box)

    def change(self,id):
        db = App.get_running_app()
        db.conn.execute("""
        UPDATE clothes_orders
        SET status='Zamówione'
        WHERE id=?
        """,(id,))
        db.conn.commit()
        self.refresh()

class ClothesStatusScreen(Screen):
    def on_enter(self):
        if not hasattr(self, 'built'):
            self.build_ui()
        self.refresh()

    def build_ui(self):
        try:
            self.clear_widgets()
        except:
            pass
        root = BoxLayout(orientation='vertical')
        sc = ScrollView()
        self.list_layout = GridLayout(cols=1, size_hint_y=None, spacing=dp(6), padding=dp(6))
        self.list_layout.bind(minimum_height=self.list_layout.setter('height'))
        sc.add_widget(self.list_layout)
        root.add_widget(sc)
        self.add_widget(root)
        self.built = True

    def refresh(self):
        self.list_layout.clear_widgets()
        rows = App.get_running_app().conn.execute("""
        SELECT id,date,plant,status FROM clothes_orders ORDER BY id DESC
        """).fetchall()
        for r in rows:
            box = BoxLayout(size_hint_y=None,height=dp(70), padding=dp(6))
            box.add_widget(Label(text=f"Zamówienie #{r[0]}  {r[2]}  {r[3]}"))
            box.add_widget(Button(text="Zmień", size_hint_x=0.2, on_press=lambda x,i=r[0]: self.change(i)))
            self.list_layout.add_widget(box)

    def change(self,id):
        db = App.get_running_app()
        db.conn.execute("""
        UPDATE clothes_orders
        SET status='Zamówione'
        WHERE id=?
        """,(id,))
        db.conn.commit()
        self.refresh()

class ClothesReportsScreen(Screen):
    def on_enter(self):
        if not hasattr(self, 'built'):
            self.build_ui()

    def build_ui(self):
        try:
            self.clear_widgets()
        except:
            pass
        root = BoxLayout(orientation='vertical', padding=dp(6), spacing=dp(6))
        header = BoxLayout(size_hint_y=None, height=dp(50))
        header.add_widget(Label(text="Raporty wydanych ubrań", bold=True))
        header.add_widget(Button(text="Generuj PDF", size_hint_x=None, width=dp(160), on_press=lambda x: self.generate()))
        root.add_widget(header)
        self.add_widget(root)
        self.built = True

    def generate(self):
        if canvas is None:
            App.get_running_app().msg("Brak biblioteki", "Brak reportlab - PDF niedostępny")
            return
        db = App.get_running_app()
        rows = db.conn.execute("""
        SELECT name,surname,item,SUM(qty)
        FROM clothes_issued
        GROUP BY name,surname,item
        """).fetchall()
        path = Path(db.user_data_dir)/"raport_clothes.pdf"
        c = canvas.Canvas(str(path))
        y = 800
        for r in rows:
            txt = f"{r[0]} {r[1]} {r[2]} {r[3]}"
            c.drawString(50,y,txt)
            y -= 20
            if y < 50:
                c.showPage()
                y = 800
        c.save()
        App.get_running_app().msg("OK", f"Zapisano: {path.name}")
        db.log(f"Generated clothes report: {path}")

def import_clothes_excel(path):
    if load_workbook is None:
        return
    db = App.get_running_app()
    wb = load_workbook(path)
    ws = wb.active
    for r in ws.iter_rows(min_row=2,values_only=True):
        vals = tuple("" if v is None else str(v) for v in r[:8])
        if len(vals) < 8:
            vals = vals + tuple("" for _ in range(8 - len(vals)))
        db.conn.execute("""
        INSERT INTO clothes_sizes
        (name,surname,plant,shirt,hoodie,pants,jacket,shoes)
        VALUES (?,?,?,?,?,?,?,?)
        """,vals)
    db.conn.commit()

class FutureApp(App):
    def build(self):
        Window.clearcolor = COLOR_BG
        if platform == "android":
            from android.permissions import request_permissions, Permission
            request_permissions([Permission.READ_EXTERNAL_STORAGE, Permission.WRITE_EXTERNAL_STORAGE])
        if not os.path.exists(self.user_data_dir): os.makedirs(self.user_data_dir, exist_ok=True)

        self.full_data, self.filtered_data, self.export_indices = [], [], []
        self.global_attachments, self.queue = [], []
        self.stats = {"ok": 0, "fail": 0, "skip": 0}
        self.idx_name, self.idx_surname, self.idx_pesel = 0, 1, -1
        self.auto_send_mode = False
        self.is_mailing_running = False
        self.mailing_paused = False
        self._log_buffer = []

        self.init_db()
        self.log_file = Path(self.user_data_dir) / "future_v19.log"
        try:
            self.log_file.touch(exist_ok=True)
        except:
            pass

        self.sm = ScreenManager(transition=SlideTransition())
        self.add_screens()
        return self.sm

    def log(self, txt):
        try:
            t = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            line = f"[{t}] {txt}\n"
            self._log_buffer.append(line)
            if len(self._log_buffer) > 200:
                self._log_buffer = self._log_buffer[-200:]
            with open(self.log_file, "a", encoding="utf-8") as f:
                f.write(line)
        except:
            pass

    def init_db(self):
        db_p = Path(self.user_data_dir) / "future_v19.db"
        self.conn = sqlite3.connect(str(db_p), check_same_thread=False)
        self.conn.execute("CREATE TABLE IF NOT EXISTS contacts (name TEXT, surname TEXT, email TEXT, pesel TEXT, phone TEXT, PRIMARY KEY(name, surname))")
        self.conn.execute("CREATE TABLE IF NOT EXISTS settings (key TEXT PRIMARY KEY, val TEXT)")
        self.conn.execute("CREATE TABLE IF NOT EXISTS reports (id INTEGER PRIMARY KEY AUTOINCREMENT, date TEXT, ok INTEGER, fail INTEGER, skip INTEGER, auto INTEGER, details TEXT)")
        self.conn.execute("""
        CREATE TABLE IF NOT EXISTS clothes_sizes(
        id INTEGER PRIMARY KEY,
        name TEXT,
        surname TEXT,
        plant TEXT,
        shirt TEXT,
        hoodie TEXT,
        pants TEXT,
        jacket TEXT,
        shoes TEXT
        )
        """)
        self.conn.execute("""
        CREATE TABLE IF NOT EXISTS clothes_orders(
        id INTEGER PRIMARY KEY,
        date TEXT,
        plant TEXT,
        status TEXT
        )
        """)
        self.conn.execute("""
        CREATE TABLE IF NOT EXISTS clothes_order_items(
        id INTEGER PRIMARY KEY,
        order_id INTEGER,
        name TEXT,
        surname TEXT,
        item TEXT,
        size TEXT,
        qty INTEGER,
        issued INTEGER DEFAULT 0
        )
        """)
        self.conn.execute("""
        CREATE TABLE IF NOT EXISTS clothes_issued(
        id INTEGER PRIMARY KEY,
        name TEXT,
        surname TEXT,
        item TEXT,
        size TEXT,
        qty INTEGER,
        date TEXT
        )
        """)
        self.conn.commit()

    def add_screens(self):
        names = ["home", "table", "email", "smtp", "tmpl", "contacts", "report", "cars", "clothes", "paski", "pracownicy", "zaklady", "settings"]
        self.sc_ref = {name: Screen(name=name) for name in names}
        self.setup_ui_all()
        for s in self.sc_ref.values():
            self.sm.add_widget(s)

    def setup_ui_all(self):
        self.sc_ref["home"].clear_widgets()
        root = BoxLayout(orientation="vertical", padding=dp(10), spacing=dp(10))
        lbl = Label(text="FUTURE ULTIMATE v19", font_size='34sp', bold=True, color=COLOR_PRIMARY, size_hint_y=None, height=dp(70))
        root.add_widget(lbl)
        sv = ScrollView(size_hint=(1, None), size=(Window.width, dp(300)))
        grid = GridLayout(cols=2, spacing=dp(12), padding=dp(10), size_hint_y=None)
        grid.bind(minimum_height=grid.setter('height'))
        btn_props = dict(size_hint_y=None, height=dp(80))
        grid.add_widget(ModernButton(text="Kontakty", on_press=lambda x: [self.refresh_contacts_list(), setattr(self.sm, 'current', 'contacts')], **btn_props))
        grid.add_widget(ModernButton(text="Samochody", on_press=lambda x: setattr(self.sm, 'current', 'cars'), **btn_props))
        grid.add_widget(ModernButton(text="Ubranie robocze", on_press=lambda x: setattr(self.sm, 'current', 'clothes'), **btn_props))
        grid.add_widget(ModernButton(text="Paski", on_press=lambda x: setattr(self.sm, 'current', 'paski'), **btn_props))
        grid.add_widget(ModernButton(text="Pracownicy", on_press=lambda x: setattr(self.sm, 'current', 'pracownicy'), **btn_props))
        grid.add_widget(ModernButton(text="Zakłady", on_press=lambda x: setattr(self.sm, 'current', 'zaklady'), **btn_props))
        grid.add_widget(ModernButton(text="Ustawienia", on_press=lambda x: setattr(self.sm, 'current', 'settings'), **btn_props))
        grid.add_widget(ModernButton(text="Wyjście", on_press=lambda x: App.get_running_app().stop(), bg_color=(0.6,0.1,0.1,1), **btn_props))
        sv.add_widget(grid)
        root.add_widget(sv)
        self.sc_ref["home"].add_widget(root)
        self.setup_table_ui(); self.setup_email_ui(); self.setup_smtp_ui(); self.setup_tmpl_ui(); self.setup_contacts_ui(); self.setup_report_ui()
        self.setup_cars_ui(); self.setup_paski_ui(); self.setup_pracownicy_ui(); self.setup_zaklady_ui(); self.setup_settings_ui()
        self.setup_clothes_container()

    def setup_clothes_container(self):
        self.sc_ref["clothes"].clear_widgets()
        container = BoxLayout(orientation='vertical')
        top = BoxLayout(size_hint_y=None, height=dp(50), spacing=dp(6), padding=dp(6))
        top.add_widget(Label(text="Ubrania robocze", bold=True))
        self.clothes_sm = ScreenManager(transition=SlideTransition())
        top.add_widget(Button(text="Rozmiary", size_hint_x=None, width=dp(120), on_press=lambda x: setattr(self.clothes_sm, 'current', 'sizes')))
        top.add_widget(Button(text="Zamówienia", size_hint_x=None, width=dp(120), on_press=lambda x: setattr(self.clothes_sm, 'current', 'orders')))
        top.add_widget(Button(text="Status", size_hint_x=None, width=dp(120), on_press=lambda x: setattr(self.clothes_sm, 'current', 'status')))
        top.add_widget(Button(text="Raporty", size_hint_x=None, width=dp(120), on_press=lambda x: setattr(self.clothes_sm, 'current', 'reports')))
        container.add_widget(top)
        self.clothes_sm.add_widget(ClothesSizesScreen(name='sizes'))
        self.clothes_sm.add_widget(ClothesOrdersScreen(name='orders'))
        self.clothes_sm.add_widget(ClothesStatusScreen(name='status'))
        self.clothes_sm.add_widget(ClothesReportsScreen(name='reports'))
        container.add_widget(self.clothes_sm)
        self.sc_ref["clothes"].add_widget(container)

    def setup_table_ui(self):
        self.sc_ref["table"].clear_widgets()
        root = BoxLayout(orientation="vertical")
        menu = BoxLayout(size_hint_y=None, height=dp(55), spacing=dp(5), padding=dp(5))
        self.ti_tab_search = ModernInput(hint_text="Szukaj w tabeli..."); self.ti_tab_search.bind(text=self.filter_table)
        menu.add_widget(self.ti_tab_search)
        menu.add_widget(Button(text="KOLUMNY", size_hint_x=0.2, on_press=self.popup_columns))
        menu.add_widget(Button(text="WRÓĆ", size_hint_x=0.2, on_press=lambda x: setattr(self.sm, 'current', 'home')))
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
        for head in headers: self.table_header_layout.add_widget(ColorSafeLabel(text=str(head), bg_color=COLOR_HEADER, bold=True, size=(w_cell, h), size_hint=(None,None), text_color=(0,0,0,1)))
        self.table_header_layout.add_widget(ColorSafeLabel(text="AKCJE", bg_color=COLOR_HEADER, bold=True, size=(w_act, h), size_hint=(None,None), text_color=(0,0,0,1)))
        for r_idx, row in enumerate(self.filtered_data[1:]):
            row_bg = COLOR_ROW_A if r_idx % 2 == 0 else COLOR_ROW_B
            for c_idx in self.export_indices:
                val = str(row[c_idx]) if c_idx < len(row) and str(row[c_idx]).strip() != "" else "0"
                self.table_content_layout.add_widget(ColorSafeLabel(text=val, bg_color=row_bg, size=(w_cell, h), size_hint=(None,None)))
            act_box = BoxLayout(size=(w_act, h), size_hint=(None,None), spacing=dp(4), padding=dp(4))
            act_box.add_widget(Button(text="ZAPISZ", on_press=lambda x, r=row: self.export_single_row(r), background_color=(0.2, 0.6, 0.2, 1)))
            act_box.add_widget(Button(text="WYŚLIJ", on_press=lambda x, r=row: self.send_individual_from_table(r), background_color=(0.1, 0.5, 0.9, 1)))
            self.table_content_layout.add_widget(act_box)

    def process_excel(self, path):
        try:
            if str(path).endswith(".xls") and xlrd:
                wb = xlrd.open_workbook(path); ws = wb.sheet_by_index(0); raw = [[str(ws.cell_value(r,c)).strip() for c in range(ws.ncols)] for r in range(ws.nrows)]
            else:
                wb = load_workbook(path, data_only=True); ws = wb.active; raw = [["" if v is None else str(v).strip() for v in r] for r in ws.iter_rows(values_only=True)]
            h_idx = 0
            for i, r in enumerate(raw[:15]):
                if any(x in " ".join([str(v) for v in r]).lower() for x in ["imię", "imie", "nazwisko"]): h_idx = i; break
            self.full_data = raw[h_idx:]
            self.filtered_data = self.full_data
            self.export_indices = list(range(len(self.full_data[0])))
            for i,v in enumerate(self.full_data[0]):
                v = str(v).lower()
                if "imi" in v: self.idx_name = i
                if "naz" in v: self.idx_surname = i
                if "pesel" in v: self.idx_pesel = i
            self.msg("OK", "Arkusz wczytany")
            self.log(f"Loaded excel: {path}")
        except Exception as e:
            self.log(f"process_excel error: {traceback.format_exc()}")
            self.msg("BŁĄD", "Plik uszkodzony")

    def send_individual_from_table(self, row):
        name, sur = str(row[self.idx_name]).strip(), str(row[self.idx_surname]).strip()
        pes = str(row[self.idx_pesel]).strip() if self.idx_pesel != -1 else ""
        res = self.conn.execute("SELECT email FROM contacts WHERE pesel=? AND pesel != ''", (pes,)).fetchone() if pes else None
        if not res: res = self.conn.execute("SELECT email FROM contacts WHERE name=? AND surname=? COLLATE NOCASE", (name.lower(), sur.lower())).fetchone()
        if not res: return self.msg("Błąd", f"Brak maila dla: {name}")
        def task():
            cfg_p = Path(self.user_data_dir)/"smtp.json"
            if not cfg_p.exists(): return Clock.schedule_once(lambda d: self.msg("!", "Brak SMTP"))
            cfg = json.load(open(cfg_p)); srv = self.connect_smtp(cfg)
            if self.send_single_email(srv, cfg, row, res[0]): Clock.schedule_once(lambda d: self.msg("OK", f"Wysłano do: {name}"))
            srv.quit()
        threading.Thread(target=task, daemon=True).start()

    def setup_email_ui(self):
        self.sc_ref["email"].clear_widgets()
        l = BoxLayout(orientation="vertical", padding=dp(25), spacing=dp(10))
        ab = BoxLayout(size_hint_y=None, height=dp(45), spacing=dp(10))
        self.cb_auto = CheckBox(size_hint_x=None, width=dp(45))
        self.cb_auto.active = self.auto_send_mode
        self.cb_auto.bind(active=self.on_auto_checkbox_changed)
        ab.add_widget(self.cb_auto); ab.add_widget(Label(text="AUTOMATYCZNA WYSYŁKA", bold=True)); l.add_widget(ab)
        self.lbl_stats = Label(text="Baza: 0", height=dp(30)); l.add_widget(self.lbl_stats)
        l.add_widget(ModernButton(text="WYCZYŚĆ ZAŁĄCZNIKI", on_press=self.clear_all_attachments, height=dp(45), size_hint_y=None, bg_color=(0.7, 0.1, 0.1, 1)))
        self.pb_label = Label(text="Gotowy", height=dp(25)); self.pb = ProgressBar(max=100, height=dp(20)); l.add_widget(self.pb_label); l.add_widget(self.pb)
        btns = [("EDYTUJ SZABLON", lambda x: setattr(self.sm, 'current', 'tmpl')), ("DODAJ ZAŁĄCZNIK", lambda x: self.open_picker("attachment")), ("WYŚLIJ JEDEN PLIK", self.start_special_send_flow), ("START MASOWA WYSYŁKA", self.start_mass_mailing)]
        for t, c in btns: l.add_widget(ModernButton(text=t, on_press=c, height=dp(50), size_hint_y=None))
        l.add_widget(ModernButton(text="PAUZA/RESUME WYSYŁKI", on_press=self.toggle_pause_mailing, height=dp(50), size_hint_y=None, bg_color=(0.6,0.6,0.1,1)))
        l.add_widget(ModernButton(text="POWRÓT", on_press=lambda x: setattr(self.sm, 'current', 'home'), bg_color=(0.3,0.3,0.3,1))); self.sc_ref["email"].add_widget(l); self.update_stats()

    def on_auto_checkbox_changed(self, instance, value):
        self.auto_send_mode = bool(value)
        try:
            if hasattr(self, 'cb_paski_auto') and self.cb_paski_auto.active != value:
                self.cb_paski_auto.active = value
        except: pass

    def process_book(self, path):
        try:
            wb = load_workbook(path, data_only=True); ws = wb.active; raw = list(ws.iter_rows(values_only=True))
            if not raw or not raw[0]:
                self.msg("Błąd", "Pusty plik")
                return
            headers = ["" if v is None else str(v).strip() for v in raw[0]]
            h_low = [h.lower() for h in headers]
            iN = iS = iE = iP = iPhone = -1
            for i, v in enumerate(h_low):
                if iN == -1 and ("imi" in v or v == "name"): iN = i
                if iS == -1 and ("naz" in v or v == "surname" or "nazw" in v): iS = i
                if iE == -1 and ("@" in v or "mail" in v): iE = i
                if iP == -1 and "pesel" in v: iP = i
                if iPhone == -1 and any(x in v for x in ["tel", "phone", "telefon"]): iPhone = i
            car_keys = ["rej", "rejestr", "plate", "nr rejestr", "nr rej", "vin", "marka", "model", "brand", "car", "samoch"]
            clothes_keys = ["ubran", "odziez", "rozmiar", "size", "typ", "kolor", "clothe", "garment"]
            car_cols = []
            clothes_cols = []
            for i, v in enumerate(h_low):
                if any(k in v for k in car_keys): car_cols.append((i, headers[i]))
                if any(k in v for k in clothes_keys): clothes_cols.append((i, headers[i]))
            if iE != -1:
                for r in raw[1:]:
                    try:
                        e = r[iE] if iE < len(r) else None
                        if e and "@" in str(e):
                            n = r[iN] if iN < len(r) and iN != -1 else ""
                            s = r[iS] if iS < len(r) and iS != -1 else ""
                            p = r[iP] if iP < len(r) and iP != -1 else ""
                            ph = r[iPhone] if iPhone < len(r) and iPhone != -1 else ""
                            self.conn.execute("INSERT OR REPLACE INTO contacts VALUES (?,?,?,?,?)", (str(n).lower(), str(s).lower(), str(e).strip(), str(p) if p is not None else "", str(ph) if ph is not None else ""))
                    except:
                        pass
            if car_cols:
                table_cols = ["id INTEGER PRIMARY KEY AUTOINCREMENT"] + [f"'{self._sanitize_col(c[1])}' TEXT" for c in car_cols]
                self.conn.execute("CREATE TABLE IF NOT EXISTS cars (" + ",".join(table_cols) + ")")
                insert_cols = [self._sanitize_col(c[1]) for c in car_cols]
                q = "INSERT INTO cars (" + ",".join([f"'{c}'" for c in insert_cols]) + ") VALUES (" + ",".join(["?"] * len(insert_cols)) + ")"
                for r in raw[1:]:
                    vals = []
                    for idx, _ in car_cols:
                        vals.append(str(r[idx]) if idx < len(r) and r[idx] is not None else "")
                    if any(v.strip() for v in vals):
                        try:
                            self.conn.execute(q, vals)
                        except:
                            pass
            if clothes_cols:
                table_cols = ["id INTEGER PRIMARY KEY AUTOINCREMENT"] + [f"'{self._sanitize_col(c[1])}' TEXT" for c in clothes_cols]
                self.conn.execute("CREATE TABLE IF NOT EXISTS clothes (" + ",".join(table_cols) + ")")
                insert_cols = [self._sanitize_col(c[1]) for c in clothes_cols]
                q = "INSERT INTO clothes (" + ",".join([f"'{c}'" for c in insert_cols]) + ") VALUES (" + ",".join(["?"] * len(insert_cols)) + ")"
                for r in raw[1:]:
                    vals = []
                    for idx, _ in clothes_cols:
                        vals.append(str(r[idx]) if idx < len(r) and r[idx] is not None else "")
                    if any(v.strip() for v in vals):
                        try:
                            self.conn.execute(q, vals)
                        except:
                            pass
            self.conn.commit()
            new_ver = self._increment_db_version()
            self.update_stats()
            self.msg("OK", f"Baza zaktualizowana. Wersja: {new_ver}")
            self.log(f"Imported book: {path}")
        except Exception as e:
            self.log(f"process_book error: {traceback.format_exc()}")
            self.msg("Błąd", f"Nieudany import: {str(e)[:120]}")

    def _sanitize_col(self, name):
        s = "".join(c if c.isalnum() else "_" for c in str(name).strip().lower())
        s = s.strip("_")
        if not s:
            s = "col"
        return s

    def _increment_db_version(self):
        cur = self.conn.execute("SELECT val FROM settings WHERE key='db_version'").fetchone()
        v = int(cur[0]) if cur and cur[0].isdigit() else 0
        v += 1
        self.conn.execute("INSERT OR REPLACE INTO settings VALUES (?,?)", ('db_version', str(v)))
        self.conn.commit()
        return v

    def mailing_worker(self):
        cfg_p = Path(self.user_data_dir)/"smtp.json"
        if not cfg_p.exists(): return self.finish_mailing("Brak SMTP")
        cfg = json.load(open(cfg_p)); b_on, b_sz, proc = cfg.get('batch', True), 30, 0
        try:
            srv = self.connect_smtp(cfg)
            while self.queue:
                if self.mailing_paused:
                    time.sleep(0.5)
                    continue
                row = self.queue.pop(0); n, s = str(row[self.idx_name]).strip(), str(row[self.idx_surname]).strip()
                p_exc = str(row[self.idx_pesel]).strip() if self.idx_pesel != -1 else ""
                res_p = self.conn.execute("SELECT email FROM contacts WHERE pesel=? AND pesel != ''", (p_exc,)).fetchone() if p_exc else None
                target, vrf = (res_p[0], False) if res_p else (None, False)
                if not target:
                    res_n = self.conn.execute("SELECT email FROM contacts WHERE name=? AND surname=? COLLATE NOCASE", (n.lower(), s.lower())).fetchone()
                    if res_n: target, vrf = res_n[0], not self.auto_send_mode
                if target:
                    if vrf:
                        self.wait_for_user = True; Clock.schedule_once(lambda dt: self.ask_before_send_worker(row, target, n, s))
                        while self.wait_for_user: time.sleep(0.5)
                        if self.user_decision == "skip": continue
                    if self.send_single_email(srv, cfg, row, target): self.stats["ok"] += 1; self.session_details.append(f"OK: {n} {s}")
                    else: self.stats["fail"] += 1; srv.quit(); srv = self.connect_smtp(cfg)
                    proc += 1
                    if self.queue:
                        if b_on and proc >= b_sz: srv.quit(); time.sleep(60); srv = self.connect_smtp(cfg); proc = 0
                        else: time.sleep(random.uniform(3, 7))
                else: self.stats["skip"] += 1; self.session_details.append(f"SKIP: {n} {s}")
                Clock.schedule_once(lambda dt: self.update_progress(self.total_q - len(self.queue)))
            srv.quit(); self.finish_mailing("Zakończono")
        except Exception as e:
            self.log(f"mailing_worker error: {traceback.format_exc()}")
            self.finish_mailing(f"Error: {e}")

    def connect_smtp(self, cfg):
        s = smtplib.SMTP(cfg.get('h','smtp.gmail.com'), int(cfg.get('port',587)), timeout=25); s.starttls(); s.login(cfg['u'], cfg['p']); return s

    def send_single_email(self, srv, cfg, row_data, target):
        try:
            nx, sx = str(row_data[self.idx_name]).title(), str(row_data[self.idx_surname]).title()
            msg = EmailMessage(); ts, tb = self.conn.execute("SELECT val FROM settings WHERE key='t_sub'").fetchone(), self.conn.execute("SELECT val FROM settings WHERE key='t_body'").fetchone()
            msg["Subject"] = (ts[0] if ts else "Raport").replace("{Imię}", nx); msg["From"], msg["To"] = cfg['u'], target
            msg.set_content((tb[0] if tb else "Dzień dobry").replace("{Imię}", nx).replace("{Data}", datetime.now().strftime("%d.%m.%Y")))
            t_f = Path(self.user_data_dir)/f"r_{nx}.xlsx"; wb = Workbook(); ws = wb.active
            ws.append([self.full_data[0][k] for k in self.export_indices]); ws.append([str(row_data[k]) if (str(row_data[k]).strip()!="") else "0" for k in self.export_indices])
            self.style_xlsx(ws); wb.save(t_f)
            with open(t_f, "rb") as f: msg.add_attachment(f.read(), maintype="application", subtype="xlsx", filename=f"Raport_{nx}_{sx}.xlsx")
            for p in self.global_attachments:
                if os.path.exists(p):
                    ct, _ = mimetypes.guess_type(p); mn, sb = (ct or 'application/octet-stream').split('/', 1)
                    with open(p,"rb") as f: msg.add_attachment(f.read(), maintype=mn, subtype=sb, filename=os.path.basename(p))
            srv.send_message(msg); return True
        except Exception:
            self.log(f"send_single_email error: {traceback.format_exc()}")
            return False

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
        self.sc_ref["smtp"].clear_widgets()
        l = BoxLayout(orientation="vertical", padding=dp(25), spacing=dp(8)); p = Path(self.user_data_dir)/"smtp.json"; d = json.load(open(p)) if p.exists() else {}
        self.ti_h, self.ti_pt = ModernInput(hint_text="Host", text=d.get('h','')), ModernInput(hint_text="Port", text=str(d.get('port','587')))
        self.ti_u, self.ti_p = ModernInput(hint_text="Email/Login", text=d.get('u','')), ModernInput(hint_text="Hasło/Klucz", password=True, text=d.get('p',''))
        l.add_widget(Label(text="USTAWIENIA POCZTY", bold=True)); l.add_widget(self.ti_h); l.add_widget(self.ti_pt); l.add_widget(self.ti_u); l.add_widget(self.ti_p)
        bx = BoxLayout(size_hint_y=None, height=dp(45)); self.cb_b = CheckBox(size_hint_x=None, width=dp(45), active=d.get('batch', True)); bx.add_widget(self.cb_b); bx.add_widget(Label(text="Batching (przerwa 60s/30 maili)")); l.add_widget(bx)
        l.add_widget(ModernButton(text="ZAPISZ KONFIGURACJĘ", on_press=lambda x: [json.dump({'h':self.ti_h.text,'port':self.ti_pt.text,'u':self.ti_u.text,'p':self.ti_p.text,'batch':self.cb_b.active}, open(p,"w")), self.msg("OK","Zapisano")]))
        l.add_widget(ModernButton(text="TEST POŁĄCZENIA", on_press=lambda x: self.test_smtp_direct(), bg_color=(.1,.7,.4,1)))
        l.add_widget(ModernButton(text="POKAŻ LOGI", on_press=self.show_logs))
        l.add_widget(ModernButton(text="POWRÓT", on_press=lambda x: setattr(self.sm,'current','home'), bg_color=(.3,.3,.3,1))); self.sc_ref["smtp"].add_widget(l)

    def test_smtp_direct(self):
        try: s = self.connect_smtp({'h':self.ti_h.text,'port':self.ti_pt.text,'u':self.ti_u.text,'p':self.ti_p.text}); s.quit(); self.msg("OK", "Serwer SMTP Działa!"); self.log("SMTP test succeeded")
        except Exception as e: self.log(f"test_smtp_direct error: {traceback.format_exc()}"); self.msg("BŁĄD", str(e)[:60])

    def start_special_send_flow(self, _): self.open_picker("special_send")

    def special_send_step_2(self, path):
        self.selected_emails = []; box = BoxLayout(orientation="vertical", padding=dp(15), spacing=dp(10)); ti = ModernInput(hint_text="Szukaj..."); box.add_widget(ti)
        sc = ScrollView(); gl = GridLayout(cols=1, size_hint_y=None, spacing=dp(5)); gl.bind(minimum_height=gl.setter('height')); sc.add_widget(gl); box.add_widget(sc)
        def rf(v=""):
            gl.clear_widgets(); rows = self.conn.execute("SELECT name, surname, email FROM contacts").fetchall()
            for r in rows:
                if v and v.lower() not in f"{r[0]} {r[1]} {r[2]}".lower(): continue
                bx = BoxLayout(size_hint_y=None, height=dp(50)); cb = CheckBox(size_hint_x=None, width=dp(50))
                cb.bind(active=lambda inst, val, m=r[2]: self.selected_emails.append(m) if val else self.selected_emails.remove(m))
                bx.add_widget(cb); bx.add_widget(Label(text=f"{r[0].title()} {r[1].title()}")); gl.add_widget(bx)
        ti.bind(text=lambda i,v: rf(v)); rf(); btn = ModernButton(text="DALEJ", on_press=lambda x: [p.dismiss(), self.special_send_step_3(path)] if self.selected_emails else None); box.add_widget(btn); p = Popup(title="Odbiorcy", content=box, size_hint=(.95,.9)); p.open()

    def special_send_step_3(self, path):
        b = BoxLayout(orientation="vertical", padding=dp(15), spacing=dp(10)); ti_s = ModernInput(hint_text="Temat"); ti_b = ModernInput(hint_text="Treść", multiline=True); b.add_widget(ti_s); b.add_widget(ti_b)
        def run(_):
            def task():
                cfg = json.load(open(Path(self.user_data_dir)/"smtp.json")); srv = self.connect_smtp(cfg)
                for m in self.selected_emails:
                    msg = EmailMessage(); msg["Subject"], msg["From"], msg["To"] = ti_s.text, cfg['u'], m; msg.set_content(ti_b.text)
                    with open(path, "rb") as f: msg.add_attachment(f.read(), maintype="application", subtype="octet-stream", filename=os.path.basename(path))
                    srv.send_message(msg)
                srv.quit(); Clock.schedule_once(lambda d: self.msg("OK", "Wysłano")); self.log(f"Special send file {path} to {len(self.selected_emails)} recipients")
            threading.Thread(target=task, daemon=True).start(); p.dismiss()
        b.add_widget(ModernButton(text="WYŚLIJ PLIK", on_press=run)); p = Popup(title="Wiadomość", content=b, size_hint=(.9, .8)); p.open()

    def filter_table(self, i, v): self.filtered_data = [self.full_data[0]] + [r for r in self.full_data[1:] if any(v.lower() in str(c).lower() for c in r)]; self.refresh_table()

    def start_mass_mailing(self, _):
        if self.is_mailing_running: return
        self.stats, self.session_details, self.queue = {"ok": 0, "fail": 0, "skip": 0}, [], list(self.full_data[1:])
        self.total_q = len(self.queue); self.is_mailing_running = True; self.mailing_paused = False; threading.Thread(target=self.mailing_worker, daemon=True).start()
        self.log(f"Mass mailing started: {self.total_q} items")

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
                            while (n := stream.read(buf)) > 0: f.write(buf[:n])
                        stream.close()
                        if mode == "data": self.process_excel(loc)
                        elif mode == "book": self.process_book(loc)
                        elif mode == "attachment":
                            self.global_attachments.append(str(loc))
                        elif mode == "special_send": Clock.schedule_once(lambda dt: self.special_send_step_2(str(loc)))
                    except:
                        pass
                self.update_stats()
        activity.bind(on_activity_result=cb); PA.mActivity.startActivityForResult(intent, 1001)

    def setup_tmpl_ui(self):
        self.sc_ref["tmpl"].clear_widgets()
        l, ti_s, ti_b = BoxLayout(orientation="vertical", padding=dp(25), spacing=dp(10)), ModernInput(hint_text="Temat {Imię}"), ModernInput(hint_text="Treść...", multiline=True)
        ts, tb = self.conn.execute("SELECT val FROM settings WHERE key='t_sub'").fetchone(), self.conn.execute("SELECT val FROM settings WHERE key='t_body'").fetchone()
        ti_s.text, ti_b.text = (ts[0] if ts else ""), (tb[0] if tb else "")
        l.add_widget(Label(text="SZABLON EMAIL", bold=True)); l.add_widget(ti_s); l.add_widget(ti_b)
        l.add_widget(ModernButton(text="ZAPISZ", on_press=lambda x: [self.conn.execute("INSERT OR REPLACE INTO settings VALUES (?,?)", ('t_sub',ti_s.text)), self.conn.execute("INSERT OR REPLACE INTO settings VALUES (?,?)", ('t_body',ti_b.text)), self.conn.commit(), self.msg("OK","Wzór zapisany")]))
        l.add_widget(ModernButton(text="POWRÓT", on_press=lambda x: setattr(self.sm, 'current', 'email'))); self.sc_ref["tmpl"].add_widget(l)

    def setup_contacts_ui(self):
        self.sc_ref["contacts"].clear_widgets()
        l, top = BoxLayout(orientation="vertical", padding=dp(10)), BoxLayout(size_hint_y=None, height=dp(55), spacing=dp(5))
        self.ti_cs = TextInput(hint_text="Szukaj..."); self.ti_cs.bind(text=self.refresh_contacts_list); top.add_widget(self.ti_cs)
        top.add_widget(Button(text="+", size_hint_x=0.15, on_press=lambda x: self.form_contact())); top.add_widget(Button(text="Wróć", size_hint_x=0.2, on_press=lambda x: setattr(self.sm, 'current', 'home')))
        self.c_ls = GridLayout(cols=1, size_hint_y=None, spacing=dp(10)); self.c_ls.bind(minimum_height=self.c_ls.setter('height'))
        sc = ScrollView(); sc.add_widget(self.c_ls); l.add_widget(top); l.add_widget(sc); self.sc_ref["contacts"].add_widget(l)

    def refresh_contacts_list(self, *args):
        self.c_ls.clear_widgets(); sv = self.ti_cs.text.lower(); rows = self.conn.execute("SELECT name, surname, email, pesel, phone FROM contacts ORDER BY surname ASC").fetchall()
        for d in rows:
            if sv and sv not in f"{d[0]} {d[1]} {d[2]}".lower(): continue
            r = BoxLayout(size_hint_y=None, height=dp(125), padding=dp(10))
            with r.canvas.before: Color(*COLOR_CARD); Rectangle(pos=r.pos, size=r.size)
            inf, acts = BoxLayout(orientation="vertical"), BoxLayout(size_hint_x=0.3, orientation="vertical", spacing=dp(4))
            inf.add_widget(Label(text=f"{d[0]} {d[1]}".title(), bold=True, halign="left", text_size=(dp(250),None)))
            inf.add_widget(Label(text=f"E: {d[2]}\nP: {d[3]}\nT: {d[4] if d[4] else '-'}", font_size='11sp', halign="left", text_size=(dp(250),None), color=(0.7,0.7,0.7,1)))
            r.add_widget(inf); acts.add_widget(Button(text="Edytuj", on_press=lambda x, data=d: self.form_contact(*data))); acts.add_widget(Button(text="Usuń", background_color=(0.8,0.2,0.2,1), on_press=lambda x, n=d[0], s=d[1]: self.delete_contact(n, s))); r.add_widget(acts); self.c_ls.add_widget(r)

    def msg(self, tit, txt):
        b = BoxLayout(orientation="vertical", padding=dp(20)); b.add_widget(Label(text=txt, halign="center")); b.add_widget(ModernButton(text="OK", on_press=lambda x: p.dismiss(), height=dp(50), size_hint_y=None)); p = Popup(title=tit, content=b, size_hint=(0.85, 0.45)); p.open()

    def update_stats(self, *a):
        try:
            count = self.conn.execute('SELECT count(*) FROM contacts').fetchone()[0]
            s = f"Baza: {count} | Załączniki: {len(self.global_attachments)}"
            if hasattr(self, 'lbl_stats'):
                self.lbl_stats.text = s
            if hasattr(self, 'lbl_stats_paski'):
                self.lbl_stats_paski.text = s
        except:
            pass

    def update_progress(self, d):
        try:
            val = int((d/self.total_q)*100) if self.total_q else 0
            if hasattr(self, 'pb'):
                self.pb.value = val
            if hasattr(self, 'pb_paski'):
                self.pb_paski.value = val
            if hasattr(self, 'pb_label'):
                self.pb_label.text = f"Postęp: {d}/{self.total_q}"
            if hasattr(self, 'pb_label_paski'):
                self.pb_label_paski.text = f"Postęp: {d}/{self.total_q}"
        except:
            pass

    def finish_mailing(self, s):
        self.is_mailing_running = False; det = "\n".join(self.session_details); self.conn.execute("INSERT INTO reports (date, ok, fail, skip, auto, details) VALUES (?,?,?,?,?,?)", (datetime.now().strftime("%Y-%m-%d %H:%M"), self.stats['ok'], self.stats['fail'], self.stats['skip'], 0, det)); self.conn.commit()
        Clock.schedule_once(lambda dt: self.msg("Mailing", f"{s}\nSukces: {self.stats['ok']}"))
        self.log(f"Mailing finished: {s} | ok={self.stats['ok']} fail={self.stats['fail']} skip={self.stats['skip']}")

    def popup_columns(self, _):
        box, gr, checks = BoxLayout(orientation="vertical", padding=dp(10)), GridLayout(cols=1, size_hint_y=None, spacing=dp(5)), []
        gr.bind(minimum_height=gr.setter('height'))
        for i, h in enumerate(self.full_data[0]):
            r, cb = BoxLayout(size_hint_y=None, height=dp(45)), CheckBox(active=(i in self.export_indices), size_hint_x=None, width=dp(50)); checks.append((i, cb)); r.add_widget(cb); r.add_widget(Label(text=str(h))); gr.add_widget(r)
        sc = ScrollView(); sc.add_widget(gr); box.add_widget(sc); box.add_widget(ModernButton(text="ZASTOSUJ", on_press=lambda x: [setattr(self, 'export_indices', [idx for idx, c in checks if c.active]), p.dismiss(), self.refresh_table()], height=dp(50), size_hint_y=None)); p = Popup(title="Kolumny", content=box, size_hint=(0.9, 0.9)); p.open()

    def setup_report_ui(self):
        self.sc_ref["report"].clear_widgets()
        l, self.r_grid = BoxLayout(orientation="vertical", padding=dp(15), spacing=dp(10)), GridLayout(cols=1, size_hint_y=None, spacing=dp(10))
        self.r_grid.bind(minimum_height=self.r_grid.setter('height')); sc = ScrollView(); sc.add_widget(self.r_grid); l.add_widget(Label(text="HISTORIA SESJI", bold=True, height=dp(40), size_hint_y=None)); l.add_widget(sc); l.add_widget(ModernButton(text="POWRÓT", on_press=lambda x: setattr(self.sm, 'current', 'home'), height=dp(55), size_hint_y=None)); self.sc_ref["report"].add_widget(l)

    def refresh_reports(self, *a):
        self.r_grid.clear_widgets(); rows = self.conn.execute("SELECT date, ok, fail, skip, details FROM reports ORDER BY id DESC").fetchall()
        for d, ok, fl, sk, det in rows:
            row = BoxLayout(orientation="vertical", size_hint_y=None, height=dp(110), padding=dp(10))
            with row.canvas.before: Color(0.15, 0.2, 0.25, 1); Rectangle(pos=row.pos, size=row.size)
            row.add_widget(Label(text=f"Sesja: {d}", bold=True, color=COLOR_PRIMARY)); row.add_widget(Button(text="Pokaż logi", size_hint_y=None, height=dp(35), on_press=lambda x, t=det: self.show_details(t))); self.r_grid.add_widget(row)

    def show_details(self, t):
        b = BoxLayout(orientation="vertical", padding=dp(10)); ti = TextInput(text=str(t), readonly=True, font_size='11sp'); b.add_widget(ti); b.add_widget(Button(text="ZAMKNIJ", size_hint_y=0.2, on_press=lambda x: p.dismiss())); p = Popup(title="Logi", content=b, size_hint=(.9,.8)); p.open()

    def ask_before_send_worker(self, row, email, n, s):
        def dec(v): self.user_decision = "send" if v else "skip"; self.wait_for_user = False; px.dismiss()
        box = BoxLayout(orientation="vertical", padding=dp(20), spacing=dp(10)); box.add_widget(Label(text=f"POTWIERDŹ:\n[b]{n} {s}[/b]\n{email}", markup=True, halign="center"))
        btns = BoxLayout(size_hint_y=None, height=dp(55), spacing=dp(10)); btns.add_widget(Button(text="WYŚLIJ", on_press=lambda x: dec(True), background_color=(0,0.6,0,1))); btns.add_widget(Button(text="POMIŃ", on_press=lambda x: dec(False), background_color=(0.7,0,0,1)))
        box.add_widget(btns); px = Popup(title="Weryfikacja", content=box, size_hint=(0.9, 0.45), auto_dismiss=False); px.open()

    def export_single_row(self, r):
        p = Path("/storage/emulated/0/Documents/FutureExport") if platform=="android" else Path("./exports"); p.mkdir(parents=True, exist_ok=True)
        nx, sx = str(r[self.idx_name]).title(), str(r[self.idx_surname]).title(); wb = Workbook(); ws = wb.active
        ws.append([self.full_data[0][k] for k in self.export_indices]); ws.append([str(r[k]) if (k < len(r) and str(r[k]).strip() != "") else "0" for k in self.export_indices])
        self.style_xlsx(ws); wb.save(p/f"Raport_{nx}_{sx}.xlsx"); self.msg("OK", f"Zapisano PDF dla: {nx}"); self.log(f"Export single row for {nx} {sx}")

    def delete_contact(self, n, s):
        def pr(_): [self.conn.execute("DELETE FROM contacts WHERE name=? AND surname=?", (n, s)), self.conn.commit(), px.dismiss(), self.refresh_contacts_list(), self.update_stats()]
        px = Popup(title="Usuń?", content=Button(text="USUŃ KONTAKT", on_press=pr, background_color=(1,0,0,1)), size_hint=(0.7,0.3)); px.open()

    def form_contact(self, n="", s="", e="", pes="", ph=""):
        b, f_ins = BoxLayout(orientation="vertical", padding=dp(15), spacing=dp(10)), [TextInput(text=str(n), hint_text="Imię"), TextInput(text=str(s), hint_text="Nazwisko"), TextInput(text=str(e), hint_text="Email"), TextInput(text=str(pes), hint_text="PESEL"), TextInput(text=str(ph), hint_text="Telefon")]
        for f in f_ins: b.add_widget(f)
        def save(_): [self.conn.execute("INSERT OR REPLACE INTO contacts VALUES (?,?,?,?,?)", (f_ins[0].text.lower(), f_ins[1].text.lower(), f_ins[2].text.strip(), f_ins[3].text.strip(), f_ins[4].text.strip())), self.conn.commit(), px.dismiss(), self.refresh_contacts_list(), self.update_stats()]
        b.add_widget(ModernButton(text="ZAPISZ", on_press=save)); px = Popup(title="Kontakt", content=b, size_hint=(0.9, 0.85)); px.open()

    def clear_all_attachments(self, _):
        [self.global_attachments.clear(), self.update_stats(), self.log("Cleared attachments")]

    def setup_cars_ui(self):
        self.sc_ref["cars"].clear_widgets()
        b = BoxLayout(orientation="vertical", padding=dp(20), spacing=dp(10))
        b.add_widget(Label(text="Moduł Samochody", bold=True))
        b.add_widget(Label(text="Placeholder - tu będzie rozwijany moduł Samochody"))
        b.add_widget(ModernButton(text="Powrót", on_press=lambda x: setattr(self.sm, 'current', 'home')))
        self.sc_ref["cars"].add_widget(b)

    def setup_paski_ui(self):
        self.sc_ref["paski"].clear_widgets()
        l = BoxLayout(orientation="vertical", padding=dp(15), spacing=dp(10))
        header = BoxLayout(size_hint_y=None, height=dp(40))
        header.add_widget(Label(text="Moduł Paski", bold=True))
        l.add_widget(header)
        ab = BoxLayout(size_hint_y=None, height=dp(45), spacing=dp(10))
        self.cb_paski_auto = CheckBox(size_hint_x=None, width=dp(45))
        self.cb_paski_auto.active = self.auto_send_mode
        self.cb_paski_auto.bind(active=self.on_auto_checkbox_changed)
        ab.add_widget(self.cb_paski_auto); ab.add_widget(Label(text="AUTOMATYCZNA WYSYŁKA", bold=True))
        l.add_widget(ab)
        self.lbl_stats_paski = Label(text="Baza: 0 | Załączniki: 0", height=dp(30)); l.add_widget(self.lbl_stats_paski)
        self.pb_label_paski = Label(text="Gotowy", height=dp(25)); self.pb_paski = ProgressBar(max=100, height=dp(20)); l.add_widget(self.pb_label_paski); l.add_widget(self.pb_paski)
        l.add_widget(ModernButton(text="Wczytaj arkusz płac", on_press=lambda x: self.open_picker("data"), height=dp(50), size_hint_y=None))
        l.add_widget(ModernButton(text="Podgląd i eksport", on_press=lambda x: [self.refresh_table(), setattr(self.sm, 'current', 'table')] if self.full_data else self.msg("!", "Wczytaj arkusz!"), height=dp(50), size_hint_y=None))
        l.add_widget(ModernButton(text="Edytuj szablon", on_press=lambda x: setattr(self.sm, 'current', 'tmpl'), height=dp(50), size_hint_y=None))
        l.add_widget(ModernButton(text="Dołącz załącznik", on_press=lambda x: self.open_picker("attachment"), height=dp(50), size_hint_y=None))
        l.add_widget(ModernButton(text="Wyślij jeden plik", on_press=self.start_special_send_flow, height=dp(50), size_hint_y=None))
        l.add_widget(ModernButton(text="Start masowa wysyłka", on_press=self.start_mass_mailing, height=dp(50), size_hint_y=None))
        l.add_widget(ModernButton(text="PAUZA/RESUME WYSYŁKI", on_press=self.toggle_pause_mailing, height=dp(50), size_hint_y=None, bg_color=(0.6,0.6,0.1,1)))
        l.add_widget(ModernButton(text="Raporty sesji", on_press=lambda x: [self.refresh_reports(), setattr(self.sm, 'current', 'report')], height=dp(50), size_hint_y=None))
        l.add_widget(ModernButton(text="Wyczyść załączniki", on_press=self.clear_all_attachments, height=dp(50), size_hint_y=None))
        l.add_widget(ModernButton(text="Powrót", on_press=lambda x: setattr(self.sm, 'current', 'home'), height=dp(55), size_hint_y=None, bg_color=(0.3,0.3,0.3,1)))
        self.sc_ref["paski"].add_widget(l)
        self.update_stats()

    def setup_pracownicy_ui(self):
        self.sc_ref["pracownicy"].clear_widgets()
        b = BoxLayout(orientation="vertical", padding=dp(20), spacing=dp(10))
        b.add_widget(Label(text="Moduł Pracownicy", bold=True))
        b.add_widget(Label(text="Placeholder - moduł Pracownicy do późniejszego rozwinięcia"))
        b.add_widget(ModernButton(text="Powrót", on_press=lambda x: setattr(self.sm, 'current', 'home')))
        self.sc_ref["pracownicy"].add_widget(b)

    def setup_zaklady_ui(self):
        self.sc_ref["zaklady"].clear_widgets()
        b = BoxLayout(orientation="vertical", padding=dp(20), spacing=dp(10))
        b.add_widget(Label(text="Moduł Zakłady", bold=True))
        b.add_widget(Label(text="Placeholder - moduł Zakłady do późniejszego rozwinięcia"))
        b.add_widget(ModernButton(text="Powrót", on_press=lambda x: setattr(self.sm, 'current', 'home')))
        self.sc_ref["zaklady"].add_widget(b)

    def setup_settings_ui(self):
        self.sc_ref["settings"].clear_widgets()
        l = BoxLayout(orientation="vertical", padding=dp(15), spacing=dp(10))
        l.add_widget(Label(text="Ustawienia", bold=True))
        l.add_widget(ModernButton(text="Dodaj bazę danych", on_press=lambda x: self.open_picker("book"), height=dp(50), size_hint_y=None))
        l.add_widget(ModernButton(text="Ustawienia SMTP", on_press=lambda x: setattr(self.sm, 'current', 'smtp'), height=dp(50), size_hint_y=None))
        l.add_widget(ModernButton(text="Edytuj szablon email", on_press=lambda x: setattr(self.sm, 'current', 'tmpl'), height=dp(50), size_hint_y=None))
        l.add_widget(ModernButton(text="Wczytaj arkusz płac", on_press=lambda x: self.open_picker("data"), height=dp(50), size_hint_y=None))
        l.add_widget(ModernButton(text="Pokaż logi", on_press=self.show_logs, height=dp(50), size_hint_y=None))
        l.add_widget(ModernButton(text="Powrót", on_press=lambda x: setattr(self.sm, 'current', 'home'), height=dp(55), size_hint_y=None, bg_color=(0.3,0.3,0.3,1)))
        self.sc_ref["settings"].add_widget(l)

    def toggle_pause_mailing(self, _=None):
        self.mailing_paused = not self.mailing_paused
        if self.mailing_paused: self.msg("OK", "Wysyłka zatrzymana"); self.log("Mailing paused")
        else: self.msg("OK", "Wysyłka wznowiona"); self.log("Mailing resumed")

    def show_logs(self, _=None):
        try:
            text = ""
            if self.log_file.exists():
                with open(self.log_file, "r", encoding="utf-8") as f:
                    text = f.read()[-40000:]
            else:
                text = "\n".join(self._log_buffer)
            b = BoxLayout(orientation="vertical", padding=dp(10))
            ti = TextInput(text=text, readonly=True, font_size='11sp')
            b.add_widget(ti)
            b.add_widget(Button(text="ZAMKNIJ", size_hint_y=0.2, on_press=lambda x: p.dismiss()))
            p = Popup(title="Logi aplikacji", content=b, size_hint=(.95,.95)); p.open()
        except Exception:
            self.log(f"show_logs error: {traceback.format_exc()}")
            self.msg("Błąd", "Nie można otworzyć logów")

if __name__ == "__main__":
    FutureApp().run()
