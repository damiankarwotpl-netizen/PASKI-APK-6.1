import os
import threading
from datetime import datetime

from kivy.app import App
from kivy.clock import Clock
from kivy.core.window import Window
from kivy.metrics import dp
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.button import Button
from kivy.uix.label import Label
from kivy.uix.progressbar import ProgressBar
from kivy.uix.popup import Popup
from kivy.uix.screenmanager import ScreenManager, Screen

from plyer import filechooser

from openpyxl import load_workbook, Workbook


APP_TITLE = "Paski Future"


class PremiumButton(Button):
    pass


class HomeScreen(Screen):
    pass


class ResultScreen(Screen):
    pass


class PaskiApp(App):

    def build(self):

        self.title = APP_TITLE
        Window.clearcolor = (0.08, 0.1, 0.15, 1)

        self.full_data = []
        self.filtered_data = []
        self.current_file = None

        self.export_dir = None

        sm = ScreenManager()

        self.home = HomeScreen(name="home")
        self.result = ResultScreen(name="result")

        sm.add_widget(self.home)
        sm.add_widget(self.result)

        self._build_home()
        self._build_result()

        return sm

    def _build_home(self):

        layout = BoxLayout(orientation="vertical", padding=dp(30), spacing=dp(20))

        layout.add_widget(Label(size_hint_y=1))

        load = PremiumButton(text="📂 Wczytaj Excel")
        load.bind(on_press=self.load_file)

        layout.add_widget(load)

        layout.add_widget(Label(size_hint_y=1))

        self.home.add_widget(layout)

    def _build_result(self):

        layout = BoxLayout(orientation="vertical", padding=dp(20), spacing=dp(10))

        self.info = Label(text="")

        export = PremiumButton(text="📦 Eksport")
        export.bind(on_press=self.export_files)

        self.progress = ProgressBar(max=100)

        back = PremiumButton(text="⬅ Powrót")
        back.bind(on_press=lambda x: self.root.current="home")

        layout.add_widget(self.info)
        layout.add_widget(self.progress)
        layout.add_widget(export)
        layout.add_widget(back)

        self.result.add_widget(layout)

    def load_file(self, _):

        filechooser.open_file(on_selection=self._file_selected)

    def _file_selected(self, selection):

        if not selection:
            return

        self.current_file = selection[0]

        threading.Thread(target=self._load_excel).start()

    def _load_excel(self):

        try:

            wb = load_workbook(self.current_file)
            ws = wb.active

            data = []

            for row in ws.iter_rows(values_only=True):
                data.append(list(row))

            self.full_data = data
            self.filtered_data = data

            Clock.schedule_once(lambda dt: self._show_results())

        except Exception as e:

            Clock.schedule_once(lambda dt: self._popup("Błąd", str(e)))

    def _show_results(self):

        self.info.text = f"Wczytano {len(self.full_data)-1} rekordów"

        self.root.current = "result"

    def export_files(self, _):

        if not self.export_dir:

            filechooser.choose_dir(on_selection=self._set_export_dir)

        else:

            threading.Thread(target=self._export_thread).start()

    def _set_export_dir(self, selection):

        if selection:

            self.export_dir = selection[0]

            threading.Thread(target=self._export_thread).start()

    def _export_thread(self):

        if len(self.filtered_data) < 2:
            return

        documents = self.export_dir

        os.makedirs(documents, exist_ok=True)

        header = self.full_data[0]
        rows = self.filtered_data[1:]

        total = len(rows)
        done = 0

        for row in rows:

            wb = Workbook()
            ws = wb.active

            ws.append(header)
            ws.append(row)

            name = row[1] if len(row) > 1 else "brak"

            now = datetime.now().strftime("%Y%m%d_%H%M%S")

            filename = f"{name}_{now}.xlsx"

            path = os.path.join(documents, filename)

            wb.save(path)

            done += 1

            percent = int((done / total) * 100)

            Clock.schedule_once(lambda dt, p=percent: setattr(self.progress, "value", p))

        Clock.schedule_once(lambda dt: self._popup("Sukces", f"Wyeksportowano {done} plików"))

    def _popup(self, title, text):

        popup = Popup(
            title=title,
            content=Label(text=text),
            size_hint=(0.8,0.4)
        )

        popup.open()


if __name__ == "__main__":
    PaskiApp().run()
