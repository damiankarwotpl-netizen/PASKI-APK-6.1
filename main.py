import os
import json
import smtplib
import threading
from pathlib import Path
from datetime import datetime
from email.message import EmailMessage

from kivy.app import App
from kivy.clock import Clock
from kivy.core.window import Window
from kivy.metrics import dp
from kivy.utils import platform

from kivy.uix.boxlayout import BoxLayout
from kivy.uix.button import Button
from kivy.uix.label import Label
from kivy.uix.scrollview import ScrollView
from kivy.uix.gridlayout import GridLayout
from kivy.uix.textinput import TextInput
from kivy.uix.popup import Popup
from kivy.uix.progressbar import ProgressBar
from kivy.uix.screenmanager import ScreenManager, Screen
from kivy.uix.checkbox import CheckBox


APP_TITLE = "Paski Future 6.3"
CONFIG_FILE = "smtp_config.json"
EMAIL_COLUMN_INDEX = 3


class HomeScreen(Screen): pass
class TableScreen(Screen): pass
class EmailScreen(Screen): pass
class SMTPScreen(Screen): pass


class PremiumButton(Button):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.background_normal = ""
        self.background_color = (0.2,0.4,0.8,1)
        self.color = (1,1,1,1)
        self.size_hint_y = None
        self.height = dp(48)


class PaskiFutureApp(App):

    def build(self):

        Window.clearcolor=(0.08,0.1,0.15,1)

        self.full_data=[]
        self.filtered_data=[]
        self.current_file=None

        self.export_columns=None
        self.email_column=EMAIL_COLUMN_INDEX

        self.sm=ScreenManager()

        self.home=HomeScreen(name="home")
        self.table=TableScreen(name="table")
        self.email=EmailScreen(name="email")
        self.smtp=SMTPScreen(name="smtp")

        self._build_home()
        self._build_table()
        self._build_email()
        self._build_smtp()

        self.sm.add_widget(self.home)
        self.sm.add_widget(self.table)
        self.sm.add_widget(self.email)
        self.sm.add_widget(self.smtp)

        return self.sm


# =================================================
# HOME
# =================================================

    def _build_home(self):

        layout=BoxLayout(orientation="vertical",padding=dp(30),spacing=dp(20))

        title=Label(text=APP_TITLE,font_size=26)

        open_btn=PremiumButton(text="📂 Otwórz Excel")
        open_btn.bind(on_press=self.open_excel_picker)

        load_btn=PremiumButton(text="📊 Wczytaj dane")
        load_btn.bind(on_press=self.load_full_excel)

        smtp_btn=PremiumButton(text="⚙ SMTP")
        smtp_btn.bind(on_press=lambda x:setattr(self.sm,"current","smtp"))

        self.home_status=Label(text="Gotowy")

        layout.add_widget(title)
        layout.add_widget(open_btn)
        layout.add_widget(load_btn)
        layout.add_widget(smtp_btn)
        layout.add_widget(self.home_status)

        self.home.add_widget(layout)


# =================================================
# PICKER
# =================================================

    def open_excel_picker(self,_):

        if platform!="android":
            self.home_status.text="Picker tylko Android"
            return

        from jnius import autoclass
        from android import activity

        PythonActivity=autoclass("org.kivy.android.PythonActivity")
        Intent=autoclass("android.content.Intent")

        intent=Intent(Intent.ACTION_OPEN_DOCUMENT)
        intent.setType("*/*")
        intent.addCategory(Intent.CATEGORY_OPENABLE)

        activity.bind(on_activity_result=self._on_activity_result)
        PythonActivity.mActivity.startActivityForResult(intent,999)


    def _on_activity_result(self,request_code,result_code,intent):

        if request_code!=999 or not intent:
            return

        from android import activity
        activity.unbind(on_activity_result=self._on_activity_result)

        from jnius import autoclass
        PythonActivity=autoclass("org.kivy.android.PythonActivity")

        resolver=PythonActivity.mActivity.getContentResolver()
        uri=intent.getData()
        stream=resolver.openInputStream(uri)

        local_file=Path(self.user_data_dir)/"selected.xlsx"

        with open(local_file,"wb") as out:

            buffer=bytearray(4096)

            while True:

                read=stream.read(buffer)

                if read==-1:
                    break

                out.write(buffer[:read])

        stream.close()

        self.current_file=local_file
        self.home_status.text="Plik wybrany"


# =================================================
# LOAD EXCEL
# =================================================

    def load_full_excel(self,_):

        if not self.current_file:
            return

        from openpyxl import load_workbook

        wb=load_workbook(str(self.current_file),data_only=True)
        sheet=wb.active

        self.full_data=[
            ["" if v is None else str(v) for v in row]
            for row in sheet.iter_rows(values_only=True)
        ]

        wb.close()

        self.filtered_data=self.full_data

        self.display_table()
        self.sm.current="table"


# =================================================
# TABLE
# =================================================

    def _build_table(self):

        layout=BoxLayout(orientation="vertical",padding=dp(10),spacing=dp(10))

        top=BoxLayout(size_hint=(1,0.12),spacing=dp(10))

        self.search=TextInput(hint_text="🔎 Szukaj",multiline=False)
        self.search.bind(text=self.filter_data)

        col_btn=PremiumButton(text="📑 Kolumny")
        col_btn.bind(on_press=lambda x:self.select_columns_popup())

        mail_btn=PremiumButton(text="📧 Kolumna mail")
        mail_btn.bind(on_press=lambda x:self.select_email_column_popup())

        export_btn=PremiumButton(text="📦 Eksport")
        export_btn.bind(on_press=self.export_files)

        email_btn=PremiumButton(text="📬 Email")
        email_btn.bind(on_press=lambda x:setattr(self.sm,"current","email"))

        back_btn=PremiumButton(text="⬅ Powrót")
        back_btn.bind(on_press=lambda x:setattr(self.sm,"current","home"))

        top.add_widget(self.search)
        top.add_widget(col_btn)
        top.add_widget(mail_btn)
        top.add_widget(export_btn)
        top.add_widget(email_btn)
        top.add_widget(back_btn)

        self.scroll=ScrollView()

        self.grid=GridLayout(size_hint=(None,None))
        self.grid.bind(minimum_height=self.grid.setter("height"))
        self.grid.bind(minimum_width=self.grid.setter("width"))

        self.scroll.add_widget(self.grid)

        self.progress=ProgressBar(max=100)

        layout.add_widget(top)
        layout.add_widget(self.scroll)
        layout.add_widget(self.progress)

        self.table.add_widget(layout)


# =================================================
# TABLE DISPLAY
# =================================================

    def display_table(self):

        self.grid.clear_widgets()

        if not self.filtered_data:
            return

        rows=len(self.filtered_data)
        cols=len(self.filtered_data[0])

        cell_w=dp(160)
        cell_h=dp(40)

        self.grid.cols=cols
        self.grid.width=cols*cell_w
        self.grid.height=rows*cell_h

        for row in self.filtered_data:

            for cell in row:

                lbl=Label(
                    text=str(cell),
                    size_hint=(None,None),
                    size=(cell_w,cell_h)
                )

                self.grid.add_widget(lbl)


    def filter_data(self,instance,value):

        value=value.lower()

        self.filtered_data=[
            r for r in self.full_data
            if any(value in str(c).lower() for c in r)
        ]

        self.display_table()


# =================================================
# POPUP COLUMN SELECT
# =================================================

    def select_columns_popup(self):

        header=self.full_data[0]

        layout=BoxLayout(orientation="vertical",spacing=10)

        scroll=ScrollView()

        grid=GridLayout(cols=2,size_hint_y=None)
        grid.bind(minimum_height=grid.setter("height"))

        self.col_checks=[]

        for i,name in enumerate(header):

            lbl=Label(text=name)
            cb=CheckBox()

            grid.add_widget(lbl)
            grid.add_widget(cb)

            self.col_checks.append((i,cb))

        scroll.add_widget(grid)

        btn=PremiumButton(text="OK")

        layout.add_widget(scroll)
        layout.add_widget(btn)

        popup=Popup(title="Kolumny exportu",content=layout,size_hint=(0.9,0.9))

        def apply(_):

            cols=[]

            for i,cb in self.col_checks:
                if cb.active:
                    cols.append(i)

            self.export_columns=cols
            popup.dismiss()

        btn.bind(on_press=apply)

        popup.open()


# =================================================
# EMAIL COLUMN SELECT
# =================================================

    def select_email_column_popup(self):

        header=self.full_data[0]

        layout=BoxLayout(orientation="vertical",spacing=10)

        grid=GridLayout(cols=2,size_hint_y=None)
        grid.bind(minimum_height=grid.setter("height"))

        self.mail_checks=[]

        for i,name in enumerate(header):

            lbl=Label(text=name)
            cb=CheckBox(group="mail")

            grid.add_widget(lbl)
            grid.add_widget(cb)

            self.mail_checks.append((i,cb))

        btn=PremiumButton(text="OK")

        layout.add_widget(grid)
        layout.add_widget(btn)

        popup=Popup(title="Kolumna Email",content=layout,size_hint=(0.9,0.9))

        def apply(_):

            for i,cb in self.mail_checks:
                if cb.active:
                    self.email_column=i

            popup.dismiss()

        btn.bind(on_press=apply)

        popup.open()


# =================================================
# EXPORT
# =================================================

    def export_files(self,_):
        threading.Thread(target=self._export_thread).start()


    def _export_thread(self):

        from openpyxl import Workbook
        from openpyxl.styles import Font,Border,Side
        from openpyxl.utils import get_column_letter

        rows=self.filtered_data[1:]
        header=self.full_data[0]

        if self.export_columns:
            header=[header[i] for i in self.export_columns]

        for row in rows:

            wb=Workbook()
            ws=wb.active

            if self.export_columns:
                r=[row[i] for i in self.export_columns]
            else:
                r=row

            ws.append(header)
            ws.append(r)

            bold=Font(bold=True)
            thin=Side(style="thin")
            border=Border(left=thin,right=thin,top=thin,bottom=thin)

            for cell in ws[1]:
                cell.font=bold

            for rowx in ws.iter_rows():
                for cell in rowx:
                    cell.border=border

            for col in ws.columns:
                max_len=0
                col_letter=get_column_letter(col[0].column)

                for cell in col:
                    if cell.value:
                        max_len=max(max_len,len(str(cell.value)))

                ws.column_dimensions[col_letter].width=max_len+2

            name=row[1] if len(row)>1 else "plik"
            path=f"/storage/emulated/0/Documents/{name}.xlsx"

            wb.save(path)


# =================================================
# EMAIL
# =================================================

    def _build_email(self):

        layout=BoxLayout(orientation="vertical")

        send_btn=PremiumButton(text="Wyślij wszystkie")
        send_btn.bind(on_press=self.send_bulk)

        back=PremiumButton(text="Powrót")
        back.bind(on_press=lambda x:setattr(self.sm,"current","table"))

        layout.add_widget(send_btn)
        layout.add_widget(back)

        self.email.add_widget(layout)


    def send_bulk(self,_):

        threading.Thread(target=self._send_all).start()


    def _send_all(self):

        for row in self.filtered_data[1:]:
            self._send_email_row(row)


    def _send_email_row(self,row):

        from openpyxl import Workbook

        if not os.path.exists(CONFIG_FILE):
            return

        with open(CONFIG_FILE) as f:
            cfg=json.load(f)

        wb=Workbook()
        ws=wb.active
        ws.append(self.full_data[0])
        ws.append(row)

        file=Path(self.user_data_dir)/"mail.xlsx"
        wb.save(file)

        msg=EmailMessage()
        msg["Subject"]="Dane"
        msg["From"]=cfg["email"]
        msg["To"]=row[self.email_column]

        msg.set_content("Załącznik")

        with open(file,"rb") as f:

            msg.add_attachment(
                f.read(),
                maintype="application",
                subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                filename="dane.xlsx"
            )

        s=smtplib.SMTP(cfg["server"],int(cfg["port"]))
        s.starttls()
        s.login(cfg["email"],cfg["password"])
        s.send_message(msg)
        s.quit()


# =================================================
# SMTP
# =================================================

    def _build_smtp(self):

        layout=BoxLayout(orientation="vertical")

        self.s_server=TextInput(hint_text="server")
        self.s_port=TextInput(hint_text="port")
        self.s_email=TextInput(hint_text="email")
        self.s_pass=TextInput(hint_text="hasło",password=True)

        save=PremiumButton(text="Zapisz")
        back=PremiumButton(text="Powrót")

        save.bind(on_press=self.save_smtp)
        back.bind(on_press=lambda x:setattr(self.sm,"current","home"))

        layout.add_widget(self.s_server)
        layout.add_widget(self.s_port)
        layout.add_widget(self.s_email)
        layout.add_widget(self.s_pass)
        layout.add_widget(save)
        layout.add_widget(back)

        self.smtp.add_widget(layout)


    def save_smtp(self,_):

        data={
            "server":self.s_server.text,
            "port":self.s_port.text,
            "email":self.s_email.text,
            "password":self.s_pass.text
        }

        with open(CONFIG_FILE,"w") as f:
            json.dump(data,f)


if __name__=="__main__":
    PaskiFutureApp().run()
