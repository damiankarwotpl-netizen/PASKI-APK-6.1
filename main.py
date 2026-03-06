import os
import json
import threading
from pathlib import Path

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


APP_TITLE = "Paski Future 7.0 ULTRA STABLE"
CONFIG_FILE = "smtp_config.json"


class HomeScreen(Screen): pass
class TableScreen(Screen): pass
class EmailScreen(Screen): pass
class SMTPScreen(Screen): pass


class PremiumButton(Button):

    def __init__(self, **kwargs):
        super().__init__(**kwargs)

        self.background_normal = ""
        self.background_color = (0.2,0.4,0.8,1)
        self.size_hint_y=None
        self.height=dp(55)
        self.font_size=16


class PaskiFutureApp(App):

    def build(self):

        Window.clearcolor=(0.08,0.1,0.15,1)

        self.full_data=[]
        self.filtered_data=[]
        self.current_file=None

        self.export_columns=None
        self.email_column=0

        self.sm=ScreenManager()

        self.home=HomeScreen(name="home")
        self.table=TableScreen(name="table")
        self.email=EmailScreen(name="email")
        self.smtp=SMTPScreen(name="smtp")

        self.build_home()
        self.build_table()
        self.build_email()
        self.build_smtp()

        self.sm.add_widget(self.home)
        self.sm.add_widget(self.table)
        self.sm.add_widget(self.email)
        self.sm.add_widget(self.smtp)

        return self.sm


# HOME
    def build_home(self):

        layout=BoxLayout(orientation="vertical",padding=dp(30),spacing=dp(20))

        title=Label(text=APP_TITLE,font_size=26)

        open_btn=PremiumButton(text="📂 Otwórz Excel")
        open_btn.bind(on_press=self.open_excel_picker)

        load_btn=PremiumButton(text="📊 Wczytaj dane")
        load_btn.bind(on_press=self.load_excel)

        smtp_btn=PremiumButton(text="⚙ SMTP")
        smtp_btn.bind(on_press=lambda x:setattr(self.sm,"current","smtp"))

        layout.add_widget(title)
        layout.add_widget(open_btn)
        layout.add_widget(load_btn)
        layout.add_widget(smtp_btn)

        self.home.add_widget(layout)


# PICKER
    def open_excel_picker(self,_):

        if platform!="android":
            return

        from jnius import autoclass
        from android import activity

        PythonActivity=autoclass("org.kivy.android.PythonActivity")
        Intent=autoclass("android.content.Intent")

        intent=Intent(Intent.ACTION_OPEN_DOCUMENT)
        intent.setType("*/*")

        activity.bind(on_activity_result=self.on_file_selected)

        PythonActivity.mActivity.startActivityForResult(intent,999)


    def on_file_selected(self,request,result,intent):

        if request!=999:
            return

        from android import activity
        activity.unbind(on_activity_result=self.on_file_selected)

        from jnius import autoclass

        PythonActivity=autoclass("org.kivy.android.PythonActivity")

        resolver=PythonActivity.mActivity.getContentResolver()

        uri=intent.getData()

        stream=resolver.openInputStream(uri)

        path=Path(self.user_data_dir)/"excel.xlsx"

        with open(path,"wb") as f:
            f.write(stream.read())

        self.current_file=path


# LOAD EXCEL
    def load_excel(self,_):

        if not self.current_file:
            return

        from openpyxl import load_workbook

        wb=load_workbook(self.current_file,data_only=True)

        ws=wb.active

        data=[]

        for row in ws.iter_rows(values_only=True):

            data.append(["" if v is None else str(v) for v in row])

        wb.close()

        if not data:
            return

        self.full_data=data
        self.filtered_data=data

        self.display_table()

        self.sm.current="table"


# TABLE
    def build_table(self):

        layout=BoxLayout(orientation="vertical")

        top=BoxLayout(size_hint=(1,0.12))

        self.search=TextInput(hint_text="Szukaj...",multiline=False)
        self.search.bind(text=self.filter_data)

        export_btn=PremiumButton(text="📦 Export")
        export_btn.bind(on_press=self.export_files)

        choose_cols=PremiumButton(text="☑ Kolumny")
        choose_cols.bind(on_press=self.choose_export_columns)

        choose_email=PremiumButton(text="📧 Kolumna email")
        choose_email.bind(on_press=self.choose_email_column)

        email_btn=PremiumButton(text="📬 Email")
        email_btn.bind(on_press=lambda x:setattr(self.sm,"current","email"))

        back_btn=PremiumButton(text="⬅")
        back_btn.bind(on_press=lambda x:setattr(self.sm,"current","home"))

        top.add_widget(self.search)
        top.add_widget(export_btn)
        top.add_widget(choose_cols)
        top.add_widget(choose_email)
        top.add_widget(email_btn)
        top.add_widget(back_btn)

        self.scroll=ScrollView()

        self.grid=GridLayout(size_hint=(None,None),spacing=1)
        self.grid.bind(minimum_height=self.grid.setter("height"))
        self.grid.bind(minimum_width=self.grid.setter("width"))

        self.scroll.add_widget(self.grid)

        self.progress=ProgressBar(max=100,size_hint=(1,0.05))

        layout.add_widget(top)
        layout.add_widget(self.scroll)
        layout.add_widget(self.progress)

        self.table.add_widget(layout)


# DISPLAY TABLE
    def display_table(self):

        self.grid.clear_widgets()

        if not self.filtered_data:
            return

        cols=len(self.filtered_data[0])

        self.grid.cols=cols

        for row in self.filtered_data:

            for cell in row:

                lbl=Label(
                    text=str(cell),
                    size_hint=(None,None),
                    size=(dp(160),dp(40))
                )

                self.grid.add_widget(lbl)


# FILTER
    def filter_data(self,instance,value):

        value=value.lower()

        header=self.full_data[0]

        rows=self.full_data[1:]

        filtered=[]

        for row in rows:

            if any(value in str(cell).lower() for cell in row):

                filtered.append(row)

        self.filtered_data=[header]+filtered

        self.display_table()


# CHOOSE EXPORT COLUMNS
    def choose_export_columns(self,_):

        layout=BoxLayout(orientation="vertical")

        self.export_checks=[]

        for i,col in enumerate(self.full_data[0]):

            row=BoxLayout(size_hint_y=None,height=dp(40))

            cb=CheckBox(active=True)

            lbl=Label(text=str(col))

            self.export_checks.append((i,cb))

            row.add_widget(cb)
            row.add_widget(lbl)

            layout.add_widget(row)

        btn=Button(text="OK",size_hint_y=None,height=dp(50))

        layout.add_widget(btn)

        popup=Popup(title="Kolumny exportu",content=layout,size_hint=(0.9,0.9))

        def save_cols(instance):

            self.export_columns=[i for i,cb in self.export_checks if cb.active]

            popup.dismiss()

        btn.bind(on_press=save_cols)

        popup.open()


# EMAIL COLUMN
    def choose_email_column(self,_):

        layout=BoxLayout(orientation="vertical")

        self.email_checks=[]

        for i,col in enumerate(self.full_data[0]):

            row=BoxLayout(size_hint_y=None,height=dp(40))

            cb=CheckBox(group="email")

            lbl=Label(text=str(col))

            self.email_checks.append((i,cb))

            row.add_widget(cb)
            row.add_widget(lbl)

            layout.add_widget(row)

        btn=Button(text="OK",size_hint_y=None,height=dp(50))

        layout.add_widget(btn)

        popup=Popup(title="Kolumna email",content=layout,size_hint=(0.9,0.9))

        def save_email(instance):

            for i,cb in self.email_checks:
                if cb.active:
                    self.email_column=i

            popup.dismiss()

        btn.bind(on_press=save_email)

        popup.open()


# EXPORT
    def export_files(self,_):

        thread=threading.Thread(target=self.export_thread)

        thread.start()


    def export_thread(self):

        from openpyxl import Workbook
        from datetime import datetime

        path="/storage/emulated/0/Documents/PaskiFuture"

        os.makedirs(path,exist_ok=True)

        header=self.full_data[0]

        rows=self.filtered_data[1:]

        total=len(rows)

        done=0

        for row in rows:

            wb=Workbook()

            ws=wb.active

            if self.export_columns:

                ws.append([header[i] for i in self.export_columns])

                ws.append([row[i] for i in self.export_columns])

            else:

                ws.append(header)
                ws.append(row)

            for col in ws.columns:

                max_len=0

                for cell in col:

                    if cell.value:

                        max_len=max(max_len,len(str(cell.value)))

                ws.column_dimensions[col[0].column_letter].width=max_len+4

            name=row[1] if len(row)>1 else "dane"

            now=datetime.now().strftime("%Y%m%d_%H%M%S")

            wb.save(f"{path}/{name}_{now}.xlsx")

            done+=1

            percent=int(done/total*100)

            Clock.schedule_once(lambda dt,p=percent:setattr(self.progress,"value",p))


# EMAIL SCREEN
    def build_email(self):

        layout=BoxLayout(orientation="vertical",padding=dp(30),spacing=dp(20))

        send1=PremiumButton(text="📧 Wyślij jeden")
        sendAll=PremiumButton(text="📨 Wyślij wszystkie")

        back=PremiumButton(text="⬅")

        send1.bind(on_press=self.send_single)
        sendAll.bind(on_press=self.send_all)

        back.bind(on_press=lambda x:setattr(self.sm,"current","table"))

        layout.add_widget(send1)
        layout.add_widget(sendAll)
        layout.add_widget(back)

        self.email.add_widget(layout)


# SEND EMAIL
    def send_single(self,_):

        if len(self.filtered_data)<2:
            return

        threading.Thread(target=self.send_email,args=(self.filtered_data[1],)).start()


    def send_all(self,_):

        for row in self.filtered_data[1:]:

            threading.Thread(target=self.send_email,args=(row,)).start()


    def send_email(self,row):

        import smtplib
        from email.message import EmailMessage
        from openpyxl import Workbook

        if not os.path.exists(CONFIG_FILE):
            return

        with open(CONFIG_FILE) as f:
            config=json.load(f)

        if self.email_column>=len(row):
            return

        email=row[self.email_column]

        if not email or "@" not in email:
            return

        wb=Workbook()

        ws=wb.active

        ws.append(self.full_data[0])
        ws.append(row)

        temp=Path(self.user_data_dir)/"temp.xlsx"

        wb.save(temp)

        msg=EmailMessage()

        msg["Subject"]="Dane"
        msg["From"]=config["email"]
        msg["To"]=email

        msg.set_content("W załączniku plik.")

        with open(temp,"rb") as f:

            msg.add_attachment(
                f.read(),
                maintype="application",
                subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                filename="dane.xlsx"
            )

        server=smtplib.SMTP(config["server"],int(config["port"]))

        server.starttls()

        server.login(config["email"],config["password"])

        server.send_message(msg)

        server.quit()


# SMTP
    def build_smtp(self):

        layout=BoxLayout(orientation="vertical",padding=dp(30),spacing=dp(15))

        self.s_server=TextInput(hint_text="SMTP server")
        self.s_port=TextInput(hint_text="port")
        self.s_email=TextInput(hint_text="email")
        self.s_pass=TextInput(hint_text="hasło",password=True)

        save=PremiumButton(text="Zapisz")

        back=PremiumButton(text="⬅")

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
