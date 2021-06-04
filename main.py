import datetime
import mimetypes
import os
import smtplib
import sys
import json
from email import encoders
from email.mime.base import MIMEBase
import numpy as np
import pdfkit
import requests
import pandas as pd
from PyQt5 import QtWidgets
from PyQt5.QtGui import QIcon
from config import *
from PyQt5.QtWidgets import QApplication, QMainWindow
from models import database, Data
from design import Ui_MainWindow
from widget import Ui_Form
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText


class Form(QtWidgets.QWidget, Ui_Form):
    def __init__(self, *args, **kwargs):
        QtWidgets.QWidget.__init__(self, *args, **kwargs)
        self.setupUi(self)
        self.connect_ui()

    def connect_ui(self):
        self.pushButton.clicked.connect(self.close)


class MainWindow(QMainWindow, Ui_MainWindow):

    #  Инициализация основного окна
    def __init__(self, *args, **kwargs):
        super(MainWindow, self).__init__(*args, **kwargs)
        self.setupUi(self)
        self.widget = Form()
        self.setWindowIcon(QIcon('logistic.ico'))
        database.connect()
        database.create_tables([Data])
        self.show()
        self.init_ui()

    def init_ui(self):
        self.dateEdit.setDate(datetime.datetime.now())
        self.connect_ui()

    def connect_ui(self):
        self.pushButton_3.clicked.connect(self.load_data)
        self.pushButton.clicked.connect(self.unload_data)
        self.pushButton_2.clicked.connect(self.send_to_email)

    def send_to_email(self):
        check, file = self.unload_data()
        if check:
            email = self.lineEdit_3.text()
            if not email:
                self.widget_act('Введите адрес почты получателя!')
            msg = MIMEMultipart()
            msg['From'] = ADDRESS_EMAIL
            msg['To'] = email
            msg['Subject'] = 'Отчет'
            filename = os.path.basename(file)
            if os.path.isfile(file):
                ctype, encoding = mimetypes.guess_type(file)
                maintype, subtype = ctype.split('/', 1)
                with open(file, 'rb') as fp:
                    file = MIMEBase(maintype, subtype)
                    file.set_payload(fp.read())
                    fp.close()
                encoders.encode_base64(file)
                file.add_header('Content-Disposition', 'attachment', filename=filename)
                msg.attach(file)
            body = f"Выгрузка за {datetime.datetime.now().date()}"
            msg.attach(MIMEText(body, 'plain'))
            server = smtplib.SMTP('smtp.mail.ru', 25)
            server.starttls()
            server.login(ADDRESS_EMAIL, PASSWORD_EMAIL)
            server.send_message(msg)
            server.quit()

    def load_data(self):
        if self.lineEdit_2.text() == '':
            self.widget_act('Заполните поле с вашим адресом!')
            return False
        elif self.lineEdit.text() == '':
            self.widget_act('Заполните поле с адресами!')
            return False
        else:
            if Data.get_or_none(date=datetime.datetime.now()):
                self.widget_act('Сегодня вы уже внесли данные в базу!')
                return
            else:
                final_address = self.lineEdit_2.text() + ';' + self.lineEdit.text() + ';' + self.lineEdit_2.text()
                row = final_address.split(';')
                for i in range(len(row) - 1):
                    url = f"https://maps.googleapis.com/maps/api/distancematrix/json?origins={row[i]}, Москва" \
                          f"&destinations={row[i + 1]}, Москва&key={API_KEY}&language=ru&region=ru"
                    r = requests.get(url)
                    data_list = json.loads(r.text)
                    if data_list['rows'][0]['elements'][0]['status'] == 'NOT_FOUND':
                        self.widget_act('Один из адресов не был найден\nПопробуйте ввести адреса снова!')
                        return
                    else:
                        Data(
                            date=datetime.datetime.now(),
                            from_=data_list['origin_addresses'][0],
                            to_=data_list['destination_addresses'][0],
                            distance=data_list['rows'][0]['elements'][0]['distance']['value']
                        ).save()
                self.widget_act('Данные успешно загружены в базу!')
            self.lineEdit.clear()
            self.lineEdit_2.clear()

    def widget_act(self, text_to_send):
        self.widget.label.setText(text_to_send)
        self.widget.show()

    def pandas_processing(self, date, from_, to_, distance):
        df = pd.DataFrame(
            {'Дата': [item for item in date],
             'Место отправления': [item for item in from_],
             'Место назначения': [item for item in to_],
             'Пройдено км.': [item for item in distance]
             }
        )
        writer = pd.ExcelWriter(f"unload_{datetime.datetime.now().date()}.xlsx",
                                engine='xlsxwriter')
        df.to_excel(writer, sheet_name='Sheet1', index=False)
        writer.save()
        writer.close()
        config = pdfkit.configuration(wkhtmltopdf=f'{os.getcwd()}\\wkhtmltopdf\\bin\\wkhtmltopdf.exe')
        df = pd.read_excel(f"unload_{datetime.datetime.now().date()}.xlsx")
        df1 = df.replace(np.nan, '', regex=True)
        df1.to_html(f"{os.getcwd()}\\file.html")
        pdfkit.from_file(f"{os.getcwd()}\\file.html", f'{os.getcwd()}\\unload_{datetime.datetime.now().date()}.pdf',
                         configuration=config, options={'encoding': "UTF-8"})
        os.remove(f'{os.getcwd()}\\file.html')
        os.remove(f'{os.getcwd()}\\unload_{datetime.datetime.now().date()}.xlsx')
        self.widget_act('Вы успешно выгрузили данные!')
        return f'{os.getcwd()}\\unload_{datetime.datetime.now().date()}.pdf'

    def data_processing(self, price_gas):
        intermediate_date = ''
        date = []
        from_ = []
        to_ = []
        distance = 0
        distance_list = [1]
        index = 0
        final_distance = 0
        month = self.spinBox.value()
        year = self.spinBox_2.value()
        data_from_db_with_month = Data.select().where((month == Data.date.month) & (year == Data.date.year))
        if not data_from_db_with_month:
            self.widget_act('Данных за этот месяц и год\nне существует!')
            return
        else:
            for item in data_from_db_with_month:
                final_distance += item.distance
                if intermediate_date != str(item.date):
                    distance_list[index] = int(round(distance / 1000, 0))
                    distance = 0
                distance += item.distance
                if str(item.date) not in date:
                    intermediate_date = str(item.date)
                    date.append(str(item.date))
                    index = date.index(str(item.date))
                else:
                    date.append('')
                    distance_list.append('')
                from_.append(item.from_)
                to_.append(item.to_)
            else:
                distance_list[index] = int(round(distance / 1000, 0))
                if len(distance_list) < len(date):
                    distance_list.append('')
            final_distance = int(round(final_distance / 1000))
            date += ['', '', '']
            from_ += ['', '', '']
            to_ += ['Итого:(км)', 'Итого к выплате:(руб)', 'Цена бензина за 1 л.:']
            distance_list += [final_distance, round(final_distance / 10 * price_gas, 0), price_gas]
            return self.pandas_processing(date=date, from_=from_, to_=to_, distance=distance_list)

    def unload_data(self):
        price_gas = self.doubleSpinBox.value()
        if price_gas == 0.0:
            self.widget_act('Вы не ввели цену бензина!')
            return False
        else:
            filename = self.data_processing(price_gas=price_gas)
            if not filename:
                return False, None
            else:
                return True, filename


def run():
    try:
        app = QApplication(sys.argv)
        window = MainWindow()
        app.exec_()
    except Exception as exc:
        print(exc)


if __name__ == '__main__':
    run()
