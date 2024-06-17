import os
import re
import time
from pysnmp.hlapi import *
from bs4 import BeautifulSoup
import requests
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.alert import Alert
from selenium.webdriver.support.ui import WebDriverWait

from selenium.webdriver.support import expected_conditions as EC

import sys
from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import QApplication, QMainWindow, QTableWidgetItem, QDialog
from PyQt5 import uic, QtGui

# Без привязки к версии
# driver_options = webdriver.ChromeOptions()
# driver_options.add_argument("--headless")
# driver = webdriver.Chrome(options=driver_options)

#
driver_options = Options()
driver_options.add_argument("--headless")
driver_options.add_argument("--disable-blink-features=AutomationControlled")

# Отключение автоматической проверки веб-драйвера
service = Service(executable_path="webdriver/chromedriver.exe")

# Создание экземпляра веб-драйвера Chrome с отключенной проверкой
driver = webdriver.Chrome(service=service, options=driver_options)


class MainWindow(QMainWindow):
    def __init__(self):
        super(MainWindow, self).__init__()
        uic.loadUi(r'ui\scan.ui', self)
        self.setWindowIcon(QtGui.QIcon(r'ui\icons\logo.ico'))
        self.t_scan.setColumnWidth(0, 170)
        self.b_clr_le_ip_prn.clicked.connect(self.le_ip_prn.clear)
        self.b_clear_le_username.clicked.connect(self.le_username.clear)
        self.b_clr_le_ip_arm.clicked.connect(self.le_ip_arm.clear)
        self.b_check_scan.clicked.connect(self.check_scan)
        self.b_save.clicked.connect(self.new_scan)
        self.b_test.clicked.connect(self.test_scan)
        self.b_del_scan.clicked.connect(self.delete_scan)
        self.b_edit_scan.clicked.connect(self.edit_scan)
        self.b_new_window.clicked.connect(self.open_dialog)
        self.b_copy_host_path.clicked.connect(self.send_clipboard_name)
        self.b_copy_ip_path.clicked.connect(self.send_clipboard_ip_arm)
        self.b_clear_le_name.clicked.connect(self.le_name.clear)
        self.b_clr_hostname.clicked.connect(self.le_hostname.clear)
        self.b_clear_all.clicked.connect(self.clear_all)

    def test_scan(self):
        """Проверка записи сканирования."""
        ip = self.le_ip_prn.text()
        if self.is_status_code():
            num = self.le_num.text()
            if num:
                url = (f'http://{ip}/hp/device/info_ScantoFolder_testStatus.'
                       f'html?tab=Scan&menu=ScantoCfg?entryNum={num}')
                try:
                    driver.get(url=url)
                    self.update_status(f'Ожидайте, идет проверка записи '
                                       f'"{num}", это займет 15 секунд')
                    time.sleep(15)
                    msg = driver.find_element(By.ID, "alertText")
                    text_msg = msg.text.strip('\n')[:84]
                    self.update_status(text_msg)
                except Exception as e:
                    print(e)
                    self.populate_table()
                    self.update_status('Готово')
                finally:
                    self.populate_table()
            else:
                self.update_status('Необходимо указать запись '
                                   'для проверки')

    def is_printer(self, ip):
        """Проверка IP на принтер."""
        try:
            errorIndication, errorStatus, errorIndex, varBinds = next(
                getCmd(SnmpEngine(),
                       CommunityData('public'),
                       UdpTransportTarget((ip, 161)),
                       ContextData(),
                       ObjectType(ObjectIdentity('SNMPv2-MIB', 'sysDescr', 0)))
            )

            if errorIndication or errorStatus:
                self.t_scan.clearContents()
                self.update_status('Проверьте что IP-адрес является принтером')
                return False
            else:
                for name, val in varBinds:
                    if '426' in str(val).lower():
                        return True
                self.update_status('Проверьте что IP-адрес является принтером')
                return False
        except Exception:
            self.t_scan.clearContents()
            self.update_status('Проверьте что IP-адрес является принтером')
            return False

    def clear_all(self):
        """Очистка все полей ввода."""
        self.le_ip_prn.clear()
        self.le_username.clear()
        self.le_ip_arm.clear()
        self.le_name.clear()
        self.le_num.clear()
        self.te_text.clear()
        self.le_hostname.clear()
        self.le_pref.setText('сканирование')

    def open_dialog(self):
        """Открыть новое окно с настройками поля сканирования."""
        try:
            dialog = Dialog(self)
            dialog.show()
        except Exception as e:
            print(e)

    def parsing(self):
        """Парсинг настройки сканирования."""
        try:
            num = self.le_num.text()
            if num and num.isdigit():
                ip = self.le_ip_prn.text()
                url = (f'http://{ip}/hp/device/'
                       f'set_config_folderAddNew.html?tab=Scan&menu='
                       f'ScantoCfg?entryNum={num}')
                response = requests.get(url)
                html_content = response.text
                soup = BeautifulSoup(html_content, 'lxml')
                name = soup.find(
                    'input', {'id': 'displayName'}).get('value')
                if '�' in name:
                    name = ''
                folder = soup.find(
                    'input', {'id': 'networkFolderPath'}).get('value')
                pref = soup.find(
                    'input', {'id': 'filePrefix'}).get('value')

                selected_option = soup.findAll('option', selected=True)
                param = {
                    'Scan_PDF': 'PDF',
                    'Scan_JPEG': 'JPEG',
                    'A4': 'A4',
                    'LETTER': 'Letter',
                    'LEGAL': 'Legal',
                    'DPI_75': 'Экран - 75 т/д',
                    'DPI_150': 'Фото - 150 т/д',
                    'DPI_300': 'Текст - 300 т/д',
                    'DPI_600': 'Высота - 600 т/д',
                    'SCAN_COLOR': 'Цвет',
                    'SCAN_BW': 'Черно-белый'
                }
                element = [param[i.get('value')] for i in selected_option]
                file_type, size, dpi, color = element
                self.update_status(f'Открываю настройку сканирования "{num}"')
                return name, folder, pref, file_type, size, dpi, color
            else:
                self.update_status('Ошибка. Введите номер существующей'
                                   ' записи сканирования')
        except Exception as e:
            self.update_status('Ошибка. Введите номер существующей'
                               ' записи сканирования')
            print(e)

    def update_status(self, msg):
        """Изменение сообщения статусбара."""
        self.statusbar.showMessage(msg)
        self.statusbar.repaint()

    def is_valid_ip(self, ip):
        """Проверка корректности IP-адреса принтера."""
        pattern = r'^(\d{1,3})\.(\d{1,3})\.(\d{1,3})\.(\d{1,3})$'
        match = re.match(pattern, ip)
        if match:
            parts = [int(part) for part in match.groups()]
            if all(0 <= part <= 255 for part in parts):
                return True
        self.update_status('IP-адрес не указан или он не корректный. '
                           'Проверьте поле IP-адреса принтера')
        return False

    def is_status_code(self):
        """Проверка доступности принтера."""
        ip = self.le_ip_prn.text()
        if self.is_valid_ip(ip) and self.is_printer(ip):
            try:
                url = f'http://{ip}'
                response = requests.get(url, timeout=2)
                return response.status_code == 200
            except Exception:
                self.update_status('Устройство не в сети')

    def check_model(self, ip):
        """Проверка модели принтера."""
        try:
            url = f'http://{ip}'
            response = requests.get(url, timeout=2)
            html_content = response.text
            soup = BeautifulSoup(html_content, 'lxml')
            td_element = soup.find('td', {'class': 'mastheadTitle'})
            if td_element:
                model = td_element.find('h1').text
                pattern = r'\d{3}'
                match = re.search(pattern, model)
                print(match.group())
        except Exception as e:
            print(e)

    def populate_table(self):
        """Заполнение таблицы данными."""
        try:
            ip = self.le_ip_prn.text()
            if self.is_status_code() and self.is_printer(ip):
                for x in range(1, 21):
                    cnt = 0
                    url = (f'http://{ip}/hp/device/'
                           f'set_config_folderAddNew.html?tab=Scan&menu='
                           f'ScantoCfg?entryNum={x}')
                    response = requests.get(url)
                    html_content = response.text
                    soup = BeautifulSoup(html_content, 'lxml')
                    name = soup.find(
                        'input', {'id': 'displayName'}).get('value')
                    if '�' in name:
                        name = ''
                    folder = soup.find(
                        'input', {'id': 'networkFolderPath'}).get('value')

                    self.t_scan.setItem(x - 1, cnt, QTableWidgetItem(name))
                    self.t_scan.setItem(x - 1, cnt + 1, QTableWidgetItem(folder))
            else:
                self.t_scan.clearContents()
        except Exception as e:
            self.update_status(e)

    def save_button(self):
        """Действие кнопки 'Только сохранить.'"""
        try:
            name_ = WebDriverWait(driver, 1).until(
                EC.presence_of_element_located(
                    (By.ID, "displayNameError"))
            )
            text_name = name_.text
            folder_ = WebDriverWait(driver, 1).until(
                EC.presence_of_element_located(
                    (By.ID, "networkPathError"))
            )
            text_folder = folder_.text
            if text_name:
                self.update_status(f'Ошибка названия: {text_name}')
            elif text_folder:
                self.update_status(f'Ошибка пути: {text_folder}')
        except Exception:
            self.populate_table()
            self.update_status('Готово')
        finally:
            self.populate_table()

    def fill_scan(self, url):
        """Внесение данных для сканирования."""
        try:
            if self.is_status_code():
                driver.get(url=url)

                name = driver.find_element(By.ID, 'displayName')
                name.clear()
                name.send_keys(self.le_name.text())

                path = driver.find_element(By.ID, 'networkFolderPath')
                path.clear()
                hostname = self.le_hostname.text()
                username = self.le_username.text()
                ip = self.le_ip_arm.text()
                if not hostname:
                    path.send_keys(fr'\\{ip}\scan\{username}')
                else:
                    path.send_keys(fr'\\{hostname + ".mosgorzdrav.local"}'
                                   fr'\scan\{username}')

                login = driver.find_element(By.ID, 'UserName')
                login.clear()
                login.send_keys('scan')

                password = driver.find_element(By.ID, 'PassWord')
                password.clear()
                password.send_keys('ol23lrm')

                f_type = driver.find_element(By.ID, 'fileType')
                f_type.send_keys(self.cb_format.currentText())

                size = driver.find_element(By.ID, 'DefaultPaperSize')
                size.send_keys(self.cb_size.currentText())

                dpi = driver.find_element(By.ID, 'ScanQualitySelection')
                dpi.send_keys(self.cb_dpi.currentText())

                color = driver.find_element(By.ID, 'scanColorSelection')
                color.send_keys(self.cb_color.currentText())

                pref = driver.find_element(By.ID, 'filePrefix')
                pref.clear()
                pref.send_keys(self.le_pref.text())

                save = driver.find_element(By.ID, 'Save_button')
                save.click()
                self.save_button()
        except Exception as e:
            print(e)

    def delete_scan(self):
        """Удалить запись сканирования."""
        try:
            ip = self.le_ip_prn.text()
            if self.is_status_code():
                url = (f'http://{ip}/hp/device/'
                       f'set_config_scantoConfiguration.html?tab=Scan&amp;'
                       f'menu=ScantoCfg')
                driver.get(url=url)
                num = self.le_num.text()
                if num:
                    self.update_status(f'Ожидайте, удаляю запись'
                                       f'"{num}" сканирования...')

                    check_box = driver.find_element(
                        By.NAME, f'1.{int(num) - 1}.20.Select_Entry')
                    check_box.click()

                    b_delete = driver.find_element(
                        By.ID, 'Delete_Button')
                    b_delete.click()

                    alert = Alert(driver)
                    alert.accept()

                    self.populate_table()
                    self.update_status('Готово')
                else:
                    self.update_status('Введите номер записи для удаления')
        except Exception:
            self.update_status(f'Ошибка удаления записи. '
                               f'Записи "{num}" не существует')

    def check_scan(self):
        """Проверка сканирования."""
        self.update_status('Ожидайте, проверяю сканирование...')
        if self.is_status_code():
            self.populate_table()
            self.update_status('Готово')

    def new_scan(self):
        """Создание запись сканирования."""
        self.update_status('Ожидайте, '
                           'создаю запись для нового пользователя...')
        ip = self.le_ip_prn.text()
        url = (f'http://{ip}/hp/device/'
               f'set_config_folderAddNew.html?tab=Scan&amp;menu=ScantoCfg')
        self.fill_scan(url)

    def edit_scan(self):
        """Редактировать запись сканирования."""
        ip = self.le_ip_prn.text()
        num = self.le_num.text()
        self.update_status(f'Ожидайте, начинаю процесс изменения '
                           f'записи "{num}"')
        url = (f'http://{ip}/hp/device/'
               f'set_config_folderAddNew.html?tab=Scan&menu='
               f'ScantoCfg?entryNum={num}')
        self.fill_scan(url)

    def send_clipboard_name(self):
        """Команда для настройки сканирования через 'hostname' АРМ."""
        login = self.le_username.text()
        hostname = self.le_name.text()
        command = (fr"\\{hostname}.mosgorzdrav.local\scan\{login}"
                   .replace(" ", ""))
        QApplication.clipboard().setText(command)

    def send_clipboard_ip_arm(self):
        """Команда для настройки сканирования через 'ip-адрес' АРМ."""
        login = self.le_username.text()
        ip = self.le_ip_arm.text()
        command = fr"\\{ip}\scan\{login}".replace(" ", "")
        QApplication.clipboard().setText(command)


class Dialog(QDialog):
    def __init__(self, parent=None):
        super(Dialog, self).__init__(parent)
        uic.loadUi(r'ui\dialog.ui', self)
        self.main = parent
        self.setWindowFlags(
            self.windowFlags() & ~Qt.WindowContextHelpButtonHint)
        name, folder, pref, file_type, size, dpi, color = self.main.parsing()
        self.setWindowTitle(f'Настройки сканирования "{name}"')
        self.le_name.setText(name)
        self.le_folder.setText(folder)
        self.le_type.setText(file_type)
        self.le_size.setText(size)
        self.le_dpi.setText(dpi)
        self.le_color.setText(color)
        self.le_pref.setText(pref)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
