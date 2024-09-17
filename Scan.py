import os
import re
import requests
import sys
import time

from pysnmp.hlapi import *
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.alert import Alert
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from PyQt5 import uic, QtGui
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtWidgets import (QApplication, QMainWindow,
                             QTableWidgetItem, QDialog)


# Без привязки к версии
# # driver_options = webdriver.ChromeOptions()
# # driver_options.add_argument("--headless")
# # driver = webdriver.Chrome(options=driver_options)
#
# # С привязкой к версии
# driver_options = Options()
# driver_options.add_argument("--headless")
# driver_options.add_argument("--disable-blink-features=AutomationControlled")
# #
# # # Отключение автоматической проверки веб-драйвера
# service = Service(executable_path="webdriver/chromedriver.exe")
# #
# # # Создание экземпляра веб-драйвера Chrome с отключенной проверкой
# driver = webdriver.Chrome(service=service, options=driver_options)


class ScanThread(QThread):
    status_update = pyqtSignal(str)

    def __init__(self, ip, num, driver):
        super().__init__()
        self.ip = ip
        self.num = num
        self.driver = driver

    def run(self):
        if self.num:
            url = (f'http://{self.ip}/hp/device/info_ScantoFolder_testStatus.'
                   f'html?tab=Scan&menu=ScantoCfg?entryNum={self.num}')
            try:
                self.driver.get(url=url)
                self.status_update.emit(
                    f'Ожидайте, идет проверка записи '
                    f'"{self.num}", это займет немного времени')
                time.sleep(10)
                msg = self.driver.find_element(By.ID, "alertText")
                text_msg = msg.text.strip('\n')[:84]
                self.status_update.emit(text_msg)
            except Exception as e:
                print(e)
                self.status_update.emit('Готово')
        else:
            self.status_update.emit('Необходимо указать запись для проверки')


class MainWindow(QMainWindow):
    """"Главное окно."""
    def __init__(self):
        super(MainWindow, self).__init__()
        uic.loadUi(self.resource_path(r'ui\main_window.ui'), self)
        # С привязкой к версии
        driver_path = self.resource_path("webdriver/chromedriver.exe")
        service = Service(executable_path=driver_path)
        driver_options = Options()
        driver_options.add_argument("--headless")
        driver_options.add_argument(
            "--disable-blink-features=AutomationControlled")
        self.driver = webdriver.Chrome(service=service, options=driver_options)

        # Без привязки к версии
        # driver_options = webdriver.ChromeOptions()
        # driver_options.add_argument("--headless")
        # self.driver = webdriver.Chrome(options=driver_options)

        self.setWindowIcon(QtGui.QIcon(self.resource_path(r'ui\icons\logo.ico')))
        self.t_scan.setColumnWidth(0, 170)
        self.b_clr_le_ip_prn.clicked.connect(self.le_ip_prn.clear)
        self.b_clear_le_username.clicked.connect(self.le_username.clear)
        self.b_clr_le_ip_arm.clicked.connect(self.le_ip_arm.clear)
        self.b_check_scan.clicked.connect(self.check_scan)
        self.b_create.clicked.connect(self.new_scan)
        self.b_test.clicked.connect(self.test_scan)
        self.b_set_dns.clicked.connect(self.set_dns)
        self.b_del_scan.clicked.connect(self.delete_scan)
        self.b_edit_scan.clicked.connect(self.edit_scan)
        self.b_new_window.clicked.connect(self.open_window_dialog)
        self.b_copy_host_path.clicked.connect(self.send_clipboard_name)
        self.b_copy_ip_path.clicked.connect(self.send_clipboard_ip_arm)
        self.b_clear_le_name.clicked.connect(self.le_name.clear)
        self.b_clr_hostname.clicked.connect(self.le_hostname.clear)
        self.b_clear_all.clicked.connect(self.clear_all)

    def resource_path(self, relative_path):
        """Получает абсолютный путь к ресурсу,
        работает как в режиме разработки, так и в скомпилированном виде."""
        try:
            # PyInstaller создает временную папку и сохраняет путь к ней
            # в атрибуте `_MEIPASS`
            base_path = sys._MEIPASS
        except Exception:
            base_path = os.path.abspath(".")
        return os.path.join(base_path, relative_path)

    def test_scan(self):
        """Проверка записи сканирования."""
        ip = self.le_ip_prn.text()
        num = self.le_num.text()
        self.scan_thread = ScanThread(ip, num, self.driver)
        self.scan_thread.status_update.connect(self.update_status)
        self.scan_thread.start()

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
                    if 'HP LaserJet MFP M426' in str(val):
                        return True
                self.update_status('Проверьте что IP-адрес является принтером')
                return False
        except Exception:
            self.t_scan.clearContents()
            self.update_status('Проверьте что IP-адрес является принтером')
            return False

    def clear_all(self):
        """Очистка все полей ввода."""
        fields_to_clear = [
            self.le_ip_prn,
            self.le_username,
            self.le_ip_arm,
            self.le_name,
            self.le_num,
            self.te_text,
            self.le_hostname
        ]

        for field in fields_to_clear:
            if hasattr(field, 'clear'):
                field.clear()

        self.le_pref.setText('сканирование')

    def open_window_dialog(self):
        """Открыть новое окно с настройками поля сканирования."""
        try:
            dialog = Dialog(parent=self, resource_path=self.resource_path)
            dialog.show()
        except Exception as e:
            print(e)

    def set_dns(self):
        """Установить DNS сервер."""
        try:
            if self.is_status_code():
                self.update_status('Начался процесс изменения DNS-сервера')
                ip = self.le_ip_prn.text()
                'http://192.168.35.149/set_config_netIdentification.html?tab=Networking&menu=NetIdent'
                url = (f'http://{ip}/hp/device/set_config_netIdentification.'
                       f'html?tab=Networking&menu=NetIdent')
                self.driver.get(url=url)

                dns_primary = self.driver.find_element(By.ID, 'DNSPrimaryV4')
                dns_primary.clear()
                dns_primary.send_keys('10.0.4.14')

                dns_secondary = self.driver.find_element(By.ID, 'DNSSecV4')
                dns_secondary.clear()
                dns_secondary.send_keys('10.0.4.15')

                b_apply = self.driver.find_element(By.ID, 'apply_button')
                b_apply.click()

                alert = Alert(self.driver)
                alert.accept()
                self.update_status('DNS-сервер успешно изменен')
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
        except Exception:
            self.update_status('Ошибка. Введите номер существующей'
                               ' записи сканирования')

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
            if self.is_status_code():
                for num in range(1, 21):
                    cnt = 0
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

                    self.t_scan.setItem(
                        num - 1, cnt, QTableWidgetItem(name))
                    self.t_scan.setItem(
                        num - 1, cnt + 1, QTableWidgetItem(folder))
            else:
                self.t_scan.clearContents()
        except Exception as e:
            self.update_status(e)

    def save_button(self):
        """Действие кнопки 'Создать'."""
        try:
            # Проверка наличия ошибки в названии
            name_error = WebDriverWait(self.driver, 1).until(
                EC.presence_of_element_located((By.ID, "displayNameError"))
            )
            text_name_error = name_error.text

            # Проверка наличия ошибки в пути
            folder_error = WebDriverWait(self.driver, 1).until(
                EC.presence_of_element_located((By.ID, "networkPathError"))
            )
            text_folder_error = folder_error.text

            # Обработка ошибок
            if text_name_error:
                self.update_status(f'Ошибка названия: {text_name_error}')
            elif text_folder_error:
                self.update_status(f'Ошибка пути: {text_folder_error}')
            else:
                self.populate_table()
                self.update_status('Готово')
        except Exception:
            self.populate_table()
            self.update_status('Готово')

    def fill_scan(self, url):
        """Внесение данных для сканирования."""
        try:
            if self.is_status_code():
                self.driver.get(url=url)

                name = self.driver.find_element(By.ID, 'displayName')
                name.clear()
                name.send_keys(self.le_name.text())

                path = self.driver.find_element(By.ID, 'networkFolderPath')
                path.clear()

                hostname = self.le_hostname.text()
                username = self.le_username.text()
                ip = self.le_ip_arm.text()
                if not hostname:
                    path.send_keys(fr'\\{ip}\scan\{username}')
                else:
                    path.send_keys(fr'\\{hostname + ".mosgorzdrav.local"}'
                                   fr'\scan\{username}')

                login = self.driver.find_element(By.ID, 'UserName')
                login.clear()
                login.send_keys('scan')

                password = self.driver.find_element(By.ID, 'PassWord')
                password.clear()
                password.send_keys('ol23lrm')

                f_type = self.driver.find_element(By.ID, 'fileType')
                f_type.send_keys(self.cb_format.currentText())

                size = self.driver.find_element(By.ID, 'DefaultPaperSize')
                size.send_keys(self.cb_size.currentText())

                dpi = self.driver.find_element(By.ID, 'ScanQualitySelection')
                dpi.send_keys(self.cb_dpi.currentText())

                color = self.driver.find_element(By.ID, 'scanColorSelection')
                color.send_keys(self.cb_color.currentText())

                pref = self.driver.find_element(By.ID, 'filePrefix')
                pref.clear()
                pref.send_keys(self.le_pref.text())

                save = self.driver.find_element(By.ID, 'Save_button')
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
                self.driver.get(url=url)
                num = self.le_num.text()
                if num:
                    self.update_status(f'Ожидайте, удаляю запись'
                                       f'"{num}" сканирования...')

                    check_box = self.driver.find_element(
                        By.NAME, f'1.{int(num) - 1}.20.Select_Entry')
                    check_box.click()

                    b_delete = self.driver.find_element(
                        By.ID, 'Delete_Button')
                    b_delete.click()

                    alert = Alert(self.driver)
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
    """Диалоговое окно."""
    def __init__(self, parent=None, resource_path=None):
        super(Dialog, self).__init__(parent)
        self.resource_path = resource_path
        uic.loadUi(resource_path(r'ui\settings.ui'), self)
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
