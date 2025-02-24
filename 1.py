import pandas as pd
import smtplib
import os
import re
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from email.mime.text import MIMEText
from email.header import Header  # Импортируем класс Header для кодирования имени файла


# Функция для проверки валидности email
def is_valid_email(email):
    pattern = r"^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$"
    return bool(re.match(pattern, email))


# Конфигурация SMTP сервера
smtp_server = '192.168.10.5'
smtp_port = 25  # Обычно порт 25 для SMTP

# Путь к директории с файлами
files_dir = 'C:/PY/rassilka2/pythonProject/dir/'

# Чтение Excel файла
df = pd.read_excel('C:/PY/rassilka2/pythonProject/xls/1.xlsx', header=None)


# Функция для отправки письма с вложением
def send_email(to_email, attachment_filename):
    # Создание сообщения
    msg = MIMEMultipart()
    msg['From'] = 'celebration@andruhes.ru'  # Укажите ваш email
    msg['To'] = to_email
    msg['Subject'] = 'С 23 февраля!'

    # Текст письма
    body = 'Поздравляем!'
    msg.attach(MIMEText(body, 'plain'))

    # Вложение
    attachment_path = files_dir + attachment_filename
    try:
        with open(attachment_path, 'rb') as attachment:
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(attachment.read())
            encoders.encode_base64(part)

            # Кодируем имя файла с использованием класса Header
            encoded_filename = Header(attachment_filename, 'utf-8').encode()
            part.add_header('Content-Disposition', f'attachment; filename="{encoded_filename}"')
            msg.attach(part)
    except FileNotFoundError:
        print(f'Файл {attachment_filename} не найден в директории {files_dir}. Письмо на {to_email} не отправлено.')
        return

    # Отправка письма
    try:
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.sendmail(msg['From'], msg['To'], msg.as_string())
        server.quit()
        print(f'Письмо успешно отправлено на {to_email} с вложением {attachment_filename}')
    except Exception as e:
        print(f'Ошибка при отправке письма на {to_email}: {e}')


# Проход по всем строкам в Excel файле
for index, row in df.iterrows():
    email = row.iloc[0]  # Адрес email из первого столбца
    filename = row.iloc[1]  # Имя файла из второго столбца

    # Проверка на пустые значения
    if pd.isna(email) or pd.isna(filename):
        print(f"Пустая строка в строке {index + 1}. Пропуск.")
        continue

    # Проверка корректности email
    if not is_valid_email(email):
        print(f"Адрес электронной почты {email} некорректный. Пропуск.")
        continue

    # Проверка существования файла
    attachment_path = files_dir + filename
    if not os.path.exists(attachment_path):
        print(f"Файл {filename} не найден для адреса {email}. Пропуск.")
        continue

    # Отладочный вывод
    print(f"Обработка строки {index + 1}: Email: {email}, Файл: {filename}")

    # Отправка письма
    send_email(email, filename)