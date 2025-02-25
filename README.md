# Кандидатская диссертация на тему:  
# "Программа для отправки электронных писем с вложениями на основе данных из Excel файла. С подробным разбором кода"  
---------------------
  
ЯП - Python  
  
1 часть. Чистый код.  


-----------------------------------------------
# Начало кода
  
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
  
# Конец кода
-----------------------------------------------

-----------------------------------------------

-----------------------------------------------
  
#  2. часть. Код с разбором.  
-----------------------------------------------

Условные обозначения:  
1. код: коментарий с разбором кода  
  
  
# Начало кода 
# Импорт библиотек  
  
import pandas as pd  
import smtplib  
import os  
import re  
from email.mime.multipart import MIMEMultipart  
from email.mime.base import MIMEBase  
from email import encoders  
from email.mime.text import MIMEText  
from email.header import Header  
  
1. import pandas as pd: Импортирует библиотеку pandas, которая используется для работы с данными в табличном формате, в частности, для чтения и обработки Excel файлов.  
1. import smtplib: Импортирует модуль smtplib, который предоставляет функции для отправки электронной почты через SMTP (Simple Mail Transfer Protocol).  
1. import os: Импортирует модуль os, который предоставляет функции для взаимодействия с операционной системой, например, для работы с файловой системой.  
1. import re: Импортирует модуль re, который предоставляет функции для работы с регулярными выражениями, что позволяет проверять строки на соответствие определенным шаблонам.  
1. from email.mime.multipart import MIMEMultipart: Импортирует класс MIMEMultipart, который используется для создания многочастных сообщений электронной почты (например, с текстом и вложениями).  
1. from email.mime.base import MIMEBase: Импортирует класс MIMEBase, который используется для создания базовых частей сообщения, таких как вложения.  
1. from email import encoders: Импортирует модуль encoders, который предоставляет функции для кодирования вложений в сообщения электронной почты.  
1. from email.mime.text import MIMEText: Импортирует класс MIMEText, который используется для создания текстовых частей сообщения.  
1. from email.header import Header: Импортирует класс Header, который используется для кодирования заголовков сообщений, таких как имя файла вложения.  
  
  
  
# Функция для проверки валидности email  
  
def is_valid_email(email):  
    pattern = r"^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$"  
    return bool(re.match(pattern, email))  
      
1. def is_valid_email(email):: Определяет функцию is_valid_email, которая принимает один аргумент email.  
1. pattern = r"^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$": Определяет регулярное выражение для проверки формата email. Шаблон проверяет, что email состоит из допустимых символов, за которыми следует символ @, доменное имя и доменная зона.  
1. return bool(re.match(pattern, email)): Использует re.match для проверки, соответствует ли переданный email шаблону. Возвращает True, если соответствует, и False в противном случае.  
  
  
  
# Конфигурация SMTP сервера  
  
smtp_server = '192.168.10.5'  
smtp_port = 25  # Обычно порт 25 для SMTP  
  
1. smtp_server = '192.168.10.5': Указывает IP-адрес SMTP сервера, который будет использоваться для отправки писем.  
1. smtp_port = 25: Указывает порт, используемый для подключения к SMTP серверу. Порт 25 обычно используется для SMTP.  
  
  
  
# Путь к директории с файлами  
  
files_dir = 'C:/PY/rassilka2/pythonProject/dir/'  
  
1. files_dir = 'C:/PY/rassilka2/pythonProject/dir/': Определяет путь к директории, где находятся файлы, которые будут отправлены в качестве вложений.  
  
  
  
# Чтение Excel файла  
  
df = pd.read_excel('C:/PY/rassilka2/pythonProject/xls/1.xlsx', header=None)  
  
1. df = pd.read_excel('C:/PY/rassilka2/pythonProject/xls/1.xlsx', header=None): Использует pandas для чтения Excel файла по указанному пути. Параметр header=None указывает, что в файле нет заголовков, и все строки будут считаны как данные.  
  
  
  
# Функция для отправки письма с вложением  
  
def send_email(to_email, attachment_filename):  
  
1. def send_email(to_email, attachment_filename):: Определяет функцию send_email, которая принимает два аргумента: to_email (адрес электронной почты получателя) и attachment_filename (имя файла, который будет вложен в письмо).  
  
  
  
# Создание сообщения  
  
    msg = MIMEMultipart()  
    msg['From'] = 'celebration@andruhes.ru'  # Укажите ваш email  
    msg['To'] = to_email  
    msg['Subject'] = 'С 23 февраля!'  
  
1. msg = MIMEMultipart(): Создает объект msg типа MIMEMultipart, который будет использоваться для формирования многочастного сообщения (с текстом и вложениями).  
1. msg['From'] = 'celebration@andruhes.ru': Устанавливает адрес отправителя в заголовке сообщения. Здесь указан email отправителя.  
1. msg['To'] = to_email: Устанавливает адрес получателя в заголовке сообщения, используя переданный аргумент to_email.  
1. msg['Subject'] = 'С 23 февраля!': Устанавливает тему письма в заголовке сообщения.  
  
  
  
# Текст письма  
  
    body = 'Поздравляем!'  
    msg.attach(MIMEText(body, 'plain'))  
  
1. body = 'Поздравляем!': Определяет текстовое содержимое письма.  
1. msg.attach(MIMEText(body, 'plain')): Создает текстовую часть сообщения с помощью MIMEText и прикрепляет ее к объекту msg. Параметр 'plain' указывает, что текст будет в обычном формате (без форматирования).  
  
  
  
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
  
1. attachment_path = files_dir + attachment_filename: Формирует полный путь к файлу вложения, объединяя путь к директории files_dir и имя файла attachment_filename.  
1. try:: Начинает блок try, чтобы обработать возможные исключения при работе с файлами.  
1. with open(attachment_path, 'rb') as attachment:: Открывает файл по указанному пути в бинарном режиме ('rb'). Если файл успешно открыт, он будет доступен через переменную attachment.  
1. part = MIMEBase('application', 'octet-stream'): Создает объект part типа MIMEBase, который будет использоваться для вложения. 'application' и 'octet-stream' указывают на тип содержимого.  
1. part.set_payload(attachment.read()): Читает содержимое файла и устанавливает его как полезную нагрузку для объекта part.  
1. encoders.encode_base64(part): Кодирует содержимое вложения в формате Base64, чтобы оно могло быть корректно передано по электронной почте.  
1. encoded_filename = Header(attachment_filename, 'utf-8').encode(): Кодирует имя файла с использованием класса Header, чтобы обеспечить правильное отображение символов в заголовке.  
1. part.add_header('Content-Disposition', f'attachment; filename="{encoded_filename}"'): Добавляет заголовок Content-Disposition, который указывает, что это вложение, и задает имя файла для отображения.  
1. msg.attach(part): Прикрепляет объект part (вложение) к сообщению msg.  
  
  
  
# Обработка ошибок при открытии файла  
  
    except FileNotFoundError:  
        print(f'Файл {attachment_filename} не найден в директории {files_dir}. Письмо на {to_email} не отправлено.')  
        return  
  
1. except FileNotFoundError:: Обрабатывает исключение, если файл не найден.  
1. **`print(f'Файл {attachment_filename} не найден в директории  
  
  
# Обработка ошибок при открытии файла  
  
    except FileNotFoundError:  
        print(f'Файл {attachment_filename} не найден в директории {files_dir}. Письмо на {to_email} не отправлено.')  
        return  
  
1. except FileNotFoundError:: Этот блок обрабатывает исключение, которое возникает, если файл, указанный в attachment_path, не найден. Это позволяет избежать аварийного завершения программы.  
1. print(f'Файл {attachment_filename} не найден в директории {files_dir}. Письмо на {to_email} не отправлено.'): Выводит сообщение об ошибке в консоль, информируя пользователя о том, что файл не найден и письмо не было отправлено.  
1. return: Завершает выполнение функции send_email, если файл не найден, чтобы предотвратить дальнейшие действия по отправке письма.  
  
  
  
# Отправка письма  
  
    try:  
        server = smtplib.SMTP(smtp_server, smtp_port)  
        server.sendmail(msg['From'], msg['To'], msg.as_string())  
        server.quit()  
        print(f'Письмо успешно отправлено на {to_email} с вложением {attachment_filename}')  
    except Exception as e:  
        print(f'Ошибка при отправке письма на {to_email}: {e}')  
  
1. try:: Начинает новый блок try, чтобы обработать возможные исключения при отправке письма.  
1. server = smtplib.SMTP(smtp_server, smtp_port): Создает объект server для подключения к SMTP серверу, используя указанный IP-адрес и порт.  
1. server.sendmail(msg['From'], msg['To'], msg.as_string()): Отправляет письмо, используя метод sendmail. В качестве аргументов передаются адрес отправителя, адрес получателя и строковое представление сообщения (включая текст и вложения).  
1. server.quit(): Закрывает соединение с SMTP сервером после отправки письма.  
1. print(f'Письмо успешно отправлено на {to_email} с вложением {attachment_filename}'): Выводит сообщение в консоль, подтверждающее успешную отправку письма.  
1. except Exception as e:: Обрабатывает любые исключения, которые могут возникнуть при отправке письма.  
1. print(f'Ошибка при отправке письма на {to_email}: {e}'): Выводит сообщение об ошибке в консоль, информируя пользователя о том, что произошла ошибка при отправке письма, и выводит текст ошибки.  
  
  
  
# Проход по всем строкам в Excel файле  
  
for index, row in df.iterrows():  
    email = row.iloc[0]  # Адрес email из первого столбца  
    filename = row.iloc[1]  # Имя файла из второго столбца  
  
1. for index, row in df.iterrows():: Начинает цикл, который проходит по всем строкам DataFrame df, созданному из Excel файла. index — это индекс строки, а row — это объект Series, представляющий данные в строке.  
1. email = row.iloc[0]: Извлекает адрес электронной почты из первого столбца текущей строки.  
1. filename = row.iloc[1]: Извлекает имя файла из второго столбца текущей строки.  
  
  
  
# Проверка на пустые значения  
  
    if pd.isna(email) or pd.isna(filename):  
        print(f"Пустая строка в строке {index + 1}. Пропуск.")  
        continue  
  
1. if pd.isna(email) or pd.isna(filename):: Проверяет, является ли адрес электронной почты или имя файла пустым значением (NaN).  
1. print(f"Пустая строка в строке {index + 1}. Пропуск."): Если одно из значений пустое, выводит сообщение о том, что строка пропускается.  
1. continue: Пропускает текущую итерацию цикла и переходит к следующей строке.  
  
  
  
# Проверка корректности email  
  
    if not is_valid_email(email):  
        print(f"Адрес электронной почты {email} некорректный. Пропуск.")  
        continue  
  
1. if not is_valid_email(email):: Вызывает функцию is_valid_email, передавая ей адрес электронной почты email. Если функция возвращает False (т.е. email некорректный), условие выполняется.  
1. print(f"Адрес электронной почты {email} некорректный. Пропуск."): Выводит сообщение в консоль, информируя о том, что указанный адрес электронной почты некорректен, и строка будет пропущена.  
1. continue: Пропускает текущую итерацию цикла и переходит к следующей строке в DataFrame.  
  
  
  
# Проверка существования файла  
  
    attachment_path = files_dir + filename  
    if not os.path.exists(attachment_path):  
        print(f"Файл {filename} не найден для адреса {email}. Пропуск.")  
        continue  
  
1. attachment_path = files_dir + filename: Формирует полный путь к файлу вложения, объединяя путь к директории files_dir и имя файла filename.  
1. if not os.path.exists(attachment_path):: Проверяет, существует ли файл по указанному пути attachment_path. Если файл не существует, условие выполняется.  
1. print(f"Файл {filename} не найден для адреса {email}. Пропуск."): Выводит сообщение в консоль, информируя о том, что файл не найден для указанного адреса электронной почты, и строка будет пропущена.  
1. continue: Пропускает текущую итерацию цикла и переходит к следующей строке.  
  
  
  
# Отладочный вывод  
  
    print(f"Обработка строки {index + 1}: Email: {email}, Файл: {filename}")  
  
1. print(f"Обработка строки {index + 1}: Email: {email}, Файл: {filename}"): Выводит сообщение в консоль, информируя о том, что происходит обработка текущей строки, и показывает адрес электронной почты и имя файла, которые будут использованы для отправки письма.  
  
  
  
# Отправка письма  
  
    send_email(email, filename)  
  
1. send_email(email, filename): Вызывает функцию send_email, передавая ей адрес электронной почты email и имя файла filename. Это инициирует процесс отправки письма с вложением.

# Конец кода
-----------------------------------------------
© 2025 Маркушев А.М.
Итог.  
Таким образом, данный код выполняет следующие основные задачи:  
  
Читает данные из Excel файла, содержащего адреса электронной почты и имена файлов.  
Проверяет корректность адресов электронной почты и существование файлов.  
Отправляет письма с вложениями на указанные адреса электронной почты, выводя сообщения о статусе обработки каждой строки.  
Если у вас есть дополнительные вопросы или вам нужно больше информации по какому-либо аспекту кода, дайте знать!  

