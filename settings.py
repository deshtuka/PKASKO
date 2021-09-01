from dotenv import load_dotenv
import os

class Config:
    # Название КАТАЛОГА - основная папка для вывода
    trunk = '_Результаты'

    # Название ПОДКАТАЛОГОВ - второстепенные папки указанной выше
    branches = [
        'Бэкапы страниц',
        'Скрины расчета',
        'Логи',
        'Частотные словари'
    ]

    # Название потоков
    twig = [
        'КАСКО - Предварительный расчет',
        'КАСКО - Точный расчет'
    ]

    # Название Страховых компаний по котором релизованы скрипты B2B
    ic = {
        1: 'Альфа',
        2: 'РЕСО',
        3: 'Совкомбанк',
        4: 'Ренессанс',
        5: 'МАКС',
        6: 'Ингосстрах',
        7: 'Тинькофф'
    }

    @staticmethod   # Т.к. в функции используются данные без предварительной инициализации в классе
    def key_password(view):
        dotenv_path = os.path.join(os.path.dirname(r'C:\Python_works\_!Настройки\\'), 'all_pass.env')

        if os.path.exists(dotenv_path):
            load_dotenv(dotenv_path)

            if view == 1:
                host, name, user, password = ('' for i in range(4))
                try:
                    host = os.environ.get('DB_HOST')
                    name = os.environ.get('DB_NAME')
                    user = os.environ.get('DB_USER')
                    password = os.environ.get('DB_PASS')
                except:
                    print(f'Отстутсвуют данные для подключения к Базе Данных')
                finally:
                    return host, name, user, password

            if view == 2:
                user, password = ('' for i in range(2))
                try:
                    user = os.environ.get('PKASKO_USER')
                    password = os.environ.get('PKASKO_PASS')
                except:
                    print(f'Отстутсвует Логин/Пароль для подключения к ПКАСКО')
                finally:
                    return user, password

        else:
            print(f'\nОтсутствует файл с конфигурационными настройками!\n')
