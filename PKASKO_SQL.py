# -*- coding: utf-8 -*-
from settings import Config
import mysql.connector
import pandas as pd
import os


def db_pkasko():
    try:
        host, database, user, password = Config.key_password(1)
        conn = mysql.connector.connect(
            host=host,
            database=database,
            user=user,
            password=password
        )
    except mysql.connector.errors.ProgrammingError:
        print(f'Нет связи!')

    else:
        case = (
            'SELECT python_calc_test_case.*, python_calc_test_car.*, python_calc_test_beneficiary.*'
            ' FROM python_calc_test_case'
            ' RIGHT JOIN python_calc_test_car ON python_calc_test_case.id_car = python_calc_test_car.id'
            ' RIGHT JOIN python_calc_test_beneficiary ON python_calc_test_case.id_beneficiary = python_calc_test_beneficiary.id'
        )
        people_1 = (
            'SELECT python_calc_test_case.*, python_calc_test_people.*'
            ' FROM python_calc_test_case'
            ' RIGHT JOIN python_calc_test_people ON python_calc_test_people.id = python_calc_test_case.id_people_1'
            ' WHERE python_calc_test_case.id_people_1 IS NOT NULL'
        )
        people_2 = (
            'SELECT python_calc_test_case.*, python_calc_test_people.*'
            ' FROM python_calc_test_case'
            ' RIGHT JOIN python_calc_test_people ON python_calc_test_people.id = python_calc_test_case.id_people_2'
            ' WHERE python_calc_test_case.id_people_2 IS NOT NULL'
        )

        # SQL запросы в базу
        df_case = pd.read_sql(case, conn)
        df_people_1 = pd.read_sql(people_1, conn)
        df_people_2 = pd.read_sql(people_2, conn)

        case = dict()
        # Убирается повтор ID, приводится к общему виду Словарь, где {0: Параметры кейса, 1: Человек #1, 2: Человек #2}
        for idx, one in enumerate(['df_case', 'df_people_1', 'df_people_2']):
            tmp = eval(f'{one}.T.apply(lambda x: unical_id(x))')
            case[idx] = fix_dict(tmp)

        # Отображение Тест-кейсов в консоле
        number = selection_vision_db(case)

        # Выбранный кейс
        one_case = {i: case.get(i).get(number) for i in range(3)}

        # Добавление в словарь {status: True/False} который указывает на тип (0-Выгодоприобр, 1-Водитель #1, 2-Водитель #2)
        for i in range(3):
            if one_case.get(i) is not None:
                one_case.get(i).update({'status': {i: people_status_in_db(case).get(i).get(number) for i in range(3)}.get(i)})
            else:
                del one_case[i]

        text = one_case.get(1).get('license_date')

        # Возвращаем выбранный кейс в виде словаря | Структура: {0: {Параметры кейса}, 1: {Человек #1}, 2: {Человек #2}}
        return one_case


# Добавление уникальности для ячеек ID
def unical_id(value):
    tmp = dict()
    count = 0

    for key, item in value.items():
        if key == 'id':
            tmp[f'{key}_{count}'] = item
            count += 1
        else:
            tmp[key] = item

    return tmp


# Конвертация ID кейса в название словаря
def fix_dict(body):
    test = dict()

    for name, items in body.items():
        tmp = list()
        count = 0
        for key, value in items.items():
            if key == 'id_0':
                count = value
            else:
                tmp.append([key, value])
        else:
            test[count] = dict(tmp)

    return test


# Отображение Тест-кейсов в консоле
def selection_vision_db(case):
    # Статус людей (Выгодоприобретатель / Водитель 1 / Водитель 2)
    status = people_status_in_db(case)

    # Получение: ID кейса, марки, модели, год ТС
    list_1 = [[id, items.get("make"), items.get("model"), items.get("year")] for id, items in case.get(0).items()]
    # Получение: ID кейса, Человек #1 (Фамилия, Имя)
    list_2 = [[id, items.get("lastname"), items.get("firstname")] for id, items in case.get(1).items()]
    # Получение: ID кейса, Человек #2 (Фамилия, Имя)
    list_3 = [[id, items.get("lastname"), items.get("firstname")] for id, items in case.get(2).items()]

    base = {}
    for one in list_1:
        # Авто
        base[one[0]] = [' '.join(map(str, one[1:]))]
        # Страхователь
        [base.setdefault(one[0], []).append(' '.join(two[1:])) for two in list_2 if one[0] == two[0]]
        # Водитель
        tmp = list()
        [tmp.append(' '.join(two[1:])) for two in list_2 if one[0] == two[0] and status.get(1).get(one[0])]
        [tmp.append(' '.join(three[1:])) for three in list_3 if one[0] == three[0] and status.get(2).get(three[0])]
        [base.setdefault(one[0], []).append(' & '.join(tmp)) for two in list_2 if one[0] == two[0]]
        # Выгодоприобретатель
        value = [' '.join(two[1:]) for two in list_2 if one[0] == two[0] and not status.get(0).get(one[0])]
        base.setdefault(one[0], []).extend(value if value else f'Банк [ИНН: {case.get(0).get(one[0]).get("inn")}]')

    while True:
        print('\n{:^3} | {:^25} | {:^25} | {:^45} | {:^25} |'.format('ID', 'Автомобиль', 'Страхователь', 'Водитель(-я)', 'Выгодоприобретатель'))
        print('{:^3} | {:^25} | {:^25} | {:^45} | {:^25} |'.format('---', '--------------------', '--------------------', '--------------------', '--------------------'))
        for idx, name in base.items():
            try:
                print('{:^3d} | {:<25} | {:<25} | {:<45} | {:<25} |'.format(idx, name[0], name[1], name[2], name[3]))
            except IndexError:
                print('{:^3d} | {:<25} | {:<25} | {:<25} |'.format(idx, name[0], name[1], name[2]))

        try:
            data = int(input('\nВведите ID тест-кейса = '))
        except ValueError:
            os.system('cls')
        else:
            if base.get(data):
                break

    print(f'Выбран Тест-кейс #{data}: {" | ".join(base.get(data))}\n')
    return data


# Определение статусов у людей
def people_status_in_db(case):
    status = dict()
    for idx, items in case.items():
        for id, values in items.items():

            # idx = 0 -> Выгодоприобретатель - Юр.лицо?
            # idx = 1 -> Человек #1 - Водитель?
            # idx = 2 -> Человек #2 - Водитель?
            value = (True if values.get('inn') else False) if idx == 0 else (True if values.get('driver') == 1 else False)

            status.setdefault(idx, {}).update({id: value})

    return status


# def main():
#     db_pkasko()
#
#
# if __name__ == '__main__':
#     main()
