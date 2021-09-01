# -*- coding: utf-8 -*-
from settings import Config

import pandas as pd
from collections import Counter
from xlsxwriter.utility import xl_rowcol_to_cell
from datetime import date
import time
from traceback import print_exc


# Вывод всех данных в Эксель
def export_to_excel(GL_good_2, GL_good_1):
    # # Выставляем последовательность вывода данных
    # GL_good_1, GL_good_2 = priority_setting_export(GL_good_1, GL_good_2)

    data_today = date.today().strftime('%d.%m.%Y')
    time_now = time.strftime("%H.%M", time.localtime())

    file_excel = f'{Config.trunk}/Расчет ПКАСКО [{data_today} - {time_now}].xlsx'

    # ПРЕДВАРИТЕЛЬНЫЙ расчет
    try:
        print(f'GL_good_1 = {GL_good_1}')
        good_1 = splitting_by_titles(GL_good_1, columns_export=len(GL_good_1[0][0]))        # Параметры

        # Соединение ПАРАМЕТРОВ и ПРЕМИИ
        dict_1 = dict()

        # ПАРАМЕТРЫ
        for it in range(len(good_1.values)):
            dict_1[it] = good_1.values[it]

        # Объявление Dataframe
        df_1 = pd.DataFrame(data=dict_1)

        # Транспонирование и очистка заголовков
        df_1 = df_1.T
        df_1[0] = df_1[0].apply(lambda x: panda_trash(x))

    except Exception as ex:
        print(f'ОШИБКА ВЫВОДА В ЭКСЕЛЬ - ПРЕДВАРИТЕЛЬНОГО РАСЧЕТА!\n  Сообщение: {ex}')
        print_exc()

    else:
        # ИТОГОВЫЙ расчет
        try:
            good_2 = splitting_by_titles(GL_good_2, columns_export=len(GL_good_2[0][0]))   # Параметры

            # Соединение ПАРАМЕТРОВ и ПРЕМИИ
            dict_2 = dict()

            # ПАРАМЕТРЫ
            for it in range(len(good_2.values)):
                dict_2[it] = good_2.values[it]

            df_2 = pd.DataFrame(data=dict_2)

            df_2 = df_2.T

            df_2[0] = df_2[0].apply(lambda x: panda_trash(x))

        except Exception as ex:
            # print_exc()
            # ######## Если появилась ошибка, вывод ПРЕДВАРИТЕЛЬНОГО расчета ######## #

            print(f'Ошибка в ИТОГОВОМ -> Вывод только предварительного!\n  Сообщение:{ex}')

            with pd.ExcelWriter(file_excel) as writer:
                # Преобразовать фрейм данных в объект Excel XlsxWriter
                df_1.to_excel(writer, sheet_name='Предварительный', index=False, header=False)
                sheet = writer.sheets['Предварительный']
                workbook = writer.book
                setting = FormatExport()

                fmt = setting.design(color=False, bold=False, wrap=True, valign='l')
                first_fmt = workbook.add_format(fmt)

                fmt = setting.design(color=False, bold=False, wrap=True)
                second_fmt = workbook.add_format(fmt)

                format1 = workbook.add_format({'bold': 1, 'bg_color': '#1DACD6'})

                # Находить уникальные ячейки в строке
                for x in range(0, df_1.shape[0] - 1):
                    sheet.conditional_format('B{}:F{}'.format(x, x), {'type': 'unique', 'format': format1})

                # Первый столбец применяет формат
                sheet.set_column(0, 0, 27.86, first_fmt)

                # Весь лист применяет формат
                sheet.set_column(1, len(df_1.columns) - 1, 27.86, second_fmt)

                # writer.save()

        else:
            with pd.ExcelWriter(file_excel) as writer:
                df_1.to_excel(writer, sheet_name='Предварительный', index=False, header=False)
                # df_1_prize.to_excel(writer, sheet_name='ПредПремия', index=False, header=False)
                df_2.to_excel(writer, sheet_name='Итоговый', index=False, header=False)

                sheet_1 = writer.sheets['Предварительный']
                # sheet_1_prize = writer.sheets['ПредПремия']
                sheet_2 = writer.sheets['Итоговый']
                workbook = writer.book
                setting = FormatExport()

                fmt = setting.design(color=False, bold=False, wrap=True, valign='l')
                first_fmt = workbook.add_format(fmt)

                fmt = setting.design(color=False, bold=False, wrap=True)
                second_fmt = workbook.add_format(fmt)

                format1 = workbook.add_format({'bold': 1, 'bg_color': '#1DACD6'})

                # //////////////////// СТИЛЬ ЗАГОЛОВКОВ -> ПАРАМЕТРОВ ////////////////////
                fmt = setting.design(color='#C86B85', valign='l')
                fmt_head_1 = workbook.add_format(fmt)

                fmt = setting.design(color='#FDA886')
                fmt_head_2 = workbook.add_format(fmt)

                fmt = setting.design(color='#F3D7CA')
                fmt_head_3 = workbook.add_format(fmt)

                fmt = setting.design(color='#F3EECA')
                fmt_head_4 = workbook.add_format(fmt)

                # //////////////////// СТИЛЬ ЗАГОЛОВКОВ -> ПРЕМИЙ ////////////////////
                fmt = setting.design(color='#1FAB89', valign='l')
                fmt_prime_1 = workbook.add_format(fmt)

                fmt = setting.design(color='#62D2A2')
                fmt_prime_2 = workbook.add_format(fmt)

                # Заголовки - которые нужно закрасить
                # Параметры
                heading_1_param = [
                    'Водители',
                    'Транспортное средство',
                    'Противоугонные системы',
                    'Параметры договора страхования',
                    'Несчастный случай (НС)',
                    'Несчастный случай',
                    'Доп. продукты',
                    'Базовые параметры',
                    'Скоринг',
                    'Условия страхования',
                    'Лица, допущенные к управлению',
                    'Дополнительные условия',
                    'Франшиза',
                    'Дополнительное оборудование',
                    'Гражданская ответственность',
                    'Клиент',
                    'ШАГ 1. Выберите участника(ов) договора',
                    'ШАГ 2. Транспортное средство',
                    'ШАГ 3. Противоугонные системы',
                    'ШАГ 4. Параметры договора страхования',
                    'ШАГ 5. Несчастный случай (НС)'
                ]
                heading_2_param = [
                    'Водитель №1',
                    'Водитель №2',
                    'Статус клиента',
                    'Личные данные',
                    'Документы'
                ]
                heading_3_param = [
                    'Паспорт',
                    'Данные предыдущего паспорта',
                    'Водительское удостоверение',
                    'Предыдущее водительское удостоверение',
                    'Контактные данные'
                ]

                heading_4_param = [
                    'Адрес регистрации'
                ]

                # Премии
                heading_1_prime = ['Премия, руб']

                heading_2_prime = ['Страховая компания']

                # Дублирование Заголовков Премий - избежать изменений
                heading_1_prime_2 = heading_1_prime.copy()
                heading_2_prime_2 = heading_2_prime.copy()

                # Для ПРЕДВАРИТЕЛЬНОГО РАСЧЕТА //////////////////////////////////////
                """
                df_2            - DataFrame
                sheet_1         - Название листа
                heading_1_param - Название параметров
                fmt_head_1      - Стиль
                """
                # Параметры
                style_of_design(df_1, sheet_1, heading_1_param, fmt_head_1)
                style_of_design(df_1, sheet_1, heading_2_param, fmt_head_2)
                style_of_design(df_1, sheet_1, heading_3_param, fmt_head_3)
                style_of_design(df_1, sheet_1, heading_4_param, fmt_head_4)
                # Премии
                style_of_design(df_1, sheet_1, heading_1_prime, fmt_prime_1)
                style_of_design(df_1, sheet_1, heading_2_prime, fmt_prime_2)

                # Для ИТОГОВОГО РАСЧЕТА //////////////////////////////////////
                # Параметры
                style_of_design(df_2, sheet_2, heading_1_param, fmt_head_1)
                style_of_design(df_2, sheet_2, heading_2_param, fmt_head_2)
                style_of_design(df_2, sheet_2, heading_3_param, fmt_head_3)
                style_of_design(df_2, sheet_2, heading_4_param, fmt_head_4)
                # Премии
                style_of_design(df_2, sheet_2, heading_1_prime_2, fmt_prime_1)
                style_of_design(df_2, sheet_2, heading_2_prime_2, fmt_prime_2)

                # Закраска всех уникальных - для ПРЕДВАРИТЕЛЬНОГО РАСЧЕТА - Параметры
                filling_unique(df_1, sheet_1, format1, first_fmt, second_fmt)

                # Закраска всех уникальных - для ИТОГОВОГО РАСЧЕТА
                filling_unique(df_2, sheet_2, format1, first_fmt, second_fmt)

                # writer.save()


# Выставляем последовательность вывода данных
def priority_setting_export(list_1, list_2):
    boss = {}
    for idx, name in enumerate(Config.twig):
        boss[idx] = Config.twig[idx]

    exp_1 = []
    exp_2 = []

    for key, value in boss.items():
        for idx, name in enumerate(list_1):
            if name[0][1] == value:
                exp_1 += [name]

    for key, value in boss.items():
        for idx, name in enumerate(list_2):
            if name[0][1] == value:
                exp_2 += [name]

    return exp_1, exp_2


# Настройка формата
class FormatExport:
    def design(self, color, font=True, bold=True, size=9, wrap=False, valign='c', align='c'):
        """
        :param font:    # Шрифт (True = Times New Roman | False = None)
        :param color:   # Цвет заливки
        :param bold:    # Текст - Жирный (True = Жирный | False = Обычный)
        :param size:    # Размер шрифта (int)
        :param wrap:    # Перенос текста (True = Переносить | False = Нет)
        :param valign:  # Выравнивание текста по горизонтали ('c' = По центру | 'l' = Слева)
        :param align:   # Выравнивание текста по вертикали ('c' = По центру)
        """

        dict_format = {}

        # Шрифт
        if font:
            dict_format['font_name'] = 'Times New Roman'

        # Цвет заливки
        if color:
            dict_format['bg_color'] = color

        # Текст - Жирный
        if bold:
            dict_format['bold'] = 1

        # Размер шрифта
        if size:
            dict_format['font_size'] = int(size)

        # Перенос текста
        if wrap:
            dict_format['text_wrap'] = True

        # Выравнивание текста по горизонтали
        if valign:
            valign = str(valign).lower()
            if valign or valign == 'center' or valign == 'c':
                dict_format['valign'] = 'center'
            if valign == 'left' or valign == 'l':
                dict_format['valign'] = 'left'

        # Выравнивание текста по вертикали
        if align:
            align = str(align).lower()
            if align == 'center' or align == 'c' or align == 'vcenter' or align == 'vc':
                dict_format['align'] = 'vcenter'

        return dict_format


# Заливка уникальных ячеек
def filling_unique(df, sheet, fmt_cell, fmt_first, fmt_second):
    """
    :param df:          # DataFrame
    :param sheet:       # Название листа
    :param fmt_cell:    # Формат заливки - Уникальной ячейки
    :param fmt_first:   # Формат текста - Первого столбца
    :param fmt_second:  # Формат текста - Все, кроме первого столбца
    """

    num_line = 1    # Порядковый номер строки
    prob = []       # Список уникальных значений в строке

    for i in df.itertuples():
        for j in range(2, len(i)):
            if i[j] == i[j]:
                prob.append(str(i[j]).lower())
        cnt = Counter(prob).most_common(1)
        prob.clear()
        # Если cnt[..][1] == 1 это значит все уникальные переменные, то их и красить
        for j in range(2, len(i)):
            if str(i[j]).lower() != cnt[0][0] and cnt[0][1] != 1 and i[j] == i[j] and i[j] != '':
                # Эту ячейку нужно закрасить
                cell = xl_rowcol_to_cell(num_line - 1, j - 1)
                sheet.conditional_format(cell, {'type': 'no_errors',
                                                'format': fmt_cell})
            # Все уникальные - закрасить
            elif cnt[0][1] == 1 and i[j] == i[j] and i[j] != '':
                cell = xl_rowcol_to_cell(num_line - 1, j - 1)
                sheet.conditional_format(cell, {'type': 'no_errors',
                                                'format': fmt_cell})

        num_line += 1

    # Первый столбец применяет формат
    sheet.set_column(0, 0, 27.86, fmt_first)

    # Весь лист применяет формат
    sheet.set_column(1, len(df.columns) - 1, 27.86, fmt_second)


# Разбитие общего СПИСКА по ЗАГОЛОВКам
def splitting_by_titles(list_import, columns_export):
    """
    list_import     - Список - Входные значения. Структура: [[['',''],['','']]]
    index           - Список - Содержит координаты заголовков
    list_processing - Список - Временный список, участвующий между Импортом и Экспортом
    list_export     - Список - Разбитый список по заголовкам (delimiter)
    """

    # !!!!!!!!!!
    # !Внимание! - Если добавляется Xpath в парсинг - то исправлять ИНДЕКС заголовка !
    # !!!!!!!!!!
    delimiter = [
        'Клиент',
        'Статус клиента',
        'Личные данные',
        'Документы',
        'Паспорт',
        'Данные предыдущего паспорта',
        'Водительское удостоверение',
        'Предыдущее водительское удостоверение',
        'Контактные данные',
        'Адрес регистрации',
        'ШАГ 1. Выберите участника(ов) договора',
        'Водители',
        'Водитель №1',
        'Водитель №2',
        'ШАГ 2. Транспортное средство',
        'Транспортное средство',
        'ШАГ 3. Противоугонные системы',
        'Противоугонные системы',
        'ШАГ 4. Параметры договора страхования',
        'Параметры договора страхования',
        'ШАГ 5. Несчастный случай (НС)',
        'Несчастный случай (НС)',
        'Доп. продукты',
        'Премия, руб',
        'Страховая компания',
        'Базовые параметры',
        'Скоринг',
        'Условия страхования',
        'Лица, допущенные к управлению',
        'Дополнительные условия',
        'Пролонгация',
        'Франшиза',
        'Дополнительное оборудование',
        'Гражданская ответственность',
        'Несчастный случай'
    ]

    index, list_processing, list_export = ([] for i in range(3))

    # # Поиск Заголовков и определение их координат
    # for x in range(0, len(list_import)):
    #     # Выводится - по потокам
    #     for y in range(0, len(list_import[x])):
    #         for one_delimiter in delimiter:
    #             if list_import[x][y][0] == one_delimiter:
    #                 # Заполняем список - координатами заголовка
    #                 index.append([x, y, 0])
    #
    # # Соотношение индексов заголовков с основным списком и его разбив (срез)
    # for x in range(0, len(list_import)):
    #     for y in range(0, len(list_import[x])):
    #         # При переборе основного списка -> Смотрим все координаты и соотносим
    #         for number in range(0, len(index)):
    #             # Находим заголовки по полученным ранее координатам
    #             if x == int(index[number][0]) and y == int(index[number][1]):
    #                 # Срез с ЗАГОЛОВКА (включительно) и до следующего заголовка (НЕ включительно)
    #                 section = list_import[x][y:]
    #                 section[0][1] = f'{x}' if section[0][1] == '' else section[0][1]
    #
    #                 try:
    #                     next_y = int(index[number + 1][1])
    #                 except IndexError:
    #                     list_processing += [section]
    #                 else:
    #                     # Если следующий заголовок есть, то резать до него, иначе всё порезано
    #                     if next_y != 0:
    #                         step = next_y - int(index[number][1])
    #                         section = section[:step]
    #
    #                     list_processing += [section]
    
    column = 0
    good = pd.DataFrame()

    # # Рассортировка по блокам - разбитого списка
    # for number in range(0, len(delimiter)):
    #     for x in range(0, len(list_processing)):
    #         try:
    #             if list_processing[x][0][0] == delimiter[number]:
    #                 # Создаем список БЛОКА -> отправляем на merge (pandas)
    #                 list_export += [list_processing[x]]
    #         except IndexError:
    #             # иногда падает ошибка в list_processing[x][0][0] - Якобы переменная не существует
    #             pass
    #
    #     # Определение всего колич-ва столбцов по заголовку "Параметры"
    #     if number == 0:
    #         column = len(list_export)
    #
    #     # Добавление отсутствующих (пустых) столбцов в массив для их смещения при дальнейшем merge
    #     if len(list_export) != 0:
    #         for i in range(column):
    #             try:
    #                 # Если число и оно не равно номеру потока -> нужно добавить
    #                 if str(list_export[i][0][1]).isdigit() and int(list_export[i][0][1]) != i:
    #                     list_export.insert(i, [[list_export[i][0][0], '']])
    #                 elif str(list_export[i][0][1]).isdigit():
    #                     list_export[i][0][1] = ''
    #             except IndexError:
    #                 list_export.insert(i, [[list_export[i - 1][0][0], '']])
    #
    #     # Если Блок с названием Заголовка существует
    #     if len(list_export) != 0:
    #         # -> отправляем на merge (pandas) - после зачищаем
    #         panda_export = panda_param(list_export, columns_export)
    #
    #         # Словарь - Соединение/наращивание[good] объединенных блоков
    #         good = panda_export if number == 0 else pd.concat([good, panda_export])
    #
    #         list_export.clear()

    """
    Раскоментировать код выше и удалить код ниже
    """

    list_export = list_import
    # -> отправляем на merge (pandas) - после зачищаем
    panda_export = panda_param(list_export, columns_export)

    # Словарь - Соединение/наращивание[good] объединенных блоков
    good = pd.concat([good, panda_export])

    list_export.clear()

    return good


# Функция, чтобы убрать мусор
def panda_trash(value):
    value = value.split('_')[0]
    # value = value.replace('*', '')
    return value


# Создание DataFrame из Списков
def panda_param(doc, num_columns):
    """
    :param doc:         - Список list_export
    :param num_columns: - Количество столбцов
    :return:            - dataframe объединенных столбцов
    """
    try:
        # Если 2 столбца
        if num_columns == 2:
            good = pd.DataFrame(doc[0], columns=['name', '1'])

            for x in range(1, len(doc)):
                da = pd.DataFrame(doc[x], columns=['name', f'{x + 1}'])
                good = pd.merge(good, da, on='name', how='outer')

        # Если столбцов больше 2-х
        elif num_columns > 2:
            list_def = ['name' if num == 0 else f'{num}' for num in range(0, num_columns)]
            good = pd.DataFrame(doc[0], columns=list_def)

            for x in range(1, len(doc)):
                list_add = ['name' if num == 0 else f'{num * 10 + num}' for num in range(0, num_columns)]
                da = pd.DataFrame(doc[x], columns=list_add)
                good = pd.merge(good, da, on='name', how='outer')

        else:
            print(f'Ошибка: Объединение данных! Колич-во столбцов = 1 или 0')
            good = 'error'

    except ValueError:
        # Если какой-то из потоков не завершился, то выскакивает ошибка, что столбцов 0, а нужно 2
        pass

    else:
        return good


# Заливка заголовков
def style_of_design(df, sheet, heading, fmt_head):
    """
    :param df:          # DataFrame
    :param sheet:       # Название листа
    :param heading:     # Название параметров
    :param fmt_head:    # Стиль
    """

    check = False
    for i in range(0, df.shape[0]):
        if not check:
            for j in range(0, len(heading)):
                if df.iat[i, 0] == heading[j]:
                    sheet.conditional_format(i, 0, i, df.shape[1] - 1, {'type': 'no_errors',
                                                                        'format': fmt_head})
                    # Затереть старое значение (исключает закраску повторяющихся заголовков)
                    # heading[j] = 'xxx'
                    break
        else:
            break


# def main():
#     export_to_excel([], [])
#
#
# if __name__ == '__main__':
#     main()
