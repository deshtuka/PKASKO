# -*- coding: utf-8 -*-
from settings import Config
from PKASKO_scraping import scraping
from PKASKO_SQL import db_pkasko
from PKASKO_export_to_excel import export_to_excel

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import ElementNotInteractableException
import pandas as pd
import threading
from datetime import date
import time
import os
from traceback import print_exc
from pyinstrument import Profiler


conf_dict = dict()      # dict - Содержит параметры сессии {Адрес сайта, СК, Тип расчета}
log_out = ''            # str - Путь для вывода всего (логи, скрины...)
import_dict = dict()    # dict - Словарь инпортированных данных из БД | Структура: {0: Параметры кейса, 1: Человек #1, 2: Человек #2}
# list_import = list()    # list - Список инпортированных данных из Excel
list_export = list()    # list - Список экспортируемых данных в Excel


class PkaskoXpath:
    XPATH_SIGNIN = '//a[@data-target="#modalSignIn"]'
    XPATH_LOGIN = '//form[@id="authForm"]//input[@name="login"]'
    XPATH_PASSOWRD = '//form[@id="authForm"]//input[@name="password"]'
    XPATH_GO_TO = '//button[contains(text(), "Войти")]'

    XPATH_LOADER_CONTENT = '//*[@class="loader-content"]'
    XPATH_MAIN_LOADING = '//*[@ng-if="mainLoading"]'
    XPATH_BAR = '//*[@class="bar"]'
    XPATH_SELECT_DROP = '//*[@id="select2-drop"]'
    XPATH_SELECT_INPUT = '//div[@id="select2-drop"]//input'
    XPATH_SELECT_CLICK_ONE = '//div[@id="select2-drop"]//li[1]/div'
    XPATH_SELECT_CLICK_TO_ONE = '//*[@id="select2-drop"]/ul/li[1]/div'
    XPATH_SELECT_CLICK_TWO = '//div[@id="select2-drop"]//li[2]/div'
    XPATH_SELECT_CLICK_TO_TWO = '//*[@id="select2-drop"]/ul/li[2]/div'
    XPATH_HIDE_CALENDAR = '//*[@class="text-nowrap d-flex"]/h4'
    XPATH_HIDE_CALENDAR_DATE = '//div[@class="modal-header ng-scope"]'
    XPATH_CLIENT_SAVE = '//a[@ng-disabled="!canSave()"]'
    XPATH_KASKO_CALC_BUTTON = '//*[@id="kaskoCalcButton"]'
    XPATH_GO_TO_RIGHT = '//*[@class="kasko-result"]//div[@class="ng-binding"]'
    XPATH_CALC = '//*[@class="icon-p-arrow icon-p-arrow-style"]'
    XPATH_RESULT_LAST = '//*[@class="kasko-result"]//tbody/tr'

    XPATH_FULLNAME = '//*[@ng-model="kaskoDriver.fullName"]'
    XPATH_BIRTHDATE = '//*[@ng-model="kaskoDriver.birthdate"]'
    XPATH_LICENSE_DATE = '//*[@ng-model="kaskoDriver.license.date"]'
    XPATH_LICENSE_SERIES = '//*[@ng-model="kaskoDriver.license.series"]'
    XPATH_LICENSE_NUMBER = '//*[@ng-model="kaskoDriver.license.number"]'
    XPATH_SEX = '//*[@ng-model="kaskoDriver.sex" and @btn-radio="\'w\'"]'
    XPATH_MARRIAGE = '//*[@ng-model="kaskoDriver.marriage"]'
    XPATH_CHILDREN = '//*[@ng-model="kaskoDriver.children"]'
    XPATH_ADD_DRIVER = '//*[@id="kaskoDriverLinkLayer"]//a'

    XPATH_GO_TO_KASKO_EXACT = '//*[@id="mainContent"]//*[@ui-sref="calc.kaskoExact"]'
    XPATH_EXACT_ADD_DRIVER = '//*[@id="kaskoDriverParentLayer"]//*[@ng-click="addDriver()"]'
    XPATH_STATUS_INSURANT = '//*[@for="insurant"]'
    XPATH_STATUS_OWNER = '//*[@for="owner"]'
    XPATH_STATUS_DRIVER = '//*[@for="driver"]'
    XPATH_STATUS_BENEFICIARY = '//*[@for="beneficiary"]'
    XPATH_DATA_LASTNAME = '//*[@ng-model="client.data.lastname"]'
    XPATH_DATA_FIRSTNAME = '//*[@ng-model="client.data.firstname"]'
    XPATH_DATA_MIDDLENAME = '//*[@ng-model="client.data.middlename"]'
    XPATH_DATA_BIRTHDATE = '//*[@ng-model="client.data.birthdate"]'
    XPATH_STATUS_SEX = '//label[@class="span3"]//*[@class="select2-arrow"]'
    XPATH_STATUS_SEX_WOMEN = '//*[contains(text(), "Женский")]/..//div[@class="select2-result-label"]'
    XPATH_STATUS_SEX_MEN = '//*[contains(text(), "Мужской")]/../div[@class="select2-result-label"]'
    XPATH_STATUS_MARRIAGE = '//*[contains(text(), "В браке? ")]//*[@class="select2-arrow"]'
    XPATH_STATUS_MARRIAGE_YES = '//*[contains(text(), "Да")]/../div[@class="select2-result-label"]'
    XPATH_STATUS_MARRIAGE_NO = '//*[contains(text(), "Нет")]/../div[@class="select2-result-label"]'
    XPATH_STATUS_CHILDREN = '//*[contains(text(), "Наличие детей? ")]//*[@class="select2-arrow"]'
    XPATH_STATUS_CHILDREN_YES = '//*[contains(text(), "Да")]/../div[@class="select2-result-label"]'
    XPATH_STATUS_CHILDREN_NO = '//*[contains(text(), "Нет")]/../div[@class="select2-result-label"]'
    XPATH_PASSPORT_SERIES = '//*[@ng-model="client.data.data.documents.passport.series"]'
    XPATH_PASSPORT_DATE = '//*[@ng-model="client.data.data.documents.passport.date"]'
    XPATH_PASSPORT_ISSUED = '//*[@ng-model="client.data.data.documents.passport.issued"]'
    XPATH_DATA_LICENSE_SERIES = '//*[@ng-model="client.data.data.documents.license.series"]'
    XPATH_DATA_LICENSE_NUMBER = '//*[@ng-model="client.data.data.documents.license.number"]'
    XPATH_DATA_LICENSE_DATELAST = '//*[@ng-model="client.data.data.documents.license.dateLast"]'
    XPATH_DATA_LICENSE_DATE = '//*[@ng-model="client.data.data.documents.license.date"]'
    XPATH_DATA_REGION = '//*[contains(text(), "Регион")]//*[@class="select2-arrow"]'
    XPATH_LOCALITY = '//*[contains(text(), "Населённый пункт")]//*[@class="select2-arrow"]'
    XPATH_STREET = '//*[contains(text(), "Улица")]//*[@class="select2-arrow"]'
    XPATH_HOME = '//*[@ng-model="client.data.data.address.house"]'
    XPATH_BUILDING = '//*[@ng-model="client.data.data.address.building"]'
    XPATH_APARTMENT = '//*[@ng-model="client.data.data.address.appartment"]'
    XPATH_DATA_VIN = '//label[contains(text(), "VIN")]'
    XPATH_DATA_VIN_INPUT = f'{XPATH_DATA_VIN}/input'
    XPATH_SIGN_NUMBER = '//*[@ng-model="client.data.data.car.signNumber"]'
    XPATH_PTS_TYPE = '//*[contains(text(), "Тип ПТС")]//*[@class="select2-arrow"]'
    XPATH_PTS_SERIES = '//*[@ng-model="client.data.data.car.seriesPTS"]'
    XPATH_PTS_NUMBER = '//*[@ng-model="client.data.data.car.numberPTS"]'
    XPATH_PTS_DATE = '//*[@ng-model="client.data.data.car.datePTS"]'

    XPATH_MAKE = '//*[contains(text(), "Марка")]/..//*[@class="select2-arrow"]'
    XPATH_MODEL = '//*[contains(text(), "Модель")]/..//*[@class="select2-arrow"]'
    XPATH_POWER = '//*[@ng-model="model.kasko.power"]'
    XPATH_YEAR = '//*[contains(text(), "Год выпуска")]/..//*[@class="select2-arrow"]'
    XPATH_PRICE = '//*[@ng-model="model.kasko.price"]'
    XPATH_CREDIT = '//*[contains(text(), "Кредитное ТС")]/..//*[@class="select2-arrow"]'
    XPATH_MILEAGE = '//*[@ng-model="model.kasko.mileage"]'
    XPATH_EXPLDATE = '//*[@ng-model="model.extended.data[\'explDate\']"]'
    XPATH_NEW_TS = '//*[contains(text(), "Новое ТС")]/..//*[@class="select2-arrow"]'
    XPATH_VIN = '//*[@ng-model="model.kasko.VIN"]'


# Основное тело всего кода
def work():

    # # Инициализация
    # profiler = Profiler()
    # profiler.start()

    # Создание папок
    creating_folders()

    # Авторизация
    browser = authorization()

    # Тело расчетов
    good = body(browser)

    # Переход в правую часть
    good.extend(body_right_side(browser))

    global list_export
    list_export.append(good)

    # Закрытие браузера
    exit_code(browser)

    # # Завершение - Профилировки
    # profiler.stop()
    # print(profiler.output_text(unicode=True, color=True))


# Тело всех расчетов
def body(browser):
    calc_now = int(threading.currentThread().getName())
    good = []

    try:
        # Вкладка "Расчет" - Клик
        check_on_loading_banner(browser)

        browser.find_element_by_id('calcTab').click()
        check_on_loading_page(browser)

        try:
            # КАСКО - предварительный
            if calc_now == 0:
                good = calc_kasko_preliminary(browser)

            # КАСКО - итоговый
            elif calc_now == 1:
                good = calc_kasko_exact(browser)

        except:
            print(f'ОШИБКА: [Noname] -> {Config.twig[calc_now]}!')
            print_exc()
            browser.save_screenshot(f'{log_out}ОШИБКА [Noname] - {Config.twig[calc_now]}.png')

        else:
            time.sleep(0.2)
            # Прыжок к последней записе расчета для скрина
            try:
                calc_elements = browser.find_elements_by_xpath(PkaskoXpath.XPATH_RESULT_LAST)
                link = calc_elements[len(calc_elements) - 1]
                ActionChains(browser).move_to_element(link).perform()

                browser.find_element_by_tag_name('body').send_keys(Keys.DOWN)
                time.sleep(0.2)

                browser.save_screenshot(f'{log_out}Премии - {Config.twig[calc_now]}.png')
            except IndexError:
                browser.save_screenshot(f'{log_out}Ошибка в расчете - премия отсутствует - {Config.twig[calc_now]}.png')
                print_exc()

    except:
        browser.save_screenshot(f'{log_out}Ошибка неизвестная - {Config.twig[calc_now]}.png')
        print_exc()

    finally:
        return good


# КАСКО - Предварительный расчет
def calc_kasko_preliminary(browser):
    print(f'*** КАСКО - Предварительный расчет ***')

    # Новый водитель - короткая запись
    def new_short_driver(browser_sd):
        print(f'----- Заполняю данные -> [Водитель] -----')

        try:
            # Определение кто водитель: страховщик или др.водитель
            driver_1 = import_dict.get(1).get('status')
            driver_2 = False if import_dict.get(2) is None else import_dict.get(2).get('status')

            # repeat = {1: Водитель 1 | 2: Водитель 2 | 3: Оба водители}
            repeat = 3 if all([driver_1, driver_2]) else 1 if all([driver_1, not driver_2]) else 2

            for num in range(1, repeat + 1):
                # Пауза на первое прохождение цикла
                if num == 1 and repeat == 2:
                    continue
                if num == 3 and repeat == 3:
                    break

                driver = import_dict.get(num)

                # ФИО
                value = f'{driver.get("lastname")} {driver.get("firstname")} {driver.get("middlename", "")}'
                links = browser_sd.find_elements_by_xpath(PkaskoXpath.XPATH_FULLNAME)
                links[0].send_keys(value) if repeat in [1, 2] else links[num - 1].send_keys(value)

                # Дата рождения
                value = str(driver.get("birthdate").strftime("%d.%m.%Y"))
                links = browser_sd.find_elements_by_xpath(PkaskoXpath.XPATH_BIRTHDATE)
                links[0].send_keys(value) if repeat in [1, 2] else links[num - 1].send_keys(value)

                # Закрыть всплывающий календарь, который мешает
                browser_sd.find_elements_by_xpath(PkaskoXpath.XPATH_HIDE_CALENDAR)[0].click()

                # Дата начала стажа
                value = str(driver.get("licenseDate").strftime("%d.%m.%Y"))
                links = browser_sd.find_elements_by_xpath(PkaskoXpath.XPATH_LICENSE_DATE)
                links[0].send_keys(value) if repeat in [1, 2] else links[num - 1].send_keys(value)

                # Закрыть всплывающий календарь, который мешает
                browser_sd.find_elements_by_xpath(PkaskoXpath.XPATH_HIDE_CALENDAR)[0].click()

                # ВУ - Серия
                value = driver.get("licenseSeries")
                links = browser_sd.find_elements_by_xpath(PkaskoXpath.XPATH_LICENSE_SERIES)
                links[0].send_keys(value) if repeat in [1, 2] else links[num - 1].send_keys(value)

                # ВУ - Номер
                value = driver.get("licenseNumber")
                links = browser_sd.find_elements_by_xpath(PkaskoXpath.XPATH_LICENSE_NUMBER)
                links[0].send_keys(value) if repeat in [1, 2] else links[num - 1].send_keys(value)

                # Пол
                if len([True for one in ['ж', 'w'] if str(driver.get("sex")[:1]).lower() == one]) > 0:
                    links = browser_sd.find_elements_by_xpath(PkaskoXpath.XPATH_SEX)
                    links[0].click() if repeat in [1, 2] else links[num - 1].click()

                # Брак
                if int(driver.get('marriage')) == 1:
                    links = browser_sd.find_elements_by_xpath(PkaskoXpath.XPATH_MARRIAGE)
                    links[0].click() if repeat in [1, 2] else links[num - 1].click()

                # Дети
                if int(driver.get('children')) == 1:
                    links = browser_sd.find_elements_by_xpath(PkaskoXpath.XPATH_CHILDREN)
                    links[0].click() if repeat in [1, 2] else links[num - 1].click()

                # Добавить второго водителя
                if repeat == 3 and num == 1:
                    browser_sd.find_element_by_xpath(PkaskoXpath.XPATH_ADD_DRIVER).click()
                    driver.clear()

        except:
            print(f'Ошибка: [КАСКО - Предварительный расчет] - заполнение водителя!')
            print_exc()
            browser_sd.save_screenshot(f'{log_out}Ошибка - Заполнение водителя.png')

    # Заполнение водителя - короткая запись
    new_short_driver(browser)

    # Заполняем - "Транспортное средство"
    new_auto(browser)

    # Заполняем - "Параметры договора страхования"
    new_param_doc(browser)

    # Рассчитать - клик
    list_param = calc_kasko_button_click(browser)

    return list_param


# КАСКО - Точный расчет
def calc_kasko_exact(browser):
    print(f'*** КАСКО - Точный расчет ***')
    browser.find_element_by_xpath(PkaskoXpath.XPATH_GO_TO_KASKO_EXACT).click()

    # Новый клиент - создание
    def new_user(browser_user):
        client_user = list()
        check_on_loading_page(browser_user)

        # Распределение ролей | кто страховщик, водитель, выгодоприобретатель
        status = {
            0: import_dict.get(0).get('status'),  # Выгодоприобретатель: True - если Юр.лицо
            1: import_dict.get(1).get('status'),
            2: False if import_dict.get(2) is None else import_dict.get(2).get('status')
        }

        for num in range(1, 4):
            driver = import_dict.get(num) if num in [1, 2] else import_dict.get(0)
            """
            num = 1 -> Человек 1
            num = 2 -> Человек 2
            num = 3 -> Выгодоприобретатель
            """

            # todo Временно отключен Выгодоприобретатель
            if driver is not None and num != 3:
                if num != 3:
                    print(f'----- Заполняю данные -> [Клиент #{num}: {driver.get("lastname")} {driver.get("firstname")}] -----')

                browser_user.find_element_by_xpath(PkaskoXpath.XPATH_EXACT_ADD_DRIVER).click()

                time.sleep(0.5)

                if num == 1:
                    # Страхователь
                    browser_user.find_element_by_xpath(PkaskoXpath.XPATH_STATUS_INSURANT).click()
                    # Собственник
                    browser_user.find_element_by_xpath(PkaskoXpath.XPATH_STATUS_OWNER).click()

                # Водитель
                if num in [1, 2]:
                    if status.get(num):
                        browser_user.find_element_by_xpath(PkaskoXpath.XPATH_STATUS_DRIVER).click()

                # Выгодоприобретатель
                if num in [1, 3]:
                    if (status.get(0) and num == 3) or (status.get(1) and num == 1):
                        browser_user.find_element_by_xpath(PkaskoXpath.XPATH_STATUS_BENEFICIARY).click()

                time.sleep(0.2)

                if num in [1, 2]:
                    # Фамилия
                    browser_user.find_element_by_xpath(PkaskoXpath.XPATH_DATA_LASTNAME).send_keys(driver.get('lastname'))

                    # Имя
                    browser_user.find_element_by_xpath(PkaskoXpath.XPATH_DATA_FIRSTNAME).send_keys(driver.get('firstname'))

                    # Отчество
                    browser_user.find_element_by_xpath(PkaskoXpath.XPATH_DATA_MIDDLENAME).send_keys(driver.get('middlename'))

                    # Дата рождения
                    browser_user.find_element_by_xpath(PkaskoXpath.XPATH_DATA_BIRTHDATE).send_keys(str(driver.get('birthdate').strftime("%d.%m.%Y")))
                    # data_click - нужен, чтобы закрыть всплывающий календарь, который мешает в дальнейшем
                    data_click = browser_user.find_element_by_xpath(PkaskoXpath.XPATH_HIDE_CALENDAR_DATE)
                    data_click.click()

                    # Пол
                    browser_user.find_element_by_xpath(PkaskoXpath.XPATH_STATUS_SEX).click()
                    if len([True for one in ['ж', 'w'] if str(driver.get("sex")[:1]).lower() == one]) > 0:
                        browser_user.find_element_by_xpath(PkaskoXpath.XPATH_STATUS_SEX_WOMEN).click()
                    else:
                        browser_user.find_element_by_xpath(PkaskoXpath.XPATH_STATUS_SEX_MEN).click()

                    # Брак
                    browser_user.find_element_by_xpath(PkaskoXpath.XPATH_STATUS_MARRIAGE).click()
                    if int(driver.get('marriage')) == 1:
                        # Брак - ДА
                        browser_user.find_element_by_xpath(PkaskoXpath.XPATH_STATUS_MARRIAGE_YES).click()
                    else:
                        # Брак - НЕТ
                        browser_user.find_element_by_xpath(PkaskoXpath.XPATH_STATUS_MARRIAGE_NO).click()

                    # Дети
                    browser_user.find_element_by_xpath(PkaskoXpath.XPATH_STATUS_CHILDREN).click()
                    if int(driver.get('children')) == 1:
                        # Дети - ДА
                        browser_user.find_element_by_xpath(PkaskoXpath.XPATH_STATUS_CHILDREN_YES).click()
                    else:
                        # Дети - НЕТ
                        browser_user.find_element_by_xpath(PkaskoXpath.XPATH_STATUS_CHILDREN_NO).click()

                    # У Водителя - нет паспорта
                    if num != 2:
                        # --- Паспорт ---> Серия
                        browser_user.find_element_by_xpath().send_keys(driver.get('passportSeries'))

                        # --- Паспорт ---> Номер
                        browser_user.find_element_by_xpath(PkaskoXpath.XPATH_PASSPORT_SERIES).send_keys(driver.get('passportNumber'))

                        # --- Паспорт ---> Дата выдачи
                        data = browser_user.find_element_by_xpath(PkaskoXpath.XPATH_PASSPORT_DATE)
                        data.send_keys(str(driver.get('passportDate').strftime("%d.%m.%Y")))
                        data_click.click()

                        # --- Паспорт ---> Кем выдан
                        browser_user.find_element_by_xpath(PkaskoXpath.XPATH_PASSPORT_ISSUED).send_keys('Отделом УФМС России')

                    # Если статус водителя == True
                    if status[num]:
                        # --- Водительское ---> Серия
                        az = browser_user.find_element_by_xpath(PkaskoXpath.XPATH_DATA_LICENSE_SERIES)
                        ActionChains(browser_user).move_to_element(az).perform()
                        az.send_keys(driver.get('licenseSeries'))

                        # --- Водительское ---> Номер
                        browser_user.find_element_by_xpath(PkaskoXpath.XPATH_DATA_LICENSE_NUMBER).send_keys(driver.get('licenseNumber'))

                        # --- Водительское ---> Дата выдачи
                        data = browser_user.find_element_by_xpath(PkaskoXpath.XPATH_DATA_LICENSE_DATELAST)
                        data.send_keys('01.01.2021')
                        data_click.click()

                        # --- Водительское ---> Дата стажа
                        data = browser_user.find_element_by_xpath(PkaskoXpath.XPATH_DATA_LICENSE_DATE)
                        data.send_keys(str(driver.get('licenseDate').strftime("%d.%m.%Y")))
                        data_click.click()

                    # У Водителя - нет регистрации
                    if num != 2:
                        # --- Регистрация ---> Регион
                        try:
                            az = browser_user.find_element_by_xpath(PkaskoXpath.XPATH_DATA_REGION)
                            ActionChains(browser_user).move_to_element(az).click(az).perform()
                        except:
                            print('Ошибка -> Регистрация - Регион')
                        else:
                            browser_user.find_element_by_xpath(PkaskoXpath.XPATH_SELECT_INPUT).send_keys(driver.get('addressRegion'))
                            browser_user.find_element_by_xpath(PkaskoXpath.XPATH_SELECT_CLICK_ONE).click()

                        # --- Регистрация ---> Населенный пункт
                        try:
                            az = browser_user.find_element_by_xpath(PkaskoXpath.XPATH_LOCALITY)
                            ActionChains(browser_user).move_to_element(az).click(az).perform()
                        except:
                            print('Ошибка -> Регистрация - Населенный пункт')
                        else:
                            browser_user.find_element_by_xpath(PkaskoXpath.XPATH_SELECT_INPUT).send_keys(driver.get('addressRegion'))
                            browser_user.find_element_by_xpath(PkaskoXpath.XPATH_SELECT_CLICK_ONE).click()

                        # --- Регистрация ---> Улица
                        try:
                            az = browser_user.find_element_by_xpath(PkaskoXpath.XPATH_STREET)
                            az.click()
                        except:
                            print('Ошибка -> Регистрация - Улица')
                        else:
                            browser_user.find_element_by_xpath(PkaskoXpath.XPATH_SELECT_INPUT).send_keys(driver.get('addressStreet'))
                            browser_user.find_element_by_xpath(PkaskoXpath.XPATH_SELECT_CLICK_ONE).click()

                        # --- Регистрация ---> Дом
                        browser_user.find_element_by_xpath(PkaskoXpath.XPATH_HOME).send_keys(driver.get('addressHouse'))

                        # --- Регистрация ---> Корпус
                        browser_user.find_element_by_xpath(PkaskoXpath.XPATH_BUILDING).send_keys(driver.get('addressBuilding'))

                        # --- Регистрация ---> Квартира
                        browser_user.find_element_by_xpath(PkaskoXpath.XPATH_APARTMENT).send_keys(driver.get('addressAppartment'))

                    # Второму персонажу не нужно вносить данные ТС, чтобы не нарушать целлостность данных в дальнейшем
                    if num != 2:
                        # --- ТС ---> VIN
                        if import_dict.get(0).get('VIN') is not None:
                            if len([True for one in ['нет', 'н', 'no', 'n'] if str(import_dict.get(0).get('VIN')).lower() == one]) == 0:
                                az = browser_user.find_element_by_xpath(PkaskoXpath.XPATH_DATA_VIN)
                                ActionChains(browser_user).move_to_element(az).perform()
                                browser_user.find_element_by_xpath(PkaskoXpath.XPATH_DATA_VIN_INPUT).send_keys(import_dict.get(0).get('VIN'))

                        # --- ТС ---> Гос.номер
                        if import_dict.get(0).get('bodyNumber') is not None:
                            if len([True for one in ['нет', 'н', 'no', 'n'] if str(import_dict.get(0).get('bodyNumber')).lower() == one]) == 0:
                                browser_user.find_element_by_xpath(PkaskoXpath.XPATH_SIGN_NUMBER).send_keys(import_dict.get(0).get('bodyNumber'))

                        # --- ПТС ---> Тип ПТС
                        az = browser_user.find_element_by_xpath(PkaskoXpath.XPATH_PTS_TYPE)
                        ActionChains(browser_user).move_to_element(az).click().perform()
                        # --- ПТС ---> Тип ПТС = ПТС
                        if import_dict.get(0).get('seriesPTS') is not None:
                            if len([True for one in ['нет', 'н', 'no', 'n'] if str(import_dict.get(0).get('seriesPTS')).lower() == one]) == 0:
                                browser_user.find_element_by_xpath(PkaskoXpath.XPATH_SELECT_CLICK_ONE).click()
                                # --- ПТС ---> Серия
                                browser_user.find_element_by_xpath(PkaskoXpath.XPATH_PTS_SERIES).send_keys(import_dict.get(0).get('seriesPTS'))

                            # --- ПТС ---> Тип ПТС = е-ПТС
                            else:
                                browser_user.find_element_by_xpath(PkaskoXpath.XPATH_SELECT_CLICK_TWO).click()

                        # --- ПТС ---> Номер
                        browser_user.find_element_by_xpath(PkaskoXpath.XPATH_PTS_NUMBER).send_keys(import_dict.get(0).get('numberPTS'))

                        # --- ПТС ---> Дата ПТС
                        data = browser_user.find_element_by_xpath(PkaskoXpath.XPATH_PTS_DATE)
                        data.send_keys(str(import_dict.get(0).get('datePTS').strftime("%d.%m.%Y")))
                        data_click.click()

                client_user.extend(scraping(browser, block=1))
                print(f'client_user = {client_user}')

                browser.save_screenshot(f'{log_out}{Config.twig[1]} - Данные клиента #{num} - {driver.get("lastname")} {driver.get("firstname")}.png')

                # Сохранить
                try:
                    az = browser_user.find_element_by_xpath(PkaskoXpath.XPATH_CLIENT_SAVE)
                    ActionChains(browser_user).move_to_element(az).click().perform()
                    browser_user.find_element_by_tag_name('body').send_keys(Keys.HOME)
                except:
                    print('Ошибка - Кнопка "Сохранить" - Не активна!')
                    print_exc()
                else:
                    print('Клиент - Сохранен!')

                    # todo Это ВРЕМЕННО! Убрать когда решится проблема с Выгодоприобретателем
                    if num == 2:
                        break

        return client_user

    # Создание нового пользователя
    list_param = new_user(browser)

    # Заполняем - "Транспортное средство"
    new_auto(browser)

    # Заполняем - "Параметры договора страхования"
    new_param_doc(browser)

    # Рассчитать - клик
    list_param.extend(calc_kasko_button_click(browser))

    return list_param


# Транспортное средство - заполнение данных
def new_auto(browser):
    print(f'----- Заполняю данные -> [Транспортное средство] -----')

    time.sleep(1)
    check_on_loading_page(browser)

    # Марка
    az = browser.find_element_by_xpath(PkaskoXpath.XPATH_MAKE)
    ActionChains(browser).move_to_element(az).click().perform()
    check_drop = check_on_loading_select2_drop(browser, error='ТС - Марка')
    if check_drop:
        browser.find_element_by_xpath(PkaskoXpath.XPATH_SELECT_INPUT).send_keys(import_dict.get(0).get('make'))
        browser.find_element_by_xpath(PkaskoXpath.XPATH_SELECT_CLICK_ONE).click()

    # Модель
    az = browser.find_element_by_xpath(PkaskoXpath.XPATH_MODEL)
    ActionChains(browser).move_to_element(az).click().perform()
    check_drop = check_on_loading_select2_drop(browser, error='ТС - Модель')
    if check_drop:
        browser.find_element_by_xpath(PkaskoXpath.XPATH_SELECT_INPUT).send_keys(import_dict.get(0).get('model'))
        browser.find_element_by_xpath(PkaskoXpath.XPATH_SELECT_CLICK_TO_ONE).click()

    # Мощность
    browser.find_element_by_xpath(PkaskoXpath.XPATH_POWER).send_keys(import_dict.get(0).get('power'))

    # Год выпуска
    az = browser.find_element_by_xpath(PkaskoXpath.XPATH_YEAR)
    ActionChains(browser).move_to_element(az).click().perform()
    check_drop = check_on_loading_select2_drop(browser, error='ТС - Год выпуска')
    if check_drop:
        browser.find_element_by_xpath(PkaskoXpath.XPATH_SELECT_INPUT).send_keys(import_dict.get(0).get('year'))
        browser.find_element_by_xpath(PkaskoXpath.XPATH_SELECT_CLICK_TO_ONE).click()

    # Страховая сумма
    browser.find_element_by_xpath(PkaskoXpath.XPATH_PRICE).send_keys(import_dict.get(0).get('price'))

    # Кредитное ТС - НЕТ
    az = browser.find_element_by_xpath(PkaskoXpath.XPATH_CREDIT)
    ActionChains(browser).move_to_element(az).click().perform()
    check_drop = check_on_loading_select2_drop(browser, error='ТС - Кредит')
    if check_drop:
        browser.find_element_by_xpath(PkaskoXpath.XPATH_SELECT_CLICK_TO_ONE).click()

    # Пробег
    browser.find_element_by_xpath(PkaskoXpath.XPATH_MILEAGE).send_keys(import_dict.get(0).get('mileage'))

    # Дата начала эксплуатации
    browser.find_element_by_xpath(PkaskoXpath.XPATH_EXPLDATE).clear()
    browser.find_element_by_xpath(PkaskoXpath.XPATH_EXPLDATE).send_keys(str(import_dict.get(0).get('datePTS').strftime("%d.%m.%Y")))
    # Закрыть всплывающий календарь, который мешает
    browser.find_elements_by_xpath(PkaskoXpath.XPATH_HIDE_CALENDAR)[1].click()

    # Новое ТС
    az = browser.find_element_by_xpath(PkaskoXpath.XPATH_NEW_TS)
    ActionChains(browser).move_to_element(az).click().perform()
    if int(import_dict.get(0).get('year')) != int(date.today().strftime('%Y')):
        browser.find_element_by_xpath(PkaskoXpath.XPATH_SELECT_CLICK_TO_ONE).click()
    else:
        browser.find_element_by_xpath(PkaskoXpath.XPATH_SELECT_CLICK_TO_TWO).click()

    # VIN
    if import_dict.get(0).get('VIN') is not None:
        if len([True for one in ['нет', 'н', 'no', 'n'] if str(import_dict.get(0).get('VIN')).lower() == one]) == 0:
            link = browser.find_element_by_xpath(PkaskoXpath.XPATH_VIN)
            vin = link.get_attribute('value')
            if (vin and vin == '') or not vin:
                link.send_keys(import_dict.get(0).get('VIN'))


# Параметры договора страхования - заполнение данных
def new_param_doc(browser):
    print(f'----- Заполняю данные -> [Параметры договора страхования] -----')

    # Способ возмещения убытка
    az = browser.find_element_by_xpath('//*[contains(text(), "Способ возмещения убытка")]/..//*[@class="select2-arrow"]')
    ActionChains(browser).move_to_element(az).click().perform()
    check_drop = check_on_loading_select2_drop(browser, error='Способ возмещения убытка')
    if check_drop:
        browser.find_element_by_xpath(PkaskoXpath.XPATH_SELECT_CLICK_TO_ONE).click()


# Финиш КАСКО - "Рассчитать стоимость полиса" - Клик
def calc_kasko_button_click(browser):
    calc_now = int(threading.currentThread().getName())
    start_time_local = time.time()
    browser.save_screenshot(f'{log_out}{Config.twig[calc_now]} - Заполненные данные.png')

    # Скролл
    az = browser.find_element_by_xpath(PkaskoXpath.XPATH_KASKO_CALC_BUTTON)
    ActionChains(browser).move_to_element(az).click().perform()

    print(f'"Рассчитать стоимость полиса" - Клик!')
    time.sleep(0.5)

    list_param = []
    list_param.extend(scraping(browser, block=2))

    # Ожидание загрузки расчета
    check_on_loading_calc(browser)

    list_param.extend(scraping(browser, block=4))

    minutes, seconds = timer(start_time_local, info=False)
    print(f'"Рассчитать стоимость полиса" - Загрузился! [Время расчета: {minutes}m {seconds}s]')

    return list_param


# Переход в правую часть
def body_right_side(browser):
    try:
        start_time_local = time.time()

        # Переход в правую часть согласно выбранной опции
        links = browser.find_elements_by_xpath(PkaskoXpath.XPATH_GO_TO_RIGHT)
        for idx, link in enumerate(links):
            if conf_dict.get('ic') == str(link.get_attribute('innerText')).strip():
                pass

        browser.find_elements_by_xpath(PkaskoXpath.XPATH_CALC)[3].click()
        print(f'"Расчитать правую часть" - Клик!')

        # Ожидание загрузки спинера
        check_on_loading_page(browser)

        browser.find_element_by_tag_name('body').send_keys(Keys.HOME)
        time.sleep(0.1)

        minutes, seconds = timer(start_time_local, info=False)
        print(f'"Расчет в правой части"       - Загрузился! [Время расчета: {minutes}m {seconds}s]')

    except Exception as ex:
        print(f'Ошибка при парсинге')
        browser.save_screenshot(f'Ошибка при парсинге.png')
        print_exc()
        return [['Внутренняя ошибка скрипта', ex]]

    else:
        start_time_pars = time.time()

        list_param = []
        list_param.extend(scraping(browser, block=5))
        list_param.extend(scraping(browser, block=6))

        timer(start_time_pars, info='Время парсинга Правой части')
        return list_param


# Проверка на наличие спинера ЗАГРУЗКИ - БОЛЬШОЙ (всей старницы)
def check_on_loading_page(browser):
    browser.implicitly_wait(2)
    while True:
        try:
            browser.find_element_by_xpath(PkaskoXpath.XPATH_LOADER_CONTENT)
        except NoSuchElementException:
            # Если Спинер не грузится - выйти
            browser.implicitly_wait(15)
            break
        else:
            # Если Спинер грузится - то замереть
            time.sleep(1)


def check_on_loading_banner(browser):
    browser.implicitly_wait(2)
    while True:
        try:
            browser.find_element_by_xpath(PkaskoXpath.XPATH_MAIN_LOADING)
        except NoSuchElementException:
            # Если Спинер не грузится - выйти
            time.sleep(1)
            browser.implicitly_wait(15)
            break
        else:
            # Если Спинер грузится - то замереть
            time.sleep(1)


# Проверка на наличие спинера ЗАГРУЗКИ - при расчете
def check_on_loading_calc(browser):
    browser.implicitly_wait(1)
    while True:
        try:
            browser.find_element_by_xpath(PkaskoXpath.XPATH_BAR)
            time.sleep(0.5)
        except NoSuchElementException:
            browser.implicitly_wait(15)
            break


# Проверка на отображение окна select'а
def check_on_loading_select2_drop(browser, error):
    result = False
    time_check = 0
    try:
        browser.implicitly_wait(1)
        while True:
            container = browser.find_element_by_xpath(PkaskoXpath.XPATH_SELECT_DROP)
            block = browser.execute_script('return arguments[0].style.display;', container)
            if block == 'block':
                result = True
                break
            if time_check == 8:
                break
            else:
                time_check += 1
                time.sleep(0.5)
    except NoSuchElementException:
        print(f'Ошибка: Select [{error}] - Не раскрылся!')
        browser.save_screenshot(f'{log_out}Ошибка - select [{error}] - не раскрылся.png')
    finally:
        browser.implicitly_wait(15)
        return result


# Авторизация на сайте
def authorization():
    print(f'Авторизация на сайте -> {conf_dict.get("site").split("/")[2]}')

    # Проверка - Подтверждение персональных данных
    def verification_for_confirmation_of_personal_data(browser_pd):
        try:
            browser_pd.implicitly_wait(1)
            print(f'Подтверждение персональных данных')
            time_check = 0
            while True:
                container = browser_pd.find_element_by_xpath('//*[@class="modal-dialog"]/..')
                block = browser_pd.execute_script('return arguments[0].style.display;', container)
                if block == 'block' or time_check == 8:
                    break
                else:
                    time_check += 1
                    time.sleep(0.5)

            if browser_pd.capabilities['browserName'] != 'firefox':
                time.sleep(1)

            browser_pd.find_element_by_xpath('//*[@ng-click="okIs() && ok()"]').click()

        except NoSuchElementException:
            print(f'Подтверждение персональных данных - Не требовалось!')
        except ElementNotInteractableException:
            browser_pd.save_screenshot(f'{log_out}Ошибка - Подтверждение персональных данных.png')
            print_exc()
        else:
            print(f'Подтверждение персональных данных - Успешно пройдено!')
        finally:
            browser_pd.implicitly_wait(15)

    try:
        # Браузер - Chrome
        options = webdriver.ChromeOptions()
        # options.add_argument('--start-maximized')
        options.add_argument("--headless")
        options.add_argument("--window-size=1500x1200")

        # options.add_argument('user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko)'
        #                      ' Chrome/90.0.4430.212 Safari/537.36')
        options.add_argument('--disable-blink-features=AutomationControlled')

        options.add_experimental_option("excludeSwitches", ["enable-automation"])
        options.add_experimental_option('useAutomationExtension', False)

        options.add_experimental_option(
            'prefs',
            {
                # 'profile.default_content_setting_values.notifications': 2,
                # 'profile.managed_default_content_settings.cookies': 2,        # Куки
                # 'profile.default_content_settings.popups': 2,                 # Внезапное окно с рекламой
                'profile.default_content_settings.geolocation': 2,              # Геолокация
                # 'profile.managed_default_content_settings.plugins': 2,
                # 'profile.managed_default_content_settings.fullscreen': 2,
                # 'profile.managed_default_content_settings.javascript': 2,     # JavaScript
                # 'profile.managed_default_content_settings.images': 2,         # Изображения
                # 'profile.managed_default_content_settings.mixed_script': 2,
                'profile.default_content_settings.media_stream': 2,             # Аудио/Видео
                # 'profile.managed_default_content_settings.stylesheets': 2     # CSS стили
            }
        )

        browser = webdriver.Chrome(
            executable_path=r'C:\Python_works\_!Настройки\chromedriver.exe',
            options=options
        )

        # Запуск
        browser.get(conf_dict.get('site'))

        # Неявные ожидания - до 15 секунд
        browser.implicitly_wait(15)

    except Exception as ex:
        print(f'\n------------\n'
              f'ОШИБКА: Проблема с брауером - Chrome!\n'
              f'Возможные варианты решения:\n'
              f'  - Браузер "Chrome" - Не установлен на PC;\n'
              f'  - Браузер "Chrome" - Требует обновления;\n'
              f'  - Webdriver Selenium - Не найден;\n'
              f'  - Webdriver Selenium - Требует обновления\n'
              f'------------\n'
              f'Сообщение ошибки: {ex}')

    else:
        try:
            user, password = Config.key_password(view=2)
            browser.find_element_by_xpath(PkaskoXpath.XPATH_SIGNIN).click()

            login = browser.find_element_by_xpath(PkaskoXpath.XPATH_LOGIN)
            time.sleep(0.2)
            login.send_keys(user)
            time.sleep(0.4)

            # # Проверка - Вводить Логин пока он не будет полностью указан [!!! Из-за не корректной работы в Опере убрано]
            # while login.get_attribute('value') != auth['login']:
            #     time.sleep(0.5)
            #     login.clear()
            #     time.sleep(0.5)
            #     login.send_keys(auth['login'])

            browser.find_element_by_xpath(PkaskoXpath.XPATH_PASSOWRD).send_keys(password)
            time.sleep(0.4)

            browser.find_element_by_xpath(PkaskoXpath.XPATH_GO_TO).click()

        except:
            print_exc()
            exit_code(browser)

        else:
            print(f'Авторизация на сайте -> {conf_dict.get("site").split("/")[2]} - Пройдена!')

            # Проверка - Подтверждение персональных данных
            verification_for_confirmation_of_personal_data(browser)

            return browser


# Закрытие браузера
def exit_code(browser):
    browser.close()
    browser.quit()
    try:
        raise SystemExit
    except SystemExit:
        print(f'\nРабота программы - завершена!!!\n')


# ------ Ввод ВСЕХ данных из Excel ------
def excel_import():
    print(f'----- Ввод данных из Экселя -----')

    global list_import

    # Количество строк данных в Excel - Всегда +1
    strok = 67 + 1

    xl = pd.read_excel('_Данные_для_теста_ПКАСКО.xlsx', sheet_name='Лист1', usecols='B:D', header=0, nrows=strok)
    for i in range(0, strok):
        try:
            if xl.iat[i, 0] == '' or pd.isnull(xl.iat[i, 0]):
                xl.iat[i, 0] = ''
        except IndexError as ix:
            print(f'IndexError = {ix}')
        else:
            list_import.append(xl.iat[i, 0])


# Запуск ПОТОКОВ
def threads():
    # Взаимодействие с пользователем (Выбор среды тестирования и д.р.)
    def selection():
        # Выбор сайта для тестирования
        def selection_site():
            while True:
                dict_site = {
                    1: 'https://pkasko.com/contact',
                    2: 'http://pkasko.dev0/contact',
                    3: 'http://pkasko.dev1/contact',
                    4: 'https://test-in.pkasko.com/contact',
                    5: 'http://pkasko.test0/contact',
                    6: 'http://pkasko.test1/contact'
                }

                print(f'Введите номер сайт ПКАСКО для тестирования:\n'
                      f'    1) Прод    - pkasko.com\n'
                      f'    2) Dev0    - pkasko.dev0\n'
                      f'    3) Dev1    - pkasko.dev1\n'
                      f'    4) Test-In - test-in.pkasko.com\n'
                      f'    5) Test0   - pkasko.test0\n'
                      f'    6) Test1   - pkasko.test1'
                      )
                number = input('Ввод #: ')

                if dict_site.get(int(number), 0) == 0:
                    print(f'\nОШИБКА! Повторите ввод "Cреды тестирования"!\n')
                else:
                    return dict_site[int(number)]

        # Выбор страховой компании для тестирования
        def selection_insurance_company():
            while True:
                print(f'\nВыберите Страховую компанию для тестирования:')
                [print(f'        {key}) {value}') for key, value in Config.ic.items()]
                number = input('Ввод #: ')

                if Config.ic.get(int(number), 0) == 0:
                    print(f'\nОШИБКА! Повторите ввод Страховой компании!\n')
                else:
                    return Config.ic.get(int(number))

        # Выбор программ для тестирования
        def selection_program():
            while True:
                dict_type = {
                    1: f'Только - {Config.twig[0]}',
                    2: f'Только - {Config.twig[1]}',
                    3: f'Все виды расчета сразу! (Повышается нагрузка на компьютер)'
                }

                print(f'\nВыберите вид страхования для тестирования:')
                [print(f'        {key}) {value}') for key, value in dict_type.items()]
                number = input('Ввод #: ')

                if dict_type.get(int(number), 0) == 0:
                    print(f'\nОШИБКА! Повторите ввод Типа тестирования!\n')
                else:
                    return int(number)

        global conf_dict

        # Выбор сайта для тестирования
        conf_dict['site'] = selection_site()

        # Выбор страховой компании для тестирования
        conf_dict['ic'] = selection_insurance_company()

        # Выбор программ для тестирования
        conf_dict['test'] = selection_program()

    # # Выгрузка данных из Экселя
    # excel_import()

    # Взаимодействие с пользователем
    selection()

    # Выгрузка данных из БД
    global import_dict
    import_dict = db_pkasko()

    # Потоки
    if conf_dict.get('test') != 3:
        t0 = threading.Thread(target=work, args=(), name=f'{conf_dict.get("test") - 1}')
        t0.start()
        t0.join()
    else:
        t0 = threading.Thread(target=work, args=(), name='0')
        t0.start()

        t1 = threading.Thread(target=work, args=(), name='1')
        t1.start()

        t0.join()
        t1.join()

    export_to_excel(list_export, GL_good_1=[])


# Создание папок
def creating_folders():
    data_today = date.today().strftime('%d.%m.%Y')
    time_now = time.strftime("%H.%M", time.localtime())
    folder_data_time = f'Тест [{conf_dict.get("site").split("/")[2]}] - [{conf_dict.get("ic")}] - [{data_today} - {time_now}]'

    # Базовая папка
    base = Config.trunk
    # Папка расчетов и вывода всей информации о нем
    global log_out
    log_out = f'{base}/{folder_data_time}/'

    # Создание папок
    for one in [base, log_out]:
        try:
            os.mkdir(one)
        except OSError:
            pass


# Вывод времени расчета
def timer(start_time, info):
    """
    :param info:                Информация для вывода в консоль
    :param start_time:          Время начала расчета
    :return: minutes, seconds - минуты и секунды
    """
    vremia = time.time() - start_time
    minutes = vremia // 60
    seconds = vremia - minutes * 60

    if not info:
        return int(minutes), int(seconds)
    else:
        print(f'\n@-@-@ {info}: {int(minutes)}m {int(seconds)}s @-@-@')


def main():
    start_time = time.time()

    threads()

    timer(start_time, info='Время работы программы')

    # input('Нажмите любую клавишу, чтобы выйти')


if __name__ == "__main__":
    main()
