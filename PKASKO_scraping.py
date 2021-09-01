from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from lxml import html
from bs4 import BeautifulSoup
import re


def scraping(browser, block):
    """
    ----------------------------------------------------------------
    :param browser:     - [object] Вэб-драйвер
    :param block:       - [int]    Номер блока для парсинга
    :return:            - [list]   Нарощенный/Скомпанованный список
    ----------------------------------------------------------------
    Блок 1 - Вкладка "Клиент"
    Блок 2 - Параметры в ЛЕВОЙ части (Шаг 2, Шаг 3, Шаг 4,...)
    Блок 3 - Спец. блок СК в Левой части
    Блок 4 - Премия в ЛЕВОЙ части
    Блок 5 - Параметры в ПРАВОЙ части
    Блок 6 - Премия в ПРАВОЙ части
    ----------------------------------------------------------------
    """
    all_list = []

    for number in range(1, 6+1):
        if block == number:
            all_list = eval(f'scraping_block_{number}(browser)')

    # Тут можно выводить куда-то промежуточный результат

    return all_list


# Блок 1 - Вкладка "Клиент"
def scraping_block_1(browser):
    # Из-за странной верстки ПКАСКО отдельно реализованы блоки: Контактные данные и Транспортное средство
    def scraping_client(cbrowser, name):
        browser.implicitly_wait(0.1)
        content = cbrowser.page_source.encode('utf-8')
        dom = html.fromstring(content)

        xp = '//*[@ng-switch="client.data.type"]'
        list_param = []
        xpath_data = {
            'Тип клиента': {
                'Клиент': '',
                'Тип клиента': '//*[@ng-model="client.data.type"]/..//*[@class="select2-chosen"]'
            },
            'Выберите НЕ МЕНЕЕ ОДНОГО участника договора': {
                'Статус клиента': '',
                'Страхователь': '//*[@id="insurant"]',
                'Собственник': '//*[@id="owner"]',
                'Выгодоприобретатель': '//*[@id="beneficiary"]',
                'Водитель': '//*[@id="driver"]'
            },
            'Личные данные': {
                'Личные данные': '',
                'Фамилия': '//*[@ng-model="client.data.lastname"]',
                'Имя': '//*[@ng-model="client.data.firstname"]',
                'Отчество': '//*[@ng-model="client.data.middlename"]',
                'Дата рождения': '//*[@ng-model="client.data.birthdate"]',
                'Пол': '//*[@ng-model="client.data.data.sex"]/..//*[@class="select2-chosen"]',
                'В браке?': '//*[@ng-model="client.data.data.marriage"]/..//*[@class="select2-chosen"]',
                'Наличие детей?': '//*[@ng-model="client.data.data.children"]/..//*[@class="select2-chosen"]',
                'Резидент РФ': '//*[@ng-model="client.data.data.resident"]/..//*[@class="select2-chosen"]'
            },
            'Документы': {
                'Документы': '',
                'Паспорт': '',
                'Серия_1': '//*[@ng-model="client.data.data.documents.passport.series"]',
                'Номер_1': '//*[@ng-model="client.data.data.documents.passport.number"]',
                'Дата выдачи_1': '//*[@ng-model="client.data.data.documents.passport.date"]',
                'Кем выдан_1': '//*[@ng-model="client.data.data.documents.passport.issued"]',
                'Данные предыдущего паспорта': '',
                'Фамилия': '//*[@ng-model="client.data.prev.lastname"]',
                'Имя': '//*[@ng-model="client.data.prev.firstname"]',
                'Отчество': '//*[@ng-model="client.data.prev.middlename"]',
                'Серия_2': '//*[@ng-model="client.data.data.documents.prev.passport.series"]',
                'Номер_2': '//*[@ng-model="client.data.data.documents.prev.passport.number"]',
                'Дата выдачи': '//*[@ng-model="client.data.data.documents.prev.passport.date"]',
                'Кем выдан': '//*[@ng-model="client.data.data.documents.prev.passport.issued"]',
                'Водительское удостоверение': '',
                'Серия_3': '//*[@ng-model="client.data.data.documents.license.series"]',
                'Номер_3': '//*[@ng-model="client.data.data.documents.license.number"]',
                'Дата выдачи текущ._3': '//*[@ng-model="client.data.data.documents.license.dateLast"]',
                'Дата начала стажа': '//*[@ng-model="client.data.data.documents.license.date"]',
                'Предыдущее водительское удостоверение': '',
                'Серия_4': '//*[@ng-model="client.data.data.documents.prev.license.series"]',
                'Номер_4': '//*[@ng-model="client.data.data.documents.prev.license.number"]',
                'Дата выдачи текущ._4': '//*[@ng-model="client.data.data.documents.prev.license.dateLast"]'
            },
            'Контактные данные': {
                'Контактные данные': '',
                'Номер телефона': '//*[@ng-model="client.data.data.phone"]',
                'E-mail адрес': '//*[@ng-model="client.data.data.email"]',
                'Адрес регистрации': '',
                'Индекс': '//*[@ng-model="client.data.data.address.index"]',
                'Государство': '//*[@ng-model="client.data.data.address.country"]/..//*[@class="select2-chosen"]',
                'Регион': '//*[@ng-model="client.data.data.address.regionSuggest"]/..//*[@class="select2-chosen"]',
                'Населённый пункт': '//*[@ng-model="client.data.data.address.townSuggest"]/..//*[@class="select2-chosen"]',
                'Улица': '//*[@ng-model="client.data.data.address.streetSuggest"]/..//*[@class="select2-chosen"]',
                'Дом': '//*[@ng-model="client.data.data.address.house"]',
                'Корп.': '//*[@ng-model="client.data.data.address.building"]',
                'Кв.': '//*[@ng-model="client.data.data.address.appartment"]'
            },
            'Транспортное средство (Заполните ': {
                'Транспортное средство': '',
                'Гос. номер': '//*[@ng-model="client.data.data.car.signNumber"]',
                'VIN': '//*[@ng-model="client.data.data.car.VIN"]',
                'Номер кузова': '//*[@ng-model="client.data.data.car.bodyNumber"]',
                'Документы': '',
                'Тип ПТС': '//*[@ng-model="client.data.data.car.document"]/..//*[@class="select2-chosen"]',
                'Серия ПТС': '//*[@ng-model="client.data.data.car.seriesPTS"]',
                'Номер ПТС': '//*[@ng-model="client.data.data.car.numberPTS"]',
                'Дата ПТС': '//*[@ng-model="client.data.data.car.datePTS"]',
                'Серия СТС': '//*[@ng-model="client.data.data.car.seriesSTS"]',
                'Номер СТС': '//*[@ng-model="client.data.data.car.numberSTS"]',
                'Дата СТС': '//*[@ng-model="client.data.data.car.dateSTS"]'
            },
        }

        for head, items in xpath_data.items():
            """
            1 - Просто заголовок
            2 - Select
            3 - Checked
            4 - Input
            """
            if head == name:
                for key, xpath in items.items():
                    if len(dom.xpath(f'{xp}{xpath}')) > 0:
                        try:
                            if xpath == '':
                                list_param.append([key, xpath])

                            elif xpath[-16:-2] == 'select2-chosen':
                                value = dom.xpath(f'{xp}{xpath}')
                                list_param.append([key, str(value[0].text).strip()])

                            elif name == 'Выберите НЕ МЕНЕЕ ОДНОГО участника договора':
                                clink = cbrowser.find_element_by_xpath(f'{xp}{xpath}')
                                list_param.append([key, 'Да' if clink.get_attribute('checked') == 'true' else 'Нет'])

                            else:
                                clink = cbrowser.find_element_by_xpath(f'{xp}{xpath}')
                                ActionChains(cbrowser).move_to_element(clink).perform()
                                list_param.append([key, str(clink.get_attribute('value')).strip()])

                        except:
                            pass

        browser.implicitly_wait(15)
        return list_param

    all_list = []
    website = browser.page_source.encode('utf-8')
    tree = html.fromstring(website)

    # Название блоков у клиентов
    xpath_blocks = '//*[@ng-switch="client.data.type"]//*[@class="panel-heading"]'

    # Цикл по блокам
    try:
        body_header = tree.xpath(xpath_blocks)
        tmp = body_header[0]
    except IndexError:
        pass
    else:
        # Заголовки блоков Клиентов
        list_headers = [one.text for one in body_header]

        for header in list_headers:
            all_list.extend(scraping_client(browser, header))

    return all_list


# Блок 2 Параметры в ЛЕВОЙ части (Шаг 2, Шаг 3, Шаг 4,...)
def scraping_block_2(browser):
    return scraping_block(browser, block=2)


# Блок 5 - Параметры правой части
def scraping_block_5(browser):
    return scraping_block(browser, block=5)


# Блок 2 и 5 - Параметры в ЛЕВОЙ части (Шаг 2, Шаг 3, Шаг 4,...)
def scraping_block(browser, block):
    # Водители в предварительном
    def driver(browser_d):
        all_list_d = []
        browser_d.find_element_by_tag_name('body').send_keys(Keys.HOME)
        xpath_count = '//*[@id="kaskoDriverLayer"]/div'

        for num in range(1, len(tree.xpath(xpath_count)) + 1):
            all_list_d.append([f'Водитель №{num}', ''])
            # Input'ы
            link = browser_d.find_elements_by_xpath(f'{xpath_count}[{num}]//*[@type="text"]')
            keys = [one.get_attribute('placeholder') for one in link]
            values = [one.get_attribute('value') for one in link]
            all_list_d += [[str(keys[i]).strip(), str(values[i]).strip()] for i in range(len(keys))]

            # Пол
            values = browser_d.find_elements_by_xpath(f'{xpath_count}[{num}]//*[@type="button"]')[0].get_attribute('className')
            value = 'Мужской' if len([True for one in re.sub(r'[\s]', '', str(values)) if one == 'active']) > 0 else 'Женский'
            all_list_d += [['Пол', value]]

            # CheckBox'ы
            link = browser_d.find_elements_by_xpath(f'{xpath_count}[{num}]//*[starts-with(@class,"checkbox")]')
            keys = [one.get_attribute('innerText') for one in link]
            values = ['Да' if one.get_attribute('className') == 'checkbox active' else 'Нет' for one in link]
            all_list_d += [[str(keys[i]).strip(), str(values[i]).strip()] for i in range(len(keys))]

        return all_list_d

    # Парсинг всех типов параметров
    def parameter_scraping(browser_ps, xpath_active_ps, all_list_ps, head_ps):
        treep = html.fromstring(browser_ps.page_source.encode('utf-8'))

        # Переменная с xpath заголовков
        add_key = {
            2: '//*[@class="h3-negative"]',                 # Левая часть
            5: '//*[@data-ng-bind-html="params.title"]'     # Правая часть
        }

        # Ключи значений
        add_values = {
            'input': f'/../..//input[@type="text" or @type="number"]',  # value
            'date': f'/../..//input[@ui-date-format="yy-mm-dd"]',       # value
            'select': f'/../..//span[@class="select2-chosen"]',         # innerText
            'checkbox': f'/../..//input[@type="checkbox"]',             # checked
            'radiobutton': f'/../..//input[@type="radio"]'              # checked
        }

        xpath_key = f'{xpath_active_ps}{add_key.get(head_ps)}'
        heads = treep.xpath(xpath_key)

        for header in heads:
            key = header.text
            value = ''
            if key != '':
                for values in add_values.items():
                    try:
                        # Проверка на существование пути через lxml, чтобы не затупливать Selenium на ожидания
                        xpath_values = f'{xpath_active_ps}//{"span" if head_ps == 5 else "h4"}[text() = "{key}"]{values[1] if head_ps == 5 else values[1][3:]}'
                        body = treep.xpath(xpath_values)
                        if len(body) > 0:
                            if values[0] != 'select':
                                link = browser_ps.find_elements_by_xpath(xpath_values)
                                ActionChains(browser_ps).move_to_element(link[-1 if len(body) > 1 else 0]).perform()

                                if values[0] == 'input' or values[0] == 'date':
                                    value = [one.get_attribute('value') for one in link if one.get_attribute('value') != ''][0]
                                elif values[0] == 'checkbox':
                                    value = 'Да' if link[0].get_attribute('checked') == 'true' else 'Нет'
                                elif values[0] == 'radiobutton':
                                    value = [one.get_attribute('value') for one in link if one.get_attribute('checked') == 'true'][0]
                                    value = 'Нет' if value == '' else value  # Подстраховка, если не нашли RadioButton == True

                            # Единственный костыль Selenium'a для Select'а, потому что в Марке присутствует логотип
                            elif values[0] == 'select' and key == 'Марка':
                                value = str(browser_ps.find_element_by_xpath(xpath_values).get_attribute('innerText')).strip()
                            else:
                                value = '' if body[0].text is None else body[0].text

                            break
                    except:
                        pass

            if key != value:
                all_list_ps.append([key, value])

        return all_list_ps

    all_list = []
    website = browser.page_source.encode('utf-8')
    tree = html.fromstring(website)

    # Кол-во И название блоков
    xpath_block = {
        2: '//*[starts-with(@class, "panel panel-grey")]',
        5: '//div[@class="titleblock ng-scope"]'
    }
    body_block = tree.xpath(xpath_block.get(block))

    # Название блоков
    headers = {
        2: ['//*[starts-with(@class,"panel-heading")]//h4'],
        5: [
            '//b[@data-ng-bind-html="group.group"]',
            '//b[starts-with(@class,"titleblock_title")]'
        ]
    }

    # Цикл по блокам
    for number in range(len(body_block) + 1):  # number < 5
        # Заголовок блока (string)
        header_text = ''
        for xpath_header_text in headers.get(block):
            xpath_header = f'{xpath_block.get(block)}[{number}]{xpath_header_text}'
            try:
                body_header = tree.xpath(xpath_header)
                tmp = body_header[0]
            except IndexError:
                pass
            else:
                header_text = body_header[0].text
                all_list.append([header_text, ''])
                break

        # После получения названия, работаем только с этим блоком
        if header_text != '':
            xpath_active_block = f'//{"h4" if block == 2 else "b"}[text() = "{header_text}"]/ancestor::{xpath_block.get(block)[2:]}'
            # Парсинг Водителей в предварительном
            if header_text == 'Водители' and block == 2:
                all_list += driver(browser)
            # Парисинг параметров
            all_list = parameter_scraping(browser, xpath_active_block, all_list, block)

    return all_list


# Блок 4 - Премия в ЛЕВОЙ части
def scraping_block_4(browser):
    all_list = []
    website = browser.page_source.encode('utf-8')

    all_list.append(['Премия, руб', ''])

    tree = html.fromstring(website)
    soup = BeautifulSoup(website, features="lxml")

    xpath = '//*[@class="kasko-result"]//tbody/tr'
    for i in range(1, len(tree.xpath(xpath))+1):
        xpath_row = f'{xpath}[{i}]'
        # Страховая компания
        try:
            value = soup.find('div', {'class': 'kasko-result'}).find('tbody').find_all('tr')[i-1].find('div', {'class': 'ng-binding'}).text
            all_list.append(['Страховая компания', str(value).strip()])
        except IndexError as ex:
            print(f'Ошибка: Страховая компания = {ex}')

        # Программа
        body_values = tree.xpath(f'{xpath_row}//b[@class="ng-binding"]/..')
        if len(body_values) > 0:
            all_list.append(['Программа', body_values[0].attrib['title']])

        # Премия
        body_values = tree.xpath(f'{xpath_row}//div[@class="result ng-binding ng-scope"]')
        if len(body_values) > 0:
            value = [one.strip() for one in str(body_values[0].text).split('-')]
            all_list.append(['Премия', value[1]])

        # Ошибка
        body_values = tree.xpath(f'{xpath_row}//*[@ng-if="kaskoResult.error"]//span')
        if len(body_values) > 0:
            value = body_values[0].attrib['title']
            all_list.append(['Ошибка', value])

        # НС - ...
        try:
            value_ns = soup.find('div', {'class': 'kasko-result'}).find('tbody').find_all('tr')[i-1].find('span', {'title': 'Несчастный случай'}).text
            value = [str(one).strip() for one in str(value_ns).split(':')]
            all_list.append(['Несчастный случай', value[1]])
        except AttributeError:
            pass
        except IndexError:
            pass

        # Франшиза -
        try:
            value_fr = soup.find('div', {'class': 'kasko-result'}).find('tbody').find_all('tr')[i-1].find('span', {'title': 'Франшиза'}).text
            value = [str(one).strip() for one in str(value_fr).split(':')]
            all_list.append(['Франшиза', value[1]])
        except AttributeError:
            pass
        except IndexError:
            pass

    return all_list


# Блок 6 - Премия в ПРАВОЙ части
def scraping_block_6(browser):
    all_list = []
    website = browser.page_source.encode('utf-8')

    all_list.append(['Премия, руб', ''])

    tree = html.fromstring(website)
    soup = BeautifulSoup(website, features="lxml")

    # Дополнительные параметры
    bs4_key = soup.find_all('p', {'ng-if': 'featureValue'})
    for i in range(len(bs4_key)):
        key, value = ('' for i in range(2))
        try:
            key = str(bs4_key[i].text).split(':')[0].strip()
            value = str(bs4_key[i].text).split(':')[1].strip()
        except IndexError:
            pass
        all_list.append([key, value])

    # Премии
    xpath_key = f'//*[@id="resultPanel"]/tbody/tr/td'                   # Название (Каждое третье название)
    xpath_value = f'//*[@id="resultPanel"]//*[@class="ng-binding"]'     # Премия (Каждое нечетное число)

    body_key = tree.xpath(xpath_key)
    body_value = tree.xpath(xpath_value)

    keys = [body_key[i].text for i in range(len(body_key)) if i % 3 == 0]
    values = [body_value[i].text for i in range(len(body_value)) if i % 2 != 0]

    tmp = []
    # Объединение списков (Названий/Значений) и разварот (от меньшего к максимуму)
    if len(keys) == len(values):
        # key = [str(keys[i]).strip() for i in range(len(keys))]
        # value = [str(values[i]).strip() for i in range(len(values))]
        # tmp = [[str(key[i]).strip(), str(value[i]).strip()] for i in range(len(values))]
        tmp = [[str(keys[i]).strip(), str(values[i]).strip()] for i in range(len(keys))]
    all_list += tmp[::-1]
    tmp.clear()

    # Предупреждение
    bs4_value = soup.find_all('span', {'data-ng-bind-html': 'warning'})

    tmp = []
    for values in bs4_value:
        warning = []
        for key in values.contents:
            warning.append(str(re.sub(r'[<br/>|<b>]', '', str(key))).strip())
        tmp.append('\n'.join([one for one in warning if one != '']))

    all_list.append(['Предупреждение', '\n'.join(tmp)])
    tmp.clear()

    return all_list
