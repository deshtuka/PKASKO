UI Autotest CASCO - PKASKO.

Structure:
1) main_test.py – Main file.
2) settings.py – Configuration. Configuring access.
3) PKASKO_SQL.py - Retrieving and selecting test-cases from database.
4) PKASKO_scraping.py - Website scraping
5) PKASKO_export_to_excel.py - Output test situations in Excel
6) all_pass.env – Login / Password.
7) Chromedriver – Selenium driver. You must download current version!

Import. Personal data involved in calculation taken from the database MySQL.

Task: Automatically import flow of personal data of insurer and vehicle into in Integration Calculator B2B. Subsequent receipt of insurance prize and a full breakdown of calculation parameters.

Purpose: Checking relevance of CASCO tariffs, tracking changes on the site.

Implementation: It is executed in the following sequence:
1) Import of personal data from the .xlsx file or from database;
2) Authorization, filling out forms, making calculations;
3) Saving the HTML page (necessary to track changes over time);
4) Parsing of all fields;
5) Export of all possible parameters and cases of calculation in .xlsx with formatting (convenient for analysis by users);
6) Creation of screenshots of all calculations, it is necessary to track errors and verify the correctness of the input / output of information.

Conclusion: This script allows you to get all possible variations from only one data in 2 minutes, increasing the device performance and adding more data, the implementation process is possible in large volumes.

For the script to work, you must have an account with working access (Login / Password).

-------------------------------------------------------------------------------------------------------------------------------------
-------------------------------------------------------------------------------------------------------------------------------------

UI Автотест КАСКО - ПКАСКО.

Структура:
1) main_test.py — Главный файл.
2) settings.py — Конфигурация. Настройка доступов.
3) PKASKO_SQL.py – Получение и выбор из БД тест-кейсов.
4) PKASKO_scraping.py – Парсинг сайта.
5) PKASKO_export_to_excel.py – Вывод тест ситуаций в Excel.
6) all_pass.env – Логин / Пароль.
7) Chromedriver — Драйвер Selenium. Необходимо скачать актуальную версию!

Импорт. Персональные данные участвующие в расчете берутся из БД (MySQL).

Задача: В автоматическом режиме импортировать поток персональных данных страховщика и транспортного средства в Калькуляторе интеграций B2B. Последующее получение страховой премии и полной разверстки параметров расчета.

Цель: Проверка актуальности тарифов КАСКО, отслеживание изменений на сайте.

Реализация: Выполнена в следующей последовательности:
1) Импорт персональных данных из файла .xlsx или БД;
2) Авторизация, заполнение форм, произведение расчетов;
3) Сохранение HTML страницы (необходимо для отслеживания изменений с течением времени);
4) Парсинг всех полей;
5) Экспорт всех возможных параметров и случаев расчета в .xlsx с форматированием (удобно для анализа пользователями);
6) Создание скриншотов всех расчетов, необходимо для отслеживания ошибок и удостоверении корректности ввода/вывода информации.

Вывод: Данный скрипт позволяет за 2 минуты получить все возможные вариации только по одним данным, повысив производительность устройства и добавив больше данных, процесс реализации возможен в больших объемах.

Для работы скрипта необходимо иметь учетную запись с рабочим доступ (Логин/Пароль).
