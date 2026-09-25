import pandas as pd, numpy as np, os, socket, io, glob, shutil, sqlalchemy, sys, win32com.client, subprocess, copy, msoffcrypto, threading
import pythoncom
pythoncom.CoInitializeEx(0)
from datetime import datetime

# блок импорта отправки почты
import smtplib,ssl
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import formatdate
from email import encoders

import logging

def dir_link():
    """возвращает абсолютный путь
    """
    import os
    try:
        script_dir = os.path.dirname(os.path.abspath(__file__))
        return script_dir
    except:
        script_dir_2 = os.getcwd()
        return script_dir_2
DIR = dir_link()

from functools import wraps
import time
# декоратор для times-повторного выполнения функции при неудачном выполнении 
def retry(times, sec_):
    """_summary_

    Args:
        times (_type_): попыток
        sec_ (_type_): секунд между попытками
    """
    def wrapper_fn(f):
        @wraps(f)
        def new_wrapper(*args,**kwargs):
            for i in range(times):
                try:
                    print ('---ПОПЫТКА ЧТЕНИЯ ФАЙЛА ---- %s' % (i + 1))
                    return f(*args,**kwargs)
                except Exception as e:
                    error = e
                    print(time.sleep(sec_))
            raise error
        return new_wrapper
    return wrapper_fn

# === НАСТРОЙКА ЛОГИРОВАНИЯ С ПЕРЕЗАПИСЬЮ ===
LOG_FILE = os.path.join(DIR, "log_main_kum.log")

# Удаляем старый лог, если он существует
if os.path.exists(LOG_FILE):
    try:
        os.remove(LOG_FILE)
        print(f" Старый лог удалён: {LOG_FILE}")
    except PermissionError:
        print(f"Не удалось удалить старый лог: возможно, файл открыт в другом приложении")
    except Exception as e:
        print(f"Ошибка при удалении лога: {e}")

# Теперь настраиваем логирование — файл будет создан заново
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)-8s | %(message)s",
    handlers=[
        logging.FileHandler(LOG_FILE, encoding="utf-8"),  # Создастся как новый
        logging.StreamHandler()  # вывод в консоль
    ]
)

logger = logging.getLogger(__name__)
logger.info("Запуск скрипта main_kum - новая сборка кумов (все данные)")

@retry(10, 5)
def links_main(name_file, key):
    """функция для работы с путями, ссылки, вводные данные хранятся в блокноте

    Args:
        name_file (_type_): имя файла
        key (_type_): имя ключа

    Returns:
        _type_: _description_
    """
    try:
        file = pd.read_csv(name_file, sep=';')
        result = list(file[file['ключ']==key]['значение'])[0]
        return result
    except Exception as ex_:
        logger.error(f'ошибка функции {links_main.__name__} не удалось считать файл {name_file} или данные в нем {key} ошибка {ex_}')

def update_task(link_script):
    """зупуск субпроцесса 
    Args:
        link_script (_type_): ссылка на файл скрипта .py

    Raises:
        FileNotFoundError: _description_
        RuntimeError: _description_
    """
    try:
        link_script = link_script
        # Проверка существования скриптов
        if not os.path.exists(link_script):
            raise FileNotFoundError(f"Скрипт субпроцесса не найден: {link_script}")
        # === Этап 1: Обновление  ===
        python_exe = sys.executable
        logger.info(f"Использую Python: {python_exe}")  # Отладка
        logger.info(f"Версия: {sys.version}")           # Отладка
        logger.info(f"Запуск субпроцесса {link_script}")
        result = subprocess.run(
            [python_exe, link_script],
            capture_output=True,
            check=True, 
            shell=False     
        )
        if result.returncode != 0:
            raise RuntimeError(f"Ошибка в {link_script}:\n{result.stderr}")
        logger.info("✅ Файл успешно обработан")
    except Exception as e:
        logger.error(f"Ошибка при выполнении субпроцесса: {e}")


def get_data_user_param(socket_name=True):
    """возврат данных пользователя по соккету (имени ПК) для отправки почты в формате кортежа
    name_pc server port username send_from password

    Args:
        socket_name (bool, optional): _description_. Defaults to True.

    Returns:
        _type_: _description_
    """
    try:
        df = pd.read_excel(links_main(fr"{DIR}\main_links.txt", "pass_all"), engine='calamine', sheet_name='PC')
        if len(df) > 0:
            logger.info(f"датафрейм обнаружен")
            if socket_name == True: 
                socket_name = socket.gethostname()
                logger.info(f'ищем данные по сокету {socket_name}')
                records_dict = df.to_dict('records')
                index_result = [index for index, dct in enumerate(records_dict) if dct['name_pc'] == socket_name]
                if len(index_result)>0:
                    try:
                        index_data=index_result[0]
                        result_all = records_dict[index_data]
                        name_pc = result_all['name_pc']
                        server = result_all['server']
                        port = result_all['port']
                        username = result_all['username']
                        send_from = result_all['send_from']
                        password = result_all['password']
                        logger.info(f'данные получены и возвращены кортежем в следующем порядке : name_pc server port username send_from password')
                        return name_pc, server, port, username, send_from, password
                    except Exception as e:
                        logger.error(f'Ошибка функции {get_data_user_param.__name__} в блоке получения данных пользователя по имени пк {socket_name} ошибка {e}')
            else :
                logger.info(f"Поиск данных с учетом соккета не включен, будет возвращен весь фрейм с данными")
                return df
        else: 
            logger.info(f"датафрейм не найден")
            return None
    except Exception as e:
        logger.error(f'Ошибка функции {get_data_user_param.__name__} ошибка {e}')

def test_connect_SQL_driver():
    """функция проверки доступных драйверов для работы с SQL
    от этого зависит какой драйвер будем использовать
    на серверном пк s-kao3 новый ODBC Driver 18 for SQL Server
    на рабочих пк ODBC Driver 13 for SQL Server
    так же на серверном варинате в настройке нужно отключение 
    проверки сертификатов SSL - это описывается в самих функциях подключения к SQL
    """
    try:
        import pyodbc
        wite_list_driver = ['ODBC Driver 13 for SQL Server', 'ODBC Driver 18 for SQL Server']
        logger.info(f'🔄 Проверка достутных драйверов')

        list_driver = []
        for driver in pyodbc.drivers():
            if 'SQL' in driver:
                list_driver.append(driver)

        logger.info(f'🛠️ Доступные драйвера SQL: {list_driver}')

        if len(list_driver)>0:
            found_drivers = []
            for i in wite_list_driver:
                if i in list_driver:
                    found_drivers.append(i)
            if len(found_drivers)>0:
                logger.info(f'✅ Найден драйвер: {found_drivers[0]} соответсвующий белому списку: {wite_list_driver}')
                return found_drivers[0]
            else:
                logger.error(f'⚠️ Нет драйверов соответсвующих белому списку: {wite_list_driver}')
        else:
            logger.error(f'⚠️ Доступные драйвера SQL не обнаружены')
    except Exception as e:
        logger.error(f'ошибка функции {test_connect_SQL_driver.__name__} {e}')

def connect_bd_SQL(server, database, username, password):
    """функция подключения к бд SQL
    Returns:
        _type_: _description_
    """
    # import sqlalchemy
    # Замените значения на ваши
    # если серит ошибками - проверить данные по столбцам ибо в наценке оборотки было inf и из-за этого серило ошибки
    server = server
    database = database
    username = username
    password = password
    driver = test_connect_SQL_driver() # 'ODBC Driver 13 for SQL Server' или 'ODBC Driver 18 for SQL Server' # моддет работать просто 'SQL Server'

    try:
        # connection_string = f'mssql+pyodbc://{username}:{password}@{server}/{database}?driver={driver}' # или f'mssql://{username}:{password}@{server}/{database}?driver={driver}'
        # исправленный вариант
        connection_string = (
            f'mssql+pyodbc://{username}:{password}@{server}/{database}'
            f'?driver={driver}&Encrypt=no'  # эта строка отключает проверку SSL сертификата
        )
        engine = sqlalchemy.create_engine(
                    connection_string,
                    echo=False, 
                    pool_pre_ping=True,
                    fast_executemany=True) 
        # fast_executemany=True - можно удалить эту строчку / Она оптимизирует процедуру массовых вставок, 
        # значительно сокращая количество запросов к базе данных, ускоряя запись
        logger.info(f'соединение с bd SQL - установлено')
        return engine
    except Exception as ex_:
        logger.error(f'Проблемы с подключением к SQL')
        logger.error(f'ошибка функции {connect_bd_SQL.__name__} ошибка {ex_}')


def exception_column_SQL(df, server, database, username, password):
    """проверяет наличие ошибок в df при записи в SQL
    перебирает каждый столбец и записывает в тестовую таблицу БД SQL
    при возникновении ошибки в записи - возвращает имя столбца и ошибку
    чаще такие ошибки вызван ытипом данных inf 
    Args:
        df (_type_): _description_
    """
    import sqlalchemy
    server = server
    database = database
    username = username
    password = password
    driver = test_connect_SQL_driver() # driver = 'ODBC Driver 13 for SQL Server' or # driver = 'ODBC Driver 18 for SQL Server'

    # connection_string = f'mssql+pyodbc://{username}:{password}@{server}/{database}?driver={driver}' # или f'mssql://{username}:{password}@{server}/{database}?driver={driver}'
    # исправленный вариант
    connection_string = (
        f'mssql+pyodbc://{username}:{password}@{server}/{database}'
        f'?driver={driver}&Encrypt=no'  # эта строка отключает проверку SSL сертификата
    )
    engine = sqlalchemy.create_engine(
                connection_string,
                echo=False, 
                pool_pre_ping=True,
                fast_executemany=True) # fast_executemany=True - можно удалить эту строчку / Она оптимизирует процедуру массовых вставок, значительно сокращая количество запросов к базе данных

    # перебираем каждую колонку df и пытаемся записать
    for i in df.columns:
        try:
            df[[i]].to_sql('df_ttt', con=engine, if_exists='replace', index=False)
            logger.info(f'✅ столбец {i} - ок')
        except Exception as ex_:
            logger.error(f'❌ {exception_column_SQL.__name__} ОШИБКА в столбце {i}  возможно порблемы  стипом данных {ex_}')

def return_params_in_link(df, name_col, name_serch:str, return_col:str):
    """возвращаем значение из столбца по ссылке
    Args:
        df (_type_): фрейм с ссылками
        name_col (str): имя столбца с ссылками
        name_serch (str): имя столбца в котором нужно искать значение
        return_col (str): имя столбца в котором нужно вернуть значение
    Returns:
        _type_: _description_
    """
    res = list(df[df[name_col]==name_serch][return_col])
    if len(res) != 0:
        return res[0]
    else:
        return None


def open_df_locked(link: str, password: str, lst_name = None):
    print('open_df_locked')
    """функция обработки заблокированных книг EXCEL с паролем

    Args:
        link (str): ссылка на книгу
        password (str): пароль
        lst_name (str): название листа можно ввести и откроет конкретный лист, если нет то по умолчанию

    Returns:
        _type_ (df, list): вовзращает df и список листов в книге
    """
    logging.info(f"{open_df_locked.__name__} - ЗАПУСК")
    try:
        lnk = link
        passwd = password                                 # пароль книги excel
        decrypted_workbook = io.BytesIO()
        with open(lnk, 'rb') as file:
            office_file = msoffcrypto.OfficeFile(file)
            office_file.load_key(password=passwd)
            office_file.decrypt(decrypted_workbook)

        xlsx_file = pd.ExcelFile(decrypted_workbook)
        sheet_names = xlsx_file.sheet_names             # получаем имена листов в книге  
        if lst_name == None:
            df = pd.read_excel(decrypted_workbook, sheet_name=None)
            return  df, sheet_names
        else:
            df = pd.read_excel(decrypted_workbook, sheet_name=lst_name)
            return  df, sheet_names
    except:
        logging.error(f"{open_df_locked.__name__} - ОШИБКА", exc_info=True)


def open_df_unlocked(link: str, lst_name = None):
    print('open_df_unlocked')
    """функция обработки не заблокированный книг EXCEL без пароля

    Args:
        link (str): ссылка на книгу
        lst_name (str): название листа можно ввести и откроет конкретный лист, если нет то по умолчанию

    Returns:
        _type_ (df, list): вовзращает df и список листов в книге
    """
    logging.info(f"{open_df_unlocked.__name__} - ЗАПУСК")
    try:
        lnk = link
        if lst_name == None:
            df = pd.read_excel(lnk, engine='calamine', sheet_name=None)
            sheet_names = list(pd.read_excel(lnk, sheet_name=None).keys())
            return df, sheet_names
        else:
            df = pd.read_excel(lnk, engine='calamine', sheet_name=lst_name)
            sheet_names = list(pd.read_excel(lnk, sheet_name=None).keys())
            return df, sheet_names
        
    except:
        logging.error(f"{open_df_unlocked.__name__} - ОШИБКА", exc_info=True)

def open_dataframe(link, password='0', lst_name = None):
    print('open_dataframe')
    """функция открытия книги с паролем или без (задействует две доп функции)

    Args:
        link (_type_): ссылка на книгу
        password (str, optional): пароль. Defaults to '0'.
        lst_name (_type_, optional): Имя листа. Defaults to None.

    Returns:
        _type_: _description_
    """
    logging.info(f"{open_dataframe.__name__} - ЗАПУСК")
    try:
        if len(password)>1:
            df, sheet_names = open_df_locked(link, password, lst_name)
            return df, sheet_names
        else:
            df, sheet_names = open_df_unlocked(link, lst_name)
            return df, sheet_names
    except:
        logging.error(f"{open_dataframe.__name__} - ОШИБКА", exc_info=True)


def Shapka(table, text='VIN'):
    """ функция ищет в какой строке находится шапка таблицы, путём поиска "VIN" в ограниченной таблице table.iloc[0:15,0:15]
                    и вырезает лишние куски до найденной шапки и вместе с ней
                    если ошибка, то оставляем входную таблицу оставляем без именений

    Args:
        table (_type_): таблица
        text (str, optional): поиск по ключевому слову. Defaults to 'VIN'.

    Returns:
        table: таблица только с заголовками без верхних вспомогательных строк
    """
    
    try:
        table = table.T.reset_index().T.reset_index(drop=True) # опускает имена столбцов в первую строку
        qqq = table.iloc[0:20,0:20] # кординаты поиска 20х20
    
        # находим в какой строке находится шапка таблицы
        qqq = qqq[qqq.astype(str).apply(lambda x: x.astype(str).str.upper().str.contains(text, case=False)).any(axis=1)]
        q = qqq.index[0]
        table.columns = table.loc[q]
        table = table.iloc[q+1:,:]
        table = table.reset_index(drop = True)
        table = table[table.columns.dropna()]
        return(table)
    except:
        return(table)

def file_update(link):
    
    "возвращает дату последнего обновления файла"
    try:
        from datetime import datetime, date, timedelta
        res = datetime.fromtimestamp(os.path.getmtime(link))
        return res
    except Exception as ex_:
        print(f'ошибка функции {file_update.__name__} не удалось считать метаданные файла {link} ошибка {ex_}')

def is_number(s):
    """првоерка на число"""
    try:
        float(s)
        return True
    except ValueError:
        return False

def mazda_next_msk(marka_2:str, region:str, s_lista:str):
    """распределение мазды некст МСК
    Args:
        marka_2 (str): значение марки из столбца marka_2
        region (str): регион
        s_lista (str): с какого листа
    Returns:
        _type_: _description_
    """
    try:
        if marka_2 == 'MAZDA' and region == 'MSK' and s_lista == 'Next': return 'MAZDA_NEXT'
        else: return marka_2
    except Exception as ex_:
        print(f'ошибка функции {mazda_next_msk.__name__} входные параметры {marka_2, region, s_lista} ошибка {ex_}')

def razdelenie_OVP_yar_region(marka_2, region, salon):
    """распределение регионов по ОВП ЯР на YAR и RYB
    Args:
        marka_2 (_type_): марка_2
        region (_type_): регион
        salon (_type_): салон
    Returns:
        _type_: _description_
    """
    try:
        if marka_2 == 'OVP' and region == 'YAR' and salon == 'Ярославль': return 'YAR'
        elif marka_2 == 'OVP' and region == 'YAR' and salon == 'Рыбинск': return 'RYB'
        else: return region
    except Exception as ex_:
        print(f'ошибка функции {razdelenie_OVP_yar_region.__name__} входные параметры {marka_2, region, salon} ошибка {ex_}')

def form_pay(x):
    """определяет кредит / нал 
    Args:
        x (_type_): _description_
    Returns:
        _type_: _description_
    """
    try:
        x = str(x).lower().strip()
        kredit_lst = ['кредит', 'банк', 'лизинг','кре', 'черилегко']
        if 'не для' not in x and any([i in x  for i in kredit_lst]):
            return 'кредит'
        elif 'б/н' in x or 'безнал' in x:
            return 'нал'
        else:
            return 'нал'
    except Exception as ex_:
        print(f'ошибка функции {form_pay.__name__} входные параметры {x} ошибка {ex_}')

def forma_oplaty_OVP_in_kommentary(forma_oplaty_in_col, str_in_kommentary):
    """функция корреткировки формы оплаты для ОВП проверяет комментарий на наличие слова кредит или лизинг
    Args:
        forma_oplaty_in_col (_type_): форма оплаты заполненная в столбец по факту - столбец форма оплаты
        str_in_kommentary (_type_): строка из столбца комментарий
    Returns:
        _type_: _description_
    """
    try:
        lst_triger = ['кредит', 'лизинг']
        str_in_kommentary = str(str_in_kommentary).lower()
        result_kred = any([i in str_in_kommentary for i in lst_triger])
        if result_kred:
            return 'кредит'
        else:
            return forma_oplaty_in_col
    except Exception as ex_:
        print(f'ошибка функции {forma_oplaty_OVP_in_kommentary.__name__} входные параметры {forma_oplaty_in_col, str_in_kommentary} ошибка {ex_}')
    

def fiz_yur(x):
    """разделяет на физ и юр

    Args:
        x (_type_): столбце с признаками физ и юр

    Returns:
        _type_: _description_
    """
    try:
        word = str(x).lower().split()
        if len(word) > 0:
            word = word[0]
        if any([j in word for j in ['юр','флит']]): return 'юрл'
        elif 'физ' in word: return 'физ'
        else: return x
    except Exception as ex_:
        print(f'ошибка функции {fiz_yur.__name__} входные параметры {x} ошибка {ex_}')

def fiz_yur_stage_2(x, klient):
    """второй этап распределения на физ и юрл
    если значени не равно физ или юрл - пытается определенить значение по клиенту

    Args:
        x (_type_): _description_
        klient (_type_): _description_

    Returns:
        _type_: _description_
    """
    try:
        word = str(x).lower().split()
        if len(word) > 0:
                word = word[0]
        if word not in ['физ', 'юрл'] or x == None or x == 'none':
            # проверяем клиента
            lst_klient = ['ИП', 'ООО', 'АО', 'ЗАО', 'ГК', 'ПАО', 'АНО', 'ХК', 'ЛК', 'ОО', 'ЛИЗИНГ', 'ГБУЗ', 'РЕСО', 'ВТБ', 'СПК']
            if any([j in klient for j in lst_klient]): return 'юрл'
            else: return 'физ'
        else: return x
    except Exception as ex_:
        print(f'ошибка функции {fiz_yur_stage_2.__name__} входные параметры {x} ошибка {ex_}')
        
def korp_demo(name_list:str):
    """
    если в имени листа есть признак демо то ставим True
    Args:
        name_list (str): столбец с_листа
    Returns:
        _type_: _description_
    """
    try:
        lst = ['ДЕМО', 'СВОД_ДЕМО']
        if name_list in lst: return 1
        else: return 0
    except Exception as ex_:
        print(f'ошибка функции {korp_demo.__name__} входные параметры {name_list} ошибка {ex_}')

def treid_in_auto(x):
    """трейдин да или нет - если сумма дохода за трейдин не равна 0 то 1 иначе 0
    Args:
        x (_type_): сумма по доходу трейдин
    Returns:
        _type_: 1 or 0
    """
    try:
        if x == None: return 0
        elif float(x) != 0: return 1
        else: return 0
    except Exception as ex_:
        print(f'ошибка функции {korp_demo.__name__} входные параметры {treid_in_auto} ошибка {ex_}')


def stavka_merge(date_start, date_end, sebestoimost:float, df_so_stavkami):
    """расчет простоя авто 
    создается календарь простоя авто с и по даты с его себестомостью и мерджится фрейм со ставками
    расчитывается сумма ежедневного простоя - ссумируется
    Args:
        date_start (_type_): начальная дата / дата_оплаты_счета
        date_end (_type_): конечная дата / дата_полной_оплаты_факт	
        sebestoimost (float): себестоиомость авто 
        df_so_stavkami (_type_): df сов семи ставками для merge
    primer:    stavka_merge('2024-10-10', '2024-10-15', 4000000.0, df_stavki)
    Returns:
        _type_: float
    """
    from datetime import timedelta
    from datetime import datetime

    if isinstance(date_start, str) and isinstance(date_end, str):                   # проверка на строку и преобразование в дату
        date_string_start = date_start
        date_string_end = date_end
        format_pattern = "%Y-%m-%d"
        date_start = datetime.strptime(date_string_start, format_pattern)
        date_end = datetime.strptime(date_string_end, format_pattern)
    else:
        date_start = date_start
        date_end = date_end
    try:
        if date_start<=date_end:                                                    # проверка на дату начала меньше даты конца
            date_end = date_end-timedelta(days=1)
            df_kal = pd.DataFrame({'календарь': pd.date_range(date_start,  date_end)}) # создаем календарь 
            df_kal['себестоимость'] = float(sebestoimost)                              # добавляем себестомость
            res = df_kal.merge(df_so_stavkami)                                         # мерджим фрейм со ставками
            res['ставка'] = (res['ставка']/365)/100                                    # получем коэф ставки на 1 день
            res['цена_простоя'] = df_kal['себестоимость'] * res['ставка']              # дена простоя 1 дня
            # display(res)
            return round(res['цена_простоя'].sum())                                          # результат всех дней простоя
        else:
            # date_start = date_start-timedelta(days=1)
            # df_kal = pd.DataFrame({'календарь': pd.date_range(date_end,  date_start)}) # создаем календарь 
            # df_kal['себестоимость'] = float(sebestoimost)                              # добавляем себестомость
            # res = df_kal.merge(df_so_stavkami)                                         # мерджим фрейм со ставками
            # res['ставка'] = (res['ставка']/365)/100                                    # получем коэф ставки на 1 день
            # res['цена_простоя'] = df_kal['себестоимость'] * res['ставка']              # дена простоя 1 дня
            # # display(res)
            # return -round(res['цена_простоя'].sum())                                          # результат всех дней простоя
            return 0 # Рома сходил к ЛР и ЛР сказал что это блох ловить возвращаем 0 / но рабочий код выше можно раскомментировать !!!!!!!!!!!!!!!
    except Exception as ex_:
        logger.error(f'ошибка функции {stavka_merge.__name__} дата_оплаты_счета: {date_start} дата_полной_оплаты_факт: {date_end} себестоимость: {sebestoimost} {ex_}')
        return 0


def return_result_prostoy(df, search_col_1:str, arg_search_1:str, search_col_2:str, arg_search_2:str, return_col:str):
    """возвращаем дату оплаты счета  по vin и дата_выдачи_факт
    """
    try:
        res = sum(list(df[(df[search_col_1]==arg_search_1)
                &(df[search_col_2]==arg_search_2)][return_col]))
        return res
    except Exception as ex_:
        logger.error(f'ошибка функции {return_result_prostoy.__name__}  {ex_} входные арг {search_col_1, arg_search_1, search_col_2, arg_search_2, return_col}')
        return 0

def return_result_oplata_scheta(df, search_col_1:str, arg_search_1:str, search_col_2:str, arg_search_2:str, return_col:str):
    """возвращаем дату оплаты счета  по vin и дата_выдачи_факт
    """
    try:
        res = list(df[(df[search_col_1]==arg_search_1)
                &(df[search_col_2]==arg_search_2)][return_col])
        if len(res) > 0:
            return res[0]
        else:
            return None
    except Exception as ex_:
        logger.error(f'ошибка функции {return_result_oplata_scheta.__name__}  {ex_} входные арг {search_col_1, arg_search_1, search_col_2, arg_search_2, return_col}')
        return None

def return_result_raznica_day(df, search_col_1:str, arg_search_1:str, search_col_2:str, arg_search_2:str, return_col:str):
    """возвращаем разницу дней по vin и дата_выдачи_факт
    """
    try:
        res = list(df[(df[search_col_1]==arg_search_1)
                &(df[search_col_2]==arg_search_2)][return_col])
        if len(res) > 0:
            return res[0]
        else:
            return 0
    except Exception as ex_:
        logger.error(f'ошибка функции {return_result_oplata_scheta.__name__}  {ex_} входные арг {search_col_1, arg_search_1, search_col_2, arg_search_2, return_col}')
        return 0

def stavka_treid_in(df, date, sebest, brend):
    try:
        if isinstance(date, str):
            date = datetime.strptime(date, "%Y-%m-%d")
        elif isinstance(date, (pd.Timestamp, datetime)):
            date = date
        if brend not in df['spec_for_brand'].unique():#бренд не найден
            # print('бренд не найден')
            res = df[((df['sebest_from']<=sebest)&(df['sebest_to']>=sebest)) 
                    & (df['spec_for_brand'].isna()) 
                    & (df['date_from']<=date) 
                    ]['date_from'].max()
            
            table = df[((df['sebest_from']<=sebest)&(df['sebest_to']>=sebest)) 
                    & (df['date_from']<=date)
                    & (df['date_from']==res)]
            
        elif brend in df['spec_for_brand'].unique(): #бренд найден
            # print('бренд найден')
            res = df[((df['sebest_from']<=sebest)&(df['sebest_to']>=sebest)) 
                    & (df['spec_for_brand']==brend) 
                    & (df['date_from']<=date) 
                    # & (df['date_to']<=date) 
                    ][['date_from', 'date_to']].max()
            res = max(list(res))
            if date<=res:
                # print('блок найденного бренда и дата поиска в диапазоне')
                table = df[((df['sebest_from']<=sebest)&(df['sebest_to']>=sebest)) 
                        & (df['date_from']<=date)
                        & (df['date_to']==res)]
            else:
                # print('блок найденного бренда но дата вне диапазона')
                res = df[((df['sebest_from']<=sebest)&(df['sebest_to']>=sebest)) 
                    & (df['spec_for_brand'].isna()) 
                    & (df['date_from']<=date) 
                    ]['date_from'].max()
                table = df[((df['sebest_from']<=sebest)&(df['sebest_to']>=sebest)) 
                    & (df['date_from']<=date)
                    & (df['date_from']==res)]
        res = list(table['stavka'])
        if len(res)>0:
            return res[0]*sebest # возвращаем результат себестомость * ставка
        else:
            return 0
    except Exception as ex_:
        logger.error(f'ошибка функции {stavka_treid_in.__name__}  {ex_} входные арг {date, sebest, brend}')
        return 0

def stavka_treid_in_list(df, date, sebest:list, brend):
    """рассчитывает правильную доходность то есть если за 1 новый отбали более 1 авто в трейдин считает отдельно зп каждый доход
    Args:
        df (_type_): фрейм с правилами ОВП от РОМЫ
        date (_type_): дата выдачи
        sebest (list): список сбестоимостей
        brend (_type_): бренд
    Returns:
        _type_: _description_
    """
    try:
        if len(sebest)>0:
            result = []
            for i in sebest:
                seb = float(i)
                if isinstance(date, str):
                    date = datetime.strptime(date, "%Y-%m-%d")
                elif isinstance(date, (pd.Timestamp, datetime)):
                    date = date
                if brend not in df['spec_for_brand'].unique():#бренд не найден
                    # print('бренд не найден')
                    res = df[((df['sebest_from']<=seb)&(df['sebest_to']>=seb)) 
                            & (df['spec_for_brand'].isna()) 
                            & (df['date_from']<=date) 
                            ]['date_from'].max()
                    
                    table = df[((df['sebest_from']<=seb)&(df['sebest_to']>=seb)) 
                            & (df['date_from']<=date)
                            & (df['date_from']==res)]
                    
                elif brend in df['spec_for_brand'].unique(): #бренд найден
                    # print('бренд найден')
                    res = df[((df['sebest_from']<=seb)&(df['sebest_to']>=seb)) 
                            & (df['spec_for_brand']==brend) 
                            & (df['date_from']<=date) 
                            # & (df['date_to']<=date) 
                            ][['date_from', 'date_to']].max()
                    res = max(list(res))
                    if date<=res:
                        # print('блок найденного бренда и дата поиска в диапазоне')
                        table = df[((df['sebest_from']<=seb)&(df['sebest_to']>=seb)) 
                                & (df['date_from']<=date)
                                & (df['date_to']==res)]
                    else:
                        # print('блок найденного бренда но дата вне диапазона')
                        res = df[((df['sebest_from']<=seb)&(df['sebest_to']>=seb)) 
                            & (df['spec_for_brand'].isna()) 
                            & (df['date_from']<=date) 
                            ]['date_from'].max()
                        table = df[((df['sebest_from']<=seb)&(df['sebest_to']>=seb)) 
                            & (df['date_from']<=date)
                            & (df['date_from']==res)]
                res = list(table['stavka'])
                
                if len(res)>0:
                    result.append(res[0]*seb) # возвращаем результат себестомость * ставка
                else:
                    result.append(0)
            return sum(result)
        else: return 0
    except Exception as ex_:
        logger.error(f'ошибка функции {stavka_treid_in_list.__name__}  {ex_} входные арг {date, sebest, brend}')
        return 0

def yesterday(days:int=1):
    """возвращает дату на вчера - по цморлчанию минус 1 день
    Args:
        days (int, optional): на сколько дней назад откатываемся по дате. Defaults to 1.
    Returns:
        _type_: _description_
    """
    try:
        from datetime import datetime, timedelta
        date = datetime.now()
        new_date = date - timedelta(days=days)# вычитание одного дня
        return new_date
    except Exception as ex_:
        logger.error(f'ошибка функции {yesterday.__name__}  {ex_}')

def return_fin_uslygi(df, vin, date_vidachi, name_return_column):
    """_summary_
    Args:
        df (_type_): фрейм с финуслугами из sql
        vin (_type_): vin номер авто
        date_vidachi (_type_): дата выдачи
        name_return_column (_type_): имя столбца значения которого надо вернуть
    Returns:
        _type_: _description_
    """
    try:
        res = sum(df[(df['VIN']==vin)&(df['Deregistration Date']==date_vidachi)][name_return_column])
        return res
    except Exception as ex_:
        logger.error(f'ошибка функции {return_fin_uslygi.__name__}  {ex_} входные арг {vin, date_vidachi, name_return_column}')
        return 0


def fackt_dohod_in_all_OVP(df, vin, search_date_kontrakta, serch_date_vidachy, return_name_column, flag_search_in_date=True):
    """поиск дохода за вто трейдин
    из КУМА подем VIN дату контракта и дату выдачи, если вин КУМА совпадает с ОВП принятым авто и дату контракта или дату выдачи совпадает 
    с дата_прихода ОВП возвращаем результат
    Args:
        df (_type_): врейм сборки ОВП МСК+ЯР
        vin (_type_): vin из КУМА который будеи скать в выгрузке ОВП
        search_date_kontrakta (_type_): дата контракта из КУМ
        serch_date_vidachy (_type_): дата выдачи из КУМ
        return_name_column (_type_): имя столбца - результат коотрого нужно вернуть
        flag_search_in_date (bool, optional): флаг поиска если True ищем с учетом дат, False - без учета дат только по VIN. Defaults to True.
    Returns:
        _type_: _description_
    """
    from datetime import datetime
    vin = str(vin).split('/')
    try:
        if flag_search_in_date:
            res = sum(df[(df['vin_нового'].str.contains('|'.join(vin), na=False)) & 
                    ((df['дата_прихода']>=search_date_kontrakta) & (df['дата_прихода']<=serch_date_vidachy))][return_name_column])
            return res
        else:
            res = sum(df[(df['vin_нового'].str.contains('|'.join(vin), na=False))][return_name_column])
            return res
    except Exception as ex_:
        logger.error(f'ошибка функции {fackt_dohod_in_all_OVP.__name__}  {ex_} входные арг {vin, search_date_kontrakta, serch_date_vidachy, return_name_column, flag_search_in_date}')
        return 0
    
def fackt_sebest_in_all_OVP_list(df, vin, search_date_kontrakta, serch_date_vidachy, return_name_column, flag_search_in_date=True):
    """поиск дохода за вто трейдин
    из КУМА подем VIN дату контракта и дату выдачи, если вин КУМА совпадает с ОВП принятым авто и дату контракта или дату выдачи совпадает 
    с дата_прихода ОВП возвращаем результат
    Args:
        df (_type_): врейм сборки ОВП МСК+ЯР
        vin (_type_): vin из КУМА который будеи скать в выгрузке ОВП
        search_date_kontrakta (_type_): дата контракта из КУМ
        serch_date_vidachy (_type_): дата выдачи из КУМ
        return_name_column (_type_): имя столбца - результат коотрого нужно вернуть
        flag_search_in_date (bool, optional): флаг поиска если True ищем с учетом дат, False - без учета дат только по VIN. Defaults to True.
    Returns:
        _type_: _description_
    """
    from datetime import datetime
    vin = str(vin).split('/')
    try:
        if flag_search_in_date:
            res = list(df[(df['vin_нового'].str.contains('|'.join(vin), na=False)) & 
                    ((df['дата_прихода']>=search_date_kontrakta) & (df['дата_прихода']<=serch_date_vidachy))][return_name_column])
            return res
        else:
            res = list(df[(df['vin_нового'].str.contains('|'.join(vin), na=False))][return_name_column])
            return res
    except Exception as ex_:
        logger.error(f'ошибка функции {fackt_sebest_in_all_OVP_list.__name__}  {ex_} входные арг {vin, search_date_kontrakta, serch_date_vidachy, return_name_column, flag_search_in_date}')
        return 0
    
def planiruemiy_dphod_all_OVP_list(df, vin, search_date_kontrakta, serch_date_vidachy, return_name_column_1, return_name_column_2, flag_search_in_date=True, fakt_realizacii=False):
    """поиск планируемого дохода за вто трейдин
    из КУМА подем VIN дату контракта и дату выдачи, если вин КУМА совпадает с ОВП принятым авто и дату контракта или дату выдачи совпадает 
    с дата_прихода ОВП вытаскиваем план_цена_продажи и цена_покупки и вычитаем
    Args:
        df (_type_): врейм сборки ОВП МСК+ЯР
        vin (_type_): vin из КУМА который будеи скать в выгрузке ОВП как вин_новгого
        search_date_kontrakta (_type_): дата контракта из КУМ
        serch_date_vidachy (_type_): дата выдачи из КУМ
        return_name_column_1 (_type_): имя столбца - из которого будем вычетать - план_цена_продажи
        return_name_column_2 (_type_): имя столбца - что будем вычетать - цена_покупки
        flag_search_in_date (bool, optional): флаг поиска если True ищем с учетом дат, False - без учета дат только по VIN. Defaults to True.
    Returns:
        _type_: _description_
    """
    konstant_percent = 0.95 # результат * на это значение
    from datetime import datetime
    vin = str(vin).split('/')
    try:
        if flag_search_in_date:
            if fakt_realizacii==False or str(fakt_realizacii)=='False': # если флаг False то ищем по дате фактической реализации
                res = df[(df['vin_нового'].str.contains('|'.join(vin), na=False)) & 
                        ((df['дата_прихода']>=search_date_kontrakta) & (df['дата_прихода']<=serch_date_vidachy)) & (df['план_цена_продажи']!=0)]
                if res.empty:
                    return 0
                else:
                    res_col_1 = sum(list(res[return_name_column_1]))
                    res_col_2 = sum(list(res[return_name_column_2]))
                    res = res_col_1-res_col_2
                    return res*konstant_percent
            else: return 0
        else:
            if fakt_realizacii==False or str(fakt_realizacii)=='False':
                res = df[(df['vin_нового'].str.contains('|'.join(vin), na=False))]
                if res.empty:
                    return 0
                else:
                    res_col_1 = sum(list(res[return_name_column_1]))
                    res_col_2 = sum(list(res[return_name_column_2]))
                    res = res_col_1-res_col_2
                    return res*konstant_percent
            else: return 0
    except Exception as ex_:
        logger.error(f'ошибка функции {planiruemiy_dphod_all_OVP_list.__name__}  {ex_} входные арг {vin, search_date_kontrakta, serch_date_vidachy, return_name_column_1, return_name_column_2, flag_search_in_date, fakt_realizacii}')
        return 0
    
def fackt_realizacii_auto_in_all_OVP(df, vin, search_date_kontrakta, serch_date_vidachy, return_name_column, flag_search_in_date=True):
    """поиск реализации авто трейдин - проверяет есть ли даты в столбце дата_полной_оплаты и проверяет их на datetime
    из КУМА подем VIN дату контракта и дату выдачи, если вин КУМА совпадает с ОВП принятым авто и дату контракта или дату выдачи совпадает 
    с дата_прихода ОВП возвращаем результат
    Args:
        df (_type_): врейм сборки ОВП МСК+ЯР
        vin (_type_): vin из КУМА который будеи скать в выгрузке ОВП
        search_date_kontrakta (_type_): дата контракта из КУМ
        serch_date_vidachy (_type_): дата выдачи из КУМ
        return_name_column (_type_): имя столбца - результат коотрого нужно вернуть
        flag_search_in_date (bool, optional): флаг поиска если True ищем с учетом дат, False - без учета дат только по VIN. Defaults to True.
    Returns:
        _type_: _description_
    """
    from datetime import datetime
    vin = str(vin).split('/')
    try:
        if flag_search_in_date:
            res = list(df[(df['vin_нового'].str.contains('|'.join(vin), na=False)) & 
                    ((df['дата_прихода']>=search_date_kontrakta) & (df['дата_прихода']<=serch_date_vidachy))][return_name_column]) # возможно придется добавить +5 дней к дате выдачи
            if len(res) > 0:
                return any([isinstance(i, (datetime)) and str(i) !='NaT' for i in res])
        else:
            res = list(df[(df['vin_нового'].str.contains('|'.join(vin), na=False))][return_name_column])
            if len(res) > 0:
                return any([isinstance(i, (datetime)) and str(i) !='NaT' for i in res])
            return res
    except Exception as ex_:
        logger.error(f'ошибка функции {fackt_dohod_in_all_OVP.__name__}  {ex_} входные арг {vin, search_date_kontrakta, serch_date_vidachy, return_name_column, flag_search_in_date}')
        return False
    
def vozvrat_dohoda_po_faktu_realizacii(dohod:float, fakt_realizacii:bool):
    """проверяем факт реализации и если да то подтверждаем доход
    Args:
        dohod (float): сумма дохода
        fakt_realizacii (bool): факт реализации
    Returns:
        _type_: _description_
    """
    try:
        if fakt_realizacii:
            return dohod
        elif fakt_realizacii==None:
            return 0
        else: return 0
    except Exception as ex_:
        logger.error(f'ошибка функции {vozvrat_dohoda_po_faktu_realizacii.__name__}  {ex_} входные арг {dohod, fakt_realizacii}')
        return 0

def sebest_iz_dvoh(sebest_prinvatogo, sebest_fact_iz_OVP):
    try:
        sebest_prinvatogo = float(sebest_prinvatogo)
        sebest_fact_iz_OVP = float(sebest_fact_iz_OVP)
        if sebest_fact_iz_OVP != 0:
            return sebest_fact_iz_OVP
        else: return sebest_prinvatogo
    except Exception as ex_:
        logger.error(f'ошибка функции {sebest_iz_dvoh.__name__}  {ex_} входные арг {sebest_prinvatogo, sebest_fact_iz_OVP}')


def pravka_dohoda_DO_OVP(dohod_do, dohod_rezina):
    """складывает доход_до и доход_резина - возвращает как доход_до - применять только к ОВП
    Args:
        dohod_do (_type_): _description_
        dohod_rezina (_type_): _description_
    Returns:
        _type_: _description_
    """
    try:
        if dohod_do==None:
            dohod_do=0
        if dohod_rezina==None:
            dohod_rezina=0
        result = dohod_do+dohod_rezina
        return result
    except Exception as ex_:
        logger.error(f'ошибка функции {pravka_dohoda_DO_OVP.__name__}  {ex_} входные арг {dohod_do, dohod_rezina}')
        return 0

def read_constants(name_constant:str, name_return_col:str):
    """возвращает значение из фрейма констант по имени константы
    Args:
        name_constant (str): имя константы
        name_return_col (str): имя столбца из еоторого возвращаем данные
    Returns:
        _type_: _description_
    """
    try:
        df = pd.read_excel(links_main(f'{DIR}\main_links.txt','constants'), engine='calamine', sheet_name='constants')
        df = list(df[df['name_constatn']==name_constant][name_return_col])
        if len(df)>0:
            return df[0]
        else: 
            logger.info(f"данных по запросу в константах - нет - будет передан аргумент 0 для {name_constant}")
            return 0
    except Exception as ex_:
        logger.error(f'ошибка функции {read_constants.__name__}  {ex_} входные арг {name_constant, name_return_col}')
        return 0

def read_email_users(sheet_name_, name_return_col:str):
    """возвращает значение из фрейма констант по имени константы
    Args:
        sheet_name_ (str): имя листа с которого берем данные
        name_return_col (str): имя столбца из которого возвращаем данные
    Returns:
        _type_: _description_
    """
    try:
        df = pd.read_excel(links_main(f'{DIR}\main_links.txt','email_users'), engine='calamine', sheet_name=sheet_name_)
        df = [str(i).strip() for i in list(df[name_return_col]) if str(i) != 'nan' and '@' in str(i)]
        if len(df)>0:
            return df
        else: 
            logger.info(f"email адресов нет для -  {name_return_col}")
            return None
    except Exception as ex_:
        logger.error(f'ошибка функции {read_email_users.__name__}  {ex_} входные арг {sheet_name_, name_return_col}')
        return 0

# отправка email в том числе со скрытыми получателями

def send_mail(send_to:list, 
              send_cc:list, 
              send_bcc:list, 
              topic_text:str, 
              body_text:str, 
              file_link:str, 
              file_name:str, 
              SEND_FROM:str, 
              SERVER:str, 
              PORT:int, 
              USER_NAME:str, 
              PASSWORD:str):
    """рассылка почты с вложением, пользователям в т.ч. добавление в копию и скрытую копию 

    Args:
        send_to (list):  список адресов для рассылки
        send_cc (list):  список адресов для рассылки - копии (можно не заполнять - проставить пустой список)
        send_bcc (list): список адресов для рассылки - скриыте копии (можно не заполнять - проставить пустой список)

        topic_text(str): текст темы писма
        body_text(str):  текст тела писма

        file_link(str):  ссылка на файл / если файла нет то оставить пустым file_name(str) = ''
        file_name(str):  имя файла в данном варианте нужно указывать с расширением 'BAIC_MSK.xlsx' (имя должно быть на латинице иначе придет в кодировке bin)
                         если файла нет то оставить пустым file_name(str) = ''

        SEND_FROM (str):  email пользователя от кого будет отправлено сообщение
        SERVER (str):     имя сервера
        PORT (int):       порт
        USER_NAME (str):  имя пользователя в сети
        PASSWORD (str):   пароль пользователя (учетной записи в сети)

        Пример:
        send_mail(['skrqqqo@siml-auto.ru'],             # send_to
          [],                                           # send_cc
          ['krutkosergey11111@yandex.ru'],              # send_bcc
          f'Привет привет',                             # topic_text
          f'Здравствуйте \nВо вложении файлик',         # body_text
          "//local/data/BAIC_MSK.xlsx",                 # file_link
          'BAIC_MSK.xlsx',                              # file_name
          'skrqqqo@siml-auto.ru',                       # SEND_FROM
          'server-vm23.LOCAL',                          # SERVER
          555,                                          # PORT
          'skrutko',                                    # USER_NAME
          'ZZZZZZZxxxxxx1111a')                         # PASSWORD

    """
    try:
        send_from = SEND_FROM                                                             
        subject = topic_text                                                               
        text = body_text                                                                 
        files = fr'{file_link.strip()}'
        server = SERVER
        port = PORT
        username=USER_NAME
        password = PASSWORD
        isTls=True

        msg = MIMEMultipart()
        msg['From'] = send_from
        msg['To'] = ','.join(send_to)
        msg["Cc"] = ','.join(send_cc)
        msg["Bcc"] = ','.join(send_bcc)
        msg['Date'] = formatdate(localtime = True)
        msg['Subject'] = subject
        msg.attach(MIMEText(text))

        part = MIMEBase('application', "octet-stream")
        if len(files) > 0: # если есть вложения то прикрепляем их к письму
            part.set_payload(open(files, "rb").read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', f'attachment; filename={file_name.strip()}') # имя файла должно быть на латинице иначе придет в кодировке bin
            msg.attach(part)

        smtp = smtplib.SMTP(server, port)
        if isTls:
            smtp.starttls()
        smtp.login(username, password)
        smtp.sendmail(send_from, send_to+send_cc+send_bcc, msg.as_string())
        smtp.quit()
    except Exception as ex_:
        logger.error(f'ошибка функции {send_mail.__name__} входне параметры {send_to, send_cc, send_bcc, topic_text, body_text, file_link, file_name, SEND_FROM, SERVER, PORT, USER_NAME, PASSWORD} {ex_}')

def update_file(link):
    """обновление сводной таблицы Excel
    # блок импортов для обновления сводных
    import pythoncom
    pythoncom.CoInitializeEx(0)
    import win32com.client
    Args:
        link (_type_): ссылка на файл - который нужно обновить
    """
    try:
        xlapp = win32com.client.DispatchEx("Excel.Application")
        wb = xlapp.Workbooks.Open(link)
        wb.Application.AskToUpdateLinks = False   # разрешает автоматическое  обновление связей (файл - парметры - дополнительно - общие - убирает галку запрашивать об обновлениях связей)
        wb.Application.DisplayAlerts = True  # отображает панель обновления иногда из-за перекрестного открытия предлагает ручной выбор обновления True - показать панель
        wb.RefreshAll()
        # xlapp.CalculateUntilAsyncQueriesDone() # удержит программу и дождется завершения обновления. было прописано time.sleep(30)
        time.sleep(40) # задержка 60 секунд, чтоб уж точно обновились сводные wb.RefreshAll() - иначе будет ошибка 
        wb.Application.AskToUpdateLinks = True   # запрещает автоматическое  обновление связей / то есть в настройках экселя (ставим галку обратно)
        wb.Save()
        wb.Close()
        xlapp.Quit()
        wb = None # обнуляем сслыки переменных иначе процесс эксель не завершается и висит в дистпетчере
        xlapp = None # обнуляем сслыки переменных иначе процесс эксел ь не завершается и висит в дистпетчере
        del wb # удаляем сслыки переменных иначе процесс эксель не завершается и висит в дистпетчере
        del xlapp # удаляем сслыки переменных иначе процесс эксель не завершается и висит в дистпетчере
    except Exception as ex_:
        wb.Close()
        xlapp.Quit()
        del wb # удаляем сслыки переменных иначе процесс эксель не завершается и висит в дистпетчере
        del xlapp # удаляем сслыки переменных иначе процесс эксель не завершается и висит в дистпетчере
        print(f'ошибка функции {update_file.__name__} {ex_} не удалось обновить файл по ссылке {link}')
        logger.error(f'ошибка функции {update_file.__name__} {ex_} не удалось обновить файл по ссылке {link}')

def mean_dohod_OVP(df, region_search:str, vid_postavki:str, diapozon_dayz_ago:int, return_col:str, start_fn_reverse:bool, planiruemiy_dohod_OVP:float):
    """возвращает средний доход по ОВП за период по региону
    Args:
        df (_type_): df сборки ОВП
        region_search (str): регион 'MSK'
        vid_postavki (str): вид_поставки 'трейдин'
        diapozon_dayz_ago (int): диапазон дней за который считаем средний доход 30 значит последние 30 дней
        return_col (str): по какому столбцу возвращаем среднее значение
        start_fn_reverse (bool): условие должно быть False - для старта выаолнения функции
    Returns:
        _type_: float
    """
    konstant_percent = 0.9
    try: 
        if (start_fn_reverse==False or str(start_fn_reverse)=='False') and float(planiruemiy_dohod_OVP)==0:
            res = df[(df['region']==region_search) & 
                    ((df['дата_выдачи']>=yesterday(diapozon_dayz_ago)) & (df['дата_выдачи']<=yesterday())) & (df['вид_поставки']==vid_postavki)]#[return_col].mean()
            if res.empty:
                print('Нет данных')
                return 0
            else: 
                res = res[return_col].mean()
                return float(res)*konstant_percent # 90% от среднего
        else: return 0
    except Exception as e:
        print(f'ошибка функции {mean_dohod_OVP.__name__} - {e} входные параметры {region_search, vid_postavki, diapozon_dayz_ago, return_col, start_fn_reverse, planiruemiy_dohod_OVP}') 

def read_constants(name_constant:str, name_return_col:str):
    """возвращает значение из фрейма констант по имени константы
    Args:
        name_constant (str): имя константы
        name_return_col (str): имя столбца из еоторого возвращаем данные
    Returns:
        _type_: _description_
    """
    try:
        df = pd.read_excel(links_main(fr'{DIR}\main_links.txt','constants'), sheet_name='constants')
        df = list(df[df['name_constatn']==name_constant][name_return_col])
        if len(df)>0:
            return df[0]
        else: 
            logger.info(f"данных по запросу в константах - нет - будет передан аргумент 0 для {name_constant}")
            return 0
    except Exception as ex_:
        logger.error(f'❌ ошибка функции {read_constants.__name__}  {ex_} входные арг {name_constant, name_return_col}')
        return 0

logger.info(f'считываем email-ы пользователей для отправки КУМ общего и отдельно КУМ_ДИР')  
try: 
    KUM_USER_EMAIL = read_email_users('user_kum', 'email')
    KUM_USER_EMAIL_CC = read_email_users('user_kum', 'email_cc')
    KUM_USER_EMAIL_BCC = read_email_users('user_kum', 'email_bcc')

    KUM_DIR_USER_EMAIL = read_email_users('user_kum_dir', 'email')
    KUM_DIR_USER_EMAIL_CC = read_email_users('user_kum_dir', 'email_cc')
    KUM_DIR_USER_EMAIL_BCC = read_email_users('user_kum_dir', 'email_bcc')
except Exception as ex_:
        logger.error(f'не удалось считать email-ы {ex_}')


logger.info("Получаем парметры пользователя для работы с почтой")
try:
    NAME_PC, SERVER, PORT, USER_NAME, SEND_FROM, PASSWORD = get_data_user_param()
except Exception as ex_:
    logger.error(f"Ошибка получения данных: {ex_}")

logger.info("Проверка полученных параметров")

resul_mail_param = all([i is not None for i in [NAME_PC, SERVER, PORT, USER_NAME, SEND_FROM, PASSWORD]])

if resul_mail_param:
    logger.info(f"Все данные для работы с почтой присутствуют {resul_mail_param}")
else:
    logger.error(f"Некоторые данные для работы с почтой отсутствуют {resul_mail_param} {get_data_user_param()}")


logger.info(f'считываем данные подключения SQL')
try:
    SERVER_SQL = links_main(fr"{DIR}\main_links.txt", "server_sql")
    DATABASE_SQL = links_main(fr"{DIR}\main_links.txt", "database_sql")
    USERNAME_SQL = links_main(fr"{DIR}\main_links.txt", "username_sql")
    PASSWORD_SQL = links_main(fr"{DIR}\main_links.txt", "password_sql")
except Exception as ex:
    logger.error(f'не удалось получить данные подключения SQL')

logger.info(f'считываем константы')
try:
    NACENKA_BANKA_K_STAVKE_REF = float(read_constants('NACENKA_BANKA_K_STAVKE_REF', 'result'))
    KOEF_FIN_USLUGI = float(read_constants('KOEF_FIN_USLUGI', 'result'))
    KOEF_VOSSTANOVLENIYA_PROGR_PRIVILEGIY = float(read_constants('KOEF_VOSSTANOVLENIYA_PROGR_PRIVILEGIY', 'result'))
    YEAR_COSTRACII = int(read_constants('YEAR_COSTRACII', 'result'))
    LIST_YUR_LIC = [str(i).strip() for i in read_constants('LIST_YUR_LIC', 'result').split(',')] # юр лица для рассчета перемещения
except Exception as ex:
    logger.error(f'не удалось считать константы из файла')


# новый запрос от Леонида
logger.info(f'считываем данные ДКИС МСК из SQL по запросу Леонида')   
try:
    start_day_dkis_sql = '01'
    start_month_dkis_sql = '01'
    start_year_dkis_sql = '2024'
    end_day_dkis_sql = yesterday().strftime("%d")
    end_month_dkis_sql = yesterday().strftime("%m")
    end_year_dkis_sql = yesterday().strftime("%Y")
    sql_query = f"""  
    SELECT [Order Date]
                    ,[Deregistration Date]
                    ,[Description]
                    ,[VIN]
                    ,[KL_NAIM1]
                    ,[АвтоКаско]
                    ,[ОCАГО]
                    --,[UR_FIZ]
                    --,[UR_LIC]
                    ,[Kredit]
                    ,'URLIC'= [UR_LIC]
    ,[Region_ZP]
        ,[Region]
                    ,[Name]
        ,[Podval]
                    ,[NAIM_BANK]
                    ,[BANKCOUNT]
                    ,[REGI] 
        ,[Comment]                  
                    --,[KOD_MENEDZ2]
                    ,[NOM_POL]
                    ,[FIO_MEN]
    -- ,[VINP]
        ,[Real_Date]
        ,[Department Code]
    ,[nomer]
                    ,[Make Code]
                    ,[NPOLKASKO]
                    ,[NPOLOSAGO]
    ,[Document No_],sumkomis,KBraz,KBKBzaSJ,fullKV,sumkomis*[АвтоКаско] as KVkasko,sumkomis*[ОCАГО] as KVosago,KV_DP,cliphone,KBSJfact,sumgap,sumassist,sumpriv,sumkasko,sumosago,IsPriviledge,sumporuch,sumsumgap,strnaimosago,strnaimkasko,strnaimgap
    FROM(SELECT * FROM [Q_OPPA] ( '{start_month_dkis_sql}.{start_day_dkis_sql}.{start_year_dkis_sql}' , '{end_month_dkis_sql}.{end_day_dkis_sql}.{end_year_dkis_sql}' , '' ))as tbb
    LEFT OUTER JOIN [OPPA] ON tbb.VIN collate Cyrillic_General_CI_AS = [OPPA].[VINP]
                        """
    df_dkis_MSK = pd.read_sql(sql_query, connect_bd_SQL(SERVER_SQL, 'StrahovkaSQL', USERNAME_SQL, PASSWORD_SQL))

except Exception as ex_:
        logger.error(f'не удалось считать данные ДКИС МСК из SQL по запросу Леонида {ex_}')


# новый запрос от Леонида
logger.info(f'считываем данные ДКИС ЯР из SQL по запросу Леонида')   
try:
    start_day_dkis_sql = '01'
    start_month_dkis_sql = '01'
    start_year_dkis_sql = '2024'
    end_day_dkis_sql = yesterday().strftime("%d")
    end_month_dkis_sql = yesterday().strftime("%m")
    end_year_dkis_sql = yesterday().strftime("%Y")
    sql_query = f"""  
    SELECT [Order Date]
                    ,[Deregistration Date]
                    ,[Description]
                    ,[VIN]
                    ,[KL_NAIM1]
                    ,[АвтоКаско]
                    ,[ОCАГО]
                    --,[UR_FIZ]
                    --,[UR_LIC]
                    ,[Kredit]
                    ,'URLIC'= [UR_LIC]
    ,[Region_ZP]
        ,[Region]
                    ,[Name]
        ,[Podval]
                    ,[NAIM_BANK]
                    ,[BANKCOUNT]
                    ,[REGI] 
        ,[Comment]                  
                    --,[KOD_MENEDZ2]
                    ,[NOM_POL]
                    ,[FIO_MEN]
    -- ,[VINP]
        ,[Real_Date]
        ,[Department Code]
    ,[nomer]
                    ,[Make Code]
                    ,[NPOLKASKO]
                    ,[NPOLOSAGO]
    ,[Document No_],sumkomis,KBraz,KBKBzaSJ,fullKV,sumkomis*[АвтоКаско] as KVkasko,sumkomis*[ОCАГО] as KVosago,KV_DP,cliphone,KBSJfact,sumgap,sumassist,sumpriv,sumkasko,sumosago,IsPriviledge,sumporuch,sumsumgap,strnaimosago,strnaimkasko,strnaimgap
    FROM(SELECT * FROM [Q_OPPA] ( '{start_month_dkis_sql}.{start_day_dkis_sql}.{start_year_dkis_sql}' , '{end_month_dkis_sql}.{end_day_dkis_sql}.{end_year_dkis_sql}' , '' ))as tbb
    LEFT OUTER JOIN [OPPA] ON tbb.VIN collate Cyrillic_General_CI_AS = [OPPA].[VINP]
                        """
    df_dkis_YAR = pd.read_sql(sql_query, connect_bd_SQL("SRV-Y03", 'Strahovka2016', USERNAME_SQL, PASSWORD_SQL))

except Exception as ex_:
        logger.error(f'не удалось считать данные ДКИС ЯР из SQL по запросу Леонида {ex_}')


logger.info(f'объединяем фреймы финуслуг ЯР и МСК из SQL в один') 
try:
    df_dkis = pd.concat([df_dkis_MSK, df_dkis_YAR])
except Exception as ex_:
        logger.error(f'не удалось объединить фреймы финуслуг ЯР и МСК из SQL в один {ex_}')


logger.info(f"запускаем субпроцесс - сбора ОВП ЯР и МСК")
try:
    update_task(links_main(f'{DIR}\main_links.txt','run_sub_sbor_OVP_ALL'))
except Exception as e:
    logger.error(f"ошибка в субпроцессе - сбора ОВП ЯР и МСК {e}")



logger.info(f"считываем result_OVP_all из SQL")
try:
    df_OVP_all = pd.read_sql('select * from result_OVP_all', connect_bd_SQL(SERVER_SQL, DATABASE_SQL, USERNAME_SQL, PASSWORD_SQL))
except Exception as e:
    logger.error(f"ошибка - при считывании result_OVP_all из SQL {e}")


logger.info(f"запускаем субпроцесс - считывания ставок")
try:
    update_task(links_main(f'{DIR}\main_links.txt','run_sub_stavki_ref_2'))
except Exception as e:
    logger.error(f"ошибка в субпроцессе - считывания ставок {e}")


# NACENKA_BANKA_K_STAVKE_REF = 0
logger.info(f"считываем result_stavka_ref из SQL и применяем повышение ставки + {NACENKA_BANKA_K_STAVKE_REF}")
try:
    df_stavka_ref = pd.read_sql('select * from result_stavka_ref', connect_bd_SQL(SERVER_SQL, DATABASE_SQL, USERNAME_SQL, PASSWORD_SQL))
    df_stavka_ref['ставка'] = df_stavka_ref['ставка'].apply(lambda x: x+NACENKA_BANKA_K_STAVKE_REF)
except Exception as e:
    logger.error(f"ошибка - при считывании result_stavka_ref из SQL {e}")


logger.info(f"объединяем result_np_auto и result_sclad в SQL и получем табличку для расчета дат простоя")
query_q = """
SELECT  sklad.vin as scl_vin, sklad.себестоимость_ам, sklad.цена_продажи, np.дата_выдачи_факт, np.дата_полной_оплаты_факт, sklad.дата_оплаты_счета, DATEDIFF(day, sklad.дата_оплаты_счета,np.дата_полной_оплаты_факт) AS разница_дней
FROM result_sclad AS sklad
LEFT JOIN result_np_auto AS np ON sklad.vin = np.vin and sklad.дата_прихода_на_склад = np.дата_прихода_на_склад
WHERE sklad.vin <> '0'  and np.дата_выдачи_факт IS NOT NULL
"""
try:
    resul_slivaniva = pd.read_sql(query_q, connect_bd_SQL(SERVER_SQL, DATABASE_SQL, USERNAME_SQL, PASSWORD_SQL))
    logger.info(f"результат получен и сохранен в переменную - resul_slivaniva")
except Exception as e:
    logger.error(f"ошибка - при объединении result_np_auto и result_sclad в SQL и получении таблички для расчета дат простоя {e}")


logger.info(f'считываем данные подключения SQL')
try:
    df_otchisleniva_ovp_pravilo = pd.read_excel(links_main(fr"{DIR}\main_links.txt", "df_otchisleniva_ovp_pravilo"), engine='calamine',sheet_name='otchisl_ovp')
    df_otchisleniva_ovp_pravilo = df_otchisleniva_ovp_pravilo[df_otchisleniva_ovp_pravilo['date_from'].notna()]
except Exception as ex:
    logger.error(f'не удалось получить данные подключения SQL')

logger.info(f"считаем цену простоя")
try:
    resul_slivaniva['цена_простоя'] = resul_slivaniva.apply(lambda x: stavka_merge(x.дата_оплаты_счета, x.дата_полной_оплаты_факт, x.себестоимость_ам, df_stavka_ref), axis=1)
except Exception as e:
    logger.error(f"ошибка - при расчете цены простоя {e}")


# книга в которой все ссылки и логика
# 'db' - на этом листе все ссылки и пароли и указазанием какие лситы нужно брать в книге по ссылке 
# все последующие листы - это логика, названия соответсвуют именам листов в кумах и расписаны столбцы которые нужно брать и какое эталонное имя присваивать
logger.info(f'считываем фрейм с ссылками и логикой')
try:
    df_links_name_file = pd.read_excel(links_main(fr"{DIR}\main_links.txt", "links_name_file"), engine='calamine', sheet_name=None)    # вся книга со всеми листами как словарь
    df_main_db = df_links_name_file['db']                                                                           # главный фрейм с ссылками
    main_list = [i for i in df_links_name_file.keys() if i != 'db']                                                 # список листов без главного листа со всеми ссылками
except Exception as ex:
    logger.error(f'не удалось фрейм с ссылками и логикой')

class Masturbation:
    def __init__(self, link):   
        self.link:str = str(link).strip()   # ссылка на файл
        self.name:str = None                # имя файла 
        self.password:str = None            # пароль
        self.region:str = None              # регион
        self.marka:str = None               # марка
        self.marka_2:str = None             # марка 2
        self.work_sheet:list = None         # рабочие листы которые будем обрабатывать
        self.path_file_ok:bool = None       # путь к файлу - существует ли он
        self.df_work_all = None             # считанный df по ссылке со всеми листами
        self.df_work_name_lst = None        # считанный из df по ссылке имена листов
        self.all_frames:list = []           # все собранные фреймы с листов
        self.df_sborka = None               # собранный фрейм все в один фрейм
        self.date_update = None             # дата обновления файла
        self.fn_starter()

    def fn_return_params_in_link(self):
        "возвращает параметры по ссылке имя пароль и рабочие листы"
        logger.info(f'возвращаем параметры по ссылке имя пароль и рабочие листы')
        try:
            self.name = return_params_in_link(df_main_db, 'link', self.link, 'name')
            self.password = str(return_params_in_link(df_main_db, 'link', self.link, 'pass'))
            self.region = str(return_params_in_link(df_main_db, 'link', self.link, 'region'))
            self.marka = str(return_params_in_link(df_main_db, 'link', self.link, 'marka'))
            self.marka_2 = str(return_params_in_link(df_main_db, 'link', self.link, 'marka_2'))
            self.work_sheet = [i.strip() for i in return_params_in_link(df_main_db, 'link', self.link, 'kum_work_sheet').split(',')]
        except Exception as e:
            logger.error(f'ошибка класса {Masturbation.__name__} функции {self.fn_return_params_in_link.__name__} - {e}')

    def fn_return_path_file_ok(self):
        """проверяет актуален ли путь к файлу
        """
        logger.info(f'проверяем актуален ли путь к файлу')
        try:
            self.path_file_ok = os.path.isfile(self.link)
        except Exception as e:
            logger.error(f'ошибка класса {Masturbation.__name__} функции {self.fn_return_path_file_ok.__name__} - {e}')

    def fn_read_file(self):
        """считываем кум по ссылке всю книгу со всеми листами"""
        logger.info(f'считываем кум по ссылке всю книгу со всеми листами')
        try:
            self.df_work_all, self.df_work_name_lst = open_dataframe(self.link, self.password, lst_name = None)
        except Exception as e:
            logger.error(f'ошибка класса {Masturbation.__name__} функции {self.fn_read_file.__name__} - {e}')

    def df_sborka_kum(self):
        """основная обратока фрейма
        забираем только необходимые колонки и строим эталонный фрейм
        """
        logger.info(f'основная обратока фрейма')
        try:
            for name_list in self.work_sheet:                                                                   # проходим по рабочим листам в книге и грузим только их
                df_original = df_links_name_file[name_list]                                                     # грузим лист эталоннного фрейма
                df_original_list = df_original[df_original['link']==self.link]                                  # фильтруем по входящей ссылке какие столбцы брать из оригинала и как называть эталонно
                if len(df_original_list) > 0:                                                                   # если есть совпадения по ссылке
                    df_original_list = df_original_list[[i for i in df_original_list.columns if i != 'link']]   # убираем колонку с ссылкой получаем фрейм с именами столбцов которые берем и эталонные названия
                    df_iz_knigi_po_imeni_lista = Shapka(self.df_work_all[name_list])                            # загружаем фрейм по имени листа который будем обрабатывать
                    for i in df_original_list.columns:                                                          # проходим по колонкам в эталонном фрейме
                        etalon_name_col = i                                                                     # имя колонки в эталонном фрейме
                        df_original_name_col = list(df_original_list[i])[0]                                     # имя колонки в оригинальном фрейме
                        df_iz_knigi_po_imeni_lista.rename(columns={df_original_name_col:etalon_name_col}, inplace=True)       # переименовываем колонку в фрейме по имени листа с оригинала на эатлон
                    list_etallonih_imen_stolbcoy = list(df_original_list.columns)                               # список эталонных имен столбцов
                    df_iz_knigi_po_imeni_lista = df_iz_knigi_po_imeni_lista[[i for i in df_iz_knigi_po_imeni_lista.columns if i in list_etallonih_imen_stolbcoy]] # фильтруем фрейм по эталонным именам столбцов
                    for i in list_etallonih_imen_stolbcoy:                                                      # проходим по эталонным именам столбцов
                        if i not in df_iz_knigi_po_imeni_lista.columns:                                         # если имя столбца в эталонном списке не найдено в фрейме по имени листа
                            df_iz_knigi_po_imeni_lista[i] = '0'                                                 # добавляем столбец с нулями
                    df_iz_knigi_po_imeni_lista = df_iz_knigi_po_imeni_lista[df_original_list.columns]           # сортируем по эталонным именам столбцов
                    df_iz_knigi_po_imeni_lista['link'] = self.link                                              # добавляем колонку с ссылкой на файл для последующей фильтрации по ссылке ['link']
                    df_iz_knigi_po_imeni_lista['name_list'] = name_list                                         # добавляем колонку с именем листа для последующей фильтрации по имени листа ['name_list]
                    self.all_frames.append(df_iz_knigi_po_imeni_lista)                                          # добавляем в список фреймов
                else:
                    logger.error(f'нет совпадений по ссылке {self.link} в эталонной таблице на листе {name_list} функция {self.df_sborka_kum.__name__}')
        except Exception as e:
            logger.error(f'ошибка класса {Masturbation.__name__} функции {self.df_sborka_kum.__name__} - {e}')


    def fn_concat_all_frames(self):
        """объединяем все фреймы в один фрейм"""
        logger.info(f'объединяем все фреймы в один фрейм')
        try:
            if len(self.all_frames) > 0:

                self.df_sborka = pd.concat(self.all_frames)
        except Exception as e:
            logger.error(f'ошибка класса {Masturbation.__name__} функции {self.fn_concat_all_frames.__name__} - {e}')


    def fn_date_update(self):
        """дата обновления файла"""
        logger.info(f'дата обновления файла')
        try:
            self.date_update = file_update(self.link)
        except Exception as e:
            logger.error(f'ошибка класса {Masturbation.__name__} функции {self.fn_date_update.__name__} - {e}')


    def fn_dop_columns(self):
        """добавляем дополнительные колонки"""
        logger.info(f'добавляем дополнительные колонки')
        try:
            if self.df_sborka is not None:
                self.df_sborka['name'] = self.name
                self.df_sborka['marka'] = self.marka
                self.df_sborka['marka_2'] = self.marka_2
                self.df_sborka['region'] = self.region
                self.df_sborka['date_update'] = self.date_update
        except Exception as e:
            logger.error(f'ошибка класса {Masturbation.__name__} функции {self.fn_dop_columns.__name__} - {e}')



    def fn_starter(self):
        self.fn_return_params_in_link()
        self.fn_return_path_file_ok()
        self.fn_read_file()
        self.df_sborka_kum()
        self.fn_concat_all_frames()
        self.fn_date_update()
        self.fn_dop_columns()


logger.info('создаем объекты класса Masturbation')
dct_object_Masturbation = {}
exception_links = []
for i in df_main_db.link:
    logger.info(f'обрабатываем ссылку {i}')
    try:
        dct_object_Masturbation[return_params_in_link(df_main_db, 'link', i, 'name')] = Masturbation(i)                            # создаем объект класса
        logger.info(f"создан объект класса {return_params_in_link(df_main_db, 'link', i, 'name')} по ссылке {i}")
    except Exception as ex:
        logger.error(f"не удалось создать объект класса {return_params_in_link(df_main_db, 'link', i, 'name')} по ссылке {i} ссылка будет добавлена в лист повторной попытки считвания")
        exception_links.append(i)

if len(exception_links) > 0:
    logger.info('повторная попытка создать объекты класса Masturbation из ссылок с ишибками')
    for i in  exception_links:
        try:
            dct_object_Masturbation[return_params_in_link(df_main_db, 'link', i, 'name')] = Masturbation(i)                            # создаем объект класса
            logger.info(f"создан объект класса {return_params_in_link(df_main_db, 'link', i, 'name')} по ссылке {i}")
        except Exception as ex:
            logger.error(f"не удалось создать объект класса {return_params_in_link(df_main_db, 'link', i, 'name')} по ссылке {i} ссылка будет добавлена в лист повторной попытки считвания")

logger.info(f'создано объектов: {len(dct_object_Masturbation)}')
logger.info(f'список объектов: {list(dct_object_Masturbation.keys())}')


logger.info(f'сборка всех объектов класса: {Masturbation.__name__} в один файл')
try:
    df_kum_sborka_Masturbation =  pd.concat([dct_object_Masturbation[i].df_sborka for i in dct_object_Masturbation.keys()])
except Exception as e:
    logger.error(f'ошибка сборки всех объектов класса: {Masturbation.__name__} в один файл- {e}')


logger.info(f'промежуточное сохранение сборки объектов класса: {Masturbation.__name__} в excel')
try:
    df_kum_sborka_Masturbation.to_excel(links_main(fr"{DIR}\main_links.txt", 'save_class_1_Masturbation'), index=False)
except Exception as e:
    logger.error(f'ошибка промежуточного сохранение сборки объектов класса: {Masturbation.__name__} в excel {e}')

class Pravka:
    def __init__(self, name, oblect_cl_Masturbation):
        self.name = copy.deepcopy(name)
        self.object_cl = copy.deepcopy(oblect_cl_Masturbation)
        self.region = self.object_cl.region
        self.date_update = self.object_cl.date_update
        self.df_sborka = self.object_cl.df_sborka
        self.fn_starter()
    

    def fn_change_column_type(self):
        """изменяем тип столбцов"""
        try:
            logger.info(f'изменяем тип столбцов - float')
            for i in self.df_sborka.columns:
                for name_prefix in ['авто_', 'до_', 'фин_', 'трейдин_', 'доход_']:
                    if name_prefix in i:
                        try:
                            self.df_sborka[i] = self.df_sborka[i].apply(lambda x: x if is_number(x) else 0)
                            self.df_sborka[i] = self.df_sborka[i].astype('float')
                        except Exception as e:
                            logger.error(f'ошибка класса {Pravka.__name__} функции {self.fn_change_column_type.__name__} столбец {i} правка типа на float - {e}')

            logger.info(f'изменяем тип столбцов - datetime')
            for i in self.df_sborka.columns:
                if 'дата_' in i or 'date_update' in i:
                    try:
                        self.df_sborka[i] = pd.to_datetime(self.df_sborka[i], format='mixed', errors='coerce')
                    except Exception as e:
                        logger.error(f'ошибка класса {Pravka.__name__} функции {self.fn_change_column_type.__name__} столбец {i} правка типа на datetime - {e}')
        except Exception as e:
            logger.error(f'ошибка класса {Pravka.__name__} функции {self.fn_change_column_type.__name__} - {e}')

    def fn_strip(self):
        """удаляем пробелы в столбцах"""
        try:
            logger.info(f'удаляем пробелы в строковых значениях столбцов')
            for i in self.df_sborka.columns:
                for name_column in ['vin', 'клинет', 'модель', 'продавец', 'салон','физ_юр', 'name', 'marka', 'marka_2', 'region']:
                    if name_column == i:
                        try:
                            self.df_sborka[i] = self.df_sborka[i].apply(lambda x: str(x).strip())
                        except Exception as e:
                            logger.error(f'ошибка класса {Pravka.__name__} функции {self.fn_strip.__name__} -удаление пробела в стобце {i} - {e}')               
        except Exception as e:
            logger.error(f'ошибка класса {Pravka.__name__} функции {self.fn_strip.__name__} - {e}')

    def clear_rows_frame(self):
        """удаляем пустые строки - очищаем фреймы""" 
        try:
            logger.info(f'удаляем пустые строки по - дата_выдачи')
            self.df_sborka = self.df_sborka[self.df_sborka['дата_выдачи'].notna()]
        except Exception as e:
            logger.error(f'ошибка класса {Pravka.__name__} функции {self.clear_rows_frame.__name__} - {e}')

    def fn_fillna_go(self):
        """заменяем na на 0
        """
        try:
            lst_col = ['авто_', 'фин_', 'трейдин', 'до_','фин_', 'доход_']
            logger.info(f'заменяем na на 0 в столбцах имеющих префиксы {lst_col}')
            for i in self.df_sborka.columns:
                for name_prefix in lst_col:
                    if name_prefix in i:
                        self.df_sborka[i] = self.df_sborka[i].fillna(0)
        except Exception as e:
            logger.error(f'ошибка класса {Pravka.__name__} функции {self.fn_fillna_go.__name__} - {e}')

    def fn_MAZDA_next(self):
        """выделяем мазду некст - если она с листа некст"""
        try:
            if 'MAZDA'.lower() in str(self.name).lower() and 'MSK' in self.region:
                logger.info(f'выделяем мазду некст')
                self.df_sborka['marka_2'] = self.df_sborka.apply(lambda x: mazda_next_msk(x.marka_2, x.region, x.name_list), axis=1)
        except Exception as e:
            logger.error(f'ошибка класса {Pravka.__name__} функции {self.fn_MAZDA_next.__name__} - {e}')


    def fn_OVP_ryb_yar(self):
        """разделяем ОВП ЯР РЫБ по регионам""" 
        try:
            if 'OVP'.lower() in str(self.name).lower() and 'YAR' in self.region:
                logger.info(f'разделяем ОВП ЯР РЫБ по регионам')
                self.df_sborka['region'] = self.df_sborka.apply(lambda x: razdelenie_OVP_yar_region(x.marka_2, x.region, x.салон), axis=1)
        except Exception as e:
            logger.error(f'ошибка класса {Pravka.__name__} функции {self.fn_OVP_ryb_yar.__name__} - {e}')


    def fn_kre_nal(self):
        """разделяем форму оплаты на кре нал""" 
        try:
            logger.info(f'разделяем форму оплаты на кре нал')
            self.df_sborka['форма_оплаты'] = self.df_sborka.apply(lambda x: form_pay(x.форма_оплаты), axis=1)
        except Exception as e:
            logger.error(f'ошибка класса {Pravka.__name__} функции {self.fn_kre_nal.__name__} - {e}')

    def fn_kre_nal_OVP_in_komment(self):
        """ разделяем форму оплаты на кре нал по ОВП из комментария""" 
        try:
            if 'OVP'.lower() in str(self.name).lower():
                self.df_sborka['форма_оплаты'] = self.df_sborka.apply(lambda x: forma_oplaty_OVP_in_kommentary(x.форма_оплаты, x.комментарий), axis=1)
        except Exception as e:
            logger.error(f'ошибка класса {Pravka.__name__} функции {self.fn_kre_nal_OVP_in_komment.__name__} - {e}')

    def fn_fiz_yur(self):
        """разделяем физ юрл""" 
        try:
            logger.info(f'разделяем физ юрл')
            self.df_sborka['физ_юр'] = self.df_sborka.apply(lambda x: fiz_yur(x.физ_юр), axis=1)
            logger.info(f'разделяем физ юрл - stage_2')
            self.df_sborka['физ_юр'] = self.df_sborka.apply(lambda x: fiz_yur_stage_2(x.физ_юр, x.клиент), axis=1)
        except Exception as e:
            logger.error(f'ошибка класса {Pravka.__name__} функции {self.fn_fiz_yur.__name__} - {e}')

    def fn_korp_demo(self):
            """добавляем столбце демо и проставляем статус""" 
            try:
                logger.info(f'добавляем колонку демо и проставляем статус 1/0')
                self.df_sborka['демо'] = self.df_sborka.apply(lambda x: korp_demo(x.name_list), axis=1)
            except Exception as e:
                logger.error(f'ошибка класса {Pravka.__name__} функции {self.fn_korp_demo.__name__} - {e}')

    def fn_treid_in_and_fact_am(self):
            """добавляем колонку ти_факт_ам и проставляем статус 1/0""" 
            try:
                logger.info(f'добавляем колонку ти_факт_ам и проставляем статус 1/0')
                self.df_sborka['ти_факт_ам'] = self.df_sborka['трейдин_доход'].apply(lambda x: treid_in_auto(x))
                logger.info(f'добавляем колонку факт_ам и проставляем статус 1')
                self.df_sborka['факт_ам'] = 1
            except Exception as e:
                logger.error(f'ошибка класса {Pravka.__name__} функции {self.fn_treid_in_and_fact_am.__name__} - {e}')

    def fn_pravka_dohoda_do_OVP(self):
        if 'OVP'.lower() in str(self.name).lower():
            self.df_sborka['до_доход'] = self.df_sborka.apply(lambda x: pravka_dohoda_DO_OVP(x.до_доход, x.до_втч_резина), axis=1)

    def fn_peremeschenie(self):
        logger.info(f'🔄{self.fn_peremeschenie.__name__} - расчитываем перемещение')
        try:
            lst_mask1 = ['OOO','ООО']
            mask1 = (self.df_sborka['клиент'].str.contains("|".join(lst_mask1), case=False, na=False))
            lst_mask2 = LIST_YUR_LIC # заполняется в эксель
            mask2 = (self.df_sborka['клиент'].str.contains("|".join(lst_mask2), case=False, na=False))
            mask_all = [mask1 & mask2]
            choices = [1]
            self.df_sborka['исключ_из_кум'] = np.select(mask_all, choices, default=0)
        except Exception as ex_:
            logger.error(f'❌ ошибка {self.fn_peremeschenie.__name__} {ex_}')

    


    def fn_starter(self):
        self.fn_change_column_type()
        self.fn_strip()
        self.clear_rows_frame()
        self.fn_fillna_go()
        self.fn_MAZDA_next()
        self.fn_OVP_ryb_yar()
        self.fn_kre_nal()
        self.fn_kre_nal_OVP_in_komment()
        self.fn_fiz_yur()
        self.fn_korp_demo()
        self.fn_treid_in_and_fact_am()
        self.fn_pravka_dohoda_do_OVP()
        self.fn_peremeschenie()
        
    
logger.info(f'создание объектов класса: {Pravka.__name__} на основе объектов класса: {Masturbation.__name__}')
dct_object_Pravka = {}
try:
    for name_obj in dct_object_Masturbation.keys():
        try:
            logger.info(f'создание объекта класса: {Pravka.__name__} с именем: {name_obj}')
            dct_object_Pravka[name_obj] = Pravka(name_obj, dct_object_Masturbation[name_obj])
        except Exception as e:
            logger.error(f'ошибка - создания объекта класса: {Pravka.__name__} с именем: {name_obj}')
except Exception as e:
    logger.error(f'создание объектов класса: {Pravka.__name__} на основе объектов класса: {Masturbation.__name__}')    

logger.info(f'сборка всех объектов класса: {Pravka.__name__} в один файл')
try:
    df_kum_sborka_Pravka =  pd.concat([dct_object_Pravka[i].df_sborka for i in dct_object_Pravka.keys()])
except Exception as e:
    logger.error(f'ошибка сборки всех объектов класса: {Pravka.__name__} в один файл- {e}')


logger.info(f'промежуточное сохранение сборки объектов класса: {Pravka.__name__} в excel')
try:
    df_kum_sborka_Pravka.to_excel(links_main(fr"{DIR}\main_links.txt", 'save_class_2_Pravka'), index=False)
except Exception as e:
    logger.error(f'ошибка промежуточного сохранение сборки объектов класса: {Pravka.__name__} в excel {e}')

class Raschet_KUM_dlya_dir:
    """сборка кума подвергается дополнительным вычислениям 
    цена простоя / фин услуги из SQL / новый расчет трейдин и прочее с префиксом dir"""
    
    # KOEF_FIN_USLUGI = 0.7 # применяется к сумме финуслуг

    def __init__(self, data):
        self.data = copy.deepcopy(data) # подается датафрейм сборка кума / результат concat объектов класса Pravka
        self.fn_starter()

    
    def fn_data_oplaty_scheta(self):
        logger.info(f"добавляем дату_оплаты_счета из resul_slivaniva в df столбец - dir_дата_оплаты_счета")
        try:
            self.data['dir_дата_оплаты_счета'] = self.data.apply(lambda x: return_result_oplata_scheta(resul_slivaniva, 'scl_vin', x.vin, 'дата_выдачи_факт', x.дата_выдачи, 'дата_оплаты_счета'), axis=1)
            
            try: # применить к столбцу to_datetime
                logger.info(f"применяем datetime к - dir_дата_оплаты_счета")
                self.data['dir_дата_оплаты_счета'] = pd.to_datetime(self.data['dir_дата_оплаты_счета'], format='mixed', errors='coerce')
            except Exception as ex_:
                    logger.error(f'❌ ошибка класса {Raschet_KUM_dlya_dir.__name__} функции {self.fn_data_oplaty_scheta.__name__} этап преобразования столбца в datetime - {ex_}')
        except Exception as e:
                logger.error(f'❌ ошибка класса {Raschet_KUM_dlya_dir.__name__} функции {self.fn_data_oplaty_scheta.__name__} - {e}')

    def fn_raznica_day(self):
         """разница дней между датой выдачи и датой оплаты счета"""
         logger.info(f"добавляем разницу дней из resul_slivaniva в df столбец - dir_разница_дней")
         try:
            self.data['dir_разница_дней'] = self.data.apply(lambda x: return_result_raznica_day(resul_slivaniva, 'scl_vin', x.vin, 'дата_выдачи_факт', x.дата_выдачи, 'разница_дней'), axis=1)
         except Exception as e:
                logger.error(f'❌ ошибка класса {Raschet_KUM_dlya_dir.__name__} функции {self.fn_raznica_day.__name__} - {e}')


    def fn_cena_prostoya(self):
        logger.info(f"добавляем цену простоя из resul_slivaniva в df столбец - dir_цена_простоя")
        try:
            self.data['dir_цена_простоя'] = self.data.apply(lambda x: return_result_prostoy(resul_slivaniva, 'scl_vin', x.vin, 'дата_выдачи_факт', x.дата_выдачи, 'цена_простоя'), axis=1)
        except Exception as e:
                logger.error(f'❌ ошибка класса {Raschet_KUM_dlya_dir.__name__} функции {self.fn_cena_prostoya.__name__} - {e}')




    def fn_fackt_sebest_in_all_OVP_spisok(self):
        """цепляем фактическая себестоимость из сборки ОВП мск+яр"""
        logger.info(f"цепляем фактический доход из сборки ОВП мск+яр в df столбец - dir_факт_себест_из_ОВП_список")
        try:
            self.data['dir_факт_себест_из_ОВП_список'] = self.data.apply(lambda x:fackt_sebest_in_all_OVP_list(df_OVP_all, x.vin, x.дата_контракта, x.дата_выдачи, 'цена_покупки', True), axis=1)
        except Exception as e:
                logger.error(f'❌ ошибка класса {Raschet_KUM_dlya_dir.__name__} функции {self.fn_fackt_sebest_in_all_OVP_spisok.__name__} - {e}')

    def fn_fackt_sebest_in_all_OVP(self):
        """цепляем фактическая себестоимость из сборки ОВП мск+яр"""
        logger.info(f"цепляем фактический доход из сборки ОВП мск+яр в df столбец - dir_факт_доход_из_ОВП")
        try:
            self.data['dir_факт_себест_из_ОВП'] = self.data.apply(lambda x:fackt_dohod_in_all_OVP(df_OVP_all, x.vin, x.дата_контракта, x.дата_выдачи, 'цена_покупки', True), axis=1)
        except Exception as e:
                logger.error(f'❌ ошибка класса {Raschet_KUM_dlya_dir.__name__} функции {self.fn_fackt_sebest_in_all_OVP.__name__} - {e}')


    def fn_dir_sebestoimost_treid_in_iz_dvoh(self): 
         """себестоимость трейдина из двух значений трейдин_принятый_ам или dir_факт_себест_из_ОВП"""
         logger.info(f"себестоимость трейдина из двух значений трейдин_принятый_ам или dir_факт_себест_из_ОВП в столбец dir_себест_ОВП")
         try:
            self.data['dir_себест_ОВП'] = self.data.apply(lambda x:sebest_iz_dvoh(x.трейдин_принятый_ам, x.dir_факт_себест_из_ОВП), axis=1)
         except Exception as e:
                logger.error(f'❌ ошибка класса {Raschet_KUM_dlya_dir.__name__} функции {self.fn_dir_sebestoimost_treid_in_iz_dvoh.__name__} - {e}')

    def fn_dir_treid_in(self):
        logger.info(f"добавляем расчет трейдин в df столбец - dir_трейдин_доход")
        try:
            self.data['dir_трейдин_доход'] = self.data.apply(lambda x:stavka_treid_in(df_otchisleniva_ovp_pravilo, x.дата_выдачи, x.dir_себест_ОВП, x.marka_2), axis=1)
        except Exception as e:
                logger.error(f'❌ ошибка класса {Raschet_KUM_dlya_dir.__name__} функции {self.fn_dir_treid_in.__name__} - {e}')


    def fn_dir_treid_in_razbivka_po_kajdomu_prinyztomu(self):
        """здесь считается доход когда было принято более 1-го в трейдин при выдаче одного нового - то есть за каждый авто считается отдельная сумма"""
        logger.info(f"добавляем расчет трейдин в df столбец - dir_трейдин_доход_разбивка")
        try:
            self.data['dir_трейдин_доход_разбивка'] = self.data.apply(lambda x:stavka_treid_in_list(df_otchisleniva_ovp_pravilo, x.дата_выдачи, x.dir_факт_себест_из_ОВП_список, x.marka_2), axis=1)
        except Exception as e:
                logger.error(f'❌ ошибка класса {Raschet_KUM_dlya_dir.__name__} функции {self.fn_dir_treid_in_razbivka_po_kajdomu_prinyztomu.__name__} - {e}')

    
    
    def fn_fackt_dohod_in_all_OVP(self):
        """цепляем фактический доход из сборки ОВП мск+яр"""
        logger.info(f"цепляем фактический доход из сборки ОВП мск+яр в df столбец - dir_факт_доход_из_ОВП")
        try:
            self.data['dir_факт_доход_из_ОВП'] = self.data.apply(lambda x:fackt_dohod_in_all_OVP(df_OVP_all, x.vin, x.дата_контракта, x.дата_выдачи, 'итого_доход', True), axis=1)
        except Exception as e:
                logger.error(f'❌ ошибка класса {Raschet_KUM_dlya_dir.__name__} функции {self.fn_fackt_dohod_in_all_OVP.__name__} - {e}')

    def fn_fackt_cena_prostova_in_all_OVP(self):
        """цепляем фактическую цену простоя из сборки ОВП мск+яр с учетом простоя"""
        logger.info(f"цепляем фактическую цену простоя из сборки ОВП мск+яр в df столбец - dir_цена_простоя_из_ОВП")
        try:
            self.data['dir_цена_простоя_из_ОВП'] = self.data.apply(lambda x:fackt_dohod_in_all_OVP(df_OVP_all, x.vin, x.дата_контракта, x.дата_выдачи, 'цена_простоя', True), axis=1)
        except Exception as e:
                logger.error(f'❌ ошибка класса {Raschet_KUM_dlya_dir.__name__} функции {self.fn_fackt_cena_prostova_in_all_OVP.__name__} - {e}') 


    def fn_fackt_dohod_in_all_OVP_s_prostoem(self):
        """цепляем фактический доход из сборки ОВП мск+яр с учетом простоя"""
        logger.info(f"цепляем фактический доход с учетом простоя из сборки ОВП мск+яр в df столбец - dir_факт_доход_из_ОВП_с_уч_простоя")
        try:
            self.data['dir_факт_доход_из_ОВП_с_уч_простоя'] = self.data.apply(lambda x:fackt_dohod_in_all_OVP(df_OVP_all, x.vin, x.дата_контракта, x.дата_выдачи, 'итого_доход_с_уч_простоя', True), axis=1)
        except Exception as e:
                logger.error(f'❌ ошибка класса {Raschet_KUM_dlya_dir.__name__} функции {self.fn_fackt_dohod_in_all_OVP_s_prostoem.__name__} - {e}')

    

#####################
    def fn_fackt_realizacii_auto_in_all_OVP(self):
        """цепляем факт реализации из сборки ОВП мск+яр"""
        logger.info(f"цепляем факт реализации из сборки ОВП мск+яр из сборки ОВП мск+яр в df столбец - dir_факт_реализации_авто_из_ОВП")
        try:
            self.data['dir_факт_реализации_авто_из_ОВП'] = self.data.apply(lambda x:fackt_realizacii_auto_in_all_OVP(df_OVP_all, x.vin, x.дата_контракта, x.дата_выдачи, 'дата_полной_оплаты', True), axis=1)
        except Exception as e:
                logger.error(f'❌ ошибка класса {Raschet_KUM_dlya_dir.__name__} функции {self.fn_fackt_realizacii_auto_in_all_OVP.__name__} - {e}')
################################


    def fn_planiruemiy_dohod_in_all_OVP(self):
         """планируемый доход из ОВП разница между план_цена_продажи - цена_покупки """
         logger.info(f"считаем планируемый доход (план_цена_продажи - цена_покупки) из сборки ОВП мск+яр в df столбец - dir_факт_доход_из_ОВП_с_уч_простоя") 
         try:
            self.data['dir_планируемый_доход_из_ОВП'] = self.data.apply(lambda x:planiruemiy_dphod_all_OVP_list(df_OVP_all, x.vin, x.дата_контракта, x.дата_выдачи, 'план_цена_продажи', 'цена_покупки', True, x.dir_факт_реализации_авто_из_ОВП), axis=1)
         except Exception as e:
                logger.error(f'❌ ошибка класса {Raschet_KUM_dlya_dir.__name__} функции {self.fn_planiruemiy_dohod_in_all_OVP.__name__} - {e}')

    def fn_dir_podtverjdenie_dohoda_OVP(self):
        """Проверяем реализован ли авто по факту и только тогда утвержаем доход иначе большие минуса так как авто еще не реализован
        если dir_факт_реализации_авто_из_ОВП == False то по dir_факт_доход_из_ОВП будет 0"""
        logger.info(f"Проверяем реализован ли авто по факту и только тогда утвержаем доход - dir_факт_доход_из_ОВП_подтвержденный")
        try:
            self.data['dir_факт_доход_из_ОВП_подтвержденный'] = self.data.apply(lambda x:vozvrat_dohoda_po_faktu_realizacii(x.dir_факт_доход_из_ОВП, x.dir_факт_реализации_авто_из_ОВП), axis=1)
        except Exception as e:
                logger.error(f'❌ ошибка класса {Raschet_KUM_dlya_dir.__name__} функции {self.fn_dir_podtverjdenie_dohoda_OVP.__name__} - {e}')


    # 'dir_предполагаемый_доход_из_ОВП'

    def fn_dir_predpolojitelniy_dohod_OVP(self):
         """предполагаемый доход из ОВП средний"""
         logger.info(f"Предполагаемый доход из ОВП - dir_предполагаемый_доход_из_ОВП")
         try:
            self.data['dir_предполагаемый_доход_из_ОВП'] = self.data.apply(lambda x: mean_dohod_OVP(df_OVP_all, x.region, 'трейдин', 30, 'итого_доход', x.dir_факт_реализации_авто_из_ОВП, x.dir_планируемый_доход_из_ОВП), axis=1)
         except Exception as e:
                logger.error(f'❌ ошибка класса {Raschet_KUM_dlya_dir.__name__} функции {self.fn_dir_predpolojitelniy_dohod_OVP.__name__} - {e}')


    def fn_dir_fin_uslygi(self):
         """тащим финуслуги из SQL по VIN и дате выдачи"""
         logger.info(f"добавляем финуслуги из SQL по VIN и дате выдачи в df столбец - dir_фин_услуги")
         try:
            self.data['dir_фин_услуги'] = self.data.apply(lambda x:return_fin_uslygi(df_dkis, x.vin, x.дата_выдачи, 'fullKV'), axis=1)
         except Exception as e:
                logger.error(f'❌ ошибка класса {Raschet_KUM_dlya_dir.__name__} функции {self.fn_dir_fin_uslygi.__name__} - {e}')

    def fn_dir_fin_uslygi_s_koeff(self):
         """применяем коэффициент к dir_фин_услуги"""
         logger.info(f"применяем коэффициент {KOEF_FIN_USLUGI} к dir_фин_услуги и получем столбец - dir_фин_услуги_с_коэф")
         try:
            self.data['dir_фин_услуги_с_коэф'] = self.data['dir_фин_услуги'] * KOEF_FIN_USLUGI
         except Exception as e:
                logger.error(f'❌ ошибка класса {Raschet_KUM_dlya_dir.__name__} функции {self.fn_dir_fin_uslygi_s_koeff.__name__} - {e}')

    def fn_dir_fin_uslygi_v_tch_progr_priv(self):
         """тащим финуслуги из SQL по VIN и дате выдачи"""
         logger.info(f"добавляем финуслуги программу привилегий из SQL по VIN и дате выдачи в df столбец - dir_фин_услуги")
         # важно! разница в сумме привилегий - в выгрузке SQL 90% от того что стоит в КУМе - "типа что у МУЗУРОВА в базе SQL отображается 90% стоимости программы прививлегий"("с" Роман)
         # ПРОВЕРИЛИ ТОЧНО - ЭТО В ТОЧМ ЧИСЛЕ то есть sumpriv входит в fullKV = (sumpriv + KVkasko)
         try:
            self.data['dir_фин_услуги_в_тч_прогр_прив'] = self.data.apply(lambda x:return_fin_uslygi(df_dkis, x.vin, x.дата_выдачи, 'sumpriv'), axis=1)
         except Exception as e:
                logger.error(f'❌ ошибка класса {Raschet_KUM_dlya_dir.__name__} функции {self.fn_dir_fin_uslygi_v_tch_progr_priv.__name__} - {e}')

    def fn_dir_fin_uslygi_v_tch_progr_priv_100(self): ################ !!!!!!!!!!!!!!!!!!
         """dir_фин_услуги_в_тч_прогр_прив восстанавливаем сумму деля / на 0.9"""
         logger.info(f"dir_фин_услуги_в_тч_прогр_прив восстанавливаем сумму деля / на {KOEF_VOSSTANOVLENIYA_PROGR_PRIVILEGIY} в df столбец (восстанавливаем 100%) - dir_фин_услуги_в_тч_прогр_прив_100")
         # важно! разница в сумме привилегий - в выгрузке SQL 90% от того что стоит в КУМе - "типа что у МУЗУРОВА в базе SQL отображается 90% стоимости программы прививлегий"("с" Роман)
         # ПРОВЕРИЛИ ТОЧНО - ЭТО В ТОЧМ ЧИСЛЕ то есть sumpriv входит в fullKV = (sumpriv + KVkasko)
         try:
            self.data['dir_фин_услуги_в_тч_прогр_прив_100'] = self.data['dir_фин_услуги_в_тч_прогр_прив']/KOEF_VOSSTANOVLENIYA_PROGR_PRIVILEGIY
         except Exception as e:
                logger.error(f'❌ ошибка класса {Raschet_KUM_dlya_dir.__name__} функции {self.fn_dir_fin_uslygi_v_tch_progr_priv_100.__name__} - {e}')

    
    def fn_costraciya_kuma(self):
         """обрезаем фрейм по дате выдачи"""
         start_costraciya_year = YEAR_COSTRACII
         end_costraciya_date = yesterday().strftime("%Y-%m-%d")
         logger.info(f"обрезаем фрейм по дате выдачи - с {start_costraciya_year} года по {end_costraciya_date} - оставляем")
         try:
            self.data = self.data[ (self.data['дата_выдачи'].dt.year>=start_costraciya_year) & (self.data['дата_выдачи']<=end_costraciya_date)]
         except Exception as e:
                logger.error(f'❌ ошибка класса {Raschet_KUM_dlya_dir.__name__} функции {self.fn_costraciya_kuma.__name__} - {e}')
        
    def fn_dir_dohod_itogo(self):
         """считаем доход"""
         info_lst_name_column_sum = ['авто_доход', 'до_доход', 'доход_next', 'dir_трейдин_доход', 'dir_фин_услуги_с_коэф']
         logger.info(f"считаем итого доход: суммируем {info_lst_name_column_sum} результат заносим в df столбец - dir_итого_доход")
         try:
            self.data['dir_итого_доход'] = self.data[info_lst_name_column_sum].sum(axis=1)
         except Exception as e:
                logger.error(f'❌ ошибка класса {Raschet_KUM_dlya_dir.__name__} функции {self.fn_dir_dohod_itogo.__name__} - {e}')

    def fn_dir_dohod_itogo_s_prostoy(self):
         """считаем доход c учетом простоя"""
         logger.info(f"считаем доход c учетом простоя результат заносим в df столбец - dir_итого_доход_с_уч_простоя")
         try:
            self.data['dir_итого_доход_с_уч_простоя'] = self.data['dir_итого_доход']-self.data['dir_цена_простоя']
         except Exception as e:
                logger.error(f'❌ ошибка класса {Raschet_KUM_dlya_dir.__name__} функции {self.fn_dir_dohod_itogo_s_prostoy.__name__} - {e}')


    def fn_dir_fin_uslugi_bez_progr_priv(self):
         """из dir_фин_услуги - вычетаем dir_фин_услуги_в_тч_прогр_прив"""
         logger.info(f"из dir_фин_услуги - вычетаем dir_фин_услуги_в_тч_прогр_прив результат заносим в df столбец - dir_итого_доход_с_уч_простоя")
         try:
            self.data['dir_фин_услуги_без_пр_прив'] = self.data['dir_фин_услуги'] - self.data['dir_фин_услуги_в_тч_прогр_прив']
         except Exception as e:
                logger.error(f'❌ ошибка класса {Raschet_KUM_dlya_dir.__name__} функции {self.fn_dir_fin_uslugi_bez_progr_priv.__name__} - {e}')

    def fn_dir_fin_uslygi_itogo(self):
         """dir_фин_услуги_без_пр_прив + dir_фин_услуги_в_тч_прогр_прив_100 """
         logger.info(f"dir_фин_услуги_без_пр_прив + dir_фин_услуги_в_тч_прогр_прив_100 в df столбец - dir_фин_услуги_итого")
         try:
            self.data['dir_фин_услуги_итого'] = self.data['dir_фин_услуги_без_пр_прив'] + self.data['dir_фин_услуги_в_тч_прогр_прив_100']
         except Exception as e:
                logger.error(f'❌ ошибка класса {Raschet_KUM_dlya_dir.__name__} функции {self.fn_dir_fin_uslygi_itogo.__name__} - {e}')

    def fn_dir_fin_uslugi_bez_progr_priv_s_koef(self):
         """из dir_фин_услуги - вычетаем dir_фин_услуги_в_тч_прогр_прив"""
         logger.info(f"dir_фин_услуги_без_пр_прив * коэф {KOEF_FIN_USLUGI} результат заносим в df столбец - dir_фин_услуги_без_пр_прив_с_коэф")
         try:
            self.data['dir_фин_услуги_без_пр_прив_с_коэф'] = self.data['dir_фин_услуги_без_пр_прив'] * KOEF_FIN_USLUGI
         except Exception as e:
                logger.error(f'❌ ошибка класса {Raschet_KUM_dlya_dir.__name__} функции {self.fn_dir_fin_uslugi_bez_progr_priv_s_koef.__name__} - {e}')


    def fn_dir_drop_columns(self):
         """удаляем лишние столбцы"""
         logger.info(f"удаляем лишние столбцы")
         try:
            self.data.drop(columns=['dir_факт_себест_из_ОВП_список'], inplace=True)
         except Exception as e:
                logger.error(f'❌ ошибка класса {Raschet_KUM_dlya_dir.__name__} функции {self.fn_dir_drop_columns.__name__} - {e}')

    def fn_starter(self):
        self.fn_costraciya_kuma()
        self.fn_data_oplaty_scheta()
        self.fn_raznica_day()
        self.fn_cena_prostoya()
        self.fn_fackt_sebest_in_all_OVP_spisok()
        self.fn_fackt_sebest_in_all_OVP()
        self.fn_dir_sebestoimost_treid_in_iz_dvoh()
        self.fn_dir_treid_in()
        self.fn_dir_treid_in_razbivka_po_kajdomu_prinyztomu()
        self.fn_fackt_dohod_in_all_OVP()
        self.fn_fackt_cena_prostova_in_all_OVP()
        self.fn_fackt_dohod_in_all_OVP_s_prostoem()
        self.fn_fackt_realizacii_auto_in_all_OVP()
        self.fn_planiruemiy_dohod_in_all_OVP()
        self.fn_dir_podtverjdenie_dohoda_OVP()
        self.fn_dir_predpolojitelniy_dohod_OVP()
        self.fn_dir_fin_uslygi()
        self.fn_dir_fin_uslygi_s_koeff()
        self.fn_dir_fin_uslygi_v_tch_progr_priv()
        self.fn_dir_fin_uslygi_v_tch_progr_priv_100()
        self.fn_dir_dohod_itogo()
        self.fn_dir_dohod_itogo_s_prostoy()
        self.fn_dir_fin_uslugi_bez_progr_priv()
        self.fn_dir_fin_uslugi_bez_progr_priv_s_koef()
        self.fn_dir_fin_uslygi_itogo()
        self.fn_dir_drop_columns()


logger.info(f"создание объекта класса {Raschet_KUM_dlya_dir.__name__}")
try:
    df_kum_Raschet_KUM_dlya_dir = Raschet_KUM_dlya_dir(df_kum_sborka_Pravka)
except Exception as e:
    logger.error(f'ошибка создания объекта класса {Raschet_KUM_dlya_dir.__name__} - {e}')    

class Preobrazovanie_marki:
    def __init__(self, df):
        self.df = copy.deepcopy(df)
        self.fn_starter()

    def izvlech_model_marku(self):
        logger.info(f"🔄 {self.izvlech_model_marku.__name__} - разделяем марки и модели")
        try:
            mask1 = (self.df['marka_2']=='PAR_IMP')&(self.df['marka']=='KIAIMP')
            mask2 = (self.df['marka_2']=='PAR_IMP')&(self.df['marka']=='HYUNDAIIMP') 
            mask3 = (self.df['marka_2']=='PAR_IMP')&(self.df['marka']=='MAZDAIMP')
            mask4 = (self.df['marka_2']=='PAR_IMP')&(self.df['marka']=='VOLKSWAGENIMP')
            mask5 = (self.df['marka_2']=='PAR_IMP')&(self.df['marka']=='PARIMP')
            mask6 = (self.df['marka_2']=='OVP')&(self.df['marka']=='OVP')
            mask7 = (self.df['marka_2']=='MAZDA_NEXT')&(self.df['marka']=='MAZDA')&(self.df['модель'].str.split().str.len()==1)
            mask8 = (self.df['marka_2']=='MAZDA_NEXT')&(self.df['marka']=='MAZDA')&(self.df['модель'].str.split().str.len()>1)
            mask9 = (self.df['marka_2']=='VOLKSWAGEN')&(self.df['marka']=='VOLKSWAGEN')


            yslovie = [mask1, mask2, mask3, mask4, mask5, mask6, mask7, mask8, mask9]
            # возврат по марке
            vozvrat_result_marka = ['KIA', 
                                    self.df['модель'].str.split().str[0], 
                                    self.df['модель'].str.split().str[0], 
                                    self.df['модель'].str.split().str[0], 
                                    self.df['модель'].str.split().str[0],
                                    self.df['марка_доп'].str.join(''),
                                    self.df['marka'].str.split().str[0],
                                    self.df['модель'].str.split().str[0],
                                    self.df['marka'].str.split().str[0]]
            # возврат по модели
            vozvrat_result_model = [self.df['модель'].str.split().str[0], 
                                    self.df['модель'].str.split().str[1:].str.join(' '), 
                                    self.df['модель'].str.split().str[1:].str.join(' '), 
                                    self.df['модель'].str.split().str[1:].str.join(' '), 
                                    self.df['модель'].str.split().str[1:].str.join(' '),
                                    self.df['модель'].str.split().str[0],
                                    self.df['модель'].str.split().str[0],
                                    self.df['модель'].str.split().str[1:].str.join(' '),
                                    self.df['модель'].str.lower().str.replace('volkswagen', '').str.strip().str.split().str[1:].str.join(' ')]

            # применяем все условия
            self.df['marka_test'] = np.select(yslovie, vozvrat_result_marka, default=self.df['marka'].str.split().str[0])
            self.df['model_test'] = np.select(yslovie, vozvrat_result_model, default=self.df['модель'].str.join(''))
            # все вверх рег
            self.df['marka_test'] = self.df['marka_test'].str.upper().str.strip()
            self.df['model_test'] = self.df['model_test'].str.upper().str.strip()
        except Exception as ex_:
            logger.error(f'❌ ошибка {self.izvlech_model_marku.__name__} {ex_}')

    def clear_simbols(self):
        logger.info(f"🔄 {self.clear_simbols.__name__} - очищаем марку от цифр")
        try:
            self.df['marka_test'] = self.df['marka_test'].str.replace(r'\d+', '', regex=True)
        except Exception as ex_:
            logger.error(f'❌ ошибка {self.clear_simbols.__name__} {ex_}')

    def zamena_col(self):
        logger.info(f"🔄 {self.zamena_col.__name__} - заменяем значения стобцов")
        try:
            self.df['marka'] = self.df['marka_test']
            self.df['модель'] = self.df['model_test']
        except Exception as ex_:
            logger.error(f'❌ ошибка {self.zamena_col.__name__} {ex_}')

    def clear_col(self):
        try:
            logger.info(f"🔄 {self.clear_col.__name__} - удаляем вспомогательные столбцы")
            # Список столбцов для удаления
            cols_to_drop = ['марка_доп', 'marka_test', 'model_test']
            # Оставляем только те, которые есть в DataFrame
            cols_to_drop_existing = [col for col in cols_to_drop if col in self.df.columns]
            logger.info(f"столбцы для удаления: {cols_to_drop} и найденные столбцы {cols_to_drop_existing}")
            
            if cols_to_drop_existing:
                self.df = self.df.drop(cols_to_drop_existing, axis=1, errors='ignore')
                logger.info(f"✅ Удалены столбцы: {cols_to_drop_existing}")
            else:
                logger.warning(f"⚠️ Ни один из столбцов {cols_to_drop} не найден в DataFrame")
                
        except Exception as ex_:
            logger.error(f'❌ ошибка {self.clear_col.__name__} {ex_}')


    def fn_starter(self):
        self.izvlech_model_marku()
        self.clear_simbols()
        self.zamena_col()
        self.clear_col()


logger.info(f'преобразование объекта класса {Raschet_KUM_dlya_dir.__name__} в классе {Preobrazovanie_marki.__name__}')
try:
    df_kum_Raschet_KUM_dlya_dir_ = Preobrazovanie_marki(df_kum_Raschet_KUM_dlya_dir.data)
except Exception as e:
    logger.error(f'ошибка преобразования объекта класса {Raschet_KUM_dlya_dir.__name__} в классе {Preobrazovanie_marki.__name__} - {e}')  

logger.info(f'промежуточное сохранение сборки объекта класса: {Raschet_KUM_dlya_dir.__name__} в excel')
try:
    df_kum_Raschet_KUM_dlya_dir_.df.to_excel(links_main(fr"{DIR}\main_links.txt", 'save_class_3_Raschet_KUM_dlya_dir'), index=False)
except Exception as e:
    logger.error(f'ошибка промежуточного сохранение сборки объекта класса: {Raschet_KUM_dlya_dir.__name__} в excel {e}')


logger.info(f'тестирование столбцов на запись в SQL') 
try:
    exception_column_SQL(df_kum_Raschet_KUM_dlya_dir_.df, SERVER_SQL, DATABASE_SQL, USERNAME_SQL, PASSWORD_SQL)
except Exception as ex_:
        logger.error(f'❌ ошибка тестирования столбцов  {ex_}')


name_save_file_sql = 'result_kum_dir'
logger.info(f'запись итогового кума - в SQL под именем {name_save_file_sql}') 
try:
    df_kum_Raschet_KUM_dlya_dir_.df.to_sql(name_save_file_sql, con=connect_bd_SQL(SERVER_SQL, DATABASE_SQL, USERNAME_SQL, PASSWORD_SQL), if_exists='replace', index=False)
    logger.info(f'✅ {name_save_file_sql} запиcан в SQL') 
except Exception as ex_:
        logger.error(f'❌ ошибка записи итогового кума - в SQL {ex_}')


logger.info(f"запускаем обновление кум")
try:
    update_file(links_main(f'{DIR}\main_links.txt','update_kum'))
except Exception as ex_:
        logger.error(f'не удалось обновить кум {ex_}')

logger.info(f"запускаем обновление кум_дир")
try:
    update_file(links_main(f'{DIR}\main_links.txt','update_kum_dir'))
except Exception as ex_:
        logger.error(f'не удалось обновить кум_дир {ex_}')


logger.info(f"рассылаем кум")
try:
    TOPIC_TEXT = fr"""КУМ на {yesterday().strftime('%d-%m-%Y')}  service_message"""
    BODY_TEXT = f"Здравствуйте \nВо вложении СВОД КУМ на {yesterday().strftime('%d-%m-%Y')}"
    send_mail(KUM_USER_EMAIL,
            KUM_USER_EMAIL_CC,
            KUM_USER_EMAIL_BCC,
            TOPIC_TEXT,
            BODY_TEXT,
            links_main(f'{DIR}\main_links.txt','update_kum'),
            os.path.basename(links_main(f'{DIR}\main_links.txt','update_kum')),
            SEND_FROM,
            SERVER,
            PORT,
            USER_NAME,
            PASSWORD)
    logger.info(f"рассылка кум прошла успешно для следующих пользователей {KUM_USER_EMAIL, KUM_USER_EMAIL_CC, KUM_USER_EMAIL_BCC}")
except Exception as ex_:
        logger.error(f'❌ не удалось разослать кум {ex_}')


logger.info(f"рассылаем кум_дир")
try:
    TOPIC_TEXT = fr"""КУМ_ИД на {yesterday().strftime('%d-%m-%Y')}  service_message"""
    BODY_TEXT = f"Здравствуйте \nВо вложении СВОД КУМ_ИД на {yesterday().strftime('%d-%m-%Y')}"
    send_mail(KUM_DIR_USER_EMAIL,
            KUM_DIR_USER_EMAIL_CC,
            KUM_DIR_USER_EMAIL_BCC,
            TOPIC_TEXT,
            BODY_TEXT,
            links_main(f'{DIR}\main_links.txt','update_kum_dir'),
            os.path.basename(links_main(f'{DIR}\main_links.txt','update_kum_dir')),
            SEND_FROM,
            SERVER,
            PORT,
            USER_NAME,
            PASSWORD)
    logger.info(f"рассылка кум прошла успешно для следующих пользователей {KUM_DIR_USER_EMAIL, KUM_DIR_USER_EMAIL_CC, KUM_DIR_USER_EMAIL_BCC}")
except Exception as ex_:
        logger.error(f'❌ не удалось разослать кум_дир {ex_}')


# logger.info(f"запускаем субпроцесс - сравнение выдач по темпу и ОМ")
# try:
#     update_task(links_main(f'{DIR}\main_links.txt','starter_sravnenie'))
# except Exception as e:
#     logger.error(f"❌ ошибка в субпроцессе - сравнение выдач по темпу и ОМ {e}")


logger.info(f"Выполнение основного скрипта и сбпроцессов завершено")
