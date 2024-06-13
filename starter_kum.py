# блок логирования
import logging
logging.basicConfig(level=logging.INFO, filename="//sim.local/data/Varsh/OFFICE/CAGROUP/run_python/task_scheduler/kum/py_log.log",filemode="w", format="%(asctime)s %(levelname)s %(message)s")
# https://habr.com/ru/companies/wunderfund/articles/683880/   - ссылка на статью логирования
# filemode="a" дозапись "w" - перезапись
logging.info("Запуск скрипта starter_kum.ipynb")

import pandas as pd
import os
import shutil
import io
import msoffcrypto
from datetime import datetime, date, timedelta

# обязательно использовать версию openpyxl==3.0.10 (в версиях моложе возникает ошибка "Value must be either numerical or a string containing a wildcard" - если на файлах эксель есть некоторые фильтры в столбцах)

pd.options.display.max_colwidth = 100 # увеличить максимальную ширину столбца
pd.set_option('display.max_columns', None) # макс кол-во отображ столбц

# блок импортов для обновления сводных
import pythoncom
pythoncom.CoInitializeEx(0)
import win32com.client
import time

# блок импорта отправки почты
import smtplib,ssl
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import formatdate
from email import encoders

start_time_work_skript = time.time() ## точка отсчета времени
logging.info(f"запуск точки отсчета времени: {start_time_work_skript}")

# ссылка куда сохраняем
link_save = r"\\sim.local\data\Varsh\OFFICE\CAGROUP\run_python\task_scheduler\kum"
print(f"Путь сохранения файлов: {link_save}")
logging.info(f"Путь сохранения файлов: {link_save}")

def time_wopking_skript(start):
    """функция определения время выполнения скрипта

    Args:
        start (_type_): подаем переменную считавшую время на старте скрпта: start_time_work_skript = time.time()

    Returns:
        _type_: вернет разницу в минутах и секундах
    """
    logging.info(f"{time_wopking_skript.__name__} - ЗАПУСК")
    try:
        start = start
        end = time.time() - start
        min_ = int(end//60)
        sec_ = int((end - (min_*60)))
        return f"{min_} мин {sec_} сек"
    except:
        logging.error(f"{time_wopking_skript.__name__} - ОШИБКА", exc_info=True)
        

def open_file_links(link_ = r"\\sim.local\data\Varsh\OFFICE\CAGROUP\run_python\task_scheduler\kum\links_name_file.xlsx"):
    """Открывает файл со сссылками именами файлов паролями и именами листов  
    добавляет в этот файл время последнего сохранеия истояника данных - считывает по ссылке  
    убирает лишние столбцы   
    пересохраняет файл - чтоб сервер не удалил по сроку давности  

    Args:
        link_ (regexp, optional): _description_. Defaults to r"\sim.local\data\Varsh\OFFICE\CAGROUP\run_python\task_scheduler\kum\links_name_file.xlsx".

    Returns:
        _type_: _description_
    """
    logging.info(f"{open_file_links.__name__} - ЗАПУСК")
    try:
        # важно !!!!- если в столбце work_sheet - значение lot (значит рабочих листов много)
        print('Открытие файла со ссылками и именами файлов')
        links_name = pd.read_excel(link_, sheet_name="Sheet1", dtype='str') # открываем файл
        links_name['time_update'] = links_name['link'].apply(lambda x: datetime.fromtimestamp(os.path.getmtime(x)))     # добавляем время последнего обновления источника по ссылке
        links_name = links_name[['link', 'name', 'pass', 'kum_work_sheet', 'kum_not_work_sheet', 'time_update']]                                                # оставляем только нужные столбцы
        links_name['pass'] = links_name['pass'].fillna(0).astype(int)                                                   # преобразовываем столбец с паролями
        # пересохраняем чтоб обновилось время последненго сохранения и файл не исчез с сервера по сроку давности
        links_name.to_excel(r"\\sim.local\data\Varsh\OFFICE\CAGROUP\run_python\task_scheduler\kum\links_name_file.xlsx")
        return links_name
    except:
        logging.error(f"{open_file_links.__name__} - ОШИБКА", exc_info=True)
        

def testing_links(links):
    """Проверка ссылкок на файлы  
    
    Результатом является вывод текстового сообщения с результатом проверки ссылки 

    Args:
        links (_type_): list _description_ - подается список ссылок
    """
    logging.info(f"{testing_links.__name__} - ЗАПУСК")
    
    if os.path.exists(f"{links}"):
        print(f'OK - ', links)
        logging.info(f"{testing_links.__name__} ссылка рабочая {links}")
    else:
        print(f"ОШИБКА - ", links)
        logging.error(f"{testing_links.__name__} ссылка не рабочая {links}", exc_info=True)
        
def search_region(x: str):
    """ищет регион в названии файла

    Args:
        x (str): подается строка

    Returns:
        _type_: возвращает регион
    """
    logging.info(f"{search_region.__name__} - ЗАПУСК")
    
    try:
        if "YAR" in x:
            return "YAR"
        elif "MSK" in x:
            return "MSK"
        elif "SAR" in x:
            return "SAR"
        else:
            return "неизвестно"
    except:
        logging.error(f"{search_region.__name__} - ОШИБКА", exc_info=True)
        
def marka_replace(list_data: list, repl: list =['vved', 'varsh', 'arh','MSK', 'YAR', 'SAR', 'KUM', '2021', '2022', '2023', '2024', '2025', '2026']):
    """Функция сбора марок авто из названия файлов  
    который очищаем от рудиментов названия

    Args:
        list_data (list): _description_ = подаем список с именами файлов (который будем очищать)
        repl (list , optional): _description_. Defaults to ['vved', 'varsh', 'arh','MSK', 'YAR', 'SAR', 'KUM', '2021', '2022', '2023', '2024', '2025']. - список исключений

    Returns:
        _type_: _description_
    """
    logging.info(f"{marka_replace.__name__} - ЗАПУСК")
    
    try:
        new_lst = []
        # замена прочерков
        for i in list_data:
            new_lst.append(i.replace('_',' '))
        # убираем все рудиментные значения по списку repl
        for i in range(len(new_lst)):
            for j in repl:
                new_lst[i]=new_lst[i].replace(j,'')
        # пробегаем по списку и если есть слипшиеся значения типа 'HYUNDAI BAIC UKA' разделяем их и удаляем текущее значение
        for i in range(len(new_lst)):
            if len(new_lst[i].split())>1:
                for j in new_lst[i].split():
                    new_lst.append(j)
                new_lst.remove(new_lst[i])
        # убираем пробелы
        for i in range(len(new_lst)):
            new_lst[i]=new_lst[i].strip()
        # оставляем уникальные значения марок
        new_lst = set(new_lst)
        return new_lst
    
    except:
        logging.error(f"{marka_replace.__name__} - ОШИБКА", exc_info=True)
        
def append_dict_marka_auto(dict_update: dict, lst_mark_uniq: [list, set]) -> dict:
    """дополняет справчник маркой авто   
    по названию файла определяет какие марки авто в нем есть (не имеет отношения к ОВП)

    Args:
        dict_update (dict): подаем справочник  который будем дополнять
        lst_mark_uniq (list, set]): подаем итерируемый список с марками авто 

    Returns:
        dict: возвращает обновленный словарь 
    """
    logging.info(f"{append_dict_marka_auto.__name__} - ЗАПУСК")
    
    try:
        for i in dict_update.items():
            dict_update[i[0]]['marka'] = ', '.join([unic for unic in lst_mark_uniq if unic in i[0]] )
        return dict_update
    except:
        logging.error(f"{append_dict_marka_auto.__name__} - ОШИБКА", exc_info=True)
        
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
            df = pd.read_excel(decrypted_workbook)
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
            df = pd.read_excel(lnk, dtype='str')
            sheet_names = list(pd.read_excel(lnk, sheet_name=None).keys())
            return df, sheet_names
        else:
            df = pd.read_excel(lnk, sheet_name=lst_name, dtype='str')
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
        
        
def predobrabotka_df(df):
    """первичная обработка df   
    находит шапку таблицы по значению VIN в строках и делает эту строку заголовком таблицы

    Args:
        df (_type_): подаем df

    Returns:
        _type_: возращает df
    """
    logging.info(f"{predobrabotka_df.__name__} - ЗАПУСК")
    try:
        
        count_col = 0
        for i in df.columns:
            if str(i).lower() == 'vin':
                    count_col +=1
            counter_vin = df[i].apply(lambda x: str(x).lower()).str.contains('^vin').sum() # ^ - в регулярке используется для поиска когда слово начинается с 
            name_column = i
            row_number = None

            if counter_vin >0:
                        row_number = df[df[name_column].apply(lambda x: str(x).lower())=='vin'].index[0]
                        #print(f"VIN найден в столбце {i}")
                        break
                    
        if count_col != 0:
                return df # если шапка в первой строке, ничего не изменяем
        else: 
            new_header = df.iloc[row_number] # берем первую строку как заголовок
            df = df[row_number+1:]  # отбрасываем исходный заголовок
            df.rename(columns=new_header, inplace=True) # переименовываем столбцы
            return df
    except:
        logging.error(f"{predobrabotka_df.__name__} - ОШИБКА", exc_info=True)
        
        
def head_registr_low_strip(df):
    """переводит названия столбцов в нижний регистр   
    удаляет пробелы слева и справа

    Args:
        df (_type_): _description_

    Returns:
        _type_: _description_
    """
    logging.info(f"{head_registr_low_strip.__name__} - ЗАПУСК")
    try:
        return df.rename(columns={f'{i}' : f'{str(i).lower().strip()}' for i in df.columns})
    except:
        logging.error(f"{head_registr_low_strip.__name__} - ОШИБКА", exc_info=True)
        
        
def rename_columns_individual(df, spravka_ind: dict):
    """переименовывает наименования столбцов персонально
    подаем df и словарь

    Args:
        df (_type_): _description_
        spravka_ind (dict): подаем словарь с данными формата - {'форма оплаты': ['форма оплаты', 'б/н / нал', 'кре/нал', 'кредит / нал']}
        если в списке словаря находит совпадение - возвращает ключ словаря 'форма оплаты'

    Returns:
        _type_: возвращает обработанный df
    """
    logging.info(f"{rename_columns_individual.__name__} - ЗАПУСК")
    
    try:
        for i in df.columns:
            for j in spravka_ind.keys():
                if i in spravka_ind[j]:
                    df = df.rename(columns={i:j})
        return df
    
    except:
        logging.error(f"{rename_columns_individual.__name__} - ОШИБКА", exc_info=True)
        

def df_white_list_col(df, white_list_columns: list):
    """оставляет в df Только те столбцы которые есть в блоем списке  
    если в df нет столбцов из белого списка они будут добавлены (нужно для беспроблемной конкатенации - чтоб все было одинаково)  

    Args:
        df (_type_): _description_
        white_list_columns (list): лист именами столбцов которые хотим видеть в df

    Returns:
        _type_: df
    """
    logging.info(f"{df_white_list_col.__name__} - ЗАПУСК")
    
    try:
        df = df[[i for i in df.columns if i in white_list_columns]]
        
        # спсок колонок которые есть в белом списке и если их нет в df - они будут добавлены
        add_columns = [i for i in white_list_columns if i not in df.columns]
        for i in add_columns:
            df[i] = None
        return df
    
    except:
        logging.error(f"{df_white_list_col.__name__} - ОШИБКА", exc_info=True)
        

def df_white_list_col_OVP(df, white_list_columns_individual: list, white_list_columns: list):
    """оставляет в df Только те столбцы которые есть в блоем списке   
    если в df нет столбцов из белого списка они будут добавлены (нужно для беспроблемной конкатенации - чтоб все было одинаково)  
    
    Дополнительно подгоняет ОВП под общий стандарт названия столбцов  

    Args:
        df (_type_): _description_
        white_list_columns (list): лист именами столбцов которые хотим видеть в df

    Returns:
        _type_: df
    """
    logging.info(f"{df_white_list_col_OVP.__name__} - ЗАПУСК")
    try:
        df = df[[i for i in df.columns if i in white_list_columns_individual]]
        # подгоняем названия столбцов как у всех
        df = df.rename(columns={'доход_авто_кум':'итого ам доход', 'доход_до_кум':'до доход','доход_фу_кум':'доход финуслуги',
                                'итого_кум':'кум доход итого', 'примечание':'форма оплаты', 'менеджер (продал)':'продавец'})
        
        #спсок колонок которые есть в белом списке и если их нет в df - они будут добавлены
        add_columns = [i for i in white_list_columns if i not in df.columns]
        for i in add_columns:
            df[i] = None
        return df
    
    except:
        logging.error(f"{df_white_list_col_OVP.__name__} - ОШИБКА", exc_info=True)
        

def df_white_list_col_H_B_U_v_MSK(df, white_list_columns_individual: list, white_list_columns: list):
    """оставляет в df Только те столбцы которые есть в блоем списке   
    если в df нет столбцов из белого списка они будут добавлены (нужно для беспроблемной конкатенации - чтоб все было одинаково)  
    
    Дополнительно подгоняет KUM_HYUNDAI_BAIC_UKA_varsh_MSK под общий стандарт названия столбцов  

    Args:
        df (_type_): _description_
        white_list_columns (list): лист именами столбцов которые хотим видеть в df

    Returns:
        _type_: df
    """
    logging.info(f"{df_white_list_col_H_B_U_v_MSK.__name__} - ЗАПУСК")
    
    try:
        df = df[[i for i in df.columns if i in white_list_columns_individual]]
        # подгоняем названия столбцов как у всех
        df = df.rename(columns={'дата выдачи клиенту':'дата выдачи', 'доход_ам_бонус':'итого ам доход', 'доход_до':'до доход', 'перечисления_от_окис':'доход финуслуги', 
                                '№25р дох трейд-ин':'доход трейдин', 'итого_кум':'кум доход итого', 'нал/кредит':'форма оплаты', 'марка_':'салон', 'покупатель':'клиент'})
        
        #спсок колонок которые есть в белом списке и если их нет в df - они будут добавлены
        add_columns = [i for i in white_list_columns if i not in df.columns]
        for i in add_columns:
            df[i] = None
        return df
    except:
        logging.error(f"{df_white_list_col_H_B_U_v_MSK.__name__} - ОШИБКА", exc_info=True)
        

def name_df_columns_and_marka(df, name_df:str, marka_auto:str, lst_name_istochnik:str, region_ist:str):
    """функция добавляет в df столбцы с именем источника и маркой - которые подаются

    Args:
        df (_type_): datafarame
        name_df (str): имя df (чтоб понимать источник в случае расхождения)
        marka_auto (str): марка авто (они обрабатываются ранее по имени файла)

    Returns:
        _type_: df
    """
    logging.info(f"{name_df_columns_and_marka.__name__} - ЗАПУСК")
    try:
        df['имя_бд'] = name_df
        df['марки_бд'] = marka_auto
        df['регион'] = region_ist
        df['с_листа'] = lst_name_istochnik
        return df
    except:
        logging.error(f"{name_df_columns_and_marka.__name__} - ОШИБКА", exc_info=True)
        
        
def conversorrrrrr_date(df, name_date_columns:str):
    """функция для преобразования кривых формат дат ы том числле формата 41253   
      
    Подается df и имя столбца

    Args:
        df (dataframe): df
        name_date_columns (str): имя столбца с датой (который хотим преобразовать)  

    Returns:
        _type_: возварщает преобразованный df  
    """
    logging.info(f"{conversorrrrrr_date.__name__} - ЗАПУСК")
    try:
        from datetime import datetime
        
        formating = (lambda x: datetime.fromordinal(datetime(1900, 1, 1).toordinal() + int(x) - 2))
        df[name_date_columns] = df[name_date_columns].apply(lambda x: str(x).replace('00:00:00','').strip() if '00:00:00' in str(x) else x)
        df[name_date_columns] = df[name_date_columns].apply(lambda x: formating(x) if len(str(x))==5 and str(x)[0] == '4' else x)
        df[name_date_columns] = pd.to_datetime(df[name_date_columns], format='mixed')
        return df
    except:
        logging.error(f"{conversorrrrrr_date.__name__} - ОШИБКА", exc_info=True)
        

def conversion_columns_integer(df, columns:list):
    """преобразует чиловые значения в один формат int   

    NA = 0

    Args:
        df (_type_): dataframe  
        columns (list): список столбцов которые преобразовываем  

    Returns:
        _type_: _description_
    """
    logging.info(f"{conversion_columns_integer.__name__} - ЗАПУСК")
    try:
        for i in columns:
            df[i] = df[i].astype('float')
            df[i] = df[i].astype('int', errors='ignore')
            df[i] = df[i].fillna(0)
            df[i] = df[i].astype('int')
        return df
    except:
        logging.error(f"{conversion_columns_integer.__name__} - ОШИБКА", exc_info=True)
        
        
def form_pay(x):
    """определяет кредит / нал 

    Args:
        x (_type_): _description_

    Returns:
        _type_: _description_
    """
    
    try:
        x = str(x).lower()
        kredit_lst = ['кредит','банк','лизинг','кре']
        counter = 0
        for i in kredit_lst:
            if i in x and 'не для' not in x:
                counter+=1
        if counter != 0:
            return 'кредит'
        
        elif 'б/н' in x or 'безнал' in x:
            return 'нал'
        
        else:
            return 'нал'
    except:
        logging.error(f"{form_pay.__name__} - ОШИБКА", exc_info=True)
        

def reg_test(rg, podr):
    """функция находит YAR и проверяет есть ли там RYB

    Args:
        rg (_type_): столбец регион
        podr (_type_): столбец подразделение

    Returns:
        _type_: _description_
    """
    
    try:
        if rg == 'YAR':
            if 'яр' in podr.lower():
                return 'YAR'
            elif 'рыб' in podr.lower():
                return 'RYB'
            else:
                return rg
        else:
            return rg
    except:
        logging.error(f"{reg_test.__name__} - ОШИБКА", exc_info=True)
        

def raspred_salon_marki(marki, salon, model, region):
    """функция для определения автоцентра по 4 входящим параметрам  
    
    все дело в KUM_OMODA_JAECOO_SAR  и KUM_HYUNDAI_BAIC_UKA_varsh_MSK

    Args:
        marki (_type_): столбец с марками
        salon (_type_): салон 
        model (_type_): можеди авто 
        region (_type_): регион 
        

    Returns:
        _type_: _description_
        
    ОБРАЗЕЦ ПРИМЕНЕНИЯ ФУНКЦИИ : result_svod_pred['автоцентр'] = result_svod_pred.apply(lambda x: raspred_salon_marki(x.марки_бд, x.салон, x.модель, x.регион), axis=1)
    """
    
    try:
        marki = str(marki).upper().split(',')
        
        # для всех остальынх
        if len(marki)==1:
            return marki[0]
        
        # для KUM_OMODA_JAECOO_SAR
        elif len(marki)>1 and region=="SAR":
            if 'OMODA' in str(model).upper():
                return 'OMODA'
            elif 'JAECOO' in str(model).upper():
                return 'JAECOO'
        
        # для KUM_HYUNDAI_BAIC_UKA_varsh_MSK
        elif len(marki)>1 and region=="MSK" and ('HYUNDAI' in str(salon).upper() or 'BAIC' in str(salon).upper() or 'UKA' in str(salon).upper()):
            if 'HYUNDAI' in str(salon).upper():
                return 'HYUNDAI'
            elif 'BAIC' in str(salon).upper():
                return 'BAIC'
            elif 'UKA' in str(salon).upper():
                return 'UKA'
        else:
            return 'неизвестно'
    except:
        logging.error(f"{raspred_salon_marki.__name__} - ОШИБКА", exc_info=True)
        
        
def korp_rozn(x):
    """определяет юрл и физ продажу по клиенту

    Args:
        x (_type_): _description_

    Returns:
        _type_: _description_
    """
    
    try:
        x = str(x)
        counter = 0
        lstr_test = ['АО','ООО','ВТБ','НАО','ЗАО','ИП','КФХ', 'АНО']
        for i in lstr_test:
            if i in x:
                counter+=1
        if counter !=0:
            return 'юрл'
        else:
            return 'физ'
    except:
        logging.error(f"{korp_rozn.__name__} - ОШИБКА", exc_info=True)
        
# считываем актуальный пароль 
def my_pass():
    """функция считывания пароля

    Returns:
        _type_: _description_
    """
    logging.info(f"{my_pass.__name__} - ЗАПУСК")
    
    try:
        with open(f'//sim.local/data/Varsh/OFFICE/CAGROUP/run_python/task_scheduler/temp_/password_email.txt', 'r') as actual_pass:
            return actual_pass.read()
        
    except:
        logging.error(f"{my_pass.__name__} - ОШИБКА", exc_info=True)
        


# письмо если нет ошибок
def send_mail(send_to:list):
    """рассылка почты

    Args:
        send_to (list): _description_
    """
    logging.info(f"{send_mail.__name__} - ЗАПУСК")
    
    try:
        send_from = 'skrutko@sim-auto.ru'                                                                
        subject = f"КУМ на {(datetime.now()-timedelta(1)).strftime('%d-%m-%Y')}"                                                                 
        text = f"Здравствуйте\nВо вложении КУМ на {(datetime.now()- timedelta(1)).strftime('%d-%m-%Y')}"                                                                      
        files = "//sim.local/data/Varsh/OFFICE/CAGROUP/run_python/task_scheduler/kum/КУМ_ОБЩИЙ.xlsx"  
        server = "server-vm36.SIM.LOCAL"
        port = 587
        username='skrutko'
        password=my_pass()
        isTls=True
        
        msg = MIMEMultipart()
        msg['From'] = send_from
        msg['To'] = ','.join(send_to)
        msg['Date'] = formatdate(localtime = True)
        msg['Subject'] = subject
        msg.attach(MIMEText(text))

        part = MIMEBase('application', "octet-stream")
        part.set_payload(open(files, "rb").read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', 'attachment; filename="KUM.xlsx"') # имя файла должно быть на латинице иначе придет в кодировке bin
        msg.attach(part)

        smtp = smtplib.SMTP(server, port)
        if isTls:
            smtp.starttls()
        smtp.login(username, password)
        smtp.sendmail(send_from, send_to, msg.as_string())
        smtp.quit()
        logging.info(f"{send_mail.__name__} - ВЫПОЛНЕНО")
        logging.info(f"Адреса рассылки {send_to}")
    except:
        logging.error(f"{send_mail.__name__} - ОШИБКА", exc_info=True)
    

# письмо если есть ошибки
def send_mail_danger(send_to:list):
    """расслыка почты если ошибка

    Args:
        send_to (_type_): _description_
    """
    logging.info(f"{send_mail_danger.__name__} - ЗАПУСК")
    
    try:                                                                                       
        send_from = 'skrutko@sim-auto.ru'                                                                
        subject =  f"проверьте исходники {'//sim.local/data/Varsh/OFFICE/CAGROUP/run_python/task_scheduler/kum'}"                                                                  
        text = f"проверьте исходники {'//sim.local/data/Varsh/OFFICE/CAGROUP/run_python/task_scheduler/kum'}"                                                                      
        files = '//sim.local/data/Varsh/OFFICE/CAGROUP/run_python/task_scheduler/kum/py_log.log'  
        server = "server-vm36.SIM.LOCAL"
        port = 587
        username='skrutko'
        password=my_pass()
        isTls=True
        
        msg = MIMEMultipart()
        msg['From'] = send_from
        msg['To'] = ','.join(send_to)
        msg['Date'] = formatdate(localtime = True)
        msg['Subject'] = subject
        msg.attach(MIMEText(text))

        part = MIMEBase('application', "octet-stream")
        part.set_payload(open(files, "rb").read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', 'attachment; filename="log.txt"') # имя файла должно быть на латинице иначе придет в кодировке bin
        msg.attach(part)

        smtp = smtplib.SMTP(server, port)
        if isTls:
            smtp.starttls()
        smtp.login(username, password)
        smtp.sendmail(send_from, send_to, msg.as_string())
        smtp.quit()
        logging.info(f"{send_mail_danger.__name__} - ВЫПОЛНЕНО")
        logging.info(f"Адреса рассылки {send_to}")
    except:
        logging.error(f"{send_mail_danger.__name__} - ОШИБКА", exc_info=True)
        

def detected_danger(filename_log = "//sim.local/data/Varsh/OFFICE/CAGROUP/run_python/task_scheduler/kum/py_log.log"):
    """обнаружение ошибок в логах   
    ищет 'warning'

    Returns:
        _type_: bool
    """
    logging.info(f"{detected_danger.__name__} - ЗАПУСК")
    
    try:
        with open(filename_log, '+r') as file:
            return 'error' in file.read().lower()
    except:
        logging.error(f"{detected_danger.__name__} - ОШИБКА", exc_info=True)
        
        
        
def read_email_adress(mail = fr'\\sim.local\data\Varsh\OFFICE\CAGROUP\run_python\task_scheduler\kum\Список_адресатов.xlsx'):
    """Функция считывания адресатов для рассылки

    Args:
        mail (_type_, optional): _description_. Defaults to fr'\sim.local\data\Varsh\OFFICE\CAGROUP\run_python\task_scheduler\temp_\Список_адресатов.xlsx'.

    Returns:
        _type_: возфращает строку со списком email
    """
    logging.info(f"{read_email_adress.__name__} - ЗАПУСК")
    
    try:
        em_list = pd.read_excel(mail)
        return list(em_list['email'])
    except:
        logging.error(f"{read_email_adress.__name__} - ОШИБКА", exc_info=True)
        
        
def sending_mail(lst_email, lst_email_error):
    """рассылка почты - если нет ошибок вызываем send_mail(),   
    если есть ошибки send_mail_error()   
    """
    logging.info(f"{sending_mail.__name__} - ЗАПУСК")
    
    try:
        if detected_danger()==False:
            send_mail(lst_email)
        else:
            send_mail_danger(lst_email_error)
            
        logging.info(f"{sending_mail.__name__} - ВЫПОЛНЕНО")
    except:
        logging.error(f"{sending_mail.__name__} - ОШИБКА", exc_info=True)
        
        
# открываем файл со ссылками (на источники КУМЫ)
# считываем данные пересохраняем файл для невозможности удаления сервером
links_name = open_file_links()
logging.info(f"открываем файл со ссылками (на источники КУМЫ) / считываем данные пересохраняем файл для невозможности удаления сервером")

print('Тестирвоание ссылок')
logging.info(f"Тестирвоание ссылок")
for i in range(links_name.shape[0]):
    testing_links(links_name['link'].iloc[i])
    
    

# копируем файлы и сохраняем в новой директории под новыми именами
print('Обновление баз данных')
logging.info(f"Обновление баз данных")

dict_link_name = {}
list_marki = []
for i in range(links_name.shape[0]):
    link_copy = links_name['link'].iloc[i]                                                      # ссылка откуда копируем файл
    name_new_file = links_name['name'].iloc[i]+'.'+links_name['link'].iloc[i].split('.')[-1]    # имя нового файла
    new_link = fr'{link_save}\{name_new_file}'                                                  # новый путь где лежит файл
    list_marki.append(name_new_file.split('.')[0])                                                            # собиарем лист сов семи именами для дальнейше обработки (там останутся только марки)
    
    # блок логгирования
    print(f'источник данных - {link_copy}, новое имя - {name_new_file}')
    logging.info(f'источник данных - {link_copy}, новое имя - {name_new_file}')
    print(f'сохранено в {new_link}')
    logging.info(f"сохранено в {new_link}")
    
    
    shutil.copy2(link_copy, f'{new_link}')                                                      # копируем файл из исходной директории в диреткорию для дальнейшей обработки
                                                                                                # собираем справочник с данными ссылки, имена, пароли, регионы.
    dict_link_name[links_name['name'].iloc[i]] = {'link' : new_link.replace("\\\\","//").replace("\\","/"),
                                                  'pass': links_name['pass'].iloc[i],
                                                  'region' : search_region(name_new_file),
                                                  'kum_work_sheet' : str(links_name['kum_work_sheet'].iloc[i]).split(', '),
                                                  'kum_not_work_sheet' : str(links_name['kum_not_work_sheet'].iloc[i]).split(', '),}
    
    
# уникальные марки авто
logging.info(f"Собираем уникальные марки")
unique_marki = marka_replace(list_marki)

# дополняем / обновляем словарь разделом марка
logging.info(f"дополняем / обновляем словарь разделом марка")
dict_link_name = append_dict_marka_auto(dict_link_name, unique_marki)



class Manufactory_df:
    
    # список столбцов для переименования в единый стиль названий
    dict_rename_columns = {'форма оплаты': ['форма оплаты', 'б/н / нал', 'кре/нал', 'кредит / нал'],
                        'салон': ['салон', 'марка'],
                        'доход пргм привилегий': ['доход_пргр_прив_кум', 'дох_прог_прев'],
                        'продавец':['продавец', 'Продавец', 'менеджер', 'менеджер (продал)'],
                        'доход трейдин':['доход трейдин', 'доход trade in']}
    
    # названия столбцов которые оставляем кроме ОВП и ХЕНДЭ_БАИК_МСК
    white_list_columns = ['дата выдачи', 'модель', 'vin', 'клиент', 'итого ам доход', 'до доход', 'доход финуслуги', 'доход трейдин', 'доход next', 'кум доход итого','форма оплаты', 'доход пргм привилегий', 'салон', 'продавец']
    
    # названия столбцов которые оставляем ОВП
    white_list_columns_ovp = ['дата выдачи', 'модель', 'vin', 'клиент', 'доход_авто_кум', 'доход_до_кум', 'доход_фу_кум', 'итого_кум', 'примечание', 'доход пргм привилегий', 'менеджер (продал)']
    
    # названия столбцов которые оставляем HYUNDAI_BAIC_UKA_varsh_MSK
    white_list_columns_ovp_HYUNDAI_BAIC_UKA_varsh_MSK = ['дата выдачи клиенту', 'модель', 'vin', 'покупатель', 'доход_ам_бонус', 'доход_до', 'перечисления_от_окис', '№25р дох трейд-ин', 'итого_кум', 'продавец', 'нал/кредит', 'марка_', 'доход пргм привилегий']
    
    # список тех, кто будет обрабатываться не так как все / их кумулятив значительно отличается от большинства
    list_individual_treatment = ['KUM_HYUNDAI_BAIC_UKA_varsh_MSK', 'KUM_OVP_vved_MSK']
    
    def __init__(self, name, link, password, region, kum_work_sheet, kum_not_work_sheet, marka, flag = True):
        """_summary_

        Args:
            name (_type_): _description_
            link (_type_): _description_
            password (_type_): _description_
            region (_type_): _description_
            marka (_type_): _description_
        """
        self.name = name
        self.link = link
        self.password = password
        self.region = region
        self.kum_work_sheet = kum_work_sheet
        self.kum_not_work_sheet = kum_not_work_sheet
        
        self.marka = marka
        self.df = None 
        self.sheet_names = None
        self.flag = flag
        logging.info(f"создание объекта класса {__class__.__name__} имя {self.name}")
        self.start()
        
        
    def manufactory(self):
        logging.info(f"{self.manufactory.__name__} - ЗАПУСК")
        # если 1 обрабатываемый лист и нет признака lot (условное обозначение когда листов много и они с разными названиями как у kia, тогда исключаем Справоник, прочее и обрабатываем все листы)
        if len(self.kum_work_sheet) == 1 and 'lot' not in self.kum_work_sheet and self.name not in Manufactory_df.list_individual_treatment:
            logging.info(f"обработка листа {self.kum_work_sheet[0]}")
            self.df, self.sheet_names = open_dataframe(self.link, self.password, self.kum_work_sheet[0])    # открываем df
            self.df = predobrabotka_df(self.df)                                                             # находим шапку
            self.df = head_registr_low_strip(self.df)                                                       # шапку в нижний регистр
            self.df = rename_columns_individual(self.df, Manufactory_df.dict_rename_columns)                # переименовываем столбцы в единый стиль
            self.df = df_white_list_col(self.df, Manufactory_df.white_list_columns)                         # оставляем нужные столбцы и добавляем если нет столбцов из white_list_columns
            self.df = name_df_columns_and_marka(self.df, self.name, self.marka, self.kum_work_sheet[0], self.region)     # добавляем столбцы с имененм источника и маркой которая есть в названии источника и именем листа

        
        elif len(self.kum_work_sheet) > 1 and 'lot' not in self.kum_work_sheet and self.name not in Manufactory_df.list_individual_treatment:
            bank_dataframes = []
            
            for lst_name_open in self.kum_work_sheet:
                logging.info(f"обработка листа {lst_name_open}")
                self.df, self.sheet_names = open_dataframe(self.link, self.password, lst_name_open)         # открываем df
                self.df = predobrabotka_df(self.df)                                                         # находим шапку
                self.df = head_registr_low_strip(self.df)                                                   # шапку в нижний регистр
                self.df = rename_columns_individual(self.df, Manufactory_df.dict_rename_columns)            # переименовываем столбцы в единый стиль
                self.df = df_white_list_col(self.df, Manufactory_df.white_list_columns)                     # оставляем нужные столбцы и добавляем если нет столбцов из white_list_columns
                self.df = name_df_columns_and_marka(self.df, self.name, self.marka, lst_name_open, self.region)          # добавляем столбцы с имененм источника и маркой которая есть в названии источника и именем листа
                bank_dataframes.append(self.df)                                                             # складируем наш датафрейм из обработанного листа 
                
            self.df  = pd.concat(bank_dataframes)                                                           # объединяем все df
     
        
        elif 'lot' in self.kum_work_sheet and self.name not in Manufactory_df.list_individual_treatment:
            self.df, self.sheet_names = open_dataframe(self.link, self.password)             # открываем предварительное открытие считать все листы
            
            bank_dataframes = []
            for lst_name_open in self.sheet_names:
                if lst_name_open not in self.kum_not_work_sheet:
                    logging.info(f"обработка листа {lst_name_open}")
                    self.df, self.sheet_names = open_dataframe(self.link, self.password, lst_name_open)      # открываем df
                    self.df = predobrabotka_df(self.df)                                                      # находим шапку
                    self.df = head_registr_low_strip(self.df)                                                # шапку в нижний регистр
                    self.df = rename_columns_individual(self.df, Manufactory_df.dict_rename_columns)         # переименовываем столбцы в единый стиль
                    self.df = df_white_list_col(self.df, Manufactory_df.white_list_columns)                  # оставляем нужные столбцы и добавляем если нет столбцов из white_list_columns
                    self.df = name_df_columns_and_marka(self.df, self.name, self.marka, lst_name_open, self.region)       # добавляем столбцы с имененм источника и маркой которая есть в названии источника и именем листа
                    bank_dataframes.append(self.df)                                                          # складируем наш датафрейм из обработанного листа 
                    
            self.df  = pd.concat(bank_dataframes)                                                            # объединяем все df
            
            
        elif self.name == 'KUM_OVP_vved_MSK':
            logging.info(f"обработка листа {self.kum_work_sheet[0]}")
            self.df, self.sheet_names = open_dataframe(self.link, self.password, self.kum_work_sheet[0])    # открываем предварительное открытие считать все листы
            self.df = predobrabotka_df(self.df)                                                             # находим шапку
            self.df = head_registr_low_strip(self.df)                                                       # шапку в нижний регистр
            self.df = rename_columns_individual(self.df, Manufactory_df.dict_rename_columns)                # переименовываем столбцы в единый стиль
            self.df = df_white_list_col_OVP(self.df, Manufactory_df.white_list_columns_ovp,  Manufactory_df.white_list_columns)                 # оставляем нужные столбцы и добавляем если нет столбцов из white_list_columns
            self.df = name_df_columns_and_marka(self.df, self.name, self.marka, self.kum_work_sheet[0], self.region)       # добавляем столбцы с имененм источника и маркой которая есть в названии источника и именем листа
            
        # написать для хендай баик москва
        elif self.name == 'KUM_HYUNDAI_BAIC_UKA_varsh_MSK':
            logging.info(f"обработка листа {self.kum_work_sheet[0]}")
            self.df, self.sheet_names = open_dataframe(self.link, self.password, self.kum_work_sheet[0])    # открываем предварительное открытие считать все листы
            self.df = predobrabotka_df(self.df)                                                             # находим шапку
            self.df = head_registr_low_strip(self.df)                                                       # шапку в нижний регистр
            self.df = rename_columns_individual(self.df, Manufactory_df.dict_rename_columns)                # переименовываем столбцы в единый стиль
            self.df = df_white_list_col_H_B_U_v_MSK(self.df, Manufactory_df.white_list_columns_ovp_HYUNDAI_BAIC_UKA_varsh_MSK,  Manufactory_df.white_list_columns) # оставляем нужные столбцы и добавляем если нет столбцов из white_list_columns
            self.df = name_df_columns_and_marka(self.df, self.name, self.marka, self.kum_work_sheet[0], self.region)       # добавляем столбцы с имененм источника и маркой которая есть в названии источника и именем листа
            
            
            
    def start(self):
        logging.info(f"{self.start.__name__} - ЗАПУСК")
        if self.flag == True:
            self.manufactory()
            
    
catalog_manufactory_df = {} # справочник с объектами класса
logging.info(f"заполняем справочник объектами класса")

for i in dict_link_name.keys():
    catalog_manufactory_df[i] = Manufactory_df(i, 
                                               dict_link_name[i]['link'], 
                                               str(dict_link_name[i]['pass']), 
                                               dict_link_name[i]['region'], 
                                               dict_link_name[i]['kum_work_sheet'], 
                                               dict_link_name[i]['kum_not_work_sheet'], 
                                               dict_link_name[i]['marka'])
    
    
logging.info(f"собираем все df в один список")
frames_kum = [catalog_manufactory_df[i].df for i in catalog_manufactory_df.keys()]

logging.info(f"конкатинируем в один df")
result_svod = pd.concat(frames_kum)

logging.info(f"создаем копию df")
result_svod_pred = result_svod.copy()

# очищаем df от мусора
logging.info(f"очищаем df от мусора")
result_svod_pred = result_svod_pred[(result_svod_pred['vin'].notna()) 
                                    & (result_svod_pred['дата выдачи'].notna()) 
                                    & (result_svod_pred['vin'] !=0) 
                                    & (result_svod_pred['дата выдачи'] !=0)
                                    & (result_svod_pred['vin'] !="0") 
                                    & (result_svod_pred['дата выдачи'] !="0") 
                                    & (result_svod_pred['vin'].apply(lambda x: len(str(x))>2))
                                    & (result_svod_pred['vin'].apply(lambda x: str(x) != '00:00:00'))
                                    & (result_svod_pred['дата выдачи'].apply(lambda x: str(x) != '00:00:00'))]

# приводим столбец с датами к общему формату
logging.info(f"приводим столбец с датами к общему формату")
result_svod_pred = conversorrrrrr_date(result_svod_pred, 'дата выдачи')

# столбцы с числовыми значениями к одному формату данных 
logging.info(f"приводим столбцы с числовыми значениями к одному формату данных")
result_svod_pred = conversion_columns_integer(result_svod_pred, ['итого ам доход', 'до доход', 'доход финуслуги', 'доход трейдин', 'кум доход итого', 'доход next', 'доход пргм привилегий'])

# распеределяем нал/кредит
logging.info(f"распеределяем нал/кредит - применение функции lambda & form_pay")
result_svod_pred['форма оплаты'] = result_svod_pred['форма оплаты'].apply(lambda x: 'нал' if str(x).isdigit() == True else form_pay(x))

# заменим Nan для дальнейших действий
logging.info(f"заменим Nan на - неизвестно")
result_svod_pred['салон'] = result_svod_pred['салон'].fillna('неизвестно')

# отделим Рыбинск из Ярославля
logging.info(f"отделим Рыбинск из Ярославля - применение функции lambda & reg_test")
result_svod_pred['регион'] = result_svod_pred.apply(lambda x: reg_test(x.регион, x.салон), axis=1)

logging.info(f"отделим автоцентры - добавим столбец - применение функции raspred_salon_marki")
result_svod_pred['автоцентр'] = result_svod_pred.apply(lambda x: raspred_salon_marki(x.марки_бд, x.салон, x.модель, x.регион), axis=1)

logging.info(f"объединим доход трейдин + доход next")
result_svod_pred['доход трейдин'] = result_svod_pred['доход трейдин'] + result_svod_pred['доход next']
logging.info(f"столбец - доход next - удалим")
del result_svod_pred['доход next']

logging.info(f"добавим столбцы - ти_факт_ам (дох трейд не равно 0) и факт_ам (просто 1 так как данные очищены)")
result_svod_pred['ти_факт_ам'] = result_svod_pred['доход трейдин'].apply(lambda x: 1 if int(x) != 0 else 0)
result_svod_pred['факт_ам'] = 1

logging.info(f"добавим столбец - корп_розн - применение функции lambda & korp_rozn")
result_svod_pred['корп_розн'] = result_svod_pred['клиент'].apply(lambda x: korp_rozn(x))

logging.info(f"сохраняем файл - //sim.local/data/Varsh/OFFICE/CAGROUP/run_python/task_scheduler/kum/result_svod.xlsx")
result_svod_pred.to_excel('//sim.local/data/Varsh/OFFICE/CAGROUP/run_python/task_scheduler/kum/result_svod.xlsx')

# Обновляем сводные таблицы
logging.info(f"Обновляем сводные таблицы")
try:
    xlapp = win32com.client.DispatchEx("Excel.Application")
    wb = xlapp.Workbooks.Open("//sim.local/data/Varsh/OFFICE/CAGROUP/run_python/task_scheduler/kum/КУМ_ОБЩИЙ.xlsx")
    wb.Application.AskToUpdateLinks = False   # разрешает автоматическое  обновление связей (файл - парметры - дополнительно - общие - убирает галку запрашивать об обновлениях связей)
    wb.Application.DisplayAlerts = True  # отображает панель обновления иногда из-за перекрестного открытия предлагает ручной выбор обновления True - показать панель
    wb.RefreshAll()
    #xlapp.CalculateUntilAsyncQueriesDone() # удержит программу и дождется завершения обновления. было прописано time.sleep(30)
    time.sleep(40) # задержка 60 секунд, чтоб уж точно обновились сводные wb.RefreshAll() - иначе будет ошибка 
    wb.Application.AskToUpdateLinks = True   # запрещает автоматическое  обновление связей / то есть в настройках экселя (ставим галку обратно)
    wb.Save()
    wb.Close()
    xlapp.Quit()
    wb = None # обнуляем сслыки переменных иначе процесс эксель не завершается и висит в дистпетчере
    xlapp = None # обнуляем сслыки переменных иначе процесс эксел ь не завершается и висит в дистпетчере
    del wb # удаляем сслыки переменных иначе процесс эксель не завершается и висит в дистпетчере
    del xlapp # удаляем сслыки переменных иначе процесс эксель не завершается и висит в дистпетчере
    logging.info(f"сводные таблицы - обновлены")
except:
    logging.error(f"ОШИБКА", exc_info=True)
    

# список с адресами рассылки
lst_email = read_email_adress() # 'skrutko@sim-auto.ru'
lst_email_error = ['skrutko@sim-auto.ru'] # есть ошибки

# запуск функции рассылки почты
logging.info(f"детектим ошибки, проверяем почту")
sending_mail(lst_email, lst_email_error)
logging.info(f"почта отправлена")

time_wopking_skript(start_time_work_skript)
logging.info(f"время выполнения скрипта {time_wopking_skript(start_time_work_skript)}")



