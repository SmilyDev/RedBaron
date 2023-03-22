# pip install pandas openpyxl
# pip install pandas jupyter pandarallel requests tqdm
# conda install xlwings
# conda install docplex
# conda install pandas
# pip install pandas openpyxl

import xlwings as xw
import pandas as pd
import time
import os
from docplex.mp.model import Model
from docplex.mp.conflict_refiner import ConflictRefiner
import numpy as np
import math
import xlrd
from zipfile import ZipFile
import shutil
import datetime
import sys

#Функция логов
def log_print(text, log_file):
    with open(log_file, 'a') as f:
        print(text, file=f)
    print(text)
    
#функция загрузки файлов
def load_xlsx(file_name, skiprows, log_file):
    log_print('{0} : Загружаем {1}'.format(time.strftime("%d/%m/%Y, %H:%M:%S",time.localtime()), file_name), log_file)
    # Создаем временную папку
    tmp_folder = '/tmp/convert_wrong_excel/'
    os.makedirs(tmp_folder, exist_ok=True)

    # Распаковываем excel как zip в нашу временную папку
    with ZipFile(file_name) as excel_container:
        excel_container.extractall(tmp_folder)

    # Переименовываем файл с неверным названием
    wrong_file_path = os.path.join(tmp_folder, 'xl', 'SharedStrings.xml')
    correct_file_path = os.path.join(tmp_folder, 'xl', 'sharedStrings.xml')
    os.rename(wrong_file_path, correct_file_path) 

    # Запаковываем excel обратно в zip и переименовываем в исходный файл
    shutil.make_archive('yourfile', 'zip', tmp_folder)
    os.remove(file_name)
    os.rename('yourfile.zip', file_name)
    return pd.read_excel(file_name, engine='openpyxl', header = 0, skiprows=skiprows)
    
#Определяем предыдущую операцию (функция)
def pred_op(x, BD):
    #Определяем сколько строк с такми-же показателями
    TEMP = BD[(BD['Номенклатура вых изд'] == x['Номенклатура вых изд']) 
     & (BD['Характеристика номенклатуры вых изд'] == x['Характеристика номенклатуры вых изд']) 
     & (BD['Номер операции тех карта'] == x['Номер операции тех карта'])
     & (BD.index < x.name)]
    if len(TEMP) > 0:
        #Для одинаковых номеров техопераций берем предыдущий индекс
        TEMP = TEMP[TEMP.index == TEMP.index.max()].index
    else:
        #Определяем по тем-же характиристикам с меньшими номерами операций
        TEMP = BD[(BD['Номенклатура вых изд'] == x['Номенклатура вых изд']) 
         & (BD['Характеристика номенклатуры вых изд'] == x['Характеристика номенклатуры вых изд']) 
         & (BD['Номер операции тех карта'] < x['Номер операции тех карта'])]
        TEMP = TEMP[TEMP['Номер операции тех карта'] == TEMP['Номер операции тех карта'].max()].index

        if len(TEMP) == 0:
            #Если не нашли, определяем по исх комплектующим
            if pd.notna(x['Номенклатура исх комп']):
                TEMP = BD[(BD['Номенклатура вых изд'] == x['Номенклатура исх комп']) 
                 & (BD['Характеристика номенклатуры вых изд'] == x['Характеристика номенклатуры исх комп'])]
                TEMP = TEMP[TEMP['Номер операции тех карта'] == TEMP['Номер операции тех карта'].max()].index
    if len(TEMP) > 0:
        return TEMP[-1]

#Вносим остатки (функция)
def ost_conect(x, SPEC):
    TEMP = SPEC[(SPEC['Номенклатура вых изд'] == x['Номенклатура']) & (SPEC['Характеристика номенклатуры вых изд'] == x['Характеристика номенклатуры'])]
    TEMP = TEMP[TEMP['Номер операции тех карта'] == TEMP['Номер операции тех карта'].max()].index
    if len(TEMP) > 0:
        return TEMP[-1]
     
        
#Расчет коэфициента (функция)
def K_count(x, cur_date, max_date, PZ):
    K = 0
    PZ_temp = PZ[PZ['Номенклатура'] == x['Номенклатура вых изд']]
    date_temp = x['День']
    if len(PZ_temp) > 0:
        K +=  (PZ_temp.apply(lambda x2: ((max_date - x2['Дата сдачи']).days + 1) * x2['Количество Остаток'], axis = 1).sum() 
               / PZ_temp['Количество Остаток'].sum() * (x['Номер дня'] +1))
    else:
        K = 0.01 * (x['Номер дня']+1)
    return K
    
    
#Определяем реальные сроки ПЗ (функция)
def real_PZ(x, BD2, PZ):
    temp_BD = BD2[(BD2['Номенклатура вых изд'] == x['Номенклатура']) & 
                 (BD2['Характеристика номенклатуры вых изд'] == x['Характеристика номенклатуры'])]
    sum_PZ = PZ[(PZ['Номенклатура'] == x['Номенклатура']) & 
                (PZ['Характеристика номенклатуры'] == x['Характеристика номенклатуры']) &
                (PZ['Дата сдачи'] <= x['Дата сдачи'])]['Количество Остаток'].sum() - 1
    for ind in temp_BD.index:
        if temp_BD.loc[ind, 'Остаток на конец'] >= sum_PZ:
            return temp_BD.loc[ind, 'День']  
    


#Сохраняем конфликты (функция)
def save_conflicts(res, log_file):
        """ Displays all conflicts.

        """
        log_print('conflict(s): {0}'.format(res.number_of_conflicts), log_file)
        for conflict in res.iter_conflicts():
            st = conflict.status
            elt = conflict.element
            if hasattr(conflict.element, 'as_constraint'):
                ct = conflict.element.as_constraint()
                label = elt.short_typename
            else:
                ct = elt
                label = ct.__class__.__name__
            log_print("  - status: {1}, {0}: {2!s}".format(label, st.name, ct.to_readable_string()), log_file)
            
            
 
 #Определяение промежуточной строки (функция)
def prom_str(x, SPEC):
    
    if len(x['Строки потребления']) == 1:
        if ((x['Номенклатура вых изд'] == SPEC.loc[x['Строки потребления'][0], 'Номенклатура вых изд']) 
            & (x['Характеристика номенклатуры вых изд'] == 
               SPEC.loc[x['Строки потребления'][0], 'Характеристика номенклатуры вых изд'])):
            return 1
    return 0
    
    
#Отнимает от потребностей ТМЦ остатки ТМЦ (функция)
def cor_potr(x, OST_TMC):
    ost_sum = OST_TMC[(OST_TMC['Номенклатура'] == x['Номенклатура']) & (OST_TMC['Характеристика номенклатуры'] == x['Характеристика номенклатуры'])]['Количество Остаток'].sum()
    if x['Количество Конечный остаток'] <= ost_sum:
        return 0
    else:
        return x['Количество Конечный остаток'] - ost_sum


def Start():
    
    #Подключаемся к файлу
    try:
        wb = xw.Book.caller()
    except:
        wb = xw.Book('C:\\Users\\mgvse\\КЭАЗ\\Решатель КЭАЗ.xlsb')
    #Определяем путь к файлу
    path=wb.fullname
    pathLog=path[:(path.rfind('\\')+1)]
    pathLog=pathLog+'logs'
    log_file = pathLog + '\\logs'+time.strftime("%Y%m%d-%H%M%S")+'.txt'

    #Если папки LOGS нет, то создаем
    if not os.path.exists(pathLog):
            os.mkdir(pathLog)
    #Вносим данные из настроек интерфейса
    SPEC_path = xw.Range('SPEC_path').options().value
    PZ_path = xw.Range('PZ_path').options().value
    KOLOBOR_path = xw.Range('KOLOBOR_path').options().value
    OST_path = xw.Range('OST_path').options().value
    TIME_path = xw.Range('TIME_Path').options(pd.DataFrame, index=False, header=0).value
    KALEND = xw.Range('Таб_календ').options(pd.DataFrame, index=False, header=0).value
    gor_plan = xw.Range('gor_plan').options().value
    gor_plan_max = xw.Range('gor_plan_max').options().value
    #priority = xw.Range('priority').options().value
    #raspred = xw.Range('raspred').options().value
    tab_Ogr=xw.Range('Таб_Ограничения').options(pd.DataFrame, index=False, header=0).value 
    tab_Ogr_zag=xw.Range('Таб_Ограничения_Загаловки').options(pd.DataFrame, index=False, header=0).value
    timelimit =xw.Range('timelimit').options().value
    res_ceh_flag =xw.Range('res_ceh_flag').options().value
    potr_tmc_path = xw.Range('potr_tmc_path').options().value
    ost_tmc_path = xw.Range('ost_tmc_path').options().value
    
    
    
    #Загружаем файлы
    SPEC = load_xlsx(SPEC_path[0], int(SPEC_path[1]-2), log_file)
    PZ = load_xlsx(PZ_path[0], int(PZ_path[1]-2), log_file)
    KOLOBOR = load_xlsx(KOLOBOR_path[0], int(KOLOBOR_path[1]-2), log_file)
    OST = load_xlsx(OST_path[0], int(OST_path[1]-2), log_file)
    POTR = load_xlsx(potr_tmc_path[0], int(potr_tmc_path[1]-2), log_file)
    OST_TMC = load_xlsx(ost_tmc_path[0], int(ost_tmc_path[1]-2), log_file)
    for i in TIME_path[0]:
        TIME = load_xlsx(i, 3, log_file)
        
        
    
        
    #Удаляем пустые строки в кол-ве оборудования и дубликаты
    log_print('{0} : Удаляем дубликаты и лишние столбцы'.format(time.strftime("%d/%m/%Y, %H:%M:%S",time.localtime())), log_file)
    #KOLOBOR.drop(KOLOBOR[KOLOBOR['Количество оборудования'].isna()].index, inplace= True)
    KOLOBOR= KOLOBOR.groupby(by = ['Подразделение', 'Рабочий центр'], dropna = True).sum()
    KOLOBOR.reset_index(inplace = True)
    KOLOBOR['Рабочий центр'] = KOLOBOR['Рабочий центр'].str.strip()
    TIME['Рабочий центр'] = TIME['Рабочий центр'].str.strip()
    SPEC['Рабочий центр'] = SPEC['Рабочий центр'].str.strip()
    SPEC['Номенклатура вых изд'] = SPEC['Номенклатура вых изд'].str.strip()
    PZ['Номенклатура'] = PZ['Номенклатура'].str.strip()
    SPEC['Характеристика номенклатуры вых изд'] = SPEC['Характеристика номенклатуры вых изд'].str.strip()
    PZ['Характеристика номенклатуры'] = PZ['Характеристика номенклатуры'].str.strip()
    SPEC['Номенклатура исх комп'] = SPEC['Номенклатура исх комп'].str.strip()
    SPEC['Характеристика номенклатуры исх комп'] = SPEC['Характеристика номенклатуры исх комп'].str.strip()
    PZ.drop(PZ[PZ['Количество Остаток'] <= 0].index, inplace = True) #Удаляем строки ПЗ >= 0
    POTR = POTR[POTR['Количество Конечный остаток'] > 0]
    KALEND = KALEND[0]
    KALEND = KALEND.dt.date
    #заменяем шапку времени работы на даты
    fromcol = TIME.columns[2:].values
    tocol =  pd.Series(TIME.columns[2:]).dt.date.values
    dictionary = dict(zip(fromcol, tocol))
    TIME.rename(columns=dictionary, inplace = True)
    #Удаляем лишние столбцы
    try:
        SPEC = SPEC.drop(columns = [np.nan, 'nan.1', 'nan.2'])
    except: 
        pass
    try:
        POTR = POTR.drop(columns = [np.nan, 'nan.1', 'nan.2'])
    except: 
        pass
    try:
        POTR.drop(columns = ['Unnamed: 1', 'Unnamed: 3', 'Unnamed: 4', 'Unnamed: 7', 'Unnamed: 8'], inplace = True)
    except: 
        pass
    try:
        PZ.drop(columns = ['np.nan', 'nan.1'], inplace = True)
    except: 
        pass
    PZ['Дата сдачи'] = pd.to_datetime(PZ['Дата сдачи'], format = '%d.%m.%Y') 
    PZ['Дата сдачи'] = PZ['Дата сдачи'].map(datetime.datetime.date)
    POTR['Период'] = pd.to_datetime(POTR['Период'], format = '%d.%m.%Y %H:%M:%S').dt.date
    POTR.reset_index(inplace = True, drop = True)
    POTR = POTR[POTR['Количество Конечный остаток'] > 0]
    

    
    #Удаляем из ПЗ строки, по которым нет производства в секунду
    temp = SPEC[(SPEC['Время выполнения (сек)'].isna()) & 
                (SPEC['Номенклатура вых изд'].isin(PZ['Номенклатура']))]['Номенклатура вых изд'].drop_duplicates()
    #temp = PZ[PZ['Номенклатура'].isin(temp)].reset_index(drop=True)['Номенклатура']
    if len(temp) > 0:
        log_print('Внимание!!! Нижеприведенные номенклатуры удалены из производственного задания, '
                  +'т.к. по ним нет даных времени выполнения (сек).', log_file)
        for i in temp:
            log_print(i, log_file)
            #PZ.drop(PZ['Номенклатура'] == i, inplace = True)
        PZ = PZ[~PZ['Номенклатура'].isin(temp)].reset_index(drop=True)
            
    SPEC['Время выполнения (сек)'].fillna(0.001, inplace = True) #Если нет времени работы, ставим 0.001 секунд          
        
        
        
    #Отбираем номенклатуру которая есть в ПЗ или в исходных номенклатурах
    log_print('{0} : Отбираем номенклатуру которая есть в ПЗ'.format(time.strftime("%d/%m/%Y, %H:%M:%S",time.localtime())), log_file)
    SPEC2 = pd.DataFrame()
    nom_temp = PZ['Номенклатура']
    nom_temp.dropna(inplace = True)
    SPEC_temp = SPEC[SPEC['Номенклатура вых изд'].isin(nom_temp)]
    SPEC2 = SPEC_temp.copy()

    while len(nom_temp) > 0:
        nom_temp = SPEC_temp['Номенклатура исх комп']
        nom_temp.dropna(inplace = True)
        nom_temp.drop_duplicates(inplace = True)
        SPEC_temp = SPEC[(SPEC['Номенклатура вых изд'].isin(nom_temp)) & 
                         (SPEC['Номенклатура вых изд'] != SPEC['Номенклатура исх комп'])]
        SPEC2 = SPEC2.append(SPEC_temp)
        
    SPEC = SPEC2.drop_duplicates().copy()
    SPEC.reset_index(inplace = True, drop = True)  

    nom_temp = SPEC['Номенклатура вых изд'].drop_duplicates()
    POTR = POTR[POTR['Номенклатура'].isin(nom_temp)]
    POTR.reset_index(inplace = True, drop = True)
    
    
    
    #Считаем производство в минуту
    SPEC['Производство в минуту'] = SPEC.apply(lambda x: round(x['Количество вых изд'] / x['Время выполнения (сек)'] * 60, 2), axis = 1)
    
    
     
    #Определяем предыдущую операцию
    log_print('{0} : Определяем предыдущую операцию'.format(time.strftime("%d/%m/%Y, %H:%M:%S",time.localtime())), log_file)
    SPEC['предыдущая операция'] = SPEC.apply(pred_op, BD = SPEC, axis = 1)
    
    
         
    #Вносим остатки
    log_print('{0} : Вносим остатки'.format(time.strftime("%d/%m/%Y, %H:%M:%S",time.localtime())), log_file)
    OST['Индекс в базе'] = OST.apply(ost_conect, SPEC = SPEC, axis = 1)
    for n in OST.index:
        if not np.isnan(OST.loc[n, 'Индекс в базе']):
            SPEC.loc[OST.loc[n, 'Индекс в базе'], 'Остаток на начало'] = OST.loc[n, 'Количество Остаток']
    SPEC['Остаток на начало'].fillna(0, inplace = True) 
    
    
    
    #Вносим потребление
    log_print('{0} : Вносим потребление'.format(time.strftime("%d/%m/%Y, %H:%M:%S",time.localtime())), log_file)
    SPEC['Строки потребления'] = SPEC.apply(lambda x, SPEC: SPEC[SPEC['предыдущая операция'] == x.name].index.values.tolist(), args = [SPEC], axis = 1)
    
    
    
    #Определяение промежуточной строки
    log_print('{0} : Определяение промежуточной строки'.format(time.strftime("%d/%m/%Y, %H:%M:%S",time.localtime())), log_file)
    SPEC['Промежуточная операция'] = SPEC.apply(prom_str, axis = 1, args = [SPEC])
    
    
    
    #Определяем строки в Потреблении
    log_print('{0} : Определяем строки в потреблении.'.format(time.strftime("%d/%m/%Y, %H:%M:%S",time.localtime())), log_file)
    POTR['Индекс в базе'] = POTR.apply(ost_conect, SPEC = SPEC, axis = 1)
    POTR.dropna(subset = ['Индекс в базе'], inplace = True)
    
    
    
    #Определяем строки ПЗ и кол-во оборудования
    log_print('{0} : Определяем строки ПЗ и кол-во оборудования'.format(time.strftime("%d/%m/%Y, %H:%M:%S",time.localtime())), log_file)
    PZ['Индекс в базе'] = PZ.apply(ost_conect, SPEC = SPEC, axis = 1)
    SPEC['строка ПЗ'] = SPEC.apply(lambda x, PZ: PZ[PZ['Индекс в базе'] == x.name].index.tolist(), axis = 1, args = [PZ])
    SPEC['ПЗ'] = SPEC.apply(lambda x, PZ: PZ.loc[x['строка ПЗ'], 'Количество Остаток'].sum(), axis = 1, args = [PZ])
    SPEC['Кол-во оборуд'] = SPEC.apply(lambda x, KOLOBOR: KOLOBOR[KOLOBOR['Рабочий центр'] == x['Рабочий центр']]
                                     ['Количество оборудования'].max(), axis = 1, args = [KOLOBOR])
    
    
    
    #Определяем даты
    cur_date = xw.Range('cur_date').options().value.date()
    KALEND = KALEND[KALEND >= cur_date].reset_index(drop=True)
    #max_date = KALEND[KALEND >= PZ['Дата сдачи'].max()].reset_index(drop = True)[gor_plan]

    SPEC['Кол-во дней на ПЗ'] = SPEC.apply(lambda x, PZ: PZ[PZ['Номенклатура'] == x['Номенклатура вых изд']]['Количество Остаток'].sum() *
                                           x['Время выполнения (сек)'] / x['Кол-во оборуд'] / 60 / 60 / 8 , axis = 1, args = [PZ])
    kol_days_PZ = SPEC[['Рабочий центр', 'Кол-во дней на ПЗ']].groupby('Рабочий центр').sum()
    kol_days_PZ.sort_values(by = 'Кол-во дней на ПЗ', ascending=False, inplace = True)
    kol_days_PZ = kol_days_PZ.head(10)
    kol_days_PZ = kol_days_PZ.round({'Кол-во дней на ПЗ':0})
    kol_days_PZ.reset_index(inplace = True)
    kol_date = kol_days_PZ.loc[0, 'Кол-во дней на ПЗ'] + gor_plan
    max_date = KALEND[kol_date]
    log_print('{0} : Горизонт планирования - {1} дней ({2}):'.format(time.strftime("%d/%m/%Y, %H:%M:%S",time.localtime()), 
                                                               kol_date, max_date), log_file)
    log_print(kol_days_PZ, log_file)
    if kol_date > gor_plan_max:
        log_print('{0} : Горизонт планирования превысил максимальное значение.'.format(time.strftime("%d/%m/%Y, %H:%M:%S",time.localtime())), log_file)
        sys.exit()
    KALEND = KALEND[KALEND <= max_date]
    
    
    
    #Отнимаем от потребностей ТМЦ остатки ТМЦ
    log_print('{0} : Отнимаем от потребностей ТМЦ остатки ТМЦ'.format(time.strftime("%d/%m/%Y, %H:%M:%S",time.localtime())), log_file)
    POTR = POTR[POTR['Период'] <= max_date]
    POTR['Потребность с учетом остатка'] = POTR.apply(cor_potr, axis = 1, args = [OST_TMC])
    POTR = POTR[POTR['Потребность с учетом остатка'] > 0]
    POTR.reset_index(drop = True, inplace = True)
    
    
    
    #Проверка потребностей
    log_print('{0} : Проверка потребностей'.format(time.strftime("%d/%m/%Y, %H:%M:%S",time.localtime())), log_file)
    POTR['Кол-во дней'] = POTR.apply(lambda x : KALEND[KALEND == x['Период']].index.values.max() + 1, axis = 1)
    POTR['Среднедневная потребность'] = POTR.apply(lambda x: x['Потребность с учетом остатка'] / x['Кол-во дней'], axis = 1)
    POTR_MAX = POTR[['Номенклатура', 'Среднедневная потребность']].groupby('Номенклатура').max()
    POTR_MAX.reset_index(inplace = True)
    SPEC_TEMP = SPEC[['Номенклатура вых изд', 'Рабочий центр', 'Производство в минуту', 'Кол-во оборуд']].drop_duplicates()
    SPEC_TEMP['Среднедневная потребность'] = SPEC_TEMP.apply(
        lambda x : POTR_MAX[POTR_MAX['Номенклатура'] == x['Номенклатура вых изд']]['Среднедневная потребность'].max(), axis = 1
    ).fillna(0)
    SPEC_TEMP['Средняя потребность, ч'] = (SPEC_TEMP['Среднедневная потребность'] 
                                                / SPEC_TEMP['Кол-во оборуд'] 
                                                / SPEC_TEMP['Производство в минуту']
                                                / 60)
    RC_TEMP_F = SPEC_TEMP[['Рабочий центр', 'Средняя потребность, ч']].groupby('Рабочий центр').sum()
    RC_TEMP_F.sort_values(by = 'Средняя потребность, ч', ascending=False, inplace = True)
    RC_TEMP = RC_TEMP_F[RC_TEMP_F['Средняя потребность, ч'] > 8]
    RC_TEMP.reset_index(inplace = True)
    if len(RC_TEMP) > 0:
        log_print('{0} : Потребности сборочного цеха неосуществимы.'.format(time.strftime("%d/%m/%Y, %H:%M:%S",time.localtime()))
                  , log_file)
        log_print(RC_TEMP, log_file)
        log_print('----------------------------------------', log_file)
        log_print(SPEC_TEMP[['Номенклатура вых изд', 'Рабочий центр', 'Средняя потребность, ч']]
                  [SPEC_TEMP['Рабочий центр'].isin(RC_TEMP['Рабочий центр'])], log_file)
        sys.exit()
    
    
    
    #Разбиваем по дням
    log_print('{0} : Формируем базу данных'.format(time.strftime("%d/%m/%Y, %H:%M:%S",time.localtime())), log_file)
    BD = pd.DataFrame()
    day_len = len(SPEC)
    for day in KALEND.index:
        log_print(KALEND[day] , log_file)
        TEMP = SPEC.copy()
        TEMP['День'] = KALEND[day]
        TEMP['Номер дня'] = day
        #определяем сумму коэффициента превышения ПЗ
        #Ищем ПЗ с наименьшей датой
        TEMP['минимальная дата ПЗ'] = TEMP.apply(lambda x, PZ: PZ.loc[x['строка ПЗ']]['Дата сдачи'].min(), 
                                                 axis = 1, args = [PZ])
        #TEMP['минимальная дата ПЗ'].fillna(max_date, inplace = True)
        TEMP['Коэффициент'] =TEMP.apply(K_count, axis = 1, args = [cur_date, max_date, PZ])
        TEMP['предыдущая операция'] = SPEC['предыдущая операция'] + day_len * day
        TEMP['Строки потребления'] = TEMP.apply(lambda x, day_len, day: [i + day_len * day for i in x['Строки потребления']], args = [day_len, day], axis = 1)
        BD = pd.concat([BD, TEMP])
    BD.reset_index(inplace = True, drop=True)
    
    
    
    #Инициализируем модель
    log_print('{0} : Инициализируем модель'.format(time.strftime("%d/%m/%Y, %H:%M:%S",time.localtime())), log_file)
    model = Model(name='Optim_KEAZ')
        
        
        
    #Определяем переменные
    log_print('{0} : Определяем переменные'.format(time.strftime("%d/%m/%Y, %H:%M:%S",time.localtime())), log_file)
    vars_time = []
    vars_ost = []
    for n_str in range(len(BD)):
        vars_time.append(model.continuous_var(lb = 0, name = 'var_time_'+str(n_str%day_len)+'_'+str(n_str//day_len))) #Время производства
        vars_ost.append(model.continuous_var(lb = 0, name = 'var_ost_'+str(n_str%day_len)+'_'+str(n_str//day_len))) #Остаток на конец дня
        
        
     
    #Расчет остатка
    log_print('{0} : Расчет остатка'.format(time.strftime("%d/%m/%Y, %H:%M:%S",time.localtime())), log_file)
    for n_str in range(len(BD)):   
        model.add_constraint(vars_ost[n_str] == (SPEC.loc[n_str, 'Остаток на начало'] if BD.loc[n_str, 'Номер дня'] == 0 else
                            vars_ost[n_str - day_len]) + vars_time[n_str] * BD.loc[n_str, 'Производство в минуту']  - 
                             (model.sum([vars_time[i] * BD.loc[i, 'Производство в минуту'] for i in BD.loc[n_str, 'Строки потребления']]) 
                              if len(BD.loc[n_str, 'Строки потребления']) > 0 else 0), ctname = 'con_ost_' +str(n_str%day_len)+'_'+str(n_str//day_len))
                              
        
    
    #Ограничения (время в день)
    log_print('{0} : Вносим ограничения (время в день)'.format(time.strftime("%d/%m/%Y, %H:%M:%S",time.localtime())), log_file)
    TEMP_TIME = SPEC['Рабочий центр'].drop_duplicates()
    TEMP_TIME.reset_index(inplace = True, drop = True)
    n = 0
    for day in KALEND:
        log_print(day, log_file)
        TEMP_DAY = BD[BD['День'] == day]
        for n_rc in TEMP_TIME.index:
            n+=1
            kol_ob = KOLOBOR[(KOLOBOR['Рабочий центр'] == TEMP_TIME.loc[n_rc]) 
                            ]['Количество оборудования'].max()
            if np.isnan(kol_ob):
                kol_ob = 1
            time_work_temp = TIME[(TIME['Рабочий центр'] == TEMP_TIME.loc[n_rc]) 
                                 ].reset_index(drop=True)
            try:
                time_work = time_work_temp.loc[0, day]
            except:
                time_work = 8
            
            timeForAll = kol_ob * time_work * 60
            var_ind_temp = TEMP_DAY[(TEMP_DAY['Рабочий центр'] == TEMP_TIME.loc[n_rc]) 
                                   ].index.to_list()
            if len(var_ind_temp) > 0:
                model.add_constraint(model.sum(vars_time[i] for i in var_ind_temp) <= int(timeForAll), ctname = 'con_time_'+str(n))
    
    
    
    # Ограничение кол-во ПЗ <= остаток на последний день
    log_print('{0} : Вносим ограничения кол-во ПЗ <= остаток на последний день'
              .format(time.strftime("%d/%m/%Y, %H:%M:%S",time.localtime())), log_file)
    TEMP_DAY = BD[(BD['День'] == max_date) & (BD['строка ПЗ'].map(len) > 0)]
    for n_str in TEMP_DAY.index:
        PZ_sum = PZ.loc[TEMP_DAY.loc[n_str, 'строка ПЗ'], 'Количество Остаток'].sum()
        model.add_constraint(vars_ost[n_str] >= PZ_sum, ctname = 'con_PZ_'+str(n_str))
    
    
    
    #Ограничение по потребности ТМЦ
    log_print('{0} : Вносим ограничения по потребности ТМЦ для сборочного цеха'
              .format(time.strftime("%d/%m/%Y, %H:%M:%S",time.localtime())), log_file)\

    POTR['Индекс в базе 2'] = POTR.apply(lambda x: x['Индекс в базе'] + day_len * (x['Кол-во дней'] - 1), axis = 1)

    ind = POTR['Индекс в базе 2']
    potr = POTR['Потребность с учетом остатка']

    for n_str in POTR.index:
        model.add_constraint(vars_ost[int(ind[n_str])] >= potr[n_str], ctname = 'con_TMC_' + str(n_str))
        
        
        
    # Ограничение все промежуточные операции имеют конечный остаток 0
    log_print('{0} : Ограничение все промежуточные операции имеют конечный остаток 0'
              .format(time.strftime("%d/%m/%Y, %H:%M:%S",time.localtime())), log_file)

    TEMP_BD = BD[BD['Промежуточная операция'] == 1]
    for n_str in TEMP_BD.index:
        model.add_constraint(vars_ost[n_str] == 0, ctname = 'con_PO_'+str(n_str))    
        
        
        
        
    #Вносим таблицу ограничений
    log_print('{0} : Вносим универсальные ограничения'.format(time.strftime("%d/%m/%Y, %H:%M:%S",time.localtime())), log_file)
    tab_Ogr_zag_list= tab_Ogr_zag.values.tolist()
    tab_Ogr.columns = tab_Ogr_zag_list[0]
    #del tab_Ogr_zag_list
    #del tab_Ogr_zag
    try:
        tab_Ogr['День'] = tab_Ogr['День'].map(datetime.datetime.date)
    except:
        pass
        
        
        
    
    #Универсальный ограничитель
    for n_ogr in tab_Ogr.index:
        ogr_col_str = ''
        for ogr_col in tab_Ogr.columns[1:-2]:
            if pd.notnull(tab_Ogr.at[n_ogr, ogr_col]): 
                ogr_col_str = ogr_col_str + ' & ' if ogr_col_str != '' else '' #Добавляем к строке &
                ogr_value_str =   str(tab_Ogr.at[n_ogr, ogr_col]) if (not isinstance(tab_Ogr.at[n_ogr, ogr_col], str)) and (not isinstance(tab_Ogr.at[n_ogr, ogr_col], datetime.date)) else '"' +  tab_Ogr.at[n_ogr, ogr_col] + '"'        #Ограничение
                ogr_col_str = ogr_col_str + '`' +str(ogr_col) + '` == ' + ogr_value_str #Формируем строку формулы

        if ogr_col_str == "":
            BD_ind = BD.index.values.tolist()
        else:
            BD_TEMP = BD.query(ogr_col_str)
            if isinstance(tab_Ogr.loc[n_ogr, 'День'], datetime.date):
                BD_ind = BD_TEMP[BD_TEMP['День'] == tab_Ogr.loc[n_ogr, 'День']].index.values.tolist()
            else:
                BD_ind = BD_TEMP.index.values.tolist()
        if pd.notnull(tab_Ogr.at[n_ogr, 'Максимум, мин']):
            if tab_Ogr.at[n_ogr, 'Максимум, мин'] >= 0:
                model.add_constraint(model.sum(vars_time[i] for i in BD_ind) <= round(tab_Ogr.at[n_ogr, 'Максимум, мин']), ctname = 'Gr_const_max_' + str(n_ogr+2)) 
        if pd.notnull(tab_Ogr.at[n_ogr, 'Минимум, мин']):
            if tab_Ogr.at[n_ogr, 'Минимум, мин'] > 0:
                model.add_constraint(model.sum(vars_time[i] for i in BD_ind) >= round(tab_Ogr.at[n_ogr, 'Минимум, мин']), ctname = 'Gr_const_min_' + str(n_ogr+2))
                
                
    
    
    #Целевая функция
    func = model.sum(vars_time[i] * BD.loc[i, 'Коэффициент'] for i in BD.index)




    #Ограничения времени 
    model.parameters.timelimit.set(timelimit)
    model.parameters.read.datacheck = 2

    #Выводим конфликты
    refiner = ConflictRefiner() 
    res = refiner.refine_conflict(model)

    log_print('{0} : Запускаем решатель'.format(time.strftime("%d/%m/%Y, %H:%M:%S",time.localtime())), log_file)

    #Запускаем считалку
    model.minimize(func)
    model.solve(log_output = pathLog + "\\docplex_logs_"+time.strftime("%Y%m%d-%H%M%S")+".txt")
    model.export_to_stream(pathLog + "\\formulas_"+time.strftime("%Y%m%d-%H%M%S")+' model')
    save_conflicts(res, log_file)
    
    
    
    
    #Выгружаем результат

    log_print('{0} : Формируем результат'.format(time.strftime("%d/%m/%Y, %H:%M:%S",time.localtime())), log_file)
    xw.Range('GAP_факт').options(index=False, header=0).value=model.solve_details.mip_relative_gap
    xw.Range('res_time').options(index=False, header=0).value=model.solve_details.time
    BD['Переменная остатка'] = BD.apply(lambda x, vars_ost: vars_ost[x.name].name, axis = 1, args = [vars_ost])
    BD['Переменная времени'] = BD.apply(lambda x, vars_time: vars_time[x.name].name, axis = 1, args = [vars_time])

    try:
        BD['Время работы'] = pd.Series(model.solution.get_values(vars_time))#.apply(np.ceil)
        BD['Остаток на конец'] = pd.Series(model.solution.get_values(vars_ost))
        BD2 = BD[BD['Время работы'] > 0]
        #BD2 = BD.copy()
    except:
        BD2 = BD[BD['Номер дня'] == 0]
        pass
        

                
    #Определяем реальные сроки ПЗ
    log_print('{0} : Определяем реальные сроки ПЗ'.format(time.strftime("%d/%m/%Y, %H:%M:%S",time.localtime())), log_file)
    try:
        PZ['Дата сдачи (план)'] = PZ.apply(real_PZ, args = [BD2, PZ], axis = 1)
        PZ['Остаток на начало'] = PZ.apply(lambda x, OST: OST[(OST['Номенклатура'] == x['Номенклатура']) & 
                                (OST['Характеристика номенклатуры'] == 
                                 x['Характеристика номенклатуры'])]['Количество Остаток'].sum()
                                           ,axis = 1, args = [OST])
        BD2['Дата сдачи (план)'] = BD2.apply(lambda x, PZ: PZ.loc[x['строка ПЗ'], 'Дата сдачи (план)'].max(), args = [PZ], axis = 1)
        BD2['Объем производства'] = BD2.apply(lambda x: x['Время работы'] * x['Производство в минуту'], axis = 1).round()
        PZ['Объем производства'] = PZ.apply(lambda x, BD2: BD2[(BD2['Номенклатура вых изд'] == x['Номенклатура']) & 
                                                               (BD2['Характеристика номенклатуры вых изд'] == x['Характеристика номенклатуры']) &
                                                              (BD2['ПЗ'] > 0)]['Объем производства'].sum()
                                           , axis = 1 , args = [BD2])
        PZ['Остаток на конец'] = PZ.apply(lambda x: x['Остаток на начало'] + x['Объем производства'], axis = 1)
        BD2['Остаток на начало'] = BD2.apply(lambda x, BD, day_len: BD.loc[x.name - day_len]['Остаток на конец'] if x['Номер дня'] > 0 else
                                            x['Остаток на начало'], axis=1, args = [BD, day_len])
    except:
        pass

    



    #Сохраняем результат в отдельном файле
    fname = pathLog + "\\result_"+time.strftime("%Y%m%d-%H%M%S")+".xlsx"
    log_print('{0} : Сохраняем результат в отдельном файле {1}'.format(time.strftime("%d/%m/%Y, %H:%M:%S",time.localtime()), fname), log_file)
    writer = pd.ExcelWriter(fname, engine = 'openpyxl')
    BD2.to_excel(writer, sheet_name = 'План производства')
    PZ.to_excel(writer, sheet_name = 'ПЗ')
    POTR.to_excel(writer, sheet_name = 'Потребности ТМЦ')
    writer.close()
    xw.Range('last_result').options(index=False, header=0).value= fname
    
    
    
    
    #Разбиваем результат по цехам
    pathRep = pathLog + '\\results_'+time.strftime("%Y%m%d-%H%M%S")
    if not os.path.exists(pathRep):
            os.mkdir(pathRep)

    if res_ceh_flag:
        BD3 = BD2[['Технологическая карта.Подразделение', 'День', 'Номенклатура вых изд', 'Характеристика номенклатуры вых изд',
                  'Рабочий центр', 'Объем производства']].reset_index(drop = True)
        BD3.columns = ['Подразделение', 'Дата', 'Номенклатура', 'Характеристика', 'Рабочий центр', 'Кол-во']
        
        cehs = BD3['Подразделение'].drop_duplicates()
        for ceh in cehs:
            BD3[BD3['Подразделение'] == ceh].reset_index(drop = True).to_excel(pathRep + '\\' + ceh + '_'+time.strftime("%Y%m%d-%H%M%S") + '.xlsx', engine = 'openpyxl')

    
    