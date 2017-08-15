# -*- coding: utf-8 -*-
#### Склеивание сводных таблиц в HTML, сгенерированных FastReport в отдельных документах,
#### в общий документ для открытия в Word
#### Многопоточный вариант
#### Авторы: Konstantin Ostrovsky, Ivan Kush
#### НИЦ СтаДиО 2017

import fnmatch
import os
import os.path
import re
import argparse
import concurrent.futures as confu
import multiprocessing
import time
import logging

# TODO Сделать однократную печать шапки (имя проекта, название объекта и пр.) в начале объединенного документа
# TODO Поправить документацию функций
# TODO удалять директорию

## Константы
LOG_FILE_NAME = 'log/html_merger.log'    # Путь к файлу для печати логов
LOGGER_NAME = 'logger'                   # Общее имя логгера
BREAKING_LINE = '='*50                   # Строка-разделитель для форматирования
DEFAULT_DEBUG_MODE = True                # Режим дебага по умолчанию (для вывода исключений, возникших при парсинге аргументов)

# Глобальные параметры
log = None           # Объект-логгер

def read_styles(file_name, input_dir):
    """
        Извлечение названий классов css-стилей из документа
        
        Args:
            file_name: имя файла
            input_dir: рабочая директория
            
        Returns:
            css_styles (tuple): список стилей, содержит
                class_label: имя класса
                properties_text: свойства стиля
    """
    
    log.info('чтение стилей из файла \'%s\'' % (file_name))
    file = open(input_dir + '/' + file_name, 'r', encoding="utf8")
    # Объединение всех строк, считанных из файла
    content = '\n'.join([line.strip() for line in file.readlines()])   
    
    # Считывание стилей
    m = re.search(r'<style\s+type\s*=\s*"text/css">(.+)</style>', content, flags=re.DOTALL)
    if m == None or len(m.groups()) < 1:
        raise TablesComposerException('Ошибка при считывании стилей в analyze_styles_in_file для файла \'%s\'' % (file_name))
    styles_content = m.group(1)
    
    css_styles = []
    for m in re.finditer(r'\.(-?[_a-zA-Z]+[_a-zA-Z0-9-]*)\s*\{([^{]+)\}', styles_content, flags=re.DOTALL):
        class_label = m.group(1)
        properties_text = m.group(2)
        if class_label == 'page_break':
            continue
        
        css_styles.append((class_label, properties_text))        
        
    file.close()
    log.info('файл \'%s\' обработан' % (file_name))
    return css_styles 

def parse_file(file_name, input_dir, css_styles, class_num_start, document_num=1):
    """
    Обработка одного HTML-документа. Выделение стилей и таблиц
    
    Args:
        file_name: имя файла
        input_dir: рабочая директория
        class_num_start: номер класса стилей, с которого начинать нумерацию
        document_num: номер документа (используется для нумерации таблиц)        
        
    Returns:
        dict {styles_content, tables_content, tables_count}
        styles_content: строка с описанием стилей CSS, подготовленная для вставки в общий документ
        tables_content: строка с описанием HTML-таблиц из текущего документа с исправленными номерами стилей для вставки в общий документ
        tables_count: количество таблиц в данном документе
    """
    
    log.info('Начало обработки файла \'%s\'' % (file_name))
    file = open(input_dir + '/' + file_name, 'r', encoding="utf8")
    # Объединение всех строк, считанных из файла
    content = '\n'.join([line.strip() for line in file.readlines()])
    
    # Считывание стилей
    m = re.search(r'<style\s+type\s*=\s*"text/css">(.+)</style>', content, flags=re.DOTALL)
    if m == None:
        raise TablesComposerException('Ошибка при обработке регулярного выражения \'<style\s+type\s*=\s*"text/css">(.+)</style>\'')
    styles_content = m.group(1)
    
    # Удаляем символы комментария        
    m = re.search(r'<!--\n*(.+)\n*-->',styles_content,flags=re.DOTALL)
    if(m != None):
       styles_content = m.group(1)
       
    # Извлекаем названия классов
    class_labels = [lab for lab, _ in css_styles]
       
    # Корректируем имена классов (продолжаем нумерацию относительно предыдущего документа)
    # Формат имен классов: s1, s2, ...
    css_styles_mod = []     # Измененный список стилей
    for i, (_, css_props) in enumerate(css_styles):
        css_styles_mod.append('.s%d {%s}\n'%(class_num_start + i, css_props))
    styles_content = ''.join(css_styles_mod)            
    
    # Считывание таблиц
    table_num = 0  # Номер таблицы
    tables_content = ''
    for m in re.finditer(r'(<table[^<]+>.+?</table>)', content, flags=re.DOTALL):  # Not greedy match '+?'
        if m == None or len(m.groups()) < 1:
            raise TablesComposerException('Ошибка при считывании таблиц в функции parse_file для файла \'%s\'' % (file_name))
        tables_content_cur = m.group(1)
        
        # Корректировка имен классов стилей
        classes_count = len(css_styles)
        for i in range(classes_count-1, -1, -1):
            tables_content_cur = re.sub(r'class\s*=\s*"%s"'%(class_labels[i]), 'class="s%d"'%(i+class_num_start), tables_content_cur)
        
        # Исправление номера таблицы  
        table_num += 1
        tables_content_cur = re.sub(r'Т\d+-\d+', 'T%d-%d'%(document_num, table_num), tables_content_cur)
        
        tables_content += tables_content_cur
    tables_count = table_num  # Количество таблиц
    
    file.close()
    log.info('Завершение обработки файла \'%s\'. Стилей: %d, таблиц: %d' % (file_name, len(class_labels), tables_count))
    
    return {'styles_content' : styles_content, 'tables_content': tables_content, 'tables_count': tables_count}    
    
def compose_astra_html_tables(input_dir, target_path, files_list=[], multithread=True):
    """
    Объединение набора HTML-таблиц сгенерированных с помощью FastReport
    Требования: классы во всех документах должны обозначаться s0,s1,s2,...,
    нумерация должна быть без пропусков
    
    Args:
        input_dir: Директория, откуда брать таблицы
        target_path: Имя файла с полным путем с объединенной таблицей 
        files_list: Список файлов для объединения, если пустой, то используются все HTML-файлы из input_dir         
        multithread (boolean): Флаг многопоточности        
    """
        
    # Если передан пустой список, берем все html-файлы из директории
    if files_list == None or len(files_list) == 0:
        files_list = []
        files_list += fnmatch.filter(os.listdir(input_dir), "*.html")
        files_list += fnmatch.filter(os.listdir(input_dir), "*.htm")
        files_list += fnmatch.filter(os.listdir(input_dir), "*.txt")
        # Убираем вложенные директории из списка
        files_list = [file for file in files_list if os.path.isfile(input_dir + '/' + file)]
        # Сортируем по номерам (предполагается, что названия файлов: 1.html, 2.html, 3.html, ...)
        files_list = sorted(files_list, key = lambda s : int(s[:s.find('.')]))
        
    files_count = len(files_list)    # Количество файлов
    
    cores_used = 1 # Количество используемых ядер
    if multithread:
        cores_used = multiprocessing.cpu_count()           
    
    start_time = time.time()
    log.info('++++++++++++++++++Объединение HTML-таблиц в один документ+++++++++++++++++++++++')
    log.info('Рабочая директория:')
    log.info(input_dir)    
    log.info('Список файлов для обработки (количество: %d):' % (files_count))
    log.info(files_list)
    log.info('Путь к выходному (объединенному) файлу:')
    log.info(target_path)
    
    log.info('%s обработка. Доступно ядер: %d. Используется: %d ' % ('Многопоточная' if multithread else 'Однопоточная', multiprocessing.cpu_count(), cores_used))
    
    # Определение количества классов css-стилей в каждом из документов, сохранения названий классов в список
    log.info('-------Подсчет числа классов во всех документах-------')
    css_styles = [[]]*files_count   # Стили
    
    with confu.ThreadPoolExecutor(max_workers = cores_used) as executor:
        futures_to_idx = {executor.submit(read_styles, file_name, input_dir) : idx for idx,file_name in enumerate(files_list)}
    for future in confu.as_completed(futures_to_idx):
        idx = futures_to_idx[future]        
        try:
            css_styles[idx] = future.result()
        except Exception as exc:
            raise TablesComposerException(exc)
    log.info(BREAKING_LINE)
#    class_labels = [[lab for lab,_ in entry] for entry in css_styles]    
    
    all_styles_content = [[]]*files_count    # Список строк с описанием стилей для всех документов
    all_tables_content = [[]]*files_count     # Список строк с HTML-кодом таблиц по всем документам
#    max_class_num_prev = 0   # Максимальный номер класса в предыдущем обработанном файле
#    tables_count = 0    
    log.info('-------Начало обработки файлов-------')
    with confu.ThreadPoolExecutor(max_workers = cores_used) as executor:
        class_num_start = 1
        futures_to_idx = {}
        for idx,file_name in enumerate(files_list):
            # TODO Преобразовать в вызов функции для устранения повтора кода
            futures_to_idx[executor.submit(parse_file, file_name, input_dir, css_styles[idx], class_num_start, idx+1)] = idx
            class_num_start += len(css_styles[idx])
    
    tables_count = 0  # Количество таблиц во всех файлах
    for future in confu.as_completed(futures_to_idx):
        idx = futures_to_idx[future]        
        try:
            data = future.result()
            all_styles_content[idx] = data['styles_content']
            all_tables_content[idx] = data['tables_content']
            tables_count += data['tables_count']
        except Exception as exc:
            raise TablesComposerException(exc)
        
    log.info('Завершение обработки всех файлов')
    log.info('Обработано: документов %d, таблиц %d' % (len(files_list), tables_count))
    log.info(BREAKING_LINE)
    
    log.info('Генерация выходного файла')
    output_lines = []
    output_lines.append('<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">\n')
    output_lines.append('<html>\n')
    output_lines.append('<head>\n')
    output_lines.append('<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">\n')
    output_lines.append('<meta name="Generator" content="FastReport 5.0 http://www.fast-report.com">\n')    
    output_lines.append('<title></title>\n')
    output_lines.append('<style type="text/css">\n')
    for style_content in all_styles_content:
        output_lines.append(style_content+'\n')
    output_lines.append('</style>\n')
    output_lines.append('</head>\n')
    
    output_lines.append('<body text="#000000" bgcolor="#FFFFFF">\n')
    
    for table_content in all_tables_content:
        output_lines.append(table_content+'<br>\n')
    
    output_lines.append('</body>\n')
    output_lines.append('</html>\n')

    out_file = open(target_path, 'w', encoding="utf8")
    out_file.writelines(output_lines)
    out_file.close()        
    log.info('Выходной файл создан')
    finish_time = time.time()
    log.info('Затраченное время: %5.2f сек' % (finish_time - start_time))
    return

def parse_args():
    """
    Считывание аргументов командной строки
    
    Args:    
        none
    
    Returns:
        dict:
        {
        'dir': директория с HTML-файлами
        'output': путь к выходному файлу
        'files': список HTML-файлов
        'multithread': флаг многопоточной обработки
        }
    """
    
    try:
        parser = argparse.ArgumentParser(description='Склеивание сводных таблиц в HTML')
        parser.add_argument('-d', '--dir', type=str, help='директория с HTML-файлами')
        parser.add_argument('-f', '--file', type=str, nargs='+', help='список HTML-файлов')
        parser.add_argument('-o', '--output', type=str, help='путь к выходному файлу')
        parser.add_mutually_exclusive_group(required=False)
        parser.add_argument('--mthread', dest='mthread', action='store_true', help='использовать многопоточность')
        parser.add_argument('--no-mthread', dest='mthread', action='store_false', help='однопоточный режим')
#        parser.set_defaults(mthread=True)
        parser.add_argument('--debug', dest='debug', action='store_true', required=False, help='режим отладки')
        parser.set_defaults(debug=False)
        
        args_parser_result = parser.parse_args()
    
        # Словарь с результатами обработки
        args = {}
        if(args_parser_result.dir and args_parser_result.output):
            args['dir'] = args_parser_result.dir
            args['output'] = args_parser_result.output
        else:
            raise ArgsParserException('Не переданы необходимые аргументы: -d, -o')
        
        # Парсинг списка файлов
        if (args_parser_result.file):
            files_string = ' '.join(args_parser_result.file)
            if(files_string.find("\"") == -1):
                args['files'] = files_string.split(' ')
            else:
                files_list = []
                for m in re.finditer(r'"([\w\s]+\.\w+)"', files_string):
                    if(m == None):
                        raise ArgsParserException('Ошибка при обработке re.finditer(r'("[\w\s]+")+', files_string)')
                    file_name = m.group(1)
                    files_list.append(file_name)
                args['files'] = files_list
        else:
            args['files'] = []
        
        # Многопоточность
        args['multithread'] = args_parser_result.mthread    
    
        # Режим вывода ошибок
        args['debug'] = args_parser_result.debug

    except Exception as e:
        print(e)
        raise ArgsParserException(e)  
        
    return args
    
def configure_logging(debug=False):
    """
    Инициализация логирования
    
    Args:
        debug: флаг отладки
    Returns:
        logger: объект для вывода в лог
    
    """
    try:
        logger = logging.getLogger(LOGGER_NAME)

        # Вывод в консоль        
        console_log_handler = logging.StreamHandler()
        logger.addHandler(console_log_handler)
        
        # Вывод в файл
        if not os.path.exists('log'):
            os.makedirs('log')
        file_log_handler = logging.FileHandler(LOG_FILE_NAME, mode='w')
        logger.addHandler(file_log_handler)
        
        formatter = logging.Formatter('%(message)s')
        file_log_handler.setFormatter(formatter)
        console_log_handler.setFormatter(formatter)
        
        logger.setLevel('DEBUG' if debug else 'INFO')
        
        logger.propagate = False
        
    except Exception as e:
        raise LoggerException(e)
    return logger
    
def run_from_command_line():
    """
    Запуск из командной строки        
    """
    global log
    try:
#        sys.stdout = open(LOG_FILE_NAME, "w")
        args = parse_args()
        log = configure_logging(args['debug'])
        compose_astra_html_tables(args['dir'], args['output'], args['files'], args['multithread'])
    except ArgsParserException as e:
        print('[ОШИБКА] Ошибка при парсинге аргументов командной строки')
        if DEFAULT_DEBUG_MODE:
            print(e)
            logging.exception(e)
    except LoggerException as e:
        print('[ОШИБКА] Ошибка при настройке логирования')
        if args['debug']:
            logging.exception(e)
    except TablesComposerException as e:
        log.error('[ОШИБКА] Ошибка при обработке таблиц')
        log.debug(e)
    
    except Exception as e:
        log.error('[ОШИБКА] Ошибка обшего содержания')
        log.debug(e)
    finally:
        # Очистка обработчиков у логгера
        if log != None:
            handlers = log.handlers[:]
            for handler in handlers:
                handler.close()
                log.removeHandler(handler)
#        sys.stdout = sys.__stdout__
        
# Исключения    
class TablesComposerException(Exception):
    pass
class ArgsParserException(Exception):
    pass
class LoggerException(Exception):
    pass

if __name__ == '__main__':
    run_from_command_line()
    
    # Пример аргументов командной строки 
    # -d ../test/input -f "макс напряж в отводах.HTML" "макс напряж в тройниках.HTML" "макс напряж в прямых.HTML" "нагрузки на патрубки арматуры.HTML" -o ../test/CombinedTable.html
    