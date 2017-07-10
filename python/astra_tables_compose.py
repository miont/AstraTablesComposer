# -*- coding: utf-8 -*-
import fnmatch
import os
import re

def compose_astra_HTML_tables(input_dir, target_path, files_list=[]):
    """
    Объединение набора HTML-таблиц сгенерированных с помощью FastReport
    Требования: классы во всех документах должны обозначаться s0,s1,s2,...,
    нумерация должна быть без пропусков
    
    Args:
        input_dir: Директория, откуда брать таблицы
        target_path: Имя файла с полным путем с объединенной таблицей 
        files_list: Список файлов для объединения, если пустой, то используются все HTML-файлы из input_dir         
        
    Returns:
        none
    """
    
    # TODO предусмотреть что стили могут начинаться не с s
    # TODO предусмотреть что стили могут нумероваться не подряд
    # TODO реализовать многопоточную обработку HTML-документов    

    # Считывание аргументов командной строки
    # TODO реализовать с использованием argparse    
    
    # Если передан пустой список, берем все html-файлы из директории
    if files_list == None or len(files_list) == 0:
        files_list = fnmatch.filter(os.listdir(input_dir), "*.html")
    
    print('++++++++++++++++++Объединение HTML-таблиц в один документ+++++++++++++++++++++++')
    print('Рабочая директория:')
    print(input_dir)    
    print('Список файлов для обработки (количество: %d):' % (len(files_list)))
    print(files_list)
    print('Путь к выходному (объединенному) файлу:')
    print(target_path)
    
    all_styles_content = []    # Список строк с описанием стилей для всех таблиц
    all_tables_content = []    # Список строк с HTML-кодом по всем таблицам
    max_class_num_prev = -1   # Максимальный номер класса в предыдущем обработанном файле
    table_num = 0 # Номер таблицы
    print('Начало обработки')    
    for idx,file_name in enumerate(files_list):
        print('Обработка файла \'%s\'' % (file_name))
        file = open(input_dir + '/' + file_name, 'r', encoding="utf8")
        # Объединение всех строк, считанных из файла
        content = '\n'.join([line.strip() for line in file.readlines()])
        
        # Считывание стилей
        m = re.search(r'<style\s+type\s*=\s*"text/css">(.+)</style>', content, flags=re.DOTALL)
        if m == None:
            raise Exception('Error during processing re.search(r\'<style\s+type\s*=\s*"text/css">(.+)</style>\'...')
        styles_content = m.group(1)
        
        # Удаляем символы комментария        
        m = re.search(r'<!--\n*(.+)\n*-->',styles_content,flags=re.DOTALL)
        if(m != None):
           styles_content = m.group(1)
           
        # Корректируем номера классов (продолжаем нумерацию относительно предыдущего документа)
        css_styles = []        
        for m in re.finditer(r'\.s(\d+)\s*\{([^{]+)\}', styles_content, flags=re.DOTALL):
            class_num = int(m.group(1))
            properties_text = m.group(2)
            css_style = '.s%d {%s}\n'%(class_num+max_class_num_prev+1, properties_text)
            css_styles.append(css_style)
        classes_count_cur = len(css_styles)
        styles_content = ''.join(css_styles)
        all_styles_content.append(styles_content)
        
        # Считывание таблиц
        for m in re.finditer(r'(<table[^<]+>.+</table>)', content, flags=re.DOTALL):
            if m == None:
                raise Exception('Error during processing re.search(r\'<table[^<]+>.+</table>\'')
            table_content = m.group(1)
            
            # Корректировка номеров стилей
            for i in range(classes_count_cur, -1, -1):
                table_content = re.sub(r'class\s*=\s*"s%i"'%(i), 'class=s%d'%(i+max_class_num_prev+1), table_content)
            
            # Исправление номера таблицы
            table_num += 1
            table_content = re.sub(r'Т\d+-\d+', r'T%d-1'%(table_num), table_content) 
            
            all_tables_content.append(table_content)
        
        file.close()
        max_class_num_prev += classes_count_cur
        
    print('Завершение обработки')
    print('Обработано %d документов, %d таблиц' % (idx+1, table_num))
    print('Генерация выходного файла')
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
        output_lines.append(table_content+'\n')
    
    output_lines.append('</body>\n')
    output_lines.append('</html>\n')

    out_file = open(target_path, 'w', encoding="utf8")
    out_file.writelines(output_lines)
    out_file.close()        
    print('Выходной файл создан') 
    return

if __name__ == '__main__':
    
#    compose_astra_HTML_tables('../test','../test/CombinedTable.html',['макс напряж в отводах.HTML', 'макс напряж в тройниках.HTML', 'макс напряж в прямых.HTML', 'нагрузки на патрубки арматуры.HTML', 'нагрузки на оборуд и конструкции.HTML'])
    compose_astra_HTML_tables('../test','../test/CombinedTable.html')
    