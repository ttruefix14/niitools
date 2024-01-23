import xml.etree.ElementTree as ET
import pandas as pd
import re
import os
import sys
from dateutil.parser import parse as dtparse

def xml_parse(xml, dirname):
    types = {'extract_about_property_land': 1, 
             'extract_transfer_rights_property': 2, 
             'exract_notice_absence_request_info_12': 3,
             'extract_base_params_land': 1}
    tree = ET.parse(xml)
    root = tree.getroot()
    xml_type = types[root.tag]
    # Для выписки
    if xml_type == 1:
        # Кад номер
        cad_number = next(next(root.iter('land_record')).iter('cad_number')).text
#         print(cad_number)
        # Предыдущие кад номера
        try:
            prevs = next(root.iter('ascendant_cad_numbers'))
        except:
            prevs = None
        prev_cads = [next(i.iter('cad_number')).text for i in prevs] if prevs else []
#         print(prev_cads)
        # Категория
        try:
            cat = next(next(root.iter('category')).iter('value')).text
        except:
            cat = None
#         print(cat)
        # Права
        try:
            rights = next(root.iter('right_records'))
        except:
            rights = None
        regs = []
        reg = (None, None)
        if rights:
            regs = [(next(i.iter('value')).text, dtparse(next(i.iter('registration_date')).text).replace(tzinfo=None)) for i in rights]
            reg = min(regs, key = lambda x: x[1])
#         print(reg)
        # Аренда
        try:
            restricts = next(root.iter('restrict_records'))
        except:
            restricts = None
        rents = []
        rent = (None, None)
        if restricts:
            rents = [(next(i.iter('value')).text, dtparse(next(i.iter('registration_date')).text).replace(tzinfo=None)) for i in restricts]
            rents = list(filter(lambda x: re.search(r'ренда', x[0]), rents))
            if len(rents) > 0:
                rent = min(rents, key = lambda x: x[1])
#         print(rent)
        # Оксы
        try:
            objs = next(root.iter('included_objects'))
        except:
            objs = None
        oks = [next(i.iter('cad_number')).text for i in objs] if objs else []
#         print(oks)
        return ['reg', cad_number, reg[0], reg[1], dirname, ', '.join(prev_cads), rent[0], rent[1], cat, ', '.join(oks), root.tag]
    # Для выписки о переходе прав
    elif xml_type == 2:
        # Кад номер
        cad_number = next(next(root.iter('land_record')).iter('cad_number')).text
#         print(cad_number)        
        # Права
        try:
            rights = next(root.iter('right_records'))
        except:
            rights = None
        regs = []
        reg = ('Нет данных', 'Нет данных')
        if rights:
            regs = [(next(i.iter('value')).text, dtparse(next(i.iter('registration_date')).text).replace(tzinfo=None)) for i in rights]
            reg = min(regs, key = lambda x: x[1])
#         print(reg)        
        
        return ['ip', cad_number, reg[0], reg[1], dirname]
    
    # Для выписки без информации
    elif xml_type == 3:
        cad_number = re.search(r'\d{2}:\d{2}:\d{7}:\d+', next(root.iter('content_request')).text)[0]
#         print(cad_number)
        return ['ip', cad_number, 'Нет данных', 'Нет данных', dirname]
#         for elem in next(root.iter('land_record')):
#             print(elem.tag)
#             for child in elem:
#                 print(child.tag, end='    ')
#             print()
    else:
        print(xml, 'Неизвестный тип выписки!')


def main(dirname, xlsx):
    result = {'reg': [], 'ip': []}
    for dirpath, _, filenames in os.walk(dirname):
        # перебрать файлы
            for filename in filenames:
                if 'xml' in filename.lower():
                    res = xml_parse(os.path.join(dirpath, filename), os.path.join(dirpath))
                    result[res[0]].append(res[1:])


    writer = pd.ExcelWriter(xlsx)
    df = pd.DataFrame(result['reg'], columns=['Кадастровый номер', 'Собственность', 'Дата_собственности', 'Сельское_поселение', 'Предыдущие номера', 'Обременение', 'Дата_обременения', 'Категория земель', 'Объекты капитального строительства', 'Тип выписки'])
    df.to_excel(writer, engine='openpyxl', sheet_name='Собственность', index=False)
    df = pd.DataFrame(result['ip'], columns=['Кадастровый номер', 'Собственность', 'Дата_собственности', 'Сельское_поселение'])
    df.to_excel(writer, engine='openpyxl', sheet_name='История права', index=False)
    writer.close()

if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2])