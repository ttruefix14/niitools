import os
import sys
import re
import xml.etree.ElementTree as ET

def xml_rename(xml, dirname):
    with open(xml, 'rb') as xml_file:
        tree = ET.parse(xml_file)
        
    types = {'extract_about_property_land': 1, 
             'extract_transfer_rights_property': 2, 
             'exract_notice_absence_request_info_12': 3,
             'extract_base_params_land': 1}
    # print(xml)
    root = tree.getroot()
    xml_type = types.get(root.tag, "Неизвестен")
    # Для выписки
    if xml_type == 1:
        # Кад номер
        cad_number = next(next(root.iter('land_record')).iter('cad_number')).text
        cad_type = ''
#         print(cad_number)

    # Для выписки о переходе прав
    elif xml_type == 2:
        # Кад номер
        cad_number = next(next(root.iter('land_record')).iter('cad_number')).text
        cad_type = 'ип_'

    # Для выписки без информации
    elif xml_type == 3:
        cad_number = re.search(r'\d{2}:\d{2}:\d{7}:\d+', next(root.iter('content_request')).text)[0]
        cad_type = 'ип_нд_'
#         print(cad_number)

    else:
        print(xml, 'Неизвестный тип выписки!')
        return
    new_filename = cad_type + cad_number.replace(':', '_') + '.xml'
    new_path = os.path.join(dirname, new_filename)
    try:
        os.rename(xml, new_path)
    except FileExistsError:
        print('Уже есть:' + cad_number)   


def main(dirname):
    for dirpath, _, filenames in os.walk(dirname):
        # перебрать файлы
            for filename in filenames:
                #print("Файл:", os.path.join(dirpath, dirname, filename))
                xml_rename(os.path.join(dirpath, filename), os.path.join(dirpath))

if __name__ == '__main__':
    main(sys.argv[1])