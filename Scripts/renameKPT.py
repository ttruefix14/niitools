import xml.etree.ElementTree as ET
import os
import re
from dateutil import parser

def kptCadNumber(path):
    with open(path, 'r', encoding='utf-8') as f:
        n = 0
        while True:
            text = f.readline()
            if n == 0:
              n += 1
            elif n == 1:
                n += 1
                if "<extract_cadastral_plan_territory>" not in text:
                    return None, None
            
            searchDate = re.search(r'<date_formation>(.+)</date_formation>', text)
            if searchDate:
                formDate = parser.parse(searchDate.groups()[0])
            searchCad = re.search(r'<cadastral_number>(.+)</cadastral_number>', text)
            if searchCad:
                cadNumber = searchCad.groups()[0].replace(':', '_')
                return cadNumber, formDate
            
def main(dirname):
    cadNumbers = {}
    if not os.path.exists(os.path.join(dirname, 'duplicates')):
        os.mkdir(os.path.join(dirname, 'duplicates'))

    if not os.path.exists(os.path.join(dirname, 'notKpt')):
        os.mkdir(os.path.join(dirname, 'notKpt'))

    n = 1
    for filename in os.listdir(dirname):
        n += 1
        
        path = os.path.join(dirname, filename)
        if not filename.endswith('.xml'):# and not filename.startswith('report'):
            continue
        cadNumber, formDate = kptCadNumber(path)
        if cadNumber is None:
            cadNumbers.setdefault("notKpt", []).append([filename, path])
            continue
        cadNumbers.setdefault(cadNumber, []).append([formDate, path])

    for cad, dates in cadNumbers.items():
        for i, item in enumerate(sorted(dates, reverse=True)):
            path = item[1]
            sig = path + '.sig'
            if cad == 'notKpt':
                os.rename(path, os.path.join(dirname, 'notKpt', item[0]))
                if os.path.isfile(sig):
                    os.rename(sig, os.path.join(dirname, 'notKpt', item[0] + '.sig'))
                continue
            
            if i == 0:
                os.rename(path, os.path.join(dirname, cad + '.xml'))
                if os.path.isfile(sig):
                    os.rename(sig, os.path.join(dirname, cad + '.xml.sig'))
                
            else:
                os.rename(path, os.path.join(dirname, 'duplicates', cad + '_' + str(i) + '.xml'))
                if os.path.isfile(sig):
                    os.rename(sig, os.path.join(dirname, 'duplicates', cad + '_' + str(i) + '.xml.sig'))


if __name__ == '__main__':
    dirname = arcpy.GetParameterAsText(0)
    main(dirname)