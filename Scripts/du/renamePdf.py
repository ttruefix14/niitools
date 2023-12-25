import fitz
import os
import sys
import re

def main(inDir):
    text = ''
    for dirname, _, filenames in os.walk(inDir):
        for filename in filenames:
            if '.pdf' not in filename.lower():
                continue
            dop = ''
            with fitz.open(f'{dirname}\\{filename}') as doc:
                try:
                    text = doc[0].get_text()
                except:
                    print(filename)
                t_list = text.split('\n')
                if 'о переходе прав на объект' in text:
                    dop = 'ип_'
                if 'Уведомление об отсутствии в Едином' not in text:
                    for i, row in enumerate(t_list):
                        if 'Кадастровый номер' in row:
                            cad = t_list[i+1]
                            break
                if 'Уведомление об отсутствии в Едином' in text:
                    dop = 'ип_нд_'
                    cad = re.search(r'[0-9]{2}:[0-9]{2}:[0-9]{7}:[0-9]+', text).group(0)
                cad = cad.replace(':', '_')
                print(cad)
            try:
                os.rename(f'{dirname}\\{filename}', f'{dirname}\\{dop + cad}.pdf')
            except FileExistsError:
                print(f'Уже есть: {cad}')

if __name__ == '__main__':
    main(sys.argv[1])