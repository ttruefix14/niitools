import os
import sys
from PIL import Image
Image.MAX_IMAGE_PIXELS = 1_500_000_000

import importlib
arcpy_loader = importlib.find_loader('arcpy')
arcpyLoaded = arcpy_loader is not None
if arcpyLoaded:
    import arcpy

def printMessage(text, *args):
    if arcpyLoaded:
        arcpy.AddMessage(text)
    else:
        print(text, *args)

def isImage(f: str) -> bool:
    if f.lower().endswith('.jpg'):
        return True
    
def main(dirname):
    if not os.path.exists(dirname):
        raise ValueError(f"{dirname}: директория не существует")
    for r, dirs, files in os.walk(dirname):
        for f in files:
            if not isImage(f):
                continue
            try:
                file = os.path.join(r, f)
                img = Image.open(file)
                dpi = img.info['dpi'][0]
                img.save(os.path.splitext(file)[0] + ".pdf", "PDF", resolution=dpi)
                printMessage(f"{f}: успешно конвертирован!")
            except Exception as e:
                if hasattr(e, 'message'):
                    printMessage(f'{f}: {e.message}', file=sys.stderr)
                else:
                    printMessage(f'{f}: {str(e)}', file=sys.stderr)

if __name__ == '__main__':
    if len(sys.argv) == 1:
        raise TypeError("Необходимо указать путь к директории")
    main(sys.argv[1])

