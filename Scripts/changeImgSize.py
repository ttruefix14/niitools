import os
import arcpy
from PIL import Image
Image.MAX_IMAGE_PIXELS = 933120000

def isImage(f: str) -> bool:
    if f.endswith('.jpg'):
        return True

def changeSize(img, new_dpi):
    dpi = img.info['dpi'][0]
    dpiDelta = new_dpi / dpi
    new_width = int(round(img.width * dpiDelta))
    new_height = int(round(img.height * dpiDelta))
    return img.resize((new_width, new_height), Image.LANCZOS)

def execute():
    dirname = arcpy.GetParameterAsText(0)
    dpi = arcpy.GetParameter(1)
    for r, dirs, files in os.walk(dirname):
        for f in files:
            if not isImage(f):
                continue
            try:
                file = os.path.join(r, f)
                img = Image.open(file)
                img = changeSize(img, dpi)
                img.save(file, dpi=(dpi,dpi))
                raise NameError('НЕ ПАШЕТ')
            except Exception as e:
                if hasattr(e, 'message'):
                    arcpy.AddMessage(f'{f}: {e.message}')
                else:
                    arcpy.AddMessage(f'{f}: {str(e)}')

if __name__ == '__main__':
    execute()
