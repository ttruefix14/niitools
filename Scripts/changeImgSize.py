import os
import arcpy
from PIL import Image
Image.MAX_IMAGE_PIXELS = 933120000

def isImage(f: str) -> bool:
    if f.endswith('.jpg'):
        return True

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
                img.save(file, dpi=(dpi,dpi))
                raise NameError('НЕ ПАШЕТ')
            except Exception as e:
                if hasattr(e, 'message'):
                    arcpy.AddMessage(f'{f}: {e.message}')
                else:
                    arcpy.AddMessage(f'{f}: {str(e)}')

if __name__ == '__main__':
    execute()
