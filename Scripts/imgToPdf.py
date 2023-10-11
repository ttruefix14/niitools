import os
import sys
from PIL import Image
Image.MAX_IMAGE_PIXELS = 933120000


def isImage(f: str) -> bool:
    if f.endswith('.jpg'):
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
                print(f"{f}: успешно конвертирован!")
            except Exception as e:
                if hasattr(e, 'message'):
                    print(f'{f}: {e.message}', file=sys.stderr)
                else:
                    print(f'{f}: {str(e)}', file=sys.stderr)

if __name__ == '__main__':
    if len(sys.argv) == 1:
        raise TypeError("Необходимо указать путь к директории")
    main(sys.argv[1])

