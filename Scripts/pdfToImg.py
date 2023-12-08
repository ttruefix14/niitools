import fitz
import os
import sys

def isPdf(f: str) -> bool:
    if f.endswith('.pdf'):
        return True

def main(dirname, dpi=200):
    if not os.path.exists(dirname):
        raise ValueError(f"{dirname}: директория не существует")
    for r, dirs, files in os.walk(dirname):
        for f in files:
            if not isPdf(f):
                continue
            try:
                file = os.path.join(r, f)
                print(f"{f}: Начало конвертации")
                with fitz.open(file) as doc:
                    page = doc.load_page(0)  # number of page
                    pix = page.get_pixmap(dpi=dpi)

                output = os.path.splitext(file)[0] + ".jpg"
                pix.save(output)
                print(f"{f}: успешно конвертирован!")
            except Exception as e:
                if hasattr(e, 'message'):
                    print(f'{f}: {e.message}', file=sys.stderr)
                else:
                    print(f'{f}: {str(e)}', file=sys.stderr)

if __name__ == '__main__':
    dpi = 200
    if len(sys.argv) == 1:
        raise TypeError("Необходимо указать путь к директории")
    elif len(sys.argv) > 2:
        if sys.argv[2].isdigit():
            dpi = int(sys.argv[2])
    main(sys.argv[1], dpi)