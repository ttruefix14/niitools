from xml.etree import ElementTree as et
import os
import pandas as pd

def parseXml(filename: str) -> list:
    tree = et.parse(filename)
    root = tree.getroot()
    result = []
    for elem in root.iter("land_record"):
        if isEzp(elem):
            result.append(parseEzp(elem) + [filename])
    return result

def isEzp(elem: et.Element) -> bool:
    try:
        subtype = next(elem.iter("subtype"))
        if next(subtype.iter("value")).text == "Единое землепользование":
            return True
    except:
        return False

def parseEzp(elem: et.Element) -> list:
    cadNumber = next(elem.iter("cad_number")).text
    
    cat = next(next(elem.iter("category")).iter("value")).text
    bydoc = next(elem.iter("by_document")).text
    area = float(next(next(elem.iter("area")).iter("value")).text)
    address = next(elem.iter("readable_address")).text
    cost= float(next(next(elem.iter("cost")).iter("value")).text)

    return [cadNumber, address, cat, bydoc, area, cost]

def execute():
    dirname = arcpy.GetParameterAsText(0)
    files = [os.path.abspath(os.path.join(dirname, f)) for f in os.listdir(dirname) if '.xml' in f and '.sig' not in f]
    
    result = []
    for f in files:
        print(f)
        result += parseXml(f)
    df = pd.DataFrame(result, columns=['par', 'adr', 'cat', 'bydoc', 'area', 'cost', 'filename'])
    df.to_excel(arcpy.GetParameterAsText(1), index=False)

if __name__ == "__main__":
    execute()