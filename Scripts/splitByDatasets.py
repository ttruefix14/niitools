import os
import pandas as pd
import sys

def isP10Layer(layer, layers_dict):
    if layer not in layers_dict:
        return False
    else:
        return True

def moveToFolder(path, dirname, layers_dict):
    filename = os.path.split(path)[1]
    layerName = os.path.splitext(filename)[0]
    rightName = layerName.split('.')[0].split('_')[0]
    if not isP10Layer(rightName, layers_dict):
        return
    dataset = layers_dict[rightName]
    newFolder = os.path.join(dirname, dataset)
    if not os.path.exists(newFolder):
        os.mkdir(newFolder)
    os.rename(path, os.path.join(newFolder, filename))

def main(dirname):
    p10 = "..\Defaults\p10.xlsx"
    with pd.ExcelFile(p10) as xls:
        layers = pd.read_excel(xls, sheet_name="Слои")

    layers_dict = {fc: ds for fc, ds in zip(layers.Layer, layers.Dataset)}


    for filename in os.listdir(dirname):
        path = os.path.join(dirname, filename)
        moveToFolder(path, dirname, layers_dict)

if __name__ == '__main__':
    main(sys.argv[1])