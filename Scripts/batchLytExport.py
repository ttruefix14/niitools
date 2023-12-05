import arcpy
import os

# def getLyts(value, aprx):
#     lyt_name = value.value
#     lyt = aprx.listLayouts(lyt_name)[0]
#     return lyt

def exportToJPEG(lyts, dpi, quality, outputFolder):
    for lyt in lyts:
        try:
            lyt.exportToJPEG(os.path.join(outputFolder, lyt.name), dpi, jpeg_quality=quality)
            arcpy.AddMessage(f'{lyt.name}: Успешно экспортирован')
        except Exception as e:
            arcpy.AddMessage(f'{lyt.name}: {e}')

def exportToPDF(lyts, dpi, quality, output_as_image, outputFolder):
    for lyt in lyts:
        try:
            lyt.exportToPDF(os.path.join(outputFolder, lyt.name), dpi, image_compression='LZW', layers_attributes='LAYERS_ONLY', georef_info=False, jpeg_compression_quality=quality, output_as_image=output_as_image)
            arcpy.AddMessage(f'{lyt.name}: Успешно экспортирован')
        except Exception as e:
            arcpy.AddMessage(f'{lyt.name}: {e}')



def execute():
    lyt_names = arcpy.GetParameterAsText(0)
    outputFormat = arcpy.GetParameterAsText(1)
    outputAsImage = arcpy.GetParameter(2)
    dpi = arcpy.GetParameter(3)
    quality = arcpy.GetParameter(4)
    outputFolder = arcpy.GetParameterAsText(5)

    aprx = arcpy.mp.ArcGISProject("CURRENT")
    lyts = [lyt for lyt in aprx.listLayouts() if lyt.name in lyt_names]

    if outputFormat == 'JPEG':
        exportToJPEG(lyts, dpi, quality, outputFolder)
    elif outputFormat == 'PDF':
        exportToPDF(lyts, dpi, quality, outputAsImage, outputFolder)
    else:
        arcpy.AddMessage(f'Формат {outputFormat} не поддерживается')

if __name__ == '__main__':
    execute()