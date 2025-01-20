import arcpy
import colorsys
import re
from osgeo import gdal
from osgeo import osr
import os

   
def hsv2rgb(h,s,v):
    return tuple(round(i * 255) for i in colorsys.hsv_to_rgb(h/360,s/100,v/100))

def get_brush(rgb, is_hs=False):
    color_code = (rgb[0]*65536) + (rgb[1]*256) + rgb[2]
    if is_hs:
        return f"Brush (107, 13421772, {color_code})" # можно поставить штриховку 5
    return f"Brush (2, {color_code})"

def get_brushes(input_layer):
    colors_dict = {}
    cym = input_layer.getDefinition('V2').renderer

    for g in cym.groups:
        for item in g.classes:
            values = item.values[0].fieldValues
            for l in item.symbol.symbol.symbolLayers:
                if not l.enable:
                    continue
                if 'CIMSolidFill' in repr(l):
                    if True:#l.color.values[:-1] != [0, 0, 0]:
                        if 'CIMHSV' in repr(l.color):
                            colors_dict[tuple(values)] = hsv2rgb(*l.color.values[:-1])
                        else:
                            colors_dict[tuple(values)] = l.color.values[:-1]
    return colors_dict, cym.fields

in_layer = arcpy.GetParameter(0)
cols = ["Ext_Zone_Code", "CLASSID"]

out_dir = arcpy.GetParameterAsText(1)
out_name = arcpy.GetParameterAsText(2)
special_field = arcpy.GetParameterAsText(3)
special_field_value = arcpy.GetParameterAsText(4)

out_mif = os.path.join(out_dir, out_name + ".mif")

colors_dict, fields = get_brushes(in_layer)

arcpy.AddMessage(str(colors_dict))

output = arcpy.conversion.FeatureClassToFeatureClass(in_layer, out_dir, out_name)

srs = osr.SpatialReference()
srs.ImportFromWkt('LOCAL_CS["Nonearth",UNIT["Meter",1.0]]')
gdal.VectorTranslate(gdal.GetDriverByName("MapInfo File").Create(out_dir, 0, 0, 0, 0, ["FORMAT=MIF"]), 
                     gdal.OpenEx(os.path.join(out_dir, out_name + ".shp")), 
                     format="MapInfo Table", 
                     layerCreationOptions=["ENCODING=cp1251", "BOUNDS=0,0,10000000,10000000"],
                     dstSRS=srs, reproject=False)

with open(out_mif, "r") as mif_file:
    mif_content = mif_file.readlines()
    
with arcpy.da.SearchCursor(in_layer, fields) as cursor:
    i = 0
    for row in cursor:
        brush = get_brush(colors_dict[tuple(str(i) for i in row)], special_field_value in str(row[fields.index(special_field)]) if special_field else False)
        while i < len(mif_content):
            if "Brush" in mif_content[i]:
                break
            i += 1

            mif_content[i] = re.sub(r"(Brush \(.*\))", brush, mif_content[i])
        i += 1 

with open(out_mif, "w") as w:
    w.writelines(mif_content)