''' На карте должны быть все слои, которые экспортируются в GML в любой прямоугольной системе координат (на нужную поменяется автоматически)
Названия слоев берутся из базы - если делаете из шейпов - нужно называть их также как они называются по десятому приказу
'''
import arcpy
import shapely
import geopandas as gpd
# from osgeo import ogr
# from osgeo import osr
from osgeo import gdal

import os
import shutil

import difflib
import re
import regex

import pandas as pd
from numpy import int64, float64
import json
from io import BytesIO
from math import ceil


class Params:
    """Входные параметры инструмента"""
    def __init__(self, params):
        with open(r'..\Defaults\xsds.json', encoding='utf-8') as f:
            xsds = json.load(f)

        self.xsd = xsds[params[0].valueAsText] 

        self.p10 = params[1].valueAsText

        self.outputSRS = params[2].value

        rename_cols = [i.split('=') for i in params[3].values] if params[3].values else None
        self.rename_cols = {i[0]: i[1] for i in rename_cols} if rename_cols else None
        
        self.mask_layer = params[4].value # Введите название
        self.mask_exceptions = params[5].values if params[5].values else []

        self.OKTMO = params[6].valueAsText # Введите ОКТМО поселения, если где-то октмо будет не заполнено - оно заполнится этим автоматически
        self.empty_value = params[7].value

        self.ignore_cid = params[8].values

        self.omz_definition = [i.split('<>') for i in params[9].values] if params[9].value else None
        
        self.output_dirname = params[10].valueAsText
        
        self.geoTransfm = params[11].valueAsText if params[11].value else None

        self.customGeoTransfm = params[12].valueAsText

        self.splitSize = params[13].value

        self.gml_version = params[14].valueAsText

       

class P10:
    def __init__(self, p10):
        xls = pd.ExcelFile(p10)
        cid_table = pd.read_excel(xls, 'ClassID')
        cid_table = cid_table.where(pd.notnull(cid_table), None)
        atr_table = pd.read_excel(xls, 'Общий')
        atr_table = atr_table.where(pd.notnull(atr_table), None)
        dom_table = pd.read_excel(xls, 'Справочники')
        dom_table = dom_table.where(pd.notnull(dom_table), None)
        xls.close()

        self.p10 = {i:{} for i in atr_table['Layer'].to_list()}

        for i in zip(atr_table['Layer'].to_list(), 
                     atr_table['Name'].to_list(), 
                     atr_table['Check'].to_list(), 
                     atr_table['Type'].to_list(),
                     atr_table['Domain'].to_list(),
                     atr_table['Condition'].to_list(),
                     atr_table['Default'].to_list()):
            self.p10[i[0]][i[1]] = [i[2], 
                                    i[3], 
                                    [d for d in dom_table['Code'].loc[dom_table['Domain'] == i[4]].to_list()] if i[4] else [c for c in cid_table['CLASSID'].loc[cid_table['Layer'] == i[0]].to_list()] if i[1] == 'CLASSID' else None, 
                                    i[5], 
                                    i[4], 
                                    i[6]]
    
    def __repr__(self):
        return self.p10.__repr__()
    
    def __getitem__(self, item):
        return self.p10[item]
    
    def is_p10_layer(self, r_name):
        if r_name not in self.p10.keys():
            return False
        else:
            return True
    
    def check_columns(self, df, r_name, b_name, rename_cols):
        if rename_cols:
            df = df.rename(columns=rename_cols)
        df.columns = df.columns.str.replace('_$', '', regex=True).str.replace('.*\.', '', regex=True).str.upper()
        required_columns_list = [col for col in self.p10[r_name].keys()]
        required_columns = set(required_columns_list)
        rename_columns = dict()
        # for col in required_columns:
        #     if col not in df.columns:
        #         for col2 in df.columns:
        #             if col2 in col:
        #                 rename_columns[col2] = col
        missing_columns = set(required_columns) - set(df.columns.to_list())
        last_columns = set(df.columns.to_list()) - set(required_columns)
        for col in missing_columns:
            closest = difflib.get_close_matches(col, last_columns, 1)
            if len(closest) == 0:
                continue
            closest = closest[0]
            rename_columns[closest] = col
            last_columns.remove(closest)

        df = df.rename(columns=rename_columns)
        df = df.rename(columns={'GEOMETRY': 'geometry'})
        missing_columns = set(required_columns) - set(df.columns.to_list())
        # временное решение для прописывания дефолтов
        for column in missing_columns:
            df[column] = None
        check_columns = required_columns# - missing_columns
        check_columns = sorted(list(check_columns), key=lambda x: required_columns_list.index(x))
        return df[check_columns + ['geometry']]
    
def tab_to_gdf(in_table, input_fields=None, where_clause=None, spatial_reference=None, datum_transformation=None):
        """Function will convert an arcgis table into a pandas dataframe with an object ID index, and the selected
        input fields using an arcpy.da.SearchCursor."""
        OIDFieldName = arcpy.Describe(in_table).OIDFieldName
        try:
            shapeFieldName = arcpy.Describe(in_table).shapeFieldName
        except:
            shapeFieldName = None
        if input_fields:
            input_fields = [field.upper() for field in input_fields]
            if shapeFieldName:
                final_fields = [OIDFieldName] + ['SHAPE@WKT'] + input_fields
            else:
                final_fields = [OIDFieldName] + input_fields
        else:
            final_fields = [field.name.upper() if field.name != shapeFieldName else 'SHAPE@WKT' for field in arcpy.Describe(in_table).fields]
        data = [row for row in arcpy.da.SearchCursor(in_table, final_fields, where_clause=where_clause, spatial_reference=spatial_reference, datum_transformation=datum_transformation)]
        df = pd.DataFrame(data, columns=final_fields)
        gs = gpd.GeoSeries.from_wkt(df['SHAPE@WKT'])
        del df['SHAPE@WKT']
        srs = spatial_reference or arcpy.Describe(in_table).spatialReference
        fc_dataframe = gpd.GeoDataFrame(df, geometry=gs, crs=srs.factoryCode) #srs.exportToString()
        # fc_dataframe = fc_dataframe.set_index(OIDFieldName, drop=True)
        return fc_dataframe, srs.factoryCode


def create_custom_transformation(inputSRS, outputSRS, customGeoTransfm, geoTransfm):
    if not outputSRS:
        return None
    if inputSRS == outputSRS or inputSRS.exportToString() == outputSRS.exportToString() or inputSRS.factoryCode == outputSRS.factoryCode:
        return None
    if inputSRS.factoryCode == 4326:
        return None
    dirname = os.path.join(os.getenv('APPDATA'), 'Esri', 'ArcGISPro', 'ArcToolbox', 'CustomTransformations')
    transfList = arcpy.ListTransformations(inputSRS, outputSRS)
    if len(transfList) > 0:
        if geoTransfm in transfList:
            return geoTransfm
        return transfList[0]
    
    geotransfName = 'MSK_to_WGS{}'
    n = 1
    while os.path.isfile(os.path.join(dirname, geotransfName.format(n) + '.gtf')):
        n += 1
    arcpy.AddMessage(geotransfName.format(n))
    arcpy.management.CreateCustomGeoTransformation(geotransfName.format(n), inputSRS, outputSRS, customGeoTransfm)
    return geotransfName.format(n)

def get_dimension(shape_type, d_dict={'point': 1, 'polyline': 2, 'polygon': 4}):
    return d_dict[shape_type]


def makeValidForP10(row_object, name, OKTMO, p10):
    name = name.split('_')[0]

    for col, value in row_object.items():
        if col not in p10.get(name):
            continue

        if col == 'OKTMO':
            if type(value) == str:
                if len(value) < 6:
                    value = OKTMO
                else:
                    pass
            else:
                value = OKTMO
        # Устранить
        if value == 0:
            value = None
        if value is None or value != value:
            value = p10.get(name).get(col)[5]
        if type(value) == str:
            if value.isspace() or value == "":
                value = p10.get(name).get(col)[5]

        if p10.get(name).get(col)[1] == 'Целое':
            try:
                value = int(value)
            except:
                if p10.get(name).get(col)[4] is not None:
                    value = 0

        if p10.get(name).get(col)[1] == 'Вещественное':
            try:
                value = round(value, 2)
            except:
                value = 0

        if isinstance(value, (float, float64)):
            value = round(value, 2)

        row_object[col] = value
    return row_object

def to_multigeo(geom):
    gtype = geom.geom_type
    if 'Multi' in gtype:
        return geom
    if gtype == 'Polygon':
        return shapely.MultiPolygon([geom])
    if gtype == 'LineString':
        return shapely.MultiLineString([geom])
    if gtype == 'Point':
        return geom
    if gtype == 'GeometryCollection':
        return geom
        
def removeLowerDimension(geom, gtype):
    new_gtype = geom.geom_type
    if new_gtype == gtype:
        return to_multigeo(geom)
    elif 'Multi' + gtype == new_gtype:
        return geom
    if new_gtype == 'GeometryCollection':
        geoms = {geom.geom_type: geom for geom in geom.geoms}
        if gtype in geoms:
            return to_multigeo(geoms[gtype])
        elif 'Multi' + gtype in geoms:
            return geoms['Multi' + gtype]
    return None

def escape_xml_illegal_chars(unicodeString, replaceWith=r''):
	return re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1F\uD800-\uDFFF\uFFFE\uFFFF]', replaceWith, unicodeString)

# def create_gml(dirname, filename, xsd, gml_version):
#     path = os.path.join(dirname, filename)
#     if os.path.isfile(path):
#         os.remove(path)
    
#     driver = ogr.GetDriverByName("GML")
#     outDataSource = driver.CreateDataSource(path, options=['PREFIX=fgistp', f'FORMAT={gml_version}', r'TARGET_NAMESPACE=http://fgistp', fr'XSISCHEMAURI=http://fgistp {xsd}', 'WRITE_FEATURE_BOUNDED_BY=NO'])
#     return outDataSource

def create_gml(dirname, filename, xsd, gml_version):
    path = os.path.join(dirname, filename)
    if os.path.isfile(path):
        os.remove(path)

    driver = gdal.GetDriverByName("GML")
    outDataSource = driver.Create(path, 0,0,0,0, ['PREFIX=fgistp', f'FORMAT={gml_version}', r'TARGET_NAMESPACE=http://fgistp', fr'XSISCHEMAURI=http://fgistp fgistp-10-izm-698.xsd', 'WRITE_FEATURE_BOUNDED_BY=NO']) #{xsd}
    return outDataSource

def fc_to_gml(outSource, layerName, gdf, epsg, mask, oktmo, p10):
    if len(gdf) == 0:
        return
    gdf = gdf.loc[~gdf.geometry.isnull()]
    gdf.geometry = gdf.apply(lambda row: shapely.make_valid(row.geometry), axis=1)
    if len(gdf) == 0:
        return
    
    if mask:
        gdf.geometry = gdf.intersection(mask)
    gdf = gdf.loc[~gdf.geometry.is_empty]
    if len(gdf) == 0:
        return
    
    gtype = layerName.split('_')[1]

    gdf.geometry = gdf.apply(lambda row: removeLowerDimension(row.geometry, gtype), axis=1)
    gdf = gdf.loc[~gdf.geometry.isnull()]

    if len(gdf) == 0:
        return

    gdf = gdf.apply(lambda row: makeValidForP10(row, layerName, oktmo, p10), axis=1)


    for col in gdf.columns:
        if gdf[col].dtype == 'datetime64[ns]':
            gdf[col] = gdf.apply(lambda row: row[col].strftime('%Y/%m/%d'), axis=1)
    
    valid = [layerName, [[gid, shapely.is_valid(geom)] for gid, geom in zip(gdf.GLOBALID, gdf.geometry)]]
    if not all([i[1] for i in valid[1]]):
        not_valid_geoms.append(valid)

    temp = BytesIO()
    gdf.to_file(temp, driver='GeoJSON')
    # gdf.to_file(fr"C:\Users\ya.shatalov\Desktop\Data\1_Проекты\Апшеронское\Apsheronskoe\ФГИСТП_Апшеронское\{layerName}.geojson", driver='GeoJSON')
    temp.seek(0)
    
    # driver = gdal.GetDriverByName("GeoJSON")
    # inSource = driver.Open(temp.read())
    inSource = gdal.OpenEx(temp.read())
    # inSource = driver.Open(fr"C:\Users\ya.shatalov\Desktop\Data\1_Проекты\Апшеронское\Apsheronskoe\ФГИСТП_Апшеронское\{layerName}.geojson")
    temp.close()

    gdal.VectorTranslate(outSource, inSource, accessMode="append", geometryType="PROMOTE_TO_MULTI", layerName=layerName)
    
    #old way
    # inLayer = inSource.GetLayer()

    # srs = osr.SpatialReference()
    # srs.ImportFromEPSG(epsg)
    # outLayer = outSource.GetLayerByName(layerName)
    # if not outLayer:
    #     outLayer = outSource.CreateLayer(layerName, srs)

    # layerDefn = inLayer.GetLayerDefn()
    # outLayerDefn = outLayer.GetLayerDefn()
    # for n in range(0, layerDefn.GetFieldCount()):
    #     defn = layerDefn.GetFieldDefn(n)
    #     if outLayerDefn.GetFieldIndex(defn.GetName()) == -1:
    #         outLayer.CreateField(defn)
            
    
    # for feature in inLayer:
    #     geom = feature.GetGeometryRef()
        
    #     # geom = feature.GetGeometryRef()#.RemoveLowerDimensionSubGeoms()
    #     geom.AssignSpatialReference(srs)
    #     feature.SetGeometry(geom)

    #     print(geom.GetSpatialReference())
    #     outLayer.CreateFeature(feature)

def splitXml(path, splitSize, fileSize):
    splitSize = int((fileSize / ceil(fileSize / splitSize)) + 1)
    filename, extension = os.path.splitext(path)
    with open(path, 'r', encoding='utf-8') as f:
        xmlstr = f.read()

    blank = [xmlstr[:regex.search(r"</.*boundedBy>\n*\t*", xmlstr).end()], regex.search(r"(</.*FeatureCollection>)", xmlstr).groups()[0]]


    blank_size = len(blank[0].encode('utf-8')) + len(blank[1].encode('utf-8'))

    current = xmlstr[regex.search(r"</.*boundedBy>\n*\t*", xmlstr).end():regex.search(r"(</.*FeatureCollection>)", xmlstr).start()].encode('utf-8')

    n = 1

    byteSplitSize = (1024 * 1024) * splitSize
    while byteSplitSize < (len(current) + blank_size):
        if len(current) <= 2:
            break
        if len(current) > (byteSplitSize - blank_size):
            chunk = current[:byteSplitSize - blank_size].decode('utf-8')
        else:
            chunk = current.decode('utf-8')
        current_bord = regex.search(r"</.*featureMember>\n*\t*", chunk, flags=regex.REVERSE)

        result = blank[0] + chunk[:current_bord.end()] + blank[1]
        result = re.sub(r"(\t+)(</.*FeatureCollection>)", r"\2", result)

        with open(os.path.join(f'{filename}_{n}{extension}'), 'wb') as f:
            f.write(result.encode('utf-8'))
        current = current.decode('utf-8')[current_bord.end():].encode('utf-8')

        n += 1
    else:
        if len(current) > 2:
            if len(current) > (byteSplitSize - blank_size):
                chunk = current[:byteSplitSize - blank_size].decode('utf-8')
            else:
                chunk = current.decode('utf-8')
            current_bord = regex.search(r"</.*featureMember>\n*\t*", chunk, flags=regex.REVERSE)
            result = blank[0] + chunk[:current_bord.end()] + blank[1]
            with open(os.path.join(f'{filename}_{n}{extension}'), 'wb') as f:
                f.write(result.encode('utf-8'))


def execute():
    params = Params(arcpy.GetParameterInfo())
    p10 = P10(params.p10)

    aprx = arcpy.mp.ArcGISProject("CURRENT")
    m = aprx.activeMap
    layers = m.listLayers()

    test_layer = layers[0]
    inputSRS = arcpy.Describe(test_layer).spatialReference
    outputSRS = params.outputSRS # if params.outputSRS else inputSRS # inputSRS # 
    # equalSRS = (inputSRS == outputSRS or inputSRS.exportToString() == outputSRS.exportToString())
    # if equalSRS:
    #     outputSRS == inputSRS

    border_loc = p10['AdmBorder']['CLASSID'][2] + p10['AdmeNP']['CLASSID'][2] + p10['AdmeMO']['CLASSID'][2]
    omz_loc = params.ignore_cid
    fz_loc = p10['FunctionalZone']['CLASSID'][2]
    mo_loc = border_loc + fz_loc

    mask_layer = params.mask_layer

    geotransfName = create_custom_transformation(inputSRS, outputSRS, params.customGeoTransfm, params.geoTransfm)
    if params.mask_layer:
        clipping_mask = shapely.from_wkt([row[0] for row in arcpy.da.SearchCursor(params.mask_layer, ['SHAPE@WKT'], spatial_reference=outputSRS, datum_transformation=geotransfName)][0])
    else:
        clipping_mask = None

    border_gml = create_gml(params.output_dirname, 'Карта границ населенных пунктов (в том числе границ образуемых населенных пунктов).gml', params.xsd[0], params.gml_version)
    omz_gml = create_gml(params.output_dirname, 'Карта планируемого размещения объектов.gml', params.xsd[0], params.gml_version)
    fz_gml = create_gml(params.output_dirname, 'Карта функциональных зон поселения или городского округа.gml', params.xsd[0], params.gml_version)
    mo_gml = create_gml(params.output_dirname, 'Материалы по обоснованию в виде карт.gml', params.xsd[0], params.gml_version)


    for lyr in layers:
        if not lyr.isFeatureLayer:
            continue
        if mask_layer:
            if lyr.name == params.mask_layer.name:
                continue
        arcpy.AddMessage(lyr.name)
            # ПЕРЕПРОЕЦИРОВАНИЕ

        lyrDesc = arcpy.Describe(lyr)

        inputSRS = lyrDesc.spatialReference
        # if equalSRS:
        #     outputSRS = inputSRS
        geotransfName = create_custom_transformation(inputSRS, outputSRS, params.customGeoTransfm, params.geoTransfm)
        
        table, epsg = tab_to_gdf(lyr, spatial_reference=outputSRS, datum_transformation=geotransfName)

        if len(table) == 0:
            continue
            
        # Проверка соответствия названия слоя десятому приказу
        b_name = lyrDesc.baseName
        r_name = b_name.split('.')[-1].split('_')[0]
        
        if not p10.is_p10_layer(r_name):
            continue

        # Проверка соответствия названий столбцов десятому приказу
        table = p10.check_columns(table, r_name, b_name, params.rename_cols)
        
        # Смотрим тип слоя
        shape = lyrDesc.shapeType

        if shape == 'Polygon':
            name = r_name + '_Polygon'
            shape_type = 'polygon'
        elif shape == 'Polyline':
            name = r_name + '_LineString'
            shape_type = 'polyline'
        elif shape == 'Point':
            name = r_name + '_Point'
            shape_type = 'point'
        else:
            print('Неверный тип векторных данных')

        
        border_table = table.loc[table['CLASSID'].isin(border_loc)]

        no_clip = lyr.name in [lyr.name for lyr in params.mask_exceptions]

        fc_to_gml(border_gml, name, border_table, epsg, clipping_mask if not no_clip else None, params.OKTMO, p10.p10)
        
        fz_table = table.loc[table['CLASSID'].isin(fz_loc)]

        fc_to_gml(fz_gml, name, fz_table, epsg, clipping_mask if not no_clip else None, params.OKTMO, p10.p10)
        
        if 'STATUS' in table.columns and 'REG_STATUS' in table.columns:
            omz_table = table.loc[table['STATUS'].isin([2, 3]) & table['REG_STATUS'].isin(params.xsd[1]) & ~table['CLASSID'].isin(fz_loc)]
            if omz_loc:
                omz_table = omz_table.loc[~omz_table['CLASSID'].isin(omz_loc)]
            if params.omz_definition:
                for field, filt in params.omz_definition:
                    if field not in omz_table:
                        continue
                    omz_table = omz_table.loc[omz_table[field] != filt]

            fc_to_gml(omz_gml, name, omz_table, epsg, clipping_mask if not no_clip else None, params.OKTMO, p10.p10)
            
            mo_table = table.loc[~table['CLASSID'].isin(mo_loc) & (~table['STATUS'].isin([2, 3]) | (table['STATUS'].isin([2, 3]) & ~table['REG_STATUS'].isin(params.xsd[1])))]
        else:
            mo_table = table.loc[~table['CLASSID'].isin(mo_loc)]
            
        fc_to_gml(mo_gml, name, mo_table, epsg, clipping_mask if not no_clip else None, params.OKTMO, p10.p10)
        
        
    # border_gml.Release()
    # omz_gml.Release()
    # fz_gml.Release()
    # mo_gml.Release()

    border_gml = None
    omz_gml = None
    fz_gml = None
    mo_gml = None

    for filename in os.listdir(params.output_dirname):
        if not filename.endswith(".gml"):
            continue
        path = os.path.join(params.output_dirname, filename)
        fileSize = os.path.getsize(path) / (1024 * 1024)
        if params.splitSize and fileSize > params.splitSize:
            splitXml(path, params.splitSize, fileSize)

    for filename in os.listdir(params.output_dirname):
        if not filename.endswith(".gml"):
            continue
        if 'Карта планируемого размещения объектов' in filename:
            if filename == 'Карта планируемого размещения объектов.gml':
                shutil.copy(os.path.join(params.output_dirname, 'Карта планируемого размещения объектов.gml'), os.path.join(params.output_dirname, 'Приложение к положению о территориальном планировании в форме электронного документа.xml'))
            else:
                shutil.copy(os.path.join(params.output_dirname, filename), os.path.join(params.output_dirname, 'Приложение к положению о территориальном планировании в форме электронного документа_' + re.search(r'_(.*)\.', filename).groups(0)[0] + '.xml'))
        
        if 'Материалы по обоснованию в виде карт' in filename:
            if filename == 'Материалы по обоснованию в виде карт.gml':
                shutil.copy(os.path.join(params.output_dirname, 'Материалы по обоснованию в виде карт.gml'), os.path.join(params.output_dirname, 'Материалы по обоснованию в формате xml.xml'))
            else:
                shutil.copy(os.path.join(params.output_dirname, filename), os.path.join(params.output_dirname, 'Материалы по обоснованию в формате xml_' + re.search(r'_(.*)\.', filename).groups(0)[0] + '.xml'))
                

if __name__ == '__main__':
    not_valid_geoms = []
    execute()

    with open('not_valid.txt', 'w') as f:
        f.write(str(not_valid_geoms))

