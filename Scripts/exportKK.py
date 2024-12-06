''' На карте должны быть все слои, которые экспортируются в GML в любой прямоугольной системе координат (на нужную поменяется автоматически)
Названия слоев берутся из базы - если делаете из шейпов - нужно называть их также как они называются по десятому приказу
'''
import arcpy
import shapely
import geopandas as gpd
# from osgeo import ogr
from osgeo import osr
from osgeo import gdal
gdal.UseExceptions()

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
        self.project_type = params[0].valueAsText

        self.p10 = params[1].valueAsText

        rename_cols = [i.split('=') for i in params[2].values] if params[2].values else None
        self.rename_cols = {i[0]: i[1] for i in rename_cols} if rename_cols else None
        
        self.mask_layer = params[3].value # Введите название
        self.mask_exceptions = params[4].values if params[5].values else []

        self.OKTMO = params[5].valueAsText # Введите ОКТМО поселения, если где-то октмо будет не заполнено - оно заполнится этим автоматически
        self.empty_value = params[6].value

        self.ignore_cid = params[7].values

        self.omz_definition = [i.split('<>') for i in params[9].values] if params[8].value else None
        
        self.output_dirname = params[9].valueAsText

        self.vector_format = params[10].valueAsText
        

       

class P10:
    def __init__(self, p10):
        xls = pd.ExcelFile(p10)
        cid_table = pd.read_excel(xls, 'ClassID')
        cid_table = cid_table.where(pd.notnull(cid_table), None)
        atr_table = pd.read_excel(xls, 'Общий')
        atr_table = atr_table.where(pd.notnull(atr_table), None)
        dom_table = pd.read_excel(xls, 'Справочники')
        dom_table = dom_table.where(pd.notnull(dom_table), None)
        catalog_table = pd.read_excel(xls, 'Слои')
        catalog_table = catalog_table.where(pd.notnull(catalog_table), None)
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
            
        self.catalog = {row["Layer"]: {"Folder": row["Папка"], "Geometry": row["GeomKK"] if row["GeomKK"] is None else row["GeomKK"].split(";")} for i, row in catalog_table.iterrows()}
    
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
        df.columns=df.columns.str.slice(0, 10) # обрезаем до 10 символов
        check_columns = [column[:10] for column in check_columns]
        return df[check_columns + ['geometry']]
    
def tab_to_gdf(in_table, input_fields=None, where_clause=None):
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
        data = [row for row in arcpy.da.SearchCursor(in_table, final_fields, where_clause=where_clause)]
        df = pd.DataFrame(data, columns=final_fields)
        gs = gpd.GeoSeries.from_wkt(df['SHAPE@WKT'])
        del df['SHAPE@WKT']
        srs = arcpy.Describe(in_table).spatialReference
        fc_dataframe = gpd.GeoDataFrame(df, geometry=gs, crs=srs.factoryCode) #srs.exportToString()
        # fc_dataframe = fc_dataframe.set_index(OIDFieldName, drop=True)
        return fc_dataframe, srs.factoryCode


def get_dimension(shape_type, d_dict={'point': 1, 'polyline': 2, 'polygon': 4}):
    return d_dict[shape_type]


def makeValidForP10(row_object, name, OKTMO, p10):
    name = name.split('_')[0]

    for col, value in row_object.items():
        # если колонка не по приказу пропускаем
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

        # Нули преобразуем в null, для шейпов
        if value == 0:
            value = None
        # NaN Заменяем на None
        if value != value:
            value = None
        # Для null вписываем значения по дефолту
        if value is None:
            value = p10.get(name).get(col)[5]
        # Для пустых строк вписываем значения по дефолту
        if type(value) == str:
            if value.isspace() or value == "":
                value = p10.get(name).get(col)[5]

        # Если тип поля целое число - пробуем конвертировать в целое число либо заменяем на ноль для значений со справочниками
        if p10.get(name).get(col)[1] == 'Целое':
            try:
                value = int(value)
            except:
                if p10.get(name).get(col)[4] is not None: # если домен то None если не домен то 0?? фига знает как лучше
                    value = 0

        if p10.get(name).get(col)[1] == 'Вещественное':
            try:
                value = round(value, 6)
            except:
                value = 0

        if isinstance(value, (float, float64)):
            value = round(value, 6)

        # Преобразуем null в пустые строки
        # if value != value or value is None:
        #     value = ""
        if p10.get(name).get(col)[1] == 'Символьное':
            if value is None:
                value = ""
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


def get_shapefile(dirname, filename, vector_format):
    if not os.path.exists(dirname):
        os.makedirs(dirname)

    # driver = gdal.GetDriverByName("ESRI Shapefile")
    # driver = gdal.GetDriverByName("MapInfo File")
    driver = gdal.GetDriverByName(vector_format)

    path = os.path.join(dirname, filename)

    if os.path.exists(dirname):
        outDataSource = gdal.OpenEx(dirname, gdal.OF_UPDATE) #allowed_drivers=["MapInfo File"])
    else:
        outDataSource = driver.Create(dirname, 0,0,0,0)#, ["FORMAT=MIF"])

    return outDataSource



def fc_to_gml(outSource, layerName, gdf, epsg, mask, oktmo, p10, gtype):
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
    
    # gdal.VectorTranslate(outSource, inSource, accessMode="append", geometryType="PROMOTE_TO_MULTI", layerName=layerName, layerCreationOptions=["ENCODING=utf-8"])
    srs = osr.SpatialReference()
    srs.ImportFromWkt('LOCAL_CS["Nonearth",UNIT["Meter",1.0]]')
    gdal.VectorTranslate(outSource, inSource, accessMode="append", layerName=layerName)#, geometryType="PROMOTE_TO_MULTI")#, layerCreationOptions=["ENCODING=cp1251", "BOUNDS=-10000000,-10000000,10000000,10000000"], dstSRS=srs, reproject=False)
    
def get_layer_name(name, shape_type, include_shape_type, status):
    return name + ("_pr" if status == 1 and name != "AreaBaseDevelopment" else "") + (shape_type if include_shape_type else "")

def execute():
    params = Params(arcpy.GetParameterInfo())
    p10 = P10(params.p10)

    aprx = arcpy.mp.ArcGISProject("CURRENT")
    m = aprx.activeMap
    layers = m.listLayers()

    mask_layer = params.mask_layer

    if params.mask_layer:
        clipping_mask = shapely.from_wkt([row[0] for row in arcpy.da.SearchCursor(params.mask_layer, ['SHAPE@WKT'])][0])
    else:
        clipping_mask = None



    for lyr in layers:
        if not lyr.isFeatureLayer:
            continue
        if mask_layer:
            if lyr.name == params.mask_layer.name:
                continue
        arcpy.AddMessage(lyr.name)
            # ПЕРЕПРОЕЦИРОВАНИЕ

        lyrDesc = arcpy.Describe(lyr)
        
        table, epsg = tab_to_gdf(lyr)

        if len(table) == 0:
            continue
            
        # Проверка соответствия названия слоя десятому приказу
        b_name = lyrDesc.baseName
        r_name = b_name.split('.')[-1].split('_')[0]
        
        if not p10.is_p10_layer(r_name):
            continue

        # Проверка соответствия названий столбцов десятому приказу
        table = p10.check_columns(table, r_name, b_name, params.rename_cols)
        


        

        layer_dict = p10.catalog[r_name]
        
        status_field = "STATUS" if "STATUS" in table.columns else "STATUS_ADM" if "STATUS_ADM" in table.columns else None

        op_table = table
        gp_table = None
        if r_name == "TerritorialZone":
            gp_table = table
            op_table = None
        else:
            if status_field:
                op_table = table.loc[table[status_field] == 1]
                gp_table = table.loc[table[status_field] != 1]


        no_clip = lyr.name in [lyr.name for lyr in params.mask_exceptions]

        # Смотрим тип слоя
        shape = lyrDesc.shapeType

        if shape == 'Polygon':
            gtype = 'Polygon'
            shape_type = '_pol'
        elif shape == 'Polyline':
            gtype = 'LineString'
            shape_type = '_lin'
        elif shape == 'Point':
            gtype = 'Point'
            shape_type = '_pnt'
        else:
            print('Неверный тип векторных данных')
            
        op_name = get_layer_name(r_name, shape_type, layer_dict["Geometry"] is not None, 0)
        gp_name = get_layer_name(r_name, shape_type, layer_dict["Geometry"] is not None, 1)

        if op_table is not None and op_table.shape[0] > 0:
            fc_to_gml(get_shapefile(os.path.join(params.output_dirname, params.project_type, "Опорный план", layer_dict["Folder"]), op_name, params.vector_format), op_name, op_table, epsg, clipping_mask if not no_clip else None, params.OKTMO, p10.p10, gtype)
        if gp_table is not None and gp_table.shape[0] > 0:
            fc_to_gml(get_shapefile(os.path.join(params.output_dirname, params.project_type, params.project_type, layer_dict["Folder"]), gp_name, params.vector_format), gp_name, gp_table, epsg, clipping_mask if not no_clip else None, params.OKTMO, p10.p10, gtype)
        
                

if __name__ == '__main__':
    not_valid_geoms = []
    execute()

    with open('not_valid.txt', 'w') as f:
        f.write(str(not_valid_geoms))

