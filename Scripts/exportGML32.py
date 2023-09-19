''' На карте должны быть все слои, которые экспортируются в GML в любой прямоугольной системе координат (на нужную поменяется автоматически)
Названия слоев берутся из базы - если делаете из шейпов - нужно называть их также как они называются по десятому приказу
'''
import arcpy
import shapely
from osgeo import ogr

import os
import shutil
from itertools import count

import pandas as pd
import json

import xml.etree.ElementTree as et
import xml.dom.minidom as minidom

class Params:
    """Входные параметры инструмента"""
    def __init__(self, params):
        xsds = {'Проект схемы территориального планирования Российской Федерации в области федерального транспорта (железнодорожного, воздушного, морского, внутреннего водного, трубопроводного) и автомобильных дорог федерального значения': 'Doc.10501010000.xsd', 'Проект схемы территориального планирования Российской Федерации в области обороны страны и безопасности государства': 'Doc.10503000000.xsd', 'Проект схемы территориального планирования Российской Федерации в области энергетики': 'Doc.10504000000.xsd', 'Проект схемы территориального планирования Российской Федерации в области высшего образования': 'Doc.10505000000.xsd', 'Проект схемы территориального планирования Российской Федерации в области здравоохранения': 'Doc.10506000000.xsd', 'Проект схемы территориального планирования Российской Федерации в иной области': 'Doc.10507000000.xsd', 'Проект схемы территориального планирования Российской Федерации в нескольких областях': 'Doc.10508000000.xsd', 'Проект схемы территориального планирования на часть территории Российской Федерации в одной или нескольких областях, подготовленный по решению Президента Российской Федерации или Правительства Российской Федерации': 'Doc.10509000000.xsd', 'Проекты схем территориального планирования двух и более субъектов Российской Федерации': 'Doc.10801000000.xsd', 'Проекты схем территориального планирования субъектов Российской Федерации': 'Doc.10803010000.xsd', 'Проекты генеральных планов городов федерального значения': 'Doc.10804010000.xsd', 'Проекты схем территориального планирования муниципальных районов': 'Doc.20101000000.xsd', 'Проекты генеральных планов поселений и генеральных планов городских округов': 'Doc.20201000000.xsd', 'Схема территориального планирования Российской Федерации в области федерального транспорта (железнодорожного, воздушного, морского, внутреннего водного, трубопроводного) и автомобильных дорог федерального значения': 'Doc.10601010000.xsd', 'Схема территориального планирования Российской Федерации в области обороны страны и безопасности государства': 'Doc.10603000000.xsd', 'Схема территориального планирования Российской Федерации в области энергетики': 'Doc.10604000000.xsd', 'Схема территориального планирования Российской Федерации в области высшего образования': 'Doc.10605000000.xsd', 'Схема территориального планирования Российской Федерации в области здравоохранения': 'Doc.10606000000.xsd', 'Схема территориального планирования Российской Федерации в иной области': 'Doc.10607000000.xsd', 'Схема территориального планирования Российской Федерации в нескольких областях': 'Doc.10608000000.xsd', 'Схема территориального планирования на часть территории Российской Федерации в одной или нескольких областях, подготовленная по решению Президента Российской Федерации или Правительства Российской Федерации': 'Doc.10609000000.xsd', 'Схемы территориального планирования двух и более субъектов Российской Федерации': 'Doc.10802000000.xsd', 'Схемы территориального планирования субъектов Российской Федерации': 'Doc.10803050000.xsd', 'Генеральные планы городов федерального значения': 'Doc.10804040000.xsd', 'Схемы территориального планирования муниципальных районов': 'Doc.20104000000.xsd', 'Генеральные планы поселений и генеральные планы городских округов': 'Doc.20204000000.xsd'}
        
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
                     atr_table['Condition'].to_list()):
            self.p10[i[0]][i[1]] = [i[2], i[3], [d for d in dom_table['Code'].loc[dom_table['Domain'] == i[4]].to_list()] if i[4] else [c for c in cid_table['CLASSID'].loc[cid_table['Layer'] == i[0]].to_list()] if i[1] == 'CLASSID' else None, i[5], i[4]]
    
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
        for col in required_columns:
            if col not in df.columns:
                for col2 in df.columns:
                    if col2 in col:
                        rename_columns[col2] = col
        df = df.rename(columns=rename_columns)
        missing_columns = set(required_columns) - set(df.columns.to_list())
        check_columns = required_columns - missing_columns
        check_columns = sorted(list(check_columns), key=lambda x: required_columns_list.index(x))
        return df, check_columns
    
def tab_to_df(in_table, input_fields=None, where_clause=None, spatial_reference=None, datum_transformation=None):
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
        fc_dataframe = pd.DataFrame(data, columns=final_fields)
        fc_dataframe = fc_dataframe.set_index(OIDFieldName, drop=True)
        return fc_dataframe



def create_custom_transformation(inputSRS, outputSRS, customGeoTransfm, geoTransfm):
    if inputSRS == outputSRS or inputSRS.exportToString() == outputSRS.exportToString():
        return None
    dirname = os.path.join(os.getenv('APPDATA'), 'Esri', 'ArcGISPro', 'ArcToolbox', 'CustomTransformations')
    transfList = arcpy.ListTransformations(inputSRS, outputSRS)
    if len(transfList) > 0:
        if geoTransfm in transfList:
            return geoTransfm
        return transfList[0]
    
    geotransfName = 'MSK_to_WGS{}'
    n = 1
    while os.path.isfile(os.path.join(dirname, geotransfName.format(n))):
        n += 1
    print(geotransfName.format(n))
    arcpy.management.CreateCustomGeoTransformation(geotransfName.format(n), inputSRS, outputSRS, customGeoTransfm)
    return geotransfName.format(n)

def get_dimension(shape_type, d_dict={'point': 1, 'polyline': 2, 'polygon': 4}):
    return d_dict[shape_type]

def update_extent(extent):
    global gml_extent
    if gml_extent:
        gml_extent = [max(i) for i in zip(gml_extent, [round(extent.XMin, 2), round(extent.YMin, 2), round(extent.XMax, 2), round(extent.YMax, 2)])]
    else:
        gml_extent = [round(extent.XMin, 2), round(extent.YMin, 2), round(extent.XMax, 2), round(extent.YMax, 2)]

def make_valid(row_object, empty_value, OKTMO, name, p10):
    name = name.split('_')[0]

    for col, value in row_object.items():

        if col == 'OKTMO':
            if type(value) == str:
                if len(value) < 2:
                    value = OKTMO
                else:
                    pass
            else:
                value = OKTMO

        if value is None and p10.get(name).get(col)[1] != 'Символьное':
            value = empty_value
        elif value is None and p10.get(name).get(col)[1] == 'Символьное':
            value = ''
        
        if p10.get(name).get(col)[1] == 'Целое':
            if isinstance(value, (int, numpy.int64)):
                value = int(value)
            else:
                value = 0

        if p10.get(name).get(col)[1] == 'Вещественное':
            if isinstance(value, (float, numpy.float64)):
                value = round(value, 2)
            else:
                value = 0            

        if col in ['DATE_START', 'DATE_CLOSE']:
            try:
                value = value.strftime('%Y/%m/%d')
            except:
                pass    
            

                
        if isinstance(value, (float, numpy.float64)):
            value = round(value, 2)
            

        row_object[col] = str(value)
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
        return shapely.MultiPoint([geom])
    if gtype == 'GeometryCollection':
        return geom

def geom_to_gml(geom, clipping_mask, no_clip):
    if not geom:
        return
    
    if clipping_mask and not no_clip:
        if geom.disjoint(clipping_mask):
            return 
        geom = geom.intersect(clipping_mask, get_dimension(geom.type))
    
    if geom.area == 0 and geom.type == 'polygon':
        return
    
    update_extent(geom.extent)
    
    shapely_obj = to_multigeo(shapely.from_wkt(geom.WKT))
    if not shapely_obj.is_valid:
        # print(gml_id)
        shapely_type = shapely_obj.geom_type
        shapely_obj = shapely.remove_repeated_points(shapely_obj)
        shapely_obj = shapely.make_valid(shapely_obj)
        shapely_obj = to_multigeo(shapely_obj)
        
        shapely_new_type = shapely_obj.geom_type
        if shapely_new_type != shapely_type:
            if shapely_new_type == 'GeometryCollection':
                geoms = {to_multigeo(geom).geom_type: to_multigeo(geom) for geom in shapely_obj.geoms}
                if shapely_type in geoms:
                    shapely_obj = geoms[shapely_type]
                else:
                    bad_geoms.append([geom, 'Потерян тип объекта при исправлении'])
                    # print('Объект потерян джонни')
                    return
            else:
                bad_geoms.append([geom, 'Потерян тип объекта при исправлении'])
                # print('Объект потерян джонни')
                return
        if not shapely_obj.is_valid:
            bad_geoms.append([geom, 'Ошибка после исправления геометрии'])
            return
    if shapely_obj.geom_type == 'MultiPolygon' and shapely_obj.area == 0:
        null_geoms.append(geom)
    wkt = shapely_obj.wkt

    ogr_shp = ogr.CreateGeometryFromWkt(wkt)
    return ogr_shp.ExportToGML(options=["FORMAT=GML32", 'NAMESPACE_DECL=YES', 'GMLID=change']) # 'GML3_LINESTRING_ELEMENT=curve'

def get_gml(gml):
    # et.register_namespace('xmlns:gml', "http://www.opengis.net/gml")
#     gml_namespace = ' xmlns:gml="http://www.opengis.net/gml">'
#     gml_string = gml.replace('>', gml_namespace, 1)
    gml_string = gml
    gml_elem = et.fromstring(gml_string)
    return gml_elem


def rows_as_dicts(table, columns, empty_value, OKTMO, clipping_mask, no_clip, name, p10):
    '''
    Yields rows from passed Arcpy da cursor as dicts
    '''
    table = table
    for row in range(len(table)):
        row_object = dict(zip(columns, table[columns].iloc[row]))
        row_object = make_valid(row_object, empty_value, OKTMO, name, p10)
        
        row_object['GML'] = geom_to_gml(table['SHAPE@'].iloc[row], clipping_mask, no_clip)
        yield row_object
#         if uc:
#             cursor.updateRow([row_object[colname] for colname in colnames])

def dump2xml(row, stands, elelst, gml_name, gml_id):
    '''
    Builds the xml tree from the passed row dict
    '''

    # stand level creation
    fm = et.Element("fgistp:featureMember")
    
    stand = et.SubElement(fm, 'fgistp:' + gml_name)

    stand.set("gml:id", str(next(gml_id)))

    # add field elements with their values
    for e in elelst:
        xele = et.SubElement(stand, 'fgistp:' + e)
        xele.text = str(row[e])
    
    
    stand.append(get_gml(row['GML']))
    
    

    #add to top level stands element
    stands.append(fm)

def fc_to_gml(table, fields, stands, gml_name, empty_value, OKTMO, clipping_mask, gml_id, no_clip, p10):
    
    for row in rows_as_dicts(table, fields, empty_value, OKTMO, clipping_mask, no_clip, gml_name, p10):
        if not row['GML']:
            continue
        dump2xml(row, stands, fields, gml_name, gml_id)

    #throw the entire xml tree to a file        
    xmltree = et.ElementTree(stands)
    # 
    
def create_gml(xsd):
    et.register_namespace("", "http://www.opengis.net/gml/3.2")
    et.register_namespace("gml", "http://www.opengis.net/gml/3.2")
    gml_id = count(1)
    name_space = {
        'xmlns:xsi': "http://www.w3.org/2001/XMLSchema-instance",
        'xmlns:xlink': "http://www.w3.org/1999/xlink",
        # 'xmlns:gml': "http://www.opengis.net/gml",
        'xmlns:fgistp': "http://fgistp",
        'xsi:schemaLocation': f"http://fgistp {xsd}", # Doc.20204000000.xsd
        'gml:id': "FeatureCollection.1"
    }
    # extent = clipping_mask.extent
    envelope = '<gml:boundedBy  xmlns:gml="http://www.opengis.net/gml/3.2"><gml:Envelope srsName="EPSG:3857"><gml:lowerCorner>{} {}</gml:lowerCorner><gml:upperCorner>{} {}</gml:upperCorner></gml:Envelope></gml:boundedBy>' #.format(round(extent.XMin, 2), round(extent.YMin, 2), round(extent.XMax, 2), round(extent.YMax, 2))
    stands = et.Element("fgistp:FeatureCollection", name_space)
    stands.append(et.fromstring(envelope))
    return stands, gml_id

class CommentedTreeBuilder(et.TreeBuilder):
    """Класс для сохранения комментариев при парсинге XML"""
    def comment(self, data):
        self.start(et.Comment, {})
        self.data(data)
        self.end(et.Comment)

def get_ID(name, gml_id):

    return f'{name}.{next(gml_id)}'

def save_gml(gml, dirname, filename, p10):
    lowerCorner = list(list(list(gml)[0])[0])[0]
    lowerCorner.text = lowerCorner.text.format(gml_extent[0], gml_extent[1])
    upperCorner = list(list(list(gml)[0])[0])[1]
    upperCorner.text = upperCorner.text.format(gml_extent[2], gml_extent[3])

    layers_ids = {}
    for elem in gml.iter():
        if elem.tag == 'fgistp:featureMember':
            name = list(elem)[0].tag.replace('fgistp:', '')
            layers_ids.setdefault(name, count(1))
        if 'gml:id' in elem.attrib and 'FeatureCollection' not in elem.tag:
            elem.attrib['gml:id'] = name + '.' + str(next(layers_ids[name]))
        if '{http://www.opengis.net/gml/3.2}id' in elem.attrib and 'FeatureCollection' not in elem.tag:
            elem.attrib['{http://www.opengis.net/gml/3.2}id'] = name + '.' + str(next(layers_ids[name]))
    
    xmltree = et.ElementTree(gml)

    xmlstr = et.tostring(xmltree.getroot(), encoding='UTF-8', short_empty_elements=False)
    xmlpretty = minidom.parseString(xmlstr).toprettyxml(indent='\t', newl='\n', encoding='UTF-8')
    with open(os.path.join(dirname, filename), 'wb') as f:
        f.write(xmlpretty)
        
    gml_path = os.path.join(dirname, filename)
    gfs_path = os.path.join(os.path.split(gml_path)[0], os.path.split(gml_path)[1].replace('gml', 'gfs'))

    # НОВАЯ ПОПЫТКА
    
    gfs = et.Element('GMLFeatureClassList')

    layer_list = []
    geom_types = {'Point': '4', 'Line': '5', 'Polygon': '6'}
    #geom_types = {'Point': '4', 'LineString': '5', 'Polygon': '6'}
    for fm in gml.iter('fgistp:featureMember'):
        layer_name = list(fm)[0].tag.replace('fgistp:', '')
        layer_params = layer_name.split('_')
        #layer_type = list(list(list(list(fm)[0])[-1])[0])[0].tag.split('}')[1]
        #layer_params = [layer_name, layer_type]
        if layer_name in layer_list:
            continue
        else:
            layer_list.append(layer_name)

        fields = [elem.tag.replace('fgistp:', '') for elem in list(filter(lambda x: 'fgistp:' in x.tag, list(list(fm)[0])))]
        stand = et.SubElement(gfs, 'GMLFeatureClass')
        xele = et.SubElement(stand, 'Name')
        xele.text = layer_name
        xele = et.SubElement(stand, 'ElementPath')
        xele.text = layer_name
        xele = et.SubElement(stand, 'GeometryType')
        xele.text = geom_types[layer_params[1]]
        for field in fields:
            arcpy.AddMessage(layer_params[0] + ' ' + field)
            xele = et.SubElement(stand, 'PropertyDefn')
            fele = et.SubElement(xele, 'Name')
            fele.text = field
            fele = et.SubElement(xele, 'ElementPath')
            fele.text = field
            field_type = et.SubElement(xele, 'Type')
            required_type = p10.get(layer_params[0]).get(field)[1]
            if required_type == 'Целое' or field == 'CLASSID':
                field_type.text = 'Integer'
            elif required_type == 'Вещественное':
                field_type.text = 'Real'
            else:
                field_type.text = 'String'
        
    xmlstr = et.tostring(gfs, short_empty_elements=False)
    xmlpretty = minidom.parseString(xmlstr).toprettyxml(indent='  ', newl='\n')
    with open(gfs_path, 'w') as f:
        f.write(xmlpretty)
    # НОВАЯ ПОПЫТКА

    # if os.path.isfile(gfs_path):
    #     os.remove(gfs_path)

    # ogr.Open(gml_path) # открываем GML для создания файла GFS
    
    # parser = et.XMLParser(target=CommentedTreeBuilder(), encoding='utf-8')
    # gfs = et.parse(gfs_path, parser)
    # for layer in gfs.iter('GMLFeatureClass'):
    #     layer_name = layer[1].text
    #     for field in layer.iter('PropertyDefn'): 
    #         field_name, field_type = field[1], field[2]
    #         required_type = p10.get(layer_name.split('_')[0]).get(field_name.text)[1]
    #         if required_type == 'Целое' or field_name.text == 'CLASSID':
    #             field_type.text = 'Integer'
    #         elif required_type == 'Вещественное':
    #             field_type.text = 'Real'
    #         else:
    #             field_type.text = 'String'
    # xmlstr = et.tostring(gfs.getroot(), encoding='UTF-8')
    # with open(os.path.join(gfs_path), 'wb') as f:
    #     f.write(xmlstr)

def execute():
    params = Params(arcpy.GetParameterInfo())
    p10 = P10(params.p10)

    aprx = arcpy.mp.ArcGISProject("CURRENT")
    m = aprx.activeMap
    layers = m.listLayers()

    test_layer = layers[0]
    inputSRS = arcpy.Describe(test_layer).spatialReference
    outputSRS = params.outputSRS if params.outputSRS else inputSRS # inputSRS # 
    equalSRS = not outputSRS

    bad_geoms = []
    null_geoms = []


    border_loc = p10['AdmBorder']['CLASSID'][2] + p10['AdmeNP']['CLASSID'][2] + p10['AdmeMO']['CLASSID'][2]
    omz_loc = params.ignore_cid
    fz_loc = p10['FunctionalZone']['CLASSID'][2]
    mo_loc = border_loc + fz_loc

    mask_layer = params.mask_layer

    geotransfName = create_custom_transformation(inputSRS, outputSRS, params.customGeoTransfm, params.geoTransfm)
    if params.mask_layer:
        clipping_mask = [row[0] for row in arcpy.da.SearchCursor(params.mask_layer, ['SHAPE@'], spatial_reference=outputSRS, datum_transformation=geotransfName)][0]
    else:
        clipping_mask = None

    border_gml, border_counter = create_gml(params.xsd)
    omz_gml, omz_counter = create_gml(params.xsd)
    fz_gml, fz_counter = create_gml(params.xsd)
    mo_gml, mo_counter = create_gml(params.xsd)

    for tab in layers:
        if mask_layer:
            if tab.name == params.mask_layer.name:
                continue
        print(tab.name)
            # ПЕРЕПРОЕЦИРОВАНИЕ


        inputSRS = arcpy.Describe(tab).spatialReference
        if equalSRS:
            outputSRS = inputSRS
        geotransfName = create_custom_transformation(inputSRS, outputSRS, params.customGeoTransfm, params.geoTransfm)
        
        table = tab_to_df(tab, spatial_reference=outputSRS, datum_transformation=geotransfName)

        table['SHAPE@'] = table.apply(lambda row: arcpy.FromWKT(row['SHAPE@WKT']) if row['SHAPE@WKT'] else row['SHAPE@WKT'], axis=1, result_type='reduce')
        table = table.astype(object).where(pd.notnull(table), None)
        if len(table) == 0:
            continue
            
        # Проверка соответствия названия слоя десятому приказу
        b_name = arcpy.Describe(tab).baseName
        r_name = b_name.split('.')[-1].split('_')[0]
        
        if not p10.is_p10_layer(r_name):
            continue
        
        # Проверка соответствия названий столбцов десятому приказу
        table, columns_to_gml = p10.check_columns(table, r_name, b_name, params.rename_cols)
        
        # Смотрим тип слоя
        shape = arcpy.Describe(tab).shapeType

        if shape == 'Polygon':
            name = r_name + '_Polygon'
            shape_type = 'polygon'
        elif shape == 'Polyline':
            name = r_name + '_Line'
            shape_type = 'polyline'
        elif shape == 'Point':
            name = r_name + '_Point'
            shape_type = 'point'
        else:
            print('Неверный тип векторных данных')
        # ПРОВЕРКА
        #name = r_name
        
        border_table = table.loc[table['CLASSID'].isin(border_loc)]

        no_clip = tab.name in [lyr.name for lyr in params.mask_exceptions]

        fc_to_gml(border_table, columns_to_gml, border_gml, name, params.empty_value, params.OKTMO, clipping_mask, border_counter, no_clip, p10.p10)
        
        fz_table = table.loc[table['CLASSID'].isin(fz_loc)]

        fc_to_gml(fz_table, columns_to_gml, fz_gml, name, params.empty_value, params.OKTMO, clipping_mask, fz_counter, no_clip, p10.p10)
        
        if 'STATUS' in table.columns and 'REG_STATUS' in table.columns:
            omz_table = table.loc[table['STATUS'].isin([2, 3]) & table['REG_STATUS'].isin([4, 5]) & ~table['CLASSID'].isin(fz_loc)]
            if omz_loc:
                omz_table = omz_table.loc[~omz_table['CLASSID'].isin(omz_loc)]
            if params.omz_definition:
                for field, filt in params.omz_definition:
                    if field not in omz_table:
                        continue
                    omz_table = omz_table.loc[omz_table[field] != filt]

            fc_to_gml(omz_table, columns_to_gml, omz_gml, name, params.empty_value, params.OKTMO, clipping_mask, omz_counter, no_clip, p10.p10)
            
            mo_table = table.loc[~table['CLASSID'].isin(mo_loc) & (~table['STATUS'].isin([2, 3]) | (table['STATUS'].isin([2, 3]) & ~table['REG_STATUS'].isin([4, 5])))]
        else:
            mo_table = table.loc[~table['CLASSID'].isin(mo_loc)]
            
        fc_to_gml(mo_table, columns_to_gml, mo_gml, name, params.empty_value, params.OKTMO, clipping_mask, mo_counter, no_clip, p10.p10)
        
        
    if len(bad_geoms) > 1 or len(null_geoms) > 1:
        arcpy.AddMessage('ПРИСУТСТВУЮТ НЕ ИСПРАВЛЕННЫЕ ОШИБКИ ГЕОМЕТРИИ')

    save_gml(border_gml, params.output_dirname, 'Карта границ населенных пунктов (в том числе границ образуемых населенных пунктов).gml', p10.p10)
    save_gml(omz_gml, params.output_dirname, 'Карта планируемого размещения объектов.gml', p10.p10)
    save_gml(fz_gml, params.output_dirname, 'Карта функциональных зон поселения или городского округа.gml', p10.p10)
    save_gml(mo_gml, params.output_dirname, 'Материалы по обоснованию в виде карт.gml', p10.p10)

    shutil.copy(os.path.join(params.output_dirname, 'Карта планируемого размещения объектов.gml'), os.path.join(params.output_dirname, 'Приложение к положению о территориальном планировании в форме электронного документа.xml'))
    shutil.copy(os.path.join(params.output_dirname, 'Материалы по обоснованию в виде карт.gml'), os.path.join(params.output_dirname, 'Материалы по обоснованию в формате xml.xml'))    

if __name__ == '__main__':

    gml_extent = None
    execute()