import pandas as pd
import arcpy
import os

class Param:
    def __init__(self, param):
        self.value = param
        self.valueAsText = str(param)
        
class Params:
    def __init__(self, params):
        self.layer_1 = params[0].valueAsText
        self.field_1 = params[1].valueAsText
        self.border = params[2].valueAsText
        self.layer_2 = params[3].valueAsText
        self.field_2 = params[4].valueAsText if params[4].valueAsText != params[1].valueAsText else params[4].valueAsText + '_1'
        self.min_intersection = params[5].value
        self.exceptions = set(params[6].valueAsText) if params[6].valueAsText else None
        self.output_fc = params[7].valueAsText
        if arcpy.Describe(os.path.split(self.output_fc)[0]).dataType == 'Folder':
            self.output_fc += '.shp'
            self.field_1 = self.field_1[:10] if len(self.field_1) > 10 else self.field_1
            self.field_2 = self.field_2[:10] if len(self.field_2) > 10 else self.field_2
            if self.field_1 == self.field_2:
                self.field_2 = self.field_2[:8] + '_1'
        self.output_xls = params[8].valueAsText
        self.onlyMany = params[9].value
        

def table_to_data_frame(in_table, input_fields=None, where_clause=None):
    """Function will convert an arcgis table into a pandas dataframe with an object ID index, and the selected
    input fields using an arcpy.da.SearchCursor."""
    describe = arcpy.Describe(in_table)
    OIDFieldName = describe.OIDFieldName
    shapeFieldName = describe.shapeFieldName if hasattr(describe, 'shapeFieldName') else None
    if input_fields:
        final_fields = [OIDFieldName] + input_fields
        if shapeFieldName:
            final_fields += ['SHAPE@']
    else:
        final_fields = ['SHAPE@' if field.name == shapeFieldName else field.name for field in arcpy.ListFields(in_table)]
    data = [row for row in arcpy.da.SearchCursor(in_table, final_fields, where_clause=where_clause)]
    fc_dataframe = pd.DataFrame(data, columns=final_fields)
    fc_dataframe = fc_dataframe.set_index(OIDFieldName, drop=True)
    return fc_dataframe


def execute():
    params = Params(arcpy.GetParameterInfo())

    memFc = None
    if params.border:
        geoms = []
        erase = [row[0] for row in arcpy.da.SearchCursor(params.layer_1, 'SHAPE@')]
        for row in arcpy.da.SearchCursor(params.border, 'SHAPE@'):
            geom = row[0]
            for other in erase:
                geom = geom.difference(other)
            geoms.append([geom, 'Вне границ'])
        df = pd.DataFrame(geoms, columns=['SHAPE', params.field_1])
        arcpy.AddMessage(df)
        memFc = r'memory/temp_merge'
        arcpy.management.CopyFeatures(df, r'memory/temp')
        arcpy.management.Merge((params.layer_1, r'memory/temp'), memFc, add_source='ADD_SOURCE_INFO')

    arcpy.analysis.PairwiseIntersect((memFc or params.layer_1, params.layer_2), params.output_fc, "all", "", "input")
    

    df = table_to_data_frame(params.output_fc)
    df['AREA'] = df.apply(lambda row: row['SHAPE@'].area, axis=1, result_type='reduce')

    layer_2 = table_to_data_frame(params.layer_2)
    layer_2['AREA'] = layer_2.apply(lambda row: row['SHAPE@'].area if row['SHAPE@'] else 0, axis=1, result_type='reduce')

    df = df[[params.field_2, params.field_1, 'AREA', 'SHAPE@']]
    df = df.groupby([params.field_2, params.field_1]).agg({'AREA': 'sum'}).reset_index()
    layer_2 = layer_2.groupby(params.field_2).agg({'AREA': 'sum'}).reset_index()
    layer_2 = layer_2.rename(columns={'AREA': 'AREA_old'})
    df = df.merge(layer_2, 'left', params.field_2)
    df = df.loc[df['AREA'] > params.min_intersection]
    df['part'] = df['AREA'] / df['AREA_old'] * 100

    result = {}
    for n, row in df.iterrows():
        result[row[params.field_2]] = result.setdefault(row[params.field_2], dict())
        result[row[params.field_2]][row[params.field_1]] = round(row['AREA'], 2)

    df['objects'] = df.apply(lambda row: result[row[params.field_2]], axis=1, result_type='reduce')
    df = df.sort_values(by='part', ascending=False)
    df = df.drop_duplicates(subset=params.field_2, keep="first")

    df = df.round(2)

    if params.onlyMany:
        df['delete'] = df.apply(lambda row: 1 if (len(row['objects']) == 1 or set(row['objects'].keys()) - params.exceptions == set()) else 0, axis=1, result_type='reduce')
        df = df.loc[df['delete'] == 0]
        del df['delete']

    writer = pd.ExcelWriter(params.output_xls)
    df = df.rename(
        columns={
            params.field_1: f'Преобладающий элемент {params.field_1}', 
            'AREA': 'Площадь пересечения', 
            'AREA_old': 'Исходная площадь', 
            'part': 'Доля площади преобладающего элемента', 
            'objects': 'Площади всех элементов'
        }
                )
    df.to_excel(writer, sheet_name = 'Лист1', index = False)
    writer.close()

if __name__ == '__main__':
    execute()