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
        self.fields = [field.value for field in params[4].values]
        self.min_intersection = params[5].value
        self.exceptions = set(params[6].values) | {'Вне границ'} if params[6].values else {'Вне границ'}
        self.output_fc = params[7].valueAsText
        if arcpy.Describe(os.path.split(self.output_fc)[0]).dataType == 'Folder':
            self.output_fc += '.shp'
            self.field_1 = self.field_1[:10] if len(self.field_1) > 10 else self.field_1
            for i, field in enumerate(self.fields):
                self.fields[i] = self.fields[i][:10] if len(self.fields[i]) > 10 else self.fields[i]
                if self.field_1 == self.fields[i]:
                    self.fields[i] = self.fields[i][:8] + '_1'
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
    writer = pd.ExcelWriter(params.output_xls)

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
        memFc = r'memory/temp'
        arcpy.management.CopyFeatures(params.layer_1, memFc)
        with arcpy.da.InsertCursor(memFc, ['SHAPE@', params.field_1]) as cursor:
            for i, row in df.iterrows():
                cursor.insertRow(row.to_list())

    arcpy.analysis.PairwiseIntersect((memFc or params.layer_1, params.layer_2), params.output_fc, "all", "", "input")
    

    df = table_to_data_frame(params.output_fc)
    df['AREA'] = df.apply(lambda row: row['SHAPE@'].area, axis=1, result_type='reduce')

    df_2 = table_to_data_frame(params.layer_2)
    df_2['AREA'] = df_2.apply(lambda row: row['SHAPE@'].area if row['SHAPE@'] else 0, axis=1, result_type='reduce')

    for field in params.fields:
        if field == params.field_1:
            field = field + '_1'
        current_df = df[[field, params.field_1, 'AREA', 'SHAPE@']]
        current_df = current_df.groupby([field, params.field_1]).agg({'AREA': 'sum'}).reset_index()
        current_df_2 = df_2.groupby(field).agg({'AREA': 'sum'}).reset_index()
        current_df_2 = current_df_2.rename(columns={'AREA': 'AREA_old'})
        current_df = current_df.merge(current_df_2, 'left', field)
        current_df = current_df.loc[current_df['AREA'] > params.min_intersection]
        current_df['part'] = current_df['AREA'] / current_df['AREA_old'] * 100

        result = {}
        for n, row in current_df.iterrows():
            result[row[field]] = result.setdefault(row[field], dict())
            result[row[field]][row[params.field_1]] = round(row['AREA'], 2)

        current_df['objects'] = current_df.apply(lambda row: result[row[field]], axis=1, result_type='reduce')
        current_df = current_df.sort_values(by='part', ascending=False)
        current_df = current_df.drop_duplicates(subset=field, keep="first")

        current_df = current_df.round(2)

        if params.onlyMany:
            current_df['delete'] = current_df.apply(lambda row: 1 if (len(row['objects']) == 1 or set(row['objects'].keys()) - params.exceptions == set()) else 0, axis=1, result_type='reduce')
            current_df = current_df.loc[current_df['delete'] == 0]
            del current_df['delete']

    
        current_df = current_df.rename(
            columns={
                params.field_1: f'Преобладающий элемент {params.field_1}', 
                'AREA': 'Площадь пересечения', 
                'AREA_old': 'Исходная площадь', 
                'part': 'Доля площади преобладающего элемента', 
                'objects': 'Площади всех элементов'
            }
                    )
        current_df.to_excel(writer, sheet_name = field, index = False)
    writer.close()

if __name__ == '__main__':
    execute()