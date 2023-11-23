import pandas as pd
import shapely
import arcpy

class Params:
    def __init__(self, params):
        self.input_fs = params[0].value
        self.output_xls = params[1].valueAsText
        self.join_np_types = params[2].value

def getRingCoords(ring):
    xx, yy = ring.coords.xy
    return list(zip(xx, yy))

def getPolygonCoords(geom):
    exterior = getRingCoords(geom.exterior)
    interiors = []
    for interior in geom.interiors:
        interiors.append(getRingCoords(interior))
    return [exterior] + interiors

def getCoords(geom):
    geoms = []
    if geom.geom_type == 'Polygon':
        return [getPolygonCoords(geom)]
    for g in geom.geoms:
        geoms.append(getPolygonCoords(g))
    return geoms

def merge(geometry):
    merged_geom = shapely.from_wkt(geometry.iloc[0])
    for i, geom in enumerate(geometry):
        if i == 0:
            continue
        merged_geom = shapely.union(merged_geom, shapely.from_wkt(geom))
    return merged_geom

def tab_to_gdf(in_table, input_fields=None, where_clause=None, spatial_reference=None, datum_transformation=None):
        """Function will convert an arcgis table into a pandas dataframe with an object ID index, and the selected
        input fields using an arcpy.da.SearchCursor."""
        OIDFieldName = arcpy.Describe(in_table).OIDFieldName
        try:
            shapeFieldName = arcpy.Describe(in_table).shapeFieldName
        except:
            shapeFieldName = None
        if input_fields:
            input_fields = [field for field in input_fields]
            if shapeFieldName:
                final_fields = [OIDFieldName] + ['SHAPE@WKT'] + input_fields
            else:
                final_fields = [OIDFieldName] + input_fields
        else:
            final_fields = [field.name if field.name != shapeFieldName else 'SHAPE@WKT' for field in arcpy.Describe(in_table).fields]
        data = [row for row in arcpy.da.SearchCursor(in_table, final_fields, where_clause=where_clause, spatial_reference=spatial_reference, datum_transformation=datum_transformation)]
        df = pd.DataFrame(data, columns=final_fields)
        df = df.rename(columns={'SHAPE@WKT': 'geometry'})
        srs = spatial_reference or arcpy.Describe(in_table).spatialReference
        fc_dataframe = df
        # fc_dataframe = fc_dataframe.set_index(OIDFieldName, drop=True)
        return fc_dataframe, srs.factoryCode

def execute():
    params = Params(arcpy.GetParameterInfo())
    df, _ = tab_to_gdf(params.input_fs)

    settl_type = pd.read_excel(r"..\Defaults\p10.xlsx", sheet_name='Справочники')
    settl_type = settl_type.loc[settl_type['Domain'] == 'SETTL_TYPE']

    def sort_NP(ds):
        sort_list = set(ds)
        sort_list = sort_list - {'МО'}
        sort_list = sorted(sort_list)
        sort_list.append('МО')
        return ds.apply(lambda value: sort_list.index(value))
    
    arcpy.AddMessage(str(df['NP']))
    # result = df[['CadNumber', 'NP', 'SETTL_TYPE', 'NAME_UCHLE', 'geometry']].groupby(['CadNumber', 'NP', 'SETTL_TYPE', 'NAME_UCHLE'], dropna=False).agg(merge).reset_index()
    result = df[['CadNumber', 'NP', 'SETTL_TYPE', 'geometry']].groupby(['CadNumber', 'NP', 'SETTL_TYPE'], dropna=False).agg(merge).reset_index()
    if params.join_np_types:
        result = result.merge(settl_type, how = 'left', left_on='SETTL_TYPE', right_on='Code')
        result['NP_NAME'] = result.apply(lambda row: 'МО' if row['NP'] == 'МО' else row['Value'] + ' ' + row['NP'], axis=1)
    else:
        result['NP_NAME'] = result['NP']
    # result = result.sort_values(['NP', 'NAME_UCHLE', 'CadNumber'])
    result = result.sort_values(['NP', 'CadNumber'])
    result = result.sort_values(['NP'], key=sort_NP)

    coords_table = []
    for i, row in result.iterrows():
        coords = getCoords(row.geometry)
        for ci, geom in enumerate(coords):
            n = 1
            print_n = n
            for ri, part in enumerate(geom):
                for pi, xy in enumerate(part):
                    if pi == 0 and ri > 0:
                        current_n = n
                    if pi == len(part) - 1 and ri == 0:
                        print_n = 1
                        n -= 1
                    elif pi == len(part) - 1 and ri > 0:
                        print_n = current_n
                        n -= 1
                    x, y = xy
                    # coords_table.append([row.NP_NAME, row.NAME_UCHLE, row.CadNumber, ci + 1, print_n, y, x])
                    coords_table.append([row.NP_NAME, row.CadNumber, ci + 1, print_n, y, x])
                    n += 1
                    print_n = n

    # df = pd.DataFrame(coords_table, columns=['Населенный пункт', 'Участковое лесничество', 'Кадастровый номер', 'Номер контура', 'Номер точки', 'X, м', 'Y, м'])
    df = pd.DataFrame(coords_table, columns=['Населенный пункт', 'Кадастровый номер', 'Номер контура', 'Номер точки', 'X, м', 'Y, м'])
    df.to_excel(params.output_xls, index=False)

if __name__ == '__main__':
    execute()