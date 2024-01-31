import pandas as pd
import arcpy
import os

from datetime import datetime, timezone

# class Param:
#     def __init__(self, param):
#         self.valueAsText = str(param)
#         self.value = str(param)
        
class Params:
    def __init__(self, params):
        self.input_xls = params[0].valueAsText
        
        self.NP = params[1].valueAsText
        self.ZU = params[2].valueAsText
        self.FZ = params[3].valueAsText
        self.LU = params[4].valueAsText
        self.DU = params[5].valueAsText
        self.du_type = params[6].valueAsText
        
        self.output_xls = params[7].valueAsText
                
def table_to_data_frame(in_table, input_fields=None, where_clause=None):
    """Function will convert an arcgis table into a pandas dataframe with an object ID index, and the selected
    input fields using an arcpy.da.SearchCursor."""
    OIDFieldName = arcpy.Describe(in_table).OIDFieldName
    try:
        shapeFieldName = arcpy.Describe(in_table).shapeFieldName
    except:
        shapeFieldName = None
    if input_fields:
        if shapeFieldName:
            final_fields = [OIDFieldName] + ['SHAPE@'] + input_fields
        else:
            final_fields = [OIDFieldName] + input_fields
    else:
        final_fields = [field.name if field.name != shapeFieldName else 'SHAPE@' for field in arcpy.ListFields(in_table)]
    data = [row for row in arcpy.da.SearchCursor(in_table, final_fields, where_clause=where_clause)]
    fc_dataframe = pd.DataFrame(data, columns=final_fields)
    fc_dataframe = fc_dataframe.set_index(OIDFieldName, drop=True)
    return fc_dataframe

def count_transfers(df, df_du, right_on, np_dict, lu_dict, fz_dict, du_type=True):

    fields = ['name', 'np', 'cadnumber', 'category', 'category_plan']
    df['cadnumber'] = df.apply(lambda row: row['cadnumber'] if row['cadnumber'] else 'Территория, собственность на которую не разграничена', axis=1, result_type='reduce')
    df['area'] = df.apply(lambda row: row["SHAPE@"].area / 10000, axis=1, result_type='reduce')
    df['OLD_FID'] = df.index

    
    df_du['area_du'] = df_du.apply(lambda row: row["SHAPE@"].area / 10000, axis=1, result_type='reduce')

    df = df.merge(df_du, 'left', left_on='OLD_FID', right_on=right_on, validate="one_to_many")
    df['category'] = df.apply(lambda row: row["cat"] if (row["cat"] and not row["cat"].startswith("Земли лесного")) else lu_dict[row['CLASSID']], axis=1, result_type='reduce')
    
#     df['category'] = df.apply(lambda row: lu_dict[row['CLASSID']], axis=1, result_type='reduce')
    
   
    df['iskl_v_np'] = False
    if 'name' in df.columns:
        df['iskl_v_np'] = df.apply(lambda row: True if row['SETTL_TYPE'] > 100 else False, axis=1, result_type='reduce')
        df['SETTL_TYPE'] = df.apply(lambda row: row['SETTL_TYPE'] if row['SETTL_TYPE'] < 101 else row['SETTL_TYPE'] - 100, axis=1, result_type='reduce')
        df['np'] = df.apply(lambda row: np_dict[row['SETTL_TYPE']] + ' ' + row['name'], axis=1, result_type='reduce')
    else:
        df['name'] = 'МО'
        df['np'] = 'МО'

    
    if 'CLASSID_12' in df.columns:
        df['category_plan'] = df.apply(lambda row: row['category'] if not row['CLASSID_12'] else lu_dict[row['CLASSID_12']], axis=1, result_type='reduce')
        df['category_plan'] = df.apply(lambda row: lu_dict[702010500] if abs(row['area'] - row['area_du']) < 0.2 and not row['CLASSID_12'] and not row['iskl_v_np'] and row['DU_TYPE'] in ['искл_лес', 'лес'] else row['category_plan'], axis=1, result_type='reduce')
        # костыль для не заполненных DU_TYPE
        if not du_type:
            df['category_plan'] = df.apply(lambda row: lu_dict[702010500] if abs(row['area'] - row['area_du']) < 0.2 and not row['CLASSID_12'] and not row['iskl_v_np'] and row['cadnumber'] == 'Территория, собственность на которую не разграничена' else row['category_plan'], axis=1, result_type='reduce')
    else:
        df['category_plan'] = 'Земли населенных пунктов'
        
    df['zones']  = df.apply(lambda row: fz_dict[row['Ext_Zone_Code']], axis=1, result_type='reduce')

        
    def agg_strings(strings):
        return ', '.join(set(strings))
    df_agr = df.groupby(fields).agg({'zones': agg_strings, 'area': 'sum', 'area_du': 'sum'}).reset_index()
    df_agr['area'] = df_agr['area'].round(2)
    df_agr['area_du'] = df_agr['area_du'].round(3)
    df_agr = df_agr.loc[df_agr['area'] > 0].reset_index()
    df_agr['zones'] = df_agr.apply(lambda row: row['zones'].capitalize(), axis=1, result_type='reduce')
    del df_agr['index']
    df_agr = df_agr.rename(columns=dict(zip(df_agr.columns, ['НП для сортировки',
                                           'Населенный пункт', 
                                           'Кадастровый номер', 
                                           'Категория земель', 
                                           'Планируемая категория земель', 
                                           'Функциональные зоны', 
                                           'Площадь перевода, га',
                                           'Площадь пересечения с границами лесничества, га'])))
    return df_agr

def main():
    folder = arcpy.mp.ArcGISProject("CURRENT").homeFolder

    params = Params(arcpy.GetParameterInfo())

    with pd.ExcelFile(params.input_xls) as xls:
        df_np_types = pd.read_excel(xls, "Типы НП")
        np_dict = {i[1]["SETTL_TYPE"]: i[1]["NP_TYPE"] for i in df_np_types.iterrows()}
        df_lu = pd.read_excel(xls, "Категории земель")
        lu_dict = {i[1]["CLASSID"]: i[1]["CATEGORY"] for i in df_lu.iterrows()}
        df_fz = pd.read_excel(xls, "Функциональные зоны")
        fz_dict = {i[1]["Ext_Zone_Code"]: i[1]["Zone"] for i in df_fz.iterrows()}
        df_np = pd.read_excel(xls, "Населенные пункты")
        

    db_name = "temp_" + datetime.now(timezone.utc).replace(microsecond=0).astimezone().isoformat().replace(":", "") + ".gdb"
    output_db = os.path.join(folder, db_name)

    if arcpy.Exists(output_db):
        arcpy.management.Delete(output_db)
    arcpy.management.CreateFileGDB(folder, db_name)
    
    ###ЗА ГРАНИЦАМИ
    arcpy.env.addOutputsToMap = False

    # aprx = arcpy.mp.ArcGISProject("CURRENT")
    # m = aprx.activeMap

    # Выгружаем слои в локальную базу
    arcpy.conversion.FeatureClassToFeatureClass(params.NP, output_db, "NP_plan", "STATUS_ADM = 2")
    arcpy.conversion.FeatureClassToFeatureClass(params.NP, output_db, "NP_ex", "STATUS_ADM = 1")
    arcpy.conversion.FeatureClassToFeatureClass(params.ZU, output_db, "ZU")
    arcpy.conversion.FeatureClassToFeatureClass(params.FZ, output_db, "FZ")
    arcpy.conversion.FeatureClassToFeatureClass(params.LU, output_db, "LU_ex", "STATUS = 1 And (Note <> 'Двойной учет' Or Note IS NULL)")
    arcpy.conversion.FeatureClassToFeatureClass(params.LU, output_db, "LU_plan", "STATUS = 2")
    arcpy.conversion.FeatureClassToFeatureClass(params.DU, output_db, "DU_temp")#, "DU_TYPE IN ('После 2016/Нет информации', 'искл') And (Note <> 'Амнистия' Or Note IS NULL)")
    arcpy.management.Dissolve(output_db + "\\DU_temp", output_db + '\\DU', ['DU_TYPE'])

    # Прописываем названия слоев
    NP_plan = os.path.join(output_db, "NP_plan")
    NP_ex = os.path.join(output_db, "NP_ex")
    ZU = os.path.join(output_db, "ZU")
    FZ = os.path.join(output_db, "FZ")
    LU_ex = os.path.join(output_db, "LU_ex")
    LU_plan = os.path.join(output_db, "LU_plan")
    DU = os.path.join(output_db, "DU")

    # получаем слой с включениями
    arcpy.analysis.Erase(NP_plan, NP_ex, output_db + '\\NP_vkl', "")
    arcpy.analysis.PairwiseIntersect((NP_plan, NP_ex), output_db + "\\NP_plan_NP_ex", "all", "", "input")
    # arcpy.management.CalculateField(output_db + "\\NP_ex_NP_plan", "SETTL_TYPE_1", "!SETTL_TYPE_1! + 100", "PYTHON3")
    arcpy.management.Append([output_db + "\\NP_plan_NP_ex"], output_db + '\\NP_vkl', "NO_TEST", expression="NAME <> NAME_1")

    # получаем слой с исключениями
    arcpy.analysis.Erase(NP_ex, NP_plan, output_db + '\\NP_iskl', "")
    arcpy.analysis.PairwiseIntersect((NP_ex, NP_plan), output_db + "\\NP_ex_NP_plan", "all", "", "input")
    arcpy.management.CalculateField(output_db + "\\NP_ex_NP_plan", "SETTL_TYPE", "!SETTL_TYPE! + 100", "PYTHON3")
    arcpy.management.Append([output_db + "\\NP_ex_NP_plan"], output_db + '\\NP_iskl', "NO_TEST", expression="NAME <> NAME_1")


    # Общие операции
    arcpy.analysis.PairwiseIntersect((LU_ex, ZU), output_db + "\\LU_ex_ZU")

    arcpy.analysis.Erase(LU_ex, ZU, output_db + '\\LU_ex_Erase_ZU', "")

    arcpy.management.Merge([output_db + '\\LU_ex_Erase_ZU', output_db + "\\LU_ex_ZU"], output_db + "\\LU_all", "", "NO_SOURCE_INFO")

    arcpy.analysis.PairwiseIntersect((output_db + "\\LU_all", FZ), output_db + "\\LU_all_FZ")


    arcpy.analysis.PairwiseIntersect((output_db + "\\LU_all_FZ", LU_plan), output_db + "\\LU_all_FZ_LU_plan")


    # За границами
    arcpy.analysis.Erase(output_db + "\\LU_all_FZ_LU_plan", NP_ex, output_db + '\\LU_Erase_NP_ex', "")
    arcpy.analysis.Erase(output_db + '\\LU_Erase_NP_ex', NP_plan, output_db + '\\MO', "")
    arcpy.analysis.PairwiseIntersect((output_db + "\\MO", DU), output_db + "\\MO_DU")

    # Включаемые
    arcpy.analysis.PairwiseIntersect((output_db + "\\LU_all_FZ", output_db + '\\NP_vkl'), output_db + "\\VKL")
    arcpy.analysis.PairwiseIntersect((output_db + "\\VKL", DU), output_db + "\\VKL_DU")

    # Исключаемые
    arcpy.analysis.Erase(output_db + "\\LU_all_FZ", LU_plan, output_db + '\\LU_all_FZ_Erase_LU_plan', "")
    arcpy.management.Merge([output_db + '\\LU_all_FZ_Erase_LU_plan', output_db + "\\LU_all_FZ_LU_plan"], output_db + "\\LU_all_Merge_LU_plan", "", "NO_SOURCE_INFO")

    arcpy.analysis.PairwiseIntersect((output_db + "\\LU_all_Merge_LU_plan", output_db + '\\NP_iskl'), output_db + "\\ISKL")
    arcpy.analysis.PairwiseIntersect((output_db + "\\ISKL", DU), output_db + "\\ISKL_DU")

    mo = table_to_data_frame(output_db + "\\MO", ["cadnumber", "CLASSID", "cat", "CLASSID_12", "Ext_Zone_Code"])
    mo_du = table_to_data_frame(output_db + "\\MO_DU", ["FID_MO", "DU_TYPE"])
    result_mo = count_transfers(mo, mo_du, 'FID_MO', np_dict, lu_dict, fz_dict, params.du_type)
    
    vkl = table_to_data_frame(output_db + "\\VKL", ["cadnumber", "CLASSID", "cat", "Ext_Zone_Code", "name", "SETTL_TYPE"])
    vkl_du = table_to_data_frame(output_db + "\\VKL_DU", ["FID_VKL", "DU_TYPE"])
    result_vkl = count_transfers(vkl, vkl_du, 'FID_VKL', np_dict, lu_dict, fz_dict, params.du_type)
    
    iskl = table_to_data_frame(output_db + "\\ISKL", ["cadnumber", "CLASSID", "cat", "CLASSID_12", "Ext_Zone_Code", "name", "SETTL_TYPE"])
    iskl_du = table_to_data_frame(output_db + "\\ISKL_DU", ["FID_ISKL", "DU_TYPE"])
    result_iskl = count_transfers(iskl, iskl_du, 'FID_ISKL', np_dict, lu_dict, fz_dict, params.du_type)
    
    writer = pd.ExcelWriter(params.output_xls)
    result_vkl.to_excel(writer, engine='openpyxl', sheet_name='Включается', index=False)
    result_iskl.to_excel(writer, engine='openpyxl', sheet_name='Исключается', index=False)
    result_mo.to_excel(writer, engine='openpyxl', sheet_name='За границами', index=False)
    writer.close()

if __name__ == '__main__':
    main()