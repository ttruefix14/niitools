import arcpy
import pandas as pd
import re

arcpy.SetLogMetadata(False)

class Params:
    """Входные параметры инструмента"""
    def __init__(self, params):
        self.p10 = params[0].valueAsText

        rename_cols = [i.split('=') for i in params[1].values] if params[1].value else None
        self.rename_cols = {i[0]: i[1] for i in rename_cols} if rename_cols else None

        self.output_xl = params[2].valueAsText

        
class Errors:
    """Класс объектов, содержащий таблицу с найденными ошибками"""
    columns = ['Слой', 'ObjectID', 'Атрибут', 'Словарь', 'Ошибка', 'Значение', 'Тип поля (обязательное, условное, необязательное)', 'Недопустимые символы']
    errors = []
    gid_dict = dict()
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

    
    def append(self, error):
        if error:
            self.errors.append(error)
        else:
            return
        
    def is_p10_layer(self, r_name, b_name):
        if r_name not in self.p10.keys():
            self.append([b_name, None, None, None, 'Название слоя не соответствует десятому приказу', None, None, None]) 
            return False
        else:
            return True
        
    def update_gid(self, df, b_name):
        current_gid = df['GLOBALID'].tolist()
        for gid in current_gid:
            self.gid_dict.setdefault(gid, []).append(b_name)
            
    def get_double_gid(self):
        filt_dict = dict(filter(lambda x: len(x[1]) > 1, self.gid_dict.items()))
        filt_dict = {key: ', '.join(value) for key, value in filt_dict.items()}
        return filt_dict.items() if filt_dict.items() else [[None, None]]

    def check_columns(self, df, r_name, b_name, rename_cols):
        
        if rename_cols:
            df = df.rename(columns=rename_cols)
        df.columns = df.columns.str.replace('_$', '', regex=True).str.replace('.*\.', '', regex=True).str.upper()
        required_columns = {col for col in self.p10[r_name].keys()}
        rename_columns = dict()
        for col in required_columns:
            if col not in df.columns:
                for col2 in df.columns:
                    if col2 in col:
                        rename_columns[col2] = col
        df = df.rename(columns=rename_columns)
        missing_columns = set(required_columns) - set(df.columns.to_list())
        check_columns = required_columns - missing_columns

        if missing_columns:
            self.append([b_name, None, ', '.join(missing_columns), None, 'Столбцы отсутствуют в слое', '', 'О', None])
        else:
            pass
        
        return df, check_columns
            
    def check_classid(self, cid, row_index, r_name, b_name):
        if cid not in self.p10[r_name]['CLASSID'][2]:
            self.errors.append([b_name, row_index, 'CLASSID', None, 'Код объекта не соответствует десятому приказу', cid, 'О', None])
        else:
            pass
        
    def check_oktmo(self, oktmo, row_index, r_name, b_name):
        if isinstance(oktmo, str):
            if oktmo.isdigit():
                return
        else:
            self.append([b_name, row_index, 'OKTMO', 'OKTMO', 'Код ОКТМО заполнен неверно', repr(oktmo)[1:-1] if oktmo else None, 'О', None])
    
    def is_correct_uuid(self, gid, row_index, r_name, b_name):
        if not re.search(r'[a-fA-F0-9]{8}-[a-fA-F0-9]{4}-[a-fA-F0-9]{4}-[a-fA-F0-9]{4}-[a-fA-F0-9]{12}', gid):
            self.append([b_name, row_index, 'GLOBALID', None, 'Неверный формат GLOBALID', gid, 'О', None])
        else:
            pass
        
    def is_date(self, date, row_index, col, col_required, r_name, b_name):
        try:
            date.strftime('%Y/%m/%d')
        except:
            self.append([b_name, row_index, col, None, 'Неверный формат даты', date, col_required, None])
        return
        
    def good_str(self, string):
        if string == None or string.strip() == '':
            return False, "Пустая строка", ''
        string = str(string)
        bad = re.findall(r'[&\n\t\r\'<>\u0008\x02]+', string)
        bad = re.findall(r'[\x00-\x08\x0b\x0c\x0e-\x1F\uD800-\uDFFF\uFFFE\uFFFF]+', string)
        if bad:
            return False, "Недопустимые символы в строке (Знак табуляции \\t, знак абзаца \\n, знак возврата каретки \\r, одинарные кавычки \', знаки < >, &, x08(бекспейс в юникоде))", bad
        else:
            return True, 'Все верно', ''
        
    def check_str(self, string, row_index, col, col_required, r_name, b_name):
        str_check = self.good_str(string)
        if str_check[0]:
            return
        else:
            if col_required == 'Н' and str_check[1] == 'Пустая строка':
                return
            else:
                self.append([b_name, row_index, col, None, str_check[1], repr(string)[1:-1] if string else None, col_required, str_check[2]])
                return


def tab_to_df(in_table, input_fields=None, where_clause=None):
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
                final_fields = [OIDFieldName] + ['SHAPE@'] + input_fields
            else:
                final_fields = [OIDFieldName] + input_fields
        else:
            final_fields = [field.name.upper() if field.name != shapeFieldName else 'SHAPE@' for field in arcpy.Describe(in_table).fields]
        data = [row for row in arcpy.da.SearchCursor(in_table, final_fields, where_clause=where_clause)]
        fc_dataframe = pd.DataFrame(data, columns=final_fields)
        fc_dataframe = fc_dataframe.set_index(OIDFieldName, drop=True)
        return fc_dataframe

def execute():
    params = Params(arcpy.GetParameterInfo())
    errors = Errors(params.p10)

    aprx = arcpy.mp.ArcGISProject('CURRENT')
    m = aprx.activeMap
    layers = m.listLayers()


    for layer in layers:
        arcpy.AddMessage(layer.name)

        df = tab_to_df(layer)
        if len(df) == 0:
            continue
            
        # Проверка соответствия названия слоя десятому приказу
        b_name = arcpy.Describe(layer).baseName
        r_name = b_name.split('.')[-1].split('_')[0]
        
        if not errors.is_p10_layer(r_name, b_name):
            continue
        
        # Проверка соответствия названий столбцов десятому приказу
        df, check_columns = errors.check_columns(df, r_name, b_name, params.rename_cols)
        
        if 'GLOBALID' in df:
            errors.update_gid(df, b_name)
        
        for row in range(len(df)):
            row_index = df.index[row]
            # Задаем переменные
            for col in df:
                vars()[col] = df[col].iloc[row]
            
            if 'CLASSID' in df.columns:
                errors.check_classid(vars()['CLASSID'], row_index, r_name, b_name)
            
            if 'GLOBALID' in df.columns:
                errors.is_correct_uuid(vars()['GLOBALID'], row_index, r_name, b_name)
            
            for col in check_columns:
                # Проверяем есть ли столбец в слое
                if col in ['CLASSID', 'GLOBALID']:
                    continue
                    
                if col == 'OKTMO':
                    errors.check_oktmo(vars()['OKTMO'], row_index, r_name, b_name)
                    continue
                    
                col_type =  errors.p10[r_name][col][1]
                col_required = errors.p10[r_name][col][0]
                
                if col in ['DATE_START', 'DATE_CLOSE']:
                    errors.is_date(vars()[col], row_index, col, col_required, r_name, b_name)
                    continue
                        
                if col_type == 'Символьное':
                    errors.check_str(vars()[col], row_index, col, col_required, r_name, b_name)
                    continue
                    
                # Проверяем является ли столбец обязательным
                if col_required not in ['О', 'У']:
                    continue
                else:
                    pass
                
                # Проверяем есть ли условие
                if not eval(errors.p10[r_name][col][3]):
                    continue
                else:
                    pass
                # Проверяем используется ли словарь для столбца
                
                domain = errors.p10[r_name][col][2]
                if domain:
                    # Проверяем есть ли значение в словаре
                    if vars()[col] in domain:
                        continue
                    else:
                        errors.append([b_name, row_index, col, errors.p10[r_name][col][4], 'Значение не входит в словарь', vars()[col], col_required, None])
                        continue
                # Словаря нет, проверяем обычным способом
                else:
                    pass
                # Проверяем по типу данных
                if col_type == 'Целое':
                    if vars()[col] == None or vars()[col] == 0:
                        errors.append([b_name, row_index, col, None, 'Поле должно быть заполнено целым числом', vars()[col], col_required, None])
                    continue
                elif col_type == 'Вещественное':
                    if vars()[col] == None or vars()[col] == 0:
                        errors.append([b_name, row_index, col, None, 'Поле должно быть заполнено вещественным числом', vars()[col], col_required, None])
                    continue          
                else:
                    errors.append([b_name, row_index, col, None, f'Неизвестный тип поля {col_type}', vars()[col], col_required, None])

    df = pd.DataFrame(errors.errors, columns = errors.columns)
    gid_df = pd.DataFrame(errors.get_double_gid(), columns = ['GLOBALID', 'Слои'])
    writer = pd.ExcelWriter(params.output_xl)

    df.to_excel(writer, engine='openpyxl', sheet_name='Ошибки', index=False)
    gid_df.to_excel(writer, engine='openpyxl', sheet_name='Повторные GLOBALID', index=False)

    writer.close()

if __name__ == '__main__':
    execute()
