import arcpy
def execute():
    inFc = arcpy.GetParameter(0) or "23_02_0804000_2023-09-20"
    cad_field = arcpy.GetParameterAsText(1)
    part_field = arcpy.GetParameterAsText(2)
    fields = {field.name: field for field in arcpy.ListFields(inFc)}
    if not cad_field:
        cad_field = 'CadNumber'
        cad_field_n = 0
        while cad_field + str(cad_field_n or "") in fields:
            cad_field_n += 1
        cad_field = cad_field + str(cad_field_n or "")
        arcpy.management.AddField(inFc, cad_field, 'TEXT')
        
    if not part_field:
        part_field = 'part'
        part_field_n = 0
        while part_field + str(part_field_n or "") in fields:
            part_field_n += 1
        part_field = part_field + str(part_field_n or "")
        arcpy.management.AddField(inFc, part_field, 'TEXT')
        
    type_expression = 'get_type(!TypeObjRus!, !cat!, !toks!, !type!, !typez!)'
    type_expression_code = """
def get_type(current_type, cat, toks, type_, typez):
    if current_type != 'Контур':
        return current_type
    if cat != ' ':
        return 'Земельный участок'
    if toks == 'Объект незавершенного строительства':
        return 'Незавершенное строительство'
    elif toks != ' ':
        return toks
    if type_ in ['Граница муниципального образования', 'Граница населенного пункта']:
        return 'Границы субьекта, МО, населенного пункта'
    elif type_ != ' ':
        return 'Зона территориальная или иная'
    return current_type
"""

    arcpy.management.CalculateField(inFc, 'TypeObjRus', type_expression, code_block=type_expression_code)

    cad_expression = "!CadastralN! if '(' not in !CadastralN! else !CadastralN![:str.find(!CadastralN!, '(')]"
    arcpy.management.CalculateField(inFc, cad_field, cad_expression)

    part_expression = "'' if '(' not in !CadastralN! else !CadastralN![str.find(!CadastralN!, '(')+1:str.find(!CadastralN!, ')')]"
    arcpy.management.CalculateField(inFc, part_field, part_expression)

if __name__ == '__main__':
    execute()