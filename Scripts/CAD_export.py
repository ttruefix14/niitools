import arcpy
import os
import pandas as pd
import ezdxf
import colorsys

def hsv2rgb(h,s,v):
    return tuple(round(i * 255) for i in colorsys.hsv_to_rgb(h/360,s/100,v/100))

def execute():
    
    input_layer = arcpy.GetParameter(0)
    cols = [arcpy.GetParameterAsText(1), arcpy.GetParameterAsText(2)]
    cemetery = arcpy.GetParameterAsText(3)
    output_file = arcpy.GetParameterAsText(4)

    arcpy.env.addOutputsToMap = False

    

    unique_set = set()
    with arcpy.da.SearchCursor(input_layer, cols) as cursor:
        for row in cursor:
            unique_set.add(row)
    unique = sorted(list(unique_set), key=lambda x: x[1])
    unique = sorted(unique)
    sql_dict = {f'{zone}_{status}': "{} = {} AND {} = {}".format(cols[0], f"'{zone}'", cols[1], status) for zone, status in unique}
    
    colors_dict = {}
    cym = input_layer.getDefinition('V2').renderer

    cym_fields = [cym.fields.index(cols[0]), cym.fields.index(cols[1])]
    for g in cym.groups:
        for item in g.classes:
            values = item.values[0].fieldValues
            for l in item.symbol.symbol.symbolLayers:
                if not l.enable:
                    continue
                if 'CIMSolidFill' in repr(l):
                    if l.color.values[:-1] != [0, 0, 0]:
                        if 'CIMHSV' in repr(l.color):
                            colors_dict[values[cym_fields[0]] + '_' + values[cym_fields[1]]] = hsv2rgb(*l.color.values[:-1])
                        else:
                            colors_dict[values[cym_fields[0]] + '_' + values[cym_fields[1]]] = l.color.values[:-1]

    layer = arcpy.conversion.FeatureClassToFeatureClass(input_layer, 'memory', 'temp_layer').getOutput(0)
    if 'layer' in [field.name for field in arcpy.ListFields(layer)]:
        arcpy.management.DeleteField(layer, ['layer'])
    arcpy.management.AddField(layer, 'layer', 'TEXT')
    arcpy.management.CalculateField(layer, 'layer', f'str(!{cols[0]}!) + "_" + str(!{cols[1]}!)', 'PYTHON3')
    
    if os.path.isfile(output_file):
        os.remove(output_file)

    arcpy.conversion.ExportCAD(layer, 'DXF_R2018', output_file, 'Use_Filenames_in_Tables', 'Overwrite_Existing_Files')
    arcpy.management.Delete(layer)

    dxf = ezdxf.readfile(output_file)
    msp = dxf.modelspace()

    for layer in dxf.layers:
        if layer.dxf.name in sql_dict:
            lines = msp.query(f'LWPOLYLINE[layer=="{layer.dxf.name}"]')
            hatch = msp.add_hatch(dxfattribs={"layer": layer.dxf.name, "color": 0})
            hatch.rgb = colors_dict[layer.dxf.name]
            if layer.dxf.name[-1] == '2': 
                hatch2 = msp.add_hatch(dxfattribs={"layer": layer.dxf.name, "color": 0})
                hatch2.set_pattern_fill(name='ANSI31', color=0, scale=150, angle=0, double=0, style=1, pattern_type=1, definition=None)
                hatch2.rgb = (0, 0, 0)
            if layer.dxf.name[:-2] == cemetery: 
                hatch3 = msp.add_hatch(dxfattribs={"layer": layer.dxf.name, "color": 0})
                hatch3.set_pattern_fill(name='CROSS', color=0, scale=100, angle=0, double=0, style=1, pattern_type=1, definition=None)
                hatch3.rgb = (0, 0, 0)
            for line in lines:
                hatch.paths.add_polyline_path([i for i in line.vertices()])
                if layer.dxf.name[-1] == '2': 
                    hatch2.paths.add_polyline_path([i for i in line.vertices()])
                if layer.dxf.name[:-2] == cemetery:
                    hatch3.paths.add_polyline_path([i for i in line.vertices()])


    lines = msp.query(f'LWPOLYLINE')
    for line in lines.entities:
        line.rgb = (0, 0, 0)
    msp.entity_space.entities = sorted(msp.entity_space.entities, key=lambda entity: ['HATCH', 'LWPOLYLINE'].index(entity.dxf.dxftype))


    dxf.saveas(output_file)
    arcpy.env.addOutputsToMap = True

if __name__ == '__main__':
    execute()