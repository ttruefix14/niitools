import arcpy

class Params:
    def __init__(self, params):
        self.sde = params[0].valueAsText
        self.inputSR = params[1].value
        self.regVer = params[2].value
        self.enableEdit = params[3].value
        self.addGuid = params[4].value

def execute():
    params = Params(arcpy.GetParameterInfo())
    arcpy.env.workspace = params.sde
    inputSR = params.inputSR

    arcpy.env.addOutputsToMap = False

    if inputSR:
        for ds in arcpy.ListDatasets():
            if arcpy.Describe(ds).spatialReference.WKT != inputSR.WKT: 
                arcpy.management.DefineProjection(ds, inputSR)


    for ds in arcpy.ListDatasets():
        if params.regVer:
            arcpy.management.RegisterAsVersioned(ds)

        if params.enableEdit:
            arcpy.management.EnableEditorTracking(ds)

    if params.addGuid:
        arcpy.management.AddGlobalIDs(arcpy.ListDatasets())

if __name__ == '__main__':
    execute()
