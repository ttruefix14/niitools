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

    arcpy.env.addOutputsToMap = False

    for ds in arcpy.ListDatasets():
        if params.inputSR:
            if arcpy.Describe(ds).spatialReference.exportToString() != params.inputSR.exportToString: 
                arcpy.management.DefineProjection(ds, params.inputSR)

        if params.regVer:
            arcpy.management.RegisterAsVersioned(ds)

        if params.enableEdit:
            arcpy.management.EnableEditorTracking(ds, "created_user", "created_date", "last_edited_user", "last_edited_date", "ADD_FIELDS")

        arcpy.AddMessage(f"Датасет {ds} успешно зарегистрирован!")

    if params.addGuid:
        arcpy.management.AddGlobalIDs(arcpy.ListDatasets())

if __name__ == '__main__':
    execute()
