import arcpy
import os

arcpy.env.workspace = arcpy.GetParameterAsText(0)
arcpy.env.addOutputsToMap = False

for r, dirname, filenames in os.walk(arcpy.env.workspace):
    arcpy.env.workspace = r

    shapefiles = arcpy.ListFeatureClasses()

    for shapefile in shapefiles:
        layerName = arcpy.Describe(shapefile).baseName
        arcpy.AddMessage(layerName)
        arcpy.MakeFeatureLayer_management(shapefile, layerName)
        notHasFeatures = arcpy.GetCount_management(layerName).getOutput(0)
        arcpy.AddMessage(str(notHasFeatures))
        if int(notHasFeatures) == 0:
            arcpy.AddMessage("Deleting shapefile: " + shapefile)
            arcpy.management.Delete(shapefile)