import arcpy

aprx = arcpy.mp.ArcGISProject("CURRENT")
aprx.updateConnectionProperties(arcpy.GetParameterAsText(0), arcpy.GetParameterAsText(1))