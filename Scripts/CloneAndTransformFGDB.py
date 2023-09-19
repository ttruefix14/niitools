import arcpy, os, shutil

os.system("cls")

origFgdbLocation = r"\\Server\Share"
origFgdbNames = ["myFileGeodatabase"] # DON'T append the suffix ".gdb"
newFgdbLocation = r"\\OtherServer\OtherShare"
resolution = "0.00005 Meters"
tolerance = "0.0001 Meters"
newCrs = arcpy.SpatialReference(2056) # EPSG code of "CH1903+ LV95"
ignoreList = [] # Feature Classes and Tables to ignore

def CopyFcl(fcl, path, crs = None):
    arcpy.FeatureClassToFeatureClass_conversion(fcl, path, fcl)
    print "Copied feature class '" + fcl + "'"

    newFclPath = path + os.path.sep + fcl
    arcpy.DeleteFeatures_management(newFclPath)
    print "Truncated feature class '" + fcl + "'"

    if crs is not None:
        arcpy.DefineProjection_management(newFclPath, crs)
        print "Set CRS of feature class '" + fcl + "' to '" + crs.name + "'"

def CopyTable(table, path):
    arcpy.TableToTable_conversion(table, path, table)
    print "Copied table '" + table + "'"
    
    arcpy.DeleteRows_management(path + os.path.sep + table)
    print "Truncated table '" + table + "'"

def CopyFcl2(fcl, path, template, crs = None):
    desc = arcpy.Describe(fcl)
    shapeType = desc.shapeType.upper()
    
    arcpy.CreateFeatureclass_management(path, fcl, shapeType, template, "SAME_AS_TEMPLATE", "SAME_AS_TEMPLATE", crs)
    print "Copied feature class '" + fcl + "'"

def CopyTable2(table, path):
    arcpy.TableToTable_conversion(table, path, table)
    print "Copied table '" + table + "'"
    
    arcpy.DeleteRows_management(path + os.path.sep + table)
    print "Truncated table '" + table + "'"

for fgdbName in origFgdbNames:
    print "Start FGDB '" + fgdbName + "'"
    
    newFgdb = fgdbName + ".gdb"
    newFgdbPath = newFgdbLocation + os.path.sep + newFgdb
    origFgdbPath = origFgdbLocation + os.path.sep + fgdbName + ".gdb"

    if os.path.exists(newFgdbPath):
        # looks weird, but look here: http://stackoverflow.com/questions/16373747/permission-denied-doing-os-mkdird-after-running-shutil-rmtreed-in-python
        tempPath = newFgdbPath + "2"
        os.rename(newFgdbPath, tempPath)
        shutil.rmtree(tempPath)
        print "FGDB '" + newFgdbPath + "' deleted"
    
    arcpy.CreateFileGDB_management(newFgdbLocation, newFgdb)
    print "FGDB '" + newFgdbPath + "' created"

    arcpy.env.workspace = origFgdbPath
    arcpy.env.XYResolution = resolution
    arcpy.env.XYTolerance = tolerance  

    for fds in arcpy.ListDatasets():
        print "Start feature dataset '" + fds + "'"
        arcpy.CreateFeatureDataset_management(newFgdbPath, fds, newCrs)

        arcpy.env.workspace = os.path.join(origFgdbPath, fds)
        fcls = arcpy.ListFeatureClasses()
        tables = arcpy.ListTables()
        
        for fcl in fcls:
            if os.path.join(fds, fcl) in ignoreList:
                continue

            template = os.path.join(origFgdbPath, fds, fcl)
            CopyFcl2(fcl, os.path.join(newFgdbPath, fds), template, newCrs)
        for table in tables:
            if os.path.join(fds, table) in ignoreList:
                continue
            
            CopyTable2(fcl, os.path.join(newFgdbPath, fds))

        arcpy.env.workspace = origFgdbPath
        
        print "Finish feature dataset '" + fds + "'"

    print "Start copying root feature classes"
    arcpy.env.workspace = origFgdbPath
    fcls = arcpy.ListFeatureClasses()
    for fcl in fcls:
        if fcl in ignoreList:
            continue

        template = os.path.join(origFgdbPath, fcl)
        CopyFcl2(fcl, newFgdbPath, template, newCrs)
        
    print "Finish copying root feature classes"
    
    print "Start copying root tables"
    tables = arcpy.ListTables()
    for table in tables:
        if table in ignoreList:
            continue
        
        CopyTable2(table, newFgdbPath)
        
    print "Finish copying root tables"

    print "Finish FGDB '" + fgdbName + "'"

print "Finish all"