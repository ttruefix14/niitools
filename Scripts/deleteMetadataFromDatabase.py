import arcpy
import os
from arcpy import metadata as md

# Update the following variables before running the script.
myWorkspace = arcpy.GetParameterAsText(0)
db_type = arcpy.GetParameterAsText(1) #Set this to either "SQL", "Oracle or Postgres" if your db has spatial views. If not you may set it to "".

def RemoveHistory(myWorkspace):
##Removes GP History for feature dataset stored feature classes, and feature classes in the File Geodatabase.
    arcpy.env.workspace = myWorkspace
    for fds in arcpy.ListDatasets('','feature') + ['']:
        arcpy.AddMessage(str(fds))
        data_path = os.path.join(myWorkspace, fds)
        removeMetaData(data_path)
        for fc in arcpy.ListFeatureClasses('','',fds):
            arcpy.AddMessage(str(fds) + " " + str(fc))
            data_path = os.path.join(myWorkspace, fds, fc)
            if isNotSpatialView(myWorkspace, fc):
                removeMetaData(data_path)
                print("Removed the geoprocessing metadata from: {0}".format(fc))
    removeMetaData(myWorkspace)
    print("Removed the geoprocessing metadata from: {0}".format(myWorkspace))

def isNotSpatialView(myWorkspace, fc):
##Determines if the item is a spatial view and if so returns True to listFcsInGDB()
    if db_type != "GDB":
        desc = arcpy.Describe(fc)
        fcName = desc.name
        #Connect to the GDB
        egdb_conn = arcpy.ArcSDESQLExecute(myWorkspace)
        #Execute SQL against the view table for the specified RDBMS
        if db_type == "SQL":
            db, schema, tableName = fcName.split(".")
            sql = r"IF EXISTS(select * FROM sys.views where name = '{0}') SELECT 1 ELSE SELECT 0".format(tableName)
        elif db_type == "Oracle":
            schema, tableName = fcName.split(".")
            sql = r"SELECT count(*) from dual where exists (select * from user_views where view_name = '{0}')".format(tableName)
        elif db_type == "Postgres":
            db, schema, tableName = fcName.split(".")
            sql = r"SELECT count(*) from information_schema.views where table_schema NOT IN ('information_schema', 'pg_catalog') and table_name = '{0}'".format(tableName)
        egdb_return = egdb_conn.execute(sql)
        if egdb_return == 0:
            return True
        else:
            return False
    else:
        return True

def removeMetaData(data_path):
    # Get the metadata for the dataset
    tgt_item_md = md.Metadata(data_path)
    # Delete all geoprocessing history from the item's metadata
    if not tgt_item_md.isReadOnly:
        tgt_item_md.deleteContent('GPHISTORY')
        tgt_item_md.deleteContent('THUMBNAIL')
        tgt_item_md.save()
    else:
        arcpy.AddMessage("is readonly")
        arcpy.AddMessage("---")

if __name__ == "__main__":
    RemoveHistory(myWorkspace)
    print("Done Done")