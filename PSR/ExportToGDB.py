import arcpy
import timeit
import logging
try:
    start = timeit.default_timer() 
    logfile = r'C:\Users\HKiavarz\Documents\log_wetland.txt'
    wetland_gdb_FC = '//cabcvan1gis006/GISData/Data/PSR/PSR.gdb/wetland_merged'
    # wetland_sde_FC = '//cabcvan1gis006/GISData/GIS_Test.sde/ERIS_GIS.WETLAND'
    wetland_sde_FC = '//cabcvan1gis006/GISData/GIS_Prod.sde/ERIS_GIS.WETLAND'
    
    logger = logging.getLogger(__name__)
    logger.setLevel(logging.WARNING)
    handler = logging.FileHandler(logfile)
    handler.setLevel(logging.WARNING)
    logger.addHandler(handler)
    
    fields = ["OBJECTID","ATTRIBUTE", 'WETLAND_TYPE','ACRES','ERIS_STATE','GLOBALID','SHAPE@']
    expression = "OBJECTID >= 7916162 AND OBJECTID <= 7916162"
    # expression = "OBJECTID = 7916161"
    wetland_gdb_rows = arcpy.da.SearchCursor(wetland_gdb_FC,fields,expression)
    cursor_insert = arcpy.da.InsertCursor(wetland_sde_FC, fields)
    OBJECTID = 0
    srWGS84 = arcpy.SpatialReference('WGS 1984')
    print(srWGS84.name)
    i = 0
    for row in wetland_gdb_rows:
        OBJECTID = row[0]
        count_vertices = row[6].pointCount-row[6].partCount
    #     # count_vertices = row[5].pointCount
        if count_vertices < 524287:
            if row[6] != None:
                geom_WGS84 = row[6].projectAs(srWGS84)
                row_values = [row[0],row[1],row[2],row[3],row[4],row[5],geom_WGS84]
                print("OBJECTID: %s - VCount: %s" % (row[0],count_vertices))
                cursor_insert.insertRow(row_values)
        else:
            logger.warning("OBJECTID: %s - VCount: %s" % (OBJECTID,count_vertices))
    end = timeit.default_timer()
    arcpy.AddMessage(('End PSR report process. Duration:', round(end -start,4)))
except:
        msgs = "ArcPy ERRORS %s:\n %s\n"% (str(OBJECTID),arcpy.GetMessages(2))
        arcpy.AddError(msgs)
        raise