import arcpy, os
from ast import literal_eval
import numpy as np

def createGeometry(pntCoords,geometry_type):
    geometry = arcpy.Geometry()
    spatialRef = arcpy.SpatialReference(4326)
    if geometry_type.lower()== 'point':
        geometry = arcpy.Multipoint(arcpy.Array([arcpy.Point(*coords) for coords in pntCoords]),spatialRef)
    elif geometry_type.lower() =='polyline':        
        geometry = arcpy.Polyline(arcpy.Array([arcpy.Point(*coords) for coords in pntCoords]),spatialRef)
    elif geometry_type.lower() =='polygon':
        geometry = arcpy.Polygon(arcpy.Array([arcpy.Point(*coords) for coords in pntCoords]),spatialRef)
    return geometry
try:
    orderFC = "//cabcvan1gis006/GISData/GIS_Test.sde/ERIS.ERIS_ORDER_GEOMETRY"
    order_polygon_fc = '//cabcvan1gis006/GISData/GIS_Test.sde/ERIS_GIS.Order_Geometry_Polygon'
    # orderFC = '//cabcvan1gis006/GISData/ERIS_TEST.sde/ERIS.ERIS_ORDER_GEOMETRY'
    where = "GEO_SPATIAL IS NULL and GEOMETRY_TYPE = 'POLYGON' and order_id = '354268'"
    spatialRef = arcpy.SpatialReference(4326)
   
    # i = 0
    # row = arcpy.da.SearchCursor(orderFC,("ORDER_ID","GEOMETRY_TYPE","GEOMETRY"),where).next()
    # coord_string = ((row[2])[1:-1])
    # coordinates = np.array(literal_eval(coord_string))
    # geometry_polygon = arcpy.Polygon(arcpy.Array([arcpy.Point(*coords) for coords in coordinates]),spatialRef)
    
    # fields = ['ORDER_ID','SHAPE@']
    # insert_cursor = arcpy.da.InsertCursor(order_polygon_fc, fields)
    # row_values = [int(row[0]),geometry_polygon]
    # desc = arcpy.Describe(order_polygon_fc)
    # print("Shape Type :   " + desc.shapeType)
    # print(geometry_polygon.area)
    # insert_cursor.insertRow(row_values)
    
except:
    msgs = "ArcPy ERRORS %s:\n" % (arcpy.GetMessages(1))
    arcpy.AddError(msgs)
    raise