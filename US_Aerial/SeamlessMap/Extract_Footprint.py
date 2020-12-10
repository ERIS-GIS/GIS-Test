import arcpy, os ,timeit
from arcpy.sa import *
from multiprocessing import Pool
import os.path
import logging
import numpy as np
import pandas as pd
arcpy.CheckOutExtension("Spatial")

def GetUTMFromRaster(inputRaster):
    r = arcpy.Raster(inputRaster)
    x_centre = r.extent.XMin + (r.extent.XMax - r.extent.XMin)/2 # center of inpout image (x)
    y_centre = r.extent.YMin + (r.extent.YMax - r.extent.YMin)/2 # center of inpout image (y)
    input_sr_pcs = arcpy.GetUTMFromLocation(x_centre,y_centre)
    return input_sr_pcs
def create_image_extent_FC(inputRaster,input_sr,tmpGDB):
    arcpy.AddMessage('Start creating image"s extent polygon...')
    starto = timeit.default_timer()
    if input_sr.linearUnitCode == 0: ### if input_sr is GCS find PCS for it
        input_sr_pcs = GetUTMFromRaster(inputRaster)
    else:
        input_sr_pcs = input_sr
    image_extent_FC_org = arcpy.CreateFeatureclass_management(tmpGDB,"image_extent_FC_org", "POLYGON", "", "DISABLED", "DISABLED", input_sr)
    cursor = arcpy.InsertCursor(image_extent_FC_org)
    point = arcpy.Point()
    coords = arcpy.Array()
    corners = ["lowerLeft", "lowerRight", "upperRight", "upperLeft"]
    feat = cursor.newRow() 
    r = arcpy.Raster(inputRaster) 
    for corner in corners:    
        point.X = getattr(r.extent, "%s" % corner).X
        point.Y = getattr(r.extent, "%s" % corner).Y
        coords.add(point)
    coords.add(coords.getObject(0))
    polygon = arcpy.Polygon(coords)
    feat.shape = polygon
    cursor.insertRow(feat)
    del feat
    del cursor 
    ### project the image extent polygon if the inpot sr is GCS
    if input_sr.linearUnitCode == 0:
        image_extent_FC = arcpy.Project_management(image_extent_FC_org, os.path.join(tmpGDB,'image_extent_FC'), input_sr_pcs)
    else:
        image_extent_FC = image_extent_FC_org
    endo = timeit.default_timer()
    arcpy.AddMessage(('End creating image"s extent polygon. Duration:', round(endo -starto,4)))
    return image_extent_FC

def create_excluded_polygon_FC(inputRaster,excluded_rows_OIDs,polygon_with_holes,input_sr,tmpGDB):
    if input_sr.linearUnitCode == 0: ### if input_sr is GCS find PCS for it
        input_sr_pcs = GetUTMFromRaster(inputRaster)
    else:
        input_sr_pcs = input_sr
    excluded_polygons_FC = arcpy.CreateFeatureclass_management(tmpGDB,"excluded_polygons_FC", "POLYGON", "", "DISABLED", "DISABLED", input_sr_pcs)
    ### find the main polygon (has image content)
    sql_clause = (None, 'ORDER BY Shape_Area DESC')
    main_oid = arcpy.da.SearchCursor(polygon_with_holes,('OBJECTID'),'gridcode = 1',None,False,sql_clause).next()[0]
    ### populate excluded polygon in FC
    for oid in  excluded_rows_OIDs:
        if oid+1 != main_oid:
            expression = arcpy.AddFieldDelimiters(polygon_with_holes, 'OBJECTID') + ' = ' + str(oid+1)
            donuts_geom = (arcpy.da.SearchCursor(polygon_with_holes,('SHAPE@'),expression).next())[0]
            cursor = arcpy.da.InsertCursor(excluded_polygons_FC, ['SHAPE@'])
            cursor.insertRow([donuts_geom.projectAs(input_sr_pcs)])
            del cursor
    return excluded_polygons_FC
def detect_excluded_OIDS(infc):
    arcpy.AddMessage('Start extracting excluded polygons OIDs...')
    startd = timeit.default_timer()
    threshold = 0.5
    array = arcpy.da.TableToNumPyArray(infc, ('OBJECTID','Shape_Area'), skip_nulls=True)
    mean = np.mean(array['Shape_Area'])
    std = np.std(array['Shape_Area'])
    z_score = (array['Shape_Area'] - mean)/std
    indices = [i for i,v in enumerate(z_score > threshold) if v]
    endd = timeit.default_timer()
    arcpy.AddMessage(('End extracting excluded polygons OIDs. Duration:', round(endd -startd,4)))
    return indices 
def simplify_polygon(inputFC):
    arcpy.AddMessage('Start simplifying polygon...')
    start6 = timeit.default_timer()
    arcpy.Generalize_edit(inputFC, '100 Meter')
    end6 = timeit.default_timer()
    arcpy.AddMessage(('End simplifying polygon. Duration:', round(end6 -start6,4)))
def get_Footprint(inputRaster):
    try:
        ws = arcpy.env.scratchFolder
        arcpy.env.workspace = ws
        srWGS84 = arcpy.SpatialReference('WGS 1984')
        tmpGDB =os.path.join(ws,r"temp.gdb")
        arcpy.env.overwriteOutput = True
        if not os.path.exists(tmpGDB):
            arcpy.CreateFileGDB_management(ws,r"temp.gdb")
        print(ws)
        ### set input and output paths
        resampleRaster = os.path.join(ws,'resampleRaster' + '.tif')
        bin_Raster = os.path.join(ws,'bin_Raster' + '.tif')
        polygon_with_holes= os.path.join(tmpGDB,'polygon_with_holes')
        footprint_FC = os.path.join(tmpGDB,'footprint_FC')
        
        input_raster_desc = arcpy.Describe(inputRaster)
        input_sr = input_raster_desc.spatialReference
        multi_band = 0
        if input_raster_desc.bandCount > 1:
            multi_band = 1
            Band1 = Raster(inputRaster + "/Band_1")
            #Save the raster
            arcpy.AddMessage("saving one band of input image")
            Band1.save(os.path.join(ws,"band_01.tif"))
            arcpy.AddMessage("saved.")
        ### create featureclass from input image extent
        image_extent_FC = create_image_extent_FC(inputRaster,input_sr,tmpGDB)
        ### resampling the input image beacuse of increasing performance
        arcpy.AddMessage('Start resampling the input raster...')
        start1 = timeit.default_timer()
        rasterProp = arcpy.GetRasterProperties_management(inputRaster, "CELLSIZEX")
        if input_sr.linearUnitCode == 0:
            res_size = float(rasterProp.getOutput(0)) * 10
        else:
            if float(rasterProp.getOutput(0)) < 4:
                res_size = 4
            else:
                res_size = float(rasterProp.getOutput(0))
        resampleRaster = arcpy.Resample_management(inputRaster,resampleRaster ,res_size, "Cubic")
        
        end1 = timeit.default_timer()
        arcpy.AddMessage(('End resampling the input raster. Duration:', round(end1 -start1,4)))

        arcpy.AddMessage('Start creating binary raster (Raster Calculator)...')
        start2 = timeit.default_timer()
        expression = 'Con(' + '"' + 'resampleRaster' + '.tif' + '"' + ' >= 10 , 1 , 0)'
        bin_Raster = arcpy.gp.RasterCalculator_sa(expression, bin_Raster)
        end2 = timeit.default_timer()
        arcpy.AddMessage(('End creating binary raster. Duration:', round(end2 -start2,4)))

        ### Convert binary raster to polygon
        arcpy.AddMessage('Start creating prime polygon from raster...')
        start3 = timeit.default_timer()
        polygon_with_holes =  arcpy.RasterToPolygon_conversion(in_raster= bin_Raster, out_polygon_features=polygon_with_holes, simplify="SIMPLIFY", raster_field="Value", create_multipart_features="SINGLE_OUTER_PART", max_vertices_per_feature="")
        end3 = timeit.default_timer()
        arcpy.AddMessage(('End creating polygon. Duration:', round(end3 -start3,4)))
        ### extract the polygons of black,white and gap area of input image 
        excluded_rows_OIDs = detect_excluded_OIDS(polygon_with_holes)
        if(len(excluded_rows_OIDs) > 0):
            excluded_polygons_FC = create_excluded_polygon_FC(inputRaster,excluded_rows_OIDs,polygon_with_holes,input_sr,tmpGDB)
            simplify_polygon(excluded_polygons_FC) 
            ### repair the geometry beacuse the extracted polygons have a lot of complexity
            arcpy.AddMessage('Start generating final footprint in coordinate system of WGS84...')
            start6 = timeit.default_timer()
            arcpy.RepairGeometry_management (excluded_polygons_FC) 
            footprint_FC = arcpy.Erase_analysis(image_extent_FC, excluded_polygons_FC, footprint_FC, cluster_tolerance="100 Meters")
            footprint_FC_wgs84 = arcpy.Project_management(footprint_FC, os.path.join(tmpGDB,'footprint_FC_wgs84'), srWGS84)
            end6 = timeit.default_timer()
            arcpy.AddMessage(('End  generating final footprint in coordinate system of WGS84. Duration:', round(end6 -start6,4)))
        else: ### footprint equal to envelope
            footprint_FC = image_extent_FC
    except:
        msgs = "ArcPy ERRORS:\n %s\n"%arcpy.GetMessages(2)
        arcpy.AddError(msgs)
        raise

if __name__ == '__main__':
    arcpy.AddMessage('Start footprint extraction process...')
    startTotal = timeit.default_timer()
    # image_Path = '//cabcvan1nas003/doqq/FULL/IL/1998/z16/il_1m_1998-z16-east - copy 2.sid'
    # image_Path = '//cabcvan1nas003/doqq/FULL/WA/1990/z10/wa_1m_1990_z10-6.sid'
    # image_Path = '//CABCVAN1NAS003/doqq/201x/AZ/GILA/2010/ortho_1-2_1n_s_az007_2010_1.sid'
    # image_Path = '//cabcvan1nas003/doqq/osa/CA/Los Angeles/2012/ortho_1-1_1n_s_ca037_2012_2.sid'
    image_Path  = '//CABCVAN1NAS003/doqq/FULL/CA/1994/z11/ca_1m_1994-z11centralsouth.sid'
    footprint_Polygon = get_Footprint(image_Path)
    endTotal= timeit.default_timer()
    arcpy.AddMessage(('End footprint extraction process. Total Duration:', round(endTotal -startTotal,4)))