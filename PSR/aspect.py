
import arcpy, os, sys
import timeit
import psr_utility as utility
import psr_config as config
from numpy import gradient
from numpy import arctan2, arctan, sqrt
sys.path.insert(1,os.path.join(os.getcwd(),'DB_Framework'))
reload(sys)
import models
def calculate_aspect(order_obj,aspect_tif_pcs):
    order_geometry_pcs = order_obj.geometry.projectAs(order_obj.spatial_ref_pcs)
    centre_point_pcs = order_geometry_pcs.trueCentroid
    location = str(centre_point_pcs.X)+" "+str(centre_point_pcs.Y)
    aspect = arcpy.GetCellValue_management(aspect_tif_pcs,location)
    if aspect.getOutput((0)) != "NoData":
        aspect_text = utility.degree_direction_to_text(float(aspect.getOutput((0))))
        if float(aspect.getOutput((0))) == -1:
            aspect_text = r'N/A'
    else:
        print ('fail to use point XY to retrieve')
        aspect_text = '-9999'
        raise ValueError('No aspect retrieved CHECK data spatial reference')
    return aspect_text,centre_point_pcs.X, centre_point_pcs.Y
def generate_aspect_map(order_obj):
    arcpy.AddMessage('  -- Start generating aspect map ...')
    start = timeit.default_timer() 
    ### create buffer map based on order geometry
    arcpy.Buffer_analysis(config.order_geometry_pcs_shp, config.order_aspect_buffer, config.buffer_dist_aspect) 
    ### select dem itersected by order buffer
    master_dem_layer = arcpy.MakeFeatureLayer_management(config.master_lyr_dem, 'master_dem_layer') 
    arcpy.SelectLayerByLocation_management(master_dem_layer, 'intersect',  config.order_aspect_buffer)
    # arcpy.CopyFeatures_management(master_dem_layer, config.dem_selectedby_order)
    dem_rasters = []
    dem_raster = None
    if (int((arcpy.GetCount_management(master_dem_layer).getOutput(0)))== 0):
        arcpy.AddMessage('  -- NO records selected for US')
    else:
        dem_rows = arcpy.SearchCursor(master_dem_layer)
        for row in dem_rows:
            dem_raster_file = row.getValue("image_name")
            if(dem_raster_file != ''):
                dem_rasters.append(os.path.join(config.img_dir_dem,dem_raster_file))
        del row
        del dem_rows
    if len(dem_rasters) == 0:
        master_dem_layer = arcpy.mapping.Layer(config.master_lyr_dem)
        arcpy.SelectLayerByLocation_management(master_dem_layer, 'intersect',  config.order_aspect_buffer)
        if (int((arcpy.GetCount_management(master_dem_layer).getOutput(0)))!= 0):
            for row in dem_rows:
                dem_raster_file = row.getValue("image_name")
                if(dem_raster_file != ''):
                    dem_rasters.append(os.path.join(config.img_dir_dem,dem_raster_file))
    if len(dem_rasters) >= 1:
        dem_raster = "img.img"
        if len(dem_rasters) == 1:
            ras = dem_rasters[0]
            arcpy.Clip_management(ras, "#",os.path.join(config.scratch_folder,dem_raster),config.order_aspect_buffer,"#","NONE", "MAINTAIN_EXTENT")
        else:
            clipped_ras=''
            i = 1
            for ras in dem_rasters:
                clip_name ="clip_ras_"+str(i)+".img"
                arcpy.Clip_management(ras, "#",os.path.join(config.scratch_folder, clip_name),config.order_aspect_buffer,"#","NONE", "MAINTAIN_EXTENT")
                clipped_ras = clipped_ras + os.path.join(config.order_aspect_buffer, clip_name)+ ";"
                i+=1

            arcpy.MosaicToNewRaster_management(clipped_ras[0:-1], config.scratch_folder, dem_raster, order_obj.spatial_ref_pcs, "32_BIT_FLOAT", "#","1", "FIRST", "#")
                
        numpy_array =  arcpy.RasterToNumPyArray(os.path.join(config.scratch_folder, dem_raster))
        x,y = gradient(numpy_array)
        slope = 57.29578*arctan(sqrt(x*x + y*y))
        aspect = 57.29578*arctan2(-x,y)

        for i in range(len(aspect)):
            for j in range(len(aspect[i])):
                if -180 <=aspect[i][j] <= -90:
                    aspect[i][j] = -90-aspect[i][j]
                else :
                    aspect[i][j] = 270 - aspect[i][j]
                if slope[i][j] ==0:
                    aspect[i][j] = -1
        # gather some information on the original file
        cell_size_h  = arcpy.Describe(os.path.join(config.scratch_folder,dem_raster)).meanCellHeight
        cell_size_w  = arcpy.Describe(os.path.join(config.scratch_folder,dem_raster)).meanCellWidth
        extent     = arcpy.Describe(os.path.join(config.scratch_folder,dem_raster)).Extent
        pnt        = arcpy.Point(extent.XMin,extent.YMin)
        
        # save the raster
        aspect_tif = os.path.join(config.scratch_folder,"aspect.tif")
        aspect_ras = arcpy.NumPyArrayToRaster(aspect,pnt,cell_size_h,cell_size_w)
        arcpy.CopyRaster_management(aspect_ras,aspect_tif)
        arcpy.DefineProjection_management(aspect_tif, order_obj.spatial_ref_gcs)
        # slope map
        slope_tif = os.path.join(config.scratch_folder,"slope.tif")
        slope_ras = arcpy.NumPyArrayToRaster(slope,pnt,cell_size_h,cell_size_w)
        arcpy.CopyRaster_management(slope_ras,slope_tif)
        arcpy.DefineProjection_management(slope_tif, order_obj.spatial_ref_gcs)
        
        aspect_tif_pcs = os.path.join(config.scratch_folder,"aspect_pcs.tif")
        arcpy.ProjectRaster_management(aspect_tif,aspect_tif_pcs, order_obj.spatial_ref_pcs)
        
        #Calculate aspect
        aspect_text, utm_x, utm_y = calculate_aspect(order_obj,aspect_tif_pcs)
        site_elevation =  utility.get_elevation(order_obj.geometry.trueCentroid.X,order_obj.geometry.trueCentroid.Y)
        
        #Update DB
        psr_obj = models.PSR()
        psr_obj.update_order(order_obj.id,str(utm_x),str(utm_y),order_obj.spatial_ref_pcs.name, site_elevation, aspect_text)
                    
    end = timeit.default_timer()
    arcpy.AddMessage(('-- End generating PSR aspect map. Duration:', round(end -start,4)))
    