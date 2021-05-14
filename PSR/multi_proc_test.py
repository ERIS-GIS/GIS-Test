from multiprocessing import Process, Queue, Pool, cpu_count, current_process, Manager
import arcpy
from imp import reload
import arcpy, os, sys
from datetime import datetime
import timeit,time
import shutil
import psr_utility as utility
import psr_config as config
file_path =os.path.dirname(os.path.abspath(__file__))
sys.path.insert(1,os.path.join(os.path.dirname(file_path),'DB_Framework'))
import models
reload(sys)
class Setting:
    grid_layer_mm = None
    df_mm_flood = None
    mxd_multi_flood = None
    output_jpg_flood = None
    order_obj = None

def simpleFunction(arg): 
    return arg*2 
def export_to_jpg(param_dic):
    page_num = param_dic[0]
    mxd_mm_file_flood = param_dic[1][0]
    output_jpg_flood = param_dic[1][1]
    mxd_multi_flood = arcpy.mapping.MapDocument(mxd_mm_file_flood)
    df_mm_flood = arcpy.mapping.ListDataFrames(mxd_multi_flood,"Flood*")[0]
    grid_layer_mm = arcpy.mapping.ListLayers(mxd_multi_flood,"Grid" ,df_mm_flood)[0]
    arcpy.SelectLayerByAttribute_management(grid_layer_mm, "NEW_SELECTION", ' "PageNumber" =  ' + str(page_num))
    df_mm_flood.extent = grid_layer_mm.getSelectedExtent(True)
    df_mm_flood.scale = df_mm_flood.scale * 1.1
    arcpy.SelectLayerByAttribute_management(grid_layer_mm, "CLEAR_SELECTION")

    title_text = arcpy.mapping.ListLayoutElements(mxd_multi_flood, "TEXT_ELEMENT", "title")[0]
    title_text.text = '      - Flood Hazard Zones - Page ' + str(page_num)
    title_text.elementPositionX = 0.5946
    arcpy.RefreshTOC()
    arcpy.mapping.ExportToJPEG(mxd_multi_flood, output_jpg_flood[0:-4]+str(page_num)+".jpg", "PAGE_LAYOUT", 480, 640, 75, "False", "24-BIT_TRUE_COLOR", 40)
    # shutil.copy(output_jpg_flood[0:-4]+str(page_num)+".jpg", os.path.join(config.report_path, 'PSRmaps', order_obj.number))
    return output_jpg_flood[0:-4]+str(page_num)+".jpg"
def execute_multi_task():
    list = [1, 2, 3, 4]
    arcpy.AddMessage(" Multiprocessing test...") 
    pool = Pool(processes=cpu_count())
    result = pool.map(simpleFunction, list)
    pool.close()
    pool.join()
    print(result)
    
def generate_multi_page_multi_processing(param_dic):
    arcpy.AddMessage(" Multiprocessing test...") 
    # for itm in param_dic:
    #     export_to_jpg(itm)
    pool = Pool(processes=cpu_count())
    # pool = Pool(processes=10)
    result = pool.map(export_to_jpg, param_dic)
    pool.close()
    pool.join()
    return  result
    
def generate_flood_report(order_obj):
    arcpy.AddMessage('  -- Start generating PSR flood report...')
    start = timeit.default_timer() 
    eris_id = 0
    Setting.output_jpg_flood = config.output_jpg(order_obj,config.Report_Type.flood)
    page = 1
    
    config.buffer_dist_flood = str(order_obj.psr.search_radius['10683']) + ' MILES'
    ### create buffer map based on order geometry
    arcpy.Buffer_analysis(config.order_geometry_pcs_shp, config.order_buffer_shp, config.buffer_dist_flood) 
    
    arcpy.MakeFeatureLayer_management(config.data_flood, 'flood_lyr') 
    arcpy.SelectLayerByLocation_management('flood_lyr', 'intersect',  config.order_buffer_shp)
    arcpy.CopyFeatures_management('flood_lyr', config.flood_selectedby_order_shp)
    
    arcpy.MakeFeatureLayer_management(config.data_flood_panel, 'flood_panel_lyr') 
    arcpy.SelectLayerByLocation_management('flood_panel_lyr', 'intersect',  config.order_buffer_shp)
    arcpy.CopyFeatures_management('flood_panel_lyr', config.flood_panel_selectedby_order_shp)
    
    arcpy.Statistics_analysis(config.flood_selectedby_order_shp, os.path.join(config.scratch_folder,"summary_flood.dbf"), [['FLD_ZONE','FIRST'], ['ZONE_SUBTY','FIRST']],'ERIS_CLASS')
    arcpy.Sort_management(os.path.join(config.scratch_folder,"summary_flood.dbf"), os.path.join(config.scratch_folder,"summary_sorted_flood.dbf"), [["ERIS_CLASS", "ASCENDING"]])
    
    mxd_flood = arcpy.mapping.MapDocument(config.mxd_file_flood)
    df_flood = arcpy.mapping.ListDataFrames(mxd_flood,"Flood*")[0]
    df_flood.spatialReference = order_obj.spatial_ref_pcs
    
    df_flood_small = arcpy.mapping.ListDataFrames(mxd_flood,"Study*")[0]
    df_flood_small.spatialReference = order_obj.spatial_ref_pcs
    del df_flood_small
    
    utility.add_layer_to_mxd("order_buffer",df_flood,config.buffer_lyr_file, 1.1)
    utility.add_layer_to_mxd("order_geometry_pcs", df_flood,config.order_geom_lyr_file,1)
    arcpy.RefreshActiveView()
    
    if not config.if_multi_page: # single-page
        mxd_flood.saveACopy(os.path.join(config.scratch_folder, "mxd_flood.mxd"))  
        arcpy.mapping.ExportToJPEG(mxd_flood, Setting.output_jpg_flood, "PAGE_LAYOUT", resolution=75, jpeg_quality=40)
        if not os.path.exists(os.path.join(config.report_path, 'PSRmaps', order_obj.number)):
            os.mkdir(os.path.join(config.report_path, 'PSRmaps', order_obj.number))
        arcpy.AddMessage('      - output jpg image path (overview map): %s' % os.path.join(config.report_path, 'PSRmaps', order_obj.number,os.path.basename(Setting.output_jpg_flood)))
        shutil.copy(Setting.output_jpg_flood, os.path.join(config.report_path, 'PSRmaps', order_obj.number))
        del mxd_flood
        del df_flood
    else: # multi-page
        grid_lyr_shp = os.path.join(config.scratch_folder, 'grid_lyr_flood.shp')
        arcpy.GridIndexFeatures_cartography(grid_lyr_shp, config.order_buffer_shp, "", "", "", config.grid_size, config.grid_size)
        
        # part 1: the overview map
        # add grid layer
        grid_layer = arcpy.mapping.Layer(config.grid_lyr_file)
        grid_layer.replaceDataSource(config.scratch_folder,"SHAPEFILE_WORKSPACE","grid_lyr_flood")
        arcpy.mapping.AddLayer(df_flood,grid_layer,"Top")

        df_flood.extent = grid_layer.getExtent()
        df_flood.scale = df_flood.scale * 1.1

        mxd_flood.saveACopy(os.path.join(config.scratch_folder, "mxd_flood.mxd"))
        arcpy.mapping.ExportToJPEG(mxd_flood, Setting.output_jpg_flood, "PAGE_LAYOUT", 480, 640, 75, "False", "24-BIT_TRUE_COLOR", 40)
        if not os.path.exists(os.path.join(config.report_path, 'PSRmaps', order_obj.number)):
            os.mkdir(os.path.join(config.report_path, 'PSRmaps', order_obj.number))
        shutil.copy(Setting.output_jpg_flood, os.path.join(config.report_path, 'PSRmaps', order_obj.number))
        arcpy.AddMessage('      - output jpg image page 1: %s' % os.path.join(config.report_path, 'PSRmaps', order_obj.number,os.path.basename(Setting.output_jpg_flood)))
        del mxd_flood
        del df_flood
        
        ### part 2: the data driven pages
        
        page = int(arcpy.GetCount_management(grid_lyr_shp).getOutput(0))  + page
        arcpy.AddMessage('  -- number of pages: %s' % str(page))
        Setting.mxd_multi_flood = arcpy.mapping.MapDocument(config.mxd_mm_file_flood)

        Setting.df_mm_flood = arcpy.mapping.ListDataFrames(Setting.mxd_multi_flood,"Flood*")[0]
        Setting.df_mm_flood.spatialReference = order_obj.spatial_ref_pcs
        utility.add_layer_to_mxd("order_buffer",Setting.df_mm_flood,config.buffer_lyr_file,1.1)
        utility.add_layer_to_mxd("order_geometry_pcs", Setting.df_mm_flood,config.order_geom_lyr_file,1)
        
        Setting.grid_layer_mm = arcpy.mapping.ListLayers(Setting.mxd_multi_flood,"Grid" ,Setting.df_mm_flood)[0]
        Setting.grid_layer_mm.replaceDataSource(config.scratch_folder,"SHAPEFILE_WORKSPACE","grid_lyr_flood")
        arcpy.CalculateAdjacentFields_cartography(grid_lyr_shp, "PageNumber")
        Setting.mxd_multi_flood.saveACopy(os.path.join(config.scratch_folder, "mxd_mm_flood.mxd"))
       
        if not os.path.exists(os.path.join(config.report_path, 'PSRmaps', order_obj.number)):
            os.mkdir(os.path.join(config.report_path, 'PSRmaps', order_obj.number))
        
        ### execute multi-page report bu multi-task processing
        # page = 10
        parameter_dic = {}
        path_list = [os.path.join(config.scratch_folder, "mxd_mm_flood.mxd"), Setting.output_jpg_flood]
        for i in range(1,page):
            parameter_dic[i] = path_list
            
        # execute_multi_task()
        result = generate_multi_page_multi_processing(parameter_dic.items())
