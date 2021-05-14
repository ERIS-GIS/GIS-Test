
from imp import reload
import arcpy, os, sys
from datetime import datetime
import timeit,time
import shutil
import psr_utility as utility
import psr_config as config
sys.path.insert(1,os.path.join(os.getcwd(),'DB_Framework'))
import models
reload(sys)

def generate_singlepage_report(order_obj,mxd_wetland,outputjpg_wetland,scratch_folder):
    mxd_wetland.saveACopy(os.path.join(scratch_folder, "mxd_wetland.mxd"))
    arcpy.mapping.ExportToJPEG(mxd_wetland, outputjpg_wetland, "PAGE_LAYOUT", 480, 640, 150, "False", "24-BIT_TRUE_COLOR", 85)
    if not os.path.exists(os.path.join(config.report_path, 'PSRmaps', str(order_obj.number))):
        os.mkdir(os.path.join(config.report_path, 'PSRmaps', str(order_obj.number)))
    shutil.copy(outputjpg_wetland, os.path.join(config.report_path, 'PSRmaps', str(order_obj.number)))
    arcpy.AddMessage('      - Wetland Output: %s' % os.path.join(config.report_path, 'PSRmaps', str(order_obj.number)))
    del mxd_wetland
def generate_multipage_report(order_obj,mxd_wetland,output_jpg_wetland,df_wetland):
    grid_lyr_shp = os.path.join(config.scratch_folder, 'grid_lyr_wetland.shp')
    arcpy.GridIndexFeatures_cartography(grid_lyr_shp, config.order_buffer_shp, "", "", "", config.grid_size, config.grid_size)

    # part 1: the overview map
    # add grid layer
    grid_layer = arcpy.mapping.Layer(config.grid_lyr_file)
    grid_layer.replaceDataSource(config.scratch_folder,"SHAPEFILE_WORKSPACE","grid_lyr_wetland")
    arcpy.mapping.AddLayer(df_wetland,grid_layer,"Top")
    
    df_wetland.extent = grid_layer.getExtent()
    df_wetland.scale = df_wetland.scale * 1.1
    
    mxd_wetland.saveACopy(os.path.join(config.scratch_folder, "mxd_wetland.mxd"))
    arcpy.mapping.ExportToJPEG(mxd_wetland, output_jpg_wetland, "PAGE_LAYOUT", 480, 640, 150, "False", "24-BIT_TRUE_COLOR", 85)
    
    
    if not os.path.exists(os.path.join(config.report_path, 'PSRmaps', order_obj.number)):
        os.mkdir(os.path.join(config.report_path, 'PSRmaps', order_obj.number))
    shutil.copy(output_jpg_wetland, os.path.join(config.report_path, 'PSRmaps', order_obj.number))
    arcpy.AddMessage('      - Wetland Output: %s' % os.path.join(config.report_path, 'PSRmaps', str(order_obj.number)))
    del mxd_wetland
    del df_wetland

    ### part 2: the data driven pages
    page = int(arcpy.GetCount_management(grid_lyr_shp).getOutput(0))  + 1
    mxd_mm_wetland = arcpy.mapping.MapDocument(config.mxd_mm_file_wetland)
    df_mm_wetland = arcpy.mapping.ListDataFrames(mxd_mm_wetland,"big")[0]
    df_mm_wetland.spatialReference = order_obj.spatial_ref_pcs
    
    utility.add_layer_to_mxd("order_buffer",df_mm_wetland,config.buffer_lyr_file, 1.1)
    utility.add_layer_to_mxd("order_geometry_pcs", df_mm_wetland,config.order_geom_lyr_file,1)
    
    grid_layer_mm= arcpy.mapping.ListLayers(mxd_mm_wetland,"Grid" ,df_mm_wetland)[0]
    grid_layer_mm.replaceDataSource(config.scratch_folder, "SHAPEFILE_WORKSPACE","grid_lyr_wetland")
    arcpy.CalculateAdjacentFields_cartography(grid_lyr_shp, "PageNumber")
    mxd_mm_wetland.saveACopy(os.path.join(config.scratch_folder, "grid_layer_mm.mxd"))

    for i in range(1,int(arcpy.GetCount_management(grid_lyr_shp).getOutput(0))+1):
        arcpy.SelectLayerByAttribute_management(grid_layer_mm, "NEW_SELECTION", ' "PageNumber" =  ' + str(i))
        df_mm_wetland.extent = grid_layer_mm.getSelectedExtent(True)
        df_mm_wetland.scale = df_mm_wetland.scale * 1.1
        arcpy.SelectLayerByAttribute_management(grid_layer_mm, "CLEAR_SELECTION")

        titleTextE = arcpy.mapping.ListLayoutElements(mxd_mm_wetland, "TEXT_ELEMENT", "title")[0]
        titleTextE.text = "Wetland Type - Page " + str(i)
        titleTextE.elementPositionX = 0.468
        arcpy.RefreshTOC()

        arcpy.mapping.ExportToJPEG(mxd_mm_wetland, output_jpg_wetland[0:-4]+str(i)+".jpg", "PAGE_LAYOUT", 480, 640, 150, "False", "24-BIT_TRUE_COLOR", 85)
        if not os.path.exists(os.path.join(config.report_path, 'PSRmaps', order_obj.number)):
            os.mkdir(os.path.join(config.report_path, 'PSRmaps', order_obj.number))
        shutil.copy(output_jpg_wetland[0:-4]+str(i)+".jpg", os.path.join(config.report_path, 'PSRmaps', order_obj.number))
    del mxd_mm_wetland
    del df_mm_wetland
    
def generate_wetland_report(order_obj):
    arcpy.AddMessage('  -- Start generating PSR wetland report...')
    start = timeit.default_timer()   
    ### set scratch folder
    arcpy.env.workspace = config.scratch_folder
    arcpy.env.overwriteOutput = True   
    
    output_jpg_wetland = config.output_jpg(order_obj,config.Report_Type.wetland)
    output_jpg_ny_wetland = os.path.join(config.scratch_folder, str(order_obj.number)+'_NY_WETL.jpg')

    ### Wetland Map
    if '10684' not in order_obj.psr.search_radius.keys():
        arcpy.AddMessage('      -- Wetland search radius is not availabe')
        return
    config.buffer_dist_wetland = str(order_obj.psr.search_radius['10684']) + ' MILES'
    arcpy.Buffer_analysis(config.order_geometry_pcs_shp, config.order_buffer_shp,  config.buffer_dist_wetland)
    mxd_wetland = arcpy.mapping.MapDocument(config.mxd_file_wetland)
    df_wetland = arcpy.mapping.ListDataFrames(mxd_wetland,"big")[0]
    df_wetland.spatialReference = order_obj.spatial_ref_pcs
    df_wetland_small = arcpy.mapping.ListDataFrames(mxd_wetland,"small")[0]
    df_wetland_small.spatialReference = order_obj.spatial_ref_pcs
    del df_wetland_small
    ### add order and order_buffer layers to wetland mxd file
    utility.add_layer_to_mxd("order_buffer",df_wetland,config.buffer_lyr_file, 1.1)
    utility.add_layer_to_mxd("order_geometry_pcs", df_wetland,config.order_geom_lyr_file,1)
   
    # print the maps
    if not config.if_multi_page:  # sinle page report
        generate_singlepage_report(order_obj,mxd_wetland,output_jpg_wetland,config.scratch_folder)
    else:                           # multipage report
        generate_multipage_report(order_obj,mxd_wetland,output_jpg_wetland, df_wetland)
   ### Create wetland report for Newyork province
    if order_obj.province =='NY':
        arcpy.AddMessage ("      - Starting NY wetland section: " + time.strftime('%Y-%m-%d %H:%M:%S', time.localtime()))
        buffer_wetland_shp = os.path.join(config.scratch_folder,"buffer_wetland.shp")
        mxd_wetland_ny = arcpy.mapping.MapDocument(config.mxd_file_wetlandNY)
        df_wetland_ny = arcpy.mapping.ListDataFrames(mxd_wetland_ny,"big")[0]
        df_wetland_ny.spatialReference = order_obj.spatial_ref_pcs
        ### add order and order_buffer layers to wetland Newyork mxd file
        utility.add_layer_to_mxd("order_buffer",df_wetland_ny,config.buffer_lyr_file,1.1)
        utility.add_layer_to_mxd("order_geometry_pcs", df_wetland_ny,config.order_geom_lyr_file,1)

        page = 1
        ### print the maps
        if not config.if_multi_page:
            mxd_wetland_ny.saveACopy(os.path.join(config.scratch_folder, "mxd_wetland_ny.mxd"))
            arcpy.mapping.ExportToJPEG(mxd_wetland_ny, output_jpg_ny_wetland, "PAGE_LAYOUT", 480, 640, 150, "False", "24-BIT_TRUE_COLOR", 85)
            shutil.copy(output_jpg_ny_wetland, os.path.join(config.report_path, 'PSRmaps', order_obj.number))
            arcpy.AddMessage('      - Wetland Output for NY state: %s' % os.path.join(config.report_path, 'PSRmaps', order_obj.number))
            del mxd_wetland_ny
            del df_wetland_ny
        else:                           # multipage
            grid_lyr_shp = os.path.join(config.scratch_folder, 'grid_lyr_wetland.shp')
            arcpy.GridIndexFeatures_cartography(grid_lyr_shp, buffer_wetland_shp, "", "", "", config.grid_size, config.grid_size)  #note the tool takes featureclass name only, not the full path

            # part 1: the overview map
            # add grid layer
            grid_layer = arcpy.mapping.Layer(config.grid_lyr_file)
            grid_layer.replaceDataSource(config.scratch_folder,"SHAPEFILE_WORKSPACE","grid_lyr_wetland")
            arcpy.mapping.AddLayer(df_wetland_ny,grid_layer,"Top")

            df_wetland_ny.extent = grid_layer.getExtent()
            df_wetland_ny.scale = df_wetland_ny.scale * 1.1

            mxd_wetland_ny.saveACopy(os.path.join(config.scratch_folder, "mxd_wetland_ny.mxd"))
            arcpy.mapping.ExportToJPEG(mxd_wetland_ny, output_jpg_ny_wetland, "PAGE_LAYOUT", 480, 640, 150, "False", "24-BIT_TRUE_COLOR", 85)
            if not os.path.exists(os.path.join(config.report_path, 'PSRmaps', order_obj.number)):
                os.mkdir(os.path.join(config.report_path, 'PSRmaps', order_obj.number))
            shutil.copy(output_jpg_ny_wetland, os.path.join(config.report_path, 'PSRmaps', order_obj.number))

            del mxd_wetland_ny
            del df_wetland_ny

            # part 2: the data driven pages
            
            page = int(arcpy.GetCount_management(grid_lyr_shp).getOutput(0))  + page
            mxd_mm_wetland_NY = arcpy.mapping.MapDocument(config.mxdMMfile_wetlandNY)

            df_mm_wetland_ny = arcpy.mapping.ListDataFrames(mxd_mm_wetland_NY,"big")[0]
            df_mm_wetland_ny.spatialReference = order_obj.spatial_ref_pcs
            ### add order and order_buffer layers to wetland Newyork multipage mxd file
            utility.add_layer_to_mxd("order_buffer",df_mm_wetland_ny,config.buffer_lyr_file,1.1)
            utility.add_layer_to_mxd("order_geometry_pcs", df_mm_wetland_ny,config.order_geom_lyr_file,1)
         
            grid_layer_mm = arcpy.mapping.ListLayers(mxd_mm_wetland_NY,"Grid" ,df_mm_wetland_ny)[0]
            grid_layer_mm.replaceDataSource(config.scratch_folder, "SHAPEFILE_WORKSPACE","grid_lyr_wetland")
            arcpy.CalculateAdjacentFields_cartography(grid_lyr_shp, "PageNumber")
            mxd_mm_wetland_NY.saveACopy(os.path.join(config.scratch_folder, "mxd_mm_wetland_NY.mxd"))

            for i in range(1,int(arcpy.GetCount_management(grid_lyr_shp).getOutput(0))+1):
                arcpy.SelectLayerByAttribute_management(grid_layer_mm, "NEW_SELECTION", ' "PageNumber" =  ' + str(i))
                df_mm_wetland_ny.extent = grid_layer_mm.getSelectedExtent(True)
                df_mm_wetland_ny.scale = df_mm_wetland_ny.scale * 1.1
                arcpy.SelectLayerByAttribute_management(grid_layer_mm, "CLEAR_SELECTION")
                title_text = arcpy.mapping.ListLayoutElements(df_mm_wetland_ny, "TEXT_ELEMENT", "title")[0]
                title_text.text = "NY Wetland Type - Page " + str(i)
                title_text.elementPositionX = 0.468
                arcpy.RefreshTOC()
                arcpy.mapping.ExportToJPEG(mxd_mm_wetland_NY, output_jpg_ny_wetland[0:-4]+str(i)+".jpg", "PAGE_LAYOUT", 480, 640, 150, "False", "24-BIT_TRUE_COLOR", 85)
                if not os.path.exists(os.path.join(config.report_path, 'PSRmaps', order_obj.number)):
                    os.mkdir(os.path.join(config.report_path, 'PSRmaps', order_obj.number))
                shutil.copy(output_jpg_ny_wetland[0:-4]+str(i)+".jpg", os.path.join(config.report_path, 'PSRmaps', order_obj.number))
                arcpy.AddMessage('      -- Wetland Output for NY state: %s' % os.path.join(config.report_path, 'PSRmaps', order_obj.number))
            del mxd_mm_wetland_NY
            del df_mm_wetland_ny
            
        psr_obj = models.PSR()
        for i in range(1,page): #insert generated .jpg report path into eris_maps_psr table
            psr_obj.insert_map(order_obj.id, 'WETLAND', order_obj.number +'_NY_WETL'+str(i)+'.jpg', int(i)+1)
           
    end = timeit.default_timer()
    arcpy.AddMessage((' -- End generating PSR Wetland report. Duration:', round(end -start,4)))