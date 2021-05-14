from imp import reload
import arcpy, os, sys
import timeit
import shutil
import psr_utility as utility
import psr_config as config
sys.path.insert(1,os.path.join(os.getcwd(),'DB_Framework'))
reload(sys)
import models

def generate_geology_report(order_obj):
    arcpy.AddMessage('  -- Start generating PSR geology report...')
    start = timeit.default_timer() 
    ### set scratch folder
    arcpy.env.workspace = config.scratch_folder
    arcpy.env.overwriteOutput = True   
    output_jpg_geology = config.output_jpg(order_obj,config.Report_Type.geology)
    page = 1
    if '10685' not in order_obj.psr.search_radius.keys():
        arcpy.AddMessage('      -- Geology search radius is not availabe')
        return
    config.buffer_dist_geology =  str(order_obj.psr.search_radius['10685']) + ' MILES'

    ### create buffer map based on order geometry
    arcpy.Buffer_analysis(config.order_geometry_pcs_shp, config.order_buffer_shp, config.buffer_dist_geology) 
    
    arcpy.MakeFeatureLayer_management(config.data_geology, 'geology_lyr') 
    arcpy.SelectLayerByLocation_management('geology_lyr', 'intersect',  config.order_buffer_shp)
    arcpy.CopyFeatures_management('geology_lyr', config.geology_selectedby_order_shp)
    
    arcpy.Statistics_analysis(config.geology_selectedby_order_shp, os.path.join(config.scratch_folder,"summary_geology.dbf"), [['UNIT_NAME','FIRST'], ['UNIT_AGE','FIRST'], ['ROCKTYPE1','FIRST'], ['ROCKTYPE2','FIRST'], ['UNITDESC','FIRST'], ['ERIS_KEY_1','FIRST']],'ORIG_LABEL')
    arcpy.Sort_management(os.path.join(config.scratch_folder,"summary_geology.dbf"), os.path.join(config.scratch_folder,"summary_sorted_geol.dbf"), [["ORIG_LABEL", "ASCENDING"]])

    mxd_geology = arcpy.mapping.MapDocument(config.mxd_file_geology)
    df_geology = arcpy.mapping.ListDataFrames(mxd_geology,"*")[0]
    df_geology.spatialReference = order_obj.spatial_ref_pcs
    
    ### add order and order_buffer layers to geology mxd file
    utility.add_layer_to_mxd("order_buffer",df_geology,config.buffer_lyr_file, 1.1)
    utility.add_layer_to_mxd("order_geometry_pcs", df_geology,config.order_geom_lyr_file,1)
    
    if not config.if_multi_page: # single-page
        #df.scale = 5000
        mxd_geology.saveACopy(os.path.join(config.scratch_folder, "mxd_geology.mxd"))
        arcpy.mapping.ExportToJPEG(mxd_geology, output_jpg_geology, "PAGE_LAYOUT", 480, 640, 150, "False", "24-BIT_TRUE_COLOR", 85)
        if not os.path.exists(os.path.join(config.report_path, 'PSRmaps',  order_obj.number)):
            os.mkdir(os.path.join(config.report_path, 'PSRmaps',  order_obj.number))
        shutil.copy(output_jpg_geology, os.path.join(config.report_path, 'PSRmaps', order_obj.number))
        arcpy.AddMessage('      - output jpg image: %s' % os.path.join(config.report_path, 'PSRmaps', order_obj.number,os.path.basename(output_jpg_geology)))
        del mxd_geology
        del df_geology
    else:    # multipage
        grid_lyr_shp = os.path.join(config.scratch_folder, 'grid_lyr_geology.shp')
        arcpy.GridIndexFeatures_cartography(grid_lyr_shp, config.order_buffer_shp, "", "", "", config.grid_size, config.grid_size)
        
        # part 1: the overview map
        # add grid layer
        grid_layer = arcpy.mapping.Layer(config.grid_lyr_file)
        grid_layer.replaceDataSource(config.scratch_folder,"SHAPEFILE_WORKSPACE","grid_lyr_geology")
        arcpy.mapping.AddLayer(df_geology,grid_layer,"Top")
        
        df_geology.extent = grid_layer.getExtent()
        df_geology.scale = df_geology.scale * 1.1
        
        mxd_geology.saveACopy(os.path.join(config.scratch_folder, "mxd_geology.mxd"))
        arcpy.mapping.ExportToJPEG(mxd_geology, output_jpg_geology, "PAGE_LAYOUT", 480, 640, 150, "False", "24-BIT_TRUE_COLOR", 85)
        
        if not os.path.exists(os.path.join(config.report_path, 'PSRmaps', order_obj.number)):
            os.mkdir(os.path.join(config.report_path, 'PSRmaps', order_obj.number))
        shutil.copy(output_jpg_geology, os.path.join(config.report_path, 'PSRmaps', order_obj.number))
        arcpy.AddMessage('      - output jpg image page 1: %s' % os.path.join(config.report_path, 'PSRmaps', order_obj.number,os.path.basename(output_jpg_geology)))
        del mxd_geology
        del df_geology
        
        # part 2: the data driven pages
        
        page = int(arcpy.GetCount_management(grid_lyr_shp).getOutput(0))  + 1
        mxd_mm_geology = arcpy.mapping.MapDocument(config.mxd_mm_file_geology)
        
        df_mm_geology = arcpy.mapping.ListDataFrames(mxd_mm_geology,"*")[0]
        df_mm_geology.spatialReference = order_obj.spatial_ref_pcs
        utility.add_layer_to_mxd("order_buffer",df_mm_geology,config.buffer_lyr_file,1.1)
        utility.add_layer_to_mxd("order_geometry_pcs", df_mm_geology,config.order_geom_lyr_file,1)
        
        grid_layer_mm = arcpy.mapping.ListLayers(mxd_mm_geology,"Grid" ,df_mm_geology)[0]
        grid_layer_mm.replaceDataSource(config.scratch_folder,"SHAPEFILE_WORKSPACE","grid_lyr_geology")
        arcpy.CalculateAdjacentFields_cartography(grid_lyr_shp, "PageNumber")
        mxd_mm_geology.saveACopy(os.path.join(config.scratch_folder, "mxd_mm_geology.mxd"))
        
        for i in range(1,int(arcpy.GetCount_management(grid_lyr_shp).getOutput(0))+1):
            arcpy.SelectLayerByAttribute_management(grid_layer_mm, "NEW_SELECTION", ' "PageNumber" =  ' + str(i))
            df_mm_geology.extent = grid_layer_mm.getSelectedExtent(True)
            df_mm_geology.scale = df_mm_geology.scale * 1.1
            arcpy.SelectLayerByAttribute_management(grid_layer_mm, "CLEAR_SELECTION")

            title_text = arcpy.mapping.ListLayoutElements(mxd_mm_geology, "TEXT_ELEMENT", "title")[0]
            title_text.text = "Geologic Units - Page " + str(i)
            title_text.elementPositionX = 0.6303
            arcpy.RefreshTOC()

            arcpy.mapping.ExportToJPEG(mxd_mm_geology, output_jpg_geology[0:-4]+str(i)+".jpg", "PAGE_LAYOUT", 480, 640, 150, "False", "24-BIT_TRUE_COLOR", 85)
            if not os.path.exists(os.path.join(config.report_path, 'PSRmaps', order_obj.number)):
                os.mkdir(os.path.join(config.report_path, 'PSRmaps', order_obj.number))
            shutil.copy(output_jpg_geology[0:-4]+str(i)+".jpg", os.path.join(config.report_path, 'PSRmaps', order_obj.number))
            # arcpy.AddMessage('      - output jpg image: %s' % os.path.join(config.report_path, 'PSRmaps', order_obj.number, os.path.basename(output_jpg_geology[0:-4]+str(i)+".jpg")))
        del mxd_mm_geology
        del df_mm_geology
        psr_obj = models.PSR()
        for i in range(1,page):
            psr_obj.insert_map(order_obj.id, 'GEOL', order_obj.number +'_US_GEOL'+str(i)+'.jpg', i+1)
        
    if (int(arcpy.GetCount_management(os.path.join(config.scratch_folder,"summary_sorted_geol.dbf")).getOutput(0))== 0):
        # no geology polygon selected...., need to send in map only
        arcpy.AddMessage('No geology polygon is selected....')
        psr_obj = models.PSR()
        psr_obj.insert_map(order_obj.id, 'GEOL', order_obj.number + '_US_GEOLOGY.jpg', 1)  #note type 'SOIL' or 'GEOL' is used internally
    else:
        eris_id = 0
        psr_obj = models.PSR()
        in_rows = arcpy.SearchCursor(os.path.join(config.scratch_folder,config.geology_selectedby_order_shp))
        for in_row in in_rows:
            # note the column changed in summary dbf
            # arcpy.AddMessage("Unit label is: " + in_row.ORIG_LABEL)
            # arcpy.AddMessage(in_row.UNIT_NAME)     # unit name
            # arcpy.AddMessage(in_row.UNIT_AGE)      # unit age
            # arcpy.AddMessage( in_row.ROCKTYPE1)      # rocktype 1
            # arcpy.AddMessage( in_row.ROCKTYPE2)      # rocktype2
            # arcpy.AddMessage( in_row.UNITDESC)       # unit description
            # arcpy.AddMessage( in_row.ERIS_KEY_1)     # eris key created from upper(unit_link)
            eris_id = eris_id + 1
            config.geology_ids.append([in_row.ERIS_KEY_1,eris_id])
            psr_obj.insert_order_detail(order_obj.id,eris_id, '10685')   
            psr_obj.insert_flex_rep(order_obj.id, eris_id, '10685', 2, 'S1', 1, 'Geologic Unit ' + in_row.ORIG_LABEL, '')
            psr_obj.insert_flex_rep(order_obj.id, eris_id, '10685', 2, 'N', 2, 'Unit Name: ', in_row.UNIT_NAME)
            psr_obj.insert_flex_rep(order_obj.id, eris_id, '10685', 2, 'N', 3,'Unit Age: ', in_row.UNIT_AGE)
            psr_obj.insert_flex_rep(order_obj.id, eris_id, '10685', 2, 'N', 4, 'Primary Rock Type ', in_row.ROCKTYPE1)
            psr_obj.insert_flex_rep(order_obj.id, eris_id, '10685', 2, 'N', 4, 'Secondary Rock Type: ', in_row.ROCKTYPE2)
            if in_row.UNITDESC == None:
                node_scr = 'No description available.'
                psr_obj.insert_flex_rep(order_obj.id, eris_id, '10685', 2, 'N', 6, 'Unit Description: ', node_scr)
            else:
                psr_obj.insert_flex_rep(order_obj.id, eris_id, '10685', 2, 'N', 6, 'Unit Description: ', in_row.UNITDESC.encode('utf-8'))
            del in_row
        del in_rows
        psr_obj.insert_map(order_obj.id, 'GEOL', order_obj.number + '_US_GEOLOGY.jpg', 1)
    
    end = timeit.default_timer()
    arcpy.AddMessage((' -- End generating PSR geology report. Duration:', round(end -start,4)))