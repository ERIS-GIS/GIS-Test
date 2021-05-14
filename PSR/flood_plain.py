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
def export_to_jpg(page_num):
    arcpy.SelectLayerByAttribute_management(Setting.grid_layer_mm, "NEW_SELECTION", ' "PageNumber" =  ' + str(page_num))
    Setting.df_mm_flood.extent = Setting.grid_layer_mm.getSelectedExtent(True)
    Setting.df_mm_flood.scale = Setting.df_mm_flood.scale * 1.1
    arcpy.SelectLayerByAttribute_management(Setting.grid_layer_mm, "CLEAR_SELECTION")

    title_text = arcpy.mapping.ListLayoutElements(Setting.mxd_multi_flood, "TEXT_ELEMENT", "title")[0]
    title_text.text = '      - Flood Hazard Zones - Page ' + str(page_num)
    title_text.elementPositionX = 0.5946
    arcpy.RefreshTOC()

    arcpy.mapping.ExportToJPEG(Setting.mxd_multi_flood, Setting.output_jpg_flood[0:-4]+str(page_num)+".jpg", "PAGE_LAYOUT", 480, 640, 75, "False", "24-BIT_TRUE_COLOR", 40)
    
    shutil.copy(Setting.output_jpg_flood[0:-4]+str(page_num)+".jpg", os.path.join(config.report_path, 'PSRmaps', Setting.order_obj.number))


def generate_flood_report(order_obj):
    arcpy.AddMessage('  -- Start generating PSR flood report...')
    start = timeit.default_timer()   
  
    eris_id = 0
    Setting.output_jpg_flood = config.output_jpg(order_obj,config.Report_Type.flood)
    page = 1
    if '10683' not in order_obj.psr.search_radius.keys():
        arcpy.AddMessage('      -- Flood search radius is not availabe')
        return
    config.buffer_dist_flood = str(order_obj.psr.search_radius['10683']) + ' MILES'
    arcpy.Buffer_analysis(config.order_geometry_pcs_shp, config.order_buffer_shp, config.buffer_dist_flood) ### create buffer map based on order geometry
    
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
    
    psr_obj = models.PSR()
    
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
        # part 2: the data driven pages
        
        page = int(arcpy.GetCount_management(grid_lyr_shp).getOutput(0))  + page
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

        for i in range(1,page):
            arcpy.SelectLayerByAttribute_management(Setting.grid_layer_mm, "NEW_SELECTION", ' "PageNumber" =  ' + str(i))
            Setting.df_mm_flood.extent = Setting.grid_layer_mm.getSelectedExtent(True)
            Setting.df_mm_flood.scale = Setting.df_mm_flood.scale * 1.1
            arcpy.SelectLayerByAttribute_management(Setting.grid_layer_mm, "CLEAR_SELECTION")

            title_text = arcpy.mapping.ListLayoutElements(Setting.mxd_multi_flood, "TEXT_ELEMENT", "title")[0]
            title_text.text = '      - Flood Hazard Zones - Page ' + str(i)
            title_text.elementPositionX = 0.5946
            arcpy.RefreshTOC()

            arcpy.mapping.ExportToJPEG(Setting.mxd_multi_flood, Setting.output_jpg_flood[0:-4]+str(i)+".jpg", "PAGE_LAYOUT", 480, 640, 75, "False", "24-BIT_TRUE_COLOR", 40)
          
            shutil.copy(Setting.output_jpg_flood[0:-4]+str(i)+".jpg", os.path.join(config.report_path, 'PSRmaps', order_obj.number))
        
        del Setting.mxd_multi_flood
        del Setting.df_mm_flood
        
        ### update tables in DB
        for i in range(1,page):
            psr_obj.insert_map(order_obj.id, 'FLOOD', order_obj.number + '_US_FLOOD' + str(i) + '.jpg', i + 1)
        
    flood_panels = ''
    if (int(arcpy.GetCount_management(os.path.join(config.scratch_folder,"summary_flood.dbf")).getOutput(0))== 0):
        # no floodplain records selected....
        arcpy.AddMessage('      - No floodplain records are selected....')
        if (int(arcpy.GetCount_management(config.flood_panel_selectedby_order_shp).getOutput(0))== 0):
            # no panel available, means no data
            arcpy.AddMessage('      - no panels available in the area')
        else:
            # panel available, just not records in area
            in_rows = arcpy.SearchCursor(config.flood_panel_selectedby_order_shp)
            for in_row in in_rows:
                # arcpy.AddMessage('      - : ' + in_row.FIRM_PAN)    # panel number
                # arcpy.AddMessage('      - %s' % in_row.EFF_DATE)      # effective date
                flood_panels = flood_panels + in_row.FIRM_PAN+'(effective:' + str(in_row.EFF_DATE)[0:10]+') '
                del in_row
            del in_rows
        
        if len(flood_panels) > 0:
            eris_id += 1
            # arcpy.AddMessage('      - erisid for flood_panels is ' + str(eris_id))
            psr_obj.insert_order_detail(order_obj.id,eris_id, '10683')   
            psr_obj.insert_flex_rep(order_obj, eris_id, '10683', 2, 'N', 1, 'Available FIRM Panels in area: ', flood_panels)
        psr_obj.insert_map(order_obj.id, 'FLOOD', order_obj.number + '_US_FLOOD.jpg', 1)
    else:
        in_rows = arcpy.SearchCursor(config.flood_panel_selectedby_order_shp)
        for in_row in in_rows:
            # arcpy.AddMessage('      : ' + in_row.FIRM_PAN)      # panel number
            # arcpy.AddMessage('      - %s' %in_row.EFF_DATE)             # effective date
            flood_panels = flood_panels + in_row.FIRM_PAN+'(effective:' + str(in_row.EFF_DATE)[0:10]+') '
            del in_row
        del in_rows

        in_rows = arcpy.SearchCursor(os.path.join(config.scratch_folder,"summary_flood.dbf"))
        eris_id += 1
        psr_obj.insert_order_detail(order_obj.id , eris_id, '10683')
        psr_obj.insert_flex_rep(order_obj.id, eris_id, '10683', 2, 'N', 1, 'Available FIRM Panels in area: ', flood_panels)
        
        for in_row in in_rows:
            # note the column changed in summary dbf
            # arcpy.AddMessage('      : ' + in_row.ERIS_CLASS)    # eris label
            # arcpy.AddMessage('      : ' + (in_row.FIRST_FLD_))           # zone type
            # arcpy.AddMessage('      : '+ (in_row.FIRST_ZONE))           # subtype

            eris_id += 1
            config.flood_ids.append([in_row.ERIS_CLASS,eris_id])
            
            psr_obj.insert_order_detail(order_obj.id,eris_id, '10683')   
            psr_obj.insert_flex_rep(order_obj.id, eris_id, '10683', 2, 'S1', 1, 'Flood Zone ' + in_row.ERIS_CLASS, '')
            psr_obj.insert_flex_rep(order_obj.id, eris_id, '10683', 2, 'N', 2, 'Zone: ', in_row.FIRST_FLD_)
            psr_obj.insert_flex_rep(order_obj.id, eris_id, '10683', 2, 'N', 3, 'Zone subtype: ', in_row.FIRST_ZONE)
            del in_row
        del in_rows
        psr_obj.insert_map(order_obj.id, 'FLOOD', order_obj.number + '_US_FLOOD.jpg'+'.jpg', 1)
    end = timeit.default_timer()
    arcpy.AddMessage((' -- End generating PSR flood report. Duration:', round(end -start,4)))