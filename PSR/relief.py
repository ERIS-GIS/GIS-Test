from imp import reload
import arcpy, os, sys
import timeit
import shutil
import psr_utility as utility
import psr_config as config
sys.path.insert(1,os.path.join(os.getcwd(),'DB_Framework'))
reload(sys)
import models

def generate_singlepage_report(order_obj, mxd_relief, output_jpg_relief):
    
    mxd_relief.saveACopy(os.path.join(config.scratch_folder,"mxd_relief.mxd"))
    arcpy.mapping.ExportToJPEG(mxd_relief, output_jpg_relief, "PAGE_LAYOUT", 480, 640, 150, "False", "24-BIT_TRUE_COLOR", 85)
    if not os.path.exists(os.path.join(config.report_path, 'PSRmaps', order_obj.number)):
        os.mkdir(os.path.join(config.report_path, 'PSRmaps', order_obj.number))
    shutil.copy(output_jpg_relief, os.path.join(config.report_path, 'PSRmaps', order_obj.number))
    arcpy.AddMessage('      - output jpg image: %s' % os.path.join(config.report_path, 'PSRmaps', order_obj.number,os.path.basename(output_jpg_relief)))
    del mxd_relief
    
def generate_multipage_report(order_obj,cellids_selected, mxd_relief,df_relief, output_jpg_relief):
    df_mm_relief = None
    grid_lyr_shp = os.path.join(config.scratch_folder, 'grid_lyr_relief.shp')
    arcpy.GridIndexFeatures_cartography(grid_lyr_shp, config.order_buffer_shp, "", "", "", config.grid_size, config.grid_size)    
    # part 1: the overview map
    # add grid layer
    grid_layer = arcpy.mapping.Layer(config.grid_lyr_file)
    grid_layer.replaceDataSource(config.scratch_folder,"SHAPEFILE_WORKSPACE","grid_lyr_relief")
    arcpy.mapping.AddLayer(df_relief,grid_layer,"Top")
    
    df_relief.extent = grid_layer.getExtent()
    df_relief.scale = df_relief.scale * 1.1

    mxd_relief.saveACopy(os.path.join(config.scratch_folder, "mxd_relief.mxd"))
    arcpy.mapping.ExportToJPEG(mxd_relief, output_jpg_relief, "PAGE_LAYOUT", 480, 640, 150, "False", "24-BIT_TRUE_COLOR", 85)
    
    if not os.path.exists(os.path.join(config.report_path, 'PSRmaps', order_obj.number)):
        os.mkdir(os.path.join(config.report_path, 'PSRmaps', order_obj.number))
    shutil.copy(output_jpg_relief, os.path.join(config.report_path, 'PSRmaps', order_obj.number))
    arcpy.AddMessage('      - output jpg image page 1: %s' % os.path.join(config.report_path, 'PSRmaps', order_obj.number,os.path.basename(output_jpg_relief)))
    del mxd_relief
    del df_relief
    
    # part 2: the data driven pages
    page = int(arcpy.GetCount_management(grid_lyr_shp).getOutput(0))  + 1
    mxd_mm_relief = arcpy.mapping.MapDocument(config.mxd_mm_file_relief)
    
    df_mm_relief = arcpy.mapping.ListDataFrames(mxd_mm_relief,"*")[0]
    df_mm_relief.spatialReference = order_obj.spatial_ref_pcs
    
    ### add order geometry and it's bugger to mxd
    utility.add_layer_to_mxd("order_buffer",df_mm_relief,config.buffer_lyr_file,1.1)
    utility.add_layer_to_mxd("order_geometry_pcs", df_mm_relief,config.order_geom_lyr_file,1)

    grid_layer_mm = arcpy.mapping.ListLayers(mxd_mm_relief,"Grid" ,df_mm_relief)[0]
    grid_layer_mm.replaceDataSource(config.scratch_folder, "SHAPEFILE_WORKSPACE","grid_lyr_relief")
    arcpy.CalculateAdjacentFields_cartography(grid_lyr_shp, "PageNumber")
    mxd_mm_relief.saveACopy(os.path.join(config.scratch_folder, "mxd_mm_relief.mxd"))
    
    for item in cellids_selected:
        item =item[:-4]
        relief_layer = arcpy.mapping.Layer(config.relief_lyr_file)
        shutil.copyfile(os.path.join(config.path_shaded_relief,item+'_hs.img'),os.path.join(config.scratch_folder,item+'_hs.img'))   #make a local copy, will make it run faster
        relief_layer.replaceDataSource(config.scratch_folder,"RASTER_WORKSPACE",item+'_hs.img')
        relief_layer.name = item
        arcpy.mapping.AddLayer(df_mm_relief, relief_layer, "BOTTOM")

        for i in range(1,int(arcpy.GetCount_management(grid_lyr_shp).getOutput(0))+1):
            arcpy.SelectLayerByAttribute_management(grid_layer_mm, "NEW_SELECTION", ' "PageNumber" =  ' + str(i))
            df_mm_relief.extent = grid_layer_mm.getSelectedExtent(True)
            df_mm_relief.scale = df_mm_relief.scale * 1.1
            arcpy.SelectLayerByAttribute_management(grid_layer_mm, "CLEAR_SELECTION")
            arcpy.mapping.ExportToJPEG(mxd_mm_relief, output_jpg_relief[0:-4]+str(i)+".jpg", "PAGE_LAYOUT", 480, 640, 150, "False", "24-BIT_TRUE_COLOR", 85)

            if not os.path.exists(os.path.join(config.report_path, 'PSRmaps', order_obj.number)):
                os.mkdir(os.path.join(config.report_path, 'PSRmaps', order_obj.number))
            shutil.copy(output_jpg_relief[0:-4]+str(i)+".jpg", os.path.join(config.report_path, 'PSRmaps', order_obj.number))
            
    del mxd_mm_relief
    del df_mm_relief
    #insert generated .jpg report path into eris_maps_psr table
    psr_obj = models.PSR()
    if os.path.exists(os.path.join(config.report_path, 'PSRmaps', order_obj.number, order_obj.number + '_US_RELIEF.jpg')):
        psr_obj.insert_map(order_obj.id, 'RELIEF', order_obj.number + '_US_RELIEF.jpg', 1)
        for i in range(1,page):
            psr_obj.insert_map(order_obj.id, 'RELIEF', order_obj.number + '_US_RELIEF' + str(i) + '.jpg', i + 1)
    else:
        arcpy.AddMessage('No Relief map is available')
       
def generate_relief_report(order_obj):
    arcpy.AddMessage('  -- Start generating PSR shaded relief map report...')
    start = timeit.default_timer() 

    arcpy.env.workspace = config.scratch_folder
    arcpy.env.overwriteOutput = True   
    ### set paths
    output_jpg_relief = config.output_jpg(order_obj,config.Report_Type.relief)

    point = arcpy.Point()
    array = arcpy.Array()
    feature_list = []
    ### set mxd variable
    mxd_relief = arcpy.mapping.MapDocument(config.mxd_file_relief)
    df_relief = arcpy.mapping.ListDataFrames(mxd_relief,"*")[0]
    df_relief.spatialReference = order_obj.spatial_ref_pcs
    
    ### create buffer map based on order geometry and topo redius
    arcpy.Buffer_analysis(config.order_geometry_pcs_shp, config.order_buffer_shp, config.buffer_dist_relief) 

    ### add order geometry and it's bugger to mxd
    utility.add_layer_to_mxd("order_buffer",df_relief,config.buffer_lyr_file,1.1)
    utility.add_layer_to_mxd("order_geometry_pcs", df_relief,config.order_geom_lyr_file,1)

    # locate and add relevant shadedrelief tiles
    width = arcpy.Describe(config.order_buffer_shp).extent.width/2
    height = arcpy.Describe(config.order_buffer_shp).extent.height/2

    if (width/height > 5/4.4): # wider shape
        height = width/5*4.4
    else: # longer shape
        width = height/4.4*5

    x_centr_oid = (arcpy.Describe(config.order_buffer_shp).extent.XMax + arcpy.Describe(config.order_buffer_shp).extent.XMin)/2
    y_centr_oid = (arcpy.Describe(config.order_buffer_shp).extent.YMax + arcpy.Describe(config.order_buffer_shp).extent.YMin)/2
    width = width + 6400     # add 2 miles to each side, for multipage
    height = height + 6400   # add 2 miles to each side, for multipage

    point.X = x_centr_oid - width
    point.Y = y_centr_oid + height
    array.add(point)
    point.X = x_centr_oid + width
    point.Y = y_centr_oid + height
    array.add(point)
    point.X = x_centr_oid + width
    point.Y = y_centr_oid - height
    array.add(point)
    point.X = x_centr_oid - width
    point.Y = y_centr_oid - height
    array.add(point)
    point.X = x_centr_oid - width
    point.Y = y_centr_oid + height
    array.add(point)
    feat = arcpy.Polygon(array,order_obj.spatial_ref_pcs)
    array.removeAll()
    feature_list.append(feat)

    arcpy.CopyFeatures_management(feature_list, config.relief_frame)
    master_layer_relief = arcpy.mapping.Layer(config.master_lyr_dem)
    arcpy.SelectLayerByLocation_management(master_layer_relief,'intersect',config.relief_frame)

    cellids_selected = []
    if(int((arcpy.GetCount_management(master_layer_relief).getOutput(0))) ==0):
        arcpy.AddMessage ("NO records selected")
        master_layer_relief = None
    else:
        cellid = ''
        # loop through the relevant records, locate the selected cell IDs
        rows = arcpy.SearchCursor(master_layer_relief)    # loop through the selected records
        for row in rows:
            cellid = str(row.getValue("image_name")).strip()
            if cellid !='':
                cellids_selected.append(cellid)
            del row
        del rows
        master_layer_relief = None

        for item in cellids_selected:
            item =item[:-4]
            relief_layer = arcpy.mapping.Layer(config.relief_lyr_file)
            shutil.copyfile(os.path.join(config.path_shaded_relief,item+'_hs.img'),os.path.join(config.scratch_folder,item+'_hs.img'))
            relief_layer.replaceDataSource(config.scratch_folder,"RASTER_WORKSPACE",item+'_hs.img')
            relief_layer.name = item
            arcpy.mapping.AddLayer(df_relief, relief_layer, "BOTTOM")
            relief_layer = None

            arcpy.RefreshActiveView()
        if not config.if_multi_page:
            generate_singlepage_report (order_obj, mxd_relief, output_jpg_relief)
        else:
            generate_multipage_report (order_obj,cellids_selected, mxd_relief,df_relief, output_jpg_relief)
    end = timeit.default_timer()
    arcpy.AddMessage(('-- End generating PSR shaded relief report. Duration:', round(end -start,4)))
    