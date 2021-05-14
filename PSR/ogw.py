from imp import reload
import arcpy, os, sys
import timeit
import shutil
import psr_utility
import psr_config as config
file_path = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(1,os.path.join(os.path.dirname(file_path),'GIS_Utility'))
sys.path.insert(1,os.path.join(os.path.dirname(file_path),'DB_Framework'))
import gis_utility
import models
reload(sys)

def add_fields():
    
    check_elevation_field = arcpy.ListFields(config.wells_sja,"Elevation")
    if len(check_elevation_field) == 0:
        arcpy.AddField_management(config.wells_sja, "Elevation", "DOUBLE", "12", "6", "", "", "NULLABLE", "NON_REQUIRED", "")
    # Add mapkey with script from ERIS toolbox
    check_elevation_map_key = arcpy.ListFields(config.wells_sja,"MapKeyNo")
    if len(check_elevation_map_key) == 0:
        arcpy.AddField_management(config.wells_sja, "MapKeyNo", "SHORT", "", "", "", "", "NULLABLE", "NON_REQUIRED", "")
    #  Add Field for mapkey rank storage based on location and total number of keys at one location
    check_elevation_map_key_loc = arcpy.ListFields(config.wells_sja,"MapKeyLoc")
    if len(check_elevation_map_key_loc) == 0:
        arcpy.AddField_management(config.wells_sja, "MapKeyLoc", "SHORT", "", "", "", "", "NULLABLE", "NON_REQUIRED", "")
    check_elevation_map_key_tot = arcpy.ListFields(config.wells_sja,"MapKeyTot")
    if len(check_elevation_map_key_tot) == 0:
        arcpy.AddField_management(config.wells_sja, "MapKeyTot", "SHORT", "", "", "", "", "NULLABLE", "NON_REQUIRED", "")
    check_direction = arcpy.ListFields(config.wells_sja,"Direction")
    if len(check_direction) == 0:
        arcpy.AddField_management(config.wells_sja, "Direction", "TEXT", "", "", "3", "", "NULLABLE", "NON_REQUIRED", "")
def add_direction():
    desc = arcpy.Describe(config.wells_sja)
    shape_field_name = desc.ShapeFieldName
    rows = arcpy.UpdateCursor(config.wells_sja)
    for row in rows:
        if(row.Distance < 0.001):         # give onsite, give "-" in Direction field
            direction_text = '-'
        else:
            geometry = row.SHAPE
            ref_x =  geometry.trueCentroid.X        # field is directly accessible
            ref_y = geometry.trueCentroid.Y
            feat = row.getValue(shape_field_name)
            pnt = feat.getPart()
            direction_text = psr_utility.get_direction_text(ref_x,ref_y,pnt.X,pnt.Y)
        row.Direction = direction_text   # field is directly accessible
        rows.updateRow(row)
    del rows
def elevation_ranking():
    cur = arcpy.UpdateCursor(config.wells_display)
    for row in cur:
        if row.Ele_Diff < 0.000:
            row.Elev_Rank = -1
        elif row.Ele_Diff == 0:
            row.Elev_Rank = 0
        elif row.Ele_Diff > 0.0000:
            row.Elev_Rank = 1
        else:
            row.Elev_Rank = 100
        cur.updateRow(row)
    # release the layer from locks
    del row, cur
def save_wells_to_db(order_obj, page):
    psr_obj = models.PSR()
    if (int(arcpy.GetCount_management(config.wells_sja).getOutput(0))== 0):
        arcpy.AddMessage('      - No well records are selected.')
        psr_obj.insert_map(order_obj.id, 'WELLS', order_obj.number + '_US_WELLS.jpg', 1)
    else:
        in_rows = arcpy.SearchCursor(config.wells_sja)
        for in_row in in_rows:
            eris_id = str(int(in_row.ID))
            ds_oid=str(int(in_row.DS_OID))
            distance = str(float(in_row.Distance))
            direction = str(in_row.Direction)
            elevation = str(float(in_row.Elevation))
            site_elevation = str(float(in_row.Elevation) - float(in_row.Site_Z))
            map_key_loc = str(int(in_row.MapKeyLoc))
            map_key_no = str(int(in_row.MapKeyNo))
            psr_obj.insert_order_detail(order_obj.id,eris_id, ds_oid, '', distance, direction, elevation, site_elevation, map_key_loc,map_key_no )
        #note type 'SOIL' or 'GEOL' is used internally
        psr_obj.insert_map(order_obj.id, 'WELLS', order_obj.number + '_US_WELLS.jpg', 1)
        if not config.if_multi_page:
            for i in range(1,page):
                psr_obj.insert_map(order_obj.id, 'WELLS', order_obj.number +'_US_WELLS'+str(i)+'.jpg', i + 1)
def generate_ogw_report(order_obj):
    arcpy.AddMessage('  -- Start generating PSR Oil, Gas and Water wells map report...')
    start = timeit.default_timer() 
    ### set scratch folder
    arcpy.env.workspace = config.scratch_folder
    arcpy.env.overwriteOutput = True   
    centre_point = order_obj.geometry.trueCentroid
    elevation = psr_utility.get_elevation(centre_point.X, centre_point.Y)
    if elevation != None:
        centre_point.Z = float(elevation)
    ### create order geometry center shapefile
    order_rows = arcpy.SearchCursor(config.order_geometry_pcs_shp)
    point = arcpy.Point()
    array = arcpy.Array()
    feature_list = []
    arcpy.CreateFeatureclass_management(config.scratch_folder, os.path.basename(config.order_center_pcs), "POINT", "", "DISABLED", "DISABLED", order_obj.spatial_ref_pcs)
    arcpy.AddField_management(config.order_center_pcs, "Site_z", "DOUBLE", "12", "6", "", "", "NULLABLE", "NON_REQUIRED", "")
    insert_cursor = arcpy.InsertCursor(config.order_center_pcs)
    feat = insert_cursor.newRow()
    for order_row in order_rows:
        # Set X and Y for start and end points
        geometry = order_row.SHAPE
        geometry_gcs = geometry.projectAs(order_obj.spatial_ref_gcs)
        site_elevation = psr_utility.get_elevation(geometry_gcs.trueCentroid.X,geometry_gcs.trueCentroid.Y)
        point.X = geometry.trueCentroid.X
        point.Y = geometry.trueCentroid.Y
        array.add(point)
        center_point = arcpy.Multipoint(array)
        array.removeAll()
        feature_list.append(center_point)
        feat.shape = point
        feat.Site_Z = float(site_elevation)
        insert_cursor.insertRow(feat)
    del feat
    del insert_cursor
    del order_row
    del order_rows
    del point
    del array
    
    output_jpg_wells = config.output_jpg(order_obj,config.Report_Type.wells)
    if '10685' not in order_obj.psr.search_radius.keys():
        arcpy.AddMessage('      -- OGW search radius is not availabe')
        return
    config.buffer_dist_ogw =  str(order_obj.psr.search_radius['10685']) + ' MILES'
    ds_oid_wells_max_radius = '10093'     # 10093 is a federal source, PWSV
    ds_oid_wells = []
    for key in order_obj.psr.search_radius:
        if key not in ['9334', '10683', '10684', '10685', '10688','10689', '10695', '10696']:       #10695 is US topo, 10696 is HTMC, 10688 and 10689 are radons
            ds_oid_wells.append(key)
            if (order_obj.psr.search_radius[key] > order_obj.psr.search_radius[ds_oid_wells_max_radius]):
                ds_oid_wells = key

    merge_list = []
    for ds_oid in ds_oid_wells:
        buffer_wells_fc = os.path.join(config.scratch_folder,"order_buffer_" + str(ds_oid) + ".shp")
        arcpy.Buffer_analysis(config.order_geometry_pcs_shp, buffer_wells_fc, str(order_obj.psr.search_radius[ds_oid]) + " MILES")
        wells_clip = os.path.join(config.scratch_folder,'wells_clip_' + str(ds_oid) + '.shp')
        arcpy.Clip_analysis(config.eris_wells, buffer_wells_fc, wells_clip)
        arcpy.Select_analysis(wells_clip, os.path.join(config.scratch_folder,'wells_selected_' + str(ds_oid) + '.shp'), "DS_OID =" + str(ds_oid))
        merge_list.append(os.path.join(config.scratch_folder,'wells_selected_' + str(ds_oid) + '.shp'))
    arcpy.Merge_management(merge_list, config.wells_merge)
    del config.eris_wells
    
    # Calculate Distance with integration and spatial join- can be easily done with Distance tool along with direction if ArcInfo or Advanced license
    wells_merge_pcs= os.path.join(config.scratch_folder,"wells_merge_pcs.shp")
    arcpy.Project_management(config.wells_merge, wells_merge_pcs, order_obj.spatial_ref_pcs)
    arcpy.Integrate_management(wells_merge_pcs, ".5 Meters")
    
    # Add distance to selected wells
    arcpy.SpatialJoin_analysis(wells_merge_pcs, config.order_geometry_pcs_shp, config.wells_sj, "JOIN_ONE_TO_MANY", "KEEP_ALL","#", "CLOSEST","5000 Kilometers", "Distance")   # this is the reported distance
    arcpy.SpatialJoin_analysis(config.wells_sj, config.order_center_pcs, config.wells_sja, "JOIN_ONE_TO_MANY", "KEEP_ALL","#", "CLOSEST","5000 Kilometers", "Dist_cent")  # this is used for mapkey calculation
    if int(arcpy.GetCount_management(os.path.join(config.wells_merge)).getOutput(0)) != 0:
        arcpy.AddMessage('      - Water Wells section, exists water wells')
        add_fields()
        with arcpy.da.UpdateCursor(config.wells_sja, ['X','Y','Elevation']) as update_cursor:
            for row in update_cursor:
                row[2] = psr_utility.get_elevation(row[0],row[1])
                update_cursor.updateRow(row)
        # generate map key
        gis_utility.generate_map_key(config.wells_sja)
        # Add Direction to ERIS sites
        add_direction()
        
        arcpy.Select_analysis(config.wells_sja, config.wells_final, '"MapKeyTot" = 1')
        arcpy.Sort_management(config.wells_final, config.wells_display, [["MapKeyLoc", "ASCENDING"]])

        arcpy.AddField_management(config.wells_display, "Ele_Diff", "DOUBLE", "12", "6", "", "", "NULLABLE", "NON_REQUIRED", "")
        arcpy.CalculateField_management(config.wells_display, 'Ele_Diff', '!Elevation!-!Site_z!', "PYTHON_9.3", "")
        
        arcpy.AddField_management(config.wells_display, "Elev_Rank", "SHORT", "12", "6", "", "", "NULLABLE", "NON_REQUIRED", "")
        # categorize elevation for symbology
        elevation_ranking()
        ## create a map with water wells and ogw wells
        mxd_wells = arcpy.mapping.MapDocument(config.mxd_file_wells)
        df_wells = arcpy.mapping.ListDataFrames(mxd_wells,"*")[0]
        df_wells.spatialReference = order_obj.spatial_ref_pcs
        
        lyr = arcpy.mapping.ListLayers(mxd_wells, "wells", df_wells)[0]
        lyr.replaceDataSource(config.scratch_folder,"SHAPEFILE_WORKSPACE", "wells_display")
    else:
        arcpy.AddMessage('  - WaterWells section, no water wells exists')
        mxd_wells = arcpy.mapping.MapDocument(config.mxd_file_wells)
        df_wells = arcpy.mapping.ListDataFrames(mxd_wells,"*")[0]
        df_wells.spatialReference = order_obj.spatial_ref_pcs
    for item in ds_oid_wells:
        psr_utility.add_layer_to_mxd("order_buffer_" + str(item), df_wells, config.buffer_lyr_file,1.1)
    psr_utility.add_layer_to_mxd("order_geometry_pcs", df_wells,config.order_geom_lyr_file,1)
    # create single-page
    if not config.if_multi_page or int(arcpy.GetCount_management(config.wells_sja).getOutput(0))== 0: 
        mxd_wells.saveACopy(os.path.join(config.scratch_folder, "mxd_wells.mxd"))
        arcpy.mapping.ExportToJPEG(mxd_wells, output_jpg_wells, "PAGE_LAYOUT", 480, 640, 150, "False", "24-BIT_TRUE_COLOR", 85)
        if not os.path.exists(os.path.join(config.report_path, 'PSRmaps', order_obj.number)):
            os.mkdir(os.path.join(config.report_path, 'PSRmaps', order_obj.number))
        shutil.copy(output_jpg_wells, os.path.join(config.report_path, 'PSRmaps', order_obj.number))
        # arcpy.AddMessage('      - output jpg image: %s' % os.path.join(config.report_path, 'PSRmaps', order_obj.number,os.path.basename(output_jpg_wells)))
        del mxd_wells
        del df_wells
    else: # multipage
        grid_lyr_shp = os.path.join(config.scratch_folder, 'grid_lyr_wells.shp')
        #note the tool takes featureclass name only, not the full path
        arcpy.GridIndexFeatures_cartography(grid_lyr_shp, os.path.join(config.scratch_folder,"order_buffer_"+ ds_oid_wells_max_radius + '.shp'), "", "", "",config.grid_size, config.grid_size)  
        # part 1: the overview map
        #add grid layer
        grid_layer = arcpy.mapping.Layer(config.grid_lyr_file)
        grid_layer.replaceDataSource(config.scratch_folder,"SHAPEFILE_WORKSPACE","grid_lyr_wells")
        arcpy.mapping.AddLayer(df_wells,grid_layer,"Top")
        # turn the site label off
        well_lyr = arcpy.mapping.ListLayers(mxd_wells, "wells", df_wells)[0]
        well_lyr.showLabels = False
        df_wells.extent = grid_layer.getExtent()
        df_wells.scale = df_wells.scale * 1.1
        mxd_wells.saveACopy(os.path.join(config.scratch_folder, "mxd_wells.mxd"))
        arcpy.mapping.ExportToJPEG(mxd_wells, output_jpg_wells, "PAGE_LAYOUT", 480, 640, 150, "False", "24-BIT_TRUE_COLOR", 85)
        if not os.path.exists(os.path.join(config.report_path, 'PSRmaps', order_obj.number)):
            os.mkdir(os.path.join(config.report_path, 'PSRmaps', order_obj.number))
        shutil.copy(output_jpg_wells, os.path.join(config.report_path, 'PSRmaps', order_obj.number))
        arcpy.AddMessage('      - output jpg image page 1: %s' % os.path.join(config.report_path, 'PSRmaps', order_obj.number,os.path.basename(output_jpg_wells)))
        del mxd_wells
        del df_wells
        # part 2: the data driven pages
        page = int(arcpy.GetCount_management(grid_lyr_shp).getOutput(0))  + 1
        mxd_mm_wells = arcpy.mapping.MapDocument(config.mxd_mm_file_wells)
        df_mm_wells = arcpy.mapping.ListDataFrames(mxd_mm_wells)[0]
        df_mm_wells.spatialReference = order_obj.spatial_ref_pcs
        for item in ds_oid_wells:
            psr_utility.add_layer_to_mxd("order_buffer_" + str(item), df_mm_wells, config.buffer_lyr_file,1.1)
        psr_utility.add_layer_to_mxd("order_geometry_pcs", df_mm_wells, config.order_geom_lyr_file,1)
        
        grid_layer_mm = arcpy.mapping.ListLayers(mxd_mm_wells,"Grid" ,df_mm_wells)[0]
        grid_layer_mm.replaceDataSource(config.scratch_folder, "SHAPEFILE_WORKSPACE","grid_lyr_wells")
        arcpy.CalculateAdjacentFields_cartography(grid_lyr_shp, 'PageNumber')
        lyr = arcpy.mapping.ListLayers(mxd_mm_wells, "wells", df_mm_wells)[0]   #"wells" or "Wells" doesn't seem to matter
        lyr.replaceDataSource(config.scratch_folder,"SHAPEFILE_WORKSPACE", "wells_display")
        
        for i in range(1,int(arcpy.GetCount_management(grid_lyr_shp).getOutput(0))+1):
            arcpy.SelectLayerByAttribute_management(grid_layer_mm, "NEW_SELECTION", ' "PageNumber" =  ' + str(i))
            df_mm_wells.extent = grid_layer_mm.getSelectedExtent(True)
            df_mm_wells.scale = df_mm_wells.scale * 1.1
            arcpy.SelectLayerByAttribute_management(grid_layer_mm, "CLEAR_SELECTION")
            title_text = arcpy.mapping.ListLayoutElements(mxd_mm_wells, "TEXT_ELEMENT", "MainTitleText")[0]
            title_text.text = "Wells & Additional Sources - Page " + str(i)
            title_text.elementPositionX = 0.6438
            arcpy.RefreshTOC()
            arcpy.mapping.ExportToJPEG(mxd_mm_wells, output_jpg_wells[0:-4]+str(i)+".jpg", "PAGE_LAYOUT", 480, 640, 150, "False", "24-BIT_TRUE_COLOR", 85)
            if not os.path.exists(os.path.join(config.report_path, 'PSRmaps', order_obj.number)):
                os.mkdir(os.path.join(config.report_path, 'PSRmaps', order_obj.number))
            shutil.copy(output_jpg_wells[0:-4]+str(i)+".jpg", os.path.join(config.report_path, 'PSRmaps', order_obj.number))
        del mxd_mm_wells
        del df_mm_wells
        ### Save wells data in database
        save_wells_to_db(order_obj, page)
            
    end = timeit.default_timer()
    arcpy.AddMessage((' -- End generating PSR Oil, Gas and Water wells report. Duration:', round(end -start,4)))