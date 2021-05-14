from imp import reload
import arcpy, os, sys
import timeit
import shutil
import psr_utility as utility
import psr_config as config
sys.path.insert(1,os.path.join(os.getcwd(),'DB_Framework'))
reload(sys)
import models

def create_map(order_obj):
    point = arcpy.Point()
    array = arcpy.Array()
    feature_list = []
    width = arcpy.Describe(config.order_buffer_shp).extent.width/2
    height = arcpy.Describe(config.order_buffer_shp).extent.height/2
    
    if (width > 662 or height > 662):
        if (width/height > 1):
            # buffer has a wider shape
            width = width * 1.1
            height = width

        else:
            # buffer has a vertically elonged shape
            height = height * 1.1
            width = height
    else:
        width = 662*1.1
        height = 662*1.1
    width = width + 6400     #add 2 miles to each side, for multipage soil
    height = height + 6400   #add 2 miles to each side, for multipage soil
    xCentroid = (arcpy.Describe(config.order_buffer_shp).extent.XMax + arcpy.Describe(config.order_buffer_shp).extent.XMin)/2
    yCentroid = (arcpy.Describe(config.order_buffer_shp).extent.YMax + arcpy.Describe(config.order_buffer_shp).extent.YMin)/2
    point.X = xCentroid-width
    point.Y = yCentroid+height
    array.add(point)
    point.X = xCentroid+width
    point.Y = yCentroid+height
    array.add(point)
    point.X = xCentroid+width
    point.Y = yCentroid-height
    array.add(point)
    point.X = xCentroid-width
    point.Y = yCentroid-height
    array.add(point)
    point.X = xCentroid-width
    point.Y = yCentroid+height
    array.add(point)
    feat = arcpy.Polygon(array,order_obj.spatial_ref_pcs)
    array.removeAll()
    feature_list.append(feat)
        
    frame_soil = os.path.join(config.scratch_folder, "frame_soil.shp")
    arcpy.CopyFeatures_management(feature_list, frame_soil)
    arcpy.SelectLayerByLocation_management(config.soil_lyr,'intersect',frame_soil)
    arcpy.CopyFeatures_management(config.soil_lyr, config.soil_selectedby_frame)
    
    # add another column to soil_disp just for symbology purpose
    arcpy.AddField_management(os.path.join(config.scratch_folder, config.soil_selectedby_frame), "FIDCP", "TEXT", "", "", "", "", "NON_NULLABLE", "REQUIRED", "")
    arcpy.CalculateField_management(os.path.join(config.scratch_folder, config.soil_selectedby_frame), 'FIDCP', '!FID!', "PYTHON_9.3")
    
    mxd_soil = arcpy.mapping.MapDocument(config.mxd_file_soil)
    df_soil = arcpy.mapping.ListDataFrames(mxd_soil,"*")[0]
    df_soil.spatialReference = order_obj.spatial_ref_pcs
    
    ssurgo_lyr = arcpy.mapping.ListLayers(mxd_soil, "SSURGO*", df_soil)[0]
    ssurgo_lyr.replaceDataSource(config.scratch_folder,"SHAPEFILE_WORKSPACE", 'soil_selectedby_frame')
    ssurgo_lyr.symbology.addAllValues()
    config.soil_lyr = ssurgo_lyr
    
    utility.add_layer_to_mxd("order_buffer",df_soil,config.buffer_lyr_file, 1.1)
    utility.add_layer_to_mxd("order_geometry_pcs", df_soil,config.order_geom_lyr_file,1)
    arcpy.RefreshActiveView()
    output_jpg_soil = config.output_jpg(order_obj,config.Report_Type.soil)
    if not config.if_multi_page: # single-page
        mxd_soil.saveACopy(os.path.join(config.scratch_folder, "mxd_soil.mxd"))
        arcpy.mapping.ExportToJPEG(mxd_soil, output_jpg_soil, "PAGE_LAYOUT", 480, 640, 150, "False", "24-BIT_TRUE_COLOR", 85)
        if not os.path.exists(os.path.join(config.report_path, 'PSRmaps', order_obj.number)):
            os.mkdir(os.path.join(config.report_path, 'PSRmaps', order_obj.number))
        shutil.copy(output_jpg_soil, os.path.join(config.report_path, 'PSRmaps', order_obj.number))
        arcpy.AddMessage('      - output jpg image %s' % os.path.join(config.report_path, 'PSRmaps', order_obj.number,os.path.basename(output_jpg_soil)))
        del mxd_soil
        del df_soil
    else: # multipage
        grid_lyr_shp = os.path.join(config.scratch_folder, 'grid_lyr_soil.shp')
        arcpy.GridIndexFeatures_cartography(grid_lyr_shp, config.order_buffer_shp, "", "", "", config.grid_size, config.grid_size)
        # part 1: the overview map
        # add grid layer
        grid_layer = arcpy.mapping.Layer(config.grid_lyr_file)
        grid_layer.replaceDataSource(config.scratch_folder,"SHAPEFILE_WORKSPACE","grid_lyr_soil")
        arcpy.mapping.AddLayer(df_soil, grid_layer,"Top")

        df_soil.extent = grid_layer.getExtent()
        df_soil.scale = df_soil.scale * 1.1

        mxd_soil.saveACopy(os.path.join(config.scratch_folder, "mxd_soil.mxd"))
        arcpy.mapping.ExportToJPEG(mxd_soil, output_jpg_soil, "PAGE_LAYOUT", 480, 640, 150, "False", "24-BIT_TRUE_COLOR", 85)
        if not os.path.exists(os.path.join(config.report_path, 'PSRmaps', order_obj.number)):
            os.mkdir(os.path.join(config.report_path, 'PSRmaps', order_obj.number))
        shutil.copy(output_jpg_soil, os.path.join(config.report_path, 'PSRmaps', order_obj.number))
        arcpy.AddMessage('      - output jpg image page 1: %s' % os.path.join(config.report_path, 'PSRmaps', order_obj.number,os.path.basename(output_jpg_soil)))
        del mxd_soil
        del df_soil
        
        # part 2: the data driven pages maps
        page = int(arcpy.GetCount_management(grid_lyr_shp).getOutput(0))  + 1
        mxd_mm_soil = arcpy.mapping.MapDocument(config.mxd_mm_file_soil)

        df_mm_soil = arcpy.mapping.ListDataFrames(mxd_mm_soil,"*")[0]
        df_mm_soil.spatialReference = order_obj.spatial_ref_pcs
        
        utility.add_layer_to_mxd("order_buffer",df_mm_soil,config.buffer_lyr_file,1.1)
        utility.add_layer_to_mxd("order_geometry_pcs", df_mm_soil, config.order_geom_lyr_file,1)
        
        ssurgo_lyr = arcpy.mapping.ListLayers(mxd_mm_soil, "SSURGO*", df_mm_soil)[0]
        ssurgo_lyr.replaceDataSource(config.scratch_folder,"SHAPEFILE_WORKSPACE", "soil_selectedby_frame")
        ssurgo_lyr.symbology.addAllValues()
        config.soil_lyr = ssurgo_lyr

        grid_layer_mm = arcpy.mapping.ListLayers(mxd_mm_soil,"Grid" ,df_mm_soil)[0]
        grid_layer_mm.replaceDataSource(config.scratch_folder, "SHAPEFILE_WORKSPACE","grid_lyr_soil")
        arcpy.CalculateAdjacentFields_cartography(grid_lyr_shp, 'PageNumber')
        mxd_mm_soil.saveACopy(os.path.join(config.scratch_folder, "mxd_mm_soil.mxd"))

        for i in range(1,int(arcpy.GetCount_management(grid_lyr_shp).getOutput(0))+1):
            arcpy.SelectLayerByAttribute_management(grid_layer_mm, "NEW_SELECTION", ' "PageNumber" =  ' + str(i))
            df_mm_soil.extent = grid_layer_mm.getSelectedExtent(True)
            df_mm_soil.scale = df_mm_soil.scale * 1.1
            arcpy.SelectLayerByAttribute_management(grid_layer_mm, "CLEAR_SELECTION")

            title_text = arcpy.mapping.ListLayoutElements(mxd_mm_soil, "TEXT_ELEMENT", "title")[0]
            title_text.text = "SSURGO Soils - Page " + str(i)
            title_text.elementPositionX = 0.6156
            arcpy.RefreshTOC()

            arcpy.mapping.ExportToJPEG(mxd_mm_soil, output_jpg_soil[0:-4]+str(i)+".jpg", "PAGE_LAYOUT", 480, 640, 150, "False", "24-BIT_TRUE_COLOR", 85)
            if not os.path.exists(os.path.join(config.report_path, 'PSRmaps', order_obj.number)):
                os.mkdir(os.path.join(config.report_path, 'PSRmaps', order_obj.number))
            shutil.copy(output_jpg_soil[0:-4]+str(i)+".jpg", os.path.join(config.report_path, 'PSRmaps', order_obj.number))
        del mxd_mm_soil
        del df_mm_soil
        ### Update DB if multiple pages
        psr_obj = models.PSR()
        for i in range(1,page):
            psr_obj.insert_map(order_obj.id, 'SOIL', order_obj.number + str(i) + '.jpg', i + 1)

    eris_id = 0
    ### Udpate DB
    psr_obj = models.PSR()
    for map_unit in config.report_data:
        eris_id = eris_id + 1
        mu_key = str(map_unit['Mukey'])
        config.soil_ids.append([map_unit['Musym'],eris_id])
        psr_obj.insert_order_detail(order_obj.id,eris_id, mu_key)
        psr_obj.insert_flex_rep(order_obj.id, eris_id, '9334', 2, 'S1', 1, 'Map Unit ' + map_unit['Musym'] + " (%s)" %map_unit["Soil_Percent"], '')  
        psr_obj.insert_flex_rep(order_obj.id, eris_id, '9334', 2, 'N', 2, 'Map Unit Name:', map_unit['Map Unit Name']) 
        if len(map_unit) < 6 :    #for Water, Urbanland and Gravel Pits
            psr_obj.insert_flex_rep(order_obj.id, eris_id, '9334', 2, 'N', 3, 'No more attributes available for this map unit','')
        else:           # not do for Water or urban land
            psr_obj.insert_flex_rep(order_obj.id, eris_id, '9334', 2, 'N', 3, 'Bedrock Depth - Min:',map_unit['Bedrock Depth - Min']) 
            psr_obj.insert_flex_rep(order_obj.id, eris_id, '9334', 2, 'N', 4, 'Watertable Depth - Annual Min:', map_unit['Watertable Depth - Annual Min'])
            if (map_unit['Drainage Class - Dominant'] == '-99'):
                psr_obj.insert_flex_rep(order_obj.id, eris_id, '9334', 2, 'N', 5, 'Drainage Class - Dominant:', 'null')
            else: 
                psr_obj.insert_flex_rep(order_obj.id, eris_id, '9334', 2, 'N', 5, 'Drainage Class - Dominant:', map_unit['Drainage Class - Dominant'])
            if (map_unit['Hydrologic Group - Dominant'] == '-99'):
                psr_obj.insert_flex_rep(order_obj.id, eris_id, '9334', 2, 'N', 6, 'Hydrologic Group - Dominant:', 'null')
            else:
                psr_obj.insert_flex_rep(order_obj.id, eris_id, '9334', 2, 'N', 6, 'Hydrologic Group - Dominant:', map_unit['Hydrologic Group - Dominant'] + ' - ' + config.hydrologic_dict[map_unit['Hydrologic Group - Dominant']])
            psr_obj.insert_flex_rep(order_obj.id, eris_id, '9334', 2, 'N', 7, 'Major components are printed below', '')

            k = 7
            if 'component' in map_unit.keys():
                k = k + 1
                for comp in map_unit['component']:
                    psr_obj.insert_flex_rep(order_obj.id, eris_id, '9334', 2, 'S2', k, comp[0][0], '')
                    for i in range(1,len(comp)):
                        k = k+1
                        psr_obj.insert_flex_rep(order_obj.id, eris_id, '9334', 2, 'S3',k, comp[i][0], comp[i][1])
    psr_obj.insert_map(order_obj.id, 'SOIL', order_obj.number + 'US_SOIL.jpg', 1)
                
def generate_soil_report(order_obj):
    arcpy.AddMessage('  -- Start generating PSR soil report...')
    start = timeit.default_timer() 
    ### set scratch folder
    arcpy.env.workspace = config.scratch_folder
    arcpy.env.overwriteOutput = True   
    ### extract buffer size for soil report
    eris_id = 0
    
    if '9334' not in order_obj.psr.search_radius.keys():
        arcpy.AddMessage('      -- Soil search radius is not availabe')
        return
    config.buffer_dist_soil = str(order_obj.psr.search_radius['9334']) + ' MILES'
    arcpy.Buffer_analysis(config.order_geometry_pcs_shp, config.order_buffer_shp, config.buffer_dist_soil) 
    
    if order_obj.province == 'HI':
        config.data_path_soil = config.data_path_soil_HI
    elif order_obj.province == 'AK':
        config.data_path_soil = config.data_path_soil_AK
    else:
        config.data_path_soil = config.data_path_soil_CONUS
        
    data_soil = os.path.join(config.data_path_soil,'MUPOLYGON')
    # select soil data by using spatial query of order buffere layer
    config.soil_lyr = arcpy.MakeFeatureLayer_management(data_soil,'soil_lyr') 
    arcpy.SelectLayerByLocation_management(config.soil_lyr, 'intersect',  config.order_buffer_shp)
    arcpy.CopyFeatures_management(config.soil_lyr, config.soil_selectedby_order_shp)
    
    table_muaggatt = os.path.join(config.data_path_soil,'muaggatt')
    table_component = os.path.join(config.data_path_soil,'component')
    table_chorizon = os.path.join(config.data_path_soil,'chorizon')
    table_chtexturegrp = os.path.join(config.data_path_soil,'chtexturegrp')
    
    stable_muaggatt = os.path.join(config.temp_gdb,"muaggatt")
    stable_component = os.path.join(config.temp_gdb,"component")
    stable_chorizon = os.path.join(config.temp_gdb,"chorizon")
    stable_chtexture_grp = os.path.join(config.temp_gdb,"chtexturegrp")
    data_array = []
    if (int(arcpy.GetCount_management('soil_lyr').getOutput(0)) == 0):   # no soil polygons selected
        arcpy.AddMessage('no soil data in order geometry buffer')
        psr_obj = models.PCR()
        eris_id = eris_id + 1
        psr_obj.insert_flex_rep(order_obj.id, eris_id, '9334', 2, 'N', 1, 'No soil data available in the project area.', '')
    else:
        soil_selectedby_order_pcs_shp = arcpy.Project_management(config.soil_selectedby_order_shp, config.soil_selectedby_order_pcs_shp, order_obj.spatial_ref_pcs)
        # create map keys
        arcpy.Statistics_analysis(soil_selectedby_order_pcs_shp, os.path.join(config.scratch_folder,"summary_soil.dbf"), [['mukey','FIRST'],["Shape_Area","SUM"]],'musym')
        arcpy.Sort_management(os.path.join(config.scratch_folder,"summary_soil.dbf"), os.path.join(config.scratch_folder,"summary_sorted_soil.dbf"), [["musym", "ASCENDING"]])
        seq_array = arcpy.da.TableToNumPyArray(os.path.join(config.scratch_folder,'summary_sorted_soil.dbf'), '*')    #note: it could contain 'NOTCOM' record
        # retrieve attributes
        unique_MuKeys = utility.return_unique_setstring_musym(soil_selectedby_order_pcs_shp)
        
        if len(unique_MuKeys) > 0:    # special case: order only returns one "NOTCOM" category, filter out
            where_clause_select_table = "muaggatt.mukey in " + unique_MuKeys
            arcpy.TableSelect_analysis(table_muaggatt, stable_muaggatt, where_clause_select_table)

            where_clause_select_table = "component.mukey in " + unique_MuKeys
            arcpy.TableSelect_analysis(table_component, stable_component, where_clause_select_table)

            unique_co_keys = utility.return_unique_set_string(stable_component, 'cokey')
            where_clause_select_table = "chorizon.cokey in " + unique_co_keys
            arcpy.TableSelect_analysis(table_chorizon, stable_chorizon, where_clause_select_table)

            unique_achkeys = utility.return_unique_set_string(stable_chorizon,'chkey')
            if len(unique_achkeys) > 0:       # special case: e.g. there is only one Urban Land polygon
                where_clause_select_table = "chorizon.chkey in " + unique_achkeys
                arcpy.TableSelect_analysis(table_chtexturegrp, stable_chtexture_grp, where_clause_select_table)

                table_list = [stable_muaggatt, stable_component,stable_chorizon, stable_chtexture_grp]
                field_list  = config.fc_soils_field_list #[['muaggatt.mukey','mukey'], ['muaggatt.musym','musym'], ['muaggatt.muname','muname'],['muaggatt.drclassdcd','drclassdcd'],['muaggatt.hydgrpdcd','hydgrpdcd'],['muaggatt.hydclprs','hydclprs'], ['muaggatt.brockdepmin','brockdepmin'], ['muaggatt.wtdepannmin','wtdepannmin'], ['component.cokey','cokey'],['component.compname','compname'], ['component.comppct_r','comppct_r'], ['component.majcompflag','majcompflag'],['chorizon.chkey','chkey'],['chorizon.hzname','hzname'],['chorizon.hzdept_r','hzdept_r'],['chorizon.hzdepb_r','hzdepb_r'], ['chtexturegrp.chtgkey','chtgkey'], ['chtexturegrp.texdesc1','texdesc'], ['chtexturegrp.rvindicator','rv']]
                key_list = config.fc_soils_key_list #['muaggatt.mukey', 'component.cokey','chorizon.chkey','chtexturegrp.chtgkey']
            
                where_clause_query_table = config.fc_soils_where_clause_query_table#"muaggatt.mukey = component.mukey and component.cokey = chorizon.cokey and chorizon.chkey = chtexturegrp.chkey"
                #Query tables may only be created using data from a geodatabase or an OLE DB connection
                query_table = arcpy.MakeQueryTable_management(table_list,'query_table','USE_KEY_FIELDS', key_list, field_list, where_clause_query_table)  #note: outTable is a table view and won't persist

                arcpy.TableToTable_conversion(query_table, config.temp_gdb, 'soil_table')  #note: 1. <null> values will be retained using .gdb, will be converted to 0 using .dbf; 2. domain values, if there are any, will be retained by using .gdb
                data_array = arcpy.da.TableToNumPyArray(os.path.join(config.temp_gdb,'soil_table'), '*', null_value = -99)
                
        for i in range (0, len(seq_array)):
            map_unit_data = {}
            mukey = seq_array['FIRST_MUKE'][i]   #note the column name in the .dbf output was cut off
            # arcpy.AddMessage('      - multiple pages map unit ' + str(i))
            # arcpy.AddMessage('      - musym is ' + str(seq_array['MUSYM'][i]))
            # arcpy.AddMessage('      - mukey is ' + str(mukey))
            map_unit_data['Seq'] = str(i+1)    # note i starts from 0, but we want labels to start from 1

            if (seq_array['MUSYM'][i].upper() == 'NOTCOM'):
                map_unit_data['Map Unit Name'] = 'No Digital Data Available'
                map_unit_data['Mukey'] = mukey
                map_unit_data['Musym'] = 'NOTCOM'
            else:
                if 'data_array' not in locals():           #there is only one special polygon(urban land or water)
                    cursor = arcpy.SearchCursor(stable_muaggatt, "mukey = '" + str(mukey) + "'")
                    row = cursor.next()
                    map_unit_data['Map Unit Name'] = row.muname
                    # arcpy.AddMessage('      -  map unit name: ' + row.muname)
                    map_unit_data['Mukey'] = mukey          #note
                    map_unit_data['Musym'] = row.musym
                    row = None
                    cursor = None

                elif ((utility.return_map_unit_attribute(data_array, mukey, 'muname')).upper() == '?'):  #Water or Unrban Land
                    cursor = arcpy.SearchCursor(stable_muaggatt, "mukey = '" + str(mukey) + "'")
                    row = cursor.next()
                    map_unit_data['Map Unit Name'] = row.muname
                    # arcpy.AddMessage('      -  map unit name: ' + row.muname)
                    map_unit_data['Mukey'] = mukey          #note
                    map_unit_data['Musym'] = row.musym
                    row = None
                    cursor = None
                else:
                    map_unit_data['Mukey'] = utility.return_map_unit_attribute(data_array, mukey, 'mukey')
                    map_unit_data['Musym'] = utility.return_map_unit_attribute(data_array, mukey, 'musym')
                    map_unit_data['Map Unit Name'] = utility.return_map_unit_attribute(data_array, mukey, 'muname')
                    map_unit_data['Drainage Class - Dominant'] = utility.return_map_unit_attribute(data_array, mukey, 'drclassdcd')
                    map_unit_data['Hydrologic Group - Dominant'] = utility.return_map_unit_attribute(data_array, mukey, 'hydgrpdcd')
                    map_unit_data['Hydric Classification - Presence'] = utility.return_map_unit_attribute(data_array, mukey, 'hydclprs')
                    map_unit_data['Bedrock Depth - Min'] = utility.return_map_unit_attribute(data_array, mukey, 'brockdepmin')
                    map_unit_data['Watertable Depth - Annual Min'] = utility.return_map_unit_attribute(data_array, mukey, 'wtdepannmin')

                    component_data = utility.return_componen_attribute_rv_indicator_Y(data_array,mukey)
                    map_unit_data['component'] = component_data
            map_unit_data["Soil_Percent"]  ="%s"%round(seq_array['SUM_Shape_'][i]/sum(seq_array['SUM_Shape_'])*100,2)+r'%'
            config.report_data.append(map_unit_data)
        # create the map
        create_map(order_obj)
        
    end = timeit.default_timer()
    arcpy.AddMessage((' -- End generating PSR soil report. Duration:', round(end -start,4)))