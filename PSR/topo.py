
from imp import reload
import arcpy, os, sys
import timeit
import csv
import shutil
import psr_utility as utility
import psr_config as config
import models
sys.path.insert(1,os.path.join(os.getcwd(),'DB_Framework'))
reload(sys)

def generate_topo_report(order_obj):
    arcpy.AddMessage('  -- Start generating PSR topo report...')
    start = timeit.default_timer()  
    ### set scratch folder
    arcpy.env.workspace = config.scratch_folder
    arcpy.env.overwriteOutput = True   
    ### set paths
    output_jpg_topo = config.output_jpg(order_obj,config.Report_Type.topo)

    ### create buffer map based on order geometry
    arcpy.Buffer_analysis(config.order_geometry_pcs_shp, config.order_buffer_shp, config.buffer_dist_topo) 
    
    point = arcpy.Point()
    array = arcpy.Array()
    feature_list = []

    width = arcpy.Describe(config.order_buffer_shp).extent.width/2
    height = arcpy.Describe(config.order_buffer_shp).extent.height/2
    
    if (width/height > 7/7):    # 7/7 now since adjusted the frame to square
        # wider shape
        height = width/7*7
    else:
        # longer shape
        width = height/7*7
    x_centre_oid = (arcpy.Describe(config.order_buffer_shp).extent.XMax + arcpy.Describe(config.order_buffer_shp).extent.XMin)/2
    y_centre_oid = (arcpy.Describe(config.order_buffer_shp).extent.YMax + arcpy.Describe(config.order_buffer_shp).extent.YMin)/2
    
    if config.if_multi_page:
        width = width + 6400     # add 2 miles to each side, for multipage
        height = height + 6400   # add 2 miles to each side, for multipage
        
    point.X = x_centre_oid-width
    point.Y = y_centre_oid+height
    array.add(point)
    point.X = x_centre_oid+width
    point.Y = y_centre_oid+height
    array.add(point)
    point.X = x_centre_oid+width
    point.Y = y_centre_oid-height
    array.add(point)
    point.X = x_centre_oid-width
    point.Y = y_centre_oid-height
    array.add(point)
    point.X = x_centre_oid-width
    point.Y = y_centre_oid+height
    array.add(point)
    feat = arcpy.Polygon(array,order_obj.spatial_ref_pcs)
    array.removeAll()
    feature_list.append(feat)
    arcpy.CopyFeatures_management(feature_list, config.topo_frame)
    arcpy.Project_management(config.topo_frame, config.topo_frame_gcs, order_obj.spatial_ref_gcs)
    
    topo_master_lyr = arcpy.mapping.Layer(config.topo_master_lyr)
    arcpy.SelectLayerByLocation_management(topo_master_lyr,'intersect',config.topo_frame_gcs)
    
    if(int((arcpy.GetCount_management(topo_master_lyr).getOutput(0))) == 0):
        arcpy.AddMessage('      - NO records selected')
        topo_master_lyr = None
    else:
        cell_ids_selected = []
        infomatrix = []
        # loop through the relevant records, locate the selected cell IDs
        rows = arcpy.SearchCursor(topo_master_lyr)    # loop through the selected records
        
        for row in rows:
            cell_id = str(int(row.getValue("CELL_ID")))
            cell_ids_selected.append(cell_id)
        del row
        del rows

        topo_master_lyr = None
    
        with open(config.topo_csv_file, "rb") as f:
            reader = csv.reader(f)
            for row in reader:
                if row[9] in cell_ids_selected:
                    pdf_name = row[15].strip()

                    # for current topos, read the year from the geopdf file name
                    temp_list = pdf_name.split("_")
                    year_to_use = temp_list[len(temp_list)-3][0:4]

                    if year_to_use[0:2] != "20":
                        arcpy.AddMessage('      - Error in the year of the map!!')

                    # arcpy.AddMessage('      -%s' % row[9] + " " + row[5] + "  " + row[15] + "  " + year_to_use)
                    infomatrix.append([row[9],row[5],row[15],year_to_use])

        mxd_topo = arcpy.mapping.MapDocument(config.mxd_file_topo) if order_obj.province !='WA' else arcpy.mapping.MapDocument(config.mxd_file_topo_Tacoma)#mxdfile_topo_Tacoma
        df_topo = arcpy.mapping.ListDataFrames(mxd_topo,"*")[0]
        df_topo.spatialReference = order_obj.spatial_ref_pcs
        mxd_mm_topo = None
        df_mm_topo = None
        if config.if_multi_page:
            mxd_mm_topo = arcpy.mapping.MapDocument(config.mxd_mm_file_topo) if order_obj.province !='WA' else arcpy.mapping.MapDocument(config.mxd_mm_file_topo_Tacoma)
            df_mm_topo = arcpy.mapping.ListDataFrames(mxd_mm_topo,"*")[0]
            df_mm_topo.spatialReference =  order_obj.spatial_ref_pcs

        quadrangles = ''
        for item in infomatrix:
            pdf_name = item[2]
            tif_name = pdf_name[0:-4]   # note without .tif part
            tif_name_bk = tif_name
            year = item[3]
            if os.path.exists(os.path.join(config.topo_tif_dir,tif_name+ "_t.tif")):
                if '.' in tif_name:
                    tif_name = tif_name.replace('.','')

                # need to make a local copy of the tif file for fast data source replacement
                name_comps = tif_name.split('_')
                name_comps.insert(-2,year)
                new_tif_name = '_'.join(name_comps)

                shutil.copyfile(os.path.join(config.topo_tif_dir,tif_name_bk+"_t.tif"),os.path.join(config.scratch_folder,new_tif_name+'.tif'))

                topo_layer = arcpy.mapping.Layer(config.topo_white_lyr_file)
                topo_layer.replaceDataSource(config.scratch_folder, "RASTER_WORKSPACE", new_tif_name)
                topo_layer.name = new_tif_name
                arcpy.mapping.AddLayer(df_topo, topo_layer, "BOTTOM")
                if config.if_multi_page:
                    arcpy.mapping.AddLayer(df_mm_topo, topo_layer, "BOTTOM")

                comps = pdf_name.split('_')
                quad_name = " ".join(comps[1:len(comps)-3])+","+comps[0]

                if quadrangles == '':
                    quadrangles = quad_name
                else:
                    quadrangles = quadrangles + "; " + quad_name
                topo_layer = None

            else:
                arcpy.AddMessage("      - tif file doesn't exist " + tif_name)
                if not os.path.exists(config.topo_tif_dir):
                    arcpy.AddMessage("      - tif dir doesn't exist " + config.topo_tif_dir)
                else:
                    arcpy.AddMessage("      - tif dir does exist " + config.topo_tif_dir)
        # possibly no topo returned. Seen one for EDR Alaska order. = True even for topoLayer = None                

        utility.add_layer_to_mxd("order_buffer",df_topo,config.buffer_lyr_file,1.1)
        utility.add_layer_to_mxd("order_geometry_pcs", df_topo,config.order_geom_lyr_file,1)

        year_text = arcpy.mapping.ListLayoutElements(mxd_topo, "TEXT_ELEMENT", "year")[0]
        year_text.text = "Current USGS Topo"
        year_text.elementPositionX = 0.4959

        quadrangle_text = arcpy.mapping.ListLayoutElements(mxd_topo, "TEXT_ELEMENT", "quadrangle")[0]
        quadrangle_text.text = "Quadrangle(s): " + quadrangles

        source_text = arcpy.mapping.ListLayoutElements(mxd_topo, "TEXT_ELEMENT", "source")[0]
        source_text.text = "Source: USGS 7.5 Minute Topographic Map"

        arcpy.RefreshTOC()
        if not config.if_multi_page :
            arcpy.mapping.ExportToJPEG(mxd_topo, output_jpg_topo, "PAGE_LAYOUT")#, resolution=200, jpeg_quality=90)
            if not os.path.exists(os.path.join(config.report_path, 'PSRmaps', order_obj.number)):
                os.mkdir(os.path.join(config.report_path, 'PSRmaps', order_obj.number))
            shutil.copy(output_jpg_topo, os.path.join(config.report_path, 'PSRmaps', order_obj.number))
            # arcpy.AddMessage('      - output jpg image path: %s' % output_jpg_topo)
            mxd_topo.saveACopy(os.path.join(config.scratch_folder,"mxd_topo.mxd"))
            del mxd_topo
            del df_topo
        else:                           # multipage
            grid_lyr_shp = os.path.join(config.scratch_folder, 'grid_lyr_topo.shp')
            arcpy.GridIndexFeatures_cartography(grid_lyr_shp, config.order_buffer_shp, "", "", "", config.grid_size, config.grid_size)

            # part 1: the overview map
            # add grid layer
            grid_layer = arcpy.mapping.Layer(config.grid_lyr_file)
            grid_layer.replaceDataSource(config.scratch_folder,"SHAPEFILE_WORKSPACE","grid_lyr_topo")
            arcpy.mapping.AddLayer(df_topo,grid_layer,"Top")

            df_topo.extent = grid_layer.getExtent()
            df_topo.scale = df_topo.scale * 1.1

            mxd_topo.saveACopy(os.path.join(config.scratch_folder, "mxd_topo.mxd"))
            arcpy.mapping.ExportToJPEG(mxd_topo, output_jpg_topo, "PAGE_LAYOUT", 480, 640, 150, "False", "24-BIT_TRUE_COLOR", 85)
            if not os.path.exists(os.path.join(config.report_path, 'PSRmaps', order_obj.number)):
                os.mkdir(os.path.join(config.report_path, 'PSRmaps', order_obj.number))
            shutil.copy(output_jpg_topo, os.path.join(config.report_path, 'PSRmaps', order_obj.number))
            arcpy.AddMessage('      - output jpg image : %s' % os.path.join(config.report_path,os.path.basename(output_jpg_topo)))
            del mxd_topo
            del df_topo

            # part 2: the data driven pages
            page = int(arcpy.GetCount_management(grid_lyr_shp).getOutput(0))  + 1
            
            utility.add_layer_to_mxd("order_buffer",df_mm_topo,config.buffer_lyr_file,1.1)
            utility.add_layer_to_mxd("order_geometry_pcs", df_mm_topo,config.order_geom_lyr_file,1)

            grid_layer_mm = arcpy.mapping.ListLayers(mxd_mm_topo,"Grid" ,df_mm_topo)[0]
            grid_layer_mm.replaceDataSource(config.scratch_folder, "SHAPEFILE_WORKSPACE","grid_lyr_topo")
            arcpy.CalculateAdjacentFields_cartography(grid_lyr_shp, "PageNumber")
            mxd_mm_topo.saveACopy(os.path.join(config.scratch_folder, "mxd_mm_topo.mxd"))

            for i in range(1,int(arcpy.GetCount_management(grid_lyr_shp).getOutput(0))+1):
                arcpy.SelectLayerByAttribute_management(grid_layer_mm, "NEW_SELECTION", ' "PageNumber" =  ' + str(i))
                df_mm_topo.extent = grid_layer_mm.getSelectedExtent(True)
                df_mm_topo.scale = df_mm_topo.scale * 1.1

                # might want to select the quad name again
                quadrangles_mm = ""
                images = arcpy.mapping.ListLayers(mxd_mm_topo, "*TM_geo", df_mm_topo)
                for image in images:
                    if image.getExtent().overlaps(grid_layer_mm.getSelectedExtent(True)) or image.getExtent().contains(grid_layer_mm.getSelectedExtent(True)):
                        temp = image.name.split('_20')[0]    # e.g. VA_Port_Royal
                        comps = temp.split('_')
                        quadname = " ".join(comps[1:len(comps)])+","+comps[0]

                        if quadrangles_mm == "":
                            quadrangles_mm = quadname
                        else:
                            quadrangles_mm = quadrangles_mm + "; " + quadname

                arcpy.SelectLayerByAttribute_management(grid_layer_mm, "CLEAR_SELECTION")

                year_text = arcpy.mapping.ListLayoutElements(mxd_mm_topo, "TEXT_ELEMENT", "year")[0]
                year_text.text = "Current USGS Topo - Page " + str(i)
                year_text.elementPositionX = 0.4959

                quadrangle_text = arcpy.mapping.ListLayoutElements(mxd_mm_topo, "TEXT_ELEMENT", "quadrangle")[0]
                quadrangle_text.text = "Quadrangle(s): " + quadrangles_mm

                quadrangle_text = arcpy.mapping.ListLayoutElements(mxd_mm_topo, "TEXT_ELEMENT", "source")[0]
                quadrangle_text.text = "Source: USGS 7.5 Minute Topographic Map"

                arcpy.RefreshTOC()

                arcpy.mapping.ExportToJPEG(mxd_mm_topo, output_jpg_topo[0:-4]+str(i)+".jpg", "PAGE_LAYOUT", 480, 640, 150, "False", "24-BIT_TRUE_COLOR", 85)
                if not os.path.exists(os.path.join(config.report_path, 'PSRmaps', order_obj.number)):
                    os.mkdir(os.path.join(config.report_path, 'PSRmaps', order_obj.number))
                shutil.copy(output_jpg_topo[0:-4]+str(i)+".jpg", os.path.join(config.report_path, 'PSRmaps', order_obj.number))
                # arcpy.AddMessage('      - output jpg image: %s' % os.path.join(config.report_path, 'PSRmaps', order_obj.number, os.path.basename(output_jpg_topo[0:-4]+str(i)+".jpg")))
            del mxd_mm_topo
            del df_mm_topo
            #insert generated .jpg report path into eris_maps_psr table
            psr_obj = models.PSR()
            if os.path.exists(os.path.join(config.report_path, 'PSRmaps', order_obj.number, order_obj.number + '_US_TOPO.jpg')):
                psr_obj.insert_map(order_obj.id, 'RELIEF', order_obj.number + '_US_TOPO.jpg', 1)
                for i in range(1,page):
                    psr_obj.insert_map(order_obj.id, 'TOPO', order_obj.number + '_US_TOPO' + str(i) + '.jpg', i + 1)
            else:
                arcpy.AddMessage('No Relief map is available')
    end = timeit.default_timer()
    arcpy.AddMessage(('-- End generating PSR topo report. Duration:', round(end -start,4)))