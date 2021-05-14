from imp import reload
import arcpy, os, sys
import glob, urllib
import timeit
import shutil
import psr_config as config
file_path = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(1,os.path.join(os.path.dirname(file_path),'DB_Framework'))
import models
reload(sys)

class Kml_Config:
    viewer_dir_kml = None
    viewer_temp = None
    viewer_dir_topo = None
    viewer_dir_relief = None
    
def wetland_to_kml(order_obj):
    wetland_clip = os.path.join(config.scratch_folder, "wetland_clip.shp")
    wetland_mxd_path = os.path.join(config.scratch_folder,'mxd_wetland.mxd')
    if os.path.exists(wetland_mxd_path):
        wetland_mxd = arcpy.mapping.MapDocument(wetland_mxd_path)
        df = arcpy.mapping.ListDataFrames(wetland_mxd,"big")[0]    # the spatial reference here is UTM zone #, need to change to WGS84 Web Mercator
        df.spatialReference = order_obj.spatial_ref_gcs
        if order_obj.province == 'AK':
            df.spatialReference = config.spatial_ref_mercator
        #re-focus using Buffer layer for multipage
        if config.if_multi_page == True:
            buffer_layer = arcpy.mapping.ListLayers(wetland_mxd, "Buffer", df)[0]
            df.extent = buffer_layer.getSelectedExtent(False)
            df.scale = df.scale * 1.1
        #df.spatialReference is currently UTM. dfAsFeature is a feature, not even a layer
        df_as_feature = arcpy.Polygon(arcpy.Array([df.extent.lowerLeft, df.extent.lowerRight, df.extent.upperRight, df.extent.upperLeft]), df.spatialReference)
        del df, wetland_mxd
        wetland_boudnary = os.path.join(config.scratch_folder,"wetland_kml_extend.shp")
        arcpy.Project_management(df_as_feature, wetland_boudnary, order_obj.spatial_ref_gcs)
        arcpy.Clip_analysis(config.data_lyr_wetland, wetland_boudnary, wetland_clip)
        del df_as_feature
        
        wetland_clip_final = None
        if int(arcpy.GetCount_management(wetland_clip).getOutput(0)) != 0:
            arcpy.AddField_management(wetland_boudnary,"TYPE", "TEXT", "", "", "15", "", "NULLABLE", "NON_REQUIRED", "")
            wetland_clip_final = os.path.join(config.scratch_folder, "wetland_clip_final.shp")
            arcpy.Union_analysis([wetland_clip,wetland_boudnary],wetland_clip_final)

            keepFieldList = ("TYPE")
            fieldInfo = ""
            fieldList = arcpy.ListFields(wetland_clip_final)
            for field in fieldList:
                if field.name in keepFieldList:
                    if field.name == 'TYPE':
                        fieldInfo = fieldInfo + field.name + " " + "Wetland Type" + " VISIBLE;"
                    else:
                        pass
                else:
                    fieldInfo = fieldInfo + field.name + " " + field.name + " HIDDEN;"

            arcpy.MakeFeatureLayer_management(wetland_clip_final, 'wetland_clip_lyr', "", "", fieldInfo[:-1])
            arcpy.ApplySymbologyFromLayer_management('wetland_clip_lyr', config.data_lyr_wetland)
            arcpy.LayerToKML_conversion('wetland_clip_lyr', os.path.join(Kml_Config.viewer_dir_kml,"wetland.kmz"))
            arcpy.AddMessage('      -- Create Wetland kmz map: %s' % os.path.join(Kml_Config.viewer_dir_kml,"wetland.kmz"))
            arcpy.Delete_management('wetland_clip_lyr')
        else:
            arcpy.AddMessage('      -- no wetland data')
            arcpy.MakeFeatureLayer_management(wetland_clip, 'wetland_clip_lyr')
            arcpy.LayerToKML_conversion('wetland_clip_lyr', os.path.join(Kml_Config.viewer_dir_kml,"wetland_nodata.kmz"))
            arcpy.AddMessage('      -- Create Wetland kmz map: %s' % os.path.join(Kml_Config.viewer_dir_kml,"wetland_nodata.kmz"))
            arcpy.Delete_management('wetland_clip_lyr')
    else:
        arcpy.AddMessage('  -- Wetland report is not generatated therfore the wetland kml file cannot be exported.')

def wetland_ny_to_kml(order_obj):
    wetland_ny_clip = os.path.join(config.scratch_folder, "wetland_ny_clip")
    wetland_ny_mxd_path = os.path.join(config.scratch_folder,'mxd_wetlandNY.mxd')
    if os.path.exists(wetland_ny_mxd_path):
        wetland_ny_mxd = arcpy.mapping.MapDocument(wetland_ny_mxd_path)
        df = arcpy.mapping.ListDataFrames(wetland_ny_mxd,"big")[0]    # the spatial reference here is UTM zone #, need to change to WGS84 Web Mercator
        df.spatialReference = order_obj.spatial_ref_gcs
        if config.if_multi_page == True:
            buffer_layer = arcpy.mapping.ListLayers(wetland_ny_mxd, "Buffer", df)[0]
            df.extent = buffer_layer.getSelectedExtent(False)
            df.scale = df.scale * 1.1
            
        wetland_boudnary = os.path.join(config.scratch_folder,"wetland_kml_extend.shp")
        arcpy.Clip_analysis(config.data_lyr_wetland_ny_kml, wetland_boudnary, wetland_ny_clip)
        del df, wetland_ny_mxd
        
        wetland_clip_ny_final = None
        if int(arcpy.GetCount_management(wetland_ny_clip).getOutput(0)) != 0:
            wetland_clip_ny_final = os.path.join(config.scratch_folder, "wetland_clip_ny_final.shp")
            arcpy.Union_analysis([wetland_ny_clip,wetland_boudnary],wetland_clip_ny_final)

            keep_field_list = ("CLASS")
            field_info = ""
            field_list = arcpy.ListFields(wetland_clip_ny_final)
            for field in field_list:
                if field.name in keep_field_list:
                    if field.name == 'CLASS':
                        field_info = field_list + field.name + " " + "Wetland CLASS" + " VISIBLE;"
                    else:
                        pass
                else:
                    field_info = field_info + field.name + " " + field.name + " HIDDEN;"

            arcpy.MakeFeatureLayer_management(wetland_clip_ny_final, 'wetland_ny_clip_lyr', "", "", field_info[:-1])
            arcpy.ApplySymbologyFromLayer_management('wetland_ny_clip_lyr', config.data_lyr_wetland_ny_kml)
            arcpy.LayerToKML_conversion('wetland_ny_clip_lyr', os.path.join(Kml_Config.viewer_dir_kml,"w_NY_wetland.kmz"))
            arcpy.AddMessage('      -- Create Wetland NY kmz map: %s' % os.path.join(Kml_Config.viewer_dir_kml,"wetland_nodata.kmz"))
            arcpy.Delete_management('wetland_ny_clip_lyr')
        else:
            arcpy.AddMessage('no wetland data, no kml to folder')
            arcpy.MakeFeatureLayer_management(wetland_ny_clip, 'wetland_ny_clip_lyr')
            arcpy.LayerToKML_conversion('wetland_ny_clip_lyr', os.path.joinKml_config.viewer_dir_kml,"w_NY_wetland_nodata.kmz")
            arcpy.AddMessage('      -- Create Wetland NY kmz map: %s' % os.path.join(Kml_Config.viewer_dir_kml,"wetland_nodata.kmz"))
            arcpy.Delete_management('wetland_ny_clip_lyr')
        ### APA Wetland
        wetland_ny_apa_clip = os.path.join(config.scratch_folder, "wetland_ny_apa_clip")
        arcpy.Clip_analysis(config.data_lyr_wetland_ny_apa_kml, wetland_boudnary, wetland_ny_apa_clip)
        
        if int(arcpy.GetCount_management(wetland_ny_apa_clip).getOutput(0)) != 0:
            wetland_ny_apa_clip_final = os.path.join(config.scratch_folder, "wetland_ny_apa_clip_final")
            arcpy.Union_analysis([wetland_ny_apa_clip,wetland_boudnary],wetland_ny_apa_clip_final)

            keep_field_list = ("ERIS_WTLD")
            field_info = ""
            field_list = arcpy.ListFields(wetland_ny_apa_clip)
            for field in field_list:
                    if field.name in keep_field_list:
                        if field.name == 'ERIS_WTLD':
                            field_info = field_info + field.name + " " + "Wetland CLASS" + " VISIBLE;"
                        else:
                            pass
                    else:
                        field_info = field_info + field.name + " " + field.name + " HIDDEN;"
            arcpy.MakeFeatureLayer_management(wetland_ny_apa_clip_final, 'wetland_ny_apa_clip_final_lyr', "", "", field_info[:-1])
            arcpy.ApplySymbologyFromLayer_management('wetland_ny_apa_clip_final_lyr', config.data_lyr_wetland_ny_apa_kml)
            arcpy.LayerToKML_conversion('wetland_ny_apa_clip_final_lyr', os.path.join(Kml_Config.viewer_dir_kml,"w_apa_wetland.kmz"))
            arcpy.AddMessage('      -- Create APA Wetland kmz map: %s' % os.path.join(Kml_Config.viewer_dir_kml,"w_apa_wetland.kmz"))
            arcpy.Delete_management('wetland_ny_apa_clip_final_lyr')
        else:
            arcpy.AddMessage('no wetland data, no kml to folder')
            arcpy.MakeFeatureLayer_management(wetland_ny_apa_clip, 'wetland_ny_apa_clip_lyr')
            arcpy.LayerToKML_conversion('wetland_ny_apa_clip_lyr', os.path.join(Kml_Config.viewer_dir_kml,"w_apa_wetland_nodata.kmz"))
            arcpy.AddMessage('      -- Create APA Wetland kmz map: %s' % os.path.join(Kml_Config.viewer_dir_kml,"w_apa_wetland_nodata.kmz"))
            arcpy.Delete_management('wetland_ny_apa_clip_lyr')
    else:
        arcpy.AddMessage('  -- Wetland NY report is not generatated therfore the wetland NY kml file cannot be exported.')

def flood_to_kml(order_obj):
    if os.path.exists(config.flood_selectedby_order_shp):
        if int(arcpy.GetCount_management(config.flood_selectedby_order_shp).getOutput(0)) != 0:
            arcpy.AddField_management(config.flood_selectedby_order_shp, "CLASS", "TEXT", "", "", "15", "", "NULLABLE", "NON_REQUIRED", "")
            arcpy.AddField_management(config.flood_selectedby_order_shp,"ERISBIID", "TEXT", "", "", "15", "", "NULLABLE", "NON_REQUIRED", "")
            rows = arcpy.UpdateCursor(config.flood_selectedby_order_shp)
            for row in rows:
                row.CLASS = row.ERIS_CLASS
                ID = [id[1] for id in config.flood_ids if row.ERIS_CLASS==id[0]]
                if ID !=[]:
                    row.ERISBIID = ID[0]
                    rows.updateRow(row)
                rows.updateRow(row)
            del rows
            keep_field_list = ("ERISBIID","CLASS", "FLD_ZONE","ZONE_SUBTY")
            field_info = ""
            field_list = arcpy.ListFields(config.flood_selectedby_order_shp)
            for field in field_list:
                if field.name in keep_field_list:
                    if field.name =='ERISBIID':
                        field_info = field_info + field.name + " " + "ERISBIID" + " VISIBLE;"
                    elif field.name == 'CLASS':
                        field_info = field_info + field.name + " " + "Flood Zone Label" + " VISIBLE;"
                    elif field.name == 'FLD_ZONE':
                        field_info = field_info + field.name + " " + "Flood Zone" + " VISIBLE;"
                    elif field.name == 'ZONE_SUBTY':
                        field_info = field_info + field.name + " " + "Zone Subtype" + " VISIBLE;"
                    else:
                        pass
                else:
                    field_info = field_info + field.name + " " + field.name + " HIDDEN;"
            arcpy.MakeFeatureLayer_management(config.flood_selectedby_order_shp, 'selected_flood_data_lyr', "", "", field_info[:-1])
            arcpy.ApplySymbologyFromLayer_management('selected_flood_data_lyr', config.data_lyr_flood)
            arcpy.LayerToKML_conversion('selected_flood_data_lyr', os.path.join(Kml_Config.viewer_dir_kml,"flood.kmz"))
            arcpy.AddMessage('      -- Create flood kmz map: %s' % os.path.join(Kml_Config.viewer_dir_kml,"flood.kmz"))
            arcpy.Delete_management('selected_flood_data_lyr')
        else:
            arcpy.AddMessage('no flood data to kml')
            arcpy.MakeFeatureLayer_management(config.flood_selectedby_order_shp, 'selected_flood_data_lyr')
            arcpy.LayerToKML_conversion('selected_flood_data_lyr', os.path.join(Kml_Config.viewer_dir_kml,"flood_nodata.kmz"))
            arcpy.AddMessage('      -- Create flood kmz map: %s' % os.path.join(Kml_Config.viewer_dir_kml,"flood_nodata.kmz"))
            arcpy.Delete_management('selected_flood_data_lyr')
    else:
        arcpy.AddMessage('  -- Flood report is not generatated therfore the flood kml file cannot be exported.')
        
def geology_to_kml(order_obj):    
    geology_clip = os.path.join(config.scratch_folder, "geology_clip.shp")
    geology_mxd_path = os.path.join(config.scratch_folder,'mxd_geology.mxd')
    if os.path.exists(geology_mxd_path):
        geology_mxd = arcpy.mapping.MapDocument(geology_mxd_path)
        df = arcpy.mapping.ListDataFrames(geology_mxd,"*")[0]    # the spatial reference here is UTM zone #, need to change to WGS84 Web Mercator
        df.spatialReference = order_obj.spatial_ref_gcs
        if order_obj.province == 'AK':
            df.spatialReference = config.spatial_ref_mercator
        #re-focus using Buffer layer for multipage
        if config.if_multi_page == True:
            buffer_layer = arcpy.mapping.ListLayers(geology_mxd, "Buffer", df)[0]
            df.extent = buffer_layer.getSelectedExtent(False)
            df.scale = df.scale * 1.1
        df_as_feature = arcpy.Polygon(arcpy.Array([df.extent.lowerLeft, df.extent.lowerRight, df.extent.upperRight, df.extent.upperLeft]), df.spatialReference)
        del df, geology_mxd
        
        arcpy.Project_management(df_as_feature, os.path.join(Kml_Config.viewer_temp,"Extent_geol_WGS84.shp"), order_obj.spatial_ref_gcs)
        arcpy.Clip_analysis(config.data_lyr_geology, os.path.join(Kml_Config.viewer_temp,"Extent_geol_WGS84.shp"), geology_clip)
        del df_as_feature
        
        if int(arcpy.GetCount_management(geology_clip).getOutput(0)) != 0:
            arcpy.AddField_management(geology_clip,"ERISBIID", "TEXT", "", "", "15", "", "NULLABLE", "NON_REQUIRED", "")
            rows = arcpy.UpdateCursor(geology_clip)
            for row in rows:
                id_final = [id[1] for id in config.geology_ids if row.ERIS_KEY==id[0]]
                if id_final !=[]:
                    row.ERISBIID = id_final[0]
                    rows.updateRow(row)
            del rows
            keep_field_list = ("ERISBIID","ORIG_LABEL", "UNIT_NAME", "UNIT_AGE","ROCKTYPE1", "ROCKTYPE2", "UNITDESC")
            field_info = ""
            field_list = arcpy.ListFields(geology_clip)
            for field in field_list:
                if field.name in keep_field_list:
                    if field.name =='ERISBIID':
                        field_info = field_info + field.name + " " + "ERISBIID" + " VISIBLE;"
                    elif field.name == 'ORIG_LABEL':
                        field_info = field_info + field.name + " " + "Geologic_Unit" + " VISIBLE;"
                    elif field.name == 'UNIT_NAME':
                        field_info = field_info + field.name + " " + "Name" + " VISIBLE;"
                    elif field.name == 'UNIT_AGE':
                        field_info = field_info + field.name + " " + "Age" + " VISIBLE;"
                    elif field.name == 'ROCKTYPE1':
                        field_info = field_info + field.name + " " + "Primary_Rock_Type" + " VISIBLE;"
                    elif field.name == 'ROCKTYPE2':
                        field_info = field_info + field.name + " " + "Secondary_Rock_Type" + " VISIBLE;"
                    elif field.name == 'UNITDESC':
                        field_info = field_info + field.name + " " + "Unit_Description" + " VISIBLE;"
                    else:
                        pass
                else:
                    field_info = field_info + field.name + " " + field.name + " HIDDEN;"
            arcpy.MakeFeatureLayer_management(geology_clip, 'geologyclip_lyr', "", "", field_info[:-1])
            arcpy.ApplySymbologyFromLayer_management('geologyclip_lyr', config.data_lyr_geology)
            arcpy.LayerToKML_conversion('geologyclip_lyr', os.path.join(Kml_Config.viewer_dir_kml,"geology.kmz"))
            arcpy.AddMessage('      -- Create geology kmz map: %s' % os.path.join(Kml_Config.viewer_dir_kml,"geology.kmz"))
            arcpy.Delete_management('geologyclip_lyr')
        else:
            # print "no geology data to kml"
            arcpy.MakeFeatureLayer_management(geology_clip, 'geologyclip_lyr')
            arcpy.LayerToKML_conversion('geology_clip_lyr', os.path.join(Kml_Config.viewer_dir_kml,"geology_nodata.kmz"))
            arcpy.AddMessage('      -- Create geology kmz map: %s' % os.path.join(Kml_Config.viewer_dir_kml,"geology_nodata.kmz"))
            arcpy.Delete_management('geology_clip_lyr')
    else:
        arcpy.AddMessage('  -- Geology report is not generatated therfore the geology kml file cannot be exported.')

def soil_to_kml(order_obj):
    soil_clip = os.path.join(config.scratch_folder,'temp.gdb', "soil_clip")
    soil_mxd_path = os.path.join(config.scratch_folder,'mxd_soil.mxd')
    if os.path.exists(soil_mxd_path):
        soil_mxd = arcpy.mapping.MapDocument(soil_mxd_path)
        df = arcpy.mapping.ListDataFrames(soil_mxd,"*")[0]    # the spatial reference here is UTM zone #, need to change to WGS84 Web Mercator
        df.spatialReference = order_obj.spatial_ref_gcs
        if order_obj.province == 'AK':
            df.spatialReference = config.spatial_ref_mercator
        #re-focus using Buffer layer for multipage
        if config.if_multi_page == True:
            buffer_layer = arcpy.mapping.ListLayers(soil_mxd, "Buffer", df)[0]
            df.extent = buffer_layer.getSelectedExtent(False)
            df.scale = df.scale * 1.1
        df_as_feature = arcpy.Polygon(arcpy.Array([df.extent.lowerLeft, df.extent.lowerRight, df.extent.upperRight, df.extent.upperLeft]), df.spatialReference)
        del df, soil_mxd
        
        arcpy.Project_management(df_as_feature, os.path.join(Kml_Config.viewer_temp,"Extent_soil_WGS84.shp"), order_obj.spatial_ref_gcs)
        arcpy.Clip_analysis(os.path.join(config.data_path_soil,'MUPOLYGON'), os.path.join(Kml_Config.viewer_temp,"Extent_soil_WGS84.shp"), soil_clip)
        del df_as_feature
        
        if int(arcpy.GetCount_management(soil_clip).getOutput(0)) != 0:
            arcpy.AddField_management(soil_clip, "Map_Unit", "TEXT", "", "", "1500", "", "NULLABLE", "NON_REQUIRED", "")
            arcpy.AddField_management(soil_clip, "Map_Unit_Name", "TEXT", "", "", "1500", "", "NULLABLE", "NON_REQUIRED", "")
            arcpy.AddField_management(soil_clip, "Dominant_Drainage_Class", "TEXT", "", "", "1500", "", "NULLABLE", "NON_REQUIRED", "")
            arcpy.AddField_management(soil_clip, "Dominant_Hydrologic_Group", "TEXT", "", "", "1500", "", "NULLABLE", "NON_REQUIRED", "")
            arcpy.AddField_management(soil_clip, "Presence_Hydric_Classification", "TEXT", "", "", "1500", "", "NULLABLE", "NON_REQUIRED", "")
            arcpy.AddField_management(soil_clip, "Min_Bedrock_Depth", "TEXT", "", "", "1500", "", "NULLABLE", "NON_REQUIRED", "")
            arcpy.AddField_management(soil_clip, "Annual_Min_Watertable_Depth", "TEXT", "", "", "1500", "", "NULLABLE", "NON_REQUIRED", "")
            arcpy.AddField_management(soil_clip, "component", "TEXT", "", "", "2500", "", "NULLABLE", "NON_REQUIRED", "")
            arcpy.AddField_management(soil_clip,"ERISBIID", "TEXT", "", "", "15", "", "NULLABLE", "NON_REQUIRED", "")
            rows = arcpy.UpdateCursor(soil_clip)
            for row in rows:
                for map_unit in config.report_data:
                    if row.musym == map_unit["Musym"]:
                        id_final = [id[1] for id in config.soil_ids if row.MUSYM==id[0]]
                        if id_final !=[]:
                            row.ERISBIID = id_final[0]
                            rows.updateRow(row)
                        for key in map_unit.keys():
                            if key =="Musym":
                                row.Map_Unit = map_unit[key]
                            elif key == "Map Unit Name":
                                row.Map_Unit_Name = map_unit[key]
                            elif key == "Bedrock Depth - Min":
                                row.Min_Bedrock_Depth = map_unit[key]
                            elif key =="Drainage Class - Dominant":
                                row.Dominant_Drainage_Class = map_unit[key]
                            elif key =="Hydric Classification - Presence":
                                row.Presence_Hydric_Classification = map_unit[key]
                            elif key =="Hydrologic Group - Dominant":
                                row.Dominant_Hydrologic_Group = map_unit[key]
                            elif key =="Watertable Depth - Annual Min":
                                row.Annual_Min_Watertable_Depth = map_unit[key]
                            elif key =="component":
                                new = ''
                                component = map_unit[key]
                                for i in range(len(component)):
                                    for j in range(len(component[i])):
                                        for k in range(len(component[i][j])):
                                            new = new+component[i][j][k]+" "
                                row.component = new
                            else:
                                pass
                            rows.updateRow(row)
            del rows
            keep_field_list = ("ERISBIID","Map_Unit", "Map_Unit_Name", "Dominant_Drainage_Class","Dominant_Hydrologic_Group", "Presence_Hydric_Classification", "Min_Bedrock_Depth","Annual_Min_Watertable_Depth","component")
            field_info = ""
            fieldList = arcpy.ListFields(soil_clip)
            for field in fieldList:
                if field.name in keep_field_list:
                    if field.name =='ERISBIID':
                        field_info = field_info + field.name + " " + "ERISBIID" + " VISIBLE;"
                    elif field.name == 'Map_Unit':
                        field_info = field_info + field.name + " " + "Map_Unit" + " VISIBLE;"
                    elif field.name == 'Map_Unit_Name':
                        field_info = field_info + field.name + " " + "Map_Unit_Name" + " VISIBLE;"
                    elif field.name == 'Dominant_Drainage_Class':
                        field_info = field_info + field.name + " " + "Dominant_Drainage_Class" + " VISIBLE;"
                    elif field.name == 'Dominant_Hydrologic_Group':
                        field_info = field_info + field.name + " " + "Dominant_Hydrologic_Group" + " VISIBLE;"
                    elif field.name == 'Presence_Hydric_Classification':
                        field_info = field_info + field.name + " " + "Presence_Hydric_Classification" + " VISIBLE;"
                    elif field.name == 'Min_Bedrock_Depth':
                        field_info = field_info + field.name + " " + "Min_Bedrock_Depth" + " VISIBLE;"
                    elif field.name == 'Annual_Min_Watertable_Depth':
                        field_info = field_info + field.name + " " + "Annual_Min_Watertable_Depth" + " VISIBLE;"
                    elif field.name == 'component':
                        field_info = field_info + field.name + " " + "component" + " VISIBLE;"
                    else:
                        pass
                else:
                    field_info = field_info + field.name + " " + field.name + " HIDDEN;"
            arcpy.MakeFeatureLayer_management(soil_clip, 'soil_clip_lyr',"", "", field_info[:-1])
            soil_symbol_local = os.path.join(config.scratch_folder,"soil_lyr_local.lyr")
            arcpy.SaveToLayerFile_management(config.soil_lyr,soil_symbol_local[:-4])
            arcpy.ApplySymbologyFromLayer_management('soil_clip_lyr', soil_symbol_local)
            arcpy.LayerToKML_conversion('soil_clip_lyr', os.path.join(Kml_Config.viewer_dir_kml,"soil_clip.kmz"))
            arcpy.Delete_management('soil_clip_lyr')
        else:
            arcpy.AddMessage('no soil data to kml')
            arcpy.MakeFeatureLayer_management(soil_clip, 'soil_clip_lyr')
            arcpy.LayerToKML_conversion('soil_clip_lyr', os.path.join(Kml_Config.viewer_dir_kml,"soil_clip_nodata.kmz"))
            arcpy.AddMessage('      -- Create soil kmz map: %s' % os.path.join(Kml_Config.viewer_dir_kml,"soil_clip_nodata.kmz"))
            arcpy.Delete_management('soil_clip_lyr')
    else:
        arcpy.AddMessage('  -- Soil report is not generatated therfore the soil kml file cannot be exported.')

def clip_contour_lines(order_obj):
    contour_clip = os.path.join(config.scratch_folder, "contour_clip.shp")
    arcpy.Clip_analysis(config.data_lyr_contour,config.topo_frame, contour_clip)

    if int(arcpy.GetCount_management(contour_clip).getOutput(0)) != 0:
        keep_field_list = ("CONTOURELE")
        field_info = ""
        field_list = arcpy.ListFields(contour_clip)
        for field in field_list:
            if field.name in keep_field_list:
                if field.name == 'CONTOURELE':
                    field_info = field_info + field.name + " " + "elevation" + " VISIBLE;"
                else:
                    pass
            else:
                field_info = field_info + field.name + " " + field.name + " HIDDEN;"
        arcpy.MakeFeatureLayer_management(contour_clip, 'contour_clip_lyr', "", "", field_info[:-1])
        arcpy.ApplySymbologyFromLayer_management('contour_clip_lyr', config.data_lyr_contour)
        arcpy.LayerToKML_conversion('contour_clip_lyr', os.path.join(Kml_Config.viewer_dir_relief,"contour_clip.kmz"))
        arcpy.AddMessage('      -- Create contour kmz map: %s' % os.path.join(Kml_Config.viewer_dir_relief,"contour_clip.kmz"))
        arcpy.Delete_management('contour_clip_lyr')
    else:
        arcpy.AddMessage('no contour data, no kml to folder')
        arcpy.MakeFeatureLayer_management(contour_clip, 'contour_clip_lyr')
        arcpy.LayerToKML_conversion('contour_clip_lyr', os.path.join(Kml_Config.viewer_dir_relief,'contour_clip_nodata.kmz'))
        arcpy.AddMessage('      -- Create contour kmz map: %s' % os.path.join(Kml_Config.viewer_dir_relief,'contour_clip_nodata.kmz'))
        arcpy.Delete_management('contour_clip_lyr')
        
    if os.path.exists(os.path.join(config.viewer_path, order_obj.number +"_psr_kml")):
        shutil.rmtree(os.path.join(config.viewer_path, order_obj.number +"_psr_kml"))
    shutil.copytree(os.path.join(config.scratch_folder, order_obj.number +"_psr_kml"), os.path.join(config.viewer_path, order_obj.number+"_psr_kml"))
    url = config.upload_link + 'PSRKMLUpload?ordernumber=' + order_obj.number
    urllib.urlopen(url)
    
    if os.path.exists(os.path.join(config.viewer_path, order_obj.number + "_psr_topo")):
            shutil.rmtree(os.path.join(config.viewer_path, order_obj.number + "_psr_topo"))
    shutil.copytree(os.path.join(config.scratch_folder, order_obj.number + "_psr_topo"), os.path.join(config.viewer_path, order_obj.number + "_psr_topo"))
    url = config.upload_link + "PSRTOPOUpload?ordernumber=" + order_obj.number
    urllib.urlopen(url)
    
    if os.path.exists(os.path.join(config.viewer_path, order_obj.number + '_psr_relief')):
        shutil.rmtree(os.path.join(config.viewer_path, order_obj.number + '_psr_relief'))
    shutil.copytree(os.path.join(config.scratch_folder, order_obj.number + '_psr_relief'), os.path.join(config.viewer_path, order_obj.number + '_psr_relief'))
    url = config.upload_link + "ReliefUpload?ordernumber=" + order_obj.number
    urllib.urlopen(url)
    
def topo_to_kml(order_obj):
    topo_mxd_path = os.path.join(config.scratch_folder,'mxd_topo.mxd')
    if os.path.exists(topo_mxd_path):
        topo_mxd = arcpy.mapping.MapDocument(topo_mxd_path)
        df = arcpy.mapping.ListDataFrames(topo_mxd,"*")[0]    # the spatial reference here is UTM zone #, need to change to WGS84 Web Mercator
        df.spatialReference = order_obj.spatial_ref_gcs
        if order_obj.province == 'AK':
            df.spatialReference = config.spatial_ref_mercator
        #re-focus using Buffer layer for multipage
        if config.if_multi_page == True:
            buffer_layer = arcpy.mapping.ListLayers(topo_mxd, "Buffer", df)[0]
            df.extent = buffer_layer.getSelectedExtent(False)
            df.scale = df.scale * 1.1
        df_as_feature = arcpy.Polygon(arcpy.Array([df.extent.lowerLeft, df.extent.lowerRight, df.extent.upperRight, df.extent.upperLeft]), df.spatialReference)
        del df, topo_mxd
        
        arcpy.Project_management(df_as_feature, os.path.join(Kml_Config.viewer_temp,"Extent_topo_WGS84.shp"), order_obj.spatial_ref_gcs)
        del df_as_feature
        
        n = 0
        mosaics = []
        for item in glob.glob(os.path.join(config.scratch_folder,'*_TM_geo.tif')):
            try:
                arcpy.Clip_management(item,"",os.path.join(Kml_Config.viewer_temp, "topo"+str(n)+".jpg"),os.path.join(Kml_Config.viewer_temp,"Extent_topo_WGS84.shp"),"255","ClippingGeometry")
                mosaics.append(os.path.join(Kml_Config.viewer_temp, "topo"+str(n)+".jpg"))
                n = n+1
            except Exception as e:
                arcpy.AddMessage(str(e) + item )  #possibly not in the clip_frame
      
        if len(mosaics) != 0:
            raster = mosaics[0]
            raster_desc = arcpy.Describe(raster)
            raster_band = raster_desc.bandCount
            arcpy.MosaicToNewRaster_management(mosaics, Kml_Config.viewer_dir_relief,config.relief_image_name,config.spatial_ref_mercator,"","",raster_band,"MINIMUM","MATCH")
            desc = arcpy.Describe(os.path.join(Kml_Config.viewer_dir_relief, config.relief_image_name))
            boundary_feature = arcpy.Polygon(arcpy.Array([desc.extent.lowerLeft, desc.extent.lowerRight, desc.extent.upperRight, desc.extent.upperLeft]),desc.spatialReference)
            del desc
            if 'year' not in locals():
                year = '0'
            temp_fc = os.path.join(config.scratch_folder, "img_bnd_"+str(year)+ ".shp")

            arcpy.Project_management(boundary_feature, temp_fc, order_obj.spatial_ref_gcs) #function requires output not be in_memory
            del boundary_feature
            desc = arcpy.Describe(temp_fc)
            meta_item = {}
            meta_item['type'] = 'psrrelief'
            meta_item['imagename'] = config.relief_image_name

            meta_item['lat_sw'] = desc.extent.YMin
            meta_item['long_sw'] = desc.extent.XMin
            meta_item['lat_ne'] = desc.extent.YMax
            meta_item['long_ne'] = desc.extent.XMax
            ### update DB
            overlay_image_obj = models.Overlay_Image(order_obj, meta_item)
            overlay_image_obj.delete()
            overlay_image_obj.insert()
            ### Extract Contour data from topo frame
            clip_contour_lines(order_obj)
    else:
        arcpy.AddMessage('  -- Topo report is not generatated therfore the topo kml file cannot be exported.')    

def convert_to_kml(order_obj):
    
    arcpy.AddMessage('  -- Start generating kml for explorer viewer...')
    start = timeit.default_timer() 
    
    Kml_Config.viewer_dir_kml = os.path.join(config.scratch_folder,order_obj.number + '_psr_kml')
    if not os.path.exists(Kml_Config.viewer_dir_kml):
        os.mkdir(Kml_Config.viewer_dir_kml)
        
    Kml_Config.viewer_temp =os.path.join(config.scratch_folder,'viewer_temp')
    if not os.path.exists(Kml_Config.viewer_temp):
        os.mkdir(Kml_Config.viewer_temp)
        
    Kml_Config.viewer_dir_topo = os.path.join(config.scratch_folder,order_obj.number+'_psr_topo')
    if not os.path.exists(Kml_Config.viewer_dir_topo):
        os.mkdir(Kml_Config.viewer_dir_topo)
    
    Kml_Config.viewer_dir_relief = os.path.join(config.scratch_folder,order_obj.number+'_psr_relief')
    if not os.path.exists(Kml_Config.viewer_dir_relief):
        os.mkdir( Kml_Config.viewer_dir_relief)
    
    ### generate kml for wetland map
    if config.if_wetland_report:
        wetland_to_kml(order_obj)
    ## generate kml for wetland newyork map
    if order_obj.province == 'NY' and config.if_wetland_report:
        wetland_ny_to_kml(order_obj)
    
    ## generate kml(kmz) for flood map
    if config.if_flood_report:
        flood_to_kml(order_obj)
    
    ## generate kml(kmz) for geology map
    if config.if_geology_report:
        geology_to_kml(order_obj)
    
    ## generate kml(kmz) for soil map
    if config.if_soil_report:
        soil_to_kml(order_obj)
    
    ## generate kml(kmz) for topo map
    if config.if_topo_report:
        topo_to_kml(order_obj)
    
    end = timeit.default_timer()
    arcpy.AddMessage((' -- End of generating kml for explorer viewer. Duration:', round(end -start,4)))