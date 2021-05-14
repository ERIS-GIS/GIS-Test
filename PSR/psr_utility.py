import arcpy,os, sys
import psr_config as config
import numpy as np
file_path =os.path.dirname(os.path.abspath(__file__))
sys.path.insert(1,os.path.join(os.path.dirname(file_path),'Elevation'))
import dem_footprints
import dem_imgs_ll
import urllib, contextlib
import json as simple_json
import math

def if_multipage(geometry_pcs_shp, input_report_type = None):
    multi_page = None
    geomExtent = arcpy.Describe(geometry_pcs_shp).extent
    if geomExtent.width > 1300 or geomExtent.height > 1300:
        multi_page = True
    else:
        multi_page = False
    if geomExtent.width > 500 or geomExtent.height > 500 and multi_page == None:
        if input_report_type in [config.Report_Type.topo, config.Report_Type.relief, config.Report_Type.wells]:
            multi_page = True
    return multi_page
def add_layer_to_mxd(layer_name,data_frame,lyr_file, scale):
    layer = arcpy.mapping.Layer(lyr_file)
    layer.replaceDataSource(arcpy.env.scratchFolder,"SHAPEFILE_WORKSPACE",layer_name)
    arcpy.mapping.AddLayer(data_frame,layer,"Top")
    data_frame.extent = layer.getSelectedExtent(False)
    data_frame.scale = data_frame.scale * scale
def set_order_geometry(order_obj):
    ### set order geometry layer file
    if order_obj.geometry.type.lower() in ['point','multipoint']:
        config.order_geom_lyr_file = config.order_geom_lyr_point
    elif order_obj.geometry.type.lower() =='polyline':
        config.order_geom_lyr_file = config.order_geom_lyr_polyline
    else: #polygon
        config.order_geom_lyr_file = config.order_geom_lyr_polygon
    ### calculate order geometry in PCS
    order_geometry_pcs = order_obj.geometry.projectAs(order_obj.spatial_ref_pcs)
    arcpy.CopyFeatures_management(order_obj.geometry, config.order_geometry_gcs_shp)
    arcpy.DefineProjection_management(config.order_geometry_gcs_shp, order_obj.spatial_ref_gcs)
    arcpy.CopyFeatures_management(order_geometry_pcs, config.order_geometry_pcs_shp)
    arcpy.DefineProjection_management(config.order_geometry_gcs_shp, order_obj.spatial_ref_pcs)
def return_unique_setstring_musym(table_name):
    data = arcpy.da.TableToNumPyArray(table_name, ['mukey', 'musym'])
    uniques = np.unique(data[data['musym']!='NOTCOM']['mukey'])
    if len(uniques) == 0:
        return ''
    else:
        output_string = '('
        for item in uniques:
            output_string = output_string + "'" + str(item) + "', "
        output_string = output_string[0:-2] + ")"
        return output_string
def return_unique_set_string(table_name, field_name):
    data = arcpy.da.TableToNumPyArray(table_name, [field_name])
    uniques = np.unique(data[field_name])
    if len(uniques) == 0:
        return ''
    else:
        output_string = '('
        for item in uniques:
            output_string = output_string + "'" + str(item) + "', "
        output_string = output_string[0:-2] + ")"
        return output_string
#check if an array contain the same values
def check_if_unique_value(input_array):
    value = input_array[0]
    for i in range(0,len(input_array)):
        if(input_array[i] != value):
            return False
    return True
def return_map_unit_attribute(data_array, mukey, attribute_name):   #water, urban land is not in data_array, so will return '?'
    data = data_array[data_array['mukey'] == mukey][attribute_name]
    if (len(data) == 0):
        return "?"
    else:
        if(check_if_unique_value):
            if (attribute_name == 'brockdepmin' or attribute_name == 'wtdepannmin'):
                if data[0] == -99:
                    return 'null'
                else:
                    return str(data[0]) + 'cm'
            return str(data[0])  #will convert to str no matter what type
        else:
            return "****ERROR****"
def return_componen_attribute_rv_indicator_Y(data_array,mukey):
    result_array = []
    dataarray1 = data_array[data_array['mukey'] == mukey]
    data = dataarray1[dataarray1['majcompflag'] =='Yes']      # 'majcompfla' needs to be used for .dbf table
    comps = data[['cokey','compname','comppct_r']]
    comps_sorted = np.sort(np.unique(comps), order = 'comppct_r')[::-1]     #[::-1] gives descending order
    for comp in comps_sorted:
        horizon_array = []
        keyname = comp[1] + '('+str(comp[2])+'%)'
        horizon_array.append([keyname])

        selection = data[data['cokey']==comp[0]][['mukey','cokey','compname','comppct_r','hzname','hzdept_r','hzdepb_r','texdesc']]
        selection_sorted = np.sort(selection, order = 'hzdept_r')
        for item in selection_sorted:
            horizon_label = 'horizon ' + item['hzname'] + '(' + str(item['hzdept_r']) + 'cm to '+ str(item['hzdepb_r']) + 'cm)'
            horizon_texture = item['texdesc']
            horizon_array.append([horizon_label,horizon_texture])
        result_array.append(horizon_array)

    return result_array

def find_dem_parameters(master_grids, (x, y)): # one point at once
    temp={}
    for cell in master_grids:
        [x_max,x_min,y_max,y_min] =[max([_[0] for _ in cell[0]]),min([_[0] for _ in cell[0]]),max([_[1] for _ in cell[0]]),min([_[1] for _ in cell[0]])]
        if (x_min < x < x_max and y_min< y < y_max) :
            if cell[2] not in temp.keys():
                temp[cell[2]]=[cell[1]]
            elif cell[1] not in temp[cell[2]]:
                temp[cell[2]].append(cell[1])
    return temp
def get_single_elevation((x,y),dem_list,path,key):
    if dem_list!=[]:
        exec("mch = dem_imgs_ll.mch%s"%(key))
        exec("mcw = dem_imgs_ll.mcw%s"%(key))
        for dem in dem_list:
            dem_path = os.path.join(path,dem)
            ele = arcpy.RasterToNumPyArray(dem_path,arcpy.Point(x,y),1,1)
            if len(ele)==1 and ele[0,0] >-50:
                return ele[0,0]
            del ele,dem_path,dem
    return None
def get_google_elevation((x,y)):
    google_url = 'https://maps.googleapis.com/maps/api/elevation/json?locations='
    url = google_url + str(x)+','+str(y) + '&key='+ config.google_key

    with contextlib.closing(urllib.urlopen(url)) as x:
        response = simple_json.load(x)
        try:
            elevation = response['results'][0]['elevation']
            return str(int(elevation))
        except KeyError:
            elevation = ''
            return elevation
def get_elevation(long,lat):
    ### This function extract elevation(z) value by passing lat and long information
    elevation = None
    xy = (long, lat)
    # read module
    master_grids = dem_footprints.dem_masterGrids
    # find DEM for one point
    dem_params = find_dem_parameters(master_grids,xy)
    
    if 10 in dem_params.keys():
        # Calculate elevation based on 10m collection
        elevation = get_single_elevation(xy,dem_params[10],config.img_dir_dem,10)

        if elevation is None and 30 in dem_params.keys():
            # 3.1 Calculate Elevation based on 30m collection
            elevation = get_single_elevation(xy,dem_params[30],config.img_dir_dem_CA,30)

        if elevation is None:
            elevation = get_single_elevation(xy)
        
        if elevation is None:
            elevation = ''
    return elevation
def get_file_size(image_path):
    return os.stat(image_path).st_size
def get_direction_text(x1,y1,x2,y2):
    angle = calculate_angle(x1,y1,x2,y2)
    return dgr_dir_to_txt(angle)

def calculate_angle(x1,y1,x2,y2):
    #(x1,y1) is the reference point.
    #return an angle in degrees [0,360]
    a = y2-y1
    b = x2-x1
    angle_180 = math.degrees(math.atan2(a,b))
    angle_360= (450-angle_180) % 360

    return angle_360
#convert the degree angle into a letter direction
def dgr_dir_to_txt(dgr_360):
    if (dgr_360 >= 348.75) or (dgr_360 < 11.25):
        dir_text = "N"
    elif (dgr_360 >= 11.25) and (dgr_360 < 33.75):
        dir_text = "NNE"
    elif (dgr_360 >= 33.75) and (dgr_360 < 56.25):
        dir_text = "NE"
    elif (dgr_360 >= 56.25) and (dgr_360 < 78.75):
        dir_text = "ENE"
    elif (dgr_360 >= 78.75) and (dgr_360 < 101.25):
        dir_text = "E"
    elif (dgr_360 >= 101.25) and (dgr_360 < 123.75):
        dir_text = "ESE"
    elif (dgr_360 >= 123.75) and (dgr_360 < 146.25):
        dir_text = "SE"
    elif (dgr_360 >= 146.25) and (dgr_360 < 168.75):
        dir_text = "SSE"
    elif (dgr_360 >= 168.75) and (dgr_360 < 191.25):
        dir_text = "S"
    elif (dgr_360 >= 191.25) and (dgr_360 < 213.75):
        dir_text = "SSW"
    elif (dgr_360 >= 213.75) and (dgr_360 < 236.25):
        dir_text = "SW"
    elif (dgr_360 >= 236.25) and (dgr_360 < 258.75):
        dir_text = "WSW"
    elif (dgr_360 >= 258.75) and (dgr_360 < 281.25):
        dir_text = "W"
    elif (dgr_360 >= 281.25) and (dgr_360 < 303.75):
        dir_text = "WNW"
    elif (dgr_360 >= 303.75) and (dgr_360 < 326.25):
        dir_text = "NW"
    elif (dgr_360 >= 326.25) and (dgr_360 < 348.75):
        dir_text = "NNW"
    else:
        dir_text = "NUL"    # this line should not be reached

    return dir_text
def degree_direction_to_text(dgr_360):
    if (dgr_360 >= 348.75) or (dgr_360 < 11.25):
        dir_text = "N"
    elif (dgr_360 >= 11.25) and (dgr_360 < 33.75):
        dir_text = "NNE"
    elif (dgr_360 >= 33.75) and (dgr_360 < 56.25):
        dir_text = "NE"
    elif (dgr_360 >= 56.25) and (dgr_360 < 78.75):
        dir_text = "ENE"
    elif (dgr_360 >= 78.75) and (dgr_360 < 101.25):
        dir_text = "E"
    elif (dgr_360 >= 101.25) and (dgr_360 < 123.75):
        dir_text = "ESE"
    elif (dgr_360 >= 123.75) and (dgr_360 < 146.25):
        dir_text = "SE"
    elif (dgr_360 >= 146.25) and (dgr_360 < 168.75):
        dir_text = "SSE"
    elif (dgr_360 >= 168.75) and (dgr_360 < 191.25):
        dir_text = "S"
    elif (dgr_360 >= 191.25) and (dgr_360 < 213.75):
        dir_text = "SSW"
    elif (dgr_360 >= 213.75) and (dgr_360 < 236.25):
        dir_text = "SW"
    elif (dgr_360 >= 236.25) and (dgr_360 < 258.75):
        dir_text = "WSW"
    elif (dgr_360 >= 258.75) and (dgr_360 < 281.25):
        dir_text = "W"
    elif (dgr_360 >= 281.25) and (dgr_360 < 303.75):
        dir_text = "WNW"
    elif (dgr_360 >= 303.75) and (dgr_360 < 326.25):
        dir_text = "NW"
    elif (dgr_360 >= 326.25) and (dgr_360 < 348.75):
        dir_text = "NNW"
    else:
        dir_text = "NUL"    # this line should not be reached

    return dir_text
    