from imp import reload
import arcpy, os, sys
from datetime import datetime
import timeit,time
import shutil
# import psr_utility as utility
import psr_config as config
file_path =os.path.dirname(os.path.abspath(__file__))
sys.path.insert(1,os.path.join(os.path.dirname(file_path),'DB_Framework'))
import models
reload(sys)
def generate_radon_report(order_obj):
    arcpy.AddMessage('  -- Start generating PSR Radon report...')
    start = timeit.default_timer()   
    ### set scratch folder
    arcpy.env.workspace = config.scratch_folder
    arcpy.env.overwriteOutput = True  
    if '10689' not in order_obj.psr.search_radius.keys():
        arcpy.AddMessage('      -- Radon search radius is not availabe')
        return
    config.buffer_dist_radon =  str(order_obj.psr.search_radius['10689']) + ' MILES'
    
    ### create buffer map based on order geometry
    arcpy.Buffer_analysis(config.order_geometry_pcs_shp, config.order_buffer_shp, config.buffer_dist_radon) 
    
    states_lyr = arcpy.MakeFeatureLayer_management(config.master_lyr_states, 'states_lyr') 
    arcpy.SelectLayerByLocation_management(states_lyr, 'intersect',  config.order_buffer_shp)
    
    counties_lyr = arcpy.MakeFeatureLayer_management(config.master_lyr_counties, 'counties_lyr') 
    arcpy.SelectLayerByLocation_management('counties_lyr', 'intersect',  config.order_buffer_shp)
    
    cities_lyr = arcpy.MakeFeatureLayer_management(config.master_lyr_cities, 'cities_lyr') 
    arcpy.SelectLayerByLocation_management('cities_lyr', 'intersect',  config.order_buffer_shp)
    
    zip_codes_lyr = arcpy.MakeFeatureLayer_management(config.master_lyr_zip_codes, 'zip_codes_lyr') 
    arcpy.SelectLayerByLocation_management('zip_codes_lyr', 'intersect',  config.order_buffer_shp)
    
    state_list = ''
    in_rows = arcpy.SearchCursor(states_lyr)
    for in_row in in_rows:
        state_list = state_list+ ',' + in_row.STUSPS
    state_list = state_list.strip(',')        #two letter state
    state_list_str = str(state_list)
    del in_rows
    del in_row
    
    county_list = ''
    in_rows = arcpy.SearchCursor(counties_lyr)
    for in_row in in_rows:
        county_list = county_list + ',' + in_row.NAME
    county_list = county_list.strip(',')
    county_list_str = str(county_list.replace(u'\xed','i').replace(u'\xe1','a').replace(u'\xf1','n'))
    del in_rows
    if 'in_row' in locals():     #sometimes returns no city
        del in_row
    
    city_list = ''
    in_rows = arcpy.SearchCursor(cities_lyr)
    for in_row in in_rows:
        city_list = city_list + ',' + in_row.NAME
    city_list = city_list.strip(',')
    del in_rows
    if 'in_row' in locals():     #sometimes returns no city
        del in_row
    
    if 'NH' in state_list:
        town_lyr = arcpy.MakeFeatureLayer_management(config.master_lyr_nh_towns, 'town_lyr') 
        arcpy.SelectLayerByLocation_management('town_lyr', 'intersect',  config.order_buffer_shp)

        in_rows = arcpy.SearchCursor(town_lyr)
        for in_row in in_rows:
            city_list = city_list + ','+in_row.NAME
        city_list = city_list.strip(',')
        del in_rows
        if 'in_row' in locals():     #sometimes returns no city
            del in_row
    city_list_str = str(city_list.replace(u'\xed','i').replace(u'\xe1','a').replace(u'\xf1','n'))
    
    zip_list = ''
    in_rows = arcpy.SearchCursor(zip_codes_lyr)
    for in_row in in_rows:
        zip_list = zip_list + ',' + in_row.ZIP
    zip_list = zip_list.strip(',')
    zip_list_str = str(zip_list)
    del in_rows
    if 'in_row' in locals():
        del in_row
    ### update DB
    psr_obj = models.PSR()
    psr_obj = models.PSR()
    psr_obj.get_radon(order_obj.id, state_list_str, zip_list_str, county_list_str, city_list_str)
    
    end = timeit.default_timer()
    arcpy.AddMessage((' -- End generating PSR Radon report. Duration:', round(end -start,4)))