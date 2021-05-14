import arcpy,os
import ConfigParser
import arcpy
def server_loc_config(config_path):
    config_parser = ConfigParser.RawConfigParser()
    config_parser.read(config_path)
    report_check_test = config_parser.get('server-config','reportcheck_test')
    report_viewer_test = config_parser.get('server-config','reportviewer_test')
    report_instant_test = config_parser.get('server-config','instant_test')
    report_noninstant_test = config_parser.get('server-config','noninstant_test')
    upload_viewer = config_parser.get('url-config','uploadviewer')
    server_config = {'reportcheck':report_check_test,'viewer':report_viewer_test,'instant':report_instant_test,'noninstant':report_noninstant_test,'viewer_upload':upload_viewer}
    return server_config

class Report_Type:
    wetland = 'wetland'
    ny_wetland = 'ny_wetland'
    flood = 'flood'
    topo = 'topo'
    relief = 'relief'
    wells = 'wells'
    geology = 'geology'
    soil = 'soil'
    wells = 'wells'
    
server_config_file = r'\\cabcvan1gis006\GISData\ERISServerConfig.ini'
server_config = server_loc_config(server_config_file)
connection_string = 'eris_gis/gis295@cabcvan1ora006.glaciermedia.inc:1521/GMTESTC'
report_path = server_config['noninstant']
viewer_path = server_config['viewer']
upload_link = server_config['viewer_upload']+r"/ErisInt/BIPublisherPortal_test/Viewer.svc/"
report_check_path = server_config['reportcheck']
connection_path = r"\\cabcvan1gis005\GISData\PSR\python"
#Set the type of report
if_relief_report = None
if_topo_report = None
if_wetland_report = None
if_flood_report = None
if_geology_report = None
if_soil_report = None
if_ogw_report = None
if_radon_report = None
if_aspect_map = None
if_kml_output = None

scratch_folder = arcpy.env.scratchFolder
# temp gdb in scratch folder
temp_gdb = os.path.join(scratch_folder,r"temp.gdb")

order_geom_lyr_point = r"\\cabcvan1gis005\GISData\PSR\python\mxd\SiteMaker.lyr"
order_geom_lyr_polyline = r"\\cabcvan1gis005\GISData\PSR\python\mxd\orderLine.lyr"
order_geom_lyr_polygon = r"\\cabcvan1gis005\GISData\PSR\python\mxd\orderPoly.lyr"
buffer_lyr_file = r"\\cabcvan1gis005\GISData\PSR\python\mxd\buffer.lyr"
grid_lyr_file = r"\\cabcvan1gis005\GISData\PSR\python\mxd\Grid_hollow.lyr"

# grid size
grid_size = "2 MILES"
# Explorer
data_lyr_wetland = r"\\cabcvan1gis005\GISData\PSR\python\mxd\wetland_kml.lyr"
data_lyr_flood = r"\\cabcvan1gis005\GISData\PSR\python\mxd\flood.lyr"
data_lyr_geology = r"\\cabcvan1gis005\GISData\PSR\python\mxd\geology.lyr"
data_lyr_contour = r"\\cabcvan1gis005\GISData\PSR\python\mxd\contours_largescale.lyr"
data_lyr_plumet_acoma = r"\\cabcvan1gis005\GISData\PSR\python\mxd\Plume.lyr"

def output_jpg(order_obj, report_type):
    if report_type == Report_Type.wetland :
        return os.path.join(scratch_folder, str(order_obj.number) + '_US_WETL.jpg')
    elif report_type == Report_Type.ny_wetland :
         return os.path.join(scratch_folder, str(order_obj.number) + '_NY_WETL.jpg')
    elif report_type == Report_Type.flood:
        return os.path.join(scratch_folder, order_obj.number + '_US_FLOOD.jpg')
    elif report_type == Report_Type.topo:
        return os.path.join(scratch_folder, order_obj.number + '_US_TOPO.jpg')
    elif report_type == Report_Type.relief:
        return os.path.join(scratch_folder, order_obj.number + '_US_RELIEF.jpg')
    elif report_type == Report_Type.geology:
        return os.path.join(scratch_folder, order_obj.number + '_US_GEOLOGY.jpg')
    elif report_type == Report_Type.soil:
        return os.path.join(scratch_folder, order_obj.number + '_US_SOIL.jpg')
    elif report_type == Report_Type.wells:
        return os.path.join(scratch_folder, order_obj.number + '_US_WELLS.jpg')
    
### Basemaps
img_dir_dem_CA = r"\\Cabcvan1fpr009\US_DEM\DEM1"
master_lyr_dem_CA = r"\\Cabcvan1fpr009\US_DEM\Canada_DEM_edited.shp"
img_dir_dem = r"\\Cabcvan1fpr009\US_DEM\DEM13"
master_lyr_dem = r"\\cabcvan1gis005\GISData\Data\US_DEM\CellGrid_1X1Degree_NW_wgs84.shp"
master_lyr_states = r"\\cabcvan1gis005\GISData\PSR\python\mxd\USStates.lyr"
master_lyr_counties = r"\\cabcvan1gis005\GISData\PSR\python\mxd\USCounties.lyr"
master_lyr_cities = r"\\cabcvan1gis005\GISData\PSR\python\mxd\USCities.lyr"
master_lyr_nh_towns = r"\\cabcvan1gis005\GISData\PSR\python\mxd\NHTowns.lyr"
master_lyr_zip_codes = r"\\cabcvan1gis005\GISData\PSR\python\mxd\USZipcodes.lyr"
# Common Variables
if_multi_page = None
### order geometry paths config
order_geometry_pcs_shp =  os.path.join(scratch_folder,'order_geometry_pcs.shp')
order_geometry_gcs_shp =  os.path.join(scratch_folder,'order_geometry_gcs.shp')
order_buffer_shp =  os.path.join(scratch_folder,'order_buffer.shp')
order_geom_lyr_file = None
spatial_ref_mercator = arcpy.SpatialReference(3857) #web mercator
### wetland report config
buffer_dist_wetland = None
mxd_file_wetland = r"\\cabcvan1gis005\GISData\PSR\python\mxd\wetland.mxd"
mxd_file_wetland_ny = r"\\cabcvan1gis005\GISData\PSR\python\mxd\wetlandNY_CC.mxd"
mxd_mm_file_wetland = r"\\cabcvan1gis005\GISData\PSR\python\mxd\wetlandMM.mxd"
mxd_mm_file_wetland_ny = r"\\cabcvan1gis005\GISData\PSR\python\mxd\wetlandMMNY.mxd"
data_wetland = r'\\cabcvan1gis005\GISData\Data\PSR\PSR.gdb\Merged_wetland_Final'
data_lyr_wetland = r"\\cabcvan1gis005\GISData\PSR\python\mxd\wetland_kml.lyr"
data_lyr_wetland_ny_kml = r"\\cabcvan1gis005\GISData\PSR\python\mxd\wetlandNY_kml.lyr"
data_lyr_wetland_ny_apa_kml = r"\\cabcvan1gis005\GISData\PSR\python\mxd\wetlandNYAPA_kml.lyr"
### relief report paths config
buffer_dist_relief = '1 MILES'
mxd_file_relief =  r"\\cabcvan1gis005\GISData\PSR\python\mxd\shadedrelief.mxd"
mxd_mm_file_relief =  r"\\cabcvan1gis005\GISData\PSR\python\mxd\shadedreliefMM.mxd"
path_shaded_relief = r"\\cabcvan1fpr009\US_DEM\hillshade13"
relief_lyr_file = r"\\cabcvan1gis005\GISData\PSR\python\mxd\relief.lyr"
data_shaded_relief = r"\\cabcvan1fpr009\US_DEM\CellGrid_1X1Degree_NW.shp"
relief_frame = os.path.join(scratch_folder, "relief_frame.shp")
relief_image_name = "relief.jpg"
### topo report paths config
buffer_dist_topo = '1 MILES'
mxd_file_topo = r"\\cabcvan1gis005\GISData\PSR\python\mxd\topo.mxd"
mxd_file_topo_Tacoma = r"\\cabcvan1gis005\GISData\PSR\python\mxd\topo_tacoma.mxd"
mxd_mm_file_topo = r"\\cabcvan1gis005\GISData\PSR\python\mxd\topoMM.mxd"
mxd_mm_file_topo_Tacoma = r"\\cabcvan1gis005\GISData\PSR\python\mxd\topoMM_tacoma.mxd"
topo_master_lyr = r"\\cabcvan1gis005\GISData\Topo_USA\masterfile\_ARCHIVE\CellGrid_7_5_Minute_wgs84.shp"
data_topo = r"\\cabcvan1gis005\GISData\Topo_USA\masterfile\Cell_PolygonAll.shp"
topo_white_lyr_file = r"\\cabcvan1gis005\GISData\PSR\python\mxd\topo_white.lyr"
topo_csv_file = r"\\cabcvan1gis005\GISData\Topo_USA\masterfile\All_USTopo_T_7.5_gda_results.csv"
topo_tif_dir = r"\\cabcvan1fpr009\USGS_Topo\USGS_currentTopo_Geotiff"
topo_frame = os.path.join(scratch_folder, "topo_frame.shp")
topo_frame_gcs = os.path.join(scratch_folder, "topo_frame_gcs.shp")
### flood report paths config
buffer_dist_flood = None
data_flood = r'\\cabcvan1gis005\GISData\Data\PSR\PSR.gdb\flood_map_wgs84'
data_flood_panel = r'\\cabcvan1gis005\GISData\Data\PSR\PSR.gdb\flood_panel_map_wgs84'
mxd_file_flood = r"\\cabcvan1gis005\GISData\PSR\python\mxd\flood.mxd"
mxd_mm_file_flood = r"\\cabcvan1gis005\GISData\PSR\python\mxd\floodMM.mxd"
order_buffer_flood_shp = os.path.join(scratch_folder,'order_buffer_flood.shp')
flood_selectedby_order_shp = os.path.join(scratch_folder,"flood_selectedby_order.shp")
flood_panel_selectedby_order_shp = os.path.join(scratch_folder,"flood_panel_selectedby_order.shp")
flood_selectedby_frame = os.path.join(scratch_folder,"flood_selectedby_frame.shp")
data_lyr_flood = r"\\cabcvan1gis005\GISData\PSR\python\mxd\flood.lyr"
flood_ids = []
### geology report paths config
buffer_dist_geology = None
data_geology = r'\\cabcvan1gis005\GISData\Data\PSR\PSR.gdb\GEOL_DD_MERGE' ## WGS84
geology_selectedby_order_shp = os.path.join(scratch_folder,"geology_selectedby_order.shp")
mxd_file_geology = r"\\cabcvan1gis005\GISData\PSR\python\mxd\geology.mxd"
mxd_mm_file_geology = r"\\cabcvan1gis005\GISData\PSR\python\mxd\geologyMM.mxd"
geology_ids = []
### soil report paths config
buffer_dist_soil = None
data_path_soil_HI =r'\\cabcvan1gis005\GISData\Data\PSR\gSSURGO_HI.gdb'  ## WGS84
data_path_soil_AK =r'\\cabcvan1gis005\GISData\Data\PSR\gSSURGO_AK.gdb'  ## WGS84
data_path_soil_CONUS =r'\\cabcvan1gis005\GISData\Data\PSR\gSSURGO_CONUS.gdb'  ## WGS84
data_path_soil = None
report_data = []
soil_ids = []
soil_lyr = None
# data_path_soil_HI =r'\\cabcvan1fpr009\SSURGO\CONUS_2015\gSSURGO_HI.gdb'
# data_path_soil_AK =r'\\cabcvan1fpr009\SSURGO\CONUS_2015\gSSURGO_AK.gdb'
# data_path_soil_CONUS =r'\\cabcvan1fpr009\SSURGO\CONUS_2015\gSSURGO_CONUS_10m.gdb'

soil_selectedby_order_shp = os.path.join(scratch_folder,"soil_selectedby_order.shp")
soil_selectedby_order_pcs_shp =  os.path.join(scratch_folder,"soil_selectedby_order_pcs.shp")
soil_selectedby_frame =  os.path.join(scratch_folder,"soil_selectedby_frame.shp")
mxd_file_soil = r"\\cabcvan1gis005\GISData\PSR\python\mxd\soil.mxd"
mxd_mm_file_soil = r"\\cabcvan1gis005\GISData\PSR\python\mxd\soilMM.mxd"
hydrologic_dict = {
        "A":'Soils in this group have low runoff potential when thoroughly wet. Water is transmitted freely through the soil.',
        "B":'Soils in this group have moderately low runoff potential when thoroughly wet. Water transmission through the soil is unimpeded.',
        "C":'Soils in this group have moderately high runoff potential when thoroughly wet. Water transmission through the soil is somewhat restricted.',
        "D":'Soils in this group have high runoff potential when thoroughly wet. Water movement through the soil is restricted or very restricted.',
        "A/D":'These soils have low runoff potential when drained and high runoff potential when undrained.',
        "B/D":'These soils have moderately low runoff potential when drained and high runoff potential when undrained.',
        "C/D":'These soils have moderately high runoff potential when drained and high runoff potential when undrained.',
        }

hydric_dict = {
        '1':'All hydric',
        '2':'Not hydric',
        '3':'Partially hydric',
        '4':'Unknown',
        }

fc_soils_field_list  = [['muaggatt.mukey','mukey'], ['muaggatt.musym','musym'], ['muaggatt.muname','muname'],['muaggatt.drclassdcd','drclassdcd'],['muaggatt.hydgrpdcd','hydgrpdcd'],['muaggatt.hydclprs','hydclprs'], ['muaggatt.brockdepmin','brockdepmin'], ['muaggatt.wtdepannmin','wtdepannmin'], ['component.cokey','cokey'],['component.compname','compname'], ['component.comppct_r','comppct_r'], ['component.majcompflag','majcompflag'],['chorizon.chkey','chkey'],['chorizon.hzname','hzname'],['chorizon.hzdept_r','hzdept_r'],['chorizon.hzdepb_r','hzdepb_r'], ['chtexturegrp.chtgkey','chtgkey'], ['chtexturegrp.texdesc1','texdesc'], ['chtexturegrp.rvindicator','rv']]
fc_soils_key_list = ['muaggatt.mukey', 'component.cokey','chorizon.chkey','chtexturegrp.chtgkey']
fc_soils_where_clause_query_table = "muaggatt.mukey = component.mukey and component.cokey = chorizon.cokey and chorizon.chkey = chtexturegrp.chkey"
### ogw report paths config
dem_server = r"\\cabcvan1fpr009"
us_dem = 'US_DEM'
dir_dem = os.path.join(dem_server, us_dem,"DEM13")
dir_dem_ca = os.path.join(dem_server, us_dem,"DEM1")
master_lyr_dem = os.path.join(dem_server, us_dem,"DEM13.shp")
master_lyr_dem_ca = os.path.join(dem_server, us_dem,"DEM1.shp")
google_key = r'AIzaSyBmub_p_nY5jXrFMawPD8jdU0DgSrWfBic'

### Rado report config
buffer_dist_radon = None
states_selectedby_order = os.path.join(scratch_folder,"states_selectedby_order.shp")
counties_selectedby_order = os.path.join(scratch_folder,"counties_selectedby_order.shp")
cities_selectedby_order = os.path.join(scratch_folder,"cities_selectedby_order.shp")

### OGW report config
buffer_dist_ogw= None
order_center_pcs = os.path.join(scratch_folder, "order_center_pcs.shp")
eris_wells = r"\\cabcvan1gis005\GISData\PSR\python\mxd\ErisWellSites.lyr"   #which contains water, oil/gas wells etc.
wells_merge = os.path.join(scratch_folder, "wells_merge.shp")
wells_sj= os.path.join(scratch_folder,"wells_sj.shp")
wells_sja= os.path.join(scratch_folder,"wells_sja.shp")
wells_final= os.path.join(scratch_folder,"wells_fin.shp")
wells_display= os.path.join(scratch_folder,"wells_display.shp")
mxd_file_wells = r"\\cabcvan1gis005\GISData\PSR\python\mxd\wells.mxd"
mxd_mm_file_wells = r"\\cabcvan1gis005\GISData\PSR\python\mxd\wellsMM.mxd"

### Aspect map config
buffer_dist_aspect = '0.25 MILES'
order_aspect_buffer =  os.path.join(scratch_folder,'order_spect_buffer.shp')
dem_selectedby_order = os.path.join(scratch_folder,'aspect_selectedby_order.shp')