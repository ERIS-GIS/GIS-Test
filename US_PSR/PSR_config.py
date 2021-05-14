#-------------------------------------------------------------------------------
# Name:        module1
# Purpose:
#
# Author:      cchen
#
# Created:     16/01/2018
# Copyright:   (c) cchen 2018
# Licence:     <your licence>
#-------------------------------------------------------------------------------

import arcpy,os
import ConfigParser

def server_loc_config(configpath,environment):
    configParser = ConfigParser.RawConfigParser()
    configParser.read(configpath)
    if environment == 'test':
        reportcheck = configParser.get('server-config','reportcheck_test')
        reportviewer = configParser.get('server-config','reportviewer_test')
        reportinstant = configParser.get('server-config','instant_test')
        reportnoninstant = configParser.get('server-config','noninstant_test')
        upload_viewer = configParser.get('url-config','uploadviewer')
        server_config = {'reportcheck':reportcheck,'viewer':reportviewer,'instant':reportinstant,'noninstant':reportnoninstant,'viewer_upload':upload_viewer}
        return server_config
    elif environment == 'prod':
        reportcheck = configParser.get('server-config','reportcheck_prod')
        reportviewer = configParser.get('server-config','reportviewer_prod')
        reportinstant = configParser.get('server-config','instant_prod')
        reportnoninstant = configParser.get('server-config','noninstant_prod')
        upload_viewer = configParser.get('url-config','uploadviewer_prod')
        server_config = {'reportcheck':reportcheck,'viewer':reportviewer,'instant':reportinstant,'noninstant':reportnoninstant,'viewer_upload':upload_viewer}
        return server_config
    else:
        return 'invalid server configuration'

server_environment = 'test'
server_config_file = r'\\cabcvan1gis006\GISData\ERISServerConfig.ini'
server_config = server_loc_config(server_config_file,server_environment)
connectionString = 'eris_gis/gis295@cabcvan1ora006.glaciermedia.inc:1521/GMTESTC'
report_path = server_config['noninstant']
viewer_path = server_config['viewer']
upload_link = server_config['viewer_upload']+r"/ErisInt/BIPublisherPortal_prod/Viewer.svc/"
#production: upload_link = r"http://CABCVAN1OBI002/ErisInt/BIPublisherPortal_prod/Viewer.svc/"
reportcheck_path = server_config['reportcheck']
connectionPath = r"\\cabcvan1gis005\GISData\PSR\python"

orderGeomlyrfile_point = r"\\cabcvan1gis005\GISData\PSR\python\mxd\SiteMaker.lyr"
orderGeomlyrfile_polyline = r"\\cabcvan1gis005\GISData\PSR\python\mxd\orderLine.lyr"
orderGeomlyrfile_polygon = r"\\cabcvan1gis005\GISData\PSR\python\mxd\orderPoly.lyr"
bufferlyrfile = r"\\cabcvan1gis005\GISData\PSR\python\mxd\buffer.lyr"
topowhitelyrfile = r"\\cabcvan1gis005\GISData\PSR\python\mxd\topo_white.lyr"
gridlyrfile = r"\\cabcvan1gis005\GISData\PSR\python\mxd\Grid_hollow.lyr"
relieflyrfile = r"\\cabcvan1gis005\GISData\PSR\python\mxd\relief.lyr"

masterlyr_topo = r"\\cabcvan1gis005\GISData\Topo_USA\masterfile\CellGrid_7_5_Minute.shp"
data_topo = r"\\cabcvan1gis005\GISData\Topo_USA\masterfile\Cell_PolygonAll.shp"
csvfile_topo = r"\\cabcvan1gis005\GISData\Topo_USA\masterfile\All_USTopo_T_7.5_gda_results.csv"
tifdir_topo = r"\\cabcvan1fpr009\USGS_Topo\USGS_currentTopo_Geotiff"

data_shadedrelief = r"\\cabcvan1fpr009\US_DEM\CellGrid_1X1Degree_NW.shp"
data_geol = r'\\cabcvan1gis005\GISData\Data\PSR\PSR.gdb\GEOL_DD_MERGE'
data_flood = r'\\cabcvan1gis005\GISData\Data\PSR\PSR.gdb\S_Fld_haz_Ar_merged2018'
data_floodpanel = r'\\cabcvan1gis005\GISData\Data\PSR\PSR.gdb\S_FIRM_PAN_MERGED2018'
data_wetland = r'\\cabcvan1gis005\GISData\Data\PSR\PSR.gdb\Merged_wetland_Final'
eris_wells = r"\\cabcvan1gis005\GISData\PSR\python\mxd\ErisWellSites.lyr"   #which contains water, oil/gas wells etc.
path_shadedrelief = r"\\cabcvan1fpr009\US_DEM\hillshade13"
datalyr_wetland = r"\\cabcvan1gis005\GISData\PSR\python\mxd\wetland_kml.lyr"
##datalyr_wetlandNY = r"E:\GISData\PSR\python\mxd\wetlandNY.lyr"
datalyr_wetlandNYkml = r"\\cabcvan1gis005\GISData\PSR\python\mxd\wetlandNY_kml.lyr"
datalyr_wetlandNYAPAkml = r"\\cabcvan1gis005\GISData\PSR\python\mxd\wetlandNYAPA_kml.lyr"
datalyr_flood = r"\\cabcvan1gis005\GISData\PSR\python\mxd\flood.lyr"
datalyr_geology = r"\\cabcvan1gis005\GISData\PSR\python\mxd\geology.lyr"
datalyr_contour = r"\\cabcvan1gis005\GISData\PSR\python\mxd\contours_largescale.lyr"
datalyr_plumetacoma = r"\\cabcvan1gis005\GISData\PSR\python\mxd\Plume.lyr"

imgdir_demCA = r"\\Cabcvan1fpr009\US_DEM\DEM1"
masterlyr_demCA = r"\\Cabcvan1fpr009\US_DEM\Canada_DEM_edited.shp"
imgdir_dem = r"\\Cabcvan1fpr009\US_DEM\DEM13"
masterlyr_dem = r"\\Cabcvan1fpr009\US_DEM\CellGrid_1X1Degree_NW_imagename_update.shp"
masterlyr_states = r"\\cabcvan1gis005\GISData\PSR\python\mxd\USStates.lyr"
masterlyr_counties = r"\\cabcvan1gis005\GISData\PSR\python\mxd\USCounties.lyr"
masterlyr_cities = r"\\cabcvan1gis005\GISData\PSR\python\mxd\USCities.lyr"
masterlyr_NHTowns = r"\\cabcvan1gis005\GISData\PSR\python\mxd\NHTowns.lyr"
masterlyr_zipcodes = r"\\cabcvan1gis005\GISData\PSR\python\mxd\USZipcodes.lyr"

mxdfile_topo = r"\\cabcvan1gis005\GISData\PSR\python\mxd\topo.mxd"
mxdfile_topo_Tacoma = r"\\cabcvan1gis005\GISData\PSR\python\mxd\topo_tacoma.mxd"
mxdMMfile_topo = r"\\cabcvan1gis005\GISData\PSR\python\mxd\topoMM.mxd"
mxdMMfile_topo_Tacoma = r"\\cabcvan1gis005\GISData\PSR\python\mxd\topoMM_tacoma.mxd"
mxdfile_relief =  r"\\cabcvan1gis005\GISData\PSR\python\mxd\shadedrelief.mxd"
mxdMMfile_relief =  r"\\cabcvan1gis005\GISData\PSR\python\mxd\shadedreliefMM.mxd"
mxdfile_wetland = r"\\cabcvan1gis005\GISData\PSR\python\mxd\wetland.mxd"
mxdfile_wetlandNY = r"\\cabcvan1gis005\GISData\PSR\python\mxd\wetlandNY_CC.mxd"
mxdMMfile_wetland = r"\\cabcvan1gis005\GISData\PSR\python\mxd\wetlandMM.mxd"
mxdMMfile_wetlandNY = r"\\cabcvan1gis005\GISData\PSR\python\mxd\wetlandMMNY.mxd"
mxdfile_flood = r"\\cabcvan1gis005\GISData\PSR\python\mxd\flood.mxd"
mxdMMfile_flood = r"\\cabcvan1gis005\GISData\PSR\python\mxd\floodMM.mxd"
mxdfile_geol = r"\\cabcvan1gis005\GISData\PSR\python\mxd\geology.mxd"
mxdMMfile_geol = r"\\cabcvan1gis005\GISData\PSR\python\mxd\geologyMM.mxd"
mxdfile_soil = r"\\cabcvan1gis005\GISData\PSR\python\mxd\soil.mxd"
mxdMMfile_soil = r"\\cabcvan1gis005\GISData\PSR\python\mxd\soilMM.mxd"
mxdfile_wells = r"\\cabcvan1gis005\GISData\PSR\python\mxd\wells.mxd"
mxdMMfile_wells = r"\\cabcvan1gis005\GISData\PSR\python\mxd\wellsMM.mxd"

srGCS83 = arcpy.SpatialReference(os.path.join(connectionPath, r"projections\GCSNorthAmerican1983.prj"))

datapath_soil_HI =r'\\cabcvan1fpr009\SSURGO\CONUS_2015\gSSURGO_HI.gdb'
datapath_soil_AK =r'\\cabcvan1fpr009\SSURGO\CONUS_2015\gSSURGO_AK.gdb'
datapath_soil_CONUS =r'\\cabcvan1fpr009\SSURGO\CONUS_2015\gSSURGO_CONUS_10m.gdb'

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

fc_soils_fieldlist  = [['muaggatt.mukey','mukey'], ['muaggatt.musym','musym'], ['muaggatt.muname','muname'],['muaggatt.drclassdcd','drclassdcd'],['muaggatt.hydgrpdcd','hydgrpdcd'],['muaggatt.hydclprs','hydclprs'], ['muaggatt.brockdepmin','brockdepmin'], ['muaggatt.wtdepannmin','wtdepannmin'], ['component.cokey','cokey'],['component.compname','compname'], ['component.comppct_r','comppct_r'], ['component.majcompflag','majcompflag'],['chorizon.chkey','chkey'],['chorizon.hzname','hzname'],['chorizon.hzdept_r','hzdept_r'],['chorizon.hzdepb_r','hzdepb_r'], ['chtexturegrp.chtgkey','chtgkey'], ['chtexturegrp.texdesc1','texdesc'], ['chtexturegrp.rvindicator','rv']]
fc_soils_keylist = ['muaggatt.mukey', 'component.cokey','chorizon.chkey','chtexturegrp.chtgkey']
fc_soils_whereClause_queryTable = "muaggatt.mukey = component.mukey and component.cokey = chorizon.cokey and chorizon.chkey = chtexturegrp.chkey"

tbx = r"\\cabcvan1gis005\GISData\PSR\python\DEV\ERIS.tbx"

# Explorer
datalyr_wetland = r"\\cabcvan1gis005\GISData\PSR\python\mxd\wetland_kml.lyr"
datalyr_flood = r"\\cabcvan1gis005\GISData\PSR\python\mxd\flood.lyr"
datalyr_geology = r"\\cabcvan1gis005\GISData\PSR\python\mxd\geology.lyr"