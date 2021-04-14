import os
import sys
import arcpy
import ConfigParser

file_path =os.path.dirname(os.path.abspath(__file__))
if 'arcgisserver' in file_path:
    model_path = os.path.join(r'\\cabcvan1gis005\arcgisserver\directories\arcgissystem\arcgisinput\GPtools\DB_Framework')
else:
    main_path = os.path.abspath(os.path.join(__file__, os.pardir))
    model_path = os.path.join(main_path.split('GIS_Test')[0],'GIS_Test','DB_Framework')

sys.path.insert(1,model_path)
import models

def server_loc_config(configpath,environment):
    configParser = ConfigParser.RawConfigParser()
    configParser.read(configpath)
    dbconnection = configParser.get('server-config','dbconnection_%s'%environment)
    reportcheck = configParser.get('server-config','reportcheck_%s'%environment)
    reportviewer = configParser.get('server-config','reportviewer_%s'%environment)
    reportinstant = configParser.get('server-config','instant_%s'%environment)
    reportnoninstant = configParser.get('server-config','noninstant_%s'%environment)
    upload_viewer = configParser.get('url-config','uploadviewer_%s'%environment)
    server_config = {'dbconnection':dbconnection,'reportcheck':reportcheck,'viewer':reportviewer,'instant':reportinstant,'noninstant':reportnoninstant,'viewer_upload':upload_viewer}
    return server_config

# def createScratch():
#     scratch = os.path.join(r"\\cabcvan1gis005\MISC_DataManagement\_AW\TOPO_US_word_scratchy", "test2")
#     scratchgdb = "scratch.gdb"
#     if not os.path.exists(scratch):
#         os.mkdir(scratch)
#     if not os.path.exists(os.path.join(scratch, scratchgdb)):
#         arcpy.CreateFileGDB_management(scratch, "scratch.gdb")
#     return scratch, scratchgdb

# arcpy parameters
OrderIDText = arcpy.GetParameterAsText(0)
yesBoundary = (arcpy.GetParameterAsText(1)).lower()
scratch = arcpy.env.scratchFolder
scratchgdb = arcpy.env.scratchGDB

# order info
order_obj = models.Order().get_order(OrderIDText)

# # flags
# yesBoundary = "yes"                 # fixed/yes/no/arrow
# delyearFlag = "N"                   # Y/N, for internal use only, blank maps, etc.
custom_profile = False

# scratch file/folder outputs
# scratch, scratchgdb = createScratch()
orderGeometry = os.path.join(scratch, scratchgdb, "orderGeometry")
orderGeometryPR = os.path.join(scratch, scratchgdb, "orderGeometryPR")
orderBuffer = os.path.join(scratch, scratchgdb, "buffer")
extent = os.path.join(scratch, scratchgdb, "extent")

boundaryEMF = os.path.join(scratch, "orderGeometry.emf")
coverDOCX = os.path.join(scratch,'cover.docx')
summaryDOCX = os.path.join(scratch,'summary.docx')
coverEMF = os.path.join(scratch, "cover.emf")
summaryEMF = os.path.join(scratch, "summary.emf")
zipCover= os.path.join(scratch,'tozip_cover')
zipSummary = os.path.join(scratch,'tozip_summary')
zipMap = os.path.join(scratch,'tozip')
coverMXD = os.path.join(scratch,"cover.mxd")
summaryMXD = os.path.join(scratch, "summary.mxd")

# connection/report output
server_environment = 'test'
serverpath = r"\\cabcvan1gis006"
server_config_file = os.path.join(serverpath, r"GISData\ERISServerConfig.ini")
server_config = server_loc_config(server_config_file,server_environment)

reportcheckFolder = server_config["reportcheck"]
viewerFolder = server_config["viewer"]
topouploadurl =  server_config["viewer_upload"] + r"/TopoUpload?ordernumber="
connectionString = server_config["dbconnection"]

# folder
connectionPath = os.path.join(serverpath, r"GISData\Topo_US_Word_test")

# master data file\folder
masterfolder = r'\\cabcvan1fpr009\USGS_Topo'
csvfile_h = os.path.join(masterfolder, r"USGS_MapIndice\All_HTMC_all_all_gda_results.csv")
csvfile_c = os.path.join(masterfolder, r"USGS_MapIndice\All_USTopo_T_7.5_gda_results.csv")
mastergdb = os.path.join(masterfolder, r"USGS_MapIndice\MapIndices_National_GDB.gdb")
tifdir_h = os.path.join(masterfolder, "USGS_HTMC_Geotiff")
tifdir_c = os.path.join(masterfolder, "USGS_currentTopo_Geotiff")

# layer
topolyrfile_none = os.path.join(connectionPath,r"layer\topo.lyr")
topolyrfile_b = os.path.join(connectionPath,r"layer\topo_black.lyr")
topolyrfile_w = os.path.join(connectionPath,r"layer\topo_white.lyr")
bufferlyrfile = os.path.join(connectionPath,r"layer\buffer_extent.lyr")
orderGeomlyrfile_point = os.path.join(connectionPath,r"layer\SiteMaker.lyr")
orderGeomlyrfile_polyline = os.path.join(connectionPath,r"layer\orderLine.lyr")
orderGeomlyrfile_polygon = os.path.join(connectionPath,r"layer\orderPoly.lyr")

# docx
coverTemplate = os.path.join(connectionPath,r"templates\CoverPage")
summaryTemplate = os.path.join(connectionPath,r"templates\SummaryPage")
marginTemplate = os.path.join(connectionPath,r"templates\margin.docx")
customdocxTemplate =os.path.join(connectionPath,r"templates\template_noarrow")
docxTemplate =os.path.join(connectionPath,r"templates\template")

# mxd
Summarymxdfile = os.path.join(connectionPath, r"mxd\SummaryPage.mxd")
Covermxdfile = os.path.join(connectionPath, r"mxd\CoverPage.mxd")
mxdfile = os.path.join(connectionPath,r"mxd\template.mxd")

# other
readmefile = os.path.join(connectionPath,r"templates\readme.txt")
logfile = os.path.join(connectionPath, r"log\USTopoSearch_Terracon_Log.txt")
logname = "TOPO_US_word_%s"%server_environment