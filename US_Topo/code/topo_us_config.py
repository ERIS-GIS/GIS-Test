import os
import sys
import arcpy
import ConfigParser

file_path =os.path.dirname(os.path.abspath(__file__))
if 'arcgisserver' in file_path:
    model_path = os.path.join(r'D:/arcgisserver/directories/arcgissystem/arcgisinput/GPtools/DB_Framework')
else:
    main_path = os.path.abspath(os.path.join(__file__, os.pardir))
    model_path = os.path.join(main_path.split('GIS_Test')[0],'GIS_Test','DB_Framework')

sys.path.insert(1,model_path)
import models

def server_loc_config(configpath,environment):
    configParser = ConfigParser.RawConfigParser()
    configParser.read(configpath)
    if environment == 'test':
        dbconnection = configParser.get('server-config','dbconnection_test')
        reportcheck = configParser.get('server-config','reportcheck_test')
        reportviewer = configParser.get('server-config','reportviewer_test')
        reportinstant = configParser.get('server-config','instant_test')
        reportnoninstant = configParser.get('server-config','noninstant_test')
        upload_viewer = configParser.get('url-config','uploadviewer_test')
        server_config = {'dbconnection':dbconnection,'reportcheck':reportcheck,'viewer':reportviewer,'instant':reportinstant,'noninstant':reportnoninstant,'viewer_upload':upload_viewer}
        return server_config
    else:
        return 'invalid server configuration'

# def createScratch():
#     scratch = os.path.join(r"\\cabcvan1gis005\MISC_DataManagement\_AW\TOPO_US_SCRATCHY", "test1")
#     scratchgdb = "scratch.gdb"
#     if not os.path.exists(scratch):
#         os.mkdir(scratch)
#     if not os.path.exists(os.path.join(scratch, scratchgdb)):
#         arcpy.CreateFileGDB_management(scratch, "scratch.gdb")
#     return scratch, scratchgdb

# arcpy parameters
OrderIDText = arcpy.GetParameterAsText(0)
yesBoundary = (arcpy.GetParameterAsText(1)).lower()
BufsizeText = arcpy.GetParameterAsText(2)
multipage = arcpy.GetParameterAsText(3)
gridsize = arcpy.GetParameterAsText(4)
scratch = arcpy.env.scratchWorkspace
scratchgdb = arcpy.env.scratchGDB

# order info
order_obj = models.Order().get_order(OrderIDText)

# # flags
# multipage = "Y"                     # Y/N, for multipages
# gridsize =  0 # "3 KiloMeters"           # for multipage grid
# yesBoundary = "yes"                 # fixed/yes/no
# BufsizeText = "2.4"
# delyearFlag = "N"                   # Y/N, for internal use only, blank maps, etc.

# scratch file/folder outputs
# scratch, scratchgdb = createScratch()
summaryPdf = os.path.join(scratch,'summary.pdf')
coverPdf = os.path.join(scratch,"cover.pdf")
shapePdf = os.path.join(scratch, 'shape.pdf')
annotPdf = os.path.join(scratch, "annot.pdf")
orderGeometry = os.path.join(scratch, scratchgdb, "orderGeometry")
orderGeometryPR = os.path.join(scratch, scratchgdb, "orderGeometryPR")
orderBuffer = os.path.join(scratch, scratchgdb, "buffer")
extent = os.path.join(scratch, scratchgdb, "extent")

# connections/report outputs
server_environment = 'test'
server_config_file = r"\\cabcvan1gis006\GISData\ERISServerConfig.ini"
server_config = server_loc_config(server_config_file,server_environment)

reportcheckFolder = server_config["reportcheck"]
viewerFolder = server_config["viewer"]
topouploadurl =  server_config["viewer_upload"] + r"/TopoUpload?ordernumber="
connectionString = server_config["dbconnection"] #con.connection_string #'eris_gis/gis295@cabcvan1ora006.glaciermedia.inc:1521/GMTESTC'

# folders
connectionPath = r"\\cabcvan1gis006\GISData\Topo_USA"
mxdpath = os.path.join(connectionPath, r"mxd")

# master data files\folders
mastergdb = os.path.join(connectionPath, r"masterfile\MapIndices_National_GDB.gdb")
csvfile_h = os.path.join(connectionPath, r"masterfile\All_HTMC_all_all_gda_results.csv")
csvfile_c = os.path.join(connectionPath, r"masterfile\All_USTopo_T_7.5_gda_results.csv")
tifdir_h = r'\\cabcvan1fpr009\USGS_Topo\USGS_HTMC_Geotiff'
tifdir_c = r'\\cabcvan1fpr009\USGS_Topo\USGS_currentTopo_Geotiff'

# mxds
mxdfile = os.path.join(mxdpath,"template.mxd")
mxdfile_nova = os.path.join(mxdpath,'template_nova.mxd')

# layers
topolyrfile_none = os.path.join(mxdpath,"topo.lyr")
topolyrfile_b = os.path.join(mxdpath,"topo_black.lyr")
topolyrfile_w = os.path.join(mxdpath,"topo_white.lyr")
bufferlyrfile = os.path.join(mxdpath,"buffer.lyr")
orderGeomlyrfile_point = os.path.join(mxdpath,"orderPoint.lyr")
orderGeomlyrfile_polyline = os.path.join(mxdpath,"orderLine.lyr")
orderGeomlyrfile_polygon = os.path.join(mxdpath,"orderPoly.lyr")
gridlyr = os.path.join(mxdpath,"grid.lyr")

# pdfs
annot_poly = os.path.join(mxdpath,"annot_poly.pdf")
annot_line = os.path.join(mxdpath,"annot_line.pdf")
annot_poly_c = os.path.join(mxdpath,"annot_poly_red.pdf")
annot_line_c = os.path.join(mxdpath,"annot_line_red.pdf")
pdfsymbolfile = os.path.join(mxdpath, "US Topo Map Symbols v7.4.pdf")

# logos
logopath = os.path.join(mxdpath,"logos")

# covers
coverPic = os.path.join(mxdpath, "coverPic", "ERIS_2018_ReportCover_Topographic Maps_F.jpg")
summarypage = os.path.join(mxdpath, "coverPic", "ERIS_2018_ReportCover_Second Page_F.jpg")

# other
logfile = os.path.join(connectionPath, r"log\USTopoSearch_Log.txt")
logname = "TOPO_US_test"
readmefile = os.path.join(mxdpath,"readme.txt")