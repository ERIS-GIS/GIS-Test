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
#     scratch = os.path.join(r'\\cabcvan1gis005\MISC_DataManagement\_AW\FIM_US_SCRATCHY', "test2")
#     scratchgdb = "scratch.gdb"
#     if not os.path.exists(scratch):
#         os.mkdir(scratch)
#     if not os.path.exists(os.path.join(scratch, scratchgdb)):
#         arcpy.CreateFileGDB_management(scratch, "scratch.gdb")
#     return scratch, scratchgdb

# arcpy parameter
OrderIDText = arcpy.GetParameterAsText(0) 
BufsizeText = arcpy.GetParameterAsText(1)
yesBoundary = arcpy.GetParameterAsText(2)
multipage = arcpy.GetParameterAsText(3)
gridsize = arcpy.GetParameterAsText(4)
scratch = arcpy.env.scratchFolder
scratchgdb = arcpy.env.scratchGDB

# order info
order_obj = models.Order().get_order(OrderIDText)

# # parameters
# gridsize = "0.3 KiloMeters"
# BufsizeText ='0.17'
resolution = "600"

# # flags
# multipage = False                   # True/False        
# yesBoundary = "yes"                 # yes/no/fixed
# delyearFlag = "Y"                   # Y/N
nrf = 'N'                           # Y/N

# scratch file
# scratch, scratchgdb = createScratch()
orderGeometry= os.path.join(scratch, scratchgdb, "orderGeometry")
orderGeometryPR = os.path.join(scratch, scratchgdb, "orderGeometryPR")
outBufferSHP = os.path.join(scratch, scratchgdb, "buffer")
selectedmain = os.path.join(scratch, scratchgdb, "selectedmain")
selectedadj = os.path.join(scratch, scratchgdb, "selectedadj")
extent = os.path.join(scratch, scratchgdb, "extent")
summaryfile = os.path.join(scratch, "summary.pdf")
coverfile = os.path.join(scratch, "cover.pdf")
shapePdf = os.path.join(scratch, "shape.pdf")
annotPdf = os.path.join(scratch, "annot.pdf")

# connection/report output
server_environment = 'test'
serverpath = r"\\cabcvan1gis006"
server_config_file = os.path.join(serverpath, r"GISData\ERISServerConfig.ini")
server_config = server_loc_config(server_config_file,server_environment)

connectionString = server_config["dbconnection"]
reportcheckFolder = server_config["reportcheck"]
viewerFolder = server_config["viewer"]
uploadlink =  server_config["viewer_upload"] + r"/FIMUpload?ordernumber="

# folder
connectionPath = os.path.join(serverpath, r"GISData\FIMS_USA_test")
logopath = os.path.join(connectionPath, "logo")

# master file/folder
mastergdb = r"\\Cabcvan1fpr009\fim_data_usa\FIM_MASTERGDB\FIM_US_STATES.gdb"

# layer
imagelyr = os.path.join(connectionPath,r"layer\mosaic_jpg_255.lyr")
bndylyrfile = os.path.join(connectionPath,r"layer\boundary.lyr")
orderGeomlyrfile_point = os.path.join(connectionPath,r"layer\SiteMaker.lyr")
orderGeomlyrfile_polyline = os.path.join(connectionPath,r"layer\orderLine.lyr")
orderGeomlyrfile_polygon = os.path.join(connectionPath,r"layer\orderPoly.lyr")
sheetLayer = os.path.join(connectionPath, r"layer\hallowsheet.lyr")

# mxd
FIMmxdfile = os.path.join(connectionPath, r"mxd\FIMLayout.mxd")

# pdf
annot_poly = os.path.join(connectionPath,r"mxd\annot_poly.pdf")
annot_line = os.path.join(connectionPath,r"mxd\annot_line.pdf")
annot_point = os.path.join(connectionPath,r"mxd\annot_point.pdf")

# coverPic
coverPic = os.path.join(connectionPath, r"coverPic\ERIS_2018_ReportCover_Fire Insurance Maps_F.jpg")
secondPic = os.path.join(connectionPath, r"coverPic\ERIS_2018_ReportCover_Second Page_F.jpg")

# log
logfile = os.path.join(connectionPath, r"log\USFIM_Log.txt")
logname = "FIM_US_%s"%server_environment