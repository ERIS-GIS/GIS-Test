
#-------------------------------------------------------------------------------
# Name:        ERIS DATABASE MAP
# Purpose:    Generates the maps for the ERIS DATABASE Reprt.
#
# Author:      cchen
#
# Created:     18/01/2019
# Copyright:   (c) cchen 2019
# Licence:     <your licence>
#-------------------------------------------------------------------------------
import os
import traceback
import timeit
import arcpy
import json
import cx_Oracle
import shutil
import urllib
import topo_image_path
import ConfigParser
import csv
import xml.etree.ElementTree as ET
start1 = timeit.default_timer()
arcpy.env.overwriteOutput = True

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

server_environment = 'prod'
server_config_file = r'\\cabcvan1gis007\gptools\ERISServerConfig.ini'
server_config = server_loc_config(server_config_file,server_environment)
eris_report_path = r"gptools\ERISReport"
us_topo_path = r"gptools\Topo_USA"
eris_aerial_ca_path = r"gptools\Aerial_CAN"
tifdir_topo = r"\\cabcvan1fpr009\USGS_Topo\USGS_currentTopo_Geotiff"
world_aerial_arcGIS_online_URL = r"https://services.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/0/query?f=json&returnGeometry=false&spatialRel=esriSpatialRelIntersects&maxAllowableOffset=0&geometryType=esriGeometryPoint&inSR=4326&outFields=SRC_DATE"

class Machine:
    machine_test = r"\\cabcvan1gis006"
    machine_prod = r"\\cabcvan1gis007"

class Credential:
    oracle_test = r"eris_gis/gis295@GMTESTC.glaciermedia.inc"
    oracle_production = r'eris_gis/gis295@GMPRODC.glaciermedia.inc"

class ReportPath:
    noninstant_reports_test = server_config['noninstant']
    noninstant_reports_prod = server_config['noninstant']
    instant_report_test = server_config['instant']
    instant_report_prod = server_config['instant']

class TestConfig:
    machine_path=Machine.machine_test
    instant_reports =ReportPath.instant_report_test
    noninstant_reports = ReportPath.noninstant_reports_test

    def __init__(self,code):
        machine_path=self.machine_path
        self.LAYER=LAYER(machine_path)
        self.DATA=DATA(machine_path)
        self.MXD=MXD(machine_path,code)

class ProdConfig:
    machine_path=Machine.machine_prod
    instant_reports =ReportPath.instant_report_prod
    noninstant_reports = ReportPath.noninstant_reports_prod

    def __init__(self,code):
        machine_path=self.machine_path
        self.LAYER=LAYER(machine_path)
        self.DATA=DATA(machine_path)
        self.MXD=MXD(machine_path,code)

class Map(object):
    def __init__(self,mxdPath,dfname=''):
        self.mxd = arcpy.mapping.MapDocument(mxdPath)
        self.df= arcpy.mapping.ListDataFrames(self.mxd,('%s*')%(dfname))[0]

    def addLayer(self,lyr,workspace_path, dataset_name='',workspace_type="SHAPEFILE_WORKSPACE",add_position="TOP"):
        lyr = arcpy.mapping.Layer(lyr)
        if dataset_name !='':
            lyr.replaceDataSource(workspace_path, workspace_type, os.path.splitext(dataset_name)[0])
        arcpy.mapping.AddLayer(self.df, lyr, add_position)

    def replaceLayerSource(self,lyr_name,to_path, dataset_name='',workspace_type="SHAPEFILE_WORKSPACE"):
        for _ in arcpy.mapping.ListLayers(self.mxd):
            if _.name == lyr_name:
                _.replaceDataSource(to_path, workspace_type,dataset_name)
                return

    def toScale(self,value):
        self.df.scale=value
        self.scale =self.df.scale

    def zoomToTopLayer(self,position =0):
        self.df.extent = arcpy.mapping.ListLayers(self.mxd)[0].getExtent()

    def zoomToLayer(self,lyr_name):
        for _ in arcpy.mapping.ListLayers(self.mxd):
            if _.name ==lyr_name:
                self.df.extent = _.getExtent()
                break
            elif lyr_name in _.name:
                self.df.extent = _.getExtent()
                break
        arcpy.RefreshActiveView()

    def turnOnLayer(self):
        layers = arcpy.mapping.ListLayers(self.mxd, "*", self.df)
        for layer in layers:
            layer.visible = True
        arcpy.RefreshTOC()
        arcpy.RefreshActiveView()

    def turnLabel(self,lyr_name,visibility =True):
        layers = arcpy.mapping.ListLayers(self.mxd, "*", self.df)
        for layer in layers:
            if layer.name ==lyr_name or layer.name ==arcpy.mapping.Layer(lyr_name).name:
                layer.showLabels = visibility
        arcpy.RefreshTOC()
        arcpy.RefreshActiveView()

    def addTextoMap(self,textName,textValue, x=None,y=None):
        textElements =arcpy.mapping.ListLayoutElements(self.mxd,"TEXT_ELEMENT")
        for element in textElements:
            if textName.lower() in (element.name).lower():
                element.text = textValue
                if x!=None or y!=None:
                    element.elementPositionX=x
                    element.elementPositionY=y

class LAYER():
    def __init__(self,machine_path):
        self.machine_path = machine_path
        self.get()

    def get(self):
        machine_path = self.machine_path
        self.buffer = os.path.join(machine_path,eris_report_path,'layer','buffer.lyr')
        self.point = os.path.join(machine_path,eris_report_path,r"layer","SiteMaker.lyr")
        self.polyline = os.path.join(machine_path,eris_report_path,r"layer","orderLine.lyr")
        self.polygon = os.path.join(machine_path,eris_report_path,r"layer","orderPoly.lyr")
        self.buffer = os.path.join(machine_path,eris_report_path,'layer','buffer.lyr')
        self.grid = os.path.join(machine_path,eris_report_path,r"layer","GridCC.lyr")
        self.erisPoints = os.path.join(machine_path,eris_report_path,r"layer","ErisClipCC.lyr")
        self.topowhite = os.path.join(machine_path,eris_report_path,'layer',"topo_white.lyr")
        self.road = os.path.join(machine_path,eris_report_path,r"layer","Roadadd_notransparency.lyr")

class DATA():
    def __init__(self,machine_path):
        self.machine_path = machine_path
        self.get()

    def get(self):
        machine_path = self.machine_path
        self.data_topo = os.path.join(machine_path,us_topo_path,"masterfile","CellGrid_7_5_Minute.shp")
        self.road = os.path.join(machine_path,eris_report_path,r"layer","US","Roads2.lyr")

class MXD():
    def __init__(self,machine_path,code):
        self.machine_path = machine_path
        self.get(code)

    def get(self,code):
        machine_path = self.machine_path
        if code == 9093:        # USA
            self.mxdtopo = os.path.join(machine_path,eris_report_path,r"mxd","USTopoMapLayoutCC.mxd")
            self.mxdbing = os.path.join(machine_path,eris_report_path,r"mxd","USBingMapLayoutCC.mxd")
            self.mxdMM = os.path.join(machine_path,eris_report_path,'mxd','USLayoutMMCC.mxd')
        elif code == 9036:      # CAN
            self.mxdtopo = os.path.join(machine_path,eris_report_path,r"mxd","TopoMapLayoutCC.mxd")
            self.mxdbing = os.path.join(machine_path,eris_report_path,r"mxd","BingMapLayoutCC.mxd")
            self.mxdMM = os.path.join(machine_path,eris_report_path,'mxd','DMTILayoutMMCC.mxd')
        elif code == 9049:      # MEX
            self.mxdtopo = os.path.join(machine_path,eris_report_path,r"mxd","TopoMapLayoutCC.mxd")
            self.mxdbing = os.path.join(machine_path,eris_report_path,r"mxd","USBingMapLayoutCC.mxd")
            self.mxdMM = os.path.join(machine_path,eris_report_path,'mxd','MXLayoutMMCC.mxd')

class Oracle:
    # static variable: oracle_functions
    oracle_functions = {
    'getorderinfo':"eris_gis.getOrderInfo",
    'printtopo':"eris_gis.printTopo",
    'geterispointdetails':"eris_gis.getErisPointDetails"}

    oracle_procedures ={
    'xplorerflag':"eris_gis.getOrderXplorer"}

    def __init__(self,machine_name):
        # initiate connection credential
        if machine_name.lower() =='test':
            self.oracle_credential = Credential.oracle_test
        elif machine_name.lower()=='prod':
            self.oracle_credential = Credential.oracle_production
        else:
            raise ValueError("Bad machine name")

    def connect_to_oracle(self):
        try:
            self.oracle_connection = cx_Oracle.connect(self.oracle_credential)
            self.cursor = self.oracle_connection.cursor()
        except cx_Oracle.Error as e:
            print(e,'Oracle connection failed, review credentials.')

    def close_connection(self):
        self.cursor.close()
        self.oracle_connection.close()

    def call_function(self,function_name,OrderIDText):
        self.connect_to_oracle()
        cursor = self.cursor
        try:
            outType = cx_Oracle.CLOB
            func = [self.oracle_functions[_] for _ in self.oracle_functions.keys() if function_name.lower() ==_.lower()]
            if func != [] and len(func) == 1:
                try:
                    output=json.loads(cursor.callfunc(func[0],outType,((str(OrderIDText)),)).read())
                except ValueError:
                    output = cursor.callfunc(func[0],outType,((str(OrderIDText)),)).read()
                except AttributeError:
                    output = cursor.callfunc(func[0],outType,((str(OrderIDText)),))
            return output
        except cx_Oracle.Error as e:
            raise Exception(("Oracle Failure",e.message))
        except Exception as e:
            raise Exception(("JSON Failure",e.message))
        except NameError as e:
            raise Exception("Bad Function")
        finally:
            self.close_connection()

    def call_procedure(self,procedure_name,OrderIDText):
        self.connect_to_oracle()
        cursor = self.cursor
        try:
            outValue = 'Y'
            func = [self.oracle_procedures[_] for _ in self.oracle_procedures.keys() if procedure_name.lower() ==_.lower()]
            if func !=[] and len(func)==1:
                try:
                    output = cursor.callproc(func[0],[outValue,str(OrderIDText),])
                except ValueError:
                    output = cursor.callproc(func[0],[outValue,str(OrderIDText),])
                except AttributeError:
                    output = cursor.callproc(func[0],[outValue,str(OrderIDText),])
            return output
        except cx_Oracle.Error as e:
            raise Exception(("Oracle Failure",e.message))
        except Exception as e:
            raise Exception(("JSON Failure",e.message))
        except NameError as e:
            raise Exception("Bad Function")
        finally:
            self.close_connection()

    def insert_overlay(self,delete_query,insert_query):
        self.connect_to_oracle()
        cursor = self.cursor
        try:
            cursor.execute(delete_query)
            cursor.execute("commit")
            cursor.execute(insert_query)
            cursor.execute("commit")
            return "Oracle successfully populated image overlay info"
        except cx_Oracle.Error as e:
            raise Exception(("Oracle Failure",e.message))
        except Exception as e:
            raise Exception(("JSON Failure",e.message))
        except NameError as e:
            raise Exception("Bad Function")
        finally:
            self.close_connection()

def createBuffers(orderBuffers,output_folder,buffer_name=r"buffer_%s.shp"):
    buffer_dict={}
    buffer_sizes_dict ={}
    for i in range(len(orderBuffers)):
        buffer_dict[i]=createGeometry(eval(orderBuffers[i].values()[0])[0],"polygon",output_folder,buffer_name%i)
        buffer_sizes_dict[i] =float(orderBuffers[i].keys()[0])
    print(buffer_dict,buffer_sizes_dict)
    return [buffer_dict,buffer_sizes_dict]

def createGeometry(pntCoords,geometry_type,output_folder,output_name, spatialRef = arcpy.SpatialReference(4269)):
    outputSHP = os.path.join(output_folder,output_name)
    if geometry_type.lower()== 'point':
        arcpy.CreateFeatureclass_management(output_folder, output_name, "MULTIPOINT", "", "DISABLED", "DISABLED", spatialRef)
        cursor = arcpy.da.InsertCursor(outputSHP, ['SHAPE@'])
        cursor.insertRow([arcpy.Multipoint(arcpy.Array([arcpy.Point(*coords) for coords in pntCoords]),spatialRef)])
    elif geometry_type.lower() =='polyline':
        arcpy.CreateFeatureclass_management(output_folder, output_name, "POLYLINE", "", "DISABLED", "DISABLED", spatialRef)
        cursor = arcpy.da.InsertCursor(outputSHP, ['SHAPE@'])
        cursor.insertRow([arcpy.Polyline(arcpy.Array([arcpy.Point(*coords) for coords in pntCoords]),spatialRef)])
    elif geometry_type.lower() =='polygon':
        arcpy.CreateFeatureclass_management(output_folder,output_name, "POLYGON", "", "DISABLED", "DISABLED", spatialRef)
        cursor = arcpy.da.InsertCursor(outputSHP, ['SHAPE@'])
        cursor.insertRow([arcpy.Polygon(arcpy.Array([arcpy.Point(*coords) for coords in pntCoords]),spatialRef)])
    del cursor
    return outputSHP

def addERISpoint(pointInfo,mxd,output_folder,out_points=r'points.shp'):
    out_pointsSHP = os.path.join(output_folder,out_points)
    erisPointsLayer = config.LAYER.erisPoints
    #erisIDs_4points = dict((_.get('DATASOURCE_POINTS')[0].get('ERIS_DATA_ID'),[('m%sc'%(_.get("MAP_KEY_LOC"))) if _.get("MAP_KEY_NO_TOT")==1 else ('m%sc(%s)'%(_.get("MAP_KEY_LOC"), _.get("MAP_KEY_NO_TOT"))) ,float('%s'%(1 if round(_.get("ELEVATION_DIFF"),2)>0.0 else 0 if round(_.get("ELEVATION_DIFF"),2)==0.0 else -1 if round(_.get("ELEVATION_DIFF"),2)<0.0 else 100))]) for _ in pointInfo)
    erisIDs_4points = dict((_.get('DATASOURCE_POINTS')[0].get('ERIS_DATA_ID'),[('m%sc'%(_.get("MAP_KEY_LOC"))) if _.get("MAP_KEY_NO_TOT")==1 else ('m%sc(%s)'%(_.get("MAP_KEY_LOC"), _.get("MAP_KEY_NO_TOT"))) ,float('%s'%(1 if _.get("ELEVATION_DIFF")>0.0 else 0 if _.get("ELEVATION_DIFF")==0.0 else -1 if _.get("ELEVATION_DIFF")<0.0 else 100))]) for _ in pointInfo)
    erispoints = dict((int(_.get('DATASOURCE_POINTS')[0].get('ERIS_DATA_ID')),(_.get("X"),_.get("Y"))) for _ in pointInfo)
    # print(erisIDs_4points)
    if erisIDs_4points != {}:
        arcpy.CreateFeatureclass_management(output_folder, out_points, "MULTIPOINT", "", "DISABLED", "DISABLED", arcpy.SpatialReference(4269))
        check_field = arcpy.ListFields(out_pointsSHP,"ERISID")
        if check_field==[]:
            arcpy.AddField_management(out_pointsSHP, "ERISID", "LONG", field_length='40')
        cursor = arcpy.da.InsertCursor(out_pointsSHP, ['SHAPE@','ERISID'])
        for point in erispoints.keys():
            cursor.insertRow([arcpy.Multipoint(arcpy.Array([arcpy.Point(*erispoints[point])]),arcpy.SpatialReference(4269)),point])
        del cursor
        check_field = arcpy.ListFields(out_pointsSHP,"mapkey")
        if check_field==[]:
            arcpy.AddField_management(out_pointsSHP, "mapkey", "TEXT", "", "", "20", "", "NULLABLE", "NON_REQUIRED", "")
        check_field = arcpy.ListFields(out_pointsSHP,"eleRank")
        if check_field==[]:
            arcpy.AddField_management(out_pointsSHP, "eleRank", "SHORT", "12", "6", "", "", "NULLABLE", "NON_REQUIRED", "")
        rows = arcpy.UpdateCursor(out_pointsSHP)
        for row in rows:
            row.mapkey = erisIDs_4points[int(row.ERISID)][0]
            row.eleRank = erisIDs_4points[int(row.ERISID)][1]
            rows.updateRow(row)
        del rows
        mxd.addLayer(erisPointsLayer,output_folder,out_points)
    return erisPointsLayer

def addRoadLayer(mxd,buffer_name, output_folder):
    road_clip = r"road_clip"
    arcpy.Clip_analysis(config.DATA.road, buffer_name, os.path.join(output_folder,road_clip), "0.3 Meters")
    mxd.addLayer(config.LAYER.road,output_folder,road_clip)

def addOrderGeometry(mxd,geometry_type,output_folder,name):
    geometryLayer = eval('config.LAYER.%s'%(geometry_type.lower()))
    if arcpy.mapping.ListLayoutElements(mxd.mxd, "LEGEND_ELEMENT", "Legend") !=[]:
        legend = arcpy.mapping.ListLayoutElements(mxd.mxd, "LEGEND_ELEMENT", "Legend")[0]
        legend.autoAdd = True
        mxd.addLayer(geometryLayer,output_folder,name)
        legend.autoAdd = False
    else:
        mxd.addLayer(geometryLayer,output_folder,name)
    mxd.zoomToTopLayer()

def getMaps(mxd, output_folder,map_name,buffer_dict, buffer_sizes_list,unit_code,buffer_name=r"buffer_%s.shp", multiPage = False):
    temp = []
    if buffer_name.endswith(".shp"):
        buffer_name = buffer_name[:-4]
    bufferLayer = config.LAYER.buffer
    for i in buffer_dict.keys():
        if buffer_sizes_list[i]>=0.04:
            mxd.addLayer(bufferLayer,output_folder,"buffer_%s"%(i))
        if i in buffer_dict.keys()[-3:]:
            mxd.zoomToLayer("Grid") if i == buffer_dict.keys()[-1] and multiPage == True else mxd.zoomToTopLayer()
            mxd.df.scale = ((int(1.1*mxd.df.scale)/100)+1)*100
            unit = 'Kilometer' if unit_code ==9036 else 'Mile'
            mxd.addTextoMap("Map","Map: %s %s Radius"%(buffer_sizes_list[i],unit))
            arcpy.mapping.ExportToPDF(mxd.mxd,os.path.join(output_folder,map_name%(i)))
            temp.append(os.path.join(output_folder,map_name%(i)))
    if temp==[] and buffer_dict=={}:
        arcpy.mapping.ExportToPDF(mxd.mxd,os.path.join(output_folder,map_name%0))
        temp.append(os.path.join(output_folder,map_name%0))
    return temp

def exportMap(mxd,output_folder,map_name,UTMzone,buffer_dict,buffer_sizes_list,unit_code, buffer_name=r"buffer_%s.shp"):
    mxd.df.spatialReference = arcpy.SpatialReference('WGS 1984 UTM Zone %sN'%UTMzone)
    mxd.resolution =250
    temp = getMaps(mxd, output_folder,map_name, buffer_dict, buffer_sizes_list,unit_code, buffer_name=r"buffer_%s.shp", multiPage = False)
    mxd.mxd.saveACopy(os.path.join(output_folder,"mxd.mxd"))
    return temp

def exportMultipage(mxd,output_folder,map_name,UTMzone,grid_size,erisPointLayer,buffer_dict,buffer_sizes_list,unit_code, buffer_name=r"buffer_%s.shp"):
    bufferLayer = config.LAYER.buffer
    gridlr = "gridlr"
    gridlrSHP = os.path.join(output_folder, gridlr+'.shp')

    arcpy.GridIndexFeatures_cartography(gridlrSHP, buffer_dict[buffer_dict.keys()[-1]], "", "", "", grid_size, grid_size)
    mxd.replaceLayerSource("Grid",output_folder,gridlr)
    arcpy.CalculateAdjacentFields_cartography(gridlrSHP, u'PageNumber')
    mxd.turnLabel(erisPointLayer,False)
    mxd.df.spatialReference = arcpy.SpatialReference('WGS 1984 UTM Zone %sN'%UTMzone)
    mxd.resolution =250
    temp = getMaps(mxd, output_folder,map_name, buffer_dict, buffer_sizes_list, unit_code, buffer_name=r"buffer_%s.shp", multiPage = True)
    mxd.turnLabel(erisPointLayer,True)
    mxd.addTextoMap("Map","Grid: ")
    mxd.addTextoMap("Grid",'<dyn type="page"  property="number"/>')
    ddMMDDP = mxd.mxd.dataDrivenPages
    ddMMDDP.refresh()
    ddMMDDP.exportToPDF(os.path.join(output_folder,map_name%("GRID")), "ALL",resolution=200,layers_attributes='LAYERS_ONLY',georef_info=False)
    mxd.mxd.saveACopy(os.path.join(output_folder,"mxdMM.mxd"))
    return [temp,os.path.join(output_folder,map_name%("GRID"))]

def exportTopo(mxd,output_folder,geometry_name,geometry_type, output_pdf,unit_code,bufferSHP,UTMzone):
    geometryLayer = eval('config.LAYER.%s'%geometry_type.lower())
    addOrderGeometry(mxd,geometry_type,output_folder,geometry_name)
    mxd.df.spatialReference = arcpy.SpatialReference('WGS 1984 UTM Zone %sN'%UTMzone)
    topoYear = '2020'
    if unit_code == 9093:
        topoLayer = config.LAYER.topowhite    
        topolist = getCurrentTopo(config.DATA.data_topo,bufferSHP,output_folder)
        topoYear = getTopoQuadnYear(topolist)[1]
        mxd.addTextoMap("Year", "Year: %s"%topoYear)
        mxd.addTextoMap("Quadrangle","Quadrangle(s): %s"%getTopoQuadnYear(topolist)[0])
        for topo in topolist:
            mxd.addLayer(topoLayer,output_folder,topo.split('.')[0],"RASTER_WORKSPACE","BOTTOM")
    elif unit_code == 9049:
        mxd.addTextoMap("Logo", "\xa9 ERIS Information Inc.")
    mxd.toScale(24000) if mxd.df.scale<24000 else mxd.toScale(1.1*mxd.df.scale)
    mxd.resolution=300
    arcpy.mapping.ExportToPDF(mxd.mxd,output_pdf)
    if xplorerflag == 'Y' :
        df = arcpy.mapping.ListDataFrames(mxd.mxd,'')[0]
        mxd.df.spatialReference = arcpy.SpatialReference(3857)
        projectproperty = arcpy.mapping.ListLayers(mxd.mxd,"Project Property",df)[0]
        projectproperty.visible = False
        arcpy.mapping.ExportToJPEG(mxd.mxd,os.path.join(scratchviewer,topoYear+'_topo.jpg'), df, 3825, 4950, world_file = True, jpeg_quality = 85)
        exportViewerTable(os.path.join(scratchviewer,topoYear+'_topo.jpg'),topoYear+'_topo.jpg')
    mxd.mxd.saveACopy(os.path.join(output_folder,"maptopo.mxd"))

def getCurrentTopo(masterfile_topo,inputSHP,output_folder): # copy current topo images that intersect with input shapefile to output folder
    masterLayer_topo = arcpy.mapping.Layer(masterfile_topo)
    arcpy.SelectLayerByLocation_management(masterLayer_topo,'intersect',inputSHP)
    if(int((arcpy.GetCount_management(masterLayer_topo).getOutput(0))) ==0):
        return None
    else:
        cellids_selected = []
        infomatrix = []
        rows = arcpy.SearchCursor(masterLayer_topo) # loop through the relevant records, locate the selected cell IDs
        for row in rows:
            cellid = str(int(row.getValue("CELL_ID")))
            cellids_selected.append(cellid)
        del row
        del rows
        masterLayer_topo = None        
        
        for cellid in cellids_selected:
            try:
                exec("info =  topo_image_path.topo_%s"%(cellid))
                infomatrix.append(info)
                print(infomatrix)
            except  AttributeError as ae:
                print("AttributeError: No current topo available")
                print(ae)

                newmastertopo = r'\\cabcvan1gis006\GISData\Topo_USA\masterfile\Cell_PolygonAll.shp'                
                csvfile_h = r'\\cabcvan1gis006\GISData\Topo_USA\masterfile\All_HTMC_all_all_gda_results.csv'
                global tifdir_topo
                tifdir_topo = r'\\cabcvan1fpr009\USGS_Topo\USGS_HTMC_Geotiff'
                masterLayer = arcpy.mapping.Layer(newmastertopo)
                #arcpy.SelectLayerByLocation_management(masterLayer,'intersect', inputSHP, '0.25 KILOMETERS')  #it doesn't seem to work without the distance
                arcpy.SelectLayerByLocation_management(masterLayer,'INTERSECT', inputSHP)
                cellids_selected = []
                cellids = []
                infomatrix = []

                if(int((arcpy.GetCount_management(masterLayer).getOutput(0))) ==0):
                    print ("NO records selected")
                    masterLayer = None
                else:
                    # loop through the relevant records, locate the selected cell IDs
                    rows = arcpy.SearchCursor(masterLayer)    # loop through the selected records
                    for row in rows:
                        cellid = str(int(row.getValue("CELL_ID")))
                        cellids_selected.append(cellid)
                        cellsize = str(int(row.getValue("CELL_SIZE")))
                        cellids.append(cellid)
                        # cellsizes.append(cellsize)
                    del row
                    del rows
                    masterLayer = None
                    
                    with open(csvfile_h, "rb") as f:
                        reader = csv.reader(f)
                        for row in reader:
                            if row[9] in cellids:
                                # print "#2 " + time.strftime('%Y-%m-%d %H:%M:%S', time.localtime())
                                pdfname = row[15].strip()
                                # read the year from .xml file
                                xmlname = pdfname[0:-3] + "xml"
                                xmlpath = os.path.join(tifdir_topo,xmlname)
                                tree = ET.parse(xmlpath)
                                root = tree.getroot()
                                procsteps = root.findall("./dataqual/lineage/procstep")
                                yeardict = {}
                                for procstep in procsteps:
                                    procdate = procstep.find("./procdate")
                                    if procdate != None:
                                        procdesc = procstep.find("./procdesc")
                                        yeardict[procdesc.text.lower()] = procdate.text

                                year2use = ""
                                yearcandidates = []
                                if "edit year" in yeardict.keys():
                                    yearcandidates.append(int(yeardict["edit year"]))

                                if "aerial photo year" in yeardict.keys():
                                    yearcandidates.append(int(yeardict["aerial photo year"]))

                                if "photo revision year" in yeardict.keys():
                                    yearcandidates.append(int(yeardict["photo revision year"]))

                                if "field check year" in yeardict.keys():
                                    yearcandidates.append(int(yeardict["field check year"]))

                                if "photo inspection year" in yeardict.keys():
                                    # print "photo inspection year is " + yeardict["photo inspection year"]
                                    yearcandidates.append(int(yeardict["photo inspection year"]))

                                if "date on map" in yeardict.keys():
                                    # print "date on  map " + yeardict["date on map"]
                                    yearcandidates.append(int(yeardict["date on map"]))

                                if len(yearcandidates) > 0:
                                    # print "***** length of yearcnadidates is " + str(len(yearcandidates))
                                    year2use = str(max(yearcandidates))
                                if year2use == "":
                                    print ("################### cannot determine the year of the map!!")
                                
                                # ONLY GET 7.5 OR 15 MINUTE MAP SERIES
                                if row[5] == "7.5X7.5 GRID" or row[5] == "15X15 GRID":
                                    infomatrix.append([row[15],year2use])  # [64818, 15X15 GRID,  LA_Zachary_335142_1963_62500_geo.pdf,  1963]
                    # print(infomatrix)
                    # GET MAX YEAR ONLY
                    infomatrix = [item for item in infomatrix if item[1] == max(item[1] for item in infomatrix)]             
        _=[]
        for item in infomatrix:
            tifname = item[0][0:-4]   # note without .tif part
            topofile = os.path.join(tifdir_topo,tifname+"_t.tif")
            year = item[1]

            if os.path.exists(topofile):
                if '.' in tifname:
                    tifname = tifname.replace('.','')
                temp = tifname.split('_')
                temp.insert(-2,item[1])
                newtopo = '_'.join(temp)+'.tif'
                shutil.copyfile(topofile,os.path.join(output_folder,newtopo))
                _.append(newtopo)
        return _
        
def getTopoYear(name):
    for year in range(1900,2030):
        if str(year) in name:
            return str(year)
    return None

def getTopoQuadnYear(topo_filelist):
    quadrangles=set()
    year=set()
    for topo in topo_filelist:
        name = topo.split("_")
        z = 0
        for i in range(len(name)):
            year_value = getTopoYear(name[i])
            if year_value:
                if z == 0:
                    quadrangles.add('%s, %s'%(' '.join([name[j] for j in range(1,i)]), name[0]))
                year.add(year_value)
                z = z + 1
    # GET MAX YEAR FROM FILE NAME
    year = [y for y in year if y == max(y for y in year)]
    return ('; '.join(quadrangles),'; '.join(year))

def exportAerial(mxd,output_folder,geometry_name,geometry_type,centroid,scale,output_pdf,UTMzone):
    geometryLayer = eval('config.LAYER.%s'%geometry_type.lower())
    addOrderGeometry(mxd,geometry_type,output_folder,geometry_name)
    aerialYear = getWorldAerialYear(centroid)
    mxd.addTextoMap("Year","Year: %s"%aerialYear)
    mxd.df.spatialReference = arcpy.SpatialReference('WGS 1984 UTM Zone %sN'%UTMzone)
    mxd.toScale(10000) if mxd.df.scale<10000 else mxd.toScale(1.1*mxd.df.scale)
    mxd.resolution=200
    arcpy.mapping.ExportToPDF(mxd.mxd,output_pdf)

    if xplorerflag == 'Y' :
        df = arcpy.mapping.ListDataFrames(mxd.mxd,'')[0]
        mxd.df.spatialReference = arcpy.SpatialReference(3857)
        projectproperty = arcpy.mapping.ListLayers(mxd.mxd,"Project Property",df)[0]
        projectproperty.visible = False
        arcpy.mapping.ExportToJPEG(mxd.mxd,os.path.join(scratchviewer,aerialYear+'_aerial.jpg'), df, 3825, 4950, world_file = True, jpeg_quality = 85)
        exportViewerTable(os.path.join(scratchviewer,aerialYear+'_aerial.jpg'),aerialYear+'_aerial.jpg')
    mxd.mxd.saveACopy(os.path.join(output_folder,"mapbing.mxd"))

def getWorldAerialYear((centroid_X,centroid_Y)):
    fsURL = r"https://services.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/0/query?f=json&returnGeometry=false&spatialRel=esriSpatialRelIntersects&maxAllowableOffset=0&geometryType=esriGeometryPoint&inSR=4326&outFields=SRC_DATE"
    params = urllib.urlencode({'geometry':{'x':float(centroid_X),'y':float(centroid_Y)}})
    resultBing = urllib.urlopen(fsURL,params).read()

    if "error" not in resultBing:
        for year in list(reversed(range(1900,2020))):
            if str(year) in resultBing :
                return str(year)
    else:
        tries = 5
        key = False
        while tries >= 0:
            if "error" not in resultBing:
                for year in list(reversed(range(1900,2020))):
                    if str(year) in resultBing:
                        return str(year)
            elif tries == 0:
                    return ""
            else:
                time.sleep(5)
                tries -= 1

def exportViewerTable(ImagePath,FileName):
    srGoogle = arcpy.SpatialReference(3857)
    srWGS84 = arcpy.SpatialReference(4326)
    metaitem = {}
    arcpy.DefineProjection_management(ImagePath,srGoogle)
    desc = arcpy.Describe(ImagePath)
    featbound = arcpy.Polygon(arcpy.Array([desc.extent.lowerLeft, desc.extent.lowerRight, desc.extent.upperRight, desc.extent.upperLeft]),srGoogle)
    del desc

    tempfeat = os.path.join(scratch, "imgbnd_"+FileName[:-4]+ ".shp")
    arcpy.Project_management(featbound, tempfeat, srWGS84)
    desc = arcpy.Describe(tempfeat)
    metaitem['type'] = 'cur'+ str(FileName.split('_')[1]).split('.')[0]
    metaitem['imagename'] = FileName
    metaitem['lat_sw'] = desc.extent.YMin
    metaitem['long_sw'] = desc.extent.XMin
    metaitem['lat_ne'] = desc.extent.YMax
    metaitem['long_ne'] = desc.extent.XMax

    delete_query = "delete from overlay_image_info where order_id = '%s' and type = '%s' and filename = '%s'"%(OrderIDText,metaitem['type'],FileName)
    insert_query = "insert into overlay_image_info values (%s, %s, %s, %.5f, %.5f, %.5f, %.5f, %s, '', '')" % (str(OrderIDText), orderInfo['ORDER_NUM'], "'" + metaitem['type']+"'", metaitem['lat_sw'], metaitem['long_sw'], metaitem['lat_ne'], metaitem['long_ne'],"'"+metaitem['imagename']+"'" )
    image_info = Oracle('prod').insert_overlay(delete_query,insert_query)

if __name__ == '__main__':
    try:
        # INPUT #####################################
        OrderIDText = '988390'#'#arcpy.GetParameterAsText(0).strip()#'736799'#
        multipage = False #True if (arcpy.GetParameterAsText(1).lower()=='yes' or arcpy.GetParameterAsText(1).lower()=='y') else False
        gridsize = '2 Miles'#arcpy.GetParameterAsText(2).strip()#0#
        code = 'usa'#arcpy.GetParameterAsText(3).strip()#'usa'#
        isInstant = False #True if arcpy.GetParameterAsText(4).strip().lower()=='yes'else False
        scratch = r'C:\Users\JLoucks\Documents\JL\test2'

        # Server Setting ############################
        code = 9093 if code.strip().lower()=='usa' else 9036 if code.strip().lower()=='can' else 9049 if code.strip().lower()=='mex' else ValueError
        config = ProdConfig(code)

        # PARAMETERS ################################
        orderGeometry = r'orderGeometry.shp'
        orderGeometrySHP = os.path.join(scratch,orderGeometry)
        map_name = 'map_%s.pdf'
        mapMM_name = 'mapMM_%s.pdf'
        bufferMax = ''
        buffer_name = "buffer_%s.shp"
        mapMM = os.path.join(scratch,mapMM_name)
        aerial_pdf = os.path.join(scratch,'mapbing.pdf')
        topo_pdf = os.path.join(scratch,"maptopo.pdf")
        pdfreport = os.path.join(scratch,map_name)
        grid_unit = 'Kilometers' if code == 9036 and float(gridsize.strip())<100 else 'Meters' if code ==9036 else 'Miles'
        gridsize = '%s %s'%(gridsize,grid_unit)
        viewerpath = server_config['viewer']
        currentuploadurl = server_config['viewer_upload']+r"/ErisInt/BIPublisherPortal_prod/Viewer.svc/CurImageUpload?ordernumber="

        # STEPS ####################################
        # 1  get order info by Oracle call
        orderInfo= Oracle('prod').call_function('getorderinfo',OrderIDText)
        needTopo= Oracle('prod').call_function('printTopo',OrderIDText)
        xplorerflag= Oracle('prod').call_procedure('xplorerflag',OrderIDText)[0]
        end = timeit.default_timer()
        arcpy.AddMessage(('call oracle', round(end -start1,4)))
        start=end

        # 2 create xplorer directory
        if xplorerflag == 'Y':
            scratchviewer = os.path.join(scratch,orderInfo['ORDER_NUM']+'_current')
            os.mkdir(scratchviewer)

        # 2 create order geometry
        orderGeometrySHP = createGeometry(eval(orderInfo[u'ORDER_GEOMETRY'][u'GEOMETRY'])[0],orderInfo[u'ORDER_GEOMETRY'][u'GEOMETRY_TYPE'],scratch,orderGeometry)
        end = timeit.default_timer()
        arcpy.AddMessage(('create geometry shp', round(end -start,4)))
        start=end

        # 3 create buffers
        [buffers, buffer_sizes] = createBuffers(orderInfo['BUFFER_GEOMETRY'],scratch,buffer_name)
        end = timeit.default_timer()
        arcpy.AddMessage(('create buffer shps', round(end -start,4)))
        start=end

        # 4 Maps
        # 4-0 initial Map
        map1 = Map(config.MXD.mxdMM)
        end = timeit.default_timer()
        arcpy.AddMessage(('4-0 initiate object Map', round(end -start,4)))
        start=end

        if code == 9093:
            # 3.1 MAx Buffer
            bufferMax = os.path.join(scratch,buffer_name%(len(orderInfo['BUFFER_GEOMETRY'])+1))
            maxBuffer = max([float(_.keys()[0]) for _ in orderInfo['BUFFER_GEOMETRY']]) if orderInfo['BUFFER_GEOMETRY'] !=[] else 0#
            maxBuffer ="%s MILE"%(2*maxBuffer if maxBuffer>0.2 else 2)
            # print(orderInfo['BUFFER_GEOMETRY'])
            arcpy.Buffer_analysis(orderGeometrySHP,bufferMax,maxBuffer)
            end = timeit.default_timer()
            arcpy.AddMessage(('create max buffer', round(end -start,4)))
            start=end
            # 4-1 add Road US
            addRoadLayer(map1, bufferMax,scratch)
            end = timeit.default_timer()
            arcpy.AddMessage(('4-1 clip and add road', round(end -start,4)))
            start=end

        # 4-2 add ERIS points
        erisPointsInfo = Oracle('prod').call_function('geterispointdetails',OrderIDText)
        # for i in erisPointsInfo:
        #     print(i)
        erisPointsLayer=addERISpoint(erisPointsInfo,map1,scratch)
        end = timeit.default_timer()
        arcpy.AddMessage(('4-3 add ERIS points to Map object', round(end -start,4)))
        start=end

        # 4-2 add Order Geometry
        addOrderGeometry(map1,orderInfo['ORDER_GEOMETRY']['GEOMETRY_TYPE'],scratch,orderGeometry)
        end = timeit.default_timer()
        arcpy.AddMessage(('4-2 add Geometry layer to Map object', round(end -start,4)))
        start=end

        # 4-3 Add Address n Order Number Turn on Layers
        map1.addTextoMap('Address',"Address: %s, %s, %s"%(orderInfo['ADDRESS'],orderInfo["CITY"],orderInfo['PROVSTATE']))
        map1.addTextoMap("OrderNum","Order Number: %s"%orderInfo['ORDER_NUM'])
        map1.turnOnLayer()
        end = timeit.default_timer()
        arcpy.AddMessage(('4-3 Add Address n turn on source layers', round(end -start,4)))
        start=end

        # 4-4 Optional Multipage add Buffer Export Map
        zoneUTM = orderInfo['ORDER_GEOMETRY']['UTM_ZONE']
        if zoneUTM<10:
            zoneUTM =' %s'%zoneUTM

        if multipage==True:
            [maplist,mapMM] = exportMultipage(map1,scratch,mapMM_name,zoneUTM,gridsize,erisPointsLayer,buffers,buffer_sizes,code,buffer_name)
            end = timeit.default_timer()
            arcpy.AddMessage(('4-4 MM map to pdf', round(end -start,4)))
            start=end
        else:
            # 4-4 add Buffer Export Map
            maplist = exportMap(map1,scratch,map_name,zoneUTM,buffers,buffer_sizes,code,buffer_name)
            end = timeit.default_timer()
            arcpy.AddMessage(('4-4 maps to 3 pdfs', round(end -start,4)))
            start=end
        scale = map1.df.scale
        del map1,erisPointsLayer

        # 5 Aerial
        mapbing = Map(config.MXD.mxdbing)
        end = timeit.default_timer()
        arcpy.AddMessage(('5-1 inital aerial', round(end -start,4)))
        start=end
        mapbing.addTextoMap('Address',"Address: %s, %s, %s"%(orderInfo['ADDRESS'],orderInfo["CITY"],orderInfo['PROVSTATE']))
        mapbing.addTextoMap("OrderNum","Order Number: %s"%orderInfo['ORDER_NUM'])
        exportAerial(mapbing,scratch,orderGeometry,orderInfo['ORDER_GEOMETRY']['GEOMETRY_TYPE'],eval(orderInfo['ORDER_GEOMETRY']['CENTROID'].strip('[]')),scale, aerial_pdf,zoneUTM)
        del mapbing
        end = timeit.default_timer()
        arcpy.AddMessage(('5-2 aerial', round(end -start,4)))
        start=end

        # 6 Topo
        if needTopo =='Y':
            maptopo = Map(config.MXD.mxdtopo)
            end = timeit.default_timer()
            arcpy.AddMessage(('6-1 topo', round(end -start,4)))
            start=end
            maptopo.addTextoMap('Address',"Address: %s, %s"%(orderInfo['ADDRESS'],orderInfo['PROVSTATE']))
            maptopo.addTextoMap("OrderNum","Order Number: %s"%orderInfo['ORDER_NUM'])
            
            exportTopo(maptopo,scratch,orderGeometry,orderInfo['ORDER_GEOMETRY']['GEOMETRY_TYPE'],topo_pdf,code,bufferMax,zoneUTM)
            del maptopo,orderGeometry
            end = timeit.default_timer()
            arcpy.AddMessage(('6 Topo', round(end -start,4)))
            start=end

        # 7 Report
        maplist.sort(reverse=True)
        if multipage ==True:
            maplist.append(mapMM)
        maplist.append(aerial_pdf)
        maplist.append(topo_pdf) if needTopo =='Y' else None
        end = timeit.default_timer()
        arcpy.AddMessage(('7 maplist', round(end -start,4)))
        start=end

        pdfreport =pdfreport%(orderInfo['ORDER_NUM'])
        outputPDF = arcpy.mapping.PDFDocumentCreate(pdfreport)
        for page in maplist:
            outputPDF.appendPages(str(page))
        outputPDF.saveAndClose()

        if isInstant:
            shutil.copy(pdfreport,config.instant_reports)
        else:
            shutil.copy(pdfreport,config.noninstant_reports)
        end = timeit.default_timer()
        arcpy.AddMessage(('7 Bundle', round(end -start,4)))
        start=end
        arcpy.SetParameterAsText(5,pdfreport)

        # Xplorer
        if xplorerflag == 'Y':
            if os.path.exists(os.path.join(viewerpath,orderInfo['ORDER_NUM']+'_current')):
                shutil.rmtree(os.path.join(viewerpath,orderInfo['ORDER_NUM']+'_current'))
            shutil.copytree(scratchviewer, os.path.join(viewerpath, orderInfo['ORDER_NUM']+'_current'))
            url = currentuploadurl + orderInfo['ORDER_NUM']
            urllib.urlopen(url)

    except:
        tb = sys.exc_info()[2]
        tbinfo = traceback.format_tb(tb)[0]
        pymsg = "PYTHON ERRORS:\nTraceback info:\n %s \nError Info:\n %s"%(tbinfo,str(sys.exc_info()[1]))
        msgs = "ArcPy ERRORS:\n %s\n"%arcpy.GetMessages(2)
        arcpy.AddError("hit CC's error code in except: Order ID %s"%OrderIDText)
        arcpy.AddError(pymsg)
        arcpy.AddError(msgs)
        raise