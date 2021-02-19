import arcpy, cx_Oracle, os
import json, shutil
class Machine:
    machine_test = r"\\cabcvan1gis006"
    machine_prod = r"\\cabcvan1gis007"
class Credential:
    oracle_test = r"ERIS_GIS/gis295@GMTESTC.glaciermedia.inc"
    oracle_production = r"ERIS_GIS/gis295@GMPRODC.glaciermedia.inc"
class ReportPath:
    caaerial_prod= r'\\CABCVAN1OBI007\ErisData\prod\aerial_ca'
    caaerial_test= r'\\CABCVAN1OBI007\ErisData\\test\aerial_ca'
class TestConfig:
    machine_path=Machine.machine_test
    caaerial_path = ReportPath.caaerial_test

    def __init__(self):
        machine_path=self.machine_path
        self.LAYER=LAYER(machine_path)
        self.MXD=MXD(machine_path)
class ProdConfig:
    machine_path=Machine.machine_prod
    caaerial_path = ReportPath.caaerial_prod

    def __init__(self):
        machine_path=self.machine_path
        self.LAYER=LAYER(machine_path)
        self.MXD=MXD(machine_path)
class Oracle:
    # static variable: oracle_functions
    oracle_functions = {'getorderinfo':"eris_gis.getOrderInfo"}
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
    def call_function(self,function_name,orderID):
        self.connect_to_oracle()
        cursor = self.cursor
        try:
            outType = cx_Oracle.CLOB
            func = [self.oracle_functions[_] for _ in self.oracle_functions.keys() if function_name.lower() ==_.lower()]
            if func !=[] and len(func)==1:
                try:
                    if type(orderID) !=list:
                        orderID = [orderID]
                    output=json.loads(cursor.callfunc(func[0],outType,orderID).read())
                except ValueError:
                    output = cursor.callfunc(func[0],outType,orderID).read()
                except AttributeError:
                    output = cursor.callfunc(func[0],outType,orderID)
            return output
        except cx_Oracle.Error as e:
            raise Exception(("Oracle Failure",e.message))
        except Exception as e:
            raise Exception(("JSON Failure",e.message))
        except NameError as e:
            raise Exception("Bad Function")
        finally:
            self.close_connection()
    def call_erisapi(self,erisapi_input):
        self.connect_to_oracle()
        cursor = self.cursor
        arg1 = erisapi_input
        arg2 = cursor.var(cx_Oracle.CLOB)
        arg3 = cursor.var(cx_Oracle.CLOB)
        arg4 = cursor.var(str)
        try:
            func = ['eris_api.callOracle']
            if func !=[] and len(func)==1:
                try:
                    output = cursor.callproc(func[0],[arg1,arg2,arg3,arg4])
                except ValueError:
                    output = cursor.callproc(func[0],[arg1,arg2,arg3,arg4])
                except AttributeError:
                    output = cursor.callproc(func[0],[arg1,arg2,arg3,arg4])
            return [output[0],cx_Oracle.LOB.read(output[1]),cx_Oracle.LOB.read(output[2]),output[3]]
        except cx_Oracle.Error as e:
            raise Exception(("Oracle Failure",e.message))
        except Exception as e:
            raise Exception(("JSON Failure",e.message))
        except NameError as e:
            raise Exception("Bad Function")
        finally:
            self.close_connection()
    def pass_values(self,function_name,value):#(self,function_name,data_type,value):
        self.connect_to_oracle()
        cursor = self.cursor
        try:
            func = [self.oracle_functions[_] for _ in self.oracle_functions.keys() if function_name.lower() ==_.lower()]
            if func !=[] and len(func)==1:
                try:
                    #output= cursor.callfunc(func[0],oralce_object,value)
                    output= cursor.callproc(func[0],value)
                    return 'pass'
                except ValueError:
                    raise
            return 'failed'
        except cx_Oracle.Error as e:
            raise Exception(("Oracle Failure",e.message))
        except Exception as e:
            raise Exception(e.message)
        except NameError as e:
            raise Exception("Bad Function")
        finally:
            self.close_connection()
def createGeometry(pntCoords,geometry_type,output_folder,output_name, spatialRef = arcpy.SpatialReference(4326)):
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

OrderID = arcpy.GetParameterAsText(0)#'968849'#arcpy.GetParameterAsText(0)'1014745'
scale = arcpy.GetParameterAsText(1)#'2000'#arcpy.GetParameterAsText(1)'1200'
unit = arcpy.GetParameterAsText(2)#'ft'#arcpy.GetParameterAsText(2)'ft'
scratch = arcpy.env.scratchF0older#r'C:\Users\JLoucks\Documents\JL\test9'
init_env = 'test'
jobfolder = os.path.join(r'\\cabcvan1eap003\v2_usaerial\JobData', init_env)
mxdtemplate = r'\\cabcvan1gis006\GISData\Aerial_US\mxd\Aerial_US_Export_new.mxd'
#supported units ft ratio | default ratio

if unit == 'ft':
    scale = int(round(int(scale)*12,-2))
    print scale

orderInfo = Oracle(init_env).call_function('getorderinfo',OrderID)
OrderNumber = orderInfo['ORDER_NUM']
site_feat = os.path.join(os.path.join(scratch,'OrderGeometry.shp'))
centroidX = eval(orderInfo['ORDER_GEOMETRY']['CENTROID'])[0][0][0]
centroidY = eval(orderInfo['ORDER_GEOMETRY']['CENTROID'])[0][0][1]

if os.path.exists(os.path.join(jobfolder,OrderNumber,'OrderGeometry.shp')):
    arcpy.CopyFeatures_management(os.path.join(jobfolder,OrderNumber,'OrderGeometry.shp'),site_feat)
else:
    OrderGeometry = createGeometry(eval(orderInfo[u'ORDER_GEOMETRY'][u'GEOMETRY'])[0],orderInfo['ORDER_GEOMETRY']['GEOMETRY_TYPE'],scratch,'OrderGeometry.shp')
mxdextent = os.path.join(scratch,'extent.mxd')
shutil.copy(mxdtemplate,mxdextent)
mxd = arcpy.mapping.MapDocument(mxdextent)
df = arcpy.mapping.ListDataFrames(mxd,'*')[0]
sr = arcpy.GetUTMFromLocation(centroidX,centroidY)
df.spatialReference = sr
geo_lyr = arcpy.mapping.Layer(site_feat)
arcpy.mapping.AddLayer(df,geo_lyr,'TOP')
geometry_layer = arcpy.mapping.ListLayers(mxd,'OrderGeometry',df)[0]
geometry_layer.visible = False
geo_extent = geometry_layer.getExtent(True)
df.extent = geo_extent
df.scale = int(scale)
arcpy.RefreshActiveView()
arcpy.mapping.ExportToJPEG(mxd,os.path.join(scratch,'extent.jpg'),df,df_export_width=170,df_export_height=220,world_file = True,jpeg_quality = 10)
arcpy.DefineProjection_management(os.path.join(scratch,'extent.jpg'),sr)
arcpy.ProjectRaster_management(os.path.join(scratch,'extent.jpg'),os.path.join(scratch,'extentwgs84.jpg'),4326)
extentdesc = arcpy.Describe(os.path.join(scratch,'extentwgs84.jpg')).extent
extentout = [[extentdesc.XMin,extentdesc.YMax],[extentdesc.XMax,extentdesc.YMax],[extentdesc.XMax,extentdesc.YMin],[extentdesc.XMin,extentdesc.YMin],[extentdesc.XMin,extentdesc.YMax]]

arcpy.AddMessage("{0}: {1}".format('Extent Output', extentout))
arcpy.SetParameter(3,extentout)