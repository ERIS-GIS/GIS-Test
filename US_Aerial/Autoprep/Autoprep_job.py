#-------------------------------------------------------------------------------
# Name:        GIS US Aerial Autoprep
# Purpose:
#
# Author:      jloucks
#
# Created:     07/20/2020
# Copyright:   (c) jloucks 2020
#-------------------------------------------------------------------------------

## Create job folder, pull images from nas, copy images to job folder/clip doqq images for process
## may need to pass back json to FE or DB
import sys
import arcpy
import cx_Oracle
import contextlib
import json
import os
import shutil
import timeit
import urllib
start1 = timeit.default_timer()
arcpy.env.overwriteOutput = True

class Machine:
    machine_test = r"\\cabcvan1gis006"
    machine_prod = r"\\cabcvan1gis007"
class Credential:
    oracle_test = r"ERIS_GIS/gis295@GMTESTC.glaciermedia.inc"
    oracle_production = r"ERIS_GIS/gis295@GMPRODC.glaciermedia.inc"
class ReportPath:
    caaerial_prod= r"\\CABCVAN1OBI007\ErisData\prod\aerial_ca"
    caaerial_test= r"\\CABCVAN1OBI007\ErisData\test\aerial_ca"
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
    oracle_functions = {'getorderinfo':"eris_gis.getOrderInfo"
    }
    erisapi_procedures = {'getaeriallist':'flow_autoprep.getAerialImageJson','passclipextent': 'flow_autoprep.setClipImageDetail'}
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
            raise Exception(("Oracle Failure",e.message.message))
        except Exception as e:
            raise Exception(("JSON Failure",e.message.message))
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
            raise Exception(("Oracle Failure",e.message.message))
        except Exception as e:
            raise Exception(e.message)
        except NameError as e:
            raise Exception("Bad Function")
        finally:
            self.close_connection()
## Custom Exceptions ##
class SingleEmptyImage(Exception):
    pass
class DoqqEmptyImage(Exception):
    pass
class OracleBadReturn(Exception):
    pass
class NoAvailableImage(Exception):
    pass
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
def export_reportimage(imagepath,ordergeometry,auid):
    ## In memory
    if os.path.exists(imagepath) == False:
        arcpy.AddWarning(imagepath+' DOES NOT EXIST')
    else:
        mxd = arcpy.mapping.MapDocument(mxdexport_template)
        df = arcpy.mapping.ListDataFrames(mxd,'*')[0]
        sr = arcpy.SpatialReference(4326)
        df.SpatialReference = sr
        lyrpath = os.path.join(scratch,str(auid) + '.lyr')
        arcpy.MakeRasterLayer_management(imagepath,lyrpath)
        image_lyr = arcpy.mapping.Layer(lyrpath)
        geo_lyr = arcpy.mapping.Layer(ordergeometry)
        arcpy.mapping.AddLayer(df,geo_lyr,'TOP')
        arcpy.mapping.AddLayer(df,image_lyr,'TOP')
        image_layer = arcpy.mapping.ListLayers(mxd,"",df)[0]
        geometry_layer = arcpy.mapping.ListLayers(mxd,'OrderGeometry',df)[0]
        geometry_layer.visible = False
        image_extent = image_layer.getExtent()
        geo_extent = geometry_layer.getExtent()
        df.extent = geo_extent
        print df.scale
        """if df.scale <= 7500 or image_source == 'DOQQ':
            if df.scale <= 7500 or image_source != 'DOQQ':
                df.extent = geo_extent
                df.scale = MapScale
                w_res = 2550
                h_res = 3300
            elif image_source == 'DOQQ' and df.scale > 7500:
                df.extent = geo_extent
                df.scale = ((df.scale/100)+1)*100
                w_res = 2550
                h_res = 3300
        elif df.scale > 7500 and image_source != 'DOQQ':
            df.extent = image_extent
            df.scale = df.scale*0.91 #very important setting as it defines how much of the image will be displayed to FE
            w_res=2550
            h_res= int((image_extent.height/image_extent.width)*w_res)"""
        if image_collection == 'DOQQ':
            df.extent = geo_extent
            df.scale = ((df.scale/100)+1)*1500
            w_res = 5100
            h_res = 6600
        elif image_collection != 'DOQQ':
            df.extent = image_extent
            df.scale = ((df.scale/100))*85 #very important setting as it defines how much of the image will be displayed to FE
            w_res=10200
            h_res= int((image_extent.height/image_extent.width)*w_res)
            #if image_extent.width < 0 and image_extent.height < 0:
                #w_res=int((image_extent.width*1000)*3)
                #h_res=int((image_extent.height*1000)*3)
            #elif image_extent.width > 1000 and image_extent.height > 1000:
                #w_res=int((image_extent.width/1000)*3)
                #h_res=int((image_extent.height/1000)*3)
            #else:
                #w_res = 5100
                #h_res = 5100
        ###############################
        ## NEED TO EXPORT DF EXTENT TO ORACLE HERE
        #sr = arcpy.SpatialReference(3857)
        #df.SpatialReference = sr
        #mxd.save()
        print image_extent.width, image_extent.height
        print w_res, h_res
        arcpy.RefreshActiveView()
        arcpy.overwriteOutput = True
        sr2 = arcpy.SpatialReference(4326)
        #decribe # of bands for raster to determine output
        desc = arcpy.Describe(lyrpath)
        bandcount = desc.bandcount
        if bandcount == 1:
            arcpy.mapping.ExportToJPEG(mxd,os.path.join(jpg_image_folder,image_year + '_' + image_source + '_' +auid + '.jpg'),df,df_export_width=w_res,df_export_height=h_res,world_file=True,color_mode = '8-BIT_GRAYSCALE', jpeg_quality = 70)
        else:
            arcpy.mapping.ExportToJPEG(mxd,os.path.join(jpg_image_folder,image_year + '_' + image_source + '_' +auid + '.jpg'),df,df_export_width=w_res,df_export_height=h_res,world_file=True,color_mode = '24-BIT_TRUE_COLOR', jpeg_quality = 70)
        arcpy.DefineProjection_management(os.path.join(jpg_image_folder,image_year + '_' + image_source + '_' +auid + '.jpg'),sr2)
        #arcpy.env.compression = "JPEG 1"
        #arcpy.env.pyramid = "NONE"
        #print "projecting"
        #arcpy.ProjectRaster_management(os.path.join(jpg_image_folder,image_year + '_' + image_source + '_' +auid + '.jpg'),os.path.join(scratch,image_year + '_' + image_source + '_' +auid + '_2.jpg'),sr)
        #extent =arcpy.Describe(os.path.join(scratch,image_year + '_' + image_source + '_' +auid + '_2.jpg')).extent
        extent =arcpy.Describe(os.path.join(jpg_image_folder,image_year + '_' + image_source + '_' +auid + '.jpg')).extent
        print "done projecting"
        NW_corner= str(extent.XMin) + ',' +str(extent.YMin)
        NE_corner= str(extent.XMax) + ',' +str(extent.YMax)
        SW_corner= str(extent.XMin) + ',' +str(extent.YMin)
        SE_corner= str(extent.XMax) + ',' +str(extent.YMin)
        print NW_corner, NE_corner, SW_corner, SE_corner
        try:
            image_extents = str({"PROCEDURE":Oracle.erisapi_procedures['passclipextent'], "ORDER_NUM" : OrderNumText,"AUI_ID":auid,"SWLAT":str(extent.YMin),"SWLONG":str(extent.XMin),"NELAT":(extent.YMax),"NELONG":str(extent.XMax)})
            message_return = Oracle('test').call_erisapi(image_extents)
            if message_return[3] != 'Y':
                raise OracleBadReturn
        except OracleBadReturn:
            arcpy.AddError('status: '+message_return[3]+' - '+message_return[2])
        ##############################
        #arcpy.mapping.ExportToJPEG(mxd,os.path.join(job_folder,'year'+'_source'+auid + '.jpg'),df,df_export_width=5100,df_export_height=6600,world_file=True,color_mode = '24-BIT_TRUE_COLOR', jpeg_quality = 50)
        #arcpy.DefineProjection_management(os.path.join(job_folder,'year'+'_source'+auid + '.jpg'), 3857)
        #shutil.copy(os.path.join(job_folder,'year'+'_source'+auid + '.jpg'),os.path.join(jpg_image_folder,auid + '.jpg'))
        mxd.saveACopy(os.path.join(scratch,auid+'_export.mxd'))
        del mxd


if __name__ == '__main__':
    start = timeit.default_timer()
    orderID = '968770'#arcpy.GetParameterAsText(0)#'934465'#arcpy.GetParameterAsText(0)
    AUI_ID = '7424603'#arcpy.GetParameterAsText(1)#''#arcpy.GetParameterAsText(1)
    scratch = r'C:\Users\JLoucks\Documents\JL\psr2'#arcpy.env.scratchFolder#r'C:\Users\JLoucks\Documents\JL\psr2'#arcpy.env.scratchFolder
    job_directory = r'\\192.168.136.164\v2_usaerial\JobData\test'
    mxdexport_template = r'\\cabcvan1gis006\GISData\Aerial_US\mxd\Aerial_US_Export.mxd'
    conversion_input = r'\\192.168.136.164\v2_usaerial\input'
    conversion_output = r'\\192.168.136.164\v2_usaerial\output'
    Conversion_URL = r'http://erisservice3.ecologeris.com/ErisInt/USAerialAppService_test/USAerial.svc/USAerialImagePromote_temp?inputfile='
    MapScale = 6000

    ##get info for order from oracle
    orderInfo = Oracle('test').call_function('getorderinfo',orderID)
    OrderNumText = str(orderInfo['ORDER_NUM'])

    ## Get order geometry & mxd
    shutil.copy(mxdexport_template,os.path.join(scratch,'template.mxd'))
    mxdexport_template = os.path.join(scratch,'template.mxd')

    job_folder = os.path.join(job_directory,OrderNumText)
    if AUI_ID == '':
        ## Return aerial list from oracle
        oracle_autoprep = str({"PROCEDURE":Oracle.erisapi_procedures['getaeriallist'],"ORDER_NUM":OrderNumText})
        aerial_list_return = Oracle('test').call_erisapi(oracle_autoprep)
        aerial_list_json = json.loads(aerial_list_return[1])

        ## Seperate processes for singleframe and DOQQ
        single_image_candidates = aerial_list_json['INHOUSE_IMAGE']
        doqq_image_candidates = aerial_list_json['DOQQ_IMAGE']
        index_image_candidates = aerial_list_json['INDEX_IMAGES']

        ##Create job folder and copy images
        try:
            org_image_folder = os.path.join(job_folder,'org')
            jpg_image_folder = os.path.join(job_folder,'jpg')
            if os.path.exists(job_folder):
                shutil.rmtree(job_folder)
            os.mkdir(job_folder)
            os.mkdir(org_image_folder)
            os.mkdir(jpg_image_folder)
            OrderGeometry = createGeometry(eval(orderInfo[u'ORDER_GEOMETRY'][u'GEOMETRY'])[0],orderInfo['ORDER_GEOMETRY']['GEOMETRY_TYPE'],job_folder,'OrderGeometry.shp')
            if len(doqq_image_candidates) == 0 and len(single_image_candidates) == 0:
                raise NoAvailableImage
            if len(single_image_candidates) == 0:
                arcpy.AddWarning('No singleframe image candidates')
            
            try:
                for inhouse_image in single_image_candidates:
                    image_auid = str(inhouse_image['AUI_ID'])
                    image_name = inhouse_image['ORIGINAL_IMAGEPATH']
                    image_year = str(inhouse_image['AERIAL_YEAR'])
                    image_source = inhouse_image['IMAGE_SOURCE']
                    image_collection = inhouse_image['IMAGE_COLLECTION_TYPE']
                    selected_flag = inhouse_image['SELECTED_FLAG']
                    if image_source == '':
                        image_source = 'UNKWN'
                    if selected_flag == 'Y':
                        export_reportimage(image_name,OrderGeometry,image_auid)


                if len(doqq_image_candidates) == 0:
                    arcpy.AddWarning('No DOQQ image candidates')

                for inhouse_image in doqq_image_candidates:
                    image_auid = str(inhouse_image['AUI_ID'])
                    image_name = inhouse_image['ORIGINAL_IMAGEPATH']
                    image_year = str(inhouse_image['AERIAL_YEAR'])
                    image_source = inhouse_image['IMAGE_SOURCE']
                    image_collection = inhouse_image['IMAGE_COLLECTION_TYPE']
                    selected_flag = inhouse_image['SELECTED_FLAG']
                    if image_source == '':
                        image_source = 'UNKWN'
                    if selected_flag == 'Y':
                        export_reportimage(image_name,OrderGeometry,image_auid)
                
                if len(index_image_candidates) == 0:
                    arcpy.AddWarning('No INDEX image candidates')

                for inhouse_image in index_image_candidates:
                    image_auid = str(inhouse_image['AUI_ID'])
                    image_name = inhouse_image['ORIGINAL_IMAGEPATH']
                    image_year = str(inhouse_image['AERIAL_YEAR'])
                    image_source = inhouse_image['IMAGE_SOURCE']
                    image_collection = inhouse_image['IMAGE_COLLECTION_TYPE']
                    selected_flag = inhouse_image['SELECTED_FLAG']
                    if image_source == '':
                        image_source = 'UNKWN'
                    if selected_flag == 'Y':
                        export_reportimage(image_name,OrderGeometry,image_auid)
            except KeyError as k:
                arcpy.AddError('JSON missing key: ' + k.message)
        except NoAvailableImage:
            arcpy.AddError('No available images for location')
            sys.exit()
    else:
        AUI_IDtext = str(AUI_ID)
        oracle_singleprep = str({"PROCEDURE":Oracle.erisapi_procedures['getaeriallist'],"ORDER_NUM":OrderNumText,"AUI_ID":AUI_IDtext})
        aerial_list_return = Oracle('test').call_erisapi(oracle_singleprep)
        aerial_list_json = json.loads(aerial_list_return[1])

        single_image_candidates = aerial_list_json['INHOUSE_IMAGE']
        doqq_image_candidates = aerial_list_json['DOQQ_IMAGE']
        index_image_candidates = aerial_list_json['INDEX_IMAGES']

        if os.path.exists(job_folder) == False:
            arcpy.AddError('Job Folder does not exist - Reinitialize order')
            sys.exit()

        org_image_folder = os.path.join(job_folder,'org')
        jpg_image_folder = os.path.join(job_folder,'jpg')
        OrderGeometry = os.path.join(job_folder,'OrderGeometry.shp')

        if len(single_image_candidates) == 1:
            for inhouse_image in single_image_candidates:
                image_auid = str(inhouse_image['AUI_ID'])
                image_name = inhouse_image['ORIGINAL_IMAGEPATH']
                image_year = str(inhouse_image['AERIAL_YEAR'])
                image_source = inhouse_image['IMAGE_SOURCE']
                image_collection = inhouse_image['IMAGE_COLLECTION_TYPE']
                selected_flag = inhouse_image['SELECTED_FLAG']
                if image_source == '':
                    image_source = 'UNKWN'
                export_reportimage(image_name,OrderGeometry,image_auid)
        elif len(doqq_image_candidates) == 1:
            for inhouse_image in doqq_image_candidates:
                image_auid = str(inhouse_image['AUI_ID'])
                image_name = inhouse_image['ORIGINAL_IMAGEPATH']
                image_year = str(inhouse_image['AERIAL_YEAR'])
                image_source = inhouse_image['IMAGE_SOURCE']
                image_collection = inhouse_image['IMAGE_COLLECTION_TYPE']
                selected_flag = inhouse_image['SELECTED_FLAG']
                if image_source == '':
                    image_source = 'UNKWN'
                export_reportimage(image_name,OrderGeometry,image_auid)
        elif len(index_image_candidates) == 1:
            for inhouse_image in index_image_candidates:
                image_auid = str(inhouse_image['AUI_ID'])
                image_name = inhouse_image['ORIGINAL_IMAGEPATH']
                image_year = str(inhouse_image['AERIAL_YEAR'])
                image_source = inhouse_image['IMAGE_SOURCE']
                image_collection = inhouse_image['IMAGE_COLLECTION_TYPE']
                selected_flag = inhouse_image['SELECTED_FLAG']
                if image_source == '':
                    image_source = 'UNKWN'
                export_reportimage(image_name,OrderGeometry,image_auid)
        else:
            arcpy.AddError('No Available Image for that AUI ID')