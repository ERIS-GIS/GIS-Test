import sys
import os
import arcpy
import cx_Oracle
import json
import shutil
import timeit
import os
import json
reload(sys)

class Machine:
    machine_test = r"\\cabcvan1gis006"
    machine_prod = r"\\cabcvan1gis007"
class Credential:
    oracle_test = r"ERIS_GIS/gis295@GMTESTC.glaciermedia.inc"
    oracle_production = r"ERIS_GIS/gis295@GMPRODC.glaciermedia.inc"
class ImageBasePath:
    caaerial_test= r"\\CABCVAN1OBI007\ErisData\test\aerial_ca"
    caaerial_prod= r"\\CABCVAN1OBI007\ErisData\prod\aerial_ca"
class OutputDirectory:
    job_directory_test = r'\\192.168.136.164\v2_usaerial\JobData\test'
    job_directory_prod = r'\\192.168.136.164\v2_usaerial\JobData\prod'
    georef_images_test = r'\\cabcvan1nas003\historical\Georeferenced_Aerial_test'
    georef_images_prod = r'\\cabcvan1nas003\historical\Georeferenced_Aerial'
class TransformationType():
    POLYORDER0 = "POLYORDER0"
    POLYORDER1 = "POLYORDER1"
    POLYORDER2 = "POLYORDER2"
    POLYORDER3 = "POLYORDER3"
    SPLINE = "ADJUST SPLINE"
    PROJECTIVE = "PROJECTIVE "
class ResamplingType():
    NEAREST  = "NEAREST"
    BILINEAR = "BILINEAR"
    CUBIC = "CUBIC"
    MAJORITY = "MAJORITY"
## Custom Exceptions ##
class OracleBadReturn(Exception):
    pass
class NoAvailableImage(Exception):
    pass
class Oracle:
    #static variable: oracle_functions
    oracle_functions = {'getorderinfo':"eris_gis.getOrderInfo"}
    erisapi_procedures = {'getGeoreferencingInfo':'flow_gis.getGeoreferencingInfo','passclipextent': 'flow_autoprep.setClipImageDetail',
                          'getImageInventoryInfo':'flow_gis.getImageInventoryInfo','UpdateInventoryImagePath':'Flow_gis.UpdateInventoryImagePath'}
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
    def call_function(self,function_name,order_id):
        self.connect_to_oracle()
        cursor = self.cursor
        try:
            outType = cx_Oracle.CLOB
            func = [self.oracle_functions[_] for _ in self.oracle_functions.keys() if function_name.lower() ==_.lower()]
            if func !=[] and len(func)==1:
                try:
                    if type(order_id) !=list:
                        order_id = [order_id]
                    output=json.loads(cursor.callfunc(func[0],outType,order_id).read())
                except ValueError:
                    output = cursor.callfunc(func[0],outType,order_id).read()
                except AttributeError:
                    output = cursor.callfunc(func[0],outType,order_id)
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
        self.connect_to_oracle()
        cursor = self.cursor
        arg1 = erisapi_input
        arg2 = cursor.var(cx_Oracle.CLOB)
        arg3 = cursor.var(cx_Oracle.CLOB) ## Message
        arg4 = cursor.var(str)  ## Status
        try:
            func = ['eris_api.callOracle']
            if func !=[] and len(func)==1:
                try:
                    output = cursor.callproc(func[0],[arg1,arg2,arg3,arg4])
                except ValueError:
                    output = cursor.callproc(func[0],[arg1,arg2,arg3,arg4])
                except AttributeError:
                    output = cursor.callproc(func[0],[arg1,arg2,arg3,arg4])
            return output[0],cx_Oracle.LOB.read(output[1]),cx_Oracle.LOB.read(output[2]),output[3]
        except cx_Oracle.Error as e:
            raise Exception(("Oracle Failure",e.message))
        except Exception as e:
            raise Exception(("JSON Failure",e.message))
        except NameError as e:
            raise Exception("Bad Function")
        finally:
            self.close_connection()
def CoordToString(inputObj):
    coordPts_string = ""
    for i in range(len(inputObj)-1):
        coordPts_string +=  "'" + " ".join(str(i) for i in  inputObj[i]) + "';"
    result =  coordPts_string[:-1]
    return result
def get_file_size(image_path):
    return int(os.stat(image_path).st_size)
def createGeometry(pntCoords,geometry_type,output_folder,output_name, spatialRef = arcpy.SpatialReference(4326)):
    outputFC = os.path.join(output_folder,output_name)
    if geometry_type.lower()== 'point':
        arcpy.CreateFeatureclass_management(output_folder, output_name, "MULTIPOINT", "", "DISABLED", "DISABLED", spatialRef)
        cursor = arcpy.da.InsertCursor(outputFC, ['SHAPE@'])
        cursor.insertRow([arcpy.Multipoint(arcpy.Array([arcpy.Point(*coords) for coords in pntCoords]),spatialRef)])
    elif geometry_type.lower() =='polyline':        
        arcpy.CreateFeatureclass_management(output_folder, output_name, "POLYLINE", "", "DISABLED", "DISABLED", spatialRef)
        cursor = arcpy.da.InsertCursor(outputFC, ['SHAPE@'])
        cursor.insertRow([arcpy.Polyline(arcpy.Array([arcpy.Point(*coords) for coords in pntCoords]),spatialRef)])
    elif geometry_type.lower() =='polygon':
        arcpy.CreateFeatureclass_management(output_folder,output_name, "POLYGON", "", "DISABLED", "DISABLED", spatialRef)
        cursor = arcpy.da.InsertCursor(outputFC, ['SHAPE@'])
        cursor.insertRow([arcpy.Polygon(arcpy.Array([arcpy.Point(*coords) for coords in pntCoords]),spatialRef)])
    del cursor
    return outputFC
def apply_georeferencing(inputRaster,srcPoints,gcpPoints,output_folder,out_img_name,transformation_type, resampling_type):
    arcpy.AddMessage('Start Georeferencing...')
    sr_wgs_84 = arcpy.SpatialReference(4326)
    # georeference to WGS84
    output_geo_ref_image = os.path.join(output_folder,out_img_name + '.tif')
    
    arcpy.Warp_management(inputRaster, srcPoints,gcpPoints, output_geo_ref_image, transformation_type, resampling_type)
    # arcpy.Warp_management(inputRaster, srcPoints,gcpPoints,output_geo_ref_image)
    
    arcpy.DefineProjection_management(output_geo_ref_image,sr_wgs_84)
    set_raster_background(output_geo_ref_image)
    arcpy.AddMessage('Output Image(.tif): %s' % output_geo_ref_image)
    arcpy.AddMessage('--Georeferencing Done.')
    return output_geo_ref_image
def export_to_jpg(env,imagepath,outputImage_jpg,order_geometry,auid):
    mxd = arcpy.mapping.MapDocument(mxdexport_template)
    df = arcpy.mapping.ListDataFrames(mxd,'*')[0]
    sr_wgs84 = arcpy.SpatialReference(4326)
    df.SpatialReference = sr_wgs84
    lyrpath = os.path.join(arcpy.env.scratchFolder,auid + '.lyr')
    arcpy.MakeRasterLayer_management(imagepath,lyrpath)
    image_lyr = arcpy.mapping.Layer(lyrpath)
    geo_lyr = arcpy.mapping.Layer(order_geometry)
    arcpy.mapping.AddLayer(df,image_lyr,'TOP')
    arcpy.mapping.AddLayer(df,geo_lyr,'TOP')
    geometry_layer = arcpy.mapping.ListLayers(mxd,'OrderGeometry',df)[0]
    geometry_layer.visible = False
    geo_extent = geometry_layer.getExtent(True)
    image_extent = geo_lyr.getExtent(True)
    df.extent = image_extent
    if df.scale <= MapScale:
        df.scale = MapScale
        export_width = 5100
        export_height = 6600
    elif df.scale > MapScale:
        df.scale = ((int(df.scale)/100)+1)*100
        export_width = int(5100*1.4)
        export_height = int(6600*1.4)

    arcpy.RefreshActiveView()
    message_return = None
    try:
        image_extents = str({"PROCEDURE":Oracle.erisapi_procedures['passclipextent'], "ORDER_NUM" : order_num,"AUI_ID":auid,"SWLAT":str(df.extent.YMin),"SWLONG":str(df.extent.XMin),"NELAT":(df.extent.XMax),"NELONG":str(df.extent.XMax)})
        message_return = Oracle(env).call_erisapi(image_extents)
        if message_return[3] != 'Y':
            raise OracleBadReturn
    except OracleBadReturn:
        arcpy.AddError('status: ' + message_return[3]+ ' - '+ message_return[2])

    mxd.saveACopy(os.path.join(arcpy.env.scratchFolder,auid+'_export.mxd'))
    arcpy.mapping.ExportToJPEG(mxd,outputImage_jpg,df,df_export_width=export_width,df_export_height=export_height,world_file=False,color_mode = '24-BIT_TRUE_COLOR', jpeg_quality = 50)
    del mxd
def set_raster_background(input_raster):
    desc = arcpy.Describe(input_raster)
    for i in range(desc.bandCount):
        arcpy.SetRasterProperties_management(input_raster ,nodata= str(i+1) + ' 255')
def export_to_outputs(env,geroref_Image,outputImage_jpg,out_img_name,orderGeometry):
    ### Export georefed image as jpg file to jpg folder for US Aerial UI app
    output_image_jpg = os.path.join(outputImage_jpg,out_img_name + '.jpg')
    export_to_jpg(env,geroref_Image,output_image_jpg,orderGeometry,str(auid))
    arcpy.AddMessage('Output Image(.jpg): %s' % output_image_jpg)
    set_raster_background(output_image_jpg)
if __name__ == '__main__':
    ### set input parameters
    order_id = arcpy.GetParameterAsText(0)
    auid = arcpy.GetParameterAsText(1)
    order_id = '1106337'
    auid = '1056732'
    env = 'test'
    ## set scratch folder
    scratch_folder = arcpy.env.scratchFolder
    # arcpy.env.workspace = scratchFolder
    arcpy.env.overwriteOutput = True 
    arcpy.env.pyramid = 'NONE'
    if str(order_id) != '' and str(auid) != '':
        mxdexport_template = r'\\cabcvan1gis006\GISData\Aerial_US\mxd\Aerial_US_Export.mxd'
        MapScale = 6000
        try:
            start = timeit.default_timer()

            ##get info for order from oracle
            orderInfo = Oracle(env).call_function('getorderinfo',str(order_id))
            order_num = str(orderInfo['ORDER_NUM'])
            job_folder = ''
            if env == 'test':
                job_folder = os.path.join(OutputDirectory.job_directory_test,order_num)
            elif env == 'prod':
                job_folder = os.path.join(OutputDirectory.job_directory_prod,order_num)
            ### get georeferencing info from oracle
            oracle_georef = str({"PROCEDURE":Oracle.erisapi_procedures["getGeoreferencingInfo"],"ORDER_NUM": order_num, "AUI_ID": str(auid)})
            aerial_us_georef = Oracle(env).call_erisapi(oracle_georef)
            aerial_georefjson = json.loads(aerial_us_georef[1])
            if  (len(aerial_georefjson)) == 0:
                arcpy.AddWarning('The  georeferencing information is not availabe!')
                if len(aerial_georefjson) > 0 and len(aerial_georefjson[2]) > 0:
                    arcpy.AddWarning(aerial_georefjson[2])
            else:  
                org_image_folder = os.path.join(job_folder,'org')
                jpg_image_folder = os.path.join(job_folder,'jpg')
                gc_image_folder = os.path.join(job_folder,'gc')
                if not os.path.exists(job_folder):
                #     shutil.rmtree(job_folder)
                    os.mkdir(job_folder)
                    os.mkdir(org_image_folder)
                    os.mkdir(jpg_image_folder)  
                ### get input image from inventory
                aerial_inventory = str({"PROCEDURE":Oracle.erisapi_procedures["getImageInventoryInfo"],"ORDER_NUM":order_num , "AUI_ID": str(auid)})
                aerial_us_inventory = Oracle(env).call_erisapi(aerial_inventory)
                aerial_inventoryjson = json.loads(aerial_us_inventory[1])
                
                if  (len(aerial_inventoryjson)) == 0:
                    arcpy.AddWarning('There is no data for Image in inventory!')
                    arcpy.AddWarning(aerial_us_inventory[2])
                else:
                    image_input_path_inv = aerial_inventoryjson[0]['ORIGINAL_IMAGEPATH'] # image path from inventory  RAW_IMAGEPATH
                    result_input_image_job = os.path.join(job_folder,'gc',aerial_georefjson['imgname'])
                    order_geometry = os.path.join(job_folder,'OrderGeometry.shp')
                    year = aerial_inventoryjson[0]['AERIAL_YEAR'] 
                    img_source = aerial_inventoryjson[0]['IMAGE_SOURCE']
                    ## setup image custom name year_DOQQ_AUI_ID
                    out_img_name = '%s_%s_%s'%(year,img_source,str(auid))
                    ## Read input image from job folder if there is a clip image, use it as input. If not use the original image in job folder)
                    raw_input_image_job = os.path.join(gc_image_folder , out_img_name + '.jpg')
                    if os.path.exists(result_input_image_job): ## this is the image which is alreary georeferenced by us aerial app
                        input_image = result_input_image_job
                    else: # read from inventory
                        arcpy.AddMessage('Input image is not availabe in the job folder!')
                    arcpy.AddMessage('Input Image : %s' % input_image)
                    gcp_points = CoordToString(aerial_georefjson['envelope']) # footPrint
                    ### Source point from input extent
                    top = str(arcpy.GetRasterProperties_management(input_image,"TOP").getOutput(0))
                    left = str(arcpy.GetRasterProperties_management(input_image,"LEFT").getOutput(0))
                    right = str(arcpy.GetRasterProperties_management(input_image,"RIGHT").getOutput(0))
                    bottom = str(arcpy.GetRasterProperties_management(input_image,"BOTTOM").getOutput(0))
                    src_points = "'" + left + " " + bottom + "';" + "'" + right + " " + bottom + "';" + "'" + right + " " + top + "';" + "'" + left + " " + top + "'"
                    ### Georeferencing
                    img_georeferenced = apply_georeferencing(input_image, src_points, gcp_points,OutputDirectory.georef_images_prod, out_img_name, '', ResamplingType.BILINEAR)
                    ### ExportToOutputs
                    export_to_outputs(env,img_georeferenced, jpg_image_folder,out_img_name,order_geometry)
                    
                    file_size = get_file_size(img_georeferenced)
                    strprod_update_path = str({"PROCEDURE":Oracle.erisapi_procedures["UpdateInventoryImagePath"],"AUI_ID": str(auid), "ORIGINAL_IMAGEPATH":img_georeferenced, "FILE_SIZE":file_size})
                    message_return = Oracle(env).call_erisapi(strprod_update_path.replace('u','')) ## remove unicode chrachter u from json before calling strprod
            end = timeit.default_timer()
            arcpy.AddMessage(('Duration:', round(end -start,4)))
        except:
            msgs = "ArcPy ERRORS:\n %s\n"%arcpy.GetMessages(2)
            arcpy.AddError(msgs)
            raise
    else:
        arcpy.AddWarning('Order Id and Auid are not availabe')
    