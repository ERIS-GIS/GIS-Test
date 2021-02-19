import os, arcpy, shutil
import cx_Oracle
import json

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
    erisapi_procedures = {'getreworkaerials':"FLOW_GEOREFERENCE.GetAerialnewAdd","passimagedetail":"flow_inventory.setImageDetail","setimagename":"Flow_inventory.setImageName"}
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
def export_reportimage(imagepath,auid):
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
        arcpy.mapping.AddLayer(df,image_lyr,'TOP')
        img_extent = image_lyr.getExtent()
        print img_extent
        df.extent = img_extent
        arcpy.RefreshActiveView()
        ###############################
        ## NEED TO EXPORT DF EXTENT TO ORACLE HERE
        print 'exporting image'
        arcpy.mapping.ExportToJPEG(mxd,os.path.join(scratch,'temp.jpg'),df,world_file=True,color_mode = '24-BIT_TRUE_COLOR', jpeg_quality = 10)
        arcpy.DefineProjection_management(os.path.join(scratch,'temp.jpg'), 3857)
        print 'done exporting'
        extent =arcpy.Describe(os.path.join(scratch,'temp.jpg')).extent
        NW_corner= str(extent.XMin) + ',' +str(extent.YMin)
        NE_corner= str(extent.XMax) + ',' +str(extent.YMax)
        SW_corner= str(extent.XMin) + ',' +str(extent.YMin)
        SE_corner= str(extent.XMax) + ',' +str(extent.YMin)
        print NW_corner, NE_corner, SW_corner, SE_corner
        ##############################
        #arcpy.mapping.ExportToJPEG(mxd,os.path.join(job_folder,'year'+'_source'+auid + '.jpg'),df,df_export_width=5100,df_export_height=6600,world_file=True,color_mode = '24-BIT_TRUE_COLOR', jpeg_quality = 50)
        #arcpy.DefineProjection_management(os.path.join(job_folder,'year'+'_source'+auid + '.jpg'), 3857)
        #shutil.copy(os.path.join(job_folder,'year'+'_source'+auid + '.jpg'),os.path.join(jpg_image_folder,auid + '.jpg'))
        mxd.saveACopy(os.path.join(scratch,str(auid)+'_export.mxd'))
        return extent
        del mxd

if __name__ == '__main__':
    OrderID = arcpy.GetParameterAsText(0)#'934404'#arcpy.GetParameterAsText(0)
    AUI_ID = arcpy.GetParameterAsText(1)
    ee_oid = arcpy.GetParameterAsText(2)#'408212'#arcpy.GetParameterAsText(2)
    scratch = arcpy.env.scratchFolder#r'C:\Users\JLoucks\Documents\JL\test1'#arcpy.env.scratchFolder
    arcpy.CreateFileGDB_management(scratch, 'temp.gdb')
    job_directory = r'\\192.168.136.164\v2_usaerial\JobData\test'
    georeferenced_historical = r'\\cabcvan1nas003\historical\Georeferenced_Aerial_test'
    georeferenced_doqq = r'\\cabcvan1nas003\doqq\Georeferenced_DOQQ_test'
    mxdexport_template = r'\\cabcvan1gis006\GISData\Aerial_US\mxd\Aerial_US_Export.mxd'
    arcpy.env.OverwriteOutput = True

    orderInfo = Oracle('test').call_function('getorderinfo',OrderID)
    OrderNumText = str(orderInfo['ORDER_NUM'])
    job_folder = os.path.join(job_directory,OrderNumText)
    uploaded_dir = os.path.join(job_folder,"OrderImages")

    ### Get image path info ###
    inv_infocall = str({"PROCEDURE":Oracle.erisapi_procedures['getreworkaerials'],"ORDER_NUM":str(OrderNumText),"AUI_ID":str(AUI_ID),"PARENT_EE_OID":str(ee_oid)})
    rework_return = Oracle('test').call_erisapi(inv_infocall)
    rework_list_json = json.loads(rework_return[1])
    print rework_list_json

    if rework_list_json == []:
        arcpy.AddError('Image list is empty')
    try:
        for image in rework_list_json:
            auid = image['AUI_ID']
            imagename = image['IMAGE_NAME']
            aerialyear = image['AERIAL_YEAR']
            imagesource = image['IMAGE_SOURCE']
            imagecollection = image['IMAGE_COLLECTION_TYPE']
            originalpath = image['ORIGINAL_IMAGE_PATH']

            imageuploadpath = originalpath
            TAB_upload_path = imageuploadpath.split('.')[0]+'.TAB'
            #job_image_name = str(aerialyear)+'_'+imagesource+'_'+str(auid)+'.'+str(imagename.split('.')[-1])
            job_image_name = str(aerialyear)+'_'+imagesource+'_'+str(auid)+'.'+str(originalpath[-5:].split('.')[1])
            TAB_image_name = str(aerialyear)+'_'+imagesource+'_'+str(auid)+'.TAB'

            if imagecollection == 'DOQQ':
                mosaicfp = os.path.join(scratch,'image_boundary.shp')
                arcpy.CreateMosaicDataset_management(os.path.join(scratch,'temp.gdb'), 'doqq', 4326)
                arcpy.AddRastersToMosaicDataset_management (os.path.join(scratch,'temp.gdb','doqq'), "Raster Dataset", imageuploadpath, 'NO_CELL_SIZES', True, False)
                arcpy.ExportMosaicDatasetGeometry_management (os.path.join(scratch,'temp.gdb','doqq'), mosaicfp,geometry_type = 'BOUNDARY')
                cellsizeX = arcpy.GetRasterProperties_management(imageuploadpath,'CELLSIZEX')
                cellsizeY = arcpy.GetRasterProperties_management(imageuploadpath,'CELLSIZEY')
                if cellsizeY > cellsizeX:
                    spatial_res = cellsizeY
                else:
                    spatial_res = cellsizeX
                desc = arcpy.Describe(mosaicfp)
                result_top = desc.extent.YMax
                result_bot = desc.extent.YMin
                result_left = desc.extent.XMin
                result_right = desc.extent.XMax
                arcpy.AddMessage(result_top)
                arcpy.AddMessage(result_bot)
                arcpy.AddMessage(result_left)
                arcpy.AddMessage(result_right)
                #Rename image and TAB
                if os.path.exists(TAB_upload_path):
                    shutil.copy(os.path.join(uploaded_dir,TAB_image_name),os.path.join(georeferenced_doqq,TAB_image_name)) #copy TAB if exists
                #Copy image to inventory folder/
                arcpy.Copy_management(originalpath,os.path.join(scratch,job_image_name))
                if os.path.exists(os.path.join(georeferenced_doqq,job_image_name)):
                    arcpy.Delete_management(os.path.join(georeferenced_doqq,job_image_name))
                arcpy.Copy_management(os.path.join(scratch,job_image_name),os.path.join(georeferenced_doqq,job_image_name))
                image_inv_path = os.path.join(georeferenced_doqq,job_image_name)
            else:
                mosaicfp = os.path.join(scratch,'image_boundary.shp')
                arcpy.CreateMosaicDataset_management(os.path.join(scratch,'temp.gdb'), 'raster', 4326)
                arcpy.AddRastersToMosaicDataset_management (os.path.join(scratch,'temp.gdb','raster'), "Raster Dataset", imageuploadpath, 'NO_CELL_SIZES', True, False)
                arcpy.ExportMosaicDatasetGeometry_management (os.path.join(scratch,'temp.gdb','raster'), mosaicfp,geometry_type = 'BOUNDARY')
                cellsizeX = arcpy.GetRasterProperties_management(imageuploadpath,'CELLSIZEX')
                cellsizeY = arcpy.GetRasterProperties_management(imageuploadpath,'CELLSIZEY')
                if cellsizeY > cellsizeX:
                    spatial_res = cellsizeY
                else:
                    spatial_res = cellsizeX
                desc = arcpy.Describe(mosaicfp)
                result_top = desc.extent.YMax
                result_bot = desc.extent.YMin
                result_left = desc.extent.XMin
                result_right = desc.extent.XMax
                arcpy.AddMessage(result_top)
                arcpy.AddMessage(result_bot)
                arcpy.AddMessage(result_left)
                arcpy.AddMessage(result_right)
                #Rename image and TAB
                if os.path.exists(TAB_upload_path):
                    shutil.copy(os.path.join(uploaded_dir,TAB_image_name),os.path.join(georeferenced_historical,TAB_image_name)) #copy TAB if exists
                #Copy image to inventory folder/
                arcpy.Copy_management(originalpath,os.path.join(scratch,job_image_name))
                if os.path.exists(os.path.join(georeferenced_historical,job_image_name)):
                    arcpy.Delete_management(os.path.join(georeferenced_historical,job_image_name))
                arcpy.Copy_management(os.path.join(scratch,job_image_name),os.path.join(georeferenced_historical,job_image_name))
                image_inv_path = os.path.join(georeferenced_historical,job_image_name)                
            image_metadata = str({"PROCEDURE":Oracle.erisapi_procedures['passimagedetail'],"ORDER_NUM":OrderNumText,"AUI_ID":str(auid),"SWLAT":str(result_bot),"SWLONG":str(result_left),"NELAT":str(result_top),"NELONG":str(result_right),"SPATIAL_RESOLUTION":str(spatial_res),"ORIGINAL_IMAGE_PATH":str(image_inv_path)})
            Oracle('test').call_erisapi(image_metadata)
            new_image_name = str(aerialyear)+'_'+imagesource+'_'+str(auid)+'.'+imagename.split('.')[1]
            rename_call = str({"PROCEDURE":Oracle.erisapi_procedures['setimagename'],"ORDER_NUM":OrderNumText,"AUI_ID":auid,"IMAGE_NAME":str(new_image_name)})
            rename_return = Oracle('test').call_erisapi(rename_call)
            print json.loads(rework_return[1])
    except IOError as e:
        arcpy.AddError('Issue converting image: '+e.message)