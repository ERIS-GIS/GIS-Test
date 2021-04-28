import os, arcpy, shutil
import cx_Oracle
import json
from PIL import Image

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
    erisapi_procedures = {'getnewaerials':"FLOW_GEOREFERENCE.GetAerialnewPNG","setimagename":"Flow_inventory.setImageName"}
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

if __name__ == '__main__':
    OrderID = arcpy.GetParameterAsText(0)#'983883'
    AUI_ID = arcpy.GetParameterAsText(1)#''#arcpy.GetParameterAsText(1)'7367074'
    ee_oid = arcpy.GetParameterAsText(2)#'408212'#arcpy.GetParameterAsText(2)
    scratch = arcpy.env.scratchFolder#r'C:\Users\JLoucks\Documents\JL\test2'#arcpy.env.scratchFolder
    job_directory = r'\\192.168.136.164\v2_usaerial\JobData\test'
    arcpy.env.OverwriteOutput = True
    Image.MAX_IMAGE_PIXELS = 10000000000

    orderInfo = Oracle('test').call_function('getorderinfo',OrderID)
    OrderNumText = str(orderInfo['ORDER_NUM'])
    job_folder = os.path.join(job_directory,OrderNumText)
    uploaded_dir = os.path.join(job_folder,"OrderImages")

    ### Get image path info ###
    inv_infocall = str({"PROCEDURE":Oracle.erisapi_procedures['getnewaerials'],"ORDER_NUM":str(OrderNumText),"AUI_ID":str(AUI_ID),"PARENT_EE_OID":str(ee_oid)})
    rework_return = Oracle('test').call_erisapi(inv_infocall)
    rework_list_json = json.loads(rework_return[1])
    print rework_list_json
    if rework_list_json == []:
        arcpy.AddError('Image list is empty')

    try:
        if not os.path.exists(os.path.join(job_folder,'gc')):
            os.mkdir(os.path.join(job_folder,'gc'))
        for image in rework_list_json:
            auid = image['AUI_ID']
            imagename = image['IMAGE_NAME']
            aerialyear = image['AERIAL_YEAR']
            imagesource = image['IMAGE_SOURCE']
            imagecollection = image['IMAGE_COLLECTION_TYPE']
            originalpath = image['ORIGINAL_IMAGE_PATH']
            imageuploadpath = originalpath
            if imagecollection == 'DOQQ':
                arcpy.AddWarning('Cannot convert DOQQ image '+originalpath)
            else:
                if os.path.exists(imageuploadpath):
                    job_image_name = str(aerialyear)+'_'+imagesource+'_'+str(auid)+'.jpg'
                    """PNG is copied to gc folder for FE with new naming convention"""
                    if os.path.exists(os.path.join(job_folder,'gc',job_image_name)):
                        os.remove(os.path.join(job_folder,'gc',job_image_name))
                    if imageuploadpath[-3:] in ['jpg','png','tif']:
                        im = Image.open(imageuploadpath)
                        subject_im = im.convert("RGB")
                        subject_im.save(os.path.join(job_folder,'gc',job_image_name))
                    else:
                        arcpy.env.compression = "JPEG 50"
                        arcpy.CopyRaster_management(imageuploadpath,os.path.join(job_folder,'gc',job_image_name),colormap_to_RGB='ColormapToRGB',pixel_type='8_BIT_UNSIGNED',format='JPEG',transform='NONE')
                    """Rename original uploaded file with new naming convention
                    BUT ORIGINAL EXTENSION!!! And call oracle to update the name. Path to the image
                    will be updated once georeferencing in complete in that gp service"""
                    new_image_name = str(aerialyear)+'_'+imagesource+'_'+str(auid)+'.'+imagename.split('.')[1]
#                    if imageuploadpath == os.path.join(uploaded_dir,new_image_name):
#                        arcpy.AddMessage('Image name already matches naming convention rules')
#                    else:
#                        os.rename(imageuploadpath,os.path.join(uploaded_dir,new_image_name))
                    rename_call = str({"PROCEDURE":Oracle.erisapi_procedures['setimagename'],"ORDER_NUM":OrderNumText,"AUI_ID":auid,"IMAGE_NAME":str(new_image_name)})
                    rename_return = Oracle('test').call_erisapi(rename_call)
                    print json.loads(rework_return[1])
                elif not os.path.exists(imageuploadpath):
                    arcpy.AddError('cannot find image in OrderImages folder to convert, PLEASE CHECK PATH: '+imageuploadpath)
    except Exception as e:
        arcpy.AddError('Issue converting image: '+str(e.message, imageuploadpath))