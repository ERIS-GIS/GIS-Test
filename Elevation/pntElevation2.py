#
# Author:      cchen
#
# Created:     10/01/2019
# Copyright:   (c) cchen 2019
# Licence:     <your licence>
#-------------------------------------------------------------------------------
import os,arcpy,sys,traceback
import dem_footprints
import dem_imgs_ll
import math
import json as simplejson
import urllib, contextlib
US_DEM=r'US_DEM'
class Machine:
    data_server = r"\\cabcvan1fpr009"
class DEM():
    def __init__(self):
        self.imgdir_dem = os.path.join(Machine.data_server,US_DEM,"DEM13")
        self.imgdir_demCA = os.path.join(Machine.data_server,US_DEM,"DEM1")
        self.masterlyr_dem = os.path.join(Machine.data_server,US_DEM,"DEM13.shp")
        self.masterlyr_demCA = os.path.join(Machine.data_server,US_DEM,"DEM1.shp")
def createSHP(pntPairs):
    try:
        outshpP = os.path.join(r'in_memory',r'outshpP')
        arcpy.Delete_management(outshpP)
        _ = outshpP
        srGCS83 = arcpy.SpatialReference(4269)
        arcpy.CreateFeatureclass_management(r'in_memory', r"outshpP", "MULTIPOINT", "", "DISABLED", "DISABLED", srGCS83)
        cursor = arcpy.da.InsertCursor(_, ['SHAPE@'])
        cursor.insertRow([arcpy.Multipoint(arcpy.Array([arcpy.Point(*coords) for coords in pntPairs.values()]),srGCS83)])

        return _
    except Exception as e:
        print e.message
        return _
def findDEMshp(pntshp,masterfile,dempath):
    try:
        _={}
        masterLayer_dem = arcpy.mapping.Layer(masterfile)
        arcpy.SelectLayerByLocation_management(masterLayer_dem, 'intersect', pntshp)

        if int((arcpy.GetCount_management(masterLayer_dem).getOutput(0))) != 0:
            columns = arcpy.SearchCursor(masterLayer_dem)
            for column in columns:
                img = column.getValue("image_name")
                if img.strip() !="" and img[:-4] not in _.keys():
                    _[img[:-4]]=os.path.join(dempath,img)
            del column
            del columns
        return _
    except:
        return _
def getElevation(pntPair,imglist,key):
    temp={}
    keysDone=[]
    if imglist !={}:
        exec("mch = dem_imgs_ll.mch%s"%(key))
        exec("mcw = dem_imgs_ll.mcw%s"%(key))
        [aXMax,aXMin,aYMax,aYMin] =[max([_[0] for _ in pntPair.values()]),min([_[0] for _ in pntPair.values()]),max([_[1] for _ in pntPair.values()]),min([_[1] for _ in pntPair.values()])]
        for img in imglist.keys():
            exec("adem = dem_imgs_ll.%s"%(img))
            [ademLLX,ademLLY] =[min([_[0] for _ in adem]),min([_[1] for _ in adem])]

            if aXMin>=ademLLX:
                ulx = ademLLX+ int((aXMin-ademLLX)/mcw)*mcw
                uly = ademLLY+ int((aYMin-ademLLY)/mch)*mch
            else:
                ulx = ademLLX- math.ceil((ademLLX-aXMin)/mcw)*mcw
                uly = ademLLY- math.ceil((ademLLY-aYMin)/mch)*mch
            ele = arcpy.RasterToNumPyArray(imglist[img],arcpy.Point(aXMin,aYMin),math.ceil((aXMax-ulx)/mch),math.ceil((aYMax-uly)/mcw))
            if len(ele)==1 and ele[0,0] >-50:
                for key in pntPair.keys():
                    if key not in keysDone:
                        temp[key]=ele[0,0]
                        keysDone.append(key)
            elif len(ele)!=0:
                for key in pntPair.keys():
                    if key not in keysDone:
                        (x,y) = pntPair[key]
                        deltaX = x - ulx
                        deltaY = uly- y
                        arow =int(math.floor(deltaY/mch))
                        acol = int(math.floor(deltaX/mcw))
                        if ele[arow,acol] >-50:
                            temp[key]= ele[arow,acol]
                            keysDone.append(key)
    return [temp,[_ for _ in pntPair.keys() if _ not in keysDone]]

def getGoogleElevation((X,Y)):
    GOOGLE_URL = 'https://maps.googleapis.com/maps/api/elevation/json?locations='
    googlekey = r'AIzaSyBmub_p_nY5jXrFMawPD8jdU0DgSrWfBic'
    url = GOOGLE_URL + str(Y)+','+str(X) + '&key='+googlekey

    with contextlib.closing(urllib.urlopen(url)) as x:
        response = simplejson.load(x)
        try:
            elevation = response['results'][0]['elevation']
            return str(int(elevation))
        except KeyError:
            elevation = ''
            return elevation

if __name__ == '__main__':
    try:
        config = DEM()
        pntlist={}
        pntlist10={}
        pntlist30={}
        pntlist430={}
        pntOutput={}
        keys30=[]

        pntlist=eval(arcpy.GetParameterAsText(0))
        #pntlist={"1":[-77.47146845,42.99941746],"2":[-77.47457624,43.00008015],"3":[-77.4746448,42.99972976],"4":[-77.46952318,43.0002948],"5":[-77.46820239,42.9978862],"6":[-77.47449927,42.99680039]}
        for _ in pntlist.keys():
            pntlist[_] = eval((str(pntlist[_]).strip("[]")))

        # 1 create pnt shapefile
        pntShapefile = createSHP(pntlist)

        #3 find 10m DEM for multi Points key-value
        imgs10 = findDEMshp(pntShapefile,config.masterlyr_dem,config.imgdir_dem)

        if imgs10 !={}:
            #4 calculate Elevation for multi Points key-value
            [pntlist10,keys30] = getElevation(pntlist,imgs10,10)
            pntlist430={_: pntlist[_] for _ in keys30}
        else:
            #3 find 30m DEM for multi Points key-value
            imgs30 = findDEMshp(pntShapefile,config.masterlyr_demCA,config.imgdir_demCA)
            [pntlist30,_] = getElevation(pntlist,imgs30,30)

        if pntlist430 !={}:
            pntShapefile = createSHP(pntlist430)

            imgs30 = findDEMshp(pntShapefile,config.masterlyr_demCA,config.imgdir_demCA)

            [pntlist30_1,_] = getElevation(pntlist430,imgs30,30) #[done, N/A]
            for _ in pntlist30_1.keys():
                pntlist30[_]= pntlist30_1[_]
            del imgs30

        for _ in pntlist.keys():
            if _ in pntlist10.keys():
                pntOutput[_]=pntlist10[_]
            elif _ in pntlist30.keys():
                pntOutput[_]=pntlist30[_]
            else:
                pntOutput[_]=''
                X = pntlist[_][0]
                Y =  pntlist[_][1]
                pntOutput[_]=getGoogleElevation((X,Y))

        arcpy.SetParameter(1,pntOutput)
        print pntOutput
        for _ in pntlist.keys():
            arcpy.AddMessage("{0}: {1}".format(_, pntOutput[_]))

        del imgs10,pntlist,pntlist10,pntlist30,pntlist430,_,keys30

    except Exception as e:
        tb = sys.exc_info()[2]
        tbinfo = traceback.format_tb(tb)[0]
        pymsg = "PYTHON ERRORS:\nTraceback info:\n %s \nError Info:\n %s"%(tbinfo,str(sys.exc_info()[1]))
        msgs = "ArcPy ERRORS:\n %s\n"%arcpy.GetMessages(2)
        arcpy.AddError("hit CC's error code in except: ")
        arcpy.AddError(pymsg)
        arcpy.AddError(msgs)
        raise











