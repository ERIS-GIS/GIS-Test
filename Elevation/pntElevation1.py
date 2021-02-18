#
# Author:      jloucks
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

US_DEM = r"US_DEM"
class Machine:
    data_server = r"\\cabcvan1fpr009"
class DEM():
    def __init__(self):
        self.imgdir_dem = os.path.join(Machine.data_server,US_DEM,"DEM13")
        self.imgdir_demCA = os.path.join(Machine.data_server,US_DEM,"DEM1")
        self.masterlyr_dem = os.path.join(Machine.data_server,US_DEM,"DEM13.shp")
        self.masterlyr_demCA = os.path.join(Machine.data_server,US_DEM,"DEM1.shp")
def findDEMmath(masterGrids,(X,Y)): # one point at once
    temp={}
    for cell in masterGrids:
        [XMax,XMin,YMax,YMin] =[max([_[0] for _ in cell[0]]),min([_[0] for _ in cell[0]]),max([_[1] for _ in cell[0]]),min([_[1] for _ in cell[0]])]
        if (XMin <X < XMax and YMin< Y< YMax) :
            if cell[2] not in temp.keys():
                temp[cell[2]]=[cell[1]]
            elif cell[1] not in temp[cell[2]]:
                temp[cell[2]].append(cell[1])
    del XMax,XMin,YMax,YMin,cell,X,Y
    return temp
def getSingleElevation((X,Y),imglist,path,key):
    if imglist!=[]:
        exec("mch = dem_imgs_ll.mch%s"%(key))
        exec("mcw = dem_imgs_ll.mcw%s"%(key))
        for img in imglist:
            imgpath = os.path.join(path,img)
            ele = arcpy.RasterToNumPyArray(imgpath,arcpy.Point(X,Y),1,1)
            if len(ele)==1 and ele[0,0] >-50:
                return ele[0,0]
            del ele,imgpath,img
    return None
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
        input = '[[-79.38736389999991, 43.702032600000166]]'
        xy=eval((str(input).strip("[]")))
        # xy=eval((str(arcpy.GetParameterAsText(0)).strip("[]")))
        config = DEM()
        # 1 read Module
        masterGrids = dem_footprints.dem_masterGrids
        #2 find DEM for One Point
        imgs = findDEMmath(masterGrids,xy)
        del masterGrids
        elevation = None
        if 10 in imgs.keys():
            #3 Calculate Elevation based on 10m collection
            elevation = getSingleElevation(xy,imgs[10],config.imgdir_dem,10)

        if elevation is None and 30 in imgs.keys():
            # 3.1 Calculate Elevation based on 30m collection
            elevation = getSingleElevation(xy,imgs[30],config.imgdir_demCA,30)

        if elevation is None:
            elevation = getGoogleElevation(xy)
        
        if elevation is None:
            elevation = ''

        arcpy.AddMessage("{0}: {1}".format('Elevation Output', elevation))
        arcpy.SetParameter(1,elevation)

    except Exception as e:
        tb = sys.exc_info()[2]
        tbinfo = traceback.format_tb(tb)[0]
        pymsg = "PYTHON ERRORS:\nTraceback info:\n %s \nError Info:\n %s"%(tbinfo,str(sys.exc_info()[1]))
        msgs = "ArcPy ERRORS:\n %s\n"%arcpy.GetMessages(2)
        arcpy.AddError("hit CC's error code in except: ")
        arcpy.AddError(pymsg)
        arcpy.AddError(msgs)
        raise











