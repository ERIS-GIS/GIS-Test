#-------------------------------------------------------------------------------
# Name:        module1
# Purpose:
#
# Author:      cchen
#
# Created:     03/03/2017
# Copyright:   (c) cchen 2017
# Licence:     <your licence>
#-------------------------------------------------------------------------------
import sys, os, string, arcpy, logging
from arcpy import env, mapping
import json as simplejson
import urllib

## if Elevation field does not exist, create one and insert value. project the shapefile to be GCS 83

def getElevation(path="36.578581,-118.291994",sensor="false", **elvtn_args):
        elevationArray_tem = []
        ELEVATION_BASE_URL = 'https://maps.googleapis.com/maps/api/elevation/json?key=AIzaSyBmub_p_nY5jXrFMawPD8jdU0DgSrWfBic'
        elvtn_args.update({
           'locations': path,
           'sensor': sensor
         })
        url = ELEVATION_BASE_URL + '&' + urllib.urlencode(elvtn_args)
        response = simplejson.load(urllib.urlopen(url))

        if response['status'] != 'OK':
             arcpy.AddMessage('Google Elevation returns NOT OK status')
             raise Exception("Google elevation status is " + response['status'])
             #print 'Elevation is blocked, should fail here' + 1

        for resultset in response['results']:
            elevationArray_tem.append(resultset['elevation'])

        return elevationArray_tem


try:
# scratch#
        scratch = arcpy.env.scratchWorkspace
        #scratch = r"E:\GISData_testing\Christy\scratch\google"
################ static ##############################
        connectionPath = r"E:\GISData\ERISReport\ERISReport\PDFToolboxes"
        srGCS83 = arcpy.SpatialReference(4269)
        arcpy.env.OverWriteOutput = True
        all_merge = arcpy.GetParameter(0)
        all_merge = str(all_merge)

        check_field = arcpy.ListFields(all_merge,"Elevation")
        if len(check_field)==0:
            arcpy.AddField_management(all_merge, "Elevation", "DOUBLE", "12", "6", "", "", "NULLABLE", "NON_REQUIRED", "")
        locationArray = []
        elevationArray = []
        elevationArray1 = []
        elevationArray11 = []

        all_merge_1 = all_merge[:-4]+"_google.shp"
        arcpy.Project_management(all_merge,all_merge_1,srGCS83)
        all_merge = all_merge_1
        del all_merge_1

        rows = arcpy.da.SearchCursor(all_merge,["ID","SHAPE@XY","Elevation"])
        for row in rows:
            xCoord, yCoord = row[1]
            xID = row[0]
            locationArray.append([xID, str(yCoord)+','+str(xCoord)+'|'])
        del row
        del rows

        a =[0]
        n= 0
        result = int(arcpy.GetCount_management(all_merge).getOutput(0))
        for j in range(1,result):
            if (j%50==0):
                n= n+1
                a.append(j)
        n=n+1
        a.append(result)


        for k in range(0,n):
             s = ''
             for l in range(a[k], a[k+1]):
                s=s+ locationArray[l][1]
             s = s[:-1]
             #print "i = 0, call Google" + locationArray[l][1]
             if s != '':
                elevationArray1= getElevation(s)
                elevationArray11 += elevationArray1
        for j in range (0, result):
            elevationArray.append([locationArray[j][0],elevationArray11[j]])

        t = 0
        rows = arcpy.da.UpdateCursor(all_merge,["ID","Elevation"])
        for row in rows:
            for t in range(0,len(elevationArray)):
                if row[0] ==elevationArray[t][0]:
                    row[1] = elevationArray[t][1]
                    rows.updateRow(row)
        del row
        del rows

        arcpy.SetParameter(1,all_merge)


except:
   # If an error occurred, print the message to the screen
   arcpy.AddMessage(arcpy.GetMessages())
   raise





