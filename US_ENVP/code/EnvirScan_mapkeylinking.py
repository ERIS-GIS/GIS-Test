# Esri start of added imports
import sys, os, arcpy
# Esri end of added imports

# Esri start of added variables
g_ESRI_variable_1 = u'\\\\cabcvan1gis006\\GISData\\ERISReport\\ERISReport\\PDFToolboxes\\layer\\Streets\\Streets_US.lyr'
g_ESRI_variable_2 = u'\\\\cabcvan1gis006\\GISData\\ERISReport\\ERISReport\\PDFToolboxes\\layer\\Streets\\Streets_CA.lyr'
g_ESRI_variable_3 = u'in_memory/temp'
g_ESRI_variable_4 = u'in_memory/temp1'
g_ESRI_variable_5 = u'!SHAPE.CENTROID.X!'
g_ESRI_variable_6 = u'!SHAPE.CENTROID.Y!'
g_ESRI_variable_7 = u'xCentroid'
g_ESRI_variable_8 = u'yCentroid'
g_ESRI_variable_9 = u'UTM'
g_ESRI_variable_10 = u'page'
g_ESRI_variable_11 = u'Lon_X'
g_ESRI_variable_12 = u'!POINT_X!'
g_ESRI_variable_13 = u'Lat_Y'
g_ESRI_variable_14 = u'!POINT_Y!'
g_ESRI_variable_15 = u'Buffer'
g_ESRI_variable_16 = u'"Onsite"'
g_ESRI_variable_17 = u' "FID_orderg" = -1'
g_ESRI_variable_18 = u'"SmallBuffer"'
g_ESRI_variable_19 = u'Id'
g_ESRI_variable_20 = u'\\\\cabcvan1gis006\\GISData\\PSR\\python'
g_ESRI_variable_21 = u'Elevation'
g_ESRI_variable_22 = u'"MapKeyTot" = 1'
g_ESRI_variable_23 = u'ERISID'
g_ESRI_variable_24 = u'!ERIS_ID!'
g_ESRI_variable_25 = u'in_memory\\tempP'
g_ESRI_variable_26 = u'in_memory\\tempI'
# Esri end of added variables

# Import required modules
import shutil, csv,json
import cx_Oracle, urllib, glob
import arcpy, os, numpy
from datetime import datetime
import getDirectionText
import gc, time
import traceback
from numpy import gradient
from numpy import arctan2, arctan, sqrt
import EnvirScan_config
import arcpy
## function to form the datasources list
def ds(DataSources):
    DataSources = DataSources.replace(" ","")
    DataSources = DataSources.replace("'","")
    ds1= DataSources.split(',')
    sourceC= len(ds1)
    where_clause= ''
    for i in range(0,sourceC,1):
        where_clause = where_clause+'"SOURCE" = \''+ds1[i] + '\' OR '
    where_clause= where_clause[:-4]
    return where_clause

def getStreetList(query, unit1):

    if country== 'US':
        streetlyr = g_ESRI_variable_1
        streetFieldName = "FULLNAME"
    else:
        streetlyr = g_ESRI_variable_2
        streetFieldName = "STREET"

    try:
        streetLayer = arcpy.mapping.Layer(streetlyr)
        clippedStreet= g_ESRI_variable_3       # had to use clippedStreet in 10.0 for speed
        arcpy.Clip_analysis(streetLayer, query, clippedStreet)

    except:
        raise

    streetList = ""
    nSelected = int(arcpy.GetCount_management(clippedStreet).getOutput(0))      #note int() here is EXTREMELY necessary
    if nSelected == 0 :
        streetList = "***"
    else:
        streetArray = []

        rows = arcpy.SearchCursor(clippedStreet)
        for row in rows:
            value = row.getValue(streetFieldName)
            if(value.strip() != ''):
                streetArray.append(value.upper())
        streetSet = set(streetArray)
        streetList = '|'.join(streetSet)
        del row
        del rows
    return streetList


def calMapkey(fclass):
  arcpy.env.OverWriteOutput = True

  temp = arcpy.CopyFeatures_management(fclass, g_ESRI_variable_4)
  cur = arcpy.UpdateCursor(temp,"Distance = 0" ,"","Dist_cent; Distance; MapKeyLoc; MapKeyNo", 'Dist_cent A; Source A')

  lastMapkeyloc = 0
  row = cur.next()
  if row is not None:
      last = row.getValue('Dist_cent') # the last value in field A
      #print str(row.getValue('Distance')) +", " + str(last)
      row.setValue('mapkeyloc', 1)
      row.setValue('mapkeyno', 1)

      cur.updateRow(row)
      run = 1 # how many values in this run
      count = 1 # how many runs so far, including the current one

    # the for loop should begin from row 2, since
    # cur.next() has already been called once.
      for row in cur:
        current = row.getValue('Dist_cent')
        #print str(row.getValue('Distance')) + ", " + str(current)
        if current == last:
            run += 1
        else:
            run = 1
            count += 1
        row.setValue('mapkeyloc', count)
        row.setValue('mapkeyno', run)
        cur.updateRow(row)

        last = current
      lastMapkeyloc = count
    # release the layer from locks
  del cur
  if 'row' in locals():
        del row

  cur = arcpy.UpdateCursor(temp,"Distance > 0" ,"","Distance; MapKeyLoc; MapKeyNo", 'Distance A; Source A')

  row = cur.next()
  if row is not None:
      last = row.getValue('Distance') # the last value in field A
      #print "Part 2 start " + str(last)+ "   lastmaykeyloc is " + str(lastMapkeyloc)
      row.setValue('mapkeyloc', lastMapkeyloc + 1)
      row.setValue('mapkeyno', 1)


      cur.updateRow(row)
      run = 1 # how many values in this run
      count = lastMapkeyloc + 1 # how many runs so far, including the current one

    # the for loop should begin from row 2, since
    # cur.next() has already been called once.
      for row in cur:
        current = row.getValue('Distance')
        #print "Part 2 start " + str(last)+ "   lastmapkeyloc is " + str(lastMapkeyloc)
        if current == last:
            run += 1
        else:
            run = 1
            count += 1
        row.setValue('mapkeyloc', count)
        row.setValue('mapkeyno', run)
        cur.updateRow(row)

        last = current

    # release the layer from locks
  del cur
  if 'row' in locals():
        del row

  cur = arcpy.UpdateCursor(temp, "", "", 'MapKeyLoc; mapKeyNo; MapkeyTot', 'MapKeyLoc D; mapKeyNo D')

  row = cur.next()
  if row is not None:
      last = row.getValue('mapkeyloc') # the last value in field A
      max= 1
      row.setValue('mapkeytot', max)
      cur.updateRow(row)

      for row in cur:
        current = row.getValue('mapkeyloc')

        if current < last:
            max= 1
        else:
            max= 0
        row.setValue('mapkeytot', max)
        cur.updateRow(row)

        last = current

# release the layer from locks
  del cur
  if 'row' in locals():
    del row
  arcpy.CopyFeatures_management(temp, fclass)

try:

    starttime = time.time()
    print starttime
    # Setting the environment

    arcpy.env.OverWriteOutput = True
    arcpy.env.overwriteOutput = True

    # ----------------
    # Set local variables.
    # ----------------
#    scratchfolder = arcpy.env.scratchFolder
#    env.workspace = scratchfolder

    connectionPath = EnvirScan_config.connectionPath
    outPointFileName = "SiteMarker.shp"
    outPointPR = "SiteMarkerPR.shp"

    #to access the projection files
    srGCS83 = arcpy.SpatialReference(4269)
    srCanadaAlbers = arcpy.SpatialReference(102001)
    WGS84 = arcpy.SpatialReference(4326)
    srUSAlbers = arcpy.SpatialReference(102003)

    # Pull in the Geoporcessing parameters in to local valiables.

#------------------------------------------------------------------------------------------------------------------
    # Pull in the Geoporcessing parameters in to local valiables.
    OrderIDText = arcpy.GetParameterAsText(0)
    scratch = arcpy.env.scratchWorkspace
    # OrderIDText = '1014924'
    # scratch = r"\\cabcvan1gis005\MISC_DataManagement\_AW\ENVP_US_scratchy\test_test"
    Buffer1 = "0.125"
#------------------------------------------------------------------------------------------------------------------    
    count1 = 0
    countI = 0
    countP = 0
    country = ''
    try:
        con = cx_Oracle.connect(EnvirScan_config.connectionString)
        cur = con.cursor()

        cur.execute("select order_num, address1, city, provstate, country from orders where order_id =" + OrderIDText)
        t = cur.fetchone()

        OrderNumText = str(t[0])
        AddressText = str(t[1])+", "+str(t[2])+", "+str(t[3])
        country = str(t[4])

        cur.execute("select geometry_type, geometry,radius_type from eris_order_geometry where order_id =" + OrderIDText)
        t = cur.fetchone()
        OrderType = str(t[0])
        OrderCoord = eval(str(t[1]))
        Radiustype = str(t[2])

        cur.execute("select small_radius from orders where order_id=" + OrderIDText)
        t = cur.fetchone()
        Buffer1 = str(t[0])
    finally:
        cur.close()
        con.close()

    ## We are trying to get a large buffer to clip all the base layers
    MaxBuffer= 2.0*float(Buffer1)
    nBuffers = 1
    Bufferout = str(MaxBuffer)
    #1. Create a shape file based on the OrderType


    point = arcpy.Point()
    array = arcpy.Array()
    sr = arcpy.SpatialReference()
# The country codes are defined by JOHN CAMPBELL; 9036 is Canada and 9093 is USA
    #if str(unit1).lower()== '9093':
    unit = " MILES"
    #else:
    #    unit = " KILOMETERS"

    sr.factoryCode = 4269
    sr.XYTolerance = .00000001
    sr.scaleFactor = 2000
    sr.create()
    featureList = []
    for feature in OrderCoord:
        # For each coordinate pair, set the x,y properties and add to the Array object.
        for coordPair in feature:
            point.X = coordPair[0]
            point.Y = coordPair[1]
            sr.setDomain (point.X, point.X, point.Y, point.Y)
            array.add(point)
        if OrderType.lower()== 'point':
            feaERIS = arcpy.Multipoint(array, sr)
        elif OrderType.lower() =='polyline':
            feaERIS  = arcpy.Polyline(array, sr)
        else :
            feaERIS = arcpy.Polygon(array,sr)
        array.removeAll()

        # Append to the list of Polygon objects
        featureList.append(feaERIS)

     # Create a copy of the Polygon objects, by using featureList as input to the CopyFeatures tool.
    outshp= os.path.join(scratch,"orderGeoName.shp")
    #outshp = r"in_memory/orderGeoName"

    arcpy.CopyFeatures_management(featureList, outshp)
    arcpy.DefineProjection_management(outshp, srGCS83)

    del point
    del array

#2. Calculate Centroid of Geometry
    passLockTest = arcpy.TestSchemaLock(outshp)
    while not passLockTest:
        arcpy.AddWarning("There is a lock, wait, in #2")
        time.sleep(10)
        passLockTest = arcpy.TestSchemaLock(outshp)
    arcpy.AddField_management(outshp, "xCentroid", "DOUBLE", 18, 11)
    arcpy.AddField_management(outshp, "yCentroid", "DOUBLE", 18, 11)
    arcpy.AddField_management(outshp, "ERIS_ID","LONG",10)

    xExpression = g_ESRI_variable_5
    yExpression = g_ESRI_variable_6

    arcpy.CalculateField_management(outshp, g_ESRI_variable_7, xExpression, "PYTHON_9.3")
    arcpy.CalculateField_management(outshp, g_ESRI_variable_8, yExpression, "PYTHON_9.3")

#3.  Create a shapefile of centroid

    in_rows = arcpy.SearchCursor(outshp)
    outPointSHP = os.path.join(scratch, outPointFileName)
    #outPointSHP = r'in_memory/SiteMarker'
    point1 = arcpy.Point()
    array1 = arcpy.Array()

    arcpy.CreateFeatureclass_management(scratch, outPointFileName, "POINT", "", "DISABLED", "DISABLED", srGCS83)
    #arcpy.CreateFeatureclass_management('in_memory', 'SiteMarker', "POINT", "", "DISABLED", "DISABLED", srGCS83)

    cursor = arcpy.InsertCursor(outPointSHP)
    feat = cursor.newRow()

    for in_row in in_rows:
        # Set X and Y for start and end points
        point1.X = in_row.xCentroid
        point1.Y = in_row.yCentroid
        array1.add(point1)

        centerpoint = arcpy.Multipoint(array1)
        array1.removeAll()

        feat.shape = point1
        cursor.insertRow(feat)
    del feat
    del cursor
    del in_row
    del in_rows
    del point1
    del array1
    passLockTest = arcpy.TestSchemaLock(outPointSHP)
    while not passLockTest:
        arcpy.AddWarning("There is a lock, wait, before #4")
        time.sleep(10)
        passLockTest = arcpy.TestSchemaLock(outPointSHP)
    arcpy.AddXY_management(outPointSHP)

#4. Steps to know UTM zone and repoject centroid point to UTM projection (Northing and Easting )

     # Process: Add Field to store UTM information
    passLockTest = arcpy.TestSchemaLock(outPointSHP)
    while not passLockTest:
        arcpy.AddWarning("There is a lock, wait, in #4")
        time.sleep(10)
        passLockTest = arcpy.TestSchemaLock(outPointSHP)
    arcpy.AddField_management(outPointSHP, "UTM", "TEXT", "", "", "1500", "", "NULLABLE", "NON_REQUIRED", "")
     # Process: Calculate UTM Zone- you need to UTM as a value to reproject
    arcpy.CalculateUTMZone_cartography(outPointSHP, g_ESRI_variable_9)
    UT= arcpy.SearchCursor(outPointSHP)
    for row in UT:
        UTMvalue = str(row.getValue('UTM'))[41:43]
    del UT
    del row
    # Process: Project
    if len(UTMvalue)==1:
        UTMcode = int('2690'+UTMvalue)
    elif len(UTMvalue)==2:
        UTMcode = int('269'+UTMvalue)
    else:
        UTMcode = 26901

    out_coordinate_system = arcpy.SpatialReference(UTMcode)
    projPointSHP = os.path.join(scratch, outPointPR)
    arcpy.Project_management(outPointSHP, projPointSHP, out_coordinate_system)
    arcpy.AddField_management(projPointSHP, "page", "SHORT")
    arcpy.CalculateField_management(projPointSHP, g_ESRI_variable_10, '{0}'.format(2), "PYTHON_9.3", "")


#5. Calculate UTM coordinates for XML as separate columns

     # Process: Add Field
    passLockTest = arcpy.TestSchemaLock(projPointSHP)
    while not passLockTest:
        arcpy.AddWarning("There is a lock, wait, in #5")
        time.sleep(10)
        passLockTest = arcpy.TestSchemaLock(projPointSHP)
    arcpy.AddField_management(projPointSHP, "Lon_X", "DOUBLE", "12", "6", "", "", "NULLABLE", "NON_REQUIRED", "")
    arcpy.AddField_management(projPointSHP, "Lat_Y", "DOUBLE", "12", "6", "", "", "NULLABLE", "NON_REQUIRED", "")
    arcpy.AddField_management(projPointSHP, "ERIS_ID","LONG",10)

     # Process: Calculate Field- XY Coordinates- UTM Northing and UTM easting in meters
    arcpy.CalculateField_management(projPointSHP, g_ESRI_variable_11, g_ESRI_variable_12, "PYTHON_9.3", "")
    arcpy.CalculateField_management(projPointSHP, g_ESRI_variable_13, g_ESRI_variable_14, "PYTHON_9.3", "")

    passLockTest = arcpy.TestSchemaLock(projPointSHP)
    while not passLockTest:
        arcpy.AddWarning("There is a lock, wait, before #6")
        time.sleep(10)
        passLockTest = arcpy.TestSchemaLock(projPointSHP)
    arcpy.AddXY_management(projPointSHP)

#7. Create Buffers (can be centre or edge for polygons) and intersect with ERIS sites and appending to one shapefile
#for multiple buffers- calculate the area of rings or polygon and make selct based on datasources; Also add the Buffer or onsite values to field
     # Create the Buffers-
     # for data sources for final selection

    X_YTolerance = "0.3 Meters"
    projorderSHP = os.path.join(scratch, "ordergeoNamePR.shp")
    arcpy.Project_management(outshp, projorderSHP, out_coordinate_system)
    outshp= projorderSHP
    ERIS= EnvirScan_config.ERIScanPoint


    if OrderType.lower()== 'polygon':
        ERIS_sel= os.path.join(scratch,"ERISPoly"+OrderIDText+".shp")
        arcpy.Clip_analysis(ERIS, outshp, ERIS_sel)
        arcpy.AddField_management(ERIS_sel, "Buffer", "TEXT", "", "", "20", "", "NULLABLE", "NON_REQUIRED", "")
        arcpy.CalculateField_management(ERIS_sel, g_ESRI_variable_15, g_ESRI_variable_16, "PYTHON_9.3", "")
        ERIS_select = ERIS_sel
        streetList0 = getStreetList(outshp,country)

    if OrderType.lower()== 'polygon':
        if Radiustype.lower()== 'centre' :
             outshp= projPointSHP

    if Buffer1!= '0':
        outBuffer1FileName = "Buffer1.shp"
        buffer1Distance = Buffer1 + unit
        outBuffer = os.path.join(scratch, outBuffer1FileName)
        arcpy.Buffer_analysis(outshp, outBuffer, buffer1Distance)
        arcpy.DefineProjection_management(outBuffer, out_coordinate_system)
        ERIS_select= os.path.join(scratch,"ERIS"+OrderIDText+".shp")
        if OrderType.lower()== 'polygon':
            outBufferjoin = os.path.join(scratch, "Bufferjoin.shp")
            arcpy.Union_analysis ([[projorderSHP,2], [outBuffer,1]],outBufferjoin)
            outBuffersel = os.path.join(scratch, "Buffersel.shp")
            arcpy.Select_analysis(outBufferjoin,outBuffersel,g_ESRI_variable_17)
            arcpy.Clip_analysis(ERIS, outBuffersel, ERIS_select)
        else:
            arcpy.Clip_analysis(ERIS, outBuffer, ERIS_select, X_YTolerance)
        arcpy.AddField_management(ERIS_select, "Buffer", "TEXT", "", "", "20", "", "NULLABLE", "NON_REQUIRED", "")
        arcpy.CalculateField_management(ERIS_select, g_ESRI_variable_15, g_ESRI_variable_18, "PYTHON_9.3", "")
        ##polygon reports has either one or no buffer so appending to what is inside the polygon
        if OrderType.lower()== 'polygon':
            arcpy.Append_management(ERIS_sel, ERIS_select)
        streetList0 = getStreetList(outBuffer,country)

#8 add orderID and UTM zone to projected centroid
    arcpy.DeleteField_management(projPointSHP,g_ESRI_variable_19)
    arcpy.AddField_management(projPointSHP,"ID","Long", "10")
    cur = arcpy.UpdateCursor(projPointSHP)
    for row in cur:
        print(OrderIDText)
        row.setValue("ID", float(OrderIDText))
        cur.updateRow(row)
    del row
    del cur

    UT= arcpy.SearchCursor(projPointSHP)
    for row in UT:
        UTMvalue = str(row.getValue('UTM'))[41:43]
    del UT
    del row

#9 Project ERIS sites with the UTM projection

    ERISPR = os.path.join(scratch,"ERISPR.shp")
    arcpy.Project_management(ERIS_select, ERISPR, out_coordinate_system )

#9-10 Remove possible duplicated sites due to selection with donut shape and tolerance
    sites = arcpy.UpdateCursor(ERISPR)
    siteIDs = []
    for row in sites:
        siteID = int(row.getValue("ID_CHAR"))   #get the ERIS site ID
        #logger.info("processing siteID below " + str(siteID))
        if siteID in siteIDs:
            sites.deleteRow(row)
            #logger.info("delete row with ID: " + str(siteID))
        else:
            siteIDs.append(int(row.getValue("ID_CHAR")))
            #logger.info("stored siteID " + str(int(row.getValue("ID"))) + " for search")
    #del row
    del sites

#6. Calculate Elevation of centroid for XML as separate columns- Elevation script is part of ERIS toolbox
# img directory

    ERIS_clipcopy= os.path.join(scratch,"ERISCC.shp")
    arcpy.CopyFeatures_management(ERISPR, ERIS_clipcopy)
    arcpy.Integrate_management(ERIS_clipcopy, ".3 Meters")

#10. Calculate Distance with integration and spatial join- can be easily done with Distance tool along with direction if ArcInfo or Advanced license

    arcpy.ImportToolbox(os.path.join(g_ESRI_variable_20,"ERIS.tbx"))
    projPointSHP = arcpy.inhouseElevation_ERIS(projPointSHP).getOutput(0)

    elevationArray=[]
    Call_Google = ''
    rows = arcpy.SearchCursor(projPointSHP)
    for row in rows:
        #print row.Elevation
        if row.Elevation == -999:
            Call_Google = 'YES'
            break
    del row
    del rows

    if Call_Google == 'YES':
        projPointSHP = arcpy.googleElevation_ERIS(projPointSHP).getOutput(0)
    projPointSHP_final = os.path.join(scratch,outPointPR[:-4]+"_PR.shp")
    arcpy.Project_management(projPointSHP,projPointSHP_final,out_coordinate_system)
    projPointSHP = projPointSHP_final

    arcpy.AddField_management(outshp, "Elevation", "DOUBLE", "12", "6", "", "", "NULLABLE", "NON_REQUIRED", "")
    # copy the elevation of the center point to the orderGeometry (could be polygon, polyline)
    cursor = arcpy.SearchCursor(projPointSHP)
    row = cursor.next()
    elev_marker = row.getValue("Elevation")
    del cursor
    del row
    arcpy.CalculateField_management(outshp, g_ESRI_variable_21, '{0}'.format(eval(str(elev_marker))), "PYTHON_9.3", "")

    ERIS_sj= os.path.join(scratch,"ERISSA.shp")
    ERIS_sja= os.path.join(scratch,"ERISSAT_temp.shp")
    arcpy.SpatialJoin_analysis(ERIS_clipcopy, outshp, ERIS_sj, "JOIN_ONE_TO_MANY", "KEEP_ALL","#", "CLOSEST","5000 Kilometers", "Distance")   # this is the reported distance
    arcpy.SpatialJoin_analysis(ERIS_sj,projPointSHP,ERIS_sja, "JOIN_ONE_TO_MANY", "KEEP_ALL","#", "CLOSEST","5000 Kilometers", "Dist_cent")
    ERIS_sja_final= os.path.join(scratch,"ERISSAT.shp")
    arcpy.ImportToolbox(os.path.join(g_ESRI_variable_20,"ERIS.tbx"))
    ERIS_sja = arcpy.inhouseElevation_ERIS(ERIS_sja).getOutput(0)

    elevationArray=[]
    Call_Google = ''
    rows = arcpy.SearchCursor(ERIS_sja)
    for row in rows:
        #print row.Elevation
        if row.Elevation == -999:
            Call_Google = 'YES'
            break
    if 'row' in locals():
        del row
    del rows

    if Call_Google == 'YES':
        ERIS_sja = arcpy.googleElevation_ERIS(ERIS_sja).getOutput(0)
    else:
        ERIS_sja = ERIS_sja
    # Return tool error messages for use with a script tool
    #print elevationArray
    arcpy.Project_management(ERIS_sja,ERIS_sja_final,out_coordinate_system)

#11. Add mapkey with script from ERIS toolbox
    arcpy.AddField_management(ERIS_sja_final, "MapKeyNo", "SHORT", "", "", "", "", "NULLABLE", "NON_REQUIRED", "")
     # Process: Add Field for mapkey rank storage based on location and total number of keys at one location
    arcpy.AddField_management(ERIS_sja_final, "MapKeyLoc", "SHORT", "", "", "", "", "NULLABLE", "NON_REQUIRED", "")
    arcpy.AddField_management(ERIS_sja_final, "MapKeyTot", "SHORT", "", "", "", "", "NULLABLE", "NON_REQUIRED", "")
    calMapkey(ERIS_sja_final)

    arcpy.AddField_management(ERIS_sja_final, "Direction", "TEXT", "", "", "3", "", "NULLABLE", "NON_REQUIRED", "")
    arcpy.AddField_management(ERIS_sja_final, "ERISID", "LONG", "", "", "", "", "NULLABLE", "NON_REQUIRED", "")
    desc = arcpy.Describe(ERIS_sja_final)
    shapefieldName = desc.ShapeFieldName
    rows = arcpy.UpdateCursor(ERIS_sja_final)
    
    for row in rows:
        erisid = int(row.getValue("ID_CHAR"))
        row.setValue("ERISID", float(erisid))
        if(row.Distance<0.001):  #give onsite, give "-" in Direction field
            directionText = '-'
        else:
            ref_x = row.POINT_X      #field is directly accessible
            ref_y = row.POINT_Y

            feat = row.getValue(shapefieldName)
            pnt = feat.getPart()
            directionText = getDirectionText.getDirectionText(ref_x,ref_y,pnt.X,pnt.Y)

        row.Direction = directionText #field is directly accessible
        rows.updateRow(row)
    if 'row' in locals():
        del row
    del rows
     # Process: Select
    ERIS_fin= os.path.join(scratch,"ErisClip1.shp")
    arcpy.Select_analysis(ERIS_sja_final, ERIS_fin, "\"MapKeyTot\" = 1")
    ERIS_disp= os.path.join(scratch,"ErisClip.shp")
    arcpy.Sort_management(ERIS_sja_final, ERIS_disp, [["MapKeyLoc", "ASCENDING"]])
#14
    # Process: xmlWriter
    xmlName= os.path.join(scratch,"XML"+OrderIDText+".xml")
    log= os.path.join(scratch,  "log"+OrderIDText+".txt")
    arcpy.ImportToolbox(os.path.join(g_ESRI_variable_20,"ERIS.tbx"))
    arcpy.xmlWriter_ERIS(projPointSHP, ERIS_sja_final,xmlName, log)

#write streetlist to Oracle
    try:
        con = cx_Oracle.connect(EnvirScan_config.connectionString)
        cur=con.cursor()

        cv1 = cur.var(cx_Oracle.STRING)
        cv1.setvalue(0,'STREET')

        cv2 = cur.var(cx_Oracle.CLOB)
        cv2.setvalue(0,str(streetList0))     #note it's important to have str()

        cv9 = cur.var(cx_Oracle.STRING)      #orderID
        cv9.setvalue(0,str(OrderIDText))

        r=cur.callfunc('ERIS_INSTANT.RunInstant',cx_Oracle.STRING,(cv1,cv2,cv9,))

    except:
        raise

    finally:
        cur.close()
        con.close()

    mxd = arcpy.mapping.MapDocument(EnvirScan_config.mxd)
    df = arcpy.mapping.ListDataFrames(mxd,"*")[0]
    df.spatialReference = out_coordinate_system

    bufferLayer = arcpy.mapping.Layer(EnvirScan_config.bufferlyrfile)
    bufferLayer.replaceDataSource(scratch,"SHAPEFILE_WORKSPACE","Buffer1")
    arcpy.mapping.AddLayer(df,bufferLayer,"Bottom")
#
    ds_oids = []
    rows = arcpy.SearchCursor(ERIS_sja_final)
    for row in rows:
        ds_oids.append(int(row.DS_OID))
    #count1 = len(ds_oids)
    IncidentDic = {}
    if 'row' in locals():
        del row
    del rows
    try:
        con = cx_Oracle.connect(EnvirScan_config.connectionString)
        cur = con.cursor()
        for ds_oid in list(set(ds_oids)):
            cur.execute("select AST.incident_permit from eris_data_source EDS,astm_type_reference AST where eds.ds_oid = %s and eds.at_oid = ast.at_oid and AST.incident_permit is not null"%ds_oid)
            t = cur.fetchone()
            if t != None:
                incident_permit = str(t[0])
                if incident_permit not in IncidentDic.keys():
                    IncidentDic[incident_permit]= [ds_oid]
                else:
                    temp = IncidentDic[incident_permit]
                    temp.append(ds_oid)
                    IncidentDic[incident_permit]= temp

    finally:
        cur.close()
        con.close()



    if 'PERMIT' in IncidentDic.keys():
        permitText=' '
        countP = sum([ ds_oids.count(a) for a in IncidentDic['PERMIT']])
        for ds_oid in IncidentDic['PERMIT']:
            permitText = permitText+ str(ds_oid)+' OR DS_OID ='
        permitText = permitText[:-12]
        ErisClip_permit= os.path.join(scratch,"ErisClip_permit.shp")
        arcpy.Select_analysis(ERIS_sja_final, "in_memory\\tempP", "\"DS_OID\" ="+permitText)
        arcpy.Select_analysis(g_ESRI_variable_25, ErisClip_permit, g_ESRI_variable_22)
        newLayerERIS = arcpy.mapping.Layer(EnvirScan_config.ERIScanPermit)
        newLayerERIS.replaceDataSource(scratch, "SHAPEFILE_WORKSPACE", "ErisClip_permit")
        arcpy.mapping.AddLayer(df, newLayerERIS, "TOP")
        arcpy.DeleteFeatures_management(g_ESRI_variable_25)
    if 'INCIDENT' in IncidentDic.keys():
        inciText=' '
        countI = sum([ ds_oids.count(a) for a in IncidentDic['INCIDENT']])
        for ds_oid in IncidentDic['INCIDENT']:
            inciText = inciText+ str(ds_oid)+' OR DS_OID ='
        inciText = inciText[:-12]
        ErisClip_incident= os.path.join(scratch,"ErisClip_incident.shp")
        arcpy.Select_analysis(ERIS_sja_final, "in_memory\\tempI", "\"DS_OID\" ="+inciText)
        arcpy.Select_analysis(g_ESRI_variable_26, ErisClip_incident, g_ESRI_variable_22)
        newLayerERIS1 = arcpy.mapping.Layer(EnvirScan_config.ERIScanIncident)
        newLayerERIS1.replaceDataSource(scratch, "SHAPEFILE_WORKSPACE", "ErisClip_incident")
        arcpy.mapping.AddLayer(df, newLayerERIS1, "TOP")
        arcpy.DeleteFeatures_management(g_ESRI_variable_26)

##add geometry on main Map
    if OrderType.lower()== 'point':
        geom=EnvirScan_config.orderGeomlyrfile_point
    elif OrderType.lower() =='polyline':
        geom=EnvirScan_config.orderGeomlyrfile_polyline
    else :
        geom=EnvirScan_config.orderGeomlyrfile_polygon
    newLayerordergeo = arcpy.mapping.Layer(geom)
    newLayerordergeo.replaceDataSource(scratch, "SHAPEFILE_WORKSPACE", "orderGeoName" )
    arcpy.mapping.AddLayer(df, newLayerordergeo , "TOP")
    print "## df.scale is " + str(df.scale)
    df.extent = newLayerordergeo.getSelectedExtent(False)
    print "### df.scale is " + str(df.scale)

    df.scale = 2500
    df.extent = bufferLayer.getSelectedExtent(False)
    AddressTextElement = arcpy.mapping.ListLayoutElements(mxd, "TEXT_ELEMENT", "AddressText")[0]
    AddressTextElement.text = " " + AddressText + ""
    count1 = countI+countP
    if count1 !=0:
        CountTextElement = arcpy.mapping.ListLayoutElements(mxd, "TEXT_ELEMENT", "Count")[0]
        CountTextElement.text = "This report found "+str(count1) + " environmental records within 1/8 mile of the property located at:  "
    if 'INCIDENT' in IncidentDic.keys():
        Incident_count = arcpy.mapping.ListLayoutElements(mxd, "TEXT_ELEMENT", "Incident_count")[0]
        Incident_count.text = " " + str(countI) + ""
    if 'PERMIT' in IncidentDic.keys():
        Permit_count = arcpy.mapping.ListLayoutElements(mxd, "TEXT_ELEMENT", "Permit_count")[0]
        Permit_count.text = " " + str(countP) + ""

    outputLayoutPDF1 = os.path.join(scratch, "map_" + OrderNumText + ".pdf")
    arcpy.mapping.ExportToPDF(mxd, outputLayoutPDF1, "PAGE_LAYOUT", 640, 480, 250, "BEST", "RGB", True, "ADAPTIVE", "VECTORIZE_BITMAP", False, True, "LAYERS_AND_ATTRIBUTES", True, 90)
    mxd.saveACopy(os.path.join(scratch, "mxd.mxd"))
    del mxd
    shutil.copy(os.path.join(scratch, "map_" + OrderNumText + ".pdf"), EnvirScan_config.report_path)#"\\cabcvan1obi002\ErisData\Reports\test\noninstant_reports")
    arcpy.SetParameterAsText(1, os.path.join(scratch, "map_" + OrderNumText + ".pdf"))

except:
    # Get the traceback object
    #
    tb = sys.exc_info()[2]
    tbinfo = traceback.format_tb(tb)[0]

    # Concatenate information together concerning the error into a message string
    #
    pymsg = "PYTHON ERRORS:\nTraceback info:\n" + tbinfo + "\nError Info:\n" + str(sys.exc_info()[1])
    msgs = "ArcPy ERRORS:\n" + arcpy.GetMessages(2) + "\n"

    # Return python error messages for use in script tool or Python Window
    #
    arcpy.AddError("hit CC's error code in except: ")
    arcpy.AddError(pymsg)
    arcpy.AddError(msgs)

    # Print Python error messages for use in Python / Python Window
    #
    print pymsg + "\n"
    print msgs

